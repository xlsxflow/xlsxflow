export type XmlToken =
  | { type: 'startElement', name: string, attributes: Record<string, string> }
  | { type: 'endElement', name: string }
  | { type: 'text', value: string };

function codePoint(cp: number, fallback: string): string {
  // Out-of-range references (e.g. &#x110000;) are left as written instead of aborting the parse
  return cp <= 0x10ffff ? String.fromCodePoint(cp) : fallback;
}

function unescapeXml(text: string): string {
  if (text.indexOf('&') === -1) return text;
  return text.replace(/&(?:amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (match) => {
    switch (match) {
      case '&amp;': return '&';
      case '&lt;': return '<';
      case '&gt;': return '>';
      case '&quot;': return '"';
      case '&apos;': return "'";
      default:
        if (match.startsWith('&#x')) return codePoint(parseInt(match.slice(3, -1), 16), match);
        return codePoint(parseInt(match.slice(2, -1), 10), match);
    }
  });
}

// XML end-of-line handling: literal CRLF / CR become LF (escaped &#13; survives, it's decoded later)
function normalizeEol(text: string): string {
  return text.indexOf('\r') === -1 ? text : text.replace(/\r\n?/g, '\n');
}

function stripNamespace(name: string): string {
  const colonIdx = name.indexOf(':');
  return colonIdx !== -1 ? name.slice(colonIdx + 1) : name;
}

type Pending = 'text' | 'tag' | 'cdata' | 'comment' | null;

// Emits one array of tokens per input chunk. Per-token stream chunks cost a promise round-trip
// each (~5 per cell), which dominated read time; batching removes that overhead.
export function createXmlBatchParser(): TransformStream<Uint8Array, XmlToken[]> {
  let decoder: TextDecoder | undefined;
  let buffer = '';
  let isFirstChunk = true;

  // When a construct spans chunks, remember how far it was already scanned (and the quote state
  // inside a tag) so the next chunk resumes there. Rescanning from the start is O(n^2) on a
  // hostile file with one huge tag / text node / CDATA section.
  let pending: Pending = null;
  let scanFrom = 0;
  let pendingQuote = '';

  return new TransformStream({
    transform(chunk, controller) {
      // XML parsers must read UTF-16 as well as UTF-8; a UTF-16 part starts with its byte-order mark
      decoder ??= new TextDecoder(chunk[0] === 0xff && chunk[1] === 0xfe ? 'utf-16le' : chunk[0] === 0xfe && chunk[1] === 0xff ? 'utf-16be' : 'utf-8');
      buffer += decoder.decode(chunk, { stream: true });
      if (isFirstChunk) {
        if (buffer.charCodeAt(0) === 0xFEFF) {
          buffer = buffer.slice(1);
        }
        isFirstChunk = false;
      }

      const out: XmlToken[] = [];
      let index = 0;
      const resumed = pending;
      pending = null;

      while (index < buffer.length) {
        const resumeAt = (kind: Pending, from: number) => (index === 0 && resumed === kind ? Math.max(from, scanFrom) : from);

        const tagStart = buffer.indexOf('<', resumeAt('text', index));
        if (tagStart === -1) {
          // No more tags: keep the rest (from index) for the next chunk
          pending = 'text';
          scanFrom = buffer.length;
          break;
        }

        if (tagStart > index) {
          out.push({ type: 'text', value: unescapeXml(normalizeEol(buffer.slice(index, tagStart))) });
        }

        if (buffer.startsWith('<![CDATA[', tagStart)) {
          const cdataEnd = buffer.indexOf(']]>', resumeAt('cdata', tagStart + 9) - (resumed === 'cdata' && index === 0 ? 2 : 0));
          if (cdataEnd === -1) {
            pending = 'cdata';
            scanFrom = buffer.length;
            index = tagStart; // incomplete: resume here next chunk
            break;
          }
          out.push({ type: 'text', value: normalizeEol(buffer.slice(tagStart + 9, cdataEnd)) }); // CDATA is not escaped
          index = cdataEnd + 3;
          continue;
        }

        if (buffer.startsWith('<!--', tagStart)) {
          const commentEnd = buffer.indexOf('-->', resumeAt('comment', tagStart + 4) - (resumed === 'comment' && index === 0 ? 2 : 0));
          if (commentEnd === -1) {
            pending = 'comment';
            scanFrom = buffer.length;
            index = tagStart; // incomplete: resume here next chunk
            break;
          }
          index = commentEnd + 3;
          continue;
        }

        // Find the closing '>' outside quoted attribute values
        const resuming = index === 0 && resumed === 'tag';
        let tagEnd = -1;
        let quote = resuming ? pendingQuote : '';
        for (let i = resuming ? Math.max(scanFrom, tagStart + 1) : tagStart + 1; i < buffer.length; i++) {
          const char = buffer.charCodeAt(i);
          if (quote) {
            if (char === quote.charCodeAt(0)) quote = '';
          } else if (char === 34 /* " */ || char === 39 /* ' */) {
            quote = buffer[i];
          } else if (char === 62 /* > */) {
            tagEnd = i;
            break;
          }
        }

        if (tagEnd === -1) {
          pending = 'tag';
          pendingQuote = quote;
          scanFrom = buffer.length;
          index = tagStart; // incomplete: resume here next chunk
          break;
        }

        const tagContent = buffer.slice(tagStart + 1, tagEnd);
        index = tagEnd + 1;

        if (tagContent.startsWith('/')) {
          out.push({ type: 'endElement', name: stripNamespace(tagContent.slice(1).trim()) });
        } else if (tagContent.startsWith('?')) {
          // XML declaration, ignore
        } else if (tagContent.startsWith('!')) {
          // DOCTYPE or other DTD declarations, ignore
        } else {
          // Parse start element and attributes; any XML whitespace may separate the name from attributes
          const isSelfClosing = tagContent.endsWith('/');
          const content = isSelfClosing ? tagContent.slice(0, -1) : tagContent;
          const ws = content.search(/\s/);
          const name = stripNamespace(ws === -1 ? content : content.slice(0, ws));
          const attributes: Record<string, string> = {};

          if (ws !== -1) readAttributes(content, ws, attributes);

          out.push({ type: 'startElement', name, attributes });
          if (isSelfClosing) out.push({ type: 'endElement', name });
        }
      }
      buffer = buffer.slice(index);
      scanFrom -= index;
      if (out.length) controller.enqueue(out);
    },
    flush(controller) {
      buffer += decoder?.decode() ?? '';
      // Handle trailing text if any
      if (buffer.length > 0 && pending !== 'tag' && pending !== 'cdata' && pending !== 'comment') {
         controller.enqueue([{ type: 'text', value: unescapeXml(normalizeEol(buffer)) }]);
      }
    }
  });
}

// name="value" pairs from i on, in one pass: a regex here backtracks quadratically on junk like <c aaaa…>
const isSpace = (c: string) => c === ' ' || c === '\n' || c === '\t' || c === '\r';
function readAttributes(s: string, i: number, attributes: Record<string, string>) {
  const n = s.length;
  while (i < n) {
    while (i < n && isSpace(s[i])) i++;
    const start = i;
    while (i < n && !isSpace(s[i]) && s[i] !== '=') i++;
    const name = s.slice(start, i);
    while (i < n && isSpace(s[i])) i++;
    if (s[i] !== '=') { if (i === start) i++; continue; }
    i++;
    while (i < n && isSpace(s[i])) i++;
    const quote = s[i];
    if (quote !== '"' && quote !== "'") continue;
    const end = s.indexOf(quote, i + 1);
    if (end === -1) return;
    if (/^[a-zA-Z0-9_:.-]+$/.test(name)) attributes[stripNamespace(name)] = unescapeXml(normalizeEol(s.slice(i + 1, end)));
    i = end + 1;
  }
}

// Token-at-a-time view of createXmlBatchParser, kept for API compatibility.
export function createXmlStreamParser(): { readable: ReadableStream<XmlToken>; writable: WritableStream<Uint8Array> } {
  const batches = createXmlBatchParser();
  const flatten = new TransformStream<XmlToken[], XmlToken>({
    transform(batch, controller) {
      for (const token of batch) controller.enqueue(token);
    }
  });
  return { writable: batches.writable, readable: batches.readable.pipeThrough(flatten) };
}
