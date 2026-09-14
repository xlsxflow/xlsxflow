export type XmlToken = 
  | { type: 'startElement', name: string, attributes: Record<string, string> }
  | { type: 'endElement', name: string }
  | { type: 'text', value: string };

export function createXmlStreamParser(): TransformStream<Uint8Array, XmlToken> {
  const decoder = new TextDecoder();
  let buffer = '';

  return new TransformStream({
    transform(chunk, controller) {
      buffer += decoder.decode(chunk, { stream: true });
      let index = 0;

      while (index < buffer.length) {
        const tagStart = buffer.indexOf('<', index);
        if (tagStart === -1) {
          // No more tags, save remaining string to buffer
          buffer = buffer.slice(index);
          break;
        }

        if (tagStart > index) {
          const text = buffer.slice(index, tagStart).trim();
          if (text) {
            controller.enqueue({ type: 'text', value: text });
          }
        }

        const tagEnd = buffer.indexOf('>', tagStart);
        if (tagEnd === -1) {
          // Tag is incomplete, wait for more chunks
          buffer = buffer.slice(tagStart);
          break;
        }

        const tagContent = buffer.slice(tagStart + 1, tagEnd);
        index = tagEnd + 1;

        if (tagContent.startsWith('/')) {
          controller.enqueue({ type: 'endElement', name: tagContent.slice(1).trim() });
        } else if (tagContent.startsWith('?')) {
          // XML declaration, ignore
        } else {
          // Parse start element and attributes
          const spaceIdx = tagContent.indexOf(' ');
          const isSelfClosing = tagContent.endsWith('/');
          const contentWithoutSlash = isSelfClosing ? tagContent.slice(0, -1) : tagContent;
          
          if (spaceIdx === -1 || (isSelfClosing && spaceIdx === tagContent.length - 1)) {
            const name = contentWithoutSlash.trim();
            controller.enqueue({ type: 'startElement', name, attributes: {} });
            if (isSelfClosing) controller.enqueue({ type: 'endElement', name });
          } else {
            const name = contentWithoutSlash.slice(0, spaceIdx).trim();
            const attrString = contentWithoutSlash.slice(spaceIdx).trim();
            const attributes: Record<string, string> = {};
            
            // Simple attribute regex parsing
            const attrRegex = /([a-zA-Z0-9_:-]+)\s*=\s*"([^"]*)"/g;
            let match;
            while ((match = attrRegex.exec(attrString)) !== null) {
              attributes[match[1]] = match[2];
            }
            
            controller.enqueue({ type: 'startElement', name, attributes });
            if (isSelfClosing) controller.enqueue({ type: 'endElement', name });
          }
        }
      }
      buffer = buffer.slice(index);
    },
    flush(controller) {
      buffer += decoder.decode();
      // Handle trailing text if any
      if (buffer.trim()) {
         controller.enqueue({ type: 'text', value: buffer.trim() });
      }
    }
  });
}
