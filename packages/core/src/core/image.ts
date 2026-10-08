import type { SheetImage } from './types';
import { colIndex, colLetter, MAX_COLUMNS, MAX_ROWS, partRelationships, unescapeXml, xmlElements } from './utils';

const EMU_PER_PX = 9525;

export interface ImageInfo { ext: 'png' | 'jpeg' | 'gif'; width: number; height: number }

// Format and pixel size from the file header
export function imageInfo(b: Uint8Array): ImageInfo {
  const be16 = (i: number) => (b[i] << 8) | b[i + 1];
  const be32 = (i: number) => ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0;
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    return { ext: 'png', width: be32(16), height: be32(20) };
  }
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
    return { ext: 'gif', width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    // Walk the marker segments to the frame header (SOF0-15, minus DHT/JPG/DAC)
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      if (marker === 0xff) { i++; continue; }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { ext: 'jpeg', width: be16(i + 7), height: be16(i + 5) };
      }
      i += 2 + be16(i + 2);
    }
    throw new Error('JPEG image has no frame header.');
  }
  throw new Error('Unsupported image: only PNG, JPEG and GIF can be embedded.');
}

// The pictures of a worksheet's drawing, in the writer's SheetImage shape. Charts and shapes are
// skipped. Pictures sharing a media file share one `data` array.
// Offsets inside the anchor cell are ignored, and absoluteAnchor pictures are skipped
export async function readSheetImages(
  readText: (path: string) => Promise<string>,
  readBytes: (path: string) => Promise<Uint8Array | undefined>,
  sheetPath: string,
): Promise<SheetImage[]> {
  const images: SheetImage[] = [];
  const media = new Map<string, Uint8Array | undefined>();
  for (const rel of await partRelationships(readText, sheetPath)) {
    if (rel.external || !rel.type.endsWith('/drawing')) continue;
    const xml = await readText(rel.path);
    const targets = new Map((await partRelationships(readText, rel.path)).filter(r => !r.external).map(r => [r.id, r.path]));
    for (const { whole: anchor, name: kind } of xmlElements(xml, 'twoCellAnchor|oneCellAnchor')) {
      const pic = xmlElements(anchor, 'pic').next().value?.whole;
      const embed = pic && /<(?:\w+:)?blip\b[^>]*?\s(?:\w+:)?embed="([^"]*)"/.exec(pic)?.[1];
      const path = embed && targets.get(embed);
      if (!path) continue;
      if (!media.has(path)) media.set(path, await readBytes(path));
      const data = media.get(path);
      if (!data) continue;

      const pos = (tag: string) => {
        const m = new RegExp(`<(?:\\w+:)?${tag}>([\\s\\S]*?)</(?:\\w+:)?${tag}>`).exec(anchor)?.[1] ?? '';
        const num = (name: string) => parseInt(new RegExp(`<(?:\\w+:)?${name}>(-?\\d+)<`).exec(m)?.[1] ?? '0', 10);
        return { col: num('col'), row: num('row'), offset: num('colOff') + num('rowOff') };
      };
      const from = pos('from');
      const at = colLetter(from.col) + (from.row + 1);
      const descr = pic && /<(?:\w+:)?cNvPr\b[^>]*?\sdescr="([^"]*)"/.exec(pic)?.[1];
      const altText = descr ? unescapeXml(descr) : undefined;
      let image: SheetImage;
      if (kind === 'twoCellAnchor') {
        // `to` is the cell the far edge lands in; at offset 0 the picture ends on the cell before it
        const to = pos('to');
        const end = to.offset === 0 ? { col: Math.max(from.col, to.col - 1), row: Math.max(from.row, to.row - 1) } : to;
        image = { data, range: `${at}:${colLetter(end.col)}${end.row + 1}` };
      } else {
        const ext = /<(?:\w+:)?ext\b[^>]*?\scx="(\d+)"[^>]*?\scy="(\d+)"/.exec(anchor);
        image = ext ? { data, at, width: Math.round(+ext[1] / EMU_PER_PX), height: Math.round(+ext[2] / EMU_PER_PX) } : { data, at };
      }
      if (altText) image.altText = altText;
      images.push(image);
    }
  }
  return images;
}


const escapeXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const cellPos = (ref: string) => {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  const col = m ? colIndex(m[1]) : -1, row = m ? parseInt(m[2], 10) - 1 : -1;
  if (col < 0 || col >= MAX_COLUMNS || row < 0 || row >= MAX_ROWS) throw new Error(`Invalid anchor cell "${ref}".`);
  return { col, row };
};
const marker = (tag: string, { col, row }: { col: number; row: number }) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

// Where a drawing object goes: stretched over cells, or at a cell with a size in pixels
export type Placement = { range: string } | { at: string; width?: number; height?: number };

export const DRAWING_NS = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

// One anchor (xdr: prefix) holding `body`. `natural` is the size used when the placement gives
// none; a missing side keeps its aspect ratio. `attrs` lands on the anchor element, e.g. DRAWING_NS
// when the anchor is added to a drawing that declares other prefixes.
export function anchorXml(place: Placement, natural: { width: number; height: number }, body: string, attrs = ''): string {
  if ('range' in place) {
    const [from, to = from] = place.range.split(':');
    const a = cellPos(from), b = cellPos(to);
    // The object fills the range, given either way round: it ends at the far edge of the last cell
    const start = { col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) };
    const end = { col: Math.max(a.col, b.col) + 1, row: Math.max(a.row, b.row) + 1 };
    return `<xdr:twoCellAnchor editAs="oneCell"${attrs}>${marker('from', start)}${marker('to', end)}${body}<xdr:clientData/></xdr:twoCellAnchor>`;
  }
  const { width: w, height: h } = natural;
  const width = place.width ?? (place.height ? w * place.height / h : w);
  const height = place.height ?? h * width / w;
  return `<xdr:oneCellAnchor${attrs}>${marker('from', cellPos(place.at))}<xdr:ext cx="${Math.round(width * EMU_PER_PX)}" cy="${Math.round(height * EMU_PER_PX)}"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
}

export const drawingPartXml = (anchors: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr ${DRAWING_NS}>${anchors}</xdr:wsDr>`;

// One drawing part per sheet; `rIds[i]` is the drawing relationship of images[i]'s media
export function drawingXml(images: SheetImage[], infos: ImageInfo[], rIds: string[]): string {
  return drawingPartXml(images.map((img, i) => anchorXml(img, infos[i],
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="Picture ${i + 1}"${img.altText ? ` descr="${escapeXml(img.altText)}"` : ''}/>` +
    `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${rIds[i]}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`)).join(''));
}
