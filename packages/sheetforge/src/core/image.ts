import type { SheetImage } from './types';
import { colIndex } from './utils';

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

const escapeXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const cellPos = (ref: string) => {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!m) throw new Error(`Invalid image cell "${ref}".`);
  return { col: colIndex(m[1]), row: parseInt(m[2], 10) - 1 };
};
const marker = (tag: string, { col, row }: { col: number; row: number }) =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

// One drawing part per sheet; `rIds[i]` is the drawing relationship of images[i]'s media
export function drawingXml(images: SheetImage[], infos: ImageInfo[], rIds: string[]): string {
  const anchors = images.map((img, i) => {
    const pic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 2}" name="Picture ${i + 1}"${img.altText ? ` descr="${escapeXml(img.altText)}"` : ''}/>` +
      `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip r:embed="${rIds[i]}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>`;
    if ('range' in img) {
      const [from, to = from] = img.range.split(':');
      const end = cellPos(to);
      // The picture fills the range: it ends at the far edge of the last cell
      return `<xdr:twoCellAnchor editAs="oneCell">${marker('from', cellPos(from))}${marker('to', { col: end.col + 1, row: end.row + 1 })}${pic}</xdr:twoCellAnchor>`;
    }
    // A missing side keeps the aspect ratio; neither given means the image's own size
    const { width: w, height: h } = infos[i];
    const width = img.width ?? (img.height ? w * img.height / h : w);
    const height = img.height ?? h * width / w;
    return `<xdr:oneCellAnchor>${marker('from', cellPos(img.at))}<xdr:ext cx="${Math.round(width * EMU_PER_PX)}" cy="${Math.round(height * EMU_PER_PX)}"/>${pic}</xdr:oneCellAnchor>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`;
}
