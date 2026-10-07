import { CellStyle, CellFont, CellFill, GradientFill, CellBorder, CellAlignment, HighlightStyle } from './types';

function escapeXml(val: string): string {
  return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Child elements of a <font> (styles) or, with nameTag "rFont", of a rich text run's <rPr>
export function fontXml(font: CellFont, nameTag = 'name'): string {
  let xml = '';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.underline) xml += '<u/>';
  if (font.size) xml += `<sz val="${font.size}"/>`;
  if (font.color) xml += `<color rgb="${escapeXml(font.color)}"/>`;
  if (font.name) xml += `<${nameTag} val="${escapeXml(font.name)}"/>`;
  return xml;
}

export const borderSideXml = (tag: string, s?: { style: string; color?: string }) =>
  s ? `<${tag} style="${s.style}">${s.color ? `<color rgb="${escapeXml(s.color)}"/>` : ''}</${tag}>` : `<${tag}/>`;

export function borderXml(b: CellBorder): string {
  const side = borderSideXml;
  return `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`;
}

export function fillXml(fill: CellFill | GradientFill): string {
  if (fill.type === 'gradient') {
    const stops = fill.stops.map(s => `<stop position="${s.position}"><color rgb="${escapeXml(s.color)}"/></stop>`).join('');
    return `<fill><gradientFill degree="${fill.degree ?? 0}">${stops}</gradientFill></fill>`;
  }
  return `<fill><patternFill patternType="solid"><fgColor rgb="${escapeXml(fill.fgColor)}"/></patternFill></fill>`;
}

// Attributes of an <alignment> element
export function alignmentAttrs(alignment: CellAlignment): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (alignment.horizontal) attrs.horizontal = alignment.horizontal;
  // OOXML has no "middle"; Excel repairs the file if it sees one
  if (alignment.vertical) attrs.vertical = alignment.vertical === 'middle' ? 'center' : alignment.vertical;
  if (alignment.wrapText) attrs.wrapText = '1';
  return attrs;
}

// Registry of unique styles: maps to integer index for OOXML styleSheet
interface StyleRegistry {
  fonts: Map<string, number>;
  fills: Map<string, number>;
  borders: Map<string, number>;
  numFmts: Map<string, number>;
  cellXfs: { fontId: number; fillId: number; borderId: number; numFmtId: number; alignmentXml: string }[];
}

export class StyleEngine {
  private registry: StyleRegistry = {
    fonts: new Map(),
    fills: new Map(),
    borders: new Map(),
    numFmts: new Map(),
    cellXfs: [],
  };
  private xfIndex = new Map<string, number>();
  // Differential formats used by conditional formatting, in dxfId order
  private dxfs = new Map<string, number>();

  // dxfId for a highlight style
  registerDxf(style: HighlightStyle): number {
    const key = JSON.stringify(style);
    let id = this.dxfs.get(key);
    if (id === undefined) {
      this.dxfs.set(key, id = this.dxfs.size);
      if (style.numFmt) this.registerNumFmt(style.numFmt);
    }
    return id;
  }

  private dxfXml(style: HighlightStyle): string {
    let xml = style.font ? `<font>${fontXml(style.font)}</font>` : '';
    if (style.numFmt) xml += `<numFmt numFmtId="${this.registerNumFmt(style.numFmt)}" formatCode="${escapeXml(style.numFmt)}"/>`;
    if (style.fill?.type === 'solid') {
      // A differential solid fill takes its colour from bgColor
      const c = escapeXml(style.fill.fgColor);
      xml += `<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor rgb="${c}"/></patternFill></fill>`;
    }
    if (style.border) xml += borderXml(style.border);
    return `<dxf>${xml}</dxf>`;
  }

  registerStyle(style: CellStyle): number {
    const fontId = style.font ? this.registerFont(style.font) : 0;
    const fillId = style.fill ? this.registerFill(style.fill) : 0;
    const borderId = style.border ? this.registerBorder(style.border) : 0;
    const numFmtId = style.numFmt ? this.registerNumFmt(style.numFmt) : 0;
    const alignmentXml = style.alignment ? this.buildAlignmentXml(style.alignment) : '';

    const key = `${fontId}|${fillId}|${borderId}|${numFmtId}|${alignmentXml}`;
    const existing = this.xfIndex.get(key);
    if (existing !== undefined) return existing;

    this.registry.cellXfs.push({ fontId, fillId, borderId, numFmtId, alignmentXml });
    const id = this.registry.cellXfs.length; // 1-based because 0 is the default xf
    this.xfIndex.set(key, id);
    return id;
  }

  private registerFont(font: CellFont): number {
    const key = JSON.stringify(font);
    if (this.registry.fonts.has(key)) return this.registry.fonts.get(key)!;
    const id = this.registry.fonts.size + 1; // slot 0 = default font
    this.registry.fonts.set(key, id);
    return id;
  }

  private registerFill(fill: CellFill | GradientFill): number {
    const key = JSON.stringify(fill);
    if (this.registry.fills.has(key)) return this.registry.fills.get(key)!;
    const id = this.registry.fills.size + 2; // slots 0-1 = required none/gray125
    this.registry.fills.set(key, id);
    return id;
  }

  private registerBorder(border: CellBorder): number {
    const key = JSON.stringify(border);
    if (this.registry.borders.has(key)) return this.registry.borders.get(key)!;
    const id = this.registry.borders.size + 1; // slot 0 = empty border
    this.registry.borders.set(key, id);
    return id;
  }

  private registerNumFmt(fmt: string): number {
    if (this.registry.numFmts.has(fmt)) return this.registry.numFmts.get(fmt)!;
    const id = 164 + this.registry.numFmts.size;
    this.registry.numFmts.set(fmt, id);
    return id;
  }

  private buildAlignmentXml(alignment: CellAlignment): string {
    const attrs = Object.entries(alignmentAttrs(alignment)).map(([k, v]) => `${k}="${v}"`);
    return attrs.length ? `<alignment ${attrs.join(' ')}/>` : '';
  }

  // Serialize entire style registry to OOXML styleSheet XML
  toXml(): string {
    const numFmtsXml = this.buildNumFmtsXml();
    const fontsXml = this.buildFontsXml();
    const fillsXml = this.buildFillsXml();
    const bordersXml = this.buildBordersXml();
    const cellXfsXml = this.buildCellXfsXml();

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${numFmtsXml}
${fontsXml}
${fillsXml}
${bordersXml}
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${this.registry.cellXfs.length + 1}">
${cellXfsXml}
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
${this.dxfs.size ? `<dxfs count="${this.dxfs.size}">${[...this.dxfs.keys()].map(k => this.dxfXml(JSON.parse(k))).join('')}</dxfs>` : ''}
</styleSheet>`;
  }

  private buildNumFmtsXml(): string {
    if (this.registry.numFmts.size === 0) return '';
    const entries: string[] = [];
    for (const [fmt, id] of this.registry.numFmts) {
      entries.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(fmt)}"/>`);
    }
    return `<numFmts count="${entries.length}">${entries.join('')}</numFmts>`;
  }

  private buildFontsXml(): string {
    const fontEntries: string[] = [
      '<font><sz val="11"/><name val="Calibri"/></font>' // default
    ];

    for (const [key] of this.registry.fonts) {
      fontEntries.push(`<font>${fontXml(JSON.parse(key))}</font>`);
    }

    return `<fonts count="${fontEntries.length}">${fontEntries.join('')}</fonts>`;
  }

  private buildFillsXml(): string {
    const fillEntries: string[] = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
    ];

    for (const [key] of this.registry.fills) fillEntries.push(fillXml(JSON.parse(key)));

    return `<fills count="${fillEntries.length}">${fillEntries.join('')}</fills>`;
  }

  private buildBordersXml(): string {
    const borderEntries: string[] = [
      '<border><left/><right/><top/><bottom/><diagonal/></border>'
    ];

    for (const [key] of this.registry.borders) borderEntries.push(borderXml(JSON.parse(key)));

    return `<borders count="${borderEntries.length}">${borderEntries.join('')}</borders>`;
  }

  private buildCellXfsXml(): string {
    // Default xf
    const entries: string[] = [
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    ];

    for (const xf of this.registry.cellXfs) {
      entries.push(
        `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${xf.alignmentXml ? ' applyAlignment="1"' : ''}>${xf.alignmentXml}</xf>`
      );
    }

    return entries.join('\n');
  }
}
