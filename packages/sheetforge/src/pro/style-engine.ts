import { CellStyle, CellFont, CellFill, GradientFill, CellBorder, CellAlignment } from './types';

// Registry of unique styles — maps to integer index for OOXML styleSheet
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

  registerStyle(style: CellStyle): number {
    const fontId = style.font ? this.registerFont(style.font) : 0;
    const fillId = style.fill ? this.registerFill(style.fill) : 0;
    const borderId = style.border ? this.registerBorder(style.border) : 0;
    const numFmtId = style.numFmt ? this.registerNumFmt(style.numFmt) : 0;
    const alignmentXml = style.alignment ? this.buildAlignmentXml(style.alignment) : '';

    const existingIdx = this.registry.cellXfs.findIndex(xf => 
      xf.fontId === fontId && 
      xf.fillId === fillId && 
      xf.borderId === borderId && 
      xf.numFmtId === numFmtId && 
      xf.alignmentXml === alignmentXml
    );

    if (existingIdx !== -1) {
      return existingIdx + 1;
    }

    const xf = { fontId, fillId, borderId, numFmtId, alignmentXml };
    this.registry.cellXfs.push(xf);
    return this.registry.cellXfs.length; // 1-based because 0 is the default xf
  }

  private registerFont(font: CellFont): number {
    const key = JSON.stringify(font);
    if (this.registry.fonts.has(key)) return this.registry.fonts.get(key)!;
    const id = this.registry.fonts.size;
    this.registry.fonts.set(key, id);
    return id;
  }

  private registerFill(fill: CellFill | GradientFill): number {
    const key = JSON.stringify(fill);
    if (this.registry.fills.has(key)) return this.registry.fills.get(key)!;
    const id = this.registry.fills.size;
    this.registry.fills.set(key, id);
    return id;
  }

  private registerBorder(border: CellBorder): number {
    const key = JSON.stringify(border);
    if (this.registry.borders.has(key)) return this.registry.borders.get(key)!;
    const id = this.registry.borders.size;
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
    const attrs: string[] = [];
    if (alignment.horizontal) attrs.push(`horizontal="${alignment.horizontal}"`);
    if (alignment.vertical) attrs.push(`vertical="${alignment.vertical}"`);
    if (alignment.wrapText) attrs.push(`wrapText="1"`);
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
<cellXfs count="${this.registry.cellXfs.length + 1}">
${cellXfsXml}
</cellXfs>
</styleSheet>`;
  }

  private buildNumFmtsXml(): string {
    if (this.registry.numFmts.size === 0) return '';
    const entries: string[] = [];
    for (const [fmt, id] of this.registry.numFmts) {
      entries.push(`<numFmt numFmtId="${id}" formatCode="${fmt}"/>`);
    }
    return `<numFmts count="${entries.length}">${entries.join('')}</numFmts>`;
  }

  private buildFontsXml(): string {
    const fontEntries: string[] = [
      '<font><sz val="11"/><name val="Calibri"/></font>' // default
    ];

    for (const [key] of this.registry.fonts) {
      const font: CellFont = JSON.parse(key);
      const parts: string[] = [];
      if (font.bold) parts.push('<b/>');
      if (font.italic) parts.push('<i/>');
      if (font.underline) parts.push('<u/>');
      if (font.size) parts.push(`<sz val="${font.size}"/>`);
      if (font.color) parts.push(`<color rgb="${font.color}"/>`);
      if (font.name) parts.push(`<name val="${font.name}"/>`);
      fontEntries.push(`<font>${parts.join('')}</font>`);
    }

    return `<fonts count="${fontEntries.length}">${fontEntries.join('')}</fonts>`;
  }

  private buildFillsXml(): string {
    const fillEntries: string[] = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
    ];

    for (const [key] of this.registry.fills) {
      const fill: CellFill | GradientFill = JSON.parse(key);
      if (fill.type === 'solid') {
        fillEntries.push(
          `<fill><patternFill patternType="solid"><fgColor rgb="${fill.fgColor}"/></patternFill></fill>`
        );
      } else if (fill.type === 'gradient') {
        const gf = fill as GradientFill;
        const stops = gf.stops.map(s =>
          `<gradientStop position="${s.position}"><color rgb="${s.color}"/></gradientStop>`
        ).join('');
        fillEntries.push(`<fill><gradientFill degree="${gf.degree ?? 0}">${stops}</gradientFill></fill>`);
      }
    }

    return `<fills count="${fillEntries.length}">${fillEntries.join('')}</fills>`;
  }

  private buildBordersXml(): string {
    const buildSide = (side?: { style: string; color?: string }) =>
      side
        ? `<border style="${side.style}">${side.color ? `<color rgb="${side.color}"/>` : ''}</border>`
        : '<border/>';

    const borderEntries: string[] = [
      '<border><left/><right/><top/><bottom/><diagonal/></border>'
    ];

    for (const [key] of this.registry.borders) {
      const b: CellBorder = JSON.parse(key);
      borderEntries.push(
        `<border><left${b.left ? ` style="${b.left.style}"` : ''}/><right${b.right ? ` style="${b.right.style}"` : ''}/><top${b.top ? ` style="${b.top.style}"` : ''}/><bottom${b.bottom ? ` style="${b.bottom.style}"` : ''}/></border>`
      );
    }

    return `<borders count="${borderEntries.length}">${borderEntries.join('')}</borders>`;
  }

  private buildCellXfsXml(): string {
    // Default xf
    const entries: string[] = [
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    ];

    for (const xf of this.registry.cellXfs) {
      entries.push(
        `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1">${xf.alignmentXml}</xf>`
      );
    }

    return entries.join('\n');
  }
}
