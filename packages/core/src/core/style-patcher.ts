import { CellStyle, CellFont } from './types';
import { fontXml, fillXml, borderSideXml, alignmentAttrs } from './style-engine';
import { attr } from './utils';

const escapeAttr = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

const FONT_TAGS: [keyof CellFont, string][] = [
  ['bold', 'b'], ['italic', 'i'], ['underline', 'u'], ['size', 'sz'], ['color', 'color'], ['name', 'name'],
];
const SIDES = ['left', 'right', 'top', 'bottom'] as const;
// CT_Stylesheet child order, to place a section the file does not have yet
const SECTIONS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

function setAttr(open: string, name: string, value: string): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  return re.test(open) ? open.replace(re, ` ${name}="${value}"`) : open.replace(/\s*(\/?)>$/, ` ${name}="${value}"$1>`);
}

// Adds cell formats to an existing workbook's styles part. A patch changes only what it names: font
// properties, border sides and alignment settings are merged into the cell's current format, while
// a fill or number format replaces the old one.
export class StylePatcher {
  private p: string;
  private fonts: string[];
  private fills: number;
  private borders: string[];
  private xfs: string[];
  private numFmts = new Map<string, number>();
  private nextNumFmt = 164;
  private added: Record<'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs', string[]> = { numFmts: [], fonts: [], fills: [], borders: [], cellXfs: [] };
  private memo = new Map<string, number>();

  constructor(private xml: string) {
    this.p = /<((?:\w+:)?)styleSheet\b/.exec(xml)?.[1] ?? '';
    this.fonts = this.items('fonts', 'font');
    this.fills = this.items('fills', 'fill').length;
    this.borders = this.items('borders', 'border');
    this.xfs = this.items('cellXfs', 'xf');
    for (const tag of this.items('numFmts', 'numFmt')) {
      const id = parseInt(attr(tag, 'numFmtId') ?? '', 10);
      const code = attr(tag, 'formatCode');
      if (code !== null && !isNaN(id)) this.numFmts.set(code, id);
      if (id >= this.nextNumFmt) this.nextNumFmt = id + 1;
    }
  }

  get changed(): boolean {
    return this.added.cellXfs.length > 0;
  }

  // The cellXfs index of format `base` with `style` applied
  patch(base: number, style: CellStyle): number {
    const key = `${base}|${JSON.stringify(style)}`;
    let id = this.memo.get(key);
    if (id !== undefined) return id;

    const p = this.p;
    const xf = this.xfs[base] ?? this.xfs[0] ?? `<${p}xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
    let open = /^<[^>]*>/.exec(xf)![0];
    let inner = open.endsWith('/>') ? '' : xf.slice(open.length, xf.lastIndexOf('<'));
    open = open.replace(/\s*\/>$/, '>');

    if (style.font) {
      const old = this.fonts[parseInt(attr(open, 'fontId') ?? '0', 10)] ?? '';
      let font = old.endsWith('/>') ? '' : old.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '');
      for (const [key, tag] of FONT_TAGS) {
        if (!(key in style.font)) continue;
        font = font.replace(new RegExp(`<${p}${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}${tag}>)`, 'g'), '');
        if (style.font[key] !== undefined && style.font[key] !== false) font += this.prefix(fontXml({ [key]: style.font[key] }));
      }
      open = setAttr(setAttr(open, 'fontId', String(this.add('fonts', `<${p}font>${font}</${p}font>`))), 'applyFont', '1');
    }
    if (style.fill) {
      open = setAttr(setAttr(open, 'fillId', String(this.add('fills', this.prefix(fillXml(style.fill))))), 'applyFill', '1');
    }
    if (style.border) {
      const old = this.borders[parseInt(attr(open, 'borderId') ?? '0', 10)] ?? `<${p}border/>`;
      const element = (tag: string) => new RegExp(`<${p}${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}${tag}>)`).exec(old)?.[0];
      const sides = SIDES.map(side => side in style.border!
        ? this.prefix(borderSideXml(side, style.border![side])) : element(side) ?? `<${p}${side}/>`);
      const border = `${/^<[^>]*?(?=\/?>)/.exec(old)![0]}>${sides.join('')}${element('diagonal') ?? `<${p}diagonal/>`}</${p}border>`;
      open = setAttr(setAttr(open, 'borderId', String(this.add('borders', border))), 'applyBorder', '1');
    }
    if (style.numFmt !== undefined) {
      let fmt = this.numFmts.get(style.numFmt);
      if (fmt === undefined) {
        this.numFmts.set(style.numFmt, fmt = this.nextNumFmt++);
        this.added.numFmts.push(`<${p}numFmt numFmtId="${fmt}" formatCode="${escapeAttr(style.numFmt)}"/>`);
      }
      open = setAttr(setAttr(open, 'numFmtId', String(fmt)), 'applyNumberFormat', '1');
    }
    if (style.alignment) {
      const re = new RegExp(`<${p}alignment\\b[^>]*?/>|<${p}alignment\\b[^>]*>[\\s\\S]*?</${p}alignment>`);
      const old = re.exec(inner)?.[0] ?? '';
      const attrs: Record<string, string> = {};
      for (const m of old.matchAll(/\s([\w:]+)="([^"]*)"/g)) attrs[m[1]] = m[2];
      Object.assign(attrs, alignmentAttrs(style.alignment));
      if ('wrapText' in style.alignment && !style.alignment.wrapText) delete attrs.wrapText;
      const element = `<${p}alignment${Object.entries(attrs).map(([k, v]) => ` ${k}="${v}"`).join('')}/>`;
      inner = old ? inner.replace(old, element) : element + inner; // alignment is the first child
      open = setAttr(open, 'applyAlignment', '1');
    }

    id = this.add('cellXfs', inner ? `${open}${inner}</${p}xf>` : open.replace(/>$/, '/>'));
    this.memo.set(key, id);
    return id;
  }

  toXml(): string {
    let xml = this.xml;
    for (const section of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs'] as const) {
      const items = this.added[section];
      if (!items.length) continue;
      const p = this.p;
      const m = new RegExp(`<${p}${section}\\b([^>]*?)(/?)>`).exec(xml);
      if (m) {
        const count = this.existing(section) + items.length;
        const open = setAttr(`<${p}${section}${m[1]}>`, 'count', String(count));
        if (m[2]) {
          xml = xml.replace(m[0], `${open}${items.join('')}</${p}${section}>`);
        } else {
          const close = xml.indexOf(`</${p}${section}>`, m.index);
          xml = xml.slice(0, m.index) + open + xml.slice(m.index + m[0].length, close) + items.join('') + xml.slice(close);
        }
      } else {
        // Before the first later section, or the end of the stylesheet
        const later = SECTIONS.slice(SECTIONS.indexOf(section) + 1).map(s => `<${p}${s}\\b`).join('|');
        const at = xml.search(new RegExp(`${later}|</${p}styleSheet>`));
        xml = xml.slice(0, at) + `<${p}${section} count="${items.length}">${items.join('')}</${p}${section}>` + xml.slice(at);
      }
    }
    return xml;
  }

  private existing(section: 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs'): number {
    return section === 'fonts' ? this.fonts.length : section === 'fills' ? this.fills
      : section === 'borders' ? this.borders.length : section === 'cellXfs' ? this.xfs.length
        : this.items('numFmts', 'numFmt').length;
  }

  // Appends an element to a section, returning its index
  private add(section: 'fonts' | 'fills' | 'borders' | 'cellXfs', element: string): number {
    this.added[section].push(element);
    return this.existing(section) + this.added[section].length - 1;
  }

  private items(section: string, tag: string): string[] {
    const p = this.p;
    const m = new RegExp(`<${p}${section}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${p}${section}>)`).exec(this.xml);
    return [...(m?.[1] ?? '').matchAll(new RegExp(`<${p}${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}${tag}>)`, 'g'))].map(x => x[0]);
  }

  // Our builders write unprefixed elements; a prefixed stylesheet needs its prefix on them
  private prefix(xml: string): string {
    return this.p ? xml.replace(/<(\/?)([A-Za-z])/g, `<$1${this.p}$2`) : xml;
  }
}
