import type { CellFont } from './types';

// Excel's default indexed palette (ECMA-376 Part 1, 18.8.27); 64/65 are the system text/background colours
const DEFAULT_INDEXED = (
  '000000FFFFFFFF000000FF000000FFFFFF00FF00FF00FFFF' +
  '000000FFFFFFFF000000FF000000FFFFFF00FF00FF00FFFF' +
  '800000008000000080808000800080008080C0C0C0808080' +
  '9999FF993366FFFFCCCCFFFF660066FF80800066CCCCCCFF' +
  '000080FF00FFFFFF0000FFFF8000808000000080800000FF' +
  '00CCFFCCFFFFCCFFCCFFFF9999CCFFFF99CCCC99FFFFCC99' +
  '3366FF33CCCC99CC00FFCC00FF9900FF6600666699969696' +
  '003366339966003300333300993300993366333399333333' +
  '000000FFFFFF'
).match(/.{6}/g)!.map(hex => 'FF' + hex);

// Theme colour slots in the order Excel indexes them: lt1/dk1 and lt2/dk2 are swapped relative to the XML
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

export function parseThemeColors(themeXml: string): string[] {
  const scheme = /<(?:\w+:)?clrScheme\b[\s\S]*?<\/(?:\w+:)?clrScheme>/.exec(themeXml)?.[0] ?? '';
  return THEME_SLOTS.map(slot => {
    const body = new RegExp(`<(?:\\w+:)?${slot}>([\\s\\S]*?)</(?:\\w+:)?${slot}>`).exec(scheme)?.[1] ?? '';
    const hex = /srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(body)?.[1] ?? /lastClr="([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    return hex ? 'FF' + hex.toUpperCase() : '';
  });
}

// Excel's tint: scale HSL lightness towards black (tint < 0) or white (tint > 0).
// Float HSL lands within 2 per channel of Excel's rendering; Excel's exact rounding is undocumented.
function applyTint(argb: string, tint: number): string {
  const [r, g, b] = [2, 4, 6].map(i => parseInt(argb.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    t = (t + 1) % 1;
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const rgb = s === 0 ? [l, l, l] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  return argb.slice(0, 2) + rgb.map(c => Math.round(c * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

export type ColorResolver = (attrs: Record<string, string>) => string | undefined;

// Turns a CT_Color (rgb / theme / indexed, optional tint) into ARGB hex; "auto" and unknown slots give undefined
export function colorResolver(theme: string[], indexed: string[] = DEFAULT_INDEXED): ColorResolver {
  return a => {
    let argb: string | undefined;
    if (a['rgb']) argb = a['rgb'].length === 6 ? 'FF' + a['rgb'] : a['rgb'];
    else if (a['theme'] !== undefined) argb = theme[parseInt(a['theme'], 10)] || undefined;
    else if (a['indexed'] !== undefined) argb = (indexed[parseInt(a['indexed'], 10)] ?? DEFAULT_INDEXED[parseInt(a['indexed'], 10)]) || undefined;
    if (!argb || !/^[0-9A-Fa-f]{8}$/.test(argb)) return undefined;
    const tint = Number(a['tint'] || 0);
    return tint ? applyTint(argb.toUpperCase(), Math.max(-1, Math.min(1, tint))) : argb.toUpperCase();
  };
}

// One child element of <font> (styles) or <rPr> (rich text runs). Returns false for unrelated elements.
export function applyFontElement(font: CellFont, name: string, a: Record<string, string>, color: (a: Record<string, string>) => void): boolean {
  const on = a['val'] !== '0' && a['val'] !== 'false';
  switch (name) {
    case 'b': font.bold = on; return true;
    case 'i': font.italic = on; return true;
    case 'u': font.underline = a['val'] !== 'none'; return true;
    case 'sz': font.size = Number(a['val']); return true;
    case 'name': case 'rFont': font.name = a['val']; return true;
    case 'color': color(a); return true;
  }
  return false;
}
