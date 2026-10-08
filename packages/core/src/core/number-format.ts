// Renders cell values the way Excel (en-US) displays them for a number format code.

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// t: lit, num (0 # ? . , % / E+ E-), date (y m d h s, and n for minutes), elapsed ([h] [m] [s]), ampm, text (@), general
interface Tok { t: string; v: string; n?: number }
interface Section { toks: Tok[]; cond?: [string, number]; date: boolean; text: boolean }

const cache = new Map<string, Section[]>();

export function formatValue(value: number | string | boolean | Date | null | undefined, code: string, date1904 = false): string {
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) value = toSerial(value, date1904);
  let secs = cache.get(code);
  if (!secs) {
    if (cache.size >= 1000) cache.clear(); // codes come from files, so keep the cache bounded
    cache.set(code, secs = splitSections(code.trim() ? code : 'General').map(parseSection));
  }
  const text = secs.length === 4 ? secs[3] : secs[secs.length - 1].text ? secs[secs.length - 1] : undefined;
  if (typeof value === 'string') return text ? text.toks.map(t => t.t === 'text' ? value : t.v).join('') : value;
  if (!Number.isFinite(value)) return String(value);
  const [sec, signed] = pick(secs.slice(0, text ? secs.length - 1 : 3), value);
  if (!sec) return general(value);
  const neg = value < 0 && signed, x = Math.abs(value);
  if (sec.date) return neg ? general(value) : fmtDate(sec.toks, x, date1904);
  return (neg ? '-' : '') + fmtNumber(sec.toks, x);
}

function toSerial(d: Date, date1904: boolean): number {
  const days = d.getTime() / 864e5;
  if (date1904) return days + 24107;
  const s = days + 25569;
  return s < 61 ? s - 1 : s;  // Excel's phantom 1900-02-29 is serial 60
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '', quoted = false, bracket = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === ';' && !quoted && !bracket) { out.push(cur); cur = ''; continue; }
    cur += c;
    if (quoted) quoted = c !== '"';
    else if (c === '"') quoted = true;
    else if ((c === '\\' || c === '_' || c === '*') && i + 1 < code.length) cur += code[++i];
    else if (c === '[') bracket = true;
    else if (c === ']') bracket = false;
  }
  out.push(cur);
  return out.slice(0, 4);
}

function parseSection(s: string): Section {
  const toks: Tok[] = [];
  let cond: [string, number] | undefined;
  const lit = (v: string) => toks.push({ t: 'lit', v });
  // fractional seconds like ss.00 attach to the seconds token
  const subsec = (i: number, tok: Tok) => {
    const m = /^\.0+/.exec(s.slice(i + 1));
    if (m) tok.n = m[0].length - 1;
    toks.push(tok);
    return m ? i + m[0].length : i;
  };
  for (let i = 0; i < s.length; i++) {
    const c = s[i], lc = c.toLowerCase();
    if (c === '"') {
      const j = s.indexOf('"', i + 1), end = j < 0 ? s.length : j;
      lit(s.slice(i + 1, end));
      i = end;
    } else if (c === '\\') lit(s[++i] ?? '');
    else if (c === '_') { i++; lit(' '); }
    else if (c === '*') i++;
    else if (c === '[') {
      const j = s.indexOf(']', i), end = j < 0 ? s.length : j, b = s.slice(i + 1, end);
      i = end;
      const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)$/i.exec(b);
      if (m) cond = [m[1], +m[2]];
      else if (/^(h+|m+|s+)$/i.test(b)) i = subsec(i, { t: 'elapsed', v: b.toLowerCase() });
      else if (b[0] === '$') lit(b.slice(1).split('-')[0]);
      // colors and other tags are ignored
    } else if (s.slice(i, i + 7).toLowerCase() === 'general') { toks.push({ t: 'general', v: '' }); i += 6; }
    else if (s.slice(i, i + 5).toUpperCase() === 'AM/PM') { toks.push({ t: 'ampm', v: 'AM/PM' }); i += 4; }
    else if (/^a\/p$/i.test(s.slice(i, i + 3))) { toks.push({ t: 'ampm', v: s.slice(i, i + 3) }); i += 2; }
    else if ('ymdhs'.includes(lc)) {
      let j = i;
      while (s[j + 1]?.toLowerCase() === lc) j++;
      i = subsec(j, { t: 'date', v: lc.repeat(j - i + 1) });
    } else if (lc === 'e' && (s[i + 1] === '+' || s[i + 1] === '-')) toks.push({ t: 'num', v: 'E' + s[++i] });
    else if ('0#?.,%/'.includes(c)) toks.push({ t: 'num', v: c });
    else if (c === '@') toks.push({ t: 'text', v: '' });
    else lit(c);
  }
  // m means minutes right after an hour or right before a second
  const dt = toks.filter(t => t.t === 'date' || t.t === 'elapsed');
  dt.forEach((t, k) => {
    if (t.t === 'date' && t.v[0] === 'm' && (dt[k - 1]?.v[0] === 'h' || dt[k + 1]?.v[0] === 's')) t.v = 'n'.repeat(t.v.length);
  });
  return { toks, cond, date: dt.length > 0 || toks.some(t => t.t === 'ampm'), text: toks.some(t => t.t === 'text') };
}

function test([op, n]: [string, number], v: number): boolean {
  return op === '<' ? v < n : op === '<=' ? v <= n : op === '>' ? v > n : op === '>=' ? v >= n : op === '<>' ? v !== n : v === n;
}

// Returns the section for v and whether a minus sign should be shown
function pick([a, b, c]: Section[], v: number): [Section | undefined, boolean] {
  if (!a) return [undefined, true];
  if (a.cond || b?.cond) {
    const s = a.cond && test(a.cond, v) ? a : b?.cond && test(b.cond, v) ? b : a.cond && b?.cond ? c : b;
    // a section meant for negatives writes its own sign
    return [s, !(s?.cond && s.cond[0][0] === '<' && s.cond[1] <= 0)];
  }
  if (v < 0 && b) return [b, false];
  if (v === 0 && c) return [c, true];
  return [a, true];
}

// |x| as 15 significant digits and a decimal exponent, the precision Excel works in
function sig15(x: number): [string, number] {
  const [m, e] = Math.abs(x).toExponential(14).split('e');
  return [m.replace('.', ''), +e];
}

// |x| rounded half away from zero to n decimals, as [integer digits without leading zeros, n decimals]
function fixed(x: number, n: number): [string, string] {
  let [ds, e] = sig15(x), p = e + 1;
  if (p < 0) { ds = '0'.repeat(-p) + ds; p = 0; }
  ds = ds.padEnd(p + n + 1, '0');
  let keep = ds.slice(0, p + n);
  if (ds[p + n] >= '5') keep = (BigInt(keep || '0') + BigInt(1)).toString().padStart(keep.length, '0');
  return [keep.slice(0, keep.length - n).replace(/^0+/, ''), n ? keep.slice(-n) : ''];
}

function general(v: number): string {
  if (!v) return '0';
  const sign = v < 0 ? '-' : '';
  let [ds, e] = sig15(v);
  // fits in 11 characters as a plain number
  if (e >= -4 && e <= 10) {
    const [i, f] = fixed(v, Math.max(0, 9 - Math.max(e, 0)));
    const t = f.replace(/0+$/, '');
    if (i.length <= 11) return sign + (i || '0') + (t ? '.' + t : '');
  }
  let m = ds.slice(0, 6);
  if (ds[6] >= '5') m = String(+m + 1);
  if (m.length > 6) { m = m.slice(0, 6); e++; }
  return sign + m[0] + ('.' + m.slice(1)).replace(/\.?0+$/, '') + 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(2, '0');
}

const isPh = (t: Tok | undefined) => t?.t === 'num' && '0#?'.includes(t.v);

// Right-aligns digits into placeholders; the first placeholder takes any overflow
function fillInt(toks: Tok[], idx: number[], digits: string, out: string[], group = false) {
  let j = digits.length;
  for (let k = idx.length - 1; k >= 0; k--) {
    const p = toks[idx[k]].v;
    if (j > 0) { out[idx[k]] = k ? digits[--j] : digits.slice(0, j); if (!k) j = 0; }
    else out[idx[k]] = p === '0' ? '0' : p === '?' ? ' ' : '';
  }
  if (!group) return;
  let n = 0;
  for (let k = idx.length - 1; k >= 0; k--) {
    let s = '';
    for (const c of [...out[idx[k]]].reverse()) {
      if (c >= '0' && c <= '9') { if (n && n % 3 === 0) s = ',' + s; n++; }
      s = c + s;
    }
    out[idx[k]] = s;
  }
}

// Trailing zeros are dropped for # and blanked for ?
function fillDec(toks: Tok[], idx: number[], frac: string, out: string[]) {
  let trailing = true;
  for (let k = idx.length - 1; k >= 0; k--) {
    const p = toks[idx[k]].v;
    if (trailing && frac[k] === '0' && p !== '0') out[idx[k]] = p === '?' ? ' ' : '';
    else { trailing = false; out[idx[k]] = frac[k]; }
  }
}

function fmtNumber(toks: Tok[], x: number): string {
  const out = toks.map(t => t.t === 'general' ? general(x) : t.t === 'num' && (t.v === ',' || isPh(t)) ? '' : t.v);
  const idx = toks.flatMap((t, i) => isPh(t) ? [i] : []);
  if (!idx.length) return out.join('');
  x *= 100 ** toks.filter(t => t.t === 'num' && t.v === '%').length;

  const slash = toks.findIndex(t => t.t === 'num' && t.v === '/');
  if (slash > 0 && isPh(toks[slash - 1])) return fraction(toks, idx, slash, x, out);

  const e = toks.findIndex(t => t.t === 'num' && t.v[0] === 'E');
  const end = e >= 0 ? e : toks.length;
  const dot = toks.findIndex((t, i) => t.t === 'num' && t.v === '.' && i < end);
  const intEnd = dot >= 0 ? dot : end;
  const intIdx = idx.filter(i => i < intEnd), decIdx = idx.filter(i => i > intEnd && i < end);

  if (e >= 0) {
    const ip = Math.max(intIdx.length, 1);
    // with # in the integer part the exponent is a multiple of the integer digit count (engineering)
    const eng = ip > 1 && intIdx.some(i => toks[i].v === '#');
    const expOf = (p: number) => eng ? Math.floor(p / ip) * ip : p - ip + 1;
    let [ds, p] = x ? sig15(x) : ['0', 0];
    let exp = x ? expOf(p) : 0;
    let [mi, mf] = fixed(Number(`${ds[0]}.${ds.slice(1)}e${p - exp}`), decIdx.length);
    if (x && mi.length > p - exp + 1) { exp = expOf(++p); [mi, mf] = fixed(10 ** (p - exp), decIdx.length); }
    fillInt(toks, intIdx, mi, out);
    fillDec(toks, decIdx, mf, out);
    out[e] = 'E' + (exp < 0 ? '-' : toks[e].v === 'E+' ? '+' : '');
    fillInt(toks, idx.filter(i => i > e), String(Math.abs(exp)), out);
    return out.join('');
  }

  // a comma with no digit placeholder after it scales by 1000, one between digits groups thousands
  let scale = 0, group = false;
  toks.forEach((t, i) => {
    if (t.t !== 'num' || t.v !== ',') return;
    const bound = i < intEnd ? intEnd : end;
    if (!idx.some(j => j > i && j < bound)) scale++;
    else if (i < intEnd && idx.some(j => j < i)) group = true;
  });
  const [int, frac] = fixed(x / 1000 ** scale, decIdx.length);
  fillInt(toks, intIdx, int, out, group);
  fillDec(toks, decIdx, frac, out);
  return out.join('');
}

function fraction(toks: Tok[], idx: number[], slash: number, x: number, out: string[]): string {
  let a = slash - 1;
  while (isPh(toks[a - 1])) a--;
  let b = slash + 1;
  while (b < toks.length && (isPh(toks[b]) || toks[b].t === 'lit' && /^\d$/.test(toks[b].v))) b++;
  const numIdx = idx.filter(i => i >= a && i < slash), intIdx = idx.filter(i => i < a);
  const denIdx = Array.from({ length: b - slash - 1 }, (_, k) => slash + 1 + k);
  const denText = denIdx.map(i => toks[i].v).join('');
  const fixedDen = /^[1-9]\d*$/.test(denText) ? +denText : 0;

  let whole = intIdx.length ? Math.floor(x) : 0, n: number, d: number;
  const f = x - whole;
  if (fixedDen) { d = fixedDen; n = Math.round(f * d); }
  else {
    [n, d] = bestFraction(f, 10 ** Math.min(denIdx.length, 15) - 1);
  }
  if (intIdx.length && n === d) { whole++; n = 0; }

  fillInt(toks, intIdx, whole || !n ? String(whole) : '', out);
  fillInt(toks, numIdx, String(n), out);
  const ds = String(d);
  denIdx.forEach((i, k) => {
    const p = toks[i].v;
    out[i] = fixedDen ? p : ds[k] ?? (p === '?' ? ' ' : p === '0' ? '0' : '');
  });
  // a whole number blanks out the fraction but keeps its width
  if (intIdx.length && !n) for (let i = a; i < b; i++) out[i] = out[i].replace(/./g, ' ');
  return out.join('');
}

// Best approximation of f in [0, 1) with a denominator up to max, smallest denominator on ties.
// Continued fractions keep this to a few dozen steps; trying every denominator hung on codes like # ?/??????????.
export function bestFraction(f: number, max: number): [number, number] {
  let n = Math.round(f), d = 1, err = Math.abs(f - n);
  const consider = (p: number, q: number) => {
    const e = Math.abs(f - p / q);
    if (q >= 1 && q <= max && (e < err - 1e-12 || Math.abs(e - err) <= 1e-12 && q < d)) { n = p; d = q; err = e; }
  };
  let [p0, q0, p1, q1] = [0, 1, 1, 0];
  let x = f;
  for (let i = 0; i < 64 && err; i++) {
    const a = Math.floor(x);
    const [p2, q2] = [a * p1 + p0, a * q1 + q0];
    if (q2 > max) {
      // the best within max may be a semiconvergent between the last two convergents
      const k = Math.floor((max - q0) / q1);
      for (const j of [k, Math.ceil(a / 2)]) if (j >= 1 && j <= k) consider(j * p1 + p0, j * q1 + q0);
      break;
    }
    consider(p2, q2);
    [p0, q0, p1, q1] = [p1, q1, p2, q2];
    if (x - a < 1e-15) break;
    x = 1 / (x - a);
  }
  return [n, d];
}

// Calendar date for an integer serial: [year, month, day, weekday]
function ymd(days: number, date1904: boolean): number[] {
  if (date1904) {
    const dt = new Date(Date.UTC(1904, 0, 1 + days));
    return [dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate(), dt.getUTCDay()];
  }
  // Excel treats 1900 as a leap year: serial 1 is 1900-01-01 (a Sunday to Excel), 60 is 1900-02-29, 0 is 1900-01-00
  const wd = (days + 6) % 7;
  if (days === 0) return [1900, 1, 0, wd];
  if (days === 60) return [1900, 2, 29, wd];
  const dt = new Date(Date.UTC(1899, 11, days < 60 ? 31 + days : 30 + days));
  return [dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate(), wd];
}

function fmtDate(toks: Tok[], v: number, date1904: boolean): string {
  // Excel rounds to the smallest shown unit (whole seconds unless fractions are shown) before splitting
  const k = Math.max(0, ...toks.map(t => t.n ?? 0)), per = 10 ** k;
  const ticks = Math.round(v * 86400 * per);
  const days = Math.floor(ticks / (86400 * per)), rem = ticks - days * 86400 * per;
  const sub = rem % per, secs = (rem - sub) / per;
  const H = Math.floor(secs / 3600), M = Math.floor(secs / 60) % 60, S = secs % 60;
  const [y, mo, d, wd] = ymd(days, date1904);
  const h12 = toks.some(t => t.t === 'ampm');
  const pad = (n: number, w: number) => String(n).padStart(w, '0');
  const frac = (t: Tok) => t.n ? '.' + pad(sub, k).slice(0, t.n) : '';
  return toks.map(t => {
    const w = t.v.length;
    if (t.t === 'elapsed') {
      const total = t.v[0] === 'h' ? days * 24 + H : t.v[0] === 'm' ? (days * 24 + H) * 60 + M : days * 86400 + secs;
      return pad(total, w) + frac(t);
    }
    if (t.t === 'ampm') return w > 3 ? (H < 12 ? 'AM' : 'PM') : H < 12 ? t.v[0] : t.v[2];
    if (t.t !== 'date') return t.v;
    switch (t.v[0]) {
      case 'y': return w > 2 ? pad(y, 4) : pad(y % 100, 2);
      case 'm': return w === 1 ? String(mo) : w === 2 ? pad(mo, 2) : w === 3 ? MONTHS[mo - 1].slice(0, 3) : w === 5 ? MONTHS[mo - 1][0] : MONTHS[mo - 1];
      case 'd': return w === 1 ? String(d) : w === 2 ? pad(d, 2) : w === 3 ? DAYS[wd].slice(0, 3) : DAYS[wd];
      case 'h': { const h = h12 ? H % 12 || 12 : H; return w > 1 ? pad(h, 2) : String(h); }
      case 'n': return w > 1 ? pad(M, 2) : String(M);
      default: return (w > 1 ? pad(S, 2) : String(S)) + frac(t);
    }
  }).join('');
}
