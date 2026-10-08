import {
  mapFormulaRefs, shiftFormula, colIndex, colLetter, attr, escapeXml, unescapeXml, MAX_ROWS, MAX_COLUMNS, type RefRole,
} from './utils';

// Rows or columns inserted before `at`, or at..at+count-1 deleted (1-based; column 1 is A). Later
// operations use the numbers left by earlier ones, as when they are done one after another in Excel.
export interface ShiftOp { at: number; count: number; insert: boolean }

// Where the rows (or columns) of a sheet end up
export class AxisMap {
  constructor(readonly ops: ShiftOp[]) {}

  // A cell in a deleted row is gone (null); a range keeps what remains: its start moves to the row
  // after the deletion and its end to the row before it, which turns a range of deleted rows inside out.
  map(n: number, role: RefRole): number | null {
    for (const { at, count, insert } of this.ops) {
      if (insert) {
        if (n >= at) n += count;
      } else if (n >= at + count) n -= count;
      else if (n >= at) {
        if (role === 'single') return null;
        n = role === 'start' ? at : at - 1;
      }
    }
    return n;
  }
}

export interface SheetShift { rows?: AxisMap; cols?: AxisMap }
// Sheet name (lower case: Excel compares them so) -> how its rows and columns move
export type ShiftMaps = Map<string, SheetShift>;

// Column mappers work on 0-based columns, as mapFormulaRefs passes them
const mapCol = (s: SheetShift | undefined, c: number, role: RefRole) => {
  if (!s?.cols) return c;
  const n = s.cols.map(c + 1, role);
  return n === null ? null : n - 1;
};
const mapRow = (s: SheetShift | undefined, r: number, role: RefRole) => (s?.rows ? s.rows.map(r, role) : r);

const firstRow = (area: string) => parseInt(/\d+/.exec(area)?.[0] ?? '1', 10);
const lastRow = (area: string) => parseInt(/(\d+)\D*$/.exec(area)?.[1] ?? '1', 10);
const firstCol = (area: string) => colIndex(/[A-Za-z]+/.exec(area)?.[0] ?? 'A');
const lastCol = (area: string) => colIndex(/([A-Za-z]+)\$?\d*\s*$/.exec(area)?.[1] ?? 'A');

// An area (A1, A1:B5, 3:7, C:D) after rows and columns move, or null when nothing of it is left
export function mapArea(area: string, s: SheetShift): string | null {
  const out = mapFormulaRefs(area.includes(':') ? area : `${area}:${area}`,
    (c, _abs, role) => mapCol(s, c, role), (r, _abs, role) => mapRow(s, r, role));
  if (out.includes('#REF!')) return null;
  const [a, b] = out.split(':');
  return a === b ? a : out;
}

// A space-separated list of areas (sqref); null when none is left
const mapAreas = (sqref: string, s: SheetShift) => {
  const areas = sqref.trim().split(/\s+/).map(a => mapArea(a, s)).filter((a): a is string => a !== null);
  return areas.length ? areas.join(' ') : null;
};

// A cell that should land somewhere even when its row or column is deleted (a selection, a scroll
// position): it moves to the row or column after the deleted ones
function mapCellStart(ref: string, s: SheetShift): string {
  const out = mapFormulaRefs(ref, c => mapCol(s, c, 'start'), r => mapRow(s, r, 'start'));
  return out.includes('#REF!') ? ref : out;
}

// The top-left cell of an area before and after the move, for formulas counted relative to it
type Anchor = { row: number; toRow: number; col: number; toCol: number };

// A formula on sheet `sheet` ('' for workbook-level names and charts). `anchor` is set for conditional
// formats and validations, whose relative references count from the top-left cell of their range: one
// that points into deleted rows or columns keeps its distance from that cell instead of becoming #REF!.
export function mapFormula(f: string, sheet: string, maps: ShiftMaps, anchor?: Anchor): string {
  const own = maps.get(sheet.toLowerCase());
  if (!own && !f.includes('!')) return f;
  const target = (q: string | undefined) => (q === undefined ? own : maps.get(q.toLowerCase()));
  return mapFormulaRefs(f,
    (c, abs, role, q) => {
      const s = target(q);
      const n = mapCol(s, c, role);
      return n === null && anchor && !abs && q === undefined ? c - anchor.col + anchor.toCol : n;
    },
    (r, abs, role, q) => {
      const s = target(q);
      const n = mapRow(s, r, role);
      return n === null && anchor && !abs && q === undefined ? r - anchor.row + anchor.toRow : n;
    });
}

function anchorOf(area: string, s: SheetShift): Anchor {
  const row = firstRow(area), col = firstCol(area);
  return { row, toRow: mapRow(s, row, 'start')!, col, toCol: mapCol(s, col, 'start')! };
}

// Counts the children of a wrapper (<mergeCells count="n">), removing the wrapper when none are left
function recount(xml: string, wrapper: string, child: string, countAttr = 'count'): string {
  return xml.replace(new RegExp(`<((?:\\w+:)?)${wrapper}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</\\1${wrapper}>)`, 'g'),
    (_whole, p: string, attrs: string, inner = '') => {
      const n = inner.match(new RegExp(`<(?:\\w+:)?${child}\\b`, 'g'))?.length ?? 0;
      if (!n) return '';
      const counted = attrs.replace(new RegExp(`\\s${countAttr}="\\d+"`), ` ${countAttr}="${n}"`);
      return `<${p}${wrapper}${counted}>${inner}</${p}${wrapper}>`;
    });
}

// An autoFilter keeps its column filters on the columns they filter; colId counts from its first column
function mapAutoFilter(xml: string, s: SheetShift): string {
  return xml.replace(/<((?:\w+:)?)autoFilter\b([^>]*?)(\/>|>([\s\S]*?)<\/\1autoFilter>)/g, (whole, p: string, attrs: string, rest: string, inner?: string) => {
    const ref = attr(attrs, 'ref');
    if (ref === null) return whole;
    const mapped = mapArea(ref, s);
    if (mapped === null) return '';
    const from = firstCol(ref), to = firstCol(mapped);
    if (inner !== undefined && s.cols) {
      inner = inner.replace(/<((?:\w+:)?)filterColumn\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1filterColumn>)/g, (column, _fp, fAttrs: string) => {
        const id = parseInt(attr(fAttrs, 'colId') ?? '0', 10);
        const c = mapCol(s, from + id, 'single');
        return c === null ? '' : column.replace(/\scolId="\d+"/, ` colId="${c - to}"`);
      });
      rest = `>${inner}</${p}autoFilter>`;
    }
    return `<${p}autoFilter${attrs.replace(/\sref="[^"]*"/, ` ref="${mapped}"`)}${rest}`;
  });
}

// The worksheet XML around <sheetData>: column widths, merges, conditional formats, validations,
// links, filters, page breaks, sparklines and selections follow their cells; formulas follow the
// cells they point at
export function mapSheetXml(xml: string, sheet: string, maps: ShiftMaps): string {
  const own = maps.get(sheet.toLowerCase());

  xml = xml.replace(/<((?:\w+:)?)(conditionalFormatting|dataValidation)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/\1\2>)/g,
    (whole, p: string, tag: string, attrs: string, inner = '') => {
      // x14 extensions keep the range in an <xm:sqref> child
      const child = /<((?:\w+:)?)sqref>([^<]*)<\/\1sqref>/.exec(inner);
      const sqref = attr(attrs, 'sqref') ?? child?.[2];
      let anchor: Anchor | undefined;
      if (own && sqref) {
        const mapped = mapAreas(sqref, own);
        if (mapped === null) return '';
        anchor = anchorOf(sqref, own);
        attrs = attrs.replace(/\ssqref="[^"]*"/, ` sqref="${mapped}"`);
        if (child) inner = inner.replace(child[0], `<${child[1]}sqref>${mapped}</${child[1]}sqref>`);
      }
      inner = inner.replace(/<((?:\w+:)?)(formula1|formula2|formula|f)>([^<]*)<\/\1\2>/g, (_m: string, fp: string, ftag: string, f: string) =>
        `<${fp}${ftag}>${escapeXml(mapFormula(unescapeXml(f), sheet, maps, anchor))}</${fp}${ftag}>`);
      return whole.endsWith('/>') && !inner ? `<${p}${tag}${attrs}/>` : `<${p}${tag}${attrs}>${inner}</${p}${tag}>`;
    });
  xml = recount(xml, 'dataValidations', 'dataValidation');

  // Sparklines (x14): the data range follows its cells; a sparkline whose own cell is deleted goes
  const mapF = (s: string) => s.replace(/<((?:\w+:)?)f>([^<]*)<\/\1f>/g, (_m: string, p: string, f: string) =>
    `<${p}f>${escapeXml(mapFormula(unescapeXml(f), sheet, maps))}</${p}f>`);
  xml = xml.replace(/<((?:\w+:)?)sparklineGroup\b([^>]*)>([\s\S]*?)<\/\1sparklineGroup>/g, (whole, p: string, attrs: string, inner: string) => {
    const list = /<((?:\w+:)?)sparklines>([\s\S]*?)<\/\1sparklines>/.exec(inner);
    if (!list) return whole;
    const lines = list[2].replace(/<((?:\w+:)?)sparkline>([\s\S]*?)<\/\1sparkline>/g, (_m: string, sp: string, body: string) => {
      const loc = /<((?:\w+:)?)sqref>([^<]*)<\/\1sqref>/.exec(body);
      if (own && loc) {
        const mapped = mapArea(loc[2], own);
        if (mapped === null) return '';
        body = body.replace(loc[0], `<${loc[1]}sqref>${mapped}</${loc[1]}sqref>`);
      }
      return `<${sp}sparkline>${mapF(body)}</${sp}sparkline>`;
    });
    if (!/<(?:\w+:)?sparkline>/.test(lines)) return '';
    // The group's own <xm:f> (a date axis) sits outside the list
    return `<${p}sparklineGroup${attrs}>${mapF(inner.slice(0, list.index))}<${list[1]}sparklines>${lines}</${list[1]}sparklines>` +
      `${mapF(inner.slice(list.index + list[0].length))}</${p}sparklineGroup>`;
  });
  xml = recount(xml, 'sparklineGroups', 'sparklineGroup', 'nonexistent');
  // An extension left empty, and then an empty extension list, go too
  xml = xml.replace(/<((?:\w+:)?)ext\b[^>]*>\s*<\/\1ext>/g, '').replace(/<((?:\w+:)?)extLst>\s*<\/\1extLst>/g, '');
  if (!own) return xml;

  // One element type at a time, so that one nested in another (sortState in autoFilter) is reached.
  // An element whose area is gone is removed with its children.
  const mapAttr = (tag: string, name: string) => {
    xml = xml.replace(new RegExp(`<((?:\\w+:)?)${tag}\\b([^>]*?)(/>|>([\\s\\S]*?)</\\1${tag}>)`, 'g'), (whole, p: string, attrs: string, rest: string) => {
      const value = attr(attrs, name);
      if (value === null) return whole;
      const mapped = name === 'sqref' ? mapAreas(value, own) : mapArea(value, own);
      return mapped === null ? '' : `<${p}${tag}${attrs.replace(new RegExp(`\\s${name}="[^"]*"`), ` ${name}="${mapped}"`)}${rest}`;
    });
  };
  mapAttr('mergeCell', 'ref');
  xml = xml.replace(/<(?:\w+:)?mergeCell\b[^>]*?\sref="[^":]*"[^>]*\/>/g, ''); // a merge left with one cell
  xml = recount(xml, 'mergeCells', 'mergeCell');
  mapAttr('hyperlink', 'ref');
  xml = recount(xml, 'hyperlinks', 'hyperlink', 'nonexistent');
  xml = mapAutoFilter(xml, own);
  for (const tag of ['dimension', 'sortState', 'sortCondition']) mapAttr(tag, 'ref');
  for (const tag of ['ignoredError', 'protectedRange']) mapAttr(tag, 'sqref');
  xml = recount(xml, 'ignoredErrors', 'ignoredError', 'nonexistent');
  xml = recount(xml, 'protectedRanges', 'protectedRange', 'nonexistent');

  // Column widths and styles: inserted columns get the default width, deleted ones take theirs along
  if (own.cols) {
    const cols = own.cols;
    xml = xml.replace(/<((?:\w+:)?)col\b([^>]*?)\/>/g, (_m, p: string, attrs: string) => {
      const min = cols.map(parseInt(attr(attrs, 'min') ?? '1', 10), 'start')!;
      const max = cols.map(parseInt(attr(attrs, 'max') ?? '1', 10), 'end')!;
      if (max < min) return '';
      return `<${p}col${attrs.replace(/\smin="\d+"/, ` min="${min}"`).replace(/\smax="\d+"/, ` max="${Math.min(max, MAX_COLUMNS)}"`)}/>`;
    });
    xml = recount(xml, 'cols', 'col', 'nonexistent');
  }

  // The view: a selection or scroll position in deleted cells moves to the cells after them
  xml = xml.replace(/<((?:\w+:)?)(selection|pane)\b([^>]*?)\/>/g, (_m, p: string, tag: string, attrs: string) => {
    attrs = attrs.replace(/\s(activeCell|topLeftCell)="([^"]*)"/g, (_a, name: string, ref: string) => ` ${name}="${mapCellStart(ref, own)}"`);
    const sqref = attr(attrs, 'sqref');
    if (sqref !== null) attrs = attrs.replace(/\ssqref="[^"]*"/, ` sqref="${mapAreas(sqref, own) ?? attr(attrs, 'activeCell') ?? 'A1'}"`);
    return `<${p}${tag}${attrs}/>`;
  });

  // A break's id is the 0-based row (column) it sits above (left of)
  for (const [section, axis] of [['rowBreaks', own.rows], ['colBreaks', own.cols]] as const) {
    if (!axis) continue;
    xml = xml.replace(new RegExp(`<((?:\\w+:)?)${section}\\b[\\s\\S]*?</\\1${section}>`), list => {
      list = list.replace(/<(?:\w+:)?brk\b([^>]*?)\/>/g, (whole, attrs: string) => {
        const n = axis.map(parseInt(attr(attrs, 'id') ?? '0', 10) + 1, 'start')!;
        return whole.replace(/\sid="\d+"/, ` id="${n - 1}"`);
      });
      const n = list.match(/<(?:\w+:)?brk\b/g)?.length ?? 0;
      return n ? list.replace(/\scount="\d+"/, ` count="${n}"`).replace(/\smanualBreakCount="\d+"/, ` manualBreakCount="${n}"`) : '';
    });
  }
  return xml;
}

// Streams a worksheet, moving its rows and cells and the references in its formulas. Sheets whose
// own cells do not move only have formulas that point at other sheets rewritten.
export function createShiftTransform(sheet: string, maps: ShiftMaps): TransformStream<string, string> {
  const own = maps.get(sheet.toLowerCase());
  let buffer = '';
  let p = '';
  let state: 'head' | 'rows' | 'tail' = 'head';
  // Shared formulas are written out per cell when the cells under them move unevenly
  const shared = new Map<string, { text: string; row: number; col: number; expand: boolean }>();
  const mapF = (f: string) => mapFormula(f, sheet, maps);

  const rewriteRow = (whole: string, attrs: string, inner: string | undefined): string => {
    const r = parseInt(attr(attrs, 'r') ?? '0', 10);
    const to = r ? mapRow(own, r, 'single') : r;
    if (to === null) return '';
    if (to > MAX_ROWS) throw new Error(`Sheet "${sheet}": inserting rows pushes row ${r} past Excel's last row.`);
    if (to === r && !own?.cols && !inner?.includes('<' + p + 'f')) return whole;

    let col = 0;
    const cells = (inner ?? '').replace(new RegExp(`<${p}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}c>)`, 'g'), (cell, cAttrs: string, body?: string) => {
      const ref = /\sr="([A-Za-z]+)\d+"/.exec(cAttrs);
      if (ref) col = colIndex(ref[1]);
      const c = col++;
      const toCol = mapCol(own, c, 'single');
      if (toCol === null) return '';
      if (toCol >= MAX_COLUMNS) throw new Error(`Sheet "${sheet}": inserting columns pushes column ${colLetter(c)} past Excel's last column.`);
      if (to !== r || toCol !== c) cell = cell.replace(/\sr="[A-Za-z]+\d+"/, ` r="${colLetter(toCol)}${to}"`);
      const f = body && new RegExp(`<${p}f\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}f>)`).exec(body);
      if (!f) return cell;
      const text = unescapeXml(f[2] ?? '');
      const kind = attr(f[1], 't');
      const si = attr(f[1], 'si');
      let xml: string | undefined;
      if (kind === 'shared' && si !== null) {
        if (text) {
          const mapped = mapF(text);
          const expand = !!own || mapped !== text;
          shared.set(si, { text, row: r, col: c, expand });
          if (expand) xml = `<${p}f>${escapeXml(mapped)}</${p}f>`;
        } else {
          const a = shared.get(si);
          if (a) {
            // Even when the anchor's own references stay, this cell's may point into moved rows
            const own = shiftFormula(a.text, r - a.row, c - a.col);
            const mapped = mapF(own);
            if (a.expand || mapped !== own) xml = `<${p}f>${escapeXml(mapped)}</${p}f>`;
          }
        }
      } else if (kind === 'dataTable') {
        // A What-If data table: its result range and its input cells (r1, r2) follow their cells
        let fAttrs = f[1];
        const range = attr(fAttrs, 'ref');
        if (own && range) fAttrs = fAttrs.replace(/\sref="[^"]*"/, ` ref="${mapArea(range, own) ?? range}"`);
        fAttrs = fAttrs.replace(/\s(r1|r2)="([^"]*)"/g, (_m, name: string, ref: string) => ` ${name}="${escapeXml(mapF(unescapeXml(ref)))}"`);
        xml = f[0].replace(f[1], () => fAttrs);
      } else if (text) {
        let fAttrs = f[1];
        const range = attr(fAttrs, 'ref');
        if (own && range) fAttrs = fAttrs.replace(/\sref="[^"]*"/, ` ref="${mapArea(range, own) ?? range}"`);
        xml = `<${p}f${fAttrs}>${escapeXml(mapF(text))}</${p}f>`;
      }
      return xml === undefined ? cell : cell.replace(f[0], () => xml!);
    });
    let open = to === r ? attrs : attrs.replace(/\sr="\d+"/, ` r="${to}"`);
    if (own?.cols) open = open.replace(/\sspans="[^"]*"/, ''); // an optional hint, now stale
    return inner === undefined ? `<${p}row${open}/>` : `<${p}row${open}>${cells}</${p}row>`;
  };

  return new TransformStream({
    transform(chunk, controller) {
      buffer += chunk;
      let out = '';
      while (true) {
        if (state === 'head') {
          const m = /<((?:\w+:)?)sheetData\b[^>]*?(\/?)>/.exec(buffer);
          if (!m) break;
          p = m[1];
          out += mapSheetXml(buffer.slice(0, m.index), sheet, maps) + m[0];
          buffer = buffer.slice(m.index + m[0].length);
          state = m[2] ? 'tail' : 'rows';
        } else if (state === 'rows') {
          const m = new RegExp(`^\\s*(?:<${p}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}row>)|</${p}sheetData>)`).exec(buffer);
          if (!m) break;
          if (m[1] === undefined) {
            out += m[0];
            state = 'tail';
          } else out += rewriteRow(m[0].trimStart(), m[1], m[2]);
          buffer = buffer.slice(m[0].length);
        } else break; // the tail is mapped whole at the end
      }
      if (out) controller.enqueue(out);
    },
    flush(controller) {
      if (state !== 'tail') {
        controller.error(new Error(`Sheet "${sheet}" has no complete <sheetData> element.`));
        return;
      }
      controller.enqueue(mapSheetXml(buffer, sheet, maps));
    },
  });
}

// Notes (legacy and threaded) on moved cells move with them; notes on deleted cells go
export function mapComments(xml: string, s: SheetShift): string {
  return xml.replace(/<((?:\w+:)?)(comment|threadedComment)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1\2>)/g, (whole, _p, _t, attrs: string) => {
    const ref = attr(attrs, 'ref');
    if (!ref) return whole;
    const mapped = mapArea(ref, s);
    return mapped === null ? '' : whole.replace(/(\sref=")[^"]*"/, `$1${mapped}"`);
  });
}

// The note boxes drawn for legacy notes (VML). Prefixes vary: v:/x: from Excel, ns1:/ns2: from openpyxl.
export function mapVml(xml: string, s: SheetShift): string {
  return xml.replace(/<((?:\w+:)?)shape\b[\s\S]*?<\/\1shape>/g, shape => {
    const row = /<((?:\w+:)?)Row>(\d+)<\/\1Row>/.exec(shape);
    const col = /<((?:\w+:)?)Column>(\d+)<\/\1Column>/.exec(shape);
    if (!row || !col) return shape;
    const r = parseInt(row[2], 10), c = parseInt(col[2], 10);
    const toR = mapRow(s, r + 1, 'single'), toC = mapCol(s, c, 'single');
    if (toR === null || toC === null) return '';
    const dr = toR - 1 - r, dc = toC - c;
    if (!dr && !dc) return shape;
    return shape
      .replace(row[0], `<${row[1]}Row>${toR - 1}</${row[1]}Row>`)
      .replace(col[0], `<${col[1]}Column>${toC}</${col[1]}Column>`)
      .replace(/<((?:\w+:)?)Anchor>([^<]*)<\/\1Anchor>/, (_m, p: string, a: string) => {
        const v = a.split(',').map(x => parseInt(x.trim(), 10));
        v[0] += dc; v[4] += dc; v[2] += dr; v[6] += dr;
        return `<${p}Anchor>${v.join(', ')}</${p}Anchor>`;
      });
  });
}

// Pictures and charts move with their cells; one spanning inserted rows or columns grows unless it is
// set to keep its size (oneCell) or position (absolute). One whose cells are all deleted shrinks.
export function mapDrawing(xml: string, s: SheetShift): string {
  const at = (el: string, axis: string) => new RegExp(`(<(?:\\w+:)?${el}>[\\s\\S]*?<(?:\\w+:)?${axis}>)(\\d+)(<)`);
  return xml.replace(/<((?:\w+:)?)(twoCellAnchor|oneCellAnchor)\b([^>]*)>([\s\S]*?)<\/\1\2>/g, (whole, p: string, tag: string, attrs: string, inner: string) => {
    const editAs = attr(attrs, 'editAs') ?? (tag === 'oneCellAnchor' ? 'oneCell' : 'twoCell');
    if (editAs === 'absolute') return whole;
    for (const [axis, map] of [['row', (n: number, role: RefRole) => mapRow(s, n + 1, role)! - 1], ['col', (n: number, role: RefRole) => mapCol(s, n, role)!]] as const) {
      const from = at('from', axis).exec(inner);
      if (!from) continue;
      const f = parseInt(from[2], 10);
      const nf = map(f, 'start');
      const to = at('to', axis).exec(inner);
      inner = inner.replace(at('from', axis), `$1${nf}$3`);
      if (to) {
        const t = parseInt(to[2], 10);
        inner = inner.replace(at('to', axis), `$1${Math.max(nf, editAs === 'oneCell' ? t + nf - f : map(t, 'end'))}$3`);
      }
    }
    return `<${p}${tag}${attrs}>${inner}</${p}${tag}>`;
  });
}

// A table's range follows its cells. Its columns follow theirs: deleted ones go, and inserted ones get
// a new column whose name is returned as a header cell to write. Deleting the header row or every data
// row is refused.
export function mapTable(xml: string, sheet: string, maps: ShiftMaps): { xml: string; headers: { ref: string; name: string }[] } {
  const own = maps.get(sheet.toLowerCase());
  const headers: { ref: string; name: string }[] = [];
  if (own) {
    const table = /<(?:\w+:)?table\b[^>]*>/.exec(xml)?.[0] ?? '';
    const name = attr(table, 'displayName') ?? attr(table, 'name') ?? '';
    const ref = attr(table, 'ref') ?? '';
    const header = attr(table, 'headerRowCount') !== '0';
    if (header && mapRow(own, firstRow(ref), 'single') === null) {
      throw new Error(`Deleting rows on sheet "${sheet}" would remove the header row of table "${name}".`);
    }
    const mapped = mapArea(ref, own);
    if (mapped === null) throw new Error(`Deleting cells on sheet "${sheet}" would remove table "${name}".`);
    const totals = parseInt(attr(table, 'totalsRowCount') ?? '0', 10);
    if (lastRow(mapped) - firstRow(mapped) + 1 < (header ? 1 : 0) + totals + 1) {
      throw new Error(`Deleting rows on sheet "${sheet}" would leave table "${name}" without data rows.`);
    }

    if (own.cols) {
      const from = firstCol(ref), to = firstCol(mapped), width = lastCol(mapped) - to + 1;
      xml = xml.replace(/<((?:\w+:)?)tableColumns\b([^>]*)>([\s\S]*?)<\/\1tableColumns>/, (_m, p: string, attrs: string, inner: string) => {
        const old = [...inner.matchAll(/<((?:\w+:)?)tableColumn\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1tableColumn>)/g)];
        const names = new Set(old.map(c => (attr(c[2], 'name') ?? '').toLowerCase()));
        let id = Math.max(0, ...old.map(c => parseInt(attr(c[2], 'id') ?? '0', 10)));
        const slots: string[] = new Array(width);
        old.forEach((c, i) => {
          const at = mapCol(own, from + i, 'single');
          if (at !== null) slots[at - to] = c[0];
        });
        let k = 1;
        for (let i = 0; i < width; i++) {
          if (slots[i]) continue;
          while (names.has(`column${k}`)) k++;
          names.add(`column${k}`);
          slots[i] = `<${p}tableColumn id="${++id}" name="Column${k}"/>`;
          if (header) headers.push({ ref: colLetter(to + i) + firstRow(mapped), name: `Column${k}` });
        }
        return `<${p}tableColumns${attrs.replace(/\scount="\d+"/, ` count="${width}"`)}>${slots.join('')}</${p}tableColumns>`;
      });
    }
    xml = mapAutoFilter(xml, own).replace(/(<(?:\w+:)?(?:table|sortState|sortCondition)\b[^>]*?\sref=")([^"]*)"/g,
      (_m, pre: string, area: string) => `${pre}${mapArea(area, own) ?? area}"`);
  }
  xml = xml.replace(/<((?:\w+:)?)(calculatedColumnFormula|totalsRowFormula)\b([^>]*)>([^<]*)<\/\1\2>/g,
    (_m, p: string, tag: string, attrs: string, f: string) => `<${p}${tag}${attrs}>${escapeXml(mapFormula(unescapeXml(f), sheet, maps))}</${p}${tag}>`);
  return { xml, headers };
}

// Chart series follow the cells they plot. Cached values are dropped when a series moves, and Excel
// rebuilds them from the cells.
export function mapChart(xml: string, maps: ShiftMaps): string {
  let changed = false;
  const out = xml.replace(/(<(?:\w+:)?f>)([^<]*)(<\/(?:\w+:)?f>)/g, (whole, open: string, f: string, close: string) => {
    const mapped = mapFormula(unescapeXml(f), '', maps);
    if (mapped === unescapeXml(f)) return whole;
    changed = true;
    return open + escapeXml(mapped) + close;
  });
  return changed ? out.replace(/<((?:\w+:)?)(numCache|strCache)\b[\s\S]*?<\/\1\2>/g, '') : xml;
}

// A pivot table on a sheet whose rows or columns move is drawn where its cells went
export function mapPivotTable(xml: string, s: SheetShift): string {
  return xml.replace(/(<(?:\w+:)?location\b[^>]*?\sref=")([^"]*)"/, (_m, pre: string, ref: string) => `${pre}${mapArea(ref, s) ?? ref}"`);
}

// A pivot cache over moved cells reads the new range and refreshes when the file opens
export function mapPivotCache(xml: string, maps: ShiftMaps): string {
  let changed = false;
  const out = xml.replace(/<((?:\w+:)?)worksheetSource\b([^>]*?)\/>/, (whole, p: string, attrs: string) => {
    const ref = attr(attrs, 'ref');
    const s = maps.get((attr(attrs, 'sheet') ?? '').toLowerCase());
    const mapped = ref && s ? mapArea(ref, s) : null;
    if (!mapped || mapped === ref) return whole;
    changed = true;
    return `<${p}worksheetSource${attrs.replace(/\sref="[^"]*"/, ` ref="${mapped}"`)}/>`;
  });
  if (!changed) return xml;
  return out.replace(/<((?:\w+:)?)pivotCacheDefinition\b([^>]*?)>/, (_m, p: string, attrs: string) =>
    `<${p}pivotCacheDefinition${attrs.replace(/\srefreshOnLoad="[^"]*"/, '')} refreshOnLoad="1">`);
}

// Workbook defined names (print areas, named ranges) follow the cells they point at
export function mapDefinedNames(wb: string, maps: ShiftMaps): string {
  return wb.replace(/(<(?:\w+:)?definedName\b[^>]*>)([^<]*)(<\/(?:\w+:)?definedName>)/g, (_m, open: string, f: string, close: string) =>
    open + escapeXml(mapFormula(unescapeXml(f), '', maps)) + close);
}
