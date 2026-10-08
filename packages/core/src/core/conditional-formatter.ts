import { argb, escapeXml as escapeText } from './utils';
import { ConditionalFormat, DataBarRule, ColorScaleRule, ConditionalFormatRule, HighlightStyle } from './types';

const escapeXml = (val: unknown) => escapeText(String(val));

const color = (rgb: string) => `<color rgb="${argb(rgb)}"/>`;
const bound = (type: 'min' | 'max', value?: number) =>
  value === undefined ? `<cfvo type="${type}"/>` : `<cfvo type="num" val="${escapeXml(value)}"/>`;
const formula = (f: string | number) => `<formula>${escapeXml(String(f).replace(/^=/, ''))}</formula>`;
// Excel string literal inside a formula
const literal = (s: string) => `"${s.replace(/"/g, '""')}"`;

export class ConditionalFormatter {
  // `dxf` registers a highlight style and returns its dxfId
  constructor(private dxf: (style: HighlightStyle) => number = () => 0) {}

  // Priorities must be unique within a sheet: rule i gets priority i + 1
  toXml(formats: ConditionalFormat[]): string {
    return formats.map((cf, i) => {
      // The first cell of the first area: relative formulas are written for it
      const topLeft = cf.range.split(/[\s:]/)[0].replace(/\$/g, '');
      const rule = this.ruleXml(cf.rule, i + 1, topLeft);
      return `<conditionalFormatting sqref="${escapeXml(cf.range)}">${rule}</conditionalFormatting>`;
    }).join('\n');
  }

  private ruleXml(rule: ConditionalFormatRule, priority: number, cell: string): string {
    const head = (type: string, extra = '') => {
      const dxf = 'style' in rule ? ` dxfId="${this.dxf(rule.style)}"` : '';
      return `<cfRule type="${type}"${dxf} priority="${priority}"${extra}>`;
    };
    switch (rule.type) {
      case 'dataBar': return `${head('dataBar')}${this.buildDataBar(rule)}</cfRule>`;
      case 'colorScale': return `${head('colorScale')}${this.buildColorScale(rule)}</cfRule>`;
      case 'cellIs':
        return `${head('cellIs', ` operator="${escapeXml(rule.operator)}"`)}${rule.formulae.map(formula).join('')}</cfRule>`;
      case 'expression': return `${head('expression')}${formula(rule.formula)}</cfRule>`;
      case 'top10':
        return `${head('top10', `${rule.percent ? ' percent="1"' : ''}${rule.bottom ? ' bottom="1"' : ''} rank="${escapeXml(rule.rank)}"`)}</cfRule>`;
      case 'aboveAverage': return `${head('aboveAverage', rule.below ? ' aboveAverage="0"' : '')}</cfRule>`;
      case 'duplicateValues':
      case 'uniqueValues': return `${head(rule.type)}</cfRule>`;
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith': {
        // Excel stores the test as a formula too, and evaluates that one
        const t = literal(rule.text);
        const test = {
          containsText: `NOT(ISERROR(SEARCH(${t},${cell})))`,
          notContainsText: `ISERROR(SEARCH(${t},${cell}))`,
          beginsWith: `LEFT(${cell},LEN(${t}))=${t}`,
          endsWith: `RIGHT(${cell},LEN(${t}))=${t}`,
        }[rule.type];
        const operator = rule.type === 'notContainsText' ? 'notContains' : rule.type;
        return `${head(rule.type, ` operator="${operator}" text="${escapeXml(rule.text)}"`)}${formula(test)}</cfRule>`;
      }
      case 'iconSet': {
        const n = parseInt(rule.iconSet, 10);
        const cfvos = Array.from({ length: n }, (_, k) => `<cfvo type="percent" val="${Math.round(k * 100 / n)}"/>`).join('');
        const attrs = `${rule.reverse ? ' reverse="1"' : ''}${rule.showValue === false ? ' showValue="0"' : ''}`;
        return `${head('iconSet')}<iconSet iconSet="${escapeXml(rule.iconSet)}"${attrs}>${cfvos}</iconSet></cfRule>`;
      }
      default:
        throw new Error(`Unknown conditional format type "${(rule as { type: string }).type}".`);
    }
  }

  private buildDataBar(rule: DataBarRule): string {
    return `<dataBar>${bound('min', rule.minValue)}${bound('max', rule.maxValue)}${color(rule.color)}</dataBar>`;
  }

  private buildColorScale(rule: ColorScaleRule): string {
    const mid = rule.midColor !== undefined;
    return `<colorScale><cfvo type="min"/>${mid ? '<cfvo type="percentile" val="50"/>' : ''}<cfvo type="max"/>` +
      `${color(rule.minColor)}${mid ? color(rule.midColor!) : ''}${color(rule.maxColor)}</colorScale>`;
  }
}
