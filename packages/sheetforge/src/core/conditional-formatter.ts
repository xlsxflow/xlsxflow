import { ConditionalFormat, DataBarRule, ColorScaleRule } from './types';

function escapeXml(val: string): string {
  return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const color = (rgb: string) => `<color rgb="${escapeXml(rgb)}"/>`;
const bound = (type: 'min' | 'max', value?: number) =>
  value === undefined ? `<cfvo type="${type}"/>` : `<cfvo type="num" val="${value}"/>`;

export class ConditionalFormatter {
  // Priorities must be unique within a sheet: rule i gets priority i + 1
  toXml(formats: ConditionalFormat[]): string {
    return formats.map((cf, i) => {
      const rule = cf.rule;
      let body = '';
      if (rule.type === 'dataBar') body = this.buildDataBar(rule);
      else if (rule.type === 'colorScale') body = this.buildColorScale(rule);
      if (!body) return '';
      return `<conditionalFormatting sqref="${escapeXml(cf.range)}"><cfRule type="${rule.type}" priority="${i + 1}">${body}</cfRule></conditionalFormatting>`;
    }).join('\n');
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
