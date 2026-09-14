import { ConditionalFormat, DataBarRule, ColorScaleRule } from './types';

export class ConditionalFormatter {
  toXml(formats: ConditionalFormat[]): string {
    return formats.map(cf => {
      const rule = cf.rule;
      if (rule.type === 'dataBar') {
        return this.buildDataBar(cf.range, rule);
      } else if (rule.type === 'colorScale') {
        return this.buildColorScale(cf.range, rule);
      }
      return '';
    }).join('\n');
  }

  private buildDataBar(range: string, rule: DataBarRule): string {
    return `<conditionalFormatting sqref="${range}">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="${rule.color}"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>`;
  }

  private buildColorScale(range: string, rule: ColorScaleRule): string {
    const mid = rule.midColor
      ? `<cfvo type="percentile" val="50"/>`
      : '';
    const midColor = rule.midColor
      ? `<color rgb="${rule.midColor}"/>`
      : '';

    return `<conditionalFormatting sqref="${range}">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      ${mid}
      <cfvo type="max"/>
      <color rgb="${rule.minColor}"/>
      ${midColor}
      <color rgb="${rule.maxColor}"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>`;
  }
}
