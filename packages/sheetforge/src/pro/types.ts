export type CellValue = string | number | boolean | null;

export interface BorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
  color?: string; // ARGB hex e.g. "FF000000"
}

export interface CellBorder {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

export interface CellFill {
  type: 'solid';
  fgColor: string; // ARGB hex
}

export interface GradientStop {
  position: number; // 0.0 to 1.0
  color: string;    // ARGB hex
}

export interface GradientFill {
  type: 'gradient';
  degree?: number;
  stops: GradientStop[];
}

export interface CellFont {
  name?: string;         // e.g. "Calibri"
  size?: number;         // pt
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;        // ARGB hex
}

export interface CellAlignment {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify';
  vertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
}

export interface CellStyle {
  font?: CellFont;
  fill?: CellFill | GradientFill;
  border?: CellBorder;
  alignment?: CellAlignment;
  numFmt?: string; // e.g. "#,##0.00", "dd/mm/yyyy"
}

export interface DataBarRule {
  type: 'dataBar';
  color: string; // ARGB hex
  minValue?: number;
  maxValue?: number;
}

export interface ColorScaleRule {
  type: 'colorScale';
  minColor: string;
  midColor?: string;
  maxColor: string;
}

export type ConditionalFormatRule = DataBarRule | ColorScaleRule;

export interface ConditionalFormat {
  range: string; // e.g. "A1:A100"
  rule: ConditionalFormatRule;
}

export interface StyledCell {
  value: CellValue;
  style?: CellStyle;
  formula?: string; // e.g. "=SUM(A1:A10)"
}

export type Row = (CellValue | StyledCell)[];

export interface SheetOptions {
  name?: string;
  columnWidths?: number[]; // in characters
  conditionalFormats?: ConditionalFormat[];
}
