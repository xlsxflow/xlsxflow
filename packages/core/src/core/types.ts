export type CellValue = string | number | boolean | Date | null;

export interface BorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed'
    | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
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
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'middle' | 'bottom' | 'justify' | 'distributed'; // 'middle' is written as 'center'
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

// What a highlighting rule applies to matching cells (Excel's "differential" format)
export type HighlightStyle = Pick<CellStyle, 'font' | 'fill' | 'border' | 'numFmt'>;

export type ComparisonOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan'
  | 'greaterThanOrEqual' | 'lessThanOrEqual';

export interface CellIsRule {
  type: 'cellIs';
  operator: ComparisonOperator;
  formulae: (string | number)[]; // one value, two for between/notBetween; strings are formulas ("$B$1", "\"text\"")
  style: HighlightStyle;
}

export interface ExpressionRule {
  type: 'expression';
  formula: string; // relative to the range's top-left cell, e.g. "$C2>100"
  style: HighlightStyle;
}

export interface Top10Rule {
  type: 'top10';
  rank: number;
  percent?: boolean;
  bottom?: boolean;
  style: HighlightStyle;
}

export interface AboveAverageRule {
  type: 'aboveAverage';
  below?: boolean;
  style: HighlightStyle;
}

export interface TextRule {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
  text: string;
  style: HighlightStyle;
}

export interface DuplicateRule {
  type: 'duplicateValues' | 'uniqueValues';
  style: HighlightStyle;
}

export interface IconSetRule {
  type: 'iconSet';
  iconSet: '3Arrows' | '3ArrowsGray' | '3Flags' | '3TrafficLights1' | '3TrafficLights2' | '3Signs' | '3Symbols'
    | '3Symbols2' | '4Arrows' | '4ArrowsGray' | '4RedToBlack' | '4Rating' | '4TrafficLights' | '5Arrows'
    | '5ArrowsGray' | '5Rating' | '5Quarters';
  reverse?: boolean;
  showValue?: boolean; // default true
}

export type ConditionalFormatRule = DataBarRule | ColorScaleRule | CellIsRule | ExpressionRule | Top10Rule
  | AboveAverageRule | TextRule | DuplicateRule | IconSetRule;

export interface ConditionalFormat {
  range: string; // e.g. "A1:A100"
  rule: ConditionalFormatRule;
}

export interface StyledCell {
  value: CellValue;
  style?: CellStyle;
  formula?: string; // e.g. "=SUM(A1:A10)"
  hyperlink?: string; // URL, or "#Sheet2!A1" for a location in the workbook
  richText?: RichTextRun[]; // written instead of value (value still feeds formulas)
  comment?: string | CellComment; // a note, shown when hovering the cell
}

export interface CellComment {
  text: string | RichTextRun[];
  author?: string;
}

export interface RichTextRun {
  text: string;
  font?: CellFont;
}

export type Row = (CellValue | StyledCell)[];

export interface DataValidation {
  sqref: string; // e.g. "A1:A10"
  type?: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  operator?: ComparisonOperator; // for whole, decimal, date, time and textLength
  formula1?: string;
  formula2?: string;
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  error?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
}

export interface RowOptions {
  height?: number;       // points
  hidden?: boolean;
  outlineLevel?: number; // 1-7: grouped rows
}

export interface ColumnOptions {
  width?: number;        // characters
  hidden?: boolean;
  outlineLevel?: number; // 1-7: grouped columns
}

export interface SheetProtection {
  password?: string; // Excel's legacy 16-bit hash; deters edits, is not encryption
  // Actions still allowed on the protected sheet (Excel's defaults: only selecting cells)
  formatCells?: boolean; formatColumns?: boolean; formatRows?: boolean; insertRows?: boolean;
  insertColumns?: boolean; deleteRows?: boolean; deleteColumns?: boolean; sort?: boolean; autoFilter?: boolean;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: number;    // Excel code: 1 Letter, 9 A4, 5 Legal...
  fitToWidth?: number;   // pages; set fitToWidth/fitToHeight to scale to fit (0 = automatic)
  fitToHeight?: number;
  scale?: number;        // percent
  margins?: { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number }; // inches
  header?: string;       // Excel header/footer codes, e.g. "&CPage &P of &N"
  footer?: string;
  printArea?: string;    // e.g. "A1:F40"
  printTitleRows?: string; // rows repeated on every page, e.g. "1:2"
  gridLines?: boolean;
  horizontalCentered?: boolean;
}

export interface TableOptions {
  name: string;          // unique in the workbook, no spaces
  ref: string;           // header row through last data row, e.g. "A1:D20"
  columns?: string[];    // header names; read from the header row when rows are an array
  style?: string;        // e.g. "TableStyleMedium2" (default)
  showRowStripes?: boolean; // default true
  autoFilter?: boolean;  // default true
}

export interface SheetOptions {
  name?: string;
  columnWidths?: number[]; // in characters
  conditionalFormats?: ConditionalFormat[];
  dataValidations?: DataValidation[];
  mergeCells?: string[]; // e.g. ["A1:C1", "D1:E2"]
  autoFitColumns?: boolean;
  freezePanes?: { row?: number, col?: number };
  autoFilter?: string; // e.g. "A1:D1"
  images?: SheetImage[];
  rows?: Record<number, RowOptions>;  // by 1-based row number
  columns?: ColumnOptions[];          // by column index; width here overrides columnWidths
  protection?: SheetProtection | true;
  pageSetup?: PageSetup;
  tables?: TableOptions[];
  tabColor?: string;                  // ARGB hex
}

// A PNG, JPEG or GIF picture (format detected from the bytes). Passing the same `data` object
// to several images stores the file once.
export type SheetImage = {
  data: Uint8Array | ArrayBuffer;
  altText?: string;
} & (
  | { range: string }                                // stretched over cells, e.g. "B2:D8"
  | { at: string; width?: number; height?: number }  // top-left cell; size in pixels, default the image's own
);
