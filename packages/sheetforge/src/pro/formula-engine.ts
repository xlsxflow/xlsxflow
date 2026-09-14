export type FormulaResult = number | string | boolean | null;

// Client-side formula evaluator that calculates values before exporting.
// This injects both <f> (formula string) and <v> (cached value) for cross-platform fidelity.
export class FormulaEngine {
  private cells: Map<string, FormulaResult> = new Map();

  clear() {
    this.cells.clear();
  }

  // Load a grid of values for formula context (keyed by cell ref e.g. "A1")
  loadData(data: (FormulaResult)[][], startRow = 1) {
    data.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const ref = this.toRef(ri + startRow, ci);
        this.cells.set(ref, cell);
      });
    });
  }

  // Evaluate a formula and return its result
  evaluate(formula: string): FormulaResult {
    const clean = formula.replace(/^=/, '').trim();

    // SUM(range)
    const sumMatch = clean.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (sumMatch) return this.sum(sumMatch[1], sumMatch[2]);

    // AVERAGE(range)
    const avgMatch = clean.match(/^AVERAGE\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (avgMatch) return this.average(avgMatch[1], avgMatch[2]);

    // COUNT(range)
    const countMatch = clean.match(/^COUNT\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (countMatch) return this.count(countMatch[1], countMatch[2]);

    // MAX(range)
    const maxMatch = clean.match(/^MAX\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (maxMatch) return this.max(maxMatch[1], maxMatch[2]);

    // MIN(range)
    const minMatch = clean.match(/^MIN\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
    if (minMatch) return this.min(minMatch[1], minMatch[2]);

    // Literal cell reference
    if (/^[A-Z]+\d+$/i.test(clean)) {
      return this.cells.get(clean.toUpperCase()) ?? null;
    }

    return null;
  }

  private getRangeValues(startRef: string, endRef: string): number[] {
    const [startCol, startRow] = this.parseRef(startRef);
    const [endCol, endRow] = this.parseRef(endRef);
    const values: number[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const ref = this.toRef(r, c - 1);
        const val = this.cells.get(ref);
        if (typeof val === 'number') values.push(val);
      }
    }
    return values;
  }

  private sum(startRef: string, endRef: string): number {
    return this.getRangeValues(startRef, endRef).reduce((a, b) => a + b, 0);
  }

  private average(startRef: string, endRef: string): number {
    const vals = this.getRangeValues(startRef, endRef);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  }

  private count(startRef: string, endRef: string): number {
    return this.getRangeValues(startRef, endRef).length;
  }

  private max(startRef: string, endRef: string): number {
    return Math.max(...this.getRangeValues(startRef, endRef));
  }

  private min(startRef: string, endRef: string): number {
    return Math.min(...this.getRangeValues(startRef, endRef));
  }

  // Convert "A1" -> [colNum, rowNum]
  private parseRef(ref: string): [number, number] {
    const match = ref.match(/^([A-Z]+)(\d+)$/i)!;
    const col = match[1].toUpperCase().split('').reduce((acc, ch) =>
      acc * 26 + (ch.charCodeAt(0) - 64), 0);
    const row = parseInt(match[2], 10);
    return [col, row];
  }

  // colIndex is 0-based
  private toRef(row: number, colIndex: number): string {
    let col = '';
    let c = colIndex + 1;
    while (c > 0) {
      col = String.fromCharCode(((c - 1) % 26) + 65) + col;
      c = Math.floor((c - 1) / 26);
    }
    return `${col}${row}`;
  }
}
