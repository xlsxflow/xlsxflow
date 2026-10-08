export type FormulaResult = number | string | boolean | null;

// A formula cell handed to loadData; its value is computed when a formula reads it
export interface FormulaCell { formula: string }

type TokenType = 'NUMBER' | 'STRING' | 'CELL' | 'RANGE' | 'IDENTIFIER' | 'OP' | 'PAREN_L' | 'PAREN_R' | 'COMMA';

interface Token {
  type: TokenType;
  value: string;
}

interface ASTNode {
  type: 'NUMBER' | 'STRING' | 'BOOL' | 'CELL' | 'RANGE' | 'CALL' | 'BINARY' | 'NEG' | 'PERCENT';
  value?: any;
  left?: ASTNode;
  right?: ASTNode;
  operator?: string;
  name?: string;
  args?: ASTNode[];
}

// An Excel error value (#DIV/0!, #VALUE!), kept apart from text that happens to start with "#"
export class FormulaError {
  constructor(readonly code: string) {}
}

export const ERROR_CODES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);

// Syntax or functions this engine doesn't know: no cached value, Excel computes it on open
class Unsupported extends Error {}
// A formula that reads its own cell, directly or through others
class Circular extends Error {}

type Value = number | string | boolean | null | FormulaError;

const DIV0 = new FormulaError('#DIV/0!');
const VALUE = new FormulaError('#VALUE!');

// Numbers turned into text keep Excel's 15 significant digits ("0.333333333333333", 1E+21)
export function numberText(n: number): string {
  return String(Number(n.toPrecision(15))).replace('e', 'E');
}

export class FormulaEngine {
  private cells = new Map<string, FormulaResult | FormulaCell>();
  private results = new Map<string, Value>();
  private evaluating = new Set<string>();

  clear() {
    this.cells.clear();
    this.results.clear();
  }

  loadData(data: (FormulaResult | FormulaCell)[][], startRow = 1) {
    data.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const ref = this.toRef(ri + startRow, ci);
        this.cells.set(ref, cell);
      });
    });
  }

  // Errors come back as their code ("#DIV/0!")
  evaluate(formula: string): FormulaResult {
    const v = this.evaluateRaw(formula);
    return v instanceof FormulaError ? v.code : v;
  }

  // Like evaluate, but errors stay FormulaError so text such as "#1 pick" is told apart from them
  evaluateRaw(formula: string): Value {
    try {
      return this.run(formula);
    } catch (err) {
      // A circular reference is stored as 0, as Excel does
      if (err instanceof Circular) return 0;
      return null;
    }
  }

  // The value of a loaded cell, computing its formula (and the formulas it reads) when it has one
  cellValue(ref: string): Value {
    try {
      return this.lookup(ref.toUpperCase());
    } catch (err) {
      if (err instanceof Circular) return 0;
      return null;
    }
  }

  private run(formula: string): Value {
    const clean = formula.replace(/^=/, '').trim();
    if (!clean) return null;
    // A formula is never blank: =A1 over an empty cell is 0
    return this.scalar(this.parse(this.tokenize(clean))) ?? 0;
  }

  private lookup(ref: string): Value {
    const cell = this.cells.get(ref);
    if (cell === null || typeof cell !== 'object') return cell ?? null;
    if (this.results.has(ref)) {
      const known = this.results.get(ref)!;
      if (known === undefined) throw new Unsupported();
      return known;
    }
    if (this.evaluating.has(ref)) throw new Circular();
    this.evaluating.add(ref);
    try {
      const v = this.run(cell.formula);
      this.results.set(ref, v);
      return v;
    } catch (err) {
      if (!(err instanceof Circular)) this.results.set(ref, undefined as any);
      throw err;
    } finally {
      this.evaluating.delete(ref);
    }
  }

  private tokenize(expr: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    while (i < expr.length) {
      const char = expr[i];
      if (/\s/.test(char)) { i++; continue; }
      if (char === '(') { tokens.push({ type: 'PAREN_L', value: '(' }); i++; continue; }
      if (char === ')') { tokens.push({ type: 'PAREN_R', value: ')' }); i++; continue; }
      if (char === ',') { tokens.push({ type: 'COMMA', value: ',' }); i++; continue; }
      if (/[+\-*/^<>=%&]/.test(char)) {
        let op = char;
        if (char === '<' || char === '>') {
          if (expr[i + 1] === '=') { op += '='; i++; }
          else if (char === '<' && expr[i + 1] === '>') { op += '>'; i++; }
        }
        tokens.push({ type: 'OP', value: op });
        i++;
        continue;
      }
      if (char === '"') {
        // "" inside a string is one quote
        let str = '';
        i++;
        for (;;) {
          if (i >= expr.length) throw new Unsupported();
          if (expr[i] === '"') {
            if (expr[i + 1] !== '"') break;
            i++;
          }
          str += expr[i++];
        }
        i++;
        tokens.push({ type: 'STRING', value: str });
        continue;
      }
      const num = /^(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?/.exec(expr.slice(i));
      if (num) {
        tokens.push({ type: 'NUMBER', value: num[0] });
        i += num[0].length;
        continue;
      }
      if (/[A-Za-z$]/.test(char)) {
        // "$" only pins a reference when copied; it does not change what it points at
        let id = '';
        while (i < expr.length && /[A-Za-z0-9$_.]/.test(expr[i])) { id += expr[i++]; }
        id = id.replace(/\$/g, '');
        if (i < expr.length && expr[i] === ':') {
          i++;
          let endId = '';
          while (i < expr.length && /[A-Za-z0-9$]/.test(expr[i])) { endId += expr[i++]; }
          endId = endId.replace(/\$/g, '');
          if (!/^[A-Z]{1,3}\d+$/i.test(id) || !/^[A-Z]{1,3}\d+$/i.test(endId)) throw new Unsupported();
          tokens.push({ type: 'RANGE', value: id.toUpperCase() + ':' + endId.toUpperCase() });
        } else if (/^[A-Z]{1,3}\d+$/i.test(id)) {
          tokens.push({ type: 'CELL', value: id.toUpperCase() });
        } else {
          tokens.push({ type: 'IDENTIFIER', value: id.toUpperCase() });
        }
        continue;
      }
      // Sheet references ('Q1'!A1, Data!A1), arrays, error literals and the rest
      throw new Unsupported();
    }
    return tokens;
  }

  private parse(tokens: Token[]): ASTNode {
    let pos = 0;
    const isOp = (...ops: string[]) => pos < tokens.length && tokens[pos].type === 'OP' && ops.includes(tokens[pos].value);
    const expect = (type: TokenType) => {
      if (tokens[pos]?.type !== type) throw new Unsupported();
      pos++;
    };

    const parseAtom = (): ASTNode => {
      const token = tokens[pos];
      if (!token) throw new Unsupported();
      if (token.type === 'NUMBER') { pos++; return { type: 'NUMBER', value: parseFloat(token.value) }; }
      if (token.type === 'STRING') { pos++; return { type: 'STRING', value: token.value }; }
      if (token.type === 'CELL') { pos++; return { type: 'CELL', value: token.value }; }
      if (token.type === 'RANGE') { pos++; return { type: 'RANGE', value: token.value }; }
      if (token.type === 'IDENTIFIER') {
        const name = token.value;
        pos++;
        if (tokens[pos]?.type === 'PAREN_L') {
          pos++;
          const args: ASTNode[] = [];
          if (tokens[pos]?.type !== 'PAREN_R') {
            args.push(parseExpression());
            while (tokens[pos]?.type === 'COMMA') {
              pos++;
              args.push(parseExpression());
            }
          }
          expect('PAREN_R');
          return { type: 'CALL', name, args };
        }
        if (name === 'TRUE') return { type: 'BOOL', value: true };
        if (name === 'FALSE') return { type: 'BOOL', value: false };
        throw new Unsupported(); // defined names
      }
      if (token.type === 'PAREN_L') {
        pos++;
        const node = parseExpression();
        expect('PAREN_R');
        return node;
      }
      throw new Unsupported();
    };

    // Unary minus binds tighter than ^ in Excel (-2^2 is 4); % divides by 100
    const parseUnary = (): ASTNode => {
      if (isOp('-')) { pos++; return { type: 'NEG', left: parseUnary() }; }
      if (isOp('+')) { pos++; return parseUnary(); }
      let node = parseAtom();
      while (isOp('%')) { pos++; node = { type: 'PERCENT', left: node }; }
      return node;
    };

    const binary = (next: () => ASTNode, ...ops: string[]) => (): ASTNode => {
      let node = next();
      while (isOp(...ops)) {
        const operator = tokens[pos++].value;
        node = { type: 'BINARY', operator, left: node, right: next() };
      }
      return node;
    };
    // Excel's ^ is left-associative: 2^3^2 is 64
    const parsePower = binary(parseUnary, '^');
    const parseFactor = binary(parsePower, '*', '/');
    const parseTerm = binary(parseFactor, '+', '-');
    const parseConcat = binary(parseTerm, '&');
    const parseExpression = binary(parseConcat, '=', '<>', '<', '>', '<=', '>=');

    const ast = parseExpression();
    if (pos !== tokens.length) throw new Unsupported();
    return ast;
  }

  // A single value: a range used where one value is expected is not supported (implicit intersection)
  private scalar(node: ASTNode): Value {
    const v = this.evaluateAst(node);
    if (Array.isArray(v)) throw new Unsupported();
    return v;
  }

  private evaluateAst(node: ASTNode): any {
    switch (node.type) {
      case 'NUMBER': case 'STRING': case 'BOOL': return node.value;
      case 'CELL': return this.lookup(node.value);
      case 'RANGE': return this.getRangeValues(node.value);
      case 'NEG': {
        const n = toNumber(this.scalar(node.left!));
        return n instanceof FormulaError ? n : -n;
      }
      case 'PERCENT': {
        const n = toNumber(this.scalar(node.left!));
        return n instanceof FormulaError ? n : n / 100;
      }
      case 'BINARY': return this.binary(node.operator!, this.scalar(node.left!), this.scalar(node.right!));
      case 'CALL': return this.call(node.name!, node.args!);
    }
    throw new Unsupported();
  }

  private binary(op: string, left: Value, right: Value): Value {
    if (left instanceof FormulaError) return left;
    if (right instanceof FormulaError) return right;
    if (op === '&') return toText(left) + toText(right);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const c = compare(left, right);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const l = toNumber(left), r = toNumber(right);
    if (l instanceof FormulaError) return l;
    if (r instanceof FormulaError) return r;
    switch (op) {
      case '+': return l + r;
      case '-': return l - r;
      case '*': return l * r;
      case '/': return r === 0 ? DIV0 : l / r;
      case '^': {
        const p = Math.pow(l, r);
        return isFinite(p) ? p : new FormulaError('#NUM!');
      }
    }
    throw new Unsupported();
  }

  private call(name: string, argNodes: ASTNode[]): Value {
    if (name === 'IF') {
      if (argNodes.length < 1 || argNodes.length > 3) throw new Unsupported();
      const cond = toBool(this.scalar(argNodes[0]));
      if (cond instanceof FormulaError) return cond;
      if (cond) return argNodes.length > 1 ? this.scalar(argNodes[1]) ?? 0 : true;
      return argNodes.length > 2 ? this.scalar(argNodes[2]) ?? 0 : false;
    }
    if (name === 'CONCATENATE') {
      let out = '';
      for (const a of argNodes) {
        const v = this.scalar(a);
        if (v instanceof FormulaError) return v;
        out += toText(v);
      }
      return out;
    }
    if (!['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].includes(name)) throw new Unsupported();
    // Values typed into the call count (TRUE as 1, "3" as 3); in referenced cells only numbers do
    const nums: number[] = [];
    for (const a of argNodes) {
      const v = this.evaluateAst(a);
      const referenced = a.type === 'CELL' || a.type === 'RANGE';
      for (const x of Array.isArray(v) ? v : [v]) {
        if (x instanceof FormulaError) {
          if (name === 'COUNT') continue;
          return x;
        }
        if (typeof x === 'number') nums.push(x);
        else if (!referenced && x !== null) {
          const n = toNumber(x);
          if (n instanceof FormulaError) {
            if (name === 'COUNT') continue;
            return n;
          }
          nums.push(n);
        }
      }
    }
    switch (name) {
      case 'SUM': return nums.reduce((a, b) => a + b, 0);
      case 'COUNT': return nums.length;
      case 'AVERAGE': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : DIV0;
      case 'MAX': return nums.length ? Math.max(...nums) : 0;
      default: return nums.length ? Math.min(...nums) : 0;
    }
  }

  private getRangeValues(range: string): Value[] {
    const [startRef, endRef] = range.split(':');
    const [c1, r1] = this.parseRef(startRef);
    const [c2, r2] = this.parseRef(endRef);
    const values: Value[] = [];
    // A3:A1 is the same range as A1:A3
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        values.push(this.lookup(this.toRef(r, c - 1)));
      }
    }
    return values;
  }

  private parseRef(ref: string): [number, number] {
    const match = ref.match(/^([A-Z]+)(\d+)$/i)!;
    const col = match[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
    const row = parseInt(match[2], 10);
    return [col, row];
  }

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

// Excel's conversions: blank is 0 / "" / FALSE, booleans are 1 and 0, text must read as a number
function toNumber(v: Value): number | FormulaError {
  if (v instanceof FormulaError) return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'number') return v;
  const t = v.trim();
  const n = t === '' ? NaN : Number(t);
  return isFinite(n) ? n : VALUE;
}

function toText(v: Value): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberText(v);
  return String(v);
}

function toBool(v: Value): boolean | FormulaError {
  if (v instanceof FormulaError) return v;
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const t = v.toUpperCase();
  return t === 'TRUE' ? true : t === 'FALSE' ? false : VALUE;
}

// Excel orders numbers < text < booleans; text compares without case. A blank takes the other side's type.
function compare(a: Value, b: Value): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (v: Value) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') {
    const x = a.toLowerCase(), y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = Number(a), y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}
