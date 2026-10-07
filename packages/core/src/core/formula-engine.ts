export type FormulaResult = number | string | boolean | null;

type TokenType = 'NUMBER' | 'STRING' | 'CELL' | 'RANGE' | 'IDENTIFIER' | 'OP' | 'PAREN_L' | 'PAREN_R' | 'COMMA';

interface Token {
  type: TokenType;
  value: string;
}

interface ASTNode {
  type: 'NUMBER' | 'STRING' | 'CELL' | 'RANGE' | 'CALL' | 'BINARY';
  value?: any;
  left?: ASTNode;
  right?: ASTNode;
  operator?: string;
  name?: string;
  args?: ASTNode[];
}

export class FormulaEngine {
  private cells: Map<string, FormulaResult> = new Map();

  clear() {
    this.cells.clear();
  }

  loadData(data: (FormulaResult)[][], startRow = 1) {
    data.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const ref = this.toRef(ri + startRow, ci);
        this.cells.set(ref, cell);
      });
    });
  }

  evaluate(formula: string): FormulaResult {
    try {
      const clean = formula.replace(/^=/, '').trim();
      if (!clean) return null;
      const tokens = this.tokenize(clean);
      const ast = this.parse(tokens);
      return this.evaluateAst(ast);
    } catch {
      // Unsupported syntax (other sheets, names, unknown functions): no cached value; Excel computes it on open
      return null;
    }
  }

  private tokenize(expr: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    while (i < expr.length) {
      let char = expr[i];
      if (/\s/.test(char)) { i++; continue; }
      if (char === '(') { tokens.push({ type: 'PAREN_L', value: '(' }); i++; continue; }
      if (char === ')') { tokens.push({ type: 'PAREN_R', value: ')' }); i++; continue; }
      if (char === ',') { tokens.push({ type: 'COMMA', value: ',' }); i++; continue; }
      if (/[+\-*/^<>=]/.test(char)) {
        let op = char;
        if (char === '<' || char === '>') {
          if (expr[i + 1] === '=') { op += '='; i++; }
          else if (char === '<' && expr[i + 1] === '>') { op += '>'; i++; }
        }
        tokens.push({ type: 'OP', value: op });
        i++;
        continue;
      }
      if (char === '"' || char === "'") {
        const quote = char;
        let str = '';
        i++;
        while (i < expr.length && expr[i] !== quote) { str += expr[i++]; }
        i++;
        tokens.push({ type: 'STRING', value: str });
        continue;
      }
      if (/[0-9.]/.test(char)) {
        let num = '';
        while (i < expr.length && /[0-9.]/.test(expr[i])) { num += expr[i++]; }
        tokens.push({ type: 'NUMBER', value: num });
        continue;
      }
      if (/[A-Za-z$]/.test(char)) {
        // "$" only pins a reference when copied; it does not change what it points at
        let id = '';
        while (i < expr.length && /[A-Za-z0-9$]/.test(expr[i])) { id += expr[i++]; }
        id = id.replace(/\$/g, '');
        if (i < expr.length && expr[i] === ':') {
          i++;
          let endId = '';
          while (i < expr.length && /[A-Za-z0-9$]/.test(expr[i])) { endId += expr[i++]; }
          endId = endId.replace(/\$/g, '');
          tokens.push({ type: 'RANGE', value: id.toUpperCase() + ':' + endId.toUpperCase() });
        } else if (/^[A-Z]+\d+$/i.test(id)) {
          tokens.push({ type: 'CELL', value: id.toUpperCase() });
        } else {
          tokens.push({ type: 'IDENTIFIER', value: id.toUpperCase() });
        }
        continue;
      }
      throw new Error(`Unknown character at ${i}: ${char}`);
    }
    return tokens;
  }

  private parse(tokens: Token[]): ASTNode {
    let pos = 0;

    const parsePrimary = (): ASTNode => {
      const token = tokens[pos];
      if (!token) throw new Error('Unexpected end of input');
      if (token.type === 'NUMBER') { pos++; return { type: 'NUMBER', value: parseFloat(token.value) }; }
      if (token.type === 'STRING') { pos++; return { type: 'STRING', value: token.value }; }
      if (token.type === 'CELL') { pos++; return { type: 'CELL', value: token.value }; }
      if (token.type === 'RANGE') { pos++; return { type: 'RANGE', value: token.value }; }
      if (token.type === 'IDENTIFIER') {
        const name = token.value;
        pos++;
        if (pos < tokens.length && tokens[pos].type === 'PAREN_L') {
          pos++; // skip '('
          const args: ASTNode[] = [];
          if (tokens[pos].type !== 'PAREN_R') {
            args.push(parseExpression());
            while (pos < tokens.length && tokens[pos].type === 'COMMA') {
              pos++;
              args.push(parseExpression());
            }
          }
          if (tokens[pos].type !== 'PAREN_R') throw new Error('Expected )');
          pos++;
          return { type: 'CALL', name, args };
        }
        // Handle true/false constants
        if (name === 'TRUE') return { type: 'NUMBER', value: true }; 
        if (name === 'FALSE') return { type: 'NUMBER', value: false };
        throw new Error(`Unknown identifier ${name}`);
      }
      if (token.type === 'PAREN_L') {
        pos++;
        const node = parseExpression();
        if (tokens[pos].type !== 'PAREN_R') throw new Error('Expected )');
        pos++;
        return node;
      }
      throw new Error(`Unexpected token ${token.value}`);
    };

    const parsePower = (): ASTNode => {
      let node = parsePrimary();
      while (pos < tokens.length && tokens[pos].value === '^') {
        const op = tokens[pos].value;
        pos++;
        node = { type: 'BINARY', operator: op, left: node, right: parsePrimary() };
      }
      return node;
    };

    const parseFactor = (): ASTNode => {
      let node = parsePower();
      while (pos < tokens.length && (tokens[pos].value === '*' || tokens[pos].value === '/')) {
        const op = tokens[pos].value;
        pos++;
        node = { type: 'BINARY', operator: op, left: node, right: parsePower() };
      }
      return node;
    };

    const parseTerm = (): ASTNode => {
      let node = parseFactor();
      while (pos < tokens.length && (tokens[pos].value === '+' || tokens[pos].value === '-')) {
        const op = tokens[pos].value;
        pos++;
        node = { type: 'BINARY', operator: op, left: node, right: parseFactor() };
      }
      return node;
    };

    const parseComparison = (): ASTNode => {
      let node = parseTerm();
      while (pos < tokens.length && ['=', '<>', '<', '>', '<=', '>='].includes(tokens[pos].value)) {
        const op = tokens[pos].value;
        pos++;
        node = { type: 'BINARY', operator: op, left: node, right: parseTerm() };
      }
      return node;
    };

    const parseExpression = (): ASTNode => {
      return parseComparison();
    };

    return parseExpression();
  }

  private evaluateAst(node: ASTNode): any {
    if (node.type === 'NUMBER') return node.value;
    if (node.type === 'STRING') return node.value;
    if (node.type === 'CELL') return this.cells.get(node.value) ?? 0;
    if (node.type === 'RANGE') return this.getRangeValues(node.value);
    
    if (node.type === 'BINARY') {
      const left = this.evaluateAst(node.left!);
      const right = this.evaluateAst(node.right!);
      
      const lNum = Number(left);
      const rNum = Number(right);

      switch (node.operator) {
        case '+': return lNum + rNum;
        case '-': return lNum - rNum;
        case '*': return lNum * rNum;
        case '/': return rNum === 0 ? '#DIV/0!' : lNum / rNum;
        case '^': return Math.pow(lNum, rNum);
        case '=': return left === right;
        case '<>': return left !== right;
        case '>': return lNum > rNum;
        case '<': return lNum < rNum;
        case '>=': return lNum >= rNum;
        case '<=': return lNum <= rNum;
      }
    }
    
    if (node.type === 'CALL') {
      const args = node.args!.map(a => this.evaluateAst(a));
      switch (node.name) {
        case 'SUM': return this.numbers(args).reduce((a, b) => a + b, 0);
        case 'AVERAGE': {
          const nums = this.numbers(args);
          return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : '#DIV/0!';
        }
        case 'COUNT': return this.numbers(args).length;
        case 'MAX': { const nums = this.numbers(args); return nums.length ? nums.reduce((a, b) => a > b ? a : b) : 0; }
        case 'MIN': { const nums = this.numbers(args); return nums.length ? nums.reduce((a, b) => a < b ? a : b) : 0; }
        case 'IF': return args[0] ? args[1] : args[2];
        case 'CONCATENATE': return this.flatten(args).join('');
      }
    }
    return null;
  }

  // Like Excel aggregates: only numeric values count; text, booleans and blanks are skipped.
  private numbers(args: any[]): number[] {
    return this.flatten(args).filter((v): v is number => typeof v === 'number' && !isNaN(v));
  }

  private flatten(arr: any[]): any[] {
    return arr.flat(Infinity);
  }

  private getRangeValues(range: string): any[] {
    const [startRef, endRef] = range.split(':');
    const [startCol, startRow] = this.parseRef(startRef);
    const [endCol, endRow] = this.parseRef(endRef);
    const values: any[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const ref = this.toRef(r, c - 1);
        values.push(this.cells.get(ref) ?? null);
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
