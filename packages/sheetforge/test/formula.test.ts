import { describe, it, expect } from 'vitest';
import { FormulaEngine } from '../src/core/formula-engine';

describe('AST Formula Engine', () => {
  it('should evaluate basic arithmetic', () => {
    const engine = new FormulaEngine();
    expect(engine.evaluate('1+2')).toBe(3);
    expect(engine.evaluate('10-5*2')).toBe(0);
    expect(engine.evaluate('(10-5)*2')).toBe(10);
    expect(engine.evaluate('2^3')).toBe(8);
  });

  it('treats absolute references like relative ones and stays quiet on unsupported syntax', () => {
    const engine = new FormulaEngine();
    engine.loadData([[2, 3], [4, 5]]);
    expect(engine.evaluate('$A$1*B$2+SUM($A1:B$1)')).toBe(15);
    expect(engine.evaluate('Other!A1')).toBeNull();
  });

  it('should evaluate boolean logic', () => {
    const engine = new FormulaEngine();
    expect(engine.evaluate('5>3')).toBe(true);
    expect(engine.evaluate('5=5')).toBe(true);
    expect(engine.evaluate('5<>5')).toBe(false);
  });

  it('should evaluate cell references', () => {
    const engine = new FormulaEngine();
    engine.loadData([
      [10, 20, 30],
      [5, 5, 5]
    ]);
    
    expect(engine.evaluate('A1+B1')).toBe(30);
    expect(engine.evaluate('A1*A2')).toBe(50);
  });

  it('should evaluate standard functions', () => {
    const engine = new FormulaEngine();
    engine.loadData([
      [10, 20, 30],
      [5, 5, 5]
    ]);

    expect(engine.evaluate('SUM(A1:C1)')).toBe(60);
    expect(engine.evaluate('AVERAGE(A1:C1)')).toBe(20);
    expect(engine.evaluate('MAX(A1:C1, 100)')).toBe(100);
    expect(engine.evaluate('MIN(A1:C2)')).toBe(5);
    expect(engine.evaluate('COUNT(A1:C2)')).toBe(6);
  });

  it('should evaluate nested complex formulas (IF, SUM)', () => {
    const engine = new FormulaEngine();
    engine.loadData([
      [10, 20, 30], // A1:C1
      [5, 5, 5]     // A2:C2
    ]);

    // IF(30 > 15, 60, 0) -> 60
    expect(engine.evaluate('IF(C1 > 15, SUM(A1:C1), 0)')).toBe(60);
    
    // Nested math inside function
    expect(engine.evaluate('SUM(A1:B1) * 2')).toBe(60);
  });
});
