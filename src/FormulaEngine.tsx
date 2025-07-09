import { Graph, alg } from 'graphlib';

export interface CellReference {
  sheet: string;
  col: number;
  row: number;
}

export interface FormulaResult {
  value: string | number;
  error?: string;
  dependencies: Set<string>;
}

export class FormulaEngine {
  private dependencyGraph: Graph;
  private cellValues: Map<string, any>;
  private formulaCache: Map<string, FormulaResult>;

  constructor() {
    this.dependencyGraph = new Graph({ directed: true });
    this.cellValues = new Map();
    this.formulaCache = new Map();
  }

  // Convert cell coordinates to A1 notation
  private coordsToA1(col: number, row: number, sheet: string = 'Sheet1'): string {
    let columnName = "";
    let dividend = col + 1;
    let modulo;

    while (dividend > 0) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    return `${sheet}!${columnName}${row + 1}`;
  }

  // Parse A1 notation to coordinates
  private a1ToCoords(a1: string): CellReference | null {
    const match = a1.match(/^(?:([^!]+)!)?([A-Z]+)(\d+)$/i);
    if (!match) return null;

    const [, sheet = 'Sheet1', colStr, rowStr] = match;
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 65 + 1);
    }
    col -= 1;

    const row = parseInt(rowStr) - 1;
    return { sheet, col, row };
  }

  // Parse formula to extract cell references
  private parseReferences(formula: string): Set<string> {
    const references = new Set<string>();
    const regex = /(?:([^!]+)!)?([A-Z]+\d+(?::[A-Z]+\d+)?)/gi;
    let match;

    while ((match = regex.exec(formula)) !== null) {
      const [, sheet = 'Sheet1', ref] = match;
      
      if (ref.includes(':')) {
        // Handle range references
        const [start, end] = ref.split(':');
        const startCoords = this.a1ToCoords(`${sheet}!${start}`);
        const endCoords = this.a1ToCoords(`${sheet}!${end}`);
        
        if (startCoords && endCoords) {
          for (let r = Math.min(startCoords.row, endCoords.row); r <= Math.max(startCoords.row, endCoords.row); r++) {
            for (let c = Math.min(startCoords.col, endCoords.col); c <= Math.max(startCoords.col, endCoords.col); c++) {
              references.add(this.coordsToA1(c, r, sheet));
            }
          }
        }
      } else {
        // Handle single cell references
        references.add(`${sheet}!${ref}`);
      }
    }

    return references;
  }

  // Update cell value and manage dependencies
  updateCell(col: number, row: number, value: any, sheet: string = 'Sheet1'): void {
    const cellA1 = this.coordsToA1(col, row, sheet);
    
    // Remove old dependencies
    if (this.dependencyGraph.hasNode(cellA1)) {
      const oldDependencies = this.dependencyGraph.predecessors(cellA1) || [];
      oldDependencies.forEach(dep => {
        this.dependencyGraph.removeEdge(dep, cellA1);
      });
    }

    // Clear cache for this cell and its dependents
    this.invalidateCache(cellA1);

    // Set the new value
    this.cellValues.set(cellA1, value);
    
    // If it's a formula, parse dependencies
    if (typeof value === 'string' && value.startsWith('=')) {
      const references = this.parseReferences(value);
      
      // Add nodes and edges to dependency graph
      this.dependencyGraph.setNode(cellA1);
      references.forEach(ref => {
        this.dependencyGraph.setNode(ref);
        this.dependencyGraph.setEdge(ref, cellA1);
      });

      // Check for circular dependencies
      const cycles = alg.findCycles(this.dependencyGraph);
      if (cycles.length > 0) {
        // Remove the problematic edges and throw error
        references.forEach(ref => {
          if (this.dependencyGraph.hasEdge(ref, cellA1)) {
            this.dependencyGraph.removeEdge(ref, cellA1);
          }
        });
        throw new Error(`Circular dependency detected: ${cycles.map(cycle => cycle.join(' → ')).join(', ')}`);
      }
    } else {
      // For non-formula values, just ensure the node exists
      this.dependencyGraph.setNode(cellA1);
    }
  }

  // Get cell value with formula evaluation
  getCellValue(col: number, row: number, sheet: string = 'Sheet1'): any {
    const cellA1 = this.coordsToA1(col, row, sheet);
    return this.evaluateCell(cellA1);
  }

  // Evaluate a single cell
  private evaluateCell(cellA1: string): any {
    // Check cache first
    if (this.formulaCache.has(cellA1)) {
      const cached = this.formulaCache.get(cellA1)!;
      if (cached.error) {
        return cached.error;
      }
      return cached.value;
    }

    const rawValue = this.cellValues.get(cellA1);
    
    if (!rawValue || typeof rawValue !== 'string' || !rawValue.startsWith('=')) {
      // Not a formula, return as-is
      const result: FormulaResult = {
        value: rawValue || '',
        dependencies: new Set()
      };
      this.formulaCache.set(cellA1, result);
      return result.value;
    }

    // It's a formula, evaluate it
    try {
      const result = this.evaluateFormula(rawValue, cellA1);
      this.formulaCache.set(cellA1, result);
      return result.error || result.value;
    } catch (error) {
      const errorResult: FormulaResult = {
        value: '',
        error: `#ERROR: ${error instanceof Error ? error.message : 'Unknown error'}`,
        dependencies: new Set()
      };
      this.formulaCache.set(cellA1, errorResult);
      return errorResult.error;
    }
  }

  // Evaluate formula string
  private evaluateFormula(formula: string, currentCell: string): FormulaResult {
    const dependencies = this.parseReferences(formula);
    
    // Remove the = sign
    const expression = formula.substring(1).trim();
    
    // Handle different formula types
    const funcMatch = expression.match(/^(\w+)\(([^)]*)\)$/i);
    if (funcMatch) {
      return this.evaluateFunction(funcMatch[1], funcMatch[2], dependencies);
    }

    // Handle simple cell references
    const cellRef = this.a1ToCoords(expression);
    if (cellRef) {
      const refA1 = this.coordsToA1(cellRef.col, cellRef.row, cellRef.sheet);
      const value = this.evaluateCell(refA1);
      return {
        value: value,
        dependencies: new Set([refA1])
      };
    }

    // Handle arithmetic expressions
    return this.evaluateArithmetic(expression, dependencies);
  }

  // Evaluate function calls
  private evaluateFunction(funcName: string, args: string, dependencies: Set<string>): FormulaResult {
    const func = funcName.toUpperCase();
    const argList = args.split(',').map(arg => arg.trim()).filter(arg => arg !== '');
    
    switch (func) {
      case 'SUM':
        return this.evaluateSumFunction(argList, dependencies);
      case 'AVERAGE':
        return this.evaluateAverageFunction(argList, dependencies);
      case 'MIN':
        return this.evaluateMinFunction(argList, dependencies);
      case 'MAX':
        return this.evaluateMaxFunction(argList, dependencies);
      case 'COUNT':
        return this.evaluateCountFunction(argList, dependencies);
      case 'IF':
        return this.evaluateIfFunction(argList, dependencies);
      case 'VLOOKUP':
        return this.evaluateVlookupFunction(argList, dependencies);
      case 'CONCATENATE':
        return this.evaluateConcatenateFunction(argList, dependencies);
      case 'ROUND':
        return this.evaluateRoundFunction(argList, dependencies);
      case 'ABS':
        return this.evaluateAbsFunction(argList, dependencies);
      case 'SQRT':
        return this.evaluateSqrtFunction(argList, dependencies);
      case 'POWER':
        return this.evaluatePowerFunction(argList, dependencies);
      default:
        throw new Error(`Unknown function: ${func}`);
    }
  }

  // Helper function to get values from arguments
  private getValuesFromArgs(argList: string[], dependencies: Set<string>): number[] {
    const values: number[] = [];
    
    for (const arg of argList) {
      if (arg.includes(':')) {
        // Range reference
        const [start, end] = arg.split(':');
        const startCoords = this.a1ToCoords(start);
        const endCoords = this.a1ToCoords(end);
        
        if (startCoords && endCoords) {
          for (let r = Math.min(startCoords.row, endCoords.row); r <= Math.max(startCoords.row, endCoords.row); r++) {
            for (let c = Math.min(startCoords.col, endCoords.col); c <= Math.max(startCoords.col, endCoords.col); c++) {
              const cellA1 = this.coordsToA1(c, r, startCoords.sheet);
              dependencies.add(cellA1);
              const value = this.evaluateCell(cellA1);
              const num = parseFloat(String(value));
              if (!isNaN(num)) {
                values.push(num);
              }
            }
          }
        }
      } else {
        // Single cell or literal value
        const cellRef = this.a1ToCoords(arg);
        if (cellRef) {
          const cellA1 = this.coordsToA1(cellRef.col, cellRef.row, cellRef.sheet);
          dependencies.add(cellA1);
          const value = this.evaluateCell(cellA1);
          const num = parseFloat(String(value));
          if (!isNaN(num)) {
            values.push(num);
          }
        } else {
          // Literal number
          const num = parseFloat(arg);
          if (!isNaN(num)) {
            values.push(num);
          }
        }
      }
    }
    
    return values;
  }

  // Function implementations
  private evaluateSumFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    const values = this.getValuesFromArgs(argList, dependencies);
    const sum = values.reduce((acc, val) => acc + val, 0);
    return { value: sum, dependencies };
  }

  private evaluateAverageFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    const values = this.getValuesFromArgs(argList, dependencies);
    if (values.length === 0) {
      throw new Error('No valid numbers for AVERAGE');
    }
    const avg = values.reduce((acc, val) => acc + val, 0) / values.length;
    return { value: avg, dependencies };
  }

  private evaluateMinFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    const values = this.getValuesFromArgs(argList, dependencies);
    if (values.length === 0) {
      throw new Error('No valid numbers for MIN');
    }
    const min = Math.min(...values);
    return { value: min, dependencies };
  }

  private evaluateMaxFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    const values = this.getValuesFromArgs(argList, dependencies);
    if (values.length === 0) {
      throw new Error('No valid numbers for MAX');
    }
    const max = Math.max(...values);
    return { value: max, dependencies };
  }

  private evaluateCountFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    const values = this.getValuesFromArgs(argList, dependencies);
    return { value: values.length, dependencies };
  }

  private evaluateIfFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length !== 3) {
      throw new Error('IF function requires exactly 3 arguments');
    }

    const [condition, trueValue, falseValue] = argList;
    
    // Evaluate condition
    const conditionResult = this.evaluateExpression(condition, dependencies);
    const isTrue = Boolean(conditionResult);

    // Evaluate the appropriate branch
    const result = isTrue ? 
      this.evaluateExpression(trueValue, dependencies) : 
      this.evaluateExpression(falseValue, dependencies);

    return { value: result, dependencies };
  }

  private evaluateVlookupFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length < 3) {
      throw new Error('VLOOKUP requires at least 3 arguments');
    }

    const [lookupValue, tableArray, colIndex, exactMatch = 'FALSE'] = argList;
    
    // This is a simplified VLOOKUP implementation
    // In a real implementation, you'd need to handle the table array properly
    return { value: '#N/A', dependencies };
  }

  private evaluateConcatenateFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    let result = '';
    
    for (const arg of argList) {
      const value = this.evaluateExpression(arg, dependencies);
      result += String(value);
    }
    
    return { value: result, dependencies };
  }

  private evaluateRoundFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length !== 2) {
      throw new Error('ROUND function requires exactly 2 arguments');
    }

    const [numberArg, digitsArg] = argList;
    const number = parseFloat(String(this.evaluateExpression(numberArg, dependencies)));
    const digits = parseInt(String(this.evaluateExpression(digitsArg, dependencies)));

    if (isNaN(number) || isNaN(digits)) {
      throw new Error('Invalid arguments for ROUND');
    }

    const result = Math.round(number * Math.pow(10, digits)) / Math.pow(10, digits);
    return { value: result, dependencies };
  }

  private evaluateAbsFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length !== 1) {
      throw new Error('ABS function requires exactly 1 argument');
    }

    const value = parseFloat(String(this.evaluateExpression(argList[0], dependencies)));
    if (isNaN(value)) {
      throw new Error('Invalid argument for ABS');
    }

    return { value: Math.abs(value), dependencies };
  }

  private evaluateSqrtFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length !== 1) {
      throw new Error('SQRT function requires exactly 1 argument');
    }

    const value = parseFloat(String(this.evaluateExpression(argList[0], dependencies)));
    if (isNaN(value) || value < 0) {
      throw new Error('Invalid argument for SQRT');
    }

    return { value: Math.sqrt(value), dependencies };
  }

  private evaluatePowerFunction(argList: string[], dependencies: Set<string>): FormulaResult {
    if (argList.length !== 2) {
      throw new Error('POWER function requires exactly 2 arguments');
    }

    const [baseArg, expArg] = argList;
    const base = parseFloat(String(this.evaluateExpression(baseArg, dependencies)));
    const exponent = parseFloat(String(this.evaluateExpression(expArg, dependencies)));

    if (isNaN(base) || isNaN(exponent)) {
      throw new Error('Invalid arguments for POWER');
    }

    return { value: Math.pow(base, exponent), dependencies };
  }

  // Evaluate simple expressions (cell references, literals)
  private evaluateExpression(expr: string, dependencies: Set<string>): any {
    const trimmed = expr.trim();
    
    // Check if it's a cell reference
    const cellRef = this.a1ToCoords(trimmed);
    if (cellRef) {
      const cellA1 = this.coordsToA1(cellRef.col, cellRef.row, cellRef.sheet);
      dependencies.add(cellA1);
      return this.evaluateCell(cellA1);
    }

    // Check if it's a number
    const num = parseFloat(trimmed);
    if (!isNaN(num)) {
      return num;
    }

    // Check if it's a string literal (remove quotes)
    if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
      return trimmed.slice(1, -1);
    }

    // Return as string
    return trimmed;
  }

  // Evaluate arithmetic expressions
  private evaluateArithmetic(expression: string, dependencies: Set<string>): FormulaResult {
    // This is a simplified arithmetic evaluator
    // In a real implementation, you'd want a proper expression parser
    try {
      // Replace cell references with their values
      let processedExpr = expression;
      const cellRefs = expression.match(/[A-Z]+\d+/g) || [];
      
      for (const ref of cellRefs) {
        const cellRef = this.a1ToCoords(ref);
        if (cellRef) {
          const cellA1 = this.coordsToA1(cellRef.col, cellRef.row, cellRef.sheet);
          dependencies.add(cellA1);
          const value = this.evaluateCell(cellA1);
          processedExpr = processedExpr.replace(new RegExp(ref, 'g'), String(value));
        }
      }

      // Evaluate the expression (WARNING: Using eval is dangerous in production)
      // In a real implementation, use a proper expression parser
      const result = Function(`"use strict"; return (${processedExpr})`)();
      
      return { value: result, dependencies };
    } catch (error) {
      throw new Error(`Invalid arithmetic expression: ${expression}`);
    }
  }

  // Invalidate cache for a cell and all its dependents
  private invalidateCache(cellA1: string): void {
    const toInvalidate = new Set<string>([cellA1]);
    const queue = [cellA1];

    while (queue.length > 0) {
      const current = queue.shift()!;
      const dependents = this.dependencyGraph.successors(current) || [];
      
      for (const dependent of dependents) {
        if (!toInvalidate.has(dependent)) {
          toInvalidate.add(dependent);
          queue.push(dependent);
        }
      }
    }

    // Clear cache for all affected cells
    toInvalidate.forEach(cell => {
      this.formulaCache.delete(cell);
    });
  }

  // Get all cells that depend on a given cell
  getDependents(col: number, row: number, sheet: string = 'Sheet1'): string[] {
    const cellA1 = this.coordsToA1(col, row, sheet);
    return this.dependencyGraph.successors(cellA1) || [];
  }

  // Get all cells that a given cell depends on
  getDependencies(col: number, row: number, sheet: string = 'Sheet1'): string[] {
    const cellA1 = this.coordsToA1(col, row, sheet);
    return this.dependencyGraph.predecessors(cellA1) || [];
  }

  // Recalculate all formulas that depend on changed cells
  recalculate(changedCells: string[]): void {
    const toRecalculate = new Set<string>();
    
    // Find all cells that need recalculation
    for (const cellA1 of changedCells) {
      const dependents = this.dependencyGraph.successors(cellA1) || [];
      dependents.forEach(dep => toRecalculate.add(dep));
    }

    // Sort cells by dependency order (topological sort)
    const sortedCells = alg.topsort(this.dependencyGraph);
    
    // Recalculate in dependency order
    for (const cellA1 of sortedCells) {
      if (toRecalculate.has(cellA1)) {
        this.formulaCache.delete(cellA1);
        this.evaluateCell(cellA1);
      }
    }
  }

  // Clear all data
  clear(): void {
    this.dependencyGraph = new Graph({ directed: true });
    this.cellValues.clear();
    this.formulaCache.clear();
  }

  // Export dependency graph for debugging
  exportDependencyGraph(): any {
    return {
      nodes: this.dependencyGraph.nodes(),
      edges: this.dependencyGraph.edges().map(edge => ({
        from: edge.v,
        to: edge.w
      }))
    };
  }
}