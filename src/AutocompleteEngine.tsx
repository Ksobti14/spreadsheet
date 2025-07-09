
import Fuse from 'fuse.js';

export interface AutocompleteSuggestion {
  value: string;
  score: number;
  type: 'existing' | 'function' | 'reference';
  description?: string;
}

export class AutocompleteEngine {
  private columnData: Map<number, Set<string>>;
  private functionList: { name: string; description: string; syntax: string }[];
  // Correctly type Fuse to work with objects of type { value: string }
  private fuse: Fuse<{ value: string }>; 

  constructor() {
    this.columnData = new Map();
    this.functionList = [
      { name: 'SUM', description: 'Adds all numbers in a range', syntax: 'SUM(range)' },
      { name: 'AVERAGE', description: 'Calculates the average of numbers', syntax: 'AVERAGE(range)' },
      { name: 'MIN', description: 'Finds the minimum value', syntax: 'MIN(range)' },
      { name: 'MAX', description: 'Finds the maximum value', syntax: 'MAX(range)' },
      { name: 'COUNT', description: 'Counts the number of cells with numbers', syntax: 'COUNT(range)' },
      { name: 'IF', description: 'Returns one value if true, another if false', syntax: 'IF(condition, true_value, false_value)' },
      { name: 'VLOOKUP', description: 'Looks up a value in a table', syntax: 'VLOOKUP(lookup_value, table_array, col_index, exact_match)' },
      { name: 'CONCATENATE', description: 'Joins text strings', syntax: 'CONCATENATE(text1, text2, ...)' },
      { name: 'ROUND', description: 'Rounds a number to specified digits', syntax: 'ROUND(number, digits)' },
      { name: 'ABS', description: 'Returns absolute value', syntax: 'ABS(number)' },
      { name: 'SQRT', description: 'Returns square root', syntax: 'SQRT(number)' }, // Fix: syntax instead of বেকার
      { name: 'POWER', description: 'Raises a number to a power', syntax: 'POWER(base, exponent)' },
    ];

    // Initialize Fuse for fuzzy searching, with 'value' as the key
    this.fuse = new Fuse<{ value: string }>([], { 
      threshold: 0.4,
      includeScore: true,
      keys: ['value'] // Search within the 'value' property of the objects
    });
  }

  // Update column data when cells change
  updateColumnData(col: number, values: string[]): void {
    const uniqueValues = new Set(values.filter(v => v && v.trim() !== ''));
    this.columnData.set(col, uniqueValues);
  }

  // Add a single value to a column
  addToColumn(col: number, value: string): void {
    if (!value || value.trim() === '') return;
    
    if (!this.columnData.has(col)) {
      this.columnData.set(col, new Set());
    }
    
    this.columnData.get(col)!.add(value.trim());
  }

  // Remove a value from a column
  removeFromColumn(col: number, value: string): void {
    if (!value || value.trim() === '') return;
    
    const columnSet = this.columnData.get(col);
    if (columnSet) {
      columnSet.delete(value.trim());
    }
  }

  // Get autocomplete suggestions for a column
  getColumnSuggestions(col: number, input: string, maxSuggestions: number = 10): AutocompleteSuggestion[] {
    const columnValues = this.columnData.get(col);
    if (!columnValues || columnValues.size === 0) {
      return [];
    }

    const inputLower = input.toLowerCase().trim();
    if (inputLower === '') {
      // Return all values if no input
      return Array.from(columnValues)
        .slice(0, maxSuggestions)
        .map(value => ({
          value,
          score: 1,
          type: 'existing' as const
        }));
    }

    // Use fuzzy search for better matching
    const searchData = Array.from(columnValues).map(value => ({ value }));
    this.fuse.setCollection(searchData); // Now setCollection expects { value: string }[]

    const results = this.fuse.search(inputLower);
    
    return results
      .slice(0, maxSuggestions)
      .map(result => ({
        // result.item is now correctly typed as { value: string }
        value: result.item.value, 
        score: 1 - (result.score || 0),
        type: 'existing' as const
      }));
  }

  // Get function suggestions
  getFunctionSuggestions(input: string, maxSuggestions: number = 10): AutocompleteSuggestion[] {
    const inputUpper = input.toUpperCase().trim();
    
    if (inputUpper === '' || !inputUpper.startsWith('=')) {
      return [];
    }

    const functionInput = inputUpper.substring(1); // Remove =
    
    return this.functionList
      .filter(func => func.name.startsWith(functionInput))
      .slice(0, maxSuggestions)
      .map(func => ({
        value: `=${func.name}(`,
        score: func.name.startsWith(functionInput) ? 1 : 0.8,
        type: 'function' as const,
        description: `${func.description} - ${func.syntax}`
      }));
  }

  // Get cell reference suggestions
  getCellReferenceSuggestions(input: string, maxRows: number = 1000, maxCols: number = 26): AutocompleteSuggestion[] {
    const inputUpper = input.toUpperCase().trim();
    
    // Check if input looks like a cell reference
    const cellRefPattern = /^[A-Z]*\d*$/;
    if (!cellRefPattern.test(inputUpper)) {
      return [];
    }

    const suggestions: AutocompleteSuggestion[] = [];
    
    // Generate column suggestions
    if (/^[A-Z]*$/.test(inputUpper)) {
      // Only letters, suggest columns
      for (let col = 0; col < maxCols; col++) {
        const colName = this.getColumnName(col);
        if (colName.startsWith(inputUpper)) {
          suggestions.push({
            value: colName,
            score: colName === inputUpper ? 1 : 0.8,
            type: 'reference'
          });
        }
      }
    }
    
    // Generate row suggestions
    const colMatch = inputUpper.match(/^([A-Z]+)(\d*)$/);
    if (colMatch) {
      const [, colPart, rowPart] = colMatch;
      
      if (rowPart === '') {
        // Suggest some common row numbers
        for (let row = 1; row <= Math.min(20, maxRows); row++) {
          suggestions.push({
            value: `${colPart}${row}`,
            score: 0.9,
            type: 'reference'
          });
        }
      } else {
        // Complete the row number
        const rowNum = parseInt(rowPart);
        if (!isNaN(rowNum) && rowNum <= maxRows) {
          suggestions.push({
            value: `${colPart}${rowNum}`,
            score: 1,
            type: 'reference'
          });
        }
      }
    }
    
    return suggestions.slice(0, 10);
  }

  // Get comprehensive suggestions based on context
  getSuggestions(input: string, col?: number, context: 'cell' | 'formula' = 'cell'): AutocompleteSuggestion[] {
    const suggestions: AutocompleteSuggestion[] = [];
    
    if (context === 'formula' || input.startsWith('=')) {
      // Formula context - suggest functions and cell references
      suggestions.push(...this.getFunctionSuggestions(input));
      
      // Extract the part after the last operator or comma for cell reference suggestions
      const lastPart = this.extractLastToken(input);
      if (lastPart && !lastPart.startsWith('=')) {
        suggestions.push(...this.getCellReferenceSuggestions(lastPart));
      }
    } else if (col !== undefined) {
      // Cell context - suggest column values
      suggestions.push(...this.getColumnSuggestions(col, input));
    }
    
    // Sort by score (descending)
    return suggestions.sort((a, b) => b.score - a.score);
  }

  // Extract the last token from a formula for context-aware suggestions
  private extractLastToken(input: string): string {
    const operators = ['+', '-', '*', '/', '(', ')', ',', ':', ';'];
    let lastIndex = -1;
    
    for (const op of operators) {
      const index = input.lastIndexOf(op);
      if (index > lastIndex) {
        lastIndex = index;
      }
    }
    
    return lastIndex >= 0 ? input.substring(lastIndex + 1).trim() : input;
  }

  // Convert column index to Excel-style column name
  private getColumnName(colIndex: number): string {
    let columnName = "";
    let dividend = colIndex + 1;
    let modulo;

    while (dividend > 0) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
  }

  // Clear all column data
  clear(): void {
    this.columnData.clear();
  }

  // Get statistics about column data
  getColumnStats(col: number): { uniqueValues: number; totalValues: number } {
    const columnSet = this.columnData.get(col);
    return {
      uniqueValues: columnSet ? columnSet.size : 0,
      totalValues: columnSet ? columnSet.size : 0 // In this implementation, they're the same
    };
  }

  // Export column data for debugging
  exportColumnData(): Record<number, string[]> {
    const result: Record<number, string[]> = {};
    this.columnData.forEach((values, col) => {
      result[col] = Array.from(values);
    });
    return result;
  }
}
