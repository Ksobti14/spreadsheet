import React, { useCallback, useRef, useState, useEffect, useMemo } from "react";
import DataEditor, {
  GridCell,
  GridCellKind,
  GridColumn,
  Item,
  EditableGridCell,
  GridSelection,
  CompactSelection,
  Rectangle,
  FillPatternEventArgs,
} from "@glideapps/glide-data-grid";
import "@glideapps/glide-data-grid/dist/index.css";
import * as FormulaJS from "formulajs";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { io } from "socket.io-client";
import { Graph, alg } from 'graphlib';
import Fuse from "fuse.js";
import {FormulaEngine} from './FormulaEngine';
const NUM_ROWS = 7000;
const NUM_COLUMNS = 100;

const socket = io("http://localhost:5000", {
  transports: ["websocket", "polling"],
  withCredentials: true
});

const getExcelColumnName = (colIndex: number): string => {
  let columnName = "";
  let dividend = colIndex + 1;
  let modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
};

const getCellName = (col: number, row: number) =>
  `${getExcelColumnName(col)}${row + 1}`;

const parseCellName = (name: string): [number, number] | null => {
  const match = name.match(/^([A-Z]+)([0-9]+)$/i);
  if (!match) return null;
  const [, colStr, rowStr] = match;
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 65 + 1);
  }
  col -= 1;

  const row = parseInt(rowStr) - 1;
  return [col, row];
};

type CellData = {
  value: string;
  formula?: string;
  fontSize?: number;
  alignment?: "left" | "center" | "right";
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  textColor?: string;
  bgColor?: string;
  borderColor?: string;
  fontFamily?: string;
  background?: string;
  displayData?: string | number;
  data?: string | number;
  comment?: string;
  link?: string;
  dataValidation?: {
    type: 'number' | 'text' | 'date' | 'list';
    operator?: 'greaterThan' | 'lessThan' | 'equalTo' | 'notEqualTo' | 'between' | 'textContains' | 'startsWith' | 'endsWith';
    value1?: string | number;
    value2?: string | number;
    sourceRange?: Rectangle;
  };
};

type SheetData = {
  [sheetName: string]: CellData[][];
};

type ConditionalFormattingRule = {
  id: string;
  range: Rectangle;
  type: 'greaterThan' | 'lessThan' | 'equalTo' | 'textContains' | 'between';
  value1: string | number;
  value2?: string | number;
  style: { bgColor?: string; textColor?: string };
};

type NamedRange = {
  id: string;
  name: string;
  range: Rectangle;
};

const createInitialSheetData = (): CellData[][] =>
  Array.from({ length: NUM_ROWS }, () =>
    Array.from({ length: NUM_COLUMNS }, () => ({ value: "" }))
  );

const getCellNumericValue = (
  col: number,
  row: number,
  data: CellData[][]
): number => {
  const cell = data[row]?.[col];
  if (!cell) return 0;
  const num = parseFloat(cell.value);
  return isNaN(num) ? 0 : num;
};
const excelCoordsToA1 = (col: number, row: number): string => {
  return `${getExcelColumnName(col)}${row + 1}`;
};

const a1ToExcelCoords = (a1: string): [number, number] | null => {
  const match = a1.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  let col = 0;
  const colStr = match[1];
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 65 + 1);
  }
  col--;

  const row = parseInt(match[2], 10) - 1;

  return [col, row];
};
const getCellRangeValues = (
  start: [number, number],
  end: [number, number],
  data: CellData[][]
): number[] => {
  const [startCol, startRow] = start;
  const [endCol, endRow] = end;
  const values: number[] = [];
  for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
    for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
      values.push(getCellNumericValue(c, r, data));
    }
  }
  return values;
};

const parseArg = (arg: string, data: CellData[][], namedRanges: NamedRange[]): number[] => {
  const namedRangeMatch = namedRanges.find(nr => nr.name.toUpperCase() === arg.toUpperCase());
  if (namedRangeMatch) {
    const { x, y, width, height } = namedRangeMatch.range;
    return getCellRangeValues([x, y], [x + width - 1, y + height - 1], data);
  }

  const rangeMatch = arg.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
  if (rangeMatch) {
    const start = parseCellName(rangeMatch[1]);
    const end = parseCellName(rangeMatch[2]);
    if (!start || !end) return [];
    return getCellRangeValues(start, end, data);
  }

  const cell = parseCellName(arg);
  if (cell) {
    const [col, row] = cell;
    return [getCellNumericValue(col, row, data)];
  }
  const num = parseFloat(arg);
  if (!isNaN(num)) return [num];
  if (arg.trim() === '') return [];
  return [];
};

const evaluateFormula = (formula: string, data: CellData[][], namedRanges: NamedRange[]): string => {
  const funcMatch = formula.match(/^=(\w+)\(([^)]*)\)$/i);
  if (!funcMatch) return formula;

  const [, func, args] = funcMatch;
  const funcUpper = func.toUpperCase();
  const parsedArgs = args.split(',').map(s => s.trim()).flatMap(arg => parseArg(arg, data, namedRanges));
  const error = (msg: string) => `#ERROR: ${msg}`;
  const argCount = args.split(',').map(s => s.trim()).filter(s => s !== '').length;

  switch (funcUpper) {
    case "SUM":
    case "AVERAGE":
    case "MIN":
    case "MAX":
    case "COUNT":
    case "PRODUCT":
      if (argCount < 1) return error("At least 1 argument required");
      break;
    case "IF":
      if (argCount !== 3) return error("IF requires 3 arguments");
      break;
    case "ROUND":
    case "POWER":
      if (argCount !== 2) return error("ROUND/POWER require 2 arguments");
      break;
    case "ABS":
    case "SQRT":
      if (argCount !== 1) return error("ABS/SQRT require 1 argument");
      else if (args.includes(":")) return error("ABS/SQRT require a single cell, not a range");
      break;
    default:
      return formula;
  }

  try {
    switch (funcUpper) {
      case "SUM":
        return FormulaJS.SUM(...parsedArgs).toString();
      case "AVERAGE":
        return parsedArgs.length > 0
          ? (FormulaJS.SUM(...parsedArgs) / parsedArgs.length).toString()
          : error("No valid numbers");
      case "MIN":
        return parsedArgs.length > 0 ? Math.min(...parsedArgs).toString() : error("No valid numbers");
      case "MAX":
        return parsedArgs.length > 0 ? Math.max(...parsedArgs).toString() : error("No valid numbers");
      case "COUNT":
        return parsedArgs.length.toString();
      case "PRODUCT":
        return parsedArgs.length > 0 ? FormulaJS.PRODUCT(...parsedArgs).toString() : "0";
      case "IF":
        const [cond, trueVal, falseVal] = args.split(',').map(s => s.trim());
        const condVal = parseArg(cond, data, namedRanges)[0] || 0;
        return condVal ? parseArg(trueVal, data, namedRanges)[0].toString() || trueVal : parseArg(falseVal, data, namedRanges)[0].toString() || falseVal;
      case "ROUND":
        const [numArg, digitsArg] = args.split(',').map(s => s.trim());
        const num = parseArg(numArg, data, namedRanges)[0] || 0;
        const digits = parseInt(digitsArg) || 0;
        return FormulaJS.ROUND(num, digits).toString();
      case "ABS":
        const absArg = parseArg(args, data, namedRanges)[0] || 0;
        return FormulaJS.ABS(absArg).toString();
      case "SQRT":
        const sqrtArg = parseArg(args, data, namedRanges)[0] || 0;
        return sqrtArg >= 0 ? FormulaJS.SQRT(sqrtArg).toString() : error("Negative number");
      case "POWER":
        const [baseArg, expArg] = args.split(',').map(s => s.trim());
        const base = parseArg(baseArg, data, namedRanges)[0] || 0;
        const exp = parseArg(expArg, data, namedRanges)[0] || 1;
        return FormulaJS.POWER(base, exp).toString();
      default:
        return formula;
    }
  } catch (e) {
    return error("Invalid formula");
  }
};

// Theme definitions
const lightTheme = {
  bg: '#ffffff',
  bg2: '#f8f9fa',
  text: '#202124',
  textLight: '#5f6368',
  border: '#dadce0',
  menuBg: '#fff',
  menuHoverBg: '#e6e6e6',
  activeTabBg: '#e8eaed',
  activeTabBorder: '#1a73e8',
  shadow: '0 1px 2px 0 rgba(60,64,67,0.08)',
  cellHighlightBg: '#E0F0FF',
  cellHighlightBorder: '#1E90FF',
};

const darkTheme = {
  bg: '#202124',
  bg2: '#3c4043',
  text: '#e8eaed',
  textLight: '#bdc1c6',
  border: '#5f6368',
  menuBg: '#3c4043',
  menuHoverBg: '#5f6368',
  activeTabBg: '#5f6368',
  activeTabBorder: '#8ab4f8',
  shadow: '0 2px 6px rgba(0,0,0,0.5)',
  cellHighlightBg: '#303030',
  cellHighlightBorder: '#8ab4f8',
};

const menuItem: React.CSSProperties = {
  padding: "8px 16px",
  textAlign: "left",
  border: "none",
  cursor: "pointer",
  width: "100%",
  fontSize: "13px",
  whiteSpace: "nowrap",
  display: "block",
};

const menuDropdownStyle: React.CSSProperties = {
  position: "absolute",
  top: "100%",
  left: 0,
  borderRadius: "4px",
  zIndex: 1000,
  minWidth: "160px",
  padding: "4px 0",
  display: "flex",
  flexDirection: "column",
};

const subMenuDropdownStyle: React.CSSProperties = {
  position: "absolute",
  top: "0",
  left: "100%",
  borderRadius: "4px",
  zIndex: 1001,
  minWidth: "160px",
  padding: "4px 0",
  display: "flex",
  flexDirection: "column",
};

const topBarButtonStyle: React.CSSProperties = {
  padding: "6px 10px",
  background: "transparent",
  border: "none",
  borderRadius: "4px",
  cursor: "pointer",
  fontSize: "13px",
  fontWeight: 500,
};

const formulaArgHints: { [key: string]: string[] } = {
  SUM: ["Number1", "[Number2]", "..."],
  AVERAGE: ["Number1", "[Number2]", "..."],
  MIN: ["Number1", "[Number2]", "..."],
  MAX: ["Number1", "[Number2]", "..."],
  COUNT: ["Number1", "[Number2]", "..."],
  PRODUCT: ["Number1", "[Number2]", "..."],
  IF: ["Condition", "Value if True", "Value if False"],
  ROUND: ["Number", "Digits"],
  POWER: ["Base", "Exponent"],
  ABS: ["Number"],
  SQRT: ["Number"],
};

const ContactGrid: React.FC = () => {
  const [showDropdown, setShowDropdown] = useState(false);
  const [selectedRanges, setSelectedRanges] = useState<Rectangle[]>([]);
  const [columnWidths, setColumnWidths] = useState<{ [key: number]: number }>({});
  const [clipboardData, setClipboardData] = useState<any[][] | null>(null);
  const gridData = useRef<Map<string, string | number>>(new Map()); // Actual grid data store
  
  const [data, setData] = useState<Map<string, string | number>>(new Map());
  // Enhanced context menu states
  const [contextMenuOpen, setContextMenuOpen] = useState(false);
  const [contextMenuPosition, setContextMenuPosition] = useState({ x: 0, y: 0 });
  const [rightClickedRow, setRightClickedRow] = useState<number | null>(null);
  const [contextMenuType, setContextMenuType] = useState<'row' | 'cell'>('cell');
  
  const [undoStack, setUndoStack] = useState<SheetData[]>([]);
  const [redoStack, setRedoStack] = useState<SheetData[]>([]);
const dependencyGraph = useRef(new Graph({ directed: true }));
  const [sheets, setSheets] = useState<SheetData>(() => {
    const initialSheetName = "Sheet1";
    return { [initialSheetName]: createInitialSheetData() };
  });
  const [activeSheet, setActiveSheet] = useState<string>("Sheet1");
  const [selection, setSelection] = useState<GridSelection>({
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  });
  const [formulaInput, setFormulaInput] = useState("");
  const [highlightRange, setHighlightRange] = useState<Rectangle | null>(null);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [formulaError, setFormulaError] = useState<string | null>(null);
  const [editingSheetName, setEditingSheetName] = useState<string | null>(null);
  const [newSheetName, setNewSheetName] = useState<string>("");
  const [saveLoadSheetName, setSaveLoadSheetName] = useState<string>('');
  const gridRef = useRef<any>(null);
 const formulaEngine = useRef(new FormulaEngine());
  const [sortColumnIndex, setSortColumnIndex] = useState<number | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc' | null>(null);
  const [showFilterRow, setShowFilterRow] = useState<boolean>(false);
  const [columnFilters, setColumnFilters] = useState<{ [key: number]: string }>({});

  const [conditionalFormattingRules, setConditionalFormattingRules] = useState<ConditionalFormattingRule[]>([]);
  const [showConditionalFormattingModal, setShowConditionalFormattingModal] = useState(false);
  const [cfType, setCfType] = useState<ConditionalFormattingRule['type']>('greaterThan');
  const [cfValue1, setCfValue1] = useState<string>('');
  const [cfValue2, setCfValue2] = useState<string>('');
  const [cfBgColor, setCfBgColor] = useState<string>('#FFFF00');
  const [cfTextColor, setCfTextColor] = useState<string>('#000000');

  const [frozenRows, setFrozenRows] = useState<number>(0);
  const [frozenColumns, setFrozenColumns] = useState<number>(0);

  const [dataUpdateKey, setDataUpdateKey] = useState(0);

  const [showFindModal, setShowFindModal] = useState(false);
  const [findSearchTerm, setFindSearchTerm] = useState('');
  const [findReplaceTerm, setFindReplaceTerm] = useState('');
  const [findCurrentMatch, setFindCurrentMatch] = useState<Item | null>(null);
  const [findMatches, setFindMatches] = useState<Item[]>([]);
  const [findMatchIndex, setFindMatchIndex] = useState(0);

  const [showDataValidationModal, setShowDataValidationModal] = useState(false);
  const [dvType, setDvType] = useState<'number' | 'text' | 'date' | 'list'>('number');
  const [dvOperator, setDvOperator] = useState<'greaterThan' | 'lessThan' | 'equalTo' | 'notEqualTo' | 'between' | 'textContains' | 'startsWith' | 'endsWith'>('greaterThan');
  const [dvValue1, setDvValue1] = useState<string>('');
  const [dvValue2, setDvValue2] = useState<string>('');
  const [dvSourceRange, setDvSourceRange] = useState<string>('');

  const [showNamedRangesModal, setShowNamedRangesModal] = useState(false);
  const [namedRanges, setNamedRanges] = useState<NamedRange[]>([]);
  const [newNamedRangeName, setNewNamedRangeName] = useState<string>('');
  const [newNamedRangeRef, setNewNamedRangeRef] = useState<string>('');
  const [editingNamedRangeId, setEditingNamedRangeId] = useState<string | null>(null);

  const [isDarkMode, setIsDarkMode] = useState(false);
  const currentTheme = isDarkMode ? darkTheme : lightTheme;

  // Formula guidance state
  const [currentFormulaFunction, setCurrentFormulaFunction] = useState<string | null>(null);
  const [currentFormulaArgIndex, setCurrentFormulaArgIndex] = useState<number | null>(null);
  const [showFormulaGuidance, setShowFormulaGuidance] = useState(false);

  const activeCell = useRef<Item | null>(null);
  const selecting = useRef<Rectangle | null>(null);
  const currentSheetData = sheets[activeSheet];

  const getDisplayedData = useMemo(() => {
    let dataWithOriginalIndex = currentSheetData.map((row, originalRowIndex) => ({
      originalRowIndex,
      data: row,
    }));

    const activeFilters = Object.entries(columnFilters).filter(([, value]) => value.trim() !== '');
    if (activeFilters.length > 0) {
      dataWithOriginalIndex = dataWithOriginalIndex.filter(item => {
        return activeFilters.every(([colIndexStr, filterValue]) => {
          const colIndex = parseInt(colIndexStr);
          const cellValue = item.data[colIndex]?.value?.toString().toLowerCase() || '';
          return cellValue.includes(filterValue.toLowerCase());
        });
      });
    }

    if (sortColumnIndex !== null && sortDirection !== null) {
      const sortedData = [...dataWithOriginalIndex];
      sortedData.sort((itemA, itemB) => {
        const valueA = itemA.data[sortColumnIndex]?.value || "";
        const valueB = itemB.data[sortColumnIndex]?.value || "";

        const numA = parseFloat(valueA);
        const numB = parseFloat(valueB);

        if (!isNaN(numA) && !isNaN(numB)) {
          return sortDirection === 'asc' ? numA - numB : numB - numA;
        } else {
          return sortDirection === 'asc'
            ? valueA.localeCompare(valueB)
            : valueB.localeCompare(valueA);
        }
      });
      while (sortedData.length < NUM_ROWS) {
        sortedData.push({ originalRowIndex: -1, data: Array(NUM_COLUMNS).fill({ value: "" }) });
      }
      return sortedData;
    }
    while (dataWithOriginalIndex.length < NUM_ROWS) {
      dataWithOriginalIndex.push({ originalRowIndex: -1, data: Array(NUM_COLUMNS).fill({ value: "" }) });
    }
    return dataWithOriginalIndex;
  }, [currentSheetData, sortColumnIndex, sortDirection, columnFilters]);

  const updateSuggestions = (inputValue: string) => {
    const suggestionsList = [
      "=SUM(",
      "=AVERAGE(",
      "=MIN(",
      "=MAX(",
      "=COUNT(",
      "=PRODUCT(",
      "=IF(",
      "=ROUND(",
      "=ABS(",
      "=SQRT(",
      "=POWER(",
    ];

    if (inputValue.startsWith("=")) {
      const filtered = suggestionsList.filter(s =>
        s.toLowerCase().startsWith(inputValue.toLowerCase())
      );
      setSuggestions(filtered);
      setShowSuggestions(filtered.length > 0);
    } else {
      setShowSuggestions(false);
    }
  };
const parseFormulaReferences = useCallback((formula: string): Set<string> => {
    const references = new Set<string>();
    const regex = /[A-Z]+\d+(?::[A-Z]+\d+)?/g;
    let match;
    while ((match = regex.exec(formula)) !== null) {
      const ref = match[0];
      if (ref.includes(':')) {
        const [start, end] = ref.split(':');
        const [startCol, startRow] = a1ToExcelCoords(start)!;
        const [endCol, endRow] = a1ToExcelCoords(end)!;

        for (let c = startCol; c <= endCol; c++) {
          for (let r = startRow; r <= endRow; r++) {
            references.add(excelCoordsToA1(c, r));
          }
        }
      } else {
        references.add(ref);
      }
    }
    return references;
  }, []);

  const evaluateCell = useCallback((cellA1: string): string | number => {
    const content = gridData.current.get(cellA1);
    if (typeof content === 'string' && content.startsWith('=')) {
      const formulaString = content.substring(1).trim().toUpperCase();
      const references = parseFormulaReferences(content); // Used for dependency tracking

      try {
        if (formulaString.startsWith('SUM(') && formulaString.endsWith(')')) {
          const rangeStr = formulaString.substring(4, formulaString.length - 1);
          const [start, end] = rangeStr.split(':');
          const [startCol, startRow] = a1ToExcelCoords(start)!;
          const [endCol, endRow] = a1ToExcelCoords(end)!;

          let sum = 0;
          for (let c = startCol; c <= endCol; c++) {
            for (let r = startRow; r <= endRow; r++) {
              const refA1 = excelCoordsToA1(c, r);
              const value = parseFloat(String(gridData.current.get(refA1) || 0));
              if (!isNaN(value)) {
                sum += value;
              }
            }
          }
          return sum;
        }
        return "#FORMULA!"; // Default for unhandled formulas
      } catch (e) {
        console.error("Error evaluating formula:", content, e);
        return "#ERROR!";
      }
    }
    return content || "";
  }, [parseFormulaReferences]);

  const recalculateGrid = useCallback((changedCellA1: string) => {
  const visited = new Set<string>();
  const queue = [changedCellA1];

  while (queue.length > 0) {
    const cell = queue.shift()!;
    if (visited.has(cell)) continue;
    visited.add(cell);

    const content = gridData.current.get(cell);
    if (typeof content === "string" && content.startsWith("=")) {
      const newValue = evaluateCell(cell);
      gridData.current.set(cell, newValue);
    }

    const dependents = dependencyGraph.current.successors(cell);
    if (dependents) queue.push(...dependents);
  }

  setData(new Map(gridData.current)); // Trigger UI update
}, [evaluateCell]);

  const getCellContent = useCallback(
    ([col, row]: Item): GridCell => {
      const displayedRowData = getDisplayedData[row];
      if (!displayedRowData) {
        return {
          kind: GridCellKind.Text,
          allowOverlay: true,
          readonly: false,
          displayData: "",
          data: "",
        };
      }
      const originalRowIndex = displayedRowData.originalRowIndex;
      if (originalRowIndex === -1) {
        return {
          kind: GridCellKind.Text,
          allowOverlay: true,
          readonly: false,
          displayData: "",
          data: "",
        };
      }
      const cell = displayedRowData.data[col] ?? { value: "" };
      let displayValue: string | number;
      if (cell.formula) {
        try {
          displayValue = formulaEngine.current.getCellValue(col, originalRowIndex, activeSheet);
        } catch (error) {
          displayValue = `#ERROR: ${error instanceof Error ? error.message : 'Unknown error'}`;
        }
      } else {
        displayValue = cell.value;
      }
      if (typeof displayValue === 'object' && displayValue !== null) {
        displayValue = JSON.stringify(displayValue);
      } else if (displayValue === undefined || displayValue === null) {
        displayValue = "";
      }

      const inHighlight =
        highlightRange &&
        col >= highlightRange.x &&
        col < highlightRange.x + highlightRange.width &&
        row >= highlightRange.y &&
        row < highlightRange.y + highlightRange.height;

      return {
        kind: GridCellKind.Text,
        allowOverlay: true,
        readonly: false,
        displayData: String(displayValue),
        data: cell.formula ?? cell.value,
        themeOverride: inHighlight
          ? { bgCell: currentTheme.cellHighlightBg, borderColor: currentTheme.cellHighlightBorder }
          : undefined,
      };
    },
    [getDisplayedData, highlightRange, currentTheme, currentSheetData, namedRanges]
  );

  const onCellEdited = useCallback(
    ([col, row]: Item, newValue: EditableGridCell) => {
      if (newValue.kind !== GridCellKind.Text) return;

      const text = newValue.data;
      const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
      if (originalRowIndex === undefined || originalRowIndex === -1) return;

      // Validate value if needed
      const cellToValidate = sheets[activeSheet][originalRowIndex]?.[col];
      const validationRule = cellToValidate?.dataValidation;
      if (validationRule) {
        let isValid = true;
        const numericValue = parseFloat(text);
        const lowerText = text.toLowerCase();

        switch (validationRule.type) {
          case 'number':
            if (isNaN(numericValue)) isValid = false;
            else if (validationRule.operator === 'greaterThan' && numericValue <= validationRule.value1!) isValid = false;
            else if (validationRule.operator === 'lessThan' && numericValue >= validationRule.value1!) isValid = false;
            else if (validationRule.operator === 'equalTo' && numericValue !== validationRule.value1!) isValid = false;
            else if (validationRule.operator === 'notEqualTo' && numericValue === validationRule.value1!) isValid = false;
            else if (validationRule.operator === 'between' && (numericValue < validationRule.value1! || numericValue > validationRule.value2!)) isValid = false;
            break;
          case 'text':
            if (validationRule.operator === 'textContains' && !lowerText.includes(String(validationRule.value1).toLowerCase())) isValid = false;
            else if (validationRule.operator === 'startsWith' && !lowerText.startsWith(String(validationRule.value1).toLowerCase())) isValid = false;
            else if (validationRule.operator === 'endsWith' && !lowerText.endsWith(String(validationRule.value1).toLowerCase())) isValid = false;
            break;
          case 'list':
            if (validationRule.sourceRange) {
              const { x, y, width, height } = validationRule.sourceRange;
              const allowedValues: string[] = [];
              for (let r = y; r < y + height; r++) {
                for (let c = x; c < x + width; c++) {
                  const sourceCell = sheets[activeSheet]?.[r]?.[c];
                  if (sourceCell) allowedValues.push(sourceCell.value);
                }
              }
              if (!allowedValues.includes(text)) {
                alert(`Invalid input. Choose from: ${allowedValues.join(', ')}`);
                return;
              }
            }
            break;
        }

        if (!isValid) {
          alert(`Invalid input: "${text}" does not meet validation.`);
          return;
        }
      }

      const updatedSheets = { ...sheets };
      const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

      try {
        // Update the formula engine
        formulaEngine.current.updateCell(col, originalRowIndex, text, activeSheet);

        // Update autocomplete engine
        if (text.startsWith("=")) {
          // It's a formula
          const evaluatedValue = formulaEngine.current.getCellValue(col, originalRowIndex, activeSheet);
          currentSheetCopy[originalRowIndex][col] = {
            ...currentSheetCopy[originalRowIndex][col],
            formula: text,
            value: String(evaluatedValue),
          };
        } else {
          // It's a regular value
          currentSheetCopy[originalRowIndex][col] = {
            ...currentSheetCopy[originalRowIndex][col],
            value: text,
            formula: undefined,
          };
          
          // Update autocomplete data
          
        }

        updatedSheets[activeSheet] = currentSheetCopy;
        pushToUndoStack(updatedSheets);
        setFormulaInput("");
        setHighlightRange(null);
        setShowSuggestions(false);
        setFormulaError(null);

        socket.emit("cell-edit", {
          sheet: activeSheet,
          row: originalRowIndex,
          col,
          value: text,
        });

        // Trigger recalculation of dependent cells
        const cellA1 = getCellName(col, originalRowIndex);
        const dependents = formulaEngine.current.getDependents(col, originalRowIndex, activeSheet);
        if (dependents.length > 0) {
          formulaEngine.current.recalculate([cellA1]);
          setDataUpdateKey(prev => prev + 1);
        }

      } catch (error) {
        if (error instanceof Error) {
          setFormulaError(error.message);
          alert(error.message);
        }
      }
    },
    [activeSheet, sheets, getDisplayedData]
  );

  const onFillPattern = useCallback(
    ({ patternSource, fillDestination }: FillPatternEventArgs) => {
      const sourceCol = patternSource.x;
      const sourceRow = patternSource.y;
      const originalSourceRowIndex = getDisplayedData[sourceRow]?.originalRowIndex;
      if (originalSourceRowIndex === undefined || originalSourceRowIndex === -1) return;

      const sourceCell = currentSheetData[originalSourceRowIndex]?.[sourceCol];
      if (!sourceCell) return;

      const fillValue = sourceCell.value;
      const fillFormula = sourceCell.formula;

      const updatedSheets = { ...sheets };
      const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

      for (let r = fillDestination.y; r < fillDestination.y + fillDestination.height; r++) {
        const originalDestRowIndex = getDisplayedData[r]?.originalRowIndex;
        if (originalDestRowIndex === undefined || originalDestRowIndex === -1) continue;

        for (let c = fillDestination.x; c < fillDestination.x + fillDestination.width; c++) {
          const isSourceCell = c >= patternSource.x && c < patternSource.x + patternSource.width &&
            r >= patternSource.y && r < patternSource.y + patternSource.height;
          if (!isSourceCell) {
            try {
              if (fillFormula) {
                formulaEngine.current.updateCell(c, originalDestRowIndex, fillFormula, activeSheet);
                const evaluatedValue = formulaEngine.current.getCellValue(c, originalDestRowIndex, activeSheet);
                currentSheetCopy[originalDestRowIndex][c] = { 
                  formula: fillFormula, 
                  value: String(evaluatedValue) 
                };
              } else {
                formulaEngine.current.updateCell(c, originalDestRowIndex, fillValue, activeSheet);
                currentSheetCopy[originalDestRowIndex][c] = { value: fillValue };
                
              }
            } catch (error) {
              console.error('Error during fill pattern:', error);
            }
          }
        }
      }

      updatedSheets[activeSheet] = currentSheetCopy;
      pushToUndoStack(updatedSheets);
    },
    [activeSheet, currentSheetData, getDisplayedData, sheets]
  );

  const onFinishSelecting = useCallback(() => {
    if (!activeCell.current || !selecting.current) return;
    const [col, row] = activeCell.current;

    const originalActiveRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalActiveRowIndex === undefined || originalActiveRowIndex === -1) return;

    const topLeft = getCellName(selecting.current.x, selecting.current.y);
    const bottomRight = getCellName(
      selecting.current.x + selecting.current.width - 1,
      selecting.current.y + selecting.current.height - 1
    );
    const formula = `=SUM(${topLeft}:${bottomRight})`;
    const value = evaluateFormula(formula, currentSheetData, namedRanges);

    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);
    currentSheetCopy[originalActiveRowIndex][col] = { formula, value };
    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);

    setFormulaInput("");
    setHighlightRange(null);
    setShowSuggestions(false);
    setFormulaError(null);
  }, [activeSheet, currentSheetData, getDisplayedData, namedRanges]);

  const handleFormulaChange = (val: string) => {
    setFormulaInput(val);
    updateSuggestions(val);
    setFormulaError(null);
    setHighlightRange(null);
    
    // Update formula guidance
    updateFormulaGuidance(val);

    if (val.startsWith("=")) {
      const funcMatch = val.match(/^=(\w+)\(([^)]*)\)?$/i);
      if (funcMatch) {
        const [, func, args] = funcMatch;
        const funcUpper = func.toUpperCase();
        const argCount = args.split(',').map(s => s.trim()).filter(s => s !== '').length;
        if (funcUpper === "ABS" || funcUpper === "SQRT") {
          if (argCount !== 1) {
            setFormulaError(`${funcUpper} requires 1 argument`);
          } else if (args.includes(":")) {
            setFormulaError(`${funcUpper} requires a single cell, not a range`);
          }
        } else if (funcUpper === "ROUND" || funcUpper === "POWER") {
          if (argCount !== 2) {
            setFormulaError(`${funcUpper} requires 2 arguments`);
          }
        } else if (funcUpper === "IF") {
          if (argCount !== 3) {
            setFormulaError(`${funcUpper} requires 3 arguments`);
          }
        } else if (
          funcUpper === "SUM" ||
          funcUpper === "AVERAGE" ||
          funcUpper === "MIN" ||
          funcUpper === "MAX" ||
          funcUpper === "COUNT" ||
          funcUpper === "PRODUCT"
        ) {
          if (argCount < 1) {
            setFormulaError(`${funcUpper} requires at least 1 argument`);
          }
        }
        const regex = /([A-Z]+\d+)(?:\s*:\s*([A-Z]+\d+))?/g;
        let match;
        while ((match = regex.exec(args)) !== null) {
          const start = parseCellName(match[1]);
          const end = match[2] ? parseCellName(match[2]) : start;
          if (start && end) {
            const [sx, sy] = start;
            const [ex, ey] = end;
            const x = Math.min(sx, ex);
            const y = Math.min(sy, ey);
            const width = Math.abs(sx - ex) + 1;
            const height = Math.abs(sy - ey) + 1;
            setHighlightRange({ x, y, width, height });

            setDataUpdateKey(prev => prev + 1);
            gridRef.current?.scrollTo(y, x);
            break;
          }
        }
      }
    }
  };

  // New function to update formula guidance
  const updateFormulaGuidance = (val: string) => {
    if (!val.startsWith("=")) {
      setShowFormulaGuidance(false);
      setCurrentFormulaFunction(null);
      setCurrentFormulaArgIndex(null);
      return;
    }

    const funcMatch = val.match(/^=(\w+)\(/i);
    if (funcMatch) {
      const fn = funcMatch[1].toUpperCase();
      if (formulaArgHints[fn]) {
        const argsPart = val.slice(funcMatch[0].length);
        const argIndex = argsPart.split(",").length - 1;
        
        setCurrentFormulaFunction(fn);
        setCurrentFormulaArgIndex(argIndex);
        setShowFormulaGuidance(true);
      } else {
        setShowFormulaGuidance(false);
      }
    } else {
      setShowFormulaGuidance(false);
    }
  };

  // Handle right-click context menu
  const handleContextMenu = (e: React.MouseEvent, row?: number) => {
    e.preventDefault();
    setContextMenuOpen(true);
    setContextMenuPosition({ x: e.clientX, y: e.clientY });
    if (row !== undefined) {
      setRightClickedRow(row);
    }
  };

  // Enhanced context menu handlers
  const handleRowContextMenu = (e: React.MouseEvent, rowIndex: number) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenuType('row');
    setRightClickedRow(rowIndex);
    setContextMenuPosition({ x: e.clientX, y: e.clientY });
    setContextMenuOpen(true);
  };

  const handleCellContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setContextMenuType('cell');
    setContextMenuPosition({ x: e.clientX, y: e.clientY });
    setContextMenuOpen(true);
  };

  // Close context menu when clicking elsewhere
  useEffect(() => {
    const handleClick = () => {
      setContextMenuOpen(false);
    };
    window.addEventListener("click", handleClick);
    return () => window.removeEventListener("click", handleClick);
  }, []);

  const handleAddSheet = () => {
    let newSheetName = `Sheet${Object.keys(sheets).length + 1}`;
    while (sheets[newSheetName]) {
      newSheetName = `Sheet${Math.floor(Math.random() * 10000)}`;
    }
    setSheets((prevSheets) => ({
      ...prevSheets,
      [newSheetName]: createInitialSheetData(),
    }));
    setActiveSheet(newSheetName);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleDeleteSheet = (sheetToDelete: string) => {
    if (Object.keys(sheets).length === 1) {
      alert("Cannot delete the last sheet!");
      return;
    }
    setSheets((prevSheets) => {
      const updatedSheets = { ...prevSheets };
      delete updatedSheets[sheetToDelete];
      return updatedSheets;
    });
    if (activeSheet === sheetToDelete) {
      setActiveSheet(Object.keys(sheets)[0]);
    }
    setDataUpdateKey(prev => prev + 1);
  };

  const handleEditSheetName = (sheetName: string) => {
    setEditingSheetName(sheetName);
    setNewSheetName(sheetName);
  };

  const handleSaveSheetName = (oldName: string) => {
    if (newSheetName.trim() === "" || newSheetName === oldName) {
      setEditingSheetName(null);
      return;
    }
    if (sheets[newSheetName]) {
      alert("Sheet name already exists!");
      return;
    }

    setSheets((prevSheets) => {
      const updatedSheets: SheetData = {};
      for (const key in prevSheets) {
        if (key === oldName) {
          updatedSheets[newSheetName] = prevSheets[key];
        } else {
          updatedSheets[key] = prevSheets[key];
        }
      }
      return updatedSheets;
    });
    if (activeSheet === oldName) {
      setActiveSheet(newSheetName);
    }
    setEditingSheetName(null);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleSheetNameInputKeyDown = (e: React.KeyboardEvent, oldName: string) => {
    if (e.key === 'Enter') {
      handleSaveSheetName(oldName);
    }
  };

  const BASE_BACKEND_URL = 'http://localhost:5000/api/sheets';

  const saveSheetData = async () => {
    if (!saveLoadSheetName.trim()) {
      alert('Please enter a name for your spreadsheet to save.');
      return;
    }
    try {
      const response = await fetch(BASE_BACKEND_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: saveLoadSheetName,
          worksheets: sheets,
        }),
      });

      if (response.ok) {
        const result = await response.json();
        alert(`Spreadsheet "${result.sheet.name}" saved successfully!`);
        setDataUpdateKey(prev => prev + 1);
      } else {
        const errorData = await response.json();
        alert(`Failed to save data: ${errorData.message || 'Unknown error'}`);
      }
    } catch (error) {
      alert('An error occurred while trying to save data. Check backend server.');
    }
  };

  const handleExportToJSON = () => {
    const fileData = JSON.stringify(sheets, null, 2);
    const blob = new Blob([fileData], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = `${saveLoadSheetName || "spreadsheet"}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    alert("Exported to JSON successfully!");
  };

  const handleImportFromJSON = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        const parsedSheets = JSON.parse(text);
        setSheets(parsedSheets);
        setActiveSheet(Object.keys(parsedSheets)[0] || "Sheet1");
        alert("Spreadsheet loaded from file successfully.");
        setDataUpdateKey(prev => prev + 1);
      } catch (err) {
        alert("Failed to import JSON file: Invalid format.");
      }
    };
    reader.readAsText(file);
  };

  const handleExportXLSX = () => {
    const data = sheets[activeSheet];
    if (!data) return;

    const worksheetData: any[][] = [];
    let maxRow = 0;
    let maxCol = 0;

    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      for (let c = 0; c < row.length; c++) {
        if (row[c]?.value?.toString().trim()) {
          maxRow = Math.max(maxRow, r);
          maxCol = Math.max(maxCol, c);
        }
      }
    }

    for (let r = 0; r <= maxRow; r++) {
      worksheetData[r] = [];
      for (let c = 0; c <= maxCol; c++) {
        const cell = data[r][c];
        if (cell?.formula) {
          try {
            worksheetData[r][c] = formulaEngine.current.getCellValue(c, r, activeSheet);
          } catch (error) {
            worksheetData[r][c] = cell.value || "";
          }
        } else {
          worksheetData[r][c] = cell?.value ?? "";
        }
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, saveLoadSheetName || activeSheet);

    const wbout = XLSX.write(wb, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true,
    });

    const blob = new Blob([wbout], {
      type: "application/octet-stream",
    });
    saveAs(blob, `${saveLoadSheetName || activeSheet}.xlsx`);
  };


  const importFromExcel = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const importedSheets: SheetData = {};

        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonSheet: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
          const newSheetData: CellData[][] = Array.from({ length: NUM_ROWS }, () =>
            Array.from({ length: NUM_COLUMNS }, () => ({ value: "" }))
          );

          jsonSheet.forEach((row, rowIndex) => {
            row.forEach((cellValue, colIndex) => {
              if (newSheetData[rowIndex] && newSheetData[rowIndex][colIndex]) {
                if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
                  newSheetData[rowIndex][colIndex] = { formula: cellValue, value: evaluateFormula(cellValue, newSheetData, namedRanges) };
                } else {
                  newSheetData[rowIndex][colIndex] = { value: String(cellValue) };
                }
              }
            });
          });
          importedSheets[sheetName] = newSheetData;
        });

        setSheets(importedSheets);
        setActiveSheet(Object.keys(importedSheets)[0] || "Sheet1");
        alert("Imported from Excel (.xlsx) successfully!");
        setDataUpdateKey(prev => prev + 1);
      } catch (err) {
        alert("Failed to import Excel file: Invalid format.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportToCSV = () => {
    const activeSheetData = sheets[activeSheet];
    if (!activeSheetData || activeSheetData.length === 0) {
      alert("No data in the current sheet to export to CSV.");
      return;
    }
    const csvContent = activeSheetData.map(row =>
      row.map(cell => {
        const displayValue = cell.formula ? evaluateFormula(cell.formula, activeSheetData, namedRanges) : cell.value;
        return `"${String(displayValue).replace(/"/g, '""')}"`;
      }).join(',')
    ).join('\n');

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    saveAs(blob, `${saveLoadSheetName || activeSheet}.csv`);
    alert("Exported to CSV successfully!");
  };

  const importFromCSV = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        const rows = text.split('\n').filter(row => row.trim() !== '');
        const newSheetData: CellData[][] = Array.from({ length: NUM_ROWS }, () =>
          Array.from({ length: NUM_COLUMNS }, () => ({ value: "" }))
        );

        rows.forEach((rowStr, rowIndex) => {
          const cells = rowStr.split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));

          cells.forEach((cellValue, colIndex) => {
            if (newSheetData[rowIndex] && newSheetData[rowIndex][colIndex]) {
              if (cellValue.startsWith('=')) {
                newSheetData[rowIndex][colIndex] = { formula: cellValue, value: evaluateFormula(cellValue, newSheetData, namedRanges) };
              } else {
                newSheetData[rowIndex][colIndex] = { value: cellValue };
              }
            }
          });
        });
        const importedSheetName = file.name.split('.')[0] || "Imported_CSV";
        setSheets(prevSheets => ({
          ...prevSheets,
          [importedSheetName]: newSheetData
        }));
        setActiveSheet(importedSheetName);
        alert("Imported from CSV successfully!");
        setDataUpdateKey(prev => prev + 1);

      } catch (err) {
        alert("Failed to import CSV file: Invalid format or content.");
      }
    };
    reader.readAsText(file);
  };

  const handleFontSizeChange = (size: number) => {
    if (!activeCell.current) return;
    const [col, row] = activeCell.current;

    const updated = { ...sheets };
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) return;

    const copy = updated[activeSheet].map((r) => [...r]);
    const cell = copy[originalRowIndex][col];
    copy[originalRowIndex][col] = {
      ...cell,
      fontSize: size,
    };
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };

  const handleAlignmentChange = (align: "left" | "center" | "right") => {
    if (!activeCell.current) return;
    const [col, row] = activeCell.current;

    const updated = { ...sheets };
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) return;

    const copy = updated[activeSheet].map((r) => [...r]);
    const cell = copy[originalRowIndex][col];
    copy[originalRowIndex][col] = {
      ...cell,
      alignment: align,
    };
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };

  const customDrawCell = useCallback((args: any) => {
    const { ctx, theme, rect, col, row, cell } = args;
    if (highlightRange) {
      const inX = col >= highlightRange.x && col < highlightRange.x + highlightRange.width;
      const inY = row >= highlightRange.y && row < highlightRange.y + highlightRange.height;

      if (inX && inY) {
        ctx.fillStyle = "rgba(30, 144, 255, 0.2)";
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);

        ctx.strokeStyle = "dodgerblue";
        ctx.lineWidth = 2;
        ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
      }
    }
    if (cell.kind !== GridCellKind.Text) return false;

    const displayedRowData = getDisplayedData[row];
    if (!displayedRowData) return false;
    const cellData = displayedRowData.data[col];

    let alignment: "left" | "center" | "right" = cellData?.alignment || "left";
    let fontSize = cellData?.fontSize || 12;
    let isBold = cellData?.bold ? "bold" : "normal";
    let isItalic = cellData?.italic ? "italic" : "normal";
    let isUnderline = cellData?.underline;
    let isStrikethrough = cellData?.strikethrough;
    let textColor = cellData?.textColor || currentTheme.text;
    let bgColor = cellData?.bgColor || currentTheme.bg;
    let borderColor = cellData?.borderColor || currentTheme.border;
    let fontFamily = cellData?.fontFamily || 'sans-serif';
    let text = String(cell.displayData ?? "");

    for (const rule of conditionalFormattingRules) {
      const { range, type, value1, value2, style } = rule;
      const originalRowIndex = displayedRowData.originalRowIndex;
      if (originalRowIndex !== -1 &&
        col >= range.x && col < range.x + range.width &&
        originalRowIndex >= range.y && originalRowIndex < range.y + range.height
      ) {
        const cellNumericValue = parseFloat(text);
        const ruleNumericValue1 = parseFloat(value1 as string);
        const ruleNumericValue2 = parseFloat(value2 as string);

        let applyRule = false;
        switch (type) {
          case 'greaterThan':
            applyRule = !isNaN(cellNumericValue) && !isNaN(ruleNumericValue1) && cellNumericValue > ruleNumericValue1;
            break;
          case 'lessThan':
            applyRule = !isNaN(cellNumericValue) && !isNaN(ruleNumericValue1) && cellNumericValue < ruleNumericValue1;
            break;
          case 'equalTo':
            applyRule = !isNaN(cellNumericValue) && !isNaN(ruleNumericValue1) && cellNumericValue === ruleNumericValue1;
            break;
          case 'between':
            applyRule = !isNaN(cellNumericValue) && !isNaN(ruleNumericValue1) && !isNaN(ruleNumericValue2) &&
              cellNumericValue >= ruleNumericValue1 && cellNumericValue <= ruleNumericValue2;
            break;
          case 'textContains':
            applyRule = typeof value1 === 'string' && text.toLowerCase().includes(value1.toLowerCase());
            break;
        }

        if (applyRule) {
          if (style.bgColor) bgColor = style.bgColor;
          if (style.textColor) textColor = style.textColor;
        }
      }
    }

    if (cell.link) {
      text = `${text}🔗`;
    }
    if (cell.comment) {
      text = `${text}💬`;
    }

    ctx.fillStyle = bgColor;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);

    ctx.strokeStyle = borderColor;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(rect.x, rect.y + rect.height);
    ctx.lineTo(rect.x + rect.width, rect.y + rect.height);
    ctx.moveTo(rect.x + rect.width, rect.y);
    ctx.lineTo(rect.x + rect.width, rect.y + rect.height);
    ctx.stroke();

    ctx.font = `${isItalic} ${isBold} ${fontSize}px ${fontFamily}`;
    ctx.fillStyle = textColor;
    ctx.textBaseline = "middle";
    const padding = 8;
    const textMetrics = ctx.measureText(text);
    let x = rect.x + padding;
    if (alignment === "center") {
      x = rect.x + (rect.width - textMetrics.width) / 2;
    } else if (alignment === "right") {
      x = rect.x + rect.width - textMetrics.width - padding;
    }
    ctx.fillText(text, x, rect.y + rect.height / 2);
    if (isUnderline) {
      const textWidth = textMetrics.width;
      const textHeight = fontSize;
      const underlineY = rect.y + rect.height / 2 + textHeight / 2 + 1;
      ctx.beginPath();
      ctx.strokeStyle = textColor;
      ctx.lineWidth = 1;
      ctx.moveTo(x, underlineY);
      ctx.lineTo(x + textWidth, underlineY);
      ctx.stroke();
    }
    if (isStrikethrough) {
      const textWidth = textMetrics.width;
      const strikethroughY = rect.y + rect.height / 2;
      ctx.beginPath();
      ctx.strokeStyle = textColor;
      ctx.lineWidth = 1;
      ctx.moveTo(x, strikethroughY);
      ctx.lineTo(x + textWidth, strikethroughY);
      ctx.stroke();
    }

    return true;
  }, [getDisplayedData, conditionalFormattingRules, currentTheme]);

  const toggleStyle = (type: "bold" | "italic") => {
    if (!selection.current) return;
    applyStyleToRange(type, undefined);
  };

  const applyCellColor = (type: "textColor" | "bgColor" | "borderColor", color: string) => {
    if (!selection.current) return;
    applyStyleToRange(type, color);
  };

  const pushToUndoStack = (newSheets: SheetData) => {
    setUndoStack(prev => [...prev, sheets]);
    setRedoStack([]);
    setSheets(newSheets);
    setDataUpdateKey(prev => prev + 1);
  };

  const insertRowAt = (index: number) => {
    const updated = { ...sheets };
    const copy = sheets[activeSheet].map((r) => [...r]);
    const emptyRow: CellData[] = Array(NUM_COLUMNS).fill({ value: "" });
    copy.splice(index, 0, [...emptyRow]);
    updated[activeSheet] = copy.slice(0, NUM_ROWS);
    pushToUndoStack(updated);
  };

  const deleteRows = (rowIndex: number, count: number = 1) => {
    const updated = { ...sheets };
    const copy = sheets[activeSheet].map((r) => [...r]);
    copy.splice(rowIndex, count);
    while (copy.length < NUM_ROWS) {
      copy.push(Array(NUM_COLUMNS).fill({ value: "" }));
    }
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };

  const insertColumnAt = (index: number) => {
    const updated = { ...sheets };
    const copy = sheets[activeSheet].map((row) => {
      const newRow = [...row];
      newRow.splice(index, 0, { value: "" });
      return newRow.slice(0, NUM_COLUMNS);
    });
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };

  const deleteColumns = (colIndex: number, count: number = 1) => {
    const updated = { ...sheets };
    const copy = sheets[activeSheet].map((row) => {
      const newRow = [...row];
      newRow.splice(colIndex, count);
      while (newRow.length < NUM_COLUMNS) {
        newRow.push({ value: "" });
      }
      return newRow;
    });
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };
const handleUndo = () => {
    if (undoStack.length === 0) return;
    const previous = undoStack[undoStack.length - 1];
    setRedoStack(prev => [...prev, sheets]);
    setUndoStack(prev => prev.slice(0, -1));
    setSheets(previous);
    
    // Clear and rebuild formula engine
    formulaEngine.current.clear();
    Object.entries(previous).forEach(([sheetName, sheetData]) => {
      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cell.value || cell.formula) {
            try {
              formulaEngine.current.updateCell(colIndex, rowIndex, cell.formula || cell.value, sheetName);
            } catch (error) {
              console.error('Error rebuilding formula engine during undo:', error);
            }
          }
        });
      });
    });
    
    setDataUpdateKey(prev => prev + 1);
  };
const handleRedo = () => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setUndoStack(prev => [...prev, sheets]);
    setRedoStack(prev => prev.slice(0, -1));
    setSheets(next);
    
    // Clear and rebuild formula engine
    formulaEngine.current.clear();
    Object.entries(next).forEach(([sheetName, sheetData]) => {
      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cell.value || cell.formula) {
            try {
              formulaEngine.current.updateCell(colIndex, rowIndex, cell.formula || cell.value, sheetName);
            } catch (error) {
              console.error('Error rebuilding formula engine during redo:', error);
            }
          }
        });
      });
    });
    
    setDataUpdateKey(prev => prev + 1);
  };

  
  
  const handleODSImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target?.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const importedSheet = (json as unknown[]).map((row: unknown) =>
        Array.isArray(row) ? row.map((cell) => ({ value: String(cell ?? "") })) : []
      );

      const updatedSheets = { ...sheets, Sheet1: importedSheet };
      pushToUndoStack(updatedSheets);
      setActiveSheet("Sheet1");
      setDataUpdateKey(prev => prev + 1);
    };
    reader.readAsBinaryString(file);
  };

  const handleODSExport = () => {
    const data = sheets[activeSheet].map((row) =>
      row.map((cell) => cell?.value || "")
    );
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, activeSheet);

    const fileName = saveLoadSheetName?.trim() !== "" ? `${saveLoadSheetName}.ods` : "spreadsheet_export.ods";
    XLSX.writeFile(workbook, fileName, { bookType: "ods" });
  };

  const NUM_COLS = sheets[activeSheet]?.[0]?.length || 10;

  const columns: GridColumn[] = useMemo(() =>
    Array.from({ length: NUM_COLS }, (_, i) => ({
      title: String.fromCharCode(65 + i),
      width: columnWidths[i] ?? 100,
      id: String(i),
    })),
    [sheets, columnWidths]);

  const onColumnResize = useCallback((column: GridColumn, newSize: number, colIndex: number) => {
    setColumnWidths(prev => ({ ...prev, [colIndex]: newSize }));
    setDataUpdateKey(prev => prev + 1);
  }, []);

  const applyStyleToRange = (key: string, value: any) => {
    if (!selection?.current) return;
    const { x, y, width, height } = selection.current.range;
    const currentSheet = sheets[activeSheet];

    let allHaveSame = true;
    for (let row = y; row < y + height; row++) {
      const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
      if (originalRowIndex === undefined || originalRowIndex === -1) continue;

      for (let col = x; col < x + width; col++) {
        const cell = currentSheet?.[originalRowIndex]?.[col];
        if (!cell || cell[key as keyof typeof cell] !== value) {
          allHaveSame = false;
          break;
        }
      }
      if (!allHaveSame) break;
    }

    const updated = { ...sheets };
    const copy = currentSheet.map(r => [...r]);

    for (let row = y; row < y + height; row++) {
      const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
      if (originalRowIndex === undefined || originalRowIndex === -1) continue;

      for (let col = x; col < x + width; col++) {
        const cell = copy[originalRowIndex]?.[col] ?? { value: "" };
        copy[originalRowIndex][col] = { ...cell, [key]: allHaveSame ? undefined : value };
      }
    }

    updated[activeSheet] = copy;
    pushToUndoStack(updated);
  };

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!selection?.current) return;

      const { range } = selection.current;
      const startX = Math.min(range.x, range.x + range.width - 1);
      const endX = Math.max(range.x, range.x + range.width - 1);
      const startY = Math.min(range.y, range.y + range.height - 1);
      const endY = Math.max(range.y, range.y + range.height - 1);

      if (e.ctrlKey && e.key === "c") {
        e.preventDefault();
        const copiedData: any[][] = [];
        for (let row = startY; row <= endY; row++) {
          const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
          if (originalRowIndex === undefined || originalRowIndex === -1) continue;

          const rowData: any[] = [];
          for (let col = startX; col <= endX; col++) {
            rowData.push(sheets[activeSheet][originalRowIndex]?.[col] ?? {
              value: "",
              background: "",
            });
          }
          copiedData.push(rowData);
        }
        navigator.clipboard.writeText(JSON.stringify(copiedData));
      }

      if (e.ctrlKey && e.key === "v") {
        e.preventDefault();
        navigator.clipboard.readText().then((text) => {
          try {
            const copiedData = JSON.parse(text);
            const newSheets = JSON.parse(JSON.stringify(sheets));

            const currentRange = selection.current?.range;
            if (!currentRange) return;

            const pasteStartX = Math.min(currentRange.x, currentRange.x + currentRange.width - 1);
            const pasteStartY = Math.min(currentRange.y, currentRange.y + currentRange.height - 1);

            for (let rowOffset = 0; rowOffset < copiedData.length; rowOffset++) {
              const targetRowDisplayIndex = pasteStartY + rowOffset;
              const originalTargetRowIndex = getDisplayedData[targetRowDisplayIndex]?.originalRowIndex;
              if (originalTargetRowIndex === undefined || originalTargetRowIndex === -1) continue;

              for (let colOffset = 0; colOffset < copiedData[rowOffset].length; colOffset++) {
                const targetCol = pasteStartX + colOffset;
                if (newSheets[activeSheet][originalTargetRowIndex] && newSheets[activeSheet][originalTargetRowIndex][targetCol]) {
                  const cellData = copiedData[rowOffset][colOffset];
                  newSheets[activeSheet][originalTargetRowIndex][targetCol] = {
                    fontSize: 14,
                    bold: false,
                    italic: false,
                    underline: false,
                    strikethrough: false,
                    alignment: "left",
                    ...cellData,
                  };
                  try {
                    const value = cellData.formula || cellData.value || "";
                    formulaEngine.current.updateCell(targetCol, originalTargetRowIndex, value, activeSheet);
                  } catch (error) {
                    console.error('Error updating formula engine during paste:', error);
                  }
                }
              }
            }
            pushToUndoStack(sheets);
            setSheets(newSheets);
            setDataUpdateKey(prev => prev + 1);
          } catch (err) {
            console.warn("Invalid clipboard format");
          }
        });
      }

      if (e.ctrlKey && e.key === "x") {
        e.preventDefault();
        const copiedData: any[][] = [];
        const newSheets = JSON.parse(JSON.stringify(sheets));

        for (let row = startY; row <= endY; row++) {
          const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
          if (originalRowIndex === undefined || originalRowIndex === -1) continue;

          const rowData: any[] = [];
          for (let col = startX; col <= endX; col++) {
            const cell = newSheets[activeSheet][originalRowIndex]?.[col] ?? {
              value: "",
              background: "",
            };
            rowData.push(cell);

            newSheets[activeSheet][originalRowIndex][col] = {
              value: "",
              background: "",
              fontSize: 14,
              bold: false,
              italic: false,
              underline: false,
              strikethrough: false,
              alignment: "left",
            };
          }
          copiedData.push(rowData);
        }

        navigator.clipboard.writeText(JSON.stringify(copiedData));
        pushToUndoStack(sheets);
        setSheets(newSheets);
        setDataUpdateKey(prev => prev + 1);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selection, sheets, activeSheet, clipboardData, getDisplayedData]);

  useEffect(() => {
    const handleCellEdit = ({ sheet, row, col, value }: any) => {
      setSheets(prevSheets => {
        if (sheet !== activeSheet) return prevSheets;
        const updated = { ...prevSheets };
        const copy = updated[sheet].map(r => [...r]);
        copy[row][col] = { value };
        updated[sheet] = copy;
        return updated;
      });
      setDataUpdateKey(prev => prev + 1);
    };
    socket.on("cell-edit", handleCellEdit);
    return () => {
      socket.off("cell-edit", handleCellEdit);
    };
  }, [activeSheet]);

  useEffect(() => {
    const handleClick = () => {
      if (
        formulaInput.startsWith("=") &&
        selecting.current &&
        activeCell.current &&
        (selection.columns.length > 0 || selection.rows.length > 0 || selection.current?.range)
      ) {
        const ref = getCellName(selecting.current.x, selecting.current.y);
        setFormulaInput((prev) => {
          if (prev.includes(ref) && !prev.endsWith("(") && !prev.endsWith(",")) return prev;
          const insert = prev.endsWith("(") || prev.endsWith(",") ? ref : `,${ref}`;
          return prev + insert;
        });
      }
    };

    window.addEventListener("mousedown", handleClick);
    return () => window.removeEventListener("mousedown", handleClick);
  }, [formulaInput, selection]);

  const handleExportTSV = () => {
    const data = sheets[activeSheet];
    const rows = data.map(row =>
      row.map(cell => cell?.value ?? "").join("\t")
    );
    const tsvContent = rows.join("\n");

    const blob = new Blob([tsvContent], { type: "text/tab-separated-values" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${saveLoadSheetName || activeSheet}.tsv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleExportPDF = () => {
    const doc = new jsPDF({ orientation: "landscape" });
    const data = sheets[activeSheet];

    if (!data || data.length === 0) {
      alert("Sheet is empty!");
      return;
    }

    let maxRow = 0;
    let maxCol = 0;

    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;

      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        const val = cell?.value ?? cell?.displayData ?? cell?.data;
        if (val && val.toString().trim() !== "") {
          maxRow = Math.max(maxRow, r);
          maxCol = Math.max(maxCol, c);
        }
      }
    }
    const trimmedData = data.slice(0, maxRow + 1).map(row =>
      Array.from({ length: maxCol + 1 }, (_, i) => row[i]?.value ?? "")
    );
    const headers = Array.from({ length: maxCol + 1 }, (_, i) =>
      String.fromCharCode(65 + i)
    );
    const tableBody = trimmedData.map((row, i) => [`${i + 1}`, ...row]);
    const tableHead = [["", ...headers]];

    autoTable(doc, {
      head: tableHead,
      body: tableBody,
      startY: 20,
      theme: "grid",
      styles: { fontSize: 10, cellPadding: 3 },
      columnStyles: {
        0: { cellWidth: 10 },
        ...Object.fromEntries(Array.from({ length: maxCol + 1 }, (_, i) => [i + 1, { cellWidth: 30 }]))
      },
      headStyles: {
        fillColor: [245, 245, 245],
        textColor: 0,
        fontStyle: "bold"
      },
      didDrawPage: (data) => {
        doc.setFontSize(12);
        doc.text(
          `Sheet: ${saveLoadSheetName || activeSheet}`,
          data.settings.margin.left,
          10
        );
        const pageCount = (doc as any).internal.getNumberOfPages?.() || 1;
        doc.text(`Page ${pageCount}`, doc.internal.pageSize.getWidth() - 30, 10);
      }
    });

    const filename = `${saveLoadSheetName?.trim() || activeSheet}.pdf`;
    doc.save(filename);
  };

  const [showFileDropdown, setShowFileDropdown] = useState(false);
  const [showEditDropdown, setShowEditDropdown] = useState(false);
  const [showInsertDropdown, setShowInsertDropdown] = useState(false);
  const [showFormatDropdown, setShowFormatDropdown] = useState(false);
  const [showDataDropdown, setShowDataDropdown] = useState(false);
  const [showViewDropdown, setShowViewDropdown] = useState(false);

  const [showImportDropdown, setShowImportDropdown] = useState(false);
  const [showExportDropdown, setShowExportDropdown] = useState(false);
  const [editingCell, setEditingCell] = useState<Item | null>(null);

  const [showLinkInput, setShowLinkInput] = useState(false);
  const [linkValue, setLinkValue] = useState('');
  const [showCommentInput, setShowCommentInput] = useState(false);
  const [commentValue, setCommentValue] = useState('');

  const handleInsertRowAbove = () => {
    if (!selection.current) {
      alert("Please select a cell to insert a row.");
      return;
    }
    const originalRowIndex = getDisplayedData[selection.current.range.y]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) {
      alert("Cannot insert row at this position.");
      return;
    }
    insertRowAt(originalRowIndex);
    setShowEditDropdown(false);
  };

  const handleInsertRowBelow = () => {
    if (!selection.current) {
      alert("Please select a cell to insert a row.");
      return;
    }
    const originalRowIndex = getDisplayedData[selection.current.range.y]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) {
      alert("Cannot insert row at this position.");
      return;
    }
    insertRowAt(originalRowIndex + selection.current.range.height);
    setShowEditDropdown(false);
  };

  const handleDeleteSelectedRows = () => {
    if (selection.rows.length === 0) {
      alert("Please select rows to delete.");
      return;
    }
    const selectedRowDisplayIndices: number[] = Array.from(selection.rows);

    const selectedOriginalRowIndices = selectedRowDisplayIndices
      .map(displayIndex => getDisplayedData[displayIndex]?.originalRowIndex)
      .filter(index => index !== undefined && index !== -1) as number[];

    selectedOriginalRowIndices.sort((a, b) => b - a);

    const updatedSheets = { ...sheets };
    let currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

    selectedOriginalRowIndices.forEach(originalRowIndex => {
      currentSheetCopy.splice(originalRowIndex, 1);
    });

    while (currentSheetCopy.length < NUM_ROWS) {
      currentSheetCopy.push(Array(NUM_COLUMNS).fill({ value: "" }));
    }

    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
    setShowEditDropdown(false);
    setSelection({
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty(),
    });
  };

  const handleInsertColumnLeft = () => {
    if (!selection.current) {
      alert("Please select a cell to insert a column.");
      return;
    }
    insertColumnAt(selection.current.range.x);
    setShowEditDropdown(false);
  };

  const handleInsertColumnRight = () => {
    if (!selection.current) {
      alert("Please select a cell to insert a column.");
      return;
    }
    insertColumnAt(selection.current.range.x + selection.current.range.width);
    setShowEditDropdown(false);
  };

  const handleDeleteSelectedColumns = () => {
    if (selection.columns.length === 0) {
      alert("Please select columns to delete.");
      return;
    }
    const selectedColIndices: number[] = Array.from(selection.columns);
    selectedColIndices.sort((a, b) => b - a);

    selectedColIndices.forEach(colIndex => {
      deleteColumns(colIndex);
    });
    setShowEditDropdown(false);
    setSelection({
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty(),
    });
  };

  const handleInsertLink = () => {
    if (!activeCell.current) {
      alert("Please select a cell to insert a link.");
      return;
    }
    setShowLinkInput(true);
    setShowInsertDropdown(false);
  };

  const handleInsertComment = () => {
    if (!activeCell.current) {
      alert("Please select a cell to insert a comment.");
      return;
    }
    setShowCommentInput(true);
    setShowInsertDropdown(false);
  };

  const applyLinkToCell = () => {
    if (!activeCell.current || !linkValue.trim()) {
      alert("Please enter a valid link.");
      return;
    }
    const [col, row] = activeCell.current;
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) return;

    const updated = { ...sheets };
    const copy = updated[activeSheet].map((r) => [...r]);
    const cell = copy[originalRowIndex][col];
    copy[originalRowIndex][col] = {
      ...cell,
      link: linkValue.trim(),
    };
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
    setLinkValue('');
    setShowLinkInput(false);
  };

  const applyCommentToCell = () => {
    if (!activeCell.current || !commentValue.trim()) {
      alert("Please enter a valid comment.");
      return;
    }
    const [col, row] = activeCell.current;
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) return;

    const updated = { ...sheets };
    const copy = updated[activeSheet].map((r) => [...r]);
    const cell = copy[originalRowIndex][col];
    copy[originalRowIndex][col] = {
      ...cell,
      comment: commentValue.trim(),
    };
    updated[activeSheet] = copy;
    pushToUndoStack(updated);
    setCommentValue('');
    setShowCommentInput(false);
  };

  const handleSort = (direction: 'asc' | 'desc') => {
    if (activeCell.current === null) {
      alert("Please select a cell in the column you want to sort.");
      return;
    }
    const colIndex = activeCell.current[0];
    setSortColumnIndex(colIndex);
    setSortDirection(direction);
    setShowDataDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const toggleFilterRow = () => {
    setShowFilterRow(prev => !prev);
    if (showFilterRow) {
      setColumnFilters({});
    }
    setShowDataDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleColumnFilterChange = (colIndex: number, value: string) => {
    setColumnFilters(prev => ({
      ...prev,
      [colIndex]: value,
    }));
    setDataUpdateKey(prev => prev + 1);
  };

  const clearAllFilters = () => {
    setColumnFilters({});
    setSortColumnIndex(null);
    setSortDirection(null);
    setShowFilterRow(false);
    setShowDataDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleApplyConditionalFormatting = () => {
    if (!selection.current) {
      alert("Please select a range to apply conditional formatting.");
      return;
    }
    setShowConditionalFormattingModal(true);
    setShowFormatDropdown(false);
  };

  const addConditionalFormattingRule = () => {
    if (!selection.current || (!cfValue1.trim() && cfType !== 'between')) {
      alert("Please select a range and enter a value for conditional formatting.");
      return;
    }
    if (cfType === 'between' && (!cfValue1.trim() || !cfValue2.trim())) {
      alert("Please enter both values for 'between' rule.");
      return;
    }

    const newRule: ConditionalFormattingRule = {
      id: Date.now().toString(),
      range: selection.current.range,
      type: cfType,
      value1: cfValue1,
      value2: cfType === 'between' ? cfValue2 : undefined,
      style: { bgColor: cfBgColor, textColor: cfTextColor },
    };
    setConditionalFormattingRules(prev => [...prev, newRule]);
    setShowConditionalFormattingModal(false);
    setCfValue1('');
    setCfValue2('');
    setCfType('greaterThan');
    setDataUpdateKey(prev => prev + 1);
  };

  const clearConditionalFormatting = () => {
    setConditionalFormattingRules([]);
    setShowFormatDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleFreezeRows = (count: number) => {
    if (activeCell.current === null && count > 0) {
      alert("Please select a cell to freeze rows up to.");
      return;
    }
    setFrozenRows(count === -1 ? activeCell.current![1] + 1 : count);
    setFrozenColumns(0);
    setShowViewDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleFreezeColumns = (count: number) => {
    if (activeCell.current === null && count > 0) {
      alert("Please select a cell to freeze columns up to.");
      return;
    }
    setFrozenColumns(count === -1 ? activeCell.current![0] + 1 : count);
    setFrozenRows(0);
    setShowViewDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleUnfreezePanes = () => {
    setFrozenRows(0);
    setFrozenColumns(0);
    setShowViewDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleFind = () => {
    if (!findSearchTerm.trim()) {
      setFindMatches([]);
      setFindCurrentMatch(null);
      setFindMatchIndex(0);
      return;
    }

    const matches: Item[] = [];
    const lowerCaseSearchTerm = findSearchTerm.toLowerCase();

    for (let r = 0; r < getDisplayedData.length; r++) {
      const rowData = getDisplayedData[r].data;
      for (let c = 0; c < rowData.length; c++) {
        const cell = rowData[c];
        const cellValue = cell?.value?.toString().toLowerCase() || '';
        if (cellValue.includes(lowerCaseSearchTerm)) {
          matches.push([c, r]);
        }
      }
    }
    setFindMatches(matches);
    if (matches.length > 0) {
      setFindMatchIndex(0);
      setFindCurrentMatch(matches[0]);
      setSelection({
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
        current: {
          cell: matches[0],
          range: { x: matches[0][0], y: matches[0][1], width: 1, height: 1 },
          rangeStack: [],
        },
      });
      gridRef.current?.scrollTo(matches[0][0], matches[0][1], "center", "center");
    } else {
      setFindCurrentMatch(null);
    }
  };

  const handleFindNext = () => {
    if (findMatches.length === 0) return;
    const nextIndex = (findMatchIndex + 1) % findMatches.length;
    setFindMatchIndex(nextIndex);
    const nextMatch = findMatches[nextIndex];
    setFindCurrentMatch(nextMatch);
    setSelection({
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty(),
      current: {
        cell: nextMatch,
        range: { x: nextMatch[0], y: nextMatch[1], width: 1, height: 1 },
        rangeStack: [],
      },
    });
    gridRef.current?.scrollTo(nextMatch[0], nextMatch[0], "center", "center");
  };

  const handleFindPrevious = () => {
    if (findMatches.length === 0) return;
    const prevIndex = (findMatchIndex - 1 + findMatches.length) % findMatches.length;
    setFindMatchIndex(prevIndex);
    const prevMatch = findMatches[prevIndex];
    setFindCurrentMatch(prevMatch);
    setSelection({
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty(),
      current: {
        cell: prevMatch,
        range: { x: prevMatch[0], y: prevMatch[1], width: 1, height: 1 },
        rangeStack: [],
      },
    });
    gridRef.current?.scrollTo(prevMatch[0], prevMatch[1], "center", "center");
  };

  const handleReplace = () => {
    if (!findCurrentMatch) {
      alert("No current match to replace.");
      return;
    }

    const [col, row] = findCurrentMatch;
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) return;

    const cell = currentSheetData[originalRowIndex]?.[col];
    if (cell && cell.value.toLowerCase().includes(findSearchTerm.toLowerCase())) {
      const newValue = cell.value.replace(new RegExp(findSearchTerm, 'gi'), findReplaceTerm);
      onCellEdited([col, row], { kind: GridCellKind.Text, data: newValue, displayData: newValue, allowOverlay: true });
      handleFind();
    } else {
      alert("Current match does not contain the search term or cell is empty.");
    }
  };

  const handleReplaceAll = () => {
    if (!findSearchTerm.trim()) {
      alert("Please enter a search term to replace all.");
      return;
    }

    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map(r => [...r]);
    let replacementsMade = 0;

    for (let r = 0; r < currentSheetCopy.length; r++) {
      for (let c = 0; c < currentSheetCopy[r].length; c++) {
        const cell = currentSheetCopy[r][c];
        if (cell && cell.value.toLowerCase().includes(findSearchTerm.toLowerCase())) {
          const newValue = cell.value.replace(new RegExp(findSearchTerm, 'gi'), findReplaceTerm);
          currentSheetCopy[r][c] = { ...cell, value: newValue, formula: undefined };
          replacementsMade++;
        }
      }
    }
    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
    alert(`${replacementsMade} replacements made.`);
    setShowFindModal(false);
    setFindSearchTerm('');
    setFindReplaceTerm('');
    setFindMatches([]);
    setFindCurrentMatch(null);
    setFindMatchIndex(0);
  };

  const handleApplyDataValidation = () => {
    if (!selection.current) {
      alert("Please select a range to apply data validation.");
      return;
    }
    setShowDataValidationModal(true);
    setShowDataDropdown(false);
  };

  const addDataValidationRule = () => {
    if (!selection.current) {
      alert("No range selected for data validation.");
      return;
    }

    const { x, y, width, height } = selection.current.range;
    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map(r => [...r]);

    for (let row = y; row < y + height; row++) {
      const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
      if (originalRowIndex === undefined || originalRowIndex === -1) continue;

      for (let col = x; col < x + width; col++) {
        const cell = currentSheetCopy[originalRowIndex]?.[col] ?? { value: "" };
        const validationRule: CellData['dataValidation'] = { type: dvType };

        if (dvType === 'number' || dvType === 'text') {
          validationRule.operator = dvOperator;
          validationRule.value1 = dvValue1;
          if (dvOperator === 'between') {
            validationRule.value2 = dvValue2;
          }
        } else if (dvType === 'list') {
          const parsedRange = parseCellName(dvSourceRange.split(':')[0]);
          const parsedEndRange = parseCellName(dvSourceRange.split(':')[1]);
          if (parsedRange && parsedEndRange) {
            validationRule.sourceRange = {
              x: Math.min(parsedRange[0], parsedEndRange[0]),
              y: Math.min(parsedRange[1], parsedEndRange[1]),
              width: Math.abs(parsedRange[0] - parsedEndRange[0]) + 1,
              height: Math.abs(parsedRange[1] - parsedEndRange[1]) + 1,
            };
          } else {
            alert("Invalid source range for list validation.");
            return;
          }
        }

        currentSheetCopy[originalRowIndex][col] = { ...cell, dataValidation: validationRule };
      }
    }
    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
    setShowDataValidationModal(false);
    setDvType('number');
    setDvOperator('greaterThan');
    setDvValue1('');
    setDvValue2('');
    setDvSourceRange('');
    setDataUpdateKey(prev => prev + 1);
  };

  const clearDataValidation = () => {
    if (!selection.current) {
      alert("Please select a range to clear data validation.");
      return;
    }
    const { x, y, width, height } = selection.current.range;
    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map(r => [...r]);

    for (let row = y; row < y + height; row++) {
      const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
      if (originalRowIndex === undefined || originalRowIndex === -1) continue;

      for (let col = x; col < x + width; col++) {
        const cell = currentSheetCopy[originalRowIndex]?.[col];
        if (cell) {
          const { dataValidation, ...rest } = cell;
          currentSheetCopy[originalRowIndex][col] = rest;
        }
      }
    }
    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
    setShowDataDropdown(false);
    setDataUpdateKey(prev => prev + 1);
  };

  const handleManageNamedRanges = () => {
    setShowNamedRangesModal(true);
    setShowDataDropdown(false);
  };

  const addNamedRange = () => {
    if (!newNamedRangeName.trim() || !newNamedRangeRef.trim()) {
      alert("Please enter both a name and a reference for the named range.");
      return;
    }
    if (namedRanges.some(nr => nr.name.toLowerCase() === newNamedRangeName.toLowerCase() && nr.id !== editingNamedRangeId)) {
      alert("A named range with this name already exists.");
      return;
    }

    const rangeMatch = newNamedRangeRef.match(/^([A-Z]+\d+)(?::([A-Z]+\d+))?$/i);
    if (!rangeMatch) {
      alert("Invalid cell or range reference format (e.g., A1 or A1:B5).");
      return;
    }
    const startCell = parseCellName(rangeMatch[1]);
    const endCell = rangeMatch[2] ? parseCellName(rangeMatch[2]) : startCell;

    if (!startCell || !endCell) {
      alert("Invalid cell or range reference.");
      return;
    }

    const [sx, sy] = startCell;
    const [ex, ey] = endCell;
    const x = Math.min(sx, ex);
    const y = Math.min(sy, ey);
    const width = Math.abs(sx - ex) + 1;
    const height = Math.abs(sy - ey) + 1;

    const newRange: NamedRange = {
      id: editingNamedRangeId || Date.now().toString(),
      name: newNamedRangeName.trim(),
      range: { x, y, width, height },
    };

    if (editingNamedRangeId) {
      setNamedRanges(prev => prev.map(nr => nr.id === editingNamedRangeId ? newRange : nr));
      setEditingNamedRangeId(null);
    } else {
      setNamedRanges(prev => [...prev, newRange]);
    }
    setNewNamedRangeName('');
    setNewNamedRangeRef('');
    setDataUpdateKey(prev => prev + 1);
  };

  const editNamedRange = (id: string) => {
    const rangeToEdit = namedRanges.find(nr => nr.id === id);
    if (rangeToEdit) {
      setNewNamedRangeName(rangeToEdit.name);
      setNewNamedRangeRef(`${getCellName(rangeToEdit.range.x, rangeToEdit.range.y)}:${getCellName(rangeToEdit.range.x + rangeToEdit.range.width - 1, rangeToEdit.range.y + rangeToEdit.range.height - 1)}`);
      setEditingNamedRangeId(id);
    }
  };

  const deleteNamedRange = (id: string) => {
    setNamedRanges(prev => prev.filter(nr => nr.id !== id));
    setDataUpdateKey(prev => prev + 1);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', fontFamily: 'Roboto, sans-serif', color: currentTheme.text, backgroundColor: currentTheme.bg }}>
      {/* Top Bar */}
      <div style={{
        backgroundColor: currentTheme.bg,
        borderBottom: `1px solid ${currentTheme.border}`,
        padding: '8px 16px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexShrink: 0,
        boxShadow: currentTheme.shadow,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <span style={{ fontSize: '24px', color: currentTheme.activeTabBorder, fontWeight: 'bold' }}>
            Sheets
          </span>
          <input
            type="text"
            placeholder="Untitled spreadsheet"
            value={saveLoadSheetName}
            onChange={(e) => setSaveLoadSheetName(e.target.value)}
            style={{
              padding: '6px 10px',
              borderRadius: '4px',
              border: `1px solid ${currentTheme.border}`,
              fontSize: '15px',
              fontWeight: 500,
              color: currentTheme.text,
              backgroundColor: currentTheme.bg,
              width: '200px',
            }}
          />

          <div style={{ display: 'flex', gap: '2px', marginLeft: '20px' }}>
            {/* File Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowFileDropdown(true)}
              onMouseLeave={() => { setShowFileDropdown(false); setShowImportDropdown(false); setShowExportDropdown(false); }}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>File</button>
              {showFileDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={saveSheetData} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Save</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>

                  <div
                    style={{ position: "relative" }}
                    onMouseEnter={() => setShowImportDropdown(true)}
                    onMouseLeave={() => setShowImportDropdown(false)}
                  >
                    <button style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Import &#x25B6;</button>
                    {showImportDropdown && (
                      <div style={{ ...subMenuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>
                          Import XLSX
                          <input type="file" accept=".xlsx" onChange={importFromExcel} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>
                          Import JSON
                          <input type="file" accept="application/json" onChange={handleImportFromJSON} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>
                          Import CSV
                          <input type="file" accept=".csv" onChange={importFromCSV} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>
                          Import ODS
                          <input type="file" accept=".ods" onChange={handleODSImport} style={{ display: "none" }} />
                        </label>
                      </div>
                    )}
                  </div>

                  <div
                    style={{ position: "relative" }}
                    onMouseEnter={() => setShowExportDropdown(true)}
                    onMouseLeave={() => setShowExportDropdown(false)}
                  >
                    <button style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export &#x25B6;</button>
                    {showExportDropdown && (
                      <div style={{ ...subMenuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportXLSX} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as XLSX</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportToJSON} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as JSON</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={exportToCSV} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as CSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleODSExport} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as ODS</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportTSV} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as TSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportPDF} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Export as PDF</button>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* Edit Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowEditDropdown(true)}
              onMouseLeave={() => setShowEditDropdown(false)}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>Edit</button>
              {showEditDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleUndo} disabled={undoStack.length === 0} style={{ ...menuItem, opacity: undoStack.length === 0 ? 0.5 : 1, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Undo</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleRedo} disabled={redoStack.length === 0} style={{ ...menuItem, opacity: redoStack.length === 0 ? 0.5 : 1, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Redo</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertRowAbove} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Insert Row Above</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertRowBelow} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Insert Row Below</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertColumnLeft} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Insert Column Left</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertColumnRight} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Insert Column Right</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleDeleteSelectedRows} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Delete Selected Rows</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleDeleteSelectedColumns} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Delete Selected Columns</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => setShowFindModal(true)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Find and Replace...</button>
                </div>
              )}
            </div>

            {/* View Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowViewDropdown(true)}
              onMouseLeave={() => setShowViewDropdown(false)}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>View</button>
              {showViewDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(1)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze 1 row</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(2)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze 2 rows</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(-1)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze up to current row</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(1)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze 1 column</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(2)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze 2 columns</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(-1)} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Freeze up to current column</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleUnfreezePanes} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>No frozen panes</button>
                </div>
              )}
            </div>

            {/* Insert Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowInsertDropdown(true)}
              onMouseLeave={() => setShowInsertDropdown(false)}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>Insert</button>
              {showInsertDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertLink} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Link</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertComment} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Comment</button>
                </div>
              )}
            </div>

            {/* Format Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowFormatDropdown(true)}
              onMouseLeave={() => setShowFormatDropdown(false)}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>Format</button>
              {showFormatDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: currentTheme.textLight }}>Text Styles</div>
                  <select
                    onChange={(e) => applyStyleToRange("fontSize", parseInt(e.target.value))}
                    style={{ ...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px', color: currentTheme.text, backgroundColor: currentTheme.bg, border: `1px solid ${currentTheme.border}` }}
                  >
                    <option value="">Font Size</option>
                    {[10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].map(size => (
                      <option key={size} value={size}>{size}px</option>
                    ))}
                  </select>
                  <select
                    onChange={(e) => applyStyleToRange("fontFamily", e.target.value)}
                    style={{ ...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px', color: currentTheme.text, backgroundColor: currentTheme.bg, border: `1px solid ${currentTheme.border}` }}
                    value={selection?.current ? sheets[activeSheet]?.[selection.current.range.y]?.[selection.current.range.x]?.fontFamily || 'Arial, sans-serif' : 'Arial, sans-serif'}
                  >
                    <option value="">Font Family</option>
                    <option value="Arial, sans-serif">Arial</option>
                    <option value="Helvetica, sans-serif">Helvetica</option>
                    <option value="Verdana, sans-serif">Verdana</option>
                    <option value="Tahoma, sans-serif">Tahoma</option>
                    <option value="Trebuchet MS, sans-serif">Trebuchet MS</option>
                    <option value="Georgia, serif">Georgia</option>
                    <option value="Times New Roman, serif">Times New Roman</option>
                    <option value="Courier New, monospace">Courier New</option>
                    <option value="Lucida Console, monospace">Lucida Console</option>
                  </select>
                  <div style={{ display: "flex", gap: "2px", padding: '4px 8px' }}>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("bold", true)} style={{ ...topBarButtonStyle, fontWeight: "bold", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>B</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("italic", true)} style={{ ...topBarButtonStyle, fontStyle: "italic", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>I</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("underline", true)} style={{ ...topBarButtonStyle, textDecoration: "underline", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>U</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("strikethrough", true)} style={{ ...topBarButtonStyle, textDecoration: "line-through", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>S</button>
                  </div>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: currentTheme.textLight }}>Colors</div>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px", padding: '4px 8px' }}>
                    <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Text:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("textColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                    <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Fill:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("bgColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                    <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Border:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("borderColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                  </div>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleApplyConditionalFormatting} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Conditional Formatting...</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearConditionalFormatting} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Clear Conditional Formatting</button>
                </div>
              )}
            </div>

            {/* Data Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowDataDropdown(true)}
              onMouseLeave={() => setShowDataDropdown(false)}
            >
              <button style={{ ...topBarButtonStyle, color: currentTheme.text }}>Data</button>
              {showDataDropdown && (
                <div style={{ ...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow }}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleSort('asc')} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Sort sheet A-Z (Current Column)</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleSort('desc')} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Sort sheet Z-A (Current Column)</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={toggleFilterRow} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Toggle Filter Row</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearAllFilters} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Clear All Filters</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleApplyDataValidation} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Data Validation...</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearDataValidation} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Clear Data Validation</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleManageNamedRanges} style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }}>Named Ranges...</button>
                </div>
              )}
            </div>
          </div>
        </div>

        <div>
          <button
            onClick={() => setIsDarkMode(prev => !prev)}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = isDarkMode ? '#5f6368' : '#e8eaed')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
            style={{
              padding: '8px 12px',
              background: 'transparent',
              color: currentTheme.text,
              border: `1px solid ${currentTheme.border}`,
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '14px',
              fontWeight: 500,
              marginRight: '50px'
            }}
          >
            {isDarkMode ? 'Light Mode' : 'Dark Mode'}
          </button>
        </div>
      </div>

      {/* Formatting Toolbar */}
      <div style={{
        backgroundColor: currentTheme.bg2,
        borderBottom: `1px solid ${currentTheme.border}`,
        padding: '8px 16px',
        display: 'flex',
        alignItems: 'center',
        flexWrap: 'wrap',
        gap: '10px',
        flexShrink: 0,
      }}>
        <button
          onClick={handleUndo}
          disabled={undoStack.length === 0}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            padding: "6px 10px",
            background: "transparent",
            color: currentTheme.text,
            border: `1px solid ${currentTheme.border}`,
            borderRadius: "4px",
            cursor: "pointer",
            opacity: undoStack.length === 0 ? 0.5 : 1,
            fontSize: "18px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          &#x21BA;
        </button>
        <button
          onClick={handleRedo}
          disabled={redoStack.length === 0}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            padding: "6px 10px",
            background: "transparent",
            color: currentTheme.text,
            border: `1px solid ${currentTheme.border}`,
            borderRadius: "4px",
            cursor: "pointer",
            opacity: redoStack.length === 0 ? 0.5 : 1,
            fontSize: "18px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          &#x21BB;
        </button>

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div>

        <select
          onChange={(e) => applyStyleToRange("fontFamily", e.target.value)}
          style={{
            padding: "6px 10px",
            borderRadius: "4px",
            border: `1px solid ${currentTheme.border}`,
            minWidth: "120px",
            fontSize: '13px',
            color: currentTheme.text,
            backgroundColor: currentTheme.bg,
          }}
          value={selection?.current ? sheets[activeSheet]?.[selection.current.range.y]?.[selection.current.range.x]?.fontFamily || 'Arial, sans-serif' : 'Arial, sans-serif'}
        >
          <option value="Arial, sans-serif">Arial</option>
          <option value="Helvetica, sans-serif">Helvetica</option>
          <option value="Verdana, sans-serif">Verdana</option>
          <option value="Tahoma, sans-serif">Tahoma</option>
          <option value="Trebuchet MS, sans-serif">Trebuchet MS</option>
          <option value="Georgia, serif">Georgia</option>
          <option value="Times New Roman, serif">Times New Roman</option>
          <option value="Courier New, monospace">Courier New</option>
          <option value="Lucida Console, monospace">Lucida Console</option>
        </select>

        <select
          onChange={(e) => applyStyleToRange("fontSize", parseInt(e.target.value))}
          style={{
            padding: "6px 10px",
            borderRadius: "4px",
            border: `1px solid ${currentTheme.border}`,
            minWidth: "70px",
            fontSize: '13px',
            color: currentTheme.text,
            backgroundColor: currentTheme.bg,
          }}
        >
          {[10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].map(size => (
            <option key={size} value={size}>{size}px</option>
          ))}
        </select>

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div>

        <button
          onClick={() => applyStyleToRange("bold", true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            fontWeight: "bold",
            border: `1px solid ${currentTheme.border}`,
            fontSize: '16px',
            color: currentTheme.text,
          }}
        >B</button>
        <button
          onClick={() => applyStyleToRange("italic", true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            fontStyle: "italic",
            border: `1px solid ${currentTheme.border}`,
            fontSize: '16px',
            color: currentTheme.text,
          }}
        >I</button>
        <button
          onClick={() => applyStyleToRange("underline", true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            textDecoration: "underline",
            border: `1px solid ${currentTheme.border}`,
            fontSize: '16px',
            color: currentTheme.text,
          }}
        >U</button>
        <button
          onClick={() => applyStyleToRange("strikethrough", true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            textDecoration: "line-through",
            border: `1px solid ${currentTheme.border}`,
            fontSize: '16px',
            color: currentTheme.text,
          }}
        >S</button>

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div>

        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Text:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("textColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: currentTheme.bg,
            }}
          />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Fill:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("bgColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: currentTheme.bg,
            }}
          />
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: currentTheme.textLight }}>Border:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("borderColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: currentTheme.bg,
            }}
          />
        </div>

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div>

        <select
          onChange={(e) => applyStyleToRange("alignment", e.target.value as "left" | "center" | "right")}
          style={{
            padding: "6px 10px",
            borderRadius: "4px",
            border: `1px solid ${currentTheme.border}`,
            minWidth: "90px",
            fontSize: '13px',
            color: currentTheme.text,
            backgroundColor: currentTheme.bg,
          }}
        >
          <option value="left">Left</option>
          <option value="center">Center</option>
          <option value="right">Right</option>
        </select>
      </div>

      {/* Formula Bar */}
      <div style={{
        backgroundColor: currentTheme.bg2,
        borderBottom: `1px solid ${currentTheme.border}`,
        padding: '8px 16px',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'flex-start',
        flexShrink: 0,
        position: 'relative',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
          <div style={{
            minWidth: '60px',
            fontWeight: 'bold',
            fontSize: '14px',
            color: currentTheme.textLight,
            marginRight: '10px',
            padding: '4px 8px',
            border: `1px solid ${currentTheme.border}`,
            borderRadius: "4px",
            backgroundColor: currentTheme.bg,
            textAlign: 'center',
          }}>
            {activeCell.current ? getCellName(activeCell.current[0], activeCell.current[1]) : 'A1'}
          </div>
          <input
            value={formulaInput}
            onChange={(e) => handleFormulaChange(e.target.value)}
            onFocus={() => {
              updateSuggestions(formulaInput);
            }}
            onBlur={() => {
              setTimeout(() => setShowSuggestions(false), 100);
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter" && activeCell.current) {
                const [col, row] = activeCell.current;
                onCellEdited([col, row], {
                  kind: GridCellKind.Text,
                  data: formulaInput,
                  displayData: formulaInput,
                  allowOverlay: true,
                });
              }
            }}
            placeholder="Enter formula or text"
            style={{
              flexGrow: 1,
              padding: "8px 12px",
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              fontSize: "14px",
              outline: 'none',
              boxShadow: `inset 0 1px 2px ${isDarkMode ? 'rgba(0,0,0,0.3)' : 'rgba(0,0,0,0.06)'}`,
              color: currentTheme.text,
              backgroundColor: currentTheme.bg,
            }}
          />
        </div>
        {formulaError && (
          <div style={{ color: "red", fontSize: "12px", paddingTop: "5px", paddingLeft: "80%", width:"100%" }}>
            {formulaError}
          </div>
        )}
        {showSuggestions && (
          <div
            style={{
              position: "absolute",
              background: currentTheme.menuBg,
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              boxShadow: currentTheme.shadow,
              zIndex: 1000,
              width: "calc(100% - 100px)",
              marginTop: "4px",
              top: 'calc(100% + 5px)',
              left: '80px',
            }}
          >
            {suggestions.map((suggestion, idx) => (
              <div
                key={idx}
                onClick={() => {
                  setFormulaInput(suggestion);
                  setShowSuggestions(false);
                }}
                style={{
                  padding: "6px 8px",
                  cursor: "pointer",
                  borderBottom: idx === suggestions.length - 1 ? "none" : `1px solid ${currentTheme.border}`,
                  backgroundColor: currentTheme.menuBg,
                  color: currentTheme.text,
                  fontSize: '13px',
                }}
                onMouseDown={(e) => e.preventDefault()}
              >
                {suggestion}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Filter Row */}
      {showFilterRow && (
        <div style={{
          backgroundColor: currentTheme.bg2,
          borderBottom: `1px solid ${currentTheme.border}`,
          padding: '4px 16px',
          display: 'flex',
          alignItems: 'center',
          flexShrink: 0,
        }}>
          <div style={{ width: '60px', flexShrink: 0 }}></div>
          {columns.map((col, index) => (
            <input
              key={col.id}
              type="text"
              placeholder={`Filter ${col.title}`}
              value={columnFilters[index] || ''}
              onChange={(e) => handleColumnFilterChange(index, e.target.value)}
              style={{
                width: columnWidths[index] ?? 100,
                minWidth: '50px',
                padding: '4px 8px',
                border: `1px solid ${currentTheme.border}`,
                borderRadius: '4px',
                marginRight: '2px',
                fontSize: '12px',
                color: currentTheme.text,
                backgroundColor: currentTheme.bg,
              }}
            />
          ))}
        </div>
      )}

      {/* Link Input Modal */}
      {showLinkInput && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "10px",
          color: currentTheme.text,
        }}>
          <h3>Insert Link</h3>
          <input
            type="text"
            value={linkValue}
            onChange={(e) => setLinkValue(e.target.value)}
            placeholder="Enter URL"
            style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={applyLinkToCell} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Apply</button>
            <button onClick={() => setShowLinkInput(false)} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Comment Input Modal */}
      {showCommentInput && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "10px",
          color: currentTheme.text,
        }}>
          <h3>Insert Comment</h3>
          <textarea
            value={commentValue}
            onChange={(e) => setCommentValue(e.target.value)}
            placeholder="Enter comment"
            rows={4}
            style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, minWidth: "250px", color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={applyCommentToCell} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Apply</button>
            <button onClick={() => setShowCommentInput(false)} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Conditional Formatting Modal */}
      {showConditionalFormattingModal && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "15px",
          color: currentTheme.text,
        }}>
          <h3>Conditional Formatting Rule</h3>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label>Apply to range:</label>
            <input
              type="text"
              value={selection.current ? `${getCellName(selection.current.range.x, selection.current.range.y)}:${getCellName(selection.current.range.x + selection.current.range.width - 1, selection.current.range.y + selection.current.range.height - 1)}` : ''}
              readOnly
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, flexGrow: 1, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label>Format cells if:</label>
            <select
              value={cfType}
              onChange={(e) => setCfType(e.target.value as ConditionalFormattingRule['type'])}
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            >
              <option value="greaterThan">Greater than</option>
              <option value="lessThan">Less than</option>
              <option value="equalTo">Equal to</option>
              <option value="between">Between</option>
              <option value="textContains">Text contains</option>
            </select>
            <input
              type={cfType === 'textContains' ? 'text' : 'number'}
              value={cfValue1}
              onChange={(e) => setCfValue1(e.target.value)}
              placeholder="Value 1"
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, width: '100px', color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            />
            {cfType === 'between' && (
              <input
                type="number"
                value={cfValue2}
                onChange={(e) => setCfValue2(e.target.value)}
                placeholder="Value 2"
                style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, width: '100px', color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
              />
            )}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label>Background Color:</label>
            <input
              type="color"
              value={cfBgColor}
              onChange={(e) => setCfBgColor(e.target.value)}
              style={{ width: "30px", height: "30px", border: "none", borderRadius: "4px", cursor: "pointer" }}
            />
            <label>Text Color:</label>
            <input
              type="color"
              value={cfTextColor}
              onChange={(e) => setCfTextColor(e.target.value)}
              style={{ width: "30px", height: "30px", border: "none", borderRadius: "4px", cursor: "pointer" }}
            />
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={addConditionalFormattingRule} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Apply Rule</button>
            <button onClick={() => setShowConditionalFormattingModal(false)} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Data Validation Modal */}
      {showDataValidationModal && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "15px",
          color: currentTheme.text,
        }}>
          <h3>Data Validation Rule</h3>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label>Apply to range:</label>
            <input
              type="text"
              value={selection.current ? `${getCellName(selection.current.range.x, selection.current.range.y)}:${getCellName(selection.current.range.x + selection.current.range.width - 1, selection.current.range.y + selection.current.range.height - 1)}` : ''}
              readOnly
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, flexGrow: 1, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <label>Criteria:</label>
            <select
              value={dvType}
              onChange={(e) => { setDvType(e.target.value as 'number' | 'text' | 'date' | 'list'); setDvOperator('greaterThan'); setDvValue1(''); setDvValue2(''); setDvSourceRange(''); }}
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            >
              <option value="number">Number</option>
              <option value="text">Text</option>
              <option value="list">List from range</option>
            </select>
            {(dvType === 'number' || dvType === 'text') && (
              <>
                <select
                  value={dvOperator}
                  onChange={(e) => setDvOperator(e.target.value as 'greaterThan' | 'lessThan' | 'equalTo' | 'notEqualTo' | 'between' | 'textContains' | 'startsWith' | 'endsWith')}
                  style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
                >
                  {dvType === 'number' && (
                    <>
                      <option value="greaterThan">Greater than</option>
                      <option value="lessThan">Less than</option>
                      <option value="equalTo">Equal to</option>
                      <option value="notEqualTo">Not equal to</option>
                      <option value="between">Between</option>
                    </>
                  )}
                  {dvType === 'text' && (
                    <>
                      <option value="textContains">Text contains</option>
                      <option value="startsWith">Starts with</option>
                      <option value="endsWith">Ends with</option>
                    </>
                  )}
                </select>
                <input
                  type={dvType === 'number' ? 'number' : 'text'}
                  value={dvValue1}
                  onChange={(e) => setDvValue1(e.target.value)}
                  placeholder="Value 1"
                  style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, width: '100px', color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
                />
                {dvOperator === 'between' && (
                  <input
                    type="number"
                    value={dvValue2}
                    onChange={(e) => setDvValue2(e.target.value)}
                    placeholder="Value 2"
                    style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, width: '100px', color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
                  />
                )}
              </>
            )}
            {dvType === 'list' && (
              <input
                type="text"
                value={dvSourceRange}
                onChange={(e) => setDvSourceRange(e.target.value)}
                placeholder="e.g., A1:A5"
                style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, flexGrow: 1, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
              />
            )}
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={addDataValidationRule} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Apply Rule</button>
            <button onClick={() => setShowDataValidationModal(false)} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Named Ranges Modal */}
      {showNamedRangesModal && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "15px",
          minWidth: '400px',
          color: currentTheme.text,
        }}>
          <h3>Manage Named Ranges</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <input
              type="text"
              value={newNamedRangeName}
              onChange={(e) => setNewNamedRangeName(e.target.value)}
              placeholder="Named Range Name"
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            />
            <input
              type="text"
              value={newNamedRangeRef}
              onChange={(e) => setNewNamedRangeRef(e.target.value)}
              placeholder="e.g., A1 or A1:B5"
              style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            />
            <button onClick={addNamedRange} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>
              {editingNamedRangeId ? 'Update Named Range' : 'Add Named Range'}
            </button>
          </div>

          <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '10px 0' }}></div>

          <h4>Existing Named Ranges:</h4>
          <div style={{ maxHeight: '200px', overflowY: 'auto', border: `1px solid ${currentTheme.border}`, borderRadius: '4px', padding: '5px' }}>
            {namedRanges.length === 0 ? (
              <p style={{ color: currentTheme.textLight }}>No named ranges defined.</p>
            ) : (
              namedRanges.map(nr => (
                <div key={nr.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '5px 0', borderBottom: `1px dashed ${currentTheme.border}` }}>
                  <span style={{ color: currentTheme.text }}>
                    {nr.name}: {getCellName(nr.range.x, nr.range.y)}:{getCellName(nr.range.x + nr.range.width - 1, nr.range.y + nr.range.height - 1)}
                  </span>
                  <div>
                    <button onClick={() => editNamedRange(nr.id)} style={{ padding: "4px 8px", background: currentTheme.activeTabBg, color: currentTheme.text, border: `1px solid ${currentTheme.border}`, borderRadius: "4px", cursor: "pointer", marginRight: '5px' }}>Edit</button>
                    <button onClick={() => deleteNamedRange(nr.id)} style={{ padding: "4px 8px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Delete</button>
                  </div>
                </div>
              ))
            )}
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px", marginTop: '10px' }}>
            <button onClick={() => {
              setShowNamedRangesModal(false);
              setNewNamedRangeName('');
              setNewNamedRangeRef('');
              setEditingNamedRangeId(null);
            }} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Close</button>
          </div>
        </div>
      )}

      {/* Find Modal */}
      {showFindModal && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: currentTheme.bg,
          padding: "20px",
          borderRadius: "8px",
          boxShadow: currentTheme.shadow,
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "10px",
          minWidth: '300px',
          color: currentTheme.text,
        }}>
          <h3>Find and Replace in Sheet</h3>
          <input
            type="text"
            value={findSearchTerm}
            onChange={(e) => setFindSearchTerm(e.target.value)}
            placeholder="Find what"
            style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                handleFind();
              }
            }}
          />
          <input
            type="text"
            value={findReplaceTerm}
            onChange={(e) => setFindReplaceTerm(e.target.value)}
            placeholder="Replace with"
            style={{ padding: "8px", borderRadius: "4px", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.bg2 }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                handleReplace();
              }
            }}
          />
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: 'center', gap: "10px" }}>
            <div style={{ fontSize: '12px', color: currentTheme.textLight }}>
              {findMatches.length > 0 ? `${findMatchIndex + 1} of ${findMatches.length}` : 'No matches'}
            </div>
            <div style={{ display: 'flex', gap: '5px' }}>
              <button onClick={handleFindPrevious} disabled={findMatches.length === 0} style={{ padding: "6px 10px", background: currentTheme.bg2, color: currentTheme.text, border: `1px solid ${currentTheme.border}`, borderRadius: "4px", cursor: "pointer" }}>&#x25C0; Prev</button>
              <button onClick={handleFindNext} disabled={findMatches.length === 0} style={{ padding: "6px 10px", background: currentTheme.bg2, color: currentTheme.text, border: `1px solid ${currentTheme.border}`, borderRadius: "4px", cursor: "pointer" }}>Next &#x25B6;</button>
            </div>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={handleReplace} disabled={findCurrentMatch === null} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Replace</button>
            <button onClick={handleReplaceAll} disabled={findMatches.length === 0} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Replace All</button>
            <button onClick={() => {
              setShowFindModal(false);
              setFindSearchTerm('');
              setFindReplaceTerm('');
              setFindMatches([]);
              setFindCurrentMatch(null);
              setFindMatchIndex(0);
            }} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Close</button>
          </div>
        </div>
      )}

      {/* Enhanced Context Menu */}
      {contextMenuOpen && (
        <div
          style={{
            position: "fixed",
            top: contextMenuPosition.y,
            left: contextMenuPosition.x,
            background: currentTheme.menuBg,
            border: `1px solid ${currentTheme.border}`,
            borderRadius: "4px",
            boxShadow: currentTheme.shadow,
            zIndex: 9999,
            minWidth: "180px",
            padding: "4px 0",
          }}
          onContextMenu={(e) => e.preventDefault()}
        >
          {contextMenuType === 'row' && rightClickedRow !== null ? (
            <>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  const originalRowIndex = getDisplayedData[rightClickedRow]?.originalRowIndex;
                  if (originalRowIndex !== undefined && originalRowIndex !== -1) {
                    insertRowAt(originalRowIndex);
                  }
                  setContextMenuOpen(false);
                }}
              >
                ✂️ Cut
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Copy functionality
                  setContextMenuOpen(false);
                }}
              >
                📋 Copy
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Paste functionality
                  setContextMenuOpen(false);
                }}
              >
                📄 Paste
              </button>
              <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  const originalRowIndex = getDisplayedData[rightClickedRow]?.originalRowIndex;
                  if (originalRowIndex !== undefined && originalRowIndex !== -1) {
                    insertRowAt(originalRowIndex);
                  }
                  setContextMenuOpen(false);
                }}
              >
                ➕ Insert 1 row above
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  const originalRowIndex = getDisplayedData[rightClickedRow]?.originalRowIndex;
                  if (originalRowIndex !== undefined && originalRowIndex !== -1) {
                    insertRowAt(originalRowIndex + 1);
                  }
                  setContextMenuOpen(false);
                }}
              >
                ➕ Insert 1 row below
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  const originalRowIndex = getDisplayedData[rightClickedRow]?.originalRowIndex;
                  if (originalRowIndex !== undefined && originalRowIndex !== -1) {
                    deleteRows(originalRowIndex, 1);
                  }
                  setContextMenuOpen(false);
                }}
              >
                🗑️ Delete row
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Clear row functionality
                  setContextMenuOpen(false);
                }}
              >
                ✖️ Clear row
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Hide row functionality
                  setContextMenuOpen(false);
                }}
              >
                👁️ Hide row
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Resize row functionality
                  setContextMenuOpen(false);
                }}
              >
                📏 Resize row
              </button>
              <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Create filter functionality
                  toggleFilterRow();
                  setContextMenuOpen(false);
                }}
              >
                🔽 Create a filter
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Conditional formatting functionality
                  handleApplyConditionalFormatting();
                  setContextMenuOpen(false);
                }}
              >
                🎨 Conditional formatting
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Data validation functionality
                  handleApplyDataValidation();
                  setContextMenuOpen(false);
                }}
              >
                ✅ Data validation
              </button>
            </>
          ) : (
            <>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Cut functionality for cells
                  setContextMenuOpen(false);
                }}
              >
                ✂️ Cut
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Copy functionality for cells
                  setContextMenuOpen(false);
                }}
              >
                📋 Copy
              </button>
              <button 
                style={{ ...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg }} 
                onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} 
                onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}
                onClick={() => {
                  // Paste functionality for cells
                  setContextMenuOpen(false);
                }}
              >
                📄 Paste
              </button>
            </>
          )}
        </div>
      )}

      <div 
        style={{ flexGrow: 1, overflow: 'hidden' }}
        onContextMenu={handleCellContextMenu}
      >
        <DataEditor
          columns={columns}
          rows={getDisplayedData.length}
          rowMarkers="both"
          getCellContent={getCellContent}
          onCellEdited={onCellEdited}
          gridSelection={selection}
          drawCell={customDrawCell}
          ref={gridRef}
          onColumnResize={onColumnResize}
          onGridSelectionChange={(sel) => {
            setSelection(sel);
            if (sel.current) {
              activeCell.current = sel.current.cell;
              selecting.current = sel.current.range;
              setSelectedRanges([sel.current.range, ...(sel.current.rangeStack || [])]);
            }
          }}
          onCellActivated={(cell) => {
            const [col, row] = cell;
            setSelection({
              columns: CompactSelection.empty(),
              rows: CompactSelection.empty(),
              current: {
                cell,
                range: { x: cell[0], y: cell[1], width: 1, height: 1 },
                rangeStack: [],
              },
            });
            const originalRowIndex = getDisplayedData[cell[1]]?.originalRowIndex;
            if (originalRowIndex !== undefined && originalRowIndex !== -1) {
              const currentCell = currentSheetData[originalRowIndex]?.[cell[0]];
              if (currentCell) {
                setFormulaInput(currentCell.formula || currentCell.value);
              } else {
                setFormulaInput("");
              }
            } else {
              setFormulaInput("");
            }
            setShowFormulaGuidance(false);
          }}
          onFillPattern={onFillPattern}
          rowHeight={(row) => {
            const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
            if (originalRowIndex === undefined || originalRowIndex === -1) {
              return 28;
            }
            const rowCells = sheets[activeSheet][originalRowIndex];
            const maxFontSize = rowCells.reduce((max, cell) => {
              const size = cell?.fontSize || 14;
              return Math.max(max, size);
            }, 14);
            return maxFontSize + 6;
          }}
          onKeyDown={(e) => {
            if (!selection?.current) return;

            const { range } = selection.current;
            const startX = Math.min(range.x, range.x + range.width - 1);
            const endX = Math.max(range.x, range.x + range.width - 1);
            const startY = Math.min(range.y, range.y + range.height - 1);
            const endY = Math.max(range.y, range.y + range.height - 1);

            if (e.key === "Enter") {
              e.preventDefault();
              const cell = selection.current?.cell;
              if (cell) {
                // Handle enter key for formula input
                if (formulaInput) {
                  onCellEdited(cell, { kind: GridCellKind.Text, data: formulaInput, displayData: formulaInput, allowOverlay: true });
                }
              }
            }
            if (e.key === "Escape") {
              e.preventDefault();
              const cell = selection.current?.cell;
              if (cell) {
                const [col, row] = cell;
                const originalRowIndex = getDisplayedData[row]?.originalRowIndex;

                if (
                  originalRowIndex !== undefined &&
                  originalRowIndex !== -1 &&
                  sheets[activeSheet][originalRowIndex]
                ) {
                  const newSheets = JSON.parse(JSON.stringify(sheets));
                  newSheets[activeSheet][originalRowIndex][col] = {
                    ...newSheets[activeSheet][originalRowIndex][col],
                    value: "",
                    formula: "",
                  };
                  pushToUndoStack(sheets);
                  setSheets(newSheets);
                  setDataUpdateKey((prev) => prev + 1);
                }
              }
            }

            if (e.ctrlKey && e.key === "c") {
              e.preventDefault();
              const copiedData: any[][] = [];
              for (let row = startY; row <= endY; row++) {
                const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
                if (originalRowIndex === undefined || originalRowIndex === -1) continue;

                const rowData: any[] = [];
                for (let col = startX; col <= endX; col++) {
                  rowData.push(sheets[activeSheet][originalRowIndex]?.[col] ?? {
                    value: "",
                    background: "",
                  });
                }
                copiedData.push(rowData);
              }
              navigator.clipboard.writeText(JSON.stringify(copiedData));
            }

            if (e.ctrlKey && e.key === "v") {
              e.preventDefault();
              navigator.clipboard.readText().then((text) => {
                try {
                  const copiedData = JSON.parse(text);
                  const newSheets = JSON.parse(JSON.stringify(sheets));

                  for (let rowOffset = 0; rowOffset < copiedData.length; rowOffset++) {
                    const targetRowDisplayIndex = startY + rowOffset;
                    const originalTargetRowIndex = getDisplayedData[targetRowDisplayIndex]?.originalRowIndex;
                    if (originalTargetRowIndex === undefined || originalTargetRowIndex === -1) continue;

                    for (let colOffset = 0; colOffset < copiedData[rowOffset].length; colOffset++) {
                      const targetCol = startX + colOffset;
                      if (newSheets[activeSheet][originalTargetRowIndex] && newSheets[activeSheet][originalTargetRowIndex][targetCol]) {
                        newSheets[activeSheet][originalTargetRowIndex][targetCol] = {
                          fontSize: 14,
                          bold: false,
                          italic: false,
                          underline: false,
                          strikethrough: false,
                          alignment: "left",
                          ...copiedData[rowOffset][colOffset],
                        };
                      }
                    }
                  }
                  pushToUndoStack(sheets);
                  setSheets(newSheets);
                  setDataUpdateKey(prev => prev + 1);
                } catch (err) {
                  console.warn("Invalid clipboard format");
                }
              });
            }

            if (e.ctrlKey && e.key === "x") {
              e.preventDefault();
              const copiedData: any[][] = [];
              const newSheets = JSON.parse(JSON.stringify(sheets));

              for (let row = startY; row <= endY; row++) {
                const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
                if (originalRowIndex === undefined || originalRowIndex === -1) continue;

                const rowData: any[] = [];
                for (let col = startX; col <= endX; col++) {
                  const cell = newSheets[activeSheet][originalRowIndex]?.[col] ?? {
                    value: "",
                    background: "",
                  };
                  rowData.push(cell);

                  newSheets[activeSheet][originalRowIndex][col] = {
                    value: "",
                    background: "",
                    fontSize: 14,
                    bold: false,
                    italic: false,
                    underline: false,
                    strikethrough: false,
                    alignment: "left",
                  };
                }
                copiedData.push(rowData);
              }

              navigator.clipboard.writeText(JSON.stringify(copiedData));
              pushToUndoStack(sheets);
              setSheets(newSheets);
              setDataUpdateKey(prev => prev + 1);
            }
          }}
          onHeaderContextMenu={(col, e) => {
  handleContextMenu(e as unknown as React.MouseEvent);
}}
onCellContextMenu={(cell, e) => {
  const [col, row] = cell;
  handleContextMenu(e as unknown as React.MouseEvent, row);
}}
          
          keybindings={{
            selectAll: true,
            search: true,
            copy: true,
            paste: true,
            cut: true,
          }}
        />
      </div>

      {/* Context Menu */}
      {contextMenuOpen && (
        <div
          style={{
            position: "fixed",
            top: contextMenuPosition.y,
            left: contextMenuPosition.x,
            backgroundColor: currentTheme.menuBg,
            border: `1px solid ${currentTheme.border}`,
            borderRadius: "4px",
            boxShadow: currentTheme.shadow,
            zIndex: 9999,
            minWidth: "180px",
            padding: "4px 0",
          }}
          onContextMenu={(e) => e.preventDefault()}
        >
          <button onClick={() => { handleInsertRowAbove(); setContextMenuOpen(false); }} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}} onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}>Insert 1 row above</button>
          <button onClick={() => { handleInsertRowBelow(); setContextMenuOpen(false); }} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}} onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}>Insert 1 row below</button>
          <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
          <button onClick={() => { if (rightClickedRow !== null) { deleteRows(rightClickedRow, 1); } setContextMenuOpen(false); }} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}} onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}>Delete row</button>
          <button onClick={() => { /* Clear row logic */ setContextMenuOpen(false); }} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}} onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)}>Clear row</button>
        </div>
      )}

      {/* Formula Guidance Tooltip */}
      {showFormulaGuidance && currentFormulaFunction && formulaArgHints[currentFormulaFunction] && (
        <div style={{
          position: "absolute",
          top: "170px", // Position below formula bar
          left: "90px", // Align with formula input
          backgroundColor: "#fffef3",
          border: "1px solid #ccc",
          borderRadius: "4px",
          padding: "8px 12px",
          fontSize: "12px",
          color: "#333",
          zIndex: 1001,
          boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
        }}>
          <strong>{currentFormulaFunction}</strong> (
          {formulaArgHints[currentFormulaFunction].map((arg, i) => (
            <span
              key={i}
              style={{
                fontWeight: i === currentFormulaArgIndex ? "bold" : "normal",
                color: i === currentFormulaArgIndex ? "#0078d7" : "#666",
                backgroundColor: i === currentFormulaArgIndex ? "#e6f3ff" : "transparent",
                padding: i === currentFormulaArgIndex ? "2px 4px" : "0",
                borderRadius: "2px",
              }}
            >
              {arg}
              {i < formulaArgHints[currentFormulaFunction].length - 1 ? ", " : ""}
            </span>
          ))}
          )
        </div>
      )}

      <div style={{
        padding: "10px",
        borderTop: `1px solid ${currentTheme.border}`,
        display: "flex",
        alignItems: "center",
        overflowX: "auto",
        flexShrink: 0,
        backgroundColor: currentTheme.bg2,
      }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          {Object.keys(sheets).map((sheetName) => (
            <span
              key={sheetName}
              style={{
                padding: "8px 12px",
                marginRight: "2px",
                cursor: "pointer",
                backgroundColor: activeSheet === sheetName ? currentTheme.activeTabBg : "transparent",
                border: "1px solid transparent",
                borderBottom: activeSheet === sheetName ? `2px solid ${currentTheme.activeTabBorder}` : `1px solid ${currentTheme.border}`,
                borderRadius: "4px 4px 0 0",
                display: "inline-flex",
                alignItems: "center",
                whiteSpace: "nowrap",
                color: activeSheet === sheetName ? currentTheme.activeTabBorder : currentTheme.textLight,
                fontWeight: activeSheet === sheetName ? "bold" : "normal",
                transition: 'background-color 0.2s, border-bottom 0.2s',
                minWidth: '80px',
                justifyContent: 'center',
              }}
              onClick={() => {
                setActiveSheet(sheetName);
              }}
              onDoubleClick={() => handleEditSheetName(sheetName)}
            >
              {editingSheetName === sheetName ? (
                <input
                  type="text"
                  value={newSheetName}
                  onChange={(e) => setNewSheetName(e.target.value)}
                  onBlur={() => handleSaveSheetName(sheetName)}
                  onKeyDown={(e) => handleSheetNameInputKeyDown(e, sheetName)}
                  style={{ border: "none", background: "transparent", outline: "none", width: "80px", textAlign: 'center', color: currentTheme.text }}
                  autoFocus
                />
              ) : (
                <>
                  {sheetName}
                  {Object.keys(sheets).length > 1 && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        handleDeleteSheet(sheetName);
                      }}
                      style={{
                        marginLeft: "8px",
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        color: currentTheme.textLight,
                        fontSize: "14px",
                      }}
                    >
                      &#x2715;
                    </button>
                  )}
                </>
              )}
            </span>
          ))}
          <button
            onClick={handleAddSheet}
            style={{
              padding: "8px 12px",
              background: currentTheme.activeTabBg,
              color: currentTheme.textLight,
              border: `1px solid ${currentTheme.border}`,
              borderRadius: "4px",
              cursor: "pointer",
              flexShrink: 0,
              marginLeft: "10px",
              fontWeight: 500,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            +
          </button>
        </div>
      </div>

      {/* Click outside handler for context menu */}
      {contextMenuOpen && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100vw",
            height: "100vh",
            zIndex: 9998,
          }}
          onClick={() => setContextMenuOpen(false)}
        />
      )}
    </div>
  );
};

export default ContactGrid;