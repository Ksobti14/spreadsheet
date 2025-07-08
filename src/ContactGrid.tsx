import React, { useCallback, useRef, useState, useEffect,useMemo } from "react";
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
import {io} from "socket.io-client"
const NUM_ROWS = 7000;
const NUM_COLUMNS = 100;
const socket=io("http://localhost:5000",{
   transports:["websocket","polling"],
   withCredentials:true
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
    value2?: string | number; // For 'between'
    sourceRange?: Rectangle; // For 'list' from range
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
  value2?: string | number; // For 'between'
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
  // Check for named ranges first
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
  // Pass namedRanges to parseArg
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
        const condVal = parseArg(cond, data, namedRanges)[0] || 0; // Pass namedRanges
        return condVal ? parseArg(trueVal, data, namedRanges)[0].toString() || trueVal : parseArg(falseVal, data, namedRanges)[0].toString() || falseVal; // Pass namedRanges
      case "ROUND":
        // ROUND(number, digits)
        const [numArg, digitsArg] = args.split(',').map(s => s.trim());
        const num = parseArg(numArg, data, namedRanges)[0] || 0; // Pass namedRanges
        const digits = parseInt(digitsArg) || 0;
        return FormulaJS.ROUND(num, digits).toString();
      case "ABS":
        const absArg = parseArg(args, data, namedRanges)[0] || 0; // Pass namedRanges
        return FormulaJS.ABS(absArg).toString();
      case "SQRT":
        const sqrtArg = parseArg(args, data, namedRanges)[0] || 0; // Pass namedRanges
        return sqrtArg >= 0 ? FormulaJS.SQRT(sqrtArg).toString() : error("Negative number");
      case "POWER":
        const [baseArg, expArg] = args.split(',').map(s => s.trim());
        const base = parseArg(baseArg, data, namedRanges)[0] || 0; // Pass namedRanges
        const exp = parseArg(expArg, data, namedRanges)[0] || 1; // Pass namedRanges
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
  padding: "8px 16px", // Slightly less padding for a compact menu
  textAlign: "left",
  border: "none",
  cursor: "pointer",
  width: "100%",
  fontSize: "13px", // Slightly smaller font size
  whiteSpace: "nowrap",
  display: "block", // Ensure buttons take full width
};

const menuDropdownStyle: React.CSSProperties = {
  position: "absolute",
  top: "100%",
  left: 0,
  borderRadius: "4px",
  zIndex: 1000,
  minWidth: "160px", // Adjusted min-width
  padding: "4px 0", // Reduced padding
  display: "flex",
  flexDirection: "column",
};

const subMenuDropdownStyle: React.CSSProperties = {
  position: "absolute",
  top: "0",
  left: "100%", // Position to the right of the parent menu item
  borderRadius: "4px",
  zIndex: 1001, // Higher z-index to appear above parent dropdown
  minWidth: "160px",
  padding: "4px 0",
  display: "flex",
  flexDirection: "column",
};

const topBarButtonStyle: React.CSSProperties = {
  padding: "6px 10px", // Smaller padding for top bar buttons
  background: "transparent", // Transparent background
  border: "none",
  borderRadius: "4px",
  cursor: "pointer",
  fontSize: "13px", // Smaller font size
  fontWeight: 500,
};

const FormulaOverlay: React.FC<{
  location: Item;
  value: string;
  onChange: (val: string) => void;
  onCommit: () => void;
  onCancel: () => void;
}> = ({ location, value, onChange, onCommit, onCancel }) => {
  const [localValue, setLocalValue] = useState(value);
  const [editingCell, setEditingCell] = useState<Item | null>(null);
const [currentArgIndex, setCurrentArgIndex] = useState<number | null>(null);
const [currentFunction, setCurrentFunction] = useState<string | null>(null);

  const inputRef = useRef<HTMLInputElement>(null);

  // ✅ Focus input on mount
  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.setSelectionRange(localValue.length, localValue.length);
  }, []);
  useEffect(() => {
  if (!localValue.startsWith("=")) {
    setCurrentFunction(null);
    setCurrentArgIndex(null);
    return;
  }

  const cursorPos = inputRef.current?.selectionStart ?? localValue.length;
  const beforeCursor = localValue.slice(0, cursorPos);

  const funcMatch = beforeCursor.match(/^=(\w+)\(/i);
  if (funcMatch) {
    const fn = funcMatch[1].toUpperCase();
    const argsPart = beforeCursor.slice(funcMatch[0].length);
    const argIndex = argsPart.split(",").length - 1;

    setCurrentFunction(fn);
    setCurrentArgIndex(argIndex);
  } else {
    setCurrentFunction(null);
    setCurrentArgIndex(null);
  }
}, [localValue]);

useEffect(() => {
  setEditingCell(location); // enter edit mode
  return () => setEditingCell(null); // exit edit mode on unmount
}, []);

  // ✅ Handle input changes
  const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setLocalValue(val);
    onChange(val);
    
  };

  // ✅ Handle Enter/Esc
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      onChange(localValue);
      onCommit();
    } else if (e.key === "Escape") {
      e.preventDefault();
      setLocalValue(value);
      onCancel();
    }
  };
const getParamHint = (fn: string, index: number | null) => {
  if (index === null) return "";

  const paramHints: Record<string, string[]> = {
    IF: ["condition", "value_if_true", "value_if_false"],
    ROUND: ["number", "num_digits"],
    POWER: ["base", "exponent"],
    SUM: ["number1", "number2", "..."],
    AVERAGE: ["number1", "number2", "..."],
    MIN: ["number1", "number2", "..."],
    MAX: ["number1", "number2", "..."],
    COUNT: ["value1", "value2", "..."],
  };

  const params = paramHints[fn];
  if (!params) return "";

  const safeIndex = Math.min(index, params.length - 1);
  return `→ ${params[safeIndex]}`;
};

  return (
  <div style={{ position: "absolute", top: 0, left: 0, zIndex: 9999 }}>
    <input
      ref={inputRef}
      value={localValue}
      onChange={handleInput}
      onKeyDown={handleKeyDown}
      className="formula-input"
      style={{
        display: "block",
        padding: "4px 8px",
        fontSize: "14px",
        border: "1px solid #ccc",
        borderRadius: "4px",
        width: "250px",
      }}
    />
    {currentFunction && (
      <div style={{
        marginTop: "4px",
        fontSize: "12px",
        background: "#f0f0f0",
        padding: "6px 10px",
        borderRadius: "4px",
        color: "#333",
        boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
        width: "fit-content",
        maxWidth: "300px"
      }}>
        Editing: <strong>{currentFunction}</strong> {" "}
        {getParamHint(currentFunction, currentArgIndex)}
      </div>
    )}
  </div>
);

};

const ContactGrid: React.FC = () => {
  const [showDropdown, setShowDropdown] = useState(false);
  const [selectedRanges, setSelectedRanges] = useState<Rectangle[]>([]);

  const [columnWidths, setColumnWidths] = useState<{ [key: number]: number }>({});
  const [clipboardData, setClipboardData] = useState<any[][] | null>(null);

  const [undoStack, setUndoStack] = useState<SheetData[]>([]);
  const [redoStack, setRedoStack] = useState<SheetData[]>([]);

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

  // New states for sorting and filtering
  const [sortColumnIndex, setSortColumnIndex] = useState<number | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc' | null>(null);
  const [showFilterRow, setShowFilterRow] = useState<boolean>(false);
  const [columnFilters, setColumnFilters] = useState<{ [key: number]: string }>({});

  // New states for Conditional Formatting and Freeze Panes
  const [conditionalFormattingRules, setConditionalFormattingRules] = useState<ConditionalFormattingRule[]>([]);
  const [showConditionalFormattingModal, setShowConditionalFormattingModal] = useState(false);
  const [cfType, setCfType] = useState<ConditionalFormattingRule['type']>('greaterThan');
  const [cfValue1, setCfValue1] = useState<string>('');
  const [cfValue2, setCfValue2] = useState<string>('');
  const [cfBgColor, setCfBgColor] = useState<string>('#FFFF00'); // Default to yellow
  const [cfTextColor, setCfTextColor] = useState<string>('#000000'); // Default to black

  const [frozenRows, setFrozenRows] = useState<number>(0);
  const [frozenColumns, setFrozenColumns] = useState<number>(0);

  // State to force DataEditor re-render on data/sort/filter changes
  const [dataUpdateKey, setDataUpdateKey] = useState(0);

  // States for Find feature
  const [showFindModal, setShowFindModal] = useState(false);
  const [findSearchTerm, setFindSearchTerm] = useState('');
  const [findReplaceTerm, setFindReplaceTerm] = useState(''); // New state for replace
  const [findCurrentMatch, setFindCurrentMatch] = useState<Item | null>(null);
  const [findMatches, setFindMatches] = useState<Item[]>([]);
  const [findMatchIndex, setFindMatchIndex] = useState(0);

  // Data Validation states
  const [showDataValidationModal, setShowDataValidationModal] = useState(false);
  const [dvType, setDvType] = useState<'number' | 'text' | 'date' | 'list'>('number'); // Changed type here
  const [dvOperator, setDvOperator] = useState<'greaterThan' | 'lessThan' | 'equalTo' | 'notEqualTo' | 'between' | 'textContains' | 'startsWith' | 'endsWith'>('greaterThan'); // Changed type here
  const [dvValue1, setDvValue1] = useState<string>('');
  const [dvValue2, setDvValue2] = useState<string>('');
  const [dvSourceRange, setDvSourceRange] = useState<string>('');

  // Named Ranges states
  const [showNamedRangesModal, setShowNamedRangesModal] = useState(false);
  const [namedRanges, setNamedRanges] = useState<NamedRange[]>([]);
  const [newNamedRangeName, setNewNamedRangeName] = useState<string>('');
  const [newNamedRangeRef, setNewNamedRangeRef] = useState<string>('');
  const [editingNamedRangeId, setEditingNamedRangeId] = useState<string | null>(null);

  // Dark Mode State
  const [isDarkMode, setIsDarkMode] = useState(false);
  const currentTheme = isDarkMode ? darkTheme : lightTheme;


  const activeCell = useRef<Item | null>(null);
  const selecting = useRef<Rectangle | null>(null);
  const currentSheetData = sheets[activeSheet];

  // Memoized data that is filtered and then sorted
  const getDisplayedData = useMemo(() => {
    let dataWithOriginalIndex = currentSheetData.map((row, originalRowIndex) => ({
      originalRowIndex,
      data: row,
    }));

    // Apply filters
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

    // Apply sorting
    if (sortColumnIndex !== null && sortDirection !== null) {
      const sortedData = [...dataWithOriginalIndex]; // Create a shallow copy to avoid mutating original
      sortedData.sort((itemA, itemB) => {
        const valueA = itemA.data[sortColumnIndex]?.value || "";
        const valueB = itemB.data[sortColumnIndex]?.value || "";

        // Attempt to convert to number for numeric sorting
        const numA = parseFloat(valueA);
        const numB = parseFloat(valueB);

        if (!isNaN(numA) && !isNaN(numB)) {
          return sortDirection === 'asc' ? numA - numB : numB - numA;
        } else {
          // Fallback to string comparison
          return sortDirection === 'asc'
            ? valueA.localeCompare(valueB)
            : valueB.localeCompare(valueA);
        }
      });
      // Pad sorted data with empty rows if it became shorter than NUM_ROWS
      while (sortedData.length < NUM_ROWS) {
        sortedData.push({ originalRowIndex: -1, data: Array(NUM_COLUMNS).fill({ value: "" }) }); // Use -1 for padded rows
      }
      return sortedData;
    }
    // Pad original data if it's shorter than NUM_ROWS
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
  const getCellContent = useCallback(
    ([col, row]: Item): GridCell => {
      // Use getDisplayedData for cell content
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
      const cell = displayedRowData.data[col] ?? { value: "" };
      // Pass namedRanges to evaluateFormula
      let displayValue = cell.formula ? evaluateFormula(cell.formula, currentSheetData, namedRanges) : cell.value;
      // If displayValue is an object, convert it to a string for display
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
        displayData: displayValue,
        data: cell.formula ?? cell.value,
        themeOverride: inHighlight
        ? { bgCell: currentTheme.cellHighlightBg, borderColor: currentTheme.cellHighlightBorder }
        : undefined,
      };
    },
    [getDisplayedData, highlightRange, currentTheme, currentSheetData, namedRanges] // Add namedRanges to dependencies
  );
 const onCellEdited = useCallback(
  ([col, row]: Item, newValue: EditableGridCell) => {
    console.log(
      `onCellEdited triggered for cell [${col}, ${row}] with value: ${
        newValue.kind === GridCellKind.Text ? newValue.data : "Non-text"
      }`
    );

    if (newValue.kind !== GridCellKind.Text) return;

    const text = newValue.data;

    // Get the original row index from the displayed data
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) {
      console.warn("Attempted to edit a non-existent or padded row.");
      return;
    }

    // --- Data Validation Check ---
    const cellToValidate = sheets[activeSheet][originalRowIndex]?.[col];
    const validationRule = cellToValidate?.dataValidation;
    if (validationRule) {
      let isValid = true;
      const numericValue = parseFloat(text);

      switch (validationRule.type) {
        case 'number':
          if (isNaN(numericValue)) {
            isValid = false;
          } else if (validationRule.operator === 'greaterThan' && numericValue <= (validationRule.value1 as number)) {
            isValid = false;
          } else if (validationRule.operator === 'lessThan' && numericValue >= (validationRule.value1 as number)) {
            isValid = false;
          } else if (validationRule.operator === 'equalTo' && numericValue !== (validationRule.value1 as number)) {
            isValid = false;
          } else if (validationRule.operator === 'notEqualTo' && numericValue === (validationRule.value1 as number)) {
            isValid = false;
          } else if (validationRule.operator === 'between' && (numericValue < (validationRule.value1 as number) || numericValue > (validationRule.value2 as number))) {
            isValid = false;
          }
          break;
        case 'text':
          const lowerCaseText = text.toLowerCase();
          if (validationRule.operator === 'textContains' && !(lowerCaseText.includes(String(validationRule.value1).toLowerCase()))) {
            isValid = false;
          } else if (validationRule.operator === 'startsWith' && !(lowerCaseText.startsWith(String(validationRule.value1).toLowerCase()))) {
            isValid = false;
          } else if (validationRule.operator === 'endsWith' && !(lowerCaseText.endsWith(String(validationRule.value1).toLowerCase()))) {
            isValid = false;
          }
          break;
        case 'list':
          if (validationRule.sourceRange) {
            const { x, y, width, height } = validationRule.sourceRange;
            const allowedValues: string[] = [];
            for (let r = y; r < y + height; r++) {
              for (let c = x; c < x + width; c++) {
                const sourceCell = sheets[activeSheet]?.[r]?.[c];
                if (sourceCell) {
                  allowedValues.push(sourceCell.value);
                }
              }
            }
            if (!allowedValues.includes(text)) {
              isValid = false;
              alert(`Invalid input. Please choose from: ${allowedValues.join(', ')}`);
            }
          }
          break;
        // Add more validation types (date, etc.) here
      }

      if (!isValid) {
        alert(`Invalid data for this cell based on validation rules: ${text}`);
        console.warn(`Data validation failed for cell [${col}, ${row}] with value: ${text}`);
        return; // Prevent update if validation fails
      }
    }
    // --- End Data Validation Check ---

    socket.emit("cell-edit", {
      sheet: activeSheet,
      row: originalRowIndex, // Use original row index for socket
      col,
      value: text,
    });

    if (text.startsWith("=")) {
      const funcMatch = text.match(/^=(\w+)\(([^)]*)\)$/i);
      if (funcMatch) {
        const [, func, args] = funcMatch;
        const funcUpper = func.toUpperCase();
        const argCount = args
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s !== "").length;

        let formulaValid = true;
        if (funcUpper === "IF") {
          if (argCount !== 3) { setFormulaError(`${func} requires 3 arguments`); formulaValid = false; }
        } else if (funcUpper === "ROUND" || funcUpper === "POWER") {
          if (argCount !== 2) { setFormulaError(`${func} requires 2 arguments`); formulaValid = false; }
        } else if (funcUpper === "ABS" || funcUpper === "SQRT") {
          if (argCount !== 1) { setFormulaError(`${func} requires 1 argument`); formulaValid = false; }
          else if (args.includes(":")) { setFormulaError(`${func} requires a single cell, not a range`); formulaValid = false; }
        } else if (
          ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "PRODUCT"].includes(funcUpper)
        ) {
          if (argCount < 1) { setFormulaError(`${func} requires at least 1 argument`); formulaValid = false; }
        } else {
          setFormulaError(null); // Clear error if function is not recognized or has no specific arg count
        }

        if (!formulaValid) {
          console.warn(`Formula validation failed for: ${text}`);
          return; // Prevent update if formula validation fails
        }
      }
    }
    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

    if (text.startsWith("=")) {
      currentSheetCopy[originalRowIndex][col] = { // Use originalRowIndex here
        formula: text,
        value: evaluateFormula(text, currentSheetCopy, namedRanges), // Pass namedRanges
      };
    } else {
      currentSheetCopy[originalRowIndex][col] = { value: text }; // Use originalRowIndex here
    }

    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
      setFormulaInput("");
      setHighlightRange(null);
      setShowSuggestions(false);
      setFormulaError(null);
    },
    [activeSheet,sheets, getDisplayedData, namedRanges] // Add namedRanges to dependencies
  );
  const onFillPattern = useCallback(
  ({ patternSource, fillDestination }: FillPatternEventArgs) => {
    const sourceCol = patternSource.x;
    const sourceRow = patternSource.y;
    // Need to get the original row index for the source cell
    const originalSourceRowIndex = getDisplayedData[sourceRow]?.originalRowIndex;
    if (originalSourceRowIndex === undefined || originalSourceRowIndex === -1) return;

    const sourceCell = currentSheetData[originalSourceRowIndex]?.[sourceCol];
    if (!sourceCell) return;

    const fillValue = sourceCell.value;
    const fillFormula = sourceCell.formula;

    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

    for (let r = fillDestination.y; r < fillDestination.y + fillDestination.height; r++) {
      // Get original row index for the destination
      const originalDestRowIndex = getDisplayedData[r]?.originalRowIndex;
      if (originalDestRowIndex === undefined || originalDestRowIndex === -1) continue;

      for (let c = fillDestination.x; c < fillDestination.x + fillDestination.width; c++) {
        const isSourceCell = c >= patternSource.x && c < patternSource.x + patternSource.width &&
                             r >= patternSource.y && r < patternSource.y + patternSource.height;
        if (!isSourceCell) {
          currentSheetCopy[originalDestRowIndex][c] = fillFormula // Use originalDestRowIndex
            ? { formula: fillFormula, value: evaluateFormula(fillFormula, currentSheetCopy, namedRanges) } // Pass namedRanges
            : { value: fillValue };
        }
      }
    }

    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
  },
  [activeSheet, currentSheetData, getDisplayedData, namedRanges] // Add namedRanges to dependencies
);
  const onFinishSelecting = useCallback(() => {
  if (!activeCell.current || !selecting.current) return;
  const [col, row] = activeCell.current;

  // Get original row index for the active cell
  const originalActiveRowIndex = getDisplayedData[row]?.originalRowIndex;
  if (originalActiveRowIndex === undefined || originalActiveRowIndex === -1) return;


  const topLeft = getCellName(selecting.current.x, selecting.current.y);
  const bottomRight = getCellName(
    selecting.current.x + selecting.current.width - 1,
    selecting.current.y + selecting.current.height - 1
  );
  const formula = `=SUM(${topLeft}:${bottomRight})`;
  const value = evaluateFormula(formula, currentSheetData, namedRanges); // Pass namedRanges

  const updatedSheets = { ...sheets };
  const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);
  currentSheetCopy[originalActiveRowIndex][col] = { formula, value }; // Use originalActiveRowIndex
  updatedSheets[activeSheet] = currentSheetCopy;
  pushToUndoStack(updatedSheets);

  setFormulaInput("");
  setHighlightRange(null);
  setShowSuggestions(false);
  setFormulaError(null);
}, [activeSheet, currentSheetData, getDisplayedData, namedRanges]); // Add namedRanges to dependencies
  const handleFormulaChange = (val: string) => {
    console.log("Formula input changed:", val);
    setFormulaInput(val);
    updateSuggestions(val);
    setFormulaError(null);
    setHighlightRange(null);
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
            console.log(`Highlighting range: [${x},${y}] to [${x + width -1},${y + height -1}]`);

            setDataUpdateKey(prev => prev + 1); 
            gridRef.current?.scrollTo(y,x);
            break;
          }
        }
      }
    }
  };
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
    console.log(`Added new sheet: ${newSheetName}`);
    setDataUpdateKey(prev => prev + 1); // Data structure changed
  };

  const handleDeleteSheet = (sheetToDelete: string) => {
    if (Object.keys(sheets).length === 1) {
      alert("Cannot delete the last sheet!");
      console.warn("Attempted to delete the last sheet.");
      return;
    }
    setSheets((prevSheets) => {
      const updatedSheets = { ...prevSheets };
      delete updatedSheets[sheetToDelete];
      return updatedSheets;
    });
    if (activeSheet === sheetToDelete) {
      setActiveSheet(Object.keys(sheets)[0]);
      console.log(`Switched active sheet to: ${Object.keys(sheets)[0]}`);
    }
    setDataUpdateKey(prev => prev + 1); // Data structure changed
  };

  const handleEditSheetName = (sheetName: string) => {
    console.log(`Editing sheet name: ${sheetName}`);
    setEditingSheetName(sheetName);
    setNewSheetName(sheetName);
  };

  const handleSaveSheetName = (oldName: string) => {
    console.log(`Saving sheet name. Old: ${oldName}, New: ${newSheetName}`);
    if (newSheetName.trim() === "" || newSheetName === oldName) {
      console.log("Sheet name not changed or empty. Aborting save.");
      setEditingSheetName(null);
      return;
    }
    if (sheets[newSheetName]) {
      alert("Sheet name already exists!");
      console.warn(`Sheet name "${newSheetName}" already exists.`);
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
      console.log(`Active sheet renamed to: ${newSheetName}`);
    }
    setEditingSheetName(null);
    setDataUpdateKey(prev => prev + 1); // Data structure changed
  };

  const handleSheetNameInputKeyDown = (e: React.KeyboardEvent, oldName: string) => {
    if (e.key === 'Enter') {
      console.log("Enter key pressed on sheet name input.");
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
      console.log(`Attempting to save sheet "${saveLoadSheetName}"...`);
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
        console.log('Save successful:', result);
        alert(`Spreadsheet "${result.sheet.name}" saved successfully!`);
        setDataUpdateKey(prev => prev + 1); // Data might change on save (e.g. if backend modifies it)
      } else {
        const errorData = await response.json();
        console.error('Failed to save data:', errorData);
        alert(`Failed to save data: ${errorData.message || 'Unknown error'}`);
      }
    } catch (error) {
      console.error('Error during save operation:', error);
      alert('An error occurred while trying to save data. Check backend server.');
    }
  };
useEffect(() => {
  const handleClick = () => setContextMenu(null);
  window.addEventListener("click", handleClick);
  return () => window.removeEventListener("click", handleClick);
}, []);



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
        setDataUpdateKey(prev => prev + 1); // Data changed
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

  // Build worksheet values only
  for (let r = 0; r <= maxRow; r++) {
    worksheetData[r] = [];
    for (let c = 0; c <= maxCol; c++) {
      worksheetData[r][c] = data[r][c]?.value ?? "";
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(worksheetData);

  // === Column Width Setup ===
  ws["!cols"] = Array.from({ length: maxCol + 1 }, (_, i) => {
    const col = columns[i];
    return {
      wpx:  100 // width in pixels
    };
  });

  // === Cell styles ===
  for (let r = 0; r <= maxRow; r++) {
    for (let c = 0; c <= maxCol; c++) {
      const cellData = data[r][c];
      const cellRef = XLSX.utils.encode_cell({ r, c });

      if (!ws[cellRef]) continue;

      const style: any = {};

      // Font size & color
      if (cellData?.fontSize || cellData?.textColor) {
        style.font = {};
        if (cellData.fontSize) style.font.sz = cellData.fontSize;
        if (cellData.textColor)
          style.font.color = {
            rgb: cellData.textColor.replace("#", "").toUpperCase(),
          };
      }

      // Background color
      if (cellData?.background) {
        style.fill = {
          fgColor: { rgb: cellData.background.replace("#", "").toUpperCase() },
        };
      }

      if (Object.keys(style).length > 0) {
        ws[cellRef].s = style;
      }
    }
  }

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
        setDataUpdateKey(prev => prev + 1); // Data changed
      } catch (err) {
        alert("Failed to import Excel file: Invalid format.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Function to export to CSV
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
        setDataUpdateKey(prev => prev + 1); // Data changed

      } catch (err) {
        console.error("Error importing CSV file:", err);
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
      ctx.fillStyle = "rgba(30, 144, 255, 0.2)"; // Light blue background
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);

      ctx.strokeStyle = "dodgerblue"; // Blue border
      ctx.lineWidth = 2;
      ctx.strokeRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
    }
  }
  if (cell.kind !== GridCellKind.Text) return false;

  // Use getDisplayedData for cell data
  const displayedRowData = getDisplayedData[row];
  if (!displayedRowData) return false; // Should not happen if rows prop is correct
  const cellData = displayedRowData.data[col];

  let alignment: "left" | "center" | "right" = cellData?.alignment || "left";
  let fontSize = cellData?.fontSize || 12;
  let isBold = cellData?.bold ? "bold" : "normal";
  let isItalic = cellData?.italic ? "italic" : "normal";
  let isUnderline = cellData?.underline;
  let isStrikethrough = cellData?.strikethrough;
  let textColor = cellData?.textColor || currentTheme.text;
  let bgColor = cellData?.bgColor || currentTheme.bg; // Use theme background color
  let borderColor = cellData?.borderColor || currentTheme.border; // Use theme border color
  let fontFamily = cellData?.fontFamily || 'sans-serif'; // Get font family, default to sans-serif
  let text = String(cell.displayData ?? "");

  // Apply conditional formatting rules
  for (const rule of conditionalFormattingRules) {
    const { range, type, value1, value2, style } = rule;
    // Check against the original row index for conditional formatting
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

  // UI Enhancement: Draw borders more subtly, similar to Excel
  ctx.strokeStyle = borderColor; // Use borderColor from cellData or theme
  ctx.lineWidth = 0.5; // Finer border line
  ctx.beginPath();
  ctx.moveTo(rect.x, rect.y + rect.height); // Bottom border
  ctx.lineTo(rect.x + rect.width, rect.y + rect.height);
  ctx.moveTo(rect.x + rect.width, rect.y); // Right border
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
    const textHeight = fontSize; // Approximate text height
    const underlineY = rect.y + rect.height / 2 + textHeight / 2 + 1; // Position below text
    ctx.beginPath();
    ctx.strokeStyle = textColor;
    ctx.lineWidth = 1;
    ctx.moveTo(x, underlineY);
    ctx.lineTo(x + textWidth, underlineY);
    ctx.stroke();
  }
  if (isStrikethrough) {
    const textWidth = textMetrics.width;
    const textHeight = fontSize;
    const strikethroughY = rect.y + rect.height / 2;
    ctx.beginPath();
    ctx.strokeStyle = textColor;
    ctx.lineWidth = 1;
    ctx.moveTo(x, strikethroughY);
    ctx.lineTo(x + textWidth, strikethroughY);
    ctx.stroke();
  }

  return true;
}, [getDisplayedData, conditionalFormattingRules, currentTheme]); // Add conditionalFormattingRules and currentTheme to dependencies

const toggleStyle = (type: "bold" | "italic") => {
  if (!selection.current) return;
  applyStyleToRange(type, undefined); // This will toggle the style
};

const applyCellColor = (type: "textColor" | "bgColor" | "borderColor", color: string) => {
  if (!selection.current) return;
  applyStyleToRange(type, color);
};
const pushToUndoStack = (newSheets: SheetData) => {
  setUndoStack(prev => [...prev, sheets]);
  setRedoStack([]);
  setSheets(newSheets);
  setDataUpdateKey(prev => prev + 1); // Increment key on data change
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

const [contextMenu, setContextMenu] = useState<{
  visible: boolean;
  x: number;
  y: number;
  targetRow: number;
} | null>(null);

const handleUndo = () => {
  if (undoStack.length === 0) return;
  const previous = undoStack[undoStack.length - 1];
  setRedoStack(prev => [...prev, sheets]);
  setUndoStack(prev => prev.slice(0, -1));
  setSheets(previous);
  setDataUpdateKey(prev => prev + 1); // Data changed
};

const handleRedo = () => {
  if (redoStack.length === 0) return;
  const next = redoStack[redoStack.length - 1];
  setUndoStack(prev => [...prev, sheets]);
  setRedoStack(prev => prev.slice(0, -1));
  setSheets(next);
  setDataUpdateKey(prev => prev + 1); // Data changed
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
    setDataUpdateKey(prev => prev + 1); // Data changed
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
    id: String(i), // ✅ converted to string
  })),
[sheets, columnWidths]);

const onColumnResize = useCallback((column: GridColumn, newSize: number, colIndex: number) => {
  setColumnWidths(prev => ({ ...prev, [colIndex]: newSize }));
  setDataUpdateKey(prev => prev + 1); // Column width change might affect layout
}, []);
 const applyStyleToRange = (key: string, value: any) => {
  if (!selection?.current) return;
  const { x, y, width, height } = selection.current.range;
  const currentSheet = sheets[activeSheet];

  let allHaveSame = true;
  for (let row = y; row < y + height; row++) {
    // Get original row index for the selected range
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) continue;

    for (let col = x; col < x + width; col++) {
      const cell = currentSheet?.[originalRowIndex]?.[col]; // Use originalRowIndex
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
      const cell = copy[originalRowIndex]?.[col] ?? { value: "" }; // Use originalRowIndex
      // Toggle the style: if all selected cells have the style, remove it; otherwise, apply it.
      copy[originalRowIndex][col] = { ...cell, [key]: allHaveSame ? undefined : value }; // Use originalRowIndex
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

    // ✅ Ctrl+C: Copy (value + background)
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

    // ✅ Ctrl+V: Paste (value + background)
    if (e.ctrlKey && e.key === "v") {
      e.preventDefault();
      navigator.clipboard.readText().then((text) => {
        try {
          const copiedData = JSON.parse(text);
          const newSheets = JSON.parse(JSON.stringify(sheets)); // deep copy

          // Ensure startX and startY are captured within the closure
          const currentRange = selection.current?.range;
          if (!currentRange) return; // Should not happen if outer check passes

          const pasteStartX = Math.min(currentRange.x, currentRange.x + currentRange.width - 1);
          const pasteStartY = Math.min(currentRange.y, currentRange.y + currentRange.height - 1);

          for (let rowOffset = 0; rowOffset < copiedData.length; rowOffset++) {
            const targetRowDisplayIndex = pasteStartY + rowOffset;
            const originalTargetRowIndex = getDisplayedData[targetRowDisplayIndex]?.originalRowIndex;
            if (originalTargetRowIndex === undefined || originalTargetRowIndex === -1) continue;

            for (let colOffset = 0; colOffset < copiedData[rowOffset].length; colOffset++) {
              const targetCol = pasteStartX + colOffset;
              if (newSheets[activeSheet][originalTargetRowIndex] && newSheets[activeSheet][originalTargetRowIndex][targetCol]) {
                newSheets[activeSheet][originalTargetRowIndex][targetCol] = {
                  fontSize: 14,
                  bold: false,
                  italic: false,
                  underline: false,
                  strikethrough: false, // Ensure default for strikethrough
                  alignment: "left",
                  ...copiedData[rowOffset][colOffset], // paste value + background
                };
              }
            }
          }
          pushToUndoStack(sheets);
          setSheets(newSheets);
          setDataUpdateKey(prev => prev + 1); // Data changed
        } catch (err) {
          console.warn("Invalid clipboard format");
        }
      });
    }

    // ✅ Ctrl+X: Cut (copy + clear value + background)
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

          // Clear the cell
          newSheets[activeSheet][originalRowIndex][col] = {
            value: "",
            background: "",
            fontSize: 14,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false, // Ensure default for strikethrough
            alignment: "left",
          };
        }
        copiedData.push(rowData);
      }

      navigator.clipboard.writeText(JSON.stringify(copiedData));
      pushToUndoStack(sheets);
      setSheets(newSheets);
      setDataUpdateKey(prev => prev + 1); // Data changed
    }
  };
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
    setDataUpdateKey(prev => prev + 1); // Data changed
  };
  socket.on("cell-edit", handleCellEdit);
  return () => {
    socket.off("cell-edit", handleCellEdit);
  };
}, [activeSheet]);

useEffect(() => {
  const handleClick = () => {
    // Only append cell reference if formula input is active and a range is being selected
    if (
      formulaInput.startsWith("=") &&
      selecting.current &&
      activeCell.current && // Ensure an active cell is selected
      (selection.columns.length > 0 || selection.rows.length > 0 || selection.current?.range) // Corrected access to selection.columns and selection.rows
    ) {
      const ref = getCellName(selecting.current.x, selecting.current.y);
      setFormulaInput((prev) => {
        // Prevent appending if the reference is already part of the formula or if it's the start of a new argument
        if (prev.includes(ref) && !prev.endsWith("(") && !prev.endsWith(",")) return prev;
        const insert = prev.endsWith("(") || prev.endsWith(",") ? ref : `,${ref}`;
        return prev + insert;
      });
    }
  };

  window.addEventListener("mousedown", handleClick);
  return () => window.removeEventListener("mousedown", handleClick);
}, [formulaInput, selection]); // Add selection to dependencies


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

// New state for dropdowns and input modals
const [showFileDropdown, setShowFileDropdown] = useState(false);
const [showEditDropdown, setShowEditDropdown] = useState(false);
const [showInsertDropdown, setShowInsertDropdown] = useState(false);
const [showFormatDropdown, setShowFormatDropdown] = useState(false);
const [showDataDropdown, setShowDataDropdown] = useState(false); // New state for Data dropdown
const [showViewDropdown, setShowViewDropdown] = useState(false); // New state for View dropdown

const [showImportDropdown, setShowImportDropdown] = useState(false);
const [showExportDropdown, setShowExportDropdown] = useState(false);
const [editingCell, setEditingCell] = useState<Item | null>(null);

const [showLinkInput, setShowLinkInput] = useState(false);
const [linkValue, setLinkValue] = useState('');
const [showCommentInput, setShowCommentInput] = useState(false);
const [commentValue, setCommentValue] = useState('');

// Functions for Edit menu actions
const handleInsertRowAbove = () => {
  if (!selection.current) {
    alert("Please select a cell to insert a row.");
    return;
  }
  // Get original row index for the selected row
  const originalRowIndex = getDisplayedData[selection.current.range.y]?.originalRowIndex;
  if (originalRowIndex === undefined || originalRowIndex === -1) {
    alert("Cannot insert row at this position.");
    return;
  }
  insertRowAt(originalRowIndex); // Use original row index for insertion
  setShowEditDropdown(false);
};

const handleInsertRowBelow = () => {
  if (!selection.current) {
    alert("Please select a cell to insert a row.");
    return;
  }
  // Get original row index for the selected row
  const originalRowIndex = getDisplayedData[selection.current.range.y]?.originalRowIndex;
  if (originalRowIndex === undefined || originalRowIndex === -1) {
    alert("Cannot insert row at this position.");
    return;
  }
  insertRowAt(originalRowIndex + selection.current.range.height); // Use original row index for insertion
  setShowEditDropdown(false);
};

const handleDeleteSelectedRows = () => {
  if (selection.rows.length === 0) { // Check length directly
    alert("Please select rows to delete.");
    return;
  }
  // Get all selected row indices
  const selectedRowDisplayIndices: number[] = Array.from(selection.rows); // Use Array.from to iterate

  // Map display indices to original indices
  const selectedOriginalRowIndices = selectedRowDisplayIndices
    .map(displayIndex => getDisplayedData[displayIndex]?.originalRowIndex)
    .filter(index => index !== undefined && index !== -1) as number[];

  // Sort in descending order to avoid index issues during deletion
  selectedOriginalRowIndices.sort((a, b) => b - a);

  const updatedSheets = { ...sheets };
  let currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

  selectedOriginalRowIndices.forEach(originalRowIndex => {
    currentSheetCopy.splice(originalRowIndex, 1);
  });

  // Add empty rows at the bottom to maintain NUM_ROWS
  while (currentSheetCopy.length < NUM_ROWS) {
    currentSheetCopy.push(Array(NUM_COLUMNS).fill({ value: "" }));
  }

  updatedSheets[activeSheet] = currentSheetCopy;
  pushToUndoStack(updatedSheets);
  setShowEditDropdown(false);
  setSelection({ // Clear selection after deletion
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
  selectedColIndices.sort((a, b) => b - a); // Sort descending to avoid index issues

  selectedColIndices.forEach(colIndex => {
    deleteColumns(colIndex);
  });
  setShowEditDropdown(false);
  setSelection({ // Clear selection after deletion
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  });
};


// Functions for Insert menu actions
const handleInsertLink = () => {
  if (!activeCell.current) {
    alert("Please select a cell to insert a link.");
    return;
  }
  setShowLinkInput(true);
  setShowInsertDropdown(false); // Close dropdown after selection
};

const handleInsertComment = () => {
  if (!activeCell.current) {
    alert("Please select a cell to insert a comment.");
    return;
  }
  setShowCommentInput(true);
  setShowInsertDropdown(false); // Close dropdown after selection
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

// New functions for Sorting
const handleSort = (direction: 'asc' | 'desc') => {
  if (activeCell.current === null) {
    alert("Please select a cell in the column you want to sort.");
    return;
  }
  const colIndex = activeCell.current[0];
  setSortColumnIndex(colIndex);
  setSortDirection(direction);
  setShowDataDropdown(false); // Close dropdown
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for sort
};

// New functions for Filtering
const toggleFilterRow = () => {
  setShowFilterRow(prev => !prev);
  // Clear filters when toggling off
  if (showFilterRow) {
    setColumnFilters({});
  }
  setShowDataDropdown(false); // Close dropdown
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for filter
};

const handleColumnFilterChange = (colIndex: number, value: string) => {
  setColumnFilters(prev => ({
    ...prev,
    [colIndex]: value,
  }));
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for filter
};

const clearAllFilters = () => {
  setColumnFilters({});
  setSortColumnIndex(null); // Also clear sort when clearing filters
  setSortDirection(null);
  setShowFilterRow(false);
  setShowDataDropdown(false); // Close dropdown
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for filter
};

// Conditional Formatting Functions
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
    id: Date.now().toString(), // Simple unique ID
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
  setCfType('greaterThan'); // Reset to default
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for CF
};

const clearConditionalFormatting = () => {
  setConditionalFormattingRules([]);
  setShowFormatDropdown(false);
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for CF
};

// Freeze Panes Functions
const handleFreezeRows = (count: number) => {
  if (activeCell.current === null && count > 0) {
    alert("Please select a cell to freeze rows up to.");
    return;
  }
  setFrozenRows(count === -1 ? activeCell.current![1] + 1 : count);
  setFrozenColumns(0); // Unfreeze columns when freezing rows
  setShowViewDropdown(false);
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for freeze
};

const handleFreezeColumns = (count: number) => {
  if (activeCell.current === null && count > 0) {
    alert("Please select a cell to freeze columns up to.");
    return;
  }
  setFrozenColumns(count === -1 ? activeCell.current![0] + 1 : count);
  setFrozenRows(0); // Unfreeze rows when freezing columns
  setShowViewDropdown(false);
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for freeze
};

const handleUnfreezePanes = () => {
  setFrozenRows(0);
  setFrozenColumns(0);
  setShowViewDropdown(false);
  setDataUpdateKey(prev => prev + 1); // Force re-render of DataEditor for unfreeze
};

// Find feature functions
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
    const rowData = getDisplayedData[r].data; // Access the actual row data
    for (let c = 0; c < rowData.length; c++) {
      const cell = rowData[c];
      const cellValue = cell?.value?.toString().toLowerCase() || '';
      if (cellValue.includes(lowerCaseSearchTerm)) {
        matches.push([c, r]); // Store display coordinates
      }
    }
  }
  setFindMatches(matches);
  if (matches.length > 0) {
    setFindMatchIndex(0);
    setFindCurrentMatch(matches[0]);
    // Set selection to the first match
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
    // alert("No matches found."); // Removed alert to avoid blocking UI
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
    // Re-run find to update matches and selection
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
        currentSheetCopy[r][c] = { ...cell, value: newValue, formula: undefined }; // Clear formula if value changes directly
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

// Data Validation Functions
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
  // Reset DV states
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
        const { dataValidation, ...rest } = cell; // Remove dataValidation property
        currentSheetCopy[originalRowIndex][col] = rest;
      }
    }
  }
  updatedSheets[activeSheet] = currentSheetCopy;
  pushToUndoStack(updatedSheets);
  setShowDataDropdown(false);
  setDataUpdateKey(prev => prev + 1);
};

// Named Ranges Functions
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
  setDataUpdateKey(prev => prev + 1); // Formulas might need re-evaluation
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
  setDataUpdateKey(prev => prev + 1); // Formulas might need re-evaluation
};


  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', fontFamily: 'Roboto, sans-serif', color: currentTheme.text, backgroundColor: currentTheme.bg }}>
      {/* Top Bar - Google Sheets like */}
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
        {/* Left Section: Logo, Title, Menu Bar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          {/* Logo Placeholder */}
          <span style={{ fontSize: '24px', color: currentTheme.activeTabBorder, fontWeight: 'bold' }}>
            Sheets
          </span>
          {/* Spreadsheet Name Input */}
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

          {/* Menu Bar */}
          <div style={{ display: 'flex', gap: '2px', marginLeft: '20px' }}>
            {/* File Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowFileDropdown(true)}
              onMouseLeave={() => { setShowFileDropdown(false); setShowImportDropdown(false); setShowExportDropdown(false); }}
            >
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>File</button>
              {showFileDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={saveSheetData} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Save</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>

                  {/* Import Sub-menu */}
                  <div
                    style={{ position: "relative" }}
                    onMouseEnter={() => setShowImportDropdown(true)}
                    onMouseLeave={() => setShowImportDropdown(false)}
                  >
                    <button style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Import &#x25B6;</button> {/* Right arrow for sub-menu */}
                    {showImportDropdown && (
                      <div style={{...subMenuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>
                          Import XLSX
                          <input type="file" accept=".xlsx" onChange={importFromExcel} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>
                          Import JSON
                          <input type="file" accept="application/json" onChange={handleImportFromJSON} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>
                          Import CSV
                          <input type="file" accept=".csv" onChange={importFromCSV} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>
                          Import ODS
                          <input type="file" accept=".ods" onChange={handleODSImport} style={{ display: "none" }} />
                        </label>
                      </div>
                    )}
                  </div>

                  {/* Export Sub-menu */}
                  <div
                    style={{ position: "relative" }}
                    onMouseEnter={() => setShowExportDropdown(true)}
                    onMouseLeave={() => setShowExportDropdown(false)}
                  >
                    <button style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export &#x25B6;</button> {/* Right arrow for sub-menu */}
                    {showExportDropdown && (
                      <div style={{...subMenuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportXLSX} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as XLSX</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportToJSON} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as JSON</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={exportToCSV} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as CSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleODSExport} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as ODS</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportTSV} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as TSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleExportPDF} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Export as PDF</button>
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
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>Edit</button>
              {showEditDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleUndo} disabled={undoStack.length === 0} style={{...menuItem, opacity: undoStack.length === 0 ? 0.5 : 1, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Undo</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleRedo} disabled={redoStack.length === 0} style={{...menuItem, opacity: redoStack.length === 0 ? 0.5 : 1, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Redo</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertRowAbove} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Insert Row Above</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertRowBelow} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Insert Row Below</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertColumnLeft} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Insert Column Left</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertColumnRight} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Insert Column Right</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleDeleteSelectedRows} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Delete Selected Rows</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleDeleteSelectedColumns} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Delete Selected Columns</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => setShowFindModal(true)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Find and Replace...</button>
                </div>
              )}
            </div>

            {/* View Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowViewDropdown(true)}
              onMouseLeave={() => setShowViewDropdown(false)}
            >
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>View</button>
              {showViewDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(1)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze 1 row</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(2)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze 2 rows</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeRows(-1)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze up to current row</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(1)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze 1 column</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(2)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze 2 columns</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleFreezeColumns(-1)} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Freeze up to current column</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleUnfreezePanes} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>No frozen panes</button>
                </div>
              )}
            </div>

            {/* Insert Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowInsertDropdown(true)}
              onMouseLeave={() => setShowInsertDropdown(false)}
            >
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>Insert</button>
              {showInsertDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertLink} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Link</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleInsertComment} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Comment</button>
                </div>
              )}
            </div>

            {/* Format Menu - Now mostly moved to toolbar, keeping as placeholder for future */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowFormatDropdown(true)}
              onMouseLeave={() => setShowFormatDropdown(false)}
            >
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>Format</button>
              {showFormatDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  {/* These were moved to the main toolbar */}
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: currentTheme.textLight }}>Text Styles</div>
                  <select
                    onChange={(e) => applyStyleToRange("fontSize",parseInt(e.target.value))}
                    style={{...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px', color: currentTheme.text, backgroundColor: currentTheme.bg, border: `1px solid ${currentTheme.border}`}}
                  >
                    <option value="">Font Size</option>
                    {[10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].map(size => (
                      <option key={size} value={size}>{size}px</option>
                    ))}
                  </select>
                  <select
                    onChange={(e) => applyStyleToRange("fontFamily", e.target.value)}
                    style={{...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px', color: currentTheme.text, backgroundColor: currentTheme.bg, border: `1px solid ${currentTheme.border}`}}
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
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("bold",true)} style={{...topBarButtonStyle, fontWeight: "bold", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>B</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("italic",true)} style={{...topBarButtonStyle, fontStyle: "italic", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>I</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("underline",true)} style={{...topBarButtonStyle, textDecoration: "underline", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>U</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => applyStyleToRange("strikethrough",true)} style={{...topBarButtonStyle, textDecoration: "line-through", border: `1px solid ${currentTheme.border}`, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>S</button>
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
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleApplyConditionalFormatting} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Conditional Formatting...</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearConditionalFormatting} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Clear Conditional Formatting</button>
                </div>
              )}
            </div>

            {/* Data Menu */}
            <div
              style={{ position: "relative" }}
              onMouseEnter={() => setShowDataDropdown(true)}
              onMouseLeave={() => setShowDataDropdown(false)}
            >
              <button style={{...topBarButtonStyle, color: currentTheme.text}}>Data</button>
              {showDataDropdown && (
                <div style={{...menuDropdownStyle, backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, boxShadow: currentTheme.shadow}}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleSort('asc')} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Sort sheet A-Z (Current Column)</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={() => handleSort('desc')} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Sort sheet Z-A (Current Column)</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={toggleFilterRow} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Toggle Filter Row</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearAllFilters} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Clear All Filters</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleApplyDataValidation} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Data Validation...</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={clearDataValidation} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Clear Data Validation</button>
                  <div style={{ borderTop: `1px solid ${currentTheme.border}`, margin: '4px 0' }}></div>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuHoverBg)} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = currentTheme.menuBg)} onClick={handleManageNamedRanges} style={{...menuItem, color: currentTheme.text, backgroundColor: currentTheme.menuBg}}>Named Ranges...</button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Toggle Theme Button */}
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
              marginRight:'50px'
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
        {/* Undo/Redo Buttons */}
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

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div> {/* Divider */}

        {/* Font Family */}
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

        {/* Font Size */}
        <select
          onChange={(e) => applyStyleToRange("fontSize",parseInt(e.target.value))}
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

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div> {/* Divider */}

        {/* Font Styles (Bold, Italic, Underline, Strikethrough) */}
        <button
          onClick={() => applyStyleToRange("bold",true)}
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
          onClick={() => applyStyleToRange("italic",true)}
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
          onClick={() => applyStyleToRange("underline",true)}
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
          onClick={() => applyStyleToRange("strikethrough",true)}
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

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div> {/* Divider */}

        {/* Text Color */}
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
        {/* Background Color */}
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
        {/* Border Color */}
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

        <div style={{ width: '1px', height: '24px', backgroundColor: currentTheme.border, margin: '0 5px' }}></div> {/* Divider */}

        {/* Alignment */}
        <select
          onChange={(e) => applyStyleToRange("alignment",e.target.value as "left" | "center" | "right")}
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
        flexDirection: 'column', /* Changed to column to stack elements */
        alignItems: 'flex-start', /* Align items to the start for better stacking */
        flexShrink: 0,
        position: 'relative', /* Added for positioning suggestions */
      }}>
        <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}> {/* New wrapper for input and cell ref */}
          {/* Cell Reference Display */}
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
          {/* Formula Input */}
          <input
            value={formulaInput}
            onChange={(e) => handleFormulaChange(e.target.value)}
            onFocus={() => {
              updateSuggestions(formulaInput);
            }}
            onBlur={() => {
              // Delay hiding suggestions to allow click on suggestion
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
          <div style={{ color: "red", fontSize: "12px", paddingTop: "5px", paddingLeft: "80px" }}> {/* Adjusted padding */}
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
              width: "calc(100% - 100px)", /* Adjust width to match input */
              marginTop: "4px",
              top: 'calc(100% + 5px)', // Position below formula bar
              left: '80px', // Adjust based on cell reference width + padding
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
                onMouseDown={(e) => e.preventDefault()} /* Prevent blur on click */
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
          backgroundColor: currentTheme.bg2, // Light grey background for filter row
          borderBottom: `1px solid ${currentTheme.border}`,
          padding: '4px 16px',
          display: 'flex',
          alignItems: 'center',
          flexShrink: 0,
        }}>
          {/* Empty cell for row markers */}
          <div style={{ width: '60px', flexShrink: 0 }}></div>
          {columns.map((col, index) => (
            <input
              key={col.id}
              type="text"
              placeholder={`Filter ${col.title}`}
              value={columnFilters[index] || ''}
              onChange={(e) => handleColumnFilterChange(index, e.target.value)}
              style={{
                width: columnWidths[index] ?? 100, // Corrected access to width
                minWidth: '50px', // Ensure min-width for filter inputs
                padding: '4px 8px',
                border: `1px solid ${currentTheme.border}`,
                borderRadius: '4px',
                marginRight: '2px', // Small gap between filter inputs
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


      <div style={{ flexGrow: 1, overflow: 'hidden' }}
        /* Removed onContextMenu from here as requested */
      >
        <DataEditor
          columns={columns}
          rows={getDisplayedData.length} // Use the length of the displayed data
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
            setSelection({
              columns: CompactSelection.empty(),
              rows: CompactSelection.empty(),
              current: {
                cell,
                range: { x: cell[0], y: cell[1], width: 1, height: 1 },
                rangeStack: [],
              },
            });
            // When a cell is activated, update the formula input with its current content
            // Need to get the original row index for the active cell to fetch its data
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
          }}
          onFillPattern={onFillPattern}
          rowHeight={(row) => {
    // Need to get the original row index to determine height based on original data
    const originalRowIndex = getDisplayedData[row]?.originalRowIndex;
    if (originalRowIndex === undefined || originalRowIndex === -1) {
      return 28; // Default height for padded rows
    }
    const rowCells = sheets[activeSheet][originalRowIndex]; // Still use original sheets data for row height
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
      const cell = selection.current.cell;
      if (cell) {
        
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
      const newSheets = JSON.parse(JSON.stringify(sheets)); // Deep copy
      newSheets[activeSheet][originalRowIndex][col] = {
        ...newSheets[activeSheet][originalRowIndex][col],
        value: "",
        formula: "",
      };
      pushToUndoStack(sheets); // Optional: add to undo stack
      setSheets(newSheets);
      setDataUpdateKey((prev) => prev + 1); // Trigger re-render
    }
  }
}

  // ✅ Ctrl+C: Copy (value + background)
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

  // ✅ Ctrl+V: Paste (value + background)
  if (e.ctrlKey && e.key === "v") {
    e.preventDefault();
    navigator.clipboard.readText().then((text) => {
      try {
        const copiedData = JSON.parse(text);
        const newSheets = JSON.parse(JSON.stringify(sheets)); // deep copy

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
                strikethrough: false, // Ensure default for strikethrough
                alignment: "left",
                ...copiedData[rowOffset][colOffset], // paste value + background
              };
            }
          }
        }
        pushToUndoStack(sheets);
        setSheets(newSheets);
        setDataUpdateKey(prev => prev + 1); // Data changed
      } catch (err) {
        console.warn("Invalid clipboard format");
      }
    });
  }

  // ✅ Ctrl+X: Cut (copy + clear value + background)
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

        // Clear the cell
        newSheets[activeSheet][originalRowIndex][col] = {
          value: "",
          background: "",
          fontSize: 14,
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false, // Ensure default for strikethrough
          alignment: "left",
        };
      }
      copiedData.push(rowData);
    }

    navigator.clipboard.writeText(JSON.stringify(copiedData));
    pushToUndoStack(sheets);
    setSheets(newSheets);
    setDataUpdateKey(prev => prev + 1); // Data changed
  }
}}


        />
      </div>
      {/* Removed old contextMenu as it's replaced by Edit menu */}
      {/* {contextMenu?.visible && (
        <div
          style={{
            position: "fixed",
            top: contextMenu.y,
            left: contextMenu.x,
            background: "#fff",
            border: "1px solid #ccc",
            borderRadius: "4px",
            boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
            zIndex: 9999,
          }}
          onContextMenu={(e) => e.preventDefault()}
        >
          <button style={menuItem} onClick={() => {
            insertRowAt(contextMenu.targetRow);
            setContextMenu(null);
          }}> Add Row Above</button>

          <button style={menuItem} onClick={() => {
            insertRowAt(contextMenu.targetRow + 1);
            setContextMenu(null);
          }}>Add Row Below</button>

          <button style={menuItem} onClick={() => {
            deleteRows(contextMenu.targetRow, 1);
            setContextMenu(null);
          }}>Delete Row</button>
        </div>
      )} */}

      <div style={{
        padding: "10px",
        borderTop: `1px solid ${currentTheme.border}`, // Google-like border
        display: "flex",
        alignItems: "center",
        overflowX: "auto",
        flexShrink: 0,
        backgroundColor: currentTheme.bg2, // Light grey background
      }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          {Object.keys(sheets).map((sheetName) => (
            <span
              key={sheetName}
              style={{
                padding: "8px 12px",
                marginRight: "2px",
                cursor: "pointer",
                backgroundColor: activeSheet === sheetName ? currentTheme.activeTabBg : "transparent", // Highlight active tab
                border: "1px solid transparent", // Transparent border
                borderBottom: activeSheet === sheetName ? `2px solid ${currentTheme.activeTabBorder}` : `1px solid ${currentTheme.border}`, // Blue underline for active
                borderRadius: "4px 4px 0 0", // Rounded top corners
                display: "inline-flex",
                alignItems: "center",
                whiteSpace: "nowrap",
                color: activeSheet === sheetName ? currentTheme.activeTabBorder : currentTheme.textLight, // Text color
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
              background: currentTheme.activeTabBg, // Light grey for add button
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
    </div>
  );
};

export default ContactGrid;
