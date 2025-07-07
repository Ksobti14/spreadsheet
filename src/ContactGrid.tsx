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
const socket=io("https://testsheet-ui6b.onrender.com",{
   transports:["polling"]
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
};

type SheetData = {
  [sheetName: string]: CellData[][];
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
const parseArg = (arg: string, data: CellData[][]): number[] => {
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
const evaluateFormula = (formula: string, data: CellData[][]): string => {
  const funcMatch = formula.match(/^=(\w+)\(([^)]*)\)$/i);
  if (!funcMatch) return formula;

  const [, func, args] = funcMatch;
  const funcUpper = func.toUpperCase();
  const parsedArgs = args.split(',').map(s => s.trim()).flatMap(arg => parseArg(arg, data));
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
        const condVal = parseArg(cond, data)[0] || 0;
        return condVal ? parseArg(trueVal, data)[0].toString() || trueVal : parseArg(falseVal, data)[0].toString() || falseVal;
      case "ROUND":
        // ROUND(number, digits)
        const [numArg, digitsArg] = args.split(',').map(s => s.trim());
        const num = parseArg(numArg, data)[0] || 0;
        const digits = parseInt(digitsArg) || 0;
        return FormulaJS.ROUND(num, digits).toString();
      case "ABS":
        const absArg = parseArg(args, data)[0] || 0;
        return FormulaJS.ABS(absArg).toString();
      case "SQRT":
        const sqrtArg = parseArg(args, data)[0] || 0;
        return sqrtArg >= 0 ? FormulaJS.SQRT(sqrtArg).toString() : error("Negative number");
      case "POWER":
        const [baseArg, expArg] = args.split(',').map(s => s.trim());
        const base = parseArg(baseArg, data)[0] || 0;
        const exp = parseArg(expArg, data)[0] || 1;
        return FormulaJS.POWER(base, exp).toString();
      default:
        return formula;
    }
  } catch (e) {
    return error("Invalid formula");
  }
};
const menuItem: React.CSSProperties = {
  padding: "8px 16px", // Slightly less padding for a compact menu
  textAlign: "left",
  background: "white",
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
  backgroundColor: "#fff",
  border: "1px solid #dadce0", // Google-like border
  borderRadius: "4px",
  boxShadow: "0 2px 6px rgba(60,64,67,0.15)", // Subtle shadow
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
  backgroundColor: "#fff",
  border: "1px solid #dadce0",
  borderRadius: "4px",
  boxShadow: "0 2px 6px rgba(60,64,67,0.15)",
  zIndex: 1001, // Higher z-index to appear above parent dropdown
  minWidth: "160px",
  padding: "4px 0",
  display: "flex",
  flexDirection: "column",
};

const topBarButtonStyle: React.CSSProperties = {
  padding: "6px 10px", // Smaller padding for top bar buttons
  background: "transparent", // Transparent background
  color: "#202124", // Dark text color
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
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.setSelectionRange(localValue.length, localValue.length);
  }, []);

  const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setLocalValue(val);
    onChange(val);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") {
      onChange(localValue);
      onCommit();
    } else if (e.key === "Escape") {
      onCancel();
    }
  };

  return (
    <input
      ref={inputRef}
      value={localValue}
      onChange={handleInput}
      onKeyDown={handleKeyDown}
      style={{
        width: "100%",
        height: "100%",
        fontSize: "14px",
        border: "none",
        outline: "none",
        padding: "0 6px",
        background: "#fff",
      }}
    />
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


  const activeCell = useRef<Item | null>(null);
  const selecting = useRef<Rectangle | null>(null);
  const currentSheetData = sheets[activeSheet];
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
      const cell = currentSheetData[row]?.[col] ?? { value: "" };
      let displayValue = cell.formula ? evaluateFormula(cell.formula, currentSheetData) : cell.value;
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
        ? { bgCell: "#E0F0FF", borderColor: "#1E90FF" }
        : undefined, 
      };
    },
    [currentSheetData, highlightRange]
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
    socket.emit("cell-edit", {
  sheet: activeSheet,
  row,
  col,
  value: text,
});

    if (text.startsWith("=")) {
      const checkArgCount = (func: string, args: string): boolean => {
        const argCount = args
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s !== "").length;

        if (func === "IF") {
          setFormulaError(argCount !== 3 ? `${func} requires 3 arguments` : null);
          return argCount === 3;
        } else if (func === "ROUND" || func === "POWER") {
          setFormulaError(argCount !== 2 ? `${func} requires 2 arguments` : null);
          return argCount === 2;
        } else if (func === "ABS" || func === "SQRT") {
          if (argCount !== 1) {
            setFormulaError(`${func} requires 1 argument`);
            return false;
          } else if (args.includes(":")) {
            setFormulaError(`${func} requires a single cell, not a range`);
            return false;
          }
          return true;
        } else if (
          func === "SUM" ||
          func === "AVERAGE" ||
          func === "MIN" ||
          func === "MAX" ||
          func === "COUNT" ||
          func === "PRODUCT"
        ) {
          setFormulaError(argCount < 1 ? `${func} requires at least 1 argument` : null);
          return argCount >= 1;
        }

        setFormulaError(null);
        return true;
      };

      const funcMatch = text.match(/^=(\w+)\(([^)]*)\)$/i);
      if (funcMatch) {
        const [, func, args] = funcMatch;
        const funcUpper = func.toUpperCase();
        if (!checkArgCount(funcUpper, args)) {
          console.warn(`Formula validation failed for: ${text}`);
          return;
        }
      }
    }
    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

    if (text.startsWith("=")) {
      currentSheetCopy[row][col] = {
        formula: text,
        value: evaluateFormula(text, currentSheetCopy),
      };
    } else {
      currentSheetCopy[row][col] = { value: text };
    }

    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
      setFormulaInput("");
      setHighlightRange(null);
      setShowSuggestions(false);
      setFormulaError(null);
    },
    [activeSheet,sheets]
  );
  const onFillPattern = useCallback(
  ({ patternSource, fillDestination }: FillPatternEventArgs) => {
    const sourceCol = patternSource.x;
    const sourceRow = patternSource.y;
    const sourceCell = currentSheetData[sourceRow]?.[sourceCol];
    if (!sourceCell) return;

    const fillValue = sourceCell.value;
    const fillFormula = sourceCell.formula;

    const updatedSheets = { ...sheets };
    const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

    for (let r = fillDestination.y; r < fillDestination.y + fillDestination.height; r++) {
      for (let c = fillDestination.x; c < fillDestination.x + fillDestination.width; c++) {
        const isSourceCell = c >= patternSource.x && c < patternSource.x + patternSource.width &&
                             r >= patternSource.y && r < patternSource.y + patternSource.height;
        if (!isSourceCell) {
          currentSheetCopy[r][c] = fillFormula
            ? { formula: fillFormula, value: evaluateFormula(fillFormula, currentSheetCopy) }
            : { value: fillValue };
        }
      }
    }

    updatedSheets[activeSheet] = currentSheetCopy;
    pushToUndoStack(updatedSheets);
  },
  [activeSheet, currentSheetData]
);
  const onFinishSelecting = useCallback(() => {
  if (!activeCell.current || !selecting.current) return;
  const [col, row] = activeCell.current;
  const topLeft = getCellName(selecting.current.x, selecting.current.y);
  const bottomRight = getCellName(
    selecting.current.x + selecting.current.width - 1,
    selecting.current.y + selecting.current.height - 1
  );
  const formula = `=SUM(${topLeft}:${bottomRight})`;
  const value = evaluateFormula(formula, currentSheetData);

  const updatedSheets = { ...sheets };
  const currentSheetCopy = sheets[activeSheet].map((r) => [...r]);
  currentSheetCopy[row][col] = { formula, value };
  updatedSheets[activeSheet] = currentSheetCopy;
  pushToUndoStack(updatedSheets);

  setFormulaInput("");
  setHighlightRange(null);
  setShowSuggestions(false);
  setFormulaError(null);
}, [activeSheet, currentSheetData]);
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
                  newSheetData[rowIndex][colIndex] = { formula: cellValue, value: evaluateFormula(cellValue, newSheetData) };
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
      } catch (err) {
        console.error("Error importing Excel file:", err);
        alert("Failed to import Excel file: Invalid format or content.");
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
        const displayValue = cell.formula ? evaluateFormula(cell.formula, activeSheetData) : cell.value;
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
                newSheetData[rowIndex][colIndex] = { formula: cellValue, value: evaluateFormula(cellValue, newSheetData) };
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
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
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
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
    ...cell,
    alignment: align,
  };
  updated[activeSheet] = copy;
  pushToUndoStack(updated);
};

const customDrawCell = (args: any) => {
  const { ctx, theme, rect, col, row, cell } = args;
  if (cell.kind !== GridCellKind.Text) return false;

  const cellData = sheets[activeSheet]?.[row]?.[col];

  const alignment: "left" | "center" | "right" = cellData?.alignment || "left";
  const fontSize = cellData?.fontSize || 12;
  const isBold = cellData?.bold ? "bold" : "normal";
  const isItalic = cellData?.italic ? "italic" : "normal";
  const isUnderline = cellData?.underline; 
  const isStrikethrough = cellData?.strikethrough;
  const textColor = cellData?.textColor || theme.textDark;
  const bgColor = cellData?.bgColor || theme.bgCell;
  const borderColor = cellData?.borderColor || theme.borderColor;
  const fontFamily = cellData?.fontFamily || 'sans-serif'; // Get font family, default to sans-serif
  let text = String(cell.displayData ?? "");

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
};
const toggleStyle = (type: "bold" | "italic") => {
  if (!activeCell.current) return;
  const [col, row] = activeCell.current;

  const updated = { ...sheets };
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
    ...cell,
    [type]: !cell?.[type],
  };
  updated[activeSheet] = copy;
  pushToUndoStack(updated);
};

const applyCellColor = (type: "textColor" | "bgColor" | "borderColor", color: string) => {
  if (!activeCell.current) return;
  const [col, row] = activeCell.current;

  const updated = { ...sheets };
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
    ...cell,
    [type]: color,
  };
  updated[activeSheet] = copy;
  pushToUndoStack(updated);
};  
const pushToUndoStack = (newSheets: SheetData) => {
  setUndoStack(prev => [...prev, sheets]);
  setRedoStack([]); 
  setSheets(newSheets);
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
};

const handleRedo = () => {
  if (redoStack.length === 0) return;
  const next = redoStack[redoStack.length - 1];
  setUndoStack(prev => [...prev, sheets]);
  setRedoStack(prev => prev.slice(0, -1));
  setSheets(next);
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
}, []);
 const applyStyleToRange = (key: string, value: any) => {
  if (!selection?.current) return;
  const { x, y, width, height } = selection.current.range;
  const currentSheet = sheets[activeSheet];

  let allHaveSame = true;
  for (let row = y; row < y + height; row++) {
    for (let col = x; col < x + width; col++) {
      const cell = currentSheet?.[row]?.[col];
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
    for (let col = x; col < x + width; col++) {
      const cell = copy[row]?.[col] ?? { value: "" };
      // Toggle the style: if all selected cells have the style, remove it; otherwise, apply it.
      copy[row][col] = { ...cell, [key]: allHaveSame ? undefined : value };
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
        const rowData: any[] = [];
        for (let col = startX; col <= endX; col++) {
          rowData.push(sheets[activeSheet][row]?.[col] ?? {
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
            for (let colOffset = 0; colOffset < copiedData[rowOffset].length; colOffset++) {
              const targetRow = pasteStartY + rowOffset;
              const targetCol = pasteStartX + colOffset;
              if (newSheets[activeSheet][targetRow] && newSheets[activeSheet][targetRow][targetCol]) {
                newSheets[activeSheet][targetRow][targetCol] = {
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
        const rowData: any[] = [];
        for (let col = startX; col <= endX; col++) {
          const cell = newSheets[activeSheet][row]?.[col] ?? {
            value: "",
            background: "",
          };
          rowData.push(cell);

          // Clear the cell
          newSheets[activeSheet][row][col] = {
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
    }
  };
}, [selection, sheets, activeSheet, clipboardData]);

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
// Removed showDataDropdown and showHelpDropdown states

const [showImportDropdown, setShowImportDropdown] = useState(false);
const [showExportDropdown, setShowExportDropdown] = useState(false);

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
  insertRowAt(selection.current.range.y);
  setShowEditDropdown(false);
};

const handleInsertRowBelow = () => {
  if (!selection.current) {
    alert("Please select a cell to insert a row.");
    return;
  }
  insertRowAt(selection.current.range.y + selection.current.range.height);
  setShowEditDropdown(false);
};

const handleDeleteSelectedRows = () => {
  if (selection.rows.length === 0) { // Check length directly
    alert("Please select rows to delete.");
    return;
  }
  // Get all selected row indices
  const selectedRowIndices: number[] = Array.from(selection.rows); // Use Array.from to iterate

  // Sort in descending order to avoid index issues during deletion
  selectedRowIndices.sort((a, b) => b - a);

  const updatedSheets = { ...sheets };
  let currentSheetCopy = sheets[activeSheet].map((r) => [...r]);

  selectedRowIndices.forEach(rowIndex => {
    currentSheetCopy.splice(rowIndex, 1);
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
  const updated = { ...sheets };
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
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
  const updated = { ...sheets };
  const copy = updated[activeSheet].map((r) => [...r]);
  const cell = copy[row][col];
  copy[row][col] = {
    ...cell,
    comment: commentValue.trim(),
  };
  updated[activeSheet] = copy;
  pushToUndoStack(updated);
  setCommentValue('');
  setShowCommentInput(false);
};


  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', fontFamily: 'Roboto, sans-serif', color: '#202124' }}>
      {/* Top Bar - Google Sheets like */}
      <div style={{
        backgroundColor: '#fff',
        borderBottom: '1px solid #dadce0',
        padding: '8px 16px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexShrink: 0,
        boxShadow: '0 1px 2px 0 rgba(60,64,67,0.08)',
      }}>
        {/* Left Section: Logo, Title, Menu Bar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          {/* Logo Placeholder */}
          <span style={{ fontSize: '24px', color: '#1a73e8', fontWeight: 'bold' }}>
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
              border: '1px solid #dadce0',
              fontSize: '15px',
              fontWeight: 500,
              color: '#202124',
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
              <button style={topBarButtonStyle}>File</button>
              {showFileDropdown && (
                <div style={menuDropdownStyle}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={saveSheetData} style={menuItem}>Save</button>
                  <div style={{ borderTop: '1px solid #eee', margin: '4px 0' }}></div>
                  
                  {/* Import Sub-menu */}
                  <div 
                    style={{ position: "relative" }}
                    onMouseEnter={() => setShowImportDropdown(true)}
                    onMouseLeave={() => setShowImportDropdown(false)}
                  >
                    <button style={menuItem}>Import</button> {/* Right arrow for sub-menu */}
                    {showImportDropdown && (
                      <div style={subMenuDropdownStyle}>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} style={menuItem}>
                          Import XLSX
                          <input type="file" accept=".xlsx" onChange={importFromExcel} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} style={menuItem}>
                          Import JSON
                          <input type="file" accept="application/json" onChange={handleImportFromJSON} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} style={menuItem}>
                          Import CSV
                          <input type="file" accept=".csv" onChange={importFromCSV} style={{ display: "none" }} />
                        </label>
                        <label onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} style={menuItem}>
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
                    <button style={menuItem}>Export</button> 
                    {showExportDropdown && (
                      <div style={subMenuDropdownStyle}>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleExportXLSX} style={menuItem}>Export as XLSX</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleExportToJSON} style={menuItem}>Export as JSON</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={exportToCSV} style={menuItem}>Export as CSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleODSExport} style={menuItem}>Export as ODS</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleExportTSV} style={menuItem}>Export as TSV</button>
                        <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleExportPDF} style={menuItem}>Export as PDF</button>
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
              <button style={topBarButtonStyle}>Edit</button>
              {showEditDropdown && (
                <div style={menuDropdownStyle}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleUndo} disabled={undoStack.length === 0} style={{...menuItem, opacity: undoStack.length === 0 ? 0.5 : 1}}>Undo</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleRedo} disabled={redoStack.length === 0} style={{...menuItem, opacity: redoStack.length === 0 ? 0.5 : 1}}>Redo</button>
                  <div style={{ borderTop: '1px solid #eee', margin: '4px 0' }}></div>
                  {/* Removed: Cut, Copy, Paste */}
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleInsertRowAbove} style={menuItem}>Insert Row Above</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleInsertRowBelow} style={menuItem}>Insert Row Below</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleDeleteSelectedRows} style={menuItem}>Delete Selected Rows</button>
                </div>
              )}
            </div>

            {/* Insert Menu */}
            <div 
              style={{ position: "relative" }}
              onMouseEnter={() => setShowInsertDropdown(true)}
              onMouseLeave={() => setShowInsertDropdown(false)}
            >
              <button style={topBarButtonStyle}>Insert</button>
              {showInsertDropdown && (
                <div style={menuDropdownStyle}>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleInsertLink} style={menuItem}>Link</button>
                  <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={handleInsertComment} style={menuItem}>Comment</button>
                  {/* Removed Row above and Row below */}
                </div>
              )}
            </div>
            <div 
              style={{ position: "relative" }}
              onMouseEnter={() => setShowFormatDropdown(true)}
              onMouseLeave={() => setShowFormatDropdown(false)}
            >
              <button style={topBarButtonStyle}>Format</button>
              {showFormatDropdown && (
                <div style={menuDropdownStyle}>
                  {/* These were moved to the main toolbar */}
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: '#5f6368' }}>Text Styles</div>
                  <select
                    onChange={(e) => applyStyleToRange("fontSize",parseInt(e.target.value))}
                    style={{...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px'}}
                  >
                    <option value="">Font Size</option>
                    {[10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].map(size => (
                      <option key={size} value={size}>{size}px</option>
                    ))}
                  </select>
                  <select
                    onChange={(e) => applyStyleToRange("fontFamily", e.target.value)}
                    style={{...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px'}}
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
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={() => applyStyleToRange("bold",true)} style={{...topBarButtonStyle, fontWeight: "bold", border: '1px solid #dadce0'}}>B</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={() => applyStyleToRange("italic",true)} style={{...topBarButtonStyle, fontStyle: "italic", border: '1px solid #dadce0'}}>I</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={() => applyStyleToRange("underline",true)} style={{...topBarButtonStyle, textDecoration: "underline", border: '1px solid #dadce0'}}>U</button>
                    <button onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e6e6e6')} onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'white')} onClick={() => applyStyleToRange("strikethrough",true)} style={{...topBarButtonStyle, textDecoration: "line-through", border: '1px solid #dadce0'}}>S</button>
                  </div>
                  <div style={{ borderTop: '1px solid #eee', margin: '4px 0' }}></div>
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: '#5f6368' }}>Colors</div>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px", padding: '4px 8px' }}>
                    <label style={{ fontSize: "12px", fontWeight: 500 }}>Text:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("textColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                    <label style={{ fontSize: "12px", fontWeight: 500 }}>Fill:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("bgColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                    <label style={{ fontSize: "12px", fontWeight: 500 }}>Border:</label>
                    <input type="color" onChange={(e) => applyStyleToRange("borderColor", e.target.value)} style={{ width: "24px", height: "24px", border: "none", borderRadius: "50%", cursor: "pointer" }} />
                  </div>
                  <div style={{ borderTop: '1px solid #eee', margin: '4px 0' }}></div>
                  <div style={{ padding: '8px 16px', fontWeight: 'bold', fontSize: '13px', color: '#5f6368' }}>Alignment</div>
                  <select
                    onChange={(e) => applyStyleToRange("alignment",e.target.value as "left" | "center" | "right")}
                    style={{...menuItem, width: 'calc(100% - 16px)', margin: '4px 8px'}}
                  >
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                  </select>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Formatting Toolbar */}
      <div style={{
        backgroundColor: '#f8f9fa',
        borderBottom: '1px solid #dadce0',
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
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
            style={{
                padding: "6px 10px",
                background: "transparent",
                color: "#202124",
                border: "1px solid #dadce0",
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
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
            style={{
                padding: "6px 10px",
                background: "transparent",
                color: "#202124",
                border: "1px solid #dadce0",
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

        <div style={{ width: '1px', height: '24px', backgroundColor: '#dadce0', margin: '0 5px' }}></div> {/* Divider */}

        {/* Font Family */}
        <select
          onChange={(e) => applyStyleToRange("fontFamily", e.target.value)}
          style={{
            padding: "6px 10px",
            borderRadius: "4px",
            border: "1px solid #dadce0",
            minWidth: "120px",
            fontSize: '13px',
            color: '#202124',
            backgroundColor: 'white',
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
            border: "1px solid #dadce0",
            minWidth: "70px",
            fontSize: '13px',
            color: '#202124',
            backgroundColor: 'white',
          }}
        >
          {[10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].map(size => (
            <option key={size} value={size}>{size}px</option>
          ))}
        </select>

        <div style={{ width: '1px', height: '24px', backgroundColor: '#dadce0', margin: '0 5px' }}></div> {/* Divider */}

        {/* Font Styles (Bold, Italic, Underline, Strikethrough) */}
        <button
          onClick={() => applyStyleToRange("bold",true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            fontWeight: "bold",
            border: '1px solid #dadce0',
            fontSize: '16px',
          }}
        >B</button>
        <button
          onClick={() => applyStyleToRange("italic",true)}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            fontStyle: "italic",
            border: '1px solid #dadce0',
            fontSize: '16px',
          }}
        >I</button>
        <button
          onClick={() => applyStyleToRange("underline",true)} 
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            textDecoration: "underline",
            border: '1px solid #dadce0',
            fontSize: '16px',
          }}
        >U</button>
        <button
          onClick={() => applyStyleToRange("strikethrough",true)} 
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#e8eaed')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          style={{
            ...topBarButtonStyle,
            textDecoration: "line-through",
            border: '1px solid #dadce0',
            fontSize: '16px',
          }}
        >S</button>

        <div style={{ width: '1px', height: '24px', backgroundColor: '#dadce0', margin: '0 5px' }}></div> {/* Divider */}

        {/* Text Color */}
        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: '#5f6368' }}>Text:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("textColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: "1px solid #dadce0",
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: 'white',
            }}
          />
        </div>
        {/* Background Color */}
        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: '#5f6368' }}>Fill:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("bgColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: "1px solid #dadce0",
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: 'white',
            }}
          />
        </div>
        {/* Border Color */}
        <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
          <label style={{ fontSize: "12px", fontWeight: 500, color: '#5f6368' }}>Border:</label>
          <input
            type="color"
            onChange={(e) => applyStyleToRange("borderColor", e.target.value)}
            style={{
              width: "28px",
              height: "28px",
              border: "1px solid #dadce0",
              borderRadius: "4px",
              cursor: "pointer",
              backgroundColor: 'white',
            }}
          />
        </div>

        <div style={{ width: '1px', height: '24px', backgroundColor: '#dadce0', margin: '0 5px' }}></div> {/* Divider */}

        {/* Alignment */}
        <select
          onChange={(e) => applyStyleToRange("alignment",e.target.value as "left" | "center" | "right")}
          style={{
            padding: "6px 10px",
            borderRadius: "4px",
            border: "1px solid #dadce0",
            minWidth: "90px",
            fontSize: '13px',
            color: '#202124',
            backgroundColor: 'white',
          }}
        >
          <option value="left">Left</option>
          <option value="center">Center</option>
          <option value="right">Right</option>
        </select>
      </div>


      {/* Formula Bar */}
      <div style={{
        backgroundColor: '#f8f9fa',
        borderBottom: '1px solid #dadce0',
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
            color: '#5f6368', 
            marginRight: '10px',
            padding: '4px 8px',
            border: '1px solid #dadce0',
            borderRadius: "4px",
            backgroundColor: '#fff',
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
              border: "1px solid #dadce0", 
              borderRadius: "4px", 
              fontSize: "14px",
              outline: 'none',
              boxShadow: 'inset 0 1px 2px rgba(0,0,0,0.06)',
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
              background: "white",
              border: "1px solid #ccc",
              borderRadius: "4px",
              boxShadow: "0 2px 8px rgba(0, 0, 0, 0.15)",
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
                  borderBottom: idx === suggestions.length - 1 ? "none" : "1px solid #eee",
                  backgroundColor: "#fff",
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

      {/* Link Input Modal */}
      {showLinkInput && (
        <div style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          backgroundColor: "white",
          padding: "20px",
          borderRadius: "8px",
          boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "10px",
        }}>
          <h3>Insert Link</h3>
          <input
            type="text"
            value={linkValue}
            onChange={(e) => setLinkValue(e.target.value)}
            placeholder="Enter URL"
            style={{ padding: "8px", borderRadius: "4px", border: "1px solid #ccc" }}
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
          backgroundColor: "white",
          padding: "20px",
          borderRadius: "8px",
          boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
          zIndex: 2000,
          display: "flex",
          flexDirection: "column",
          gap: "10px",
        }}>
          <h3>Insert Comment</h3>
          <textarea
            value={commentValue}
            onChange={(e) => setCommentValue(e.target.value)}
            placeholder="Enter comment"
            rows={4}
            style={{ padding: "8px", borderRadius: "4px", border: "1px solid #ccc", minWidth: "250px" }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
            <button onClick={applyCommentToCell} style={{ padding: "8px 12px", background: "#4CAF50", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Apply</button>
            <button onClick={() => setShowCommentInput(false)} style={{ padding: "8px 12px", background: "#f44336", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Cancel</button>
          </div>
        </div>
      )}

      <div style={{ flexGrow: 1, overflow: 'hidden' }}
        /* Removed onContextMenu from here as requested */
      >
        <DataEditor
          columns={columns}
          rows={NUM_ROWS}
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
            const currentCell = currentSheetData[cell[1]]?.[cell[0]];
            if (currentCell) {
              setFormulaInput(currentCell.formula || currentCell.value);
            } else {
              setFormulaInput("");
            }
          }}
          onFillPattern={onFillPattern}
          rowHeight={(row) => {
    const rowCells = sheets[activeSheet][row];
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

  // ✅ Ctrl+C: Copy (value + background)
  if (e.ctrlKey && e.key === "c") {
    e.preventDefault();
    const copiedData: any[][] = [];
    for (let row = startY; row <= endY; row++) {
      const rowData: any[] = [];
      for (let col = startX; col <= endX; col++) {
        rowData.push(sheets[activeSheet][row]?.[col] ?? {
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
          for (let colOffset = 0; colOffset < copiedData[rowOffset].length; colOffset++) {
            const targetRow = startY + rowOffset; // startY is accessible here
            const targetCol = startX + colOffset; // startX is accessible here
            if (newSheets[activeSheet][targetRow] && newSheets[activeSheet][targetRow][targetCol]) {
              newSheets[activeSheet][targetRow][targetCol] = {
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
      const rowData: any[] = [];
      for (let col = startX; col <= endX; col++) {
        const cell = newSheets[activeSheet][row]?.[col] ?? {
          value: "",
          background: "",
        };
        rowData.push(cell);

        // Clear the cell
        newSheets[activeSheet][row][col] = {
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
  }
}}


        />
      </div>
      <div style={{
        padding: "10px",
        borderTop: "1px solid #dadce0", // Google-like border
        display: "flex",
        alignItems: "center",
        overflowX: "auto",
        flexShrink: 0,
        backgroundColor: '#f8f9fa', // Light grey background
      }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          {Object.keys(sheets).map((sheetName) => (
            <span
              key={sheetName}
              style={{
                padding: "8px 12px",
                marginRight: "2px", 
                cursor: "pointer",
                backgroundColor: activeSheet === sheetName ? "#e8eaed" : "transparent", // Highlight active tab
                border: "1px solid transparent", // Transparent border
                borderBottom: activeSheet === sheetName ? "2px solid #1a73e8" : "1px solid #dadce0", // Blue underline for active
                borderRadius: "4px 4px 0 0", // Rounded top corners
                display: "inline-flex",
                alignItems: "center",
                whiteSpace: "nowrap",
                color: activeSheet === sheetName ? "#1a73e8" : "#5f6368", // Text color
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
                  style={{ border: "none", background: "transparent", outline: "none", width: "80px", textAlign: 'center' }}
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
                        color: "#888",
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
              background: "#e8eaed", // Light grey for add button
              color: "#5f6368",
              border: "1px solid #dadce0",
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
