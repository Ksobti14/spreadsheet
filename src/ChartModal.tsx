import React, { useState, useEffect, useMemo } from 'react';
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from 'recharts';

// Define the props for the ChartModal component
interface ChartModalProps {
  // The raw sheet data (CellData[][]) from ContactGrid
  sheetData: { originalRowIndex: number; data: { value: string; formula?: string; }[]; }[];
  // Column headers (e.g., A, B, C)
  columns: { title: string; id: string; width: number; }[];
  // Callback to close the modal
  onClose: () => void;
  // Current theme for styling the modal
  currentTheme: {
    bg: string;
    bg2: string;
    text: string;
    textLight: string;
    border: string;
    menuBg: string;
    menuHoverBg: string;
    activeTabBg: string;
    activeTabBorder: string;
    shadow: string;
    cellHighlightBg: string;
    cellHighlightBorder: string;
  };
}

// ChartModal functional component
const ChartModal: React.FC<ChartModalProps> = ({ sheetData, columns, onClose, currentTheme }) => {
  // State to manage the selected chart type (e.g., 'bar', 'line')
  const [chartType, setChartType] = useState<'bar' | 'line'>('bar');
  // State to manage the selected column for the X-axis
  const [xAxisColumn, setXAxisColumn] = useState<string | null>(null);
  // State to manage the selected columns for the Y-axis (can be multiple)
  const [yAxisColumns, setYAxisColumns] = useState<string[]>([]);

  // Memoize the processed data for charting
  const chartData = useMemo(() => {
    // Ensure we have a valid X-axis column selected
    if (!xAxisColumn) return [];

    // Find the index of the selected X-axis column
    const xAxisColIndex = columns.findIndex(col => col.title === xAxisColumn);
    if (xAxisColIndex === -1) return [];

    // Map the sheet data to a format suitable for Recharts
    // Each row becomes an object where keys are column titles and values are cell values
    return sheetData.map(rowData => {
      const obj: { [key: string]: any } = {};
      // Get the value for the X-axis
      obj[xAxisColumn] = rowData.data[xAxisColIndex]?.value || '';

      // Get values for all selected Y-axis columns
      yAxisColumns.forEach(yColTitle => {
        const yColIndex = columns.findIndex(col => col.title === yColTitle);
        if (yColIndex !== -1) {
          // Attempt to parse as number, otherwise keep as string
          const value = parseFloat(rowData.data[yColIndex]?.value || '');
          obj[yColTitle] = isNaN(value) ? rowData.data[yColIndex]?.value : value;
        }
      });
      return obj;
    }).filter(row => Object.keys(row).length > 1); // Filter out rows with only X-axis data
  }, [sheetData, columns, xAxisColumn, yAxisColumns]);

  // Effect to set initial X-axis column when columns are loaded
  useEffect(() => {
    if (columns.length > 0 && !xAxisColumn) {
      setXAxisColumn(columns[0].title); // Default to the first column for X-axis
    }
  }, [columns, xAxisColumn]);

  // Handle Y-axis column selection/deselection
  const handleYAxisColumnChange = (columnTitle: string) => {
    setYAxisColumns(prev =>
      prev.includes(columnTitle)
        ? prev.filter(col => col !== columnTitle) // Deselect if already selected
        : [...prev, columnTitle] // Select if not selected
    );
  };

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        width: '100vw',
        height: '100vh',
        backgroundColor: 'rgba(0, 0, 0, 0.7)', // Dark overlay
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        zIndex: 2000, // Ensure it's on top
      }}
    >
      <div
        style={{
          backgroundColor: currentTheme.bg,
          padding: '25px',
          borderRadius: '10px',
          boxShadow: currentTheme.shadow,
          width: '90%',
          maxWidth: '1000px',
          maxHeight: '90%',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden', // Hide overflow
          color: currentTheme.text,
        }}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
          <h2 style={{ margin: 0, color: currentTheme.text }}>Create Chart</h2>
          <button
            onClick={onClose}
            style={{
              background: 'none',
              border: 'none',
              fontSize: '24px',
              cursor: 'pointer',
              color: currentTheme.textLight,
            }}
          >
            &times;
          </button>
        </div>

        {/* Chart Options */}
        <div style={{ marginBottom: '20px', display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
          {/* Chart Type Selection */}
          <div>
            <label style={{ marginRight: '10px', fontWeight: 'bold' }}>Chart Type:</label>
            <select
              value={chartType}
              onChange={(e) => setChartType(e.target.value as 'bar' | 'line')}
              style={{ padding: '8px', borderRadius: '5px', border: `1px solid ${currentTheme.border}`, backgroundColor: currentTheme.bg2, color: currentTheme.text }}
            >
              <option value="bar">Bar Chart</option>
              <option value="line">Line Chart</option>
            </select>
          </div>

          {/* X-Axis Column Selection */}
          <div>
            <label style={{ marginRight: '10px', fontWeight: 'bold' }}>X-Axis:</label>
            <select
              value={xAxisColumn || ''}
              onChange={(e) => setXAxisColumn(e.target.value)}
              style={{ padding: '8px', borderRadius: '5px', border: `1px solid ${currentTheme.border}`, backgroundColor: currentTheme.bg2, color: currentTheme.text }}
            >
              <option value="">Select Column</option>
              {columns.map(col => (
                <option key={col.id} value={col.title}>
                  {col.title}
                </option>
              ))}
            </select>
          </div>

          {/* Y-Axis Column Selection (Multi-select) */}
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            <label style={{ marginBottom: '5px', fontWeight: 'bold' }}>Y-Axes:</label>
            <div style={{
              display: 'flex',
              flexWrap: 'wrap',
              gap: '10px',
              maxHeight: '100px',
              overflowY: 'auto',
              padding: '5px',
              border: `1px solid ${currentTheme.border}`,
              borderRadius: '5px',
              backgroundColor: currentTheme.bg2,
            }}>
              {columns.filter(col => col.title !== xAxisColumn).map(col => (
                <label key={col.id} style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    value={col.title}
                    checked={yAxisColumns.includes(col.title)}
                    onChange={() => handleYAxisColumnChange(col.title)}
                    style={{ marginRight: '5px' }}
                  />
                  {col.title}
                </label>
              ))}
            </div>
          </div>
        </div>

        {/* Chart Display Area */}
        <div style={{ flexGrow: 1, minHeight: '300px', width: '100%', overflow: 'hidden' }}>
          {chartData.length > 0 && xAxisColumn && yAxisColumns.length > 0 ? (
            <ResponsiveContainer width="100%" height="100%">
              {chartType === 'bar' ? (
                <BarChart
                  data={chartData}
                  margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke={currentTheme.border} />
                  <XAxis dataKey={xAxisColumn} stroke={currentTheme.text} />
                  <YAxis stroke={currentTheme.text} />
                  <Tooltip
                    contentStyle={{ backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, color: currentTheme.text }}
                    itemStyle={{ color: currentTheme.text }}
                    labelStyle={{ color: currentTheme.textLight }}
                  />
                  <Legend />
                  {yAxisColumns.map((yCol, index) => (
                    <Bar key={yCol} dataKey={yCol} fill={`hsl(${index * 60}, 70%, 50%)`} />
                  ))}
                </BarChart>
              ) : (
                <LineChart
                  data={chartData}
                  margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke={currentTheme.border} />
                  <XAxis dataKey={xAxisColumn} stroke={currentTheme.text} />
                  <YAxis stroke={currentTheme.text} />
                  <Tooltip
                    contentStyle={{ backgroundColor: currentTheme.menuBg, border: `1px solid ${currentTheme.border}`, color: currentTheme.text }}
                    itemStyle={{ color: currentTheme.text }}
                    labelStyle={{ color: currentTheme.textLight }}
                  />
                  <Legend />
                  {yAxisColumns.map((yCol, index) => (
                    <Line
                      key={yCol}
                      type="monotone"
                      dataKey={yCol}
                      stroke={`hsl(${index * 60}, 70%, 50%)`}
                      activeDot={{ r: 8 }}
                    />
                  ))}
                </LineChart>
              )}
            </ResponsiveContainer>
          ) : (
            <div style={{ textAlign: 'center', marginTop: '50px', color: currentTheme.textLight }}>
              Please select X-Axis and at least one Y-Axis column to display the chart.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default ChartModal;
