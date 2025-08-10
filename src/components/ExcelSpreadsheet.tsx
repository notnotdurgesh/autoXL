import React, { useState, useCallback, useRef, useEffect, useMemo, memo } from 'react';
// Vite worker import via ?worker&url
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import GridWorkerUrl from '../workers/gridWorker.ts?worker&url';
import { evaluateFormula } from '../utills/formula';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import { useUndoRedo, type Command } from '../hooks/useUndoRedo';
import ExcelToolbar from './ExcelToolbar';
import ActionSidebar from './ActionSidebar';
import FormulaInputBar from './FormulaInputBar';
import AIChat from './AIChat';
import type { SpreadsheetOperations } from '../services/gemini/spreadsheetHandlers';
import { createOptimizedHandler, runWhenIdle } from '../utills/performanceUtil';

interface CellFormatting {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  color?: string;
  backgroundColor?: string;
  numberFormat?: 'general' | 'number' | 'currency' | 'percentage' | 'date' | 'time' | 'custom';
  decimalPlaces?: number;
  link?: string;
}

interface CellData {
  value: string | number | null;
  isEditing?: boolean;
  formatting?: CellFormatting;
  displayValue?: string; // For formatted display (currency, percentage, etc.)
}

interface SheetData {
  [row: number]: {
    [col: number]: CellData;
  };
}

interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

// Stable empty formatting to keep referential equality for unformatted cells
const EMPTY_FORMATTING: CellFormatting = Object.freeze({}) as CellFormatting;

const ExcelSpreadsheet: React.FC = memo(() => {
  // Excel-like limits and defaults
  const MAX_ROWS = 1048576; // Excel's row limit
  const MAX_COLS = 16384; // Excel's column limit (XFD)
  const MIN_VISIBLE_ROWS = 50; // Minimum rows to always show
  const MIN_VISIBLE_COLS = 26; // Minimum columns to always show (A-Z)
  const EXPANSION_BUFFER = 10; // Extra rows/cols to add when expanding
  const DEFAULT_COLUMN_WIDTH = 80;
  const DEFAULT_ROW_HEIGHT = 20;
  const MIN_COLUMN_WIDTH = 30;
  const MIN_ROW_HEIGHT = 15;
  const [sheetData, setSheetData] = useState<SheetData>({});
  const [selectedCell, setSelectedCell] = useState<{row: number, col: number} | null>(null);
  const [selectedRange, setSelectedRange] = useState<CellRange | null>(null);
  const [editingCell, setEditingCell] = useState<{row: number, col: number} | null>(null);
  const [inputValue, setInputValue] = useState<string>('');
  const [isEditingInInputBar, setIsEditingInInputBar] = useState<boolean>(false);
  const [isFreshEdit, setIsFreshEdit] = useState<boolean>(false);
  const [isSelecting, setIsSelecting] = useState<boolean>(false);
  const [copiedData, setCopiedData] = useState<{data: CellData[][], range: CellRange, isCut?: boolean} | null>(null);
  const workerRef = useRef<Worker | null>(null);
  const pendingRequestsRef = useRef(new Map<string, (value: unknown) => void>());
  
  // Performance optimization: Track pending selection updates
  const pendingSelectionRef = useRef<CellRange | null>(null);
  const selectionUpdateRAFRef = useRef<number | null>(null);
  
  // Enhanced state for clipboard operations
  const [clipboardNotification, setClipboardNotification] = useState<string | null>(null);
  const [isActionSidebarOpen, setIsActionSidebarOpen] = useState<boolean>(false);
  
  // Context menu state for right-click
  const [contextMenu, setContextMenu] = useState<{x: number; y: number; visible: boolean}>({
    x: 0,
    y: 0,
    visible: false
  });
  const [chatContext, setChatContext] = useState<{range: string; values: string | number | (string | number | null)[][]} | null>(null);
  
  // Dynamic grid size state - starts with minimum but expands as needed
  const [visibleRows, setVisibleRows] = useState<number>(MIN_VISIBLE_ROWS);
  const [visibleCols, setVisibleCols] = useState<number>(MIN_VISIBLE_COLS);
  
  // Zoom functionality state
  const [zoomLevel] = useState<number>(100); // Percentage
  
  // Virtual scrolling and performance optimization
  const MAX_RENDERED_ROWS = 100; // Maximum rows to render at once
  const MAX_RENDERED_COLS = 50;  // Maximum columns to render at once
  const RENDER_BUFFER = 5;       // Extra rows/cols to render outside viewport
  const PERFORMANCE_THRESHOLD = 10000; // Max cells before virtual scrolling kicks in
  
  // Virtual viewport state
  const [viewportStartRow, setViewportStartRow] = useState<number>(0);
  const [viewportStartCol, setViewportStartCol] = useState<number>(0);
  const [viewportEndRow, setViewportEndRow] = useState<number>(MAX_RENDERED_ROWS);
  const [viewportEndCol, setViewportEndCol] = useState<number>(MAX_RENDERED_COLS);

  
  // Scroll position tracking
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const tableRef = useRef<HTMLTableElement>(null);
  
  // Column and row resizing state
  const [columnWidths, setColumnWidths] = useState<{[key: number]: number}>({});
  const [rowHeights, setRowHeights] = useState<{[key: number]: number}>({});
  const [isResizingColumn, setIsResizingColumn] = useState<number | null>(null);
  const [isResizingRow, setIsResizingRow] = useState<number | null>(null);
  const [resizeStartPos, setResizeStartPos] = useState<{x: number, y: number}>({x: 0, y: 0});
  const [resizeStartSize, setResizeStartSize] = useState<number>(0);
  const [tempColumnWidths, setTempColumnWidths] = useState<{[key: number]: number}>({});
  const [tempRowHeights, setTempRowHeights] = useState<{[key: number]: number}>({});
  
  // Fill handle functionality
  const [fillHandleActive, setFillHandleActive] = useState<boolean>(false);
  const [fillPreview, setFillPreview] = useState<CellRange | null>(null);
  const [dragStartCell, setDragStartCell] = useState<{row: number, col: number} | null>(null);
  
  // Undo/Redo functionality
  const { executeCommand, undo, redo, canUndo, canRedo, getState } = useUndoRedo();

  // Versioning state
  interface VersionSnapshot {
    id: string;
    timestamp: number;
    description: string;
    data: {
      sheetData: SheetData;
      columnWidths: { [key: number]: number };
      rowHeights: { [key: number]: number };
    };
  }

  const [versionHistory, setVersionHistory] = useState<Array<{ id: string; timestamp: number; description: string }>>([]);
  const versionSnapshotsRef = useRef<Map<string, VersionSnapshot>>(new Map());
  const snapshotRequestRef = useRef<{ requested: boolean; description: string | null }>({ requested: false, description: null });
  const [previewModal, setPreviewModal] = useState<{ visible: boolean; versionId?: string }>({ visible: false });
  
  // Cell formatting state

  // Helper functions to get current sizes (including temporary values during resize)
  const getCurrentColumnWidth = useCallback((colIndex: number) => {
    if (isResizingColumn === colIndex && tempColumnWidths[colIndex] !== undefined) {
      return tempColumnWidths[colIndex];
    }
    return columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH;
  }, [columnWidths, tempColumnWidths, isResizingColumn]);

  const getCurrentRowHeight = useCallback((rowIndex: number) => {
    if (isResizingRow === rowIndex && tempRowHeights[rowIndex] !== undefined) {
      return tempRowHeights[rowIndex];
    }
    return rowHeights[rowIndex] || DEFAULT_ROW_HEIGHT;
  }, [rowHeights, tempRowHeights, isResizingRow]);

  // Helper to deep clone using JSON for snapshot safety
  const deepClone = useCallback(<T,>(obj: T): T => {
    return JSON.parse(JSON.stringify(obj)) as T;
  }, []);

  const createVersionSnapshot = useCallback((description: string) => {
    const id = `v_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const snapshot: VersionSnapshot = {
      id,
      timestamp: Date.now(),
      description,
      data: {
        sheetData: deepClone(sheetData),
        columnWidths: deepClone(columnWidths),
        rowHeights: deepClone(rowHeights)
      }
    };
    versionSnapshotsRef.current.set(id, snapshot);
    setVersionHistory(prev => {
      const next = [{ id, timestamp: snapshot.timestamp, description: snapshot.description }, ...prev];
      // Keep last 50 versions
      return next.slice(0, 50);
    });
  }, [sheetData, columnWidths, rowHeights, deepClone]);

  // After any command execution (including undo/redo), record a snapshot once state settles
  useEffect(() => {
    if (snapshotRequestRef.current.requested && snapshotRequestRef.current.description) {
      // Take snapshot of latest state
      createVersionSnapshot(snapshotRequestRef.current.description);
      snapshotRequestRef.current.requested = false;
      snapshotRequestRef.current.description = null;
    }
  }, [sheetData, columnWidths, rowHeights, createVersionSnapshot]);

  // Ensure we have an initial snapshot once data initializes
  useEffect(() => {
    if (versionHistory.length === 0) {
      snapshotRequestRef.current = { requested: true, description: 'Initial state' };
    }
  }, [versionHistory.length]);

  // Wrapped execute/undo/redo to request snapshots with accurate descriptions
  const executeWithHistory = useCallback((command: Command) => {
    snapshotRequestRef.current = { requested: true, description: command.description };
    executeCommand(command);
  }, [executeCommand]);

  const handleUndo = useCallback(() => {
    const state = getState();
    const desc = state.undoDescription ? `Undo: ${state.undoDescription}` : 'Undo';
    snapshotRequestRef.current = { requested: true, description: desc };
    undo();
  }, [undo, getState]);

  const handleRedo = useCallback(() => {
    const state = getState();
    const desc = state.redoDescription ? `Redo: ${state.redoDescription}` : 'Redo';
    snapshotRequestRef.current = { requested: true, description: desc };
    redo();
  }, [redo, getState]);

  // Generate column labels A, B, C, ..., Z, AA, AB, etc.
  const generateColumnLabel = useMemo(() => {
    const cache = new Map<number, string>();
    return (index: number): string => {
      if (cache.has(index)) return cache.get(index)!;
      
      let result = '';
      let temp = index;
      while (temp >= 0) {
        result = String.fromCharCode(65 + (temp % 26)) + result;
        temp = Math.floor(temp / 26) - 1;
      }
      cache.set(index, result);
      return result;
    };
  }, []);

  // Render a compact preview table for a snapshot (limited size)
  const renderSnapshotPreview = useCallback((snapshotId: string) => {
    const snap = versionSnapshotsRef.current.get(snapshotId);
    if (!snap) return null;
    const data = snap.data.sheetData;
    const colWidths = snap.data.columnWidths || {};
    const rowHeights = snap.data.rowHeights || {};
    // Helper to compute values (supports simple formulas referencing A1 in the snapshot itself)
    const getSnapshotValue = (row: number, col: number): string | number => {
      const cell = data[row]?.[col];
      const raw = cell?.value;
      if (typeof raw === 'string' && raw.startsWith('=')) {
        const getA1 = (a1: string) => {
          const m = /^\s*([A-Z]+)([0-9]+)\s*$/i.exec(a1);
          if (!m) return '';
          const colLabel = m[1].toUpperCase();
          let idx = 0;
          for (let i = 0; i < colLabel.length; i++) idx = idx * 26 + (colLabel.charCodeAt(i) - 64);
          const c = idx - 1;
          const r = parseInt(m[2], 10) - 1;
          if (r < 0 || c < 0) return '';
          const v = data[r]?.[c]?.value;
          return v ?? '';
        };
        try {
          const v = evaluateFormula(raw, getA1);
          return (v as string | number) ?? '';
        } catch {
          return '#ERROR';
        }
      }
      return (raw as string | number | null | undefined) ?? '';
    };
    // Determine bounds
    const maxRowIndex = Math.max(9, ...Object.keys(data).map(n => Number(n)).filter(n => !isNaN(n)));
    const maxColIndex = Math.max(9, ...Object.values(data).map(row => Math.max(0, ...Object.keys(row).map(n => Number(n)))).filter(n => !isNaN(n)));
    const rows: number[] = Array.from({ length: Math.min(maxRowIndex + 1, 50) }, (_, i) => i);
    const cols: number[] = Array.from({ length: Math.min(maxColIndex + 1, 50) }, (_, i) => i);
    // Basic number formatting for preview (mirrors main formatter without creating dependency order issues)
    const previewFormat = (value: string | number | null, f: CellFormatting): string => {
      if (value === null || value === undefined || value === '') return '';
      const num = Number(value);
      const decimals = f.decimalPlaces ?? 2;
      const fmt = f.numberFormat || 'general';
      if (isNaN(num)) return String(value);
      switch (fmt) {
        case 'currency':
          return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(num);
        case 'percentage':
          return new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(num / 100);
        case 'number':
          return new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(num);
        case 'date':
          try { return new Date(num).toLocaleDateString(); } catch { return String(value); }
        case 'time':
          try { return new Date(num).toLocaleTimeString(); } catch { return String(value); }
        default:
          return String(value);
      }
    };

    return (
      <div style={{ overflow: 'auto', maxHeight: 400, border: '1px solid #dee2e6', borderRadius: 6 }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', background: '#ffffff', color: '#2d3748' }}>
          <thead>
            <tr>
              <th style={{ position: 'sticky', left: 0, top: 0, background: '#f8f9fa', color: '#2d3748', border: '1px solid #dee2e6', padding: 4, fontSize: 11, minWidth: 40 }} />
              {cols.map(c => (
                <th key={c} style={{ position: 'sticky', top: 0, background: '#f8f9fa', color: '#2d3748', border: '1px solid #dee2e6', padding: 4, fontSize: 11, width: (colWidths[c] ?? 80), minWidth: (colWidths[c] ?? 80) }}>
                  {generateColumnLabel(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(r => (
              <tr key={r}>
                <th style={{ position: 'sticky', left: 0, background: '#f8f9fa', color: '#2d3748', border: '1px solid #dee2e6', padding: 4, fontSize: 11, height: (rowHeights[r] ?? 20) }}>{r + 1}</th>
                {cols.map(c => {
                  const cell = data[r]?.[c];
                  const f = (cell?.formatting || {}) as CellFormatting;
                  const val = getSnapshotValue(r, c);
                  const display = previewFormat(val as string | number | null, f);
                  const textDecoration = `${f.underline ? 'underline ' : ''}${f.strikethrough ? 'line-through' : ''}`.trim() || 'none';
                  const content = f.link ? (
                    <a href={f.link} target="_blank" rel="noreferrer" style={{ color: '#0066cc', textDecoration: 'underline' }}>{String(display)}</a>
                  ) : (
                    <span>{String(display)}</span>
                  );
                  return (
                    <td
                      key={c}
                      style={{
                        border: '1px solid #e9ecef',
                        padding: '2px 6px',
                        fontSize: f.fontSize ? Math.max(10, Math.min(24, f.fontSize)) : 12,
                        fontFamily: f.fontFamily || 'inherit',
                        fontWeight: f.bold ? 700 as const : 400 as const,
                        fontStyle: f.italic ? 'italic' : 'normal',
                        textDecoration,
                        color: f.color || '#2d3748',
                        background: f.backgroundColor || 'transparent',
                        whiteSpace: 'nowrap',
                        maxWidth: 180,
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        textAlign: f.textAlign || 'left',
                        verticalAlign: f.verticalAlign || 'middle',
                        width: (colWidths[c] ?? 80),
                        minWidth: (colWidths[c] ?? 80),
                        height: (rowHeights[r] ?? 20)
                      }}
                    >
                      {content}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }, [generateColumnLabel]);

  // Get cell address (e.g., "A1", "B3")
  const getCellAddress = useCallback((row: number, col: number): string => {
    return `${generateColumnLabel(col)}${row + 1}`;
  }, [generateColumnLabel]);

  // Grid expansion functions - Excel-like behavior


  // Smart grid expansion with performance awareness - PRESERVES ALL DATA
  const smartExpandGrid = useCallback((targetRow: number, targetCol: number) => {
    // If we're already at performance threshold, be more conservative  
    const maxSafeExpansion = EXPANSION_BUFFER;
    
    const requiredRows = Math.max(
      MIN_VISIBLE_ROWS,
      Math.min(targetRow + maxSafeExpansion, MAX_ROWS)
    );
    const requiredCols = Math.max(
      MIN_VISIBLE_COLS, 
      Math.min(targetCol + maxSafeExpansion, MAX_COLS)
    );

    // Only expand if we really need to
    if (requiredRows > visibleRows || requiredCols > visibleCols) {
      const newRows = Math.max(visibleRows, requiredRows);
      const newCols = Math.max(visibleCols, requiredCols);
      
      // Prevent explosive growth
      const safeTotalCells = newRows * newCols;
      if (safeTotalCells > PERFORMANCE_THRESHOLD * 4) {
        console.warn('Grid expansion limited for performance');
        return false;
      }
      
      // CRITICAL: Grid expansion ONLY changes visible bounds, NEVER touches sheetData
      if (newRows > visibleRows) setVisibleRows(newRows);
      if (newCols > visibleCols) setVisibleCols(newCols);
      
      return true;
    }
    
    return false;
  }, [visibleRows, visibleCols, MIN_VISIBLE_ROWS, MIN_VISIBLE_COLS, MAX_ROWS, MAX_COLS, EXPANSION_BUFFER, PERFORMANCE_THRESHOLD]);

  // Define shouldUseVirtualScrolling before using it (memoized for performance)
  const shouldUseVirtualScrolling = useMemo(() => {
    return (visibleRows * visibleCols) > PERFORMANCE_THRESHOLD;
  }, [visibleRows, visibleCols]);

  // Smart auto-expansion with performance awareness
  const handleAutoExpansion = useCallback((row: number, col: number) => {
    // Expand if we're within 5 rows/cols of the edge or beyond current bounds
    const shouldExpand = (
      row >= visibleRows - 5 || 
      col >= visibleCols - 5 || 
      row >= visibleRows || 
      col >= visibleCols
    );
    
    if (shouldExpand) {
      smartExpandGrid(row, col);
    }
  }, [visibleRows, visibleCols, smartExpandGrid]);

  // Cleanup RAF on unmount
  useEffect(() => {
    return () => {
      if (selectionUpdateRAFRef.current) {
        cancelAnimationFrame(selectionUpdateRAFRef.current);
      }
    };
  }, []);

  // Virtual scrolling calculations - compute from live scroll positions
  const getViewport = useCallback(() => {
    if (!shouldUseVirtualScrolling) {
      return {
        startRow: 0,
        endRow: visibleRows,
        startCol: 0,
        endCol: visibleCols,
        totalRowHeight: visibleRows * DEFAULT_ROW_HEIGHT,
        totalColWidth: visibleCols * DEFAULT_COLUMN_WIDTH
      };
    }

    const container = scrollContainerRef.current;
    if (!container) {
      return {
        startRow: viewportStartRow,
        endRow: viewportEndRow,
        startCol: viewportStartCol,
        endCol: viewportEndCol,
        totalRowHeight: visibleRows * DEFAULT_ROW_HEIGHT,
        totalColWidth: visibleCols * DEFAULT_COLUMN_WIDTH
      };
    }

    const { scrollTop, scrollLeft, clientHeight, clientWidth } = container;
    const zoomFactor = zoomLevel / 100;

    const startRow = Math.max(0, Math.floor(scrollTop / (DEFAULT_ROW_HEIGHT * zoomFactor)) - RENDER_BUFFER);
    const endRow = Math.min(visibleRows, startRow + Math.ceil(clientHeight / (DEFAULT_ROW_HEIGHT * zoomFactor)) + RENDER_BUFFER * 2);

    const startCol = Math.max(0, Math.floor(scrollLeft / (DEFAULT_COLUMN_WIDTH * zoomFactor)) - RENDER_BUFFER);
    const endCol = Math.min(visibleCols, startCol + Math.ceil(clientWidth / (DEFAULT_COLUMN_WIDTH * zoomFactor)) + RENDER_BUFFER * 2);

    return {
      startRow,
      endRow: Math.min(endRow, startRow + MAX_RENDERED_ROWS),
      startCol,
      endCol: Math.min(endCol, startCol + MAX_RENDERED_COLS),
      totalRowHeight: visibleRows * DEFAULT_ROW_HEIGHT,
      totalColWidth: visibleCols * DEFAULT_COLUMN_WIDTH
    };
  }, [shouldUseVirtualScrolling, visibleRows, visibleCols, viewportStartRow, viewportEndRow, viewportStartCol, viewportEndCol, zoomLevel]);

  // Update viewport when scrolling (optimized with RAF and debouncing)
  const updateViewport = useCallback(() => {
    const viewport = getViewport();
    
    if (viewport.startRow !== viewportStartRow || viewport.endRow !== viewportEndRow ||
        viewport.startCol !== viewportStartCol || viewport.endCol !== viewportEndCol) {
      
      // Use requestAnimationFrame for smoother updates
      requestAnimationFrame(() => {
        setViewportStartRow(viewport.startRow);
        setViewportEndRow(viewport.endRow);
        setViewportStartCol(viewport.startCol);
        setViewportEndCol(viewport.endCol);
      });
    }
  }, [getViewport, viewportStartRow, viewportEndRow, viewportStartCol, viewportEndCol]);
  
  // Create debounced version of viewport update for scroll events
  const debouncedUpdateViewport = useMemo(
    () => createOptimizedHandler(updateViewport, { 
      delay: 16, 
      throttle: true, 
      useRAF: true 
    }),
    [updateViewport]
  );


  // Initialize with sample data using performance optimization - ONLY ONCE
  useEffect(() => {
    // Init worker lazily
    if (!workerRef.current) {
      try {
        workerRef.current = new Worker(GridWorkerUrl, { type: 'module' });
        workerRef.current.addEventListener('message', (e: MessageEvent) => {
          const { requestId, ok, result, error } = (e.data || {}) as { requestId: string; ok: boolean; result?: unknown; error?: string };
          const resolver = pendingRequestsRef.current.get(requestId);
          if (resolver) {
            pendingRequestsRef.current.delete(requestId);
            if (ok) resolver({ ok, result } as { ok: true; result: unknown });
            else resolver({ ok, error } as { ok: false; error?: string });
          }
        });
      } catch {
        // Worker optional; fallback to main thread if construction fails
        workerRef.current = null;
      }
    }
    // Defer initial data setup to improve initial render performance
    runWhenIdle(() => {
      const initialData: SheetData = {};
    
    // Set sample data exactly as in the original
    const setInitialCellValue = (row: number, col: number, value: string | number) => {
      if (!initialData[row]) initialData[row] = {};
      initialData[row][col] = { value };
    };

    setInitialCellValue(0, 1, 'Tesla');
    setInitialCellValue(0, 2, 'Volvo');
    setInitialCellValue(0, 3, 'Toyota');
    setInitialCellValue(0, 4, 'Ford');
    setInitialCellValue(0, 5, 'Total');

    setInitialCellValue(1, 0, '2019');
    setInitialCellValue(1, 1, 10);
    setInitialCellValue(1, 2, 11);
    setInitialCellValue(1, 3, 12);
    setInitialCellValue(1, 4, 13);
    setInitialCellValue(1, 5, 46);

    setInitialCellValue(2, 0, '2020');
    setInitialCellValue(2, 1, 20);
    setInitialCellValue(2, 2, 11);
    setInitialCellValue(2, 3, 14);
    setInitialCellValue(2, 4, 13);
    setInitialCellValue(2, 5, 58);

    setInitialCellValue(3, 0, '2021');
    setInitialCellValue(3, 1, 30);
    setInitialCellValue(3, 2, 15);
    setInitialCellValue(3, 3, 12);
    setInitialCellValue(3, 4, 13);
    setInitialCellValue(3, 5, 70);

    setInitialCellValue(4, 0, 'Total');
    setInitialCellValue(4, 1, 60);
    setInitialCellValue(4, 2, 37);
    setInitialCellValue(4, 3, 38);
    setInitialCellValue(4, 4, 39);
    setInitialCellValue(4, 5, 174);

    setInitialCellValue(6, 0, 'Profit');
    setInitialCellValue(6, 1, 5);
    setInitialCellValue(6, 2, 8);
    setInitialCellValue(6, 3, 6);
    setInitialCellValue(6, 4, 9);
    setInitialCellValue(6, 5, 28);

    setInitialCellValue(7, 0, 'Margin %');
    setInitialCellValue(7, 1, 8.33);
    setInitialCellValue(7, 2, 21.62);
    setInitialCellValue(7, 3, 15.79);
    setInitialCellValue(7, 4, 23.08);
    setInitialCellValue(7, 5, 16.09);

      setSheetData(initialData);
    }, 10, 1000);
  }, []); // FIXED: Empty dependency array so this only runs once on mount

  // Export/Save handlers
  const handleSave = useCallback(() => {
    try {
      const payload = {
        sheetData,
        columnWidths,
        rowHeights,
        timestamp: Date.now()
      };
      const json = JSON.stringify(payload);
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'sheet.json';
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error('Save failed', e);
    }
  }, [sheetData, columnWidths, rowHeights]);

  const handleSaveAsExcel = useCallback(() => {
    // Convert current visible data into a 2D array (values only)
    const maxRow = Math.max(visibleRows, ...Object.keys(sheetData).map(n => Number(n) + 1));
    const maxCol = Math.max(
      visibleCols,
      ...Object.values(sheetData).map(row => Math.max(0, ...Object.keys(row).map(n => Number(n) + 1)))
    );
    const aoa: (string | number)[][] = [];
    for (let r = 0; r < maxRow; r++) {
      const rowArr: (string | number)[] = [];
      for (let c = 0; c < maxCol; c++) {
        rowArr.push(sheetData[r]?.[c]?.value ?? '');
      }
      aoa.push(rowArr);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'sheet.xlsx');
  }, [sheetData, visibleRows, visibleCols]);

  const handleExportPNG = useCallback(async () => {
    if (!scrollContainerRef.current) return;
    const container = scrollContainerRef.current;
    const canvas = await html2canvas(container, { scale: 2, useCORS: true, backgroundColor: '#ffffff' });
    const dataUrl = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = 'sheet.png';
    a.click();
  }, []);

  const handleExportPDF = useCallback(async () => {
    if (!scrollContainerRef.current) return;
    const container = scrollContainerRef.current;
    const canvas = await html2canvas(container, { scale: 2, useCORS: true, backgroundColor: '#ffffff' });
    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF('l', 'pt', 'a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imgWidth = pageWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    let y = 0;
    let remaining = imgHeight;
    // Add pages if needed (simple vertical slice)
    while (remaining > 0) {
      pdf.addImage(imgData, 'PNG', 0, y ? 0 : 0, imgWidth, imgHeight);
      remaining -= pageHeight;
      if (remaining > 0) pdf.addPage();
      y += pageHeight;
    }
    pdf.save('sheet.pdf');
  }, []);

  const handleImportExcelFile = useCallback(async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      const aoa = XLSX.utils.sheet_to_json<(string | number)[]> (ws, { header: 1 }) as (string | number)[][];
      // Apply into sheetData preserving existing formatting when possible
      setSheetData(prev => {
        const newData: typeof prev = {};
        for (let r = 0; r < aoa.length; r++) {
          for (let c = 0; c < (aoa[r]?.length || 0); c++) {
            const value = aoa[r]?.[c];
            if (value !== undefined && value !== null && value !== '') {
              if (!newData[r]) newData[r] = {};
              newData[r][c] = { ...(prev[r]?.[c] || {}), value: value as string | number };
            }
          }
        }
        return newData;
      });
      setVisibleRows(Math.max(visibleRows, aoa.length));
      setVisibleCols(Math.max(visibleCols, Math.max(0, ...aoa.map(r => r.length))));
      setClipboardNotification('Imported Excel successfully');
      setTimeout(() => setClipboardNotification(null), 2000);
    } catch (e) {
      console.error('Import failed', e);
      setClipboardNotification('Import failed');
      setTimeout(() => setClipboardNotification(null), 2000);
    }
  }, [visibleRows, visibleCols]);

  // Smart scroll handling with virtual scrolling and performance optimization (optimized)
  useEffect(() => {
    const container = scrollContainerRef.current;
      if (!container) return;
      
    const handleScroll = () => {
      if (shouldUseVirtualScrolling) {
        debouncedUpdateViewport();
      }
      
      const { scrollTop, scrollHeight, clientHeight, scrollLeft, scrollWidth, clientWidth } = container;
      const verticalScrollPercent = (scrollTop + clientHeight) / scrollHeight;
      const horizontalScrollPercent = (scrollLeft + clientWidth) / scrollWidth;
      const threshold = shouldUseVirtualScrolling ? 0.95 : 0.9;
      
      if (verticalScrollPercent > threshold && visibleRows < MAX_ROWS) {
        const currentCellCount = visibleRows * visibleCols;
        if (currentCellCount < PERFORMANCE_THRESHOLD * 2) {
          smartExpandGrid(visibleRows, visibleCols);
        }
      }
      
      if (horizontalScrollPercent > threshold && visibleCols < MAX_COLS) {
        const currentCellCount = visibleRows * visibleCols;
        if (currentCellCount < PERFORMANCE_THRESHOLD * 2) {
          smartExpandGrid(visibleRows, visibleCols);
        }
      }
    };

    const optimizedHandleScroll = createOptimizedHandler(handleScroll, {
      throttle: true,
      useRAF: true
    });

    container.addEventListener('scroll', optimizedHandleScroll as EventListener, { passive: true });
    return () => {
      container.removeEventListener('scroll', optimizedHandleScroll as EventListener);
    };
  }, [visibleRows, visibleCols, shouldUseVirtualScrolling, debouncedUpdateViewport, smartExpandGrid]);

  // Container resize observer for virtual scrolling
  useEffect(() => {
    const container = scrollContainerRef.current;
    if (!container) return;

    const resizeObserver = new ResizeObserver(() => {
      // Update viewport when container size changes using debounced version
      if (shouldUseVirtualScrolling) {
        debouncedUpdateViewport();
      }
    });

    resizeObserver.observe(container);

    return () => {
      resizeObserver.disconnect();
    };
  }, [shouldUseVirtualScrolling, debouncedUpdateViewport]);





  const getCellValue = useCallback((row: number, col: number): string | number => {
    const raw = sheetData[row]?.[col]?.value;
    if (typeof raw === 'string' && raw.startsWith('=')) {
      const getA1 = (a1: string) => {
        const addr = parseA1(a1);
        if (!addr) return '';
        return sheetData[addr.row]?.[addr.col]?.value ?? '';
      };
      const parseA1 = (a1: string) => {
        const m = /^\s*([A-Z]+)([0-9]+)\s*$/i.exec(a1);
        if (!m) return null;
        const colLabel = m[1].toUpperCase();
        let idx = 0;
        for (let i = 0; i < colLabel.length; i++) idx = idx * 26 + (colLabel.charCodeAt(i) - 64);
        const c = idx - 1;
        const r = parseInt(m[2], 10) - 1;
        if (r < 0 || c < 0) return null;
        return { row: r, col: c };
      };
      try {
        const v = evaluateFormula(raw, getA1);
        return v as string | number;
      } catch {
        return '#ERROR';
      }
    }
    return (raw as string | number | null | undefined) ?? '';
  }, [sheetData]);

  // Internal function to set cell value directly (used by commands)
  const setCellValueDirect = useCallback((row: number, col: number, value: string | number) => {
    // Auto-expand grid if needed when setting cell values
    handleAutoExpansion(row, col);
    
    setSheetData(prev => {
      const newData = { ...prev };
      if (!newData[row]) newData[row] = {};
      
      // If value is empty, delete the cell data entirely
      if (value === '' || value === null || value === undefined) {
        delete newData[row][col];
        // If row is now empty, delete the row
        if (Object.keys(newData[row]).length === 0) {
          delete newData[row];
        }
      } else {
        newData[row][col] = { 
          ...newData[row]?.[col],
          value 
        };
      }
      return newData;
    });
  }, [handleAutoExpansion]);

  // Create an undoable command for setting cell value
  const setCellValue = useCallback((row: number, col: number, newValue: string | number, description?: string) => {
    const currentValue = getCellValue(row, col);
    
    // Don't create command if value hasn't changed
    if (currentValue === newValue) return;
    
    const command: Command = {
      execute: () => setCellValueDirect(row, col, newValue),
      undo: () => setCellValueDirect(row, col, currentValue),
      description: description || `Edit cell ${generateColumnLabel(col)}${row + 1}`
    };
    
    executeWithHistory(command);
  }, [setCellValueDirect, getCellValue, executeWithHistory, generateColumnLabel]);

  const isCellInRange = useCallback((row: number, col: number, range: CellRange | null): boolean => {
    if (!range) return false;
    const minRow = Math.min(range.startRow, range.endRow);
    const maxRow = Math.max(range.startRow, range.endRow);
    const minCol = Math.min(range.startCol, range.endCol);
    const maxCol = Math.max(range.startCol, range.endCol);
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  }, []);

  const isCellSelected = useCallback((row: number, col: number): boolean => {
    if (selectedRange) {
      return isCellInRange(row, col, selectedRange);
    }
    return selectedCell?.row === row && selectedCell?.col === col;
  }, [selectedRange, selectedCell, isCellInRange]);

  const isCellCopied = useCallback((row: number, col: number): boolean => {
    if (!copiedData) return false;
    return isCellInRange(row, col, copiedData.range);
  }, [copiedData, isCellInRange]);

  const isCellCut = useCallback((row: number, col: number): boolean => {
    if (!copiedData || !copiedData.isCut) return false;
    return isCellInRange(row, col, copiedData.range);
  }, [copiedData, isCellInRange]);

  const isCellInFillPreview = useCallback((row: number, col: number): boolean => {
    if (!fillPreview) return false;
    return isCellInRange(row, col, fillPreview);
  }, [fillPreview, isCellInRange]);





  // Perform the fill operation
  const performFillOperation = useCallback((sourceRange: CellRange, targetRange: CellRange) => {
    const sourceMinRow = Math.min(sourceRange.startRow, sourceRange.endRow);
    const sourceMaxRow = Math.max(sourceRange.startRow, sourceRange.endRow);
    const sourceMinCol = Math.min(sourceRange.startCol, sourceRange.endCol);
    const sourceMaxCol = Math.max(sourceRange.startCol, sourceRange.endCol);
    
    const targetMinRow = Math.min(targetRange.startRow, targetRange.endRow);
    const targetMaxRow = Math.max(targetRange.startRow, targetRange.endRow);
    const targetMinCol = Math.min(targetRange.startCol, targetRange.endCol);
    const targetMaxCol = Math.max(targetRange.startCol, targetRange.endCol);

    // We don't need to determine specific fill direction since modulo handles all cases

    // Get source values - always collect ALL values from the source range
    // The pattern will be applied based on the fill direction
    const sourceValues: (string | number)[] = [];
    const sourceRowCount = sourceMaxRow - sourceMinRow + 1;
    const sourceColCount = sourceMaxCol - sourceMinCol + 1;
    
    // Always collect all values from source range row by row
    for (let row = sourceMinRow; row <= sourceMaxRow; row++) {
      for (let col = sourceMinCol; col <= sourceMaxCol; col++) {
        sourceValues.push(getCellValue(row, col));
      }
    }

    // Get original values that will be overwritten
    const originalValues: Array<{row: number, col: number, value: string | number}> = [];
    for (let row = targetMinRow; row <= targetMaxRow; row++) {
      for (let col = targetMinCol; col <= targetMaxCol; col++) {
        // Skip source cells
        if (row >= sourceMinRow && row <= sourceMaxRow && col >= sourceMinCol && col <= sourceMaxCol) {
          continue;
        }
        originalValues.push({
          row,
          col,
          value: getCellValue(row, col)
        });
      }
    }

    const command: Command = {
      execute: async () => {
        const worker = workerRef.current;
        if (worker) {
          const requestId = Math.random().toString(36).slice(2);
          const payload = {
            sourceValues,
            sourceRowCount,
            sourceColCount,
            sourceMinRow,
            sourceMinCol,
            targetMinRow,
            targetMaxRow,
            targetMinCol,
            targetMaxCol
          };
          type FillResult = { updates: Array<{ row: number; col: number; value: string | number }> };
          const response = new Promise<{ ok: boolean; result?: FillResult; error?: string }>((resolve) => {
            pendingRequestsRef.current.set(requestId, resolve as (v: unknown) => void);
          });
          worker.postMessage({ requestId, type: 'computeFill', payload });
          const { ok, result } = (await response) as { ok: boolean; result?: FillResult };
          if (ok && result) {
            setSheetData(prev => {
              const newData = { ...prev };
              for (const u of result.updates) {
                if (!newData[u.row]) newData[u.row] = {} as { [col: number]: CellData };
                newData[u.row][u.col] = { ...(newData[u.row]?.[u.col] || {}), value: u.value } as CellData;
              }
              return newData;
            });
          }
        } else {
          // Fallback to main-thread loop
        for (let targetRow = targetMinRow; targetRow <= targetMaxRow; targetRow++) {
          for (let targetCol = targetMinCol; targetCol <= targetMaxCol; targetCol++) {
              if (targetRow >= sourceMinRow && targetRow <= sourceMaxRow && targetCol >= sourceMinCol && targetCol <= sourceMaxCol) {
              continue;
            }
            const relativeRow = (targetRow - sourceMinRow) % sourceRowCount;
            const relativeCol = (targetCol - sourceMinCol) % sourceColCount;
            const sourceIndex = relativeRow * sourceColCount + relativeCol;
            if (sourceIndex < sourceValues.length) {
              const value = sourceValues[sourceIndex];
              setCellValueDirect(targetRow, targetCol, value);
              }
            }
          }
        }
      },
      undo: () => {
        // Restore original values
        originalValues.forEach(({row, col, value}) => {
          setCellValueDirect(row, col, value);
        });
      },
      description: `Fill cells from ${generateColumnLabel(sourceMinCol)}${sourceMinRow + 1} to ${generateColumnLabel(targetMaxCol)}${targetMaxRow + 1}`
    };

    executeWithHistory(command);
  }, [getCellValue, setCellValueDirect, executeWithHistory, generateColumnLabel]);

  // Fill handle mouse event handlers
  const handleFillHandleMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    if (!selectedCell && !selectedRange) return;
    
    setFillHandleActive(true);
    setDragStartCell(selectedCell);
    
    // Don't clear copied data here - let user paste multiple times
  }, [selectedCell, selectedRange]);

  const handleFillHandleMouseMove = useCallback((row: number, col: number) => {
    if (!fillHandleActive || !dragStartCell) return;
    
    // Determine the source range
    let sourceRange: CellRange;
    if (selectedRange) {
      sourceRange = selectedRange;
    } else {
      sourceRange = {
        startRow: dragStartCell.row,
        startCol: dragStartCell.col,
        endRow: dragStartCell.row,
        endCol: dragStartCell.col
      };
    }
    
    // Create preview range from source to current mouse position
    const previewRange: CellRange = {
      startRow: Math.min(sourceRange.startRow, sourceRange.endRow, row),
      startCol: Math.min(sourceRange.startCol, sourceRange.endCol, col),
      endRow: Math.max(sourceRange.startRow, sourceRange.endRow, row),
      endCol: Math.max(sourceRange.startCol, sourceRange.endCol, col)
    };
    
    setFillPreview(previewRange);
  }, [fillHandleActive, dragStartCell, selectedRange]);

  const handleFillHandleMouseUp = useCallback(() => {
    if (!fillHandleActive || !dragStartCell || !fillPreview) {
      setFillHandleActive(false);
      setFillPreview(null);
      setDragStartCell(null);
      return;
    }
    
    // Determine the source range
    let sourceRange: CellRange;
    if (selectedRange) {
      sourceRange = selectedRange;
    } else {
      sourceRange = {
        startRow: dragStartCell.row,
        startCol: dragStartCell.col,
        endRow: dragStartCell.row,
        endCol: dragStartCell.col
      };
    }
    
    // Check if we actually moved beyond the source range
    const sourceMinRow = Math.min(sourceRange.startRow, sourceRange.endRow);
    const sourceMaxRow = Math.max(sourceRange.startRow, sourceRange.endRow);
    const sourceMinCol = Math.min(sourceRange.startCol, sourceRange.endCol);
    const sourceMaxCol = Math.max(sourceRange.startCol, sourceRange.endCol);
    
    const previewMinRow = Math.min(fillPreview.startRow, fillPreview.endRow);
    const previewMaxRow = Math.max(fillPreview.startRow, fillPreview.endRow);
    const previewMinCol = Math.min(fillPreview.startCol, fillPreview.endCol);
    const previewMaxCol = Math.max(fillPreview.startCol, fillPreview.endCol);
    
    // Only perform fill if the preview extends beyond the source
    if (previewMaxRow > sourceMaxRow || previewMaxCol > sourceMaxCol || 
        previewMinRow < sourceMinRow || previewMinCol < sourceMinCol) {
      performFillOperation(sourceRange, fillPreview);
    }
    
    // Clean up
    setFillHandleActive(false);
    setFillPreview(null);
    setDragStartCell(null);
  }, [fillHandleActive, dragStartCell, fillPreview, selectedRange, performFillOperation]);

  // Handle global mouse up to end selection and fill handle
  useEffect(() => {
    const handleGlobalMouseUp = () => {
      setIsSelecting(false);
      if (fillHandleActive) {
        handleFillHandleMouseUp();
      }
    };

    document.addEventListener('mouseup', handleGlobalMouseUp);
    return () => {
      document.removeEventListener('mouseup', handleGlobalMouseUp);
    };
  }, [fillHandleActive, handleFillHandleMouseUp]);

  const getSelectedCells = useCallback((): {row: number, col: number}[] => {
    if (selectedRange) {
      const cells: {row: number, col: number}[] = [];
      const minRow = Math.min(selectedRange.startRow, selectedRange.endRow);
      const maxRow = Math.max(selectedRange.startRow, selectedRange.endRow);
      const minCol = Math.min(selectedRange.startCol, selectedRange.endCol);
      const maxCol = Math.max(selectedRange.startCol, selectedRange.endCol);
      
      for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
          cells.push({row, col});
        }
      }
      return cells;
    }
    return selectedCell ? [selectedCell] : [];
  }, [selectedCell, selectedRange]);

  // Get highlighted columns and rows based on selection (memoized for performance)
  const getHighlightedColumns = useMemo((): Set<number> => {
    const highlightedCols = new Set<number>();
    
    if (selectedRange) {
      const minCol = Math.min(selectedRange.startCol, selectedRange.endCol);
      const maxCol = Math.max(selectedRange.startCol, selectedRange.endCol);
      for (let col = minCol; col <= maxCol; col++) {
        highlightedCols.add(col);
      }
    } else if (selectedCell) {
      highlightedCols.add(selectedCell.col);
    }
    
    return highlightedCols;
  }, [selectedCell, selectedRange]);

  const getHighlightedRows = useMemo((): Set<number> => {
    const highlightedRows = new Set<number>();
    
    if (selectedRange) {
      const minRow = Math.min(selectedRange.startRow, selectedRange.endRow);
      const maxRow = Math.max(selectedRange.startRow, selectedRange.endRow);
      for (let row = minRow; row <= maxRow; row++) {
        highlightedRows.add(row);
      }
    } else if (selectedCell) {
      highlightedRows.add(selectedCell.row);
    }
    
    return highlightedRows;
  }, [selectedCell, selectedRange]);

  // Clipboard notification helper
  const showClipboardNotification = useCallback((message: string) => {
    setClipboardNotification(message);
    setTimeout(() => setClipboardNotification(null), 3000);
  }, []);

  const copySelectedCells = useCallback((isCut: boolean = false) => {
    const selectedCells = getSelectedCells();
    if (selectedCells.length === 0) {
      showClipboardNotification('No cells selected to copy');
      return;
    }

    let minRow = selectedCells[0].row;
    let maxRow = selectedCells[0].row;
    let minCol = selectedCells[0].col;
    let maxCol = selectedCells[0].col;

    selectedCells.forEach(({row, col}) => {
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    });

    const copiedRange: CellRange = {
      startRow: minRow,
      startCol: minCol,
      endRow: maxRow,
      endCol: maxCol
    };

    const data: CellData[][] = [];
    
    for (let row = minRow; row <= maxRow; row++) {
      const rowData: CellData[] = [];
      for (let col = minCol; col <= maxCol; col++) {
        const value = getCellValue(row, col);
        const cellData = sheetData[row]?.[col];
        
        // Include both value and formatting
        rowData.push({ 
          value: value === null || value === undefined ? '' : value,
          formatting: cellData?.formatting // Include formatting from the source cell
        });
      }
      data.push(rowData);
    }
    
    const copiedDataObj = { data, range: copiedRange, isCut };
    console.log('✅ COPY: Data stored with formatting');
    setCopiedData(copiedDataObj);

    // Show notification
    const cellCount = selectedCells.length;
    const rangeText = cellCount === 1 
      ? `${generateColumnLabel(minCol)}${minRow + 1}`
      : `${generateColumnLabel(minCol)}${minRow + 1}:${generateColumnLabel(maxCol)}${maxRow + 1}`;
    
    const action = isCut ? 'Cut' : 'Copied';
    showClipboardNotification(`${action} ${cellCount} cell${cellCount > 1 ? 's' : ''} (${rangeText})`);

    // Auto-clear styling after 3 seconds for copy, but keep cut cells marked until paste
    if (!isCut) {
      setTimeout(() => {
        // Only clear if the same data is still copied (not replaced by new copy/cut)
        setCopiedData(current => {
          if (current && !current.isCut && current.range === copiedRange) {
            return null;
          }
          return current;
        });
      }, 3000);
    }
    // For cut operations, the data stays until paste

    // Also copy to system clipboard
    // Make sure empty cells are properly represented as empty strings in tab-separated format
    const textData = data.map(row => 
      row.map(cell => {
        const value = cell.value;
        // Handle null, undefined, and empty string cases properly
        if (value === null || value === undefined || value === '') {
          return '';
        }
        return String(value);
      }).join('\t')
    ).join('\n');
    
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(textData).catch(() => {
        // Fallback if clipboard API fails
        showClipboardNotification('System clipboard unavailable, data copied internally');
      });
    }
  }, [getSelectedCells, getCellValue, generateColumnLabel, showClipboardNotification, sheetData]);

  const cutSelectedCells = useCallback(() => {
    copySelectedCells(true);
    // Cut cells internally without visual feedback
  }, [copySelectedCells]);

  // Forward declaration - pasteData is defined below but needed here
  const pasteDataRef = useRef<(() => void) | null>(null);
  
  // Keep a ref to copiedData to avoid stale closures in event handlers
  const copiedDataRef = useRef(copiedData);
  useEffect(() => {
    copiedDataRef.current = copiedData;
  }, [copiedData]);

  const pasteFromText = useCallback((text: string) => {
    if (!selectedCell) {
      return;
    }
    
    // If we have cut data, paste that instead and clear the cut cells
    const currentCopiedData = copiedDataRef.current;
    if (currentCopiedData?.isCut && pasteDataRef.current) {
      pasteDataRef.current();
      return;
    }

    const rows = text.split('\n'); // Don't trim here to preserve empty rows
    const data = rows.map(row => {
      // Split by tabs and preserve empty cells (consecutive tabs = empty cells)
      const cells = row.split('\t');
      return cells; // Each empty string between tabs represents an empty cell
    });

    const startRow = selectedCell.row;
    const startCol = selectedCell.col;

    // Get original values that will be overwritten
    const originalValues: Array<{row: number, col: number, value: string | number}> = [];
    
    data.forEach((rowData, rowIndex) => {
      rowData.forEach((_, colIndex) => {
        const targetRow = startRow + rowIndex;
        const targetCol = startCol + colIndex;
        originalValues.push({
          row: targetRow,
          col: targetCol,
          value: getCellValue(targetRow, targetCol)
        });
      });
    });

    // Handle cut operation cleanup
    const cutClearValues: Array<{row: number, col: number, value: string | number}> = [];
    if (currentCopiedData?.isCut && currentCopiedData.range) {
      const { range } = currentCopiedData;
      const minRow = Math.min(range.startRow, range.endRow);
      const maxRow = Math.max(range.startRow, range.endRow);
      const minCol = Math.min(range.startCol, range.endCol);
      const maxCol = Math.max(range.startCol, range.endCol);

      for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
          // Store all cut cells - we'll handle overlap during paste execution
          cutClearValues.push({
            row,
            col,
            value: getCellValue(row, col)
          });
        }
      }
    }

    const command: Command = {
      execute: () => {
        // Do everything in a single state update for proper cut functionality
        setSheetData(prev => {
          const newData = { ...prev };
          
          // First, apply pasted data
          data.forEach((rowData, rowIndex) => {
            rowData.forEach((cellValue, colIndex) => {
              const targetRow = startRow + rowIndex;
              const targetCol = startCol + colIndex;
              
              if (!newData[targetRow]) newData[targetRow] = {};
              
              // Properly handle empty cells and conversion
              if (cellValue === '' || cellValue === null || cellValue === undefined) {
                // Delete empty cells
                delete newData[targetRow][targetCol];
              } else {
                // Convert to number if it's a valid number, otherwise keep as string
                const stringValue = String(cellValue);
                const trimmedValue = stringValue.trim();
                
                let finalValue;
                if (trimmedValue !== '' && !isNaN(Number(trimmedValue))) {
                  finalValue = Number(trimmedValue);
                } else {
                  finalValue = stringValue;
                }
                
                newData[targetRow][targetCol] = {
                  ...newData[targetRow]?.[targetCol],
                  value: finalValue
                };
              }
            });
          });
          
          // Then, if it was a cut operation, clear the original cells
          if (currentCopiedData?.isCut && cutClearValues.length > 0) {
            cutClearValues.forEach(({row, col}) => {
              // Check if this cell was part of the paste target area
              const wasPastedTo = row >= startRow && row < startRow + data.length &&
                                  col >= startCol && col < startCol + (data[0]?.length || 0);
              
              // Only clear if it wasn't pasted to
              if (!wasPastedTo) {
                if (newData[row]) {
                  delete newData[row][col];
                  // If row is now empty, delete the row
                  if (Object.keys(newData[row]).length === 0) {
                    delete newData[row];
                  }
                }
              }
            });
          }
          
          return newData;
        });

        // Clear copied data after pasting
        setCopiedData(null);
      },
      undo: () => {
        // Restore original values
        originalValues.forEach(({row, col, value}) => {
          setCellValueDirect(row, col, value);
        });

        // Restore cut cells
        cutClearValues.forEach(({row, col, value}) => {
          setCellValueDirect(row, col, value);
        });

        // Restore copied data if it was cut
        if (currentCopiedData?.isCut) {
          setCopiedData(currentCopiedData);
        }
      },
      description: `Paste to ${generateColumnLabel(startCol)}${startRow + 1}`
    };

    executeWithHistory(command);
  }, [selectedCell, getCellValue, setCellValueDirect, executeWithHistory, generateColumnLabel]);

  const pasteData = useCallback(async () => {
    if (!copiedData || !selectedCell) {
      showClipboardNotification('No data to paste');
      return;
    }

    const startRow = selectedCell.row;
    const startCol = selectedCell.col;
    const { data, range, isCut } = copiedData;
    
    console.log('✅ PASTE: Applying data with formatting');

    // Calculate paste area
    const pasteRows = data.length;
    const pasteCols = data[0]?.length || 0;
    const pasteEndRow = startRow + pasteRows - 1;
    const pasteEndCol = startCol + pasteCols - 1;

    // Show notification
    const pasteRangeText = `${generateColumnLabel(startCol)}${startRow + 1}:${generateColumnLabel(pasteEndCol)}${pasteEndRow + 1}`;
    const cellCount = pasteRows * pasteCols;
    showClipboardNotification(`Pasted ${cellCount} cell${cellCount > 1 ? 's' : ''} to ${pasteRangeText}`);

    // Get original values and formatting that will be overwritten
    const originalValues: Array<{row: number, col: number, value: string | number, formatting?: CellFormatting}> = [];
    
    data.forEach((rowData, rowIndex) => {
      rowData.forEach((_, colIndex) => {
        const targetRow = startRow + rowIndex;
        const targetCol = startCol + colIndex;
        const cellData = sheetData[targetRow]?.[targetCol];
        originalValues.push({
          row: targetRow,
          col: targetCol,
          value: getCellValue(targetRow, targetCol),
          formatting: cellData?.formatting
        });
      });
    });

    // Handle cut operation cleanup
    const cutClearValues: Array<{row: number, col: number, value: string | number}> = [];
    if (isCut && range) {
      const minRow = Math.min(range.startRow, range.endRow);
      const maxRow = Math.max(range.startRow, range.endRow);
      const minCol = Math.min(range.startCol, range.endCol);
      const maxCol = Math.max(range.startCol, range.endCol);

      for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
          // Store all cut cells - we'll handle overlap during paste execution
          cutClearValues.push({
            row,
            col,
            value: getCellValue(row, col)
          });
        }
      }
    }

    const command: Command = {
      execute: async () => {
        const requestId = Math.random().toString(36).slice(2);
        const worker = workerRef.current;
        if (worker) {
          const payload = {
            startRow,
            startCol,
            data: data.map(row => row.map(cell => ({ value: cell.value ?? '', formatting: cell.formatting }))),
            isCut,
            cutRange: range ? {
              startRow: Math.min(range.startRow, range.endRow),
              endRow: Math.max(range.startRow, range.endRow),
              startCol: Math.min(range.startCol, range.endCol),
              endCol: Math.max(range.startCol, range.endCol)
            } : null
          };
          type PasteResult = { updates: Array<{ row: number; col: number; value: string | number; formatting?: CellFormatting }>; clears: Array<{ row: number; col: number }> };
          const response = new Promise<{ ok: boolean; result?: PasteResult; error?: string }>((resolve) => {
            pendingRequestsRef.current.set(requestId, resolve as (v: unknown) => void);
          });
          worker.postMessage({ requestId, type: 'computePaste', payload });
          const { ok, result } = (await response) as { ok: boolean; result?: PasteResult };
          if (ok && result) {
            setSheetData(prev => {
              const newData = { ...prev } as { [row: number]: { [col: number]: CellData } };
              for (const u of result.updates) {
                if (!newData[u.row]) newData[u.row] = {} as { [col: number]: CellData };
                newData[u.row][u.col] = { ...(newData[u.row]?.[u.col] || {}), value: u.value, formatting: u.formatting } as CellData;
              }
              for (const cl of result.clears) {
                if (newData[cl.row]) {
                  delete newData[cl.row][cl.col];
                  if (Object.keys(newData[cl.row]).length === 0) delete newData[cl.row];
                }
              }
              return newData;
            });
          } else {
            // Fallback: do nothing special; could add main-thread fallback here
          }
        } else {
          // Fallback to main thread path if worker not available (existing logic suffices)
        setSheetData(prev => {
          const newData = { ...prev };
          data.forEach((rowData, rowIndex) => {
            rowData.forEach((cellData, colIndex) => {
              const targetRow = startRow + rowIndex;
              const targetCol = startCol + colIndex;
              if (!newData[targetRow]) newData[targetRow] = {};
              const value = cellData.value;
              if (value === null || value === undefined || value === '') {
                if (cellData.formatting && Object.keys(cellData.formatting).length > 0) {
                    newData[targetRow][targetCol] = { ...newData[targetRow]?.[targetCol], value: '', formatting: cellData.formatting };
                } else {
                  delete newData[targetRow][targetCol];
                }
              } else {
                  newData[targetRow][targetCol] = { ...newData[targetRow]?.[targetCol], value, formatting: cellData.formatting };
              }
            });
          });
          if (isCut && cutClearValues.length > 0) {
              cutClearValues.forEach(({ row, col }) => {
                const wasPastedTo = row >= startRow && row < startRow + data.length && col >= startCol && col < startCol + (data[0]?.length || 0);
                if (!wasPastedTo && newData[row]) {
                  delete newData[row][col];
                  if (Object.keys(newData[row]).length === 0) delete newData[row];
              }
            });
          }
          return newData;
        });
        }

        setCopiedData(null);
      },
      undo: () => {
        // Restore original values and formatting
        setSheetData(prev => {
          const newData = { ...prev };
          
          // Restore original values with their formatting
          originalValues.forEach(({row, col, value, formatting}) => {
            if (!newData[row]) newData[row] = {};
            
            if (value === '' || value === null || value === undefined) {
              if (formatting && Object.keys(formatting).length > 0) {
                // Keep cell with formatting but empty value
                newData[row][col] = { value: '', formatting };
              } else {
                // Delete cell entirely
                delete newData[row][col];
                if (Object.keys(newData[row]).length === 0) {
                  delete newData[row];
                }
              }
            } else {
              newData[row][col] = { value, formatting };
            }
          });
          
          // Restore cut cells
          cutClearValues.forEach(({row, col, value}) => {
            if (!newData[row]) newData[row] = {};
            if (value === '' || value === null || value === undefined) {
              delete newData[row][col];
              if (Object.keys(newData[row]).length === 0) {
                delete newData[row];
              }
            } else {
              newData[row][col] = { ...newData[row]?.[col], value };
            }
          });
          
          return newData;
        });

        // Restore copied data if it was cut
        if (isCut) {
          setCopiedData(copiedData);
        }
      },
      description: `Paste to ${generateColumnLabel(startCol)}${startRow + 1}`
    };

    executeWithHistory(command);
  }, [copiedData, selectedCell, getCellValue, executeWithHistory, generateColumnLabel, showClipboardNotification, sheetData]);

  // Set the ref for pasteData so pasteFromText can use it
  useEffect(() => {
    pasteDataRef.current = pasteData;
  }, [pasteData]);
  


  const deleteSelectedCells = useCallback(() => {
    const selectedCells = getSelectedCells();
    if (selectedCells.length === 0) return;

    // Get all current values before deletion
    const cellData = selectedCells.map(({row, col}) => ({
      row,
      col,
      value: getCellValue(row, col)
    }));

    // Only create command if there are non-empty cells to delete
    const nonEmptyCells = cellData.filter(cell => cell.value !== '' && cell.value !== null && cell.value !== undefined);
    if (nonEmptyCells.length === 0) return;

    const command: Command = {
      execute: () => {
        selectedCells.forEach(({row, col}) => {
          setCellValueDirect(row, col, '');
        });
      },
      undo: () => {
        cellData.forEach(({row, col, value}) => {
          setCellValueDirect(row, col, value);
        });
      },
      description: selectedCells.length === 1 
        ? `Delete cell ${generateColumnLabel(selectedCells[0].col)}${selectedCells[0].row + 1}`
        : `Delete ${selectedCells.length} cells`
    };

    executeWithHistory(command);
  }, [getSelectedCells, getCellValue, setCellValueDirect, executeWithHistory, generateColumnLabel]);

  const selectAllCells = useCallback(() => {
    setSelectedRange({
      startRow: 0,
      startCol: 0,
      endRow: visibleRows - 1,
      endCol: visibleCols - 1
    });
    setSelectedCell({ row: 0, col: 0 });
  }, [visibleRows, visibleCols]);

  // Column and row selection handlers
  const selectEntireColumn = useCallback((colIndex: number) => {
    setSelectedRange({
      startRow: 0,
      startCol: colIndex,
      endRow: visibleRows - 1,
      endCol: colIndex
    });
    setSelectedCell({ row: 0, col: colIndex });
    // Don't clear copied data here - let user paste multiple times
  }, [visibleRows]);

  const selectEntireRow = useCallback((rowIndex: number) => {
    setSelectedRange({
      startRow: rowIndex,
      startCol: 0,
      endRow: rowIndex,
      endCol: visibleCols - 1
    });
    setSelectedCell({ row: rowIndex, col: 0 });
    // Don't clear copied data here - let user paste multiple times
  }, [visibleCols]);

  const handleColumnHeaderClick = useCallback((colIndex: number, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    
    if (event.shiftKey && selectedRange) {
      // Extend column selection with Shift+click
      const startCol = Math.min(selectedRange.startCol, colIndex);
      const endCol = Math.max(selectedRange.endCol, colIndex);
      setSelectedRange({
        startRow: 0,
        startCol: startCol,
        endRow: visibleRows - 1,
        endCol: endCol
      });
    } else {
      selectEntireColumn(colIndex);
    }
  }, [selectEntireColumn, selectedRange, visibleRows]);

  const handleRowHeaderClick = useCallback((rowIndex: number, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    
    if (event.shiftKey && selectedRange) {
      // Extend row selection with Shift+click
      const startRow = Math.min(selectedRange.startRow, rowIndex);
      const endRow = Math.max(selectedRange.endRow, rowIndex);
      setSelectedRange({
        startRow: startRow,
        startCol: 0,
        endRow: endRow,
        endCol: visibleCols - 1
      });
    } else {
      selectEntireRow(rowIndex);
    }
  }, [selectEntireRow, selectedRange, visibleCols]);

  const handleSelectAllClick = useCallback((event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    selectAllCells();
  }, [selectAllCells]);

  const commitInputBarEdit = useCallback(() => {
    if (selectedCell && isEditingInInputBar) {
      const numValue = Number(inputValue);
      const finalValue = !isNaN(numValue) && inputValue.trim() !== '' ? numValue : inputValue;
      setCellValue(selectedCell.row, selectedCell.col, finalValue);
      
      // CRITICAL FIX: Immediately clear all editing state
      setIsEditingInInputBar(false);
      setEditingCell(null);
      // Don't clear inputValue here as it will be set by the cell selection
    }
  }, [selectedCell, isEditingInInputBar, inputValue, setCellValue]);

  const commitEdit = useCallback(() => {
    if (editingCell) {
      const numValue = Number(inputValue);
      const finalValue = !isNaN(numValue) && inputValue.trim() !== '' ? numValue : inputValue;
      setCellValue(editingCell.row, editingCell.col, finalValue);
      
      // CRITICAL FIX: Immediately clear all editing state
      setEditingCell(null);
      setIsEditingInInputBar(false);
      setIsFreshEdit(false); // Clear fresh edit flag
      setInputValue('');
    } else if (isEditingInInputBar) {
      commitInputBarEdit();
      // Ensure input bar editing state is cleared
      setIsEditingInInputBar(false);
      setEditingCell(null);
      setIsFreshEdit(false); // Clear fresh edit flag
    }
  }, [editingCell, inputValue, setCellValue, isEditingInInputBar, commitInputBarEdit]);

  const handlePasteFromClipboard = useCallback(async () => {
    const currentCopiedData = copiedDataRef.current;
    console.log('📋 PASTE: Attempting paste, hasData:', !!currentCopiedData);
    try {
      // Check if we have internal cut/copy data first
      if (currentCopiedData) {
        // Always prefer internal data if it exists (to properly handle cut operations)
        pasteData();
      } else if (navigator.clipboard && navigator.clipboard.readText) {
        // Only use clipboard if we don't have internal data
        const clipboardData = await navigator.clipboard.readText();
        if (clipboardData.trim()) {
          pasteFromText(clipboardData);
        } else {
          showClipboardNotification('No data to paste');
        }
      } else {
        showClipboardNotification('No data to paste');
      }
    } catch {
      // Fallback to internal copied data
      if (currentCopiedData) {
        pasteData();
      } else {
        showClipboardNotification('Unable to access clipboard');
      }
    }
  }, [pasteFromText, pasteData, showClipboardNotification]);

  // Cell formatting functions
  const getCurrentCellFormatting = useCallback((): CellFormatting => {
    if (!selectedCell) return {};
    
    const cell = sheetData[selectedCell.row]?.[selectedCell.col];
    return cell?.formatting || {};
  }, [selectedCell, sheetData]);

  // Helper function to format numbers based on format type
  const formatCellValue = useCallback((value: string | number | null, formatting?: CellFormatting): string => {
    if (value === null || value === undefined || value === '') return '';
    
    const numValue = Number(value);
    const format = formatting?.numberFormat || 'general';
    const decimals = formatting?.decimalPlaces ?? 2;
    
    if (isNaN(numValue)) return String(value);
    
    switch (format) {
      case 'currency':
        return new Intl.NumberFormat('en-US', {
          style: 'currency',
          currency: 'USD',
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        }).format(numValue);
      
      case 'percentage':
        return new Intl.NumberFormat('en-US', {
          style: 'percent',
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        }).format(numValue / 100);
      
      case 'number':
        return new Intl.NumberFormat('en-US', {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        }).format(numValue);
      
      case 'date':
        try {
          return new Date(numValue).toLocaleDateString();
        } catch {
          return String(value);
        }
      
      case 'time':
        try {
          return new Date(numValue).toLocaleTimeString();
        } catch {
          return String(value);
        }
      
      default:
        return String(value);
    }
  }, []);

  // Helper function to normalize URL
  const normalizeUrl = useCallback((url: string): string => {
    if (!url) return '';
    
    // If URL already has protocol, return as is
    if (url.match(/^https?:\/\//i)) {
      return url;
    }
    
    // If it looks like an email, add mailto:
    if (url.includes('@') && !url.includes(' ')) {
      return `mailto:${url}`;
    }
    
    // Default to https:// for web URLs
    return `https://${url}`;
  }, []);

  const handleInsertLink = useCallback((url: string, text?: string) => {
    if (!selectedCell) return;
    
    const currentValue = getCellValue(selectedCell.row, selectedCell.col);
    const currentFormatting = sheetData[selectedCell.row]?.[selectedCell.col]?.formatting || {};
    const normalizedUrl = normalizeUrl(url);
    
    const command: Command = {
      execute: () => {
        setSheetData(prev => ({
          ...prev,
          [selectedCell.row]: {
            ...prev[selectedCell.row],
            [selectedCell.col]: {
              ...prev[selectedCell.row]?.[selectedCell.col],
              value: text || String(currentValue),
              formatting: {
                ...currentFormatting,
                link: normalizedUrl,
                color: '#0066cc',
                underline: true
              }
            }
          }
        }));
      },
      undo: () => {
        setSheetData(prev => ({
          ...prev,
          [selectedCell.row]: {
            ...prev[selectedCell.row],
            [selectedCell.col]: {
              ...prev[selectedCell.row]?.[selectedCell.col],
              value: currentValue,
              formatting: currentFormatting
            }
          }
        }));
      },
      description: `Insert link in ${generateColumnLabel(selectedCell.col)}${selectedCell.row + 1}`
    };
    
    executeWithHistory(command);
  }, [selectedCell, getCellValue, sheetData, executeWithHistory, generateColumnLabel, normalizeUrl]);

  const formatSelectedCells = useCallback((formatting: Partial<CellFormatting>) => {
    const selectedCells = getSelectedCells();
    if (selectedCells.length === 0) return;
    


    // Get current formatting for all selected cells for undo
    const previousFormatting = selectedCells.map(({row, col}) => ({
      row,
      col,
      formatting: sheetData[row]?.[col]?.formatting || {}
    }));

    const command: Command = {
      execute: () => {
        setSheetData(prev => {
          const newData = { ...prev };
          selectedCells.forEach(({row, col}) => {
            if (!newData[row]) newData[row] = {};
            if (!newData[row][col]) newData[row][col] = { value: prev[row]?.[col]?.value ?? '' };
            
            const oldFormatting = newData[row][col].formatting || {};
            const newFormatting = {
              ...oldFormatting,
              ...formatting
            };
            
            newData[row][col] = {
              ...newData[row][col],
              formatting: newFormatting
            };
            

          });
          return newData;
        });
      },
      undo: () => {
        setSheetData(prev => {
          const newData = { ...prev };
          previousFormatting.forEach(({row, col, formatting}) => {
            if (!newData[row]) newData[row] = {};
            if (!newData[row][col]) newData[row][col] = { value: prev[row]?.[col]?.value ?? '' };
            
            newData[row][col] = {
              ...newData[row][col],
              formatting
            };
          });
          return newData;
        });
      },
      description: `Format ${selectedCells.length} cell(s)`
    };

    executeWithHistory(command);
  }, [getSelectedCells, sheetData, executeWithHistory]);



  // Handle global keyboard events for copy/paste/undo/redo
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      // Prevent shortcuts when typing in input fields or editing cells
      const target = e.target as HTMLElement;
      const isTypingInInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable;
      
      // Don't handle shortcuts when editing a cell or typing in input fields
      if (editingCell || isTypingInInput) return;
      
      // Only handle shortcuts when we have a valid selection in the spreadsheet
      if (!selectedCell && !selectedRange) return;

      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'c' || e.key === 'C') {
          copySelectedCells();
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'v' || e.key === 'V') {
          handlePasteFromClipboard();
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'x' || e.key === 'X') {
          cutSelectedCells();
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'a' || e.key === 'A') {
          selectAllCells();
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'z' || e.key === 'Z') {
          if (e.shiftKey) {
            // Ctrl+Shift+Z or Cmd+Shift+Z = Redo
            if (canRedo) {
              handleRedo();
            }
          } else {
            // Ctrl+Z or Cmd+Z = Undo
            if (canUndo) {
              handleUndo();
            }
          }
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'y' || e.key === 'Y') {
          // Ctrl+Y or Cmd+Y = Redo (alternative shortcut)
          if (canRedo) {
            handleRedo();
          }
          e.preventDefault();
          e.stopPropagation();
        }
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        deleteSelectedCells();
        e.preventDefault();
        e.stopPropagation();
      }
    };

    document.addEventListener('keydown', handleGlobalKeyDown);
    return () => {
      document.removeEventListener('keydown', handleGlobalKeyDown);
    };
  }, [editingCell, selectedCell, selectedRange, copySelectedCells, cutSelectedCells, deleteSelectedCells, handlePasteFromClipboard, selectAllCells, handleUndo, handleRedo, canUndo, canRedo]);

  // Don't interfere with natural cursor positioning
  // Let the browser handle cursor position naturally for better UX

  // Enhanced input handling for bidirectional sync
  const handleInputBarChange = useCallback((value: string) => {
    setInputValue(value);
    // Update cell value in real-time if we're editing (but don't trigger save yet)
    if (selectedCell && isEditingInInputBar) {
      setCellValueDirect(selectedCell.row, selectedCell.col, value);
    }
  }, [selectedCell, isEditingInInputBar, setCellValueDirect]);

  const handleInputBarKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    e.stopPropagation(); // Prevent spreadsheet key handlers from interfering
    
    if (e.key === 'Enter') {
      e.preventDefault();
      commitInputBarEdit();
      // Move to next row
      if (selectedCell) {
        setSelectedCell({ row: selectedCell.row + 1, col: selectedCell.col });
      }
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitInputBarEdit();
      // Move to next column
      if (selectedCell) {
        setSelectedCell({ 
          row: selectedCell.row, 
          col: e.shiftKey ? Math.max(0, selectedCell.col - 1) : selectedCell.col + 1 
        });
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      // Cancel editing - restore original value
      if (selectedCell) {
        const originalValue = getCellValue(selectedCell.row, selectedCell.col);
        setInputValue(String(originalValue || ''));
      }
      setIsEditingInInputBar(false);
      setEditingCell(null);
    } else if (e.key === 'F2') {
      e.preventDefault();
      // Switch to in-cell editing
      if (selectedCell) {
        setIsEditingInInputBar(false);
        setEditingCell(selectedCell);
      }
    }
  }, [selectedCell, getCellValue, commitInputBarEdit]);

  const handleInputBarFocus = useCallback(() => {
    if (selectedCell) {
      // Clear cell editing mode and start input bar editing
      setEditingCell(null);
      setIsEditingInInputBar(true);
      
      // Load current cell value only if not already editing in input bar
      if (!isEditingInInputBar) {
        const currentValue = getCellValue(selectedCell.row, selectedCell.col);
        setInputValue(String(currentValue || ''));
      }
    }
  }, [selectedCell, getCellValue, isEditingInInputBar]);

  const handleInputBarBlur = useCallback(() => {
    // Don't commit on blur if user is switching to cell editing
    setTimeout(() => {
      if (!editingCell) {
        commitInputBarEdit();
      }
    }, 50);
  }, [editingCell, commitInputBarEdit]);

  // Update input value when selected cell changes (only if not actively editing)
  useEffect(() => {
    if (selectedCell && !isEditingInInputBar && !editingCell) {
      const currentValue = getCellValue(selectedCell.row, selectedCell.col);
      setInputValue(String(currentValue || ''));
    }
  }, [selectedCell, isEditingInInputBar, editingCell, getCellValue]);

  const handleCellClick = useCallback((row: number, col: number, event?: React.MouseEvent) => {
    // Don't interfere with double-click events
    if (event?.detail === 2) {
      return;
    }
    
    // CRITICAL FIX: Immediately clear editing state when switching cells
    if (editingCell || isEditingInInputBar) {
      commitEdit();
      // Force immediate state cleanup to prevent value carryover
      setEditingCell(null);
      setIsEditingInInputBar(false);
    }
    
    // Auto-expand grid when navigating to cells
    handleAutoExpansion(row, col);
    
    // Don't clear copied data here - let user paste multiple times
    
    if (event?.shiftKey && selectedCell) {
      // Extend selection with Shift+click
      setSelectedRange({
        startRow: selectedCell.row,
        startCol: selectedCell.col,
        endRow: row,
        endCol: col
      });
    } else {
      // Single cell selection - load new cell value immediately
      setSelectedCell({ row, col });
      setSelectedRange(null);
      
      // CRITICAL FIX: Immediately load new cell value to prevent carryover
      if (!editingCell && !isEditingInInputBar) {
        const newCellValue = getCellValue(row, col);
        setInputValue(String(newCellValue || ''));
      }
    }
  }, [editingCell, commitEdit, handleAutoExpansion, selectedCell, isEditingInInputBar, getCellValue]);

  const handleMouseDown = useCallback((row: number, col: number, event: React.MouseEvent) => {
    // Don't interfere with double-click events or right-clicks
    if (event.detail === 2 || event.button === 2) {
      return;
    }
    
    // Don't change selection if context menu is visible
    if (contextMenu.visible) {
      setContextMenu({ x: 0, y: 0, visible: false });
      return;
    }
    
    // CRITICAL FIX: Clear editing state immediately when clicking elsewhere
    if (editingCell || isEditingInInputBar) {
      commitEdit();
      setEditingCell(null);
      setIsEditingInInputBar(false);
    }
    
    if (!event.shiftKey) {
      // Don't clear copied data here - let user paste multiple times
      setSelectedCell({ row, col });
      setSelectedRange(null);
      setIsSelecting(true);
      
      // Clear any pending selection updates
      if (selectionUpdateRAFRef.current) {
        cancelAnimationFrame(selectionUpdateRAFRef.current);
        selectionUpdateRAFRef.current = null;
      }
      pendingSelectionRef.current = null;
    }
  }, [editingCell, commitEdit, isEditingInInputBar, contextMenu.visible]);

  // Optimized selection update using RAF for smooth performance
  const updateSelection = useCallback((range: CellRange) => {
    if (selectionUpdateRAFRef.current) {
      cancelAnimationFrame(selectionUpdateRAFRef.current);
    }
    
    pendingSelectionRef.current = range;
    
    selectionUpdateRAFRef.current = requestAnimationFrame(() => {
      if (pendingSelectionRef.current) {
        setSelectedRange(pendingSelectionRef.current);
        pendingSelectionRef.current = null;
      }
      selectionUpdateRAFRef.current = null;
    });
  }, []);

  // Create throttled version of selection update for smooth dragging
  const throttledUpdateSelection = useMemo(
    () => createOptimizedHandler(
      (...args: unknown[]) => {
        const [row, col] = args as [number, number];
        if (selectedCell && !editingCell) {
          updateSelection({
            startRow: selectedCell.row,
            startCol: selectedCell.col,
            endRow: row,
            endCol: col
          });
        }
      },
      { throttle: true, delay: 16, useRAF: true } // 60fps throttling with RAF
    ),
    [selectedCell, editingCell, updateSelection]
  );

  const handleMouseEnter = useCallback((row: number, col: number) => {
    if (fillHandleActive) {
      handleFillHandleMouseMove(row, col);
    } else if (isSelecting && selectedCell && !editingCell) {
      // Use throttled selection update for better performance
      throttledUpdateSelection(row, col);
    }
  }, [fillHandleActive, handleFillHandleMouseMove, isSelecting, selectedCell, editingCell, throttledUpdateSelection]);

  const handleMouseUp = useCallback(() => {
    setIsSelecting(false);
    
    // Ensure any pending selection updates are applied
    if (pendingSelectionRef.current && selectionUpdateRAFRef.current) {
      cancelAnimationFrame(selectionUpdateRAFRef.current);
      setSelectedRange(pendingSelectionRef.current);
      pendingSelectionRef.current = null;
      selectionUpdateRAFRef.current = null;
    }
  }, []);

  // Column resize handlers
  const handleColumnResizeStart = (e: React.MouseEvent, colIndex: number) => {
    e.preventDefault();
    e.stopPropagation();
    setIsResizingColumn(colIndex);
    setResizeStartPos({ x: e.clientX, y: e.clientY });
    setResizeStartSize(columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH);
    document.body.classList.add('resizing-column');
  };

  const handleColumnResizeMove = useCallback((e: MouseEvent) => {
    if (isResizingColumn === null) return;
    
    const deltaX = e.clientX - resizeStartPos.x;
    const newWidth = Math.max(MIN_COLUMN_WIDTH, resizeStartSize + deltaX);
    
    // Update temporary state during drag
    setTempColumnWidths(prev => ({
      ...prev,
      [isResizingColumn]: newWidth
    }));
  }, [isResizingColumn, resizeStartPos.x, resizeStartSize]);

  const handleColumnResizeEnd = useCallback(() => {
    if (isResizingColumn === null) return;
    
    const colIndex = isResizingColumn;
    const oldWidth = columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH;
    const newWidth = tempColumnWidths[colIndex];
    
    // Only create command if width actually changed
    if (newWidth !== undefined && newWidth !== oldWidth) {
      const command: Command = {
        execute: () => {
          setColumnWidths(prev => ({
            ...prev,
            [colIndex]: newWidth
          }));
        },
        undo: () => {
          if (oldWidth === DEFAULT_COLUMN_WIDTH) {
            setColumnWidths(prev => {
              const updated = { ...prev };
              delete updated[colIndex];
              return updated;
            });
          } else {
            setColumnWidths(prev => ({
              ...prev,
              [colIndex]: oldWidth
            }));
          }
        },
        description: `Resize column ${generateColumnLabel(colIndex)}`
      };
      
      executeWithHistory(command);
    }
    
    // Clean up state
    setIsResizingColumn(null);
    setResizeStartPos({ x: 0, y: 0 });
    setResizeStartSize(0);
    setTempColumnWidths(prev => {
      const updated = { ...prev };
      delete updated[colIndex];
      return updated;
    });
    document.body.classList.remove('resizing-column');
  }, [isResizingColumn, columnWidths, tempColumnWidths, executeWithHistory, generateColumnLabel]);

  // Row resize handlers
  const handleRowResizeStart = (e: React.MouseEvent, rowIndex: number) => {
    e.preventDefault();
    e.stopPropagation();
    setIsResizingRow(rowIndex);
    setResizeStartPos({ x: e.clientX, y: e.clientY });
    setResizeStartSize(rowHeights[rowIndex] || DEFAULT_ROW_HEIGHT);
    document.body.classList.add('resizing-row');
  };

  const handleRowResizeMove = useCallback((e: MouseEvent) => {
    if (isResizingRow === null) return;
    
    const deltaY = e.clientY - resizeStartPos.y;
    const newHeight = Math.max(MIN_ROW_HEIGHT, resizeStartSize + deltaY);
    
    // Update temporary state during drag
    setTempRowHeights(prev => ({
      ...prev,
      [isResizingRow]: newHeight
    }));
  }, [isResizingRow, resizeStartPos.y, resizeStartSize]);

  const handleRowResizeEnd = useCallback(() => {
    if (isResizingRow === null) return;
    
    const rowIndex = isResizingRow;
    const oldHeight = rowHeights[rowIndex] || DEFAULT_ROW_HEIGHT;
    const newHeight = tempRowHeights[rowIndex];
    
    // Only create command if height actually changed
    if (newHeight !== undefined && newHeight !== oldHeight) {
      const command: Command = {
        execute: () => {
          setRowHeights(prev => ({
            ...prev,
            [rowIndex]: newHeight
          }));
        },
        undo: () => {
          if (oldHeight === DEFAULT_ROW_HEIGHT) {
            setRowHeights(prev => {
              const updated = { ...prev };
              delete updated[rowIndex];
              return updated;
            });
          } else {
            setRowHeights(prev => ({
              ...prev,
              [rowIndex]: oldHeight
            }));
          }
        },
        description: `Resize row ${rowIndex + 1}`
      };
      
      executeWithHistory(command);
    }
    
    // Clean up state
    setIsResizingRow(null);
    setResizeStartPos({ x: 0, y: 0 });
    setResizeStartSize(0);
    setTempRowHeights(prev => {
      const updated = { ...prev };
      delete updated[rowIndex];
      return updated;
    });
    document.body.classList.remove('resizing-row');
  }, [isResizingRow, rowHeights, tempRowHeights, executeWithHistory]);

  // Global mouse event listeners for resizing
  useEffect(() => {
    if (isResizingColumn !== null) {
      document.addEventListener('mousemove', handleColumnResizeMove);
      document.addEventListener('mouseup', handleColumnResizeEnd);
    }
    
    if (isResizingRow !== null) {
      document.addEventListener('mousemove', handleRowResizeMove);
      document.addEventListener('mouseup', handleRowResizeEnd);
    }
    
    return () => {
      document.removeEventListener('mousemove', handleColumnResizeMove);
      document.removeEventListener('mouseup', handleColumnResizeEnd);
      document.removeEventListener('mousemove', handleRowResizeMove);
      document.removeEventListener('mouseup', handleRowResizeEnd);
    };
  }, [isResizingColumn, isResizingRow, handleColumnResizeMove, handleColumnResizeEnd, handleRowResizeMove, handleRowResizeEnd]);

  // Handle right-click context menu
  const handleCellContextMenu = useCallback((e: React.MouseEvent, row: number, col: number) => {
    e.preventDefault();
    e.stopPropagation();
    
    // IMPORTANT: Don't change selection on right-click - preserve existing selection
    // Get selection info based on current selection (not the clicked cell)
    let selectionInfo = '';
    let selectedValues: string | number | (string | number | null)[][] | null = null;
    
    if (selectedRange) {
      const startCol = Math.min(selectedRange.startCol, selectedRange.endCol);
      const endCol = Math.max(selectedRange.startCol, selectedRange.endCol);
      const startRow = Math.min(selectedRange.startRow, selectedRange.endRow);
      const endRow = Math.max(selectedRange.startRow, selectedRange.endRow);
      
      selectionInfo = `${generateColumnLabel(startCol)}${startRow + 1}:${generateColumnLabel(endCol)}${endRow + 1}`;
      
      // Get values from selected range
      selectedValues = [];
      for (let r = startRow; r <= endRow; r++) {
        const rowValues = [];
        for (let c = startCol; c <= endCol; c++) {
          rowValues.push(getCellValue(r, c));
        }
        selectedValues.push(rowValues);
      }
    } else if (selectedCell) {
      selectionInfo = `${generateColumnLabel(selectedCell.col)}${selectedCell.row + 1}`;
      selectedValues = getCellValue(selectedCell.row, selectedCell.col);
    } else {
      // If no selection, use the clicked cell
      selectionInfo = `${generateColumnLabel(col)}${row + 1}`;
      selectedValues = getCellValue(row, col);
    }
    
    setChatContext({ range: selectionInfo, values: selectedValues });
    
    // Show context menu at mouse position
    setContextMenu({
      x: e.clientX,
      y: e.clientY,
      visible: true
    });
    
    // Prevent the click from propagating and changing selection
    return false;
  }, [selectedRange, selectedCell, generateColumnLabel, getCellValue]);

  // Handle adding to chat
  const handleAddToChat = useCallback(() => {
    setContextMenu({ x: 0, y: 0, visible: false });
    
    // Find the AI Chat component and send the context
    const event = new CustomEvent('addCellContext', { 
      detail: chatContext 
    });
    window.dispatchEvent(event);
    
    // If chat is closed, open it
    const openChatEvent = new CustomEvent('openAIChat');
    window.dispatchEvent(openChatEvent);
  }, [chatContext]);

  // Close context menu on click outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      // Check if click is outside the context menu
      const target = e.target as HTMLElement;
      const isContextMenuClick = target.closest('[data-context-menu]');
      
      if (!isContextMenuClick) {
        setContextMenu({ x: 0, y: 0, visible: false });
      }
    };
    
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setContextMenu({ x: 0, y: 0, visible: false });
      }
    };
    
    if (contextMenu.visible) {
      // Use setTimeout to avoid immediate closing on right-click
      setTimeout(() => {
        document.addEventListener('mousedown', handleClickOutside);
        document.addEventListener('keydown', handleEscape);
      }, 0);
      
      return () => {
        document.removeEventListener('mousedown', handleClickOutside);
        document.removeEventListener('keydown', handleEscape);
      };
    }
  }, [contextMenu.visible]);

  // Handle cell highlighting from AI Chat
  useEffect(() => {
    const handleHighlightCells = (event: CustomEvent) => {
      const { range, isHover } = event.detail;
      
      // Parse the range (e.g., "A1:B3" or "A1")
      const rangeMatch = range.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (!rangeMatch) return;
      
      const startCol = rangeMatch[1].charCodeAt(0) - 65; // Convert A->0, B->1, etc.
      const startRow = parseInt(rangeMatch[2]) - 1; // Convert to 0-based
      
      let endCol = startCol;
      let endRow = startRow;
      
      if (rangeMatch[3] && rangeMatch[4]) {
        endCol = rangeMatch[3].charCodeAt(0) - 65;
        endRow = parseInt(rangeMatch[4]) - 1;
      }
      
      // Set selection and scroll to view
      if (!isHover) {
        setSelectedCell({ row: startRow, col: startCol });
        if (startRow !== endRow || startCol !== endCol) {
          setSelectedRange({
            startRow,
            startCol,
            endRow,
            endCol
          });
        } else {
          setSelectedRange(null);
        }
        
        // Scroll the cell into view
        if (scrollContainerRef.current) {
          const cellElement = document.querySelector(
            `[data-row="${startRow}"][data-col="${startCol}"]`
          ) as HTMLElement;
          
          if (cellElement) {
            cellElement.scrollIntoView({ 
              behavior: 'smooth', 
              block: 'center', 
              inline: 'center' 
            });
          }
        }
      }
    };
    
    const handleClearHighlight = () => {
      // You could add a temporary highlight state here if needed
    };
    
    window.addEventListener('highlightCells', handleHighlightCells as EventListener);
    window.addEventListener('clearHighlight', handleClearHighlight);
    
    return () => {
      window.removeEventListener('highlightCells', handleHighlightCells as EventListener);
      window.removeEventListener('clearHighlight', handleClearHighlight);
    };
  }, []);

  const handleCellDoubleClick = useCallback((row: number, col: number, startWithKey?: string) => {
    const currentValue = getCellValue(row, col);
    
    // Enter editing mode - clear input bar editing
    setIsEditingInInputBar(false);
    setEditingCell({ row, col });
    setInputValue(startWithKey || String(currentValue));
    setSelectedCell({ row, col });
    setSelectedRange(null);
    // Don't clear copied data here - let user paste multiple times
    
    // CRITICAL FIX: Set fresh edit flag for proper text selection
    setIsFreshEdit(true);
    
    // Focus input with retry logic and select all text
    const focusInput = () => {
      if (inputRef.current) {
        inputRef.current.focus();
        // Select all text on fresh edit (double-click or new character)
        if (isFreshEdit) {
          inputRef.current.select();
        }
      } else {
        setTimeout(focusInput, 5);
      }
    };
    
    // Multiple focus attempts for reliability
    setTimeout(focusInput, 0);
    setTimeout(focusInput, 5);
    setTimeout(focusInput, 10);
    setTimeout(focusInput, 20);
  }, [getCellValue, isFreshEdit]);

  const handleArrowKey = useCallback((key: string, row: number, col: number, shiftKey: boolean = false) => {
    let newRow = row;
    let newCol = col;
    
    switch (key) {
      case 'ArrowUp':
        newRow = Math.max(0, row - 1);
        break;
      case 'ArrowDown':
        newRow = row + 1;
        break;
      case 'ArrowLeft':
        newCol = Math.max(0, col - 1);
        break;
      case 'ArrowRight':
        newCol = col + 1;
        break;
    }
    
    // Auto-expand grid when navigating with arrow keys
    handleAutoExpansion(newRow, newCol);
    
    if (shiftKey && selectedCell) {
      // Extend selection with Shift+Arrow
      setSelectedRange({
        startRow: selectedCell.row,
        startCol: selectedCell.col,
        endRow: newRow,
        endCol: newCol
      });
    } else {
      // Don't clear copied data here - let user paste multiple times
      // Move active cell and clear range selection
      setSelectedCell({ row: newRow, col: newCol });
      setSelectedRange(null);
    }
  }, [handleAutoExpansion, selectedCell]);

  const handleKeyDown = useCallback((e: React.KeyboardEvent, row: number, col: number) => {
    // Handle keyboard shortcuts
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'c') {
        copySelectedCells();
        e.preventDefault();
        return;
      } else if (e.key === 'v') {
        handlePasteFromClipboard();
        e.preventDefault();
        return;
      } else if (e.key === 'x') {
        cutSelectedCells();
        e.preventDefault();
        return;
      } else if (e.key === 'a') {
        selectAllCells();
        e.preventDefault();
        return;
      } else if (e.key === 'z') {
        if (e.shiftKey) {
          // Ctrl+Shift+Z or Cmd+Shift+Z = Redo
          if (canRedo) {
            handleRedo();
          }
        } else {
          // Ctrl+Z or Cmd+Z = Undo
          if (canUndo) {
            handleUndo();
          }
        }
        e.preventDefault();
        return;
      } else if (e.key === 'y') {
        // Ctrl+Y or Cmd+Y = Redo (alternative shortcut)
        if (canRedo) {
          handleRedo();
        }
        e.preventDefault();
        return;
      }
    }

    if (e.key === 'Delete' || e.key === 'Backspace') {
      if (!editingCell) {
        deleteSelectedCells();
        e.preventDefault();
        return;
      }
    }

    if (e.key === 'Enter') {
      if (editingCell || isEditingInInputBar) {
        commitEdit();
        // CRITICAL FIX: Clear editing state immediately
        setEditingCell(null);
        setIsEditingInInputBar(false);
        setSelectedCell({ row: row + 1, col });
        // Load new cell value immediately
        const newCellValue = getCellValue(row + 1, col);
        setInputValue(String(newCellValue || ''));
      } else {
        handleCellDoubleClick(row, col);
      }
      e.preventDefault();
    } else if (e.key === 'Tab') {
      if (editingCell || isEditingInInputBar) {
        commitEdit();
        // CRITICAL FIX: Clear editing state immediately
        setEditingCell(null);
        setIsEditingInInputBar(false);
      }
      const newCol = e.shiftKey ? Math.max(0, col - 1) : col + 1;
      setSelectedCell({ row, col: newCol });
      // Load new cell value immediately
      const newCellValue = getCellValue(row, newCol);
      setInputValue(String(newCellValue || ''));
      e.preventDefault();
    } else if (e.key === 'Escape') {
      if (editingCell) {
        setEditingCell(null);
        setIsFreshEdit(false); // Clear fresh edit flag
        setInputValue('');
      } else if (isEditingInInputBar) {
        setIsEditingInInputBar(false);
        setEditingCell(null);
        setIsFreshEdit(false); // Clear fresh edit flag
        // Restore original value
        const originalValue = getCellValue(row, col);
        setInputValue(String(originalValue || ''));
      }
    } else if (e.key === 'F2') {
      e.preventDefault();
      if (!editingCell && !isEditingInInputBar) {
        // Start editing in input bar
        setIsEditingInInputBar(true);
        setEditingCell({ row, col });
        const currentValue = getCellValue(row, col);
        setInputValue(String(currentValue || ''));
      } else if (editingCell && !isEditingInInputBar) {
        // Switch from cell editing to input bar editing
        setEditingCell(null);
        setIsEditingInInputBar(true);
        setEditingCell({ row, col });
      }
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || 
               e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      if (!editingCell && !isEditingInInputBar) {
        handleArrowKey(e.key, row, col, e.shiftKey);
        e.preventDefault();
      }
    } else if (!editingCell && !isEditingInInputBar && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      // Don't clear copied data here - let user paste multiple times
      // Start editing if user types a character
      handleCellDoubleClick(row, col, e.key);
    }
  }, [editingCell, copySelectedCells, handlePasteFromClipboard, cutSelectedCells, selectAllCells, canRedo, handleRedo, canUndo, handleUndo, deleteSelectedCells, handleCellDoubleClick, commitEdit, handleArrowKey, isEditingInInputBar, getCellValue]);



  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newValue = e.target.value;
    setInputValue(newValue);
    
    // Clear fresh edit flag on first keystroke
    if (isFreshEdit) {
      setIsFreshEdit(false);
    }
    
    // LIVE SYNC: Update cell value immediately for live sync with formula bar
    if (editingCell) {
      setCellValueDirect(editingCell.row, editingCell.col, newValue);
    }
    
    // Don't save cursor position - let browser handle it naturally
  };

  const handleInputKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      commitEdit();
      if (editingCell) {
        setSelectedCell({ row: editingCell.row + 1, col: editingCell.col });
      }
    } else if (e.key === 'Tab') {
      commitEdit();
      if (editingCell) {
        setSelectedCell({ 
          row: editingCell.row, 
          col: e.shiftKey ? Math.max(0, editingCell.col - 1) : editingCell.col + 1 
        });
      }
      e.preventDefault();
    } else if (e.key === 'Escape') {
      setEditingCell(null);
      setIsFreshEdit(false); // Clear fresh edit flag
      setInputValue('');
    }
  };

  // Highly optimized memoized cell component with deep comparison for optimal performance
  const CellComponent = memo<{
    row: number;
    col: number;
    cellValue: string | number;
    isSelected: boolean;
    isEditing: boolean;
    isPrimarySelection: boolean;
    isCopied: boolean;
    isCut: boolean;
    isInFillPreview: boolean;
    shouldShowFillHandle: boolean;
    formatting: CellFormatting;
    cellWidth: number;
    cellHeight: number;
    onCellClick: (row: number, col: number, event?: React.MouseEvent) => void;
    onCellDoubleClick: (row: number, col: number) => void;
    onMouseDown: (row: number, col: number, event: React.MouseEvent) => void;
    onMouseEnter: (row: number, col: number) => void;
    onMouseUp: () => void;
    onKeyDown: (e: React.KeyboardEvent, row: number, col: number) => void;
    onFillHandleMouseDown: (e: React.MouseEvent) => void;
    onLinkClick: (e: React.MouseEvent) => void;
    onContextMenu: (e: React.MouseEvent, row: number, col: number) => void;
    hasLink: boolean;
    linkUrl?: string;
  }>(({
    row,
    col,
    cellValue,
    isSelected,
    isEditing,
    isPrimarySelection,
    isCopied,
    isCut,
    isInFillPreview,
    shouldShowFillHandle,
    formatting,
    cellWidth,
    cellHeight,
    onCellClick,
    onCellDoubleClick,
    onMouseDown,
    onMouseEnter,
    onMouseUp,
    onKeyDown,
    onFillHandleMouseDown,
    onLinkClick,
    onContextMenu,
    hasLink,
    linkUrl
  }) => {
    let cellClasses = 'cell';
    if (isSelected) cellClasses += ' selected';
    if (isPrimarySelection) cellClasses += ' primary-selected';
    if (isCut) cellClasses += ' cut';
    else if (isCopied) cellClasses += ' copied';
    if (isInFillPreview) cellClasses += ' fill-preview';
    if (isEditing) cellClasses += ' editing';

    const cellStyle = useMemo<React.CSSProperties>(() => ({
      width: cellWidth,
      minWidth: MIN_COLUMN_WIDTH,
      height: cellHeight,
      minHeight: MIN_ROW_HEIGHT,
      backgroundColor: formatting.backgroundColor,
      textAlign: formatting.textAlign,
      verticalAlign: formatting.verticalAlign,
    }), [cellWidth, cellHeight, formatting.backgroundColor, formatting.textAlign, formatting.verticalAlign]);

    // Build text decoration
    const textDecoration = useMemo(() => {
      if (formatting.underline && formatting.strikethrough) return 'underline line-through';
      if (formatting.underline) return 'underline';
      if (formatting.strikethrough) return 'line-through';
      return 'none';
    }, [formatting.underline, formatting.strikethrough]);

    const contentStyle = useMemo<React.CSSProperties>(() => {
      let justifyContent: React.CSSProperties['justifyContent'];
      switch (formatting.textAlign) {
        case 'left': justifyContent = 'flex-start'; break;
        case 'center': justifyContent = 'center'; break;
        case 'right': justifyContent = 'flex-end'; break;
        case 'justify': justifyContent = 'space-between'; break;
        default: justifyContent = 'flex-start';
      }

      let alignItems: React.CSSProperties['alignItems'];
      switch (formatting.verticalAlign) {
        case 'top': alignItems = 'flex-start'; break;
        case 'middle': alignItems = 'center'; break;
        case 'bottom': alignItems = 'flex-end'; break;
        default: alignItems = 'center';
      }

      return {
        fontFamily: formatting.fontFamily,
        fontSize: formatting.fontSize ? `${formatting.fontSize}px` : undefined,
        fontWeight: formatting.bold ? 'bold' : 'normal',
        fontStyle: formatting.italic ? 'italic' : 'normal',
        textDecoration,
        color: formatting.color,
        justifyContent,
        alignItems,
        cursor: formatting.link ? 'pointer' : 'inherit',
        position: 'relative',
        userSelect: 'none',
        pointerEvents: 'auto'
      } as React.CSSProperties;
    }, [
      formatting.fontFamily,
      formatting.fontSize,
      formatting.bold,
      formatting.italic,
      textDecoration,
      formatting.color,
      formatting.textAlign,
      formatting.verticalAlign,
      formatting.link
    ]);

    return (
      <td
        className={cellClasses}
        style={cellStyle}
        data-row={row}
        data-col={col}
        onClick={(e) => {
          if (e.detail === 1) {
            onCellClick(row, col, e);
          }
        }}
        onDoubleClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          onCellDoubleClick(row, col);
        }}
        onMouseDown={(e) => onMouseDown(row, col, e)}
        onMouseEnter={() => onMouseEnter(row, col)}
        onMouseUp={onMouseUp}
        onKeyDown={(e) => onKeyDown(e, row, col)}
        onContextMenu={(e) => onContextMenu(e, row, col)}
        tabIndex={0}
      >
        {isEditing ? (
          <input
            ref={inputRef}
            type="text"
            value={inputValue}
            onChange={handleInputChange}
            onKeyDown={handleInputKeyDown}
            className="cell-input"
            style={contentStyle}
            autoFocus
            onFocus={(e) => {
              if (isFreshEdit) {
                e.target.select();
                setIsFreshEdit(false);
              }
              e.stopPropagation();
            }}
            onMouseDown={(e) => { e.stopPropagation(); }}
            onClick={(e) => { e.stopPropagation(); }}
            onSelect={(e) => { e.stopPropagation(); }}
          />
        ) : (
          <div 
            className="cell-content" 
            style={contentStyle}
            onDoubleClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              onCellDoubleClick(row, col);
            }}
            onClick={hasLink ? onLinkClick : undefined}
            title={hasLink ? `Double click to open: ${linkUrl}` : undefined}
          >
            {String(cellValue)}
            {shouldShowFillHandle && (
              <div 
                className="fill-handle"
                onMouseDown={onFillHandleMouseDown}
              />
            )}
          </div>
        )}
      </td>
    );
  }, (
    prev,
    next
  ) => {
    return (
      prev.cellValue === next.cellValue &&
      prev.isSelected === next.isSelected &&
      prev.isEditing === next.isEditing &&
      prev.isPrimarySelection === next.isPrimarySelection &&
      prev.isCopied === next.isCopied &&
      prev.isCut === next.isCut &&
      prev.isInFillPreview === next.isInFillPreview &&
      prev.shouldShowFillHandle === next.shouldShowFillHandle &&
      prev.cellWidth === next.cellWidth &&
      prev.cellHeight === next.cellHeight &&
      prev.formatting === next.formatting &&
      prev.hasLink === next.hasLink &&
      prev.linkUrl === next.linkUrl &&
      prev.onCellClick === next.onCellClick &&
      prev.onCellDoubleClick === next.onCellDoubleClick &&
      prev.onMouseDown === next.onMouseDown &&
      prev.onMouseEnter === next.onMouseEnter &&
      prev.onMouseUp === next.onMouseUp &&
      prev.onKeyDown === next.onKeyDown &&
      prev.onFillHandleMouseDown === next.onFillHandleMouseDown &&
      prev.onLinkClick === next.onLinkClick
    );
  });
  
  // (removed) getJustifyContent/getAlignItems in favor of inline computation inside CellComponent

  const renderCell = useCallback((row: number, col: number, cellWidth: number, cellHeight: number) => {
    const rawValue = getCellValue(row, col);
    const cellData = sheetData[row]?.[col];
    const formatting = (cellData?.formatting ? cellData.formatting : EMPTY_FORMATTING) as CellFormatting;
    
    const isSelected = isCellSelected(row, col);
    const isEditing = editingCell?.row === row && editingCell?.col === col;
    const isPrimarySelection = selectedCell?.row === row && selectedCell?.col === col;
    const isCopied = isCellCopied(row, col);
    const isCut = isCellCut(row, col);
    const isInFillPreview = isCellInFillPreview(row, col);

    const shouldShowFillHandle = !isEditing && (isPrimarySelection || (
      selectedRange ?
       row === Math.max(selectedRange.startRow, selectedRange.endRow) && 
        col === Math.max(selectedRange.startCol, selectedRange.endCol)
      : false
    ));

    const displayValue = (isPrimarySelection && isEditingInInputBar)
      ? String(inputValue)
      : formatCellValue(rawValue, formatting);

    const cellLinkClick = (e: React.MouseEvent) => {
      if (formatting.link) {
        e.preventDefault();
        e.stopPropagation();
        window.open(formatting.link, '_blank', 'noopener,noreferrer');
      }
    };

    return (
      <CellComponent
        key={`${row}-${col}`}
        row={row}
        col={col}
        cellValue={displayValue}
        isSelected={isSelected}
        isEditing={isEditing}
        isPrimarySelection={isPrimarySelection}
        isCopied={isCopied}
        isCut={isCut}
        isInFillPreview={isInFillPreview}
        shouldShowFillHandle={shouldShowFillHandle}
        formatting={formatting}
        cellWidth={cellWidth}
        cellHeight={cellHeight}
        onCellClick={handleCellClick}
        onCellDoubleClick={handleCellDoubleClick}
        onMouseDown={handleMouseDown}
        onMouseEnter={handleMouseEnter}
        onMouseUp={handleMouseUp}
        onKeyDown={handleKeyDown}
        onFillHandleMouseDown={handleFillHandleMouseDown}
        onLinkClick={cellLinkClick}
        onContextMenu={handleCellContextMenu}
        hasLink={!!formatting.link}
        linkUrl={formatting.link}
      />
    );
  }, [
    getCellValue,
    sheetData,
    isCellSelected,
    editingCell,
    selectedCell,
    isCellCopied,
    isCellCut,
    isCellInFillPreview,
    selectedRange,
    handleCellClick,
    handleCellDoubleClick,
    handleMouseDown,
    handleMouseEnter,
    handleMouseUp,
    handleKeyDown,
    handleFillHandleMouseDown,
    handleCellContextMenu,
    formatCellValue,
    isEditingInInputBar,
    inputValue,
    CellComponent
  ]);

  // Create spreadsheet operations for AI Chat
  const spreadsheetOperations: SpreadsheetOperations = useMemo(() => ({
    getCellValue: (row: number, col: number) => {
      return getCellValue(row, col);
    },

    setCellValue: (row: number, col: number, value: string | number | null) => {
      setCellValue(row, col, value ?? '');
    },

    getRangeValues: (startRow: number, startCol: number, endRow: number, endCol: number) => {
      const values: (string | number | null)[][] = [];
      for (let row = startRow; row <= endRow; row++) {
        const rowValues: (string | number | null)[] = [];
        for (let col = startCol; col <= endCol; col++) {
          rowValues.push(getCellValue(row, col));
        }
        values.push(rowValues);
      }
      return values;
    },

    // Get cell data including formatting
    getCellData: (row: number, col: number) => {
      return sheetData[row]?.[col] || { value: null, formatting: {} };
    },

    // Get range data including formatting
    getRangeData: (startRow: number, startCol: number, endRow: number, endCol: number) => {
      const data: CellData[][] = [];
      for (let row = startRow; row <= endRow; row++) {
        const rowData: CellData[] = [];
        for (let col = startCol; col <= endCol; col++) {
          rowData.push(sheetData[row]?.[col] || { value: null, formatting: {} });
        }
        data.push(rowData);
      }
      return data;
    },

    // High-quality screenshot of the visible sheet area for AI context
    getSheetScreenshotDataUrl: async () => {
      if (!tableRef.current) return '';
      const tableEl = tableRef.current as HTMLTableElement;
      // Use the scroll container to bound the capture to the visible grid area
      const container = scrollContainerRef.current as HTMLDivElement | null;
      const target = container ?? tableEl;
      const canvas = await html2canvas(target, {
        backgroundColor: '#ffffff',
        useCORS: true,
        scale: Math.min(2, Math.max(1, (window.devicePixelRatio || 1))),
        logging: false,
        allowTaint: true,
        width: target.clientWidth,
        height: target.clientHeight,
        x: target.scrollLeft ?? 0,
        y: target.scrollTop ?? 0,
        windowWidth: document.documentElement.clientWidth,
        windowHeight: document.documentElement.clientHeight,
      });
      return canvas.toDataURL('image/png');
    },

    setRangeValues: (startRow: number, startCol: number, values: (string | number | null)[][]) => {
      values.forEach((rowValues, rowIndex) => {
        rowValues.forEach((value, colIndex) => {
          setCellValue(startRow + rowIndex, startCol + colIndex, value ?? '');
        });
      });
    },

    setCellFormatting: (params: {
      startRow: number;
      startCol: number;
      endRow: number;
      endCol: number;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strikethrough?: boolean;
      fontSize?: number;
      fontFamily?: string;
      color?: string;
      backgroundColor?: string;
      textAlign?: 'left' | 'center' | 'right' | 'justify';
      verticalAlign?: 'top' | 'middle' | 'bottom';
      numberFormat?: string;
      decimalPlaces?: number;
    }) => {
      // Only include defined formatting properties
      const formatting: Partial<CellFormatting> = {};
      if (params.bold !== undefined) formatting.bold = params.bold;
      if (params.italic !== undefined) formatting.italic = params.italic;
      if (params.underline !== undefined) formatting.underline = params.underline;
      if (params.strikethrough !== undefined) formatting.strikethrough = params.strikethrough;
      if (params.fontSize !== undefined) formatting.fontSize = params.fontSize;
      if (params.fontFamily !== undefined) formatting.fontFamily = params.fontFamily;
      if (params.color !== undefined) formatting.color = params.color;
      if (params.backgroundColor !== undefined) formatting.backgroundColor = params.backgroundColor;
      if (params.textAlign !== undefined) formatting.textAlign = params.textAlign;
      if (params.verticalAlign !== undefined) formatting.verticalAlign = params.verticalAlign;
      if (params.numberFormat !== undefined) {
        formatting.numberFormat = params.numberFormat as ('general' | 'number' | 'currency' | 'percentage' | 'date' | 'time' | 'custom');
      }
      if (params.decimalPlaces !== undefined) formatting.decimalPlaces = params.decimalPlaces;
      
      // Directly apply formatting to the cells without changing selection
      setSheetData(prev => {
        const newData = { ...prev };
        for (let row = params.startRow; row <= params.endRow; row++) {
          for (let col = params.startCol; col <= params.endCol; col++) {
            if (!newData[row]) newData[row] = {};
            if (!newData[row][col]) {
              newData[row][col] = { 
                value: prev[row]?.[col]?.value ?? '',
                formatting: {}
              };
            }
            
            const oldFormatting = newData[row][col].formatting || {};
            const mergedFormatting = {
              ...oldFormatting,
              ...formatting
            };
            
            newData[row][col] = {
              ...newData[row][col],
              formatting: mergedFormatting
            };
          }
        }
        return newData;
      });
    },

    setCellFormula: (row: number, col: number, formula: string) => {
      setCellValue(row, col, formula);
    },

    autoFill: (params: {sourceRow: number; sourceCol: number; targetEndRow: number; targetEndCol: number; fillType?: string}) => {
      // Simulate fill operation
      const sourceValue = getCellValue(params.sourceRow, params.sourceCol);
      
      if (params.fillType === 'copy') {
        for (let r = params.sourceRow; r <= params.targetEndRow; r++) {
          for (let c = params.sourceCol; c <= params.targetEndCol; c++) {
            if (r === params.sourceRow && c === params.sourceCol) continue;
            setCellValue(r, c, sourceValue);
          }
        }
      } else if (params.fillType === 'series') {
        const numValue = Number(sourceValue);
        if (!isNaN(numValue)) {
          let currentValue = numValue;
          for (let r = params.sourceRow; r <= params.targetEndRow; r++) {
            for (let c = params.sourceCol; c <= params.targetEndCol; c++) {
              if (r === params.sourceRow && c === params.sourceCol) continue;
              currentValue += 1;
              setCellValue(r, c, currentValue);
            }
          }
        }
      }
    },

    insertRows: (position: number, count: number) => {
      setSheetData(prev => {
        const newData = { ...prev };
        // Shift existing rows down
        const keys = Object.keys(newData).map(Number).sort((a, b) => b - a);
        keys.forEach(row => {
          if (row >= position) {
            newData[row + count] = newData[row];
            delete newData[row];
          }
        });
        return newData;
      });
      // Update visible rows if needed
      if (position < visibleRows) {
        setVisibleRows(prev => prev + count);
      }
    },

    insertColumns: (position: number, count: number) => {
      setSheetData(prev => {
        const newData = { ...prev };
        // Shift existing columns right
        Object.keys(newData).forEach(rowKey => {
          const row = Number(rowKey);
          const rowData = newData[row] || {};
          const newRowData: Record<number, CellData> = {};
          
          Object.keys(rowData).forEach(colKey => {
            const col = Number(colKey);
            if (col >= position) {
              newRowData[col + count] = rowData[col];
            } else {
              newRowData[col] = rowData[col];
            }
          });
          
          newData[row] = newRowData;
        });
        return newData;
      });
      // Update visible columns if needed
      if (position < visibleCols) {
        setVisibleCols(prev => prev + count);
      }
    },

    deleteRows: (startRow: number, count: number) => {
      setSheetData(prev => {
        const newData = { ...prev };
        // Delete specified rows and shift others up
        for (let row = startRow; row < startRow + count; row++) {
          delete newData[row];
        }
        // Shift remaining rows up
        const keys = Object.keys(newData).map(Number).sort((a, b) => a - b);
        const shiftedData: typeof newData = {};
        keys.forEach(row => {
          if (row >= startRow + count) {
            shiftedData[row - count] = newData[row];
          } else if (row < startRow) {
            shiftedData[row] = newData[row];
          }
        });
        return shiftedData;
      });
    },

    deleteColumns: (startCol: number, count: number) => {
      setSheetData(prev => {
        const newData = { ...prev };
        Object.keys(newData).forEach(rowKey => {
          const row = Number(rowKey);
          const rowData = newData[row] || {};
          const newRowData: Record<number, CellData> = {};
          
          Object.keys(rowData).forEach(colKey => {
            const col = Number(colKey);
            if (col < startCol) {
              newRowData[col] = rowData[col];
            } else if (col >= startCol + count) {
              newRowData[col - count] = rowData[col];
            }
            // Skip columns being deleted
          });
          
          newData[row] = newRowData;
        });
        return newData;
      });
    },

    sortRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; sortColumn: number; order: string}) => {
      setSheetData(prev => {
        const newData = { ...prev };
        // Extract the range to sort
        const dataToSort: Array<{row: number; values: Array<string | number | null>}> = [];
        
        for (let row = params.startRow; row <= params.endRow; row++) {
          const rowValues: Array<string | number | null> = [];
          for (let col = params.startCol; col <= params.endCol; col++) {
            rowValues.push(prev[row]?.[col]?.value ?? null);
          }
          dataToSort.push({ row, values: rowValues });
        }
        
        // Sort based on the specified column
        const sortColIndex = params.sortColumn - params.startCol;
        dataToSort.sort((a, b) => {
          const aVal = a.values[sortColIndex];
          const bVal = b.values[sortColIndex];
          
          if (aVal === null || aVal === '') return 1;
          if (bVal === null || bVal === '') return -1;
          
          const comparison = aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
          return params.order === 'asc' ? comparison : -comparison;
        });
        
        // Write sorted data back
        dataToSort.forEach((item, index) => {
          const targetRow = params.startRow + index;
          item.values.forEach((value, colIndex) => {
            const targetCol = params.startCol + colIndex;
            if (!newData[targetRow]) newData[targetRow] = {};
            if (!newData[targetRow][targetCol]) {
              newData[targetRow][targetCol] = { value: '' };
            }
            newData[targetRow][targetCol] = {
              ...newData[targetRow][targetCol],
              value: value ?? ''
            };
          });
        });
        
        return newData;
      });
    },

    filterRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; filterColumn: number; criteria: string; value: string}) => {
      console.log('Filter range:', params);
    },

    findReplace: (params: {findValue: string; replaceValue: string; matchCase?: boolean; matchEntireCell?: boolean; rangeStartRow?: number; rangeStartCol?: number; rangeEndRow?: number; rangeEndCol?: number}) => {
      let count = 0;
      const startRow = params.rangeStartRow ?? 0;
      const endRow = params.rangeEndRow ?? visibleRows - 1;
      const startCol = params.rangeStartCol ?? 0;
      const endCol = params.rangeEndCol ?? visibleCols - 1;

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const cellValue = String(getCellValue(row, col) || '');
          const findValue = String(params.findValue);
          
          let matches = false;
          if (params.matchEntireCell) {
            matches = params.matchCase 
              ? cellValue === findValue 
              : cellValue.toLowerCase() === findValue.toLowerCase();
          } else {
            matches = params.matchCase 
              ? cellValue.includes(findValue)
              : cellValue.toLowerCase().includes(findValue.toLowerCase());
          }

          if (matches) {
            count++;
            if (params.matchEntireCell) {
              setCellValue(row, col, params.replaceValue);
            } else {
              const regex = new RegExp(
                params.findValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                params.matchCase ? 'g' : 'gi'
              );
              setCellValue(row, col, cellValue.replace(regex, params.replaceValue));
            }
          }
        }
      }
      return count;
    },

    createChart: (params: {chartType: string; dataStartRow: number; dataStartCol: number; dataEndRow: number; dataEndCol: number; title?: string; xAxisLabel?: string; yAxisLabel?: string}) => {
      console.log('Create chart:', params);
    },

    clearRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; clearContent?: boolean; clearFormatting?: boolean}) => {
      setSheetData(prev => {
        const newData = { ...prev };
        for (let row = params.startRow; row <= params.endRow; row++) {
          for (let col = params.startCol; col <= params.endCol; col++) {
            if (!newData[row]) continue;
            if (!newData[row][col]) continue;
            
            const currentCell = newData[row][col];
            if (params.clearContent !== false) {
              newData[row][col] = {
                ...currentCell,
                value: ''
              };
            }
            if (params.clearFormatting) {
              newData[row][col] = {
                ...newData[row][col],
                formatting: {}
              };
            }
          }
        }
        return newData;
      });
    },

    mergeRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      console.log('Merge range:', params);
    },

    unmergeRange: (row: number, col: number) => {
      console.log('Unmerge at:', row, col);
    },

    addHyperlink: (params: {row: number; col: number; url: string; displayText?: string}) => {
      setSelectedCell({ row: params.row, col: params.col });
      handleInsertLink(params.url, params.displayText);
    },

    calculateSum: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      let sum = 0;
      for (let row = params.startRow; row <= params.endRow; row++) {
        for (let col = params.startCol; col <= params.endCol; col++) {
          const value = getCellValue(row, col);
          if (typeof value === 'number') {
            sum += value;
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              sum += num;
            }
          }
        }
      }
      return sum;
    },

    calculateAverage: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      let sum = 0;
      let count = 0;
      for (let row = params.startRow; row <= params.endRow; row++) {
        for (let col = params.startCol; col <= params.endCol; col++) {
          const value = getCellValue(row, col);
          if (typeof value === 'number') {
            sum += value;
            count++;
          } else if (typeof value === 'string') {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              sum += num;
              count++;
            }
          }
        }
      }
      return count > 0 ? sum / count : 0;
    },
  }), [getCellValue, setCellValue, handleInsertLink, visibleRows, visibleCols, setSelectedCell, sheetData]);

  return (
    <div className="spreadsheet-container" tabIndex={0} onFocus={() => {}} onClick={() => {}}>
      {/* Excel Toolbar */}
      <ExcelToolbar
        onFormat={formatSelectedCells}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onCopy={() => copySelectedCells(false)}
        onCut={() => cutSelectedCells()}
        onPaste={handlePasteFromClipboard}
        currentFormatting={getCurrentCellFormatting()}
        canUndo={canUndo}
        canRedo={canRedo}
        canPaste={!!copiedData || true}
        selectedCellValue={selectedCell ? getCellValue(selectedCell.row, selectedCell.col) : null}
        onInsertLink={handleInsertLink}
        onOpenAI={() => {
          const evt = new CustomEvent('openAIChat');
          window.dispatchEvent(evt);
        }}
      />

      {/* Formula Input Bar */}
      <FormulaInputBar
        value={inputValue}
        onChange={handleInputBarChange}
        onKeyDown={handleInputBarKeyDown}
        onFocus={handleInputBarFocus}
        onBlur={handleInputBarBlur}
        isActive={isEditingInInputBar}
        selectedCellAddress={selectedCell ? getCellAddress(selectedCell.row, selectedCell.col) : 'A1'}
        className={isEditingInInputBar ? 'active' : ''}
      />

      <style>{`
        .spreadsheet-container {
          width: 100%;
          height: 100%;
          background: white;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          font-size: 11px;
          overflow: auto;
          border: 1px solid #d4d4d4;
          user-select: none;
          display: flex;
          flex-direction: column;
          outline: none;
          /* Ensure crisp rendering of all borders */
          -webkit-font-smoothing: antialiased;
          -moz-osx-font-smoothing: grayscale;
          image-rendering: -webkit-optimize-contrast;
          image-rendering: crisp-edges;
          position: relative;
        }
        
        .spreadsheet-container:focus {
          outline: none;
        }


        
        .spreadsheet-table-container {
          flex: 1;
          overflow: auto;
          transform: scale(${zoomLevel / 100}) translate3d(0, 0, 0);
          transform-origin: top left;
          will-change: transform, scroll-position;
          contain: strict;
          contain-intrinsic-size: 100% 100%;
          backface-visibility: hidden;
          -webkit-backface-visibility: hidden;
          content-visibility: auto;
          /* Force hardware acceleration for crisp borders */
          -webkit-transform: scale(${zoomLevel / 100}) translate3d(0, 0, 0);
          -moz-transform: scale(${zoomLevel / 100}) translate3d(0, 0, 0);
        }
        
        .spreadsheet-table {
          border-collapse: separate;
          border-spacing: 0;
          width: max-content;
          min-width: 100%;
          transform: translateZ(0);
          will-change: contents;
          table-layout: fixed;
          contain: strict;
          contain-intrinsic-size: 100% auto;
          /* Ensure grid lines are always visible */
          -webkit-backface-visibility: hidden;
          backface-visibility: hidden;
          image-rendering: -webkit-optimize-contrast;
          image-rendering: crisp-edges;
        }
        
        .header-row th,
        .row-header {
          background: linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%);
          border-right: 1px solid #a6a6a6;
          border-bottom: 1px solid #a6a6a6;
          padding: 4px 6px;
          font-weight: 400;
          font-size: 11px;
          color: #333;
          text-align: center;
          position: sticky;
          z-index: 100;
          min-width: 80px;
          width: 80px;
          user-select: none;
          box-sizing: border-box;
          contain: layout style;
          transition: background-color 0.15s ease;
        }
        
        .header-row th:first-child {
          border-left: 1px solid #a6a6a6;
        }
        
        .header-row:first-child th {
          border-top: 1px solid #a6a6a6;
        }
        
        .header-row th:hover,
        .row-header:hover {
          background: linear-gradient(180deg, #e8e8e8 0%, #c8c8c8 100%) !important;
        }
        
        .header-row th:active,
        .row-header:active {
          background: linear-gradient(180deg, #d8d8d8 0%, #b8b8b8 100%) !important;
        }
        
        .row-col-header {
          cursor: pointer !important;
        }
        
        .row-col-header:hover {
          background: linear-gradient(180deg, #e8e8e8 0%, #c8c8c8 100%) !important;
        }
        
        .row-col-header:active {
          background: linear-gradient(180deg, #d8d8d8 0%, #b8b8b8 100%) !important;
        }
        
        .header-row th:first-child,
        .row-header {
          min-width: 42px;
          max-width: 42px;
          width: 42px;
          position: sticky;
          left: 0;
          z-index: 101;
        }
        
        .header-row th {
          top: 0;
        }
        
        .header-row th:first-child {
          z-index: 102;
        }
        
        /* Highlighted header styles */
        .header-row th.highlighted {
          background: linear-gradient(180deg, #217346 0%, #1a5b36 100%) !important;
          color: white !important;
          font-weight: 500;
        }
        
        .row-header.highlighted {
          background: linear-gradient(180deg, #217346 0%, #1a5b36 100%) !important;
          color: white !important;
          font-weight: 500;
        }
        
        .cell {
          border-right: 1px solid #d4d4d4;
          border-bottom: 1px solid #d4d4d4;
          padding: 0;
          height: 20px;
          min-width: 80px;
          will-change: background-color, border-color;
          transition: background-color 0.1s ease, border-color 0.1s ease;
          contain: layout style paint;
          width: 80px;
          background: white;
          position: relative;
          cursor: cell;
          outline: none;
          transform: translateZ(0);
          will-change: background-color;
          contain: layout style;
          user-select: none;
          pointer-events: auto;
          box-sizing: border-box;
          /* Ensure borders are crisp */
          -webkit-backface-visibility: hidden;
          backface-visibility: hidden;
        }
        
        /* First column gets left border */
        td.cell:first-child,
        .cell:first-child {
          border-left: 1px solid #d4d4d4;
        }
        
        /* First row cells get top border */
        tbody tr:first-child .cell {
          border-top: 1px solid #d4d4d4;
        }
        
        .cell.editing {
          cursor: text;
        }
        
        .cell:first-child {
          position: sticky;
          left: 0;
          z-index: 10;
        }
        
        .cell.selected {
          box-shadow: inset 0 0 0 1px rgba(33, 115, 70, 0.6) !important;
          position: relative;
          z-index: 40;
          transition: none !important; /* Disable transition for instant selection feedback */
        }
        
        .cell.selected::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(33, 115, 70, 0.08);
          pointer-events: none;
          z-index: 1;
          will-change: opacity;
        }
        
        .cell.primary-selected {
          box-shadow: inset 0 0 0 2px #217346 !important;
          z-index: 50;
          position: relative;
        }
        
        .cell.primary-selected::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(33, 115, 70, 0.12);
          pointer-events: none;
          z-index: 1;
        }
        
        .cell.copied {
          z-index: 30;
          position: relative;
          animation: copied-border-pulse 2s ease-in-out infinite;
          box-shadow: inset 0 0 0 2px #0078d4;
        }
        
        .cell.copied::before {
          content: '';
          position: absolute;
          top: -1px;
          left: -1px;
          right: -1px;
          bottom: -1px;
          border: 2px dashed #0078d4;
          pointer-events: none;
          z-index: 1;
        }
        
        .cell.copied::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0, 120, 212, 0.15);
          pointer-events: none;
          z-index: 1;
          animation: copied-pulse-overlay 2s ease-in-out infinite;
        }
        

        
        @keyframes copied-pulse-overlay {
          0% { 
            background: rgba(0, 120, 212, 0.15);
          }
          50% { 
            background: rgba(0, 120, 212, 0.05);
          }
          100% { 
            background: rgba(0, 120, 212, 0.15);
          }
        }
        
        /* Cut animation removed - no longer needed */
        
        @keyframes copied-border-pulse {
          0% { 
            border-color: #0078d4;
          }
          50% { 
            border-color: #4a9eff;
          }
          100% { 
            border-color: #0078d4;
          }
        }
        
        /* Cut border animation removed - no longer needed */
        
        .cell:hover:not(.selected):not(.primary-selected):not(.copied):not(.cut) {
          position: relative;
        }
        
        .cell:hover:not(.selected):not(.primary-selected):not(.copied):not(.cut)::before {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0, 0, 0, 0.04);
          pointer-events: none;
          z-index: 1;
        }
        
        /* Ensure proper cleanup when copied/cut class is removed */
        .cell:not(.copied):not(.cut) {
          animation: none;
        }
        
        .cell-content {
          width: 100%;
          height: 100%;
          padding: 2px 4px;
          display: flex;
          box-sizing: border-box;
          background: transparent;
          border: none;
          outline: none;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          color: black;
          position: relative;
          z-index: 2;
          transform: translateZ(0);
          contain: layout style paint;
          cursor: inherit;
          user-select: none;
          pointer-events: auto;
        }
        
        .cell-content:hover {
          background: rgba(0, 0, 0, 0.02);
        }
        
        .cell.editing .cell-content {
          cursor: text;
        }
        
        .cell-input {
          width: 100%;
          height: 100%;
          padding: 2px 4px;
          border: 2px solid #217346;
          outline: none;
          background: white;
          font-family: inherit;
          font-size: inherit;
          box-sizing: border-box;
          position: relative;
          z-index: 200;
          color: black;
          user-select: text;
          border-radius: 0;
          cursor: text;
        }
        
        .cell-input:focus {
          border: 2px solid #217346;
          box-shadow: 0 0 0 1px rgba(33, 115, 70, 0.3);
          cursor: text;
        }
        
        .cell.fill-preview {
          background: rgba(33, 115, 70, 0.2) !important;
          box-shadow: inset 0 0 0 1px #217346;
        }
        
        .cell.fill-preview::before {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          border: 1px dashed #217346;
          pointer-events: none;
          z-index: 1;
        }

        
        .fill-handle {
          position: absolute;
          bottom: -2px;
          right: -2px;
          width: 6px;
          height: 6px;
          background: #217346;
          border: 1px solid white;
          cursor: crosshair;
          z-index: 100;
          box-sizing: border-box;
          transform: translateZ(0);
          will-change: transform, background-color;
          transition: transform 0.1s ease-out;
        }
        
        .fill-handle:hover {
          background: #1a5a37;
          transform: scale(1.2) translateZ(0);
        }
        
        .cell-content {
          position: relative;
        }
        
        /* Ensure fill handle is only visible on selected cells */
        .cell:not(.selected):not(.primary-selected) .fill-handle {
          display: none;
        }
        
        /* Column resize handle styles */
        .column-resize-handle {
          position: absolute;
          top: 0;
          right: -2px;
          width: 4px;
          height: 100%;
          cursor: col-resize;
          background: transparent;
          z-index: 200;
          opacity: 0;
          transition: opacity 0.2s ease, background-color 0.2s ease;
        }
        
        .column-header:hover .column-resize-handle {
          opacity: 1;
        }
        
        .column-resize-handle:hover {
          background: #217346;
          opacity: 1;
        }
        
        /* Row resize handle styles */
        .row-resize-handle {
          position: absolute;
          bottom: -2px;
          left: 0;
          width: 100%;
          height: 4px;
          cursor: row-resize;
          background: transparent;
          z-index: 200;
          opacity: 0;
          transition: opacity 0.2s ease, background-color 0.2s ease;
        }
        
        .row-header:hover .row-resize-handle {
          opacity: 1;
        }
        
        .row-resize-handle:hover {
          background: #217346;
          opacity: 1;
        }
        
        /* Header styles for resizing */
        .column-header {
          position: relative;
          min-width: ${MIN_COLUMN_WIDTH}px;
        }
        
        .row-header {
          position: relative;
          min-height: ${MIN_ROW_HEIGHT}px;
          background: #f0f0f0;
          border-right: 1px solid #d4d4d4;
          border-bottom: 1px solid #d4d4d4;
          text-align: center;
          font-weight: normal;
          color: #666;
          width: 40px;
          min-width: 40px;
          padding: 2px 4px;
        }
        
        .row-col-header {
          background: #f0f0f0;
          border-right: 1px solid #d4d4d4;
          border-bottom: 1px solid #d4d4d4;
          width: 40px;
          min-width: 40px;
        }
        
        /* Active resizing cursor styles */
        body.resizing-column {
          cursor: col-resize !important;
        }
        
        body.resizing-row {
          cursor: row-resize !important;
        }
        
        /* Prevent text selection during resize */
        .spreadsheet-table-container.resizing {
          user-select: none;
          -webkit-user-select: none;
          -moz-user-select: none;
          -ms-user-select: none;
        }
        
        /* Zoom Control Bar */
        .zoom-control-bar {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 4px 8px;
          background: #f8f9fa;
          border-top: 1px solid #d4d4d4;
          font-size: 11px;
          height: 24px;
          flex-shrink: 0;
        }
        
        .zoom-btn {
          width: 20px;
          height: 16px;
          border: 1px solid #ccc;
          background: white;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 12px;
          font-weight: bold;
          color: #333;
        }
        
        .zoom-btn:hover:not(:disabled) {
          background: #e9ecef;
          border-color: #217346;
        }
        
        .zoom-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }
        
        .zoom-display {
          margin: 0 8px;
          padding: 2px 6px;
          cursor: pointer;
          border: 1px solid transparent;
          border-radius: 2px;
          min-width: 40px;
          text-align: center;
        }
        
        .zoom-display:hover {
          background: #e9ecef;
          border-color: #ccc;
        }
        
        .grid-info {
          color: #666;
          font-size: 10px;
          margin-left: auto;
        }
        
        /* Clipboard notification styles */
        .clipboard-notification {
          position: fixed;
          top: 80px;
          right: 20px;
          background: #28a745;
          color: white;
          padding: 8px 16px;
          border-radius: 4px;
          font-size: 12px;
          font-weight: 500;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          z-index: 1000;
          animation: slideInFade 0.3s ease-out;
          max-width: 300px;
        }
        
        .clipboard-notification.error {
          background: #dc3545;
        }
        
        .clipboard-notification.info {
          background: #17a2b8;
        }
        
        @keyframes slideInFade {
          from {
            opacity: 0;
            transform: translateX(100%);
          }
          to {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        /* Enhanced copy/cut cell animations - these override the previous definitions */
        


        /* Ensure grid lines are always visible */
        .cell, .cell.selected, .cell.primary-selected, .cell.copied, .cell.cut, .cell.fill-preview {
          border-collapse: collapse;
        }
        
        /* Ensure proper grid line inheritance with containment */
        .spreadsheet-table td {
          border: 1px solid #d4d4d4 !important;
          border-collapse: collapse;
          box-sizing: border-box;
        }
        
        .spreadsheet-table th {
          border: 1px solid #a6a6a6 !important;
          border-collapse: collapse;
          box-sizing: border-box;
        }
        
        /* Table row containment for performance */
        .spreadsheet-table tbody tr {
          contain: layout style;
          content-visibility: auto;
          contain-intrinsic-size: auto 20px;
        }
        
        .spreadsheet-table thead tr {
          contain: layout style;
          contain-intrinsic-size: auto 20px;
        }
        
        /* Spacer element styles for optimal performance */
        .cell-spacer {
          pointer-events: none;
          user-select: none;
          cursor: default;
          contain: layout style;
          transform: translateZ(0);
          will-change: auto;
          border: 1px solid #d4d4d4 !important;
          box-sizing: border-box;
        }
        
        .column-header-spacer {
          pointer-events: none;
          user-select: none;
          cursor: default;
          contain: layout style;
          transform: translateZ(0);
          will-change: auto;
          border: 1px solid #a6a6a6 !important;
          box-sizing: border-box;
        }

        /* Hyperlink hover tooltip styles */
        .cell-content[title]:hover {
          position: relative;
        }

        .cell-content[title]:hover::after {
          content: attr(title);
          position: absolute;
          bottom: 100%;
          left: 50%;
          transform: translateX(-50%);
          background: rgba(0, 0, 0, 0.9);
          color: white;
          padding: 6px 10px;
          border-radius: 4px;
          font-size: 11px;
          white-space: nowrap;
          z-index: 1000;
          pointer-events: none;
          opacity: 0;
          animation: fadeInTooltip 0.2s ease-in-out 0.5s forwards;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        }

        @keyframes fadeInTooltip {
          from { 
            opacity: 0;
            transform: translateX(-50%) translateY(5px);
          }
          to { 
            opacity: 1;
            transform: translateX(-50%) translateY(0);
          }
        }
        /* Sidebar toggle button (Top-right) */
        .sidebar-fab {
          position: absolute;
          top: 8px;
          right: ${isActionSidebarOpen ? '252px' : '12px'};
          height: 32px;
          min-width: 32px;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 0 10px;
          border: 1px solid #ced4da;
          background: white;
          border-radius: 16px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.08);
          cursor: pointer;
          color: #495057;
          z-index: 210;
        }
        .sidebar-fab:hover { background: #f8f9fa; }
      `}</style>
      
      <button className="sidebar-fab" onClick={() => setIsActionSidebarOpen(v => !v)} title={isActionSidebarOpen ? 'Hide actions' : 'Show actions'}>
        <span style={{ fontSize: 12, fontWeight: 700 }}>Actions</span>
      </button>
      <div className="spreadsheet-table-container" ref={scrollContainerRef}>
        <table className="spreadsheet-table" ref={tableRef}>
          
          {/* Header row */}
          <thead>
            <tr className="header-row">
              <th 
                className="row-col-header"
                onClick={handleSelectAllClick}
                style={{ cursor: 'pointer' }}
                title="Select all cells"
              ></th>
              {(() => {
                const viewport = getViewport();
                const headers = [];
                const highlightedColumns = getHighlightedColumns;
                
                // Use buffered rendering for consistency
                const renderBuffer = shouldUseVirtualScrolling ? RENDER_BUFFER : 0;
                const renderStartCol = Math.max(0, viewport.startCol - renderBuffer);
                const renderEndCol = Math.min(visibleCols, viewport.endCol + renderBuffer);
                
                // Add spacer header for left offset if needed
                if (renderStartCol > 0) {
                  const spacerWidth = renderStartCol * DEFAULT_COLUMN_WIDTH;
                  headers.push(
                    <th 
                      key="spacer-left" 
                      className="column-header-spacer"
                      style={{ 
                        width: `${spacerWidth}px`,
                        border: '1px solid #a6a6a6',
                        background: 'linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%)',
                        padding: '4px 6px'
                      }}
                    ></th>
                  );
                }
                
                // Render buffered column headers
                for (let i = renderStartCol; i < renderEndCol; i++) {
                  const isHighlighted = highlightedColumns.has(i);
                  headers.push(
                    <th 
                      key={i} 
                      className={`column-header ${isHighlighted ? 'highlighted' : ''}`}
                      style={{ 
                        width: getCurrentColumnWidth(i),
                        minWidth: MIN_COLUMN_WIDTH,
                        position: 'relative',
                        cursor: 'pointer'
                      }}
                      onClick={(e) => handleColumnHeaderClick(i, e)}
                      title={`Select column ${generateColumnLabel(i)}`}
                    >
                      {generateColumnLabel(i)}
                      <div 
                        className="column-resize-handle"
                        onMouseDown={(e) => handleColumnResizeStart(e, i)}
                      />
                    </th>
                  );
                }
                
                // Add spacer header for right offset if needed
                if (renderEndCol < visibleCols) {
                  const remainingCols = visibleCols - renderEndCol;
                  const spacerWidth = remainingCols * DEFAULT_COLUMN_WIDTH;
                  headers.push(
                    <th 
                      key="spacer-right" 
                      className="column-header-spacer"
                      style={{ 
                        width: `${spacerWidth}px`,
                        border: '1px solid #a6a6a6',
                        background: 'linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%)',
                        padding: '4px 6px'
                      }}
                    ></th>
                  );
                }
                
                return headers;
              })()}
            </tr>
          </thead>
          
          {/* Data rows with virtual scrolling */}
          <tbody>
            {(() => {
              const viewport = getViewport();
              const rows = [];
              const highlightedRows = getHighlightedRows;
              
              // Use buffered rendering for better performance
              const renderBuffer = shouldUseVirtualScrolling ? RENDER_BUFFER : 0;
              const renderStartRow = Math.max(0, viewport.startRow - renderBuffer);
              const renderEndRow = Math.min(visibleRows, viewport.endRow + renderBuffer);
              const renderStartCol = Math.max(0, viewport.startCol - renderBuffer);
              const renderEndCol = Math.min(visibleCols, viewport.endCol + renderBuffer);
              
              // Add spacer row for top offset if needed
              if (renderStartRow > 0) {
                const spacerHeight = renderStartRow * DEFAULT_ROW_HEIGHT;
                rows.push(
                  <tr key="spacer-top" style={{ height: `${spacerHeight}px` }}>
                    <td className="row-header" style={{ 
                      border: '1px solid #a6a6a6',
                      background: 'linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%)',
                      padding: '4px 6px'
                    }}></td>
                    {Array.from({ length: visibleCols }, (_, colIndex) => (
                      <td 
                        key={`spacer-top-${colIndex}`}
                        className="cell-spacer"
                        style={{ 
                          width: getCurrentColumnWidth(colIndex),
                          borderRight: '1px solid #d4d4d4',
                        borderBottom: '1px solid #d4d4d4',
                          background: 'white'
                        }}
                      ></td>
                    ))}
                  </tr>
                );
              }
              
              // Render buffered rows
              for (let row = renderStartRow; row < renderEndRow; row++) {
                const isRowHighlighted = highlightedRows.has(row);
                const rowCells = [];
                
                // Row header
                rowCells.push(
                  <td 
                    key="row-header"
                    className={`row-header ${isRowHighlighted ? 'highlighted' : ''}`}
                    style={{ 
                      position: 'relative',
                      height: getCurrentRowHeight(row),
                      minHeight: MIN_ROW_HEIGHT,
                      cursor: 'pointer'
                    }}
                    onClick={(e) => handleRowHeaderClick(row, e)}
                    title={`Select row ${row + 1}`}
                  >
                    {row + 1}
                    <div 
                      className="row-resize-handle"
                      onMouseDown={(e) => handleRowResizeStart(e, row)}
                    />
                  </td>
                );
                
                // Add spacer cell for left offset if needed
                if (renderStartCol > 0) {
                  const spacerWidth = renderStartCol * DEFAULT_COLUMN_WIDTH;
                  rowCells.push(
                    <td 
                      key="cell-spacer-left"
                      className="cell-spacer"
                      style={{ 
                        width: `${spacerWidth}px`,
                        height: getCurrentRowHeight(row),
                        borderRight: '1px solid #d4d4d4',
                        borderBottom: '1px solid #d4d4d4',
                        borderLeft: '1px solid #d4d4d4',
                        borderTop: row === 0 ? '1px solid #d4d4d4' : 'none',
                        background: 'white'
                      }}
                    ></td>
                  );
                }
                
                // Render buffered cells in this row
                for (let col = renderStartCol; col < renderEndCol; col++) {
                  const cell = renderCell(
                    row,
                    col,
                    getCurrentColumnWidth(col) as number,
                    getCurrentRowHeight(row) as number
                  );
                  rowCells.push(React.cloneElement(cell, { key: `cell-${col}` }));
                }
                
                // Add spacer cell for right offset if needed
                if (renderEndCol < visibleCols) {
                  const remainingCols = visibleCols - renderEndCol;
                  const spacerWidth = remainingCols * DEFAULT_COLUMN_WIDTH;
                  rowCells.push(
                    <td 
                      key="cell-spacer-right"
                      className="cell-spacer"
                      style={{ 
                        width: `${spacerWidth}px`,
                        height: getCurrentRowHeight(row),
                        borderRight: '1px solid #d4d4d4',
                        borderBottom: '1px solid #d4d4d4',
                        borderTop: row === 0 ? '1px solid #d4d4d4' : 'none',
                        background: 'white'
                      }}
                    ></td>
                  );
                }
                
                rows.push(
                  <tr 
                    key={row}
                    style={{
                      height: getCurrentRowHeight(row),
                      minHeight: MIN_ROW_HEIGHT
                    }}
                  >
                    {rowCells}
                  </tr>
                );
              }
              
              // Add spacer row for bottom offset if needed
              if (renderEndRow < visibleRows) {
                const remainingRows = visibleRows - renderEndRow;
                const spacerHeight = remainingRows * DEFAULT_ROW_HEIGHT;
                rows.push(
                  <tr key="spacer-bottom" style={{ height: `${spacerHeight}px` }}>
                    <td className="row-header" style={{ 
                      borderRight: '1px solid #a6a6a6',
                      borderBottom: '1px solid #a6a6a6',
                      borderLeft: '1px solid #a6a6a6',
                      borderTop: '1px solid #a6a6a6',
                      background: 'linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%)',
                      padding: '4px 6px'
                    }}></td>
                    {Array.from({ length: visibleCols }, (_, colIndex) => (
                      <td 
                        key={`spacer-bottom-${colIndex}`}
                        className="cell-spacer"
                        style={{ 
                          width: getCurrentColumnWidth(colIndex),
                          borderRight: '1px solid #d4d4d4',
                          borderBottom: '1px solid #d4d4d4',
                          borderLeft: colIndex === 0 ? '1px solid #d4d4d4' : 'none',
                          borderTop: '1px solid #d4d4d4',
                          background: 'white'
                        }}
                      ></td>
                    ))}
                  </tr>
                );
              }
              
              return rows;
            })()}
          </tbody>
        </table>
      </div>
      {/* Version Preview Modal */}
      {previewModal.visible && previewModal.versionId && (
        <div
          onClick={() => setPreviewModal({ visible: false })}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)',
            backdropFilter: 'blur(4px)', WebkitBackdropFilter: 'blur(4px)',
            zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center',
            animation: 'overlayFadeIn 0.2s ease-out'
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              width: 'min(90vw, 900px)', maxHeight: '80vh', overflow: 'hidden',
              background: '#ffffff', borderRadius: 12, boxShadow: '0 20px 60px rgba(0,0,0,0.12), 0 8px 25px rgba(0,0,0,0.08)',
              border: '1px solid rgba(0,0,0,0.05)', padding: 16, display: 'flex', flexDirection: 'column', gap: 10,
              animation: 'modalSlideIn 0.25s ease-out'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: '#2d3748' }}>Version Preview</div>
                {(() => {
                  const info = versionHistory.find(v => v.id === previewModal.versionId);
                  return info ? (
                    <div style={{ fontSize: 12, color: '#6c757d' }}>
                      {new Date(info.timestamp).toLocaleString()} — {info.description}
                    </div>
                  ) : null;
                })()}
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  onClick={() => setPreviewModal({ visible: false })}
                  style={{ border: '1px solid #ced4da', background: '#ffffff', padding: '6px 10px', borderRadius: 6, fontSize: 12, cursor: 'pointer', color: '#333' }}
                >
                  Close
                </button>
                <button
                  onClick={() => {
                    const vid = previewModal.versionId!;
                    const snap = versionSnapshotsRef.current.get(vid);
                    if (snap) {
                      setSheetData(deepClone(snap.data.sheetData));
                      setColumnWidths(deepClone(snap.data.columnWidths));
                      setRowHeights(deepClone(snap.data.rowHeights));
                      snapshotRequestRef.current = { requested: true, description: `Restore: ${snap.description}` };
                    }
                    setPreviewModal({ visible: false });
                  }}
                  style={{ border: '1px solid #0d6efd', background: '#0d6efd', color: '#ffffff', padding: '6px 10px', borderRadius: 6, fontSize: 12, cursor: 'pointer' }}
                >
                  Restore this version
                </button>
              </div>
            </div>
            <div>
              {previewModal.versionId && renderSnapshotPreview(previewModal.versionId)}
            </div>
          </div>
        </div>
      )}
      {/* Action Sidebar */}
      <ActionSidebar
        isOpen={isActionSidebarOpen}
        onSave={handleSave}
        onSaveAsExcel={handleSaveAsExcel}
        onExportPDF={handleExportPDF}
        onExportPNG={handleExportPNG}
        onImportExcelFile={handleImportExcelFile}
        versionHistory={versionHistory}
        onPreviewVersion={(id) => setPreviewModal({ visible: true, versionId: id })}
        onRestoreVersion={(id) => {
          const snap = versionSnapshotsRef.current.get(id);
          if (!snap) return;
          setSheetData(deepClone(snap.data.sheetData));
          setColumnWidths(deepClone(snap.data.columnWidths));
          setRowHeights(deepClone(snap.data.rowHeights));
          snapshotRequestRef.current = { requested: true, description: `Restore: ${snap.description}` };
        }}
      />
      
      {/* Clipboard Notification */}
      {clipboardNotification && (
        <div className={`clipboard-notification ${clipboardNotification.includes('error') || clipboardNotification.includes('No') ? 'error' : clipboardNotification.includes('Unable') || clipboardNotification.includes('unavailable') ? 'info' : ''}`}>
          {clipboardNotification}
        </div>
      )}

      {/* Context Menu */}
      {contextMenu.visible && (
        <div
          data-context-menu="true"
          style={{
            position: 'fixed',
            top: Math.min(contextMenu.y, window.innerHeight - 300), // Prevent menu from going off-screen
            left: Math.min(contextMenu.x, window.innerWidth - 180),
            zIndex: 10000,
            backgroundColor: 'white',
            border: '1px solid #e5e7eb',
            borderRadius: '6px',
            boxShadow: '0 2px 8px rgba(0, 0, 0, 0.12), 0 0 1px rgba(0, 0, 0, 0.08)',
            padding: '3px',
            minWidth: '160px',
            animation: 'fadeIn 0.12s ease-out',
            fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
          }}
          onMouseDown={(e) => e.stopPropagation()}
          onClick={(e) => e.stopPropagation()}
        >
          {/* AI Chat Option */}
          <button
            onClick={handleAddToChat}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              width: '100%',
              padding: '6px 10px',
              border: 'none',
              backgroundColor: 'transparent',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '12px',
              fontWeight: '400',
              color: '#374151',
              transition: 'background-color 0.1s ease',
              textAlign: 'left'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f3f4f6';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#6366f1" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
            </svg>
            <span>Add to chat</span>
          </button>

          <div style={{ height: '1px', backgroundColor: '#e5e7eb', margin: '2px 6px' }} />

          {/* Delete Option */}
          <button
            onClick={() => {
              if (selectedCell) {
                setSheetData(prev => ({
                  ...prev,
                  [`${selectedCell.row}-${selectedCell.col}`]: { value: '' }
                }));
              } else if (selectedRange) {
                const updates: { [key: string]: CellData } = {};
                for (let row = selectedRange.startRow; row <= selectedRange.endRow; row++) {
                  for (let col = selectedRange.startCol; col <= selectedRange.endCol; col++) {
                    updates[`${row}-${col}`] = { value: '' };
                  }
                }
                setSheetData(prev => ({ ...prev, ...updates }));
              }
              setContextMenu({ visible: false, x: 0, y: 0 });
            }}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              width: '100%',
              padding: '6px 10px',
              border: 'none',
              backgroundColor: 'transparent',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '12px',
              fontWeight: '400',
              color: '#374151',
              transition: 'background-color 0.1s ease',
              textAlign: 'left'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#fef2f2';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
            <span>Clear Contents</span>
          </button>

          <div style={{ height: '1px', backgroundColor: '#e5e7eb', margin: '2px 6px' }} />

          {/* Insert Row Above Option */}
          <button
            onClick={() => {
              if (selectedCell) {
                spreadsheetOperations.insertRows(selectedCell.row, 1);
              }
              setContextMenu({ visible: false, x: 0, y: 0 });
            }}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              width: '100%',
              padding: '6px 10px',
              border: 'none',
              backgroundColor: 'transparent',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '12px',
              fontWeight: '400',
              color: '#374151',
              transition: 'background-color 0.1s ease',
              textAlign: 'left'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f3f4f6';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"></line>
              <line x1="5" y1="12" x2="19" y2="12"></line>
              <polyline points="12 5 8 9 16 9"></polyline>
            </svg>
            <span>Insert Row Above</span>
          </button>


          {/* Insert Column Left Option */}
          <button
            onClick={() => {
              if (selectedCell) {
                spreadsheetOperations.insertColumns(selectedCell.col, 1);
              }
              setContextMenu({ visible: false, x: 0, y: 0 });
            }}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              width: '100%',
              padding: '6px 10px',
              border: 'none',
              backgroundColor: 'transparent',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '12px',
              fontWeight: '400',
              color: '#374151',
              transition: 'background-color 0.1s ease',
              textAlign: 'left'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f3f4f6';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"></line>
              <line x1="5" y1="12" x2="19" y2="12"></line>
              <polyline points="5 12 9 8 9 16"></polyline>
            </svg>
            <span>Insert Column Left</span>
          </button>

        </div>
      )}

      {/* AI Chat Assistant - hide FAB when actions sidebar is open */}
      <AIChat spreadsheetOperations={spreadsheetOperations} hideFab={isActionSidebarOpen} />
    </div>
  );
});

// Set display name for better debugging
ExcelSpreadsheet.displayName = 'ExcelSpreadsheet';

export default ExcelSpreadsheet;
export type { CellFormatting };