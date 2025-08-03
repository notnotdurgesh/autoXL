import React, { useState, useCallback, useRef, useEffect, useMemo, memo } from 'react';
import { useUndoRedo, type Command } from '../hooks/useUndoRedo';
import ExcelToolbar from './ExcelToolbar';
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
  const [isSelecting, setIsSelecting] = useState<boolean>(false);
  const [copiedData, setCopiedData] = useState<{data: CellData[][], range: CellRange, isCut?: boolean} | null>(null);
  
  // Enhanced state for clipboard operations
  const [clipboardNotification, setClipboardNotification] = useState<string | null>(null);
  
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
  const cursorPositionRef = useRef<number>(0);
  
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
  const { executeCommand, undo, redo, canUndo, canRedo } = useUndoRedo();
  
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

  // Virtual scrolling calculations - PERFORMANCE OPTIMIZATION (memoized)
  const calculateViewport = useMemo(() => {
    if (!shouldUseVirtualScrolling) {
      // If grid is small enough, render everything
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

    // Calculate which rows are visible
    const startRow = Math.max(0, Math.floor(scrollTop / (DEFAULT_ROW_HEIGHT * zoomFactor)) - RENDER_BUFFER);
    const endRow = Math.min(visibleRows, startRow + Math.ceil(clientHeight / (DEFAULT_ROW_HEIGHT * zoomFactor)) + RENDER_BUFFER * 2);

    // Calculate which columns are visible  
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

  // Update viewport when scrolling (optimized with RAF)
  const updateViewport = useCallback(() => {
    const viewport = calculateViewport;
    
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
  }, [calculateViewport, viewportStartRow, viewportEndRow, viewportStartCol, viewportEndCol]);


  // Initialize with sample data using performance optimization - ONLY ONCE
  useEffect(() => {
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

  // Smart scroll handling with virtual scrolling and performance optimization (optimized)
  useEffect(() => {
    const handleScroll = (e: Event) => {
      const target = e.target as HTMLElement;
      if (!target.closest('.spreadsheet-table-container')) return;
      
      const container = target.closest('.spreadsheet-table-container') as HTMLElement;
      if (!container) return;
      
      // Update virtual viewport for performance
      if (shouldUseVirtualScrolling) {
        updateViewport();
      }
      
      const { scrollTop, scrollHeight, clientHeight, scrollLeft, scrollWidth, clientWidth } = container;
      
      // Calculate scroll percentages
      const verticalScrollPercent = (scrollTop + clientHeight) / scrollHeight;
      const horizontalScrollPercent = (scrollLeft + clientWidth) / scrollWidth;
      
      // Smart expansion - more conservative when performance is a concern
      const threshold = shouldUseVirtualScrolling ? 0.95 : 0.9; // Higher threshold when using virtual scrolling
      
      // Only expand if we're not already at performance limits
      if (verticalScrollPercent > threshold && visibleRows < MAX_ROWS) {
        const currentCellCount = visibleRows * visibleCols;
        if (currentCellCount < PERFORMANCE_THRESHOLD * 2) { // Safety check
          smartExpandGrid(visibleRows, visibleCols);
        }
      }
      
      if (horizontalScrollPercent > threshold && visibleCols < MAX_COLS) {
        const currentCellCount = visibleRows * visibleCols;
        if (currentCellCount < PERFORMANCE_THRESHOLD * 2) { // Safety check
          smartExpandGrid(visibleRows, visibleCols);
        }
      }
    };

    // Use modern performance-optimized scroll handler
    const optimizedHandleScroll = createOptimizedHandler(handleScroll as (...args: unknown[]) => void, {
      throttle: true,
      useRAF: true
    });

    document.addEventListener('scroll', optimizedHandleScroll, { passive: true, capture: true });
    
    return () => {
      document.removeEventListener('scroll', optimizedHandleScroll, true);
    };
  }, [visibleRows, visibleCols, shouldUseVirtualScrolling, updateViewport, smartExpandGrid]);

  // Container resize observer for virtual scrolling
  useEffect(() => {
    const container = scrollContainerRef.current;
    if (!container) return;

    const resizeObserver = new ResizeObserver(() => {
      // Update viewport when container size changes
      if (shouldUseVirtualScrolling) {
        updateViewport();
      }
    });

    resizeObserver.observe(container);

    return () => {
      resizeObserver.disconnect();
    };
  }, [shouldUseVirtualScrolling, updateViewport]);





  const getCellValue = useCallback((row: number, col: number): string | number => {
    return sheetData[row]?.[col]?.value ?? '';
  }, [sheetData]);

  // Internal function to set cell value directly (used by commands)
  const setCellValueDirect = useCallback((row: number, col: number, value: string | number) => {
    // Auto-expand grid if needed when setting cell values
    handleAutoExpansion(row, col);
    
    setSheetData(prev => ({
      ...prev,
      [row]: {
        ...prev[row],
        [col]: { 
          ...prev[row]?.[col],
          value 
        }
      }
    }));
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
    
    executeCommand(command);
  }, [setCellValueDirect, getCellValue, executeCommand, generateColumnLabel]);

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
      execute: () => {
        // Fill all target cells that are not part of the source range
        for (let targetRow = targetMinRow; targetRow <= targetMaxRow; targetRow++) {
          for (let targetCol = targetMinCol; targetCol <= targetMaxCol; targetCol++) {
            // Skip cells that are part of the source range
            if (targetRow >= sourceMinRow && targetRow <= sourceMaxRow && 
                targetCol >= sourceMinCol && targetCol <= sourceMaxCol) {
              continue;
            }
            
            // Calculate which source cell this target cell should copy from
            // Use modulo to wrap around the source pattern
            const relativeRow = (targetRow - sourceMinRow) % sourceRowCount;
            const relativeCol = (targetCol - sourceMinCol) % sourceColCount;
            
            // Calculate the index in the sourceValues array
            const sourceIndex = relativeRow * sourceColCount + relativeCol;
            
            // Get the value from the source pattern
            if (sourceIndex < sourceValues.length) {
              const value = sourceValues[sourceIndex];
              setCellValueDirect(targetRow, targetCol, value);
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

    executeCommand(command);
  }, [getCellValue, setCellValueDirect, executeCommand, generateColumnLabel]);

  // Fill handle mouse event handlers
  const handleFillHandleMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    if (!selectedCell && !selectedRange) return;
    
    setFillHandleActive(true);
    setDragStartCell(selectedCell);
    
    // Clear any existing selection
    setCopiedData(null);
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
        // Explicitly include empty cells (null, undefined, or empty string)
        rowData.push({ 
          value: value === null || value === undefined ? '' : value 
        });
      }
      data.push(rowData);
    }

    setCopiedData({ data, range: copiedRange, isCut });

    // Show notification
    const cellCount = selectedCells.length;
    const rangeText = cellCount === 1 
      ? `${generateColumnLabel(minCol)}${minRow + 1}`
      : `${generateColumnLabel(minCol)}${minRow + 1}:${generateColumnLabel(maxCol)}${maxRow + 1}`;
    
    const action = isCut ? 'Cut' : 'Copied';
    showClipboardNotification(`${action} ${cellCount} cell${cellCount > 1 ? 's' : ''} (${rangeText})`);

    // Auto-clear styling after 3 seconds for copy, or keep until paste for cut
    if (!isCut) {
      setTimeout(() => {
        setCopiedData(null);
      }, 3000);
    }

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
  }, [getSelectedCells, getCellValue, generateColumnLabel, showClipboardNotification]);

  const cutSelectedCells = useCallback(() => {
    copySelectedCells(true);
  }, [copySelectedCells]);

  const pasteFromText = useCallback((text: string) => {
    if (!selectedCell) return;

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
    if (copiedData?.isCut && copiedData.range) {
      const { range } = copiedData;
      const minRow = Math.min(range.startRow, range.endRow);
      const maxRow = Math.max(range.startRow, range.endRow);
      const minCol = Math.min(range.startCol, range.endCol);
      const maxCol = Math.max(range.startCol, range.endCol);

      for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
          // Only clear if it's not overlapping with paste target
          const isPasteTarget = data.some((_, rowIdx) =>
            data[0].some((_, colIdx) =>
              row === startRow + rowIdx && col === startCol + colIdx
            )
          );
          if (!isPasteTarget) {
            cutClearValues.push({
              row,
              col,
              value: getCellValue(row, col)
            });
          }
        }
      }
    }

    const command: Command = {
      execute: () => {
        // Clear cut cells first
        cutClearValues.forEach(({row, col}) => {
          setCellValueDirect(row, col, '');
        });

        // Paste new values
        data.forEach((rowData, rowIndex) => {
          rowData.forEach((cellValue, colIndex) => {
            const targetRow = startRow + rowIndex;
            const targetCol = startCol + colIndex;
            
            // Properly handle empty cells - they should be set as empty strings to overwrite destination
            if (cellValue === '' || cellValue === null || cellValue === undefined) {
              // Empty cell should overwrite destination with empty value
              setCellValueDirect(targetRow, targetCol, '');
            } else {
              // Convert to number if it's a valid number, otherwise keep as string
              const stringValue = String(cellValue);
              const trimmedValue = stringValue.trim();
              
              // Only convert to number if it's clearly a number and not an empty/whitespace string
              if (trimmedValue !== '' && !isNaN(Number(trimmedValue))) {
                const numValue = Number(trimmedValue);
                setCellValueDirect(targetRow, targetCol, numValue);
              } else {
                // Keep as string (including strings that look like numbers but have formatting)
                setCellValueDirect(targetRow, targetCol, stringValue);
              }
            }
          });
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
        if (copiedData?.isCut) {
          setCopiedData(copiedData);
        }
      },
      description: `Paste to ${generateColumnLabel(startCol)}${startRow + 1}`
    };

    executeCommand(command);
  }, [selectedCell, getCellValue, setCellValueDirect, executeCommand, generateColumnLabel, copiedData]);

  const pasteData = useCallback(() => {
    if (!copiedData || !selectedCell) {
      showClipboardNotification('No data to paste');
      return;
    }

    const startRow = selectedCell.row;
    const startCol = selectedCell.col;
    const { data, range, isCut } = copiedData;

    // Calculate paste area
    const pasteRows = data.length;
    const pasteCols = data[0]?.length || 0;
    const pasteEndRow = startRow + pasteRows - 1;
    const pasteEndCol = startCol + pasteCols - 1;

    // Show notification
    const pasteRangeText = `${generateColumnLabel(startCol)}${startRow + 1}:${generateColumnLabel(pasteEndCol)}${pasteEndRow + 1}`;
    const cellCount = pasteRows * pasteCols;
    showClipboardNotification(`Pasted ${cellCount} cell${cellCount > 1 ? 's' : ''} to ${pasteRangeText}`);

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
    if (isCut && range) {
      const minRow = Math.min(range.startRow, range.endRow);
      const maxRow = Math.max(range.startRow, range.endRow);
      const minCol = Math.min(range.startCol, range.endCol);
      const maxCol = Math.max(range.startCol, range.endCol);

      for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
          // Only clear if it's not the same cell we're pasting to
          const isPasteTarget = data.some((rowData, rowIdx) =>
            rowData.some((_, colIdx) =>
              row === startRow + rowIdx && col === startCol + colIdx
            )
          );
          if (!isPasteTarget) {
            cutClearValues.push({
              row,
              col,
              value: getCellValue(row, col)
            });
          }
        }
      }
    }

    const command: Command = {
      execute: () => {
        // Apply pasted data to the sheet
        data.forEach((rowData, rowIndex) => {
          rowData.forEach((cellData, colIndex) => {
            const targetRow = startRow + rowIndex;
            const targetCol = startCol + colIndex;
            
            // Properly handle empty cells - ensure they overwrite destination cells
            const value = cellData.value;
            if (value === null || value === undefined) {
              setCellValueDirect(targetRow, targetCol, '');
            } else {
              setCellValueDirect(targetRow, targetCol, value);
            }
          });
        });

        // If it was a cut operation, clear the original cells
        cutClearValues.forEach(({row, col}) => {
          setCellValueDirect(row, col, '');
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
        if (isCut) {
          setCopiedData(copiedData);
        }
      },
      description: `Paste to ${generateColumnLabel(startCol)}${startRow + 1}`
    };

    executeCommand(command);
  }, [copiedData, selectedCell, getCellValue, setCellValueDirect, executeCommand, generateColumnLabel, showClipboardNotification]);

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

    executeCommand(command);
  }, [getSelectedCells, getCellValue, setCellValueDirect, executeCommand, generateColumnLabel]);

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
    setCopiedData(null); // Clear any copied data
  }, [visibleRows]);

  const selectEntireRow = useCallback((rowIndex: number) => {
    setSelectedRange({
      startRow: rowIndex,
      startCol: 0,
      endRow: rowIndex,
      endCol: visibleCols - 1
    });
    setSelectedCell({ row: rowIndex, col: 0 });
    setCopiedData(null); // Clear any copied data
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

  const commitEdit = useCallback(() => {
    if (editingCell) {
      const numValue = Number(inputValue);
      const finalValue = !isNaN(numValue) && inputValue.trim() !== '' ? numValue : inputValue;
      setCellValue(editingCell.row, editingCell.col, finalValue);
      setEditingCell(null);
      setInputValue('');
    }
  }, [editingCell, inputValue, setCellValue]);

  const handlePasteFromClipboard = useCallback(async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const clipboardData = await navigator.clipboard.readText();
        if (clipboardData.trim()) {
          pasteFromText(clipboardData);
        } else if (copiedData) {
          pasteData();
        } else {
          showClipboardNotification('No data to paste');
        }
      } else if (copiedData) {
        pasteData();
      } else {
        showClipboardNotification('No data to paste');
      }
    } catch {
      // Fallback to internal copied data
      if (copiedData) {
        pasteData();
      } else {
        showClipboardNotification('Unable to access clipboard');
      }
    }
  }, [copiedData, pasteFromText, pasteData, showClipboardNotification]);

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
    
    executeCommand(command);
  }, [selectedCell, getCellValue, sheetData, executeCommand, generateColumnLabel, normalizeUrl]);

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
            
            newData[row][col] = {
              ...newData[row][col],
              formatting: {
                ...newData[row][col].formatting,
                ...formatting
              }
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

    executeCommand(command);
  }, [getSelectedCells, sheetData, executeCommand]);



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
              redo();
            }
          } else {
            // Ctrl+Z or Cmd+Z = Undo
            if (canUndo) {
              undo();
            }
          }
          e.preventDefault();
          e.stopPropagation();
        } else if (e.key === 'y' || e.key === 'Y') {
          // Ctrl+Y or Cmd+Y = Redo (alternative shortcut)
          if (canRedo) {
            redo();
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
  }, [editingCell, selectedCell, selectedRange, copySelectedCells, cutSelectedCells, deleteSelectedCells, handlePasteFromClipboard, selectAllCells, undo, redo, canUndo, canRedo]);

  // Restore cursor position after input value changes
  useEffect(() => {
    if (editingCell && inputRef.current) {
      // Use Promise.resolve to ensure cursor is set after React finishes rendering
      Promise.resolve().then(() => {
        inputRef.current?.setSelectionRange(cursorPositionRef.current, cursorPositionRef.current);
      });
    }
  }, [inputValue, editingCell]);

  const handleCellClick = useCallback((row: number, col: number, event?: React.MouseEvent) => {
    // Don't interfere with double-click events
    if (event?.detail === 2) {
      return;
    }
    
    if (editingCell) {
      commitEdit();
    }
    
    // Auto-expand grid when navigating to cells
    handleAutoExpansion(row, col);
    
    // Clear copied data when making new selections
    setCopiedData(null);
    
    if (event?.shiftKey && selectedCell) {
      // Extend selection with Shift+click
      setSelectedRange({
        startRow: selectedCell.row,
        startCol: selectedCell.col,
        endRow: row,
        endCol: col
      });
    } else {
      // Single cell selection
      setSelectedCell({ row, col });
      setSelectedRange(null);
    }
  }, [editingCell, commitEdit, handleAutoExpansion, selectedCell]);

  const handleMouseDown = useCallback((row: number, col: number, event: React.MouseEvent) => {
    // Don't interfere with double-click events
    if (event.detail === 2) {
      return;
    }
    
    if (editingCell) {
      commitEdit();
    }
    
    if (!event.shiftKey) {
      // Clear copied data when making new selections
      setCopiedData(null);
      setSelectedCell({ row, col });
      setSelectedRange(null);
      setIsSelecting(true);
    }
  }, [editingCell, commitEdit]);

  const handleMouseEnter = useCallback((row: number, col: number) => {
    if (fillHandleActive) {
      handleFillHandleMouseMove(row, col);
    } else if (isSelecting && selectedCell && !editingCell) {
      setSelectedRange({
        startRow: selectedCell.row,
        startCol: selectedCell.col,
        endRow: row,
        endCol: col
      });
    }
  }, [fillHandleActive, handleFillHandleMouseMove, isSelecting, selectedCell, editingCell]);

  const handleMouseUp = useCallback(() => {
    setIsSelecting(false);
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
      
      executeCommand(command);
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
  }, [isResizingColumn, columnWidths, tempColumnWidths, executeCommand, generateColumnLabel]);

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
      
      executeCommand(command);
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
  }, [isResizingRow, rowHeights, tempRowHeights, executeCommand]);

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

  const handleCellDoubleClick = useCallback((row: number, col: number, startWithKey?: string) => {
    const currentValue = getCellValue(row, col);
    
    // Enter editing mode
    setEditingCell({ row, col });
    setInputValue(startWithKey || String(currentValue));
    setSelectedCell({ row, col });
    setSelectedRange(null);
    setCopiedData(null);
    
    // Focus input with retry logic
    const focusInput = () => {
      if (inputRef.current) {
        inputRef.current.focus();
        inputRef.current.select();
      } else {
        setTimeout(focusInput, 5);
      }
    };
    
    // Multiple focus attempts for reliability
    setTimeout(focusInput, 0);
    setTimeout(focusInput, 5);
    setTimeout(focusInput, 10);
    setTimeout(focusInput, 20);
  }, [getCellValue]);

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
      // Clear copied data when navigating to new cells
      setCopiedData(null);
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
            redo();
          }
        } else {
          // Ctrl+Z or Cmd+Z = Undo
          if (canUndo) {
            undo();
          }
        }
        e.preventDefault();
        return;
      } else if (e.key === 'y') {
        // Ctrl+Y or Cmd+Y = Redo (alternative shortcut)
        if (canRedo) {
          redo();
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
      if (editingCell) {
        commitEdit();
        setSelectedCell({ row: row + 1, col });
      } else {
        handleCellDoubleClick(row, col);
      }
      e.preventDefault();
    } else if (e.key === 'Tab') {
      if (editingCell) {
        commitEdit();
      }
      setSelectedCell({ 
        row, 
        col: e.shiftKey ? Math.max(0, col - 1) : col + 1 
      });
      e.preventDefault();
    } else if (e.key === 'Escape') {
      if (editingCell) {
        setEditingCell(null);
        setInputValue('');
      }
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || 
               e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      if (!editingCell) {
        handleArrowKey(e.key, row, col, e.shiftKey);
        e.preventDefault();
      }
    } else if (!editingCell && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      // Clear copied data when starting to type
      setCopiedData(null);
      // Start editing if user types a character
      handleCellDoubleClick(row, col, e.key);
    }
  }, [editingCell, copySelectedCells, handlePasteFromClipboard, cutSelectedCells, selectAllCells, canRedo, redo, canUndo, undo, deleteSelectedCells, handleCellDoubleClick, commitEdit, handleArrowKey]);



  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    // Save cursor position before state update
    cursorPositionRef.current = e.target.selectionStart || 0;
    setInputValue(e.target.value);
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
      setInputValue('');
    }
  };

  // Memoized cell component for optimal performance
  const CellComponent = memo<{
    row: number;
    col: number;
    cellValue: string | number;
    cellData: CellData | undefined;
    isSelected: boolean;
    isEditing: boolean;
    isPrimarySelection: boolean;
    isCopied: boolean;
    isCut: boolean;
    isInFillPreview: boolean;
    shouldShowFillHandle: boolean;
    cellStyle: React.CSSProperties;
    contentStyle: React.CSSProperties;
    onCellClick: (row: number, col: number, event?: React.MouseEvent) => void;
    onCellDoubleClick: (row: number, col: number) => void;
    onMouseDown: (row: number, col: number, event: React.MouseEvent) => void;
    onMouseEnter: (row: number, col: number) => void;
    onMouseUp: () => void;
    onKeyDown: (e: React.KeyboardEvent, row: number, col: number) => void;
    onFillHandleMouseDown: (e: React.MouseEvent) => void;
    onLinkClick: (e: React.MouseEvent) => void;
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
    cellStyle,
    contentStyle,
    onCellClick,
    onCellDoubleClick,
    onMouseDown,
    onMouseEnter,
    onMouseUp,
    onKeyDown,
    onFillHandleMouseDown,
    onLinkClick,
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

    return (
      <td
        className={cellClasses}
        style={cellStyle}
        onClick={(e) => {
          // Only handle single clicks, not double clicks
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
        tabIndex={0}
      >
        {isEditing ? (
          <input
            ref={inputRef}
            type="text"
            value={inputValue}
            onChange={handleInputChange}
            onKeyDown={handleInputKeyDown}
            onBlur={commitEdit}
            className="cell-input"
            style={contentStyle}
            autoFocus
            onFocus={(e) => {
              e.target.select();
            }}
          />
        ) : (
          <div 
            className="cell-content" 
            style={{
              ...contentStyle,
              userSelect: 'none',
              pointerEvents: 'auto'
            }}
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
  });
  
  // Optimized helper functions for cell rendering (memoized)
  const getJustifyContent = useMemo(() => (align?: string) => {
    switch (align) {
      case 'left': return 'flex-start';
      case 'center': return 'center';
      case 'right': return 'flex-end';
      case 'justify': return 'space-between';
      default: return 'flex-start';
    }
  }, []);

  const getAlignItems = useMemo(() => (align?: string) => {
    switch (align) {
      case 'top': return 'flex-start';
      case 'middle': return 'center';
      case 'bottom': return 'flex-end';
      default: return 'center';
    }
  }, []);

  const renderCell = useCallback((row: number, col: number) => {
    const cellValue = getCellValue(row, col);
    const cellData = sheetData[row]?.[col];
    const formatting = cellData?.formatting || {};
    const isSelected = isCellSelected(row, col);
    const isEditing = editingCell?.row === row && editingCell?.col === col;
    const isPrimarySelection = selectedCell?.row === row && selectedCell?.col === col;
    const isCopied = isCellCopied(row, col);
    const isCut = isCellCut(row, col);
    const isInFillPreview = isCellInFillPreview(row, col);



    // Determine if this cell should show the fill handle
            const shouldShowFillHandle = !isEditing && (isPrimarySelection || 
      (selectedRange ? 
       row === Math.max(selectedRange.startRow, selectedRange.endRow) && 
       col === Math.max(selectedRange.startCol, selectedRange.endCol) : false));

    // Create cell style from formatting
    const cellStyle: React.CSSProperties = {
      backgroundColor: formatting.backgroundColor,
      textAlign: formatting.textAlign,
      verticalAlign: formatting.verticalAlign,
    };

    // Format the display value based on number format
    const displayValue = formatCellValue(cellValue, formatting);
    
    // Build text decoration
    let textDecoration = 'none';
    if (formatting.underline && formatting.strikethrough) {
      textDecoration = 'underline line-through';
    } else if (formatting.underline) {
      textDecoration = 'underline';
    } else if (formatting.strikethrough) {
      textDecoration = 'line-through';
    }

    // Create content style from formatting with flexbox alignment
    const contentStyle: React.CSSProperties = {
      fontFamily: formatting.fontFamily,
      fontSize: formatting.fontSize ? `${formatting.fontSize}px` : undefined,
      fontWeight: formatting.bold ? 'bold' : 'normal',
      fontStyle: formatting.italic ? 'italic' : 'normal',
      textDecoration,
      color: formatting.color,
      justifyContent: getJustifyContent(formatting.textAlign),
      alignItems: getAlignItems(formatting.verticalAlign),
                  cursor: formatting.link ? 'pointer' : 'inherit',
            position: 'relative',
    };

    // Handle link click - defined outside render callback
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
        cellData={cellData}
        isSelected={isSelected}
        isEditing={isEditing}
        isPrimarySelection={isPrimarySelection}
        isCopied={isCopied}
        isCut={isCut}
        isInFillPreview={isInFillPreview}
        shouldShowFillHandle={shouldShowFillHandle}
        cellStyle={cellStyle}
        contentStyle={contentStyle}
        onCellClick={handleCellClick}
        onCellDoubleClick={handleCellDoubleClick}
        onMouseDown={handleMouseDown}
        onMouseEnter={handleMouseEnter}
        onMouseUp={handleMouseUp}
        onKeyDown={handleKeyDown}
        onFillHandleMouseDown={handleFillHandleMouseDown}
        onLinkClick={cellLinkClick}
        hasLink={!!formatting.link}
        linkUrl={formatting.link}
      />
    );
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [getCellValue, sheetData, isCellSelected, editingCell, selectedCell, isCellCopied, isCellCut, isCellInFillPreview, selectedRange, handleCellClick, handleCellDoubleClick, handleMouseDown, handleMouseEnter, handleMouseUp, handleKeyDown, handleFillHandleMouseDown, getJustifyContent, getAlignItems, formatCellValue]);

  return (
    <div className="spreadsheet-container" tabIndex={0} onFocus={() => {}} onClick={() => {}}>
      {/* Excel Toolbar */}
      <ExcelToolbar
        onFormat={formatSelectedCells}
        onUndo={undo}
        onRedo={redo}
        onCopy={() => copySelectedCells(false)}
        // onCut={() => copySelectedCells(true)}
        onPaste={handlePasteFromClipboard}
        currentFormatting={getCurrentCellFormatting()}
        canUndo={canUndo}
        canRedo={canRedo}
        canPaste={!!copiedData || true}
        selectedCellValue={selectedCell ? getCellValue(selectedCell.row, selectedCell.col) : null}
        onInsertLink={handleInsertLink}
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
        }
        
        .spreadsheet-container:focus {
          outline: none;
        }


        
        .spreadsheet-table-container {
          flex: 1;
          overflow: auto;
          transform: scale(${zoomLevel / 100}) translateZ(0);
          transform-origin: top left;
          will-change: transform, scroll-position;
          contain: strict;
          contain-intrinsic-size: 100% 100%;
          backface-visibility: hidden;
          -webkit-backface-visibility: hidden;
          content-visibility: auto;
        }
        
        .spreadsheet-table {
          border-collapse: collapse;
          width: max-content;
          min-width: 100%;
          transform: translateZ(0);
          will-change: contents;
          table-layout: fixed;
          contain: strict;
          contain-intrinsic-size: 100% auto;
        }
        
        .header-row th,
        .row-header {
          background: linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%);
          border: 1px solid #a6a6a6 !important;
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
          border-collapse: collapse;
          contain: layout style;
          transition: background-color 0.15s ease;
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
          border: 1px solid #d4d4d4 !important;
          padding: 0;
          height: 20px;
          min-width: 80px;
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
          border-collapse: collapse;
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
          border: 1px solid #d4d4d4 !important;
          box-shadow: inset 0 0 0 1px rgba(33, 115, 70, 0.6) !important;
          position: relative;
          z-index: 40;
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
        }
        
        .cell.primary-selected {
          border: 3px solid #217346 !important;
          z-index: 50;
          box-shadow: 0 0 0 1px rgba(33, 115, 70, 0.3) !important;
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
          border: 1px solid #d4d4d4 !important;
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
        
        .cell.cut {
          border: 1px solid #d4d4d4 !important;
          z-index: 30;
          position: relative;
          animation: cut-border-pulse 1.5s ease-in-out infinite;
          box-shadow: inset 0 0 0 2px #d73502;
        }
        
        .cell.cut::before {
          content: '';
          position: absolute;
          top: -1px;
          left: -1px;
          right: -1px;
          bottom: -1px;
          border: 2px dotted #d73502;
          pointer-events: none;
          z-index: 1;
        }
        
        .cell.cut::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(215, 53, 2, 0.05);
          pointer-events: none;
          z-index: 1;
          animation: cut-pulse-overlay 1.5s ease-in-out infinite;
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
        
        @keyframes cut-pulse-overlay {
          0% { 
            background: rgba(215, 53, 2, 0.1);
          }
          50% { 
            background: rgba(215, 53, 2, 0.03);
          }
          100% { 
            background: rgba(215, 53, 2, 0.1);
          }
        }
        
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
        
        @keyframes cut-border-pulse {
          0% { 
            border-color: #d73502;
          }
          50% { 
            border-color: #ff6b3d;
          }
          100% { 
            border-color: #d73502;
          }
        }
        
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
          border: 1px solid #d4d4d4 !important;
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
      `}</style>
      
      <div className="spreadsheet-table-container" ref={scrollContainerRef}>
        <table className="spreadsheet-table">
          
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
                const viewport = calculateViewport;
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
              const viewport = calculateViewport;
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
                          border: '1px solid #d4d4d4',
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
                        border: '1px solid #d4d4d4',
                        background: 'white'
                      }}
                    ></td>
                  );
                }
                
                // Render buffered cells in this row
                for (let col = renderStartCol; col < renderEndCol; col++) {
                  const cell = renderCell(row, col);
                  rowCells.push(React.cloneElement(cell, {
                    key: `cell-${col}`,
                    style: {
                      ...cell.props.style,
                      width: getCurrentColumnWidth(col),
                      minWidth: MIN_COLUMN_WIDTH,
                      height: getCurrentRowHeight(row),
                      minHeight: MIN_ROW_HEIGHT
                    }
                  }));
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
                        border: '1px solid #d4d4d4',
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
                      border: '1px solid #a6a6a6',
                      background: 'linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%)',
                      padding: '4px 6px'
                    }}></td>
                    {Array.from({ length: visibleCols }, (_, colIndex) => (
                      <td 
                        key={`spacer-bottom-${colIndex}`}
                        className="cell-spacer"
                        style={{ 
                          width: getCurrentColumnWidth(colIndex),
                          border: '1px solid #d4d4d4',
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
      
      {/* Clipboard Notification */}
      {clipboardNotification && (
        <div className={`clipboard-notification ${clipboardNotification.includes('error') || clipboardNotification.includes('No') ? 'error' : clipboardNotification.includes('Unable') || clipboardNotification.includes('unavailable') ? 'info' : ''}`}>
          {clipboardNotification}
        </div>
      )}
    </div>
  );
});

// Set display name for better debugging
ExcelSpreadsheet.displayName = 'ExcelSpreadsheet';

export default ExcelSpreadsheet;
export type { CellFormatting };