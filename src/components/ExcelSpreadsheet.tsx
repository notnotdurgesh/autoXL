import React, { useState, useCallback, useRef, useEffect } from 'react';
import { useUndoRedo, type Command } from '../hooks/useUndoRedo';

interface CellData {
  value: string | number | null;
  isEditing?: boolean;
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

const ExcelSpreadsheet: React.FC = () => {
  const INITIAL_ROWS = 30;
const INITIAL_COLS = 15;
const DEFAULT_COLUMN_WIDTH = 80;
const DEFAULT_ROW_HEIGHT = 20;
const MIN_COLUMN_WIDTH = 30;
const MIN_ROW_HEIGHT = 15; // A-Z columns
  const [sheetData, setSheetData] = useState<SheetData>({});
  const [selectedCell, setSelectedCell] = useState<{row: number, col: number} | null>(null);
  const [selectedRange, setSelectedRange] = useState<CellRange | null>(null);
  const [editingCell, setEditingCell] = useState<{row: number, col: number} | null>(null);
  const [inputValue, setInputValue] = useState<string>('');
  const [isSelecting, setIsSelecting] = useState<boolean>(false);
  const [copiedData, setCopiedData] = useState<{data: CellData[][], range: CellRange, isCut?: boolean} | null>(null);
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
  const generateColumnLabel = useCallback((index: number): string => {
    let result = '';
    let temp = index;
    while (temp >= 0) {
      result = String.fromCharCode(65 + (temp % 26)) + result;
      temp = Math.floor(temp / 26) - 1;
    }
    return result;
  }, []);

  // Initialize with sample data
  useEffect(() => {
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
  }, []);





  const getCellValue = useCallback((row: number, col: number): string | number => {
    return sheetData[row]?.[col]?.value ?? '';
  }, [sheetData]);

  // Internal function to set cell value directly (used by commands)
  const setCellValueDirect = useCallback((row: number, col: number, value: string | number) => {
    setSheetData(prev => ({
      ...prev,
      [row]: {
        ...prev[row],
        [col]: { value }
      }
    }));
  }, []);

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

  const isCellInRange = (row: number, col: number, range: CellRange | null): boolean => {
    if (!range) return false;
    const minRow = Math.min(range.startRow, range.endRow);
    const maxRow = Math.max(range.startRow, range.endRow);
    const minCol = Math.min(range.startCol, range.endCol);
    const maxCol = Math.max(range.startCol, range.endCol);
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  };

  const isCellSelected = (row: number, col: number): boolean => {
    if (selectedRange) {
      return isCellInRange(row, col, selectedRange);
    }
    return selectedCell?.row === row && selectedCell?.col === col;
  };

  const isCellCopied = (row: number, col: number): boolean => {
    if (!copiedData) return false;
    return isCellInRange(row, col, copiedData.range);
  };

  const isCellCut = (row: number, col: number): boolean => {
    if (!copiedData || !copiedData.isCut) return false;
    return isCellInRange(row, col, copiedData.range);
  };

  const isCellInFillPreview = (row: number, col: number): boolean => {
    if (!fillPreview) return false;
    return isCellInRange(row, col, fillPreview);
  };





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

  const copySelectedCells = useCallback((isCut: boolean = false) => {
    const selectedCells = getSelectedCells();
    if (selectedCells.length === 0) return;

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
        console.log('Clipboard API failed, data copied internally');
      });
    }
  }, [getSelectedCells, getCellValue]);

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
    if (!copiedData || !selectedCell) return;

    const startRow = selectedCell.row;
    const startCol = selectedCell.col;
    const { data, range, isCut } = copiedData;

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
  }, [copiedData, selectedCell, getCellValue, setCellValueDirect, executeCommand, generateColumnLabel]);

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
      endRow: INITIAL_ROWS - 1,
      endCol: INITIAL_COLS - 1
    });
    setSelectedCell({ row: 0, col: 0 });
  }, []);

  const handlePasteFromClipboard = useCallback(async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const clipboardData = await navigator.clipboard.readText();
        pasteFromText(clipboardData);
      } else if (copiedData) {
        pasteData();
      }
    } catch {
      // Fallback to internal copied data
      if (copiedData) {
        pasteData();
      }
    }
  }, [copiedData, pasteFromText, pasteData]);

  // Handle global keyboard events for copy/paste/undo/redo
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      // Only handle if the spreadsheet area has focus
      const target = e.target as HTMLElement;
      const isSpreadsheetFocused = target.closest('.spreadsheet-container') !== null;
      
      if (!isSpreadsheetFocused || editingCell) return;

      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'c') {
          copySelectedCells();
          e.preventDefault();
        } else if (e.key === 'v') {
          handlePasteFromClipboard();
          e.preventDefault();
        } else if (e.key === 'x') {
          cutSelectedCells();
          e.preventDefault();
        } else if (e.key === 'a') {
          selectAllCells();
          e.preventDefault();
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
        } else if (e.key === 'y') {
          // Ctrl+Y or Cmd+Y = Redo (alternative shortcut)
          if (canRedo) {
            redo();
          }
          e.preventDefault();
        }
      } else if (e.key === 'Delete') {
        deleteSelectedCells();
        e.preventDefault();
      }
    };

    document.addEventListener('keydown', handleGlobalKeyDown);
    return () => {
      document.removeEventListener('keydown', handleGlobalKeyDown);
    };
  }, [editingCell, copySelectedCells, cutSelectedCells, deleteSelectedCells, handlePasteFromClipboard, selectAllCells, undo, redo, canUndo, canRedo]);

  // Restore cursor position after input value changes
  useEffect(() => {
    if (editingCell && inputRef.current) {
      // Use Promise.resolve to ensure cursor is set after React finishes rendering
      Promise.resolve().then(() => {
        inputRef.current?.setSelectionRange(cursorPositionRef.current, cursorPositionRef.current);
      });
    }
  }, [inputValue, editingCell]);

  const handleCellClick = (row: number, col: number, event?: React.MouseEvent) => {
    if (editingCell) {
      commitEdit();
    }
    
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
  };

  const handleMouseDown = (row: number, col: number, event: React.MouseEvent) => {
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
  };

  const handleMouseEnter = (row: number, col: number) => {
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
  };

  const handleMouseUp = () => {
    setIsSelecting(false);
  };

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

  const handleCellDoubleClick = (row: number, col: number, startWithKey?: string) => {
    const currentValue = getCellValue(row, col);
    setEditingCell({ row, col });
    setInputValue(startWithKey || String(currentValue));
    setSelectedCell({ row, col });
    setSelectedRange(null); // Clear range selection when editing
    setCopiedData(null); // Clear copied data when editing
    
    // Focus input on next render
    setTimeout(() => {
      inputRef.current?.focus();
      // Position cursor at the end instead of selecting all text
      if (startWithKey) {
        // If starting with a key, position cursor at the end
        inputRef.current?.setSelectionRange(startWithKey.length, startWithKey.length);
      } else {
        // If double-clicking to edit existing value, position cursor at the end
        const valueLength = String(currentValue).length;
        inputRef.current?.setSelectionRange(valueLength, valueLength);
      }
    }, 0);
  };

  const handleKeyDown = (e: React.KeyboardEvent, row: number, col: number) => {
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
  };

        const handleArrowKey = (key: string, row: number, col: number, shiftKey: boolean = false) => {
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
  };

  const commitEdit = () => {
    if (editingCell) {
      const numValue = Number(inputValue);
      const finalValue = !isNaN(numValue) && inputValue.trim() !== '' ? numValue : inputValue;
      setCellValue(editingCell.row, editingCell.col, finalValue);
      setEditingCell(null);
      setInputValue('');
    }
  };

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

        const renderCell = (row: number, col: number) => {
    const cellValue = getCellValue(row, col);
    const isSelected = isCellSelected(row, col);
    const isEditing = editingCell?.row === row && editingCell?.col === col;
    const isPrimarySelection = selectedCell?.row === row && selectedCell?.col === col;
    const isCopied = isCellCopied(row, col);
    const isCut = isCellCut(row, col);
    const isInFillPreview = isCellInFillPreview(row, col);
    
    let cellClasses = 'cell';
    if (isSelected) cellClasses += ' selected';
    if (isPrimarySelection) cellClasses += ' primary-selected';
    if (isCut) cellClasses += ' cut';
    else if (isCopied) cellClasses += ' copied';
    if (isInFillPreview) cellClasses += ' fill-preview';

    // Determine if this cell should show the fill handle
    const shouldShowFillHandle = !isEditing && (isPrimarySelection || 
      (selectedRange && 
       row === Math.max(selectedRange.startRow, selectedRange.endRow) && 
       col === Math.max(selectedRange.startCol, selectedRange.endCol)));

    return (
      <td
        key={`${row}-${col}`}
        className={cellClasses}
        onClick={(e) => handleCellClick(row, col, e)}
        onDoubleClick={() => handleCellDoubleClick(row, col)}
        onMouseDown={(e) => handleMouseDown(row, col, e)}
        onMouseEnter={() => handleMouseEnter(row, col)}
        onMouseUp={handleMouseUp}
        onKeyDown={(e) => handleKeyDown(e, row, col)}
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
          />
        ) : (
          <div className="cell-content">
            {String(cellValue)}
            {shouldShowFillHandle && (
              <div 
                className="fill-handle"
                onMouseDown={handleFillHandleMouseDown}
              />
            )}
          </div>
        )}
      </td>
    );
  };

  return (
    <div className="spreadsheet-container" tabIndex={0}>
      {/* Undo/Redo Toolbar */}
      <div className="undo-redo-toolbar">
        <button 
          className={`undo-button ${canUndo ? 'enabled' : 'disabled'}`}
          onClick={undo}
          disabled={!canUndo}
          title="Undo (Ctrl+Z)"
        >
          ↶ Undo
        </button>
        <button 
          className={`redo-button ${canRedo ? 'enabled' : 'disabled'}`}
          onClick={redo}
          disabled={!canRedo}
          title="Redo (Ctrl+Y or Ctrl+Shift+Z)"
        >
          ↷ Redo
        </button>
        <div className="keyboard-shortcuts-help">
          <small>Ctrl+Z: Undo | Ctrl+Y: Redo | Ctrl+C: Copy | Ctrl+V: Paste | Ctrl+X: Cut | Delete: Clear</small>
        </div>
      </div>
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
        }
        
        .undo-redo-toolbar {
          padding: 8px 12px;
          background: #f8f9fa;
          border-bottom: 1px solid #e9ecef;
          display: flex;
          align-items: center;
          gap: 8px;
          flex-shrink: 0;
        }
        
        .undo-button, .redo-button {
          padding: 4px 12px;
          border: 1px solid #d4d4d4;
          background: white;
          color: #333;
          border-radius: 3px;
          font-size: 11px;
          cursor: pointer;
          transition: all 0.2s ease;
        }
        
        .undo-button.enabled:hover, .redo-button.enabled:hover {
          background: #f0f0f0;
          border-color: #217346;
        }
        
        .undo-button.disabled, .redo-button.disabled {
          background: #f5f5f5;
          color: #999;
          cursor: not-allowed;
        }
        
        .keyboard-shortcuts-help {
          margin-left: auto;
          color: #666;
        }

        
        .spreadsheet-table-container {
          flex: 1;
          overflow: auto;
        }
        
        .spreadsheet-table {
          border-collapse: collapse;
          width: max-content;
          min-width: 100%;
        }
        
        .header-row th,
        .row-header {
          background: linear-gradient(180deg, #f0f0f0 0%, #d0d0d0 100%);
          border: 1px solid #a6a6a6;
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
        
        .cell {
          border: 1px solid #d4d4d4;
          padding: 0;
          height: 20px;
          min-width: 80px;
          width: 80px;
          background: white;
          position: relative;
          cursor: cell;
          outline: none;
        }
        
        .cell:first-child {
          position: sticky;
          left: 0;
          z-index: 10;
        }
        
        .cell.selected {
          background: rgba(33, 115, 70, 0.1) !important;
          z-index: 40;
        }
        
        .cell.primary-selected {
          border: 2px solid #217346 !important;
          z-index: 50;
          background: white !important;
        }
        
        .cell.copied {
          border: 2px dashed #0078d4 !important;
          animation: copied-pulse 2s ease-in-out infinite;
          z-index: 30;
        }
        
        .cell.cut {
          border: 2px dotted #d73502 !important;
          animation: cut-pulse 1.5s ease-in-out infinite;
          z-index: 30;
          background: rgba(215, 53, 2, 0.05) !important;
        }
        
        @keyframes copied-pulse {
          0% { 
            background: rgba(0, 120, 212, 0.15);
            border-color: #0078d4;
          }
          50% { 
            background: rgba(0, 120, 212, 0.05);
            border-color: #4a9eff;
          }
          100% { 
            background: rgba(0, 120, 212, 0.15);
            border-color: #0078d4;
          }
        }
        
        @keyframes cut-pulse {
          0% { 
            background: rgba(215, 53, 2, 0.1);
            border-color: #d73502;
          }
          50% { 
            background: rgba(215, 53, 2, 0.03);
            border-color: #ff6b3d;
          }
          100% { 
            background: rgba(215, 53, 2, 0.1);
            border-color: #d73502;
          }
        }
        
        .cell:hover:not(.selected):not(.primary-selected):not(.copied):not(.cut) {
          background: #f5f5f5;
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
          align-items: center;
          box-sizing: border-box;
          background: transparent;
          border: none;
          outline: none;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          color: black;
        }
        
        .cell-input {
          width: 100%;
          height: 100%;
          padding: 2px 4px;
          border: none;
          outline: none;
          background: white;
          font-family: inherit;
          font-size: inherit;
          box-sizing: border-box;
          z-index: 200;
          color: black;
          user-select: text;
        }
        
        .cell-input:focus {
          border: 2px solid #217346;
        }
        
        .cell.fill-preview {
          background: rgba(33, 115, 70, 0.2) !important;
          border: 1px dashed #217346 !important;
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
        }
        
        .fill-handle:hover {
          background: #1a5a37;
          transform: scale(1.2);
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
          z-index: 10;
        }
        
        .column-resize-handle:hover {
          background: #217346;
          opacity: 0.7;
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
          z-index: 10;
        }
        
        .row-resize-handle:hover {
          background: #217346;
          opacity: 0.7;
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
      `}</style>
      
      <div className="spreadsheet-table-container">
        <table className="spreadsheet-table">
          {/* Header row */}
          <thead>
            <tr className="header-row">
              <th className="row-col-header"></th>
              {Array.from({ length: INITIAL_COLS }, (_, i) => (
                <th 
                  key={i} 
                  className="column-header"
                  style={{ 
                    width: getCurrentColumnWidth(i),
                    minWidth: MIN_COLUMN_WIDTH,
                    position: 'relative'
                  }}
                >
                  {generateColumnLabel(i)}
                  <div 
                    className="column-resize-handle"
                    onMouseDown={(e) => handleColumnResizeStart(e, i)}
                  />
                </th>
              ))}
            </tr>
          </thead>
          
          {/* Data rows */}
          <tbody>
            {Array.from({ length: INITIAL_ROWS }, (_, row) => (
              <tr 
                key={row}
                style={{
                  height: getCurrentRowHeight(row),
                  minHeight: MIN_ROW_HEIGHT
                }}
              >
                <td 
                  className="row-header"
                  style={{ 
                    position: 'relative',
                    height: getCurrentRowHeight(row)
                  }}
                >
                  {row + 1}
                  <div 
                    className="row-resize-handle"
                    onMouseDown={(e) => handleRowResizeStart(e, row)}
                  />
                </td>
                {Array.from({ length: INITIAL_COLS }, (_, col) => {
                  const cell = renderCell(row, col);
                  return React.cloneElement(cell, {
                    style: {
                      ...cell.props.style,
                      width: getCurrentColumnWidth(col),
                      minWidth: MIN_COLUMN_WIDTH,
                      height: getCurrentRowHeight(row),
                      minHeight: MIN_ROW_HEIGHT
                    }
                  });
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default ExcelSpreadsheet;