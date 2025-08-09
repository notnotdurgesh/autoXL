
import type { CellFormatting } from '../../types/cell.types';

// Interface for spreadsheet operations
export interface SpreadsheetOperations {
  getCellValue: (row: number, col: number) => string | number | null;
  setCellValue: (row: number, col: number, value: string | number | null) => void;
  getRangeValues: (startRow: number, startCol: number, endRow: number, endCol: number) => (string | number | null)[][];
  setRangeValues: (startRow: number, startCol: number, values: (string | number | null)[][]) => void;
  getCellData?: (row: number, col: number) => {value: string | number | null; formatting?: CellFormatting};
  getRangeData?: (startRow: number, startCol: number, endRow: number, endCol: number) => {value: string | number | null; formatting?: CellFormatting}[][];
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
  }) => void;
  setCellFormula: (row: number, col: number, formula: string) => void;
  autoFill: (params: {sourceRow: number; sourceCol: number; targetEndRow: number; targetEndCol: number; fillType?: string}) => void;
  insertRows: (position: number, count: number) => void;
  insertColumns: (position: number, count: number) => void;
  deleteRows: (startRow: number, count: number) => void;
  deleteColumns: (startCol: number, count: number) => void;
  sortRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; sortColumn: number; order: string}) => void;
  filterRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; filterColumn: number; criteria: string; value: string}) => void;
  findReplace: (params: {findValue: string; replaceValue: string; matchCase?: boolean; matchEntireCell?: boolean; rangeStartRow?: number; rangeStartCol?: number; rangeEndRow?: number; rangeEndCol?: number}) => number;
  createChart: (params: {chartType: string; dataStartRow: number; dataStartCol: number; dataEndRow: number; dataEndCol: number; title?: string; xAxisLabel?: string; yAxisLabel?: string}) => void;
  clearRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number; clearContent?: boolean; clearFormatting?: boolean}) => void;
  mergeRange: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => void;
  unmergeRange: (row: number, col: number) => void;
  addHyperlink: (params: {row: number; col: number; url: string; displayText?: string}) => void;
  calculateSum: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => number;
  calculateAverage: (params: {startRow: number; startCol: number; endRow: number; endCol: number}) => number;
}

// Create function handlers for Gemini
export function createSpreadsheetHandlers(operations: SpreadsheetOperations) {
  return {
    // Sheet scanning - for "what's in the sheet?" queries
    scan_sheet: async (args: {maxRows?: number; maxCols?: number; includeFormatting?: boolean}) => {
      try {
        const maxRows = args.maxRows || 100;
        const maxCols = args.maxCols || 26; // A-Z
        const includeFormatting = args.includeFormatting || false;
        
        // If getRangeData is available and formatting is requested, use it
        const useFormatting = includeFormatting && operations.getRangeData;
        
        const dataFound: Array<{cell: string; value: string | number | null; formatting?: CellFormatting}> = [];
        let lastDataRow = 0;
        let lastDataCol = 0;
        
        if (useFormatting) {
          // Get data with formatting
          const data = operations.getRangeData!(0, 0, maxRows - 1, maxCols - 1);
          
          for (let row = 0; row < data.length; row++) {
            for (let col = 0; col < data[row].length; col++) {
              const cellData = data[row][col];
              const value = cellData.value;
              if (value !== null && value !== '') {
                lastDataRow = Math.max(lastDataRow, row);
                lastDataCol = Math.max(lastDataCol, col);
                if (dataFound.length < 50) {
                  const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;
                  dataFound.push({ 
                    cell: cellRef, 
                    value,
                    formatting: cellData.formatting || {}
                  });
                }
              }
            }
          }
        } else {
          // Get values only (backward compatible)
          const values = operations.getRangeValues(0, 0, maxRows - 1, maxCols - 1);
          
          for (let row = 0; row < values.length; row++) {
            for (let col = 0; col < values[row].length; col++) {
              const value = values[row][col];
              if (value !== null && value !== '') {
                lastDataRow = Math.max(lastDataRow, row);
                lastDataCol = Math.max(lastDataCol, col);
                if (dataFound.length < 50) {
                  const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;
                  dataFound.push({ cell: cellRef, value });
                }
              }
            }
          }
        }
        
        if (dataFound.length === 0) {
          return {
            isEmpty: true,
            message: 'Sheet is empty',
            dataCount: 0
          };
        }
        
        const dataRange = `A1:${String.fromCharCode(65 + lastDataCol)}${lastDataRow + 1}`;
        
        // Create a structured response with optional formatting info
        const cellData = dataFound.slice(0, 20).map(item => {
          const result: Record<string, unknown> = {
            cell: item.cell,
            value: String(item.value)
          };
          
          // Include formatting details if available
          if (item.formatting) {
            const format: Record<string, unknown> = {};
            if (item.formatting.bold) format.bold = true;
            if (item.formatting.italic) format.italic = true;
            if (item.formatting.color) format.textColor = item.formatting.color;
            if (item.formatting.backgroundColor) format.bgColor = item.formatting.backgroundColor;
            if (item.formatting.fontSize) format.fontSize = item.formatting.fontSize;
            if (item.formatting.textAlign) format.align = item.formatting.textAlign;
            if (Object.keys(format).length > 0) {
              result.formatting = format;
            }
          }
          
          return result;
        });
        
        return {
          isEmpty: false,
          dataRange: dataRange,
          dataCount: dataFound.length,
          cells: cellData,
          message: `Found ${dataFound.length} cells with data in range ${dataRange}${includeFormatting ? ' (with formatting)' : ''}`
        };
      } catch (error) {
        return {
          success: false,
          error: String(error),
          message: 'Failed to scan sheet'
        };
      }
    },
    
    // Cell value operations
    get_cell_value: async (args: {row: number; col: number}) => {
      try {
        const { row, col } = args;
        const value = operations.getCellValue(row - 1, col - 1); // Convert to 0-based indexing
        return {
          success: true,
          value: value ?? '',
          cell: `${String.fromCharCode(64 + col)}${row}`,
        };
      } catch (error) {
        return {
          success: false,
          error: String(error),
          cell: `${String.fromCharCode(64 + args.col)}${args.row}`,
        };
      }
    },

    // Get cell data with formatting
    get_cell_data: async (args: {row: number; col: number}) => {
      try {
        const { row, col } = args;
        
        // Use getCellData if available, otherwise fall back to getCellValue
        if (operations.getCellData) {
          const cellData = operations.getCellData(row - 1, col - 1);
          const result: Record<string, unknown> = {
            success: true,
            cell: `${String.fromCharCode(64 + col)}${row}`,
            value: cellData.value ?? '',
          };
          
          // Include formatting if present
          if (cellData.formatting && Object.keys(cellData.formatting).length > 0) {
            const format: Record<string, unknown> = {};
            if (cellData.formatting.bold) format.bold = true;
            if (cellData.formatting.italic) format.italic = true;
            if (cellData.formatting.color) format.textColor = cellData.formatting.color;
            if (cellData.formatting.backgroundColor) format.bgColor = cellData.formatting.backgroundColor;
            if (cellData.formatting.fontSize) format.fontSize = cellData.formatting.fontSize;
            if (cellData.formatting.textAlign) format.align = cellData.formatting.textAlign;
            result.formatting = format;
          }
          
          return result;
        } else {
          // Fallback to getValue
          const value = operations.getCellValue(row - 1, col - 1);
          return {
            success: true,
            cell: `${String.fromCharCode(64 + col)}${row}`,
            value: value ?? '',
          };
        }
      } catch (error) {
        return {
          success: false,
          error: String(error),
          cell: `${String.fromCharCode(64 + args.col)}${args.row}`,
        };
      }
    },

    set_cell_value: async (args: {row: number; col: number; value: string | number | null}) => {
      try {
        const { row, col, value } = args;
        operations.setCellValue(row - 1, col - 1, value);
        return {
          success: true,
          cell: `${String.fromCharCode(64 + col)}${row}`,
          value,
          message: `Set ${String.fromCharCode(64 + col)}${row} to ${value}`
        };
      } catch (error) {
        return {
          success: false,
          error: String(error),
          cell: `${String.fromCharCode(64 + args.col)}${args.row}`,
        };
      }
    },

    get_range_values: async (args: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      try {
        const { startRow, startCol, endRow, endCol } = args;
        const values = operations.getRangeValues(
          startRow - 1,
          startCol - 1,
          endRow - 1,
          endCol - 1
        );
        
        // Format the response to be more informative
        const range = `${String.fromCharCode(64 + startCol)}${startRow}:${String.fromCharCode(64 + endCol)}${endRow}`;
        const hasData = values.some(row => row.some(cell => cell !== null && cell !== ''));
        
        return {
          success: true,
          values,
          range,
          hasData,
          summary: hasData ? `Found data in ${range}` : `Range ${range} is empty`,
          rowCount: values.length,
          colCount: values[0]?.length || 0,
        };
      } catch (error) {
        return {
          success: false,
          error: String(error),
          range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
        };
      }
    },

    set_range_values: async (args: {startRow: number; startCol: number; values: (string | number | null)[][]}) => {
      const { startRow, startCol, values } = args;
      operations.setRangeValues(startRow - 1, startCol - 1, values);
      return {
        success: true,
        range: `${String.fromCharCode(64 + startCol)}${startRow}`,
        rowsUpdated: values.length,
        colsUpdated: values[0]?.length || 0,
      };
    },

    // Formatting operations
    set_cell_formatting: async (args: {startRow: number; startCol: number; endRow?: number; endCol?: number; bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean; fontSize?: number; fontFamily?: string; textColor?: string; backgroundColor?: string; textAlign?: string; verticalAlign?: string; numberFormat?: string; decimalPlaces?: number}) => {
      // Build formatting params with correct typing
      const formattingParams = {
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: (args.endRow || args.startRow) - 1,
        endCol: (args.endCol || args.startCol) - 1,
        bold: args.bold,
        italic: args.italic,
        underline: args.underline,
        strikethrough: args.strikethrough,
        fontSize: args.fontSize,
        fontFamily: args.fontFamily,
        color: args.textColor, // Map textColor to color
        backgroundColor: args.backgroundColor,
        textAlign: args.textAlign as ('left' | 'center' | 'right' | 'justify' | undefined),
        verticalAlign: args.verticalAlign as ('top' | 'middle' | 'bottom' | undefined),
        numberFormat: args.numberFormat,
        decimalPlaces: args.decimalPlaces,
      };
      
      operations.setCellFormatting(formattingParams);
      
      const range = `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + (args.endCol || args.startCol))}${args.endRow || args.startRow}`;
      return {
        success: true,
        range,
        message: `Applied formatting to ${range}`,
        appliedFormatting: formattingParams
      };
    },

    // Formula operations
    set_cell_formula: async (args: {row: number; col: number; formula: string}) => {
      const { row, col, formula } = args;
      operations.setCellFormula(row - 1, col - 1, formula);
      return {
        success: true,
        cell: `${String.fromCharCode(64 + col)}${row}`,
        formula,
      };
    },

    auto_fill: async (args: {sourceRow: number; sourceCol: number; targetEndRow: number; targetEndCol: number; fillType?: string}) => {
      operations.autoFill({
        sourceRow: args.sourceRow - 1,
        sourceCol: args.sourceCol - 1,
        targetEndRow: args.targetEndRow - 1,
        targetEndCol: args.targetEndCol - 1,
        fillType: args.fillType || 'copy',
      });
      return {
        success: true,
        sourceCell: `${String.fromCharCode(64 + args.sourceCol)}${args.sourceRow}`,
        targetRange: `${String.fromCharCode(64 + args.targetEndCol)}${args.targetEndRow}`,
      };
    },

    // Row and column operations
    insert_rows: async (args: {position: number; count: number}) => {
      const { position, count } = args;
      operations.insertRows(position - 1, count);
      return {
        success: true,
        inserted: count,
        position,
      };
    },

    insert_columns: async (args: {position: number; count: number}) => {
      const { position, count } = args;
      operations.insertColumns(position - 1, count);
      return {
        success: true,
        inserted: count,
        position: String.fromCharCode(64 + position),
      };
    },

    delete_rows: async (args: {startRow: number; count: number}) => {
      const { startRow, count } = args;
      operations.deleteRows(startRow - 1, count);
      return {
        success: true,
        deleted: count,
        startRow,
      };
    },

    delete_columns: async (args: {startCol: number; count: number}) => {
      const { startCol, count } = args;
      operations.deleteColumns(startCol - 1, count);
      return {
        success: true,
        deleted: count,
        startCol: String.fromCharCode(64 + startCol),
      };
    },

    // Data operations
    sort_range: async (args: {startRow: number; startCol: number; endRow: number; endCol: number; sortColumn: number; order: string}) => {
      operations.sortRange({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
        sortColumn: args.sortColumn - 1,
        order: args.order,
      });
      return {
        success: true,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
        sortedBy: String.fromCharCode(64 + args.sortColumn),
        order: args.order,
      };
    },

    filter_range: async (args: {startRow: number; startCol: number; endRow: number; endCol: number; filterColumn: number; criteria: string; value: string}) => {
      operations.filterRange({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
        filterColumn: args.filterColumn - 1,
        criteria: args.criteria,
        value: args.value,
      });
      return {
        success: true,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
        filteredBy: String.fromCharCode(64 + args.filterColumn),
      };
    },

    find_replace: async (args: {findValue: string; replaceValue: string; matchCase?: boolean; matchEntireCell?: boolean; rangeStartRow?: number; rangeStartCol?: number; rangeEndRow?: number; rangeEndCol?: number}) => {
      const count = operations.findReplace({
        findValue: args.findValue,
        replaceValue: args.replaceValue,
        matchCase: args.matchCase || false,
        matchEntireCell: args.matchEntireCell || false,
        rangeStartRow: args.rangeStartRow ? args.rangeStartRow - 1 : undefined,
        rangeStartCol: args.rangeStartCol ? args.rangeStartCol - 1 : undefined,
        rangeEndRow: args.rangeEndRow ? args.rangeEndRow - 1 : undefined,
        rangeEndCol: args.rangeEndCol ? args.rangeEndCol - 1 : undefined,
      });
      return {
        success: true,
        replacements: count,
      };
    },

    // Chart operations
    create_chart: async (args: {chartType: string; dataStartRow: number; dataStartCol: number; dataEndRow: number; dataEndCol: number; title?: string; xAxisLabel?: string; yAxisLabel?: string}) => {
      operations.createChart({
        chartType: args.chartType,
        dataStartRow: args.dataStartRow - 1,
        dataStartCol: args.dataStartCol - 1,
        dataEndRow: args.dataEndRow - 1,
        dataEndCol: args.dataEndCol - 1,
        title: args.title,
        xAxisLabel: args.xAxisLabel,
        yAxisLabel: args.yAxisLabel,
      });
      return {
        success: true,
        chartType: args.chartType,
        dataRange: `${String.fromCharCode(64 + args.dataStartCol)}${args.dataStartRow}:${String.fromCharCode(64 + args.dataEndCol)}${args.dataEndRow}`,
      };
    },

    // Find cell by content
    find_cell: async (args: {searchText: string; matchCase?: boolean; matchEntireCell?: boolean}) => {
      const { searchText, matchCase = false, matchEntireCell = false } = args;
      
      // Scan sheet to find the cell
      const maxRows = 100;
      const maxCols = 26;
      const values = operations.getRangeValues(0, 0, maxRows - 1, maxCols - 1);
      
      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const cellValue = String(values[row][col] || '');
          const searchValue = String(searchText);
          
          let matches = false;
          if (matchEntireCell) {
            matches = matchCase 
              ? cellValue === searchValue 
              : cellValue.toLowerCase() === searchValue.toLowerCase();
          } else {
            matches = matchCase 
              ? cellValue.includes(searchValue)
              : cellValue.toLowerCase().includes(searchValue.toLowerCase());
          }
          
          if (matches) {
            return {
              found: true,
              row: row + 1, // Convert to 1-based
              col: col + 1, // Convert to 1-based
              cell: `${String.fromCharCode(65 + col)}${row + 1}`,
              value: cellValue,
              message: `Found "${searchText}" in cell ${String.fromCharCode(65 + col)}${row + 1}`,
            };
          }
        }
      }
      
      return {
        found: false,
        message: `Could not find "${searchText}" in the sheet`,
      };
    },
    
    // Clear/delete single cell (convenience function)
    clear_cell: async (args: {row: number; col: number}) => {
      const { row, col } = args;
      operations.setCellValue(row - 1, col - 1, '');
      return {
        success: true,
        cell: `${String.fromCharCode(64 + col)}${row}`,
        message: `Cleared cell ${String.fromCharCode(64 + col)}${row}`,
      };
    },
    
    delete_cell: async (args: {row: number; col: number}) => {
      // Alias for clear_cell - deletes ENTIRE cell content
      const { row, col } = args;
      operations.setCellValue(row - 1, col - 1, '');
      return {
        success: true,
        cell: `${String.fromCharCode(64 + col)}${row}`,
        message: `Deleted entire content of cell ${String.fromCharCode(64 + col)}${row}`,
      };
    },
    
    // Utility operations
    clear_range: async (args: {startRow: number; startCol: number; endRow: number; endCol: number; clearContent?: boolean; clearFormatting?: boolean}) => {
      operations.clearRange({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
        clearContent: args.clearContent !== false,
        clearFormatting: args.clearFormatting === true,
      });
      return {
        success: true,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
      };
    },

    merge_range: async (args: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      operations.mergeRange({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
      });
      return {
        success: true,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
      };
    },

    unmerge_range: async (args: {row: number; col: number}) => {
      const { row, col } = args;
      operations.unmergeRange(row - 1, col - 1);
      return {
        success: true,
        cell: `${String.fromCharCode(64 + col)}${row}`,
      };
    },

    add_hyperlink: async (args: {row: number; col: number; url: string; displayText?: string}) => {
      operations.addHyperlink({
        row: args.row - 1,
        col: args.col - 1,
        url: args.url,
        displayText: args.displayText || args.url,
      });
      return {
        success: true,
        cell: `${String.fromCharCode(64 + args.col)}${args.row}`,
        url: args.url,
      };
    },

    // Advanced Cell Operations
    copy_cells: async (args: {
      sourceStartRow: number;
      sourceStartCol: number;
      sourceEndRow: number;
      sourceEndCol: number;
      destStartRow: number;
      destStartCol: number;
      transpose?: boolean;
      includeFormatting?: boolean;
    }) => {
      const sourceData = operations.getRangeValues(
        args.sourceStartRow - 1,
        args.sourceStartCol - 1,
        args.sourceEndRow - 1,
        args.sourceEndCol - 1
      );

      let dataToSet = sourceData;
      if (args.transpose) {
        // Transpose the data
        dataToSet = sourceData[0].map((_, colIndex) =>
          sourceData.map(row => row[colIndex])
        );
      }

      operations.setRangeValues(
        args.destStartRow - 1,
        args.destStartCol - 1,
        dataToSet
      );

      if (args.includeFormatting) {
        // Copy formatting as well - apply basic formatting
        operations.setCellFormatting({
          startRow: args.destStartRow - 1,
          startCol: args.destStartCol - 1,
          endRow: args.destStartRow + dataToSet.length - 2,
          endCol: args.destStartCol + dataToSet[0].length - 2,
          bold: true,
        });
      }

      return {
        success: true,
        sourceRange: `${String.fromCharCode(64 + args.sourceStartCol)}${args.sourceStartRow}:${String.fromCharCode(64 + args.sourceEndCol)}${args.sourceEndRow}`,
        destRange: `${String.fromCharCode(64 + args.destStartCol)}${args.destStartRow}`,
      };
    },

    move_cells: async (args: {
      sourceStartRow: number;
      sourceStartCol: number;
      sourceEndRow: number;
      sourceEndCol: number;
      destStartRow: number;
      destStartCol: number;
      shiftCells?: boolean;
    }) => {
      // Get source data
      const sourceData = operations.getRangeValues(
        args.sourceStartRow - 1,
        args.sourceStartCol - 1,
        args.sourceEndRow - 1,
        args.sourceEndCol - 1
      );

      // Set data at destination
      operations.setRangeValues(
        args.destStartRow - 1,
        args.destStartCol - 1,
        sourceData
      );

      // Clear source
      operations.clearRange({
        startRow: args.sourceStartRow - 1,
        startCol: args.sourceStartCol - 1,
        endRow: args.sourceEndRow - 1,
        endCol: args.sourceEndCol - 1,
        clearContent: true,
        clearFormatting: false,
      });

      return {
        success: true,
        sourceRange: `${String.fromCharCode(64 + args.sourceStartCol)}${args.sourceStartRow}:${String.fromCharCode(64 + args.sourceEndCol)}${args.sourceEndRow}`,
        destRange: `${String.fromCharCode(64 + args.destStartCol)}${args.destStartRow}`,
      };
    },

    reorganize_data: async (args: {
      operation: string;
      sourceRange: {
        startRow: number;
        startCol: number;
        endRow: number;
        endCol: number;
      };
      targetRange?: {
        startRow: number;
        startCol: number;
      };
      options?: Record<string, unknown>;
    }) => {
      const sourceData = operations.getRangeValues(
        args.sourceRange.startRow - 1,
        args.sourceRange.startCol - 1,
        args.sourceRange.endRow - 1,
        args.sourceRange.endCol - 1
      );

      let reorganizedData = sourceData;

      switch (args.operation) {
        case 'pivot':
        case 'transpose_blocks':
          // Transpose the data
          if (sourceData.length > 0 && sourceData[0].length > 0) {
            reorganizedData = sourceData[0].map((_, colIndex) =>
              sourceData.map(row => row[colIndex])
            );
          }
          break;
        case 'group_by_column':
          // Group by a specific column
          if (args.options?.groupByColumn) {
            const colIndex = Number(args.options.groupByColumn) - args.sourceRange.startCol;
            const grouped = new Map<string, string[][]>();
            sourceData.forEach(row => {
              const key = String(row[colIndex] || '');
              if (!grouped.has(key)) {
                grouped.set(key, []);
              }
              const rowAsStrings = row.map(cell => String(cell || ''));
              grouped.get(key)!.push(rowAsStrings);
            });
            reorganizedData = Array.from(grouped.values()).flat();
          }
          break;
        case 'consolidate':
          // Remove empty rows
          reorganizedData = sourceData.filter(row => 
            row.some(cell => cell !== '' && cell !== null)
          );
          break;
        case 'distribute':
          // Spread data across more cells
          reorganizedData = [];
          sourceData.forEach(row => {
            reorganizedData.push(row);
            reorganizedData.push(new Array(row.length).fill('')); // Add empty row between
          });
          break;
      }

      const targetRow = args.targetRange?.startRow || args.sourceRange.startRow;
      const targetCol = args.targetRange?.startCol || args.sourceRange.startCol;

      operations.setRangeValues(
        targetRow - 1,
        targetCol - 1,
        reorganizedData
      );

      return {
        success: true,
        operation: args.operation,
        targetRange: `${String.fromCharCode(64 + targetCol)}${targetRow}`,
      };
    },

    swap_cells: async (args: {
      range1StartRow: number;
      range1StartCol: number;
      range1EndRow: number;
      range1EndCol: number;
      range2StartRow: number;
      range2StartCol: number;
      range2EndRow: number;
      range2EndCol: number;
    }) => {
      // Get data from both ranges
      const range1Data = operations.getRangeValues(
        args.range1StartRow - 1,
        args.range1StartCol - 1,
        args.range1EndRow - 1,
        args.range1EndCol - 1
      );

      const range2Data = operations.getRangeValues(
        args.range2StartRow - 1,
        args.range2StartCol - 1,
        args.range2EndRow - 1,
        args.range2EndCol - 1
      );

      // Swap the data
      operations.setRangeValues(
        args.range1StartRow - 1,
        args.range1StartCol - 1,
        range2Data
      );

      operations.setRangeValues(
        args.range2StartRow - 1,
        args.range2StartCol - 1,
        range1Data
      );

      return {
        success: true,
        range1: `${String.fromCharCode(64 + args.range1StartCol)}${args.range1StartRow}:${String.fromCharCode(64 + args.range1EndCol)}${args.range1EndRow}`,
        range2: `${String.fromCharCode(64 + args.range2StartCol)}${args.range2StartRow}:${String.fromCharCode(64 + args.range2EndCol)}${args.range2EndRow}`,
      };
    },

    transform_data: async (args: {
      transformation: string;
      sourceRange: {
        startRow: number;
        startCol: number;
        endRow: number;
        endCol: number;
      };
      options?: Record<string, unknown>;
    }) => {
      const sourceData = operations.getRangeValues(
        args.sourceRange.startRow - 1,
        args.sourceRange.startCol - 1,
        args.sourceRange.endRow - 1,
        args.sourceRange.endCol - 1
      );

      let transformedData = sourceData;

      switch (args.transformation) {
        case 'split_column':
          if (args.options && args.options.delimiter) {
            const delimiter = String(args.options.delimiter);
            transformedData = sourceData.map(row => {
              const newRow: string[] = [];
              row.forEach(cell => {
                const parts = String(cell).split(delimiter);
                newRow.push(...parts);
              });
              return newRow;
            });
          }
          break;
        case 'combine_columns':
          if (args.options && args.options.targetColumns && args.options.joinSeparator) {
            const targetColumns = args.options.targetColumns as number[];
            const joinSeparator = String(args.options.joinSeparator);
            transformedData = sourceData.map(row => {
              const combined = targetColumns
                .map(col => String(row[col - args.sourceRange.startCol] || ''))
                .join(joinSeparator);
              return [combined];
            });
          }
          break;
        case 'normalize':
          // Convert to lowercase and trim
          transformedData = sourceData.map(row =>
            row.map(cell => String(cell).toLowerCase().trim())
          );
          break;
        case 'extract_pattern':
          if (args.options && args.options.pattern) {
            const pattern = String(args.options.pattern);
            const regex = new RegExp(pattern, 'g');
            transformedData = sourceData.map(row =>
              row.map(cell => {
                const matches = String(cell).match(regex);
                return matches ? matches.join(', ') : '';
              })
            );
          }
          break;
      }

      operations.setRangeValues(
        args.sourceRange.startRow - 1,
        args.sourceRange.startCol - 1,
        transformedData
      );

      return {
        success: true,
        transformation: args.transformation,
        range: `${String.fromCharCode(64 + args.sourceRange.startCol)}${args.sourceRange.startRow}:${String.fromCharCode(64 + args.sourceRange.endCol)}${args.sourceRange.endRow}`,
      };
    },

    create_pattern: async (args: {
      pattern: string;
      startRow: number;
      startCol: number;
      rows: number;
      cols: number;
      colors?: string[];
      values?: string[];
    }) => {
      const colors = args.colors || ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7'];
      const values = args.values || [];

      for (let r = 0; r < args.rows; r++) {
        for (let c = 0; c < args.cols; c++) {
          const row = args.startRow + r - 1;
          const col = args.startCol + c - 1;

          let colorIndex = 0;
          let value = '';

          switch (args.pattern) {
            case 'checkerboard':
              colorIndex = (r + c) % 2;
              break;
            case 'gradient':
              colorIndex = Math.floor((r / args.rows) * colors.length);
              break;
            case 'spiral':
              // Simple spiral pattern
              colorIndex = (r + c) % colors.length;
              break;
            case 'zigzag':
              colorIndex = r % 2 === 0 ? c % colors.length : (args.cols - c - 1) % colors.length;
              break;
            case 'diamond': {
              const centerR = Math.floor(args.rows / 2);
              const centerC = Math.floor(args.cols / 2);
              const distance = Math.abs(r - centerR) + Math.abs(c - centerC);
              colorIndex = distance % colors.length;
              break;
            }
            case 'wave':
              colorIndex = Math.floor(Math.sin((c / args.cols) * Math.PI * 2) * colors.length / 2 + colors.length / 2);
              break;
            default:
              colorIndex = (r * args.cols + c) % colors.length;
          }

          if (values.length > 0) {
            value = values[(r * args.cols + c) % values.length];
            operations.setCellValue(row, col, value);
          }

          operations.setCellFormatting({
            startRow: row,
            startCol: col,
            endRow: row,
            endCol: col,
            backgroundColor: colors[Math.min(colorIndex, colors.length - 1)],
          });
        }
      }

      return {
        success: true,
        pattern: args.pattern,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.startCol + args.cols - 1)}${args.startRow + args.rows - 1}`,
      };
    },

    // Analysis operations
    calculate_sum: async (args: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      const sum = operations.calculateSum({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
      });
      return {
        result: sum,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
      };
    },

    calculate_average: async (args: {startRow: number; startCol: number; endRow: number; endCol: number}) => {
      const average = operations.calculateAverage({
        startRow: args.startRow - 1,
        startCol: args.startCol - 1,
        endRow: args.endRow - 1,
        endCol: args.endCol - 1,
      });
      return {
        result: average,
        range: `${String.fromCharCode(64 + args.startCol)}${args.startRow}:${String.fromCharCode(64 + args.endCol)}${args.endRow}`,
      };
    },
  };
}
