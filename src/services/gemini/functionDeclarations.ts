import type { FunctionDeclaration } from '@google/generative-ai';
import { SchemaType } from '@google/generative-ai';

// Cell value operations
export const getCellValueDeclaration: FunctionDeclaration = {
  name: 'get_cell_value',
  description: 'Gets the value of a specific cell in the spreadsheet',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number (1-based indexing)',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number (1-based indexing, A=1, B=2, etc.)',
      },
    },
    required: ['row', 'col'],
  },
};

export const getCellDataDeclaration: FunctionDeclaration = {
  name: 'get_cell_data',
  description: 'Get the value AND formatting of a single cell. Use this when you need to know if a cell is bold, colored, etc.',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number (1-based indexing)',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number (1-based indexing, A=1, B=2, etc.)',
      },
    },
    required: ['row', 'col'],
  },
};

export const setCellValueDeclaration: FunctionDeclaration = {
  name: 'set_cell_value',
  description: 'Sets the value of a specific cell in the spreadsheet',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number (1-based indexing)',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number (1-based indexing, A=1, B=2, etc.)',
      },
      value: {
        type: SchemaType.STRING,
        description: 'The value to set (can be text, number, or formula starting with =)',
      },
    },
    required: ['row', 'col', 'value'],
  },
};

export const getRangeValuesDeclaration: FunctionDeclaration = {
  name: 'get_range_values',
  description: 'Gets values from a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row number',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column number',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row number',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column number',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol'],
  },
};

export const setRangeValuesDeclaration: FunctionDeclaration = {
  name: 'set_range_values',
  description: 'Sets values for a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row number',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column number',
      },
      values: {
        type: SchemaType.ARRAY,
        description: '2D array of values to set',
        items: {
          type: SchemaType.ARRAY,
          items: {
            type: SchemaType.STRING,
          },
        },
      },
    },
    required: ['startRow', 'startCol', 'values'],
  },
};

// Formatting operations
export const setCellFormattingDeclaration: FunctionDeclaration = {
  name: 'set_cell_formatting',
  description: 'Apply formatting to cells (bold, italic, color, alignment, etc.)',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row number',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column number',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row number (optional, defaults to startRow)',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column number (optional, defaults to startCol)',
      },
      bold: {
        type: SchemaType.BOOLEAN,
        description: 'Make text bold',
      },
      italic: {
        type: SchemaType.BOOLEAN,
        description: 'Make text italic',
      },
      underline: {
        type: SchemaType.BOOLEAN,
        description: 'Underline text',
      },
      strikethrough: {
        type: SchemaType.BOOLEAN,
        description: 'Strikethrough text',
      },
      fontSize: {
        type: SchemaType.NUMBER,
        description: 'Font size in pixels',
      },
      fontFamily: {
        type: SchemaType.STRING,
        description: 'Font family (Arial, Calibri, Times New Roman, etc.)',
      },
      textColor: {
        type: SchemaType.STRING,
        description: 'Text color in hex format (e.g., #FF0000)',
      },
      backgroundColor: {
        type: SchemaType.STRING,
        description: 'Background color in hex format',
      },
      textAlign: {
        type: SchemaType.STRING,
        description: 'Text alignment (left, center, right, or justify)',
      },
      verticalAlign: {
        type: SchemaType.STRING,
        description: 'Vertical alignment (top, middle, or bottom)',
      },
      numberFormat: {
        type: SchemaType.STRING,
        description: 'Number format type (general, number, currency, percentage, date, or time)',
      },
      decimalPlaces: {
        type: SchemaType.NUMBER,
        description: 'Number of decimal places for number formatting',
      },
    },
    required: ['startRow', 'startCol'],
  },
};

// Formula operations
export const setCellFormulaDeclaration: FunctionDeclaration = {
  name: 'set_cell_formula',
  description: 'Set a formula in a cell (SUM, AVERAGE, COUNT, IF, VLOOKUP, etc.)',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number',
      },
      formula: {
        type: SchemaType.STRING,
        description: 'Formula to set (must start with =, e.g., =SUM(A1:A10))',
      },
    },
    required: ['row', 'col', 'formula'],
  },
};

export const autoFillDeclaration: FunctionDeclaration = {
  name: 'auto_fill',
  description: 'Auto-fill cells based on a pattern (like dragging cell corner)',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      sourceRow: {
        type: SchemaType.NUMBER,
        description: 'Source cell row',
      },
      sourceCol: {
        type: SchemaType.NUMBER,
        description: 'Source cell column',
      },
      targetEndRow: {
        type: SchemaType.NUMBER,
        description: 'Target end row',
      },
      targetEndCol: {
        type: SchemaType.NUMBER,
        description: 'Target end column',
      },
      fillType: {
        type: SchemaType.STRING,
        description: 'Type of fill (copy, series, or format_only)',
      },
    },
    required: ['sourceRow', 'sourceCol', 'targetEndRow', 'targetEndCol'],
  },
};

// Row and column operations
export const insertRowsDeclaration: FunctionDeclaration = {
  name: 'insert_rows',
  description: 'Insert new rows at a specific position',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      position: {
        type: SchemaType.NUMBER,
        description: 'Position to insert rows (1-based)',
      },
      count: {
        type: SchemaType.NUMBER,
        description: 'Number of rows to insert',
      },
    },
    required: ['position', 'count'],
  },
};

export const insertColumnsDeclaration: FunctionDeclaration = {
  name: 'insert_columns',
  description: 'Insert new columns at a specific position',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      position: {
        type: SchemaType.NUMBER,
        description: 'Position to insert columns (1-based)',
      },
      count: {
        type: SchemaType.NUMBER,
        description: 'Number of columns to insert',
      },
    },
    required: ['position', 'count'],
  },
};

export const deleteRowsDeclaration: FunctionDeclaration = {
  name: 'delete_rows',
  description: 'Delete rows from the spreadsheet',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row to delete',
      },
      count: {
        type: SchemaType.NUMBER,
        description: 'Number of rows to delete',
      },
    },
    required: ['startRow', 'count'],
  },
};

export const deleteColumnsDeclaration: FunctionDeclaration = {
  name: 'delete_columns',
  description: 'Delete columns from the spreadsheet',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column to delete',
      },
      count: {
        type: SchemaType.NUMBER,
        description: 'Number of columns to delete',
      },
    },
    required: ['startCol', 'count'],
  },
};

// Data operations
export const sortRangeDeclaration: FunctionDeclaration = {
  name: 'sort_range',
  description: 'Sort a range of cells by a specific column',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row of range',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column of range',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row of range',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column of range',
      },
      sortColumn: {
        type: SchemaType.NUMBER,
        description: 'Column to sort by (within the range)',
      },
      order: {
        type: SchemaType.STRING,
        description: 'Sort order (asc or desc)',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol', 'sortColumn', 'order'],
  },
};

export const filterRangeDeclaration: FunctionDeclaration = {
  name: 'filter_range',
  description: 'Apply a filter to a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row of range',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column of range',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row of range',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column of range',
      },
      filterColumn: {
        type: SchemaType.NUMBER,
        description: 'Column to filter by',
      },
      criteria: {
        type: SchemaType.STRING,
        description: 'Filter criteria (equals, not_equals, contains, not_contains, greater_than, or less_than)',
      },
      value: {
        type: SchemaType.STRING,
        description: 'Value to filter by',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol', 'filterColumn', 'criteria', 'value'],
  },
};

export const findReplaceDeclaration: FunctionDeclaration = {
  name: 'find_replace',
  description: 'Find and replace values in the spreadsheet',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      findValue: {
        type: SchemaType.STRING,
        description: 'Value to find',
      },
      replaceValue: {
        type: SchemaType.STRING,
        description: 'Value to replace with',
      },
      matchCase: {
        type: SchemaType.BOOLEAN,
        description: 'Match case sensitivity',
      },
      matchEntireCell: {
        type: SchemaType.BOOLEAN,
        description: 'Match entire cell content only',
      },
      rangeStartRow: {
        type: SchemaType.NUMBER,
        description: 'Optional: Start row of search range',
      },
      rangeStartCol: {
        type: SchemaType.NUMBER,
        description: 'Optional: Start column of search range',
      },
      rangeEndRow: {
        type: SchemaType.NUMBER,
        description: 'Optional: End row of search range',
      },
      rangeEndCol: {
        type: SchemaType.NUMBER,
        description: 'Optional: End column of search range',
      },
    },
    required: ['findValue', 'replaceValue'],
  },
};

// Chart operations
export const createChartDeclaration: FunctionDeclaration = {
  name: 'create_chart',
  description: 'Create a chart from spreadsheet data',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      chartType: {
        type: SchemaType.STRING,
        description: 'Type of chart (line, bar, pie, scatter, area, or column)',
      },
      dataStartRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row of data range',
      },
      dataStartCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column of data range',
      },
      dataEndRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row of data range',
      },
      dataEndCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column of data range',
      },
      title: {
        type: SchemaType.STRING,
        description: 'Chart title',
      },
      xAxisLabel: {
        type: SchemaType.STRING,
        description: 'X-axis label',
      },
      yAxisLabel: {
        type: SchemaType.STRING,
        description: 'Y-axis label',
      },
    },
    required: ['chartType', 'dataStartRow', 'dataStartCol', 'dataEndRow', 'dataEndCol'],
  },
};

// Utility operations
export const findCellDeclaration: FunctionDeclaration = {
  name: 'find_cell',
  description: 'Find a cell by its content. Use this to locate cells containing specific text before deleting or modifying them.',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      searchText: {
        type: SchemaType.STRING,
        description: 'Text to search for in cells',
      },
      matchCase: {
        type: SchemaType.BOOLEAN,
        description: 'Whether to match case (default: false)',
        nullable: true,
      },
      matchEntireCell: {
        type: SchemaType.BOOLEAN,
        description: 'Whether to match entire cell content only (default: false)',
        nullable: true,
      },
    },
    required: ['searchText'],
  },
};

export const clearCellDeclaration: FunctionDeclaration = {
  name: 'clear_cell',
  description: 'Clear/delete the content of a single cell. Use this for deleting individual cells.',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number of the cell to clear',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number of the cell to clear',
      },
    },
    required: ['row', 'col'],
  },
};

export const deleteCellDeclaration: FunctionDeclaration = {
  name: 'delete_cell',
  description: 'Delete the content of a single cell (alias for clear_cell). Use this when user says "delete" instead of "clear".',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number of the cell to delete',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number of the cell to delete',
      },
    },
    required: ['row', 'col'],
  },
};

export const clearRangeDeclaration: FunctionDeclaration = {
  name: 'clear_range',
  description: 'Clear content and/or formatting from a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column',
      },
      clearContent: {
        type: SchemaType.BOOLEAN,
        description: 'Clear cell content',
      },
      clearFormatting: {
        type: SchemaType.BOOLEAN,
        description: 'Clear cell formatting',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol'],
  },
};

export const mergeRangeDeclaration: FunctionDeclaration = {
  name: 'merge_range',
  description: 'Merge a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol'],
  },
};

export const unmergeRangeDeclaration: FunctionDeclaration = {
  name: 'unmerge_range',
  description: 'Unmerge previously merged cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row of any cell in the merged range',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column of any cell in the merged range',
      },
    },
    required: ['row', 'col'],
  },
};

export const addHyperlinkDeclaration: FunctionDeclaration = {
  name: 'add_hyperlink',
  description: 'Add a hyperlink to a cell',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      row: {
        type: SchemaType.NUMBER,
        description: 'Row number',
      },
      col: {
        type: SchemaType.NUMBER,
        description: 'Column number',
      },
      url: {
        type: SchemaType.STRING,
        description: 'URL to link to',
      },
      displayText: {
        type: SchemaType.STRING,
        description: 'Text to display in the cell',
      },
    },
    required: ['row', 'col', 'url'],
  },
};

// Advanced Cell Operations
export const copyCellsDeclaration: FunctionDeclaration = {
  name: 'copy_cells',
  description: 'Copy cells from source to destination with optional transpose, allowing creative reorganization and duplication of data',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      sourceStartRow: {
        type: SchemaType.NUMBER,
        description: 'Source starting row',
      },
      sourceStartCol: {
        type: SchemaType.NUMBER,
        description: 'Source starting column',
      },
      sourceEndRow: {
        type: SchemaType.NUMBER,
        description: 'Source ending row',
      },
      sourceEndCol: {
        type: SchemaType.NUMBER,
        description: 'Source ending column',
      },
      destStartRow: {
        type: SchemaType.NUMBER,
        description: 'Destination starting row',
      },
      destStartCol: {
        type: SchemaType.NUMBER,
        description: 'Destination starting column',
      },
      transpose: {
        type: SchemaType.BOOLEAN,
        description: 'Transpose the data (swap rows and columns)',
      },
      includeFormatting: {
        type: SchemaType.BOOLEAN,
        description: 'Copy formatting along with values',
      },
    },
    required: ['sourceStartRow', 'sourceStartCol', 'sourceEndRow', 'sourceEndCol', 'destStartRow', 'destStartCol'],
  },
};

export const moveCellsDeclaration: FunctionDeclaration = {
  name: 'move_cells',
  description: 'Move cells from one location to another, clearing the source location. Perfect for reorganizing and restructuring data',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      sourceStartRow: {
        type: SchemaType.NUMBER,
        description: 'Source starting row',
      },
      sourceStartCol: {
        type: SchemaType.NUMBER,
        description: 'Source starting column',
      },
      sourceEndRow: {
        type: SchemaType.NUMBER,
        description: 'Source ending row',
      },
      sourceEndCol: {
        type: SchemaType.NUMBER,
        description: 'Source ending column',
      },
      destStartRow: {
        type: SchemaType.NUMBER,
        description: 'Destination starting row',
      },
      destStartCol: {
        type: SchemaType.NUMBER,
        description: 'Destination starting column',
      },
      shiftCells: {
        type: SchemaType.BOOLEAN,
        description: 'Shift existing cells at destination to make room',
      },
    },
    required: ['sourceStartRow', 'sourceStartCol', 'sourceEndRow', 'sourceEndCol', 'destStartRow', 'destStartCol'],
  },
};

export const reorganizeDataDeclaration: FunctionDeclaration = {
  name: 'reorganize_data',
  description: 'Completely reorganize and transform data layout - pivot, unpivot, group, split, or any creative data restructuring',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      operation: {
        type: SchemaType.STRING,
        description: 'Type of reorganization: pivot, unpivot, group_by_column, split_by_value, transpose_blocks, consolidate, distribute, or custom',
      },
      sourceRange: {
        type: SchemaType.OBJECT,
        description: 'Source data range',
        properties: {
          startRow: { type: SchemaType.NUMBER },
          startCol: { type: SchemaType.NUMBER },
          endRow: { type: SchemaType.NUMBER },
          endCol: { type: SchemaType.NUMBER },
        },
      },
      targetRange: {
        type: SchemaType.OBJECT,
        description: 'Target location for reorganized data',
        properties: {
          startRow: { type: SchemaType.NUMBER },
          startCol: { type: SchemaType.NUMBER },
        },
      },
      options: {
        type: SchemaType.OBJECT,
        description: 'Additional options for the operation',
        properties: {
          groupByColumn: { type: SchemaType.NUMBER },
          pivotColumn: { type: SchemaType.NUMBER },
          valueColumn: { type: SchemaType.NUMBER },
          preserveFormatting: { type: SchemaType.BOOLEAN },
          sortResult: { type: SchemaType.BOOLEAN },
        },
      },
    },
    required: ['operation', 'sourceRange'],
  },
};

export const swapCellsDeclaration: FunctionDeclaration = {
  name: 'swap_cells',
  description: 'Swap content between two ranges of cells - perfect for reordering data',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      range1StartRow: {
        type: SchemaType.NUMBER,
        description: 'First range starting row',
      },
      range1StartCol: {
        type: SchemaType.NUMBER,
        description: 'First range starting column',
      },
      range1EndRow: {
        type: SchemaType.NUMBER,
        description: 'First range ending row',
      },
      range1EndCol: {
        type: SchemaType.NUMBER,
        description: 'First range ending column',
      },
      range2StartRow: {
        type: SchemaType.NUMBER,
        description: 'Second range starting row',
      },
      range2StartCol: {
        type: SchemaType.NUMBER,
        description: 'Second range starting column',
      },
      range2EndRow: {
        type: SchemaType.NUMBER,
        description: 'Second range ending row',
      },
      range2EndCol: {
        type: SchemaType.NUMBER,
        description: 'Second range ending column',
      },
    },
    required: ['range1StartRow', 'range1StartCol', 'range1EndRow', 'range1EndCol', 'range2StartRow', 'range2StartCol', 'range2EndRow', 'range2EndCol'],
  },
};

export const transformDataDeclaration: FunctionDeclaration = {
  name: 'transform_data',
  description: 'Apply creative transformations to data - split columns, combine rows, extract patterns, normalize, denormalize, or any imaginative data manipulation',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      transformation: {
        type: SchemaType.STRING,
        description: 'Type of transformation: split_column, combine_columns, extract_pattern, normalize, denormalize, flatten_hierarchy, create_hierarchy, or custom',
      },
      sourceRange: {
        type: SchemaType.OBJECT,
        description: 'Source data range',
        properties: {
          startRow: { type: SchemaType.NUMBER },
          startCol: { type: SchemaType.NUMBER },
          endRow: { type: SchemaType.NUMBER },
          endCol: { type: SchemaType.NUMBER },
        },
      },
      options: {
        type: SchemaType.OBJECT,
        description: 'Transformation options',
        properties: {
          delimiter: { type: SchemaType.STRING },
          pattern: { type: SchemaType.STRING },
          joinSeparator: { type: SchemaType.STRING },
          targetColumns: { type: SchemaType.ARRAY, items: { type: SchemaType.NUMBER } },
          keepOriginal: { type: SchemaType.BOOLEAN },
        },
      },
    },
    required: ['transformation', 'sourceRange'],
  },
};

export const createPatternDeclaration: FunctionDeclaration = {
  name: 'create_pattern',
  description: 'Create artistic patterns, layouts, or visual arrangements with data and formatting - be creative!',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      pattern: {
        type: SchemaType.STRING,
        description: 'Pattern type: checkerboard, gradient, spiral, zigzag, diamond, wave, or custom artistic pattern',
      },
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row for pattern',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column for pattern',
      },
      rows: {
        type: SchemaType.NUMBER,
        description: 'Number of rows for pattern',
      },
      cols: {
        type: SchemaType.NUMBER,
        description: 'Number of columns for pattern',
      },
      colors: {
        type: SchemaType.ARRAY,
        description: 'Array of colors to use in pattern',
        items: { type: SchemaType.STRING },
      },
      values: {
        type: SchemaType.ARRAY,
        description: 'Array of values to use in pattern',
        items: { type: SchemaType.STRING },
      },
    },
    required: ['pattern', 'startRow', 'startCol', 'rows', 'cols'],
  },
};

// Analysis operations
export const calculateSumDeclaration: FunctionDeclaration = {
  name: 'calculate_sum',
  description: 'Calculate the sum of a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol'],
  },
};

export const calculateAverageDeclaration: FunctionDeclaration = {
  name: 'calculate_average',
  description: 'Calculate the average of a range of cells',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      startRow: {
        type: SchemaType.NUMBER,
        description: 'Starting row',
      },
      startCol: {
        type: SchemaType.NUMBER,
        description: 'Starting column',
      },
      endRow: {
        type: SchemaType.NUMBER,
        description: 'Ending row',
      },
      endCol: {
        type: SchemaType.NUMBER,
        description: 'Ending column',
      },
    },
    required: ['startRow', 'startCol', 'endRow', 'endCol'],
  },
};

export const scanSheetDeclaration: FunctionDeclaration = {
  name: 'scan_sheet',
  description: 'Scan the sheet to find all data. ALWAYS use this when user asks "what\'s in the sheet" or wants an overview. Can optionally include formatting information.',
  parameters: {
    type: SchemaType.OBJECT,
    properties: {
      maxRows: {
        type: SchemaType.NUMBER,
        description: 'Maximum rows to scan (default 100)',
        nullable: true,
      },
      maxCols: {
        type: SchemaType.NUMBER,
        description: 'Maximum columns to scan (default 26 for A-Z)',
        nullable: true,
      },
      includeFormatting: {
        type: SchemaType.BOOLEAN,
        description: 'Include cell formatting information (bold, colors, etc.) in the response',
        nullable: true,
      },
    },
    required: [],
  },
};

// Export all declarations as an array
export const allFunctionDeclarations = [
  scanSheetDeclaration,
  getCellValueDeclaration,
  getCellDataDeclaration,
  setCellValueDeclaration,
  getRangeValuesDeclaration,
  setRangeValuesDeclaration,
  setCellFormattingDeclaration,
  setCellFormulaDeclaration,
  autoFillDeclaration,
  insertRowsDeclaration,
  insertColumnsDeclaration,
  deleteRowsDeclaration,
  deleteColumnsDeclaration,
  sortRangeDeclaration,
  filterRangeDeclaration,
  findReplaceDeclaration,
  createChartDeclaration,
  findCellDeclaration,
  clearCellDeclaration,
  deleteCellDeclaration,
  clearRangeDeclaration,
  mergeRangeDeclaration,
  unmergeRangeDeclaration,
  addHyperlinkDeclaration,
  copyCellsDeclaration,
  moveCellsDeclaration,
  reorganizeDataDeclaration,
  swapCellsDeclaration,
  transformDataDeclaration,
  createPatternDeclaration,
  calculateSumDeclaration,
  calculateAverageDeclaration,
];
