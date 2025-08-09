// Cell-related type definitions
export interface CellFormatting {
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

export interface CellData {
  value: string | number | null;
  isEditing?: boolean;
  formatting?: CellFormatting;
  displayValue?: string; // For formatted display (currency, percentage, etc.)
}

export interface CellPosition {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface SheetData {
  [row: number]: {
    [col: number]: CellData;
  };
}

// Excel-like constants
export const EXCEL_CONSTANTS = {
  MAX_ROWS: 1048576,
  MAX_COLS: 16384,
  MIN_VISIBLE_ROWS: 50,
  MIN_VISIBLE_COLS: 26,
  EXPANSION_BUFFER: 10,
  DEFAULT_COLUMN_WIDTH: 80,
  DEFAULT_ROW_HEIGHT: 20,
  MIN_COLUMN_WIDTH: 30,
  MIN_ROW_HEIGHT: 15,
} as const;

// Virtual scrolling constants
export const VIRTUAL_SCROLL_CONSTANTS = {
  MAX_RENDERED_ROWS: 100,
  MAX_RENDERED_COLS: 50,
  RENDER_BUFFER: 5,
  PERFORMANCE_THRESHOLD: 10000,
} as const;
