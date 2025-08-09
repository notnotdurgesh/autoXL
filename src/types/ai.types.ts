import type { CellFormatting } from './cell.types';

// Cell attachment type
export interface CellAttachment {
  id: string;
  range: string;
  values: string | number | (string | number | null)[][] | null;
}

// File attachment type
export interface FileAttachment {
  id: string;
  name: string;
  type: string;
  size: number;
  content: string | (string | number | null)[][] | null;
  preview?: string;
}

// Message types for chat history
export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system' | 'function';
  content: string;
  timestamp: Date;
  cellAttachments?: CellAttachment[];
  fileAttachments?: FileAttachment[];
  functionCall?: FunctionCall;
  functionResponse?: FunctionResponse;
}

// Function call and response types
export interface FunctionCall {
  name: string;
  args: Record<string, unknown>;
}

export interface FunctionResponse {
  name: string;
  result: unknown;
  error?: string;
}

// Cell reference types
export interface CellReference {
  row: number;
  col: number;
}

export interface CellRange {
  start: CellReference;
  end: CellReference;
}

// Spreadsheet operation types
export interface CellUpdate {
  cell: CellReference;
  value?: string | number | null;
  formula?: string;
  formatting?: CellFormatting;
}

export interface RangeUpdate {
  range: CellRange;
  values?: (string | number | null)[][];
  formatting?: CellFormatting;
}

// Chart types
export interface ChartConfig {
  type: 'line' | 'bar' | 'pie' | 'scatter' | 'area';
  dataRange: CellRange;
  title?: string;
  xAxisLabel?: string;
  yAxisLabel?: string;
}

// Filter and sort types
export interface FilterConfig {
  range: CellRange;
  column: number;
  criteria: 'equals' | 'not_equals' | 'contains' | 'not_contains' | 'greater_than' | 'less_than';
  value: string | number;
}

export interface SortConfig {
  range: CellRange;
  column: number;
  order: 'asc' | 'desc';
}

// Sheet operation results
export interface OperationResult {
  success: boolean;
  message?: string;
  data?: unknown;
  error?: string;
}

// AI configuration
export interface AIConfig {
  apiKey: string;
  model?: string;
  temperature?: number;
  maxTokens?: number;
}

// Chat state
export interface ChatState {
  messages: ChatMessage[];
  isLoading: boolean;
  error: string | null;
}
