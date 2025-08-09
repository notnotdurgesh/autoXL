import Papa from 'papaparse';
import * as XLSX from 'xlsx';

export interface ParsedFile {
  name: string;
  type: string;
  size: number;
  content: any;
  rows?: number;
  columns?: number;
  preview: string;
  format: 'csv' | 'excel' | 'json' | 'text' | 'pdf' | 'word' | 'unknown';
}

export class FileParser {
  private static readonly MAX_PREVIEW_LENGTH = 200;
  private static readonly MAX_ROWS_PREVIEW = 10;

  static async parseFile(file: File): Promise<ParsedFile> {
    const name = file.name;
    const size = file.size;
    const type = file.type || this.getMimeType(name);
    const extension = name.split('.').pop()?.toLowerCase() || '';

    try {
      // CSV Files
      if (type === 'text/csv' || extension === 'csv') {
        return await this.parseCSV(file);
      }

      // Excel Files
      if (
        type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        type === 'application/vnd.ms-excel' ||
        extension === 'xlsx' ||
        extension === 'xls'
      ) {
        return await this.parseExcel(file);
      }

      // JSON Files
      if (type === 'application/json' || extension === 'json') {
        return await this.parseJSON(file);
      }

      // TSV Files
      if (extension === 'tsv' || extension === 'tab') {
        return await this.parseTSV(file);
      }

      // Text Files
      if (type.startsWith('text/') || ['txt', 'log', 'md'].includes(extension)) {
        return await this.parseText(file);
      }

      // Default: treat as text
      return await this.parseText(file);
    } catch (error) {
      console.error('Error parsing file:', error);
      return {
        name,
        type,
        size,
        content: null,
        preview: 'Error parsing file',
        format: 'unknown'
      };
    }
  }

  private static async parseCSV(file: File): Promise<ParsedFile> {
    return new Promise((resolve) => {
      Papa.parse(file, {
        complete: (result) => {
          const data = result.data as any[][];
          const rows = data.length;
          const columns = data[0]?.length || 0;
          
          // Create preview
          const previewRows = data.slice(0, this.MAX_ROWS_PREVIEW);
          const preview = `CSV: ${rows} rows × ${columns} columns\nHeaders: ${data[0]?.slice(0, 5).join(', ')}${data[0]?.length > 5 ? '...' : ''}`;

          resolve({
            name: file.name,
            type: file.type,
            size: file.size,
            content: data,
            rows,
            columns,
            preview,
            format: 'csv'
          });
        },
        error: (error) => {
          resolve({
            name: file.name,
            type: file.type,
            size: file.size,
            content: null,
            preview: `Error parsing CSV: ${error.message}`,
            format: 'csv'
          });
        },
        header: false,
        skipEmptyLines: true,
        dynamicTyping: true
      });
    });
  }

  private static async parseExcel(file: File): Promise<ParsedFile> {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    const rows = data.length;
    const columns = data[0]?.length || 0;
    
    // Create preview
    const preview = `Excel: ${workbook.SheetNames.length} sheet(s)\n${sheetName}: ${rows} rows × ${columns} columns`;

    return {
      name: file.name,
      type: file.type,
      size: file.size,
      content: data,
      rows,
      columns,
      preview,
      format: 'excel'
    };
  }

  private static async parseJSON(file: File): Promise<ParsedFile> {
    const text = await file.text();
    const data = JSON.parse(text);
    
    let preview = 'JSON: ';
    if (Array.isArray(data)) {
      preview += `Array with ${data.length} items`;
      if (data.length > 0 && typeof data[0] === 'object') {
        const keys = Object.keys(data[0]);
        preview += `\nKeys: ${keys.slice(0, 5).join(', ')}${keys.length > 5 ? '...' : ''}`;
      }
    } else if (typeof data === 'object') {
      const keys = Object.keys(data);
      preview += `Object with ${keys.length} properties\nKeys: ${keys.slice(0, 5).join(', ')}${keys.length > 5 ? '...' : ''}`;
    } else {
      preview += typeof data;
    }

    return {
      name: file.name,
      type: file.type,
      size: file.size,
      content: data,
      preview,
      format: 'json'
    };
  }

  private static async parseTSV(file: File): Promise<ParsedFile> {
    return new Promise((resolve) => {
      Papa.parse(file, {
        complete: (result) => {
          const data = result.data as any[][];
          const rows = data.length;
          const columns = data[0]?.length || 0;
          
          const preview = `TSV: ${rows} rows × ${columns} columns`;

          resolve({
            name: file.name,
            type: file.type,
            size: file.size,
            content: data,
            rows,
            columns,
            preview,
            format: 'csv'
          });
        },
        delimiter: '\t',
        header: false,
        skipEmptyLines: true,
        dynamicTyping: true
      });
    });
  }

  private static async parseText(file: File): Promise<ParsedFile> {
    const text = await file.text();
    const lines = text.split('\n');
    const preview = `Text: ${lines.length} lines\n${text.substring(0, this.MAX_PREVIEW_LENGTH)}${text.length > this.MAX_PREVIEW_LENGTH ? '...' : ''}`;

    return {
      name: file.name,
      type: file.type,
      size: file.size,
      content: text,
      preview,
      format: 'text'
    };
  }

  private static getMimeType(filename: string): string {
    const ext = filename.split('.').pop()?.toLowerCase();
    const mimeTypes: Record<string, string> = {
      'csv': 'text/csv',
      'tsv': 'text/tab-separated-values',
      'txt': 'text/plain',
      'json': 'application/json',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'xls': 'application/vnd.ms-excel',
      'xml': 'text/xml',
      'html': 'text/html',
      'md': 'text/markdown',
      'log': 'text/plain'
    };
    return mimeTypes[ext || ''] || 'application/octet-stream';
  }

  static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
}
