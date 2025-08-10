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
  format: 'csv' | 'excel' | 'json' | 'text' | 'pdf' | 'word' | 'image' | 'unknown';
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

      // PDF Files
      if (type === 'application/pdf' || extension === 'pdf') {
        return await this.parsePDF(file);
      }

      // Word (DOCX) Files
      if (
        type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        extension === 'docx'
      ) {
        return await this.parseWord(file);
      }

      // TSV Files
      if (extension === 'tsv' || extension === 'tab') {
        return await this.parseTSV(file);
      }

      // Text Files
      if (type.startsWith('text/') || ['txt', 'log', 'md'].includes(extension)) {
        return await this.parseText(file);
      }

      // Images (PNG, JPEG, WEBP, GIF, BMP)
      if (type.startsWith('image/') || ['png','jpg','jpeg','webp','gif','bmp'].includes(extension)) {
        return await this.parseImage(file);
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

  // Best-effort PDF text extraction for preview (first ~1500 chars)
  // Falls back gracefully if PDF.js fails (e.g., worker issues)
  private static async parsePDF(file: File): Promise<ParsedFile> {
    const arrayBuffer = await file.arrayBuffer();
    let preview = 'PDF document';
    let textContent = '';
    try {
      // Lazy load pdfjs-dist to avoid bundling cost until needed
      // Use modern build; bundlers handle worker if configured. Graceful catch otherwise.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pdfjs: any = await import('pdfjs-dist/build/pdf');
      // Attempt to set worker if available in ESM build; ignore if not
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worker = await import('pdfjs-dist/build/pdf.worker.mjs');
        if (worker && pdfjs && pdfjs.GlobalWorkerOptions) {
          pdfjs.GlobalWorkerOptions.workerSrc = (worker as unknown as { default: string }).default || (worker as unknown as string);
        }
      } catch {
        // Ignore worker setup failures; pdfjs may still work in-thread in some envs
      }

      const loadingTask = pdfjs.getDocument({ data: arrayBuffer });
      const doc = await loadingTask.promise;
      preview = `PDF: ${doc.numPages} page${doc.numPages > 1 ? 's' : ''}`;
      // Extract text from first 1-2 pages for a concise preview
      const pagesToRead = Math.min(2, doc.numPages);
      for (let i = 1; i <= pagesToRead; i++) {
        const page = await doc.getPage(i);
        const tc = await page.getTextContent();
        const pageText = tc.items.map((it: { str?: string }) => it.str || '').join(' ');
        textContent += pageText + (i < pagesToRead ? '\n\n' : '');
      }
      if (textContent.length > 1500) {
        textContent = textContent.substring(0, 1500) + '...';
      }
    } catch {
      // Ignore and fall back to generic info
      preview = `PDF (${(file.size / 1024).toFixed(1)} KB)`;
      textContent = '';
    }

    return {
      name: file.name,
      type: file.type || 'application/pdf',
      size: file.size,
      content: textContent || null,
      preview,
      format: 'pdf',
    };
  }

  // DOCX preview using mammoth (first ~1500 chars)
  private static async parseWord(file: File): Promise<ParsedFile> {
    let preview = 'Word document';
    let bodyText = '';
    try {
      const arrayBuffer = await file.arrayBuffer();
      // Lazy import to reduce initial bundle
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const mammoth = await import('mammoth');
      const result = await (mammoth as unknown as { extractRawText: (opts: { arrayBuffer: ArrayBuffer }) => Promise<{ value: string }> }).extractRawText({ arrayBuffer });
      bodyText = result.value || '';
      if (bodyText.length > 1500) bodyText = bodyText.substring(0, 1500) + '...';
      preview = 'Word (DOCX)';
    } catch {
      preview = `Word (DOCX) (${(file.size / 1024).toFixed(1)} KB)`;
    }

    return {
      name: file.name,
      type: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      size: file.size,
      content: bodyText || null,
      preview,
      format: 'word',
    };
  }

  private static async parseImage(file: File): Promise<ParsedFile> {
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });

    // Try to obtain dimensions for preview
    const dimensions = await new Promise<{width: number; height: number} | null>((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ width: img.width, height: img.height });
      img.onerror = () => resolve(null);
      img.src = dataUrl;
    });

    const preview = dimensions
      ? `Image: ${dimensions.width}×${dimensions.height} (${file.type || 'image'})`
      : `Image: ${file.type || 'image'}`;

    return {
      name: file.name,
      type: file.type || 'image',
      size: file.size,
      content: dataUrl, // Data URL for inlineData; caller extracts base64
      preview,
      format: 'image'
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
