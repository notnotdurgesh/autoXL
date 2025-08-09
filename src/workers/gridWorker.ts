/*
  Grid worker: computes heavy operations (paste/fill) off the main thread.
  Messages follow the shape: { requestId, type, payload }
  Responses: { requestId, ok: true, result } or { requestId, ok: false, error }
*/

export type CellFormatting = {
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
};

type CellPayload = { value: string | number | null; formatting?: CellFormatting };

type ComputePasteRequest = {
  startRow: number;
  startCol: number;
  data: CellPayload[][];
  isCut?: boolean;
  cutRange?: { startRow: number; endRow: number; startCol: number; endCol: number } | null;
};

type ComputeFillRequest = {
  sourceValues: (string | number)[];
  sourceRowCount: number;
  sourceColCount: number;
  sourceMinRow: number;
  sourceMinCol: number;
  targetMinRow: number;
  targetMaxRow: number;
  targetMinCol: number;
  targetMaxCol: number;
};

type Update = { row: number; col: number; value: string | number; formatting?: CellFormatting };
type Clear = { row: number; col: number };

function computePaste(req: ComputePasteRequest): { updates: Update[]; clears: Clear[] } {
  const { startRow, startCol, data, isCut, cutRange } = req;
  const updates: Update[] = [];
  const clears: Clear[] = [];

  const pasteRows = data.length;
  const pasteCols = data[0]?.length || 0;

  // Apply pasted data
  for (let r = 0; r < pasteRows; r++) {
    for (let c = 0; c < pasteCols; c++) {
      const payload = data[r][c];
      const targetRow = startRow + r;
      const targetCol = startCol + c;
      const value = payload?.value ?? '';
      const formatting = payload?.formatting;

      if ((value === '' || value === null) && (!formatting || Object.keys(formatting).length === 0)) {
        // No value and no formatting → clear
        clears.push({ row: targetRow, col: targetCol });
      } else {
        updates.push({ row: targetRow, col: targetCol, value: value as string | number, formatting });
      }
    }
  }

  // Handle cut cleanup (exclude overlap with pasted area)
  if (isCut && cutRange) {
    const minRow = Math.min(cutRange.startRow, cutRange.endRow);
    const maxRow = Math.max(cutRange.startRow, cutRange.endRow);
    const minCol = Math.min(cutRange.startCol, cutRange.endCol);
    const maxCol = Math.max(cutRange.startCol, cutRange.endCol);

    for (let row = minRow; row <= maxRow; row++) {
      for (let col = minCol; col <= maxCol; col++) {
        const inPaste = row >= startRow && row < startRow + pasteRows && col >= startCol && col < startCol + pasteCols;
        if (!inPaste) {
          clears.push({ row, col });
        }
      }
    }
  }

  return { updates, clears };
}

function computeFill(req: ComputeFillRequest): { updates: Update[] } {
  const {
    sourceValues,
    sourceRowCount,
    sourceColCount,
    sourceMinRow,
    sourceMinCol,
    targetMinRow,
    targetMaxRow,
    targetMinCol,
    targetMaxCol,
  } = req;

  const updates: Update[] = [];
  for (let targetRow = targetMinRow; targetRow <= targetMaxRow; targetRow++) {
    for (let targetCol = targetMinCol; targetCol <= targetMaxCol; targetCol++) {
      // Skip cells that are part of the source range
      if (
        targetRow >= sourceMinRow &&
        targetRow <= sourceMinRow + (sourceRowCount - 1) &&
        targetCol >= sourceMinCol &&
        targetCol <= sourceMinCol + (sourceColCount - 1)
      ) {
        continue;
      }
      const relativeRow = (targetRow - sourceMinRow) % sourceRowCount;
      const relativeCol = (targetCol - sourceMinCol) % sourceColCount;
      const sourceIndex = relativeRow * sourceColCount + relativeCol;
      if (sourceIndex >= 0 && sourceIndex < sourceValues.length) {
        const value = sourceValues[sourceIndex] ?? '';
        updates.push({ row: targetRow, col: targetCol, value: value as string | number });
      }
    }
  }
  return { updates };
}

self.addEventListener('message', (e: MessageEvent) => {
  const { requestId, type, payload } = e.data as { requestId: string; type: string; payload: unknown };
  try {
    if (type === 'computePaste') {
      const result = computePaste(payload as ComputePasteRequest);
      (self as unknown as Worker).postMessage({ requestId, ok: true, result });
      return;
    }
    if (type === 'computeFill') {
      const result = computeFill(payload as ComputeFillRequest);
      (self as unknown as Worker).postMessage({ requestId, ok: true, result });
      return;
    }
    (self as unknown as Worker).postMessage({ requestId, ok: false, error: 'Unknown worker message type' });
  } catch (err) {
    const message = (err as Error)?.message || 'Worker error';
    (self as unknown as Worker).postMessage({ requestId, ok: false, error: message });
  }
});


