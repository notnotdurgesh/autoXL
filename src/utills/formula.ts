// Lightweight formula evaluation utilities (MVP)

export type GetA1Value = (a1: string) => string | number | null | undefined;

function colLabelToIndex(label: string): number {
  let idx = 0;
  for (let i = 0; i < label.length; i++) {
    idx = idx * 26 + (label.charCodeAt(i) - 64);
  }
  return idx - 1; // zero-based
}

export function parseA1Address(a1: string): { row: number; col: number } | null {
  const m = /^\s*([A-Z]+)([0-9]+)\s*$/i.exec(a1);
  if (!m) return null;
  const col = colLabelToIndex(m[1].toUpperCase());
  const row = parseInt(m[2], 10) - 1;
  if (row < 0 || col < 0) return null;
  return { row, col };
}

export function expandRange(a1Range: string): string[] {
  const m = /^\s*([A-Z]+[0-9]+)\s*:\s*([A-Z]+[0-9]+)\s*$/i.exec(a1Range);
  if (!m) return [];
  const start = parseA1Address(m[1]);
  const end = parseA1Address(m[2]);
  if (!start || !end) return [];
  const r1 = Math.min(start.row, end.row);
  const r2 = Math.max(start.row, end.row);
  const c1 = Math.min(start.col, end.col);
  const c2 = Math.max(start.col, end.col);
  const out: string[] = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      out.push(indexToA1(r, c));
    }
  }
  return out;
}

export function indexToA1(row: number, col: number): string {
  return `${indexToColLabel(col)}${row + 1}`;
}

function indexToColLabel(n: number): string {
  let s = '';
  n = Math.floor(n);
  n += 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function toNumber(val: unknown): number {
  if (typeof val === 'number') return val;
  const num = Number(val);
  return Number.isFinite(num) ? num : 0;
}

function flattenArgs(args: Array<string>, getA1Value: GetA1Value): Array<string | number> {
  const out: Array<string | number> = [];
  for (const raw of args) {
    const s = raw.trim();
    if (/^[A-Z]+[0-9]+\s*:\s*[A-Z]+[0-9]+$/i.test(s)) {
      for (const a1 of expandRange(s)) out.push(getA1Value(a1) ?? '');
    } else if (/^[A-Z]+[0-9]+$/i.test(s)) {
      out.push(getA1Value(s) ?? '');
    } else if (/^".*"$/.test(s)) {
      out.push(s.slice(1, -1));
    } else if (/^[0-9]+(\.[0-9]+)?$/.test(s)) {
      out.push(parseFloat(s));
    } else {
      out.push(s);
    }
  }
  return out;
}

function evalIfCondition(cond: string, getA1Value: GetA1Value): boolean {
  // Simple comparisons: A1>10, A1<=B1, 5=5, A1<>"x"
  const m = /(.+?)(>=|<=|<>|=|>|<)(.+)/.exec(cond);
  if (!m) return false;
  const left = m[1].trim();
  const op = m[2];
  const right = m[3].trim();
  const resolve = (expr: string): string | number => {
    if (/^[A-Z]+[0-9]+$/i.test(expr)) {
      const val = getA1Value(expr);
      return val !== null && val !== undefined ? val : '';
    }
    if (/^".*"$/.test(expr)) return expr.slice(1, -1);
    if (/^[0-9]+(\.[0-9]+)?$/.test(expr)) return parseFloat(expr);
    return expr;
  };
  const l = resolve(left);
  const r = resolve(right);
  switch (op) {
    case '>': return toNumber(l) > toNumber(r);
    case '<': return toNumber(l) < toNumber(r);
    case '>=': return toNumber(l) >= toNumber(r);
    case '<=': return toNumber(l) <= toNumber(r);
    case '=': return String(l) === String(r);
    case '<>': return String(l) !== String(r);
    default: return false;
  }
}

export function evaluateFormula(formula: string, getA1Value: GetA1Value): string | number {
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  // Functions first: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF
  const funcMatch = /^([A-Z]+)\((.*)\)$/.exec(f.trim());
  if (funcMatch) {
    const name = funcMatch[1].toUpperCase();
    // Split arguments by commas, but not inside quotes
    const args = [] as string[];
    let buf = '';
    let inStr = false;
    for (let i = 0; i < funcMatch[2].length; i++) {
      const ch = funcMatch[2][i];
      if (ch === '"') inStr = !inStr;
      if (ch === ',' && !inStr) {
        args.push(buf);
        buf = '';
      } else {
        buf += ch;
      }
    }
    if (buf.length) args.push(buf);

    const list = flattenArgs(args, getA1Value);

    switch (name) {
      case 'SUM':
        return list.flat().reduce((acc: number, v) => acc + toNumber(v), 0);
      case 'AVERAGE': {
        const nums = list.flat().map(toNumber);
        return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
      }
      case 'MIN':
        return Math.min(...list.flat().map(toNumber));
      case 'MAX':
        return Math.max(...list.flat().map(toNumber));
      case 'COUNT':
        return list.flat().filter(v => typeof v === 'number' || /^[0-9]+(\.[0-9]+)?$/.test(String(v))).length;
      case 'COUNTA':
        return list.flat().filter(v => v !== '' && v !== null && v !== undefined).length;
      case 'IF': {
        // IF(condition, trueVal, falseVal)
        if (args.length < 3) return '';
        const cond = args[0];
        const yes = flattenArgs([args[1]], getA1Value)[0];
        const no = flattenArgs([args[2]], getA1Value)[0];
        return evalIfCondition(cond, getA1Value) ? yes : no;
      }
      default:
        return '#NAME?';
    }
  }

  // Simple cell or number or string literal
  if (/^[A-Z]+[0-9]+$/i.test(f.trim())) {
    const val = getA1Value(f.trim());
    return val !== null && val !== undefined ? val : '';
  }
  if (/^".*"$/.test(f.trim())) return f.trim().slice(1, -1);
  if (/^[0-9]+(\.[0-9]+)?$/.test(f.trim())) return parseFloat(f.trim());

  // Unsupported expression form
  return '#N/A';
}


