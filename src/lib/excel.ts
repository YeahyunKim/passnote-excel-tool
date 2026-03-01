import * as XLSX from 'xlsx';

export type Row = Record<string, unknown>;

export type SheetData = {
  rows: Row[];
  columns: string[];
  sheetName: string;
};

export function normalizeIsbn(value: unknown): string {
  if (value == null) return '';
  const raw = String(value).trim().toUpperCase();
  if (!raw) return '';
  return raw.replace(/[^0-9X]/g, '');
}

export function getYearFromExcelValue(value: unknown): number | null {
  if (value == null) return null;

  if (value instanceof Date) {
    const t = value.getTime();
    if (!Number.isNaN(t)) return value.getFullYear();
    return null;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const d = XLSX.SSF?.parse_date_code?.(value);
    if (d && typeof d.y === 'number') return d.y;
    return null;
  }

  if (typeof value === 'string') {
    const s = value.trim();
    if (!s) return null;

    const m = s.match(/(19|20)\d{2}/);
    if (m) return Number(m[0]);

    const dt = new Date(s);
    if (!Number.isNaN(dt.getTime())) return dt.getFullYear();
  }

  return null;
}

/** 파일에서 시트 이름 목록만 읽습니다. */
export async function readSheetNames(file: File): Promise<string[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  return wb.SheetNames?.length ? [...wb.SheetNames] : [];
}

/** 지정한 인덱스(0부터)의 시트만 읽습니다. */
export async function readSheetByIndex(file: File, sheetIndex: number): Promise<SheetData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(
    buf,
    ({
    type: 'array',
    cellDates: true,
    dense: true,
    nodim: true, // 파일이 보고한 범위 무시, 실제 셀 기준으로 전체 행 읽기 (200행 등으로 잘리는 현상 방지)
    } as XLSX.ParsingOptions & { nodim?: boolean }),
  );
  const names = wb.SheetNames || [];
  const sheetName = names[sheetIndex] ?? names[0] ?? 'Sheet1';
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    throw new Error(`시트를 찾을 수 없습니다: ${sheetName}`);
  }

  const headerAoa = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    raw: true,
    blankrows: false,
  });
  const headerRow = (headerAoa[0] || []) as unknown[];
  const columns = headerRow
    .map((v) => String(v ?? '').trim())
    .filter((v) => v.length > 0);

  const rows = XLSX.utils.sheet_to_json<Row>(ws, {
    defval: '',
    raw: true,
    blankrows: false,
  });

  return { rows, columns, sheetName };
}

/** 첫 번째 시트만 읽습니다. (기존 동작 호환) */
export async function readFirstSheet(file: File): Promise<SheetData> {
  return readSheetByIndex(file, 0);
}

export function unionColumns(left: string[], right: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const c of left) {
    if (!c) continue;
    if (seen.has(c)) continue;
    seen.add(c);
    out.push(c);
  }
  for (const c of right) {
    if (!c) continue;
    if (seen.has(c)) continue;
    seen.add(c);
    out.push(c);
  }
  return out;
}

/** Excel 셀 문자 수 제한 (초과 시 에러 방지용) */
const EXCEL_CELL_MAX_LENGTH = 32767;

function truncateCellValue(value: unknown): unknown {
  if (value == null) return value;
  if (typeof value === 'string') {
    return value.length <= EXCEL_CELL_MAX_LENGTH ? value : value.slice(0, EXCEL_CELL_MAX_LENGTH);
  }
  return value;
}

function sanitizeRowsForExcel(rows: Row[]): Row[] {
  return rows.map((row) => {
    const out: Row = {};
    for (const [k, v] of Object.entries(row)) {
      out[k] = truncateCellValue(v);
    }
    return out;
  });
}

export function downloadRowsAsXlsx(params: { rows: Row[]; columns?: string[]; filename: string; sheetName?: string }) {
  const sheetName = params.sheetName || 'Sheet1';
  const columns = params.columns && params.columns.length > 0 ? params.columns : undefined;
  const rows = sanitizeRowsForExcel(params.rows);

  const ws = XLSX.utils.json_to_sheet(rows, columns ? { header: columns } : undefined);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, params.filename, { compression: true });
}

export function downloadSheetsAsXlsx(params: {
  filename: string;
  sheets: Array<{ rows: Row[]; columns?: string[]; sheetName: string }>;
}) {
  const wb = XLSX.utils.book_new();

  for (const s of params.sheets) {
    const columns = s.columns && s.columns.length > 0 ? s.columns : undefined;
    const rows = sanitizeRowsForExcel(s.rows);
    const ws = XLSX.utils.json_to_sheet(rows, columns ? { header: columns } : undefined);
    XLSX.utils.book_append_sheet(wb, ws, s.sheetName || 'Sheet');
  }

  XLSX.writeFile(wb, params.filename, { compression: true });
}

