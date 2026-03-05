import { useMemo, useState } from 'react';
import './App.css';
import { FileDrop } from './components/FileDrop';
import { VirtualTable } from './components/VirtualTable';
import {
  downloadRowsAsXlsx,
  getYearFromExcelValue,
  normalizeIsbn,
  readSheetByIndex,
  readSheetNames,
  unionColumns,
  type Row,
} from './lib/excel';

const YEAR_MIN = 2000;
const YEAR_MAX = 2030;
const AUTO_NO_BASE = '순번';

type DataPack = { rows: Row[]; columns: string[]; label: string };
type DedupeMode = 'none' | 'isbn_keep_first' | 'isbn_keep_last';
type PageMode = 'inspect' | 'db_upload';

type DataPackWithSheet = DataPack & { sheetName?: string };
type IsbnDuplicateGroup = { isbn: string; rows: Row[] };

function hasSubstring(value: unknown, needle: string): boolean {
  const s = String(value ?? '');
  return s.includes(needle);
}

function defaultColumnsFrom(rows: Row[], fallback: string[] = []) {
  if (fallback.length > 0) return fallback;
  const first = rows[0];
  if (!first) return [];
  return Object.keys(first);
}

function getAutoNoLabel(columns: string[]) {
  return columns.includes(AUTO_NO_BASE) ? `${AUTO_NO_BASE}(자동)` : AUTO_NO_BASE;
}

function normalizeProductCode(value: unknown): string {
  if (value == null) return '';
  return String(value).trim();
}

function pickColumn(columns: string[], candidates: string[]): string | null {
  for (const c of candidates) {
    if (columns.includes(c)) return c;
  }
  return null;
}

function formatDateYmd(value: unknown): string {
  if (value == null) return '';
  if (value instanceof Date) {
    const t = value.getTime();
    if (Number.isNaN(t)) return '';
    return value.toISOString().slice(0, 10);
  }
  return String(value).trim();
}

function normalizeDiscountRate(value: unknown): number | '' {
  if (value == null) return '';
  if (typeof value === 'number' && Number.isFinite(value)) {
    // 스마트스토어 엑셀에서 0.1 = 10% 형태로 내려오는 케이스 보정
    if (value > 0 && value <= 1) return Math.round(value * 100);
    return value;
  }
  const s = String(value).trim();
  if (!s) return '';
  const normalized = s.replace(/,/g, '');
  const pct = normalized.endsWith('%') ? normalized.slice(0, -1).trim() : normalized;
  const n = Number(pct);
  if (!Number.isFinite(n)) return '';
  if (n > 0 && n <= 1) return Math.round(n * 100);
  return n;
}

/** A 엑셀 "분철 1" 값에서 권수 추출. e.g. "스프링(3권)" → 3, 없으면 null */
function parseBuncheolKwons(value: unknown): number | null {
  if (value == null) return null;
  const s = String(value).trim();
  if (!s) return null;
  const m = s.match(/\((\d+)권\)/);
  return m ? Math.max(1, parseInt(m[1], 10)) : null;
}

/** B 행에서 페이지 수/쪽수 컬럼 값 숫자로 (페이지 수 / 페이지수 / 쪽수 등) */
function getPageCount(row: Row, columns: string[]): number | null {
  const pageCol = columns.find((c) => c === '페이지 수' || c === '페이지수' || c === '쪽수');
  if (!pageCol) return null;
  const v = row[pageCol];
  if (v == null) return null;
  let n: number;

  if (typeof v === 'number') {
    n = v;
  } else {
    const s = String(v).trim();
    if (!s) return null; // 엑셀에서 빈 셀(defval: '')은 여기서 걸러야 함

    // "1,234", "300쪽", "약 200" 같은 값도 최대한 숫자로 해석
    const normalized = s.replace(/,/g, '');
    const direct = Number(normalized);
    if (Number.isFinite(direct)) {
      n = direct;
    } else {
      const m = normalized.match(/\d+/);
      if (!m) return null;
      n = Number(m[0]);
    }
  }

  if (!Number.isFinite(n)) return null;
  const pages = Math.floor(n);
  return pages > 0 ? pages : null;
}

/** 권수 N으로 옵션 컬럼 값 생성 (엔터 포함) */
function buildOptionCells(kwons: number) {
  const price = 1500 * kwons;
  return {
    옵션명: '구매 도서 스캔 서비스 이용 신청\n제본/분철',
    옵션값: `[무료] 스캔 서비스 신청(O),신청 안함(X)\n제본/분철 안함,스프링 제본 (${kwons}권)`,
    옵션가: `0,0\n0,${price}`,
    '옵션 재고수량': '999,999',
  };
}

/** 연도 필터에 사용할 수 있는 날짜 컬럼 후보 (엑셀에 따라 출간일/출판일 등으로 되어 있을 수 있음) */
const DATE_COLUMN_CANDIDATES = ['출판날짜', '출간일', '출판일'];

function applyCommonFilters(params: {
  rows: Row[];
  columns: string[];
  excludeReservation: boolean;
  selectedYears: Set<number>;
  excludeEmptyPageCount?: boolean;
}) {
  const hasProductName = params.columns.includes('상품명');
  const dateCol = params.columns.find((c) => DATE_COLUMN_CANDIDATES.includes(c));

  let out = params.rows;
  if (params.excludeEmptyPageCount) {
    out = out.filter((r) => getPageCount(r, params.columns) != null);
  }
  if (params.excludeReservation && hasProductName) {
    out = out.filter((r) => !hasSubstring(r['상품명'], '예약판매'));
  }

  if (params.selectedYears.size === 0) return [];
  if (dateCol) {
    out = out.filter((r) => {
      const y = getYearFromExcelValue(r[dateCol]);
      return y != null && params.selectedYears.has(y);
    });
  }

  return out;
}

export default function App() {
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isInfoOpen, setIsInfoOpen] = useState(false);
  const [pageMode, setPageMode] = useState<PageMode>('inspect');

  // 1) 비교 영역
  const [compareA, setCompareA] = useState<File | null>(null);
  const [compareB, setCompareB] = useState<File | null>(null);
  const [compareASheetNames, setCompareASheetNames] = useState<string[]>([]);
  const [compareBSheetNames, setCompareBSheetNames] = useState<string[]>([]);
  const [compareASelectedIndex, setCompareASelectedIndex] = useState(0);
  const [compareBSelectedIndex, setCompareBSelectedIndex] = useState(0);
  const [compareBusy, setCompareBusy] = useState(false);
  const [comparePack, setComparePack] = useState<DataPackWithSheet | null>(null);

  // 2) 필터 영역
  const [filterFile, setFilterFile] = useState<File | null>(null);
  const [filterSheetNames, setFilterSheetNames] = useState<string[]>([]);
  const [filterSelectedIndex, setFilterSelectedIndex] = useState(0);
  const [filterBusy, setFilterBusy] = useState(false);
  const [filterPack, setFilterPack] = useState<DataPackWithSheet | null>(null);

  function handleCompareAChange(file: File | null) {
    setCompareA(file);
    setComparePack(null);
    if (file) {
      readSheetNames(file).then((names) => {
        setCompareASheetNames(names);
        setCompareASelectedIndex(0);
      });
    } else {
      setCompareASheetNames([]);
      setCompareASelectedIndex(0);
    }
  }

  function handleCompareBChange(file: File | null) {
    setCompareB(file);
    setComparePack(null);
    if (file) {
      readSheetNames(file).then((names) => {
        setCompareBSheetNames(names);
        setCompareBSelectedIndex(0);
      });
    } else {
      setCompareBSheetNames([]);
      setCompareBSelectedIndex(0);
    }
  }

  function handleFilterFileChange(file: File | null) {
    setFilterFile(file);
    setFilterPack(null);
    if (file) {
      readSheetNames(file).then((names) => {
        setFilterSheetNames(names);
        setFilterSelectedIndex(0);
      });
    } else {
      setFilterSheetNames([]);
      setFilterSelectedIndex(0);
    }
  }

  function handleDbAChange(file: File | null) {
    setDbAFile(file);
    setDbPack(null);
    setDbErrorMessage(null);
    if (file) {
      readSheetNames(file).then((names) => {
        setDbASheetNames(names);
        setDbASelectedIndex(0);
      });
    } else {
      setDbASheetNames([]);
      setDbASelectedIndex(0);
    }
  }

  function handleDbBChange(file: File | null) {
    setDbBFile(file);
    setDbPack(null);
    setDbErrorMessage(null);
    if (file) {
      readSheetNames(file).then((names) => {
        setDbBSheetNames(names);
        setDbBSelectedIndex(0);
      });
    } else {
      setDbBSheetNames([]);
      setDbBSelectedIndex(0);
    }
  }

  // 공통 옵션(1/2에 적용)
  const [excludeReservation, setExcludeReservation] = useState(true);
  const [yearStart, setYearStart] = useState(YEAR_MIN);
  const [yearEnd, setYearEnd] = useState(YEAR_MAX);
  const selectedYears = useMemo(() => {
    if (yearStart > yearEnd) return new Set<number>();
    const set = new Set<number>();
    for (let y = yearStart; y <= yearEnd; y++) set.add(y);
    return set;
  }, [yearStart, yearEnd]);
  /** 비교 시 A 분철1 비었을 때 권수 산정: B 페이지수 / 이 값 (0~1000, 기본 300) */
  const [pagesPerBook, setPagesPerBook] = useState(300);
  /** 비교 시 쪽수(페이지 수)가 비어 있는 B 행 제외 여부 */
  const [excludeEmptyPageCount, setExcludeEmptyPageCount] = useState(true);

  // 3) 취합 영역
  const [aggregateRows, setAggregateRows] = useState<Row[]>([]);
  const [aggregateColumns, setAggregateColumns] = useState<string[]>([]);
  const [dedupeMode, setDedupeMode] = useState<DedupeMode>('isbn_keep_first');

  // 4) DB 업로드 엑셀
  const [dbErrorMessage, setDbErrorMessage] = useState<string | null>(null);
  const [dbAFile, setDbAFile] = useState<File | null>(null);
  const [dbBFile, setDbBFile] = useState<File | null>(null);
  const [dbASheetNames, setDbASheetNames] = useState<string[]>([]);
  const [dbBSheetNames, setDbBSheetNames] = useState<string[]>([]);
  const [dbASelectedIndex, setDbASelectedIndex] = useState(0);
  const [dbBSelectedIndex, setDbBSelectedIndex] = useState(0);
  const [dbBusy, setDbBusy] = useState(false);
  const [dbPack, setDbPack] = useState<DataPack | null>(null);
  const [dbAggregateRows, setDbAggregateRows] = useState<Row[]>([]);
  const [isDbDedupeOpen, setIsDbDedupeOpen] = useState(false);
  const [dbDupGroups, setDbDupGroups] = useState<IsbnDuplicateGroup[]>([]);
  const [dbChosenByIsbn, setDbChosenByIsbn] = useState<Record<string, Row>>({});

  const canCompare = !!compareA && !!compareB && !compareBusy;
  const canFilter = !!filterFile && !filterBusy;
  const totalAggregate = aggregateRows.length;
  const aggregateNoLabel = useMemo(() => getAutoNoLabel(aggregateColumns), [aggregateColumns]);

  const selectedYearText = useMemo(() => {
    if (selectedYears.size === 0) return '선택 없음';
    if (yearStart === YEAR_MIN && yearEnd === YEAR_MAX) return '2000 ~ 2030 전체';
    return `${yearStart} ~ ${yearEnd}`;
  }, [selectedYears.size, yearStart, yearEnd]);

  async function runCompare() {
    if (!compareA || !compareB) return;
    setErrorMessage(null);
    setCompareBusy(true);
    setComparePack(null);

    try {
      const [a, b] = await Promise.all([
        readSheetByIndex(compareA, compareASelectedIndex),
        readSheetByIndex(compareB, compareBSelectedIndex),
      ]);
      const out = [...b.rows];

      const aIsbnToBuncheol = new Map<string, unknown>();
      const buncheolCol = a.columns.find((c) => c === '분철 1');
      for (const r of a.rows) {
        const k = normalizeIsbn(r['ISBN']);
        if (k && !aIsbnToBuncheol.has(k) && buncheolCol != null) {
          aIsbnToBuncheol.set(k, r[buncheolCol]);
        }
      }

      const perBook = Math.max(1, Math.min(1000, pagesPerBook));
      const rowsWithOptions: Row[] = out.map((row) => {
        const isbn = normalizeIsbn(row['ISBN']);
        const buncheol1 = isbn ? aIsbnToBuncheol.get(isbn) : undefined;
        let kwons = parseBuncheolKwons(buncheol1);
        if (kwons == null) {
          const pages = getPageCount(row, b.columns);
          kwons = pages != null && perBook > 0 ? Math.max(1, Math.ceil(pages / perBook)) : 1;
        }
        const cells = buildOptionCells(kwons);
        return { ...row, ...cells };
      });

      const filtered = applyCommonFilters({
        rows: rowsWithOptions,
        columns: b.columns,
        excludeReservation,
        selectedYears,
        excludeEmptyPageCount,
      });

      setComparePack({
        rows: filtered,
        columns: defaultColumnsFrom(filtered, b.columns),
        label: `B 전체 · 옵션 반영 · ${filtered.length.toLocaleString()}건`,
        sheetName: b.sheetName,
      });
    } catch (e) {
      setErrorMessage(e instanceof Error ? e.message : '비교 중 오류가 발생했습니다.');
    } finally {
      setCompareBusy(false);
    }
  }

  async function runFilter() {
    if (!filterFile) return;
    setErrorMessage(null);
    setFilterBusy(true);
    setFilterPack(null);

    try {
      const src = await readSheetByIndex(filterFile, filterSelectedIndex);
      const filtered = applyCommonFilters({
        rows: src.rows,
        columns: src.columns,
        excludeReservation,
        selectedYears,
        excludeEmptyPageCount,
      });

      const perBook = Math.max(1, Math.min(1000, pagesPerBook));
      const rowsWithOptions: Row[] = filtered.map((row) => {
        const pages = getPageCount(row, src.columns);
        const kwons =
          pages != null && perBook > 0 ? Math.max(1, Math.ceil(pages / perBook)) : 1;
        const cells = buildOptionCells(kwons);
        return { ...row, ...cells };
      });

      setFilterPack({
        rows: rowsWithOptions,
        columns: defaultColumnsFrom(rowsWithOptions, src.columns),
        label: `필터 결과 · ${rowsWithOptions.length.toLocaleString()}건`,
        sheetName: src.sheetName,
      });
    } catch (e) {
      setErrorMessage(e instanceof Error ? e.message : '필터링 중 오류가 발생했습니다.');
    } finally {
      setFilterBusy(false);
    }
  }

  function addToAggregate(pack: DataPackWithSheet) {
    setErrorMessage(null);
    const hasSheet = !!pack.sheetName;
    const nextPackColumns = hasSheet ? unionColumns(pack.columns, ['카테고리코드']) : pack.columns;
    setAggregateColumns((prev) => unionColumns(prev, nextPackColumns));

    setAggregateRows((prev) => {
      const rowsToAdd: Row[] = hasSheet
        ? pack.rows.map((r) => ({ ...(r as Row), 카테고리코드: pack.sheetName! } as Row))
        : (pack.rows as Row[]);

      if (dedupeMode === 'none') return [...prev, ...rowsToAdd];

      const indexByIsbn = new Map<string, number>();
      prev.forEach((r, i) => {
        const k = normalizeIsbn(r['ISBN']);
        if (k) indexByIsbn.set(k, i);
      });

      const next = [...prev];
      for (const r of rowsToAdd) {
        const k = normalizeIsbn(r['ISBN']);
        if (!k) {
          next.push(r);
          continue;
        }

        const existingIdx = indexByIsbn.get(k);
        if (existingIdx == null) {
          indexByIsbn.set(k, next.length);
          next.push(r);
          continue;
        }

        if (dedupeMode === 'isbn_keep_last') {
          next[existingIdx] = r;
        }
      }
      return next;
    });
  }

  const EXPORT_CHUNK_SIZE = 500;
  const EXPORT_FULL_CHUNK_SIZE = 5000;

  function downloadAggregateSingle() {
    if (aggregateRows.length === 0) return;
    const noLabel = getAutoNoLabel(aggregateColumns);
    const exportColumns =
      aggregateColumns.includes('판매자 상품코드')
        ? [noLabel, ...aggregateColumns]
        : [noLabel, '판매자 상품코드', ...aggregateColumns];

    const totalChunks = Math.max(1, Math.ceil(aggregateRows.length / EXPORT_FULL_CHUNK_SIZE));
    for (let c = 0; c < totalChunks; c++) {
      const start = c * EXPORT_FULL_CHUNK_SIZE;
      const end = Math.min(start + EXPORT_FULL_CHUNK_SIZE, aggregateRows.length);
      const chunkRows = aggregateRows.slice(start, end);
      const rowsForExport = chunkRows.map((r, i) => ({
        [noLabel]: start + i + 1,
        ...r,
        '판매자 상품코드': 100001 + start + i,
      }));
      const chunkNum = String(c + 1).padStart(2, '0');
      setTimeout(() => {
        downloadRowsAsXlsx({
          rows: rowsForExport,
          columns: exportColumns,
          filename: `전체취합_${chunkNum}.xlsx`,
          sheetName: 'Aggregate',
        });
      }, c * 350);
    }
  }

  function downloadAggregateChunked() {
    if (aggregateRows.length === 0) return;
    const noLabel = getAutoNoLabel(aggregateColumns);
    const exportColumns =
      aggregateColumns.includes('판매자 상품코드')
        ? [noLabel, ...aggregateColumns]
        : [noLabel, '판매자 상품코드', ...aggregateColumns];

    const totalChunks = Math.max(1, Math.ceil(aggregateRows.length / EXPORT_CHUNK_SIZE));
    for (let c = 0; c < totalChunks; c++) {
      const start = c * EXPORT_CHUNK_SIZE;
      const end = Math.min(start + EXPORT_CHUNK_SIZE, aggregateRows.length);
      const chunkRows = aggregateRows.slice(start, end);
      const rowsForExport = chunkRows.map((r, i) => ({
        [noLabel]: start + i + 1,
        ...r,
        '판매자 상품코드': 100001 + start + i,
      }));
      const chunkNum = String(c + 1).padStart(2, '0');
      setTimeout(() => {
        downloadRowsAsXlsx({
          rows: rowsForExport,
          columns: exportColumns,
          filename: `취합_${chunkNum}.xlsx`,
          sheetName: 'Aggregate',
        });
      }, c * 350);
    }
  }

  const DB_EXPORT_COLUMNS = [
    'product_id',
    'product_code',
    'price',
    'discount_rate',
    'is_discount_applied',
    'status',
    'pages',
    'title',
    'author',
    'description',
    'isbn',
    'publisher',
    'file_key',
    'thumbnail_image_url',
    'publication_date',
    'smartstore_book_category_id',
  ];

  async function runDbUploadBuild() {
    if (!dbAFile || !dbBFile) return;
    setDbErrorMessage(null);
    setDbBusy(true);
    setDbPack(null);

    try {
      const [a, b] = await Promise.all([readSheetByIndex(dbAFile, dbASelectedIndex), readSheetByIndex(dbBFile, dbBSelectedIndex)]);

      const aProductIdCol = pickColumn(a.columns, ['상품번호(스마트스토어)', '상품번호']);
      const aProductCodeCol = pickColumn(a.columns, ['판매자상품코드', '판매자 상품코드', '판매자상품 코드']);
      const aPriceCol = pickColumn(a.columns, ['판매가']);
      const aDiscountCol = pickColumn(a.columns, ['판매자할인', '할인율', '할인율(%)']);
      const aStatusCol = pickColumn(a.columns, ['판매상태', '판매 상태']);
      const aTitleCol = pickColumn(a.columns, ['상품명']);
      const aThumbCol = pickColumn(a.columns, ['대표이미지 URL', '대표이미지URL']);
      const aIsbnCol = pickColumn(a.columns, ['ISBN', 'ISBN13', '판매자바코드']);
      const aSubCategoryCol = pickColumn(a.columns, ['소분류', '카테고리(소분류)', '카테고리 소분류']);

      const bProductCodeCol = pickColumn(b.columns, ['판매자 상품코드', '판매자상품코드']);
      const bAuthorCol = pickColumn(b.columns, ['글작가', '저자', '저자명']);
      const bDescCol = pickColumn(b.columns, ['상세설명', '상세 설명', '상품상세', '상세']);
      const bIsbnCol = pickColumn(b.columns, ['ISBN', 'ISBN13']);
      const bPublisherCol = pickColumn(b.columns, ['출판사', '출판사명']);
      const bPubDateCol = pickColumn(b.columns, ['출간일', '출판날짜', '출판일']);

      if (!aProductCodeCol) throw new Error('A 엑셀에서 판매자상품코드 컬럼을 찾지 못했습니다.');
      if (!bProductCodeCol) throw new Error('B 엑셀에서 판매자 상품코드 컬럼을 찾지 못했습니다.');
      if (!aProductIdCol) throw new Error('A 엑셀에서 상품번호(스마트스토어) 컬럼을 찾지 못했습니다.');

      const bByCode = new Map<string, Row>();
      for (const r of b.rows) {
        const code = normalizeProductCode(r[bProductCodeCol]);
        if (!code) continue;
        bByCode.set(code, r);
      }

      const outRows: Row[] = [];
      let matched = 0;
      let missingB = 0;

      for (const ar of a.rows) {
        const code = normalizeProductCode(ar[aProductCodeCol]);
        if (!code) continue;
        const br = bByCode.get(code);
        if (br) matched++;
        else missingB++;

        const pages = br ? getPageCount(br, b.columns) : null;
        const isbnFromA = aIsbnCol ? normalizeIsbn(ar[aIsbnCol]) : '';
        const isbnFromB = bIsbnCol ? normalizeIsbn(br?.[bIsbnCol]) : '';

        outRows.push({
          product_id: ar[aProductIdCol] ?? '',
          product_code: code,
          price: aPriceCol ? ar[aPriceCol] ?? '' : '',
          discount_rate: aDiscountCol ? normalizeDiscountRate(ar[aDiscountCol]) : '',
          is_discount_applied: true,
          status: aStatusCol ? ar[aStatusCol] ?? '' : '',
          pages: pages ?? '',
          title: aTitleCol ? ar[aTitleCol] ?? '' : '',
          author: bAuthorCol ? br?.[bAuthorCol] ?? '' : '',
          description: bDescCol ? br?.[bDescCol] ?? '' : '',
          isbn: isbnFromA || isbnFromB || '',
          publisher: bPublisherCol ? br?.[bPublisherCol] ?? '' : '',
          file_key: '',
          thumbnail_image_url: aThumbCol ? ar[aThumbCol] ?? '' : '',
          publication_date: bPubDateCol ? formatDateYmd(br?.[bPubDateCol]) : '',
          smartstore_book_category_id: aSubCategoryCol ? ar[aSubCategoryCol] ?? '' : '',
          __matched: !!br,
        });
      }

      setDbPack({
        rows: outRows,
        columns: DB_EXPORT_COLUMNS,
        label: `생성 결과 · ${outRows.length.toLocaleString()}건 (매칭 ${matched.toLocaleString()} · B없음 ${missingB.toLocaleString()})`,
      });
    } catch (e) {
      setDbErrorMessage(e instanceof Error ? e.message : '생성 중 오류가 발생했습니다.');
    } finally {
      setDbBusy(false);
    }
  }

  function downloadDbUploadXlsx() {
    if (!dbPack || dbPack.rows.length === 0) return;
    downloadRowsAsXlsx({
      rows: dbPack.rows,
      columns: DB_EXPORT_COLUMNS,
      filename: 'DB업로드.xlsx',
      sheetName: 'DBUpload',
    });
  }

  function addDbToAggregate(pack: DataPack) {
    const matchedRows = pack.rows.filter((r) => Boolean((r as Row)['__matched']));
    const normalized = matchedRows.map((r) => ({
      ...r,
      discount_rate: normalizeDiscountRate((r as Row)['discount_rate']),
    }));
    setDbAggregateRows((prev) => [...prev, ...normalized]);
  }

  function resetDbAggregate() {
    setDbAggregateRows([]);
  }

  function getDbIsbnKey(row: Row): string {
    // DB업로드 탭에서는 보통 `isbn`(snake_case)로 들어오지만, 방어적으로 `ISBN`도 허용
    return normalizeIsbn((row as Row)['isbn'] ?? (row as Row)['ISBN']);
  }

  function getDbRowKey(row: Row): string {
    const isbn = getDbIsbnKey(row);
    const pid = String((row as Row)['product_id'] ?? '');
    const pcode = String((row as Row)['product_code'] ?? '');
    const title = String((row as Row)['title'] ?? '');
    return `${isbn}__${pid}__${pcode}__${title}`;
  }

  function getDbRowTitle(row: Row): string {
    const t = String((row as Row)['title'] ?? '').trim();
    return t || '(제목 없음)';
  }

  function getDbRowMeta(row: Row): { productId: string; productCode: string; status: string } {
    return {
      productId: String((row as Row)['product_id'] ?? '').trim(),
      productCode: String((row as Row)['product_code'] ?? '').trim(),
      status: String((row as Row)['status'] ?? '').trim(),
    };
  }

  function downloadDbAggregateXlsx(rows: Row[] = dbAggregateRows) {
    if (rows.length === 0) return;
    const DB_EXPORT_CHUNK_SIZE = 3000;
    const totalChunks = Math.max(1, Math.ceil(rows.length / DB_EXPORT_CHUNK_SIZE));
    for (let c = 0; c < totalChunks; c++) {
      const start = c * DB_EXPORT_CHUNK_SIZE;
      const end = Math.min(start + DB_EXPORT_CHUNK_SIZE, rows.length);
      const chunkRows = rows.slice(start, end);
      const chunkNum = String(c + 1).padStart(2, '0');
      setTimeout(() => {
        downloadRowsAsXlsx({
          rows: chunkRows,
          columns: DB_EXPORT_COLUMNS,
          filename: `DB업로드_취합_${chunkNum}.xlsx`,
          sheetName: 'DBUpload',
        });
      }, c * 350);
    }
  }

  function openDbIsbnDedupeIfNeeded() {
    if (dbAggregateRows.length === 0) return;

    const byIsbn = new Map<string, Row[]>();
    for (const r of dbAggregateRows) {
      const isbn = getDbIsbnKey(r);
      if (!isbn) continue;
      const arr = byIsbn.get(isbn);
      if (arr) arr.push(r);
      else byIsbn.set(isbn, [r]);
    }

    const dupGroups: IsbnDuplicateGroup[] = [];
    for (const [isbn, rows] of byIsbn.entries()) {
      if (rows.length > 1) dupGroups.push({ isbn, rows });
    }

    if (dupGroups.length === 0) {
      downloadDbAggregateXlsx(dbAggregateRows);
      return;
    }

    dupGroups.sort((a, b) => a.isbn.localeCompare(b.isbn));
    setDbDupGroups(dupGroups);
    setDbChosenByIsbn({});
    setIsDbDedupeOpen(true);
  }

  function chooseDbRow(isbn: string, row: Row) {
    setDbChosenByIsbn((prev) => ({ ...prev, [isbn]: row }));
  }

  function unchooseDbRow(isbn: string) {
    setDbChosenByIsbn((prev) => {
      const next = { ...prev };
      delete next[isbn];
      return next;
    });
  }

  function confirmDbDedupeAndSave() {
    if (dbDupGroups.length === 0) return;
    const chosenIsbns = new Set(Object.keys(dbChosenByIsbn));
    const requiredIsbns = new Set(dbDupGroups.map((g) => g.isbn));
    for (const isbn of requiredIsbns) {
      if (!chosenIsbns.has(isbn)) return;
    }

    const chosenByIsbn = new Map<string, Row>();
    for (const g of dbDupGroups) {
      const chosen = dbChosenByIsbn[g.isbn];
      if (chosen) chosenByIsbn.set(g.isbn, chosen);
    }

    const dupIsbnSet = new Set(dbDupGroups.map((g) => g.isbn));
    const seen = new Set<string>();
    const out: Row[] = [];
    for (const r of dbAggregateRows) {
      const isbn = getDbIsbnKey(r);
      if (!isbn || !dupIsbnSet.has(isbn)) {
        out.push(r);
        continue;
      }
      if (seen.has(isbn)) continue;
      const chosen = chosenByIsbn.get(isbn);
      if (chosen) out.push(chosen);
      seen.add(isbn);
    }

    setDbAggregateRows(out);
    setIsDbDedupeOpen(false);
    setDbDupGroups([]);
    setDbChosenByIsbn({});
    downloadDbAggregateXlsx(out);
  }

  return (
    <div className="app">
      <header className="top">
        <div className="topBar">
          <div className="topLeft">
            <h1 className="topTitle">PassNote Excel Tool</h1>
            <p className="topSub">브라우저에서만 동작 · 파일은 외부로 전송되지 않습니다.</p>
          </div>
          <button
            className="infoIconBtn"
            type="button"
            aria-label="안내"
            onClick={() => setIsInfoOpen(true)}
          >
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" focusable="false">
              <path
                fill="currentColor"
                d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20Zm0 4.7a1.25 1.25 0 1 1 0 2.5 1.25 1.25 0 0 1 0-2.5ZM10.9 11a1 1 0 0 1 1-1h.2a1 1 0 0 1 1 1v6a1 1 0 0 1-2 0v-5Z"
              />
            </svg>
          </button>
        </div>
        <div className="pageTabs" role="tablist" aria-label="페이지 선택">
          <button
            className={['tabBtn', pageMode === 'inspect' ? 'tabBtn--active' : ''].filter(Boolean).join(' ')}
            type="button"
            role="tab"
            aria-selected={pageMode === 'inspect'}
            onClick={() => setPageMode('inspect')}
          >
            엑셀 검수
          </button>
          <button
            className={['tabBtn', pageMode === 'db_upload' ? 'tabBtn--active' : ''].filter(Boolean).join(' ')}
            type="button"
            role="tab"
            aria-selected={pageMode === 'db_upload'}
            onClick={() => setPageMode('db_upload')}
          >
            DB업로드 엑셀
          </button>
        </div>
      </header>

      {pageMode === 'inspect' && errorMessage ? (
        <div className="alert alert--error" role="alert">
          <div className="alertTitle">처리 실패</div>
          <div className="alertBody">{errorMessage}</div>
        </div>
      ) : null}

      {pageMode === 'inspect' ? (
        <>
          <div className="layout">
        <div className="layoutLeft">
          <section className="step step--1">
            <div className="stepHead">
              <span className="stepTitle">1. 비교</span>
              <button className="btn btn--primary" type="button" onClick={runCompare} disabled={!canCompare}>
                {compareBusy ? '처리 중…' : '비교 실행'}
              </button>
            </div>
            <div className="twoCol">
              <div className="fieldGroup">
                <FileDrop label="A 엑셀" value={compareA} onChange={handleCompareAChange} disabled={compareBusy} />
                {compareASheetNames.length > 0 && (
                  <div className="sheetSelect">
                    <label className="sheetSelectLabel" htmlFor="sheet-a">시트</label>
                    <select
                      id="sheet-a"
                      className="select sheetSelectInput"
                      value={compareASelectedIndex}
                      onChange={(e) => {
                        setCompareASelectedIndex(Number(e.target.value));
                        setComparePack(null);
                      }}
                      disabled={compareBusy}
                      aria-label="A 엑셀 시트 선택"
                    >
                      {compareASheetNames.map((name, i) => (
                        <option key={i} value={i}>{name}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
              <div className="fieldGroup">
                <FileDrop label="B 엑셀" value={compareB} onChange={handleCompareBChange} disabled={compareBusy} />
                {compareBSheetNames.length > 0 && (
                  <div className="sheetSelect">
                    <label className="sheetSelectLabel" htmlFor="sheet-b">시트</label>
                    <select
                      id="sheet-b"
                      className="select sheetSelectInput"
                      value={compareBSelectedIndex}
                      onChange={(e) => {
                        setCompareBSelectedIndex(Number(e.target.value));
                        setComparePack(null);
                      }}
                      disabled={compareBusy}
                      aria-label="B 엑셀 시트 선택"
                    >
                      {compareBSheetNames.map((name, i) => (
                        <option key={i} value={i}>{name}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            </div>
            <div className="stepResult">
              <span className="stepResultLabel">{comparePack ? comparePack.label : '결과'}</span>
              <button
                className="btn btn--secondary"
                type="button"
                onClick={() => comparePack && addToAggregate(comparePack)}
                disabled={!comparePack || comparePack.rows.length === 0}
              >
                취합에 추가
              </button>
            </div>
            <VirtualTable
              rows={comparePack?.rows ?? []}
              columns={comparePack?.columns ?? []}
              showRowNumbers
              rowNumberHeader={getAutoNoLabel(comparePack?.columns ?? [])}
              height={320}
              emptyText="A/B 파일 올린 뒤 비교 실행"
            />
          </section>

          <section className="step step--2">
            <div className="stepHead">
              <span className="stepTitle">2. 필터</span>
              <button className="btn btn--primary" type="button" onClick={runFilter} disabled={!canFilter}>
                {filterBusy ? '처리 중…' : '필터 실행'}
              </button>
            </div>
            <div className="fieldGroup">
              <FileDrop label="엑셀" value={filterFile} onChange={handleFilterFileChange} disabled={filterBusy} />
              {filterSheetNames.length > 0 && (
                <div className="sheetSelect">
                  <label className="sheetSelectLabel" htmlFor="sheet-filter">시트</label>
                  <select
                    id="sheet-filter"
                    className="select sheetSelectInput"
                    value={filterSelectedIndex}
                    onChange={(e) => {
                      setFilterSelectedIndex(Number(e.target.value));
                      setFilterPack(null);
                    }}
                    disabled={filterBusy}
                    aria-label="필터 엑셀 시트 선택"
                  >
                    {filterSheetNames.map((name, i) => (
                      <option key={i} value={i}>{name}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>
            <div className="stepResult">
              <span className="stepResultLabel">{filterPack ? filterPack.label : '결과'}</span>
              <button
                className="btn btn--secondary"
                type="button"
                onClick={() => filterPack && addToAggregate(filterPack)}
                disabled={!filterPack || filterPack.rows.length === 0}
              >
                취합에 추가
              </button>
            </div>
            <VirtualTable
              rows={filterPack?.rows ?? []}
              columns={filterPack?.columns ?? []}
              showRowNumbers
              rowNumberHeader={getAutoNoLabel(filterPack?.columns ?? [])}
              height={320}
              emptyText="파일 올린 뒤 필터 실행"
            />
          </section>
        </div>

        <aside className="layoutRight">
          <section className="step step--opt">
            <div className="stepTitle">옵션</div>
            <label className="switch">
              <input
                type="checkbox"
                checked={excludeReservation}
                onChange={(e) => setExcludeReservation(e.target.checked)}
              />
              <span className="switchTrack" aria-hidden="true">
                <span className="switchThumb" />
              </span>
              <span className="switchText">예약판매 행 제외</span>
            </label>
            <label className="switch">
              <input
                type="checkbox"
                checked={excludeEmptyPageCount}
                onChange={(e) => setExcludeEmptyPageCount(e.target.checked)}
              />
              <span className="switchTrack" aria-hidden="true">
                <span className="switchThumb" />
              </span>
              <span className="switchText">쪽수 비어 있는 행 제외</span>
            </label>
            <div className="pagesPerBookBlock">
              <div className="pagesPerBookHead">
                <span className="yearLabel">페이지/권</span>
                <span className="yearMeta">{pagesPerBook} (분철1 비었을 때)</span>
              </div>
              <input
                type="range"
                min={1}
                max={1000}
                value={pagesPerBook}
                onChange={(e) => setPagesPerBook(Number(e.target.value))}
                className="pagesPerBookSlider"
                aria-label="페이지당 권수 (0~1000)"
              />
            </div>
            <div className="yearBlock">
              <div className="yearTop">
                <span className="yearLabel">연도</span>
                <span className="yearMeta">{selectedYearText}</span>
                <div className="yearQuick">
                  <button className="btn btn--secondary btn--sm" type="button" onClick={() => { setYearStart(YEAR_MIN); setYearEnd(YEAR_MAX); }}>
                    전체
                  </button>
                  <button className="btn btn--secondary btn--sm" type="button" onClick={() => { setYearStart(YEAR_MAX); setYearEnd(YEAR_MIN); }}>
                    해제
                  </button>
                </div>
              </div>
              <div className="yearRangeTrack" role="group" aria-label="출판연도 범위 (2000~2030)">
                <span className="yearRangeEdge">{YEAR_MIN}</span>
                <div className="yearRangeSliders">
                  <input
                    type="range"
                    min={YEAR_MIN}
                    max={YEAR_MAX}
                    value={yearStart}
                    onChange={(e) => setYearStart(Math.min(Number(e.target.value), yearEnd))}
                    className="yearRangeInput yearRangeInput--left"
                    aria-label="시작 연도 (왼쪽 드래그)"
                  />
                  <input
                    type="range"
                    min={YEAR_MIN}
                    max={YEAR_MAX}
                    value={yearEnd}
                    onChange={(e) => setYearEnd(Math.max(Number(e.target.value), yearStart))}
                    className="yearRangeInput yearRangeInput--right"
                    aria-label="끝 연도 (오른쪽 드래그)"
                  />
                </div>
                <span className="yearRangeEdge">{YEAR_MAX}</span>
              </div>
              <p className="yearRangeHint">왼쪽: 시작 연도 · 오른쪽: 끝 연도 (드래그로 조절)</p>
            </div>
            <p className="optHint">옵션 변경 후 실행 버튼을 다시 누르세요.</p>
          </section>
        </aside>
      </div>

      <section className="step step--3">
        <div className="stepHead stepHead--bar">
          <span className="stepTitle">3. 취합</span>
          <span className="stepCount">{totalAggregate.toLocaleString()}건</span>
          <select
            className="select"
            value={dedupeMode}
            onChange={(e) => setDedupeMode(e.target.value as DedupeMode)}
            aria-label="중복 처리"
          >
            <option value="isbn_keep_first">중복: 최초 유지</option>
            <option value="isbn_keep_last">중복: 최신 유지</option>
            <option value="none">중복 허용</option>
          </select>
          <button
            className="btn btn--secondary"
            type="button"
            onClick={() => { setAggregateRows([]); setAggregateColumns([]); }}
            disabled={totalAggregate === 0}
          >
            초기화
          </button>
          <button className="btn btn--primary" type="button" onClick={downloadAggregateSingle} disabled={totalAggregate === 0}>
            전체 저장
          </button>
          <button className="btn btn--secondary" type="button" onClick={downloadAggregateChunked} disabled={totalAggregate === 0}>
            분할 저장
          </button>
        </div>
        <VirtualTable
          rows={aggregateRows}
          columns={aggregateColumns}
          showRowNumbers
          rowNumberHeader={aggregateNoLabel}
          height={480}
          emptyText="취합된 데이터 없음"
        />
      </section>
        </>
      ) : (
        <>
          {dbErrorMessage ? (
            <div className="alert alert--error" role="alert">
              <div className="alertTitle">처리 실패</div>
              <div className="alertBody">{dbErrorMessage}</div>
            </div>
          ) : null}

          <section className="step step--db">
            <div className="stepHead">
              <span className="stepTitle">DB 업로드 엑셀 생성</span>
              <button
                className="btn btn--primary"
                type="button"
                onClick={runDbUploadBuild}
                disabled={!dbAFile || !dbBFile || dbBusy}
              >
                {dbBusy ? '처리 중…' : '생성'}
              </button>
              <button
                className="btn btn--secondary"
                type="button"
                onClick={downloadDbUploadXlsx}
                disabled={!dbPack || dbPack.rows.length === 0}
              >
                다운로드
              </button>
            </div>

            <div className="twoCol">
              <div className="fieldGroup">
                <FileDrop label="A (스마트스토어 다운로드)" value={dbAFile} onChange={handleDbAChange} disabled={dbBusy} />
                {dbASheetNames.length > 0 && (
                  <div className="sheetSelect">
                    <label className="sheetSelectLabel" htmlFor="sheet-db-a">시트</label>
                    <select
                      id="sheet-db-a"
                      className="select sheetSelectInput"
                      value={dbASelectedIndex}
                      onChange={(e) => {
                        setDbASelectedIndex(Number(e.target.value));
                        setDbPack(null);
                      }}
                      disabled={dbBusy}
                      aria-label="A 엑셀 시트 선택"
                    >
                      {dbASheetNames.map((name, i) => (
                        <option key={i} value={i}>{name}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              <div className="fieldGroup">
                <FileDrop label="B (취합 엑셀)" value={dbBFile} onChange={handleDbBChange} disabled={dbBusy} />
                {dbBSheetNames.length > 0 && (
                  <div className="sheetSelect">
                    <label className="sheetSelectLabel" htmlFor="sheet-db-b">시트</label>
                    <select
                      id="sheet-db-b"
                      className="select sheetSelectInput"
                      value={dbBSelectedIndex}
                      onChange={(e) => {
                        setDbBSelectedIndex(Number(e.target.value));
                        setDbPack(null);
                      }}
                      disabled={dbBusy}
                      aria-label="B 엑셀 시트 선택"
                    >
                      {dbBSheetNames.map((name, i) => (
                        <option key={i} value={i}>{name}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            </div>

            <div className="stepResult">
              <span className="stepResultLabel">{dbPack ? dbPack.label : '결과'}</span>
              <button
                className="btn btn--secondary"
                type="button"
                onClick={() => dbPack && addDbToAggregate(dbPack)}
                disabled={!dbPack || dbPack.rows.length === 0}
              >
                취합에 추가
              </button>
            </div>

            <VirtualTable
              rows={dbPack?.rows ?? []}
              columns={dbPack?.columns ?? DB_EXPORT_COLUMNS}
              showRowNumbers
              rowNumberHeader="순번"
              height={520}
              emptyText="A/B 파일 올린 뒤 생성"
            />
          </section>

          <section className="step step--dbAgg">
            <div className="stepHead stepHead--bar">
              <span className="stepTitle">다운로드</span>
              <span className="stepCount">{dbAggregateRows.length.toLocaleString()}건</span>
              <button
                className="btn btn--secondary"
                type="button"
                onClick={resetDbAggregate}
                disabled={dbAggregateRows.length === 0}
              >
                초기화
              </button>
              <button
                className="btn btn--primary"
                type="button"
                onClick={openDbIsbnDedupeIfNeeded}
                disabled={dbAggregateRows.length === 0}
              >
                저장
              </button>
            </div>
            <VirtualTable
              rows={dbAggregateRows}
              columns={DB_EXPORT_COLUMNS}
              showRowNumbers
              rowNumberHeader="순번"
              height={420}
              emptyText="취합된 데이터 없음"
            />
          </section>
        </>
      )}

      <footer className="footer">
        <p className="footerText">ISBN · 상품명 · 출판날짜/출간일 컬럼 사용</p>
      </footer>

      {isDbDedupeOpen ? (
        <div className="modalOverlay" role="presentation" onClick={() => setIsDbDedupeOpen(false)}>
          <div
            className="modal modal--wide"
            role="dialog"
            aria-modal="true"
            aria-label="ISBN 중복 선택"
            onClick={(e) => e.stopPropagation()}
          >
            <button className="modalClose" type="button" aria-label="닫기" onClick={() => setIsDbDedupeOpen(false)}>
              ×
            </button>
            <div className="modalTitle">ISBN 동일 상품 선택</div>
            <div className="modalBody">
              <p className="dedupeIntro">
                동일한 <b>ISBN</b>을 가진 상품이 있습니다. 각 ISBN 그룹에서 <b>저장할 상품 1개</b>를 선택하세요.
                (중복 없는 상품은 자동으로 저장 대상입니다.)
              </p>

              <div className="dedupeGrid">
                <div className="dedupePanel">
                  <div className="dedupePanelTitle">ISBN 동일 상품</div>
                  <div className="dedupePanelBody">
                    {dbDupGroups.map((g) => {
                      const selected = dbChosenByIsbn[g.isbn];
                      const selectedKey = selected ? getDbRowKey(selected) : null;
                      const candidates = g.rows.filter((r) => getDbRowKey(r) !== selectedKey);
                      const isLocked = !!selected;
                      return (
                        <div key={g.isbn} className="dedupeGroup">
                          <div className="dedupeGroupHead">
                            <div className="dedupeGroupIsbn">ISBN {g.isbn}</div>
                            <div className="dedupeGroupMeta">{g.rows.length}개 중 1개 선택</div>
                          </div>
                          <div className="dedupeList">
                            {candidates.map((r) => {
                              const key = getDbRowKey(r);
                              const { productId, productCode, status } = getDbRowMeta(r);
                              return (
                                <div key={key} className="dedupeItem">
                                  <div className="dedupeItemText">
                                    <div className="dedupeItemTitle" title={getDbRowTitle(r)}>
                                      {getDbRowTitle(r)}
                                    </div>
                                    <div className="dedupeItemMeta">
                                      <div className="dedupeItemMetaLine">· 스마트스토어 상품 번호: {productId || '-'}</div>
                                      <div className="dedupeItemMetaLine">· 판매자 상품 번호: {productCode || '-'}</div>
                                      {status ? <div className="dedupeItemMetaLine">· 판매상태: {status}</div> : null}
                                    </div>
                                  </div>
                                  <button
                                    className="dedupeMoveBtn"
                                    type="button"
                                    onClick={() => chooseDbRow(g.isbn, r)}
                                    aria-label="저장할 상품으로 이동"
                                    disabled={isLocked}
                                    title={isLocked ? '이미 선택된 상품이 있습니다. 오른쪽에서 ←로 되돌린 뒤 선택하세요.' : '저장할 상품으로 이동'}
                                  >
                                    →
                                  </button>
                                </div>
                              );
                            })}
                            {candidates.length === 0 ? <div className="dedupeEmpty">선택할 항목 없음</div> : null}
                            {isLocked ? <div className="dedupeHint">오른쪽에 이미 선택됨 · ←로 되돌린 뒤 다른 상품을 선택하세요.</div> : null}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="dedupePanel">
                  <div className="dedupePanelTitle">저장할 상품</div>
                  <div className="dedupePanelBody">
                    {dbDupGroups.map((g) => {
                      const selected = dbChosenByIsbn[g.isbn];
                      return (
                        <div key={g.isbn} className="dedupeGroup">
                          <div className="dedupeGroupHead">
                            <div className="dedupeGroupIsbn">ISBN {g.isbn}</div>
                            <div className="dedupeGroupMeta">{selected ? '선택됨' : '선택 필요'}</div>
                          </div>
                          {selected ? (
                            <div className="dedupeItem dedupeItem--selected">
                              <div className="dedupeItemText">
                                <div className="dedupeItemTitle" title={getDbRowTitle(selected)}>
                                  {getDbRowTitle(selected)}
                                </div>
                                <div className="dedupeItemMeta">
                                  {(() => {
                                    const { productId, productCode, status } = getDbRowMeta(selected);
                                    return (
                                      <>
                                        <div className="dedupeItemMetaLine">· 스마트스토어 상품 번호: {productId || '-'}</div>
                                        <div className="dedupeItemMetaLine">· 판매자 상품 번호: {productCode || '-'}</div>
                                        {status ? <div className="dedupeItemMetaLine">· 판매상태: {status}</div> : null}
                                      </>
                                    );
                                  })()}
                                </div>
                              </div>
                              <button className="dedupeMoveBtn dedupeMoveBtn--back" type="button" onClick={() => unchooseDbRow(g.isbn)} aria-label="선택 해제">
                                ←
                              </button>
                            </div>
                          ) : (
                            <div className="dedupeEmpty">오른쪽으로 이동해서 1개 선택</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              <div className="dedupeActions">
                <button className="btn btn--secondary" type="button" onClick={() => setIsDbDedupeOpen(false)}>
                  취소
                </button>
                <button
                  className="btn btn--primary"
                  type="button"
                  onClick={confirmDbDedupeAndSave}
                  disabled={dbDupGroups.some((g) => !dbChosenByIsbn[g.isbn])}
                >
                  저장
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {isInfoOpen ? (
        <div className="modalOverlay" role="presentation" onClick={() => setIsInfoOpen(false)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label="안내"
            onClick={(e) => e.stopPropagation()}
          >
            <button className="modalClose" type="button" aria-label="닫기" onClick={() => setIsInfoOpen(false)}>
              ×
            </button>
            <div className="modalTitle">안내</div>
            <div className="modalBody">
              <p>소호, 프로그램이나 크롤링이 잘못됐을 때,</p>
              <p>꼭 연락해주세요. 빨리 고치고 도와드리겠습니다.</p>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
