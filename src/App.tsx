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

const YEAR_OPTIONS = Array.from({ length: 11 }, (_, i) => 2016 + i).reverse();
const AUTO_NO_BASE = '순번';

type DataPack = { rows: Row[]; columns: string[]; label: string };
type DedupeMode = 'none' | 'isbn_keep_first' | 'isbn_keep_last';

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

/** A 엑셀 "분철 1" 값에서 권수 추출. e.g. "스프링(3권)" → 3, 없으면 null */
function parseBuncheolKwons(value: unknown): number | null {
  if (value == null) return null;
  const s = String(value).trim();
  if (!s) return null;
  const m = s.match(/\((\d+)권\)/);
  return m ? Math.max(1, parseInt(m[1], 10)) : null;
}

/** B 행에서 페이지 수 컬럼 값 숫자로 (페이지 수 / 페이지수 등) */
function getPageCount(row: Row, columns: string[]): number | null {
  const pageCol = columns.find((c) => c === '페이지 수' || c === '페이지수');
  if (!pageCol) return null;
  const v = row[pageCol];
  if (v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : null;
}

/** 권수 N으로 옵션 컬럼 값 생성 (엔터 포함) */
function buildOptionCells(kwons: number) {
  const price = 1500 * kwons;
  return {
    옵션명: '책 스캔 서비스 신청 (무료 체험)\n제본/분철',
    옵션값: `스캔 신청(O),신청 안함(X)\n제본/분철 안함,스프링 제본 (${kwons}권)`,
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
}) {
  const hasProductName = params.columns.includes('상품명');
  const dateCol = params.columns.find((c) => DATE_COLUMN_CANDIDATES.includes(c));

  let out = params.rows;
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

  // 1) 비교 영역
  const [compareA, setCompareA] = useState<File | null>(null);
  const [compareB, setCompareB] = useState<File | null>(null);
  const [compareASheetNames, setCompareASheetNames] = useState<string[]>([]);
  const [compareBSheetNames, setCompareBSheetNames] = useState<string[]>([]);
  const [compareASelectedIndex, setCompareASelectedIndex] = useState(0);
  const [compareBSelectedIndex, setCompareBSelectedIndex] = useState(0);
  const [compareBusy, setCompareBusy] = useState(false);
  const [comparePack, setComparePack] = useState<DataPack | null>(null);

  // 2) 필터 영역
  const [filterFile, setFilterFile] = useState<File | null>(null);
  const [filterSheetNames, setFilterSheetNames] = useState<string[]>([]);
  const [filterSelectedIndex, setFilterSelectedIndex] = useState(0);
  const [filterBusy, setFilterBusy] = useState(false);
  const [filterPack, setFilterPack] = useState<DataPack | null>(null);

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

  // 공통 옵션(1/2에 적용)
  const [excludeReservation, setExcludeReservation] = useState(true);
  const [selectedYears, setSelectedYears] = useState<Set<number>>(() => new Set(YEAR_OPTIONS));
  /** 비교 시 A 분철1 비었을 때 권수 산정: B 페이지수 / 이 값 (0~1000, 기본 300) */
  const [pagesPerBook, setPagesPerBook] = useState(300);

  // 3) 취합 영역
  const [aggregateRows, setAggregateRows] = useState<Row[]>([]);
  const [aggregateColumns, setAggregateColumns] = useState<string[]>([]);
  const [dedupeMode, setDedupeMode] = useState<DedupeMode>('isbn_keep_first');

  const canCompare = !!compareA && !!compareB && !compareBusy;
  const canFilter = !!filterFile && !filterBusy;
  const totalAggregate = aggregateRows.length;
  const aggregateNoLabel = useMemo(() => getAutoNoLabel(aggregateColumns), [aggregateColumns]);

  const selectedYearText = useMemo(() => {
    const years = Array.from(selectedYears).sort((a, b) => b - a);
    if (years.length === 0) return '선택 없음';
    if (years.length === YEAR_OPTIONS.length) return '전체';
    if (years.length <= 4) return years.join(', ');
    return `${years[0]} ~ ${years[years.length - 1]} 외 ${years.length - 2}개`;
  }, [selectedYears]);

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
      });

      setComparePack({
        rows: filtered,
        columns: defaultColumnsFrom(filtered, b.columns),
        label: `B 전체 · 옵션 반영 · ${filtered.length.toLocaleString()}건`,
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
      const out = applyCommonFilters({
        rows: src.rows,
        columns: src.columns,
        excludeReservation,
        selectedYears,
      });

      setFilterPack({
        rows: out,
        columns: defaultColumnsFrom(out, src.columns),
        label: `필터 결과 · ${out.length.toLocaleString()}건`,
      });
    } catch (e) {
      setErrorMessage(e instanceof Error ? e.message : '필터링 중 오류가 발생했습니다.');
    } finally {
      setFilterBusy(false);
    }
  }

  function addToAggregate(pack: DataPack) {
    setErrorMessage(null);
    setAggregateColumns((prev) => unionColumns(prev, pack.columns));

    setAggregateRows((prev) => {
      if (dedupeMode === 'none') return [...prev, ...pack.rows];

      const indexByIsbn = new Map<string, number>();
      prev.forEach((r, i) => {
        const k = normalizeIsbn(r['ISBN']);
        if (k) indexByIsbn.set(k, i);
      });

      const next = [...prev];
      for (const r of pack.rows) {
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

  function downloadAggregate() {
    if (aggregateRows.length === 0) return;
    const noLabel = getAutoNoLabel(aggregateColumns);
    const exportColumns =
      aggregateColumns.includes('판매자 상품코드')
        ? [noLabel, ...aggregateColumns]
        : [noLabel, '판매자 상품코드', ...aggregateColumns];

    if (aggregateRows.length <= EXPORT_CHUNK_SIZE) {
      const rowsForExport = aggregateRows.map((r, i) => ({
        [noLabel]: i + 1,
        ...r,
        '판매자 상품코드': 100001 + i,
      }));
      const today = new Date().toISOString().slice(0, 10);
      downloadRowsAsXlsx({
        rows: rowsForExport,
        columns: exportColumns,
        filename: `passnote_aggregate_${today}.xlsx`,
        sheetName: 'Aggregate',
      });
      return;
    }

    const totalChunks = Math.ceil(aggregateRows.length / EXPORT_CHUNK_SIZE);
    for (let c = 0; c < totalChunks; c++) {
      const start = c * EXPORT_CHUNK_SIZE;
      const end = Math.min(start + EXPORT_CHUNK_SIZE, aggregateRows.length);
      const chunkRows = aggregateRows.slice(start, end);
      const rowsForExport = chunkRows.map((r, i) => ({
        [noLabel]: start + i + 1,
        ...r,
        '판매자 상품코드': start + i + 1,
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

  return (
    <div className="app">
      <header className="top">
        <h1 className="topTitle">PassNote Excel Tool</h1>
        <p className="topSub">브라우저에서만 동작 · 파일은 외부로 전송되지 않습니다.</p>
      </header>

      {errorMessage ? (
        <div className="alert alert--error" role="alert">
          <div className="alertTitle">처리 실패</div>
          <div className="alertBody">{errorMessage}</div>
        </div>
      ) : null}

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
                  <button className="btn btn--secondary btn--sm" type="button" onClick={() => setSelectedYears(new Set(YEAR_OPTIONS))}>
                    전체
                  </button>
                  <button className="btn btn--secondary btn--sm" type="button" onClick={() => setSelectedYears(new Set())}>
                    해제
                  </button>
                </div>
              </div>
              <div className="yearChips" role="group" aria-label="출판연도 선택">
                {YEAR_OPTIONS.map((y) => (
                  <button
                    key={y}
                    type="button"
                    className={['chip', selectedYears.has(y) ? 'chip--active' : ''].filter(Boolean).join(' ')}
                    onClick={() => {
                      setSelectedYears((prev) => {
                        const next = new Set(prev);
                        if (next.has(y)) next.delete(y);
                        else next.add(y);
                        return next;
                      });
                    }}
                  >
                    {y}
                  </button>
                ))}
              </div>
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
          <button className="btn btn--primary" type="button" onClick={downloadAggregate} disabled={totalAggregate === 0}>
            저장 (XLSX)
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

      <footer className="footer">
        <p className="footerText">ISBN · 상품명 · 출판날짜/출간일 컬럼 사용</p>
      </footer>
    </div>
  );
}
