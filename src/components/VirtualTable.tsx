import { useMemo, useRef } from 'react';
import { useVirtualizer } from '@tanstack/react-virtual';
import type { Row } from '../lib/excel';

type Props = {
  rows: Row[];
  columns: string[];
  showRowNumbers?: boolean;
  rowNumberHeader?: string;
  rowNumberStart?: number;
  height?: number;
  rowHeight?: number;
  stickyHeader?: boolean;
  emptyText?: string;
};

export function VirtualTable({
  rows,
  columns,
  showRowNumbers = false,
  rowNumberHeader = '순번',
  rowNumberStart = 1,
  height = 420,
  rowHeight = 38,
  stickyHeader = true,
  emptyText = '표시할 데이터가 없습니다.',
}: Props) {
  const parentRef = useRef<HTMLDivElement | null>(null);
  const cols = useMemo(() => columns.filter(Boolean), [columns]);
  const gridTemplateColumns = useMemo(() => {
    const noCol = showRowNumbers ? '72px ' : '';
    if (cols.length === 0) return showRowNumbers ? '72px' : '1fr';
    return noCol + cols.map(() => '160px').join(' ');
  }, [cols, showRowNumbers]);

  const virtualizer = useVirtualizer({
    count: rows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => rowHeight,
    overscan: 12,
  });

  if (rows.length === 0) {
    return <div className="tableEmpty">{emptyText}</div>;
  }

  return (
    <div className="tableShell" style={{ height }}>
      <div ref={parentRef} className="tableScroll" role="table">
        <div
          className="tableRow tableRow--head"
          role="row"
          style={{
            position: stickyHeader ? 'sticky' : 'static',
            top: 0,
            zIndex: 2,
            width: 'max-content',
            minWidth: '100%',
            gridTemplateColumns,
          }}
        >
          {showRowNumbers ? (
            <div className="tableCell tableCell--head tableCell--no" role="columnheader" title={rowNumberHeader}>
              {rowNumberHeader}
            </div>
          ) : null}
          {cols.map((c) => (
            <div key={c} className="tableCell tableCell--head" role="columnheader" title={c}>
              {c}
            </div>
          ))}
        </div>

        <div role="rowgroup" style={{ height: virtualizer.getTotalSize(), position: 'relative' }}>
          {virtualizer.getVirtualItems().map((v) => {
            const r = rows[v.index] || {};
            return (
              <div
                key={v.key}
                className="tableRow"
                role="row"
                style={{
                  position: 'absolute',
                  top: 0,
                  left: 0,
                  width: 'max-content',
                  minWidth: '100%',
                  height: v.size,
                  transform: `translateY(${v.start}px)`,
                  gridTemplateColumns,
                }}
              >
                {showRowNumbers ? (
                  <div className="tableCell tableCell--no" role="cell" title={String(rowNumberStart + v.index)}>
                    {rowNumberStart + v.index}
                  </div>
                ) : null}
                {cols.map((c) => (
                  <div key={c} className="tableCell" role="cell" title={String(r[c] ?? '')}>
                    {String(r[c] ?? '')}
                  </div>
                ))}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

