import { useMemo } from 'react';

const PAGE_SIZES = [20, 50, 100];

// Shared pagination — จำลอง PG engine เดิม (20/50/100, reset page 1 on data change)
export default function Pagination({ total, page, onPage, pageSize, onPageSize }) {
  const totalPages = useMemo(() => Math.max(1, Math.ceil(total / pageSize)), [total, pageSize]);
  const pageNumbers = useMemo(() => {
    const winS = Math.max(1, page - 2);
    const winE = Math.min(totalPages, winS + 4);
    const s = Math.max(1, winE - 4);
    const arr = [];
    for (let p = s; p <= winE; p++) arr.push(p);
    return arr;
  }, [page, totalPages]);

  if (total === 0) return null;
  const start = (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, total);

  return (
    <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-3 border-t border-[var(--g100)] bg-[var(--surface2)]">
      <div className="text-[12px] text-[var(--tmuted)]">
        แสดง <strong className="text-[var(--tsub)]">{start}–{end}</strong> จาก <strong className="text-[var(--tsub)]">{total}</strong>
      </div>
      <div className="flex items-center gap-2">
        <select
          value={pageSize}
          onChange={(e) => onPageSize(parseInt(e.target.value))}
          className="h-8 px-2 rounded-lg border border-[var(--g200)] text-[12px] bg-white"
        >
          {PAGE_SIZES.map((s) => (
            <option key={s} value={s}>{s} รายการ/หน้า</option>
          ))}
        </select>
        <div className="flex items-center gap-1">
          <PageBtn disabled={page === 1} onClick={() => onPage(page - 1)}>‹</PageBtn>
          {page > 3 && <PageBtn onClick={() => onPage(1)}>1</PageBtn>}
          {page > 4 && <span className="px-1 text-[12px] text-[var(--tmuted)]">…</span>}
          {pageNumbers.map((p) => (
            <PageBtn key={p} active={p === page} onClick={() => onPage(p)}>{p}</PageBtn>
          ))}
          {page < totalPages - 3 && <span className="px-1 text-[12px] text-[var(--tmuted)]">…</span>}
          {page < totalPages - 2 && <PageBtn onClick={() => onPage(totalPages)}>{totalPages}</PageBtn>}
          <PageBtn disabled={page === totalPages} onClick={() => onPage(page + 1)}>›</PageBtn>
        </div>
      </div>
    </div>
  );
}

function PageBtn({ children, onClick, disabled, active }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`min-w-8 h-8 px-2 rounded-lg text-[12px] font-medium border transition-colors ${
        active
          ? 'bg-[var(--blue)] text-white border-[var(--blue)]'
          : 'bg-white border-[var(--g200)] text-[var(--tsub)] hover:bg-[var(--blue-l)] disabled:opacity-40'
      }`}
    >
      {children}
    </button>
  );
}