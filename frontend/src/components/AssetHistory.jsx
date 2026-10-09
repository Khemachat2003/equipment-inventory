import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';

export function AssetHistoryModal({ serial, onClose }) {
  const [data, setData] = useState([]);
  const [status, setStatus] = useState('loading');

  useEffect(() => {
    let active = true;
    axios.get(`/api/asset-history/${encodeURIComponent(serial)}`)
      .then(({ data: rows }) => {
        if (!active) return;
        setData(Array.isArray(rows) ? rows : []);
        setStatus('ready');
      })
      .catch(() => { if (active) setStatus('error'); });
    return () => { active = false; };
  }, [serial]);

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-4 backdrop-blur-sm" onClick={onClose}>
      <section role="dialog" aria-modal="true" aria-labelledby="asset-history-title" className="my-8 w-full max-w-2xl overflow-hidden rounded-2xl bg-white shadow-xl" onClick={(e) => e.stopPropagation()}>
        <header className="flex items-center justify-between gap-3 border-b border-[var(--g100)] px-5 py-4">
          <div className="flex min-w-0 items-center gap-2">
            <span className="text-[var(--blue)]"><Icon name="assignment" size="sm" /></span>
            <h2 id="asset-history-title" className="text-[15px] font-bold">ประวัติการเคลื่อนย้าย</h2>
            <span className="truncate rounded-full bg-[var(--blue-l)] px-2 py-0.5 font-mono text-[11px] text-[var(--blue)]">{serial}</span>
            {status === 'ready' && <span className="shrink-0 text-[12px] text-[var(--tmuted)]">{data.length} รายการ</span>}
          </div>
          <div className="flex shrink-0 items-center gap-2">
            <button type="button" onClick={onClose} aria-label="ปิดประวัติ" className="flex h-9 w-9 items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
          </div>
        </header>
        <div className="max-h-[70vh] overflow-y-auto bg-[var(--bg)] p-4 sm:p-5">
          {status === 'loading' && <div className="py-12 text-center text-[13px] text-[var(--tmuted)]">กำลังโหลดประวัติ...</div>}
          {status === 'error' && <div className="py-12 text-center text-[13px] text-[var(--red)]">โหลดประวัติไม่ได้ กรุณาลองใหม่</div>}
          {status === 'ready' && data.length === 0 && <div className="py-12 text-center text-[13px] text-[var(--tmuted)]">ยังไม่มีประวัติในระบบ</div>}
          {status === 'ready' && data.length > 0 && <AssetHistoryTimeline data={data} />}
        </div>
        <footer className="border-t border-[var(--g100)] bg-white p-3 sm:px-5">
          <a href={`/trace/${encodeURIComponent(serial)}`} target="_blank" rel="noreferrer" className="flex min-h-11 w-full items-center justify-center gap-1.5 rounded-lg bg-[var(--blue)] px-4 text-[13px] font-semibold text-white hover:bg-[var(--blue-d)] sm:ml-auto sm:w-auto">
            ดูประวัติเต็ม <Icon name="arrow_forward" size="xs" />
          </a>
        </footer>
      </section>
    </div>
  );
}

export function AssetHistoryTimeline({ data }) {
  return (
    <div className="relative space-y-3.5 pl-11 before:absolute before:bottom-6 before:left-[15px] before:top-6 before:w-0.5 before:bg-gradient-to-b before:from-[var(--blue)] before:to-[var(--blue-l)]">
      {data.map((item, idx) => <HistoryEntry key={`${item.date}-${item.serialNumber || ''}-${idx}`} item={item} isFirst={idx === 0} />)}
    </div>
  );
}

function HistoryEntry({ item, isFirst }) {
  const hasRoute = item.from && item.from !== '-';
  return (
    <article className="relative">
      <span className={`absolute -left-11 top-4 z-[1] flex h-[26px] w-[26px] items-center justify-center rounded-full border-2 text-[13px] shadow-[0_0_0_4px_var(--bg)] ${isFirst ? 'border-[#93c5fd] bg-[var(--blue)] text-white' : dotStyle(item.action)}`}>
        {isFirst ? <Icon name="place" size="xs" /> : <Icon name={dotIcon(item.action)} size="xs" />}
      </span>
      <div className="rounded-xl border border-[var(--g200)] bg-white px-4 py-3.5 shadow-[var(--sh-sm)] sm:px-5 sm:py-4">
        <div className="mb-2.5 flex flex-wrap items-start justify-between gap-2">
          <div className="flex items-center gap-2 text-[14px] font-bold text-[var(--text)]">
            {item.action || '-'}
            {isFirst && <span className="rounded-full bg-[var(--blue)] px-2 py-0.5 text-[10px] font-bold text-white">ล่าสุด</span>}
          </div>
          <div className="whitespace-nowrap rounded-full border border-[var(--g200)] bg-[var(--surface2)] px-2.5 py-0.5 text-[11px] text-[var(--tmuted)]"><Icon name="event" size="xs" /> {item.date || '-'}</div>
        </div>
        {hasRoute ? (
          <div className="mb-2 flex flex-wrap items-center gap-2">
            <RouteChip marker="from">{item.from}</RouteChip><span className="text-[var(--tmuted)]"><Icon name="arrow_forward" size="sm" /></span><RouteChip marker="dest">{item.to || '-'}</RouteChip>
          </div>
        ) : item.to && item.to !== '-' ? <div className="mb-2"><RouteChip marker="dest"><Icon name="place" size="xs" /> {item.to}</RouteChip></div> : null}
        <div className="flex flex-wrap items-center gap-2.5">
          <span className="flex h-[22px] w-[22px] items-center justify-center rounded-full border border-[var(--blue-l)] bg-[var(--blue-l)] text-[10px] font-bold text-[var(--blue)]">{initials(item.user)}</span>
          <span className="text-[12px] text-[var(--tsub)]">{item.user || 'System'}</span>
        </div>
        {item.remark && item.remark !== '-' && <div className="mt-2.5 rounded-r-md border-l-[3px] border-[var(--blue)] bg-[var(--surface2)] px-3 py-2"><div className="mb-0.5 text-[10px] font-semibold uppercase tracking-wide text-[var(--tmuted)]">หมายเหตุ</div><div className="whitespace-pre-wrap break-words text-[12px] text-[var(--tsub)]">{item.remark}</div></div>}
      </div>
    </article>
  );
}

function RouteChip({ marker, children }) {
  return <span className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-1 text-[12px] font-medium ${marker === 'dest' ? 'border-[var(--blue-l)] bg-[var(--blue-l)] text-[var(--blue)]' : 'border-[var(--g200)] bg-[var(--surface2)] text-[var(--text)]'}`}>{children}</span>;
}

function dotStyle(action = '') {
  const text = action.toLowerCase();
  if (text.includes('ลงทะเบียน') || text.includes('เพิ่ม')) return 'border-[#4ade80] bg-[var(--emerald-l)] text-[var(--emerald-d)]';
  if (text.includes('ซ่อม')) return 'border-[#f87171] bg-[var(--red-l)] text-[var(--red)]';
  if (text.includes('คืน')) return 'border-[#fbbf24] bg-[var(--amber-l)] text-[var(--amber-d)]';
  return 'border-[var(--blue)] bg-[var(--blue-l)] text-[var(--blue)]';
}

function dotIcon(action = '') {
  const text = action.toLowerCase();
  if (text.includes('ลงทะเบียน') || text.includes('เพิ่ม')) return 'add';
  if (text.includes('ซ่อม')) return 'build';
  if (text.includes('คืน')) return 'undo';
  return 'arrow_forward';
}

function initials(name = '') {
  const parts = String(name || '').trim().split(' ').filter(Boolean);
  if (!parts.length) return '?';
  return parts.length > 1 ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase() : parts[0].slice(0, 2).toUpperCase();
}
