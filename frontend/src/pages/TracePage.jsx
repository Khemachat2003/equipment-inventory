import { useEffect, useState } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import TransferModal from '../components/TransferModal.jsx';
import { useAuth } from '../context/AuthContext.jsx';

// Public Trace Page — ดูประวัติการเคลื่อนย้ายอุปกรณ์ตาม Serial (ไม่ต้องล็อกอิน)
// route: /trace/:serial  |  /trace  (search)
//
// ประวัติมาจาก /api/public-asset-history/:serial (public endpoint)
// ปุ่ม "โอนย้ายอุปกรณ์" แสดงเฉพาะเมื่อล็อกอิน (API สำหรับโอนย้ายต้องการ session)

export default function TracePage() {
  const { serial: serialParam } = useParams();
  const navigate = useNavigate();
  const { user } = useAuth();

  const [serial, setSerial] = useState(serialParam || '');
  const [searchVal, setSearchVal] = useState(serialParam || '');
  const [data, setData] = useState(null);
  const [status, setStatus] = useState('idle'); // idle | loading | ok | empty | error
  const [transferOpen, setTransferOpen] = useState(false);

  useEffect(() => {
    const s = (serialParam || '').trim();
    setSerial(s);
    setSearchVal(s);
    fetchHistory(s);
  }, [serialParam]);

  async function fetchHistory(s) {
    if (!s) { setStatus('idle'); setData(null); return; }
    setStatus('loading');
    try {
      const { data: rows } = await axios.get(`/api/public-asset-history/${encodeURIComponent(s)}`);
      if (!Array.isArray(rows) || rows.length === 0) {
        setStatus('empty');
        setData(null);
      } else {
        setStatus('ok');
        setData(rows);
      }
    } catch (e) {
      console.error('Trace load error:', e);
      setStatus('error');
      setData(null);
    }
  }

  function go() {
    const v = searchVal.trim();
    if (!v) return;
    navigate(`/trace/${encodeURIComponent(v)}`);
  }

  function onTransferred() {
    setTransferOpen(false);
    fetchHistory(serial);
  }

  if (status === 'idle') {
    return <Shell searchVal={searchVal} setSearchVal={setSearchVal} onGo={go}>
      <SearchHint />
    </Shell>;
  }

  if (status === 'loading') {
    return <Shell searchVal={searchVal} setSearchVal={setSearchVal} onGo={go}>
      <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-md)] p-16 text-center">
        <div className="flex justify-center mb-4">
          <div className="w-9 h-9 border-2 border-[var(--g200)] border-t-[var(--blue)] rounded-full animate-spin" />
        </div>
        <div className="font-bold text-[var(--text)]">กำลังโหลดข้อมูล...</div>
        <div className="text-[13px] text-[var(--tmuted)] mt-1 font-mono">{serial}</div>
      </div>
    </Shell>;
  }

  if (status === 'empty' || status === 'error') {
    return <Shell searchVal={searchVal} setSearchVal={setSearchVal} onGo={go}>
      <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-md)] p-16 text-center">
        <div className="text-[40px] mb-3 opacity-60 text-[var(--tmuted)]">
          <Icon name={status === 'empty' ? 'mail' : 'error'} size="2xl" />
        </div>
        <div className="text-[18px] font-bold text-[var(--text)]">
          {status === 'empty' ? 'ไม่พบข้อมูล' : 'เกิดข้อผิดพลาด'}
        </div>
        <div className="text-[13px] text-[var(--tsub)] mt-1">
          {status === 'empty' ? (
            <>Serial Number <span className="font-mono font-semibold text-[var(--blue)]">{serial}</span> ไม่พบในระบบ<br />กรุณาตรวจสอบและลองใหม่อีกครั้ง</>
          ) : (
            <>ไม่สามารถโหลดข้อมูลได้ กรุณาลองใหม่</>
          )}
        </div>
      </div>
    </Shell>;
  }

  const latest = data[0];
  const oldest = data[data.length - 1];
  const displaySerial = latest.serialNumber && latest.serialNumber !== '-' ? latest.serialNumber : serial;

  return (
    <Shell searchVal={searchVal} setSearchVal={setSearchVal} onGo={go}>
      <AssetCard latest={latest} oldest={oldest} data={data} displaySerial={displaySerial} />

      <div className="flex flex-wrap gap-2 mb-6">
        <button onClick={() => window.print()} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-white border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]">
          <Icon name="print" size="sm" /> พิมพ์รายงาน
        </button>
        <button onClick={() => navigate(-1)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[13px] font-semibold border border-[var(--emerald-d)]/30 hover:bg-[var(--emerald-l)]/70">
          ← ย้อนกลับ
        </button>
        {user && (
          <button onClick={() => setTransferOpen(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
            <Icon name="local_shipping" size="sm" /> โอนย้ายอุปกรณ์
          </button>
        )}
        {!user && (
          <Link to="/login" className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[var(--g200)] text-[var(--tsub)] text-[13px] font-semibold hover:bg-[var(--g300)]">
            <Icon name="lock" size="sm" /> เข้าสู่ระบบเพื่อโอนย้าย
          </Link>
        )}
      </div>

      <div className="flex items-center gap-2 mb-4">
        <span className="text-[15px] font-bold text-[var(--text)] flex items-center gap-2">
          <span className="text-[var(--blue)]"><Icon name="assignment" /></span> ประวัติการเคลื่อนย้าย
        </span>
        <span className="px-2.5 py-0.5 rounded-full bg-[var(--blue-l)] text-[var(--blue)] text-[12px] font-600 border border-[var(--blue-l)]">
          {data.length} รายการ
        </span>
      </div>

      <Timeline data={data} />

      <TransferModal
        open={transferOpen}
        onClose={() => setTransferOpen(false)}
        onSuccess={onTransferred}
        serial={displaySerial}
        current={{ siteName: latest.to || '', location: latest.from || '', user: latest.user || '', status: latest.action || '' }}
      />
    </Shell>
  );
}

// ── Shell: standalone header (ไม่ใช้ Layout เพราะเป็นหน้า public) ──
function Shell({ searchVal, setSearchVal, onGo, children }) {
  return (
    <div className="min-h-screen bg-[var(--bg)]">
      <header className="bg-[var(--blue)] sticky top-0 z-40 shadow-md print:hidden">
        <div className="max-w-3xl mx-auto px-4 py-3 flex flex-col sm:flex-row items-start sm:items-center gap-3">
          <div className="flex items-center gap-2.5">
            <span className="w-9 h-9 flex items-center justify-center rounded-lg bg-white/15 text-white">
              <Icon name="inventory_2" />
            </span>
            <div>
              <div className="text-white text-[15px] font-semibold leading-tight">Equipment Management</div>
              <div className="text-white/60 text-[11px]">Asset Traceability</div>
            </div>
          </div>
          <div className="flex flex-1 gap-2 w-full sm:max-w-sm ml-auto">
            <input
              value={searchVal}
              onChange={(e) => setSearchVal(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && onGo()}
              placeholder="ค้นหา Serial Number..."
              className="flex-1 h-9 px-3 rounded-lg bg-white/15 text-white text-[13px] placeholder:text-white/40 border border-white/25 focus:bg-white/25 focus:outline-none"
            />
            <button onClick={onGo} className="h-9 px-3.5 rounded-lg bg-white text-[var(--blue)] text-[13px] font-semibold hover:bg-[#e0e7ff] flex items-center gap-1.5">
              <Icon name="search" size="sm" /> ค้นหา
            </button>
          </div>
        </div>
      </header>
      <main className="max-w-3xl mx-auto px-4 py-8 pb-16">{children}</main>
    </div>
  );
}

function SearchHint() {
  return (
    <div className="bg-[var(--blue-l)] border border-[var(--blue-l)] rounded-2xl p-8 text-center">
      <div className="text-[36px] mb-3 text-[var(--blue)] opacity-70"><Icon name="search" size="2xl" /></div>
      <p className="text-[var(--blue)] text-[14px] leading-relaxed">
        <strong>ระบุ Serial Number</strong> ในช่องค้นหาด้านบน<br />เพื่อดูประวัติการเคลื่อนย้ายอุปกรณ์ทั้งหมด
      </p>
    </div>
  );
}

function AssetCard({ latest, oldest, data, displaySerial }) {
  // ชื่อ = text หลัง ":" ใน remark ถ้ามี (เหมือนหน้าเดิม)
  const name = latest.remark && latest.remark.includes(':')
    ? latest.remark.split(':').pop().trim()
    : displaySerial;

  return (
    <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-md)] overflow-hidden mb-6 print:shadow-none">
      <div className="bg-gradient-to-r from-[var(--blue)] to-[#1d4ed8] px-6 py-5 flex items-start justify-between gap-4 flex-wrap print:hidden">
        <div>
          <div className="text-white text-[19px] font-bold break-all">{name}</div>
          <div className="text-white/75 text-[13px] font-mono mt-1">S/N : {displaySerial}</div>
        </div>
        <span className="inline-flex items-center gap-1 px-3.5 py-1 rounded-full bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[12px] font-bold whitespace-nowrap border border-[var(--emerald-d)]/40 print:hidden">
          <Icon name="place" size="xs" /> {latest.to || 'ในระบบ'}
        </span>
      </div>
      <div className="border-t border-[var(--g100)] grid grid-cols-2 sm:grid-cols-4 divide-x divide-y sm:divide-y-0 divide-[var(--g100)] bg-white">
        <Meta label="รายการทั้งหมด" value={`${data.length} รายการ`} />
        <Meta label="ลงทะเบียนเมื่อ" value={oldest.date || '-'} />
        <Meta label="อัปเดตล่าสุด" value={latest.date || '-'} />
        <Meta label="ตำแหน่งล่าสุด" value={latest.to || '-'} />
      </div>
    </div>
  );
}

function Meta({ label, value }) {
  return (
    <div className="px-5 py-3.5">
      <div className="text-[11px] uppercase tracking-wide text-[var(--tmuted)] font-semibold mb-1">{label}</div>
      <div className="text-[13px] font-medium text-[var(--text)] break-all">{value}</div>
    </div>
  );
}

function Timeline({ data }) {
  return (
    <div className="space-y-3.5 relative pl-11 before:absolute before:left-[15px] before:top-6 before:bottom-6 before:w-0.5 before:bg-gradient-to-b before:from-[var(--blue)] before:to-[var(--blue-l)]">
      {data.map((item, idx) => (
        <Entry key={idx} item={item} isFirst={idx === 0} />
      ))}
    </div>
  );
}

function Entry({ item, isFirst }) {
  const dotClass = isFirst
    ? 'bg-[var(--blue)] border-[#93c5fd] text-white'
    : dotStyle(item.action);
  const hasRoute = item.from && item.from !== '-';

  return (
    <div className="relative">
      <span className={`absolute -left-11 top-4 w-[26px] h-[26px] rounded-full flex items-center justify-center text-[13px] border-2 shadow-[0_0_0_4px_var(--bg)] z-[1] ${dotClass}`}>
        {isFirst ? <Icon name="place" size="xs" /> : dotIcon(item.action)}
      </span>

      <div className="bg-white border border-[var(--g200)] rounded-xl px-5 py-4 shadow-[var(--sh-sm)] hover:shadow-[var(--sh-md)]">
        <div className="flex items-start justify-between gap-3 flex-wrap mb-2.5">
          <div className="text-[14px] font-bold text-[var(--text)] flex items-center gap-2">
            {item.action || '-'}
            {isFirst && <span className="px-2 py-0.5 rounded-full bg-[var(--blue)] text-white text-[10px] font-bold">ล่าสุด</span>}
          </div>
          <div className="px-2.5 py-0.5 rounded-full bg-[var(--surface2)] border border-[var(--g200)] text-[11px] text-[var(--tmuted)] whitespace-nowrap">
            <Icon name="event" size="xs" /> {item.date || '-'}
          </div>
        </div>

        {hasRoute ? (
          <div className="flex items-center gap-2 flex-wrap mb-2">
            <RouteChip marker="from">{item.from}</RouteChip>
            <span className="text-[var(--tmuted)]"><Icon name="arrow_forward" size="sm" /></span>
            <RouteChip marker="dest">{item.to || '-'}</RouteChip>
          </div>
        ) : item.to && item.to !== '-' ? (
          <div className="flex items-center gap-2 flex-wrap mb-2">
            <RouteChip marker="dest"><Icon name="place" size="xs" /> {item.to}</RouteChip>
          </div>
        ) : null}

        <div className="flex items-center gap-2.5 flex-wrap">
          <span className="w-[22px] h-[22px] rounded-full bg-[var(--blue-l)] text-[var(--blue)] text-[10px] font-bold flex items-center justify-center border border-[var(--blue-l)]">
            {initials(item.user)}
          </span>
          <span className="text-[12px] text-[var(--tsub)]">{item.user || 'System'}</span>
        </div>

        {item.remark && item.remark !== '-' && (
          <div className="mt-2.5 bg-[var(--surface2)] border-l-[3px] border-[var(--blue)] rounded-r-md px-3 py-2">
            <div className="text-[10px] font-semibold uppercase tracking-wide text-[var(--tmuted)] mb-0.5">หมายเหตุ</div>
            <div className="text-[12px] text-[var(--tsub)] whitespace-pre-wrap break-words">{item.remark}</div>
          </div>
        )}
      </div>
    </div>
  );
}

function RouteChip({ marker, children }) {
  return (
    <span className={`px-2.5 py-1 rounded-lg text-[12px] font-medium inline-flex items-center gap-1 border ${marker === 'dest' ? 'bg-[var(--blue-l)] border-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--surface2)] border-[var(--g200)] text-[var(--text)]'}`}>
      {children}
    </span>
  );
}

function dotStyle(action = '') {
  const a = action.toLowerCase();
  if (a.includes('ลงทะเบียน') || a.includes('เพิ่ม')) return 'bg-[var(--emerald-l)] border-[#4ade80] text-[var(--emerald-d)]';
  if (a.includes('ซ่อม')) return 'bg-[var(--red-l)] border-[#f87171] text-[var(--red)]';
  if (a.includes('คืน')) return 'bg-[var(--amber-l)] border-[#fbbf24] text-[var(--amber-d)]';
  return 'bg-[var(--blue-l)] border-[var(--blue)] text-[var(--blue)]';
}

function dotIcon(action = '') {
  const a = action.toLowerCase();
  if (a.includes('ลงทะเบียน') || a.includes('เพิ่ม')) return '✦';
  if (a.includes('ซ่อม')) return <Icon name="build" size="sm" />;
  if (a.includes('คืน')) return '↩';
  return '→';
}

function initials(name = '') {
  const parts = String(name || '').trim().split(' ').filter(Boolean);
  if (!parts.length) return '?';
  return parts.length > 1
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : parts[0].slice(0, 2).toUpperCase();
}