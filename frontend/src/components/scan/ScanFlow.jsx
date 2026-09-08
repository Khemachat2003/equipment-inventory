import { useState } from 'react';
import axios from 'axios';
import Icon from '../ui/Icon.jsx';
import StatusBadge from '../ui/StatusBadge.jsx';
import TransferModal from '../TransferModal.jsx';
import ScanModal from './ScanModal.jsx';

// ปุ่มสแกน global (topbar) + flow: สแกน → ระบุอุปกรณ์/ตำแหน่ง → ดูประวัติ / โอนย้าย
export default function ScanFlow() {
  const [scanOpen, setScanOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [single, setSingle] = useState(null);      // เจอตัวเดียว
  const [multiple, setMultiple] = useState(null);  // เจอหลายตัว (เช่น scan รหัส Part)
  const [notFound, setNotFound] = useState('');    // ไม่พบ
  const [history, setHistory] = useState(null);    // ประวัติของ serial ตัวนั้น
  const [transfer, setTransfer] = useState(null);  // เปิด TransferModal

  async function resolve(q) {
    const key = (q || '').trim();
    if (!key) return;
    setBusy(true);
    setHistory(null);
    try {
      const { data } = await axios.get('/api/assets');
      const list = data || [];
      const eq = (s) => (s || '').trim().toLowerCase() === key.toLowerCase();
      let matches = list.filter((a) => eq(a.serialNumber) || eq(a.code) || eq(a.assetId));
      if (matches.length === 0) {
        matches = list.filter((a) => (a.serialNumber || '').trim().toLowerCase().endsWith(key.toLowerCase()));
      }
      if (matches.length === 0) { setSingle(null); setMultiple(null); setNotFound(key); }
      else if (matches.length === 1) { setSingle(matches[0]); setMultiple(null); setNotFound(''); }
      else { setMultiple(matches); setSingle(null); setNotFound(''); }
    } catch (e) {
      console.error('resolve asset', e);
      setSingle(null); setMultiple(null); setNotFound(key);
    } finally {
      setBusy(false);
    }
  }

  function onScanned(q) {
    setScanOpen(false);
    resolve(q);
  }

  function closeAll() {
    setScanOpen(false);
    setSingle(null);
    setMultiple(null);
    setNotFound('');
    setHistory(null);
    setTransfer(null);
  }

  async function openHistory(serial) {
    if (history && history.serial === serial) return setHistory(null);
    try {
      const { data } = await axios.get(`/api/asset-history/${encodeURIComponent(serial)}`);
      setHistory({ serial, logs: data || [] });
    } catch (e) {
      alert('โหลดประวัติไม่ได้');
    }
  }

  function refresh() {
    if (single) resolve(single.serialNumber);
  }

  return (
    <>
      <button
        onClick={() => setScanOpen(true)}
        title="สแกน Barcode / Serial"
        className="flex items-center gap-1.5 h-9 px-3 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold hover:bg-[var(--emerald-d)] transition-colors shrink-0"
      >
        <Icon name="document_scanner" size="sm" />
        <span className="hidden min-[430px]:inline">สแกน</span>
      </button>

      {scanOpen && <ScanModal onClose={() => setScanOpen(false)} onResult={onScanned} />}

      {busy && (
        <Centered>
          <Icon name="hourglass_top" size="lg" className="animate-spin text-[var(--blue)]" />
          <span className="text-[13px] text-[var(--tmuted)]">กำลังค้นหา...</span>
        </Centered>
      )}

      {/* เจอตัวเดียว */}
      {single && !busy && (
        <ResultModal
          asset={single}
          history={history}
          onHistory={() => openHistory(single.serialNumber)}
          onTransfer={() => setTransfer(single)}
          onScanAgain={() => { setHistory(null); setScanOpen(true); }}
          onClose={() => { setSingle(null); setHistory(null); }}
          onResolve={refresh}
        />
      )}

      {/* สแกนแล้วตรงหลายตัว (Part code) */}
      {multiple && !busy && (
        <ModalShell title="พบหลายรายการ" onClose={() => setMultiple(null)}>
          <p className="text-[12px] text-[var(--tmuted)]">บาร์โค้ดนี้ตรงกับอุปกรณ์หลายชิ้น — เลือกชิ้นที่ต้องการ</p>
          <div className="max-h-72 overflow-y-auto space-y-2 mt-3">
            {multiple.map((a) => (
              <button
                key={a.serialNumber + a.assetId}
                onClick={() => { setSingle(a); setMultiple(null); }}
                className="w-full text-left px-3 py-2.5 rounded-lg border border-[var(--g200)] hover:bg-[var(--surface2)] flex items-center justify-between gap-3"
              >
                <div className="min-w-0">
                  <div className="text-[13px] font-semibold truncate">{a.name}</div>
                  <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5">{a.serialNumber}</div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <span className="text-[11px] text-[var(--tmuted)]">{a.siteName}: {a.location}</span>
                  <StatusBadge status={a.status} />
                </div>
              </button>
            ))}
          </div>
        </ModalShell>
      )}

      {/* ไม่พบ */}
      {notFound && !busy && (
        <ModalShell title="ไม่พบอุปกรณ์ในระบบ" onClose={() => setNotFound('')}>
          <div className="flex flex-col items-center gap-2 text-center py-2">
            <span className="w-12 h-12 flex items-center justify-center rounded-full bg-[var(--red-l)] text-[var(--red)]">
              <Icon name="search_off" size="lg" />
            </span>
            <div className="text-[15px] font-bold text-[var(--text)]">"{notFound}"</div>
            <p className="text-[12px] text-[var(--tmuted)]">
              ไม่พบ Serial/Code นี้ในทะเบียนอุปกรณ์ (Asset_List)
              <br />ลองสแกนใหม่ หรือตรวจว่าเป็นรหัสของอุปกรณ์ชุดนี้
            </p>
          </div>
          <div className="flex justify-end gap-2">
            <button onClick={() => setNotFound('')} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px]">ปิด</button>
            <button onClick={() => { setNotFound(''); setScanOpen(true); }} className="px-4 py-2 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold">
              <Icon name="document_scanner" size="sm" /> สแกนอีกครั้ง
            </button>
          </div>
        </ModalShell>
      )}

      {transfer && (
        <TransferModal
          open
          onClose={() => setTransfer(null)}
          onSuccess={() => { setTransfer(null); refresh(); }}
          serial={transfer.serialNumber}
          current={{ status: transfer.status, location: transfer.location, siteName: transfer.siteName, user: transfer.user }}
        />
      )}
    </>
  );
}

function ResultModal({ asset, history, onHistory, onTransfer, onScanAgain, onClose, onResolve }) {
  return (
    <ModalShell title="สแกนสำเร็จ" onClose={onClose} accent="success">
      <div className="rounded-xl border border-[var(--g200)] overflow-hidden">
        <div className="p-3.5 bg-[var(--surface2)] space-y-1.5">
          <div className="flex items-start justify-between gap-3">
            <div>
              <div className="text-[15px] font-bold text-[var(--text)] leading-snug">{asset.name}</div>
              <div className="text-[11px] text-[var(--tmuted)] mt-0.5">{asset.assetId} · {asset.code}</div>
            </div>
            <StatusBadge status={asset.status} />
          </div>
          <div className="pt-1.5">
            <div className="text-[12px] text-[var(--tmuted)]">Serial Number</div>
            <div className="font-mono text-[13px] text-[var(--blue)]">{asset.serialNumber}</div>
          </div>
        </div>
        <div className="p-3.5 space-y-2">
          <LocBlock label="ตำแหน่งปัจจุบัน" icon="place" value={`${asset.siteName} · ${asset.location}`} highlight />
          <div className="grid grid-cols-2 gap-2 text-[12px]">
            <div>
              <div className="text-[var(--tmuted)]">Part Number</div>
              <div className="font-mono text-[var(--text)]">{asset.partNumber || '-'}</div>
            </div>
            <div>
              <div className="text-[var(--tmuted)]">ผู้รับผิดชอบ</div>
              <div className="text-[var(--text)]">{asset.user || '-'}</div>
            </div>
          </div>
        </div>
      </div>

      {history && (
        <div className="mt-3 rounded-xl border border-[var(--g200)] overflow-hidden">
          <div className="flex items-center justify-between px-3.5 py-2.5 bg-[var(--surface2)] border-b border-[var(--g100)]">
            <span className="text-[12px] font-semibold text-[var(--text)]">ประวัติอุปกรณ์นี้ ({history.logs.length})</span>
            <button onClick={onHistory} className="text-[11px] text-[var(--blue)] font-medium">ซ่อน</button>
          </div>
          <div className="max-h-56 overflow-y-auto p-3.5 space-y-3">
            {history.logs.length === 0 && <div className="text-center py-6 text-[12px] text-[var(--tmuted)]">ยังไม่มีประวัติ</div>}
            {history.logs.map((l, i) => {
              let icon = 'chevron_right';
              if (i === 0) icon = 'place';
              else if ((l.action || '').includes('คืน')) icon = 'undo';
              else if ((l.action || '').includes('ซ่อม')) icon = 'build';
              else if ((l.action || '').includes('ลงทะเบียน')) icon = 'add';
              return (
                <div key={i} className="flex gap-2.5 items-start">
                  <span className={`w-6 h-6 flex items-center justify-center rounded-full shrink-0 ${i === 0 ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
                    <Icon name={icon} size="xs" />
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-[12px] font-semibold">{l.action}</span>
                      <span className="text-[10px] text-[var(--tmuted)] whitespace-nowrap">{l.date}</span>
                    </div>
                    <div className="text-[11px] text-[var(--tsub)] mt-0.5">
                      {l.from && l.from !== '-' ? `${l.from} → ` : ''}{l.to}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div className="grid grid-cols-2 gap-2 mt-4">
        <button onClick={onHistory} className="flex items-center justify-center gap-1.5 h-10 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]">
          <Icon name="history" size="sm" /> ประวัติ
        </button>
        <button onClick={onScanAgain} className="flex items-center justify-center gap-1.5 h-10 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]">
          <Icon name="document_scanner" size="sm" /> สแกนอีก
        </button>
      </div>
      <button
        onClick={onTransfer}
        className="w-full mt-2 flex items-center justify-center gap-1.5 h-11 rounded-lg bg-[var(--blue)] text-white text-[14px] font-semibold hover:bg-[var(--blue-d)]"
      >
        <Icon name="local_shipping" size="sm" /> โอนย้าย / คืนอุปกรณ์นี้
      </button>
    </ModalShell>
  );
}

function LocBlock({ icon, label, value, highlight }) {
  return (
    <div className={`flex items-center gap-2 px-3 py-2.5 rounded-lg ${highlight ? 'bg-[var(--emerald-l)] border border-[var(--emerald-b)]' : ''}`}>
      <span className={`flex items-center justify-center w-7 h-7 rounded-lg shrink-0 ${highlight ? 'bg-[var(--emerald)] text-white' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
        <Icon name={icon} size="sm" />
      </span>
      <div className="min-w-0">
        <div className="text-[10px] text-[var(--tmuted)] uppercase tracking-wide">{label}</div>
        <div className="text-[13px] font-bold text-[var(--text)] truncate">{value}</div>
      </div>
    </div>
  );
}

function ModalShell({ title, accent, onClose, children }) {
  const dot = accent === 'success'
    ? 'bg-[var(--emerald-l)] text-[var(--emerald-d)]'
    : 'bg-[var(--blue-l)] text-[var(--blue)]';
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-md bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className={`w-7 h-7 flex items-center justify-center rounded-lg ${dot}`}>
              <Icon name="qr_code_scanner" size="sm" />
            </span>
            <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
          </div>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]">
            <Icon name="close" size="sm" />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Centered({ children }) {
  return (
    <div className="fixed inset-0 z-50 flex flex-col items-center justify-center gap-3 bg-black/40 backdrop-blur-sm">
      <div className="flex flex-col items-center gap-3 px-8 py-6 rounded-2xl bg-white shadow-xl">
        {children}
      </div>
    </div>
  );
}