import { useEffect, useRef, useState } from 'react';
import axios from 'axios';
import { BrowserMultiFormatReader } from '@zxing/browser';
import { BarcodeFormat } from '@zxing/library';
import Icon from '../components/ui/Icon.jsx';
import StatusBadge from '../components/ui/StatusBadge.jsx';
import TransferModal from '../components/TransferModal.jsx';

const FORMATS = [
  BarcodeFormat.QR_CODE,
  BarcodeFormat.CODE_128,
  BarcodeFormat.CODE_39,
  BarcodeFormat.DATA_MATRIX,
  BarcodeFormat.EAN_13,
  BarcodeFormat.EAN_8,
  BarcodeFormat.CODABAR,
];

function silenceZxing() {
  if (globalThis.__zxingWarn) return;
  globalThis.__zxingWarn = console.warn.bind(console);
  console.warn = (...args) => {
    const first = typeof args[0] === 'string' ? args[0] : '';
    if (first.startsWith('MultiFormatReader')) return;
    globalThis.__zxingWarn(...args);
  };
}
function restoreConsole() {
  if (!globalThis.__zxingWarn) return;
  console.warn = globalThis.__zxingWarn;
  globalThis.__zxingWarn = null;
}

function extractSerial(q) {
  const t = (q || '').trim();
  if (!t) return '';
  try { const u = new URL(t); const s = u.searchParams.get('serial'); if (s) return s.trim(); } catch (e) {}
  const m = t.match(/serial=([^&]+)/i);
  if (m) return decodeURIComponent(m[1]).trim();
  return t;
}

let capCtx = null, capCanvas = null;
function captureFrame(vid) {
  if (!capCanvas) {
    capCanvas = document.createElement('canvas');
    capCtx = capCanvas.getContext('2d', { willReadFrequently: true });
  }
  capCanvas.width = vid.videoWidth || 320;
  capCanvas.height = vid.videoHeight || 240;
  capCtx.drawImage(vid, 0, 0, capCanvas.width, capCanvas.height);
  return capCanvas;
}

export default function ScanPage() {
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const readerRef = useRef(null);
  const pendingRef = useRef(false);
  const foundRef = useRef(false);

  async function getCameraStream() {
    const c = { video: { facingMode: 'environment', width: { ideal: 1280 } } };
    return navigator.mediaDevices.getUserMedia(c);
  }

  async function startCam() {
    setStarting(true);
    setError('');
    silenceZxing();
    let stream = null;
    try {
      // 1) เปิด MediaStream เอง + attach กับ <video> ของเรา → control preview เต็มที่
      stream = await getCameraStream();
      streamRef.current = stream;
      const vid = videoRef.current;
      vid.srcObject = stream;
      try { await vid.play(); } catch (e) {}

      // 2) ส่ง stream ให้ ZXing decode (ไม่ให้ ZXing จัดการ video/Auto play)
      const reader = new BrowserMultiFormatReader();
      reader.possibleFormats = FORMATS;
      readerRef.current = reader;

      // ZXing decodeFromStream จะ attach video ใหม่ detached → เราเปลี่ยนมาใช้ video ของเราควบคุมเองไม่ได้
      // ใช้ API ต่ำกว่า: reader.sharpens... แต่ decodeFromStream วาดลง video ตัวเดียวกับที่เราส่งไม่ได้
      // => เราใช้ stream + video ของเราเอง และใช้ reader.decodeFromCanvas ต่อเฟรมจาก video ของเราเอง

      const loop = async () => {
        const vid2 = videoRef.current;
        if (!vid2 || !streamRef.current) return;
        if (vid2.readyState >= 2) {
          try {
            const result = reader.decodeFromCanvas(captureFrame(vid2));
            if (result && result.getText && !foundRef.current) {
              const text = result.getText().trim();
              if (text) {
                foundRef.current = true;
                pauseCam();
                resolve(text);
                return;
              }
            }
          } catch (e) { /* ไม่เจอบาร์โค้ดในเฟรมนี้ — ปล่อยผ่าน */ }
        }
        pendingRef.current = requestAnimationFrame(loop);
      };
      pendingRef.current = requestAnimationFrame(loop);
      setStarting(false);
    } catch (e) {
      console.error('scan start error:', e);
      restoreConsole();
      setError('เปิดกล้องไม่ได้ | ตรวจสิทธิ์กล้องแล้วลองใหม่ หรือพิมพ์ Serial ด้านล่างแทน');
      setStarting(false);
    }
  }

  function pauseCam() {
    // หยุด decode loop (แต่ยังค้าง preview ไว้ให้ user เห็นภาพนิ่ง)
    if (pendingRef.current) { cancelAnimationFrame(pendingRef.current); pendingRef.current = null; }
  }

  function stop() {
    pauseCam();
    try { readerRef.current = null; } catch (e) {}
    try { streamRef.current?.getTracks?.().forEach((t) => t.stop()); } catch (e) {}
    const vid = videoRef.current;
    if (vid) { vid.srcObject = null; try { vid.load?.(); } catch (e) {} }
    streamRef.current = null;
    restoreConsole();
  }

  const [starting, setStarting] = useState(true);
  const [camOn, setCamOn] = useState(true);
  const [error, setError] = useState('');
  const [manual, setManual] = useState('');
  const [status, setStatus] = useState('idle'); // idle | searching | found | multiple | notfound
  const [asset, setAsset] = useState(null);
  const [multi, setMulti] = useState(null);
  const [miss, setMiss] = useState('');
  const [history, setHistory] = useState(null);
  const [transfer, setTransfer] = useState(null);

  function reset() {
    foundRef.current = false;
    setStatus('idle');
    setAsset(null);
    setMulti(null);
    setMiss('');
    setHistory(null);
  }

  async function toggleCam() {
    if (camOn) {
      stop();
      setCamOn(false);
      setError('');
    } else {
      setCamOn(true);
      await startCam();
    }
  }

  useEffect(() => {
    startCam();
    return stop;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function resolve(q) {
    const key = extractSerial(q);
    if (!key) return;
    setStatus('searching');
    setAsset(null); setMulti(null); setMiss(''); setHistory(null);
    try {
      const { data } = await axios.get('/api/assets');
      const list = data || [];
      const eq = (s) => (s || '').trim().toLowerCase() === key.toLowerCase();
      let matches = list.filter((a) => eq(a.serialNumber) || eq(a.code) || eq(a.assetId));
      if (!matches.length) matches = list.filter((a) => (a.serialNumber || '').trim().toLowerCase().endsWith(key.toLowerCase()));
      if (!matches.length) { setStatus('notfound'); setMiss(key); }
      else if (matches.length === 1) { setStatus('found'); setAsset(matches[0]); }
      else { setStatus('multiple'); setMulti(matches); }
    } catch (e) {
      console.error('resolve', e);
      setStatus('error');
    }
  }

  async function openHistory(serial) {
    if (history && history.serial === serial) return setHistory(null);
    try {
      const { data } = await axios.get(`/api/asset-history/${encodeURIComponent(serial)}`);
      setHistory({ serial, logs: data || [] });
    } catch (e) { alert('โหลดประวัติไม่ได้'); }
  }

  function refresh() { if (asset) resolve(asset.serialNumber); }

  return (
    <div className="max-w-2xl mx-auto space-y-4">
      {/* กล้องสแกน */}
      <div className="relative rounded-2xl overflow-hidden bg-black shadow-[var(--sh-md)]">
        {camOn ? (
          <video ref={videoRef} className="w-full h-72 sm:h-80 object-cover" muted playsInline />
        ) : (
          <div className="w-full h-72 sm:h-80 flex flex-col items-center justify-center gap-3 text-white/70">
            <Icon name="videocam_off" size="2xl" />
            <div className="text-[14px] font-medium">กล้องปิดอยู่</div>
            <div className="text-[12px] text-white/50 text-center px-8">ใช้ช่อง "พิมพ์ Serial" ด้านล่างกับเครื่องยิงบาร์โค้ด (USB Scanner) ได้เลย</div>
          </div>
        )}
        <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
          <div className={`w-4/5 h-40 border-2 border-white/70 rounded-xl ${camOn ? '' : 'hidden'}`} />
        </div>
        {starting && camOn && (
          <div className="absolute inset-0 flex flex-col items-center justify-center gap-2 text-white text-[13px] bg-black/50">
            <Icon name="hourglass_top" className="animate-spin" />
            กำลังเปิดกล้อง...
          </div>
        )}
      </div>

      <div className="flex flex-wrap items-center justify-center gap-3">
        <button
          onClick={toggleCam}
          className={`flex items-center gap-2 h-10 px-5 rounded-xl font-semibold transition-colors ${
            camOn
              ? 'bg-[var(--red-l)] text-[var(--red)] border border-[var(--red-b)] hover:bg-[var(--red)] hover:text-white'
              : 'bg-[var(--emerald)] text-white border border-[var(--emerald-d)] hover:bg-[var(--emerald-d)]'
          }`}
        >
          <Icon name={camOn ? 'videocam_off' : 'videocam'} size="sm" />
          {camOn ? 'ปิดกล้อง' : 'เปิดกล้อง'}
        </button>
        {status === 'idle' && (
          <button onClick={reset} className="flex items-center gap-2 h-10 px-5 rounded-xl bg-[var(--emerald)] text-white text-[14px] font-semibold hover:bg-[var(--emerald-d)]">
            <Icon name="document_scanner" size="sm" /> พร้อมสแกน — วางบาร์โค้ดในกรอบ
          </button>
        )}
      </div>

      {error && (
        <div className="flex items-start gap-2.5 px-3.5 py-3 rounded-xl bg-[var(--red-l)] border border-[var(--red-b)] text-[13px] text-[var(--red)]">
          <Icon name="videocam_off" size="sm" />
          <div>
            <div className="font-semibold">ไม่สามารถเปิดกล้องได้</div>
            <div className="text-[12px] text-[var(--tsub)]">{error}</div>
          </div>
        </div>
      )}

      {/* พิมพ์เอง */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
        <div className="text-[13px] font-semibold text-[var(--text)] mb-2.5">พิมพ์ Serial / Code และกด Enter</div>
        <div className="flex gap-2">
          <input
            value={manual}
            onChange={(e) => setManual(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') resolve(manual); }}
            placeholder="เช่น SN-SEN.TEMP-RS485-01082026-0001"
            className="flex-1 h-11 px-3 rounded-lg border border-[var(--g300)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
          />
          <button onClick={() => resolve(manual)} disabled={!manual.trim() || status === 'searching'} className="h-11 px-5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-40">
            {status === 'searching' ? 'ค้นหา...' : 'ค้นหา'}
          </button>
        </div>
        <p className="text-[11px] text-[var(--tmuted)] mt-2">รับทั้ง Serial แบบเต็ม ตัวย่อท้าย หรือรองรับ USB Scanner (พิมพ์แล้ว Enter อัตโนมัติ)</p>
      </div>

      {status === 'searching' && (
        <div className="flex flex-col items-center gap-3 py-6 text-[var(--tmuted)]">
          <Icon name="hourglass_top" size="lg" className="animate-spin text-[var(--blue)]" /> กำลังค้นหา...
        </div>
      )}

      {status === 'found' && asset && (
        <ResultCard
          asset={asset}
          history={history}
          onHistory={() => openHistory(asset.serialNumber)}
          onTransfer={() => setTransfer(asset)}
          onScanAgain={() => { reset(); startCam(); }}
          onRefresh={refresh}
        />
      )}

      {status === 'multiple' && multi && (
        <Card title="พบหลายรายการ">
          <p className="text-[12px] text-[var(--tmuted)]">บาร์โค้ดนี้ตรงกับอุปกรณ์หลายชิ้น — เลือกรายการที่ต้องการ</p>
          <div className="mt-3 space-y-2 max-h-80 overflow-y-auto">
            {multi.map((a) => (
              <button key={a.serialNumber + a.assetId} onClick={() => { setAsset(a); setStatus('found'); setMulti(null); }}
                className="w-full text-left px-3.5 py-3 rounded-xl border border-[var(--g200)] hover:bg-[var(--surface2)] flex items-center justify-between gap-3">
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
        </Card>
      )}

      {status === 'notfound' && (
        <Card title="ไม่พบอุปกรณ์ในระบบ">
          <div className="flex flex-col items-center gap-2 text-center py-2">
            <span className="w-12 h-12 flex items-center justify-center rounded-full bg-[var(--red-l)] text-[var(--red)]"><Icon name="search_off" size="lg" /></span>
            <div className="text-[15px] font-bold text-[var(--text)]">"{miss}"</div>
            <p className="text-[12px] text-[var(--tmuted)]">ไม่พบ Serial/Code นี้ในทะเบียนอุปกรณ์</p>
          </div>
          <div className="flex justify-center gap-2 mt-3">
            <button onClick={() => { reset(); startCam(); }} className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold">
              <Icon name="document_scanner" size="sm" /> สแกนอีกครั้ง
            </button>
          </div>
        </Card>
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
    </div>
  );
}

function ResultCard({ asset, history, onHistory, onTransfer, onScanAgain, onRefresh }) {
  return (
    <Card title="สแกนสำเร็จ" accent="success">
      <div className="rounded-xl border border-[var(--g200)] overflow-hidden">
        <div className="p-4 bg-[var(--surface2)] space-y-1.5">
          <div className="flex items-start justify-between gap-3">
            <div>
              <div className="text-[16px] font-bold text-[var(--text)] leading-snug">{asset.name}</div>
              <div className="text-[11px] text-[var(--tmuted)] mt-0.5">{asset.assetId} · {asset.code}</div>
            </div>
            <StatusBadge status={asset.status} />
          </div>
          <div className="pt-1">
            <div className="text-[12px] text-[var(--tmuted)]">Serial Number</div>
            <div className="font-mono text-[14px] text-[var(--blue)]">{asset.serialNumber}</div>
          </div>
        </div>
        <div className="p-4 space-y-2">
          <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-[var(--emerald-l)] border border-[var(--emerald-b)]">
            <span className="flex items-center justify-center w-7 h-7 rounded-lg shrink-0 bg-[var(--emerald)] text-white"><Icon name="place" size="sm" /></span>
            <div className="min-w-0">
              <div className="text-[10px] text-[var(--tmuted)] uppercase tracking-wide">ตำแหน่งปัจจุบัน</div>
              <div className="text-[14px] font-bold text-[var(--text)] truncate">{asset.siteName} · {asset.location}</div>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-2 text-[12px] pt-1">
            <div><div className="text-[var(--tmuted)]">Part Number</div><div className="font-mono text-[var(--text)]">{asset.partNumber || '-'}</div></div>
            <div><div className="text-[var(--tmuted)]">ผู้รับผิดชอบ</div><div className="text-[var(--text)]">{asset.user || '-'}</div></div>
          </div>
        </div>
      </div>

      {history && (
        <div className="mt-3 rounded-xl border border-[var(--g200)] overflow-hidden">
          <div className="flex items-center justify-between px-4 py-2.5 bg-[var(--surface2)] border-b border-[var(--g100)]">
            <span className="text-[12px] font-semibold text-[var(--text)]">ประวัติอุปกรณ์นี้ ({history.logs.length})</span>
            <button onClick={onHistory} className="text-[11px] text-[var(--blue)] font-medium">ซ่อน</button>
          </div>
          <div className="max-h-60 overflow-y-auto p-4 space-y-3">
            {history.logs.length === 0 && <div className="text-center py-6 text-[12px] text-[var(--tmuted)]">ยังไม่มีประวัติ</div>}
            {history.logs.map((l, i) => {
              let icon = 'chevron_right';
              if (i === 0) icon = 'place';
              else if ((l.action || '').includes('คืน')) icon = 'undo';
              else if ((l.action || '').includes('ซ่อม')) icon = 'build';
              else if ((l.action || '').includes('ลงทะเบียน')) icon = 'add';
              return (
                <div key={i} className="flex gap-2.5 items-start">
                  <span className={`w-6 h-6 flex items-center justify-center rounded-full shrink-0 ${i === 0 ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}><Icon name={icon} size="xs" /></span>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-[12px] font-semibold">{l.action}</span>
                      <span className="text-[10px] text-[var(--tmuted)] whitespace-nowrap">{l.date}</span>
                    </div>
                    <div className="text-[11px] text-[var(--tsub)] mt-0.5">{l.from && l.from !== '-' ? `${l.from} → ` : ''}{l.to}</div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div className="grid grid-cols-2 gap-2 mt-4">
        <button onClick={onHistory} className="flex items-center justify-center gap-1.5 h-10 rounded-xl border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="history" size="sm" /> ประวัติ</button>
        <button onClick={onScanAgain} className="flex items-center justify-center gap-1.5 h-10 rounded-xl border border-[var(--g300)] text-[13px] text-[var(--tsub)] hover:bg-[var(--surface2)]"><Icon name="document_scanner" size="sm" /> สแกนอีก</button>
      </div>
      <button onClick={onTransfer} className="w-full mt-2 flex items-center justify-center gap-1.5 h-11 rounded-xl bg-[var(--blue)] text-white text-[14px] font-semibold hover:bg-[var(--blue-d)]">
        <Icon name="local_shipping" size="sm" /> โอนย้าย / คืนอุปกรณ์นี้
      </button>
    </Card>
  );
}

function Card({ title, accent, children }) {
  const dot = accent === 'success' ? 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' : 'bg-[var(--blue-l)] text-[var(--blue)]';
  return (
    <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
      <div className="flex items-center gap-2 mb-4">
        <span className={`w-7 h-7 flex items-center justify-center rounded-lg ${dot}`}><Icon name="qr_code_scanner" size="sm" /></span>
        <span className="text-[15px] font-bold text-[var(--text)]">{title}</span>
      </div>
      {children}
    </div>
  );
}