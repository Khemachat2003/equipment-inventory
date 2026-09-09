import { useEffect, useMemo, useRef, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import axios from 'axios';
import JsBarcode from 'jsbarcode';
import Icon from '../components/ui/Icon.jsx';

// QrPage — สร้างฉลาก QR/Barcode สำหรับอุปกรณ์ (แปลงจาก public/qr.html เป็น React)
// ฟีเจอร์ครบ: เลือกอุปกรณ์ + presets ขนาด + A4 layout + ดีไซน์ + templates + live preview + PDF/พิมพ์

const MM2IN = 25.4;
const SCREEN_DPI = 96;
const A4W_MM = 210;
const A4H_MM = 297;
const SCALE_PREV = 1.8;

const PRESETS = {
  small: [38, 20],
  medium: [50, 25],
  large: [60, 30],
  rack: [100, 40],
};

const STATUS_MAP = {
  'ใช้งานได้': { label: 'ใช้งาน', dotColor: '#10b981', badgeBg: '#10b981', badgeText: '#fff' },
  'สำรอง': { label: 'สำรอง', dotColor: '#f59e0b', badgeBg: '#f59e0b', badgeText: '#000' },
  'ส่งซ่อม': { label: 'ส่งซ่อม', dotColor: '#ef4444', badgeBg: '#ef4444', badgeText: '#fff' },
  'ชำรุด/สูญหาย': { label: 'ชำรุด', dotColor: '#94a3b8', badgeBg: '#94a3b8', badgeText: '#fff' },
};
const FALLBACK_STATUS = STATUS_MAP['ใช้งานได้'];

const TMPL_KEY = 'ems_label_templates_v2';

const mmToPx = (mm, sc = 1) => (mm / MM2IN) * SCREEN_DPI * sc;

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function isLightColor(hex) {
  const h = String(hex || '#ffffff').replace('#', '');
  if (h.length < 6) return true;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return (r * 299 + g * 587 + b * 114) / 1000 > 128;
}

function bcDataFor(serial) {
  if (!serial) return 'PREVIEW001';
  const p = serial.split('-');
  return p.length >= 2 ? p.slice(-2).join('-') : serial;
}

export default function QrPage() {
  const [searchParams] = useSearchParams();
  const [assets, setAssets] = useState([]);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [selected, setSelected] = useState(new Set());
  const [mode, setMode] = useState('label');
  const [lW, setLW] = useState(50);
  const [lH, setLH] = useState(25);
  const [lMargin, setLMargin] = useState(2);
  const [pm, setPm] = useState(10);
  const [gapX, setGapX] = useState(3);
  const [gapY, setGapY] = useState(3);
  const [bgColor, setBgColor] = useState('#ffffff');
  const [showBadge, setShowBadge] = useState(false);
  const [templates, setTemplates] = useState([]);
  const [tmplName, setTmplName] = useState('');
  const [loaded, setLoaded] = useState(false);

  const nameOf = useMemo(() => {
    const m = new Map();
    assets.forEach((a) => m.set(a.serial, a.name));
    return (s) => m.get(s) || '';
  }, [assets]);
  const statusOf = useMemo(() => {
    const m = new Map();
    assets.forEach((a) => m.set(a.serial, a.status));
    return (s) => m.get(s) || 'ใช้งานได้';
  }, [assets]);

  useEffect(() => {
    let cancel = false;
    axios.get('/api/assets', { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
      .then(({ data }) => {
        if (cancel) return;
        const rows = (data || []).map((a) => ({
          serial: a.serial || a.serialNumber || '',
          name: a.name || a.assetName || '',
          status: a.status || 'ใช้งานได้',
        }));
        setAssets(rows);
        const urlSerial = searchParams.get('serial');
        if (urlSerial) {
          setSelected((prev) => new Set(prev).add(urlSerial));
          setMode('label');
        }
        setLoaded(true);
      })
      .catch(() => {
        if (!cancel) {
          setAssets([
            { serial: 'SN-PCB001-21062025-0001', name: 'อุปกรณ์ควบคุมฟาร์ม A', status: 'ใช้งานได้' },
            { serial: 'SN-PCB001-21062025-0002', name: 'อุปกรณ์ควบคุมฟาร์ม B', status: 'ใช้งานได้' },
            { serial: 'SN-SENS-21062025-0001', name: 'เซนเซอร์อุณหภูมิ', status: 'ใช้งานได้' },
            { serial: 'SN-SENS-21062025-0002', name: 'เซนเซอร์ความชื้น', status: 'สำรอง' },
            { serial: 'SN-GATE-21062025-0001', name: 'Gateway IoT หลัก', status: 'ใช้งานได้' },
          ]);
          setLoaded(true);
        }
      });
    return () => { cancel = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    try { setTemplates(JSON.parse(localStorage.getItem(TMPL_KEY) || '[]')); } catch { setTemplates([]); }
  }, []);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return assets.filter((a) => {
      const okStatus = statusFilter === 'all' || a.status === statusFilter;
      const okQ = !q || a.serial.toLowerCase().includes(q) || a.name.toLowerCase().includes(q);
      return okStatus && okQ;
    });
  }, [assets, search, statusFilter]);

  const firstSerial = selected.size ? [...selected][0] : null;

  function toggleSn(serial) {
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(serial) ? next.delete(serial) : next.add(serial);
      return next;
    });
  }
  function toggleAll(checked) {
    setSelected((prev) => {
      const next = new Set(prev);
      filtered.forEach((a) => (checked ? next.add(a.serial) : next.delete(a.serial)));
      return next;
    });
  }
  function selectAllFiltered() {
    setSelected((prev) => new Set([...prev, ...filtered.map((a) => a.serial)]));
  }

  function applyPreset(id) {
    if (!PRESETS[id]) return;
    setLW(PRESETS[id][0]);
    setLH(PRESETS[id][1]);
  }

  function saveTemplate() {
    const name = tmplName.trim();
    if (!name) { toast('⚠ กรุณาใส่ชื่อเทมเพลต', 'warn'); return; }
    const cfg = { mode, lW, lH, lMargin, a4PM: pm, a4GX: gapX, a4GY: gapY, bgColor, showBadge };
    const arr = [{ id: Date.now(), name, config: cfg }, ...templates].slice(0, 20);
    setTemplates(arr);
    localStorage.setItem(TMPL_KEY, JSON.stringify(arr));
    setTmplName('');
    toast('บันทึกเทมเพลต: ' + name);
  }
  function loadTemplate(id) {
    const t = templates.find((x) => x.id === id);
    if (!t) return;
    const c = t.config || {};
    if (c.mode) setMode(c.mode);
    if (c.lW) setLW(Number(c.lW));
    if (c.lH) setLH(Number(c.lH));
    if (c.lMargin) setLMargin(Number(c.lMargin));
    if (c.a4PM) setPm(Number(c.a4PM));
    if (c.a4GX) setGapX(Number(c.a4GX));
    if (c.a4GY) setGapY(Number(c.a4GY));
    if (c.bgColor) setBgColor(c.bgColor);
    if (c.showBadge != null) setShowBadge(Boolean(c.showBadge));
    toast('โหลดเทมเพลต: ' + t.name);
  }
  function deleteTemplate(id) {
    const arr = templates.filter((x) => x.id !== id);
    setTemplates(arr);
    localStorage.setItem(TMPL_KEY, JSON.stringify(arr));
    toast('ลบเทมเพลตแล้ว');
  }

  function downloadLabelPdf() {
    const items = [...selected];
    if (!items.length) { toast('⚠ เลือกอุปกรณ์ก่อน', 'warn'); return; }
    const q = new URLSearchParams({
      serial: items[0],
      name: nameOf(items[0]) || '',
      width: lW,
      height: lH,
      margin: lMargin,
      bgColor,
      showBadge: showBadge ? 1 : 0,
      status: statusOf(items[0]),
      theme: 'light',
    });
    window.open('/api/label/pdf?' + q, '_blank');
  }

  async function downloadA4Pdf() {
    if (!selected.size) { toast('⚠ เลือกอุปกรณ์ก่อน', 'warn'); return; }
    const items = [...selected].map((s) => ({ serial: s, name: nameOf(s) || '', status: statusOf(s) }));
    const payload = {
      items,
      labelW: lW,
      labelH: lH,
      labelMargin: lMargin,
      pageMarginTop: pm,
      pageMarginBottom: pm,
      pageMarginLeft: pm,
      pageMarginRight: pm,
      gapX,
      gapY,
      bgColor,
      showBadge,
      theme: 'light',
    };
    try {
      const r = await axios.post('/api/label/a4pdf', payload, { responseType: 'blob' });
      const url = URL.createObjectURL(r.data);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'labels_a4.pdf';
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch {
      toast('เกิดข้อผิดพลาดในการสร้าง PDF', 'err');
    }
  }

  function toast(msg, type = 'ok') {
    const el = document.getElementById('qtoast');
    if (!el) return alert(msg);
    el.textContent = msg;
    el.className = 'qtoast show ' + (type === 'err' ? 'q-err' : type === 'warn' ? 'q-warn' : '');
    setTimeout(() => { el.className = 'qtoast'; }, 2500);
  }

  return (
    <div>
      {/* Page toolbar (convention: header แสดง title อยู่แล้ว, page มีแค่ action row) */}
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <span className="text-[12px] text-[var(--tmuted)]">สร้างฉลาก Barcode สำหรับอุปกรณ์ — รองรับทั้งพิมพ์เดี่ยวและ A4 หลายดวงต่อแผ่น</span>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-[1fr_300px] gap-4 items-start">
        {/* ── LEFT / เลือกอุปกรณ์ ── */}
        <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
          <div className="flex items-center gap-2 px-5 py-3.5 border-b border-[var(--g100)]">
            <span className="text-[var(--blue)]"><Icon name="assignment" /></span>
            <h2 className="text-[15px] font-bold text-[var(--text)]">เลือกอุปกรณ์เพื่อพิมพ์</h2>
            <span className="ml-auto px-2.5 py-0.5 rounded-full bg-[var(--blue-l)] text-[var(--blue)] text-[12px] font-semibold">{selected.size} รายการ</span>
            <span className="px-2.5 py-0.5 rounded-full bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[12px] font-semibold">{assets.length} ทั้งหมด</span>
          </div>

          <div className="p-4 space-y-3">
            <div className="relative">
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]"><Icon name="search" size="sm" /></span>
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="ค้นหา Serial / ชื่ออุปกรณ์..."
                className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
              />
            </div>

            <div className="flex flex-wrap gap-1.5">
              {[['all', 'ทั้งหมด'], ['ใช้งานได้', '● ใช้งาน'], ['สำรอง', '● สำรอง'], ['ส่งซ่อม', '● ส่งซ่อม'], ['ชำรุด/สูญหาย', '● ชำรุด']].map(([k, lab]) => (
                <button
                  key={k}
                  onClick={() => setStatusFilter(k)}
                  className={`px-2.5 py-1 rounded-full text-[12px] font-medium border transition-colors ${
                    statusFilter === k
                      ? 'bg-[var(--blue)] text-white border-[var(--blue)]'
                      : 'bg-white text-[var(--tsub)] border-[var(--g200)] hover:bg-[var(--surface2)]'
                  }`}
                >
                  {lab}
                </button>
              ))}
            </div>

            <div className="overflow-x-auto rounded-xl border border-[var(--g100)]">
              <table className="w-full text-[13px]">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wide text-[var(--tmuted)] border-b border-[var(--g100)]">
                    <th className="px-3 py-2 w-9">
                      <input
                        type="checkbox"
                        checked={filtered.length > 0 && filtered.every((a) => selected.has(a.serial))}
                        onChange={(e) => toggleAll(e.target.checked)}
                        className="accent-[var(--blue)]"
                      />
                    </th>
                    <th className="px-3 py-2">Serial Number</th>
                    <th className="px-3 py-2">ชื่ออุปกรณ์</th>
                    <th className="px-3 py-2">สถานะ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[var(--g100)]">
                  {!loaded ? (
                    <tr><td colSpan="4" className="px-3 py-10 text-center text-[var(--tmuted)]">⏳ กำลังโหลด...</td></tr>
                  ) : filtered.length === 0 ? (
                    <tr><td colSpan="4" className="px-3 py-10 text-center text-[var(--tmuted)]">ไม่พบข้อมูล</td></tr>
                  ) : filtered.map((a) => {
                    const st = STATUS_MAP[a.status] || FALLBACK_STATUS;
                    const isSel = selected.has(a.serial);
                    return (
                      <tr key={a.serial} onClick={() => toggleSn(a.serial)} className={`cursor-pointer ${isSel ? 'bg-[var(--blue-l)]/40' : 'hover:bg-[var(--surface2)]'}`}>
                        <td className="px-3 py-2">
                          <input type="checkbox" checked={isSel} onChange={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()} className="accent-[var(--blue)]" />
                        </td>
                        <td className="px-3 py-2 font-mono text-[12px] text-[var(--text)]">{a.serial}</td>
                        <td className="px-3 py-2 text-[var(--tsub)]">{a.name}</td>
                        <td className="px-3 py-2">
                          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-[var(--surface2)] border border-[var(--g200)]">
                            <span className="w-2 h-2 rounded-full" style={{ background: st.dotColor }} />
                            {st.label}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="flex items-center justify-between text-[11.5px] text-[var(--tmuted)]">
              <span>{filtered.length} รายการ</span>
              <button onClick={selectAllFiltered} className="text-[var(--blue)] font-semibold hover:underline">เลือกทั้งหมดในหน้า</button>
            </div>

            <div className="flex flex-wrap gap-1.5 min-h-[30px] items-center">
              {!selected.size ? (
                <span className="text-[12px] text-[var(--tmuted)]">ยังไม่ได้เลือกอุปกรณ์</span>
              ) : [...selected].map((s) => (
                <span key={s} className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-[var(--blue-l)] border border-[var(--blue-l)] text-[var(--blue)] text-[12px] font-mono">
                  {s}
                  <button onClick={() => toggleSn(s)} title="ยกเลิก" className="text-[var(--blue)] hover:text-[var(--red)]">
                    <Icon name="close" size="xs" />
                  </button>
                </span>
              ))}
            </div>
          </div>
        </div>

        {/* ── RIGHT / ตัวควบคุม + Preview ── */}
        <div className="space-y-4">
          {/* โหมด */}
          <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
            <div className="text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] mb-2.5">โหมดการพิมพ์</div>
            <div className="grid grid-cols-2 gap-2">
              <ModeTab active={mode === 'label'} onClick={() => setMode('label')} icon="label" title="Label Printer" sub="พิมพ์ทีละใบ" />
              <ModeTab active={mode === 'a4'} onClick={() => setMode('a4')} icon="description" title="A4 Sheet" sub="หลายดวงต่อแผ่น" />
            </div>
          </div>

          {/* ขนาด Label */}
          <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
            <div className="text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] mb-2.5">ขนาด Label</div>
            <div className="grid grid-cols-2 gap-1.5 mb-3">
              {Object.keys(PRESETS).map((id) => (
                <button
                  key={id}
                  onClick={() => applyPreset(id)}
                  className={`text-left px-3 py-2 rounded-lg border text-[12px] font-medium transition-colors ${
                    lW === PRESETS[id][0] && lH === PRESETS[id][1]
                      ? 'bg-[var(--blue-l)] border-[var(--blue)] text-[var(--blue)]'
                      : 'bg-white border-[var(--g200)] text-[var(--tsub)] hover:bg-[var(--surface2)]'
                  }`}
                >
                  <div className="font-semibold capitalize">{id}</div>
                  <div className="text-[11px] opacity-70">{PRESETS[id][0]} × {PRESETS[id][1]} mm</div>
                </button>
              ))}
            </div>
            <div className="h-px bg-[var(--g100)] my-3" />
            <div className="space-y-2.5">
              <div className="grid grid-cols-2 gap-2">
                <NumField label="กว้าง" value={lW} onChange={setLW} min={20} max={200} unit="mm" />
                <NumField label="สูง" value={lH} onChange={setLH} min={12} max={100} unit="mm" />
              </div>
              <NumField label="ขอบ Label" value={lMargin} onChange={setLMargin} min={0.5} max={10} step={0.5} unit="mm" />
            </div>
          </div>

          {/* A4 Layout */}
          {mode === 'a4' && (
            <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
              <div className="text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] mb-2.5">การจัดวาง A4</div>
              <div className="space-y-2.5">
                <NumField label="ขอบกระดาษ (ทุกด้าน)" value={pm} onChange={setPm} min={0} max={25} step={0.5} unit="mm" />
                <div className="grid grid-cols-2 gap-2">
                  <NumField label="ช่องว่าง H" value={gapX} onChange={setGapX} min={0} max={15} step={0.5} unit="mm" />
                  <NumField label="ช่องว่าง V" value={gapY} onChange={setGapY} min={0} max={15} step={0.5} unit="mm" />
                </div>
              </div>
              <div className="mt-3 px-3 py-2 bg-[var(--blue-l)] border border-[var(--blue-l)] rounded-lg text-[12px] text-[var(--blue)]" id="a4info">
                <A4Info items={[...selected]} lw={lW} lh={lH} pm={pm} gapX={gapX} gapY={gapY} />
              </div>
            </div>
          )}

          {/* ดีไซน์ */}
          <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
            <div className="text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] mb-2.5">ดีไซน์</div>
            <div className="text-[12px] font-medium text-[var(--tsub)] mb-1.5">สีพื้นหลัง Label</div>
            <div className="flex items-center gap-2.5 mb-2">
              <input type="color" value={bgColor} onChange={(e) => setBgColor(e.target.value)} className="w-8 h-8 rounded cursor-pointer border border-[var(--g200)]" />
              <span className="text-[12px] font-mono text-[var(--tsub)]">{bgColor}</span>
              <span className={`text-[10.5px] ${isLightColor(bgColor) ? 'text-[var(--tmuted)]' : 'text-[var(--amber-d)]'}`}>
                {isLightColor(bgColor) ? '— บาร์โค้ดปกติ (dark on light)' : 'Dark bg: White patch + dark bars'}
              </span>
            </div>
            <div className="flex flex-wrap gap-1.5 mb-3">
              {['#ffffff', '#f8fafc', '#fef9c3', '#dbeafe', '#dcfce7', '#1e293b', '#0f172a', '#111827'].map((c) => (
                <button
                  key={c}
                  onClick={() => setBgColor(c)}
                  title={c}
                  className={`w-6 h-6 rounded-full border transition-transform ${bgColor.toLowerCase() === c ? 'ring-2 ring-[var(--blue)] scale-110' : 'border-[var(--g200)] hover:scale-105'}`}
                  style={{ background: c }}
                />
              ))}
            </div>
            <div className="h-px bg-[var(--g100)] my-3" />
            <div className="flex items-center justify-between">
              <div>
                <div className="text-[13px] font-medium text-[var(--text)]">แสดง Status Badge บนป้าย</div>
                <div className="text-[11px] text-[var(--tmuted)]">แถบสีสถานะอุปกรณ์ (เปิด/ปิดได้)</div>
              </div>
              <button
                onClick={() => setShowBadge(!showBadge)}
                className={`w-10 h-5.5 rounded-full relative transition-colors ${showBadge ? 'bg-[var(--blue)]' : 'bg-[var(--g300)]'}`}
                aria-pressed={showBadge}
              >
                <span className={`absolute top-0.5 w-4.5 h-4.5 rounded-full bg-white shadow transition-all ${showBadge ? 'left-5' : 'left-0.5'}`} />
              </button>
            </div>
          </div>

          {/* Templates */}
          <div className="bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
            <div className="text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] mb-2.5">เทมเพลตการตั้งค่า</div>
            {!templates.length ? (
              <div className="text-[12px] text-[var(--tmuted)] py-2">ยังไม่มีเทมเพลต</div>
            ) : (
              <div className="space-y-1.5 mb-3">
                {templates.map((t) => (
                  <div key={t.id} className="flex items-center gap-2 px-2.5 py-2 rounded-lg border border-[var(--g100)] cursor-pointer hover:bg-[var(--surface2)]" onClick={() => loadTemplate(t.id)}>
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] font-semibold text-[var(--text)] truncate">{t.name}</div>
                      <div className="text-[11px] text-[var(--tmuted)]">{t.config?.lW}×{t.config?.lH} mm | {t.config?.mode === 'a4' ? 'A4 Sheet' : 'Label'}</div>
                    </div>
                    <button onClick={(e) => { e.stopPropagation(); deleteTemplate(t.id); }} className="p-1 text-[var(--tmuted)] hover:text-[var(--red)]" title="ลบ">
                      <Icon name="close" size="sm" />
                    </button>
                  </div>
                ))}
              </div>
            )}
            <div className="flex gap-2">
              <input
                value={tmplName}
                onChange={(e) => setTmplName(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && saveTemplate()}
                placeholder="ชื่อเทมเพลต..."
                maxLength={40}
                className="flex-1 h-8.5 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[12px] focus:outline-none focus:border-[var(--blue)]"
              />
              <button onClick={saveTemplate} className="flex items-center gap-1 px-3 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)]">
                <Icon name="save" size="sm" /> บันทึก
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ── PREVIEW ── */}
      <div className="mt-4 bg-white rounded-2xl border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="flex items-center justify-between px-5 py-3 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <h3 className="text-[14px] font-bold text-[var(--text)]">Live Preview</h3>
            <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)] animate-pulse" />
          </div>
          <div className="flex gap-2">
            <button onClick={downloadLabelPdf} className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)]">
              <Icon name="description" size="sm" /> PDF
            </button>
            <button onClick={() => window.print()} className="flex items-center gap-1 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[var(--tsub)] text-[12px] font-semibold hover:bg-[var(--surface2)]">
              <Icon name="print" size="sm" /> พิมพ์
            </button>
          </div>
        </div>

        <div className="p-5">
          {mode === 'label' ? (
            <LabelPreview
              serial={firstSerial}
              name={firstSerial ? nameOf(firstSerial) : ''}
              status={firstSerial ? statusOf(firstSerial) : 'ใช้งานได้'}
              lw={lW} lh={lH} lm={lMargin}
              bg={bgColor} showBadge={showBadge}
            />
          ) : (
            <A4Preview
              items={[...selected]}
              nameOf={nameOf} statusOf={statusOf}
              lw={lW} lh={lH} lm={lMargin} pm={pm} gapX={gapX} gapY={gapY}
              bg={bgColor} showBadge={showBadge}
              onDownload={downloadA4Pdf}
              onPrint={printA4}
            />
          )}
        </div>
      </div>

      <div id="qtoast" className="qtoast" />

      {/* Print pool — ซ่อนบนจอ แต่มองเห็นเฉพาะตอนพิมพ์ (visible) */}
      <div id="printroot" aria-hidden="true">
        {mode === 'a4' && selected.size > 0 && (
          <PrintSheet
            items={[...selected].map((s) => ({ serial: s, name: nameOf(s), status: statusOf(s) }))}
            lw={lW} lh={lH} lm={lMargin} pm={pm} gapX={gapX} gapY={gapY}
            bg={bgColor} showBadge={showBadge}
          />
        )}
      </div>

      <style>{`
        .qtoast{position:fixed;left:50%;bottom:28px;transform:translateX(-50%) translateY(20px);background:#0f172a;color:#fff;padding:10px 18px;border-radius:10px;font-size:13px;opacity:0;pointer-events:none;transition:.25s;z-index:999;box-shadow:0 8px 24px rgba(0,0,0,.25)}
        .qtoast.show{opacity:1;transform:translateX(-50%) translateY(0)}
        .qtoast.q-err{border:1px solid rgba(239,68,68,.5)}
        .qtoast.q-warn{border:1px solid rgba(245,158,11,.5)}
        print-sheet{position:relative;background:#fff;overflow:hidden}
        print-sheet .guide{position:absolute;border:1px dashed rgba(59,130,246,.35);pointer-events:none;box-sizing:border-box}
        print-sheet .lcell{position:absolute;overflow:hidden;display:flex;flex-direction:column}
        print-sheet .lcell .bc{display:flex;justify-content:center;align-items:flex-end}
        print-sheet .lcell .lsep{height:1px;flex-shrink:0}
        print-sheet .lcell .ltxt{display:flex;flex-direction:column;align-items:center;justify-content:center;overflow:hidden}
        print-sheet .lcell .lbadge{display:flex;align-items:center;justify-content:center;text-transform:uppercase;letter-spacing:.5px;font-family:'Sarabun',sans-serif}
        #printroot{display:none}
        @media print{
          @page{size:A4;margin:0}
          body>*{visibility:hidden}
          #printroot{display:block;position:absolute;left:0;top:0;width:210mm}
          #printroot,#printroot *{visibility:visible}
        }
      `}</style>
    </div>
  );

  function printA4() {
    if (!selected.size) { toast('⚠ เลือกอุปกรณ์ก่อน', 'warn'); return; }
    window.print();
  }
}

// ── helpers components ──

function ModeTab({ active, onClick, icon, title, sub }) {
  return (
    <button
      onClick={onClick}
      className={`text-left px-3 py-2.5 rounded-xl border transition-colors ${active ? 'bg-[var(--blue-l)] border-[var(--blue)]' : 'bg-white border-[var(--g200)] hover:bg-[var(--surface2)]'}`}
    >
      <div className={`flex items-center gap-1.5 text-[13px] font-semibold ${active ? 'text-[var(--blue)]' : 'text-[var(--text)]'}`}>
        <Icon name={icon} size="sm" /> {title}
      </div>
      <div className="text-[11px] text-[var(--tmuted)] pl-6 mt-0.5">{sub}</div>
    </button>
  );
}

function NumField({ label, value, onChange, min, max, step = 1, unit }) {
  return (
    <label className="block">
      <span className="block text-[12px] font-medium text-[var(--tsub)] mb-1">{label}</span>
      <div className="flex items-center gap-1.5">
        <input
          type="number"
          value={value}
          min={min}
          max={max}
          step={step}
          onChange={(e) => onChange(parseFloat(e.target.value))}
          className="flex-1 h-8.5 px-2.5 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
        />
        <span className="text-[11px] text-[var(--tmuted)]">{unit}</span>
      </div>
    </label>
  );
}

function A4Info({ items, lw, lh, pm, gapX, gapY }) {
  const innerW = A4W_MM - pm * 2;
  const innerH = A4H_MM - pm * 2;
  const cols = Math.max(1, Math.floor((innerW + gapX) / (lw + gapX)));
  const rows = Math.max(1, Math.floor((innerH + gapY) / (lh + gapY)));
  const perPage = cols * rows;
  const pages = Math.max(1, Math.ceil(items.length / perPage));
  return (
    <>จัด <strong>{cols} คอลัมน์ × {rows} แถว = {perPage} ดวง/หน้า</strong> | เลือก {items.length} ดวง → <strong>{pages} หน้า A4</strong></>
  );
}

// ── Label preview (โหมด Label) ──
function LabelPreview({ serial, name, status, lw, lh, lm, bg, showBadge }) {
  const svgRef = useRef(null);
  const isLight = isLightColor(bg);
  const textColor = isLight ? '#0f172a' : '#f1f5f9';
  const textColorDim = isLight ? '#475569' : '#94a3b8';
  const bcFg = isLight ? '#000000' : '#ffffff';
  const bcBg = isLight ? bg : '#ffffff';

  const sc = Math.min(3.4, 480 / mmToPx(lw));
  const pxW = Math.max(80, Math.round(mmToPx(lw, sc)));
  const pxH = Math.round(mmToPx(lh, sc));
  const mPx = mmToPx(lm, sc);

  const innerH = pxH - mPx * 2;
  const badgeH = showBadge ? 12 : 0;
  const bcZoneH = Math.round((innerH - badgeH) * 0.6);
  const txtZoneH = innerH - bcZoneH - 1 - badgeH;

  const bcData = bcDataFor(serial);
  const estMods = bcData.length * 11 + 35;
  const innerW = pxW - mPx * 2;
  const quietZoneMm = Math.max(3, Math.min(6, lw * 0.05));
  const quietPx = Math.round(mmToPx(quietZoneMm, sc));
  const availBcW = innerW - quietPx * 2;
  const bwPx = Math.max(1, Math.floor(availBcW / estMods));
  const barWidthMm = (bwPx / (SCREEN_DPI * sc)) * MM2IN;

  const serialPx = Math.max(8, Math.min(14, Math.round(txtZoneH * 0.46)));
  const namePx = Math.max(7, Math.min(12, Math.round(txtZoneH * 0.34)));

  useEffect(() => {
    if (!svgRef.current) return;
    try {
      JsBarcode(svgRef.current, bcData, {
        format: 'CODE128', width: bwPx, height: Math.round(bcZoneH * 0.75),
        displayValue: false, background: isLight ? bg : '#ffffff', lineColor: bcFg,
        marginLeft: quietPx, marginRight: quietPx, marginTop: Math.round(bcZoneH * 0.08), marginBottom: Math.round(bcZoneH * 0.05),
      });
    } catch (e) { console.error('JsBarcode:', e); }
  }, [bcData, bwPx, bcZoneH, quietPx, isLight, bg, bcFg]);

  const st = STATUS_MAP[status] || FALLBACK_STATUS;

  let hintClass = 'text-[var(--emerald-d)] bg-[var(--emerald-l)] border border-[var(--emerald-l)]';
  let hintMsg = `${lw}×${lh} mm | X-dim ≈ ${barWidthMm.toFixed(2)} mm | Quiet zone ${quietZoneMm.toFixed(1)} mm`;
  if (!isLight) { hintClass = 'text-[var(--amber-d)] bg-[var(--amber-l)] border border-[var(--amber-l)]'; hintMsg += ' | Dark bg: White patch'; }
  else if (barWidthMm < 0.17) { hintClass = 'text-[var(--red)] bg-[var(--red-l)] border border-[var(--red-l)]'; hintMsg = `X-dim ${barWidthMm.toFixed(2)} mm ต่ำกว่ามาตรฐาน — เพิ่มความกว้าง Label`; }
  else if (barWidthMm < 0.20) { hintClass = 'text-[var(--amber-d)] bg-[var(--amber-l)] border border-[var(--amber-l)]'; hintMsg = `X-dim ${barWidthMm.toFixed(2)} mm (แนะนำ ≥0.20) — อาจสแกนยากบางอุปกรณ์`; }

  return (
    <div className="flex flex-col items-center gap-3">
      <div className="rounded-xl bg-[var(--surface2)] p-5 flex items-center justify-center w-full">
        <div style={{ width: pxW, height: pxH, borderRadius: 4, background: bg, overflow: 'hidden', display: 'flex', flexDirection: 'column', boxShadow: '0 2px 10px rgba(0,0,0,.12)' }}>
          <div style={{ height: bcZoneH + mPx, minHeight: bcZoneH + mPx, display: 'flex', justifyContent: 'center', alignItems: 'flex-end', background: isLight ? bg : '#ffffff', padding: `${Math.round(mPx * 0.4)}px ${Math.round(mPx)}px 0` }}>
            <svg ref={svgRef} />
          </div>
          <div style={{ height: 1, background: isLight ? 'rgba(0,0,0,.12)' : 'rgba(255,255,255,.2)', flexShrink: 0 }} />
          <div style={{ height: txtZoneH, minHeight: txtZoneH, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', background: bg, padding: `0 ${Math.round(mPx * 0.5)}px`, gap: Math.max(1, Math.round(txtZoneH * 0.06)) }}>
            <div style={{ fontSize: serialPx, color: textColor, fontFamily: "'IBM Plex Mono',monospace", fontWeight: 600, letterSpacing: '.8px', lineHeight: 1.2, textAlign: 'center', whiteSpace: 'nowrap' }}>
              {serial || 'SN-0000001'}
            </div>
            <div style={{ fontSize: namePx, color: textColorDim, fontFamily: 'Sarabun,sans-serif', lineHeight: 1.2, textAlign: 'center', whiteSpace: 'nowrap', maxWidth: '96%', overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {name || 'ชื่ออุปกรณ์'}
            </div>
          </div>
          {showBadge && (
            <div style={{ background: st.badgeBg, color: st.badgeText, fontSize: 7, fontWeight: 700, padding: '1px 5px', borderRadius: '0 0 4px 4px', width: '100%', textAlign: 'center', letterSpacing: '.5px', flexShrink: 0 }}>
              {st.label.toUpperCase()}
            </div>
          )}
        </div>
      </div>

      <div className={`w-full px-3 py-2 rounded-lg text-[12px] ${hintClass}`}>{hintMsg}</div>

      <div className="w-full grid grid-cols-2 gap-2 text-[12px]">
        <Stat k="ขนาดจริง" v={`${lw} × ${lh} mm`} />
        <Stat k="ขอบ Label" v={`${lm} mm`} />
        <Stat k="X-dimension" v={`${barWidthMm.toFixed(3)} mm`} />
        <Stat k="โหมดสี" v={isLight ? 'Light' : 'Dark (White Patch)'} />
      </div>
    </div>
  );
}

function Stat({ k, v }) {
  return (
    <div className="flex justify-between items-center px-3 py-1.5 rounded-lg bg-[var(--surface2)] border border-[var(--g100)]">
      <span className="text-[var(--tmuted)]">{k}</span>
      <span className="font-semibold text-[var(--text)]">{v}</span>
    </div>
  );
}

// ── A4 preview ──
function A4Preview({ items, nameOf, statusOf, lw, lh, lm, pm, gapX, gapY, bg, showBadge, onDownload, onPrint }) {
  const isLight = isLightColor(bg);
  const innerW = A4W_MM - pm * 2;
  const innerH = A4H_MM - pm * 2;
  const cols = Math.max(1, Math.floor((innerW + gapX) / (lw + gapX)));
  const rows = Math.max(1, Math.floor((innerH + gapY) / (lh + gapY)));
  const perPage = cols * rows;
  const items2 = items.length ? items.map((s) => ({ serial: s, name: nameOf(s), status: statusOf(s) })) : [{ serial: 'SN-PREVIEW-0001', name: 'Preview Label', status: 'ใช้งานได้' }];
  const pages = Math.max(1, Math.ceil(items2.length / perPage));
  const pageItems = items2.slice(0, perPage);
  const SC = SCALE_PREV;

  const cells = pageItems.map((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = pm + col * (lw + gapX);
    const y = pm + row * (lh + gapY);
    const cellW = lw * SC;
    const cellH = lh * SC;
    const mPx = lm * SC;
    const badgeH = showBadge ? 8 : 0;
    const innerCellH = cellH - mPx * 2 - badgeH;
    const bcZoneH = Math.round(innerCellH * 0.6);
    const txtZoneH = innerCellH - bcZoneH - 1;
    const bcData = bcDataFor(item.serial);
    const estMods = bcData.length * 11 + 35;
    const innerCellW = cellW - mPx * 2;
    const quietZoneMm = Math.max(2, lw * 0.05);
    const quietPx = Math.round((quietZoneMm / MM2IN) * SCREEN_DPI * SC);
    const availBcW = innerCellW - quietPx * 2;
    const bwPx = Math.max(1, Math.floor(availBcW / estMods));
    const serialPx = Math.max(4, Math.round(txtZoneH * 0.42));
    const namePx = Math.max(4, Math.round(txtZoneH * 0.32));
    const textColor = isLight ? '#0f172a' : '#f1f5f9';
    const textColorDim = isLight ? '#475569' : '#94a3b8';
    const st = STATUS_MAP[item.status] || FALLBACK_STATUS;
    return { item, x, y, cellW, cellH, mPx, badgeH, bcZoneH, txtZoneH, bcData, bwPx, quietPx, serialPx, namePx, textColor, textColorDim, st };
  });

  return (
    <div className="space-y-3">
      <div className="overflow-x-auto pb-2">
        <div className="mx-auto" style={{ width: A4W_MM * SC }}>
          <div className="print-sheet rounded-xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)]" style={{ width: A4W_MM * SC, height: A4H_MM * SC }}>
            <div className="absolute border border-dashed border-[var(--blue)]/25" style={{ left: pm * SC, top: pm * SC, width: innerW * SC, height: innerH * SC }} />
            {cells.map((c, i) => (
              <div key={c.item.serial + i} className="lcell" style={{ left: c.x * SC, top: c.y * SC, width: c.cellW, height: c.cellH, background: isLight ? bg : bg }}>
                <div className="bc" style={{ height: c.bcZoneH + c.mPx, minHeight: c.bcZoneH + c.mPx, background: isLight ? bg : '#ffffff', padding: `${Math.round(c.mPx * 0.4)}px ${Math.round(c.mPx)}px 0` }}>
                  <BarcodeSvg data={c.bcData} width={c.bwPx} height={Math.round(c.bcZoneH * 0.75)} marginL={c.quietPx} marginR={c.quietPx} marginT={Math.round(c.bcZoneH * 0.08)} marginB={Math.round(c.bcZoneH * 0.05)} bg={isLight ? bg : '#ffffff'} line="#000" />
                </div>
                <div className="lsep" style={{ background: isLight ? 'rgba(0,0,0,.1)' : 'rgba(255,255,255,.2)' }} />
                <div className="ltxt" style={{ height: c.txtZoneH, minHeight: c.txtZoneH, background: bg, padding: '0 2px', gap: Math.max(1, Math.round(c.txtZoneH * 0.06)) }}>
                  <div className="text-center whitespace-nowrap overflow-hidden max-w-[96%] text-ellipsis" style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: c.serialPx, fontWeight: 600, color: c.textColor, letterSpacing: '.4px' }}>{c.item.serial}</div>
                  <div className="text-center whitespace-nowrap overflow-hidden max-w-[96%] text-ellipsis" style={{ fontFamily: 'Sarabun,sans-serif', fontSize: c.namePx, color: c.textColorDim }}>{c.item.name || ''}</div>
                </div>
                {showBadge && (
                  <div className="lbadge" style={{ height: c.badgeH, minHeight: c.badgeH, background: c.st.badgeBg, color: c.st.badgeText, fontSize: 6, fontWeight: 700 }}>{c.st.label}</div>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="flex items-center justify-between flex-wrap gap-2">
        <span className="text-[12px] text-[var(--tmuted)]">แสดงหน้า 1/{pages} ({pageItems.length} ดวงในหน้านี้)</span>
        <div className="flex gap-2">
          <button onClick={onDownload} disabled={!items.length} className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)] disabled:opacity-50 disabled:cursor-not-allowed">
            <Icon name="description" size="sm" /> PDF A4
          </button>
          <button onClick={onPrint} disabled={!items.length} className="flex items-center gap-1 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[var(--tsub)] text-[12px] font-semibold hover:bg-[var(--surface2)] disabled:opacity-50 disabled:cursor-not-allowed">
            <Icon name="print" size="sm" /> พิมพ์
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Print sheet (mm-size, ใช้กับ window.print) ──
function PrintSheet({ items, lw, lh, lm, pm, gapX, gapY, bg, showBadge }) {
  const innerW = A4W_MM - pm * 2;
  const innerH = A4H_MM - pm * 2;
  const cols = Math.max(1, Math.floor((innerW + gapX) / (lw + gapX)));
  const rows = Math.max(1, Math.floor((innerH + gapY) / (lh + gapY)));
  const perPage = cols * rows;
  const pages = Math.max(1, Math.ceil(items.length / perPage));
  const PX = 3.78; // px per mm at 96dpi

  return (
    <>
      {Array.from({ length: pages }).map((_, pg) => {
        const pageItems = items.slice(pg * perPage, (pg + 1) * perPage);
        return (
          <div key={pg} className="print-sheet" style={{ width: A4W_MM * PX, height: A4H_MM * PX, pageBreakAfter: pages > 1 ? 'always' : undefined, background: '#fff', margin: '0 auto' }}>
            {pageItems.map((item, i) => {
              const col = i % cols;
              const row = Math.floor(i / cols);
              const x = pm + col * (lw + gapX);
              const y = pm + row * (lh + gapY);
              const cellW = lw * PX;
              const cellH = lh * PX;
              const mPx = lm * PX;
              const isLight = isLightColor(bg);
              const badgeH = showBadge ? 8 : 0;
              const innerCellH = cellH - mPx * 2 - badgeH;
              const bcZoneH = Math.round(innerCellH * 0.6);
              const txtZoneH = innerCellH - bcZoneH - 1;
              const bcData = bcDataFor(item.serial);
              const estMods = bcData.length * 11 + 35;
              const innerCellW = cellW - mPx * 2;
              const quietPx = Math.max(1, Math.round(((Math.max(2, lw * 0.05)) / MM2IN) * SCREEN_DPI));
              const availBcW = innerCellW - quietPx * 2;
              const bwPx = Math.max(1, Math.floor(availBcW / estMods));
              const textColor = isLight ? '#0f172a' : '#f1f5f9';
              const textColorDim = isLight ? '#475569' : '#94a3b8';
              const st = STATUS_MAP[item.status] || FALLBACK_STATUS;
              return (
                <div key={item.serial + i} className="lcell" style={{ left: x * PX, top: y * PX, width: cellW, height: cellH, background: bg }}>
                  <div className="bc" style={{ height: bcZoneH + mPx, minHeight: bcZoneH + mPx, background: isLight ? bg : '#ffffff', padding: `${Math.round(mPx * 0.4)}px ${Math.round(mPx)}px 0` }}>
                    <BarcodeSvg data={bcData} width={bwPx} height={Math.round(bcZoneH * 0.75)} marginL={quietPx} marginR={quietPx} marginT={Math.round(bcZoneH * 0.08)} marginB={Math.round(bcZoneH * 0.05)} bg={isLight ? bg : '#ffffff'} line={isLight ? '#000' : '#fff'} />
                  </div>
                  <div className="lsep" style={{ background: isLight ? 'rgba(0,0,0,.1)' : 'rgba(255,255,255,.2)' }} />
                  <div className="ltxt" style={{ height: txtZoneH, minHeight: txtZoneH, background: bg, padding: '0 2px', gap: 1 }}>
                    <div className="text-center whitespace-nowrap overflow-hidden max-w-[96%] text-ellipsis" style={{ fontFamily: "'IBM Plex Mono',monospace", fontSize: Math.max(4, Math.round(txtZoneH * 0.42)), fontWeight: 600, color: textColor }}>{item.serial}</div>
                    <div className="text-center whitespace-nowrap overflow-hidden max-w-[96%] text-ellipsis" style={{ fontFamily: 'Sarabun,sans-serif', fontSize: Math.max(4, Math.round(txtZoneH * 0.32)), color: textColorDim }}>{item.name || ''}</div>
                  </div>
                  {showBadge && (
                    <div className="lbadge" style={{ height: badgeH, minHeight: badgeH, background: st.badgeBg, color: st.badgeText, fontSize: 6, fontWeight: 700 }}>{st.label}</div>
                  )}
                </div>
              );
            })}
          </div>
        );
      })}
    </>
  );
}

function BarcodeSvg({ data, width, height, marginL, marginR, marginT, marginB, bg, line }) {
  const ref = useRef(null);
  useEffect(() => {
    if (!ref.current) return;
    try {
      JsBarcode(ref.current, data, { format: 'CODE128', width, height, displayValue: false, background: bg, lineColor: line, marginLeft: marginL, marginRight: marginR, marginTop: marginT, marginBottom: marginB });
    } catch (e) { console.error(e); }
  }, [data, width, height, marginL, marginR, marginT, marginB, bg, line]);
  return <svg ref={ref} />;
}