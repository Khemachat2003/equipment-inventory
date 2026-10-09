// AddDeviceModal — เพิ่มอุปกรณ์ใหม่ (ทะเบียนรายชิ้น + Serial) โฟลว์เดียว ใช้ได้ทุก user
// การเลือก Part: พิมพ์แล้วมี suggestion — เจอ Part เดิมกดเลือก, ไม่เจอกด "สร้าง Part ใหม่"
// Serial สร้างโดย backend (SN-PART-DDMMYYYY-XXXX) + backend ตรวจซ้ำแบบ all-or-nothing
// (preview ในฟอร์มเป็นเพียง "คำเดาก่อนส่ง" — ตัวจริงคืนจาก backend ตอน success)

import { useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import Icon from './ui/Icon.jsx';
import { useBusy, BusyOverlay } from './ui/Busy.jsx';
import { showToast, ToastHost } from './ui/Toast.jsx';

const ain = 'w-full h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)]';

function sanitizePartNumber(v) {
  return (v || '').trim().toUpperCase().replace(/\s+/g, '-');
}

export default function AddDeviceModal({ open, presetName = '', onClose, onDone }) {
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(null); // { partNumber, partName, isNew }
  const [qty, setQty] = useState('1');
  const [advanced, setAdvanced] = useState(false);
  const [status, setStatus] = useState('ใช้งานได้');
  const [siteName, setSiteName] = useState('Intranin');
  const [location, setLocation] = useState('Stock');
  const [user, setUser] = useState('');
  const [sites, setSites] = useState([]);
  const [parts, setParts] = useState([]);
  const [result, setResult] = useState(null); // { serials, added, partCreated }
  const [newPart, setNewPart] = useState(null); // โหมดกรอก Part ใหม่ { partNumber, partName }
  const busy = useBusy();

  useEffect(() => {
    if (!open) return;
    setQuery(presetName || '');
    setSelected(null);
    setQty('1');
    setStatus('ใช้งานได้');
    setSiteName('Intranin');
    setLocation('Stock');
    setUser('');
    setAdvanced(false);
    setResult(null);
    setNewPart(null);
    axios.get('/api/farm-sites').then(({ data }) => setSites(data || [])).catch(() => {});
    axios.get('/api/part-catalog').then(({ data }) => setParts(data || [])).catch(() => {});
  }, [open, presetName]);

  const pn = selected
    ? selected.partNumber
    : newPart
      ? sanitizePartNumber(newPart.partNumber)
      : sanitizePartNumber(query);
  const qtyNum = Math.max(1, parseInt(qty) || 0);

  const suggestions = useMemo(() => {
    if (selected) return null;
    const kk = query.trim().toLowerCase();
    if (!kk) return null;
    const hits = parts
      .filter((p) =>
        (p.partNumber || '').toLowerCase().includes(kk) ||
        (p.partName || '').toLowerCase().includes(kk)
      )
      .slice(0, 6);
    const exact = parts.some((p) => sanitizePartNumber(p.partNumber) === pn && pn);
    return { hits, showCreate: pn.length >= 2 && !exact };
  }, [query, selected, parts, pn]);

  // Serial preview — คำเดาก่อนส่ง (ตัวจริง backend สร้าง + ตรวจซ้ำ all-or-nothing)
  const serialPreview = useMemo(() => {
    if (!pn || !qtyNum) return [];
    const now = new Date();
    const ds = String(now.getDate()).padStart(2, '0') + String(now.getMonth() + 1).padStart(2, '0') + now.getFullYear();
    const prefix = `SN-${pn}-${ds}`;
    return Array.from({ length: Math.min(qtyNum, 4) }, (_, i) => `${prefix}-${String(i + 1).padStart(4, '0')}`);
  }, [pn, qtyNum]);

  async function confirmNewPart() {
    const pnum = sanitizePartNumber(newPart.partNumber);
    const pname = (newPart.partName || '').trim();
    if (!pnum) return showToast('กรอก Part Number ก่อน (ตัวย่อภาษาอังกฤษ เช่น SENWT)', { type: 'warn' });
    if (!/^[A-Z0-9][A-Z0-9._-]*$/.test(pnum)) {
      return showToast('Part Number ใช้ได้เฉพาะ A-Z 0-9 . _ - เท่านั้น (ห้ามช่องว่าง/ภาษาไทย) — เพราะจะถูกใช้สร้าง Serial', { type: 'warn' });
    }
    if (parts.some((p) => sanitizePartNumber(p.partNumber) === pnum)) {
      return showToast('Part Number นี้มีอยู่แล้วในระบบ — กดเลือกจากรายการด้านบนแทน', { type: 'warn' });
    }
    setSelected({ partNumber: pnum, partName: pname || pnum, isNew: true });
    setNewPart(null);
  }

  async function submit() {
    if (!selected) return showToast('เลือกอุปกรณ์จากรายการ หรือกด "สร้าง Part ใหม่" ก่อน', { type: 'warn' });
    if (!qtyNum || qtyNum < 1) return showToast('ระบุจำนวนชิ้นให้ถูกต้อง', { type: 'warn' });
    const partNumber = selected.partNumber;
    const partName = selected.partName || partNumber;
    await busy.run('กำลังเพิ่มอุปกรณ์...', async () => {
      try {
        // ① Part ใหม่ → สร้างใน Part_Catalog ก่อน (bulk-add จะ +total ให้เองเมื่อ Part มีอยู่)
        if (selected.isNew) {
          const r1 = await axios.post('/api/add-part', { partNumber, partName, unit: 'ชิ้น' });
          if (r1.data && r1.data.error) throw new Error('สร้าง Part ไม่สำเร็จ: ' + r1.data.error);
        }
        // ② สร้าง Serial + ทะเบียนรายชิ้น (backend ตรวจซ้ำ all-or-nothing — ซ้ำตัวเดียวไม่เขียนเลย)
        const r2 = await axios.post('/api/bulk-add-asset', {
          partNumber,
          partName,
          qty: qtyNum,
          status,
          siteName,
          location: location.trim() || '-',
          user: user.trim(),
        });
        if (!r2.data.success) throw new Error(r2.data.error || 'เพิ่ม Asset ไม่สำเร็จ');
        setResult({ serials: r2.data.serials || [], added: r2.data.added || qtyNum, partCreated: !!selected.isNew });
      } catch (e) {
        showToast('เกิดข้อผิดพลาด: ' + (e.response?.data?.error || e.message), { type: 'err' });
      }
    });
  }

  function startAnother() {
    // เพิ่มอุปกรณ์ถัดไปทันที — เคลียร์ฟอร์มให้ใหม่ (คง modal เปิดอยู่)
    setResult(null);
    setSelected(null);
    setQuery('');
    setQty('1');
    setStatus('ใช้งานได้');
    setSiteName('Intranin');
    setLocation('Stock');
    setUser('');
    setAdvanced(false);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/40 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="w-full max-w-lg bg-white rounded-2xl shadow-xl my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--g100)]">
          <div className="flex items-center gap-2">
            <span className="flex items-center justify-center w-8 h-8 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]"><Icon name="add" size="sm" /></span>
            <span className="text-[15px] font-bold text-[var(--text)]">เพิ่มอุปกรณ์ใหม่</span>
          </div>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-[var(--tmuted)] hover:bg-[var(--surface2)]"><Icon name="close" size="sm" /></button>
        </div>

        {result ? (
          /* ── สำเร็จ — โชว์ Serial จริงที่ backend สร้าง (ตรวจซ้ำแล้ว) ── */
          <div className="p-5">
            <div className="flex flex-col items-center gap-2 py-3">
              <span className="w-12 h-12 flex items-center justify-center rounded-full bg-[var(--emerald-l)] text-[var(--emerald-d)]"><Icon name="check_circle" size="2xl" /></span>
              <div className="text-[15px] font-bold text-[var(--text)]">เพิ่ม {result.added} ชิ้นสำเร็จ</div>
              {result.partCreated && <div className="text-[12px] text-[var(--tmuted)]">สร้าง Part "{selected.partNumber}" ให้อัตโนมัติแล้ว</div>}
            </div>
            <div className="px-3 py-2.5 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)]">
              <div className="text-[11px] font-bold text-[var(--blue)] mb-1">Serial ที่สร้าง (ตรวจซ้ำโดยระบบแล้ว):</div>
              <div className="max-h-40 overflow-y-auto space-y-0.5">
                {result.serials.map((s) => <div key={s} className="font-mono text-[12px] text-[var(--blue)]">{s}</div>)}
              </div>
            </div>
            <div className="flex justify-end gap-2 mt-4">
              <button
                onClick={startAnother}
                title="เคลียร์ฟอร์มเพื่อเพิ่มอุปกรณ์ถัดไปทันที"
                className="flex items-center gap-1.5 px-4 py-2 rounded-lg border border-[var(--blue-b)] bg-[var(--blue-l)] text-[var(--blue)] text-[13px] font-semibold hover:bg-[var(--blue)] hover:text-white"
              >
                <Icon name="add" size="sm" /> เพิ่มอุปกรณ์อื่นต่อ
              </button>
              <button onClick={() => onDone(result)} className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
                <Icon name="check" size="sm" /> เสร็จ
              </button>
            </div>
          </div>
        ) : (
          /* ── ฟอร์ม ── */
          <>
          <div className="p-5 space-y-3">
            {/* ① อุปกรณ์อะไร — พิมพ์แล้วเลือกจาก suggestion หรือสร้างใหม่ */}
            <div className="space-y-1">
              <label className="block text-[12px] font-medium text-[var(--tsub)]">① อุปกรณ์อะไร? (พิมพ์ชื่อหรือ Part Number)</label>
              {newPart ? (
                /* ฟอร์มสร้าง Part ใหม่ — แยก "ชื่ออุปกรณ์" กับ "Part Number (ตัวย่ออังกฤษ)" ออกจากกัน
                   เพราะ Part Number จะถูกใช้สร้าง Serial (SN-<PART>-...) ต้องเป็น A-Z 0-9 เท่านั้น */
                <div className="p-3 rounded-lg border border-[var(--emerald-b)] bg-[var(--emerald-l)] space-y-2">
                  <div className="text-[12px] font-bold text-[var(--emerald-d)]">สร้าง Part ใหม่ — กรอกให้ครบ 2 ช่อง</div>
                  <div className="space-y-1">
                    <label className="block text-[11px] font-medium text-[var(--tsub)]">ชื่ออุปกรณ์ (โชว์ในระบบ — ภาษาไทยได้)</label>
                    <input
                      value={newPart.partName}
                      onChange={(e) => setNewPart({ ...newPart, partName: e.target.value })}
                      placeholder="เช่น เซนเซอร์วัดน้ำ"
                      className={ain}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[11px] font-medium text-[var(--tsub)]">Part Number (ตัวย่อ — ใช้ใน Serial)</label>
                    <input
                      value={newPart.partNumber}
                      onChange={(e) => setNewPart({ ...newPart, partNumber: e.target.value })}
                      placeholder="เช่น SENWT"
                      className={ain + ' font-mono uppercase'}
                    />
                    <div className="text-[10px] text-[var(--tmuted)]">
                      ตัวอย่าง Serial ที่จะได้: SN-{sanitizePartNumber(newPart.partNumber) || 'PART'}-01102026-0001
                    </div>
                  </div>
                  <div className="flex justify-end gap-2">
                    <button onClick={() => setNewPart(null)} className="px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] text-[var(--tsub)]">ยกเลิก</button>
                    <button onClick={confirmNewPart} className="px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[12px] font-semibold hover:bg-[var(--emerald-d)]">ยืนยัน Part นี้</button>
                  </div>
                </div>
              ) : selected ? (
                <div className="flex items-center justify-between gap-2 px-3 h-9 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)]">
                  <span className="text-[13px] font-semibold text-[var(--blue)] truncate">
                    {selected.partNumber} — {selected.partName}
                    {selected.isNew && <span className="ml-1.5 px-1.5 py-px rounded bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[10px] font-bold">ใหม่</span>}
                  </span>
                  <button onClick={() => { setSelected(null); setQuery(''); }} className="text-[11px] text-[var(--tmuted)] hover:text-[var(--red)] shrink-0">เปลี่ยน</button>
                </div>
              ) : (
                <div className="relative">
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="เช่น เซนเซอร์อุณหภูมิ หรือ SEN.TEMP-RS485"
                    className={ain}
                    autoFocus
                  />
                  {suggestions && !newPart && (
                    <div className="absolute z-10 mt-1 w-full rounded-xl border border-[var(--g200)] bg-white shadow-[var(--sh-md)] overflow-hidden max-h-56 overflow-y-auto">
                      {suggestions.hits.map((p) => (
                        <button
                          key={p.partNumber}
                          onClick={() => setSelected({ partNumber: sanitizePartNumber(p.partNumber), partName: p.partName || p.partNumber, isNew: false })}
                          className="w-full text-left px-3 py-2 hover:bg-[var(--surface2)] flex items-center justify-between gap-2"
                        >
                          <span className="min-w-0">
                            <span className="text-[13px] font-medium text-[var(--text)] block truncate">{p.partName || p.partNumber}</span>
                            <span className="text-[11px] font-mono text-[var(--tmuted)]">{p.partNumber}</span>
                          </span>
                          <span className="text-[11px] text-[var(--tmuted)] shrink-0">มี {p.totalQty || 0} ชิ้น</span>
                        </button>
                      ))}
                      {suggestions.showCreate && (
                        <button
                          onClick={() => setNewPart({ partName: query.trim(), partNumber: '' })}
                          className="w-full text-left px-3 py-2 hover:bg-[var(--emerald-l)] flex items-center gap-2 border-t border-[var(--g100)]"
                        >
                          <Icon name="add_circle" size="sm" />
                          <span className="text-[13px] font-semibold text-[var(--emerald-d)]">สร้าง Part ใหม่ "{query.trim()}" — กรอก Part Number ต่อ</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* ② จำนวนชิ้น */}
            <div className="space-y-1">
              <label className="block text-[12px] font-medium text-[var(--tsub)]">② จะเพิ่มกี่ชิ้น?</label>
              <input type="number" min="1" value={qty} onChange={(e) => setQty(e.target.value)} className={ain} />
            </div>

            {/* ③ Serial preview — คำเดาก่อนส่ง (ตัวจริง backend ตรวจซ้ำให้) */}
            {pn && qtyNum > 0 && (
              <div className="px-3 py-2 rounded-lg bg-[var(--blue-l)] border border-[var(--blue-b)] text-[11px] text-[var(--blue)]">
                <div className="font-bold mb-1">Serial ที่จะได้ (ระบบสร้างให้ + ตรวจซ้ำอัตโนมัติ):</div>
                {serialPreview.map((s) => <div key={s} className="font-mono">{s}</div>)}
                {qtyNum > serialPreview.length && <div>... รวมทั้งหมด {qtyNum} ชิ้น</div>}
                <div className="mt-1 text-[10px] opacity-75">เลขรันต่อท้ายต่อจากที่มีอยู่จริงของวันนี้ — ระบบจะไม่สร้าง Serial ซ้ำเด็ดขาด</div>
              </div>
            )}

            {/* ขั้นสูง — พับเก็บ มี default ให้ครบ ไม่ต้องแตะก็เพิ่มได้ */}
            <button onClick={() => setAdvanced(!advanced)} className="text-[12px] font-semibold text-[var(--blue)] flex items-center gap-1">
              <Icon name="chevron_right" size="xs" className={advanced ? 'rotate-90' : ''} /> รายละเอียดเพิ่ม (สถานะ · ตำแหน่ง · ผู้รับผิดชอบ)
            </button>
            {advanced && (
              <div className="space-y-3 pl-3 border-l-2 border-[var(--g100)]">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">สถานะ</label>
                    <select value={status} onChange={(e) => setStatus(e.target.value)} className={ain}>
                      <option>ใช้งานได้</option><option>สำรอง</option><option>ส่งซ่อม</option><option>ชำรุด/สูญหาย</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">ไซต์งาน</label>
                    <select value={siteName} onChange={(e) => setSiteName(e.target.value)} className={ain}>
                      <option value="Intranin">บริษัท Intranin (คลังกลาง)</option>
                      {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
                    </select>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">จุดติดตั้ง</label>
                    <select value={location} onChange={(e) => setLocation(e.target.value)} className={ain}>
                      <option value="Stock">Stock (คลังกลาง)</option>
                      {sites.filter((s) => s.siteName && s.siteName !== 'Intranin').map((s) => <option key={s.siteId} value={s.siteName}>{s.siteName}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="block text-[12px] font-medium text-[var(--tsub)]">ผู้รับผิดชอบ</label>
                    <input value={user} onChange={(e) => setUser(e.target.value)} placeholder="เว้นว่าง = ชื่อผู้ใช้ของคุณ" className={ain} />
                  </div>
                </div>
              </div>
            )}
          </div>
          <div className="flex justify-end gap-2 px-5 py-4 border-t border-[var(--g100)] bg-[var(--surface2)] rounded-b-2xl">
            <button onClick={onClose} className="px-4 py-2 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ยกเลิก</button>
            <button onClick={submit} disabled={busy.busy || !selected} className="flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-50">
              <Icon name="add" size="sm" /> เพิ่ม {qtyNum || 0} ชิ้น
            </button>
          </div>
          </>
        )}
        <ToastHost />
        <BusyOverlay label={busy.busyLabel} />
      </div>
    </div>
  );
}