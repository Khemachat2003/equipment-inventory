import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext}?v=3`;
const PAGE_SIZES = [20, 50, 100];

export default function Stock() {
  const [items, setItems] = useState([]);
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  const [returnOpen, setReturnOpen] = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [notice, setNotice] = useState('');

  const load = useCallback(async () => {
    try {
      const { data } = await axios.get('/api/stock');
      setItems(data);
    } catch (e) {
      console.error('load stock', e);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  // Search filter (client-side, matches original filterTable)
  const filtered = useMemo(() => {
    const k = search.toLowerCase();
    if (!k) return items;
    return items.filter((i) => i.code.toLowerCase().includes(k) || i.name.toLowerCase().includes(k));
  }, [items, search]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paged = filtered.slice((page - 1) * pageSize, page * pageSize);

  // Reset page when filter/size changes
  useEffect(() => {
    setPage(1);
  }, [search, pageSize]);

  async function submitTransfer(code, name, qty, type) {
    if (!qty || qty <= 0) return alert('กรอกจำนวนให้ถูกต้อง');
    try {
      const { data } = await axios.post('/api/transfer', { code, name, qty, type });
      if (data.error) alert(data.error);
      else {
        notice(`✔ ${code} ${type} ${qty} ชิ้นแล้ว`);
        await load();
      }
    } catch (e) {
      alert('เกิดข้อผิดพลาด');
    }
  }

  async function editTotal(code, current) {
    const n = prompt('แก้ไขจำนวนทั้งหมด:', current);
    if (n === null) return;
    try {
      const { data } = await axios.post('/api/update-total', { code, newTotal: parseInt(n) });
      if (data.error) alert(data.error);
      else {
        await load();
      }
    } catch (e) {
      alert('เกิดข้อผิดพลาด');
    }
  }

  function flash(msg) {
    setNotice(msg);
    setTimeout(() => setNotice(''), 3000);
  }

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <div className="text-lg font-bold text-[var(--text)]">Stock</div>
          <div className="text-[12px] text-[var(--tmuted)]">รายการอุปกรณ์ทั้งหมด</div>
        </div>
        <button
          onClick={() => setAddOpen(true)}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] transition-colors"
        >
          <Icon name="add_circle" size="sm" /> เพิ่มอุปกรณ์
        </button>
      </div>

      {notice && <div className="px-4 py-2.5 rounded-lg bg-[var(--emerald-l)] text-[var(--emerald-d)] text-[13px]">{notice}</div>}

      {/* Panel */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        {/* Toolbar */}
        <div className="flex items-center gap-3 p-3.5 border-b border-[var(--g100)]">
          <div className="relative max-w-xs flex-1">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--tmuted)]">
              <Icon name="search" size="sm" />
            </span>
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="ค้นหารหัสหรือชื่ออุปกรณ์..."
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px] focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
            />
          </div>
          <button
            onClick={() => setReturnOpen(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] font-medium text-[var(--tsub)] hover:bg-[var(--surface2)]"
          >
            <Icon name="replay" size="sm" /> คืนจาก Site
          </button>
        </div>

        {/* Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[var(--tmuted)] border-b border-[var(--g100)] bg-[var(--surface2)]">
                <th className="px-4 py-3 font-medium">รหัส</th>
                <th className="px-4 py-3 font-medium">รูป</th>
                <th className="px-4 py-3 font-medium">ชื่ออุปกรณ์</th>
                <th className="px-4 py-3 font-medium">ทั้งหมด</th>
                <th className="px-4 py-3 font-medium">Office</th>
                <th className="px-4 py-3 font-medium">Site</th>
                <th className="px-4 py-3 font-medium">จัดการ</th>
              </tr>
            </thead>
            <tbody>
              {paged.length === 0 && (
                <tr>
                  <td colSpan={7} className="text-center py-12 text-[var(--tmuted)]">
                    ไม่มีข้อมูล
                  </td>
                </tr>
              )}
              {paged.map((item) => {
                const oq = parseInt(item.office) || 0;
                return (
                  <tr key={item.code} className={`border-b border-[var(--g100)] ${oq < 1 ? 'opacity-60' : ''}`}>
                    <td className="px-4 py-2.5 font-mono font-semibold">{item.code}</td>
                    <td className="px-4 py-2.5">
                      <img
                        src={IMAGE_URL(item.code, item.ext || 'jpg')}
                        width="46"
                        height="46"
                        loading="lazy"
                        onError={(e) => {
                          e.currentTarget.onerror = null;
                          e.currentTarget.src = '/image/noimage.jpg';
                        }}
                        className="object-cover rounded-lg border border-[var(--g200)]"
                        alt={item.name}
                      />
                    </td>
                    <td className="px-4 py-2.5 font-medium text-[var(--text)]">{item.name}</td>
                    <td className="px-4 py-2.5">{item.total}</td>
                    <td className="px-4 py-2.5">
                      {oq < 1 ? <span className="text-[var(--red)] font-semibold">หมด</span> : oq}
                    </td>
                    <td className="px-4 py-2.5">{item.site}</td>
                    <td className="px-4 py-2.5">
                      <RowActions
                        code={item.code}
                        name={item.name}
                        total={item.total}
                        onTransfer={submitTransfer}
                        onEdit={editTotal}
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        {filtered.length > 0 && (
          <div className="flex items-center justify-between gap-3 p-3.5 border-t border-[var(--g100)] text-[12px]">
            <span className="text-[var(--tmuted)]">
              แสดง {Math.min(pageSize, filtered.length)} จาก {filtered.length} รายการ
            </span>
            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                disabled={page <= 1}
                className="px-2.5 py-1 rounded border border-[var(--g300)] disabled:opacity-40 hover:bg-[var(--surface2)]"
              >
                ‹
              </button>
              <span className="px-2 text-[var(--tsub)]">
                {page} / {totalPages}
              </span>
              <button
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                disabled={page >= totalPages}
                className="px-2.5 py-1 rounded border border-[var(--g300)] disabled:opacity-40 hover:bg-[var(--surface2)]"
              >
                ›
              </button>
              <select
                value={pageSize}
                onChange={(e) => setPageSize(Number(e.target.value))}
                className="ml-2 h-7 px-1.5 rounded border border-[var(--g300)] text-[12px]"
              >
                {PAGE_SIZES.map((s) => (
                  <option key={s} value={s}>
                    {s} รายการ/หน้า
                  </option>
                ))}
              </select>
            </div>
          </div>
        )}
      </div>

      {returnOpen && <ReturnModal onClose={() => setReturnOpen(false)} onDone={() => { load(); setReturnOpen(false); }} />}
      {addOpen && <AddModal onClose={() => setAddOpen(false)} onDone={() => { load(); setAddOpen(false); }} />}
    </div>
  );
}

function RowActions({ code, name, total, onTransfer, onEdit }) {
  const [qty, setQty] = useState('');
  const [type, setType] = useState('เบิก');

  return (
    <div className="flex items-center gap-1.5">
      <input
        type="number"
        min="1"
        value={qty}
        onChange={(e) => setQty(e.target.value)}
        placeholder="จำนวน"
        className="w-16 h-8 px-2 rounded border border-[var(--g300)] text-[12px]"
      />
      <select
        value={type}
        onChange={(e) => setType(e.target.value)}
        className="h-8 px-1.5 rounded border border-[var(--g300)] text-[12px]"
      >
        <option value="เบิก">เบิก</option>
        <option value="คืน">คืน</option>
      </select>
      <button
        onClick={() => onTransfer(code, name, Number(qty), type)}
        className="w-8 h-8 rounded-lg bg-[var(--blue)] text-white flex items-center justify-center hover:bg-[var(--blue-d)]"
        title="โอน"
      >
        <Icon name="check" size="sm" />
      </button>
      <button
        onClick={() => onEdit(code, total)}
        className="w-8 h-8 rounded-lg border border-[var(--g300)] text-[var(--tsub)] flex items-center justify-center hover:bg-[var(--surface2)]"
        title="แก้ไข"
      >
        <Icon name="edit" size="sm" />
      </button>
    </div>
  );
}

function ReturnModal({ onClose, onDone }) {
  const [items, setItems] = useState([]);
  const [selected, setSelected] = useState({});
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    (async () => {
      try {
        const { data } = await axios.get('/api/get-site-items');
        setItems(data.items || []);
      } catch (e) {
        alert('โหลดข้อมูลไม่ได้');
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  function toggle(code) {
    setSelected((s) => ({ ...s, [code]: !s[code] }));
  }
  function allToggle() {
    const all = items.every((i) => selected[i.code]);
    setSelected(Object.fromEntries(items.map((i) => [i.code, !all])));
  }

  async function confirm() {
    const toReturn = items.filter((i) => selected[i.code]).map((i) => ({ code: i.code, qty: parseInt(i.qty) }));
    if (!toReturn.length) return alert('กรุณาเลือกรายการ');
    try {
      const { data } = await axios.post('/api/return-selected-site', { items: toReturn });
      if (data.success) onDone();
      else alert('เกิดข้อผิดพลาด');
    } catch (e) {
      alert('เกิดข้อผิดพลาด');
    }
  }

  const selectedCount = items.filter((i) => selected[i.code]).length;

  return (
    <Modal title="คืนจาก Site" onClose={onClose}>
      <div className="flex items-center justify-between mb-3">
        <button onClick={allToggle} className="text-[12px] text-[var(--blue)] font-medium">
          เลือกทั้งหมด / ยกเลิก
        </button>
        <span className="text-[12px] text-[var(--tmuted)]">เลือก {selectedCount} รายการ</span>
      </div>
      <div className="max-h-72 overflow-y-auto space-y-2">
        {loading && <div className="text-center py-8 text-[var(--tmuted)]">กำลังโหลด...</div>}
        {!loading && items.length === 0 && (
          <div className="text-center py-8 text-[var(--tmuted)]">ไม่มีอุปกรณ์ใน Site</div>
        )}
        {items.map((i) => (
          <label
            key={i.code}
            className="flex items-center justify-between gap-2 px-3 py-2.5 rounded-lg border border-[var(--g200)] hover:bg-[var(--surface2)] cursor-pointer"
          >
            <div className="flex items-center gap-2.5">
              <input type="checkbox" checked={!!selected[i.code]} onChange={() => toggle(i.code)} />
              <span className="text-[13px] font-medium">{i.name}</span>
            </div>
            <span className="text-[12px] text-[var(--tmuted)]">{i.qty} ชิ้น</span>
          </label>
        ))}
      </div>
      <div className="mt-4 flex justify-end gap-2">
        <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
          ปิด
        </button>
        <button onClick={confirm} className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold">
          คืนอุปกรณ์
        </button>
      </div>
    </Modal>
  );
}

function AddModal({ onClose, onDone }) {
  const [code, setCode] = useState('');
  const [name, setName] = useState('');
  const [qty, setQty] = useState('');
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);

  async function submit() {
    if (!code || !name || !qty || !file) return alert('กรอกข้อมูลและเลือกรูปให้ครบ');
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['jpg', 'jpeg', 'png'].includes(ext)) return alert('รองรับ JPG/PNG เท่านั้น');
    setBusy(true);
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const up = await (await fetch('/upload-image', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileName: code + '.' + ext, base64: e.target.result }),
          })).json();
          if (!up.success) return alert('อัปโหลดรูปไม่สำเร็จ');
          await axios.post('/api/add-item', { code, name, total: parseInt(qty), office: 0, site: 0, ext });
          alert('เพิ่มอุปกรณ์สำเร็จ');
          onDone();
        } catch (e) {
          alert('เกิดข้อผิดพลาด');
        } finally {
          setBusy(false);
        }
      };
      reader.readAsDataURL(file);
    } catch (e) {
      alert('เกิดข้อผิดพลาด');
      setBusy(false);
    }
  }

  return (
    <Modal title="เพิ่มอุปกรณ์" onClose={onClose}>
      <Field label="รหัสอุปกรณ์" value={code} onChange={setCode} placeholder="เช่น AC-001" />
      <Field label="ชื่ออุปกรณ์" value={name} onChange={setName} />
      <Field label="จำนวน" type="number" value={qty} onChange={setQty} />
      <div className="mb-3">
        <label className="block text-[12px] font-medium text-[var(--tsub)] mb-1">รูปอุปกรณ์ (JPG/PNG)</label>
        <input type="file" accept=".jpg,.jpeg,.png" onChange={(e) => setFile(e.target.files[0])} className="text-[12px]" />
      </div>
      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
          ยกเลิก
        </button>
        <button
          onClick={submit}
          disabled={busy}
          className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60"
        >
          {busy ? 'กำลังเพิ่ม...' : 'เพิ่มอุปกรณ์'}
        </button>
      </div>
    </Modal>
  );
}

function Field({ label, value, onChange, placeholder, type = 'text' }) {
  return (
    <div className="mb-3">
      <label className="block text-[12px] font-medium text-[var(--tsub)] mb-1">{label}</label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
      />
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40">
      <div className="w-full max-w-md rounded-2xl bg-white shadow-[var(--sh-lg)]">
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-[var(--g100)]">
          <div className="text-[14px] font-semibold text-[var(--text)]">{title}</div>
          <button onClick={onClose} className="text-[var(--tmuted)] hover:text-[var(--text)]">
            <Icon name="close" size="md" />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}
