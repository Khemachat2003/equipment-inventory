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
  const [borrowOpen, setBorrowOpen] = useState(false);
  const [confirmBorrow, setConfirmBorrow] = useState(null);
  const [cart, setCart] = useState({});
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

  // ---------- ตะกร้าเบิก (bulk borrow) ----------
  const itemName = useCallback(
    (code) => (items.find((i) => i.code === code) || {}).name || code,
    [items]
  );
  const cartCount = Object.keys(cart).length;
  const cartTotalQty = Object.values(cart).reduce((a, v) => a + (v.qty || 0), 0);

  function addToCart(code) {
    setCart((prev) => {
      const cur = prev[code]?.qty || 0;
      return { ...prev, [code]: { qty: cur + 1 } };
    });
  }
  function setCartQty(code, qty) {
    const v = Math.max(0, parseInt(qty) || 0);
    setCart((prev) => {
      const next = { ...prev };
      if (v === 0) delete next[code];
      else next[code] = { qty: v };
      return next;
    });
  }
  function removeFromCart(code) {
    setCart((prev) => {
      const next = { ...prev };
      delete next[code];
      return next;
    });
  }
  function clearCart() {
    setCart({});
  }

  async function submitTransfer(code, name, qty, type) {
    if (!qty || qty <= 0) return alert('กรอกจำนวนให้ถูกต้อง');
    try {
      const { data } = await axios.post('/api/transfer', { code, name, qty, type });
      if (data.error) alert(data.error);
      else {
        flash(`${code} ${type} ${qty} ชิ้นแล้ว`);
        await load();
      }
    } catch (e) {
      alert('เกิดข้อผิดพลาด');
    }
  }

  async function confirmBorrowSubmit() {
    const carts = Object.entries(cart).map(([code, v]) => ({ code, qty: v.qty }));
    if (!carts.length) return;
    setBorrowOpen(false);
    try {
      const { data } = await axios.post('/api/borrow-selected', { items: carts });
      if (data.success) {
        const ok = (data.borrowed || []).length;
        const short = data.short || [];
        if (ok) flash(`เบิกสำเร็จ ${ok} รายการ${short.length ? `, ไม่พอ ${short.length} รายการ` : ''}`);
        else flash('ไม่มีรายการที่เบิกได้');
        if (short.length) setConfirmBorrow({ items: short });
        clearCart();
        await load();
      } else {
        flash('เกิดข้อผิดพลาด');
      }
    } catch (e) {
      flash('เกิดข้อผิดพลาด');
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
    setTimeout(() => setNotice(''), 4000);
  }

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-3">
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

      {/* ตะกร้าเบิก */}
      {cartCount > 0 && (
        <div className="rounded-2xl bg-white border border-[var(--emerald)] shadow-[var(--sh-sm)] overflow-hidden">
          <div className="flex items-center gap-2 px-4 py-2.5 bg-[var(--emerald-l)] border-b border-[var(--emerald-b)]">
            <Icon name="add_shopping_cart" size="sm" />
            <span className="text-[13px] font-semibold text-[var(--emerald-d)]">
              ตะกร้าเบิก ({cartCount} รายการ, รวม {cartTotalQty} ชิ้น)
            </span>
            <div className="ml-auto flex items-center gap-2">
              <button
                onClick={clearCart}
                className="text-[12px] text-[var(--tmuted)] hover:text-[var(--red)]"
              >
                เคลียร์
              </button>
              <button
                onClick={() => setBorrowOpen(true)}
                disabled={cartCount === 0}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold hover:bg-[var(--emerald-d)] transition-colors disabled:opacity-50"
              >
                <Icon name="logout" size="sm" /> ยืนยันเบิกทั้งหมด
              </button>
            </div>
          </div>
          <div className="max-h-56 overflow-y-auto divide-y divide-[var(--g100)]">
            {Object.entries(cart).map(([code, v]) => {
              const oq = parseInt((items.find((i) => i.code === code) || {}).office) || 0;
              return (
                <div key={code} className="flex items-center gap-2 px-4 py-2 text-[13px]">
                  <span className="font-mono font-semibold">{code}</span>
                  <span className="flex-1 truncate text-[var(--tsub)]">{itemName(code)}</span>
                  <span className="text-[12px] text-[var(--tmuted)]">Office: {oq}</span>
                  <input
                    type="number"
                    min="1"
                    value={v.qty}
                    onChange={(e) => setCartQty(code, e.target.value)}
                    className="w-16 h-8 px-2 rounded border border-[var(--g300)] text-[12px]"
                  />
                  <button
                    onClick={() => removeFromCart(code)}
                    className="w-7 h-7 rounded flex items-center justify-center text-[var(--tmuted)] hover:text-[var(--red)] hover:bg-[var(--red-l)]"
                    title="ลบ"
                  >
                    <Icon name="close" size="sm" />
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ยืนยันเบิก (หน้าสรุปเช็คก่อนยิง) */}
      {borrowOpen && (
        <BorrowConfirmModal
          cart={cart}
          itemName={itemName}
          onClose={() => setBorrowOpen(false)}
          onConfirm={confirmBorrowSubmit}
        />
      )}

      {/* สรุปผลหลังเบิก (มีของไม่พอ) */}
      {confirmBorrow && (
        <BorrowSummaryModal
          short={confirmBorrow.items || []}
          onClose={() => setConfirmBorrow(false)}
        />
      )}

      {/* Panel */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        {/* Toolbar */}
        <div className="flex flex-wrap items-center gap-2 p-3 sm:p-3.5 border-b border-[var(--g100)]">
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

        {/* Table (จอใหญ่) */}
        <div className="hidden md:block overflow-x-auto">
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
                        officeQty={oq}
                        onTransfer={submitTransfer}
                        onEdit={editTotal}
                        onAddToCart={addToCart}
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Mobile card list */}
        <div className="md:hidden divide-y divide-[var(--g100)]">
          {paged.length === 0 && (
            <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">ไม่มีข้อมูล</div>
          )}
          {paged.map((item) => {
            const oq = parseInt(item.office) || 0;
            return (
              <div key={item.code} className={`p-3.5 space-y-2.5 ${oq < 1 ? 'opacity-60' : ''}`}>
                <div className="flex items-start gap-3">
                  <img
                    src={IMAGE_URL(item.code, item.ext || 'jpg')}
                    width="48"
                    height="48"
                    loading="lazy"
                    onError={(e) => {
                      e.currentTarget.onerror = null;
                      e.currentTarget.src = '/image/noimage.jpg';
                    }}
                    className="rounded-lg border border-[var(--g200)] object-cover shrink-0"
                    alt={item.name}
                  />
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] font-semibold text-[var(--text)] leading-snug">{item.name}</div>
                    <div className="text-[12px] font-mono text-[var(--blue)] mt-0.5">{item.code}</div>
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-2 text-center text-[12px]">
                  <div className="rounded-lg bg-[var(--surface2)] border border-[var(--g100)] py-1.5">
                    <div className="text-[var(--tmuted)]">ทั้งหมด</div>
                    <div className="font-bold text-[var(--text)]">{item.total}</div>
                  </div>
                  <div className="rounded-lg bg-[var(--surface2)] border border-[var(--g100)] py-1.5">
                    <div className="text-[var(--tmuted)]">Office</div>
                    <div className="font-bold text-[var(--text)]">{oq < 1 ? <span className="text-[var(--red)]">หมด</span> : oq}</div>
                  </div>
                  <div className="rounded-lg bg-[var(--surface2)] border border-[var(--g100)] py-1.5">
                    <div className="text-[var(--tmuted)]">Site</div>
                    <div className="font-bold text-[var(--text)]">{item.site}</div>
                  </div>
                </div>
                <RowActions
                  code={item.code}
                  name={item.name}
                  total={item.total}
                  officeQty={oq}
                  onTransfer={submitTransfer}
                  onEdit={editTotal}
                  onAddToCart={addToCart}
                />
              </div>
            );
          })}
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

function RowActions({ code, name, total, officeQty, onTransfer, onEdit, onAddToCart }) {
  const [qty, setQty] = useState('');
  const [type, setType] = useState('เบิก');

  return (
    <div className="flex items-center gap-1.5">
      <button
        onClick={() => onAddToCart(code)}
        disabled={!officeQty || officeQty < 1}
        className="h-8 px-2.5 rounded-lg bg-[var(--emerald-l)] text-[var(--emerald-d)] border border-[var(--emerald-b)] text-[12px] font-semibold hover:bg-[var(--emerald)] hover:text-white transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
        title="เพิ่มใส่ตะกร้าเบิก"
      >
        <Icon name="add" size="sm" /> เบิก
      </button>
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

function BorrowConfirmModal({ cart, itemName, onClose, onConfirm }) {
  const entries = Object.entries(cart);
  const [busy, setBusy] = useState(false);

  async function handle() {
    setBusy(true);
    await onConfirm();
  }

  return (
    <Modal title="ยืนยันเบิกทั้งหมด" onClose={onClose}>
      <div className="mb-3 text-[12px] text-[var(--tmuted)]">
        ตรวจสอบรายการเบิก {entries.length} รายการ ก่อนยืนยัน
      </div>
      <div className="max-h-72 overflow-y-auto space-y-2">
        {entries.map(([code, v]) => (
          <div
            key={code}
            className="flex items-center justify-between gap-2 px-3 py-2.5 rounded-lg border border-[var(--g100)]"
          >
            <div className="min-w-0">
              <div className="text-[13px] font-medium">{itemName(code)}</div>
              <div className="text-[11px] font-mono text-[var(--blue)]">{code}</div>
            </div>
            <div className="text-[13px] font-semibold text-[var(--emerald-d)]">{v.qty} ชิ้น</div>
          </div>
        ))}
      </div>
      <div className="mt-4 flex justify-end gap-2">
        <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
          ยกเลิก
        </button>
        <button
          onClick={handle}
          disabled={busy}
          className="px-4 py-2 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold hover:bg-[var(--emerald-d)] disabled:opacity-60"
        >
          {busy ? 'กำลังเบิก...' : 'ยืนยันเบิก'}
        </button>
      </div>
    </Modal>
  );
}

function BorrowSummaryModal({ short, onClose }) {
  return (
    <Modal title="แจ้งเตือน: เบิกได้ไม่เต็มจำนวน" onClose={onClose}>
      <div className="mb-3 text-[13px] text-[var(--red)]">
        {short.length} รายการ มีของใน Office ไม่เพียงพอตามจำนวนที่ขอ → เบิกได้เท่าที่มีจริง
      </div>
      <div className="max-h-72 overflow-y-auto space-y-2">
        {short.map((s) => (
          <div
            key={s.code}
            className="flex items-center justify-between gap-2 px-3 py-2.5 rounded-lg border border-[var(--red-b)] bg-[var(--red-l)]"
          >
            <div className="min-w-0">
              <div className="text-[13px] font-medium">{s.name}</div>
              <div className="text-[11px] font-mono text-[var(--tmuted)]">{s.code}</div>
            </div>
            <div className="text-[13px] text-[var(--red)] font-medium">
              ขอ {s.requested} / ได้ {s.actual}
            </div>
          </div>
        ))}
      </div>
      <div className="mt-4 flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold">
          รับทราบ
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
