import { useEffect, useState, useMemo, useCallback } from 'react';
import axios from 'axios';
import Icon from '../components/ui/Icon.jsx';
import { useBusy, BusyOverlay } from '../components/ui/Busy.jsx';

const IMAGE_URL = (code, ext) =>
  `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image@main/images/${code}.${ext}?v=4`;
const PAGE_SIZES = [20, 50, 100];

export default function Stock() {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  const [returnOpen, setReturnOpen] = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [borrowOpen, setBorrowOpen] = useState(false);
  const [confirmBorrow, setConfirmBorrow] = useState(null);
  const [cart, setCart] = useState({});
  const [notice, setNotice] = useState('');
  const [addPrompt, setAddPrompt] = useState(null);
  const [editPrompt, setEditPrompt] = useState(null);
  const busyBorrow = useBusy();

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const { data } = await axios.get('/api/stock');
      setItems(data);
    } catch (e) {
      console.error('load stock', e);
    } finally {
      setLoading(false);
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
  function adjustCart(code, delta) {
    setCartQty(code, (cart[code]?.qty || 0) + delta);
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

  async function confirmBorrowSubmit() {
    const carts = Object.entries(cart).map(([code, v]) => ({ code, qty: v.qty }));
    if (!carts.length) return;
    // ครอบด้วย busy: modal ค้างอยู่พร้อมสถานะ + overlay กันกดซ้ำระหว่างเบิกเสร็จ
    await busyBorrow.run('กำลังเบิกของ...', async () => {
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
      } finally {
        setBorrowOpen(false);
      }
    });
  }

  // ปุ่ม "แก้ไข" — เปิด modal กรอก Office + Site (prompt() โดนบล็อกใน iframe จึงใช้ modal)
  function editTotal(code, office, site) {
    setEditPrompt({ code, office, site });
  }

  // ปุ่ม "+" — เปิด modal กรอกจำนวนที่จะเพิ่ม (บวกเข้ากับของเดิมที่ Office)
  function addQty(code, total) {
    setAddPrompt({ code, total });
  }

  function flash(msg) {
    setNotice(msg);
    setTimeout(() => setNotice(''), 4000);
  }

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-3">
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
              {loading ? (
                <tr>
                  <td colSpan={7} className="text-center py-12 text-[var(--tmuted)]">กำลังโหลด...</td>
                </tr>
              ) : paged.length === 0 ? (
                <tr>
                  <td colSpan={7} className="text-center py-12 text-[var(--tmuted)]">
                    ไม่มีข้อมูล
                  </td>
                </tr>
              ) : null}
              {!loading &&
                paged.map((item) => {
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
                        total={item.total}
                        officeQty={oq}
                        siteQty={item.site}
                        cartQty={cart[item.code]?.qty || 0}
                        onEdit={editTotal}
                        onAddQty={addQty}
                        onAddToCart={addToCart}
                        onAdjust={(d) => adjustCart(item.code, d)}
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
          {loading ? (
            <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">กำลังโหลด...</div>
          ) : paged.length === 0 ? (
            <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">ไม่มีข้อมูล</div>
          ) : null}
          {!loading &&
            paged.map((item) => {
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
                  total={item.total}
                  officeQty={oq}
                  siteQty={item.site}
                  cartQty={cart[item.code]?.qty || 0}
                  onEdit={editTotal}
                  onAddQty={addQty}
                  onAddToCart={addToCart}
                  onAdjust={(d) => adjustCart(item.code, d)}
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
      {addPrompt && (
        <AddQtyModal
          code={addPrompt.code}
          total={addPrompt.total}
          onClose={() => setAddPrompt(null)}
          onDone={async (a, res) => {
            setAddPrompt(null);
            await load();
            flash(`เพิ่ม ${a} ชิ้น → Office ${res.office} | Site ${res.site} | รวม ${res.total}`);
          }}
          onError={(msg) => alert(msg)}
        />
      )}
      {editPrompt && (
        <EditQtyModal
          code={editPrompt.code}
          initialOffice={editPrompt.office}
          initialSite={editPrompt.site}
          onClose={() => setEditPrompt(null)}
          onDone={async (res) => {
            setEditPrompt(null);
            await load();
            flash(`บันทึกแล้ว → Office ${res.office} | Site ${res.site} | รวม ${res.total}`);
          }}
          onError={(msg) => alert(msg)}
        />
      )}
      <BusyOverlay label={busyBorrow.busyLabel} />
    </div>
  );
}

function RowActions({ code, total, officeQty, siteQty, cartQty, onEdit, onAddQty, onAddToCart, onAdjust }) {
  return (
    <div className="flex flex-wrap items-center gap-1.5">
      {cartQty > 0 ? (
        <div className="flex items-center gap-1 h-8 shrink-0">
          <button
            onClick={() => onAdjust(-1)}
            className="w-7 h-8 rounded-lg bg-[var(--emerald)] text-white flex items-center justify-center hover:bg-[var(--emerald-d)]"
            title={cartQty === 1 ? 'นำออกจากตะกร้า' : 'ลดจำนวน'}
          >
            <Icon name="remove" size="sm" />
          </button>
          <span className="min-w-[2rem] px-1 text-center text-[13px] font-bold text-[var(--emerald-d)] tabular-nums">
            {cartQty}
          </span>
          <button
            onClick={() => onAdjust(1)}
            disabled={!officeQty || officeQty < 1}
            className="w-7 h-8 rounded-lg bg-[var(--emerald)] text-white flex items-center justify-center hover:bg-[var(--emerald-d)] disabled:opacity-40"
            title="เพิ่มจำนวน"
          >
            <Icon name="add" size="sm" />
          </button>
        </div>
      ) : (
        <button
          onClick={() => onAddToCart(code)}
          disabled={!officeQty || officeQty < 1}
          className="h-8 px-2.5 rounded-lg bg-[var(--emerald-l)] text-[var(--emerald-d)] border border-[var(--emerald-b)] text-[12px] font-semibold hover:bg-[var(--emerald)] hover:text-white transition-colors disabled:opacity-40 disabled:cursor-not-allowed shrink-0 whitespace-nowrap"
          title="เพิ่มใส่ตะกร้าเบิก"
        >
          <Icon name="add" size="sm" /> เบิก
        </button>
      )}
      <button
        onClick={() => onAddQty(code, total)}
        className="w-8 h-8 rounded-lg border border-[var(--emerald-b)] bg-[var(--emerald-l)] text-[var(--emerald-d)] flex items-center justify-center hover:bg-[var(--emerald)] hover:text-white shrink-0"
        title="เพิ่มจำนวน (บวกเข้ากับของเดิมที่ Office)"
      >
        <Icon name="add" size="sm" />
      </button>
      <button
        onClick={() => onEdit(code, officeQty, siteQty || 0)}
        className="w-8 h-8 rounded-lg border border-[var(--g300)] text-[var(--tsub)] flex items-center justify-center hover:bg-[var(--surface2)] shrink-0"
        title="แก้ไข (ตั้งค่า Office และ Site)"
      >
        <Icon name="edit" size="sm" />
      </button>
    </div>
  );
}

function ReturnModal({ onClose, onDone }) {
  const [items, setItems] = useState([]);
  const [selected, setSelected] = useState({});
  const [qtys, setQtys] = useState({});
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
    setSelected((s) => {
      const next = { ...s, [code]: !s[code] };
      if (next[code]) {
        const it = items.find((i) => i.code === code);
        if (it) setQtys((q) => ({ ...q, [code]: parseInt(it.qty) }));
      }
      return next;
    });
  }
  function allToggle() {
    const all = items.length > 0 && items.every((i) => selected[i.code]);
    const nextSel = Object.fromEntries(items.map((i) => [i.code, !all]));
    setSelected(nextSel);
    if (!all) {
      setQtys(Object.fromEntries(items.map((i) => [i.code, parseInt(i.qty)])));
    }
  }
  function setQty(code, val) {
    const it = items.find((i) => i.code === code);
    const max = it ? parseInt(it.qty) : 99999;
    const v = Math.max(1, Math.min(parseInt(val) || 1, max));
    setQtys((q) => ({ ...q, [code]: v }));
  }

  async function confirm() {
    const toReturn = items
      .filter((i) => selected[i.code])
      .map((i) => ({ code: i.code, qty: parseInt(qtys[i.code]) || parseInt(i.qty) }))
      .filter((r) => r.qty >= 1);
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
  const selectedTotal = items
    .filter((i) => selected[i.code])
    .reduce((a, i) => a + (parseInt(qtys[i.code]) || 0), 0);

  return (
    <Modal title="คืนจาก Site" onClose={onClose}>
      <div className="flex items-center justify-between mb-3">
        <button onClick={allToggle} className="text-[12px] text-[var(--blue)] font-medium">
          เลือกทั้งหมด / ยกเลิก
        </button>
        <span className="text-[12px] text-[var(--tmuted)]">
          เลือก {selectedCount} รายการ{selectedCount ? ` · คืน ${selectedTotal} ชิ้น` : ''}
        </span>
      </div>
      <div className="max-h-72 overflow-y-auto space-y-2">
        {loading && <div className="text-center py-8 text-[var(--tmuted)]">กำลังโหลด...</div>}
        {!loading && items.length === 0 && (
          <div className="text-center py-8 text-[var(--tmuted)]">ไม่มีอุปกรณ์ใน Site</div>
        )}
        {items.map((i) => (
          <div
            key={i.code}
            className={`flex items-center justify-between gap-2 px-3 py-2.5 rounded-lg border ${
              selected[i.code] ? 'border-[var(--emerald-b)] bg-[var(--emerald-l)]' : 'border-[var(--g200)]'
            }`}
          >
            <label className="flex items-center gap-2.5 cursor-pointer min-w-0">
              <input type="checkbox" checked={!!selected[i.code]} onChange={() => toggle(i.code)} />
              <span className="text-[13px] font-medium truncate">{i.name}</span>
            </label>
            <div className="flex items-center gap-2 shrink-0">
              <span className="text-[12px] text-[var(--tmuted)]">มี {i.qty} ชิ้น</span>
              <input
                type="number"
                min="1"
                max={i.qty}
                value={selected[i.code] ? (qtys[i.code] ?? i.qty) : i.qty}
                onChange={(e) => setQty(i.code, e.target.value)}
                disabled={!selected[i.code]}
                className="w-16 h-8 px-2 rounded border border-[var(--g300)] text-[12px] disabled:opacity-40 disabled:cursor-not-allowed"
              />
            </div>
          </div>
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
  const [office, setOffice] = useState('');
  const [site, setSite] = useState('');
  const [file, setFile] = useState(null);
  const busy = useBusy();

  // ถ้ากรอก Office/Site → รวม = Office+Site, ไม่กรอก → รวม = จำนวนทั้งหมด (ของอยู่ Office ทั้งหมด)
  const officeN = office === '' || office == null ? null : parseInt(office);
  const siteN = site === '' || site == null ? null : parseInt(site);
  const liveTotal =
    officeN == null && siteN == null ? parseInt(qty) || 0 : (officeN || 0) + (siteN || 0);

  async function submit() {
    if (!code || !name || !file) return alert('กรอกข้อมูลและเลือกรูปให้ครบ');
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['jpg', 'jpeg', 'png'].includes(ext)) return alert('รองรับ JPG/PNG เท่านั้น');
    if (officeN == null && siteN == null && (qty === '' || parseInt(qty) < 0))
      return alert('กรอกจำนวนให้ถูกต้อง');
    await busy.run('กำลังเพิ่มอุปกรณ์...', async () => {
      try {
        const base64 = await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target.result);
          reader.readAsDataURL(file);
        });
        const up = await (await fetch('/upload-image', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fileName: code + '.' + ext, base64 }),
        })).json();
        if (!up.success) return alert('อัปโหลดรูปไม่สำเร็จ: ' + (up.error || 'ไม่ทราบสาเหตุ'));
        const off = officeN == null && siteN == null ? parseInt(qty) || 0 : officeN || 0;
        const sit = siteN == null ? 0 : siteN || 0;
        try {
          const { data } = await axios.post('/api/add-item', {
            code,
            name,
            total: off + sit,
            office: off,
            site: sit,
            ext,
          });
          if (!data.success) return alert(data.error || 'เพิ่มอุปกรณ์ไม่สำเร็จ');
          alert('เพิ่มอุปกรณ์สำเร็จ');
          onDone();
        } catch (err) {
          alert(err?.response?.data?.error || 'เกิดข้อผิดพลาด');
        }
      } catch (e) {
        alert('เกิดข้อผิดพลาด');
      }
    });
  }

  return (
    <Modal title="เพิ่มอุปกรณ์" onClose={onClose} busyLabel={busy.busyLabel}>
      <Field label="รหัสอุปกรณ์" value={code} onChange={setCode} placeholder="เช่น AC-001" />
      <Field label="ชื่ออุปกรณ์" value={name} onChange={setName} />
      <Field label="จำนวนทั้งหมด" type="number" value={qty} onChange={setQty} placeholder="0" />
      <div className="grid grid-cols-2 gap-2">
        <Field label="จำนวน Office (ทางเลือก)" type="number" value={office} onChange={setOffice} placeholder="ใส่ก็ได้ หากมีที่ Office" />
        <Field label="จำนวน Site (ทางเลือก)" type="number" value={site} onChange={setSite} placeholder="ใส่ก็ได้ หากมีที่ Site" />
      </div>
      <div className="mb-3 px-3 py-2 rounded-lg bg-[var(--surface2)] border border-[var(--g100)] text-[12px] text-[var(--tsub)]">
        รวมทั้งหมด: <span className="font-bold text-[var(--blue)]">{liveTotal}</span> ชิ้น
        <div className="mt-1 text-[11px] text-[var(--tmuted)]">
          ไม่กรอก Office/Site → ของใหม่ทั้งหมดเริ่มอยู่ที่ Office (Site = 0); ถ้ากรอก จำนวนรวมจะเท่า Office + Site
        </div>
      </div>
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
          disabled={busy.busy}
          className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60"
        >
          {busy.busy ? 'กำลังเพิ่ม...' : 'เพิ่มอุปกรณ์'}
        </button>
      </div>
    </Modal>
  );
}

function AddQtyModal({ code, total, onClose, onDone, onError }) {
  const [addNum, setAddNum] = useState('');
  const busy = useBusy();

  async function confirm() {
    const a = parseInt(addNum);
    if (!a || a < 1) return onError('กรอกจำนวนที่จะเพิ่มอย่างน้อย 1');
    await busy.run('กำลังเพิ่มจำนวน...', async () => {
      try {
        const { data } = await axios.post('/api/add-total', { code, addQty: a });
        if (data.error) onError(data.error);
        else onDone(a, data);
      } catch (e) {
        onError(e?.response?.data?.error || 'เกิดข้อผิดพลาด');
      }
    });
  }

  return (
    <Modal title={`เพิ่มจำนวน ${code}`} onClose={onClose} busyLabel={busy.busyLabel}>
      <div className="mb-3 text-[12px] text-[var(--tsub)]">
        จำนวนปัจจุบัน: <b>{total}</b> ชิ้น — จำนวนที่เพิ่มจะบวกเข้ากับของเดิมที่ Office (Site คงเดิม)
      </div>
      <input
        type="number"
        min="1"
        value={addNum}
        onChange={(e) => setAddNum(e.target.value)}
        placeholder="จำนวนที่จะเพิ่ม"
        autoFocus
        className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
        onKeyDown={(e) => e.key === 'Enter' && !busy.busy && confirm()}
      />
      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
          ยกเลิก
        </button>
        <button
          onClick={confirm}
          disabled={busy.busy}
          className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60"
        >
          {busy.busy ? 'กำลังเพิ่ม...' : 'เพิ่มจำนวน'}
        </button>
      </div>
    </Modal>
  );
}

function EditQtyModal({ code, initialOffice, initialSite, onClose, onDone, onError }) {
  const [office, setOffice] = useState(String(initialOffice ?? 0));
  const [site, setSite] = useState(String(initialSite ?? 0));
  const busy = useBusy();

  const officeN = office === '' ? 0 : parseInt(office);
  const siteN = site === '' ? 0 : parseInt(site);
  const liveTotal = (isNaN(officeN) ? 0 : officeN) + (isNaN(siteN) ? 0 : siteN);

  async function confirm() {
    const off = office.replace(/\D/g, '') === '' ? 0 : parseInt(office);
    const sit = site.replace(/\D/g, '') === '' ? 0 : parseInt(site);
    await busy.run('กำลังบันทึก...', async () => {
      try {
        const { data } = await axios.post('/api/set-quantities', { code, office: off, site: sit });
        if (data.error) onError(data.error);
        else onDone(data);
      } catch (e) {
        onError(e?.response?.data?.error || 'เกิดข้อผิดพลาด');
      }
    });
  }

  return (
    <Modal title={`แก้ไขจำนวน ${code}`} onClose={onClose} busyLabel={busy.busyLabel}>
      <div className="grid grid-cols-2 gap-2">
        <Field label="Office" type="number" value={office} onChange={setOffice} placeholder="0" />
        <Field label="Site" type="number" value={site} onChange={setSite} placeholder="0" />
      </div>
      <div className="mb-3 px-3 py-2 rounded-lg bg-[var(--surface2)] border border-[var(--g100)] text-[12px] text-[var(--tsub)]">
        รวมทั้งหมด: <span className="font-bold text-[var(--blue)]">{liveTotal}</span> ชิ้น (คำนวณจาก Office + Site อัตโนมัติ)
      </div>
      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
          ยกเลิก
        </button>
        <button
          onClick={confirm}
          disabled={busy.busy}
          className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60"
        >
          {busy.busy ? 'กำลังบันทึก...' : 'บันทึก'}
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

function Modal({ title, onClose, children, busyLabel }) {
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
      <BusyOverlay label={busyLabel} />
    </div>
  );
}
