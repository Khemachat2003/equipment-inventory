import { useState, useEffect, useCallback, useRef } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import Pagination from '../components/ui/Pagination.jsx';

export default function History() {
  const navigate = useNavigate();
  const [rows, setRows] = useState([]);
  const [start, setStart] = useState('');
  const [end, setEnd] = useState('');
  const [loading, setLoading] = useState(false);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const didInit = useRef(false);

  const load = useCallback(async (s, e) => {
    setLoading(true);
    let url = '/api/history';
    if (s || e) url += `?start=${s}&end=${e}`;
    try {
      const { data } = await axios.get(url);
      setRows(data || []);
      setPage(1);
    } catch (err) {
      console.error('history error', err);
      setRows([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!didInit.current) {
      didInit.current = true;
      load(start, end);
    }
  }, [load, start, end]);

  useEffect(() => { setPage(1); }, [pageSize]);

  const total = rows.length;
  const paged = rows.slice((page - 1) * pageSize, page * pageSize);

  function applyFilter() { load(start, end); }
  function clearFilter() { setStart(''); setEnd(''); load('', ''); }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="text-lg font-bold text-[var(--text)]">ประวัติการเบิก–คืน</div>
          <div className="text-[12px] text-[var(--tmuted)]">ดูประวัติการโอนย้ายทั้งหมดในระบบ</div>
        </div>
        <button onClick={() => navigate('/report')} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)]">
          <Icon name="description" size="sm" /> Export PDF
        </button>
      </div>

      {/* Filter bar */}
      <div className="flex flex-wrap items-end gap-3 p-4 rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)]">
        <F label="จากวันที่">
          <input type="date" value={start} onChange={(e) => setStart(e.target.value)} className={inp} />
        </F>
        <F label="ถึงวันที่">
          <input type="date" value={end} onChange={(e) => setEnd(e.target.value)} className={inp} />
        </F>
        <button onClick={applyFilter} className="h-9 px-4 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold">กรอง</button>
        <button onClick={clearFilter} className="h-9 px-4 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">ล้าง</button>
        <button onClick={() => load('', '')} className="ml-auto h-9 flex items-center gap-1.5 px-3 rounded-lg border border-[var(--g300)] text-[13px] text-[var(--tsub)]">
          <Icon name="refresh" size="sm" /> รีเฟรช
        </button>
      </div>

      {/* Table */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="px-4 py-3 border-b border-[var(--g100)] text-[12px] text-[var(--tmuted)]">
          พบ <strong className="text-[var(--tsub)]">{total}</strong> รายการ{total > 0 && ` · เรียงจากล่าสุด → เก่า`}
        </div>
        <div className="hidden md:block overflow-x-auto">
          <table className="w-full text-[12px]">
            <thead>
              <tr className="text-left text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
                <th className="px-3 py-2.5 font-medium">วันที่</th>
                <th className="px-3 py-2.5 font-medium">รหัส</th>
                <th className="px-3 py-2.5 font-medium">ชื่ออุปกรณ์</th>
                <th className="px-3 py-2.5 font-medium">ประเภท</th>
                <th className="px-3 py-2.5 font-medium text-center">จำนวน</th>
                <th className="px-3 py-2.5 font-medium">จาก</th>
                <th className="px-3 py-2.5 font-medium">ไปยัง</th>
                <th className="px-3 py-2.5 font-medium">ผู้ทำรายการ</th>
              </tr>
            </thead>
            <tbody>
              {loading && <tr><td colSpan={8} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td></tr>}
              {!loading && total === 0 && (
                <tr><td colSpan={8} className="text-center py-12 text-[var(--tmuted)]">
                  <div className="flex justify-center mb-2"><Icon name="inbox" size="2xl" /></div>ไม่พบข้อมูลในช่วงเวลานี้
                </td></tr>
              )}
              {paged.map((x, i) => (
                <tr key={i} className="border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                  <td className="px-3 py-2 text-[var(--tsub)] whitespace-nowrap">{x.date || '-'}</td>
                  <td className="px-3 py-2 font-mono text-[11px] text-[var(--blue)]">{x.code || '-'}</td>
                  <td className="px-3 py-2 font-medium">{x.name || '-'}</td>
                  <td className="px-3 py-2"><TypeBadge type={x.type} /></td>
                  <td className="px-3 py-2 text-center font-semibold">{x.qty || '-'}</td>
                  <td className="px-3 py-2 text-[var(--tsub)]">{x.from || '-'}</td>
                  <td className="px-3 py-2 text-[var(--tsub)]">{x.to || '-'}</td>
                  <td className="px-3 py-2 text-[var(--tsub)]">{x.user || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Mobile card list */}
        <div className="md:hidden divide-y divide-[var(--g100)]">
          {loading && <div className="text-center py-10 text-[var(--tmuted)] text-[13px]">กำลังโหลด...</div>}
          {!loading && total === 0 && (
            <div className="text-center py-12 text-[var(--tmuted)] text-[13px]">ไม่พบข้อมูลในช่วงเวลานี้</div>
          )}
          {paged.map((x, i) => (
            <div key={i} className="p-3.5 space-y-2">
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <div className="text-[13px] font-semibold text-[var(--text)] leading-snug">{x.name || '-'}</div>
                  <div className="text-[11px] font-mono text-[var(--blue)] mt-0.5">{x.code || '-'}</div>
                </div>
                <TypeBadge type={x.type} />
              </div>
              <div className="grid grid-cols-2 gap-1 text-[12px]">
                <div className="text-[var(--tsub)]">วันที่: <span className="text-[var(--text)]">{x.date || '-'}</span></div>
                <div className="text-[var(--tsub)]">จำนวน: <span className="font-semibold text-[var(--text)]">{x.qty || '-'}</span></div>
                <div className="text-[var(--tsub)]">จาก: <span className="text-[var(--text)]">{x.from || '-'}</span></div>
                <div className="text-[var(--tsub)]">ไปยัง: <span className="text-[var(--text)]">{x.to || '-'}</span></div>
                <div className="text-[var(--tsub)]">ผู้ทำรายการ: <span className="text-[var(--text)]">{x.user || '-'}</span></div>
              </div>
            </div>
          ))}
        </div>
        <Pagination total={total} page={page} onPage={setPage} pageSize={pageSize} onPageSize={setPageSize} />
      </div>
    </div>
  );
}

function TypeBadge({ type }) {
  const t = type || '-';
  const cls =
    t === 'เบิก' ? 'bg-[var(--blue-l)] text-[var(--blue)]' :
    t === 'คืน' ? 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' :
    'bg-[var(--g100)] text-[var(--tsub)]';
  return <span className={`px-2 py-0.5 rounded-full text-[11px] font-semibold ${cls}`}>{t}</span>;
}

function F({ label, children }) {
  return <div className="space-y-1"><label className="block text-[12px] font-medium text-[var(--tsub)]">{label}</label>{children}</div>;
}

const inp = 'h-9 px-3 rounded-lg border border-[var(--g200)] bg-[var(--surface2)] text-[13px]';