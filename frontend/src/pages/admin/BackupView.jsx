import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from '../../components/ui/Icon.jsx';

// Backup View — ดูข้อมูล Backup PostgreSQL และ Export CSV (Admin เท่านั้น)
// แปลงจาก legacy public/backup-view.html เป็น React ตาม convention ระบบ

const BACKUP_TABLES = [
  { value: 'backup_stock_master', label: 'Stock Master', icon: 'inventory_2' },
  { value: 'backup_stock_office', label: 'Stock Office', icon: 'apartment' },
  { value: 'backup_stock_site', label: 'Stock Site', icon: 'factory' },
  { value: 'backup_transfer_log', label: 'Transfer Log', icon: 'assignment' },
  { value: 'backup_asset_list', label: 'Asset List', icon: 'archive' },
  { value: 'backup_asset_history', label: 'Asset History', icon: 'description' },
  { value: 'backup_damaged_assets', label: 'อุปกรณ์ชำรุด/สูญหาย', icon: 'report' },
  { value: 'backup_part_catalog', label: 'Part Catalog', icon: 'build' },
  { value: 'backup_farm_sites', label: 'Farm Sites', icon: 'eco' },
  { value: 'backup_farm_houses', label: 'Farm Houses', icon: 'home' },
  { value: 'backup_users', label: 'Users', icon: 'person' },
  { value: 'backup_audit_log', label: 'Audit Log', icon: 'fact_check' },
];

export default function BackupView() {
  const [table, setTable] = useState('');
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [loadedFor, setLoadedFor] = useState(null);

  const load = async (tbl) => {
    if (!tbl) return;
    setLoading(true);
    setError('');
    try {
      const { data } = await axios.get(`/api/backup-data?table=${encodeURIComponent(tbl)}`);
      if (data.success && Array.isArray(data.rows)) {
        setRows(data.rows);
      } else {
        setRows([]);
      }
      setLoadedFor(tbl);
    } catch (e) {
      const msg = e.response?.data?.error || e.message;
      if (e.response?.status === 401 || e.response?.status === 403) {
        setError('ต้องเข้าสู่ระบบในฐานะ Admin ก่อนดูข้อมูล Backup');
      } else {
        setError(`เกิดข้อผิดพลาด: ${msg}`);
      }
      setRows([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    // โหลดตารางแรกอัตโนมัติ (เหมือน legacy)
    setTable('backup_stock_master');
    load('backup_stock_master');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onDownload = () => {
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    const csv = [
      columns.join(','),
      ...rows.map((row) =>
        columns
          .map((col) => {
            const val = row[col] ?? '';
            return `"${String(val).replace(/"/g, '""')}"`;
          })
          .join(',')
      ),
    ].join('\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${table}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  };

  const columns = rows.length ? Object.keys(rows[0]) : [];

  return (
    <div className="space-y-4">
      {/* Toolbar */}
      <div className="flex flex-wrap items-center gap-2">
        <select
          value={table}
          onChange={(e) => setTable(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && load(table)}
          className="h-9 px-3 rounded-lg border border-[var(--g200)] bg-white text-[13px] focus:outline-none focus:border-[var(--blue)] min-w-[220px]"
        >
          <option value="">-- เลือกตาราง --</option>
          {BACKUP_TABLES.map((t) => (
            <option key={t.value} value={t.value}>{t.label}</option>
          ))}
        </select>
        <button
          onClick={() => load(table)}
          className="flex items-center gap-1.5 h-9 px-3.5 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-50"
          disabled={!table}
        >
          <Icon name="search" size="sm" /> โหลด
        </button>
        <button
          onClick={() => load(table)}
          className="flex items-center gap-1.5 h-9 px-3.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)] disabled:opacity-50"
          disabled={!table}
        >
          <Icon name="refresh" size="sm" /> Refresh
        </button>
        <button
          onClick={onDownload}
          disabled={!rows.length}
          className="flex items-center gap-1.5 h-9 px-3.5 rounded-lg bg-[var(--emerald)] text-white text-[13px] font-semibold hover:bg-[var(--emerald-d)] disabled:opacity-50 disabled:cursor-not-allowed"
        >
          <Icon name="download" size="sm" /> Export CSV
        </button>
        {loadedFor && rows.length > 0 && (
          <span className="ml-auto text-[12px] text-[var(--tmuted)]">
            อัปเดตล่าสุด: <strong className="text-[var(--text)]">{rows[0]?.backup_date || '-'}</strong>
          </span>
        )}
      </div>

      {/* Card */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="flex items-center gap-2 px-4 py-3 border-b border-[var(--g100)]">
          <span className="text-[var(--blue)]"><Icon name="analytics" /></span>
          <h2 className="text-[14px] font-bold text-[var(--text)]">
            {loadedFor ? BACKUP_TABLES.find((t) => t.value === loadedFor)?.label || loadedFor : 'ข้อมูล Backup'}
          </h2>
          <span className="ml-auto px-2.5 py-0.5 rounded-full bg-[var(--blue-l)] text-[var(--blue)] text-[12px] font-semibold">
            {rows.length} แถว
          </span>
        </div>

        {/* เนื้อหา */}
        {error && (
          <div className="p-10 text-center">
            <div className="text-[36px] mb-2 text-[var(--tmuted)]"><Icon name="lock" size="2xl" /></div>
            <div className="text-[13px] text-[var(--tsub)]">{error}</div>
          </div>
        )}

        {!error && !table && (
          <div className="p-10 text-center text-[var(--tmuted)] text-[13px]">เลือกตารางเพื่อดูข้อมูล</div>
        )}

        {!error && table && loading && (
          <div className="p-10 text-center text-[var(--tmuted)] text-[13px]">กำลังโหลด...</div>
        )}

        {!error && table && !loading && !rows.length && (
          <div className="p-12 text-center text-[var(--tmuted)]">
            <div className="text-[36px] mb-2"><Icon name="mail" size="2xl" /></div>
            <div className="text-[13px]">ไม่มีข้อมูลในตารางนี้ — ยังไม่เคยกด Backup ทั้งระบบ หรือตารางว่าง</div>
          </div>
        )}

        {!error && table && !loading && rows.length > 0 && (
          <div className="overflow-x-auto max-h-[70vh] overflow-y-auto">
            <table className="w-full text-[12px]">
              <thead className="sticky top-0 z-10">
                <tr className="text-left text-[11px] font-bold uppercase tracking-wide text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
                  {columns.map((col) => (
                    <th key={col} className="px-4 py-2.5 whitespace-nowrap">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-[var(--g100)]">
                {rows.map((row, i) => (
                  <tr key={i} className="hover:bg-[var(--surface2)]">
                    {columns.map((col) => (
                      <td key={col} className="px-4 py-2 text-[var(--text)] whitespace-nowrap max-w-[320px] overflow-hidden text-ellipsis">{row[col] ?? '-'}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}