import { useEffect, useState, useCallback } from 'react';
import axios from 'axios';
import Icon from '../../components/ui/Icon.jsx';

export default function AuditLog() {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const { data } = await axios.get('/api/audit-log');
      setLogs(data);
    } catch (e) {
      console.error('audit log', e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const filtered = search
    ? logs.filter(
        (l) =>
          (l.user || '').toLowerCase().includes(search.toLowerCase()) ||
          (l.action || '').toLowerCase().includes(search.toLowerCase()) ||
          (l.module || '').toLowerCase().includes(search.toLowerCase()) ||
          (l.detail || '').toLowerCase().includes(search.toLowerCase())
      )
    : logs.slice(0, 500);

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <div className="text-lg font-bold text-[var(--text)]">Audit Log</div>
          <div className="text-[12px] text-[var(--tmuted)]">บันทึกการใช้งานระบบ (Admin)</div>
        </div>
        <button
          onClick={load}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)]"
        >
          <Icon name="refresh" size="sm" /> รีเฟรช
        </button>
      </div>

      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="p-3.5 border-b border-[var(--g100)]">
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="ค้นหา user / action / module / detail..."
            className="w-full max-w-sm h-9 pl-9 rounded-lg border border-[var(--g200)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
          />
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-[12px]">
            <thead>
              <tr className="text-left text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
                <th className="px-4 py-2.5 font-medium">เวลา</th>
                <th className="px-4 py-2.5 font-medium">ผู้ใช้</th>
                <th className="px-4 py-2.5 font-medium">การกระทำ</th>
                <th className="px-4 py-2.5 font-medium">โมดูล</th>
                <th className="px-4 py-2.5 font-medium">รายละเอียด</th>
                <th className="px-4 py-2.5 font-medium">IP</th>
              </tr>
            </thead>
            <tbody>
              {loading && (
                <tr>
                  <td colSpan={6} className="text-center py-10 text-[var(--tmuted)]">
                    กำลังโหลด...
                  </td>
                </tr>
              )}
              {!loading && filtered.length === 0 && (
                <tr>
                  <td colSpan={6} className="text-center py-10 text-[var(--tmuted)]">
                    ไม่มีข้อมูล
                  </td>
                </tr>
              )}
              {!loading &&
                filtered.map((l, i) => (
                  <tr key={i} className="border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                    <td className="px-4 py-2 whitespace-nowrap font-mono text-[11px] text-[var(--tsub)]">{l.timestamp}</td>
                    <td className="px-4 py-2 font-medium">{l.user}</td>
                    <td className="px-4 py-2">
                      <span className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${actionTone(l.action)}`}>
                        {l.action}
                      </span>
                    </td>
                    <td className="px-4 py-2 text-[var(--tsub)]">{l.module}</td>
                    <td className="px-4 py-2 text-[var(--tsub)] max-w-[340px] truncate">{l.detail}</td>
                    <td className="px-4 py-2 font-mono text-[11px] text-[var(--tmuted)]">{l.ip}</td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function actionTone(action) {
  const a = String(action || '').toLowerCase();
  if (a.includes('ลบ') || a.includes('delete')) return 'bg-[var(--red-l)] text-[var(--red)]';
  if (a.includes('แก้ไข') || a.includes('update') || a.includes('edit')) return 'bg-[var(--amber-l)] text-[var(--amber)]';
  if (a.includes('เพิ่ม') || a.includes('add') || a.includes('create')) return 'bg-[var(--emerald-l)] text-[var(--emerald-d)]';
  return 'bg-[var(--blue-l)] text-[var(--blue)]';
}