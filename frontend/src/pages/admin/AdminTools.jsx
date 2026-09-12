import { useEffect, useState } from 'react';
import axios from 'axios';
import Icon from '../../components/ui/Icon.jsx';
import { useBusy, BusyOverlay } from '../../components/ui/Busy.jsx';

export default function AdminTools() {
  const [backupResult, setBackupResult] = useState('');
  const [fullBackupResult, setFullBackupResult] = useState('');
  const busy = useBusy();

  async function doFullBackup() {
    setFullBackupResult('');
    await busy.run('กำลังทำ Full Backup...', async () => {
      try {
        const { data } = await axios.post('/api/full-backup');
        setFullBackupResult(data.message || 'Backup สำเร็จ');
      } catch (e) {
        setFullBackupResult('ไม่สำเร็จ: ' + (e.response?.data?.error || 'Backup ล้มเหลว'));
      }
    });
  }

  async function doBackup() {
    setBackupResult('');
    await busy.run('กำลังทำ Quick Backup...', async () => {
      try {
        const { data } = await axios.post('/api/backup');
        setBackupResult(data.message || 'Backup สำเร็จ');
      } catch (e) {
        setBackupResult('ไม่สำเร็จ: ' + (e.response?.data?.error || 'Backup ล้มเหลว'));
      }
    });
  }

  async function clearCache() {
    await busy.run('กำลังล้าง Cache...', async () => {
      try {
        await axios.post('/api/clear-cache');
        alert('Cache ถูกเคลียร์แล้ว');
      } catch (e) {
        alert('เกิดข้อผิดพลาด');
      }
    });
  }

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {/* Full backup */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center gap-2 mb-2">
            <span className="flex items-center justify-center w-9 h-9 rounded-lg bg-[var(--blue-l)] text-[var(--blue)]">
              <Icon name="database" size="sm" />
            </span>
            <div>
              <div className="text-[14px] font-semibold text-[var(--text)]">Full System Backup</div>
              <div className="text-[11px] text-[var(--tmuted)]">สำรองข้อมูลทั้งหมด → PostgreSQL</div>
            </div>
          </div>
          <button
            onClick={doFullBackup}
            disabled={busy.busy}
            className="mt-3 px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold hover:bg-[var(--blue-d)] disabled:opacity-60"
          >
            {busy.busy ? 'กำลัง Backup...' : 'Start Full Backup'}
          </button>
          {fullBackupResult && <div className="mt-3 text-[12px] text-[var(--tsub)]">{fullBackupResult}</div>}
        </div>

        {/* Quick backup */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center gap-2 mb-2">
            <span className="flex items-center justify-center w-9 h-9 rounded-lg bg-[var(--amber-l)] text-[var(--amber)]">
              <Icon name="save" size="sm" />
            </span>
            <div>
              <div className="text-[14px] font-semibold text-[var(--text)]">Quick Backup</div>
              <div className="text-[11px] text-[var(--tmuted)]">สำรองข้อมูลหลัก → PostgreSQL</div>
            </div>
          </div>
          <button
            onClick={doBackup}
            disabled={busy.busy}
            className="mt-3 px-4 py-2 rounded-lg bg-[var(--amber)] text-white text-[13px] font-semibold hover:bg-[var(--amber)] dark disabled:opacity-60"
          >
            {busy.busy ? 'กำลัง Backup...' : 'Start Backup'}
          </button>
          {backupResult && <div className="mt-3 text-[12px] text-[var(--tsub)]">{backupResult}</div>}
        </div>

        {/* Clear cache */}
        <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-5">
          <div className="flex items-center gap-2 mb-2">
            <span className="flex items-center justify-center w-9 h-9 rounded-lg bg-[var(--red-l)] text-[var(--red)]">
              <Icon name="cleaning_services" size="sm" />
            </span>
            <div>
              <div className="text-[14px] font-semibold text-[var(--text)]">Clear Cache</div>
              <div className="text-[11px] text-[var(--tmuted)]">ล้าง cache ของระบบ (โหลดข้อมูลใหม่จาก Sheet)</div>
            </div>
          </div>
          <button
            onClick={clearCache}
            disabled={busy.busy}
            className="mt-3 px-4 py-2 rounded-lg bg-[var(--red)] text-white text-[13px] font-semibold hover:bg-[var(--red-b)] disabled:opacity-60"
          >
            <Icon name="refresh" size="sm" /> Clear Cache
          </button>
        </div>
      </div>

      <BusyOverlay label={busy.busyLabel} />
    </div>
  );
}