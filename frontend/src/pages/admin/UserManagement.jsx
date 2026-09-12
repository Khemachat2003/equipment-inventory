import { useEffect, useState, useCallback } from 'react';
import axios from 'axios';
import Icon from '../../components/ui/Icon.jsx';
import { useBusy, BusyOverlay } from '../../components/ui/Busy.jsx';

export default function UserManagement() {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [resetTarget, setResetTarget] = useState(null);
  const busy = useBusy();

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const { data } = await axios.get('/api/admin/users');
      setUsers(data);
    } catch (e) {
      console.error('load users', e);
      alert('โหลดผู้ใช้ไม่ได้ (ต้องเป็น admin)');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  async function changeRole(username, role) {
    await busy.run('กำลังเปลี่ยนสิทธิ์...', async () => {
      try {
        await axios.put(`/api/admin/users/${encodeURIComponent(username)}`, { role });
        await load();
      } catch (e) {
        alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
      }
    });
  }

  async function resetPassword(username) {
    setResetTarget(username);
  }

  async function deleteUser(username) {
    if (!confirm(`ยืนยันลบผู้ใช้ ${username}?`)) return;
    await busy.run('กำลังลบผู้ใช้...', async () => {
      try {
        await axios.delete(`/api/admin/users/${encodeURIComponent(username)}`);
        await load();
      } catch (e) {
        alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
      }
    });
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-end gap-2">
        <button
          onClick={load}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[13px] hover:bg-[var(--surface2)]"
        >
          <Icon name="refresh" size="sm" /> รีเฟรช
        </button>
      </div>

      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] overflow-hidden">
        <div className="overflow-x-auto">
        <table className="w-full text-[13px]">
          <thead>
            <tr className="text-left text-[var(--tmuted)] bg-[var(--surface2)] border-b border-[var(--g100)]">
              <th className="px-4 py-3 font-medium">ผู้ใช้</th>
              <th className="px-4 py-3 font-medium">สิทธิ์</th>
              <th className="px-4 py-3 font-medium text-right">จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {loading && (
              <tr>
                <td colSpan={3} className="text-center py-10 text-[var(--tmuted)]">กำลังโหลด...</td>
              </tr>
            )}
            {!loading &&
              users.map((u) => (
                <tr key={u.username} className="border-b border-[var(--g100)] hover:bg-[var(--surface2)]">
                  <td className="px-4 py-2.5 font-medium">{u.username}</td>
                  <td className="px-4 py-2.5">
                    <select
                      value={u.role}
                      onChange={(e) => changeRole(u.username, e.target.value)}
                      className="h-8 px-2 rounded border border-[var(--g300)] text-[12px]"
                    >
                      <option value="user">user</option>
                      <option value="admin">admin</option>
                    </select>
                  </td>
                  <td className="px-4 py-2.5 text-right space-x-2">
                    <button
                      onClick={() => resetPassword(u.username)}
                      className="px-2.5 py-1.5 rounded-lg border border-[var(--g300)] text-[11px] text-[var(--tsub)] hover:bg-[var(--surface2)]"
                    >
                      รีเซ็ตรหัส
                    </button>
                    <button
                      onClick={() => deleteUser(u.username)}
                      className="px-2.5 py-1.5 rounded-lg bg-[var(--red-l)] text-[var(--red)] text-[11px] font-medium hover:bg-[var(--red-b)]"
                    >
                      ลบ
                    </button>
                  </td>
                </tr>
              ))}
          </tbody>
        </table>
        </div>
      </div>
      <BusyOverlay label={busy.busyLabel} />
      {resetTarget && (
        <PasswordResetModal
          username={resetTarget}
          onClose={() => setResetTarget(null)}
          onDone={() => setResetTarget(null)}
        />
      )}
    </div>
  );
}

function PasswordResetModal({ username, onClose, onDone }) {
  const [pwd, setPwd] = useState('');
  const [busy, setBusy] = useState(false);

  async function confirm() {
    if (!pwd) return alert('กรอกรหัสผ่านใหม่');
    if (pwd.length < 4) return alert('รหัสผ่านสั้นเกินไป (ขั้นต่ำ 4)');
    setBusy(true);
    try {
      await axios.post(`/api/admin/users/${encodeURIComponent(username)}/reset-password`, { newPassword: pwd });
      alert('เปลี่ยนรหัสผ่านสำเร็จ');
      onDone();
    } catch (e) {
      alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40">
      <div className="w-full max-w-sm rounded-2xl bg-white shadow-[var(--sh-lg)]">
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-[var(--g100)]">
          <div className="text-[14px] font-semibold text-[var(--text)]">ตั้งรหัสผ่านใหม่ · {username}</div>
          <button onClick={onClose} className="text-[var(--tmuted)] hover:text-[var(--text)]">
            <Icon name="close" size="md" />
          </button>
        </div>
        <div className="p-5">
          <label className="block text-[12px] font-medium text-[var(--tsub)] mb-1">รหัสผ่านใหม่ (ขั้นต่ำ 4)</label>
          <input
            type="password"
            value={pwd}
            onChange={(e) => setPwd(e.target.value)}
            placeholder="รหัสผ่านใหม่"
            autoFocus
            className="w-full h-9 px-3 rounded-lg border border-[var(--g200)] text-[13px] focus:outline-none focus:border-[var(--blue)]"
            onKeyDown={(e) => e.key === 'Enter' && !busy && confirm()}
          />
          <div className="flex justify-end gap-2 mt-4">
            <button onClick={onClose} className="px-3 py-2 rounded-lg border border-[var(--g300)] text-[13px]">
              ยกเลิก
            </button>
            <button
              onClick={confirm}
              disabled={busy}
              className="px-4 py-2 rounded-lg bg-[var(--blue)] text-white text-[13px] font-semibold disabled:opacity-60"
            >
              {busy ? 'กำลังบันทึก...' : 'บันทึก'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}