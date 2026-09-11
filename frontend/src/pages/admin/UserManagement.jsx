import { useEffect, useState, useCallback } from 'react';
import axios from 'axios';
import Icon from '../../components/ui/Icon.jsx';

export default function UserManagement() {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);

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
    try {
      await axios.put(`/api/admin/users/${encodeURIComponent(username)}`, { role });
      await load();
    } catch (e) {
      alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
    }
  }

  async function resetPassword(username) {
    const newPassword = prompt(`ตั้งรหัสผ่านใหม่สำหรับ ${username}`);
    if (!newPassword) return;
    if (newPassword.length < 4) return alert('รหัสผ่านสั้นเกินไป (ขั้นต่ำ 4)');
    try {
      await axios.post(`/api/admin/users/${encodeURIComponent(username)}/reset-password`, { newPassword });
      alert('เปลี่ยนรหัสผ่านสำเร็จ');
    } catch (e) {
      alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
    }
  }

  async function deleteUser(username) {
    if (!confirm(`ยืนยันลบผู้ใช้ ${username}?`)) return;
    try {
      await axios.delete(`/api/admin/users/${encodeURIComponent(username)}`);
      await load();
    } catch (e) {
      alert(e.response?.data?.error || 'เกิดข้อผิดพลาด');
    }
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
    </div>
  );
}