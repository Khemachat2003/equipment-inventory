import { useNavigate } from 'react-router-dom';
import { useState } from 'react';
import { useAuth } from '../context/AuthContext.jsx';
import Icon from '../components/ui/Icon.jsx';

export default function Login() {
  const { login, error } = useAuth();
  const navigate = useNavigate();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (!username.trim() || !password.trim()) return;
    setBusy(true);
    const res = await login(username, password);
    setBusy(false);
    if (res.ok) navigate('/');
  }

  return (
    <div className="min-h-screen flex bg-[var(--ink)] text-white">
      {/* Left brand panel */}
      <div className="hidden lg:flex flex-1 flex-col justify-between p-10 max-w-[520px]">
        <div className="flex items-center gap-3">
          <span className="flex items-center justify-center w-11 h-11 rounded-xl bg-[var(--blue)]">
            <Icon name="inventory_2" size="md" />
          </span>
          <div>
            <div className="text-lg font-bold tracking-wide">INTRANIN</div>
            <div className="text-[12px] text-white/50">Equipment Management System</div>
          </div>
        </div>

        <div>
          <h1 className="text-3xl font-bold leading-snug">
            ระบบบริหาร<br />
            <span className="text-[var(--blue-b)]">จัดการอุปกรณ์</span>
          </h1>
          <p className="mt-3 text-white/60 text-sm leading-relaxed">
            ติดตาม ตรวจสอบ และบริหารอุปกรณ์ครบวงจร
            <br />
            แบบ Real-time ทุกฟาร์มทุกไซต์งาน
          </p>
          <ul className="mt-6 space-y-2.5 text-sm text-white/70">
            <li className="flex items-center gap-2.5">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)]" /> Asset Tracking แบบ Real-time
            </li>
            <li className="flex items-center gap-2.5">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)]" /> Monitor แยกฟาร์มแต่ละไซต์
            </li>
            <li className="flex items-center gap-2.5">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)]" /> สแกน Barcode → ค้นหาอัตโนมัติ
            </li>
            <li className="flex items-center gap-2.5">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)]" /> Export รายงาน PDF
            </li>
          </ul>
        </div>

        <div className="text-[11px] text-white/30">© INTRANIN EMS</div>
      </div>

      {/* Right login form */}
      <div className="flex-1 flex items-center justify-center p-6">
        <form onSubmit={handleSubmit} className="w-full max-w-sm">
          <div className="mb-6 lg:hidden">
            <div className="text-sm font-bold tracking-wide text-[--ink]">INTRANIN</div>
          </div>
          <div className="text-[11px] font-semibold tracking-widest text-[var(--blue-b)] uppercase mb-1">
            INTRANIN EMS
          </div>
          <h2 className="text-xl font-bold">ยินดีต้อนรับ</h2>
          <p className="text-sm text-white/50 mb-6">กรุณาเข้าสู่ระบบเพื่อดำเนินการต่อ</p>

          <label className="block text-[13px] font-medium text-white/80 mb-1">Username</label>
          <input
            type="text"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            placeholder="กรอก Username"
            autoComplete="username"
            className="w-full h-11 px-3.5 rounded-lg bg-white/5 border border-white/10 text-white placeholder-white/30 text-sm mb-4 focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
          />

          <label className="block text-[13px] font-medium text-white/80 mb-1">Password</label>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="กรอก Password"
            autoComplete="current-password"
            className="w-full h-11 px-3.5 rounded-lg bg-white/5 border border-white/10 text-white placeholder-white/30 text-sm mb-4 focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)]"
          />

          {error && (
            <div className="mb-4 text-[13px] text-[var(--red-b)]">
              <Icon name="error" size="sm" /> {error}
            </div>
          )}

          <button
            type="submit"
            disabled={busy}
            className="w-full h-11 rounded-lg bg-[var(--blue)] hover:bg-[var(--blue-d)] text-white text-sm font-semibold transition-colors disabled:opacity-60"
          >
            {busy ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ'}
          </button>
        </form>
      </div>
    </div>
  );
}
