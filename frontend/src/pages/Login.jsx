import { useNavigate } from 'react-router-dom';
import { useState } from 'react';
import { useAuth } from '../context/AuthContext.jsx';
import Icon from '../components/ui/Icon.jsx';

// Login (Step 4 — Enterprise/Industrial polish):
// พื้นหลัง Slate เข้ม + เส้นกริดจางๆ (industrial blueprint) + ambient glow,
// ฟอร์มยกการ์ดแก้ว (glass card) — คง logic การเข้าสู่ระบบเดิมทุกบิต
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

  const FEATURES = [
    { icon: 'monitoring', text: 'Farm Monitor — เห็นสถานะทุกฟาร์มแบบ Real-time' },
    { icon: 'grid_on', text: 'ควบคุมตู้ชุดอุปกรณ์ (Bundle) แบบครบวงจร' },
    { icon: 'document_scanner', text: 'สแกน Barcode → ตำแหน่งฟาร์ม/โรงเรือนทันที' },
    { icon: 'summarize', text: 'รายงาน PDF และ Audit Trail ครบถ้วน' },
  ];

  return (
    <div className="min-h-screen flex bg-[var(--ink)] text-white relative overflow-hidden">
      {/* ── Ambient: blueprint grid + glows ── */}
      <div className="absolute inset-0 pointer-events-none" aria-hidden="true">
        <div
          className="absolute inset-0 opacity-[0.5]"
          style={{
            backgroundImage:
              'linear-gradient(rgba(148,163,184,.05) 1px, transparent 1px), linear-gradient(90deg, rgba(148,163,184,.05) 1px, transparent 1px)',
            backgroundSize: '48px 48px',
          }}
        />
        <div className="absolute -top-40 -left-32 w-[480px] h-[480px] rounded-full bg-[var(--blue)]/20 blur-3xl" />
        <div className="absolute -bottom-48 right-[-120px] w-[520px] h-[520px] rounded-full bg-[var(--emerald)]/10 blur-3xl" />
        <div className="absolute top-1/3 right-1/4 w-64 h-64 rounded-full bg-[var(--purple)]/10 blur-3xl" />
      </div>

      {/* ── Left brand panel (จอใหญ่) ── */}
      <div className="hidden lg:flex flex-1 flex-col justify-between p-10 max-w-[560px] relative">
        <div className="flex items-center gap-3">
          <span className="flex items-center justify-center w-11 h-11 rounded-xl bg-[var(--blue)] shadow-[0_0_24px_var(--blue-glow)]">
            <Icon name="inventory_2" size="md" />
          </span>
          <div>
            <div className="text-lg font-bold tracking-wide">INTRANIN</div>
            <div className="text-[12px] text-white/50">Equipment Management System</div>
          </div>
        </div>

        <div>
          <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full border border-white/15 bg-white/[0.06] text-[10px] font-bold uppercase tracking-[0.2em] text-white/70 mb-4">
            <Icon name="verified" size="xs" /> Enterprise Platform
          </div>
          <h1 className="text-3xl font-bold leading-snug">
            ระบบบริหาร<br />
            <span className="text-[var(--blue-b)]">จัดการอุปกรณ์</span> ระดับองค์กร
          </h1>
          <p className="mt-3 text-white/60 text-sm leading-relaxed">
            ติดตาม ตรวจสอบ และบริหารอุปกรณ์ครบวงจร
            <br />
            แบบ Real-time ทุกฟาร์มทุกไซต์งาน
          </p>
          <ul className="mt-7 space-y-3 text-[13px] text-white/75">
            {FEATURES.map((f) => (
              <li key={f.icon} className="flex items-center gap-3">
                <span className="w-8 h-8 rounded-lg bg-white/[0.07] border border-white/10 flex items-center justify-center text-[var(--emerald)] shrink-0">
                  <Icon name={f.icon} size="xs" />
                </span>
                {f.text}
              </li>
            ))}
          </ul>
        </div>

        <div className="flex items-center gap-4 text-[11px] text-white/35">
          <span className="flex items-center gap-1.5"><Icon name="shield" size="xs" /> Secure Access</span>
          <span className="flex items-center gap-1.5"><Icon name="cloud_done" size="xs" /> Cloud Sync</span>
          <span>© INTRANIN EMS</span>
        </div>
      </div>

      {/* ── Right login form (glass card) ── */}
      <div className="flex-1 flex items-center justify-center p-5 sm:p-6 relative">
        <div className="w-full max-w-[400px]">
          {/* แบรนด์ย่อ (จอเล็ก) */}
          <div className="mb-5 lg:hidden flex items-center gap-2.5">
            <span className="flex items-center justify-center w-9 h-9 rounded-lg bg-[var(--blue)]">
              <Icon name="inventory_2" size="sm" />
            </span>
            <div className="leading-tight">
              <div className="text-sm font-bold tracking-wide">INTRANIN</div>
              <div className="text-[11px] text-white/50">Equipment Management System</div>
            </div>
          </div>

          <div className="rounded-2xl border border-white/10 bg-white/[0.05] backdrop-blur-md shadow-[var(--sh-lg)] p-6 sm:p-7">
            <div className="text-[10px] font-bold tracking-[0.25em] text-[var(--blue-b)] uppercase mb-1.5">
              INTRANIN EMS
            </div>
            <h2 className="text-xl font-bold">ยินดีต้อนรับกลับ</h2>
            <p className="text-[13px] text-white/50 mt-1 mb-6">กรุณาเข้าสู่ระบบเพื่อดำเนินการต่อ</p>

            <form onSubmit={handleSubmit}>
              <label htmlFor="login-username" className="block text-[12px] font-semibold text-white/75 mb-1.5">Username</label>
              <div className="relative mb-4">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-white/35 pointer-events-none">
                  <Icon name="person" size="sm" />
                </span>
                <input
                  id="login-username"
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="กรอก Username"
                  autoComplete="username"
                  className="w-full h-11 pl-9 pr-3.5 rounded-xl bg-white/[0.06] border border-white/12 text-white placeholder-white/30 text-sm focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)] transition-colors"
                />
              </div>

              <label htmlFor="login-password" className="block text-[12px] font-semibold text-white/75 mb-1.5">Password</label>
              <div className="relative mb-4">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-white/35 pointer-events-none">
                  <Icon name="lock" size="sm" />
                </span>
                <input
                  id="login-password"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="กรอก Password"
                  autoComplete="current-password"
                  className="w-full h-11 pl-9 pr-3.5 rounded-xl bg-white/[0.06] border border-white/12 text-white placeholder-white/30 text-sm focus:outline-none focus:border-[var(--blue)] focus:ring-2 focus:ring-[var(--blue-glow)] transition-colors"
                />
              </div>

              {error && (
                <div className="flex items-start gap-2 mb-4 px-3 py-2.5 rounded-lg bg-[var(--red)]/15 border border-[var(--red)]/30 text-[13px] text-[var(--red-b)]">
                  <Icon name="error" size="sm" className="shrink-0 mt-px" /> <span>{error}</span>
                </div>
              )}

              <button
                type="submit"
                disabled={busy}
                className="w-full h-11 rounded-xl bg-[var(--blue)] hover:bg-[var(--blue-d)] text-white text-sm font-semibold transition-all active:scale-[.98] disabled:opacity-60 flex items-center justify-center gap-2 shadow-[0_4px_16px_var(--blue-glow)]"
              >
                {busy ? (
                  <>
                    <span className="w-4 h-4 rounded-full border-2 border-white/40 border-t-white animate-spin" />
                    กำลังเข้าสู่ระบบ...
                  </>
                ) : (
                  <>
                    <Icon name="login" size="sm" /> เข้าสู่ระบบ
                  </>
                )}
              </button>
            </form>
          </div>

          <div className="mt-5 flex items-center justify-center gap-3 text-[10.5px] text-white/30">
            <span className="flex items-center gap-1"><Icon name="shield" size="xs" /> Secure Access</span>
            <span className="w-1 h-1 rounded-full bg-white/20" />
            <span>© INTRANIN EMS</span>
          </div>
        </div>
      </div>
    </div>
  );
}
