import { useEffect, useState } from 'react';
import axios from 'axios';
import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';
import { useAuth } from '../context/AuthContext.jsx';

export default function Settings() {
  const { user } = useAuth();
  const navigate = useNavigate();
  const [sheetsUrl, setSheetsUrl] = useState('');

  useEffect(() => {
    axios.get('/api/settings/spreadsheet-url')
      .then(({ data }) => setSheetsUrl(data.url))
      .catch(() => setSheetsUrl(''));
  }, []);

  const username = user?.username || '-';
  const isAdmin = user?.role === 'admin';

  return (
    <div className="max-w-2xl space-y-4">
      {/* บัญชีผู้ใช้ */}
      <Group label="บัญชีผู้ใช้">
        <div className="flex items-center gap-4 p-4 rounded-xl border border-[var(--g100)]">
          <div className="w-11 h-11 rounded-full bg-[var(--blue-l)] text-[var(--blue)] flex items-center justify-center text-lg font-bold uppercase">
            {(username[0] || '?')}
          </div>
          <div className="flex-1">
            <div className="text-[14px] font-bold text-[var(--text)]">{username}</div>
            <div className="text-[12px] text-[var(--tmuted)] flex items-center gap-1.5">
              <span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)] inline-block" /> Online
            </div>
          </div>
          <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold ${isAdmin ? 'bg-[var(--blue-l)] text-[var(--blue)]' : 'bg-[var(--g100)] text-[var(--tsub)]'}`}>
            {isAdmin ? 'ผู้ดูแลระบบ' : 'ผู้ใช้งาน'}
          </span>
        </div>
      </Group>

      {/* ระบบ */}
      <Group label="ระบบ">
        <div className="rounded-xl border border-[var(--g100)] divide-y divide-[var(--g100)]">
          <Row icon="info" k="เวอร์ชันระบบ" v="2.0" />
          <Row icon="dns" k="เซิร์ฟเวอร์" v="Node.js" />
          <Row icon="table_view" k="ฐานข้อมูลหลัก" v="Google Sheets" />
          <Row icon="monitor_heart" k="สถานะเซิร์ฟเวอร์" v={<Ok>Running</Ok>} />
          <Row icon="storage" k="สถานะฐานข้อมูล" v={<Ok>Connected</Ok>} />
        </div>
      </Group>

      {/* เครื่องมือ */}
      <Group label="เครื่องมือ">
        <div className="rounded-xl border border-[var(--g100)] divide-y divide-[var(--g100)]">
          <Action icon="grid_on" tone="blue" title="Open Database" desc="เปิดดู Google Sheets ของระบบ">
            <a href={sheetsUrl || 'https://docs.google.com/spreadsheets/d/1CheIF--yOt2mRxubU1000TmIIjKpuzIExH-9O0RS7FA'} target="_blank" rel="noreferrer" className={ghostBtn}>เปิด</a>
          </Action>
          <Action icon="qr_code_scanner" tone="teal" title="Quick Scan" desc="เปิดหน้าสแกน QR / Barcode และระบุตำแหน่ง">
            <button onClick={() => navigate('/scan')} className={ghostBtn}>เปิด Scanner</button>
          </Action>
          <Action icon="fact_check" tone="amber" title="Audit Log" desc="ดูประวัติการกระทำทั้งหมดในระบบ (เฉพาะ Admin)">
            {isAdmin
              ? <button onClick={() => navigate('/audit')} className={ghostBtn}>เปิด</button>
              : <span className="text-[12px] text-[var(--tmuted)]">เฉพาะ Admin</span>}
          </Action>
          <Action icon="backup" tone="violet" title="Full System Backup" desc="สำรองข้อมูลทุก Sheet ไปยัง PostgreSQL, ดูและ Export เป็น CSV">
            {isAdmin
              ? <button onClick={() => navigate('/admin-tools')} className={ghostBtn}>Backup ทั้งระบบ</button>
              : <span className="text-[12px] text-[var(--tmuted)]">เฉพาะ Admin</span>}
          </Action>
          <Action icon="visibility" tone="green" title="ดูข้อมูล Backup" desc="ดูและ Export เป็น CSV ได้จากหน้านั้น">
            {isAdmin
              ? <button onClick={() => navigate('/backup')} className={ghostBtn}>เปิดหน้าดู</button>
              : <span className="text-[12px] text-[var(--tmuted)]">เฉพาะ Admin</span>}
          </Action>
        </div>
      </Group>
    </div>
  );
}

function Group({ label, children }) {
  return (
    <div className="space-y-1.5">
      <div className="px-1 text-[12px] font-bold text-[var(--tmuted)]">{label}</div>
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">{children}</div>
    </div>
  );
}

function Row({ icon, k, v }) {
  return (
    <div className="flex items-center justify-between px-4 py-3">
      <span className="flex items-center gap-2 text-[13px] text-[var(--tsub)]">
        <Icon name={icon} size="sm" /> {k}
      </span>
      <span className="text-[13px] font-semibold text-[var(--text)]">{v}</span>
    </div>
  );
}

function Ok({ children }) {
  return <span className="flex items-center gap-1.5 text-[var(--emerald-d)]"><span className="w-1.5 h-1.5 rounded-full bg-[var(--emerald)] inline-block" />{children}</span>;
}

function Action({ icon, tone, title, desc, children }) {
  const tones = { blue: 'bg-[var(--blue-l)] text-[var(--blue)]', teal: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]', amber: 'bg-[var(--amber-l)] text-[var(--amber-d)]', violet: 'bg-[var(--purple-l)] text-[var(--purple)]', green: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]' };
  return (
    <div className="flex items-center gap-3 px-4 py-3">
      <span className={`flex items-center justify-center w-9 h-9 rounded-lg ${tones[tone]}`}><Icon name={icon} size="sm" /></span>
      <div className="flex-1 min-w-0">
        <div className="text-[13px] font-bold text-[var(--text)]">{title}</div>
        <div className="text-[11px] text-[var(--tmuted)]">{desc}</div>
      </div>
      <div className="flex-shrink-0">{children}</div>
    </div>
  );
}

const ghostBtn = 'flex items-center gap-1 px-3 py-1.5 rounded-lg border border-[var(--g300)] text-[12px] font-semibold text-[var(--tsub)] hover:bg-[var(--surface2)]';