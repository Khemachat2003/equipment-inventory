// Help — คู่มือใช้งานแบบสั้น (3 งานหลัก) — เป้าหมาย: ผู้ใช้ใหม่ทำเองได้ ไม่ต้องมีคนสอน
// เข้าถึงได้จากทุกหน้า: ปุ่ม ❓ ที่แถบบน + รายการ "วิธีใช้งาน" ใน sidebar + ทัวร์ครั้งแรก

import { useNavigate } from 'react-router-dom';
import Icon from '../components/ui/Icon.jsx';

const GUIDES = [
  {
    icon: 'inventory_2',
    box: 'bg-[var(--emerald-l)] text-[var(--emerald-d)]',
    title: 'เบิก–คืนของ',
    to: '/stock',
    cta: 'ไปหน้าเบิก–คืนของ',
    steps: [
      'เข้าเมนู "เบิก–คืนของ" (หรือกดจากหน้าแรก)',
      'พิมพ์ค้นหาชื่อ/รหัสอุปกรณ์ แล้วกดปุ่ม "เบิก" ตรงแถวนั้น',
      'เบิกหลายรายการ → ปรับจำนวนในตะกร้า แล้วกด "ยืนยันเบิกทั้งหมด"',
      'คืนของ → กดปุ่ม "คืนของจาก Site" ด้านบนตาราง',
    ],
  },
  {
    icon: 'document_scanner',
    box: 'bg-[var(--blue-l)] text-[var(--blue)]',
    title: 'ย้าย/โอนอุปกรณ์',
    to: '/scan',
    cta: 'ไปหน้าสแกน',
    steps: [
      'เข้าเมนู "ย้าย/โอนอุปกรณ์" (หรือกดปุ่ม สแกน ที่แถบบน)',
      'สแกนบาร์โค้ด/QR ที่ติดอุปกรณ์ หรือพิมพ์ Serial แล้วกด Enter',
      'เมื่อเจออุปกรณ์ → ดูตำแหน่งปัจจุบันและประวัติได้เลย',
      'กด "โอนย้าย" → เลือกปลายทาง → ยืนยัน',
    ],
  },
  {
    icon: 'qr_code_2',
    box: 'bg-[var(--amber-l)] text-[var(--amber-d)]',
    title: 'พิมพ์ฉลาก QR',
    to: '/qr',
    cta: 'ไปหน้าพิมพ์ฉลาก',
    steps: [
      'เข้าเมนู "พิมพ์ฉลาก QR" (หรือกดปุ่ม ฉลาก ที่แถบบน)',
      'เลือกอุปกรณ์ที่ต้องการจากตาราง',
      'เลือกโหมด Label (ทีละใบ) หรือ A4 (หลายดวง) แล้วกดพิมพ์',
    ],
  },
];

const TIPS = [
  { icon: 'search', text: 'หาของอยู่ไหน → ใช้ช่องค้นหาบนหน้าแรก พิมพ์ชื่อ/รหัส/Serial แล้วเห็นตำแหน่ง + ปุ่มทำต่อทันที' },
  { icon: 'place', text: 'ดูภาพรวมทั้งฟาร์ม → เมนู "เพิ่มเติม" ▸ "ของอยู่ฟาร์มไหน"' },
  { icon: 'history', text: 'ดูว่าใครเบิก/คืนอะไรไป → กด "ดูทั้งหมด" ในการ์ด ล่าสุด บนหน้าแรก' },
  { icon: 'info', text: 'ปุ่มลัดที่แถบบน: สแกน (อ่านบาร์โค้ด) · ฉลาก (พิมพ์ QR) · ❓ (คู่มือหน้านี้)' },
];

export default function Help() {
  const navigate = useNavigate();
  return (
    <div className="max-w-4xl space-y-5">
      {/* Intro */}
      <div className="rounded-2xl bg-[var(--ink)] text-white p-5 sm:p-6 shadow-[var(--sh-md)] relative overflow-hidden">
        <div className="absolute -top-20 -right-10 w-56 h-56 rounded-full bg-[var(--blue)]/25 blur-3xl pointer-events-none" />
        <div className="relative">
          <div className="text-[15px] font-bold">ใช้งานง่ายใน 3 งานหลัก</div>
          <div className="text-[13px] text-white/60 mt-1">
            ระบบนี้ออกแบบให้ทำงานจบในไม่กี่คลิก — เลือกงานที่ต้องการด้านล่าง แล้วทำตามขั้นตอนได้เลย
          </div>
        </div>
      </div>

      {/* 3 งานหลัก */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        {GUIDES.map((g) => (
          <div key={g.to} className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4 flex flex-col">
            <div className="flex items-center gap-2.5 mb-3">
              <span className={`flex items-center justify-center w-9 h-9 rounded-xl ${g.box}`}>
                <Icon name={g.icon} size="sm" />
              </span>
              <div className="text-[14px] font-bold text-[var(--text)]">{g.title}</div>
            </div>
            <ol className="space-y-2 flex-1">
              {g.steps.map((s, i) => (
                <li key={i} className="flex items-start gap-2 text-[12px] text-[var(--tsub)] leading-relaxed">
                  <span className="flex items-center justify-center w-4.5 h-4.5 rounded-full bg-[var(--surface2)] border border-[var(--g200)] text-[10px] font-bold text-[var(--tsub)] shrink-0 mt-px">
                    {i + 1}
                  </span>
                  <span>{s}</span>
                </li>
              ))}
            </ol>
            <button
              onClick={() => navigate(g.to)}
              className="mt-3.5 h-9 px-3 rounded-lg bg-[var(--blue)] text-white text-[12px] font-semibold hover:bg-[var(--blue-d)] flex items-center justify-center gap-1.5"
            >
              {g.cta} <Icon name="chevron_right" size="sm" />
            </button>
          </div>
        ))}
      </div>

      {/* เคล็ดลับ */}
      <div className="rounded-2xl bg-white border border-[var(--g200)] shadow-[var(--sh-sm)] p-4">
        <div className="flex items-center gap-2 text-[13px] font-semibold text-[var(--text)] mb-3">
          <span className="flex items-center justify-center w-7 h-7 rounded-lg bg-[var(--surface2)] text-[var(--tsub)]">
            <Icon name="lightbulb" size="sm" />
          </span>
          เคล็ดลับการใช้งาน
        </div>
        <ul className="space-y-2.5">
          {TIPS.map((t, i) => (
            <li key={i} className="flex items-start gap-2.5 text-[12px] text-[var(--tsub)] leading-relaxed">
              <span className="text-[var(--blue)] shrink-0 mt-px"><Icon name={t.icon} size="sm" /></span>
              <span>{t.text}</span>
            </li>
          ))}
        </ul>
      </div>

      {/* ติดต่อผู้ดูแล */}
      <div className="text-center text-[12px] text-[var(--tmuted)]">
        ใช้งานไม่ได้ตามคู่มือ หรือพบปัญหา → แจ้งผู้ดูแลระบบ (เข้าสู่ระบบเป็นผู้ดูแลที่เมนูด้านซ้ายล่าง)
      </div>
    </div>
  );
}