// utils/labelReprint.js
// สแนปช็อต "พิมพ์ครั้งล่าสุด" สำหรับปุ่ม พิมพ์ซ้ำ ในหน้าฉลาก (QrPage)
// แยกจาก ems_label_last_v1 (ค่าตั้งค่าที่ sync อัตโนมัติทุกครั้งที่แก้) — ตัวนี้จับตอน "กดพิมพ์จริง"
// เพื่อให้พิมพ์ซ้ำได้ตรงกับชุดที่ออกเครื่องพิมพ์ล่าสุด (ค่า + Serial) แม้หลังจากนั้นแก้อะไรไปแล้ว
export const REPRINT_KEY = 'ems_label_reprint_v1';

// cfg = { mode, lW, lH, lMargin, pm, gapX, gapY, bgColor, showBadge } · serials = [...Serial]
// คืนสแนปช็อตที่บันทึก (หรือ null ถ้าไม่มี Serial — ไม่บันทึกให้)
export function saveReprint(cfg, serials) {
  const list = (serials || []).filter(Boolean).map(String);
  if (!list.length) return null;
  const snap = { cfg: cfg || {}, serials: list, at: Date.now() };
  try { localStorage.setItem(REPRINT_KEY, JSON.stringify(snap)); } catch { /* ignore */ }
  return snap;
}

// อ่านสแนปช็อตล่าสุด — คืน null ถ้าไม่มี/ข้อมูลเสีย (ไม่ throw)
export function readReprint() {
  try {
    const s = JSON.parse(localStorage.getItem(REPRINT_KEY) || 'null');
    if (!s || !Array.isArray(s.serials) || !s.serials.length) return null;
    return { cfg: s.cfg || {}, serials: s.serials.map(String), at: Number(s.at) || 0 };
  } catch { return null; }
}