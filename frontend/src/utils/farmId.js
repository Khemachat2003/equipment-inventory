// utils/farmId.js — ตัวช่วยสร้าง "รหัสฟาร์ม / รหัสโรงเรือน" อัตโนมัติ
// ใช้ร่วมกันทุกจุดที่เพิ่มฟาร์ม/โรงเรือน: หน้า Farm, TransferModal, Bundle DeployModal
// เป้าหมาย: ลดฟิลด์ที่ user ต้องกรอก — พิมพ์ชื่อแล้วระบบเดารหัสให้ (แก้ได้)

export const FARM_TYPES = ['สัตว์ปีก', 'สัตว์บก', 'สุกร', 'อื่นๆ'];

/** เหลือเฉพาะ A-Z 0-9 + uppercase (ใช้ทำรหัส) — maxLen กันรหัสยาวเกิน */
export function cleanId(value, maxLen = 12) {
  return String(value || '')
    .replace(/[^A-Za-z0-9]/g, '')
    .toUpperCase()
    .slice(0, maxLen);
}

/** รหัสฟาร์มจากชื่อ — เอาเฉพาะอังกฤษ/ตัวเลข (ชื่อไทยล้วน → คืน '' ให้ user กรอกเอง) */
export function suggestSiteId(name) {
  return cleanId(name);
}

/**
 * รหัสโรงเรือนถัดไปของฟาร์ม เช่น SITE-H01, SITE-H02 ...
 * ดูจากรหัสที่มีอยู่แล้วในฟาร์มนั้น แล้วไล่เลขต่อ (ไม่ทับ H ที่มีอยู่)
 * @param {string} siteId รหัสฟาร์ม
 * @param {string[]} existingHouseIds รหัสโรงเรือนที่มีอยู่ในฟาร์มนี้
 */
export function suggestHouseId(siteId, existingHouseIds = []) {
  const base = cleanId(siteId, 10);
  if (!base) return '';
  let max = 0;
  (existingHouseIds || []).forEach((h) => {
    // เทียบกับค่าดิบ (ไม่ cleanId เพราะ "-" ต้องเก็บไว้ในรูปแบบ SITE-H01)
    const m = String(h || '').trim().toUpperCase().match(`^${base}-H(\\d+)$`);
    if (m) max = Math.max(max, parseInt(m[1], 10) || 0);
  });
  return `${base}-H${String(max + 1).padStart(2, '0')}`;
}