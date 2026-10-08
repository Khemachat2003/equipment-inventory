// utils/bundleGroups.js
// ก้อน ⑥ (Issue A) — เกณฑ์จัดกลุ่มหน้า "ชุดติดตั้งฟาร์ม (Bundle)":
// แยกไฟล์ตามคำแนะนำ react(only-export-components) — ไฟล์หน้า (.jsx) export เฉพาะ component
// และให้เทสต์ตรรกะ (tests/bundleGroups.test.jsx) import ตรงจุดได้

// STOCK_KEY = ค่าพิเศษของ filterFarm สำหรับ chip "คลัง" (ชุดที่ไม่ผูกกับฟาร์มใด)
export const STOCK_KEY = '__stock__';

// "ชุดนี้อยู่โซนไหน" — In Stock (หรือไม่มีข้อมูลฟาร์มเลย) = คลัง ('') / นอกนั้น = คีย์ฟาร์ม
// fallback: farmId → farmName → location (ชุดเก่าที่ไม่มี farmId ก็ยังจับกลุ่มได้จากชื่อฟาร์ม/ตำแหน่ง)
// เกณฑ์นี้ตั้งใจให้ตรงกับ BundleCard (inStock = status === 'In Stock') เพื่อไม่ให้การ์ดกับกลุ่มขัดกัน
export function farmKeyOf(b) {
  if (b.status === 'In Stock') return '';
  return (b.farmId || b.farmName || b.location || '').trim();
}

// ก้อน ⑤ (B2) — เดารหัสชุดถัดไป "BDL-XXX" จากรหัสชุดที่มีอยู่ (เลขสูงสุด +1 เติมศูนย์ 3 หลัก — ถ้าชนไล่ต่อจนว่าง)
// รูปแบบอื่นที่ผู้ใช้ตั้งเองไม่ถูกใช้เป็นฐาน — ยังไม่มีชุดหมายเลขเลย → BDL-001
export function suggestBundleId(existingIds = []) {
  const used = new Set((existingIds || []).map((x) => (x || '').trim().toUpperCase()));
  let max = 0;
  used.forEach((id) => {
    const m = /^(?:BDL-)?(\d{1,6})$/.exec(id);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  let n = max + 1;
  while (used.has(`BDL-${String(n).padStart(3, '0')}`)) n += 1;
  return `BDL-${String(n).padStart(3, '0')}`;
}