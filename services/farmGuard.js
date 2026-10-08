// services/farmGuard.js
// ตรวจซ้ำก่อนเพิ่ม Farm_Sites / Farm_Houses — pure functions (ไม่ต่อ Google Sheets)
// เดิม POST /api/add-farm-site, /api/add-farm-house เพิ่มซ้ำได้เรื่อย ๆ (ไม่เช็ค duplicate)
// ตอนนี้ route เรียกใช้ 2 ฟังก์ชันนี้ก่อน append → ซ้ำจะตอบ 409 พร้อมข้อความไทย
// เป็น pure function เพื่อให้ unit test ได้โดยไม่ต้อง mock sheets

function norm(v) {
  return String(v ?? "").trim().toUpperCase();
}

function cleanName(v) {
  return String(v ?? "").trim();
}

/**
 * เช็คว่า siteId หรือ siteName นี้มีอยู่แล้วหรือไม่ (เทียบแบบ trim + case-insensitive)
 * @param {Array<{siteId?:string, siteName?:string}>} sites แถวที่อ่านจากชีต Farm_Sites
 * @param {{siteId?:string, siteName?:string}} input
 * @returns {null | {code:409, error:string, field:string}} null = ผ่าน, ไม่ซ้ำ
 */
function findSiteDuplicate(sites, { siteId, siteName } = {}) {
  const list = Array.isArray(sites) ? sites : [];
  const id = norm(siteId);
  const name = cleanName(siteName);

  if (id) {
    for (const s of list) {
      if (norm(s && s.siteId) === id) {
        return { code: 409, error: "รหัสฟาร์มนี้มีในระบบแล้ว", field: "siteId" };
      }
    }
  }
  if (name) {
    for (const s of list) {
      if (cleanName(s && s.siteName) === name) {
        return { code: 409, error: "ชื่อฟาร์มนี้มีในระบบแล้ว", field: "siteName" };
      }
    }
  }
  return null;
}

/**
 * เช็คว่า houseId นี้มีอยู่แล้วหรือไม่ "ภายในฟาร์ม (siteId) เดิม"
 * (รหัสเดียวกันแต่คนละฟาร์มถือว่าไม่ซ้ำ — โรงเรือนเป็นลูกของฟาร์ม)
 * @param {Array<{houseId?:string, siteId?:string}>} houses แถวที่อ่านจากชีต Farm_Houses
 * @param {{siteId?:string, houseId?:string}} input
 * @returns {null | {code:409, error:string, field:string}}
 */
function findHouseDuplicate(houses, { siteId, houseId } = {}) {
  const list = Array.isArray(houses) ? houses : [];
  const sid = norm(siteId);
  const hid = norm(houseId);

  if (sid && hid) {
    for (const h of list) {
      if (norm(h && h.siteId) === sid && norm(h && h.houseId) === hid) {
        return { code: 409, error: "โรงเรือนนี้มีในฟาร์มนี้แล้ว", field: "houseId" };
      }
    }
  }
  return null;
}

module.exports = { findSiteDuplicate, findHouseDuplicate, norm };