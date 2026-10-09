// farmMonitor.js — ตัวช่วยรวมข้อมูล "สุขภาพฟาร์ม" สำหรับ Executive Overview บนหน้าแรก
// (Farm Health & Status Summary + Quick Farm Access Cards + FarmDetailModal)
//
// กฎการนับ = กฎเดียวกับหน้า Farm Monitor (/farm) — อ่านตำแหน่งด้วยกติกาเดียวกับ utils/location.js:
//   • อุปกรณ์ที่อยู่ในชุด (Bundle) ที่ Deployed → นับที่ฟาร์มตามคอลัมน์ Location ของชุด
//     (คอลัมน์ SiteName ของสมาชิกเป็นชื่อชุด ไม่ใช่ชื่อฟาร์ม)
//   • อุปกรณ์ตัวเดี่ยว → นับที่ฟาร์มตาม SiteName
//   • คลังกลาง (Intranin / Stock / ไม่ระบุไซต์) ไม่ถือเป็นฟาร์ม → ไม่โชว์เป็นการ์ดฟาร์ม
//     (ตัวเลข "รอติดตั้ง" ใช้จาก /api/dashboard-full แยกต่างหาก)

const STOCK_SITES = new Set(['Intranin', 'Stock', 'ไม่ระบุไซต์']);

function clean(value) {
  const s = (value ?? '').toString().trim();
  if (!s || s === '-' || s === 'null' || s === 'undefined') return '';
  return s;
}

/** แผนที่ bundleId → ฟาร์ม (เฉพาะชุดที่ Deployed และระบุ Location ไว้) */
export function buildBundleFarmMap(bundles = []) {
  const map = {};
  (bundles || []).forEach((b) => {
    if (!b) return;
    const id = clean(b.bundleId);
    const loc = clean(b.location);
    if (b.status === 'Deployed' && id && loc) map[id] = loc;
  });
  return map;
}

/**
 * ฟาร์มของอุปกรณ์ 1 ชิ้น (ตามกฎข้างบน) — คืน '' ถ้าไม่ได้ติดตั้งที่ฟาร์มใด (เช่น อยู่คลัง)
 * @param {object} asset แถวจาก /api/assets
 * @param {Record<string,string>} bundleFarmMap จาก buildBundleFarmMap()
 */
export function farmOfAsset(asset, bundleFarmMap = {}) {
  if (!asset) return '';
  const bundleId = clean(asset.bundleId);
  if (bundleId) return bundleFarmMap[bundleId] || '';
  const site = clean(asset.siteName);
  if (!site || STOCK_SITES.has(site) || clean(asset.location) === 'Stock') return '';
  return site;
}

/** อุปกรณ์ทั้งหมดที่ติดตั้งอยู่ "ที่ฟาร์มนี้" (ตัวเดี่ยว + สมาชิกชุดที่ deploy มาฟาร์มนี้) */
export function farmAssetsOf(farm, assets = [], bundleFarmMap = {}) {
  return (assets || []).filter((a) => farmOfAsset(a, bundleFarmMap) === farm);
}

/**
 * สรุปภาพรวมรายฟาร์ม — ใช้แสดงการ์ด Quick Farm Access บนหน้าแรก
 * ฟาร์มที่ลงทะเบียนใน /api/farms จะโชว์เสมอ (แม้ยังไม่มีอุปกรณ์)
 * ฟาร์มที่พบจากข้อมูลอุปกรณ์แต่ไม่ได้ลงทะเบียนก็โชว์เช่นกัน (กันข้อมูลหาย)
 *
 * @returns {Array<{name:string, type:string, province:string, assets:number, bundles:number, houses:number, ok:number, rep:number}>}
 */
export function buildFarmOverview({ assets = [], bundles = [], farms = [] } = {}) {
  const bundleFarm = buildBundleFarmMap(bundles);

  // จำนวนชุด (ตู้) ที่ deploy อยู่รายฟาร์ม — นับที่ตัวชุด ไม่ใช่ที่สมาชิก (กันนับซ้ำ)
  const bundleCount = {};
  Object.values(bundleFarm).forEach((farm) => {
    bundleCount[farm] = (bundleCount[farm] || 0) + 1;
  });

  const farmMeta = {};
  (farms || []).forEach((f) => {
    // รองรับ 2 รูปแบบ: /api/farms (farmName) และ /api/dashboard-full.farmSites (siteName)
    const name = clean(f?.siteName || f?.farmName);
    if (name && !STOCK_SITES.has(name)) farmMeta[name] = f;
  });

  const byFarm = {};
  const ensure = (name) => {
    if (!byFarm[name]) {
      const meta = farmMeta[name] || {};
      byFarm[name] = {
        name,
        type: clean(meta.farmType) || 'อื่นๆ',
        province: clean(meta.province),
        assets: 0,
        bundles: bundleCount[name] || 0,
        houseSet: new Set(),
        ok: 0,
        rep: 0,
      };
    }
    return byFarm[name];
  };

  // ฟาร์มที่ลงทะเบียนไว้ → การ์ดโชว์เสมอ (แม้ 0 อุปกรณ์)
  Object.keys(farmMeta).forEach((name) => ensure(name));

  (assets || []).forEach((a) => {
    const farm = farmOfAsset(a, bundleFarm);
    if (!farm) return;
    const f = ensure(farm);
    f.assets += 1;
    const h = clean(a.houseName) || clean(a.houseId);
    if (h) f.houseSet.add(h);
    const status = clean(a.status);
    if (status.includes('ใช้งานได้')) f.ok += 1;
    if (status.includes('ซ่อม')) f.rep += 1;
  });

  // ฟาร์มที่มีชุด deploy แต่ยังไม่มีสมาชิก → โผล่การ์ดให้เห็นว่าฟาร์มนี้มีตู้อยู่
  Object.values(bundleFarm).forEach((farm) => ensure(farm));

  return Object.values(byFarm)
    .map(({ houseSet, ...f }) => ({ ...f, houses: houseSet.size }))
    .sort((a, b) => (b.assets - a.assets) || a.name.localeCompare(b.name, 'th'));
}