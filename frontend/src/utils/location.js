// ตัวช่วยแปลง "ตำแหน่ง" ของอุปกรณ์ ให้เป็นเส้นทางเต็มที่คนอ่านรู้เรื่อง
//
// โครงสร้างตำแหน่ง 3 ชั้น (ไล่จากบนลงล่าง):
//   1. ฟาร์ม       = ชื่อฟาร์ม/ไซต์งาน          (Asset_List คอลัมน์ H = SiteName)
//   2. โรงเรือน     = รหัส + ชื่อโรงเรือน          (คอลัมน์ L / M = HouseID / HouseName)
//   3. จุดติดตั้ง   = ตำแหน่งย่อยลึกลงไปอีก     (คอลัมน์ G = Location)
//
// ค่าว่าง / "-" / "Stock" / "Intranin" ถือว่ายังไม่มีตำแหน่ง
// ถ้าอยู่ที่คลังกลาง จะไม่มีชั้นที่ 2 และ 3 เพราะไม่มีความหมาย

const STOCK_LABEL = 'คลังกลาง';

function clean(value) {
  const s = (value ?? '').toString().trim();
  if (!s) return '';
  if (s === '-' || s === 'null' || s === 'undefined') return '';
  return s;
}

function isStocky(siteName, location) {
  const site = clean(siteName);
  const loc = clean(location);
  if (site === 'Intranin') return true;
  if (loc === 'Stock') return true;
  if (!site && !loc) return true;
  return false;
}

/**
 * @param {object} input
 * @param {string} input.siteName   ชื่อฟาร์ม/ไซต์งาน          (คอลัมน์ H = SiteName)
 * @param {string} input.houseName  ช���่อโรงเรือน              (คอลัมน์ M = HouseName)
 * @param {string} input.houseId    รหัสโรงเรือน                (คอลัมน์ L = HouseID)
 * @param {string} input.location   จุดติดตั้ง                  (คอลัมน์ G = Location)
 * @param {string} input.bundleId   ถ้ามีค่า = ตัวนี้อยู่ในชุด  (คอลัมน์ N = BundleID)
 * @returns {{kind:string, farm:string, house:string, point:string, full:string, chips:Array}}
 */
export function buildLocation(input = {}) {
  const siteName = clean(input.siteName);
  const houseName = clean(input.houseName);
  const houseId = clean(input.houseId);
  const inBundle = !!clean(input.bundleId);
  const rawLoc = clean(input.location);

  // ⚠️ อุปกรณ์ที่อยู่ในชุด (Bundle) เก็บข้อมูล "กลับกัน" กับอุปกรณ์เดี่ยว:
  //    คอลัมน์ H = ชื่อชุด (SiteName ถูกล็อกให้เป็นชื่อ Bundle เสมอ)
  //    คอลัมน์ G = ชื่อฟาร์มที่ชุดนั้นไปติดตั้ง
  //    → ชั้น "ฟาร์ม" ของตัวนี้ต้องอ่านจาก G ไม่ใช่ H
  let farm = inBundle ? rawLoc : siteName;
  let point = inBundle ? '' : rawLoc;

  // อยู่คลังกลาง — ไม่มีชั้นโรงเรือน/จุดติดตั้ง
  if (isStocky(inBundle ? '' : siteName, inBundle ? rawLoc : rawLoc)) {
    const inStock = rawLoc === 'Stock';
    return {
      kind: 'stock',
      farm: '', house: '', point: '',
      full: inStock ? `${STOCK_LABEL} (Stock)` : `${STOCK_LABEL} (Intranin)`,
      chips: [],
    };
  }

  farm = farm || siteName;

  // ข้อมูลเก่า: เคยเก็บชื่อฟาร์มไว้ในคอลัมน์ Location — ไม่ต้องแสดงซ้ำสองชั้น
  if (point && point === farm) point = '';

  const house = houseId ? `${houseId} ${houseName}`.trim() : houseName;

  const full = [farm, house, point].filter(Boolean).join('  ›  ');

  const chips = [
    { key: 'farm', label: 'ฟาร์ม', value: farm },
    { key: 'house', label: 'โรงเรือน', value: house },
    { key: 'point', label: 'จุดติดตั้ง', value: point },
  ].filter((c) => c.value);

  return { kind: 'site', farm, house, point, full, chips };
}

/**
 * รวมตำแหน่งของชุดอุปกรณ์จากสมาชิกทุกชิ้น
 * @returns {{main:object, divergent:boolean, paths:string[]}}
 */
export function buildBundleLocation(assets = []) {
  const all = assets.map((a) => buildLocation({
    siteName: a.site ?? a.siteName,
    houseName: a.houseName,
    houseId: a.houseId,
    location: a.location,
    bundleId: a.bundleId || 'bundle',
  }));
  if (!all.length) {
    return { main: buildLocation({}), divergent: false, paths: [] };
  }
  const paths = [...new Set(all.map((l) => l.full))];
  return { main: all[0], divergent: paths.length > 1, paths };
}
