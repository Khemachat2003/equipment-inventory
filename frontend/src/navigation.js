// ── Central navigation + page meta ───────────────────────────────────────────
// แหล่งเดียวสำหรับ: Nav Pills บน Topbar (TOPBAR_PILLS), เมนูเสริมใน drawer (DRAWER_GROUPS),
// เมนูบนมือถือ (MOBILE_NAV) และชื่อ/ชื่อรองหน้า (ROUTE_META)
// ถ้าเพิ่มหน้าใหม่ ต้องลงทะเบียนที่นี่ครบ (NAV_GROUPS + ROUTE_META)
// ─────────────────────────────────────────────────────────────────────────────
// แนวคิด UX (UX simplify): เมนูเรียงตาม "งานที่ผู้ใช้ต้องทำ" ไม่ใช่ "โครงสร้างข้อมูล"
//   - งานประจำวัน = สิ่งที่ใช้ทุกวัน มองเห็นทันที
//     Step 3: **Farm Monitor** (ของอยู่ฟาร์มไหน) ยกขึ้นเมนูหลักติดอันดับ 2 — จุดขายหลักสำหรับ
//     Executive/เจ้านายและช่างหน้างาน เข้าถึงภาพรวมฟาร์มได้ 1-Click (เดิมซ่อนใน "เพิ่มเติม")
//     รวมถึง **ชุดติดตั้งฟาร์ม/Bundle** ที่ขึ้นเมนูหลักแล้วเช่นกัน (เดิมซ่อนใน "เพิ่มเติม")
//   - เพิ่มเติม = หน้าข้อมูล/เครื่องมือที่ใช้ไม่บ่อย — โชว์ตรงเป็น Clean List ทันทีเมื่อเปิด drawer
//     (ชั้นพับ/accordion "เพิ่มเติม" ถูกถอนออกถาวรตาม spec — ไม่มี flag UI_FLAGS.simpleMenu แล้ว)
//   - ผู้ดูแลระบบ (adminOnly) = ซ่อนจากผู้ใช้ทั่วไป
// Route เดิมทุกเส้นยังใช้งานได้ตามลิงก์เดิม เพียงแค่ไม่โชว์ในเมนูหลัก

export const NAV_GROUPS = [
  {
    label: 'งานประจำวัน',
    items: [
      // short = ชื่อย่อบน Nav Pills กลาง Topbar เมื่อจอกว้างระดับ lg (ชื่อเต็มโชว์ที่จอ ≥xl)
      // กติกา: short ต้องเป็น substring ของ label (กัน label-in-name ตอนใช้ aria-label)
      { to: '/', label: 'หน้าแรก', icon: 'home', end: true },
      { to: '/farm', label: 'Farm Monitor', icon: 'monitoring' },
      { to: '/bundle', label: 'ชุดติดตั้งฟาร์ม', short: 'ชุดติดตั้ง', icon: 'folder_open' },
      { to: '/stock', label: 'เบิก–คืนของ', short: 'เบิก–คืน', icon: 'inventory_2' },
      { to: '/scan', label: 'ย้าย/โอนอุปกรณ์', short: 'ย้าย/โอน', icon: 'document_scanner' },
    ],
  },
  {
    label: 'เพิ่มเติม',
    items: [
      { to: '/asset', label: 'ทะเบียนรายชิ้น', icon: 'devices' },
      { to: '/qr', label: 'พิมพ์ฉลาก QR', icon: 'qr_code_2' },
      { to: '/history', label: 'ประวัติการเบิก–คืน', icon: 'history' },
      { to: '/report', label: 'รายงาน PDF', icon: 'bar_chart' },
      // Dashboard เดิม (มีกราฟ) — เก็บไว้ให้ admin ใช้ต่อ (ดู UI_FLAGS.charts)
      { to: '/dashboard', label: 'Dashboard แบบเต็ม', icon: 'dashboard', adminOnly: true },
    ],
  },
  {
    label: 'ผู้ดูแลระบบ',
    adminOnly: true,
    items: [
      { to: '/admin-tools', label: 'Admin Tools', icon: 'admin_panel_settings' },
      { to: '/backup', label: 'ดูข้อมูล Backup', icon: 'analytics' },
      { to: '/audit', label: 'Audit Log', icon: 'fact_check' },
      { to: '/users', label: 'ผู้ใช้', icon: 'group' },
      { to: '/settings', label: 'ตั้งค่า', icon: 'settings' },
    ],
  },
];

// ── B2 Concept — Topbar Nav Pills + Drawer "เมนูเสริม" ───────────────────────
// TOPBAR_PILLS: เมนูหลัก 5 รายการ (กลุ่ม "งานประจำวัน") ขึ้นแถบเม็ดกลาง Topbar
// DRAWER_GROUPS: กลุ่มที่เหลือ ("เพิ่มเติม" + "ผู้ดูแลระบบ") — เปิดด้วยปุ่มสามขีดขวาบน
export const TOPBAR_PILLS = NAV_GROUPS[0].items;
export const DRAWER_GROUPS = NAV_GROUPS.slice(1);

// Bottom nav บนมือถือ — 3 งานหลัก (md:hidden ใน Layout)
export const MOBILE_NAV = [
  { to: '/', label: 'หน้าแรก', icon: 'home', end: true },
  { to: '/stock', label: 'เบิก–คืน', icon: 'inventory_2' },
  { to: '/scan', label: 'ย้าย/โอน', icon: 'document_scanner' },
  { to: '/qr', label: 'พิมพ์ฉลาก', icon: 'qr_code_2' },
];

// ชื่อ + ชื่อรอง ของแต่ละหน้า (ค่าควรตรงกับ header ในไฟล์ page เอง)
export const ROUTE_META = {
  '/': { title: 'หน้าแรก', subtitle: 'ค้นหาอุปกรณ์และเริ่มงานได้จากที่เดียว' },
  '/stock': { title: 'เบิก–คืนของ', subtitle: 'รายการอุปกรณ์ในคลัง — เบิก / คืน / เพิ่มของ' },
  '/asset': { title: 'ทะเบียนรายชิ้น', subtitle: 'ติดตามอุปกรณ์ทุกชิ้น — สถานะ ตำแหน่ง และประวัติ' },
  '/bundle': { title: 'ชุดติดตั้งฟาร์ม', subtitle: 'จัดการชุดอุปกรณ์ติดตั้งที่ฟาร์ม — เพิ่มอุปกรณ์เข้าชุด / ส่งไปติดตั้ง / เบิกคืนคลังกลาง' },
  '/farm': { title: 'Farm Monitor', subtitle: 'ภาพรวมสุขภาพฟาร์มและอุปกรณ์ทุกไซต์งาน — ค้นหาของว่าอยู่ฟาร์มไหน โรงเรือนใด' },
  '/scan': { title: 'ย้าย/โอนอุปกรณ์', subtitle: 'สแกน Barcode / Serial เพื่อระบุตำแหน่งหรือโอนย้าย' },
  '/qr': { title: 'พิมพ์ฉลาก QR', subtitle: 'สร้างฉลาก Barcode ตัวเดียวหรือ A4 หลายดวง' },
  '/history': { title: 'ประวัติการเบิก–คืน', subtitle: 'ดูประวัติการโอนย้ายทั้งหมดในระบบ' },
  '/report': { title: 'รายงาน PDF', subtitle: 'Export ข้อมูลเป็น PDF' },
  '/dashboard': { title: 'Dashboard แบบเต็ม', subtitle: 'ภาพรวม + กราฟสถิติ (ผู้ดูแลระบบ)' },
  '/settings': { title: 'Settings', subtitle: 'ข้อมูลระบบและผู้ใช้งาน' },
  '/admin-tools': { title: 'เครื่องมือผู้ดูแล', subtitle: 'Backup / Cache / ระบบ' },
  '/audit': { title: 'Audit Log', subtitle: 'บันทึกการใช้งานระบบ (Admin)' },
  '/backup': { title: 'ข้อมูล Backup', subtitle: 'ดูข้อมูล PostgreSQL Backup และ Export CSV (Admin)' },
  '/users': { title: 'จัดการผู้ใช้', subtitle: 'ตั้งสิทธิ์ / รีเซ็ตรหัสผ่าน / ลบผู้ใช้' },
};

// หา meta ตาม path ปัจจุบัน; path ไม่รู้จัก → กลับไปหน้าแรก
export function getRouteMeta(pathname) {
  return ROUTE_META[pathname] || ROUTE_META['/'];
}