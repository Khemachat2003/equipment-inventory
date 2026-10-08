// ── Central navigation + page meta ───────────────────────────────────────────
// แหล่งเดียวสำหรับ: เมนู sidebar (NAV_GROUPS) + ชื่อ/ชื่อรองหน้า (ROUTE_META)
// ถ้าเพิ่มหน้าใหม่ ต้องลงทะเบียนที่นี่ครบทั้ง 2 ตัว (topbar จะขึ้นชื่อให้ตรงกันเอง)
// ─────────────────────────────────────────────────────────────────────────────
// แนวคิด UX (UX simplify): เมนูเรียงตาม "งานที่ผู้ใช้ต้องทำ" ไม่ใช่ "โครงสร้างข้อมูล"
//   - งานประจำวัน (4 ปุ่ม) = สิ่งที่ใช้ทุกวัน มองเห็นทันที
//     (ชุดติดตั้งฟาร์ม/Bundle ขึ้นเมนูหลักแล้ว — เดิมซ่อนใน "เพิ่มเติม" ทั้งที่เป็นงานหลักของระบบ)
//   - เพิ่มเติม (collapsed) = หน้าข้อมูล/เครื่องมือที่ใช้ไม่บ่อย — พับเก็บ กางเมื่อต้องการ
//     (กลุ่มนี้ถูกพับหรือไม่ ขึ้นกับ UI_FLAGS.simpleMenu — ดู Layout.jsx)
//   - ผู้ดูแลระบบ (adminOnly) = ซ่อนจากผู้ใช้ทั่วไป
// Route เดิมทุกเส้นยังใช้งานได้ตามลิงก์เดิม เพียงแค่ไม่โชว์ในเมนูหลัก

export const NAV_GROUPS = [
  {
    label: 'งานประจำวัน',
    items: [
      { to: '/', label: 'หน้าแรก', icon: 'home', end: true },
      { to: '/stock', label: 'เบิก–คืนของ', icon: 'inventory_2' },
      { to: '/scan', label: 'ย้าย/โอนอุปกรณ์', icon: 'document_scanner' },
      { to: '/bundle', label: 'ชุดติดตั้งฟาร์ม (Bundle)', icon: 'folder_open' },
    ],
  },
  {
    label: 'เพิ่มเติม',
    collapsed: true,
    items: [
      { to: '/farm', label: 'ของอยู่ฟาร์มไหน', icon: 'place' },
      { to: '/asset', label: 'ทะเบียนรายชิ้น (Asset)', icon: 'devices' },
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
  '/asset': { title: 'ทะเบียนรายชิ้น (Asset)', subtitle: 'ติดตามอุปกรณ์แยกตาม Part Number' },
  '/bundle': { title: 'ชุดติดตั้งฟาร์ม (Bundle)', subtitle: 'จัดการชุดอุปกรณ์ติดตั้งที่ฟาร์ม — เพิ่มอุปกรณ์เข้าชุด / ย้ายไปฟาร์ม / คืนเข้าคลัง' },
  '/farm': { title: 'ของอยู่ฟาร์มไหน', subtitle: 'ภาพรวมอุปกรณ์และชุดอุปกรณ์ตามฟาร์ม' },
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