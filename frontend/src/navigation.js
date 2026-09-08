// ── Central navigation + page meta ───────────────────────────────────────────
// แหล่งเดียวสำหรับ: เมนู sidebar (NAV_GROUPS) + ชื่อ/ชื่อรองหน้า (ROUTE_META)
// ถ้าเพิ่มหน้าใหม่ ต้องลงทะเบียนที่นี่ครบทั้ง 2 ตัว (topbar จะขึ้นชื่อให้ตรงกันเอง)
// ─────────────────────────────────────────────────────────────────────────────

export const NAV_GROUPS = [
  {
    label: 'หลัก',
    items: [
      { to: '/', label: 'Dashboard', icon: 'dashboard', end: true },
      { to: '/stock', label: 'Stock', icon: 'inventory_2' },
      { to: '/asset', label: 'Asset', icon: 'devices' },
      { to: '/bundle', label: 'Bundle', icon: 'folder_open' },
    ],
  },
  {
    label: 'จัดการ',
    items: [
      { to: '/farm', label: 'ฟาร์ม', icon: 'agriculture' },
      { to: '/history', label: 'ประวัติ', icon: 'history' },
      { to: '/report', label: 'รายงาน', icon: 'bar_chart' },
      { to: '/settings', label: 'ตั้งค่า', icon: 'settings', adminOnly: true },
    ],
  },
  {
    label: 'ผู้ดูแลระบบ',
    adminOnly: true,
    items: [
      { to: '/admin-tools', label: 'Admin Tools', icon: 'admin_panel_settings' },
      { to: '/audit', label: 'Audit Log', icon: 'fact_check' },
      { to: '/users', label: 'ผู้ใช้', icon: 'group' },
    ],
  },
];

// ชื่อ + ชื่อรอง ของแต่ละหน้า (ค่าควรตรงกับ header ในไฟล์ page เอง)
export const ROUTE_META = {
  '/': { title: 'Dashboard', subtitle: 'ภาพรวมระบบ ณ วันนี้' },
  '/stock': { title: 'Stock', subtitle: 'รายการอุปกรณ์ทั้งหมด' },
  '/asset': { title: 'Asset Tracking', subtitle: 'ติดตามอุปกรณ์แยกตาม Part Number' },
  '/bundle': { title: 'Bundle', subtitle: 'จัดการชุดอุปกรณ์' },
  '/farm': { title: 'Farm Monitor', subtitle: 'ภาพรวมอุปกรณ์และชุดอุปกรณ์ตามฟาร์ม' },
  '/history': { title: 'ประวัติการเบิก–คืน', subtitle: 'ดูประวัติการโอนย้ายทั้งหมดในระบบ' },
  '/report': { title: 'รายงาน', subtitle: 'Export ข้อมูลเป็น PDF' },
  '/settings': { title: 'Settings', subtitle: 'ข้อมูลระบบและผู้ใช้งาน' },
  '/admin-tools': { title: 'เครื่องมือผู้ดูแล', subtitle: 'Backup / Cache / ระบบ' },
  '/audit': { title: 'Audit Log', subtitle: 'บันทึกการใช้งานระบบ (Admin)' },
  '/users': { title: 'จัดการผู้ใช้', subtitle: 'ตั้งสิทธิ์ / รีเซ็ตรหัสผ่าน / ลบผู้ใช้' },
};

// หา meta ตาม path ปัจจุบัน; path ไม่รู้จัก → กลับไปหน้าแรก
export function getRouteMeta(pathname) {
  return ROUTE_META[pathname] || ROUTE_META['/'];
}