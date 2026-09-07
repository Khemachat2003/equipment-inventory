export const CATEGORY_FALLBACK = [
  { name: 'Sensors & Weather', label: 'เซนเซอร์ & สภาพอากาศ', icon: 'sensors' },
  { name: 'Controllers & Microcontrollers', label: 'คอนโทรลเลอร์ & ไมโครคอนโทรลเลอร์', icon: 'memory' },
  { name: 'Networking & Comms', label: 'เครือข่าย & การสื่อสาร', icon: 'router' },
  { name: 'Power Supplies', label: 'แหล่งจ่ายไฟ', icon: 'power' },
  { name: 'Power Protection & Switching', label: 'อุปกรณ์ป้องกัน & สวิตชิ่ง', icon: 'shield' },
  { name: 'Display & Peripherals', label: 'จอแสดงผล & อุปกรณ์ต่อพ่วง', icon: 'monitor' },
  { name: 'IO & Automation Modules', label: 'โมดูล IO & ระบบอัตโนมัติ', icon: 'settings_input_component' },
];

export const DEFAULT_CATEGORY_ICON = 'devices_other';

export function categoryIcon(name) {
  if (!name) return DEFAULT_CATEGORY_ICON;
  const c = CATEGORY_FALLBACK.find((x) => x.name.toLowerCase() === name.toLowerCase());
  return c ? c.icon : DEFAULT_CATEGORY_ICON;
}

export function categoryLabel(name) {
  if (!name) return '';
  const c = CATEGORY_FALLBACK.find((x) => x.name.toLowerCase() === name.toLowerCase());
  return c ? c.label : name;
}