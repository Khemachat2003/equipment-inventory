// Material Symbols icon component
// Usage: <Icon name="inventory_2" size="md" weight="regular" />
// แสดงผลผ่าน codepoint (ICON_CODEPOINTS) — subset font ถูกตัด GSUB ligature ทิ้ง
// → ใช้ codepoint ตรงๆ เพื่อให้ icon แสดงได้เสมอ
// ชื่อที่ยังไม่อยู่ใน subset → fallback เป็น .msi-full (full font โหลดเฉพาะเมื่อมีการใช้จริง)
//   render จากชื่อ (ligature) → icon แสดงได้ทันทีแม้ยังไม่ได้ regen subset
//   (vite-plugin-icon-sync จะ regen subset ให้เองทั้งตอน dev และ build)
import { ICON_CODEPOINTS } from '/src/data/iconCodepoints.js';

export default function Icon({
  name,
  size = 'md',
  weight = 'regular',
  fill = 'none',
  color,
  className = '',
  ...props
}) {
  const sizeMap = { xs: 16, sm: 20, md: 24, lg: 28, xl: 32, '2xl': 40 };
  const weightMap = { light: 300, regular: 400, medium: 500, semibold: 600, bold: 700 };
  const fillMap = { none: 0, subtle: 0.25, half: 0.5, full: 1 };
  const gradeMap = { normal: 0, dark: 100, light: -25 };

  const px = sizeMap[size] ?? 24;
  const cp = ICON_CODEPOINTS[name];
  const unmapped = cp === undefined;
  if (import.meta.env.DEV && unmapped) {
    console.warn(
      `[Icon] "${name}" ยังไม่อยู่ใน icon subset — แสดงผ่าน full font (fallback) อยู่\n` +
        '  ทางแก้ถาวร: รัน python scripts/make_icon_subset.py (หรือปล่อย vite-plugin-icon-sync จัดการตอน dev/build) เพื่อใส่ icon นี้ลง subset ให้โหลดเร็วขึ้น'
    );
  }

  return (
    <span
      className={`msi ${unmapped ? 'msi-full ' : ''}${className}`}
      style={{
        fontSize: px,
        color: color,
        fontVariationSettings: `'wght' ${weightMap[weight] ?? 400}, 'FILL' ${fillMap[fill] ?? 0}, 'GRAD' ${gradeMap.normal}, 'opsz' ${px}`,
      }}
      aria-hidden="true"
      {...props}
    >
      {unmapped ? name : String.fromCodePoint(cp)}
    </span>
  );
}
