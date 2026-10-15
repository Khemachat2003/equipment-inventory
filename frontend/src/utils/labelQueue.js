// labelQueue.js — ตัวช่วย "พิมพ์ฉลาก QR" (หน้า /qr): เรียงล่าสุดอยู่บน + กรองวันรับเข้า/เลข PO + จัดกลุ่มตามวันรับเข้า
// Pure functions ทั้งหมด — เทสอยู่ที่ tests/labelQueue.test.jsx
// หลักการ: "วันที่สร้าง" ของรายชิ้นมาจาก 3 ที่ (ใช้ตัวไหนที่มีก่อน):
//   ① createdAt   — โหมด Local Inventory (แม่นระดับวินาที)
//   ② Serial      — ระบบสร้าง Serial รูปแบบ SN-<PART>-<DDMMYYYY>-<รัน> → แม่นระดับวัน
//   ③ receivedAt  — วันที่รับเข้าจากล็อต (fallback สุดท้าย)

import { receivedDateKey } from './inbound.js';

// แยกวันที่จาก Serial (SN-<PART...>-<DDMMYYYY>-<รัน>) → timestamp เที่ยงวันนั้น (local) | null ถ้าไม่ใช่รูปแบบนี้
// (Part มีขีดกลางได้ เช่น SN-SENSOR-1-09102026-0004 → ยึด segment รูปแบบวันที่ ก่อนเลขรันท้ายสุดเสมอ)
export function serialDate(serial) {
  const m = /-([0-9]{2})([0-9]{2})([0-9]{4})-[0-9]+$/.exec(String(serial || '').trim());
  if (!m) return null;
  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
  const d = new Date(yyyy, mm - 1, dd, 12, 0, 0);
  // กันวันลอย (เช่น 31/02 → ม้วนเป็นเดือนถัดไป): ต้องวนกลับได้ตรงเดิมเท่านั้น
  if (d.getDate() !== dd || d.getMonth() !== mm - 1) return null;
  return d.getTime();
}

// วันที่สร้างรายชิ้น (timestamp ms) — ① createdAt → ② วันที่ใน Serial → ③ receivedAt → 0 (ไม่รู้วัน)
export function createdTsOf(asset) {
  const created = Date.parse(asset?.createdAt || '');
  if (Number.isFinite(created)) return created;
  const fromSerial = serialDate(asset?.serialNumber);
  if (fromSerial != null) return fromSerial;
  const received = Date.parse(asset?.receivedAt || '');
  return Number.isFinite(received) ? received : 0;
}

// วันที่รับเข้า (timestamp ms) — ใช้ receivedAt จริงก่อน; ไม่มี → ถือว่า "รับเข้าวันที่สร้าง"
// (อุปกรณ์เพิ่มใหม่ต้องโผล่ในตัวกรอง "รับเข้า 2 วันล่าสุด" เสมอ แม้ยังไม่ผูกล็อต)
export function receivedTsOf(asset) {
  const received = Date.parse(asset?.receivedAt || '');
  if (Number.isFinite(received)) return received;
  return createdTsOf(asset);
}

// key 'YYYY-MM-DD' แบบเวลาไทย (Asia/Bangkok) — ตัวเดียวกับหน้าทะเบียน (utils/inbound.js)
export function dateKeyOf(ts) {
  if (!Number.isFinite(ts) || ts <= 0) return '';
  return receivedDateKey(new Date(ts).toISOString());
}

// ป้ายวันแบบไทย '15 ต.ค. 2568' (จาก timestamp) | '' ถ้าไม่รู้วัน
export function thaiDayLabel(ts) {
  const key = dateKeyOf(ts);
  if (!key) return '';
  return new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', day: 'numeric', month: 'short', year: 'numeric' }).format(new Date(ts));
}

// เรียง "ล่าสุดอยู่บนสุดเสมอ" — วันที่สร้างมาก่อน (DESC); วันเดียวกัน → เลขรันท้าย Serial มากกว่าอยู่บน
export function compareLatestFirst(a, b) {
  const diff = createdTsOf(b) - createdTsOf(a);
  if (diff !== 0) return diff;
  return String(b?.serialNumber || b?.serial || '').localeCompare(String(a?.serialNumber || a?.serial || ''), undefined, { numeric: true });
}

// ตัวกรองรวมของหน้าพิมพ์ฉลาก — ค้นหา (Serial/ชื่อ) + สถานะ + เลข PO + ช่วงวันรับเข้า
// from/to เป็น key 'YYYY-MM-DD' (จาก ThaiDateInput); ช่องว่าง = ไม่กำหนดขอบ
export function matchesLabelFilters(asset, { q = '', status = 'all', po = '', from = '', to = '' } = {}) {
  const needle = String(q).trim().toLowerCase();
  if (needle) {
    const hit = [asset?.serialNumber || asset?.serial, asset?.name].some((v) => String(v || '').toLowerCase().includes(needle));
    if (!hit) return false;
  }
  if (status !== 'all' && (asset?.status || '') !== status) return false;
  const poNeedle = String(po).trim().toLowerCase();
  if (poNeedle && !String(asset?.poNumber || '').toLowerCase().includes(poNeedle)) return false;
  if (from || to) {
    const key = dateKeyOf(receivedTsOf(asset));
    if (!key) return false; // ไม่มีวันที่ให้อ้าง → ตัดออกเมื่อเปิดตัวกรองวัน (เช่นเดียวกับหน้าทะเบียน)
    if (from && key < from) return false;
    if (to && key > to) return false;
  }
  return true;
}

// ช่วง "N วันล่าสุด" (รวมวันนี้) เป็น { from, to } แบบ key ไทย — ปุ่มด่วน "รับเข้า 2 วันล่าสุด" ใช้ N = 2
export function lastNDaysRange(n, now = new Date()) {
  const dayMs = 24 * 60 * 60 * 1000;
  const todayKey = dateKeyOf(now.getTime());
  if (!todayKey) return { from: '', to: '' };
  const start = new Date(now.getTime() - (Math.max(1, n) - 1) * dayMs);
  return { from: dateKeyOf(start.getTime()), to: todayKey };
}

// จัดกลุ่มตามวันรับเข้า (rows ต้องเรียงล่าสุดอยู่บนแล้ว) → [{ key, label, ts, items }]
// กลุ่ม "ไม่รู้วัน" ท้ายสุดเสมอ เพื่อไม่ตัดหน้ารายการล่าสุด
export function groupRowsByDate(rows) {
  const groups = [];
  const indexOf = new Map();
  (rows || []).forEach((asset) => {
    const ts = receivedTsOf(asset);
    const key = dateKeyOf(ts) || 'unknown';
    if (!indexOf.has(key)) {
      indexOf.set(key, groups.length);
      groups.push({
        key,
        ts: key === 'unknown' ? 0 : ts,
        label: key === 'unknown' ? 'ไม่ระบุวันที่รับเข้า' : thaiDayLabel(ts),
        items: [],
      });
    }
    groups[indexOf.get(key)].items.push(asset);
  });
  return groups;
}

// deep-link เข้าหน้าพิมพ์ฉลากจากหน้าอื่น:
//   ?serial=SN-1            → ['SN-1']           (แบบเดิม จากปุ่มในหน้าทะเบียน)
//   ?serials=SN-1,SN-2      → ['SN-1','SN-2']    (จากปุ่ม "พิมพ์ฉลากทันที" หลังเพิ่มอุปกรณ์ — รองรับ , หรือช่องว่าง)
export function serialsFromSearchParams(searchParams) {
  const multi = String(searchParams?.get('serials') || '').trim();
  if (multi) return [...new Set(multi.split(/[\s,]+/).filter(Boolean))];
  const single = String(searchParams?.get('serial') || '').trim();
  return single ? [single] : [];
}
