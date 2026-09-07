// middleware/portal.js
// ── PORTAL MODE ──────────────────────────────────────────────────────────────
// เมื่อเปิด PORTAL_MODE=true ระบบจะถือว่า "ผู้ใช้งานผ่านระบบ portal รวม" ของบริษัท
// ผู้ที่เข้ามาโดยยังไม่มี session จะถูก auto-login เป็น regular user (role:"user")
// ทันที — ไม่ต้องผ่านหน้า login ของเรา (login ผ่าน portal ของบริษัทแล้ว)
//
// หลักความปลอดภัย:
//  - ปลด lock เฉพาะเมื่อเปิด PORTAL_MODE โดยชัดเจน (default ปิด → พฤติกรรมเดิม)
//  - portal user จะเป็น role:"user" เสมอ (อ่าน role จากภายนอกไม่ได้ → ป้องกัน
//    คนปั้น URL หลอกเป็น admin)
//  - ถ้ามี session อยู่แล้ว (admin/user login จริง) ไม่แตะ session เดิม
//  - admin ยังต้อง login ผ่าน /api/login ปกติ (ตรวจรหัสผ่าน) + admin lock เดิม
//  - endpoint ที่ requireAdmin (backup, audit, clear-cache) ยังปกป้องอยู่
//
// วิธีกำหนด user:
//   PORTAL_USERNAME_QUERY  (env, default "username") → ชื่อ query param รับ username
//   PORTAL_USERNAME        (env, default "portal")   → ใช้ username นี้เมื่อไม่ได้ระบุ
//   request ?username=<name> (ทาง portal ส่งมา)
// ──────────────────────────────────────────────────────────────────────────────

const DEFAULT_USERNAME = process.env.PORTAL_USERNAME || "portal";
const USERNAME_QUERY_KEY = process.env.PORTAL_USERNAME_QUERY || "username";

function portalLoginEnabled() {
  const v = String(process.env.PORTAL_MODE || "").trim().toLowerCase();
  return v === "true" || v === "1" || v === "on" || v === "yes";
}

// Middleware: auto-login portal users (ไม่มี session + PORTAL_MODE on)
function portalAutoLogin(req, res, next) {
  if (!portalLoginEnabled()) return next();

  // มี session แล้ว → ไม่แตะ (admin/user ที่ login จริง ยังปกติ)
  if (req.session.user) return next();

  // ให้ /api/login ผ่านไปได้เสมอ เพื่อให้ admin/login จริงยังทำงาน
  if (req.path === "/api/login") return next();

  // อ่าน username จากที่ portal ส่งมา (บังคับ role:"user" เสมอ)
  let username = req.query[USERNAME_QUERY_KEY];
  if (typeof username === "string") username = username.trim().slice(0, 64);
  if (!username) username = DEFAULT_USERNAME;

  req.session.user = { username, role: "user", portal: true };
  req.session.save && req.session.save((err) => {
    if (err) console.error("PORTAL session save error:", err);
    next();
  });
  if (!req.session.save) next();
}

module.exports = { portalAutoLogin, portalLoginEnabled };