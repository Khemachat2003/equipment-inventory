// middleware/auth.js
const { validationResult } = require("express-validator");

// Middleware: ตรวจสอบว่าผู้ใช้ล็อกอินแล้ว
function requireLogin(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

// Middleware: ตรวจสอบว่าเป็น Admin
function requireAdmin(req, res, next) {
  if (!req.session.user || req.session.user.role !== "admin") {
    return res.status(403).json({ error: "Unauthorized: Admin only" });
  }
  next();
}

// Middleware: ตรวจสอบ Validation Results (Express-validator)
function validate(req, res, next) {
  const errors = validationResult(req);
  if (!errors.isEmpty()) {
    return res.status(400).json({ errors: errors.array() });
  }
  next();
}

// ── Admin Session Lock: ป้องกัน Admin เข้าพร้อมกัน ──
// เก็บข้อมูล { sessionId, username } ลง cache และ "ต่ออายุ" (refresh TTL) ทุกครั้ง
// ที่ Admin คนปัจจุบันยังใช้งาน API อยู่ — ถ้า session นั้นตาย(ไม่ logout) lock จะ
// ปลดปล่อยไปเองเมื่อพ้นระยะ idle หลังจากที่คนเดิมหยุดใช้งานจริง
const ADMIN_SESSION_KEY = 'active_admin_session';
const ADMIN_LOCK_TTL = parseInt(process.env.ADMIN_LOCK_TTL, 10) || 15 * 60; // 15 นาที

function createCheckAdminSessionLock(cache) {
  return function checkAdminSessionLock(req, res, next) {
    if (!req.session.user || req.session.user.role !== 'admin') {
      return next();
    }
    const currentSessionId = req.session.id;
    const stored = cache.get(ADMIN_SESSION_KEY);

    if (!stored) {
      cache.set(ADMIN_SESSION_KEY, { sessionId: currentSessionId, username: req.session.user.username }, ADMIN_LOCK_TTL);
      return next();
    }

    if (stored.sessionId === currentSessionId) {
      // Admin คนเดิมใช้งานอยู่ → ต่ออายุ lock (ไม่ปล่อยให้ตายกลางคัน)
      cache.set(ADMIN_SESSION_KEY, stored, ADMIN_LOCK_TTL);
      return next();
    }

    // มี Admin คนอื่น (session อื่น) ทำงานอยู่ → บล็อก
    return res.status(403).json({
      error: 'มี Admin กำลังใช้งานระบบอยู่',
      hint: stored.username ? `ตอนนี้ ${stored.username} กำลังใช้งานอยู่ กรุณาติดต่อ Admin คนปัจจุบัน` : undefined,
    });
  };
}

module.exports = {
  requireLogin,
  requireAdmin,
  validate,
  ADMIN_SESSION_KEY,
  createCheckAdminSessionLock
};