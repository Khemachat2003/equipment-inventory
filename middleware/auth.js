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
const ADMIN_SESSION_KEY = 'active_admin_session';

function createCheckAdminSessionLock(cache) {
  return function checkAdminSessionLock(req, res, next) {
    if (!req.session.user || req.session.user.role !== 'admin') {
      return next();
    }
    const currentSessionId = req.session.id;
    const storedSessionId = cache.get(ADMIN_SESSION_KEY);
    if (!storedSessionId) {
      cache.set(ADMIN_SESSION_KEY, currentSessionId, 3600);
      return next();
    }
    if (storedSessionId !== currentSessionId) {
      req.session.destroy((err) => {
        if (err) console.error('Destroy session error:', err);
        return res.status(403).json({ 
          error: 'มี Admin กำลังใช้งานระบบอยู่ กรุณาติดต่อ Admin คนปัจจุบัน' 
        });
      });
      return;
    }
    next();
  };
}

module.exports = {
  requireLogin,
  requireAdmin,
  validate,
  ADMIN_SESSION_KEY,
  createCheckAdminSessionLock
};