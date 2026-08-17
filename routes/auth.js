// routes/auth.js
const express = require("express");
const router = express.Router();
const bcrypt = require("bcrypt");
const { body } = require("express-validator");
const { getSheetsClient, cache, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");

const saltRounds = 10;

// -------------------- LOGIN API --------------------
router.post("/api/login",
  [
    body("username").trim().notEmpty().withMessage("Username required"),
    body("password").trim().notEmpty().withMessage("Password required"),
  ],
  validate,
  async (req, res) => {
    try {
      const { username, password } = req.body;
      const sheets = await getSheetsClient();
      const userRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users!A2:C",
      });
      const users = userRes.data.values || [];
      
      const userRow = users.find((u) => u[0]?.trim() === username.trim());
      if (!userRow) {
        return res.status(401).json({ error: "Username หรือ Password ไม่ถูกต้อง" });
      }

      const storedPassword = userRow[1]?.trim() || "";
      const role = userRow[2] || "user";
      let isMatch = false;
      let needsUpdate = false;

      if (!storedPassword.startsWith('$2b$')) {
        if (storedPassword === password.trim()) {
          isMatch = true;
          needsUpdate = true;
        }
      } else {
        isMatch = await bcrypt.compare(password.trim(), storedPassword);
      }

      if (!isMatch) {
        return res.status(401).json({ error: "Username หรือ Password ไม่ถูกต้อง" });
      }

      if (needsUpdate) {
        const hashed = await bcrypt.hash(password.trim(), saltRounds);
        const rowIndex = users.findIndex((u) => u[0]?.trim() === username.trim()) + 2;
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `Users!B${rowIndex}`,
          valueInputOption: "RAW",
          requestBody: { values: [[hashed]] },
        });
        console.log(`✅ Password for ${username} migrated to bcrypt`);
      }

      if (role === 'admin') {
        const existing = cache.get('active_admin_session');
        if (existing && existing.sessionId !== req.session.id) {
          return res.status(401).json({ 
            error: "⚠️ มี Admin กำลังใช้งานระบบอยู่",
            hint: existing.username ? `${existing.username} กำลังใช้งานอยู่ กรุณาติดต่อ Admin คนปัจจุบัน` : undefined,
          });
        }
        cache.set('active_admin_session', { sessionId: req.session.id, username }, 15 * 60);
      }

      req.session.user = { username, role };
      req.session.save(() => res.json({ success: true, role }));

    } catch (err) {
      console.error("LOGIN ERROR:", err);
      res.status(500).json({ error: "Server error", message: err.message });
    }
  }
);

// -------------------- CHECK AUTH --------------------
router.get("/api/check-auth", (req, res) => {
  if (!req.session.user) return res.json({ loggedIn: false });
  res.json({
    loggedIn: true,
    username: req.session.user.username,
    role: req.session.user.role,
  });
});

// -------------------- LOGOUT --------------------
router.post("/api/logout", (req, res) => {
  if (req.session.user && req.session.user.role === 'admin') {
    cache.del('active_admin_session');
  }
  req.session.destroy(() => res.json({ success: true }));
});

// -------------------- GET CURRENT USER --------------------
router.get("/api/me", requireLogin, (req, res) => {
  res.json({ username: req.session.user.username });
});

// -------------------- GET SPREADSHEET URL --------------------
router.get("/api/settings/spreadsheet-url", requireLogin, (req, res) => {
  res.json({ url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}` });
});

module.exports = router;