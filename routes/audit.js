// routes/audit.js
const express = require("express");
const router = express.Router();
const { getSheetsClient, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, requireAdmin } = require("../middleware/auth");

// GET /api/audit-log - ดึงข้อมูล Audit Log (เฉพาะ Admin)
router.get("/api/audit-log", requireLogin, requireAdmin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Audit_Log!A:F",
    });
    const rows = response.data.values || [];
    // แปลงข้อมูลให้เป็น Object
    const logs = rows.slice(1).map(row => ({
      timestamp: row[0] || "-",
      user: row[1] || "-",
      action: row[2] || "-",
      module: row[3] || "-",
      detail: row[4] || "-",
      ip: row[5] || "-",
    }));
    // เรียงลำดับจากล่าสุดไปเก่าสุด
    res.json(logs.reverse());
  } catch (err) {
    console.error("❌ Audit log error:", err);
    res.status(500).json({ error: "Failed to load audit log" });
  }
});

module.exports = router;