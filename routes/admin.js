// routes/admin.js
const express = require("express");
const router = express.Router();
const bcrypt = require("bcrypt");
const { body } = require("express-validator");
const { getSheetsClient, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, requireAdmin, validate } = require("../middleware/auth");

const saltRounds = 10;

// -------------------- UPDATE USER ROLE --------------------
router.put("/api/admin/users/:username",
  requireLogin,
  requireAdmin,
  [
    body("role").isIn(["admin", "user"]),
  ],
  validate,
  async (req, res) => {
    try {
      const { username } = req.params;
      const { role } = req.body;
      const sheets = await getSheetsClient();

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users!A2:C",
      });
      const rows = response.data.values || [];
      const index = rows.findIndex((r) => r[0]?.trim() === username.trim());
      if (index === -1) {
        return res.status(404).json({ error: "User not found" });
      }

      const rowIndex = index + 2;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!C${rowIndex}`,
        valueInputOption: "RAW",
        requestBody: { values: [[role]] },
      });

      res.json({ success: true, message: `Role updated for ${username}` });
    } catch (error) {
      console.error("Update role error:", error);
      res.status(500).json({ error: "Failed to update role" });
    }
  }
);

// -------------------- RESET PASSWORD --------------------
router.post("/api/admin/users/:username/reset-password",
  requireLogin,
  requireAdmin,
  [
    body("newPassword").trim().notEmpty().isLength({ min: 4 }),
  ],
  validate,
  async (req, res) => {
    try {
      const { username } = req.params;
      const { newPassword } = req.body;
      const sheets = await getSheetsClient();

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users!A2:C",
      });
      const rows = response.data.values || [];
      const index = rows.findIndex((r) => r[0]?.trim() === username.trim());
      if (index === -1) {
        return res.status(404).json({ error: "User not found" });
      }

      const hashed = await bcrypt.hash(newPassword.trim(), saltRounds);
      const rowIndex = index + 2;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!B${rowIndex}`,
        valueInputOption: "RAW",
        requestBody: { values: [[hashed]] },
      });

      res.json({ success: true, message: `Password reset for ${username}` });
    } catch (error) {
      console.error("Reset password error:", error);
      res.status(500).json({ error: "Failed to reset password" });
    }
  }
);

// -------------------- DELETE USER --------------------
router.delete("/api/admin/users/:username",
  requireLogin,
  requireAdmin,
  async (req, res) => {
    try {
      const { username } = req.params;
      const sheets = await getSheetsClient();

      if (username === req.session.user.username) {
        return res.status(400).json({ error: "Cannot delete yourself" });
      }

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users!A2:C",
      });
      const rows = response.data.values || [];
      const index = rows.findIndex((r) => r[0]?.trim() === username.trim());
      if (index === -1) {
        return res.status(404).json({ error: "User not found" });
      }

      const rowIndex = index + 2;
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `Users!A${rowIndex}:C${rowIndex}`,
      });

      res.json({ success: true, message: `User ${username} deleted` });
    } catch (error) {
      console.error("Delete user error:", error);
      res.status(500).json({ error: "Failed to delete user" });
    }
  }
);

module.exports = router;