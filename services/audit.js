// services/audit.js
const { getSheetsClient, SPREADSHEET_ID } = require("./sheets");

async function logAudit(action, module, detail, username, ip = null) {
  try {
    const sheets = await getSheetsClient();
    const now = new Date().toLocaleString("th-TH");
    const row = [
      now,
      username || "System",
      action,
      module,
      detail || "-",
      ip || "-",
    ];
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Audit_Log!A:F",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [row] },
    });
  } catch (err) {
    console.error("❌ Audit log error:", err);
  }
}

module.exports = { logAudit };