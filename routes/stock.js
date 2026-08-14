// routes/stock.js
const express = require("express");
const router = express.Router();
const { body } = require("express-validator");
const { getSheetsClient, clearStockCache, cache, SPREADSHEET_ID } = require("../services/sheets");
const { logAudit } = require("../services/audit");
const { requireLogin, requireAdmin, validate } = require("../middleware/auth");

// -------------------- GET STOCK --------------------
router.get("/api/stock", requireLogin, async (req, res) => {
  const cacheKey = "stockData";
  let stock = cache.get(cacheKey);
  if (stock) return res.json(stock);

  try {
    const sheets = await getSheetsClient();
    const master = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A2:I",
    });
    const office = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Office!A2:C",
    });
    const site = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C",
    });
    const masterData = master.data.values || [];
    const officeData = office.data.values || [];
    const siteData = site.data.values || [];

    const result = masterData.map((row) => {
      const code = row[0];
      const officeRow = officeData.find((r) => r[0] === code);
      const siteRow = siteData.find((r) => r[0] === code);
      return {
        code,
        name: row[1],
        total: parseInt(row[5] || 0),
        office: officeRow ? parseInt(officeRow[2] || 0) : 0,
        site: siteRow ? parseInt(siteRow[2] || 0) : 0,
      };
    });
    cache.set(cacheKey, result);
    res.json(result);
  } catch (err) {
    console.error("Stock error:", err);
    res.status(500).json({ error: "Stock error" });
  }
});

// -------------------- UPDATE TOTAL --------------------
router.post("/api/update-total",
  requireLogin,
  [
    body("code").trim().notEmpty(),
    body("newTotal").isInt({ min: 0 }),
  ],
  validate,
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    try {
      const { code, newTotal } = req.body;
      const sheets = await getSheetsClient();
      const masterRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Master!A2:I",
      });
      const masterData = masterRes.data.values || [];
      const index = masterData.findIndex((r) => r[0] === code);
      if (index === -1) return res.json({ error: "ไม่พบสินค้า" });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Stock_Master!F${index + 2}`,
        valueInputOption: "RAW",
        requestBody: { values: [[parseInt(newTotal)]] },
      });
    clearStockCache();

// ✅ บันทึก Audit Log (ใช้ข้อมูลที่มี)
const currentUser = req.session.user.username;
await logAudit(
  "แก้ไขจำนวน Stock",
  "Stock",
  `รหัส: ${code}, จำนวนใหม่: ${newTotal}`,
  currentUser,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });
    } catch (err) {
      console.error("Update total error:", err);
      res.status(500).json({ error: "Update error" });
    }
  }
);

// -------------------- ADD STOCK --------------------
router.post("/api/add-stock",
  requireLogin,
  [
    body("code").trim().notEmpty(),
    body("qty").isInt({ min: 1 }),
  ],
  validate,
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    try {
      const { code, qty } = req.body;
      const addQty = parseInt(qty);
      const sheets = await getSheetsClient();
      const masterRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Master!A2:I",
      });
      const masterData = masterRes.data.values || [];
      const index = masterData.findIndex((r) => r[0] === code);
      if (index === -1) return res.json({ error: "ไม่พบสินค้า" });
      let currentTotal = parseInt(masterData[index][5] || 0);
      let newTotal = currentTotal + addQty;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Stock_Master!F${index + 2}`,
        valueInputOption: "RAW",
        requestBody: { values: [[newTotal]] },
      });
      clearStockCache();

// ✅ บันทึก Audit Log
await logAudit(
  "เพิ่ม Stock",
  "Stock",
  `รหัส: ${code}, เพิ่มจำนวน: ${addQty}, จำนวนใหม่: ${newTotal}`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });
    } catch (err) {
      console.error("Add stock error:", err);
      res.status(500).json({ error: "Add stock error" });
    }
  }
);

// -------------------- TRANSFER --------------------
router.post("/api/transfer",
  requireLogin,
  [
    body("code").trim().notEmpty(),
    body("name").trim().notEmpty(),
    body("qty").isInt({ min: 1 }),
    body("type").isIn(["เบิก", "คืน"]),
  ],
  validate,
  async (req, res) => {
    try {
      const { code, name, type } = req.body;
      const qty = parseInt(req.body.qty);
      const user = req.session.user.username;
      const sheets = await getSheetsClient();

      const officeRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Office!A2:C",
      });
      const siteRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Site!A2:C",
      });
      const officeData = officeRes.data.values || [];
      const siteData = siteRes.data.values || [];

      let officeIndex = officeData.findIndex((r) => r[0] === code);
      let siteIndex = siteData.findIndex((r) => r[0] === code);
      if (officeIndex === -1 || siteIndex === -1) {
        return res.json({ error: "ไม่พบข้อมูลสินค้าใน stock" });
      }

      let officeQty = parseInt(officeData[officeIndex][2] || 0);
      let siteQty = parseInt(siteData[siteIndex][2] || 0);

      if (type === "เบิก") {
        if (officeQty < qty) return res.json({ error: "Office ไม่พอ" });
        officeQty -= qty;
        siteQty += qty;
      } else if (type === "คืน") {
        if (siteQty < qty) return res.json({ error: "Site ไม่พอ" });
        siteQty -= qty;
        officeQty += qty;
      }

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Stock_Office!C${officeIndex + 2}`,
        valueInputOption: "RAW",
        requestBody: { values: [[officeQty]] },
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Stock_Site!C${siteIndex + 2}`,
        valueInputOption: "RAW",
        requestBody: { values: [[siteQty]] },
      });

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Transfer_Log!A:H",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            [
              new Date().toLocaleString("th-TH"),
              code,
              name,
              qty,
              type,
              type === "เบิก" ? "Office" : "Site",
              type === "เบิก" ? "Site" : "Office",
              user,
            ],
          ],
        },
      });
      clearStockCache();

      // ✅ บันทึก Audit Log
      await logAudit(
        type === "เบิก" ? "เบิกอุปกรณ์" : "คืนอุปกรณ์",
        "Stock",
        `รหัส: ${code}, ชื่อ: ${name}, จำนวน: ${qty}`,
        user,
        req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
      );

      res.json({ success: true });
    } catch (err) {
      console.error("Transfer error:", err);
      res.status(500).json({ error: "Transfer error" });
    }
  }
);

// -------------------- GET SITE ITEMS --------------------
router.get("/api/get-site-items", requireLogin, async (req, res) => {
  const cacheKey = "siteItems";
  let cached = cache.get(cacheKey);
  if (cached) return res.json(cached);

  try {
    const sheets = await getSheetsClient();
    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C",
    });
    const rows = siteRes.data.values || [];
    const items = [];
    rows.forEach((r) => {
      const code = r[0];
      const name = r[1];
      const qty = parseInt(r[2] || 0);
      if (qty > 0) items.push({ code, name, qty });
    });
    const result = { items };
    cache.set(cacheKey, result, 60);
    res.json(result);
  } catch (err) {
    console.error("get-site-items error:", err);
    res.status(500).json({ items: [] });
  }
});

// -------------------- RETURN ALL SITE --------------------
router.post("/api/return-all-site", requireLogin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const user = req.session.user.username;
    const officeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Office!A2:C",
    });
    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C",
    });
    const officeData = officeRes.data.values || [];
    const siteData = siteRes.data.values || [];

    const updates = [];
    const logs = [];
    const returnedItems = []; // เก็บข้อมูลเพื่อบันทึก Audit

    for (let i = 0; i < siteData.length; i++) {
      const code = siteData[i][0];
      const name = siteData[i][1];
      let siteQty = parseInt(siteData[i][2] || 0);
      if (siteQty > 0) {
        const officeIndex = officeData.findIndex((r) => r[0] === code);
        if (officeIndex === -1) continue;
        let officeQty = parseInt(officeData[officeIndex][2] || 0);
        officeQty += siteQty;
        updates.push({
          range: `Stock_Office!C${officeIndex + 2}`,
          values: [[officeQty]],
        });
        updates.push({
          range: `Stock_Site!C${i + 2}`,
          values: [[0]],
        });
        logs.push([
          new Date().toLocaleString("th-TH"),
          code,
          name,
          siteQty,
          "คืน",
          "Site",
          "Office",
          user,
        ]);
        // เก็บข้อมูลสำหรับ Audit
        returnedItems.push(`${code} (${name}) x${siteQty}`);
      }
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "RAW", data: updates },
      });
    }
    if (logs.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Transfer_Log!A:H",
        valueInputOption: "RAW",
        requestBody: { values: logs },
      });
    }

    clearStockCache();
    cache.del("siteItems");

    // ✅ บันทึก Audit Log
    if (returnedItems.length > 0) {
      await logAudit(
        "คืนอุปกรณ์ทั้งหมดจาก Site",
        "Stock",
        `จำนวน ${returnedItems.length} รายการ: ${returnedItems.join(", ")}`,
        user,
        req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
      );
    } else {
      await logAudit(
        "คืนอุปกรณ์ทั้งหมดจาก Site (ไม่มีรายการ)",
        "Stock",
        "ไม่พบอุปกรณ์ใน Site ที่ต้องคืน",
        user,
        req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
      );
    }

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false });
  }
});

// -------------------- RETURN SELECTED SITE --------------------
router.post("/api/return-selected-site",
  requireLogin,
  [body("items").isArray({ min: 1 })],
  validate,
  async (req, res) => {
    try {
      const sheets = await getSheetsClient();
      const items = req.body.items;
      const user = req.session.user.username;
      const officeRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Office!A2:C",
      });
      const siteRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Site!A2:C",
      });
      const officeData = officeRes.data.values || [];
      const siteData = siteRes.data.values || [];

      const updates = [];
      const logs = [];
      const returnedItems = []; // เก็บข้อมูลสำหรับ Audit

      for (const item of items) {
        const code = item.code;
        const qty = parseInt(item.qty);
        const officeIndex = officeData.findIndex((r) => r[0] === code);
        const siteIndex = siteData.findIndex((r) => r[0] === code);
        if (officeIndex === -1 || siteIndex === -1) continue;
        let officeQty = parseInt(officeData[officeIndex][2] || 0);
        let siteQty = parseInt(siteData[siteIndex][2] || 0);
        if (siteQty < qty) continue;
        officeQty += qty;
        siteQty -= qty;
        updates.push({
          range: `Stock_Office!C${officeIndex + 2}`,
          values: [[officeQty]],
        });
        updates.push({
          range: `Stock_Site!C${siteIndex + 2}`,
          values: [[siteQty]],
        });
        logs.push([
          new Date().toLocaleString("th-TH"),
          code,
          officeData[officeIndex][1],
          qty,
          "คืน",
          "Site",
          "Office",
          user,
        ]);
        returnedItems.push(`${code} (${officeData[officeIndex][1]}) x${qty}`);
      }

      if (updates.length > 0) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { valueInputOption: "RAW", data: updates },
        });
      }
      if (logs.length > 0) {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: "Transfer_Log!A:H",
          valueInputOption: "RAW",
          requestBody: { values: logs },
        });
      }

      clearStockCache();
      cache.del("siteItems");

      // ✅ บันทึก Audit Log
      if (returnedItems.length > 0) {
        await logAudit(
          "คืนอุปกรณ์ที่เลือกจาก Site",
          "Stock",
          `จำนวน ${returnedItems.length} รายการ: ${returnedItems.join(", ")}`,
          user,
          req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
        );
      } else {
        await logAudit(
          "คืนอุปกรณ์ที่เลือกจาก Site (ไม่มีรายการ)",
          "Stock",
          "ไม่มีรายการที่สามารถคืนได้",
          user,
          req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
        );
      }

      res.json({ success: true });
    } catch (err) {
      console.error(err);
      res.status(500).json({ success: false });
    }
  }
);

// -------------------- ADD ITEM --------------------
router.post("/api/add-item",
  requireLogin,
  [
    body("code").trim().notEmpty(),
    body("name").trim().notEmpty(),
    body("total").isInt({ min: 0 }),
    body("ext").optional().isString(),
  ],
  validate,
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    const { code, name, total, ext } = req.body;
    try {
      const sheets = await getSheetsClient();
      const imageUrl = `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image/images/${code}.${ext}`;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Master!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[code, name, `=IMAGE("${imageUrl}")`, "", "", total]],
        },
      });
      clearStockCache();

// ✅ บันทึก Audit Log
await logAudit(
  "เพิ่มอุปกรณ์ Stock",
  "Stock",
  `รหัส: ${code}, ชื่อ: ${name}, จำนวน: ${total}`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });
    } catch (error) {
      console.error("Add item error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// -------------------- UPLOAD IMAGE --------------------
router.post("/upload-image",
  requireLogin,
  requireAdmin,
  [
    body("fileName").trim().notEmpty(),
    body("base64").notEmpty().custom((val) => val.startsWith("data:image/")),
  ],
  validate,
  async (req, res) => {
    const { Octokit } = require("@octokit/rest");
    const octokit = new Octokit({ auth: process.env.GITHUB_TOKEN });
    const { fileName, base64 } = req.body;
    const content = base64.replace(/^data:image\/\w+;base64,/, "");
    const filePath = `images/${fileName}`;
    try {
      let sha = null;
      try {
        const existingFile = await octokit.repos.getContent({
          owner: "Khemachat2003",
          repo: "stock-image",
          path: filePath,
        });
        sha = existingFile.data.sha;
      } catch (err) {}
      await octokit.repos.createOrUpdateFileContents({
        owner: "Khemachat2003",
        repo: "stock-image",
        path: filePath,
        message: "upload image",
        content: content,
        sha: sha,
      });
      res.json({ success: true });
    } catch (error) {
      console.log("GitHub error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

module.exports = router;