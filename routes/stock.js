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
        ext: row[8] || "",
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

// -------------------- BORROW (BULK) SELECTED OFFICE ITEMS --------------------
// เบิกหลายรายการพร้อมกัน (แบบตะกร้า) — ย้าย Office → Site ทีละรายการ
// หลัก: เบิกได้เท่าที่ Office มีจริง, ของไม่พอ → เบิกเต็มเท่าที่มี + แจ้งยอดจริง
router.post("/api/borrow-selected",
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
      const borrowedItems = [];   // เบิกสำเร็จ
      const shortItems = [];      // เบิกได้ไม่เต็ม (ของใน Office ไม่พอ)

      for (const item of items) {
        const code = item.code;
        const requested = parseInt(item.qty);
        if (!code || !requested || requested < 1) continue;
        const officeIndex = officeData.findIndex((r) => r[0] === code);
        const siteIndex = siteData.findIndex((r) => r[0] === code);
        if (officeIndex === -1 || siteIndex === -1) continue;
        let officeQty = parseInt(officeData[officeIndex][2] || 0);
        let siteQty = parseInt(siteData[siteIndex][2] || 0);
        if (officeQty < 1) continue; // ไม่มีของเลย → ข้าม
        const actual = Math.min(requested, officeQty); // เบิกได้เท่าที่มี
        officeQty -= actual;
        siteQty += actual;
        updates.push({
          range: `Stock_Office!C${officeIndex + 2}`,
          values: [[officeQty]],
        });
        updates.push({
          range: `Stock_Site!C${siteIndex + 2}`,
          values: [[siteQty]],
        });
        const name = officeData[officeIndex][1];
        logs.push([
          new Date().toLocaleString("th-TH"),
          code,
          name,
          actual,
          "เบิก",
          "Office",
          "Site",
          user,
        ]);
        borrowedItems.push(`${code} (${name}) x${actual}`);
        if (actual < requested) {
          shortItems.push({ code, name, requested, actual });
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

      // ✅ บันทึก Audit Log
      await logAudit(
        "เบิกอุปกรณ์ (หลายรายการ)",
        "Stock",
        `สำเร็จ ${borrowedItems.length} รายการ${shortItems.length ? `, ของไม่พอ ${shortItems.length} รายการ` : ""}: ${borrowedItems.join(", ")}`,
        user,
        req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
      );

      res.json({ success: true, borrowed: borrowedItems, short: shortItems });
    } catch (err) {
      console.error("borrow-selected error:", err);
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
    body("office").optional().isInt({ min: 0 }),
    body("site").optional().isInt({ min: 0 }),
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

      // 🔒 กันรหัสซ้ำ — รหัสเดียวกันต้องไม่เพิ่มซ้ำ (กันรูป/ข้อมูลทับกัน)
      const masterRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Master!A2:A",
      });
      const exists = (masterRes.data.values || []).some((r) => String(r[0]).trim() === code);
      if (exists) {
        return res.status(400).json({ success: false, error: `รหัส ${code} มีอยู่แล้วใน Stock แล้ว` });
      }

      // จำนวนใน Office/Site:
      //  - ถ้ากรอก Office และ/หรือ Site → ใช้ตามนั้น, Total = Office + Site
      //  - ถ้าไม่กรอก (เพิ่มแบบเดิม) → ของใหม่ทั้งหมดเริ่มที่ Office, Site = 0
      let officeQty = req.body.office != null ? parseInt(req.body.office) : null;
      let siteQty = req.body.site != null ? parseInt(req.body.site) : null;
      if (officeQty == null && siteQty == null) {
        officeQty = parseInt(total);
        siteQty = 0;
      } else {
        officeQty = officeQty || 0;
        siteQty = siteQty || 0;
      }
      const effectiveTotal = officeQty + siteQty;

      const imageUrl = `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image/images/${code}.${ext}`;

      // 1) Stock_Master (A:I) — เก็บ ext ไว้ที่คอลัมน์ I ให้หน้าเว็บรู้จักไฟล์รูป
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Master!A:I",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[code, name, `=IMAGE("${imageUrl}")`, "", "", effectiveTotal, "", "", ext]],
        },
      });

      // 2) Stock_Office (A:C) — เพิ่มรายการใหม่ทันที พร้อมจำนวนที่กรอก
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Office!A:C",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[code, name, officeQty]] },
      });

      // 3) Stock_Site (A:C) — เพิ่มรายการใหม่ทันที พร้อมจำนวนที่กรอก
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Stock_Site!A:C",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[code, name, siteQty]] },
      });

      clearStockCache();

// ✅ บันทึก Audit Log
await logAudit(
  "เพิ่มอุปกรณ์ Stock",
  "Stock",
  `รหัส: ${code}, ชื่อ: ${name}, Office: ${officeQty}, Site: ${siteQty}, รวม: ${effectiveTotal}`,
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
    if (!process.env.GITHUB_TOKEN) {
      console.error("upload-image: GITHUB_TOKEN is not set");
      return res.status(500).json({ success: false, error: "ยังไม่ได้ตั้งค่า GITHUB_TOKEN ฝั่งเซิร์ฟเวอร์" });
    }
    const { Octokit } = require("@octokit/rest");
    const octokit = new Octokit({ auth: process.env.GITHUB_TOKEN });
    const owner = process.env.GITHUB_USERNAME || "Khemachat2003";
    const repo = process.env.GITHUB_REPO || "stock-image";
    const { fileName, base64 } = req.body;
    const content = base64.replace(/^data:image\/\w+;base64,/, "");
    const filePath = `images/${fileName}`;
    try {
      let sha = null;
      // 🔒 กันชื่อไฟล์รูปซ้ำ — ถ้ามีไฟล์ชื่อนี้ใน repo แล้ว (และไม่ส่ง overwrite) → ปฏิเสธ
      // (ถ้าต้องการอัปเดตรูปเดิม ให้ส่ง overwrite:true)
      let exists = false;
      try {
        const existingFile = await octokit.repos.getContent({ owner, repo, path: filePath });
        exists = true;
        sha = existingFile.data.sha;
      } catch (err) {}
      if (exists && !req.body.overwrite) {
        return res.status(400).json({
          success: false,
          error: `ชื่อไฟล์รูปนี้มีอยู่แล้ว (${fileName}) กรุณาเปลี่ยนรหัสหรือเปิดตัวเลือกอัปเดตรูปเดิม`,
        });
      }

      const params = {
        owner,
        repo,
        path: filePath,
        message: `upload image ${fileName}`,
        content: content,
      };
      // ส่ง sha เฉพาะตอนอัปเดตไฟล์เดิม — สร้างไฟล์ใหม่ต้องไม่ส่ง sha (null ทำให้ GitHub ตอบ 422)
      if (sha) params.sha = sha;

      await octokit.repos.createOrUpdateFileContents(params);
      res.json({ success: true });
    } catch (error) {
      console.error("GitHub upload error:", error.status || "", error.message);
      const status = error.status === 401 ? 401 : 500;
      const detail =
        error.status === 401
          ? "GitHub Token หมดอายุหรือไม่ถูกต้อง กรุณาตั้งค่า GITHUB_TOKEN ใหม่"
          : error.message;
      res.status(status).json({ success: false, error: detail });
    }
  }
);

module.exports = router;