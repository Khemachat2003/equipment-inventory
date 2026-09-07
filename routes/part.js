// routes/part.js
const express = require("express");
const router = express.Router();
const { body } = require("express-validator");
const { getSheetsClient, cache, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");

// -------------------- GET PART CATALOG --------------------
router.get("/api/part-catalog", requireLogin, async (req, res) => {
  const cacheKey = "partCatalog";
  let parts = cache.get(cacheKey);
  if (parts) return res.json(parts);

  try {
    const sheets = await getSheetsClient();
    const [catRes, assetRes] = await Promise.all([
      sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Part_Catalog!A2:G",
      }),
      sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A2:M",
      }),
    ]);
    const catRows = catRes.data.values || [];
    const assetRows = assetRes.data.values || [];

    parts = catRows.map((row) => {
      const pn = row[0] || "";
      const count = assetRows.filter((a) => a[3] && a[3].trim() === pn.trim())
        .length;
      return {
        partNumber: pn,
        partName: row[1] || "",
        category: row[2] || "",
        description: row[3] || "",
        unit: row[4] || "",
        totalQty: count,
        lastUpdated: row[6] || "",
      };
    });
    cache.set(cacheKey, parts);
    res.json(parts);
  } catch (e) {
    console.error("Part catalog error:", e);
    res.status(500).json([]);
  }
});

// -------------------- GET CATEGORIES --------------------
const CATEGORY_FALLBACK = [
  ['Sensors & Weather', 'เซนเซอร์ & สภาพอากาศ', 'sensors'],
  ['Controllers & Microcontrollers', 'คอนโทรลเลอร์ & ไมโครคอนโทรลเลอร์', 'memory'],
  ['Networking & Comms', 'เครือข่าย & การสื่อสาร', 'router'],
  ['Power Supplies', 'แหล่งจ่ายไฟ', 'power'],
  ['Power Protection & Switching', 'อุปกรณ์ป้องกัน & สวิตชิ่ง', 'shield'],
  ['Display & Peripherals', 'จอแสดงผล & อุปกรณ์ต่อพ่วง', 'monitor'],
  ['IO & Automation Modules', 'โมดูล IO & ระบบอัตโนมัติ', 'settings_input_component'],
];

router.get("/api/categories", requireLogin, async (req, res) => {
  const cacheKey = "assetCategories";
  let cats = cache.get(cacheKey);
  if (cats) return res.json(cats);
  try {
    const sheets = await getSheetsClient();
    const catRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Categories!A2:C",
    });
    const rows = catRes.data.values || [];
    cats = rows.filter((r) => r[0] && r[0].trim()).map((r) => ({
      name: r[0].trim(),
      label: (r[1] || r[0]).trim(),
      icon: (r[2] || "devices_other").trim(),
    }));
    if (!cats.length) cats = CATEGORY_FALLBACK.map(([name, label, icon]) => ({ name, label, icon }));
    cache.set(cacheKey, cats);
    res.json(cats);
  } catch (e) {
    console.error("Categories error:", e);
    res.json(CATEGORY_FALLBACK.map(([name, label, icon]) => ({ name, label, icon })));
  }
});

// -------------------- ADD PART --------------------
router.post("/api/add-part",
  requireLogin,
  [
    body("partNumber").trim().notEmpty(),
    body("partName").trim().notEmpty(),
    body("category").optional().isString(),
    body("description").optional().isString(),
    body("unit").optional().isString(),
  ],
  validate,
  async (req, res) => {
    try {
      const { partNumber, partName, category, description, unit } = req.body;
      const sheets = await getSheetsClient();
      const date = new Date().toLocaleString("th-TH");
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Part_Catalog!A:G",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            [
              partNumber,
              partName,
              category || "",
              description || "",
              unit || "ชิ้น",
              0,
              date,
            ],
          ],
        },
      });
      cache.del("partCatalog");
      res.json({ success: true });
    } catch (e) {
      console.error("Add part error:", e);
      res.status(500).json({ success: false, error: e.message });
    }
  }
);

module.exports = router;