// routes/asset.js
const express = require("express");
const router = express.Router();
const QRCode = require("qrcode");
const { body } = require("express-validator");
const { 
  getSheetsClient, 
  saveAssetHistory, 
  clearAssetCache, 
  cache, 
  SPREADSHEET_ID,
  logDamagedAsset,
} = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");
const { logAudit } = require("../services/audit");

// -------------------- SYNC ASSET HISTORY --------------------
let syncDone = false;

async function syncInitialAssetHistory() {
  if (syncDone) return;
  try {
    const sheets = await getSheetsClient();
    const assetResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:M",
    });
    const assetRows = assetResponse.data.values || [];

    const historyResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_History!A2:B",
    });
    const historyRows = historyResponse.data.values || [];
    const loggedSerials = new Set(
      historyRows.map((row) => (row[1] ? row[1].trim() : ""))
    );

    for (const row of assetRows) {
      const serial = row[4] ? row[4].trim() : "";
      if (serial && !loggedSerials.has(serial)) {
        await saveAssetHistory(
          serial,
          "ลงทะเบียนอุปกรณ์ใหม่",
          "-",
          `${row[7] || "-"} (${row[6] || "-"})`,
          row[8] || "System (Auto Sync)",
          `บันทึกประวัติเริ่มต้นจริงเข้าระบบสำหรับอุปกรณ์: ${row[2] || "-"}`
        );
      }
    }
    syncDone = true;
    console.log("✅ Asset history sync completed");
  } catch (err) {
    console.error("❌ Sync asset history error:", err);
  }
}

// -------------------- GET ASSETS --------------------
router.get("/api/assets", requireLogin, async (req, res) => {
  const cacheKey = "assetData";
  let assets = cache.get(cacheKey);
  if (assets) return res.json(assets);

  try {
    await syncInitialAssetHistory();
    const sheets = await getSheetsClient();
    const assetResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:P",
    });
    const assetRows = assetResponse.data.values || [];
    assets = assetRows.map((row) => ({
      assetId: row[0] || "-",
      code: row[1] || "-",
      name: row[2] || "-",
      partNumber: row[3] || "-",
      serialNumber: row[4] || "-",
      status: row[5] || "-",
      location: row[6] || "-",
      siteName: row[7] || "-",
      user: row[8] || "-",
      farmType: row[10] || "-",
      houseId: row[11] || "-",
      houseName: row[12] || "-",
      bundleId: row[13] || "", // ว่าง = ไม่ได้อยู่ใน Bundle ไหน
    }));
    cache.set(cacheKey, assets);
    res.json(assets);
  } catch (error) {
    console.error("❌ Get Assets Error:", error);
    res.status(500).json([]);
  }
});

// -------------------- GET ASSET HISTORY --------------------
router.get("/api/asset-history/:serial", requireLogin, async (req, res) => {
  try {
    const serialNumber = req.params.serial;
    const cacheKey = `assetHistory_${serialNumber}`;
    let cached = cache.get(cacheKey);
    if (cached) return res.json(cached);

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_History!A2:G",
    });
    const rows = response.data.values || [];
    const filteredHistory = rows
      .filter((row) => row[1] && row[1].trim() === serialNumber.trim())
      .map((row) => ({
        date: row[0] || "-",
        serialNumber: row[1] || "-",
        action: row[2] || "-",
        from: row[3] || "-",
        to: row[4] || "-",
        user: row[5] || "-",
        remark: row[6] || "-",
      }));
    const result = filteredHistory.reverse();
    cache.set(cacheKey, result, 60);
    res.json(result);
  } catch (error) {
    console.error("❌ Get Asset History Error:", error);
    res.status(500).json([]);
  }
});

// -------------------- PUBLIC ASSET HISTORY --------------------
router.get("/api/public-asset-history/:serial", async (req, res) => {
  try {
    const serialInput = req.params.serial;
    const cacheKey = `publicAssetHistory_${serialInput}`;
    let cached = cache.get(cacheKey);
    if (cached) return res.json(cached);

    const sheets = await getSheetsClient();

    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:E",
    });
    const assetRows = assetRes.data.values || [];
    let fullSerial = null;

    let found = assetRows.find(r => r[4] && r[4].trim() === serialInput.trim());
    if (!found) {
      found = assetRows.find(r => {
        const s = r[4] ? r[4].trim() : "";
        return s && s.endsWith(serialInput.trim());
      });
    }
    if (found) {
      fullSerial = found[4].trim();
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_History!A2:G",
    });
    const rows = response.data.values || [];

    let filteredHistory = [];
    if (fullSerial) {
      filteredHistory = rows
        .filter((row) => row[1] && row[1].trim() === fullSerial)
        .map((row) => ({
          date: row[0] || "-",
          serialNumber: row[1] || "-",
          action: row[2] || "-",
          from: row[3] || "-",
          to: row[4] || "-",
          user: row[5] || "-",
          remark: row[6] || "-",
        }));
    } else {
      filteredHistory = rows
        .filter((row) => {
          const s = row[1] ? row[1].trim() : "";
          return s && s.endsWith(serialInput.trim());
        })
        .map((row) => ({
          date: row[0] || "-",
          serialNumber: row[1] || "-",
          action: row[2] || "-",
          from: row[3] || "-",
          to: row[4] || "-",
          user: row[5] || "-",
          remark: row[6] || "-",
        }));
    }

    const result = filteredHistory.reverse();
    cache.set(cacheKey, result, 60);
    res.json(result);
  } catch (error) {
    console.error("Public asset history error:", error);
    res.status(500).json([]);
  }
});

// -------------------- PUBLIC ASSET VIEW --------------------
router.get("/api/public/asset/:code", async (req, res) => {
  try {
    const code = req.params.code;
    const cacheKey = `publicAsset_${code}`;
    let cached = cache.get(cacheKey);
    if (cached) return res.json(cached);

    const sheets = await getSheetsClient();
    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:M",
    });
    const rows = assetRes.data.values || [];

    let row = null;
    
    row = rows.find((r) => r[4] && r[4].trim() === code.trim());
    
    if (!row) {
      row = rows.find((r) => {
        const serial = r[4] ? r[4].trim() : "";
        return serial && serial.endsWith(code.trim());
      });
    }

    if (!row) {
      row = rows.find((r) => r[1] && r[1].trim() === code.trim());
    }

    if (!row) {
      return res.status(404).json({ error: "not found" });
    }

    const qrUrl = `${req.protocol}://${req.get("host")}/a/${encodeURIComponent(code)}`;
    const qrImage = await QRCode.toDataURL(qrUrl);

    const result = {
      assetId: row[0] || "-",
      code: row[1] || "-",
      name: row[2] || "-",
      partNumber: row[3] || "-",
      serialNumber: row[4] || "-",
      status: row[5] || "-",
      location: row[6] || "-",
      siteName: row[7] || "-",
      user: row[8] || "-",
      date: row[9] || "-",
      farmType: row[10] || "-",
      houseId: row[11] || "-",
      houseName: row[12] || "-",
      qr: qrImage,
      traceUrl: qrUrl,
    };
    cache.set(cacheKey, result, 60);
    res.json(result);
  } catch (err) {
    console.error("Public asset error:", err);
    res.status(500).json({ error: "server error" });
  }
});

// -------------------- ADD ASSET --------------------
router.post("/api/add-asset",
  requireLogin,
  [
    body("assetId").trim().notEmpty(),
    body("name").trim().notEmpty(),
    body("code").optional().isString(),
    body("partNumber").optional().isString(),
    body("serialNumber").optional().isString(),
    body("status").optional().isString(),
    body("location").optional().isString(),
    body("siteName").optional().isString(),
    body("user").optional().isString(),
  ],
  validate,
  async (req, res) => {
    try {
      const {
        assetId,
        code,
        name,
        partNumber,
        serialNumber,
        status,
        location,
        siteName,
        user,
      } = req.body;
      const sheets = await getSheetsClient();
      const currentDate = new Date().toLocaleString("th-TH");
      const rowValues = [
        assetId,
        code,
        name,
        partNumber,
        serialNumber,
        status,
        location,
        siteName,
        user,
        currentDate,
      ];
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A:J",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [rowValues] },
      });
      await saveAssetHistory(
        serialNumber,
        "ลงทะเบียนอุปกรณ์ใหม่",
        "-",
        `${siteName} (${location})`,
        user,
        `บันทึกประวัติเริ่มต้นจริงเข้าระบบสำหรับอุปกรณ์: ${name}`
      );
      // routes/asset.js - POST /api/add-asset

clearAssetCache();

// ✅ บันทึก Audit Log (ต้องอยู่ก่อน res.json)
await logAudit(
  "เพิ่ม Asset",
  "Asset",
  `Asset ID: ${assetId}, Serial: ${serialNumber}, ชื่อ: ${name}, Part: ${partNumber || "-"}`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });   // ✅ ส่ง Response ทีหลัง
    } catch (error) {
      console.error("❌ Add Asset Error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// -------------------- UPDATE ASSET STATUS --------------------
router.post("/api/update-asset-status",
  requireLogin,
  [
    body("serialNumber").trim().notEmpty(),
    body("action").trim().notEmpty(),
    body("status").trim().notEmpty(),
    body("location").optional().isString(),
    body("siteName").optional().isString(),
    body("user").optional().isString(),
    body("remark").optional().isString(),
    body("fromLocation").optional().isString(),
  ],
  validate,
  async (req, res) => {
    try {
      const {
        serialNumber,
        action,
        status,
        location,
        siteName,
        user,
        remark,
        fromLocation,
      } = req.body;
      const sheets = await getSheetsClient();
      const currentDate = new Date().toLocaleString("th-TH");

      const assetResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A2:E",
      });
      const assetRows = assetResponse.data.values || [];
      const rowIndex =
        assetRows.findIndex(
          (row) => row[4] && row[4].trim() === serialNumber.trim()
        ) + 2;
      if (rowIndex === 1) {
        return res.status(400).json({
          success: false,
          error: "ไม่พบข้อมูล Serial Number นี้ในหน้าหลัก (Asset_List)",
        });
      }

      const updateValues = [status, location, siteName, user, currentDate];
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Asset_List!F${rowIndex}:J${rowIndex}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [updateValues] },
      });

      const historyValues = [
        currentDate,
        serialNumber,
        action,
        fromLocation || "-",
        `${siteName} (${location})`,
        req.session.user ? req.session.user.username : "System",
        remark || "-",
      ];
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_History!A:G",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [historyValues] },
      });
      clearAssetCache();
      cache.del(`assetHistory_${serialNumber}`);
     cache.del(`publicAssetHistory_${serialNumber}`);

// ✅ บันทึก Audit Log
await logAudit(
  "อัปเดตสถานะ Asset",
  "Asset",
  `Serial: ${serialNumber}, สถานะใหม่: ${status}, ตำแหน่ง: ${siteName || "-"} (${location || "-"})`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });
    } catch (error) {
      console.error("❌ Update Asset Status Backend Error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// -------------------- TRANSFER ASSET --------------------
router.post("/api/transfer-asset",
  requireLogin,
  [
    body("serialNumber").trim().notEmpty(),
    body("action").trim().notEmpty(),
    body("status").trim().notEmpty(),
    body("location").optional().isString(),
    body("siteName").optional().isString(),
    body("user").optional().isString(),
    body("remark").optional().isString(),
    body("fromLocation").optional().isString(),
    body("farmType").optional().isString(),
    body("houseId").optional().isString(),
    body("houseName").optional().isString(),
    body("animalType").optional().isString(),
  ],
  validate,
  async (req, res) => {
    try {
      const {
        serialNumber,
        action,
        status,
        location,
        siteName,
        user,
        remark,
        fromLocation,
        farmType,
        houseId,
        houseName,
        animalType,
      } = req.body;

      const sheets = await getSheetsClient();
      const currentDate = new Date().toLocaleString("th-TH");

      const assetRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A2:M",
      });
      const assetRows = assetRes.data.values || [];
      
      const foundIdx = assetRows.findIndex(
        (row) => row[4] && row[4].trim() === serialNumber.trim()
      );
      const rowIndex = foundIdx + 2;

      if (rowIndex === 1) {
        return res.status(400).json({
          success: false,
          error: "ไม่พบ Serial Number นี้ในระบบ",
        });
      }

      const prevRow = assetRows[foundIdx]; // ข้อมูลเดิมก่อนย้าย (A..M)

      // ── ถ้าเป็น "คืนคลังสินค้า" → บังคับกลับคลังกลาง Intranin/Stock เสมอ ──
      const isReturn = action.includes("คืนคลัง");
      if (isReturn) {
        siteName = "Intranin";
        location = "Stock";
      }

      const remarkFull = [
        remark,
        farmType ? `ประเภทฟาร์ม: ${farmType}` : "",
        animalType ? `ประเภทสัตว์: ${animalType}` : "",
        houseName ? `โรงเรือน: ${houseName}` : "",
      ]
        .filter(Boolean)
        .join(" | ");

      const updateValues = [
        status,
        location || "-",
        siteName || "-",
        user || req.session.user.username,
        currentDate,
        farmType || "-",
        houseId || "-",
        houseName || "-",
      ];

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Asset_List!F${rowIndex}:M${rowIndex}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [updateValues] },
      });

      // ── กรณี "คืนคลังสินค้า" + สถานะใหม่ "ชำรุด/สูญหาย" → บันทึกลง sheet อุปกรณ์เสีย ──
      const isDamaged = action.includes("คืนคลัง") && status.includes("ชำรุด/สูญหาย");
      if (isDamaged) {
        await logDamagedAsset({
          date: currentDate,
          serialNumber,
          assetId: prevRow[0] || "",
          code: prevRow[1] || "",
          name: prevRow[2] || "",
          partNumber: prevRow[3] || "",
          status,
          oldLocation: prevRow[6] || "",
          oldSite: prevRow[7] || "",
          user: req.session.user.username,
          remark: remarkFull || "",
          action,
        });
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_History!A:G",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            [
              currentDate,
              serialNumber,
              action,
              fromLocation || "-",
              `${siteName || "-"} (${location || "-"})${houseName ? ` [${houseName}]` : ""}`,
              req.session.user.username,
              remarkFull || "-",
            ],
          ],
        },
      });

      clearAssetCache();
      cache.del(`assetHistory_${serialNumber}`);
      cache.del(`publicAssetHistory_${serialNumber}`);

// ✅ บันทึก Audit Log
await logAudit(
  "โอนย้าย Asset",
  "Asset",
  `Serial: ${serialNumber}, Action: ${action}, จาก: ${fromLocation || "-"}, ไป: ${siteName || "-"} (${location || "-"})`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({ success: true });
    } catch (error) {
      console.error("Transfer asset error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// -------------------- QR CODE --------------------
router.get("/api/qrcode/:serial", requireLogin, async (req, res) => {
  try {
    const serial = req.params.serial;
    const sheets = await getSheetsClient();
    const url = `${req.protocol}://${req.get("host")}/trace.html?serial=${encodeURIComponent(
      serial
    )}`;
    const qrImage = await QRCode.toDataURL(url);

    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:J",
    });
    const rows = assetRes.data.values || [];
    const assetRow = rows.find((r) => r[4] && r[4].trim() === serial.trim());
    const asset = assetRow
      ? {
          assetId: assetRow[0],
          code: assetRow[1],
          name: assetRow[2],
          partNumber: assetRow[3],
          serialNumber: assetRow[4],
          status: assetRow[5],
          location: assetRow[6],
          siteName: assetRow[7],
          user: assetRow[8],
        }
      : null;

    res.json({ success: true, serial, url, qrImage, asset });
  } catch (error) {
    console.error("QR Error:", error);
    res.status(500).json({ success: false });
  }
});

// PATCH /api/assets/:assetId/location
router.patch('/api/assets/:assetId/location', requireLogin, async (req, res) => {
  try {
    const { assetId } = req.params;
    const { location, site, remark } = req.body;
    const sheets = await getSheetsClient();
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID, range: 'Asset_List!A2:M'
    });
    const rows = resp.data.values || [];
    const idx  = rows.findIndex(r => r[0] === assetId);
    if (idx === -1) return res.status(404).json({ error: 'ไม่พบ Asset' });
    const rowNum = idx + 2;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: [
        { range: `Asset_List!G${rowNum}`, values: [[location || '']] },
        { range: `Asset_List!H${rowNum}`, values: [[site || '']] },
      ]}
    });
    clearAssetCache();
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// -------------------- BULK ADD ASSET --------------------
router.post("/api/bulk-add-asset",
  requireLogin,
  [
    body("partNumber").trim().notEmpty(),
    body("partName").trim().notEmpty(),
    body("qty").isInt({ min: 1 }),
    body("status").optional().isString(),
    body("siteName").optional().isString(),
    body("location").optional().isString(),
    body("user").optional().isString(),
    body("farmType").optional().isString(),
    body("houseId").optional().isString(),
    body("houseName").optional().isString(),
  ],
  validate,
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    try {
      const {
        partNumber,
        partName,
        qty,
        status,
        siteName,
        location,
        user,
        farmType,
        houseId,
        houseName,
      } = req.body;
      const sheets = await getSheetsClient();

      const now = new Date();
      const dd = String(now.getDate()).padStart(2, "0");
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      const yyyy = now.getFullYear();
      const dateStr = `${dd}${mm}${yyyy}`;
      const prefix = `SN-${partNumber}-${dateStr}`;

      const assetRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A2:M",
      });
      const existingRows = assetRes.data.values || [];
      let maxNum = 0;
      existingRows.forEach((r) => {
        const s = r[4] || "";
        const parts = s.split("-");
        if (parts.length >= 3) {
          if (parts[parts.length - 2] === dateStr) {
            const num = parseInt(parts[parts.length - 1]);
            if (!isNaN(num) && num > maxNum) maxNum = num;
          }
        }
      });

      const assetIdRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A2:A",
      });
      const existingIds = (assetIdRes.data.values || []).map((r) => parseInt(r[0]) || 0);
      let lastId = existingIds.length > 0 ? Math.max(...existingIds) : 0;

      const currentDate = new Date().toLocaleString("th-TH");
      const newRows = [];
      const historyRows = [];

      for (let i = 0; i < qty; i++) {
        const runNum = String(maxNum + i + 1).padStart(4, "0");
        const serial = `${prefix}-${runNum}`;
        lastId += 1;
        const assetId = String(lastId).padStart(4, "0");

        newRows.push([
          assetId,
          partNumber,
          partName,
          partNumber,
          serial,
          status || "ใช้งานได้",
          location || "-",
          siteName || "-",
          user || req.session.user.username,
          currentDate,
          farmType || "-",
          houseId || "-",
          houseName || "-",
        ]);

        historyRows.push([
          currentDate,
          serial,
          "ลงทะเบียนอุปกรณ์ใหม่",
          "-",
          `${siteName || "-"} (${location || "-"})`,
          req.session.user.username,
          `Bulk Add: ${partName} (${partNumber}) จำนวน ${qty} ชิ้น`,
        ]);
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_List!A:M",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: newRows },
      });

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Asset_History!A:G",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: historyRows },
      });

      const catRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Part_Catalog!A2:G",
      });
      const catRows = catRes.data.values || [];
      const catIdx = catRows.findIndex(
        (r) => r[0] && r[0].trim() === partNumber.trim()
      );
      if (catIdx !== -1) {
        const newTotal = parseInt(catRows[catIdx][5] || 0) + qty;
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `Part_Catalog!F${catIdx + 2}:G${catIdx + 2}`,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[newTotal, currentDate]] },
        });
      }

      clearAssetCache();
      cache.del("partCatalog");

// ✅ บันทึก Audit Log
await logAudit(
  "เพิ่ม Asset หลายชิ้น",
  "Asset",
  `Part: ${partNumber}, ชื่อ: ${partName}, จำนวน: ${qty} ชิ้น, Serial: ${newRows[0][4]} ~ ${newRows[newRows.length-1][4]}`,
  req.session.user.username,
  req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress
);

res.json({
  success: true,
  added: qty,
  serials: newRows.map((r) => r[4]),
  firstSerial: newRows[0][4],
  lastSerial: newRows[newRows.length - 1][4],
});
    } catch (e) {
      console.error("Bulk add error:", e);
      res.status(500).json({ success: false, error: e.message });
    }
  }
);
// ── Attach sync function to router for external call ──
router.syncInitialAssetHistory = syncInitialAssetHistory;
module.exports = router;