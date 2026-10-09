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
const { normalizeReceivedAt, resolveInboundBatch, createInboundRows, inboundMap, summarizeInboundBatches, summarizeInboundPOs } = require("../utils/inbound");
const { readInboundRows, prepareInboundWrite, appendInboundRows } = require("../services/inboundStore");
const { readPOStatuses, savePOStatus } = require("../services/inboundPOStore");
const { isLocalInventoryMode, readLocalInventory, addLocalAsset, addLocalAssets } = require("../services/localInventory");

// -------------------- SYNC ASSET HISTORY --------------------
// แบบ batch: สร้างแถวที่ต้อง sync ทั้งหมดก่อนแล้ว append เป็นก้อน
// (เดิม append ทีละแถว → N asset = N API calls เสีย quota + บล็อก request แรกของ /api/assets)
const SHEETS_APPEND_BATCH = parseInt(process.env.SHEETS_APPEND_BATCH, 10) || 500;
let syncDone = false;
let syncInFlight = null; // กัน race: หลาย request เรียกพร้อมกัน → sync ครั้งเดียว

async function syncInitialAssetHistory() {
  if (syncDone) return;
  if (syncInFlight) return syncInFlight;
  syncInFlight = (async () => {
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

      // เก็บแถวใหม่ลง array ก่อน (เงื่อนไข/รูปแบบข้อมูลเหมือนเดิมทุกอย่าง แค่ยังไม่เขียน)
      const newRows = [];
      for (const row of assetRows) {
        const serial = row[4] ? row[4].trim() : "";
        if (serial && !loggedSerials.has(serial)) {
          newRows.push([
            new Date().toLocaleString("th-TH"),
            serial,
            "ลงทะเบียนอุปกรณ์ใหม่",
            "-",
            `${row[7] || "-"} (${row[6] || "-"})`,
            row[8] || "System (Auto Sync)",
            `บันทึกประวัติเริ่มต้นจริงเข้าระบบสำหรับอุปกรณ์: ${row[2] || "-"}`,
          ]);
        }
      }

      // append ทีละก้อน (ค่าเริ่มต้น 500 แถว/call) → API calls จาก N ลดเหลือ ceil(N/500)
      for (let i = 0; i < newRows.length; i += SHEETS_APPEND_BATCH) {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: "Asset_History!A:G",
          valueInputOption: "USER_ENTERED",
          insertDataOption: "INSERT_ROWS",
          requestBody: { values: newRows.slice(i, i + SHEETS_APPEND_BATCH) },
        });
      }

      syncDone = true;
      console.log(
        newRows.length
          ? `✅ Asset history sync completed (${newRows.length} rows, batch size ${SHEETS_APPEND_BATCH})`
          : "✅ Asset history sync completed (no new rows)"
      );
    } catch (err) {
      console.error("❌ Sync asset history error:", err);
      // syncDone ยังเป็น false → ครั้งถัดไปจะ retry (พฤติกรรมเดิม)
    } finally {
      syncInFlight = null;
    }
  })();
  return syncInFlight;
}

// -------------------- GET ASSETS --------------------
router.get("/api/assets", requireLogin, async (req, res) => {
  if (isLocalInventoryMode()) {
    try {
      const [inventory, inboundRows] = await Promise.all([readLocalInventory(), readInboundRows(null, SPREADSHEET_ID)]);
      const bySerial = inboundMap(inboundRows);
      return res.json(inventory.assets.map((asset) => ({
        ...asset,
        batchId: bySerial.get(asset.serialNumber)?.batchId || "",
        receivedAt: bySerial.get(asset.serialNumber)?.receivedAt || "",
        poNumber: bySerial.get(asset.serialNumber)?.poNumber || "",
        supplier: bySerial.get(asset.serialNumber)?.supplier || "",
      })));
    } catch (error) {
      console.error("Local asset list error:", error);
      return res.status(500).json([]);
    }
  }
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
    const inboundBySerial = inboundMap(await readInboundRows(sheets, SPREADSHEET_ID));

    // ดึงชื่อชุดอุปกรณ์มาไว้แนบ เพื่อให้หน้า Asset แสดงได้ว่า "อยู่ในชุดอะไร"
    // (Asset_List เก็บแค่ BundleID ที่คอลัมน์ N — ชื่อชุดต้องมาจาก sheet Bundles)
    let bundleNameMap = cache.get("assetBundleNames");
    if (!bundleNameMap) {
      bundleNameMap = {};
      try {
        const bRes = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: "Bundles!A2:F",
        });
        (bRes.data.values || []).forEach((r) => {
          if (r[0]) bundleNameMap[r[0]] = r[1] || r[0];
        });
        cache.set("assetBundleNames", bundleNameMap);
      } catch (e) {
        console.error("❌ Load bundle names:", e.message);
      }
    }

    assets = assetRows.map((row) => {
      const inbound = inboundBySerial.get(String(row[4] || "").trim()) || {};
      return ({
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
      bundleName: row[13] ? bundleNameMap[row[13]] || row[13] : "",
      batchId: inbound.batchId || "",
      receivedAt: inbound.receivedAt || "",
      poNumber: inbound.poNumber || "",
      supplier: inbound.supplier || "",
    });
    });
    cache.set(cacheKey, assets);
    res.json(assets);
  } catch (error) {
    console.error("❌ Get Assets Error:", error);
    res.status(500).json([]);
  }
});

// -------------------- GET ASSET LOCATIONS --------------------
// รายการ "ตำแหน่งย่อยในไซต์" (คอลัมน์ G) ที่เคยถูกใช้มาแล้ว
// ใช้เป็นตัวเดา (datalist) ให้ผู้ใช้เลือกตำแหน่งที่ใช้บ่อย แต่ยังพิมพ์ค่าใหม่เองได้
router.get("/api/asset-locations", requireLogin, async (req, res) => {
  const cacheKey = "assetLocationList";
  let list = cache.get(cacheKey);
  if (list) return res.json(list);

  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!G2:G",
    });
    const rows = r.data.values || [];
    const seen = new Set();
    rows.forEach((row) => {
      const v = (row[0] || "").trim();
      // กรองค่าที่ไม่ใช่ตำแหน่งย่อยจริงออก — Stock ไม่ใช่ตำแหน่งย่อย
      if (!v || v === "-" || v === "Stock" || v === "Intranin") return;
      seen.add(v);
    });
    list = [...seen].sort((a, b) => a.localeCompare(b, "th"));
    cache.set(cacheKey, list);
    res.json(list);
  } catch (error) {
    console.error("❌ Get Asset Locations Error:", error);
    res.json([]);
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

// -------------------- RECENT ASSET TRANSFERS (หน้าแรก) --------------------
// การโอนย้ายอุปกรณ์รายชิ้นล่าสุดทั้งระบบ (Asset_History!A2:G)
// ใช้ในการ์ด "การโอนย้ายล่าสุด" ของหน้าแรก (Home Workdesk)
router.get("/api/asset-history-recent", requireLogin, async (req, res) => {
  try {
    const cacheKey = "assetHistoryRecent";
    let cached = cache.get(cacheKey);
    if (cached) return res.json(cached);

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_History!A2:G",
    });
    const rows = response.data.values || [];
    const recent = rows
      .filter((row) => row[1])
      .slice(-10)
      .reverse()
      .map((row) => ({
        date: row[0] || "-",
        serialNumber: row[1] || "-",
        action: row[2] || "-",
        from: row[3] || "-",
        to: row[4] || "-",
        user: row[5] || "-",
        remark: row[6] || "-",
      }));
    cache.set(cacheKey, recent, 60);
    res.json(recent);
  } catch (error) {
    console.error("❌ Recent Asset History Error:", error);
    res.status(500).json([]);
  }
});

async function respondInboundBatches(req, res, limit = Infinity) {
  try {
    let sheets = null;
    if (!isLocalInventoryMode()) {
      try { sheets = await getSheetsClient(); } catch (error) { console.warn(`[inbound] Google Sheets client unavailable; using local metadata (${error.message})`); }
    }
    const rows = await readInboundRows(sheets, SPREADSHEET_ID);
    res.json(summarizeInboundBatches(rows).slice(0, limit));
  } catch (err) {
    if (String(err.message || "").includes("Unable to parse range")) return res.json([]);
    console.error("Recent inbound batches:", err);
    res.status(500).json([]);
  }
}

router.get("/api/inbound-batches", requireLogin, async (req, res) => respondInboundBatches(req, res));
router.get("/api/inbound-batches/recent", requireLogin, async (req, res) => {
  return respondInboundBatches(req, res, 12);
});

async function loadInboundPOs() {
  let sheets = null;
  if (!isLocalInventoryMode()) sheets = await getSheetsClient();
  const [inboundRows, statusRows] = await Promise.all([readInboundRows(sheets, SPREADSHEET_ID), readPOStatuses(sheets, SPREADSHEET_ID)]);
  let assets;
  if (isLocalInventoryMode()) {
    assets = (await readLocalInventory()).assets;
  } else {
    const { data } = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Asset_List!A2:P" });
    assets = (data.values || []).map((row) => ({ serialNumber: row[4] || "", status: row[5] || "", location: row[6] || "", siteName: row[7] || "" }));
  }
  return { sheets, groups: summarizeInboundPOs(inboundRows, assets, statusRows) };
}

router.get("/api/inbound-pos", requireLogin, async (req, res) => {
  try {
    const { groups } = await loadInboundPOs();
    const filtered = req.query.status === "active" || req.query.status === "closed" ? groups.filter((po) => po.status === req.query.status) : groups;
    const limit = Number.parseInt(req.query.limit, 10);
    res.json(Number.isFinite(limit) && limit > 0 ? filtered.slice(0, limit) : filtered);
  } catch (error) {
    console.error("Inbound PO list error:", error);
    res.status(500).json({ error: "Unable to load inbound PO lifecycle" });
  }
});

router.post("/api/inbound-pos/:poNumber/close", requireLogin, async (req, res) => {
  const poNumber = String(req.params.poNumber || "").trim();
  if (!poNumber) return res.status(400).json({ error: "PO number is required" });
  try {
    const { sheets, groups } = await loadInboundPOs();
    if (!groups.some((po) => po.poNumber === poNumber)) return res.status(404).json({ error: "PO not found" });
    const record = await savePOStatus(sheets, SPREADSHEET_ID, poNumber, "closed", req.session.user.username);
    clearAssetCache();
    res.json({ success: true, ...record });
  } catch (error) {
    console.error("Close inbound PO error:", error);
    res.status(500).json({ error: "Unable to close PO" });
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
        range: "Asset_List!A2:Q",
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

    // ชื่อชุดอุปกรณ์ที่ตัวนี้อยู่ (ถ้ามี) — ใช้แก้การแสดงตำแหน่งให้ถูกชั้น
    let bundleName = "";
    if (row[13]) {
      let nameMap = cache.get("assetBundleNames");
      if (!nameMap) {
        nameMap = {};
        try {
          const bRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: "Bundles!A2:F",
          });
          (bRes.data.values || []).forEach((r) => {
            if (r[0]) nameMap[r[0]] = r[1] || r[0];
          });
          cache.set("assetBundleNames", nameMap);
        } catch (e) { /* ไม่มี Bundles ก็ข้าม */ }
      }
      bundleName = nameMap[row[13]] || row[13];
    }

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
      bundleId: row[13] || "",
      bundleName: row[13] ? bundleName : "",
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
    body("receivedAt").optional().isString(),
    body("poNumber").optional().isString(),
    body("supplier").optional().isString(),
    body("batchId").optional().isString(),
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
        receivedAt: receivedAtInput,
        poNumber = "",
        supplier = "",
        batchId: requestedBatchId = "",
      } = req.body;
      if (isLocalInventoryMode()) {
        let inbound = null;
        if (serialNumber) {
          const existing = await readInboundRows(null, SPREADSHEET_ID);
          inbound = resolveInboundBatch({ batchId: requestedBatchId, receivedAt: normalizeReceivedAt(receivedAtInput), poNumber, supplier, existingRows: existing });
        }
        const created = await addLocalAsset({ assetId, code, name, partNumber, serialNumber, status, location, siteName, user }, req.session.user.username);
        if (!created.success) return res.status(409).json(created);
        if (serialNumber) {
          const [row] = createInboundRows({ ...inbound, serials: [serialNumber], assetIds: [assetId], receivedBy: req.session.user.username });
          const storage = await appendInboundRows(null, SPREADSHEET_ID, [row]);
          if (inbound.poNumber) await savePOStatus(null, SPREADSHEET_ID, inbound.poNumber, "active");
          inbound = { ...inbound, inboundStorage: storage };
        }
        clearAssetCache();
        return res.json({ success: true, ...inbound });
      }
      const sheets = await getSheetsClient();
      const receivedAt = serialNumber ? normalizeReceivedAt(receivedAtInput) : null;
      let inboundBatch = null;
      if (serialNumber) {
        const existing = await readInboundRows(sheets, SPREADSHEET_ID);
        inboundBatch = resolveInboundBatch({ batchId: requestedBatchId, receivedAt, poNumber, supplier, existingRows: existing });
        await prepareInboundWrite();
      }
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
      let inbound = null;
      if (serialNumber) {
        const [inboundRow] = createInboundRows({ ...inboundBatch, serials: [serialNumber], assetIds: [assetId], receivedBy: req.session.user.username });
        const storage = await appendInboundRows(sheets, SPREADSHEET_ID, [inboundRow]);
        if (inboundBatch.poNumber) await savePOStatus(sheets, SPREADSHEET_ID, inboundBatch.poNumber, "active");
        inbound = { ...inboundBatch, inboundStorage: storage };
      }
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

res.json({ success: true, ...(inbound || {}) });   // ✅ ส่ง Response ทีหลัง
    } catch (error) {
      console.error("❌ Add Asset Error:", error);
      res.status(error.status || 500).json({ success: false, error: error.message });
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
      cache.del("assetHistoryRecent");

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
        user,
        remark,
        fromLocation,
        farmType,
        houseId,
        houseName,
        animalType,
      } = req.body;
      let location = req.body.location;
      let siteName = req.body.siteName;

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

      // ถ้า asset นี้เคยถูกจับเข้าชุด (มี BundleID ที่ N) และถูกคืนคลัง
      // ให้ถอดออกจาก Bundle ด้วย (ล้าง N/O/P/Q) กันไม่ให้ Bundle กลับมาเขียนสถานะทับ
      const wasInBundle = !!(prevRow[13] || "");
      if (isReturn && wasInBundle) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            valueInputOption: "USER_ENTERED",
            data: [
              { range: `Asset_List!N${rowIndex}`, values: [[""]] },
              { range: `Asset_List!O${rowIndex}`, values: [[""]] },
              { range: `Asset_List!P${rowIndex}`, values: [[""]] },
              { range: `Asset_List!Q${rowIndex}`, values: [[""]] },
            ],
          },
        });
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
      cache.del("assetHistoryRecent");

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
    body("batchId").optional().isString(),
  ],
  validate,
  async (req, res) => {
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
        receivedAt: receivedAtInput,
        poNumber = "",
        supplier = "",
        batchId: requestedBatchId = "",
      } = req.body;
      if (isLocalInventoryMode()) {
        const inboundExisting = await readInboundRows(null, SPREADSHEET_ID);
        const batch = resolveInboundBatch({ batchId: requestedBatchId, receivedAt: normalizeReceivedAt(receivedAtInput), poNumber, supplier, existingRows: inboundExisting });
        const created = await addLocalAssets({ partNumber, partName, qty, status, siteName, location, user, farmType, houseId, houseName }, req.session.user.username);
        const inboundRows = createInboundRows({ ...batch, serials: created.assets.map((asset) => asset.serialNumber), assetIds: created.assets.map((asset) => asset.assetId), receivedBy: req.session.user.username });
        await appendInboundRows(null, SPREADSHEET_ID, inboundRows);
        if (batch.poNumber) await savePOStatus(null, SPREADSHEET_ID, batch.poNumber, "active");
        cache.del("partCatalog");
        return res.json({ success: true, ...created, serials: created.assets.map((asset) => asset.serialNumber), ...batch, poNumber: batch.poNumber, supplier: batch.supplier, inboundStorage: "local-file" });
      }
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

      // 🔒 ห้าม Serial ซ้ำเด็ดขาด — ตรวจทุก Serial ที่จะเขียนกับ Serial ทั้งชีตก่อนเขียนจริง
      // (all-or-nothing: ซ้ำแม้แต่ตัวเดียว → ไม่เขียนอะไรลง Asset_List เลย แล้วแจ้งกลับ)
      const existingSerials = new Set(
        existingRows.map((r) => String(r[4] || "").trim()).filter(Boolean)
      );
      const dupRow = newRows.find((r) => existingSerials.has(String(r[4]).trim()));
      if (dupRow) {
        return res.status(400).json({
          success: false,
          error: `Serial ${dupRow[4]} ซ้ำกับที่มีอยู่แล้วในระบบ — ยกเลิกการเพิ่มทั้งหมด กรุณาลองใหม่อีกครั้ง`,
        });
      }
      // กันซ้ำภายในชุดเอง (กันข้อมูลเก่ารูปแบบไม่ตรงทำให้ maxNum คำนวณผิด)
      const batchSeen = new Set();
      const dupInBatch = newRows.find((r) => {
        const s = String(r[4]).trim();
        if (batchSeen.has(s)) return true;
        batchSeen.add(s);
        return false;
      });
      if (dupInBatch) {
        return res.status(400).json({
          success: false,
          error: `เกิด Serial ซ้ำภายในชุด (${dupInBatch[4]}) — ยกเลิกการเพิ่มทั้งหมด กรุณาลองใหม่อีกครั้ง`,
        });
      }

      const inboundExisting = await readInboundRows(sheets, SPREADSHEET_ID);
      const batch = resolveInboundBatch({ batchId: requestedBatchId, receivedAt: normalizeReceivedAt(receivedAtInput), poNumber, supplier, existingRows: inboundExisting });
      const inboundRows = createInboundRows({
        ...batch,
        serials: newRows.map((row) => row[4]),
        assetIds: newRows.map((row) => row[0]),
        receivedBy: req.session.user.username,
      });
      await prepareInboundWrite();

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

      const inboundStorage = await appendInboundRows(sheets, SPREADSHEET_ID, inboundRows);
      if (batch.poNumber) await savePOStatus(sheets, SPREADSHEET_ID, batch.poNumber, "active");

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
  ...batch,
  poNumber: batch.poNumber,
  supplier: batch.supplier,
  inboundStorage,
  firstSerial: newRows[0][4],
  lastSerial: newRows[newRows.length - 1][4],
});
    } catch (e) {
      console.error("Bulk add error:", e);
      res.status(e.status || 500).json({ success: false, error: e.message });
    }
  }
);
// ── Attach sync function to router for external call ──
router.syncInitialAssetHistory = syncInitialAssetHistory;
module.exports = router;
