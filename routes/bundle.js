// routes/bundle.js
// ════════════════════════════════════════════════
//  INTRANIN EMS — Bundle System
//  จัดกลุ่มชุดอุปกรณ์ ย้ายทีเดียวทั้งชุด
//
//  Sheet: "Bundles"
//  A: BundleID  B: BundleName  C: Description
//  D: Status    E: Location    F: FarmID
//  G: AssetIDs  H: CreatedDate I: UpdatedDate  J: CreatedBy
//
//  Sheet: "Asset_List" (คอลัมน์ที่เกี่ยวข้องกับ Bundle)
//  F: Status  G: Location  H: SiteName
//  N: BundleID       — Bundle ปัจจุบันที่อุปกรณ์ชิ้นนี้อยู่ (ว่าง = ไม่ได้อยู่ใน Bundle ใดๆ)
//  O: PrevLocation    — ตำแหน่ง (Location) ก่อนถูกดึงเข้า Bundle ล่าสุด ใช้ restore ตอนถอดออก
//  P: PrevSiteName    — ไซต์ (SiteName) ก่อนถูกดึงเข้า Bundle ล่าสุด ใช้ restore ตอนถอดออก
//  ⚠️ ต้องเพิ่มหัวคอลัมน์ N/O/P นี้เองใน Google Sheet (BundleID / PrevLocation / PrevSiteName)
//     ถ้ายังไม่มี ไฟล์นี้จะยังทำงานได้ (อ่านค่าว่างแล้ว fallback) แต่การ restore ตอนถอด
//     ออกจากชุดจะไม่แม่นยำ 100%
//
//  หลักการ (ตามที่ต้องการ): Bundle คือ "สถานที่" อีกแห่งหนึ่งในระบบ
//  - อุปกรณ์ที่อยู่ใน Bundle ใดๆ: Location (G) จะ "มิเรอร์" ตาม Location ปัจจุบันของ Bundle
//    นั้นเป๊ะๆ เสมอ (ตอนสร้างใหม่ Bundle อยู่ที่ "Stock" อุปกรณ์ก็อยู่ที่นั่นด้วย,
//    พอ Bundle ถูก deploy ไปฟาร์มไหน อุปกรณ์ทุกชิ้นก็ตามไปฟาร์มนั้นทันที)
//  - SiteName (H) จะเป็น "ชื่อ Bundle" เสมอ ไม่ว่า Bundle จะ deploy อยู่ที่ไหนก็ตาม
//    เพื่อให้รู้ทันทีว่าอุปกรณ์ชิ้นนี้ถูกรวมอยู่ในชุดไหน โดยไม่ต้องเดาจาก Location
//  - เวลาย้ายทั้ง Bundle (deploy/recall) จะอัพเดต Location ของสมาชิกทุกชิ้นในทีเดียว
//    โดยไม่ต้องไปไล่ย้ายทีละตัว — และไม่แตะ SiteName เลย เพราะยังอยู่ใน Bundle เดิม
// ════════════════════════════════════════════════

const express = require("express");
const router  = express.Router();
const { body } = require("express-validator");
const {
  getSheetsClient,
  saveAssetHistory,
  clearAssetCache,
  cache,
  SPREADSHEET_ID,
} = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");
const { logAudit } = require("../services/audit");

const BUNDLE_SHEET  = "Bundles";         // ชื่อ Sheet ที่สร้างใหม่
const ASSET_SHEET   = "Asset_List";      // ชื่อเดียวกับ asset.js
const CACHE_KEY     = "bundleData";

// ─── helper: clear bundle cache ───────────────
function clearBundleCache() {
  cache.del(CACHE_KEY);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

// ─── helper: parse AssetIDs cell ──────────────
function parseIds(raw) {
  if (!raw) return [];
  return String(raw).split(",").map((s) => s.trim()).filter(Boolean);
}

// ─── helper: ป้ายชื่อ site ของอุปกรณ์เมื่ออยู่ใน Bundle — ใช้ค่านี้เดียวกันทุกจุด ──
function bundleLabel(bundleName, bundleId) {
  return `${bundleName} (${bundleId})`;
}

// ─── helper: ดึง Asset_List ทั้งหมด (สดจาก sheet ไม่ผ่าน cache
//     เพราะเรากำลังจะเขียนทับแถวพวกนี้ต่อ)
//     อ่านถึงคอลัมน์ P เพื่อให้ได้ BundleID / PrevLocation / PrevSiteName มาด้วย ──
async function getAssetRows(sheets) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${ASSET_SHEET}!A2:P`,
  });
  return resp.data.values || [];
}

// ─── helper: เมื่อ Asset ถูกเพิ่มเข้า Bundle (ตัวเดียวหรือหลายตัว)
//     - Location (G) := Location ปัจจุบันของ Bundle เป๊ะๆ (ไม่ว่า Bundle จะอยู่ office หรือฟาร์ม)
//     - SiteName (H) := ชื่อ Bundle เสมอ
//     - เก็บ Location/SiteName เดิมไว้ที่ O/P (แค่ครั้งแรกที่เข้าชุด กันไม่ให้ snapshot ทับตัวเอง
//       ถ้าเผลอเรียกซ้ำ) และ mark ว่าอยู่ Bundle ไหนที่ N
//     คืนค่า { applied, skippedOther } — skippedOther คือรายการที่ข้ามเพราะติดอยู่ใน Bundle อื่นแล้ว ──
async function applyBundleGroupingToAssets({
  sheets, assetIds, bundleId, bundleName, bundleLocation, user,
}) {
  const result = { applied: [], skippedOther: [] };
  if (!assetIds || !assetIds.length) return result;

  const assetRows = await getAssetRows(sheets);
  const updates      = [];
  const historyJobs  = [];
  const label        = bundleLabel(bundleName, bundleId);

  assetIds.forEach((assetId) => {
    const aIdx = assetRows.findIndex((r) => r[0] === assetId);
    if (aIdx === -1) return; // ไม่พบ asset นี้ใน Asset_List — ข้ามไปเงียบๆ ไม่ให้ request ทั้งก้อนล้ม

    const row = assetRows[aIdx];
    const currentBundleId = row[13] || "";
    if (currentBundleId && currentBundleId !== bundleId) {
      // อุปกรณ์นี้อยู่ใน Bundle อื่นอยู่แล้ว — ต้องถอดออกจากชุดเดิมก่อนถึงจะย้ายมาชุดนี้ได้
      result.skippedOther.push(assetId);
      return;
    }
    if (currentBundleId === bundleId) return; // อยู่ในชุดนี้อยู่แล้ว ไม่ต้องทำซ้ำ

    const aRow         = aIdx + 2;
    const serial        = row[4] || assetId;
    const prevLocation  = row[6] || "-";
    const prevSite      = row[7] || "-";

    updates.push({ range: `${ASSET_SHEET}!F${aRow}`, values: [["In Bundle"]] });
    updates.push({ range: `${ASSET_SHEET}!G${aRow}`, values: [[bundleLocation]] });
    updates.push({ range: `${ASSET_SHEET}!H${aRow}`, values: [[label]] });
    updates.push({ range: `${ASSET_SHEET}!N${aRow}`, values: [[bundleId]] });
    updates.push({ range: `${ASSET_SHEET}!O${aRow}`, values: [[prevLocation]] });
    updates.push({ range: `${ASSET_SHEET}!P${aRow}`, values: [[prevSite]] });

    result.applied.push(assetId);
    historyJobs.push(() => saveAssetHistory(
      serial,
      `เพิ่มเข้า Bundle: ${bundleId}`,
      `${prevSite} (${prevLocation})`,
      `${label} (${bundleLocation})`,
      user,
      `เพิ่มเข้าชุดอุปกรณ์ ${bundleName} (${bundleId}) — ตั้งแต่นี้ตำแหน่งของอุปกรณ์จะอ้างอิงตาม Bundle นี้`
    ));
  });

  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });
  }
  for (const job of historyJobs) await job();
  return result;
}

// ─── helper: เมื่อ Asset ถูกลบออกจาก Bundle
//     คืนสถานะกลับเป็นปกติ และคืน Location/SiteName กลับไปเป็นค่าก่อนเข้าชุด (เก็บไว้ที่ O/P)
//     ล้าง N/O/P ทิ้ง เพราะไม่ได้สังกัด Bundle ไหนแล้ว ──
async function revertAssetFromBundle({
  sheets, assetId, bundleId, bundleName, bundleLocation, user,
}) {
  const assetRows = await getAssetRows(sheets);
  const aIdx = assetRows.findIndex((r) => r[0] === assetId);
  if (aIdx === -1) return;

  const row   = assetRows[aIdx];
  const aRow  = aIdx + 2;
  const serial = row[4] || assetId;
  const label  = bundleLabel(bundleName, bundleId);

  // ถ้าไม่เคยมี snapshot เดิมเก็บไว้ (เช่นสมุดยังไม่มีคอลัมน์ O/P) ใช้ค่า default
  // ตามฐานที่ระบบใช้จริง คือของที่ยังไม่ได้ประกอบเข้า Bundle จะอยู่ที่ Stock/Intranin
  const restoreLocation = row[14] || "Stock";
  const restoreSite     = row[15] || "Intranin";

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        { range: `${ASSET_SHEET}!F${aRow}`, values: [["ใช้งานได้"]] },
        { range: `${ASSET_SHEET}!G${aRow}`, values: [[restoreLocation]] },
        { range: `${ASSET_SHEET}!H${aRow}`, values: [[restoreSite]] },
        { range: `${ASSET_SHEET}!N${aRow}`, values: [[""]] },
        { range: `${ASSET_SHEET}!O${aRow}`, values: [[""]] },
        { range: `${ASSET_SHEET}!P${aRow}`, values: [[""]] },
      ],
    },
  });

  await saveAssetHistory(
    serial,
    `ลบออกจาก Bundle: ${bundleId}`,
    `${label} (${bundleLocation})`,
    `${restoreSite} (${restoreLocation})`,
    user,
    `นำออกจากชุดอุปกรณ์ ${bundleName} (${bundleId}) — คืนตำแหน่งกลับเป็นก่อนเข้าชุด`
  );
}

// ─── helper: เมื่อทั้ง Bundle ถูกย้าย (deploy ไปฟาร์ม / recall กลับ office)
//     อัพเดต Location ของสมาชิกทุกชิ้นในทีเดียว โดยไม่แตะ SiteName เลย
//     เพราะ SiteName ต้องคงเป็นชื่อ Bundle เสมอไม่ว่าจะอยู่ที่ไหน ──
async function cascadeBundleLocationToAssets({
  sheets, assetIds, bundleId, bundleName, newLocation, action, note, user,
}) {
  if (!assetIds || !assetIds.length) return;
  const assetRows = await getAssetRows(sheets);
  const updates     = [];
  const historyJobs = [];
  const label = bundleLabel(bundleName, bundleId);

  assetIds.forEach((assetId) => {
    const aIdx = assetRows.findIndex((r) => r[0] === assetId);
    if (aIdx === -1) return;
    const row = assetRows[aIdx];
    const aRow = aIdx + 2;
    const serial      = row[4] || assetId;
    const oldLocation = row[6] || "-";

    updates.push({ range: `${ASSET_SHEET}!G${aRow}`, values: [[newLocation]] });
    // ไม่แตะ H (SiteName) — ยังคงโชว์ชื่อ Bundle เดิม ไม่ว่า Bundle จะอยู่ที่ไหน

    historyJobs.push(() => saveAssetHistory(
      serial,
      action,
      `${label} (${oldLocation})`,
      `${label} (${newLocation})`,
      user,
      note
    ));
  });

  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });
  }
  for (const job of historyJobs) await job();
}

// ════════════════════════════════════════════════
//  GET /api/bundles  — ดึงรายการ Bundle ทั้งหมด
// ════════════════════════════════════════════════
router.get("/api/bundles", requireLogin, async (req, res) => {
  let data = cache.get(CACHE_KEY);
  if (data) return res.json(data);

  try {
    const sheets = await getSheetsClient();
    const resp   = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${BUNDLE_SHEET}!A2:J`,
    });
    const rows = resp.data.values || [];

    data = rows
      .filter((r) => r[0]) // skip empty rows
      .map((r) => ({
        bundleId:    r[0] || "",
        bundleName:  r[1] || "",
        description: r[2] || "",
        status:      r[3] || "In Stock",
        location:    r[4] || "Stock",
        farmId:      r[5] || "",
        assetIds:    parseIds(r[6]),
        createdDate: r[7] || "",
        updatedDate: r[8] || "",
        createdBy:   r[9] || "",
      }));

    cache.set(CACHE_KEY, data, 60); // cache 60 วิ
    res.json(data);
  } catch (err) {
    console.error("❌ GET /api/bundles:", err);
    res.status(500).json({ error: "โหลด Bundle ล้มเหลว" });
  }
});

// ════════════════════════════════════════════════
//  POST /api/bundles  — สร้าง Bundle ใหม่
// ════════════════════════════════════════════════
router.post(
  "/api/bundles",
  requireLogin,
  [
    body("bundleId").trim().notEmpty().withMessage("bundleId required"),
    body("bundleName").trim().notEmpty().withMessage("bundleName required"),
  ],
  validate,
  async (req, res) => {
    try {
      const { bundleId, bundleName, description } = req.body;
      const user   = req.session.user.username || req.session.user.email || "system";
      const sheets = await getSheetsClient();

      // ตรวจ duplicate
      const existing = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:A`,
      });
      const ids = (existing.data.values || []).flat();
      if (ids.includes(bundleId)) {
        return res.status(400).json({ error: `Bundle ID "${bundleId}" มีอยู่แล้ว` });
      }

      const dt = today();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A:J`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[
            bundleId,
            bundleName,
            description || "",
            "In Stock",
            "Stock",
            "",          // FarmID
            "",          // AssetIDs
            dt,          // CreatedDate
            dt,          // UpdatedDate
            user,        // CreatedBy
          ]],
        },
      });

      clearBundleCache();
      await logAudit(user, "CREATE_BUNDLE", `สร้าง Bundle: ${bundleId} — ${bundleName}`);
      res.json({ success: true, message: "สร้าง Bundle สำเร็จ" });
    } catch (err) {
      console.error("❌ POST /api/bundles:", err);
      res.status(500).json({ error: "สร้าง Bundle ล้มเหลว" });
    }
  }
);

// ════════════════════════════════════════════════
//  PATCH /api/bundles/:id  — แก้ไขข้อมูล Bundle
// ════════════════════════════════════════════════
router.patch(
  "/api/bundles/:id",
  requireLogin,
  [body("bundleName").trim().notEmpty()],
  validate,
  async (req, res) => {
    try {
      const bundleId   = req.params.id;
      const { bundleName, description, status } = req.body;
      const user       = req.session.user.username || req.session.user.email;
      const sheets     = await getSheetsClient();

      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:J`,
      });
      const rows = resp.data.values || [];
      const idx  = rows.findIndex((r) => r[0] === bundleId);
      if (idx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

      const rowNum = idx + 2; // +2 เพราะ header อยู่แถว 1 และ A2 เริ่ม idx 0

      // ถ้าเปลี่ยนชื่อ Bundle ต้องอัพเดต SiteName ของสมาชิกทุกชิ้นให้ตรงชื่อใหม่ด้วย
      // เพราะ SiteName ของอุปกรณ์ = ชื่อ Bundle เสมอ
      const nameChanged = bundleName && bundleName !== rows[idx][1];
      const memberIds   = parseIds(rows[idx][6]);

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `${BUNDLE_SHEET}!B${rowNum}`, values: [[bundleName]] },
            { range: `${BUNDLE_SHEET}!C${rowNum}`, values: [[description || ""]] },
            { range: `${BUNDLE_SHEET}!D${rowNum}`, values: [[status || rows[idx][3]]] },
            { range: `${BUNDLE_SHEET}!I${rowNum}`, values: [[today()]] },
          ],
        },
      });

      if (nameChanged && memberIds.length) {
        const newLabel = bundleLabel(bundleName, bundleId);
        const assetRows = await getAssetRows(sheets);
        const updates = [];
        memberIds.forEach((assetId) => {
          const aIdx = assetRows.findIndex((r) => r[0] === assetId);
          if (aIdx === -1) return;
          updates.push({ range: `${ASSET_SHEET}!H${aIdx + 2}`, values: [[newLabel]] });
        });
        if (updates.length) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: { valueInputOption: "USER_ENTERED", data: updates },
          });
          clearAssetCache();
        }
      }

      clearBundleCache();
      await logAudit(user, "EDIT_BUNDLE", `แก้ไข Bundle: ${bundleId}`);
      res.json({ success: true, message: "อัพเดท Bundle สำเร็จ" });
    } catch (err) {
      console.error("❌ PATCH /api/bundles/:id:", err);
      res.status(500).json({ error: "แก้ไข Bundle ล้มเหลว" });
    }
  }
);

// ════════════════════════════════════════════════
//  POST /api/bundles/:id/assets  — เพิ่ม Asset เข้าชุด
// ════════════════════════════════════════════════
router.post(
  "/api/bundles/:id/assets",
  requireLogin,
  [body("assetId").trim().notEmpty()],
  validate,
  async (req, res) => {
    try {
      const bundleId = req.params.id;
      const { assetId } = req.body;
      const user     = req.session.user.username || req.session.user.email;
      const sheets   = await getSheetsClient();

      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:J`,
      });
      const rows = resp.data.values || [];
      const idx  = rows.findIndex((r) => r[0] === bundleId);
      if (idx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

      const ids = parseIds(rows[idx][6]);
      if (ids.includes(assetId)) {
        return res.status(400).json({ error: `Asset ${assetId} อยู่ในชุดนี้อยู่แล้ว` });
      }
      ids.push(assetId);

      const rowNum = idx + 2;
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `${BUNDLE_SHEET}!G${rowNum}`, values: [[ids.join(",")]] },
            { range: `${BUNDLE_SHEET}!I${rowNum}`, values: [[today()]] },
          ],
        },
      });

      // ── อัพเดตสถานะ/ตำแหน่งของ Asset ให้สะท้อนว่าถูกย้ายเข้าไปรวมในชุดนี้แล้ว ──
      const { skippedOther } = await applyBundleGroupingToAssets({
        sheets,
        assetIds: [assetId],
        bundleId,
        bundleName:     rows[idx][1],
        bundleLocation: rows[idx][4],
        user,
      });

      clearBundleCache();
      clearAssetCache();

      if (skippedOther.length) {
        return res.status(400).json({
          error: `${assetId} อยู่ใน Bundle อื่นอยู่แล้ว กรุณาถอดออกจากชุดเดิมก่อน`,
        });
      }

      await logAudit(user, "BUNDLE_ADD_ASSET", `เพิ่ม ${assetId} เข้า Bundle ${bundleId}`);
      res.json({ success: true, message: `เพิ่ม ${assetId} เข้าชุดสำเร็จ` });
    } catch (err) {
      console.error("❌ POST /api/bundles/:id/assets:", err);
      res.status(500).json({ error: "เพิ่ม Asset ล้มเหลว" });
    }
  }
);

// ════════════════════════════════════════════════
//  DELETE /api/bundles/:id/assets/:assetId  — ลบ Asset ออกจากชุด
// ════════════════════════════════════════════════
router.delete(
  "/api/bundles/:id/assets/:assetId",
  requireLogin,
  async (req, res) => {
    try {
      const { id: bundleId, assetId } = req.params;
      const user   = req.session.user.username || req.session.user.email;
      const sheets = await getSheetsClient();

      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:J`,
      });
      const rows = resp.data.values || [];
      const idx  = rows.findIndex((r) => r[0] === bundleId);
      if (idx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

      const ids      = parseIds(rows[idx][6]);
      const filtered = ids.filter((id) => id !== assetId);

      const rowNum = idx + 2;
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `${BUNDLE_SHEET}!G${rowNum}`, values: [[filtered.join(",")]] },
            { range: `${BUNDLE_SHEET}!I${rowNum}`, values: [[today()]] },
          ],
        },
      });

      // ── คืนสถานะ Asset กลับเป็นปกติ เพราะไม่ได้อยู่ในชุดนี้แล้ว ──
      await revertAssetFromBundle({
        sheets, assetId, bundleId,
        bundleName:     rows[idx][1],
        bundleLocation: rows[idx][4],
        user,
      });

      clearBundleCache();
      clearAssetCache();
      await logAudit(user, "BUNDLE_REMOVE_ASSET", `ลบ ${assetId} ออกจาก Bundle ${bundleId}`);
      res.json({ success: true, message: `ลบ ${assetId} ออกจากชุดสำเร็จ` });
    } catch (err) {
      console.error("❌ DELETE /api/bundles/:id/assets/:assetId:", err);
      res.status(500).json({ error: "ลบ Asset ล้มเหลว" });
    }
  }
);

// ════════════════════════════════════════════════
//  POST /api/bundles/:id/deploy  — ย้ายทั้งชุดไปฟาร์ม
// ════════════════════════════════════════════════
router.post(
  "/api/bundles/:id/deploy",
  requireLogin,
  [
    body("farmId").trim().notEmpty().withMessage("farmId required"),
    body("farmName").trim().notEmpty().withMessage("farmName required"),
  ],
  validate,
  async (req, res) => {
    try {
      const bundleId = req.params.id;
      const { farmId, farmName, note } = req.body;
      const user     = req.session.user.username || req.session.user.email;
      const sheets   = await getSheetsClient();

      // ── โหลด Bundle ──
      const bdlResp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:J`,
      });
      const bdlRows = bdlResp.data.values || [];
      const bdlIdx  = bdlRows.findIndex((r) => r[0] === bundleId);
      if (bdlIdx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

      const assetIds  = parseIds(bdlRows[bdlIdx][6]);
      const bundleName = bdlRows[bdlIdx][1];
      const bdlRow    = bdlIdx + 2;

      // ── 1. อัพเดท Bundle sheet — Bundle เองคือตัวที่ "เคลื่อนที่" ──
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `${BUNDLE_SHEET}!D${bdlRow}`, values: [["Deployed"]] },
            { range: `${BUNDLE_SHEET}!E${bdlRow}`, values: [[farmName]] },
            { range: `${BUNDLE_SHEET}!F${bdlRow}`, values: [[farmId]] },
            { range: `${BUNDLE_SHEET}!I${bdlRow}`, values: [[today()]] },
          ],
        },
      });

      // ── 2. ลากอุปกรณ์ทุกชิ้นในชุดตามไปที่ฟาร์มนั้นในทีเดียว (ไม่ต้องย้ายทีละตัว) ──
      //     SiteName ของอุปกรณ์ยังคงเป็นชื่อ Bundle เดิม เปลี่ยนแค่ Location ──
      await cascadeBundleLocationToAssets({
        sheets,
        assetIds,
        bundleId,
        bundleName,
        newLocation: farmName,
        action: `Deploy Bundle: ${bundleId}`,
        note: note || `ย้ายทั้งชุด Bundle ${bundleName} ไปที่ ${farmName} (${farmId})`,
        user,
      });

      clearBundleCache();
      clearAssetCache();
      await logAudit(
        user,
        "BUNDLE_DEPLOY",
        `Deploy Bundle ${bundleId} (${assetIds.length} อุปกรณ์) → ${farmName}`
      );
      res.json({
        success: true,
        message: `ย้าย Bundle ไป ${farmName} สำเร็จ (${assetIds.length} อุปกรณ์)`,
      });
    } catch (err) {
      console.error("❌ POST /api/bundles/:id/deploy:", err);
      res.status(500).json({ error: "Deploy Bundle ล้มเหลว" });
    }
  }
);

// ════════════════════════════════════════════════
//  POST /api/bundles/:id/recall  — คืน Bundle กลับ Stock
// ════════════════════════════════════════════════
router.post("/api/bundles/:id/recall", requireLogin, async (req, res) => {
  try {
    const bundleId = req.params.id;
    const user     = req.session.user.username || req.session.user.email;
    const sheets   = await getSheetsClient();

    // ── โหลด Bundle ──
    const bdlResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${BUNDLE_SHEET}!A2:J`,
    });
    const bdlRows = bdlResp.data.values || [];
    const bdlIdx  = bdlRows.findIndex((r) => r[0] === bundleId);
    if (bdlIdx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

    const prevFarm    = bdlRows[bdlIdx][4] || "ฟาร์ม";
    const assetIds    = parseIds(bdlRows[bdlIdx][6]);
    const bundleName  = bdlRows[bdlIdx][1];
    const bdlRow      = bdlIdx + 2;

    // ── 1. อัพเดท Bundle sheet ──
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: [
          { range: `${BUNDLE_SHEET}!D${bdlRow}`, values: [["In Stock"]] },
          { range: `${BUNDLE_SHEET}!E${bdlRow}`, values: [["Stock"]] },
          { range: `${BUNDLE_SHEET}!F${bdlRow}`, values: [[""]] },
          { range: `${BUNDLE_SHEET}!I${bdlRow}`, values: [[today()]] },
        ],
      },
    });

    // ── 2. ลากอุปกรณ์ทุกชิ้นกลับ Stock ในทีเดียว (SiteName ยังเป็นชื่อ Bundle เดิม) ──
    await cascadeBundleLocationToAssets({
      sheets,
      assetIds,
      bundleId,
      bundleName,
      newLocation: "Stock",
      action: `Recall Bundle: ${bundleId}`,
      note: `คืนทั้งชุด Bundle ${bundleName} กลับ Stock (จากเดิมที่ ${prevFarm})`,
      user,
    });

    clearBundleCache();
    clearAssetCache();
    await logAudit(
      user,
      "BUNDLE_RECALL",
      `Recall Bundle ${bundleId} (${assetIds.length} อุปกรณ์) ← ${prevFarm}`
    );
    res.json({
      success: true,
      message: `คืน Bundle กลับ Stock สำเร็จ (${assetIds.length} อุปกรณ์)`,
    });
  } catch (err) {
    console.error("❌ POST /api/bundles/:id/recall:", err);
    res.status(500).json({ error: "Recall Bundle ล้มเหลว" });
  }
});

// POST /api/bundles/:id/assets/bulk — เพิ่ม Asset หลายตัวพร้อมกัน
router.post(
  "/api/bundles/:id/assets/bulk",
  requireLogin,
  [body("assetIds").isArray({ min: 1 })],
  validate,
  async (req, res) => {
    try {
      const bundleId = req.params.id;
      const { assetIds } = req.body;
      const user   = req.session.user.username || req.session.user.email;
      const sheets = await getSheetsClient();

      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${BUNDLE_SHEET}!A2:J`,
      });
      const rows = resp.data.values || [];
      const idx  = rows.findIndex((r) => r[0] === bundleId);
      if (idx === -1) return res.status(404).json({ error: "ไม่พบ Bundle" });

      const existing = parseIds(rows[idx][6]);
      // รวม array โดยไม่ซ้ำ
      const newOnes = assetIds.filter((id) => !existing.includes(id));
      const merged  = Array.from(new Set(existing.concat(assetIds)));
      const added   = merged.length - existing.length;

      const rowNum = idx + 2;
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            { range: `${BUNDLE_SHEET}!G${rowNum}`, values: [[merged.join(",")]] },
            { range: `${BUNDLE_SHEET}!I${rowNum}`, values: [[today()]] },
          ],
        },
      });

      // ── อัพเดตสถานะ/ตำแหน่งของ Asset ที่เพิ่มเข้ามาใหม่จริงๆ เท่านั้น ──
      const { skippedOther } = await applyBundleGroupingToAssets({
        sheets,
        assetIds: newOnes,
        bundleId,
        bundleName:     rows[idx][1],
        bundleLocation: rows[idx][4],
        user,
      });

      clearBundleCache();
      clearAssetCache();
      await logAudit(user, "BUNDLE_BULK_ADD",
        `เพิ่ม ${added} Asset เข้า Bundle ${bundleId}: ${assetIds.join(", ")}`
      );

      let message = `เพิ่ม ${added - skippedOther.length} อุปกรณ์เข้าชุดสำเร็จ`;
      if (skippedOther.length) {
        message += ` (ข้าม ${skippedOther.join(", ")} เพราะอยู่ใน Bundle อื่นอยู่แล้ว)`;
      }
      res.json({ success: true, message });
    } catch (err) {
      console.error("❌ POST /api/bundles/:id/assets/bulk:", err);
      res.status(500).json({ error: "เพิ่ม Asset bulk ล้มเหลว" });
    }
  }
);
// ════════════════════════════════════════════════
//  GET /api/bundles/search-assets?q=xxx  — ค้นหา Asset สำหรับเพิ่มเข้าชุด
//  แสดงเฉพาะอุปกรณ์ที่ "ยังไม่ได้อยู่ใน Bundle ไหน" เท่านั้น (คอลัมน์ N ว่าง)
//  เพื่อกันเผลอหยิบอุปกรณ์ที่ถูกจับกลุ่มอยู่ในอีกชุดไปแล้วมาใส่ซ้ำ
// ════════════════════════════════════════════════
router.get("/api/bundles/search-assets", requireLogin, async (req, res) => {
  try {
    const q = (req.query.q || "").trim().toLowerCase();
    if (!q || q.length < 2) {
      return res.json([]);
    }

    const sheets = await getSheetsClient();
    const resp   = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${ASSET_SHEET}!A2:P`,
    });
    const assetRows = resp.data.values || [];

    const results = assetRows
      .filter((r) => {
        const assetId = (r[0] || "").toLowerCase();
        const name    = (r[2] || "").toLowerCase();
        const serial  = (r[4] || "").toLowerCase();
        const code    = (r[1] || "").toLowerCase();
        const inAnyBundle = !!(r[13] || "");
        const matches = assetId.includes(q) || name.includes(q) || serial.includes(q) || code.includes(q);
        return matches && !inAnyBundle;
      })
      .slice(0, 20)
      .map((r) => ({
        assetId:  r[0] || "",
        code:     r[1] || "",
        name:     r[2] || "",
        serial:   r[4] || "",
        status:   r[5] || "",
        location: r[6] || "",
        site:     r[7] || "",
      }));

    res.json(results);
  } catch (err) {
    console.error("❌ GET /api/bundles/search-assets:", err);
    res.status(500).json({ error: "ค้นหา Asset ล้มเหลว" });
  }
});

module.exports = router;