// routes/farm.js
const express = require("express");
const router = express.Router();
const { body } = require("express-validator");
// เรียกผ่าน sheetsSvc.xxx (property access) แทนการ destructure getSheetsClient ไว้ล่วงหน้า
// → tests สามารถ monkey-patch services/sheets.getSheetsClient ก่อนเรียก handler ได้
const sheetsSvc = require("../services/sheets");
const { cache, SPREADSHEET_ID } = sheetsSvc;
const { logAudit } = require("../services/audit");
const { requireLogin, validate } = require("../middleware/auth");
const { findSiteDuplicate, findHouseDuplicate } = require("../services/farmGuard");
const { isLocalInventoryMode } = require("../services/localInventory");

// -------------------- GET FARM SITES --------------------
router.get("/api/farm-sites", requireLogin, async (req, res) => {
  if (isLocalInventoryMode()) return res.json([]);
  const cacheKey = "farmSites";
  let sites = cache.get(cacheKey);
  if (sites) return res.json(sites);

  try {
    const sheets = await sheetsSvc.getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Farm_Sites!A2:F",
    });
    const rows = r.data.values || [];
    sites = rows.map((row) => ({
      siteId: row[0] || "",
      siteName: row[1] || "",
      farmType: row[2] || "",
      province: row[3] || "",
      manager: row[4] || "",
      note: row[5] || "",
    }));
    cache.set(cacheKey, sites);
    res.json(sites);
  } catch (e) {
    console.error("Farm sites error:", e);
    res.status(500).json([]);
  }
});

// -------------------- GET FARMS (alias สำหรับหน้า Bundle) --------------------
// public/js/bundle.js เรียก /api/farms และคาดหวังฟิลด์ชื่อ farmId / farmName
// (ไม่ใช่ siteId / siteName แบบที่ /api/farm-sites คืนให้) — เส้นนี้ไม่เคยมีมาก่อน
// ทำให้ dropdown เลือกฟาร์มตอน deploy ว่างเปล่าถ้าฟาร์มนั้นยังไม่เคยถูก deploy ผ่าน
// Bundle มาก่อน (fallback ฝั่ง frontend สร้าง option จาก bundle ที่เคย deploy แล้วเท่านั้น)
router.get("/api/farms", requireLogin, async (req, res) => {
  const cacheKey = "farmSites";
  let sites = cache.get(cacheKey);

  try {
    if (!sites) {
      const sheets = await sheetsSvc.getSheetsClient();
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Sites!A2:F",
      });
      const rows = r.data.values || [];
      sites = rows.map((row) => ({
        siteId: row[0] || "",
        siteName: row[1] || "",
        farmType: row[2] || "",
        province: row[3] || "",
        manager: row[4] || "",
        note: row[5] || "",
      }));
      cache.set(cacheKey, sites);
    }

    const farms = sites.map((s) => ({
      farmId: s.siteId,
      farmName: s.siteName,
      farmType: s.farmType,
    }));
    res.json(farms);
  } catch (e) {
    console.error("❌ GET /api/farms:", e);
    res.status(500).json([]);
  }
});

// -------------------- GET FARM HOUSES --------------------
router.get("/api/farm-houses/:siteId", requireLogin, async (req, res) => {
  try {
    const siteId = req.params.siteId;
    const cacheKey = `farmHouses_${siteId}`;
    let houses = cache.get(cacheKey);
    if (houses) return res.json(houses);

    const sheets = await sheetsSvc.getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Farm_Houses!A2:F",
    });
    const rows = r.data.values || [];
    houses = rows
      .filter((row) => row[1] && row[1].trim() === siteId.trim())
      .map((row) => ({
        houseId: row[0] || "",
        siteId: row[1] || "",
        houseName: row[2] || "",
        houseType: row[3] || "",
        capacity: row[4] || "",
        note: row[5] || "",
      }));
    cache.set(cacheKey, houses);
    res.json(houses);
  } catch (e) {
    console.error("Farm houses error:", e);
    res.status(500).json([]);
  }
});

// -------------------- ADD FARM SITE --------------------
router.post("/api/add-farm-site",
  requireLogin,
  [
    body("siteId").trim().notEmpty(),
    body("siteName").trim().notEmpty(),
    body("farmType").optional().isString(),
    body("province").optional().isString(),
    body("manager").optional().isString(),
    body("note").optional().isString(),
  ],
  validate,
  // สิทธิ์: ผู้ใช้ที่ล็อกอินทุกคนเพิ่มฟาร์มได้ (เดิมจำกัด admin → คนอื่นโดน 403 จนติดขั้นตอนโอนย้าย)
  // สิทธิ์เข้าถึงคุมด้วย requireLogin ด้านบนแล้ว
  async (req, res) => {
    try {
      const { siteId, siteName, farmType, province, manager, note } = req.body;
      const sheets = await sheetsSvc.getSheetsClient();

      // ── กันเพิ่มซ้ำ: siteId หรือ siteName ที่มีอยู่แล้ว → 409
      // (อ่านชีตตรง ๆ ไม่ใช้ cache เพื่อให้เห็นข้อมูลล่าสุดเสมอ — cache จะถูกล้างหลัง append)
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Sites!A2:F",
      });
      const existing = (r.data.values || []).map((row) => ({
        siteId: row[0] || "",
        siteName: row[1] || "",
      }));
      const dup = findSiteDuplicate(existing, { siteId, siteName });
      if (dup) {
        return res.status(dup.code).json({ success: false, error: dup.error });
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Sites!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[siteId, siteName, farmType || "", province || "", manager || "", note || ""]],
        },
      });
      cache.del("farmSites");

      // ✅ บันทึก Audit Log (ต้องอยู่ก่อน res.json) — เดิมเพิ่มฟาร์มไม่มี audit เลย
      await logAudit(
        "เพิ่มฟาร์ม",
        "Farm",
        `Site ID: ${siteId}, ชื่อ: ${siteName}, ประเภท: ${farmType || "-"}`,
        req.session.user.username,
        req.ip || "-"
      );

      // site แนบกลับไปให้ frontend ใช้ "เลือกฟาร์มใหม่ทันที" หลังเพิ่มเสร็จ (inline add)
      res.json({ success: true, site: { siteId, siteName, farmType: farmType || "" } });
    } catch (error) {
      console.error("Add farm site error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// -------------------- ADD FARM HOUSE --------------------
router.post("/api/add-farm-house",
  requireLogin,
  [
    body("houseId").trim().notEmpty(),
    body("siteId").trim().notEmpty(),
    body("houseName").trim().notEmpty(),
    body("houseType").optional().isString(),
    body("capacity").optional().isString(),
    body("note").optional().isString(),
  ],
  validate,
  // สิทธิ์: ผู้ใช้ที่ล็อกอินทุกคนเพิ่มโรงเรือนได้ (เดิมจำกัด admin) — คุมด้วย requireLogin ด้านบน
  async (req, res) => {
    try {
      const { houseId, siteId, houseName, houseType, capacity, note } = req.body;
      const sheets = await sheetsSvc.getSheetsClient();

      // ── กันเพิ่มซ้ำ: houseId ที่มีอยู่แล้ว "ในฟาร์มเดิม" → 409
      // (รหัสเดียวกันแต่คนละฟาร์มไม่นับซ้ำ — โรงเรือนเป็นลูกของฟาร์ม)
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Houses!A2:F",
      });
      const existing = (r.data.values || []).map((row) => ({
        houseId: row[0] || "",
        siteId: row[1] || "",
      }));
      const dup = findHouseDuplicate(existing, { siteId, houseId });
      if (dup) {
        return res.status(dup.code).json({ success: false, error: dup.error });
      }

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Houses!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[houseId, siteId, houseName, houseType || "", capacity || "", note || ""]],
        },
      });
      cache.del(`farmHouses_${siteId}`);

      // ✅ บันทึก Audit Log (ต้องอยู่ก่อน res.json)
      await logAudit(
        "เพิ่มโรงเรือน",
        "Farm",
        `House ID: ${houseId}, ชื่อ: ${houseName}, ฟาร์ม: ${siteId}`,
        req.session.user.username,
        req.ip || "-"
      );

      // house แนบกลับไปให้ frontend "เลือกโรงเรือนใหม่ทันที" หลังเพิ่มเสร็จ (inline add)
      res.json({ success: true, house: { houseId, siteId, houseName, houseType: houseType || "" } });
    } catch (error) {
      console.error("Add farm house error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

module.exports = router;
