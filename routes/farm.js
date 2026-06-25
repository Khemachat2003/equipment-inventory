// routes/farm.js
const express = require("express");
const router = express.Router();
const { body } = require("express-validator");
const { getSheetsClient, cache, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");

// -------------------- GET FARM SITES --------------------
router.get("/api/farm-sites", requireLogin, async (req, res) => {
  const cacheKey = "farmSites";
  let sites = cache.get(cacheKey);
  if (sites) return res.json(sites);

  try {
    const sheets = await getSheetsClient();
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

// -------------------- GET FARM HOUSES --------------------
router.get("/api/farm-houses/:siteId", requireLogin, async (req, res) => {
  try {
    const siteId = req.params.siteId;
    const cacheKey = `farmHouses_${siteId}`;
    let houses = cache.get(cacheKey);
    if (houses) return res.json(houses);

    const sheets = await getSheetsClient();
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
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    try {
      const { siteId, siteName, farmType, province, manager, note } = req.body;
      const sheets = await getSheetsClient();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Sites!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[siteId, siteName, farmType || "", province || "", manager || "", note || ""]],
        },
      });
      cache.del("farmSites");
      res.json({ success: true });
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
  async (req, res) => {
    if (req.session.user.role !== "admin") {
      return res.status(403).json({ error: "ไม่มีสิทธิ์" });
    }
    try {
      const { houseId, siteId, houseName, houseType, capacity, note } = req.body;
      const sheets = await getSheetsClient();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Farm_Houses!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[houseId, siteId, houseName, houseType || "", capacity || "", note || ""]],
        },
      });
      cache.del(`farmHouses_${siteId}`);
      res.json({ success: true });
    } catch (error) {
      console.error("Add farm house error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

module.exports = router;