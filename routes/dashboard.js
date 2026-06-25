// routes/dashboard.js
const express = require("express");
const router = express.Router();
const { getSheetsClient, cache, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin } = require("../middleware/auth");

// -------------------- DASHBOARD --------------------
router.get("/api/dashboard", requireLogin, async (req, res) => {
  const cacheKey = "dashboardData";
  let dash = cache.get(cacheKey);
  if (dash) return res.json(dash);

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
    const log = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A2:H",
    });
    const masterData = master.data.values || [];
    const officeData = office.data.values || [];
    const siteData = site.data.values || [];
    const logData = log.data.values || [];

    const totalItems = masterData.length;
    let totalOffice = officeData.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);
    let totalSite = siteData.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);
    const today = new Date().toLocaleDateString("th-TH");
    let todayBorrow = 0,
      todayReturn = 0;
    logData.forEach((row) => {
      if (row[0] && row[0].includes(today)) {
        if (row[4] === "เบิก") todayBorrow += parseInt(row[3] || 0);
        if (row[4] === "คืน") todayReturn += parseInt(row[3] || 0);
      }
    });
    const result = { totalItems, totalOffice, totalSite, todayBorrow, todayReturn };
    cache.set(cacheKey, result);
    res.json(result);
  } catch (err) {
    console.error("Dashboard error:", err);
    res.status(500).json({ error: "Dashboard error" });
  }
});

// -------------------- DASHBOARD STATS --------------------
router.get("/api/dashboard-stats", requireLogin, async (req, res) => {
  const cacheKey = "dashboardStats";
  let stats = cache.get(cacheKey);
  if (stats) return res.json(stats);

  try {
    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A:H",
    });
    const rows = response.data.values || [];
    const data = rows.slice(1);
    const statsMap = {};
    const today = new Date();
    for (let i = 0; i < 30; i++) {
      const d = new Date();
      d.setDate(today.getDate() - i);
      const key = d.toISOString().slice(0, 10);
      statsMap[key] = { borrow: 0, return: 0 };
    }
    data.forEach((row) => {
      if (!row[0]) return;
      const [datePart] = row[0].split(" ");
      const [day, month, buddhistYear] = datePart.split("/");
      const year = parseInt(buddhistYear) - 543;
      const d = new Date(year, month - 1, day);
      const key = d.toISOString().slice(0, 10);
      if (!statsMap[key]) return;
      if (row[4] === "เบิก") statsMap[key].borrow++;
      if (row[4] === "คืน") statsMap[key].return++;
    });
    cache.set(cacheKey, statsMap, 120); // เพิ่ม TTL 2 นาที
    res.json(statsMap);
  } catch (err) {
    console.error("Dashboard stats error:", err);
    res.status(500).json({});
  }
});

// -------------------- DASHBOARD ASSET --------------------
router.get("/api/dashboard-asset", requireLogin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Asset_List!A2:M',
    });
    const assets = assetRes.data.values || [];

    const STOCK_SITE = 'Intranin';

    let total = 0;
    let totalStock = 0;
    let totalFarm = 0;
    const bySite = {};
    const byStatus = {};
    const farmCount = {};

    assets.forEach(row => {
      const site = row[7] || '-';
      const status = row[5] || 'ไม่ระบุ';

      total++;

      byStatus[status] = (byStatus[status] || 0) + 1;

      if (site === STOCK_SITE) {
        totalStock++;
        bySite[site] = (bySite[site] || 0) + 1;
      } else if (site && site !== '-') {
        totalFarm++;
        bySite[site] = (bySite[site] || 0) + 1;
        farmCount[site] = (farmCount[site] || 0) + 1;
      }
    });

    const topFarms = Object.entries(farmCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name, count]) => ({ name, count }));

    res.json({
      total,
      totalStock,
      totalFarm,
      bySite,
      byStatus,
      topFarms,
      totalFarms: Object.keys(farmCount).length,
    });

  } catch (error) {
    console.error('Dashboard Asset Error:', error);
    res.status(500).json({ error: 'Dashboard error' });
  }
});

// -------------------- DASHBOARD STATS EXTENDED --------------------
router.get("/api/dashboard-stats-extended", requireLogin, async (req, res) => {
  const cacheKey = "dashboardStatsExtended";
  let stats = cache.get(cacheKey);
  if (stats) return res.json(stats);

  try {
    const sheets = await getSheetsClient();

    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:M",
    });
    const assets = assetRes.data.values || [];

    const farmRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Farm_Sites!A2:F",
    });
    const farms = farmRes.data.values || [];

    const totalFarms = farms.length;
    const totalFarmAssets = assets.filter(a => a[7] && a[7] !== "-" && a[7] !== "").length;
    const totalStockAssets = assets.filter(a => !a[7] || a[7] === "-" || a[7] === "").length;
    const totalAssets = assets.length;

    const statusCount = {};
    assets.forEach(a => {
      const status = a[5] || "ไม่ระบุ";
      statusCount[status] = (statusCount[status] || 0) + 1;
    });

    const farmCount = {};
    assets.forEach(a => {
      const site = a[7] || "ไม่ระบุ";
      if (site !== "-" && site !== "") {
        farmCount[site] = (farmCount[site] || 0) + 1;
      }
    });
    const topFarms = Object.entries(farmCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name, count]) => ({ name, count }));

    const result = {
      totalFarms,
      totalAssets,
      totalFarmAssets,
      totalStockAssets,
      statusCount,
      topFarms,
    };

    cache.set(cacheKey, result, 120); // 2 นาที (จาก 5 นาที)
    res.json(result);
  } catch (error) {
    console.error("Dashboard stats extended error:", error);
    res.status(500).json({ error: "Dashboard stats error" });
  }
});

// -------------------- DASHBOARD FULL --------------------
router.get("/api/dashboard-full", requireLogin, async (req, res) => {
  const cacheKey = 'dashboardFull';
  let cached = cache.get(cacheKey);
  if (cached) return res.json(cached);

  try {
    const sheets = await getSheetsClient();

    const [assetRes, stockRes, officeRes, siteRes, logRes, farmRes, partRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Asset_List!A2:M' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Stock_Master!A2:I' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Stock_Office!A2:C' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Stock_Site!A2:C' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Transfer_Log!A2:H' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Farm_Sites!A2:F' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Part_Catalog!A2:G' }),
    ]);

    const assets = assetRes.data.values || [];
    const stock = stockRes.data.values || [];
    const office = officeRes.data.values || [];
    const site = siteRes.data.values || [];
    const logs = logRes.data.values || [];
    const farms = farmRes.data.values || [];
    const parts = partRes.data.values || [];

    const totalItems = stock.length;
    const totalOffice = office.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);
    const totalSite = site.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);

    const today = new Date().toLocaleDateString('th-TH');
    let todayBorrow = 0,
      todayReturn = 0;
    logs.forEach(row => {
      if (row[0] && row[0].includes(today)) {
        if (row[4] === 'เบิก') todayBorrow += parseInt(row[3] || 0);
        if (row[4] === 'คืน') todayReturn += parseInt(row[3] || 0);
      }
    });

    const STOCK_SITE = 'Intranin';
    let totalFarmAssets = 0,
      totalStockAssets = 0;
    const farmCount = {};
    const statusCount = {};
    const siteList = new Set();

    assets.forEach(row => {
      const site = row[7] || '-';
      const status = row[5] || 'ไม่ระบุ';

      statusCount[status] = (statusCount[status] || 0) + 1;

      if (site === STOCK_SITE) {
        totalStockAssets++;
      } else if (site && site !== '-') {
        totalFarmAssets++;
        farmCount[site] = (farmCount[site] || 0) + 1;
        siteList.add(site);
      }
    });

    const topFarms = Object.entries(farmCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name, count]) => ({ name, count }));

    const statsMap = {};
    const now = new Date();
    for (let i = 0; i < 30; i++) {
      const d = new Date(now);
      d.setDate(now.getDate() - i);
      const key = d.toISOString().slice(0, 10);
      statsMap[key] = { borrow: 0, return: 0 };
    }
    logs.forEach(row => {
      if (!row[0]) return;
      const [datePart] = row[0].split(' ');
      const [day, month, buddhistYear] = datePart.split('/');
      const year = parseInt(buddhistYear) - 543;
      const d = new Date(year, month - 1, day);
      const key = d.toISOString().slice(0, 10);
      if (statsMap[key]) {
        if (row[4] === 'เบิก') statsMap[key].borrow++;
        if (row[4] === 'คืน') statsMap[key].return++;
      }
    });

    const farmSites = farms.map(row => ({
      siteId: row[0] || '',
      siteName: row[1] || '',
      farmType: row[2] || '',
      province: row[3] || '',
      manager: row[4] || '',
      note: row[5] || '',
    }));

    const partList = parts.map(row => ({
      partNumber: row[0] || '',
      partName: row[1] || '',
      totalQty: parseInt(row[5]) || 0,
    }));

    const result = {
      totalItems,
      totalOffice,
      totalSite,
      todayBorrow,
      todayReturn,
      totalFarms: farmSites.length,
      totalFarmAssets,
      totalStockAssets,
      topFarms,
      statusCount,
      chartData: statsMap,
      farmSites,
      partCatalog: partList,
      assets: assets.map(row => ({
        assetId: row[0] || '-',
        code: row[1] || '-',
        name: row[2] || '-',
        partNumber: row[3] || '-',
        serialNumber: row[4] || '-',
        status: row[5] || '-',
        location: row[6] || '-',
        siteName: row[7] || '-',
        user: row[8] || '-',
      })),
    };

    cache.set(cacheKey, result, 120); // 2 นาที (จาก 10 นาที)
    res.json(result);

  } catch (error) {
    console.error('Dashboard Full Error:', error);
    res.status(500).json({ error: 'Dashboard error' });
  }
});

module.exports = router;