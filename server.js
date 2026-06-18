require("dotenv").config();
const express = require("express");
const QRCode = require("qrcode");
const { google } = require("googleapis");
const { Octokit } = require("@octokit/rest");
const octokit = new Octokit({ auth: process.env.GITHUB_TOKEN });
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const session = require("express-session");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const rateLimit = require("express-rate-limit");
const NodeCache = require("node-cache");
const { body, validationResult } = require("express-validator");

// -------------------- CONFIG --------------------
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const PORT = process.env.PORT || 3000;
const CACHE_TTL = parseInt(process.env.CACHE_TTL) || 300;

// -------------------- CACHE --------------------
const cache = new NodeCache({ stdTTL: CACHE_TTL, checkperiod: 60 });

// -------------------- RATE LIMIT --------------------
const limiter = rateLimit({
  windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000,
  max: parseInt(process.env.RATE_LIMIT_MAX) || 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many requests, please try again later." },
});

// -------------------- GOOGLE OAUTH --------------------
console.log("ENV GOOGLE_OAUTH:", process.env.GOOGLE_OAUTH ? "present" : "missing");
let credentials = null;
if (process.env.GOOGLE_OAUTH) {
  try {
    credentials = JSON.parse(process.env.GOOGLE_OAUTH);
  } catch (e) {
    console.error("❌ Invalid GOOGLE_OAUTH JSON:", e.message);
  }
}
if (!credentials || !credentials.web) {
  throw new Error("❌ Invalid GOOGLE_OAUTH JSON format");
}
const { client_id, client_secret, redirect_uris } = credentials.web;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// -------------------- SERVICE ACCOUNT --------------------
let serviceAccount = null;
if (process.env.GOOGLE_CREDENTIALS_ACCOUNT) {
  try {
    serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS_ACCOUNT);
    console.log("✅ Google credentials loaded");
  } catch (e) {
    console.error("❌ Invalid GOOGLE_CREDENTIALS_ACCOUNT JSON:", e.message);
  }
} else {
  console.error("❌ GOOGLE_CREDENTIALS_ACCOUNT env missing");
}
const backendAuth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"],
});

// -------------------- EXPRESS APP --------------------
const app = express();

// Trust proxy (for rate limit behind reverse proxy)
app.set("trust proxy", 1);

// CORS
app.use(cors({ origin: true, credentials: true }));

// Session
app.use(
  session({
    name: "borrow-session",
    secret: process.env.SESSION_SECRET || "default-insecure-secret-change-me",
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 24 * 60 * 60 * 1000, // 1 day
    },
  })
);

// Body parser
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Static files
app.use(express.static("public"));
app.use("/image", express.static(path.join(__dirname, "image")));

// -------------------- RATE LIMIT (apply to all API routes) --------------------
app.use("/api", limiter);

// -------------------- MIDDLEWARE --------------------
function requireLogin(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

// Validation error handler
function validate(req, res, next) {
  const errors = validationResult(req);
  if (!errors.isEmpty()) {
    return res.status(400).json({ errors: errors.array() });
  }
  next();
}

// Global error handler
app.use((err, req, res, next) => {
  console.error("❌ Unhandled error:", err);
  res.status(500).json({ error: "Internal server error", message: err.message });
});

// -------------------- GOOGLE SHEETS HELPERS --------------------
async function getSheetsClient() {
  const authClient = await backendAuth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

// Save asset history (with retry logic?)
async function saveAssetHistory(serialNumber, action, fromLocation, toLocation, username, remark) {
  const sheets = await getSheetsClient();
  const currentDate = new Date().toLocaleString("th-TH");
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Asset_History!A:G",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[currentDate, serialNumber, action, fromLocation, toLocation, username, remark]],
    },
  });
}

// -------------------- SYNC ASSET HISTORY (ONCE) --------------------
let syncDone = false;
async function syncInitialAssetHistory() {
  if (syncDone) return;
  try {
    const sheets = await getSheetsClient();
    const assetResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:J",
    });
    const assetRows = assetResponse.data.values || [];

    const historyResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_History!A2:B",
    });
    const historyRows = historyResponse.data.values || [];
    const loggedSerials = new Set(historyRows.map(row => (row[1] ? row[1].trim() : "")));

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

// -------------------- ROUTES --------------------
// AUTH
app.get("/auth/google", (req, res) => {
  const state = JSON.stringify({ user: req.session.user || null });
  const url = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "profile",
      "email",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/spreadsheets.readonly",
    ],
    state,
  });
  res.redirect(url);
});

app.get("/auth/google/callback", async (req, res) => {
  try {
    const code = req.query.code;
    const state = JSON.parse(req.query.state || "{}");
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    if (state.user) req.session.user = state.user;
    req.session.tokens = tokens;
    req.session.save(() => res.redirect("/dashboard"));
  } catch (err) {
    console.error("OAuth error:", err);
    res.redirect("/");
  }
});

// LOGIN
app.post(
  "/api/login",
  [
    body("username").trim().notEmpty().withMessage("Username required"),
    body("password").trim().notEmpty().withMessage("Password required"),
  ],
  validate,
  async (req, res) => {
    try {
      const { username, password } = req.body;
      const sheets = await getSheetsClient();
      const userRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users!A2:C",
      });
      const users = userRes.data.values || [];
      const user = users.find(
        (u) => u[0]?.trim() === username.trim() && u[1]?.trim() === password.trim()
      );
      if (!user) {
        return res.status(401).json({ error: "Username หรือ Password ไม่ถูกต้อง" });
      }
      req.session.user = { username, role: user[2] };
      req.session.save(() => res.json({ success: true, role: user[2] }));
    } catch (err) {
      console.error("LOGIN ERROR:", err);
      res.status(500).json({ error: "Server error", message: err.message });
    }
  }
);

app.get("/api/check-auth", (req, res) => {
  if (!req.session.user) return res.json({ loggedIn: false });
  res.json({ loggedIn: true, username: req.session.user.username, role: req.session.user.role });
});

app.post("/api/logout", (req, res) => {
  req.session.destroy(() => res.json({ success: true }));
});

// -------------------- STOCK (with cache) --------------------
app.get("/api/stock", requireLogin, async (req, res) => {
  const cacheKey = "stockData";
  let stock = cache.get(cacheKey);
  if (stock) {
    return res.json(stock);
  }
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

app.post(
  "/api/update-total",
  requireLogin,
  [body("code").trim().notEmpty(), body("newTotal").isInt({ min: 0 })],
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
      cache.del("stockData");
      res.json({ success: true });
    } catch (err) {
      console.error("Update total error:", err);
      res.status(500).json({ error: "Update error" });
    }
  }
);

app.post(
  "/api/add-stock",
  requireLogin,
  [body("code").trim().notEmpty(), body("qty").isInt({ min: 1 })],
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
      cache.del("stockData");
      res.json({ success: true });
    } catch (err) {
      console.error("Add stock error:", err);
      res.status(500).json({ error: "Add stock error" });
    }
  }
);

app.post(
  "/api/transfer",
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
      cache.del("stockData");
      res.json({ success: true });
    } catch (err) {
      console.error("Transfer error:", err);
      res.status(500).json({ error: "Transfer error" });
    }
  }
);

// HISTORY
app.get("/api/history", requireLogin, async (req, res) => {
  try {
    const { start, end } = req.query;
    const sheets = await getSheetsClient();
    const log = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A2:H",
    });
    let logData = log.data.values || [];

    if (start || end) {
      const startDate = start ? new Date(start) : null;
      const endDate = end ? new Date(end) : null;
      logData = logData.filter((row) => {
        if (!row[0]) return false;
        const [datePart] = row[0].split(" ");
        const [day, month, yearBE] = datePart.split("/");
        const year = parseInt(yearBE) - 543;
        const rowDate = new Date(year, month - 1, day);
        if (startDate && rowDate < startDate) return false;
        if (endDate) {
          const e = new Date(endDate);
          e.setHours(23, 59, 59);
          if (rowDate > e) return false;
        }
        return true;
      });
    }

    const history = logData.reverse().map((row) => ({
      date: row[0],
      code: row[1],
      name: row[2],
      qty: row[3],
      type: row[4],
      from: row[5],
      to: row[6],
      user: row[7],
    }));
    res.json(history);
  } catch (err) {
    console.error("History error:", err);
    res.status(500).json({ error: "History error" });
  }
});

// DASHBOARD (with cache)
app.get("/api/dashboard", requireLogin, async (req, res) => {
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

// ADD ITEM
app.post(
  "/api/add-item",
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
    const { code, name, total, office, site, ext } = req.body;
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
      cache.del("stockData");
      res.json({ success: true });
    } catch (error) {
      console.error("Add item error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

// UPLOAD IMAGE
app.post(
  "/upload-image",
  [
    body("fileName").trim().notEmpty(),
    body("base64").notEmpty().custom((val) => val.startsWith("data:image/")),
  ],
  validate,
  async (req, res) => {
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

// EXPORT HISTORY PDF
app.post(
  "/api/export-history",
  requireLogin,
  [
    body("startDate").isISO8601().withMessage("Invalid start date"),
    body("endDate").isISO8601().withMessage("Invalid end date"),
    body("title").optional().isString(),
    body("locations").optional().isString(),
    body("vehicle").optional().isString(),
    body("employeeCount").optional().isInt({ min: 0 }),
    body("employees").optional().isString(),
    body("reportType").optional().isIn(["all", "borrow", "return"]),
  ],
  validate,
  async (req, res) => {
    try {
      const title = req.body.title || "รายงานประวัติการเบิก–คืนอุปกรณ์";
      const locations = req.body.locations || "-";
      const vehicle = req.body.vehicle || "-";
      const startDate = req.body.startDate || null;
      const endDate = req.body.endDate || null;
      const employeeCount = req.body.employeeCount || "0";
      const employees = req.body.employees || "";
      const reportType = req.body.reportType || "all";

      const sheets = await getSheetsClient();
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Transfer_Log!A:H",
      });
      const rows = response.data.values || [];
      const dataRows = rows.slice(1);

      let filteredData = dataRows;
      if (reportType === "borrow") {
        filteredData = filteredData.filter((row) => row[4] === "เบิก");
      } else if (reportType === "return") {
        filteredData = filteredData.filter((row) => row[4] === "คืน");
      }

      if (startDate && endDate) {
        const [startY, startM, startD] = startDate.split("-");
        const [endY, endM, endD] = endDate.split("-");
        const start = new Date(startY, startM - 1, startD, 0, 0, 0);
        const end = new Date(endY, endM - 1, endD, 23, 59, 59);
        filteredData = filteredData.filter((row) => {
          if (!row[0]) return false;
          const [datePart] = row[0].split(" ");
          const [day, month, buddhistYear] = datePart.split("/");
          const year = parseInt(buddhistYear) - 543;
          const rowDate = new Date(year, month - 1, day, 12, 0, 0);
          return rowDate >= start && rowDate <= end;
        });
      }

      const filteredRows = filteredData.map((row) => [
        row[0] || "-",
        row[2] || "-",
        row[3] || "-",
        row[4] || "-",
        row[7] || "-",
      ]);

      if (filteredRows.length === 0) {
        return res.status(400).json({ error: "ไม่มีข้อมูลสำหรับ Export" });
      }
      if (filteredRows.length > 1000) {
        return res.status(400).json({ error: "ข้อมูลมากเกินไป กรุณาเลือกช่วงวันที่" });
      }

      const fileName = `Report-${new Date().toISOString().slice(0, 10)}.pdf`;
      const doc = new PDFDocument({ size: "A4", margins: { top: 100, bottom: 60, left: 60, right: 60 } });
      const filePath = `/tmp/${fileName}`;
      const stream = fs.createWriteStream(filePath);
      doc.pipe(stream);
      doc.registerFont("THSarabun", path.join(__dirname, "fonts", "THSarabunNew.ttf"));
      doc.font("THSarabun").fontSize(14);

      const marginLeft = doc.page.margins.left;
      const marginRight = doc.page.margins.right;
      const contentWidth = doc.page.width - marginLeft - marginRight;

      function drawReportHeader() {
        doc.image(path.join(__dirname, "logo.png"), marginLeft, 40, { width: 70 });
        doc.y = 45;
        doc.fontSize(20).text("บริษัท อินทนิล ออโตเมชั่น จำกัด", { align: "center", width: contentWidth });
        doc.fontSize(16).text(title, { align: "center", width: contentWidth });
        doc.moveDown(0.5);
        doc.moveTo(marginLeft, doc.y).lineTo(doc.page.width - marginRight, doc.y).stroke();
        doc.moveDown(1);
      }
      drawReportHeader();
      doc.on("pageAdded", drawReportHeader);

      doc.fontSize(14);
      doc.x = marginLeft;
      if (reportType === "all" || reportType === "borrow") {
        if (locations) {
          const locationList = locations.split("\n").map((l) => l.trim()).filter((l) => l !== "");
          doc.text(`สถานที่: ${locationList.length} ที่`, { width: contentWidth });
          locationList.forEach((loc, index) => {
            const clean = loc.replace(/^\d+\.\s*/, "");
            doc.text(`${index + 1}. ${clean}`, { width: contentWidth, indent: 20 });
          });
        }
        doc.moveDown(0.5);
        doc.text(`ยานพาหนะ: ${vehicle}`, { width: contentWidth });
        doc.text(`จำนวนพนักงาน: ${employeeCount} คน`, { width: contentWidth });
        if (employees) {
          doc.moveDown(0.5);
          employees.split("\n").map((n) => n.trim()).filter((n) => n !== "").forEach((name, index) => {
            const clean = name.replace(/^\d+\.\s*/, "");
            doc.text(`${index + 1}. ${clean}`, { width: contentWidth, indent: 20 });
          });
        }
      }
      doc.moveDown(0.5);
      function formatDate(dateStr) {
        if (!dateStr) return "-";
        const d = new Date(dateStr);
        const day = String(d.getDate()).padStart(2, "0");
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
      }
      doc.text(`ช่วงวันที่: ${formatDate(startDate)} - ${formatDate(endDate)}`, { width: contentWidth });
      doc.text("วันที่ออกรายงาน: " + new Date().toLocaleString("th-TH"), { width: contentWidth });
      doc.moveDown(0);

      let reportTypeText = "รวมรายการเบิกและคืน";
      if (reportType === "borrow") reportTypeText = "รายการเบิก";
      if (reportType === "return") reportTypeText = "รายการคืน";
      doc.fontSize(16);
      doc.text(`ตาราง: ${reportTypeText}`, { width: contentWidth, align: "center" });
      doc.moveDown(0.5);

      let y = doc.y;
      const usableHeight = doc.page.height - doc.page.margins.bottom;
      const columns = [
        { header: "วันที่", width: 95 },
        { header: "ชื่ออุปกรณ์", width: 170 },
        { header: "จำนวน", width: 50 },
        { header: "ประเภท", width: 60 },
        { header: "ผู้ทำรายการ", width: 85 },
      ];

      function drawTableHeader() {
        doc.font("THSarabun").fontSize(14);
        const headerHeight = 25;
        if (y + headerHeight > usableHeight) {
          doc.addPage();
          y = doc.y;
          doc.font("THSarabun").fontSize(14);
        }
        let x = doc.page.margins.left;
        columns.forEach((col) => {
          doc.rect(x, y, col.width, headerHeight).fillAndStroke("#f2f2f2", "black");
          doc.fillColor("black").text(col.header, x + 5, y + 7, { width: col.width - 10, align: "center" });
          x += col.width;
        });
        y += headerHeight;
      }
      drawTableHeader();

      filteredRows.forEach((row) => {
        let maxHeight = 0;
        columns.forEach((col, i) => {
          const cellText = row[i] || "-";
          const textHeight = doc.heightOfString(cellText, { width: col.width - 10 });
          if (textHeight > maxHeight) maxHeight = textHeight;
        });
        const rowHeight = maxHeight + 10;
        if (y + rowHeight > usableHeight) {
          doc.addPage();
          y = doc.y;
          doc.font("THSarabun").fontSize(14);
          drawTableHeader();
        }
        let x = doc.page.margins.left;
        columns.forEach((col, i) => {
          const cellText = row[i] || "-";
          doc.rect(x, y, col.width, rowHeight).stroke();
          doc.text(cellText, x + 5, y + 5, { width: col.width - 10, align: i === 2 ? "center" : "left" });
          x += col.width;
        });
        y += rowHeight;
      });

      doc.moveDown(3);
      const pageWidth = doc.page.width;
      const leftX = marginLeft;
      const rightX = marginLeft + (pageWidth - marginLeft - marginRight) / 2;
      const today = new Date().toLocaleDateString("th-TH");
      doc.text("ผู้ทำรายการ", leftX, doc.y, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });
      doc.text("ผู้ตรวจสอบ", rightX, doc.y - 14, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });
      doc.moveDown(2);
      doc.text("(....................................)", leftX, doc.y, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });
      doc.text("(....................................)", rightX, doc.y - 14, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });
      doc.moveDown(1);
      doc.text(`วันที่ ${today}`, leftX, doc.y, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });
      doc.text(`วันที่ ${today}`, rightX, doc.y - 14, { width: (pageWidth - marginLeft - marginRight) / 2, align: "center" });

      doc.end();
      stream.on("finish", () => {
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.download(filePath, fileName, (err) => {
          if (err) console.error("Download error:", err);
          setTimeout(() => fs.unlink(filePath, () => {}), 3000);
        });
      });
    } catch (err) {
      console.error("Export Error:", err);
      res.status(500).json({ error: "Export ล้มเหลว" });
    }
  }
);

// SITE ITEMS
app.get("/api/get-site-items", requireLogin, async (req, res) => {
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
    res.json({ items });
  } catch (err) {
    console.error("get-site-items error:", err);
    res.status(500).json({ items: [] });
  }
});

app.post("/api/return-all-site", requireLogin, async (req, res) => {
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
    for (let i = 0; i < siteData.length; i++) {
      const code = siteData[i][0];
      const name = siteData[i][1];
      let siteQty = parseInt(siteData[i][2] || 0);
      if (siteQty > 0) {
        const officeIndex = officeData.findIndex((r) => r[0] === code);
        if (officeIndex === -1) continue;
        let officeQty = parseInt(officeData[officeIndex][2] || 0);
        officeQty += siteQty;
        updates.push({ range: `Stock_Office!C${officeIndex + 2}`, values: [[officeQty]] });
        updates.push({ range: `Stock_Site!C${i + 2}`, values: [[0]] });
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
    cache.del("stockData");
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false });
  }
});

app.post(
  "/api/return-selected-site",
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
        updates.push({ range: `Stock_Office!C${officeIndex + 2}`, values: [[officeQty]] });
        updates.push({ range: `Stock_Site!C${siteIndex + 2}`, values: [[siteQty]] });
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
      cache.del("stockData");
      res.json({ success: true });
    } catch (err) {
      console.error(err);
      res.status(500).json({ success: false });
    }
  }
);

// ASSETS
app.get("/api/assets", requireLogin, async (req, res) => {
  const cacheKey = "assetData";
  let assets = cache.get(cacheKey);
  if (assets) {
    return res.json(assets);
  }
  try {
    await syncInitialAssetHistory();
    const sheets = await getSheetsClient();
    const assetResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:J",
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
    }));
    cache.set(cacheKey, assets);
    res.json(assets);
  } catch (error) {
    console.error("❌ Get Assets Error:", error);
    res.status(500).json([]);
  }
});

app.get("/api/asset-history/:serial", requireLogin, async (req, res) => {
  try {
    const serialNumber = req.params.serial;
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
    res.json(filteredHistory.reverse());
  } catch (error) {
    console.error("❌ Get Asset History Error:", error);
    res.status(500).json([]);
  }
});

app.get("/api/public-asset-history/:serial", async (req, res) => {
  try {
    const serialNumber = req.params.serial;
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
    res.json(filteredHistory.reverse());
  } catch (error) {
    console.error(error);
    res.status(500).json([]);
  }
});

app.get("/api/dashboard-stats", requireLogin, async (req, res) => {
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
    cache.set(cacheKey, statsMap);
    res.json(statsMap);
  } catch (err) {
    console.error("Dashboard stats error:", err);
    res.status(500).json({});
  }
});

app.post(
  "/api/add-asset",
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
      cache.del("assetData");
      res.json({ success: true });
    } catch (error) {
      console.error("❌ Add Asset Error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

app.get("/api/me", requireLogin, (req, res) => {
  res.json({ username: req.session.user.username });
});

app.post(
  "/api/update-asset-status",
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
      const rowIndex = assetRows.findIndex(
        (row) => row[4] && row[4].trim() === serialNumber.trim()
      ) + 2;
      if (rowIndex === 1) {
        return res.status(400).json({ success: false, error: "ไม่พบข้อมูล Serial Number นี้ในหน้าหลัก (Asset_List)" });
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
      cache.del("assetData");
      res.json({ success: true });
    } catch (error) {
      console.error("❌ Update Asset Status Backend Error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  }
);

app.get("/api/qrcode/:serial", requireLogin, async (req, res) => {
  try {
    const serial = req.params.serial;
    const sheets = await getSheetsClient();
    const url = `${req.protocol}://${req.get("host")}/trace.html?serial=${encodeURIComponent(serial)}`;
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

app.get("/api/public/asset/:code", async (req, res) => {
  try {
    const code = req.params.code;
    const sheets = await getSheetsClient();
    const assetRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Asset_List!A2:J",
    });
    const rows = assetRes.data.values || [];
    const row = rows.find((r) => r[4] && r[4].trim() === code.trim());
    if (!row) {
      return res.status(404).json({ error: "not found" });
    }
    const qrUrl = `${req.protocol}://${req.get("host")}/trace.html?serial=${encodeURIComponent(code)}`;
    const qrImage = await QRCode.toDataURL(qrUrl);
    res.json({
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
      qr: qrImage,
      traceUrl: qrUrl,
    });
  } catch (err) {
    console.error("Public asset error:", err);
    res.status(500).json({ error: "server error" });
  }
});

// -------------------- FRONTEND ROUTES --------------------
app.get("/dashboard", (req, res) => {
  if (!req.session.user) return res.redirect("/");
  res.sendFile(__dirname + "/public/index.html");
});
app.get("/", (req, res) => {
  res.sendFile(__dirname + "/public/index.html");
});
app.get("/a/:code", (req, res) => {
  res.sendFile(path.join(__dirname, "public/asset.html"));
});
app.use(express.static("public"));

// Health check
app.get("/health", (req, res) => {
  res.status(200).json({ status: "ok", uptime: process.uptime() });
});

// -------------------- START SERVER --------------------
app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
  // Trigger initial sync (non-blocking)
  syncInitialAssetHistory().catch(console.error);
});