require("dotenv").config();
const express = require("express");
const { google } = require("googleapis");
const { Octokit } = require("@octokit/rest");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const session = require("express-session");
const axios = require("axios");
const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
const { client_id, client_secret, redirect_uris } = credentials.web;

const oAuth2Client = new google.auth.OAuth2(
  client_id,
  client_secret,
  redirect_uris[0]
);

const app = express();

let stockData = [];
let transferLog = [];

/* =========================
   🔥 IMPORTANT FIX START
========================= */

// ถ้า deploy บน https (Render / Railway) ให้เปิดบรรทัดนี้
// app.set("trust proxy", 1);

app.use(cors({
  origin: true,          // หรือใส่ URL frontend เช่น "http://localhost:5173"
  credentials: true
}));

app.use(session({
  name: "borrow-session",
  secret: "borrow-return-secret",
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: false,        // ถ้า localhost = false
    sameSite: "lax"
  }
}));
// =====================
// AUTH MIDDLEWARE
// =====================
function requireLogin(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}
/* =========================
   🔥 IMPORTANT FIX END
========================= */

app.use(express.json({ limit: "10mb" }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use("/image", express.static(path.join(__dirname, "image")));

/* =========================
   GOOGLE LOGIN
========================= */

app.get("/auth/google", (req, res) => {

  const state = JSON.stringify({
    user: req.session.user || null
  });

  const url = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "profile",
      "email",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/spreadsheets.readonly"
    ],
    state: state
  });

  res.redirect(url);
});

app.get("/auth/google/callback", async (req, res) => {

  try {
    const code = req.query.code;
    const state = JSON.parse(req.query.state || "{}");

    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);

    if (state.user) {
      req.session.user = state.user;
    }

    req.session.tokens = tokens;

    req.session.save(() => {
      res.redirect("/dashboard");
    });

  } catch (err) {
    console.error("OAuth error:", err);
    res.redirect("/");
  }
});

/* =========================
   GOOGLE SHEETS (Service Account)
========================= */

const SPREADSHEET_ID = "1xAqS4dwT91fGVqTp2b3z6VWlXug28ilUHYVJ_tHe3QE";

const backendAuth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
});

/* =========================
   LOGIN API
========================= */

app.post("/api/login", async (req, res) => {

  const { username, password } = req.body;

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const userRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users!A2:C"
    });

    const users = userRes.data.values || [];

    const user = users.find(u =>
      u[0]?.trim() === username.trim() &&
      u[1]?.trim() === password.trim()
    );

    if (!user) {
      return res.json({ error: "Username หรือ Password ไม่ถูกต้อง" });
    }

    req.session.user = {
      username,
      role: user[2]
    };

    req.session.save(() => {
      res.json({ success: true, role: user[2] });
    });

  } catch (err) {
    console.error("Login error:", err);
    res.status(500).json({ error: "Login error" });
  }
});

/* =========================
   AUTH CHECK
========================= */

app.get("/api/check-auth", (req, res) => {

  if (!req.session.user) {
    return res.json({ loggedIn: false });
  }

  res.json({
    loggedIn: true,
    username: req.session.user.username,
    role: req.session.user.role
  });
});

/* =========================
   LOGOUT
========================= */

app.post("/api/logout", (req, res) => {
  req.session.destroy(() => {
    res.json({ success: true });
  });
});

// =====================
// GET STOCK
// =====================
// =====================
// GET STOCK
// =====================
app.get("/api/stock", requireLogin, async (req, res) => {
  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const master = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A2:I"
    });

    const office = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Office!A2:C"
    });

    const site = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C"
    });

    const masterData = master.data.values || [];
    const officeData = office.data.values || [];
    const siteData = site.data.values || [];

    const result = masterData.map(row => {
      const code = row[0];
      const officeRow = officeData.find(r => r[0] === code);
      const siteRow = siteData.find(r => r[0] === code);

      return {
        code,
        name: row[1],
        total: parseInt(row[5] || 0),
        office: officeRow ? parseInt(officeRow[2] || 0) : 0,
        site: siteRow ? parseInt(siteRow[2] || 0) : 0
      };
    });

    res.json(result);

  } catch (err) {
    console.error("Stock error:", err);
    res.status(500).json({ error: "Stock error" });
  }
});
// =====================
// UPDATE TOTAL
// =====================
app.post("/api/update-total", requireLogin, async (req, res) => {

  if (req.session.user.role !== "admin") {
    return res.json({ error: "ไม่มีสิทธิ์" });
  }

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const { code, newTotal } = req.body;
    const total = parseInt(newTotal);

    if (isNaN(total) || total < 0) {
      return res.json({ error: "จำนวนไม่ถูกต้อง" });
    }

    const masterRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A2:I"
    });

    const masterData = masterRes.data.values || [];
    const index = masterData.findIndex(r => r[0] === code);

    if (index === -1) {
      return res.json({ error: "ไม่พบสินค้า" });
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Stock_Master!F${index + 2}`,
      valueInputOption: "RAW",
      requestBody: { values: [[total]] }
    });

    res.json({ success: true });

  } catch (err) {
    console.error("Update total error:", err);
    res.status(500).json({ error: "Update error" });
  }
});

// =====================
// ADD STOCK (เพิ่มจำนวน)
// =====================
app.post("/api/add-stock", requireLogin, async (req, res) => {

  if (req.session.user.role !== "admin") {
    return res.json({ error: "ไม่มีสิทธิ์" });
  }

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const { code, qty } = req.body;
    const addQty = parseInt(qty);

    if (isNaN(addQty) || addQty <= 0) {
      return res.json({ error: "จำนวนไม่ถูกต้อง" });
    }

    const masterRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A2:I"
    });

    const masterData = masterRes.data.values || [];
    const index = masterData.findIndex(r => r[0] === code);

    if (index === -1) {
      return res.json({ error: "ไม่พบสินค้า" });
    }

    let currentTotal = parseInt(masterData[index][5] || 0);
    let newTotal = currentTotal + addQty;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Stock_Master!F${index + 2}`,
      valueInputOption: "RAW",
      requestBody: { values: [[newTotal]] }
    });

    res.json({ success: true });

  } catch (err) {
    console.error("Add stock error:", err);
    res.status(500).json({ error: "Add stock error" });
  }
});

// =====================
// TRANSFER
// =====================
app.post("/api/transfer", requireLogin, async (req, res) => {

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const { code, name, type } = req.body;
    const qty = parseInt(req.body.qty);
    const user = req.session.user.username;

    if (!qty || qty <= 0) {
      return res.json({ error: "จำนวนไม่ถูกต้อง" });
    }

    const officeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Office!A2:C"
    });

    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C"
    });

    const officeData = officeRes.data.values || [];
    const siteData = siteRes.data.values || [];

    let officeIndex = officeData.findIndex(r => r[0] === code);
    let siteIndex = siteData.findIndex(r => r[0] === code);

    if (officeIndex === -1 || siteIndex === -1) {
      return res.json({ error: "ไม่พบข้อมูลสินค้าใน stock" });
    }

    let officeQty = parseInt(officeData[officeIndex][2] || 0);
    let siteQty = parseInt(siteData[siteIndex][2] || 0);

    if (type === "เบิก") {
      if (officeQty < qty) return res.json({ error: "Office ไม่พอ" });
      officeQty -= qty;
      siteQty += qty;
    }

    if (type === "คืน") {
      if (siteQty < qty) return res.json({ error: "Site ไม่พอ" });
      siteQty -= qty;
      officeQty += qty;
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Stock_Office!C${officeIndex + 2}`,
      valueInputOption: "RAW",
      requestBody: { values: [[officeQty]] }
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Stock_Site!C${siteIndex + 2}`,
      valueInputOption: "RAW",
      requestBody: { values: [[siteQty]] }
    });

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A:H",
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          new Date().toLocaleString("th-TH"),
          code,
          name,
          qty,
          type,
          type === "เบิก" ? "Office" : "Site",
          type === "เบิก" ? "Site" : "Office",
          user
        ]]
      }
    });

    res.json({ success: true });

  } catch (err) {
    console.error("Transfer error:", err);
    res.status(500).json({ error: "Transfer error" });
  }
});

// =====================
// HISTORY
// =====================
app.get("/api/history", requireLogin, async (req, res) => {

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const log = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A2:H"
    });

    const logData = log.data.values || [];

    const history = logData.reverse().map(row => ({
      date: row[0],
      code: row[1],
      name: row[2],
      qty: row[3],
      type: row[4],
      from: row[5],
      to: row[6],
      user: row[7]
    }));

    res.json(history);

  } catch (err) {
    console.error("History error:", err);
    res.status(500).json({ error: "History error" });
  }
});
// =====================
// DASHBOARD
// =====================
app.get("/api/dashboard", requireLogin, async (req, res) => {

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const master = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A2:I"
    });

    const office = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Office!A2:C"
    });

    const site = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C"
    });

    const log = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A2:H"
    });

    const masterData = master.data.values || [];
    const officeData = office.data.values || [];
    const siteData = site.data.values || [];
    const logData = log.data.values || [];

    const totalItems = masterData.length;

    let totalOffice = officeData.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);
    let totalSite = siteData.reduce((sum, r) => sum + parseInt(r[2] || 0), 0);

    const today = new Date().toLocaleDateString("th-TH");

    let todayBorrow = 0;
    let todayReturn = 0;

    logData.forEach(row => {
      if (row[0] && row[0].includes(today)) {
        if (row[4] === "เบิก") todayBorrow += parseInt(row[3] || 0);
        if (row[4] === "คืน") todayReturn += parseInt(row[3] || 0);
      }
    });

    res.json({
      totalItems,
      totalOffice,
      totalSite,
      todayBorrow,
      todayReturn
    });

  } catch (err) {
    console.error("Dashboard error:", err);
    res.status(500).json({ error: "Dashboard error" });
  }
});
app.post("/api/add-item", requireLogin, async (req, res) => {

  if (req.session.user.role !== "admin") {
    return res.json({ error: "ไม่มีสิทธิ์" });
  }

  const { code, name, total, office, site, ext } = req.body;

  try {

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version: "v4",
      auth: authClient
    });

    const imageUrl = `https://cdn.jsdelivr.net/gh/Khemachat2003/stock-image/images/${code}.${ext}`;

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Master!A:F",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          code,
          name,
          `=IMAGE("${imageUrl}")`,
          "",
          "",
          total
        ]]
      }
    });

    res.json({ success: true });

  } catch (error) {

    console.error("Add item error:", error);
    res.json({ success: false, error: error.message });

  }
});
app.post("/upload-image", async (req, res) => {

    const { fileName, base64 } = req.body;

    if (!fileName || !base64) {
        return res.json({ success: false, message: "Missing data" });
    }

    const content = base64.replace(/^data:image\/\w+;base64,/, "");
    const filePath = `images/${fileName}`;

    try {

        let sha = null;

        // เช็คก่อนว่าไฟล์มีอยู่ไหม
        try {
            const existingFile = await octokit.repos.getContent({
                owner: "Khemachat2003",
                repo: "stock-image",
                path: filePath
            });

            sha = existingFile.data.sha;

        } catch (err) {
            // ถ้า 404 = ยังไม่มีไฟล์
        }

        // สร้างหรืออัปเดตไฟล์
        await octokit.repos.createOrUpdateFileContents({
            owner: "Khemachat2003",
            repo: "stock-image",
            path: filePath,
            message: "upload image",
            content: content,
            sha: sha
        });

        res.json({ success: true });

    } catch (error) {

        console.log("GitHub error:", error);
        res.json({ success: false, error: error.message });

    }

});
const fs = require("fs");
const PDFDocument = require("pdfkit");

function formatDate(dateStr){
  if(!dateStr) return "-";

  const d = new Date(dateStr);
  const day = String(d.getDate()).padStart(2,'0');
  const month = String(d.getMonth()+1).padStart(2,'0');
  const year = d.getFullYear();

  return `${day}/${month}/${year}`;
}

app.post("/api/export-history", requireLogin, async (req, res) => {

  if (!req.session.tokens) {
    return res.status(401).json({ error: "กรุณา Login Google ก่อน Export" });
  }

  try {

    const title = req.body.title || "รายงานประวัติการเบิก–คืนอุปกรณ์";
    const locations = req.body.locations || "-";
    const vehicle = req.body.vehicle || "-";
    const startDate = req.body.startDate || null;
    const endDate = req.body.endDate || null;
    const employeeCount = req.body.employeeCount || "0";
const employees = req.body.employees || "";
const reportType = req.body.reportType || "all";

    const authClient = await backendAuth.getClient();
    const sheets = google.sheets({ version: "v4", auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Transfer_Log!A:H"
    });

    const rows = response.data.values || [];
    const dataRows = rows.slice(1);

    let filteredData = dataRows;

    if (reportType === "borrow") {
  filteredData = filteredData.filter(row => row[4] === "เบิก");
}

if (reportType === "return") {
  filteredData = filteredData.filter(row => row[4] === "คืน");
}

    if (startDate && endDate) {

      const [startY, startM, startD] = startDate.split("-");
      const [endY, endM, endD] = endDate.split("-");

      const start = new Date(startY, startM - 1, startD, 0, 0, 0);
      const end = new Date(endY, endM - 1, endD, 23, 59, 59);

      filteredData = dataRows.filter(row => {

        if (!row[0]) return false;

        const [datePart] = row[0].split(" ");
        const [day, month, buddhistYear] = datePart.split("/");
        const year = parseInt(buddhistYear) - 543;

        const rowDate = new Date(year, month - 1, day, 12, 0, 0);

        return rowDate >= start && rowDate <= end;
      });
    }

    const filteredRows = filteredData.map(row => [
      row[0] || "-",
      row[2] || "-",
      row[3] || "-",
      row[4] || "-",
      row[7] || "-"
    ]);

    if (filteredRows.length === 0) {
      return res.status(400).json({ error: "ไม่มีข้อมูลสำหรับ Export" });
    }

    const fileName = `history_${Date.now()}.pdf`;

    const doc = new PDFDocument({
      size: "A4",
      margins: { top: 100, bottom: 60, left: 60, right: 60 }
    });

    const stream = fs.createWriteStream(fileName);
    doc.pipe(stream);

    doc.registerFont("THSarabun", "./fonts/THSarabunNew.ttf");
    doc.font("THSarabun").fontSize(14);

    const marginLeft = doc.page.margins.left;
    const marginRight = doc.page.margins.right;
    const contentWidth = doc.page.width - marginLeft - marginRight;

    // ================= HEADER =================
    function drawReportHeader() {

      doc.image("./logo.png", marginLeft, 40, { width: 70 });

      doc.y = 45;

      doc.fontSize(20)
         .text("บริษัท อินทนิล ออโตเมชั่น จำกัด", {
           align: "center",
           width: contentWidth
         });

      doc.fontSize(16)
         .text(title, {
           align: "center",
           width: contentWidth
         });

      doc.moveDown(0.5);

      doc.moveTo(marginLeft, doc.y)
         .lineTo(doc.page.width - marginRight, doc.y)
         .stroke();

      doc.moveDown(1);
    }

    drawReportHeader();

    doc.on("pageAdded", () => {
      drawReportHeader();
    });

    // ================= ข้อมูลก่อนตาราง =================
    doc.fontSize(14);
doc.x = marginLeft;

/* แสดงข้อมูลพวกสถานที่/รถ/พนักงาน เฉพาะตอนที่ไม่ใช่รายการคืน */
if (reportType !== "return") {

  if (locations) {

    const locationList = locations.split("\n").filter(l => l.trim() !== "");

    doc.text(`สถานที่: ${locationList.length} ที่`, { width: contentWidth });

    doc.moveDown(0);

    locationList.forEach((loc, index) => {
      const clean = loc.replace(/^\d+\.\s*/, "");
      doc.text(`${index + 1}. ${clean}`, {
        width: contentWidth,
        indent: 20
      });
    });

  }

  doc.moveDown(0);
  doc.text(`ยานพาหนะ: ${vehicle}`, { width: contentWidth });

  doc.text(`จำนวนพนักงาน: ${employeeCount} คน`, { width: contentWidth });

  if (employees) {
    doc.moveDown(0);
    employees.split("\n").forEach((name, index) => {
      const clean = name.replace(/^\d+\.\s*/, "");
      doc.text(`${index + 1}. ${clean}`, {
        width: contentWidth,
        indent: 20
      });
    });
  }

}

/* ช่วงวันที่ต้องแสดงทุกแบบ */
doc.text(
  `ช่วงวันที่: ${formatDate(startDate)} - ${formatDate(endDate)}`,
  { width: contentWidth }
);

doc.moveDown(0);
doc.text("วันที่ออกรายงาน: " + new Date().toLocaleString("th-TH"), {
  width: contentWidth
});

doc.moveDown(0);

// แปลงค่าประเภทรายงาน
let reportTypeText = "รวมการเบิกและคืน";

if (reportType === "borrow") reportTypeText = "รายการเบิก";
if (reportType === "return") reportTypeText = "รายการคืน";

// แสดงใน PDF
doc.fontSize(16);
doc.text(`ตาราง: ${reportTypeText}`, {
  width: contentWidth,
  align: "center"
});

doc.moveDown(0.5);

// ================= ตาราง =================
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

  columns.forEach(col => {
    doc.rect(x, y, col.width, headerHeight)
       .fillAndStroke("#f2f2f2", "black");

    doc.fillColor("black")
       .text(col.header, x + 5, y + 7, {
         width: col.width - 10,
         align: "center"
       });

    x += col.width;
  });

  y += headerHeight;
}

drawTableHeader();

filteredRows.forEach(row => {

  // 🔥 คำนวณความสูงจริงของแต่ละ cell
  let maxHeight = 0;

  columns.forEach((col, i) => {

    const cellText = row[i] || "-";

    const textHeight = doc.heightOfString(cellText, {
      width: col.width - 10
    });

    if (textHeight > maxHeight) {
      maxHeight = textHeight;
    }
  });

  const rowHeight = maxHeight + 10; // padding บนล่าง

  // 🔥 เช็คก่อนวาดทั้งแถว
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

    doc.text(cellText, x + 5, y + 5, {
      width: col.width - 10,
      align: i === 2 ? "center" : "left"
    });

    x += col.width;
  });

  y += rowHeight;
});
// ================= ลายเซ็น =================

doc.moveDown(3);

const pageWidth = doc.page.width;
const marginLeftSign = doc.page.margins.left;
const marginRightSign = doc.page.margins.right;
const contentWidthSign = pageWidth - marginLeftSign - marginRightSign;

const leftX = marginLeftSign;
const rightX = marginLeftSign + contentWidthSign / 2;

const today = new Date().toLocaleDateString("th-TH");

// หัวข้อ
doc.text("ผู้ทำรายการ", leftX, doc.y, {
  width: contentWidthSign / 2,
  align: "center"
});

doc.text("ผู้ตรวจสอบ", rightX, doc.y - 14, {
  width: contentWidthSign / 2,
  align: "center"
});

doc.moveDown(2);

// เส้นเซ็น
doc.text("(....................................)", leftX, doc.y, {
  width: contentWidthSign / 2,
  align: "center"
});

doc.text("(....................................)", rightX, doc.y - 14, {
  width: contentWidthSign / 2,
  align: "center"
});

doc.moveDown(1);

// วันที่
doc.text(`วันที่ ${today}`, leftX, doc.y, {
  width: contentWidthSign / 2,
  align: "center"
});

doc.text(`วันที่ ${today}`, rightX, doc.y - 14, {
  width: contentWidthSign / 2,
  align: "center"
});
    doc.end();

    stream.on("finish", async () => {

      try {

        oAuth2Client.setCredentials(req.session.tokens);

        const drive = google.drive({
          version: "v3",
          auth: oAuth2Client
        });

        const driveResponse = await drive.files.create({
          requestBody: {
            name: fileName,
            mimeType: "application/pdf",
            parents: ["1xbSU_CSbMsq5xXOejRNMpGpRcnG2rb2o"]
          },
          media: {
            mimeType: "application/pdf",
            body: fs.createReadStream(fileName)
          },
          fields: "id, webViewLink"
        });

        fs.unlinkSync(fileName);

        res.json({
          success: true,
          link: driveResponse.data.webViewLink
        });

      } catch (error) {
        console.log("Drive Upload Error:", error);
        res.status(500).json({ error: "Upload Drive ล้มเหลว" });
      }

    });

  } catch (err) {
    console.log("Export Error:", err);
    res.status(500).json({ error: "Export ล้มเหลว" });
  }

});


app.get("/api/get-site-items", requireLogin, async (req,res)=>{

  try{

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version:"v4",
      auth:authClient
    });

    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Stock_Site!A2:C"
    });

    const rows = siteRes.data.values || [];

    const items = [];

    rows.forEach(r=>{
      const code = r[0];
      const name = r[1];
      const qty = parseInt(r[2] || 0);

      if(qty > 0){
        items.push({
          code,
          name,
          qty
        });
      }
    });

    res.json({items});

  }catch(err){

    console.error("get-site-items error:",err);
    res.status(500).json({items:[]});

  }

});
app.post("/api/return-all-site", requireLogin, async (req,res)=>{

  try{

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version:"v4",
      auth:authClient
    });

    const user = req.session.user.username;

    const officeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range:"Stock_Office!A2:C"
    });

    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range:"Stock_Site!A2:C"
    });

    const officeData = officeRes.data.values || [];
    const siteData = siteRes.data.values || [];

    const updates=[];
    const logs=[];

    for(let i=0;i<siteData.length;i++){

      const code = siteData[i][0];
      const name = siteData[i][1];
      let siteQty = parseInt(siteData[i][2] || 0);

      if(siteQty > 0){

        const officeIndex = officeData.findIndex(r=>r[0]===code);

        if(officeIndex === -1) continue;

        let officeQty = parseInt(officeData[officeIndex][2] || 0);

        officeQty += siteQty;

        /* เตรียม update */
        updates.push({
          range:`Stock_Office!C${officeIndex+2}`,
          values:[[officeQty]]
        });

        updates.push({
          range:`Stock_Site!C${i+2}`,
          values:[[0]]
        });

        /* log */
        logs.push([
          new Date().toLocaleString("th-TH"),
          code,
          name,
          siteQty,
          "คืน",
          "Site",
          "Office",
          user
        ]);

      }

    }

    /* update stock ทีเดียว */
    if(updates.length>0){

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId:SPREADSHEET_ID,
        requestBody:{
          valueInputOption:"RAW",
          data:updates
        }
      });

    }

    /* append log ทีเดียว */
    if(logs.length>0){

      await sheets.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_ID,
        range:"Transfer_Log!A:H",
        valueInputOption:"RAW",
        requestBody:{values:logs}
      });

    }

    res.json({success:true});

  }catch(err){

    console.error(err);
    res.status(500).json({success:false});

  }

});

app.post("/api/return-selected-site", requireLogin, async (req,res)=>{

  try{

    const authClient = await backendAuth.getClient();

    const sheets = google.sheets({
      version:"v4",
      auth:authClient
    });

    const items = req.body.items;
    const user = req.session.user.username;

    const officeRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range:"Stock_Office!A2:C"
    });

    const siteRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range:"Stock_Site!A2:C"
    });

    const officeData = officeRes.data.values || [];
    const siteData = siteRes.data.values || [];

    const updates=[];
    const logs=[];

    for(const item of items){

      const code=item.code;
      const qty=parseInt(item.qty);

      const officeIndex=officeData.findIndex(r=>r[0]===code);
      const siteIndex=siteData.findIndex(r=>r[0]===code);

      if(officeIndex===-1 || siteIndex===-1) continue;

      let officeQty=parseInt(officeData[officeIndex][2]||0);
      let siteQty=parseInt(siteData[siteIndex][2]||0);

      if(siteQty<qty) continue;

      officeQty+=qty;
      siteQty-=qty;

      updates.push({
        range:`Stock_Office!C${officeIndex+2}`,
        values:[[officeQty]]
      });

      updates.push({
        range:`Stock_Site!C${siteIndex+2}`,
        values:[[siteQty]]
      });

      logs.push([
        new Date().toLocaleString("th-TH"),
        code,
        officeData[officeIndex][1],
        qty,
        "คืน",
        "Site",
        "Office",
        user
      ]);

    }

    /* update stock ทีเดียว */
    if(updates.length>0){

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId:SPREADSHEET_ID,
        requestBody:{
          valueInputOption:"RAW",
          data:updates
        }
      });

    }

    /* append log ทีเดียว */
    if(logs.length>0){

      await sheets.spreadsheets.values.append({
        spreadsheetId:SPREADSHEET_ID,
        range:"Transfer_Log!A:H",
        valueInputOption:"RAW",
        requestBody:{values:logs}
      });

    }

    res.json({success:true});

  }catch(err){

    console.error(err);
    res.status(500).json({success:false});

  }

});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("Server running on port " + PORT);
});