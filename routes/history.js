// routes/history.js
const express = require("express");
const router = express.Router();
const path = require("path");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const { body } = require("express-validator");
const os = require("os");
const { getSheetsClient, cache, SPREADSHEET_ID } = require("../services/sheets");
const { requireLogin, validate } = require("../middleware/auth");

const CACHE_TTL_HISTORY = 60;

// -------------------- GET HISTORY --------------------
router.get("/api/history", requireLogin, async (req, res) => {
  try {
    const { start, end } = req.query;
    const cacheKey = `history_${start || ""}_${end || ""}`;
    let cached = cache.get(cacheKey);
    if (cached) return res.json(cached);

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
    cache.set(cacheKey, history, CACHE_TTL_HISTORY);
    res.json(history);
  } catch (err) {
    console.error("History error:", err);
    res.status(500).json({ error: "History error" });
  }
});

// Helper: format date
function formatDate(dateStr) {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

// -------------------- EXPORT HISTORY PDF --------------------
router.post("/api/export-history",
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
      const doc = new PDFDocument({
        size: "A4",
        margins: { top: 100, bottom: 60, left: 60, right: 60 },
      });
      const filePath = path.join(os.tmpdir(), fileName);
      const stream = fs.createWriteStream(filePath);
      doc.pipe(stream);

      const fontPath = path.join(__dirname, "..", "fonts", "THSarabunNew.ttf");
      if (fs.existsSync(fontPath)) {
        doc.registerFont("THSarabun", fontPath);
        doc.font("THSarabun").fontSize(14);
      } else {
        doc.font("Helvetica").fontSize(12);
        console.warn("⚠️ Font THSarabunNew.ttf not found, using Helvetica");
      }

      const marginLeft = doc.page.margins.left;
      const marginRight = doc.page.margins.right;
      const contentWidth = doc.page.width - marginLeft - marginRight;

      function drawReportHeader() {
        const logoPath = path.join(__dirname, "..", "logo.png");
        if (fs.existsSync(logoPath)) {
          doc.image(logoPath, marginLeft, 40, { width: 70 });
        }
        doc.y = 45;
        doc.fontSize(20).text("บริษัท อินทนิล ออโตเมชั่น จำกัด", {
          align: "center",
          width: contentWidth,
        });
        doc.fontSize(16).text(title, {
          align: "center",
          width: contentWidth,
        });
        doc.moveDown(0.5);
        doc.moveTo(marginLeft, doc.y)
          .lineTo(doc.page.width - marginRight, doc.y)
          .stroke();
        doc.moveDown(1);
      }
      drawReportHeader();
      doc.on("pageAdded", drawReportHeader);

      doc.fontSize(14);
      doc.x = marginLeft;

      if (reportType === "all" || reportType === "borrow") {
        if (locations) {
          const locationList = locations
            .split("\n")
            .map((l) => l.trim())
            .filter((l) => l !== "");
          doc.text(`สถานที่: ${locationList.length} ที่`, { width: contentWidth });
          locationList.forEach((loc, index) => {
            const clean = loc.replace(/^\d+\.\s*/, "");
            doc.text(`${index + 1}. ${clean}`, {
              width: contentWidth,
              indent: 20,
            });
          });
        }
        doc.moveDown(0.5);
        doc.text(`ยานพาหนะ: ${vehicle}`, { width: contentWidth });
        doc.text(`จำนวนพนักงาน: ${employeeCount} คน`, { width: contentWidth });
        if (employees) {
          doc.moveDown(0.5);
          employees
            .split("\n")
            .map((n) => n.trim())
            .filter((n) => n !== "")
            .forEach((name, index) => {
              const clean = name.replace(/^\d+\.\s*/, "");
              doc.text(`${index + 1}. ${clean}`, {
                width: contentWidth,
                indent: 20,
              });
            });
        }
      }

      doc.moveDown(0.5);
      doc.text(`ช่วงวันที่: ${formatDate(startDate)} - ${formatDate(endDate)}`, {
        width: contentWidth,
      });
      doc.text("วันที่ออกรายงาน: " + new Date().toLocaleString("th-TH"), {
        width: contentWidth,
      });
      doc.moveDown(0);

      let reportTypeText = "รวมรายการเบิกและคืน";
      if (reportType === "borrow") reportTypeText = "รายการเบิก";
      if (reportType === "return") reportTypeText = "รายการคืน";
      doc.fontSize(16);
      doc.text(`ตาราง: ${reportTypeText}`, {
        width: contentWidth,
        align: "center",
      });
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
        doc.fontSize(14);
        const headerHeight = 25;
        if (y + headerHeight > usableHeight) {
          doc.addPage();
          y = doc.y;
        }
        let x = doc.page.margins.left;
        columns.forEach((col) => {
          doc.rect(x, y, col.width, headerHeight)
            .fillAndStroke("#f2f2f2", "black");
          doc.fillColor("black").text(col.header, x + 5, y + 7, {
            width: col.width - 10,
            align: "center",
          });
          x += col.width;
        });
        y += headerHeight;
      }
      drawTableHeader();

      filteredRows.forEach((row) => {
        let maxHeight = 0;
        columns.forEach((col, i) => {
          const cellText = row[i] || "-";
          const textHeight = doc.heightOfString(cellText, {
            width: col.width - 10,
          });
          if (textHeight > maxHeight) maxHeight = textHeight;
        });
        const rowHeight = maxHeight + 10;
        if (y + rowHeight > usableHeight) {
          doc.addPage();
          y = doc.y;
          drawTableHeader();
        }
        let x = doc.page.margins.left;
        columns.forEach((col, i) => {
          const cellText = row[i] || "-";
          doc.rect(x, y, col.width, rowHeight).stroke();
          doc.text(cellText, x + 5, y + 5, {
            width: col.width - 10,
            align: i === 2 ? "center" : "left",
          });
          x += col.width;
        });
        y += rowHeight;
      });

      doc.moveDown(3);
      const pageWidth = doc.page.width;
      const leftX = marginLeft;
      const rightX = marginLeft + (pageWidth - marginLeft - marginRight) / 2;
      const today = new Date().toLocaleDateString("th-TH");
      doc.text("ผู้ทำรายการ", leftX, doc.y, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });
      doc.text("ผู้ตรวจสอบ", rightX, doc.y - 14, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });
      doc.moveDown(2);
      doc.text("(....................................)", leftX, doc.y, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });
      doc.text("(....................................)", rightX, doc.y - 14, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });
      doc.moveDown(1);
      doc.text(`วันที่ ${today}`, leftX, doc.y, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });
      doc.text(`วันที่ ${today}`, rightX, doc.y - 14, {
        width: (pageWidth - marginLeft - marginRight) / 2,
        align: "center",
      });

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

module.exports = router;