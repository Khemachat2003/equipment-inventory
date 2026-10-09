// services/sheets.js
const { google } = require("googleapis");
const NodeCache = require("node-cache");

const cache = new NodeCache({ stdTTL: 300, checkperiod: 60 });
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

let sheetsClient = null;
let serviceAccount = null;

if (process.env.GOOGLE_CREDENTIALS_ACCOUNT) {
  try {
    serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS_ACCOUNT);
    console.log("✅ Google credentials loaded (from service)");
  } catch (e) {
    console.error("❌ Invalid GOOGLE_CREDENTIALS_ACCOUNT JSON:", e.message);
  }
} else {
  console.error("❌ GOOGLE_CREDENTIALS_ACCOUNT env missing");
}

const backendAuth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});

async function getSheetsClient() {
  if (!sheetsClient) {
    const authClient = await backendAuth.getClient();
    sheetsClient = google.sheets({ version: "v4", auth: authClient });
  }
  return sheetsClient;
}

async function saveAssetHistory(
  serialNumber,
  action,
  fromLocation,
  toLocation,
  username,
  remark
) {
  const sheets = await getSheetsClient();
  const currentDate = new Date().toLocaleString("th-TH");
  const historyValues = [
    currentDate,
    serialNumber,
    action,
    fromLocation,
    toLocation,
    username,
    remark,
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Asset_History!A:G",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [historyValues] },
  });
}

function clearStockCache() {
  cache.del("stockData");
  cache.del("dashboardData");
  cache.del("dashboardStats");
  cache.del("siteItems");
}

function clearAssetCache() {
  cache.del("assetData");
  cache.del("assetLocationList");
  cache.del("assetBundleNames");
}

const DAMAGED_SHEET = "Damaged_Assets";
const DAMAGED_HEADER = [
  "Date", "SerialNumber", "AssetID", "Code", "Name", "PartNumber",
  "Status", "OldLocation", "OldSite", "TransferredBy", "Remark", "Action",
];

const INBOUND_SHEET = "Inbound_Log";
const INBOUND_HEADER = ["batchId", "receivedAt", "poNumber", "supplier", "serial", "assetId", "createdBy", "createdAt"];
const INBOUND_PO_STATUS_SHEET = "Inbound_PO_Status";
const INBOUND_PO_STATUS_HEADER = ["poNumber", "status", "closedAt", "closedBy", "updatedAt"];

async function ensureInboundPOStatusSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  if (!(meta.data.sheets || []).some((s) => s.properties?.title === INBOUND_PO_STATUS_SHEET)) {
    try {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: INBOUND_PO_STATUS_SHEET, gridProperties: { rowCount: 1000, columnCount: INBOUND_PO_STATUS_HEADER.length } } } }] } });
    } catch (err) {
      const refreshed = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
      if (!(refreshed.data.sheets || []).some((s) => s.properties?.title === INBOUND_PO_STATUS_SHEET)) throw err;
    }
  }
  const header = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${INBOUND_PO_STATUS_SHEET}!A1:E1` });
  if (!header.data.values?.length) await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${INBOUND_PO_STATUS_SHEET}!A1:E1`, valueInputOption: "RAW", requestBody: { values: [INBOUND_PO_STATUS_HEADER] } });
  return true;
}

async function ensureInboundLogSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  if ((meta.data.sheets || []).some((s) => s.properties?.title === INBOUND_SHEET)) return true;
  try {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: INBOUND_SHEET, gridProperties: { rowCount: 1000, columnCount: INBOUND_HEADER.length } } } }] } });
  } catch (err) {
    // Another request may have created it after the metadata read.
    const refreshed = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    if (!(refreshed.data.sheets || []).some((s) => s.properties?.title === INBOUND_SHEET)) throw err;
  }
  const header = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${INBOUND_SHEET}!A1:H1` });
  if (!header.data.values?.length) await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${INBOUND_SHEET}!A1:H1`, valueInputOption: "RAW", requestBody: { values: [INBOUND_HEADER] } });
  return true;
}

// ── ตรวจว่ามี sheet "Damaged_Assets" แล้วหรือยัง ถ้ายังไม่มีให้สร้าง + ตั้ง header ──
async function ensureDamagedAssetsSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = (meta.data.sheets || []).some(
    (s) => s.properties && s.properties.title === DAMAGED_SHEET
  );
  if (exists) return true;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: DAMAGED_SHEET,
              gridProperties: { rowCount: 1000, columnCount: DAMAGED_HEADER.length },
            },
          },
        },
      ],
    },
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${DAMAGED_SHEET}!A1:L1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [DAMAGED_HEADER] },
  });
  console.log("✅ Created sheet: Damaged_Assets");
  return true;
}

// ── บันทึกอุปกรณ์ที่ "ชำรุด/สูญหาย" หลังโอนย้ายลง sheet แยก ──
async function logDamagedAsset(record) {
  await ensureDamagedAssetsSheet();
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${DAMAGED_SHEET}!A:L`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        record.date || new Date().toLocaleString("th-TH"),
        record.serialNumber || "",
        record.assetId || "",
        record.code || "",
        record.name || "",
        record.partNumber || "",
        record.status || "",
        record.oldLocation || "",
        record.oldSite || "",
        record.user || "",
        record.remark || "",
        record.action || "",
      ]],
    },
  });
}

module.exports = {
  getSheetsClient,
  saveAssetHistory,
  clearStockCache,
  clearAssetCache,
  cache,
  SPREADSHEET_ID,
  DAMAGED_SHEET,
  ensureDamagedAssetsSheet,
  INBOUND_SHEET,
  ensureInboundLogSheet,
  INBOUND_PO_STATUS_SHEET,
  ensureInboundPOStatusSheet,
  logDamagedAsset,
};
