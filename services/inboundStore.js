const fs = require("node:fs/promises");
const path = require("node:path");
const { INBOUND_SHEET, ensureInboundLogSheet } = require("./sheets");
const { isLocalInventoryMode } = require("./localInventory");

const LOCAL_PATH = process.env.INBOUND_LOCAL_PATH || path.join(process.cwd(), "data", "inbound-log.local.json");
const localFallbackAllowed = () => process.env.NODE_ENV !== "production" || process.env.INBOUND_LOCAL_FALLBACK === "1";

async function readLocalInboundRows(filePath = LOCAL_PATH) {
  try {
    const parsed = JSON.parse(await fs.readFile(filePath, "utf8"));
    return Array.isArray(parsed) ? parsed.filter(Array.isArray) : [];
  } catch (error) {
    if (error.code === "ENOENT") return [];
    throw error;
  }
}

async function writeLocalInboundRows(rows, filePath = LOCAL_PATH) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(rows, null, 2)}\n`, "utf8");
  await fs.rename(tempPath, filePath);
}

function mergeInboundRows(sheetRows = [], localRows = []) {
  const bySerial = new Map();
  for (const row of [...localRows, ...sheetRows]) {
    if (row[4]) bySerial.set(String(row[4]).trim(), row);
  }
  return [...bySerial.values()];
}

async function readInboundRows(sheets, spreadsheetId, filePath = LOCAL_PATH) {
  const localRows = await readLocalInboundRows(filePath);
  if (!sheets || isLocalInventoryMode()) return localRows;
  try {
    const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${INBOUND_SHEET}!A2:H` });
    return mergeInboundRows(response.data.values || [], localRows);
  } catch (error) {
    console.warn(`[inbound] Google Sheets read unavailable; using local metadata (${error.message})`);
    return localRows;
  }
}

async function prepareInboundWrite() {
  if (isLocalInventoryMode()) return "local-file";
  try {
    await ensureInboundLogSheet();
    return "google-sheets";
  } catch (error) {
    if (!localFallbackAllowed()) throw error;
    console.warn(`[inbound] Google Sheets unavailable; local development fallback enabled (${error.message})`);
    return "local-file";
  }
}

async function appendInboundRows(sheets, spreadsheetId, rows, filePath = LOCAL_PATH, ensureSheet = ensureInboundLogSheet) {
  if (isLocalInventoryMode()) {
    const localRows = await readLocalInboundRows(filePath);
    await writeLocalInboundRows(mergeInboundRows([], [...localRows, ...rows]), filePath);
    return "local-file";
  }
  try {
    await ensureSheet();
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${INBOUND_SHEET}!A:H`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });
    return "google-sheets";
  } catch (error) {
    if (!localFallbackAllowed()) throw error;
    const localRows = await readLocalInboundRows(filePath);
    await writeLocalInboundRows(mergeInboundRows([], [...localRows, ...rows]), filePath);
    console.warn(`[inbound] Google Sheets write unavailable; saved metadata locally (${error.message})`);
    return "local-file";
  }
}

module.exports = { LOCAL_PATH, readLocalInboundRows, writeLocalInboundRows, mergeInboundRows, readInboundRows, prepareInboundWrite, appendInboundRows };
