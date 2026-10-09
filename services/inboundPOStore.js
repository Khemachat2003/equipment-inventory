const fs = require("node:fs/promises");
const path = require("node:path");
const { INBOUND_PO_STATUS_SHEET, ensureInboundPOStatusSheet } = require("./sheets");
const { isLocalInventoryMode } = require("./localInventory");

const LOCAL_PATH = process.env.INBOUND_PO_STATUS_PATH || path.join(process.cwd(), "data", "inbound-po-status.local.json");
const allowLocalFallback = () => process.env.NODE_ENV !== "production" || process.env.INBOUND_LOCAL_FALLBACK === "1";

async function readLocalStatuses(filePath = LOCAL_PATH) {
  try { const data = JSON.parse(await fs.readFile(filePath, "utf8")); return Array.isArray(data) ? data : []; }
  catch (error) { if (error.code === "ENOENT") return []; throw error; }
}

async function writeLocalStatuses(rows, filePath = LOCAL_PATH) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const temp = `${filePath}.${process.pid}.tmp`;
  await fs.writeFile(temp, `${JSON.stringify(rows, null, 2)}\n`, "utf8");
  await fs.rename(temp, filePath);
}

async function readPOStatuses(sheets, spreadsheetId) {
  const local = await readLocalStatuses();
  if (!sheets || isLocalInventoryMode()) return local;
  try {
    await ensureInboundPOStatusSheet();
    const { data } = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${INBOUND_PO_STATUS_SHEET}!A2:E` });
    const remote = (data.values || []).map((r) => ({ poNumber: r[0] || "", status: r[1] || "closed", closedAt: r[2] || "", closedBy: r[3] || "", updatedAt: r[4] || "" }));
    const map = new Map(local.map((item) => [item.poNumber, item]));
    remote.forEach((item) => map.set(item.poNumber, item));
    return [...map.values()];
  } catch (error) {
    if (!allowLocalFallback()) throw error;
    console.warn(`[inbound-po] status read fallback (${error.message})`);
    return local;
  }
}

async function savePOStatus(sheets, spreadsheetId, poNumber, status, username = "") {
  const now = new Date().toISOString();
  const record = { poNumber, status, closedAt: status === "closed" ? now : "", closedBy: status === "closed" ? username : "", updatedAt: now };
  const local = await readLocalStatuses();
  const updated = [...local.filter((row) => row.poNumber !== poNumber), record];
  if (isLocalInventoryMode() || !sheets) { await writeLocalStatuses(updated); return record; }
  try {
    await ensureInboundPOStatusSheet();
    const { data } = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${INBOUND_PO_STATUS_SHEET}!A2:A` });
    const index = (data.values || []).findIndex((row) => String(row[0] || "").trim() === poNumber);
    if (index >= 0) await sheets.spreadsheets.values.update({ spreadsheetId, range: `${INBOUND_PO_STATUS_SHEET}!A${index + 2}:E${index + 2}`, valueInputOption: "RAW", requestBody: { values: [[record.poNumber, record.status, record.closedAt, record.closedBy, record.updatedAt]] } });
    else await sheets.spreadsheets.values.append({ spreadsheetId, range: `${INBOUND_PO_STATUS_SHEET}!A:E`, valueInputOption: "RAW", requestBody: { values: [[record.poNumber, record.status, record.closedAt, record.closedBy, record.updatedAt]] } });
    return record;
  } catch (error) {
    if (!allowLocalFallback()) throw error;
    await writeLocalStatuses(updated);
    console.warn(`[inbound-po] status saved locally (${error.message})`);
    return record;
  }
}

module.exports = { LOCAL_PATH, readLocalStatuses, writeLocalStatuses, readPOStatuses, savePOStatus };
