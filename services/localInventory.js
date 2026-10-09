const fs = require("node:fs/promises");
const path = require("node:path");

const LOCAL_INVENTORY_PATH = process.env.LOCAL_INVENTORY_PATH || path.join(process.cwd(), "data", "local-inventory.dev.json");

function isLocalInventoryMode() {
  return process.env.NODE_ENV !== "production" && (process.env.INBOUND_LOCAL_MODE === "1" || !process.env.GOOGLE_CREDENTIALS_ACCOUNT);
}

async function readLocalInventory(filePath = LOCAL_INVENTORY_PATH) {
  try {
    const data = JSON.parse(await fs.readFile(filePath, "utf8"));
    return { parts: Array.isArray(data.parts) ? data.parts : [], assets: Array.isArray(data.assets) ? data.assets : [] };
  } catch (error) {
    if (error.code === "ENOENT") return { parts: [], assets: [] };
    throw error;
  }
}

async function writeLocalInventory(inventory, filePath = LOCAL_INVENTORY_PATH) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(inventory, null, 2)}\n`, "utf8");
  await fs.rename(tempPath, filePath);
}

async function addLocalPart(part, filePath = LOCAL_INVENTORY_PATH) {
  const inventory = await readLocalInventory(filePath);
  const partNumber = String(part.partNumber || "").trim().toUpperCase();
  if (inventory.parts.some((item) => item.partNumber.toUpperCase() === partNumber)) return { success: false, error: `Part ${partNumber} มีอยู่แล้ว` };
  inventory.parts.push({ ...part, partNumber, partName: String(part.partName || partNumber).trim(), totalQty: 0, lastUpdated: new Date().toISOString() });
  await writeLocalInventory(inventory, filePath);
  return { success: true };
}

async function addLocalAssets(input, username, filePath = LOCAL_INVENTORY_PATH) {
  const inventory = await readLocalInventory(filePath);
  const qty = Number.parseInt(input.qty, 10);
  const partNumber = String(input.partNumber || "").trim().toUpperCase();
  const partName = String(input.partName || partNumber).trim();
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy = now.getFullYear();
  const dateStr = `${dd}${mm}${yyyy}`;
  const prefix = `SN-${partNumber}-${dateStr}`;
  const maxSerial = inventory.assets.reduce((max, asset) => {
    const serial = String(asset.serialNumber || "");
    if (!serial.startsWith(`${prefix}-`)) return max;
    const suffix = Number.parseInt(serial.slice(prefix.length + 1), 10);
    return Number.isFinite(suffix) ? Math.max(max, suffix) : max;
  }, 0);
  let nextAssetId = inventory.assets.reduce((max, asset) => Math.max(max, Number.parseInt(asset.assetId, 10) || 0), 0);
  const created = Array.from({ length: qty }, (_, index) => {
    const serialNumber = `${prefix}-${String(maxSerial + index + 1).padStart(4, "0")}`;
    nextAssetId += 1;
    return {
      assetId: String(nextAssetId).padStart(4, "0"), code: partNumber, name: partName, partNumber, serialNumber,
      status: input.status || "ใช้งานได้", location: input.location || "Stock", siteName: input.siteName || "Intranin",
      user: String(input.user || username || "local").trim(), farmType: input.farmType || "-", houseId: input.houseId || "-", houseName: input.houseName || "-",
      bundleId: "", bundleName: "", createdAt: now.toISOString(),
    };
  });
  inventory.assets.push(...created);
  const part = inventory.parts.find((item) => item.partNumber.toUpperCase() === partNumber);
  if (part) { part.totalQty = inventory.assets.filter((asset) => asset.partNumber === partNumber).length; part.lastUpdated = now.toISOString(); }
  await writeLocalInventory(inventory, filePath);
  return { assets: created, added: created.length, firstSerial: created[0]?.serialNumber, lastSerial: created.at(-1)?.serialNumber };
}

async function addLocalAsset(input, username, filePath = LOCAL_INVENTORY_PATH) {
  const inventory = await readLocalInventory(filePath);
  const assetId = String(input.assetId || "").trim();
  const serialNumber = String(input.serialNumber || "").trim();
  if (inventory.assets.some((asset) => asset.assetId === assetId || (serialNumber && asset.serialNumber === serialNumber))) {
    return { success: false, error: "Asset ID or Serial already exists" };
  }
  const now = new Date().toISOString();
  const partNumber = String(input.partNumber || input.code || "").trim().toUpperCase();
  const asset = {
    assetId, code: String(input.code || partNumber), name: String(input.name || ""), partNumber, serialNumber,
    status: input.status || "ใช้งานได้", location: input.location || "Stock", siteName: input.siteName || "Intranin",
    user: String(input.user || username || "local").trim(), farmType: "-", houseId: "-", houseName: "-",
    bundleId: "", bundleName: "", createdAt: now,
  };
  inventory.assets.push(asset);
  const part = inventory.parts.find((item) => item.partNumber.toUpperCase() === partNumber);
  if (part) { part.totalQty = inventory.assets.filter((item) => item.partNumber === partNumber).length; part.lastUpdated = now; }
  await writeLocalInventory(inventory, filePath);
  return { success: true, asset };
}

module.exports = { LOCAL_INVENTORY_PATH, isLocalInventoryMode, readLocalInventory, writeLocalInventory, addLocalPart, addLocalAssets, addLocalAsset };
