const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { addLocalPart, addLocalAsset, addLocalAssets, readLocalInventory } = require("../services/localInventory");

test("local development inventory supports adding a part and serial assets persistently", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "local-inventory-"));
  const filePath = path.join(dir, "inventory.json");
  assert.deepEqual(await addLocalPart({ partNumber: "SENS-1", partName: "Sensor" }, filePath), { success: true });
  assert.equal((await addLocalPart({ partNumber: "sens-1", partName: "Duplicate" }, filePath)).success, false);
  const result = await addLocalAssets({ partNumber: "SENS-1", partName: "Sensor", qty: 2, location: "Stock" }, "local-user", filePath);
  const next = await addLocalAssets({ partNumber: "SENS-1", partName: "Sensor", qty: 1, location: "Stock" }, "local-user", filePath);
  assert.equal(result.added, 2);
  assert.match(result.firstSerial, /^SN-SENS-1-\d{8}-0001$/);
  assert.match(next.firstSerial, /^SN-SENS-1-\d{8}-0003$/);
  const inventory = await readLocalInventory(filePath);
  assert.equal(inventory.assets.length, 3);
  assert.equal(inventory.assets[0].user, "local-user");
  assert.equal(inventory.parts[0].totalQty, 3);
  await fs.rm(dir, { recursive: true, force: true });
});

test("local inventory accepts a single asset with its supplied serial and prevents duplicates", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "local-single-asset-"));
  const filePath = path.join(dir, "inventory.json");
  const created = await addLocalAsset({ assetId: "A-1", partNumber: "SENS-1", code: "SENS-1", name: "Sensor", serialNumber: "SN-EXACT-1" }, "local-user", filePath);
  assert.equal(created.success, true);
  assert.equal(created.asset.serialNumber, "SN-EXACT-1");
  assert.equal((await addLocalAsset({ assetId: "A-2", partNumber: "SENS-1", serialNumber: "SN-EXACT-1" }, "local-user", filePath)).success, false);
  assert.equal((await readLocalInventory(filePath)).assets.length, 1);
  await fs.rm(dir, { recursive: true, force: true });
});

test("local inventory mode stays disabled in production", () => {
  const { isLocalInventoryMode } = require("../services/localInventory");
  const previous = { nodeEnv: process.env.NODE_ENV, localMode: process.env.INBOUND_LOCAL_MODE, google: process.env.GOOGLE_CREDENTIALS_ACCOUNT };
  process.env.NODE_ENV = "production";
  delete process.env.GOOGLE_CREDENTIALS_ACCOUNT;
  process.env.INBOUND_LOCAL_MODE = "1";
  assert.equal(isLocalInventoryMode(), false);
  if (previous.nodeEnv === undefined) delete process.env.NODE_ENV; else process.env.NODE_ENV = previous.nodeEnv;
  if (previous.localMode === undefined) delete process.env.INBOUND_LOCAL_MODE; else process.env.INBOUND_LOCAL_MODE = previous.localMode;
  if (previous.google === undefined) delete process.env.GOOGLE_CREDENTIALS_ACCOUNT; else process.env.GOOGLE_CREDENTIALS_ACCOUNT = previous.google;
});
