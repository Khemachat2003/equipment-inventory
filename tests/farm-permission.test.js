// tests/farm-permission.test.js — สิทธิ์เพิ่มฟาร์ม/โรงเรือน + กันเพิ่มซ้ำ (409)
//
// พฤติกรรมใหม่ (2026-10-06): POST /api/add-farm-site และ /api/add-farm-house
// เปิดให้ผู้ใช้ที่ "ล็อกอินแล้วทุกคน" เพิ่มได้ (เดิม admin เท่านั้น → คนอื่นโดน 403)
// + ถ้า siteId / siteName / houseId (ในฟาร์มเดิม) ซ้ำ → 409 พร้อมข้อความไทย
//
// เทสนี้เรียก handler ของ routes/farm.js จริงทั้ง chain (requireLogin → validators → validate → handler)
// โดย monkey-patch services/sheets.getSheetsClient ให้เป็น fake sheets ก่อน require routes
// (node --test รันแต่ละไฟล์ใน process แยก → patch ไม่กระทบ test อื่น)
const { test, describe } = require("node:test");
const assert = require("node:assert");

// ── patch services ก่อน require routes/farm (CommonJS singleton) ──
const sheetsSvc = require("../services/sheets");
const auditSvc = require("../services/audit");

const appended = { site: [], house: [] };
const SITES_ROWS = [["SITE001", "ฟาร์มเดิม", "สัตว์ปีก", "", "", ""]];
const HOUSES_ROWS = [["SITE001-H01", "SITE001", "โรงเรือน 1", "", "", ""]];

function fakeSheets() {
  return {
    spreadsheets: {
      values: {
        get: async ({ range }) => {
          const r = String(range);
          if (r.startsWith("Farm_Sites")) return { data: { values: SITES_ROWS.map((row) => [...row]) } };
          if (r.startsWith("Farm_Houses")) return { data: { values: HOUSES_ROWS.map((row) => [...row]) } };
          return { data: { values: [] } };
        },
        append: async ({ range, requestBody }) => {
          const r = String(range);
          const v = requestBody.values[0];
          if (r.startsWith("Farm_Sites")) appended.site.push(v);
          else if (r.startsWith("Farm_Houses")) appended.house.push(v);
          return {};
        },
      },
    },
  };
}

auditSvc.logAudit = async () => {};            // ไม่ยิง audit จริงในเทส (logAudit เองก็ swallow error)
sheetsSvc.getSheetsClient = async () => fakeSheets(); // patch ก่อน require routes/farm

const farmRouter = require("../routes/farm");
const { findSiteDuplicate, findHouseDuplicate } = require("../services/farmGuard");

// ── helpers: เดิน route chain ตามลำดับ middleware จริง ──
function findRoute(path) {
  const layer = farmRouter.stack.find(
    (l) => l.route && l.route.path === path && l.route.methods && l.route.methods.post
  );
  assert.ok(layer, `ไม่พบ route ${path}`);
  return layer.route;
}

function mockRes() {
  return {
    statusCode: null,
    body: null,
    status(code) { this.statusCode = code; return this; },
    json(payload) { this.body = payload; return this; },
  };
}

// รัน middleware ตามลำดับ middleware จริง
// ⚠️ express-validator 7 เป็น async middleware: `async (req, _res, next) => { await runner.run(req); ...; next(); }`
// — มัน "คืน promise ที่ settle ก่อน chain จบ" (next ถูกเรียกภายใน) ดังนั้น:
//   • promise settle ของ handler "ตัวสุดท้าย" (controller) = จบ chain
//   • promise settle ของ validator = chain ยังวิ่งต่อผ่าน next() ของมันเอง — ห้าม finish
//   • sync middleware ที่ตอบ res แล้วไม่ next (เช่น requireLogin → 401) = จบ chain
function runRoute(route, req) {
  return new Promise((resolve) => {
    const res = mockRes();
    const handlers = route.stack.map((l) => l.handle);
    const total = handlers.length;
    let i = 0;
    let settled = false;
    const finish = () => { if (!settled) { settled = true; resolve({ res }); } };
    const responded = () => res.statusCode !== null || res.body !== null;
    const next = (err) => {
      if (err) { res.statusCode = res.statusCode ?? 500; return finish(); }
      if (i >= total) return finish();
      const h = handlers[i++];
      let advanced = false;
      const wrappedNext = (e) => { advanced = true; next(e); };
      try {
        const out = h(req, res, wrappedNext);
        if (out && typeof out.then === "function") {
          out.then(() => {
            if (i >= total && !advanced) return finish(); // controller ตัวสุดท้าย settle = จบ chain
            if (advanced) return;                          // validator: next ต่อแล้ว — ปล่อยให้ chain วิ่งเอง
            if (responded()) return finish();
            finish(); // กันค้าง: middleware จบเฉย ๆ โดยไม่ next ไม่ตอบ
          }, () => { res.statusCode = res.statusCode ?? 500; finish(); });
          return;
        }
        if (advanced) return; // sync next ต่อแล้ว — chain วิ่งต่อเอง
        let ticks = 0;
        const poll = () => {
          if (advanced) return;
          if (responded()) return finish();
          if (++ticks > 100) return finish();
          setImmediate(poll);
        };
        poll();
      } catch (e) {
        res.statusCode = res.statusCode ?? 500;
        finish();
      }
    };
    next();
  });
}

function mkReq(body, overrides = {}) {
  return {
    method: "POST",
    originalUrl: "/api/add-farm-site",
    url: "/api/add-farm-site",
    headers: {},
    body,
    session: { user: { username: "portal_user", role: "user" } },
    ip: "127.0.0.1",
    ...overrides,
  };
}

// ── tests ──
describe("สิทธิ์เพิ่มฟาร์ม/โรงเรือน — ล็อกอินแล้วเพิ่มได้ทุกคน", () => {
  test("user ทั่วไปล็อกอินแล้วเพิ่มฟาร์มได้ (ไม่โดน 403 เดิม)", async () => {
    appended.site.length = 0;
    const route = findRoute("/api/add-farm-site");
    const { res } = await runRoute(route, mkReq({
      siteId: "SITE002", siteName: "ฟาร์มใหม่", farmType: "สุกร",
    }));
    assert.strictEqual(res.statusCode, null);
    assert.strictEqual(res.body.success, true);
    assert.strictEqual(res.body.site.siteId, "SITE002");
    assert.strictEqual(appended.site.length, 1);
  });

  test("ยังไม่ล็อกอิน → 401", async () => {
    const route = findRoute("/api/add-farm-site");
    const { res } = await runRoute(route, mkReq({ siteId: "X" }, { session: null }));
    assert.strictEqual(res.statusCode, 401);
  });

  test("admin เพิ่มฟาร์มได้เหมือนเดิม", async () => {
    appended.site.length = 0;
    const route = findRoute("/api/add-farm-site");
    const { res } = await runRoute(route, mkReq(
      { siteId: "SITE003", siteName: "ฟาร์มแอดมิน", farmType: "สัตว์บก" },
      { session: { user: { username: "admin1", role: "admin" } } }
    ));
    assert.strictEqual(res.statusCode, null);
    assert.strictEqual(res.body.success, true);
    assert.strictEqual(appended.site.length, 1);
  });

  test("user ทั่วไปเพิ่มโรงเรือนได้ (ไม่โดน 403 เดิม)", async () => {
    appended.house.length = 0;
    const route = findRoute("/api/add-farm-house");
    const { res } = await runRoute(route, mkReq(
      { houseId: "SITE001-H02", siteId: "SITE001", houseName: "โรงเรือน 2" },
      { originalUrl: "/api/add-farm-house", url: "/api/add-farm-house" }
    ));
    assert.strictEqual(res.statusCode, null);
    assert.strictEqual(res.body.success, true);
    assert.strictEqual(res.body.house.houseId, "SITE001-H02");
    assert.strictEqual(appended.house.length, 1);
  });
});

describe("กันเพิ่มซ้ำ → 409", () => {
  test("siteId ซ้ำ (case-insensitive) → 409 + ข้อความไทย", async () => {
    const route = findRoute("/api/add-farm-site");
    const { res } = await runRoute(route, mkReq({
      siteId: "site001", siteName: "ฟาร์มชื่ออื่น",
    }));
    assert.strictEqual(res.statusCode, 409);
    assert.match(String(res.body.error), /มีในระบบแล้ว/);
  });

  test("siteName ซ้ำ → 409", async () => {
    const route = findRoute("/api/add-farm-site");
    const { res } = await runRoute(route, mkReq({
      siteId: "SITE009", siteName: "ฟาร์มเดิม",
    }));
    assert.strictEqual(res.statusCode, 409);
    assert.match(String(res.body.error), /มีในระบบแล้ว/);
  });

  test("houseId ซ้ำในฟาร์มเดิม → 409", async () => {
    const route = findRoute("/api/add-farm-house");
    const { res } = await runRoute(route, mkReq(
      { houseId: "SITE001-h01", siteId: "SITE001", houseName: "โรงเรือนซ้ำ" },
      { originalUrl: "/api/add-farm-house", url: "/api/add-farm-house" }
    ));
    assert.strictEqual(res.statusCode, 409);
    assert.match(String(res.body.error), /มีในฟาร์มนี้แล้ว/);
  });

  test("houseId เดิมแต่คนละฟาร์ม → เพิ่มได้", async () => {
    appended.house.length = 0;
    const route = findRoute("/api/add-farm-house");
    const { res } = await runRoute(route, mkReq(
      { houseId: "SITE001-H01", siteId: "SITE002", houseName: "โรงเรือน 1 ของ SITE002" },
      { originalUrl: "/api/add-farm-house", url: "/api/add-farm-house" }
    ));
    assert.strictEqual(res.statusCode, null);
    assert.strictEqual(res.body.success, true);
    assert.strictEqual(appended.house.length, 1);
  });
});

describe("farmGuard (pure helpers)", () => {
  test("ไม่ซ้ำ → null", () => {
    assert.strictEqual(findSiteDuplicate([{ siteId: "A", siteName: "x" }], { siteId: "B", siteName: "y" }), null);
    assert.strictEqual(findSiteDuplicate([], { siteId: "", siteName: "" }), null);
    assert.strictEqual(findHouseDuplicate([{ houseId: "A-H01", siteId: "A" }], { siteId: "B", houseId: "A-H01" }), null);
    assert.strictEqual(findHouseDuplicate([], { siteId: "A", houseId: "A-H01" }), null);
  });

  test("รหัส/ชื่อซ้ำ (trim + case-insensitive) → code 409", () => {
    assert.strictEqual(findSiteDuplicate([{ siteId: " a ", siteName: "X" }], { siteId: "A" }).code, 409);
    assert.strictEqual(findSiteDuplicate([{ siteId: "Z", siteName: " x " }], { siteId: "B", siteName: " x " }).code, 409);
    assert.strictEqual(findHouseDuplicate([{ houseId: "A-h01", siteId: "A" }], { siteId: "a", houseId: "A-H01" }).code, 409);
  });
});