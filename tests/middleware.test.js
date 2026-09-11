// tests/middleware.test.js — unit tests สำหรับ middleware/auth.js (node:test, ไม่ต้องมี deps เพิ่ม)
const { test, describe } = require("node:test");
const assert = require("node:assert");
const {
  requireLogin,
  requireAdmin,
  validate,
  createCheckAdminSessionLock,
  ADMIN_SESSION_KEY,
} = require("../middleware/auth");

// ── fake req/res/next helpers ──
function mockRes() {
  return {
    statusCode: null,
    body: null,
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(payload) {
      this.body = payload;
      return this;
    },
  };
}

function run(mw, req) {
  const res = mockRes();
  let nextCalled = false;
  mw(req, res, () => {
    nextCalled = true;
  });
  return { res, nextCalled };
}

// ── fake cache (API เดียวกับ node-cache ที่ middleware ใช้) ──
function fakeCache() {
  const m = new Map();
  return {
    get: (k) => m.get(k),
    set: (k, v) => m.set(k, v),
    del: (k) => m.delete(k),
  };
}

describe("requireLogin", () => {
  test("ไม่มี session.user → 401 และไม่เรียก next", () => {
    const { res, nextCalled } = run(requireLogin, {});
    assert.strictEqual(res.statusCode, 401);
    assert.strictEqual(nextCalled, false);
  });

  test("มี session.user → ผ่าน (next)", () => {
    const { res, nextCalled } = run(requireLogin, {
      session: { user: { username: "a" } },
    });
    assert.strictEqual(nextCalled, true);
    assert.strictEqual(res.statusCode, null);
  });
});

describe("requireAdmin", () => {
  test("role ไม่ใช่ admin → 403", () => {
    const { res, nextCalled } = run(requireAdmin, {
      session: { user: { role: "user" } },
    });
    assert.strictEqual(res.statusCode, 403);
    assert.strictEqual(nextCalled, false);
  });

  test("admin → ผ่าน", () => {
    const { nextCalled } = run(requireAdmin, {
      session: { user: { role: "admin" } },
    });
    assert.strictEqual(nextCalled, true);
  });

  test("ไม่มี session → 403", () => {
    const { res } = run(requireAdmin, {});
    assert.strictEqual(res.statusCode, 403);
  });
});

describe("validate (express-validator)", () => {
  // express-validator v7 เก็บ context ไว้ที่ req["express-validator#contexts"]
  // (ดู lib/base.js: contextsKey = 'express-validator#contexts')
  // → inject context ตรงตาม key จริง เพื่อทดสอบ logic ของ validate middleware เอง
  const CONTEXTS_KEY = "express-validator#contexts";

  test("มี validation errors → 400 และไม่เรียก next", () => {
    const req = {
      [CONTEXTS_KEY]: [
        { errors: [{ type: "field", path: "username", msg: "Username required" }] },
      ],
    };
    const { res, nextCalled } = run(validate, req);
    assert.strictEqual(res.statusCode, 400);
    assert.strictEqual(nextCalled, false);
    assert.deepStrictEqual(res.body.errors, [
      { type: "field", path: "username", msg: "Username required" },
    ]);
  });

  test("ไม่มี errors → next (ผ่าน)", () => {
    const req = { [CONTEXTS_KEY]: [{ errors: [] }] };
    const { res, nextCalled } = run(validate, req);
    assert.strictEqual(nextCalled, true);
    assert.strictEqual(res.statusCode, null);
  });

  test("request ที่ไม่ผ่าน chain ใดๆ (ไม่มี context) → next (validate ไม่บังคับเอง)", () => {
    const { nextCalled } = run(validate, {});
    assert.strictEqual(nextCalled, true);
  });
});

describe("createCheckAdminSessionLock", () => {
  test("ยังไม่มี lock → สร้าง lock และผ่าน", () => {
    const cache = fakeCache();
    const mw = createCheckAdminSessionLock(cache);
    const { nextCalled } = run(mw, {
      session: { id: "s1", user: { role: "admin", username: "admin1" } },
    });
    assert.strictEqual(nextCalled, true);
    assert.strictEqual(cache.get(ADMIN_SESSION_KEY).sessionId, "s1");
  });

  test("admin คนเดิม (session เดิม) → ต่ออายุและผ่าน", () => {
    const cache = fakeCache();
    const mw = createCheckAdminSessionLock(cache);
    run(mw, { session: { id: "s1", user: { role: "admin", username: "admin1" } } });
    const { nextCalled } = run(mw, {
      session: { id: "s1", user: { role: "admin", username: "admin1" } },
    });
    assert.strictEqual(nextCalled, true);
    assert.strictEqual(cache.get(ADMIN_SESSION_KEY).username, "admin1");
  });

  test("admin คนอื่น (session ใหม่) → 403 บล็อก + lock ยังเป็นของคนเดิม", () => {
    const cache = fakeCache();
    const mw = createCheckAdminSessionLock(cache);
    run(mw, { session: { id: "s1", user: { role: "admin", username: "admin1" } } });
    const { res, nextCalled } = run(mw, {
      session: { id: "s2", user: { role: "admin", username: "admin2" } },
    });
    assert.strictEqual(res.statusCode, 403);
    assert.strictEqual(nextCalled, false);
    assert.strictEqual(cache.get(ADMIN_SESSION_KEY).sessionId, "s1");
  });

  test("user ทั่วไปไม่โดน lock (ผ่านเสมอ)", () => {
    const cache = fakeCache();
    const mw = createCheckAdminSessionLock(cache);
    run(mw, { session: { id: "s1", user: { role: "admin", username: "admin1" } } });
    const { nextCalled } = run(mw, {
      session: { id: "s9", user: { role: "user", username: "u1" } },
    });
    assert.strictEqual(nextCalled, true);
  });
});