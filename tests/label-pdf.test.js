// tests/label-pdf.test.js — integration test สำหรับ /api/label/pdf + /api/label/a4pdf
// ยิงเข้า express จำลอง (mount routes/label.js) ซึ่งเรียก generate_label.py จริง
// ถ้าเครื่องไม่มี python + deps (reportlab/barcode/PIL) → skip ทั้ง suite (ไม่ทำ suite fail)
// (บน Render ไม่ได้รัน npm test ใน build — test นี้สำหรับ dev เครื่องที่มี python)
const { test, describe, before, after } = require("node:test");
const assert = require("node:assert");
const express = require("express");
const http = require("http");
const { execFileSync } = require("child_process");

const labelRouter = require("../routes/label");

// เช็ค python + deps พร้อม (snippet เดียวกับ scripts/ensure_python_deps.js)
function pythonReady() {
  for (const py of ["python", "python3"]) {
    try {
      execFileSync(py, ["-c", "import reportlab, barcode, PIL"], {
        timeout: 60000,
        windowsHide: true,
      });
      return true;
    } catch (_) {
      /* ลองตัวถัดไป */
    }
  }
  return false;
}

const HAVE_PY = pythonReady();

let baseUrl = "";
let server = null;

function req(method, urlPath, jsonBody) {
  return new Promise((resolve, reject) => {
    const r = http.request(
      new URL(baseUrl + urlPath),
      { method, headers: jsonBody ? { "Content-Type": "application/json" } : {} },
      (res) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () =>
          resolve({
            status: res.statusCode,
            headers: res.headers,
            body: Buffer.concat(chunks),
          })
        );
      }
    );
    r.on("error", reject);
    if (jsonBody) r.write(JSON.stringify(jsonBody));
    r.end();
  });
}

function isPdf(buf) {
  return buf.subarray(0, 4).toString("latin1") === "%PDF";
}

describe(
  "label PDF endpoints (generate_label.py)",
  { skip: HAVE_PY ? false : "ไม่พบ python + reportlab/barcode/PIL บนเครื่องนี้" },
  () => {
    before(() =>
      new Promise((resolve) => {
        const app = express();
        app.use(labelRouter);
        server = app.listen(0, () => {
          baseUrl = `http://127.0.0.1:${server.address().port}`;
          resolve();
        });
      })
    );

    after(() => server && new Promise((resolve) => server.close(resolve)));

    test("GET /api/label/pdf → 200 application/pdf เริ่มด้วย magic %PDF", async () => {
      const r = await req("GET", "/api/label/pdf?serial=SMOKE-TEST-001");
      assert.equal(r.status, 200);
      assert.match(r.headers["content-type"], /application\/pdf/);
      assert.ok(isPdf(r.body), "ต้องได้ไบต์เริ่มด้วย %PDF");
    });

    test("POST /api/label/a4pdf → 200 application/pdf หลายดวง (มีชื่อไทย)", async () => {
      const r = await req("POST", "/api/label/a4pdf", {
        items: [
          { serial: "SMOKE-TEST-001", name: "อุปกรณ์ทดสอบ", status: "ใช้งานได้" },
          { serial: "SMOKE-TEST-002", name: "อุปกรณ์ทดสอบ", status: "ชำรุด" },
        ],
      });
      assert.equal(r.status, 200);
      assert.match(r.headers["content-type"], /application\/pdf/);
      assert.ok(isPdf(r.body), "ต้องได้ไบต์เริ่มด้วย %PDF");
      assert.ok(r.body.length > 500, "PDF หลายดวงต้องมีขนาดพอสมควร");
    });

    test("POST /api/label/a4pdf ไม่มี items → 400 No items provided", async () => {
      const r = await req("POST", "/api/label/a4pdf", { items: [] });
      assert.equal(r.status, 400);
      assert.deepEqual(JSON.parse(r.body.toString()), { error: "No items provided" });
    });
  }
);
