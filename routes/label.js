// routes/label.js
const express = require("express");
const router = express.Router();
const path = require("path");
const fs = require("fs");
const os = require("os");
const { execFile } = require("child_process");

// ── Python runner ──────────────────────────────────────────────
// deps ถูกติดตั้งลง ./python_modules (path "ใน repo" — ดู render.yaml buildCommand
// และ scripts/ensure_python_deps.js) → ส่ง PYTHONPATH ให้ python import จากนั้นด้วย
const PY_MODULES = path.join(__dirname, '..', 'python_modules');

function pythonEnv() {
  const env = { ...process.env };
  if (fs.existsSync(PY_MODULES)) {
    env.PYTHONPATH = PY_MODULES + (env.PYTHONPATH ? path.delimiter + env.PYTHONPATH : '');
  }
  return env;
}

// บาง image (เช่น Render native runtime) มีแค่ python3 ไม่มี python
// → ลอง python ก่อน ถ้า ENOENT ค่อย fallback เป็น python3
function runPython(args, timeoutMs, cb) {
  const env = pythonEnv();
  const attempt = (cmd) => {
    execFile(cmd, args, { timeout: timeoutMs, env }, (err, stdout, stderr) => {
      if (err && err.code === 'ENOENT' && cmd === 'python') return attempt('python3');
      cb(err, stdout, stderr);
    });
  };
  attempt('python');
}

// รวม stderr + err.message เป็น detail เดียว (บางทั้ง error ไม่มี stderr เช่น ENOENT)
function detailOf(err, stderr) {
  return [stderr && stderr.trim(), err && err.message].filter(Boolean).join(' | ') || 'no output from python';
}

// -------------------- SINGLE LABEL PDF --------------------
router.get('/api/label/pdf', async (req, res) => {
  try {
    const {
      serial = 'SN-0000001',
      width = '50',
      height = '25',
      theme = 'light',
      fontSizeSerial = '12'
    } = req.query;

    const tmpOut = path.join(os.tmpdir(), `label_${Date.now()}.pdf`);
    const args = [
      path.join(__dirname, '..', 'generate_label.py'),
      '--mode', 'single',
      '--serial', String(serial),
      '--width', String(width),
      '--height', String(height),
      '--theme', String(theme),
      '--font-size-serial', String(fontSizeSerial),
      '--output', tmpOut,
    ];

    runPython(args, 15000, (err, stdout, stderr) => {
      if (err) {
        console.error('generate_label.py error:', detailOf(err, stderr));
        try { fs.unlinkSync(tmpOut); } catch(_) {}
        return res.status(500).json({ error: 'PDF generation failed', detail: detailOf(err, stderr) });
      }
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="label_${serial}.pdf"`);
      const stream = fs.createReadStream(tmpOut);
      stream.pipe(res);
      stream.on('end', () => { try { fs.unlinkSync(tmpOut); } catch(_) {} });
    });
  } catch (e) {
    res.status(500).json({ error: 'Server error', detail: e.message });
  }
});

// -------------------- A4 BULK PDF --------------------
router.post('/api/label/a4pdf', express.json({ limit: '2mb' }), async (req, res) => {
  try {
    const {
      items = [],
      labelW = 50,
      labelH = 25,
      marginTop = 10,
      marginBottom = 10,
      marginLeft = 10,
      marginRight = 10,
      gapX = 3,
      gapY = 3,
      theme = 'light',
      fontSizeSerial = 12
    } = req.body;

    if (!items.length) {
      return res.status(400).json({ error: 'No items provided' });
    }

    const tmpJson = path.join(os.tmpdir(), `labels_${Date.now()}.json`);
    const tmpOut = path.join(os.tmpdir(), `a4_${Date.now()}.pdf`);

    const data = {
      items,
      labelW, labelH,
      marginTop, marginBottom, marginLeft, marginRight,
      gapX, gapY,
      theme,
      fontSizeSerial
    };
    fs.writeFileSync(tmpJson, JSON.stringify(data));

    const args = [
      path.join(__dirname, '..', 'generate_label.py'),
      '--mode', 'a4',
      '--json-input', tmpJson,
      '--output', tmpOut,
    ];

    runPython(args, 30000, (err, stdout, stderr) => {
      try { fs.unlinkSync(tmpJson); } catch(_) {}
      if (err) {
        console.error('generate_label.py (a4) error:', detailOf(err, stderr));
        try { fs.unlinkSync(tmpOut); } catch(_) {}
        return res.status(500).json({ error: 'A4 PDF generation failed', detail: detailOf(err, stderr) });
      }
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="labels_a4.pdf"');
      const stream = fs.createReadStream(tmpOut);
      stream.pipe(res);
      stream.on('end', () => { try { fs.unlinkSync(tmpOut); } catch(_) {} });
    });
  } catch (e) {
    res.status(500).json({ error: 'Server error', detail: e.message });
  }
});

module.exports = router;