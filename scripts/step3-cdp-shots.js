// Step 3/4 QA: TRUE device-viewport screenshots via Chrome DevTools Protocol
// (headless CLI clamps --window-size to ~500px on Windows → media queries wrong;
//  CDP Emulation.setDeviceMetricsOverride ให้ viewport จริง 320/375/1024)
// ใช้ WebSocket ในตัวของ Node >=22 ไม่ต้องพึ่ง puppeteer
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const CHROME = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const OUT = path.join(__dirname, '..', 'artifacts', 'step3-shots');
const PROFILE = path.join(__dirname, '..', 'artifacts', 'step3-cdp3');
const PORT = 9333;
const FARM = encodeURIComponent('ฟาร์มหนองแค');

const SHOTS = [
  { name: 'home-320.png', w: 320, h: 1500, mobile: true, url: 'http://localhost:5173/' },
  { name: 'home-375.png', w: 375, h: 1500, mobile: true, url: 'http://localhost:5173/' },
  { name: 'home-1024.png', w: 1024, h: 1000, mobile: false, url: 'http://localhost:5173/' },
  { name: 'home-farm-modal-1024.png', w: 1024, h: 1000, mobile: false, url: `http://localhost:5173/?farm=${FARM}` },
  { name: 'login-1024.png', w: 1024, h: 820, mobile: false, url: 'http://localhost:5173/login' },
];

function httpRequest(url, method = 'GET') {
  return new Promise((resolve, reject) => {
    const req = http.request(url, { method }, (res) => {
      let d = '';
      res.on('data', (c) => { d += c; });
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch (e) { reject(new Error(`${method} ${url} → ${d.slice(0, 80)}`)); } });
    });
    req.on('error', reject);
    req.end();
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function newTab(cdp) {
  // สร้าง target ใหม่ต่อรอบ เพื่อไม่ให้ state ค้างจากรอบก่อน
  const t = await httpRequest(`http://127.0.0.1:${PORT}/json/new?about:blank`);
  return t;
}

async function shoot(cdpWsFactory, { name, w, h, mobile, url }) {
  const target = await httpRequest(`http://127.0.0.1:${PORT}/json/new?about:blank`, 'PUT');
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
  let id = 0;
  const pending = new Map();
  ws.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
  };
  const send = (method, params = {}) => new Promise((res) => {
    const mid = ++id;
    pending.set(mid, res);
    ws.send(JSON.stringify({ id: mid, method, params }));
  });
  try {
    await send('Emulation.setDeviceMetricsOverride', { width: w, height: h, deviceScaleFactor: 2, mobile });
    await send('Page.enable');
    await send('Page.navigate', { url });
    // รอโหลด + ให้ axios/ฟอนต์เซ็ตตัว (virtual-time ใช้ไม่ได้กับ CDP ตรงๆ → รอจริง)
    await sleep(6000);
    const shot = await send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true });
    fs.writeFileSync(path.join(OUT, name), Buffer.from(shot.result.data, 'base64'));
    console.log(`${name}: OK`);
  } finally {
    try { ws.close(); } catch { /* ignore */ }
    // ปิดแท็บทิ้ง
    try { await httpRequest(`http://127.0.0.1:${PORT}/json/close/${target.id}`, 'PUT'); } catch { /* ignore */ }
  }
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const chrome = spawn(CHROME, [
    '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
    `--remote-debugging-port=${PORT}`, `--user-data-dir=${PROFILE}`,
    '--window-size=1100,1000', 'about:blank',
  ], { stdio: 'ignore' });
  try {
    // รอ debugger พร้อม
    let targets = null;
    for (let i = 0; i < 40 && !targets; i += 1) {
      await sleep(500);
      try { targets = await httpRequest(`http://127.0.0.1:${PORT}/json/version`); } catch { /* retry */ }
    }
    if (!targets) throw new Error('Chrome debugger ไม่พร้อม');
    for (const s of SHOTS) await shoot(null, s);
  } finally {
    try { chrome.kill(); } catch { /* ignore */ }
  }
})().catch((e) => { console.error('FAILED:', e.message); process.exit(1); });