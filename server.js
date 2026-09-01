require("dotenv").config();
const express = require("express");
const { google } = require("googleapis");
const path = require("path");
const session = require("express-session");
const pgSession = require("connect-pg-simple");
const rateLimit = require("express-rate-limit");
const helmet = require("helmet");
const compression = require("compression");
const morgan = require("morgan");
const cors = require("cors");
const bodyParser = require("body-parser");
const fs = require("fs");
const { fullSystemBackup } = require('./services/backup');
// ========== IMPORT ROUTES ==========
const routes = require("./routes");
// ดึง asset router เพื่อเรียก sync (module cache จะใช้ตัวเดียวกัน)
const assetRouter = require("./routes/asset");
const { requireLogin, requireAdmin, validate } = require("./middleware/auth");
// ========== CACHE ==========
const { cache, getSheetsClient } = require("./services/sheets");

// -------------------- CONFIG --------------------
const PORT = process.env.PORT || 3000;
const CACHE_TTL = parseInt(process.env.CACHE_TTL) || 120; // 2 นาที (จาก 10)
const CACHE_TTL_HISTORY = parseInt(process.env.CACHE_TTL_HISTORY) || 60; // 1 นาที

// ========== ENV VALIDATION ==========
const requiredEnv = ["SPREADSHEET_ID", "GOOGLE_CREDENTIALS_ACCOUNT", "SESSION_SECRET"];
requiredEnv.forEach(key => {
  if (!process.env[key]) {
    console.error(`❌ Missing critical environment variable: ${key}`);
    process.exit(1);
  }
});
console.log("✅ All required environment variables are set.");

// ========== RATE LIMIT ==========
const limiter = rateLimit({
  windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000,
  max: parseInt(process.env.RATE_LIMIT_MAX) || 500,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many requests, please try again later." },
});

// ========== GOOGLE OAUTH ==========
let credentials = null;
if (process.env.GOOGLE_OAUTH) {
  try {
    credentials = JSON.parse(process.env.GOOGLE_OAUTH);
  } catch (e) {
    console.error("❌ Invalid GOOGLE_OAUTH JSON:", e.message);
  }
}
if (!credentials || !credentials.web) {
  throw new Error("❌ Invalid GOOGLE_OAUTH JSON format");
}
const { client_id, client_secret, redirect_uris } = credentials.web;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// ========== EXPRESS APP ==========
const app = express();

app.set("trust proxy", 1);

// ========== SECURITY & PERFORMANCE MIDDLEWARE ==========
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", "'unsafe-eval'", "https://cdn.jsdelivr.net", "https://unpkg.com"],
      scriptSrcAttr: ["'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com", "https://cdnjs.cloudflare.com", "https://cdn.jsdelivr.net"],
      styleSrcElem: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com", "https://cdnjs.cloudflare.com", "https://cdn.jsdelivr.net"],
      fontSrc: ["'self'", "https://fonts.gstatic.com", "https://cdnjs.cloudflare.com"],
      imgSrc: ["'self'", "data:", "https://cdn.jsdelivr.net"],
      connectSrc: ["'self'", "https://api.github.com", "https://cdn.jsdelivr.net"],
    },
  },
}));
app.use(compression());
app.use(morgan("combined"));

app.use(
  cors({
    origin: true,
    credentials: true,
  })
);

// ========== SESSION STORE ==========
// ใช้ PostgreSQL เป็นที่เก็บ session (ถ้ามี DATABASE_URL) - session จะไม่หลุด
// ทุกครั้งที่ Render restart สลับ instance และ locks ใช้ร่วมกันได้ทั้งระบบ
const DATABASE_URL = process.env.DATABASE_URL;
const sessionConfig = {
  name: "borrow-session",
  secret: process.env.SESSION_SECRET || "default-insecure-secret-change-me",
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    maxAge: 24 * 60 * 60 * 1000,
  },
};

if (DATABASE_URL) {
  const PGStore = pgSession(session);
  // ⚠️ connect-pg-simple v10 อ่านได้เฉพาะ conString/conObject ไม่ส่ง ssl ตรงๆ
  // Render ตั้ง sslmode=require ใน URL → pg ใหม่ตีความเป็น verify-full (ตรวจ cert)
  // ซึ่ง cert ของ Render เป็น self-signed → ต้อง force ssl + rejectUnauthorized:false
  let storeUrl = DATABASE_URL;
  try {
    const u = new URL(storeUrl);
    u.searchParams.delete('sslmode');
    storeUrl = u.toString();
  } catch (e) { /* keep as-is */ }
  sessionConfig.store = new PGStore({
    conObject: {
      connectionString: storeUrl,
      ssl: { rejectUnauthorized: false },
      max: 10,
    },
    createTableIfMissing: true,
    pruneSessionInterval: 60 * 60, // ล้าง session ที่หมดอายุชั่วโมงละครั้ง
  });
  console.log("✅ Session store: PostgreSQL");
} else {
  console.warn("⚠️ DATABASE_URL ไม่อยู่ - ใช้ MemoryStore (session จะหลุดเมื่อ restart)");
}

app.use(session(sessionConfig));

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(bodyParser.json());

// ========== STATIC FILE CACHING ==========
// ให้ Browser Cache ไฟล์ CSS, JS, รูป เป็นเวลา 1 วัน
app.use((req, res, next) => {
  if (req.url.match(/\.(css|js|png|jpg|jpeg|svg|ico|webp)$/)) {
    res.setHeader('Cache-Control', 'public, max-age=86400'); // 1 วัน
  }
  next();
});
// ===========================================

app.use(express.static("public"));
app.use("/image", express.static(path.join(__dirname, "image")));

// ========== RATE LIMIT ==========
app.use("/api", limiter);

// ========== ADMIN SESSION LOCK (ต่ออายุอัตโนมัติทุก request) ==========
const { createCheckAdminSessionLock } = require("./middleware/auth");
app.use("/api", createCheckAdminSessionLock(cache));

// ========== GOOGLE LOGIN ==========
app.get("/auth/google", (req, res) => {
  const state = JSON.stringify({ user: req.session.user || null });
  const url = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "profile",
      "email",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/spreadsheets.readonly",
    ],
    state,
  });
  res.redirect(url);
});

// server.js (เพิ่มตรงส่วน Routes)
const { backupToPostgres } = require('./services/backup');

// POST /api/backup - Backup ข้อมูล (Admin only)
app.post('/api/backup', requireLogin, requireAdmin, async (req, res) => {
  try {
    const result = await backupToPostgres();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// POST /api/full-backup - Full System Backup (Admin only)
app.post('/api/full-backup', requireLogin, requireAdmin, async (req, res) => {
  try {
    const result = await fullSystemBackup();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// server.js (เพิ่มตรงส่วน Routes)
const { createClient } = require('./services/pg');

// whitelist ตารางที่อนุญาตให้ดูผ่าน backup-View (กัน SQL injection)
const BACKUP_TABLES = [
  'backup_stock_master',
  'backup_users',
  'backup_transfer_log',
  'backup_audit_log',
  'backup_stock_office',
  'backup_stock_site',
  'backup_farm_sites',
  'backup_farm_houses',
  'backup_part_catalog',
  'backup_asset_list',
  'backup_asset_history',
  'backup_damaged_assets',
];

app.get('/api/backup-data', requireLogin, requireAdmin, async (req, res) => {
  const { table } = req.query;
  if (!table) {
    return res.status(400).json({ error: 'table parameter required' });
  }
  // ✅ ตรวจว่าชื่อตารางอยู่ใน whitelist ก่อน (ไม่อนุญาตให้ใส่ SQL อื่น)
  if (!BACKUP_TABLES.includes(table)) {
    return res.status(400).json({ error: 'invalid table name' });
  }

  const DATABASE_URL = process.env.DATABASE_URL;
  if (!DATABASE_URL) {
    return res.status(500).json({ error: 'DATABASE_URL not set' });
  }

  const client = createClient();

  try {
    await client.connect();
    const result = await client.query(`SELECT * FROM ${table} ORDER BY backup_date DESC LIMIT 1000`);
    await client.end();
    res.json({ success: true, rows: result.rows });
  } catch (err) {
    await client.end();
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get("/auth/google/callback", async (req, res) => {
  try {
    const code = req.query.code;
    const state = JSON.parse(req.query.state || "{}");
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    if (state.user) req.session.user = state.user;
    req.session.tokens = tokens;
    req.session.save(() => res.redirect("/dashboard"));
  } catch (err) {
    console.error("OAuth error:", err);
    res.redirect("/");
  }
});

app.get("/audit.html", (req, res) => {
  if (!req.session.user || req.session.user.role !== "admin") {
    return res.redirect("/");
  }
  res.sendFile(path.join(__dirname, "public/audit.html"));
});

// ========== USE ROUTES ==========
app.use("/", routes);

// ========== FRONTEND ROUTES ==========
app.get("/dashboard", (req, res) => {
  if (!req.session.user) return res.redirect("/");
  res.sendFile(path.join(__dirname, "public/index.html"));
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/index.html"));
});

app.get("/a/:code", (req, res) => {
  res.sendFile(path.join(__dirname, "public/asset.html"));
});

app.use(express.static("public"));

app.get("/health", (req, res) => {
  res.status(200).json({ status: "ok", uptime: process.uptime() });
});

// ========== CLEAR CACHE API ==========
app.post("/api/clear-cache", requireLogin, async (req, res) => {
  if (req.session.user.role !== "admin") {
    return res.status(403).json({ error: "Unauthorized" });
  }
  cache.flushAll();
  res.json({ success: true, message: "Cache cleared" });
});

// ========== ERROR HANDLER (ท้ายสุด) ==========
app.use((err, req, res, next) => {
  console.error("❌ Unhandled error:", err);
  res.status(500).json({ error: "Internal server error", message: err.message });
});

// ========== START SERVER ==========
app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
  // เรียก sync asset history
  assetRouter.syncInitialAssetHistory().catch(console.error);
});

// ========== SCHEDULED BACKUP ==========
// รัน Full System Backup อัตโนมัติตาม BACKUP_CRON
// แบบ "smart": จะ backup ก็ต่อเมื่อมี Audit Log ใหม่เกิดขึ้น หลัง backup ครั้งล่าสุด
// (กรณีไม่มีการเปลี่ยนแปลงข้อมูลจริงในวันนั้น จะข้ามไป ไม่สิ้นเปลือง quota)
// ปิดได้ด้วย env BACKUP_CRON=""
const cron = require("node-cron");
const BACKUP_CRON = process.env.BACKUP_CRON;
if (BACKUP_CRON) {
  try {
    cron.schedule(BACKUP_CRON, async () => {
      console.log("🔄 Running scheduled backup check...");
      if (!process.env.DATABASE_URL) {
        console.warn("⚠️ DATABASE_URL ไม่อยู่ - ข้าม scheduled backup");
        return;
      }
      try {
        const changed = await hasNewAuditSinceLastBackup();
        if (!changed) {
          console.log("ℹ️ ไม่พบ Audit Log ใหม่หลัง backup ล่าสุด — ข้าม backup วันนี้");
          return;
        }
        const result = await fullSystemBackup();
        console.log("📊 Scheduled backup result:", result.totalRows ? `OK (${result.totalRows} rows)` : result);
      } catch (err) {
        console.error("❌ Scheduled backup failed:", err.message);
      }
    });
    console.log(`✅ Scheduled backup (smart): ${BACKUP_CRON}`);
  } catch (e) {
    console.error("❌ Invalid BACKUP_CRON:", e.message);
  }
} else {
  console.log("ℹ️ Scheduled backup disabled (set BACKUP_CRON env เช่น '0 19 * * *')");
}

// ── ตรวจว่า Audit_Log มีแถวใหม่กว่า backup ครั้งล่าสุดหรือไม่ ──
async function hasNewAuditSinceLastBackup() {
  try {
    const sheets = await getSheetsClient();
    const auditResp = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.SPREADSHEET_ID,
      range: "Audit_Log!A2:A",
    });
    const sheetRows = (auditResp.data.values || []).filter((r) => r[0]).length;

    const client = createClient();
    await client.connect();

    let backupRows = 0;
    try {
      const res = await client.query('SELECT count(*) AS cnt FROM backup_audit_log');
      backupRows = parseInt(res.rows[0].cnt, 10) || 0;
    } catch (e) {
      // ตาราง backup ยังไม่เคยถูกสร้าง → ถือว่าต้อง backup
      console.warn("ℹ️ ยังไม่มี backup_audit_log (ครั้งแรก) — จะ backup");
    }
    await client.end();

    return sheetRows > backupRows;
  } catch (err) {
    console.error("❌ hasNewAuditSinceLastBackup error:", err.message);
    return true; // error → ให้ backup ไปก่อน ปลอดภัยกว่า
  }
}