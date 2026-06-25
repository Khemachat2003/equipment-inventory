require("dotenv").config();
const express = require("express");
const { google } = require("googleapis");
const path = require("path");
const session = require("express-session");
const rateLimit = require("express-rate-limit");
const helmet = require("helmet");
const compression = require("compression");
const morgan = require("morgan");
const cors = require("cors");
const bodyParser = require("body-parser");
const fs = require("fs");

// ========== IMPORT ROUTES ==========
const routes = require("./routes");
// ดึง asset router เพื่อเรียก sync (module cache จะใช้ตัวเดียวกัน)
const assetRouter = require("./routes/asset");
const { requireLogin, requireAdmin, validate } = require("./middleware/auth");
// ========== CACHE ==========
const { cache } = require("./services/sheets");

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
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      styleSrcElem: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      imgSrc: ["'self'", "data:", "https://cdn.jsdelivr.net"],
      connectSrc: ["'self'", "https://api.github.com"],
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

app.use(
  session({
    name: "borrow-session",
    secret: process.env.SESSION_SECRET || "default-insecure-secret-change-me",
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 24 * 60 * 60 * 1000,
    },
  })
);

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