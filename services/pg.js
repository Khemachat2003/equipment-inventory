// services/pg.js — PostgreSQL connection helper
// กันปัญหา self-signed certificate: ถ้า DATABASE_URL มี ?sslmode=... พวก node-postgres
// จะแปลงเป็น verify-full แล้ว fail. เราจึงตัด sslmode ออกจาก URL และระบุ ssl config เอง
const { Client } = require("pg");

function getDatabaseUrl() {
  return process.env.DATABASE_URL || "";
}

function createClient() {
  let url = getDatabaseUrl();
  url = url.replace(/[?&]sslmode=[^&]*/g, "").replace(/\?$/, "");
  return new Client({
    connectionString: url,
    ssl: { rejectUnauthorized: false },
  });
}

module.exports = { createClient, getDatabaseUrl };
