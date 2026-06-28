// scripts/export-backup.js
// ใช้สำหรับดึงข้อมูลจาก PostgreSQL กรณี Google Sheets เข้าไม่ได้

process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0'; // ✅ ปิด SSL Check

require('dotenv').config();
const { Client } = require('pg');
const fs = require('fs');
const path = require('path');

const DATABASE_URL = process.env.DATABASE_URL;

const tables = [
  'backup_stock_master',
  'backup_stock_office',
  'backup_stock_site',
  'backup_transfer_log',
  'backup_asset_list',
  'backup_asset_history',
  'backup_part_catalog',
  'backup_farm_sites',
  'backup_farm_houses',
  'backup_users',
  'backup_audit_log'
];

async function exportAll() {
  if (!DATABASE_URL) {
    console.error('❌ DATABASE_URL not set');
    process.exit(1);
  }

  // ✅ แก้ไข SSL Configuration
  const client = new Client({
    connectionString: DATABASE_URL,
    ssl: {
      rejectUnauthorized: false  // ✅ ปิดการตรวจสอบ Certificate
    }
  });

  try {
    await client.connect();
    console.log('✅ Connected to PostgreSQL');

    const exportDir = path.join(__dirname, '../backup-export');
    if (!fs.existsSync(exportDir)) {
      fs.mkdirSync(exportDir, { recursive: true });
    }

    let totalRows = 0;

    for (const table of tables) {
      try {
        const res = await client.query(`SELECT * FROM ${table}`);
        const rows = res.rows;
        const filePath = path.join(exportDir, `${table}.json`);
        fs.writeFileSync(filePath, JSON.stringify(rows, null, 2));
        totalRows += rows.length;
        console.log(`✅ ${table}: ${rows.length} แถว → ${path.basename(filePath)}`);
      } catch (err) {
        console.error(`❌ Error exporting ${table}:`, err.message);
      }
    }

    // สร้างไฟล์ CSV สรุป
    const csvPath = path.join(exportDir, 'all_tables_summary.csv');
    let csv = 'Table,RowCount,ExportDate\n';
    for (const table of tables) {
      const filePath = path.join(exportDir, `${table}.json`);
      if (fs.existsSync(filePath)) {
        const content = fs.readFileSync(filePath, 'utf8');
        const rows = JSON.parse(content);
        csv += `${table},${rows.length},${new Date().toISOString()}\n`;
      }
    }
    fs.writeFileSync(csvPath, csv);

    console.log(`\n✅ Export completed: ${totalRows} total rows`);
    console.log(`📁 Folder: ${exportDir}`);
    console.log(`📋 Summary: ${csvPath}`);

    await client.end();
  } catch (err) {
    console.error('❌ Error:', err.message);
    await client.end();
  }
}

exportAll();