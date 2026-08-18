// services/backup.js
const { getSheetsClient, SPREADSHEET_ID } = require('./sheets');
const format = require('pg-format');
const { createClient } = require('./pg');

const DATABASE_URL = process.env.DATABASE_URL;

// ============================================================
// 1. BACKUP AUDIT LOG (เดิม)
// ============================================================
async function backupToPostgres() {
  // ... (โค้ดเดิม ไม่เปลี่ยนแปลง) ...
}

// ============================================================
// 2. FULL SYSTEM BACKUP (แก้ไขแล้ว)
// ============================================================
async function fullSystemBackup() {
  if (!process.env.DATABASE_URL) {
    return { success: false, error: 'DATABASE_URL not set' };
  }

  const client = createClient();

  try {
    await client.connect();
    console.log('✅ Connected to PostgreSQL for Full System Backup');

    const sheets = await getSheetsClient();

    // ============================================================
    // กำหนด Sheets ทั้งหมด (พร้อมชื่อ Sheet ใน Range)
    // ============================================================
    const sheetsConfig = [
      // 1. Stock_Master
      { 
        name: 'Stock_Master', 
        range: `'Stock_Master'!A2:G`,  // ✅ ระบุชื่อ Sheet
        table: 'backup_stock_master', 
        columns: ['code', 'name', 'type', 'category', 'unit', 'total', 'remaining'] 
      },
      // 2. Users
      { 
        name: 'Users', 
        range: `'Users'!A2:C`,          // ✅ ระบุชื่อ Sheet
        table: 'backup_users', 
        columns: ['username', 'password', 'role'] 
      },
      // 3. Transfer_Log
      { 
        name: 'Transfer_Log', 
        range: `'Transfer_Log'!A2:H`,   // ✅ ระบุชื่อ Sheet
        table: 'backup_transfer_log', 
        columns: ['timestamp', 'code', 'name', 'qty', 'type', 'from_col', 'to_col', 'user_col'] 
      },
      // 4. Audit_Log
      { 
        name: 'Audit_Log', 
        range: `'Audit_Log'!A2:F`,      // ✅ ระบุชื่อ Sheet
        table: 'backup_audit_log', 
        columns: ['timestamp', 'user_col', 'action', 'module', 'detail', 'ip'] 
      },
      // 5. Stock_Office
      { 
        name: 'Stock_Office', 
        range: `'Stock_Office'!A2:C`,   // ✅ ระบุชื่อ Sheet
        table: 'backup_stock_office', 
        columns: ['code', 'name', 'qty'] 
      },
      // 6. Stock_Site
      { 
        name: 'Stock_Site', 
        range: `'Stock_Site'!A2:C`,     // ✅ ระบุชื่อ Sheet
        table: 'backup_stock_site', 
        columns: ['code', 'name', 'qty'] 
      },
      // 7. Farm_Sites
      { 
        name: 'Farm_Sites', 
        range: `'Farm_Sites'!A2:F`,     // ✅ ระบุชื่อ Sheet
        table: 'backup_farm_sites', 
        columns: ['site_id', 'site_name', 'farm_type', 'province', 'manager', 'note'] 
      },
      // 8. Farm_Houses
      { 
        name: 'Farm_Houses', 
        range: `'Farm_Houses'!A2:F`,    // ✅ ระบุชื่อ Sheet
        table: 'backup_farm_houses', 
        columns: ['house_id', 'site_id', 'house_name', 'house_type', 'capacity', 'note'] 
      },
      // 9. Part_Catalog
      { 
        name: 'Part_Catalog', 
        range: `'Part_Catalog'!A2:G`,   // ✅ ระบุชื่อ Sheet
        table: 'backup_part_catalog', 
        columns: ['part_number', 'part_name', 'category', 'description', 'unit', 'total_qty', 'last_updated'] 
      },
      // 10. Asset_List
      { 
        name: 'Asset_List', 
        range: `'Asset_List'!A2:M`,     // ✅ ระบุชื่อ Sheet
        table: 'backup_asset_list', 
        columns: ['asset_id', 'code', 'name', 'part_number', 'serial_number', 'status', 'location', 'site_name', 'user_col', 'date', 'farm_type', 'house_id', 'house_name'] 
      },
      // 11. Asset_History
      { 
        name: 'Asset_History', 
        range: `'Asset_History'!A2:G`,  // ✅ ระบุชื่อ Sheet
        table: 'backup_asset_history', 
        columns: ['date', 'serial_number', 'action', 'from_col', 'to_col', 'user_col', 'remark'] 
      },
      // 12. Damaged_Assets
      { 
        name: 'Damaged_Assets', 
        range: `'Damaged_Assets'!A2:L`,  // ✅ ระบุชื่อ Sheet
        table: 'backup_damaged_assets', 
        columns: ['date', 'serial_number', 'asset_id', 'code', 'name', 'part_number', 'status', 'old_location', 'old_site', 'user_col', 'remark', 'action'] 
      },
    ];

    // Mapping: คอลัมน์ที่เป็นคำสงวน
    const columnMap = {
      'from_col': '"from"',
      'to_col': '"to"',
      'user_col': '"user"'
    };

    let totalRows = 0;
    const results = {};

    for (const config of sheetsConfig) {
      try {
        // ลบตารางเก่า
        await client.query(`DROP TABLE IF EXISTS ${config.table}`);
        console.log(`🗑️ Dropped old table: ${config.table}`);

        // สร้างตารางใหม่
        const dbColumns = config.columns.map(col => {
          return columnMap[col] || col;
        });

        const createSQL = format(`
          CREATE TABLE %I (
            id SERIAL PRIMARY KEY,
            %s,
            backup_date TIMESTAMP DEFAULT NOW()
          )
        `, config.table, dbColumns.map(col => `${col} TEXT`).join(', '));

        await client.query(createSQL);
        console.log(`✅ Created table: ${config.table}`);

        // ✅ ดึงข้อมูลจาก Google Sheets (ใช้ range ที่มีชื่อ Sheet)
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: config.range,  // ✅ ตอนนี้มีชื่อ Sheet แล้ว
        });
        const rows = res.data.values || [];

        // Insert ข้อมูล
        let inserted = 0;
        if (rows.length > 0) {
          const columnNames = config.columns.map(col => columnMap[col] || col);
          const placeholders = config.columns.map((_, i) => `$${i + 1}`).join(', ');

          for (const row of rows) {
            if (!row[0]) continue;
            const values = config.columns.map((col, i) => {
              return row[i] || '';
            });

            const insertSQL = format(
              'INSERT INTO %I (%s) VALUES (%s)',
              config.table,
              columnNames.join(', '),
              placeholders
            );

            await client.query(insertSQL, values);
            inserted++;
          }
        }

        totalRows += inserted;
        results[config.name] = { inserted, table: config.table };
        console.log(`✅ ${config.name}: ${inserted} แถว`);

      } catch (err) {
        console.error(`❌ Error backing up ${config.name}:`, err.message);
        results[config.name] = { error: err.message };
      }
    }

    console.log(`✅ Full System Backup Completed: ${totalRows} total rows`);
    await client.end();
    return { success: true, totalRows, results };

  } catch (err) {
    console.error('❌ Full System Backup error:', err);
    await client.end();
    return { success: false, error: err.message };
  }
}

module.exports = { backupToPostgres, fullSystemBackup };