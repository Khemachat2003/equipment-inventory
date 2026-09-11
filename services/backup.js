// services/backup.js
const { getSheetsClient, SPREADSHEET_ID } = require('./sheets');
const format = require('pg-format');
const { createClient } = require('./pg');

// ============================================================
// VERSIONED BACKUP CONFIG
// เก็บ backup ล่าสุด N run ต่อตาราง (run เก่าถูก prune "หลัง" insert สำเร็จเท่านั้น
// → ถ้า backup ล้มกลางทาง ข้อมูล run เก่ายังครบ ไม่หายเหมือนระบบ DROP/CREATE เดิม)
// ============================================================
const BACKUP_KEEP_RUNS = parseInt(process.env.BACKUP_KEEP_RUNS, 10) || 3;

// run id รูปแบบ YYYYMMDDHHmmss (เรียงเวลาได้ + เทียบ max() ได้ตรงๆ)
function buildRunId(date = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return (
    date.getFullYear() +
    p(date.getMonth() + 1) +
    p(date.getDate()) +
    p(date.getHours()) +
    p(date.getMinutes()) +
    p(date.getSeconds())
  );
}

const DATABASE_URL = process.env.DATABASE_URL;

// ============================================================
// 1. BACKUP AUDIT LOG (เดิม)
// ============================================================
async function backupToPostgres() {
  // ... (โค้ดเดิม ไม่เปลี่ยนแปลง) ...
}

// ============================================================
// 2. FULL SYSTEM BACKUP (versioned)
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
    const runId = buildRunId();

    for (const config of sheetsConfig) {
      try {
        // 1) สร้างตารางถ้ายังไม่มี + เติมคอลัมน์ backup_run_id ให้ตารางเก่า (schema ก่อน versioned)
        //    ห้าม DROP เหมือนเดิม — เดิม backup ล้มกลางทางข้อมูล backup เก่าหายหมด
        const dbColumns = config.columns.map(col => columnMap[col] || col);
        const createSQL = format(`
          CREATE TABLE IF NOT EXISTS %I (
            id SERIAL PRIMARY KEY,
            %s,
            backup_run_id TEXT,
            backup_date TIMESTAMP DEFAULT NOW()
          )
        `, config.table, dbColumns.map(col => `${col} TEXT`).join(', '));
        await client.query(createSQL);
        await client.query(format('ALTER TABLE %I ADD COLUMN IF NOT EXISTS backup_run_id TEXT', config.table));

        // 2) ดึงข้อมูลจาก Google Sheets
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: config.range,
        });
        const rows = res.data.values || [];

        // 3) Insert แถวของ run นี้ (batch ทีละ 500 แถว — เร็วกว่า insert ทีละแถวมาก)
        let inserted = 0;
        if (rows.length > 0) {
          const columnNames = [...config.columns.map(col => columnMap[col] || col), 'backup_run_id'];
          const validRows = rows.filter(r => r[0]);
          const BATCH = 500;
          const COLS = config.columns.length;

          for (let i = 0; i < validRows.length; i += BATCH) {
            const chunk = validRows.slice(i, i + BATCH);
            const values = [];
            const rowGroups = [];
            chunk.forEach((row, ridx) => {
              const rowPh = [];
              for (let c = 0; c < COLS; c++) {
                values.push(row[c] || '');
                rowPh.push(`$${ridx * (COLS + 1) + c + 1}`);
              }
              // backup_run_id เป็นพารามิเตอร์ตัวสุดท้าย + NOW() inline สำหรับ backup_date
              values.push(runId);
              rowPh.push(`$${ridx * (COLS + 1) + COLS + 1}, NOW()`);
              rowGroups.push('(' + rowPh.join(', ') + ')');
            });

            // สร้าง SQL แบบ batch: INSERT INTO t (c1,c2) VALUES ($1,$2),($3,$4),...
            // ใช้ format() แค่กับชื่อตาราง (มาจาก whitelist config) ที่เหลือเป็น string ตรงๆ ปลอดภัย
            const insertSQL = format(
              'INSERT INTO %I (%s) VALUES %s',
              config.table,
              columnNames.join(', '),
              rowGroups.join(', ')
            );

            await client.query(insertSQL, values);
            inserted += chunk.length;
          }
        }

        // 4) ลบแถว legacy (schema เก่าก่อน versioned ไม่มี backup_run_id) แล้ว prune run เก่า
        //    เก็บไว้แค่ BACKUP_KEEP_RUNS run ล่าสุด — ทำ "หลัง" insert สำเร็จเท่านั้น
        await client.query(format('DELETE FROM %I WHERE backup_run_id IS NULL', config.table));
        const pruneRes = await client.query(format(
          `DELETE FROM %I WHERE backup_run_id IN (
             SELECT DISTINCT backup_run_id FROM %I WHERE backup_run_id IS NOT NULL
             ORDER BY backup_run_id DESC
             OFFSET $1
           )`,
          config.table, config.table
        ), [BACKUP_KEEP_RUNS]);

        totalRows += inserted;
        results[config.name] = {
          inserted,
          prunedOldRuns: pruneRes.rowCount > 0,
          table: config.table,
        };
        console.log(`✅ ${config.name}: +${inserted} แถว (run ${runId})`);

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

module.exports = { backupToPostgres, fullSystemBackup, buildRunId, BACKUP_KEEP_RUNS };