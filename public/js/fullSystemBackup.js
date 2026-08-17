// ============================================================
// BACKUP & DATABASE FUNCTIONS
// ============================================================

/**
 * ทำ Backup ทั้งระบบ (เรียก API /api/full-backup)
 */
async function fullSystemBackup() {
  const status = document.getElementById('fullBackupStatus');
  if (!status) return;
  status.textContent = '⏳ กำลัง Backup ทั้งระบบ...';
  status.style.color = '#f59e0b';
  try {
    const res = await fetch('/api/full-backup', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      status.textContent = `✅ Backup ทั้งระบบสำเร็จ (${data.totalRows} แถว)`;
      status.style.color = '#16a34a';
    } else {
      status.textContent = '❌ Backup ล้มเหลว: ' + (data.error || '');
      status.style.color = '#dc2626';
    }
  } catch (e) {
    status.textContent = '❌ Error: ' + e.message;
    status.style.color = '#dc2626';
  }
}

/**
 * เปิด Google Sheets ในแท็บใหม่
 */
function openDatabase() {
  fetch('/api/settings/spreadsheet-url')
    .then(function (r) { return r.json(); })
    .then(function (d) { window.open(d.url, '_blank'); })
    .catch(function () {
      window.open('https://docs.google.com/spreadsheets/d/1CheIF--yOt2mRxubU1000TmIIjKpuzIExH-9O0RS7FA', '_blank');
    });
}