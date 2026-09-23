/**
 * =========================================================================
 * ⚙️ SETUP SCRIPT – Rekap Sewa & Pajak
 * Jalankan fungsi ini sekali untuk inisialisasi spreadsheet dan triggers
 * =========================================================================
 */

/**
 * Jalankan fungsi ini SATU KALI setelah deploy:
 * 1. Buka Apps Script
 * 2. Pilih fungsi "setupProject"
 * 3. Klik Run
 */
function setupProject() {
  Logger.log('🚀 Memulai setup Rekap Sewa & Pajak...');

  // Pastikan spreadsheet ada
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✅ Spreadsheet ditemukan: ' + ss.getName());
  } catch(e) {
    Logger.log('❌ Spreadsheet ID tidak valid! Ganti SPREADSHEET_ID di code.gs');
    Logger.log('   Cara mendapatkan ID: buka spreadsheet → lihat URL → salin bagian panjang di tengah URL');
    return;
  }

  // Buat sheet yang dibutuhkan
  ensureSheetsExist(ss);
  Logger.log('✅ Sheet-sheet berhasil dibuat/diverifikasi');

  // Hapus trigger lama
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // Buat trigger reminder harian (jam 08:00)
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  Logger.log('✅ Trigger reminder harian dibuat (setiap hari jam 08:00)');

  // Buat trigger laporan bulanan (tanggal 1 setiap bulan)
  ScriptApp.newTrigger('sendMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .create();
  Logger.log('✅ Trigger laporan bulanan dibuat (tanggal 1 jam 07:00)');

  Logger.log('');
  Logger.log('🎉 Setup selesai! Langkah berikutnya:');
  Logger.log('   1. Deploy sebagai Web App (Deploy > New Deployment)');
  Logger.log('   2. Pilih "Execute as: Me" dan "Who has access: Anyone"');
  Logger.log('   3. Salin URL Web App dan akses dari browser');
}

/**
 * Cek status trigger yang aktif
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total trigger aktif: ' + triggers.length);
  triggers.forEach(function(t) {
    Logger.log('  - ' + t.getHandlerFunction() + ' | ' + t.getEventType());
  });
}

/**
 * Reset semua trigger
 */
function resetTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  Logger.log('Semua trigger dihapus. Jalankan setupProject() untuk membuat ulang.');
}

/**
 * Tes kirim reminder manual (tanpa cek tanggal)
 */
function testManualReminder() {
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cfg = getAdminConfig(ss);
  if (!cfg.emails.length) {
    Logger.log('❌ Belum ada email terdaftar. Tambahkan email di halaman Admin terlebih dahulu.');
    return;
  }
  sendTestReminderEmail(cfg);
  Logger.log('✅ Email tes berhasil dikirim ke: ' + cfg.emails.join(', '));
}

/**
 * Tampilkan info spreadsheet
 */
function printSpreadsheetInfo() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Nama: ' + ss.getName());
    Logger.log('URL: ' + ss.getUrl());
    var sheets = ss.getSheets();
    sheets.forEach(function(sh) {
      Logger.log('  Sheet: ' + sh.getName() + ' (' + (sh.getLastRow()-1) + ' baris data)');
    });
  } catch(e) {
    Logger.log('Error: ' + e.toString());
  }
}
