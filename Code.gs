// ============================================================
// AMSP - Aplikasi Manajemen Surat Perintah
// Code.gs — Main Router, Inisialisasi Sheet & Utilities
// ============================================================
// Cara Deploy:
// 1. Buka Google Apps Script (script.google.com)
// 2. Buat project baru, paste semua file .gs dan .html
// 3. Jalankan fungsi initializeSheets() SEKALI untuk setup database
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Salin URL Web App, bagikan ke pengguna
// ============================================================

// ============================================================
// KONFIGURASI NAMA APLIKASI
// ============================================================
var APP_CONFIG = {
  nama_aplikasi: 'AMSP PUPR',
  nama_lengkap: 'Aplikasi Manajemen Surat Perintah PUPR',
  nama_instansi: 'Dinas PUPR Kota Ambon',
  versi: '2.0.0'
};

var SHEET_NAMES = {
  USERS: 'Users',
  SURAT: 'Surat',
  DISPOSISI: 'Disposisi',
  NOTIFIKASI: 'Notifikasi'
};

// ============================================================
// OTORISASI GOOGLE DRIVE, DOKUMEN & SPREADSHEET
// ============================================================
function otorisasiDriveDanPDF() {
  Logger.log('Memeriksa izin Google Docs...');
  var testDoc = DocumentApp.create('AMSP_Otorisasi_Test');
  var docId = testDoc.getId();

  Logger.log('Memeriksa izin Google Drive...');
  var docFile = DriveApp.getFileById(docId);
  var folder = DriveApp.getFoldersByName('AMSP_Surat_PDF');
  if (!folder.hasNext()) {
    DriveApp.createFolder('AMSP_Surat_PDF');
  }

  docFile.setTrashed(true);
  Logger.log('✅ Semua izin Google Drive & Docs berhasil diotorisasi!');
  return 'Otorisasi Berhasil!';
}

// ============================================================
// MAIN ROUTER — doGet(e)
// ============================================================
function doGet(e) {
  var params = e.parameter || {};

  // === ROUTE: Halaman publik scan QR Code ===
  if (params.scan_ttd) {
    var template = HtmlService.createTemplateFromFile('ScanTTD');
    template.suratId = params.scan_ttd;
    template.webAppUrl = getWebAppUrl();
    return template.evaluate()
      .setTitle('Verifikasi Tanda Tangan — AMSP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // === ROUTE: Halaman dashboard berdasarkan parameter page ===
  var page = params.page || 'login';
  var token = params.token || '';
  var template;

  switch (page) {
    case 'admin':
      template = HtmlService.createTemplateFromFile('DashboardAdmin');
      template.initToken = token;
      break;
    case 'pimpinan':
      template = HtmlService.createTemplateFromFile('DashboardPimpinan');
      template.initToken = token;
      break;
    case 'pegawai':
      template = HtmlService.createTemplateFromFile('DashboardPegawai');
      template.initToken = token;
      break;
    default:
      template = HtmlService.createTemplateFromFile('Login');
      template.initToken = '';
      template.initRole = params.role || '';
      break;
  }

  return template.evaluate()
    .setTitle('AMSP PUPR — Aplikasi Manajemen Surat Perintah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// INCLUDE HTML PARTIAL (untuk CSS.html & JavaScript.html)
// ============================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// GET WEB APP URL
// ============================================================
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}

// ============================================================
// GET APP CONFIG (dipanggil dari frontend)
// ============================================================
function getAppConfig() {
  return APP_CONFIG;
}

// ============================================================
// INISIALISASI GOOGLE SHEETS (Jalankan sekali saat setup awal)
// ============================================================
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('ERROR: Tidak ada spreadsheet aktif. Buka GAS dari Google Sheets yang sudah ada.');
    return { success: false, message: 'Tidak ada spreadsheet aktif.' };
  }

  // ── Sheet Users ──
  setupSheet_(ss, SHEET_NAMES.USERS, [
    'id', 'nama', 'username', 'password', 'role',
    'status_approval', 'tanggal_daftar', 'email'
  ]);

  // ── Sheet Surat ──
  setupSheet_(ss, SHEET_NAMES.SURAT, [
    'id_surat', 'jenis_surat', 'nomor_surat', 'pengirim', 'penerima',
    'tembusan', 'perihal', 'isi_surat', 'tanggal', 'status',
    'file_url', 'id_pembuat', 'nama_pembuat',
    'id_pimpinan_ttd', 'nama_pimpinan_ttd', 'tanggal_ttd', 'tanggal_dibuat'
  ]);

  // ── Sheet Disposisi ──
  setupSheet_(ss, SHEET_NAMES.DISPOSISI, [
    'id_disposisi', 'id_surat', 'dari_id', 'dari_nama',
    'ke_id', 'ke_nama', 'instruksi', 'tanggal', 'status_baca'
  ]);

  // ── Sheet Notifikasi ──
  setupSheet_(ss, SHEET_NAMES.NOTIFIKASI, [
    'id_notif', 'id_user', 'pesan', 'status_baca', 'tanggal', 'tipe'
  ]);

  // Buat akun Admin default
  createDefaultAdmin_(ss);

  Logger.log('✅ AMSP: Semua sheet berhasil diinisialisasi.');
  return { success: true, message: 'Inisialisasi berhasil! Akun Admin: username=admin, password=admin123' };
}

function setupSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Set header hanya jika belum ada atau baris 1 kosong
  var firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1565c0');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // Auto-resize kolom
    for (var i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 150);
    }
  }
  return sheet;
}

function createDefaultAdmin_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();

  // Cek apakah admin sudah ada
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === 'admin' && data[i][4] === 'Admin') {
      Logger.log('Admin default sudah ada, skip.');
      return;
    }
  }

  var adminId = generateId_();
  var now = new Date();
  sheet.appendRow([
    adminId,
    'Administrator',
    'admin',
    'admin123',
    'Admin',
    'approved',
    now.toISOString(),
    'admin@amsp.local'
  ]);
  Logger.log('✅ Akun Admin default dibuat: admin / admin123');
}

// ============================================================
// UTILITY FUNCTIONS (dipakai di semua .gs files)
// ============================================================

function generateId_() {
  return Utilities.getUuid();
}
// Alias tanpa underscore untuk pemanggilan antar file
var generateId = generateId_;

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Tidak ada spreadsheet aktif. Pastikan GAS terhubung ke Google Sheets.');
  return ss;
}

function getSheet(sheetName) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Coba inisialisasi otomatis
    initializeSheets();
    sheet = ss.getSheetByName(sheetName);
  }
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');
  return sheet;
}

/**
 * Konversi data sheet menjadi array of objects menggunakan baris pertama sebagai header.
 */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var objects = [];
  for (var i = 1; i < data.length; i++) {
    // Skip baris kosong
    if (!data[i][0]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = val.toISOString();
      } else if (val === null || val === undefined) {
        val = '';
      } else {
        val = String(val);
      }
      obj[headers[j]] = val;
    }
    objects.push(obj);
  }
  return objects;
}

/**
 * Format tanggal ke format Indonesia lengkap dengan jam
 */
function formatDate(dateVal) {
  if (!dateVal) return '-';
  try {
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal);
    return Utilities.formatDate(d, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
  } catch (e) {
    return String(dateVal);
  }
}

/**
 * Format tanggal ke format Indonesia tanpa jam
 */
function formatDateOnly(dateVal) {
  if (!dateVal) return '-';
  try {
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal);
    return Utilities.formatDate(d, 'Asia/Jakarta', 'dd MMMM yyyy');
  } catch (e) {
    return String(dateVal);
  }
}
