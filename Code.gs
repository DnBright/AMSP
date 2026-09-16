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

  // ── Sheet Surat (22 Kolom Lengkap) ──
  setupSheet_(ss, SHEET_NAMES.SURAT, [
    'id_surat', 'jenis_surat', 'nomor_surat', 'pengirim', 'penerima',
    'tembusan', 'perihal', 'isi_surat', 'tanggal', 'status',
    'file_url', 'id_pembuat', 'nama_pembuat',
    'id_pimpinan_ttd', 'nama_pimpinan_ttd', 'tanggal_ttd', 'tanggal_dibuat',
    'sifat', 'waktu', 'tempat', 'lampiran', 'lampiran_file_url'
  ]);

  // ── Sheet Disposisi ──
  setupSheet_(ss, SHEET_NAMES.DISPOSISI, [
    'id_disposisi', 'id_surat', 'dari_id', 'dari_nama',
    'ke_id', 'ke_nama', 'instruksi', 'tanggal', 'status_baca', 'file_url'
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

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0 || !sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1565c0');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    for (var i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 150);
    }
  } else {
    // Sinkronisasi otomatis jika ada kolom header yang belum terdaftar
    var curRow = sheet.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0];
    var updated = false;
    for (var h = 0; h < headers.length; h++) {
      if (!curRow[h]) {
        sheet.getRange(1, h + 1).setValue(headers[h]);
        updated = true;
      }
    }
    if (updated) {
      var fullHeaderRange = sheet.getRange(1, 1, 1, headers.length);
      fullHeaderRange.setBackground('#1565c0');
      fullHeaderRange.setFontColor('#ffffff');
      fullHeaderRange.setFontWeight('bold');
      fullHeaderRange.setHorizontalAlignment('center');
      SpreadsheetApp.flush();
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

var SURAT_STANDARD_HEADERS = [
  'id_surat', 'jenis_surat', 'nomor_surat', 'pengirim', 'penerima',
  'tembusan', 'perihal', 'isi_surat', 'tanggal', 'status',
  'file_url', 'id_pembuat', 'nama_pembuat',
  'id_pimpinan_ttd', 'nama_pimpinan_ttd', 'tanggal_ttd', 'tanggal_dibuat',
  'sifat', 'waktu', 'tempat', 'lampiran', 'lampiran_file_url'
];

var DISPOSISI_STANDARD_HEADERS = [
  'id_disposisi', 'id_surat', 'dari_id', 'dari_nama',
  'ke_id', 'ke_nama', 'instruksi', 'tanggal', 'status_baca', 'file_url'
];

function getSheet(sheetName) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Coba inisialisasi otomatis
    initializeSheets();
    sheet = ss.getSheetByName(sheetName);
  }
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');

  // Pastikan kolom sheet surat selalu lengkap 22 header
  if (sheetName === SHEET_NAMES.SURAT) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < SURAT_STANDARD_HEADERS.length) {
      var row1 = sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0];
      for (var idx = 0; idx < SURAT_STANDARD_HEADERS.length; idx++) {
        if (!row1[idx]) {
          sheet.getRange(1, idx + 1).setValue(SURAT_STANDARD_HEADERS[idx]);
        }
      }
      var headerRange = sheet.getRange(1, 1, 1, SURAT_STANDARD_HEADERS.length);
      headerRange.setBackground('#1565c0');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
    }
  }

  // Pastikan kolom sheet disposisi selalu lengkap 10 header
  if (sheetName === SHEET_NAMES.DISPOSISI) {
    var lastColDisp = sheet.getLastColumn();
    if (lastColDisp < DISPOSISI_STANDARD_HEADERS.length) {
      var row1Disp = sheet.getRange(1, 1, 1, Math.max(1, lastColDisp)).getValues()[0];
      for (var dIdx = 0; dIdx < DISPOSISI_STANDARD_HEADERS.length; dIdx++) {
        if (!row1Disp[dIdx]) {
          sheet.getRange(1, dIdx + 1).setValue(DISPOSISI_STANDARD_HEADERS[dIdx]);
        }
      }
      var dispHeaderRange = sheet.getRange(1, 1, 1, DISPOSISI_STANDARD_HEADERS.length);
      dispHeaderRange.setBackground('#1565c0');
      dispHeaderRange.setFontColor('#ffffff');
      dispHeaderRange.setFontWeight('bold');
    }
  }

  return sheet;
}

/**
 * Konversi data sheet menjadi array of objects menggunakan baris pertama sebagai header.
 */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var sheetName = sheet.getName();
  var isSurat = (sheetName === SHEET_NAMES.SURAT);
  var isDisposisi = (sheetName === SHEET_NAMES.DISPOSISI);
  var objects = [];

  for (var i = 1; i < data.length; i++) {
    // Skip baris kosong
    if (!data[i][0]) continue;
    var obj = {};
    var maxCols = Math.max(headers.length, data[i].length);

    for (var j = 0; j < maxCols; j++) {
      var key = headers[j];
      if (!key && isSurat && j < SURAT_STANDARD_HEADERS.length) {
        key = SURAT_STANDARD_HEADERS[j];
      }
      if (!key && isDisposisi && j < DISPOSISI_STANDARD_HEADERS.length) {
        key = DISPOSISI_STANDARD_HEADERS[j];
      }
      if (!key) continue;

      var val = data[i][j];
      if (val instanceof Date) {
        val = val.toISOString();
      } else if (val === null || val === undefined) {
        val = '';
      } else {
        val = String(val);
      }
      obj[key] = val;
    }

    // Pastikan data kolom tambahan selalu terpetakan jika ada nilainya
    if (isSurat) {
      if (!obj.sifat && data[i][17]) obj.sifat = String(data[i][17]);
      if (!obj.waktu && data[i][18]) obj.waktu = String(data[i][18]);
      if (!obj.tempat && data[i][19]) obj.tempat = String(data[i][19]);
      if (!obj.lampiran && data[i][20]) obj.lampiran = String(data[i][20]);
      if (!obj.lampiran_file_url && data[i][21]) obj.lampiran_file_url = String(data[i][21]);
    }
    if (isDisposisi) {
      if (!obj.file_url && data[i][9]) obj.file_url = String(data[i][9]);
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
