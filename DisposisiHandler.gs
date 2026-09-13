// ============================================================
// DisposisiHandler.gs — Disposisi & Notifikasi
// ============================================================

// ============================================================
// CREATE DISPOSISI (Pimpinan → Pegawai)
// ============================================================
function createDisposisi(token, data) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  if (user.role !== 'Pimpinan') {
    return { success: false, message: 'Hanya Pimpinan yang dapat membuat disposisi.' };
  }

  if (!data.id_surat || !data.ke_id || !data.instruksi) {
    return { success: false, message: 'ID Surat, penerima, dan instruksi wajib diisi.' };
  }

  var sheet = getSheet(SHEET_NAMES.DISPOSISI);
  var disposisiId = 'DSP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  var now = new Date();

  // Ambil nama penerima
  var penerimaNama = getUserNameById_(data.ke_id);

  // ── Upload file lampiran disposisi ke Drive (jika ada) ──
  var dispFileUrl = '';
  if (data.file_base64 && data.file_name) {
    try {
      var dispFolder = getOrCreateFolder_(PDF_FOLDER_NAME);
      var dispBlob = Utilities.newBlob(
        Utilities.base64Decode(data.file_base64.replace(/^data:[^;]+;base64,/, '')),
        data.file_type || 'application/octet-stream',
        data.file_name
      );
      var dispFile = dispFolder.createFile(dispBlob);
      dispFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      dispFileUrl = 'https://drive.google.com/file/d/' + dispFile.getId() + '/view?usp=sharing';
    } catch (uploadErr) {
      Logger.log('Disposisi file upload error (non-fatal): ' + uploadErr.toString());
    }
  }

  sheet.appendRow([
    disposisiId,
    data.id_surat,
    user.userId,
    user.nama,
    data.ke_id,
    penerimaNama,
    data.instruksi,
    now.toISOString(),
    'belum_dibaca',
    dispFileUrl           // file_url (kolom ke-10)
  ]);

  // Kirim notifikasi ke penerima
  var notifMsg = '📋 Anda mendapat disposisi dari ' + user.nama + ': "' + data.instruksi + '"';
  if (dispFileUrl) notifMsg += ' (dengan lampiran file)';
  createNotifikasi(data.ke_id, notifMsg, 'disposisi');

  return { success: true, message: 'Disposisi berhasil dikirim ke ' + penerimaNama + '.', disposisiId: disposisiId };
}

// ============================================================
// GET DISPOSISI MASUK (untuk Pegawai)
// ============================================================
function getDisposisiMasuk(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.DISPOSISI);
  var all = sheetToObjects(sheet);

  var masuk = all.filter(function (d) {
    return d.ke_id === user.userId && !String(d.instruksi).startsWith('[BALASAN]');
  });

  masuk.sort(function (a, b) { return new Date(b.tanggal) - new Date(a.tanggal); });
  return { success: true, data: masuk };
}

// ============================================================
// GET DISPOSISI KELUAR (untuk Pimpinan — yang telah dia kirim)
// ============================================================
function getDisposisiKeluar(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.DISPOSISI);
  var all = sheetToObjects(sheet);

  var keluar = all.filter(function (d) {
    return d.dari_id === user.userId && !String(d.instruksi).startsWith('[BALASAN]');
  });

  // Sertakan balasan yang ditujukan ke pimpinan ini
  var balasan = all.filter(function (d) {
    return d.ke_id === user.userId && String(d.instruksi).startsWith('[BALASAN]');
  });

  keluar.sort(function (a, b) { return new Date(b.tanggal) - new Date(a.tanggal); });
  return { success: true, data: keluar, balasan: balasan };
}

// ============================================================
// BALAS DISPOSISI (Pegawai → Pimpinan)
// ============================================================
function replyDisposisi(token, disposisiId, balasan) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  if (user.role !== 'Pegawai') {
    return { success: false, message: 'Hanya Pegawai yang dapat membalas disposisi.' };
  }

  if (!balasan) return { success: false, message: 'Isi balasan tidak boleh kosong.' };

  // Cari disposisi asli
  var sheet = getSheet(SHEET_NAMES.DISPOSISI);
  var all = sheetToObjects(sheet);
  var original = null;
  for (var i = 0; i < all.length; i++) {
    if (all[i].id_disposisi === disposisiId) { original = all[i]; break; }
  }
  if (!original) return { success: false, message: 'Disposisi tidak ditemukan.' };

  // Buat reply
  var replyId = 'DSP-R-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  var now = new Date();

  sheet.appendRow([
    replyId,
    original.id_surat,
    user.userId,
    user.nama,
    original.dari_id,         // kirim ke pimpinan asal
    original.dari_nama,
    '[BALASAN] ' + balasan,
    now.toISOString(),
    'belum_dibaca'
  ]);

  // Mark disposisi asli sudah dibaca
  markDisposisiReadDirect_(disposisiId);

  // Notif ke pimpinan
  createNotifikasi(
    original.dari_id,
    '💬 ' + user.nama + ' membalas disposisi Anda: "' + balasan + '"',
    'disposisi_reply'
  );

  return { success: true, message: 'Balasan berhasil dikirim.' };
}

// ============================================================
// MARK DISPOSISI SUDAH DIBACA
// ============================================================
function markDisposisiRead(token, disposisiId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  markDisposisiReadDirect_(disposisiId);
  return { success: true };
}

function markDisposisiReadDirect_(disposisiId) {
  var sheet = getSheet(SHEET_NAMES.DISPOSISI);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('id_disposisi');
  var readIdx = headers.indexOf('status_baca');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === disposisiId) {
      sheet.getRange(i + 1, readIdx + 1).setValue('sudah_dibaca');
      return;
    }
  }
}

// ============================================================
// NOTIFIKASI
// ============================================================

/**
 * Buat notifikasi baru (dipanggil dari handler lain, tanpa token check)
 */
function createNotifikasi(userId, pesan, tipe) {
  try {
    var sheet = getSheet(SHEET_NAMES.NOTIFIKASI);
    var notifId = 'NTF-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    var now = new Date();

    sheet.appendRow([
      notifId,
      userId,
      pesan,
      'belum_dibaca',
      now.toISOString(),
      tipe || 'info'
    ]);

    return notifId;
  } catch (e) {
    Logger.log('createNotifikasi error: ' + e.toString());
    return null;
  }
}

/**
 * Ambil semua notifikasi user yang sedang login
 */
function getNotifikasi(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.NOTIFIKASI);
  var all = sheetToObjects(sheet);

  var mine = all.filter(function (n) { return n.id_user === user.userId; });
  mine.sort(function (a, b) { return new Date(b.tanggal) - new Date(a.tanggal); });

  return { success: true, data: mine };
}

/**
 * Tandai satu notifikasi sudah dibaca
 */
function markNotifRead(token, notifId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var sheet = getSheet(SHEET_NAMES.NOTIFIKASI);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('id_notif');
  var readIdx = headers.indexOf('status_baca');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === notifId) {
      sheet.getRange(i + 1, readIdx + 1).setValue('sudah_dibaca');
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * Tandai semua notifikasi user sudah dibaca
 */
function markAllNotifRead(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.NOTIFIKASI);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idUserIdx = headers.indexOf('id_user');
  var readIdx = headers.indexOf('status_baca');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idUserIdx] === user.userId && data[i][readIdx] === 'belum_dibaca') {
      sheet.getRange(i + 1, readIdx + 1).setValue('sudah_dibaca');
    }
  }
  return { success: true };
}

/**
 * Hitung jumlah notifikasi yang belum dibaca
 */
function getUnreadCount(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: true, count: 0 };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.NOTIFIKASI);
  var all = sheetToObjects(sheet);

  var count = all.filter(function (n) {
    return n.id_user === user.userId && n.status_baca === 'belum_dibaca';
  }).length;

  return { success: true, count: count };
}

// ============================================================
// HELPER: Ambil nama user berdasarkan ID
// ============================================================
function getUserNameById_(userId) {
  try {
    var sheet = getSheet(SHEET_NAMES.USERS);
    var users = sheetToObjects(sheet);
    for (var i = 0; i < users.length; i++) {
      if (users[i].id === userId) return users[i].nama;
    }
    return 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}
