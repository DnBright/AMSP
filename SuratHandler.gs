// ============================================================
// SuratHandler.gs — CRUD Surat, Statistik, Arsip, Search
// ============================================================

// ============================================================
// GET LIST SURAT (berdasarkan role & filter)
// ============================================================
function getSuratList(token, filter) {
  try {
    var session = validateSession(token);
    if (!session.valid) return { success: false, message: session.message };

    var user = session.user;
    var uid = String(user.userId || user.id || '').trim();
    var uname = String(user.username || '').trim();
    var role = String(user.role || '').trim();

    var sheet = getSheet(SHEET_NAMES.SURAT);
    var allSurat = sheetToObjects(sheet);

    var result = [];

    allSurat.forEach(function (surat) {
      var masuk = false;
      var jenis = String(surat.jenis_surat || '').trim().toLowerCase();
      var idPembuat = String(surat.id_pembuat || '').trim();
      var penerima = String(surat.penerima || '').trim();

      if (role === 'Admin') {
        masuk = true; // Admin lihat semua
      } else if (role === 'Pimpinan') {
        // Pimpinan lihat: surat masuk, permohonan, dan surat yang dibuat pimpinan
        masuk = (jenis === 'masuk' || jenis === 'permohonan' || idPembuat === uid);
      } else if (role === 'Pegawai') {
        // Pegawai lihat: surat yang dia buat, atau yang ditujukan ke dia
        masuk = (idPembuat === uid || penerima === uid || penerima === uname || !idPembuat);
      }

      if (masuk) result.push(surat);
    });

    // Apply filter tambahan
    if (filter) {
      if (filter.jenis && filter.jenis !== 'semua') {
        var fJenis = String(filter.jenis).trim().toLowerCase();
        result = result.filter(function (s) {
          return String(s.jenis_surat || '').trim().toLowerCase() === fJenis;
        });
      }
      if (filter.status && filter.status !== 'semua') {
        var fStatus = String(filter.status).trim().toLowerCase();
        result = result.filter(function (s) {
          return String(s.status || '').trim().toLowerCase() === fStatus;
        });
      }
    }

    // Sort by tanggal_dibuat descending
    result.sort(function (a, b) {
      var da = a.tanggal_dibuat ? new Date(a.tanggal_dibuat).getTime() : 0;
      var db = b.tanggal_dibuat ? new Date(b.tanggal_dibuat).getTime() : 0;
      return db - da;
    });

    return { success: true, data: result };
  } catch (err) {
    Logger.log('getSuratList error: ' + err.toString());
    return { success: false, message: 'Gagal memuat data: ' + err.toString(), data: [] };
  }
}

// ============================================================
// GET SURAT BY ID
// ============================================================
function getSuratById(suratId) {
  var sheet = getSheet(SHEET_NAMES.SURAT);
  var allSurat = sheetToObjects(sheet);

  for (var i = 0; i < allSurat.length; i++) {
    if (allSurat[i].id_surat === suratId) {
      return { success: true, data: allSurat[i] };
    }
  }
  return { success: false, message: 'Surat tidak ditemukan.' };
}

// ============================================================
// CREATE SURAT BARU
// ============================================================
function createSurat(token, data) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;

  if (!data.perihal) return { success: false, message: 'Perihal surat wajib diisi.' };
  if (!data.jenis_surat) return { success: false, message: 'Jenis surat wajib dipilih.' };

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var suratId = 'SRT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  var nomorSurat = generateNomorSurat_(data.jenis_surat);
  var now = new Date();

  // ── File upload ke Drive (jika ada) ──
  var uploadedFileUrl = '';
  if (data.file_base64 && data.file_name) {
    try {
      var folder = getOrCreateFolder_(PDF_FOLDER_NAME);
      var decoded = Utilities.newBlob(
        Utilities.base64Decode(data.file_base64.replace(/^data:[^;]+;base64,/, '')),
        data.file_type || 'application/octet-stream',
        data.file_name
      );
      var uploaded = folder.createFile(decoded);
      uploaded.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      uploadedFileUrl = 'https://drive.google.com/file/d/' + uploaded.getId() + '/view?usp=sharing';
    } catch (uploadErr) {
      Logger.log('File upload error (non-fatal): ' + uploadErr.toString());
    }
  }

  sheet.appendRow([
    suratId,                            // id_surat
    data.jenis_surat,                   // jenis_surat
    nomorSurat,                         // nomor_surat
    data.pengirim || user.nama,         // pengirim
    data.penerima || '',                // penerima
    data.tembusan || '',                // tembusan
    data.perihal,                       // perihal
    data.isi_surat || '',               // isi_surat
    data.tanggal || now.toISOString(),  // tanggal
    'draft',                            // status
    uploadedFileUrl,                    // file_url (lampiran awal)
    user.userId,                        // id_pembuat
    user.nama,                          // nama_pembuat
    '',                                 // id_pimpinan_ttd
    '',                                 // nama_pimpinan_ttd
    '',                                 // tanggal_ttd
    now.toISOString(),                  // tanggal_dibuat
    data.sifat || 'Biasa',              // sifat
    data.waktu || '',                   // waktu
    data.tempat || '',                  // tempat
    data.lampiran || ''                 // lampiran (keterangan teks)
  ]);

  return {
    success: true,
    message: 'Surat berhasil dibuat.',
    suratId: suratId,
    nomorSurat: nomorSurat
  };
}

function generateNomorSurat_(jenis) {
  var sheet = getSheet(SHEET_NAMES.SURAT);
  var rowCount = Math.max(1, sheet.getLastRow() - 1); // Jumlah data (selain header)
  var now = new Date();
  var year = now.getFullYear();
  var month = String(now.getMonth() + 1).padStart(2, '0');

  var prefix = { masuk: 'SM', keluar: 'SK', permohonan: 'SP' }[jenis] || 'SU';
  return prefix + '/' + String(rowCount + 1).padStart(4, '0') + '/' + month + '/' + year;
}

// ============================================================
// UPDATE STATUS SURAT
// ============================================================
function updateSuratStatus(token, suratId, status) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;

  // Validasi hak akses per status
  if (status === 'disetujui' && user.role !== 'Pimpinan') {
    return { success: false, message: 'Hanya Pimpinan yang dapat menyetujui surat.' };
  }
  if (status === 'verifikasi' && user.role !== 'Pegawai') {
    return { success: false, message: 'Hanya Pegawai yang dapat melakukan verifikasi awal.' };
  }

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];

  var idIdx = headers.indexOf('id_surat');
  var statusIdx = headers.indexOf('status');
  var idPimpinanIdx = headers.indexOf('id_pimpinan_ttd');
  var namaPimpinanIdx = headers.indexOf('nama_pimpinan_ttd');
  var tanggalTtdIdx = headers.indexOf('tanggal_ttd');
  var fileUrlIdx = headers.indexOf('file_url');
  var perihalIdx = headers.indexOf('perihal');
  var idPembuatIdx = headers.indexOf('id_pembuat');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === suratId) {
      var rowNum = i + 1;

      // Update status
      sheet.getRange(rowNum, statusIdx + 1).setValue(status);

      if (status === 'disetujui') {
        // Simpan data TTD pimpinan
        sheet.getRange(rowNum, idPimpinanIdx + 1).setValue(user.userId);
        sheet.getRange(rowNum, namaPimpinanIdx + 1).setValue(user.nama);
        sheet.getRange(rowNum, tanggalTtdIdx + 1).setValue(new Date().toISOString());

        // Flush agar data TTD tersimpan sebelum PDF digenerate
        SpreadsheetApp.flush();

        // Generate PDF + QR Code
        try {
          var pdfResult = generateSuratPDF(suratId);
          if (pdfResult.success && pdfResult.fileUrl) {
            sheet.getRange(rowNum, fileUrlIdx + 1).setValue(pdfResult.fileUrl);
            SpreadsheetApp.flush();
          } else {
            Logger.log('generateSuratPDF notice: ' + (pdfResult ? pdfResult.message : 'Unknown result'));
          }
        } catch (pdfErr) {
          Logger.log('PDF error (non-fatal): ' + pdfErr.toString());
        }

        // Notif ke pembuat surat
        var perihal = sheetData[i][perihalIdx];
        var idPembuat = sheetData[i][idPembuatIdx];
        createNotifikasi(idPembuat,
          '✅ Surat Anda "' + perihal + '" telah disetujui dan ditandatangani oleh ' + user.nama, 'approval');
      }

      if (status === 'arsip') {
        var perihalArsip = sheetData[i][perihalIdx];
        var idPembuatArsip = sheetData[i][idPembuatIdx];
        createNotifikasi(idPembuatArsip,
          '📁 Surat "' + perihalArsip + '" telah dipindahkan ke Arsip.', 'info');
      }

      return { success: true, message: 'Status surat berhasil diperbarui menjadi: ' + status };
    }
  }

  return { success: false, message: 'Surat tidak ditemukan.' };
}

// ============================================================
// EDIT / UPDATE SURAT (hanya draft, oleh pembuat)
// ============================================================
function updateSurat(token, suratId, data) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.SURAT);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];

  var idIdx = headers.indexOf('id_surat');
  var idPembuatIdx = headers.indexOf('id_pembuat');
  var statusIdx = headers.indexOf('status');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === suratId) {
      // Hanya pembuat atau Admin yang bisa edit
      if (sheetData[i][idPembuatIdx] !== user.userId && user.role !== 'Admin') {
        return { success: false, message: 'Anda tidak berhak mengedit surat ini.' };
      }
      // Hanya bisa edit jika masih draft
      if (sheetData[i][statusIdx] !== 'draft' && user.role !== 'Admin') {
        return { success: false, message: 'Surat yang sudah diproses tidak bisa diedit.' };
      }

      var rowNum = i + 1;
      var fields = ['pengirim', 'penerima', 'tembusan', 'perihal', 'isi_surat', 'tanggal', 'sifat', 'waktu', 'tempat', 'lampiran'];
      fields.forEach(function (f) {
        if (data[f] !== undefined) {
          var colIdx = headers.indexOf(f);
          if (colIdx >= 0) sheet.getRange(rowNum, colIdx + 1).setValue(data[f]);
        }
      });

      return { success: true, message: 'Surat berhasil diperbarui.' };
    }
  }

  return { success: false, message: 'Surat tidak ditemukan.' };
}

// ============================================================
// HAPUS SURAT
// ============================================================
function deleteSurat(token, suratId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  if (user.role !== 'Admin') {
    return { success: false, message: 'Hanya Admin yang dapat menghapus surat.' };
  }

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  var idIdx = headers.indexOf('id_surat');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === suratId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Surat berhasil dihapus.' };
    }
  }

  return { success: false, message: 'Surat tidak ditemukan.' };
}

// ============================================================
// SEARCH SURAT
// ============================================================
function searchSurat(token, keyword) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  if (!keyword) return { success: true, data: [] };

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var allSurat = sheetToObjects(sheet);
  var kw = keyword.toLowerCase();

  var hasil = allSurat.filter(function (s) {
    return (
      String(s.perihal).toLowerCase().includes(kw) ||
      String(s.pengirim).toLowerCase().includes(kw) ||
      String(s.penerima).toLowerCase().includes(kw) ||
      String(s.nomor_surat).toLowerCase().includes(kw) ||
      String(s.jenis_surat).toLowerCase().includes(kw) ||
      String(s.status).toLowerCase().includes(kw) ||
      String(s.nama_pembuat).toLowerCase().includes(kw)
    );
  });

  return { success: true, data: hasil };
}

// ============================================================
// GET ARSIP (status: disetujui atau arsip)
// ============================================================
function getArsip(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var user = session.user;
  var sheet = getSheet(SHEET_NAMES.SURAT);
  var allSurat = sheetToObjects(sheet);

  var arsip = allSurat.filter(function (s) {
    var isArsip = s.status === 'disetujui' || s.status === 'arsip';
    if (!isArsip) return false;

    // Filter sesuai role
    if (user.role === 'Admin') return true;
    if (user.role === 'Pimpinan') return true;
    return s.id_pembuat === user.userId;
  });

  // Auto-generate PDF jika ada surat yang disetujui tapi file_url-nya belum terisi
  try {
    var headers = sheet.getDataRange().getValues()[0];
    var idColIdx = headers.indexOf('id_surat');
    var fileUrlColIdx = headers.indexOf('file_url');
    var updatedAny = false;

    if (idColIdx >= 0 && fileUrlColIdx >= 0) {
      arsip.forEach(function (s) {
        if (!s.file_url && s.id_surat) {
          try {
            var resPdf = generateSuratPDF(s.id_surat);
            if (resPdf.success && resPdf.fileUrl) {
              s.file_url = resPdf.fileUrl;
              var allRows = sheet.getDataRange().getValues();
              for (var r = 1; r < allRows.length; r++) {
                if (allRows[r][idColIdx] === s.id_surat) {
                  sheet.getRange(r + 1, fileUrlColIdx + 1).setValue(resPdf.fileUrl);
                  updatedAny = true;
                  break;
                }
              }
            }
          } catch (ePdf) {
            Logger.log('Auto PDF on getArsip failed for ' + s.id_surat + ': ' + ePdf.toString());
          }
        }
      });
      if (updatedAny) {
        SpreadsheetApp.flush();
      }
    }
  } catch (errAuto) {
    Logger.log('Auto PDF loop error: ' + errAuto.toString());
  }

  arsip.sort(function (a, b) {
    return new Date(b.tanggal_ttd || b.tanggal_dibuat || 0) - new Date(a.tanggal_ttd || a.tanggal_dibuat || 0);
  });

  return { success: true, data: arsip };
}

// ============================================================
// REGENERATE PDF SURAT (Manual dari UI jika perlu)
// ============================================================
function regenerateSuratPDF(token, suratId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  if (!suratId) return { success: false, message: 'ID Surat tidak valid.' };

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  var idIdx = headers.indexOf('id_surat');
  var fileUrlIdx = headers.indexOf('file_url');

  var rowNum = -1;
  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === suratId) {
      rowNum = i + 1;
      break;
    }
  }

  if (rowNum === -1) {
    return { success: false, message: 'Surat tidak ditemukan.' };
  }

  try {
    var pdfRes = generateSuratPDF(suratId);
    if (pdfRes.success && pdfRes.fileUrl) {
      sheet.getRange(rowNum, fileUrlIdx + 1).setValue(pdfRes.fileUrl);
      SpreadsheetApp.flush();
      return {
        success: true,
        message: 'PDF berhasil dibuat!',
        fileUrl: pdfRes.fileUrl
      };
    } else {
      return { success: false, message: pdfRes.message || 'Gagal membuat PDF.' };
    }
  } catch (err) {
    return { success: false, message: 'Error membuat PDF: ' + err.toString() };
  }
}

// ============================================================
// GET STATISTIK DASHBOARD (Pimpinan & Admin)
// ============================================================
function getStatistik(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var sheet = getSheet(SHEET_NAMES.SURAT);
  var allSurat = sheetToObjects(sheet);

  var stats = {
    total: allSurat.length,
    masuk: allSurat.filter(function (s) { return s.jenis_surat === 'masuk'; }).length,
    keluar: allSurat.filter(function (s) { return s.jenis_surat === 'keluar'; }).length,
    permohonan: allSurat.filter(function (s) { return s.jenis_surat === 'permohonan'; }).length,
    sudahDiproses: allSurat.filter(function (s) { return s.status === 'disetujui' || s.status === 'arsip'; }).length,
    belumDiproses: allSurat.filter(function (s) { return s.status === 'draft' || s.status === 'verifikasi'; }).length,
    draft: allSurat.filter(function (s) { return s.status === 'draft'; }).length,
    verifikasi: allSurat.filter(function (s) { return s.status === 'verifikasi'; }).length,
    disetujui: allSurat.filter(function (s) { return s.status === 'disetujui'; }).length,
    arsip: allSurat.filter(function (s) { return s.status === 'arsip'; }).length
  };

  return { success: true, data: stats };
}

// ============================================================
// GET DAFTAR USER UNTUK DROPDOWN (penerima, disposisi, dll)
// ============================================================
function getUsersForDropdown(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };

  var sheet = getSheet(SHEET_NAMES.USERS);
  var users = sheetToObjects(sheet);

  var publicUsers = users
    .filter(function (u) { return u.status_approval === 'approved'; })
    .map(function (u) {
      return { id: u.id, nama: u.nama, role: u.role, username: u.username };
    });

  return { success: true, data: publicUsers };
}
