// ============================================================
// PDFHandler.gs — Generate PDF Surat + Embed QR Code
// ============================================================

var PDF_FOLDER_NAME = 'AMSP_Surat_PDF';

// ============================================================
// GENERATE PDF SURAT (dipanggil saat Pimpinan Sign/Setujui)
// ============================================================
function generateSuratPDF(suratId) {
  try {
    var suratResult = getSuratById(suratId);
    if (!suratResult.success) return { success: false, message: 'Surat tidak ditemukan.' };

    var surat = suratResult.data;
    var webAppUrl = getWebAppUrl();
    var qrUrl = webAppUrl + '?scan_ttd=' + suratId;

    // Nama file dokumen
    var safeName = String(surat.nomor_surat || suratId).replace(/[\/\\:*?"<>|]/g, '-');
    var docName = 'AMSP_' + safeName;

    // Buat Google Doc baru sebagai template surat
    var doc = DocumentApp.create(docName);
    var body = doc.getBody();

    // === MARGIN HALAMAN ===
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

    // ── WATERMARK "AMSP BENGKEL" (PERSIS SEPERTI CONTOH FOTO) ──
    try {
      if (typeof getWatermarkBlob === 'function') {
        var wmBlob = getWatermarkBlob();
        if (wmBlob) {
          var firstP = body.getParagraphs()[0] || body.appendParagraph('');
          var posImg = firstP.addPositionedImage(wmBlob);
          posImg.setWidth(460);
          posImg.setHeight(460);
          posImg.setLeftOffset(20);
          posImg.setTopOffset(170);
          posImg.setLayout(DocumentApp.PositionedLayout.ABOVE_TEXT);
        }
      }
    } catch (eWm) {
      Logger.log('Watermark notice (non-fatal): ' + eWm.toString());
    }

    // ── KOP SURAT POJOK KANAN ATAS (PERSIS SEPERTI CONTOH) ──
    var kopTable = body.appendTable([
      ['', '', '']
    ]);
    kopTable.setBorderWidth(0);
    var kopRow = kopTable.getRow(0);
    kopRow.getCell(0).setWidth(260); // Area kosong kiri
    kopRow.getCell(1).setWidth(46);  // Kolom logo
    kopRow.getCell(2).setWidth(174); // Kolom teks dinas

    // Logo Kota Ambon
    try {
      var logoUrl = (typeof LOGO_PEMKOT_AMBON !== 'undefined' && LOGO_PEMKOT_AMBON)
        ? LOGO_PEMKOT_AMBON
        : 'https://upload.wikimedia.org/wikipedia/commons/e/e7/Lambang_Kota_Ambon.png';
      var resLogo = UrlFetchApp.fetch(logoUrl, { muteHttpExceptions: true });
      if (resLogo.getResponseCode() === 200) {
        var pLogoCell = kopRow.getCell(1).getChild(0).asParagraph();
        pLogoCell.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var logoImg = pLogoCell.appendInlineImage(resLogo.getBlob());
        logoImg.setWidth(42).setHeight(42);
      }
    } catch (eLogo) {
      Logger.log('Logo header error: ' + eLogo.toString());
    }

    // Teks Kop: PEMERINTAH KOTA AMBON / DINAS PUPR KOTA AMBON
    var pTeksKop = kopRow.getCell(2).getChild(0).asParagraph();
    pTeksKop.setText('PEMERINTAH\nKOTA AMBON');
    pTeksKop.editAsText().setFontFamily('Arial').setFontSize(11).setBold(true);
    pTeksKop.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    var pDinas = kopRow.getCell(2).appendParagraph('DINAS PUPR KOTA AMBON');
    pDinas.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true);
    pDinas.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    body.appendParagraph('');
    body.appendParagraph('');

    // ── METADATA SURAT & TANGGAL (KIRI & KANAN) ──
    var tglSuratIndo = formatTanggalIndo_(surat.tanggal);
    var teksAmbon = tglSuratIndo ? ('Ambon, ' + tglSuratIndo) : 'Ambon, ................ 2026';

    var metaTable = body.appendTable([
      ['Nomor',    ': ' + (surat.nomor_surat || '............................................'), teksAmbon],
      ['Lampiran', ': ' + (surat.lampiran || '............................................'), ''],
      ['Sifat',    ': ' + (surat.sifat || 'Biasa/Rahasia/Penting/Segera'), ''],
      ['Hal',      ': ' + (surat.perihal || '............................................'), '']
    ]);
    metaTable.setBorderWidth(0);
    for (var m = 0; m < metaTable.getNumRows(); m++) {
      var rMeta = metaTable.getRow(m);
      rMeta.getCell(0).setWidth(70);
      rMeta.getCell(1).setWidth(240);
      rMeta.getCell(2).setWidth(170);
      rMeta.getCell(0).editAsText().setFontFamily('Arial').setFontSize(10);
      rMeta.getCell(1).editAsText().setFontFamily('Arial').setFontSize(10);
      rMeta.getCell(2).editAsText().setFontFamily('Arial').setFontSize(10);
      if (m === 0) {
        rMeta.getCell(2).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }
    }

    body.appendParagraph('');

    // ── TUJUAN SURAT ──
    body.appendParagraph('Kepada');
    body.appendParagraph('');
    var pYth = body.appendParagraph('Yth. ' + (surat.penerima || '............................................'));
    pYth.editAsText().setFontFamily('Arial').setFontSize(10);
    var pDi = body.appendParagraph('Di -');
    pDi.editAsText().setFontFamily('Arial').setFontSize(10);
    var pKota = body.appendParagraph('     Ambon');
    pKota.editAsText().setFontFamily('Arial').setFontSize(10);

    body.appendParagraph('');

    // ── PARAGRAF PEMBUKA ──
    var pBuka = body.appendParagraph(
      'Dalam rangka mendukung kelancaran pelaksanaan tugas dan menjaga tertib administrasi pada Dinas Pekerjaan Umum dan Penataan Ruang Kota Ambon, bersama ini kami menyampaikan informasi sebagai berikut:'
    );
    pBuka.editAsText().setFontFamily('Arial').setFontSize(10);
    pBuka.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

    body.appendParagraph('');

    // ── TABEL RINCIAN KEGIATAN ──
    var hariTgl = formatHariTanggalIndo_(surat.tanggal) || '............................................';
    var waktuVal = surat.waktu || '............................................';
    var tempatVal = surat.tempat || '............................................';
    var kegiatanVal = surat.perihal || surat.isi_surat || '............................................';

    var infoTable = body.appendTable([
      ['Hari / Tanggal',     ': ' + hariTgl],
      ['Waktu',              ': ' + waktuVal],
      ['Tempat / Wilayah',   ': ' + tempatVal],
      ['Kegiatan/Pekerjaan', ': ' + kegiatanVal]
    ]);
    infoTable.setBorderWidth(0);
    for (var k = 0; k < infoTable.getNumRows(); k++) {
      var rInfo = infoTable.getRow(k);
      rInfo.getCell(0).setWidth(150);
      rInfo.getCell(1).setWidth(330);
      rInfo.getCell(0).editAsText().setFontFamily('Arial').setFontSize(10);
      rInfo.getCell(1).editAsText().setFontFamily('Arial').setFontSize(10);
    }

    body.appendParagraph('');

    // ── PARAGRAF PENUTUP ──
    var pTutup1 = body.appendParagraph(
      'Untuk menunjang pelaksanaan kegiatan tersebut, seluruh pihak terkait diharapkan dapat melakukan koordinasi dan menyiapkan kebutuhan administrasi maupun teknis sesuai dengan ketentuan yang berlaku.'
    );
    pTutup1.editAsText().setFontFamily('Arial').setFontSize(10);
    pTutup1.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

    body.appendParagraph('');
    var pTutup2 = body.appendParagraph(
      'Demikian disampaikan untuk menjadi perhatian dan dapat dilaksanakan sebagaimana mestinya. Atas perhatian dan kerja samanya, kami ucapkan terima kasih.'
    );
    pTutup2.editAsText().setFontFamily('Arial').setFontSize(10);
    pTutup2.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

    body.appendParagraph('');

    // ── BLOK TANDA TANGAN & TEMBUSAN (2 Kolom Bawah) ──
    var qrBlob = null;
    if (surat.nama_pimpinan_ttd) {
      try {
        var qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=140x140&margin=2&data='
          + encodeURIComponent(qrUrl);
        var resQr = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
        if (resQr.getResponseCode() === 200) {
          qrBlob = resQr.getBlob().setName('qrcode.png');
        } else {
          var fallbackQrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(qrUrl) + '&size=140';
          var resQr2 = UrlFetchApp.fetch(fallbackQrUrl, { muteHttpExceptions: true });
          if (resQr2.getResponseCode() === 200) {
            qrBlob = resQr2.getBlob().setName('qrcode.png');
          }
        }
      } catch (qrErr) {
        Logger.log('QR Code fetch error: ' + qrErr.toString());
      }
    }

    var bottomTable = body.appendTable([
      ['', '']
    ]);
    bottomTable.setBorderWidth(0);
    var bRow = bottomTable.getRow(0);
    var leftCell = bRow.getCell(0);  // Tembusan
    var rightCell = bRow.getCell(1); // Tanda Tangan Pimpinan

    leftCell.setWidth(190);
    rightCell.setWidth(290);

    // Tanda Tangan Pimpinan (Kanan)
    var pJabatan1 = rightCell.getChild(0).asParagraph();
    pJabatan1.setText('KEPALA UPTD PERBENGKELAN DAN\nPERLENGKAPAN KENDARAAN DINAS PUPR KOTA\nAMBON');
    pJabatan1.editAsText().setFontFamily('Arial').setFontSize(9.5).setBold(true);
    pJabatan1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    if (surat.nama_pimpinan_ttd) {
      if (qrBlob) {
        var pQr = rightCell.appendParagraph('');
        pQr.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var imgQr = pQr.appendInlineImage(qrBlob);
        imgQr.setWidth(85).setHeight(85);
      } else {
        rightCell.appendParagraph('');
        rightCell.appendParagraph('');
      }

      var pNama = rightCell.appendParagraph(surat.nama_pimpinan_ttd);
      pNama.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setUnderline(true);
      pNama.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      var pNip = rightCell.appendParagraph(surat.nip ? ('NIP. ' + surat.nip) : 'NIP. ........................................');
      pNip.editAsText().setFontFamily('Arial').setFontSize(9.5);
      pNip.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else {
      rightCell.appendParagraph('');
      rightCell.appendParagraph('');
      rightCell.appendParagraph('');
      var pDraftNama = rightCell.appendParagraph('..................................................');
      pDraftNama.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true);
      pDraftNama.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      var pDraftNip = rightCell.appendParagraph('NIP. ........................................');
      pDraftNip.editAsText().setFontFamily('Arial').setFontSize(9.5);
      pDraftNip.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    // Tembusan (Kiri Bawah)
    // Sejajarkan ke bawah sesuai posisi tanda tangan
    for (var s = 0; s < 5; s++) {
      leftCell.appendParagraph('');
    }
    var pTembusanTitle = leftCell.appendParagraph('Tembusan:');
    pTembusanTitle.editAsText().setFontFamily('Arial').setFontSize(9.5).setBold(false);

    if (surat.tembusan) {
      var listTembusan = String(surat.tembusan).split(/[,;\n]/);
      var idxT = 1;
      listTembusan.forEach(function(item) {
        var tItem = item.trim();
        if (tItem) {
          var pItem = leftCell.appendParagraph('  ' + idxT + '. ' + tItem);
          pItem.editAsText().setFontFamily('Arial').setFontSize(9);
          idxT++;
        }
      });
    } else {
      var pT1 = leftCell.appendParagraph('  1. ..........');
      var pT2 = leftCell.appendParagraph('  2. ..........');
      var pT3 = leftCell.appendParagraph('  3. ..........');
      pT1.editAsText().setFontFamily('Arial').setFontSize(9);
      pT2.editAsText().setFontFamily('Arial').setFontSize(9);
      pT3.editAsText().setFontFamily('Arial').setFontSize(9);
    }

    doc.saveAndClose();

    // ── EXPORT KE PDF ──
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf').setName(docName + '.pdf');

    // Simpan ke folder AMSP di Drive
    var folder = getOrCreateFolder_(PDF_FOLDER_NAME);
    var pdfFile = folder.createFile(pdfBlob);

    // Set akses siapa saja dengan link bisa lihat
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      Logger.log('Set sharing notice: ' + eShare.toString());
    }

    var fileUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view?usp=sharing';

    // Hapus Google Doc sementara
    try {
      docFile.setTrashed(true);
    } catch (eTrash) {
      Logger.log('Set trashed notice: ' + eTrash.toString());
    }

    Logger.log('✅ PDF berhasil dibuat: ' + fileUrl);
    return { success: true, fileUrl: fileUrl, fileId: pdfFile.getId() };

  } catch (e) {
    Logger.log('generateSuratPDF ERROR: ' + e.toString() + '\n' + e.stack);
    return { success: false, message: 'Gagal generate PDF: ' + e.toString() };
  }
}

// ============================================================
// GET URL QR CODE
// ============================================================
function getQRCodeImageUrl(suratId) {
  var webAppUrl = getWebAppUrl();
  var qrUrl = webAppUrl + '?scan_ttd=' + suratId;
  return 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&margin=2&data=' + encodeURIComponent(qrUrl);
}

// ============================================================
// HELPER: Format Tanggal Indonesia
// ============================================================
function formatTanggalIndo_(dateVal) {
  if (!dateVal) return '';
  try {
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return '';
    var bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
  } catch(e) { return ''; }
}

function formatHariTanggalIndo_(dateVal) {
  if (!dateVal) return '';
  try {
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return '';
    var hari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    var bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    return hari[d.getDay()] + ' / ' + d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
  } catch(e) { return ''; }
}

// ============================================================
// GET DATA SURAT UNTUK HALAMAN SCAN (Publik)
// ============================================================
function getSuratForScan(suratId) {
  try {
    if (!suratId) return { success: false, message: 'ID Surat tidak valid.' };

    var suratResult = getSuratById(suratId);
    if (!suratResult.success) {
      return { success: false, message: 'Surat tidak ditemukan atau tidak valid.' };
    }

    var surat = suratResult.data;

    // Hanya tampilkan jika sudah disetujui
    if (surat.status !== 'disetujui' && surat.status !== 'arsip') {
      return { success: false, message: 'Surat ini belum ditandatangani secara digital.' };
    }

    // Kirim notifikasi SCAN ke pimpinan
    if (surat.id_pimpinan_ttd) {
      var scanTime = formatDate(new Date());
      createNotifikasi(
        surat.id_pimpinan_ttd,
        '⚠️ QR Code tanda tangan Anda pada surat "' + surat.perihal + '" (No: ' + surat.nomor_surat + ') telah dipindai pada ' + scanTime + '. Harap periksa jika bukan Anda.',
        'qr_scan'
      );
    }

    // Cari username pimpinan terkait
    var usernamePimpinan = '';
    if (surat.id_pimpinan_ttd) {
      try {
        var userSheet = getSheet(SHEET_NAMES.USERS);
        var uData = userSheet.getDataRange().getValues();
        var uHeaders = uData[0];
        var uIdIdx = uHeaders.indexOf('id');
        var uUserIdx = uHeaders.indexOf('username');
        for (var k = 1; k < uData.length; k++) {
          if (uData[k][uIdIdx] === surat.id_pimpinan_ttd) {
            usernamePimpinan = uData[k][uUserIdx];
            break;
          }
        }
      } catch (errUser) {
        Logger.log('Error lookup username pimpinan: ' + errUser.toString());
      }
    }

    var webAppUrl = getWebAppUrl();

    return {
      success: true,
      data: {
        perihal: surat.perihal,
        nomor_surat: surat.nomor_surat,
        jenis_surat: surat.jenis_surat,
        tanggal: surat.tanggal,
        pengirim: surat.pengirim,
        penerima: surat.penerima,
        nama_pimpinan_ttd: surat.nama_pimpinan_ttd,
        username_pimpinan: usernamePimpinan,
        tanggal_ttd: surat.tanggal_ttd,
        status: surat.status,
        login_url: webAppUrl ? (webAppUrl + '?page=login&role=pimpinan') : ''
      }
    };

  } catch (e) {
    Logger.log('getSuratForScan ERROR: ' + e.toString());
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() };
  }
}

// ============================================================
// HELPER: Get or Create Google Drive Folder
// ============================================================
function getOrCreateFolder_(folderName) {
  try {
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(folderName);
  } catch (eF) {
    Logger.log('getOrCreateFolder_ error: ' + eF.toString());
    return DriveApp.getRootFolder();
  }
}

// ============================================================
// FUNGSI OTORISASI DRIVE & DOCS (Klik Jalankan di Editor Apps Script)
// ============================================================
function testOtorisasiDriveDanPDF() {
  Logger.log('1. Memeriksa izin Google Docs...');
  var testDoc = DocumentApp.create('AMSP_Test_Doc');
  var docId = testDoc.getId();
  Logger.log('Google Docs OK, ID: ' + docId);

  Logger.log('2. Memeriksa izin Google Drive...');
  var docFile = DriveApp.getFileById(docId);
  var folder = getOrCreateFolder_(PDF_FOLDER_NAME);
  Logger.log('Google Drive Folder OK: ' + folder.getName());

  try {
    docFile.setTrashed(true);
  } catch (e) {}

  Logger.log('✅ Otorisasi Google Drive dan Dokumen Berhasil!');
  return 'Otorisasi Berhasil!';
}
