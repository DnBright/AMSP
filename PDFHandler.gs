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

    // === MARGIN ===
    body.setMarginTop(40).setMarginBottom(50).setMarginLeft(72).setMarginRight(72);

    // ── KOP HEADER DENGAN LOGO ──
    try {
      var logoUrl = (typeof LOGO_PEMKOT_AMBON !== 'undefined' && LOGO_PEMKOT_AMBON)
        ? LOGO_PEMKOT_AMBON
        : 'https://upload.wikimedia.org/wikipedia/commons/e/e7/Lambang_Kota_Ambon.png';
      
      var res = UrlFetchApp.fetch(logoUrl, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        var logoBlob = res.getBlob();
        var pLogo = body.appendParagraph('');
        pLogo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var img = pLogo.appendInlineImage(logoBlob);
        img.setWidth(55).setHeight(55);
      }
    } catch (logoErr) {
      Logger.log('Logo kop surat error (non-fatal): ' + logoErr.toString());
    }

    var p1 = body.appendParagraph((APP_CONFIG.nama_instansi || 'PEMERINTAH KOTA AMBON').toUpperCase());
    p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p1.editAsText().setFontFamily('Arial').setFontSize(15).setBold(true);

    var p2 = body.appendParagraph(APP_CONFIG.nama_lengkap || 'Aplikasi Manajemen Surat Perintah');
    p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p2.editAsText().setFontFamily('Arial').setFontSize(10).setBold(false).setItalic(true);

    body.appendHorizontalRule();
    body.appendParagraph('');

    // ── JUDUL SURAT ──
    var labelJenis = { masuk: 'SURAT MASUK', keluar: 'SURAT KELUAR', permohonan: 'SURAT PERMOHONAN' };
    var pJudul = body.appendParagraph(labelJenis[surat.jenis_surat] || 'SURAT');
    pJudul.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pJudul.editAsText().setFontFamily('Arial').setFontSize(14).setBold(true);

    var pNomor = body.appendParagraph('Nomor: ' + (surat.nomor_surat || '-'));
    pNomor.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pNomor.editAsText().setFontFamily('Arial').setFontSize(11).setBold(false);

    body.appendParagraph('');

    // ── INFO SURAT (Tabel tanpa border) ──
    var tableData = [
      ['Perihal',   ': ' + (surat.perihal || '-')],
      ['Tanggal',   ': ' + formatDateOnly(surat.tanggal)],
      ['Pengirim',  ': ' + (surat.pengirim || '-')],
      ['Penerima',  ': ' + (surat.penerima || '-')],
      ['Tembusan',  ': ' + (surat.tembusan || '-')]
    ];

    var tbl = body.appendTable(tableData);
    tbl.setBorderWidth(0);
    for (var r = 0; r < tbl.getNumRows(); r++) {
      tbl.getRow(r).getCell(0).setWidth(110);
      tbl.getRow(r).editAsText().setFontFamily('Arial').setFontSize(11);
    }

    body.appendParagraph('');
    body.appendParagraph('');

    // ── ISI SURAT ──
    var isiText = surat.isi_surat ||
      'Dengan hormat,\n\nBerkenaan dengan perihal di atas, bersama surat ini kami sampaikan untuk ditindaklanjuti sebagaimana mestinya.\n\nDemikian surat ini kami sampaikan. Atas perhatian dan kerja sama yang baik, kami ucapkan terima kasih.';

    var pIsi = body.appendParagraph(isiText);
    pIsi.editAsText().setFontFamily('Arial').setFontSize(11);

    body.appendParagraph('');
    body.appendParagraph('');

    // ── TANDA TANGAN BERBASIS BARCODE / QR CODE ──
    if (surat.nama_pimpinan_ttd) {
      body.appendParagraph('');
      var pTtdTitle = body.appendParagraph('Hormat kami,');
      pTtdTitle.editAsText().setFontFamily('Arial').setFontSize(11);

      var pJabatan = body.appendParagraph('Pimpinan,');
      pJabatan.editAsText().setFontFamily('Arial').setFontSize(11).setBold(true);

      // Embed Barcode / QR Code Tanda Tangan Digital
      var qrBlob = null;
      try {
        var qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=140x140&margin=2&data='
          + encodeURIComponent(qrUrl);
        var resQr = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
        if (resQr.getResponseCode() === 200) {
          qrBlob = resQr.getBlob().setName('qrcode.png');
        } else {
          // Fallback ke quickchart.io
          var fallbackQrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(qrUrl) + '&size=140';
          var resQr2 = UrlFetchApp.fetch(fallbackQrUrl, { muteHttpExceptions: true });
          if (resQr2.getResponseCode() === 200) {
            qrBlob = resQr2.getBlob().setName('qrcode.png');
          }
        }
      } catch (qrErr) {
        Logger.log('QR Code fetch error: ' + qrErr.toString());
      }

      if (qrBlob) {
        var pQrImg = body.appendParagraph('');
        var imgElement = pQrImg.appendInlineImage(qrBlob);
        imgElement.setWidth(95).setHeight(95);
      } else {
        body.appendParagraph('');
        body.appendParagraph('');
      }

      var pNamaTtd = body.appendParagraph(surat.nama_pimpinan_ttd);
      pNamaTtd.editAsText().setFontFamily('Arial').setFontSize(11).setBold(true).setUnderline(true);

      var pEsign = body.appendParagraph('Ditandatangani secara elektronik (Digital Signature)');
      pEsign.editAsText().setFontFamily('Arial').setFontSize(8).setItalic(true);

      body.appendParagraph('');
      body.appendHorizontalRule();

      // ── FOOTER VERIFIKASI KEABSAHAN QR ──
      var pQrTitle = body.appendParagraph('🔒 Verifikasi Keaslian Tanda Tangan Digital');
      pQrTitle.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);

      var pQrNote = body.appendParagraph(
        'Surat ini telah disetujui dan ditandatangani secara digital oleh ' + surat.nama_pimpinan_ttd +
        ' pada ' + formatDate(surat.tanggal_ttd) + '.\n' +
        'Pindai (Scan) QR Code di atas untuk memverifikasi keaslian surat melalui portal resmi AMSP.'
      );
      pQrNote.editAsText().setFontFamily('Arial').setFontSize(8).setItalic(true);
    }

    doc.saveAndClose();

    // ── EXPORT KE PDF ──
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf').setName(docName + '.pdf');

    // Simpan ke folder AMSP di Drive
    var folder = getOrCreateFolder_(PDF_FOLDER_NAME);
    var pdfFile = folder.createFile(pdfBlob);

    // Set akses siapa saja dengan link bisa lihat
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view?usp=sharing';

    // Hapus Google Doc sementara
    docFile.setTrashed(true);

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
  return 'https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + encodeURIComponent(qrUrl) + '&choe=UTF-8';
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
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}
