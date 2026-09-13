// ============================================================
// Auth.gs — Autentikasi, Session Management, Registrasi
// ============================================================

// Durasi sesi: 6 jam (dalam detik)
var SESSION_DURATION = 21600;

// ============================================================
// LOGIN
// ============================================================
function login(username, password, role) {
  try {
    if (!username || !password || !role) {
      return { success: false, message: 'Semua field wajib diisi.' };
    }

    var sheet = getSheet(SHEET_NAMES.USERS);
    var users = sheetToObjects(sheet);
    var found = null;

    for (var i = 0; i < users.length; i++) {
      if (
        String(users[i].username).trim() === String(username).trim() &&
        String(users[i].password) === String(password) &&
        String(users[i].role) === String(role)
      ) {
        found = users[i];
        break;
      }
    }

    if (!found) {
      return { success: false, message: 'Username, password, atau role tidak sesuai.' };
    }

    if (found.status_approval !== 'approved') {
      if (found.status_approval === 'pending') {
        return { success: false, message: 'Akun Anda sedang menunggu persetujuan Admin. Harap bersabar.' };
      }
      if (found.status_approval === 'rejected') {
        return { success: false, message: 'Akun Anda ditolak. Hubungi Administrator.' };
      }
      return { success: false, message: 'Akun Anda belum aktif. Hubungi Administrator.' };
    }

    // Buat session token unik
    var token = Utilities.getUuid();
    var sessionData = JSON.stringify({
      token: token,
      userId: found.id,
      nama: found.nama,
      username: found.username,
      role: found.role,
      email: found.email || '',
      loginTime: new Date().toISOString()
    });

    // Simpan di CacheService (max 6 jam)
    CacheService.getScriptCache().put('sess_' + token, sessionData, SESSION_DURATION);

    return {
      success: true,
      token: token,
      userId: found.id,
      nama: found.nama,
      username: found.username,
      role: found.role,
      message: 'Login berhasil. Selamat datang, ' + found.nama + '!'
    };

  } catch (e) {
    Logger.log('Login error: ' + e.toString());
    return { success: false, message: 'Terjadi kesalahan sistem. Coba lagi.' };
  }
}

// ============================================================
// VALIDATE SESSION
// ============================================================
function validateSession(token) {
  if (!token) return { valid: false, message: 'Token tidak ditemukan.' };

  try {
    var cached = CacheService.getScriptCache().get('sess_' + token);
    if (!cached) {
      return { valid: false, message: 'Sesi telah berakhir. Silakan login kembali.' };
    }
    var sessionData = JSON.parse(cached);
    return { valid: true, user: sessionData };
  } catch (e) {
    return { valid: false, message: 'Sesi tidak valid.' };
  }
}

// ============================================================
// LOGOUT
// ============================================================
function logout(token) {
  try {
    if (token) CacheService.getScriptCache().remove('sess_' + token);
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

// ============================================================
// REGISTRASI USER BARU (Pegawai / Pimpinan)
// ============================================================
function registerUser(data) {
  try {
    if (!data.nama || !data.username || !data.password || !data.role) {
      return { success: false, message: 'Nama, username, password, dan role wajib diisi.' };
    }

    if (data.role === 'Admin') {
      return { success: false, message: 'Tidak dapat mendaftar sebagai Admin.' };
    }

    if (data.password.length < 6) {
      return { success: false, message: 'Password minimal 6 karakter.' };
    }

    var sheet = getSheet(SHEET_NAMES.USERS);
    var users = sheetToObjects(sheet);

    // Cek duplikasi username
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].username).trim().toLowerCase() === String(data.username).trim().toLowerCase()) {
        return { success: false, message: 'Username sudah digunakan, coba username lain.' };
      }
    }

    var userId = Utilities.getUuid();
    var now = new Date();

    sheet.appendRow([
      userId,
      data.nama,
      data.username.trim(),
      data.password,
      data.role,
      'pending',         // status_approval — menunggu Admin
      now.toISOString(),
      data.email || ''
    ]);

    return {
      success: true,
      message: 'Pendaftaran berhasil! Akun Anda sedang menunggu persetujuan Administrator.'
    };

  } catch (e) {
    Logger.log('Register error: ' + e.toString());
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() };
  }
}

// ============================================================
// GET CURRENT USER (dipanggil dari frontend setelah login)
// ============================================================
function getCurrentUser(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  return { success: true, user: session.user };
}

// ============================================================
// GANTI PASSWORD (oleh user sendiri)
// ============================================================
function changePassword(token, passwordLama, passwordBaru) {
  try {
    var session = validateSession(token);
    if (!session.valid) return { success: false, message: session.message };
    var user = session.user;

    if (!passwordLama || !passwordBaru) {
      return { success: false, message: 'Password lama dan password baru wajib diisi.' };
    }
    if (passwordBaru.length < 6) {
      return { success: false, message: 'Password baru minimal 6 karakter.' };
    }

    var sheet = getSheet(SHEET_NAMES.USERS);
    var sheetData = sheet.getDataRange().getValues();
    var headers = sheetData[0];
    var idIdx = headers.indexOf('id');
    var passIdx = headers.indexOf('password');

    for (var i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][idIdx]) === String(user.userId)) {
        var currentPass = String(sheetData[i][passIdx]);
        if (currentPass !== String(passwordLama)) {
          return { success: false, message: 'Password lama tidak sesuai.' };
        }
        sheet.getRange(i + 1, passIdx + 1).setValue(passwordBaru);
        SpreadsheetApp.flush();
        return { success: true, message: 'Password berhasil diubah. Silakan login ulang.' };
      }
    }
    return { success: false, message: 'Akun tidak ditemukan.' };
  } catch (e) {
    Logger.log('changePassword error: ' + e.toString());
    return { success: false, message: 'Terjadi kesalahan: ' + e.toString() };
  }
}
