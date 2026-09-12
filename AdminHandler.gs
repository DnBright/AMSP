// ============================================================
// AdminHandler.gs — Kelola User (Admin Only)
// ============================================================

// ============================================================
// GET SEMUA USER
// ============================================================
function getUsers(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  var sheet = getSheet(SHEET_NAMES.USERS);
  var users = sheetToObjects(sheet);

  // Jangan kirim password ke frontend
  var safeUsers = users.map(function (u) {
    return {
      id: u.id,
      nama: u.nama,
      username: u.username,
      role: u.role,
      status_approval: u.status_approval,
      tanggal_daftar: u.tanggal_daftar,
      email: u.email || ''
    };
  });

  return { success: true, data: safeUsers };
}

// ============================================================
// GET USER PENDING APPROVAL
// ============================================================
function getPendingUsers(token) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  var sheet = getSheet(SHEET_NAMES.USERS);
  var users = sheetToObjects(sheet);

  var pending = users
    .filter(function (u) { return u.status_approval === 'pending'; })
    .map(function (u) {
      return {
        id: u.id,
        nama: u.nama,
        username: u.username,
        role: u.role,
        status_approval: u.status_approval,
        tanggal_daftar: u.tanggal_daftar
      };
    });

  return { success: true, data: pending };
}

// ============================================================
// CREATE USER (Admin langsung approved)
// ============================================================
function createUser(token, data) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  if (!data.nama || !data.username || !data.password || !data.role) {
    return { success: false, message: 'Nama, username, password, dan role wajib diisi.' };
  }

  var sheet = getSheet(SHEET_NAMES.USERS);
  var users = sheetToObjects(sheet);

  // Cek duplikasi username
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].username).toLowerCase() === String(data.username).toLowerCase()) {
      return { success: false, message: 'Username sudah digunakan.' };
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
    'approved',     // Admin buat → langsung aktif
    now.toISOString(),
    data.email || ''
  ]);

  return { success: true, message: 'Akun "' + data.nama + '" berhasil dibuat.', userId: userId };
}

// ============================================================
// UPDATE USER
// ============================================================
function updateUser(token, userId, data) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  var sheet = getSheet(SHEET_NAMES.USERS);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];

  var idIdx = headers.indexOf('id');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === userId) {
      var rowNum = i + 1;
      var editableFields = ['nama', 'username', 'password', 'role', 'email'];
      editableFields.forEach(function (field) {
        if (data[field] !== undefined && data[field] !== '') {
          var colIdx = headers.indexOf(field);
          if (colIdx >= 0) sheet.getRange(rowNum, colIdx + 1).setValue(data[field]);
        }
      });
      return { success: true, message: 'Data user berhasil diperbarui.' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ============================================================
// DELETE USER
// ============================================================
function deleteUser(token, userId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  // Tidak boleh hapus diri sendiri
  if (userId === session.user.userId) {
    return { success: false, message: 'Tidak dapat menghapus akun sendiri.' };
  }

  var sheet = getSheet(SHEET_NAMES.USERS);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  var idIdx = headers.indexOf('id');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === userId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'User berhasil dihapus.' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ============================================================
// APPROVE USER
// ============================================================
function approveUser(token, userId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  return setUserApprovalStatus_(userId, 'approved', 'User berhasil disetujui dan dapat login.');
}

// ============================================================
// REJECT USER
// ============================================================
function rejectUser(token, userId) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  return setUserApprovalStatus_(userId, 'rejected', 'User berhasil ditolak.');
}

function setUserApprovalStatus_(userId, status, successMsg) {
  var sheet = getSheet(SHEET_NAMES.USERS);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  var idIdx = headers.indexOf('id');
  var statusIdx = headers.indexOf('status_approval');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === userId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(status);
      return { success: true, message: successMsg };
    }
  }
  return { success: false, message: 'User tidak ditemukan.' };
}

// ============================================================
// RESET PASSWORD (Admin)
// ============================================================
function resetPassword(token, userId, newPassword) {
  var session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  if (!newPassword || newPassword.length < 6) {
    return { success: false, message: 'Password baru minimal 6 karakter.' };
  }

  var sheet = getSheet(SHEET_NAMES.USERS);
  var sheetData = sheet.getDataRange().getValues();
  var headers = sheetData[0];
  var idIdx = headers.indexOf('id');
  var passIdx = headers.indexOf('password');

  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][idIdx] === userId) {
      sheet.getRange(i + 1, passIdx + 1).setValue(newPassword);
      return { success: true, message: 'Password berhasil direset.' };
    }
  }
  return { success: false, message: 'User tidak ditemukan.' };
}
