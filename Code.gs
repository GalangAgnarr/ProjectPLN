// --- KONFIGURASI UTAMA ---
const ADMIN_DEFAULT = {
  user: 'admin',
  pass: '123',
  nama: 'Administrator Sistem',
  role: 'Admin',
  email: 'admin@example.com'
};

// Nama Folder di Google Drive tempat file disimpan
const DRIVE_FOLDER_NAME = "Arsip Digital Uploads";

/**
 * 1. HTTP GET HANDLER
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Arsip Digital & BPM Survey - PT PLN (Persero)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. INITIALIZATION
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. Setup Sheet Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['Username', 'Password', 'Nama Lengkap', 'Role', 'Status', 'Token', 'Email', 'OTP']);
    userSheet.appendRow([
      ADMIN_DEFAULT.user,
      ADMIN_DEFAULT.pass,
      ADMIN_DEFAULT.nama,
      ADMIN_DEFAULT.role,
      'Active',
      '',
      ADMIN_DEFAULT.email,
      ''
    ]);
    userSheet.setFrozenRows(1);
  }

  // B. Setup Sheet ActivityLog
  let logSheet = ss.getSheetByName('ActivityLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('ActivityLog');
    logSheet.appendRow(['Waktu', 'User', 'Aktivitas', 'Detail']);
    logSheet.setFrozenRows(1);
  }

  // C. Setup Sheet Settings (Master Data)
  let setSheet = ss.getSheetByName('Settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('Settings');
    setSheet.appendRow(['Type', 'Value']);
    const defaults = [
      ['Category', 'PFK'],
      ['Category', 'ESTETIKA'],
      ['Category', 'Pelanggan TM'],
      ['Category', 'SPKLU'],
      ['Extension', 'PDF'],
      ['Extension', 'DOCX'],
      ['Extension', 'XLSX'],
      ['Extension', 'JPG'],
      ['Lemari', 'Lemari_A'],
      ['Lemari', 'Lemari_B'],
      ['Lemari', 'Lemari_C'],
      ['Ordner', 'Ordner_01'],
      ['Ordner', 'Ordner_02']
    ];
    defaults.forEach(row => setSheet.appendRow(row));
    setSheet.setFrozenRows(1);
  }

  // D. Setup Sheet Lemari Default
  const defaultLemariList = ['Lemari_A', 'Lemari_B', 'Lemari_C'];
  defaultLemariList.forEach(lemariName => {
    getOrCreateLemariSheet(ss, lemariName);
  });

  // E. Setup Sheet Notifications
  let notifSheet = ss.getSheetByName('Notifications');
  if (!notifSheet) {
    notifSheet = ss.insertSheet('Notifications');
    notifSheet.appendRow(['ID', 'ToUser', 'Message', 'Type', 'IsRead', 'RelatedId', 'Timestamp', 'FromUser']);
    notifSheet.setFrozenRows(1);
  }

  // F. Setup Sheet BPM
  getOrCreateBpmSheet(ss);

  return "Database berhasil diinisialisasi dan siap digunakan.";
}

/**
 * HELPER: Dapatkan atau buat Sheet khusus Lemari
 */
function getOrCreateLemariSheet(ss, rawLemariName) {
  let cleanName = (rawLemariName || 'Lemari_A').trim().replace(/\s+/g, '_');
  let sheet = ss.getSheetByName(cleanName);

  if (!sheet) {
    sheet = ss.insertSheet(cleanName);
    sheet.appendRow(['ID', 'Nomor Arsip', 'Nama Arsip', 'Perihal', 'Kategori', 'Lemari', 'Ordner', 'Jenis File', 'Link File', 'Tanggal Upload', 'Pengupload', 'SharedWith']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * HELPER: Dapatkan atau buat Sheet khusus BPM (Ditambahkan Header Link di Kolom 14 / N)
 */
function getOrCreateBpmSheet(ss) {
  let sheet = ss.getSheetByName('BPM');
  if (!sheet) {
    sheet = ss.insertSheet('BPM');
    sheet.appendRow([
      'ID', 'Kode', 'Judul', 'IsPFK', 'Step', 'StatusDetail', 'Tanggal', 'Pemohon',
      'Lokasi', 'Tgl_Pengajuan', 'Tgl_Survey', 'Tgl_Manajemen', 'Tgl_Selesai', 'Link'
    ]);
    sheet.setFrozenRows(1);
  } else {
    // Pastikan header Link di kolom 14 (N) jika sheet lama belum punya
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length < 14) {
      sheet.getRange(1, 14).setValue('Link');
    }
  }
  return sheet;
}

/**
 * HELPER: Dapatkan daftar seluruh nama Lemari dari Master Data
 */
function getLemariListFromMaster(ss) {
  const setSheet = ss.getSheetByName('Settings');
  let lemaris = [];
  if (setSheet) {
    const data = setSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Lemari' && data[i][1]) {
        lemaris.push(data[i][1].toString().trim().replace(/\s+/g, '_'));
      }
    }
  }
  return lemaris.length > 0 ? lemaris : ['Lemari_A', 'Lemari_B', 'Lemari_C'];
}

/**
 * 3. API ROUTER
 */
function apiHandler(action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (action) {
      case 'checkSession': return checkSession(ss, payload);
      case 'login': return loginUser(ss, payload);
      case 'logout': return logoutUser(ss, payload);
      case 'requestPasswordOTP': return requestOtp(ss, payload);
      case 'verifyOTPAndResetPassword': return resetPassword(ss, payload);
      case 'updateProfile': return updateUserProfile(ss, payload);
      case 'getStats': return getDashboardStats(ss, payload);
      case 'getData': return getData(ss, payload);
      case 'saveData': return saveData(ss, payload);
      case 'deleteData': return deleteData(ss, payload);
      case 'getSettings': return getSettings(ss, payload);
      case 'saveSetting': return saveSetting(ss, payload);
      case 'deleteSetting': return deleteSetting(ss, payload);
      case 'generateReport': return generateReport(ss, payload);
      case 'getNotifications': return getNotifications(ss, payload);
      case 'markNotificationsRead': return markRead(ss, payload);
      case 'markRead': return markRead(ss, payload);
      default:
        throw new Error("Action tidak dikenal: " + action);
    }
  } catch (e) {
    console.error("API Error [" + action + "]: " + e.message);
    return { success: false, status: 'error', message: e.message };
  }
}

// ==========================================
// LOGIKA BISNIS & AUTHENTICATION
// ==========================================

function checkSession(ss, { token }) {
  if (!token) return { status: 'success', data: { sessionActive: false } };
  try {
    const user = validateToken(ss, token);
    return { status: 'success', data: { sessionActive: true, user: user } };
  } catch (e) {
    return { status: 'success', data: { sessionActive: false } };
  }
}

function loginUser(ss, { username, password }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      if (data[i][4] !== 'Active') throw new Error("Akun Anda dinonaktifkan/suspend.");

      const token = Utilities.getUuid();
      sheet.getRange(i + 1, 6).setValue(token);

      logActivity(ss, username, 'Login', 'User berhasil login');
      return {
        status: 'success',
        data: {
          user: {
            username: data[i][0],
            nama_lengkap: data[i][2],
            role: data[i][3],
            token: token,
            email: data[i][6] || ''
          }
        }
      };
    }
  }
  throw new Error("Username atau Password salah.");
}

function requestOtp(ss, { email }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().toLowerCase() === email.toLowerCase()) {
      const otp = Math.floor(100000 + Math.random() * 900000).toString();
      sheet.getRange(i + 1, 8).setValue(otp);

      try {
        MailApp.sendEmail({
          to: email,
          subject: "[E-Arsip PLN] Kode Reset Password",
          htmlBody: `
            <h3>Permintaan Reset Password</h3>
            <p>Halo ${data[i][2]},</p>
            <p>Gunakan kode OTP berikut untuk me-reset kata sandi Anda:</p>
            <h2 style="background: #005C9A; color: #ffffff; padding: 10px 20px; display: inline-block; letter-spacing: 5px; border-radius: 8px;">${otp}</h2>
          `
        });
      } catch (e) {
        throw new Error("Gagal mengirim email verifikasi.");
      }
      found = true;
      break;
    }
  }

  if (!found) throw new Error("Email tidak terdaftar dalam sistem.");
  return { status: 'success', message: 'Kode OTP telah dikirim ke email Anda.' };
}

function resetPassword(ss, { email, otp, newPass }) {
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][6] && data[i][6].toString().toLowerCase() === email.toLowerCase()) {
      if (String(data[i][7]) === String(otp)) {
        sheet.getRange(i + 1, 2).setValue(newPass);
        sheet.getRange(i + 1, 8).setValue("");
        logActivity(ss, data[i][0], 'Reset Password', 'Sukses reset password via OTP');
        return { status: 'success', message: 'Password berhasil diubah.' };
      } else {
        throw new Error("Kode OTP salah atau kadaluarsa.");
      }
    }
  }
  throw new Error("User tidak ditemukan.");
}

function logoutUser(ss, { token }) {
  const user = validateToken(ss, token);
  if (user) {
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.username) {
        sheet.getRange(i + 1, 6).setValue("");
        break;
      }
    }
    logActivity(ss, user.username, 'Logout', 'User logout');
  }
  return { status: 'success' };
}

function updateUserProfile(ss, { token, nama_lengkap, password_lama, password_baru }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === user.username) {
      if (String(data[i][1]) !== String(password_lama)) {
        throw new Error("Password lama salah.");
      }

      sheet.getRange(i + 1, 3).setValue(nama_lengkap);
      if (password_baru && password_baru.trim() !== "") {
        sheet.getRange(i + 1, 2).setValue(password_baru);
      }

      logActivity(ss, user.username, 'Update Profil', 'User memperbarui profil/password mandiri');
      return {
        status: 'success',
        data: {
          user: {
            username: user.username,
            nama_lengkap: nama_lengkap,
            role: user.role,
            token: token,
            email: data[i][6]
          }
        }
      };
    }
  }
  throw new Error("Data user tidak ditemukan.");
}

// --- DASHBOARD STATS ---
function getDashboardStats(ss, { token }) {
  const user = validateToken(ss, token);
  const userSheet = ss.getSheetByName('Users');
  const logSheet = ss.getSheetByName('ActivityLog');

  const totalUsers = Math.max(0, userSheet.getLastRow() - 1);
  let myArsip = 0;
  let sharedWithMe = 0;
  let userVisibleTotal = 0;
  let totalArsipSystem = 0;
  let categoryCounts = {};

  const lemariList = getLemariListFromMaster(ss);

  lemariList.forEach(lemariName => {
    const lemariSheet = ss.getSheetByName(lemariName);
    if (lemariSheet) {
      const arsipData = lemariSheet.getDataRange().getValues();
      for (let i = 1; i < arsipData.length; i++) {
        const row = arsipData[i];
        if (!row[0]) continue;

        totalArsipSystem++;
        const uploader = row[10];
        const sharedVal = row[11] ? row[11].toString() : '';
        const category = row[4] || 'Tanpa Kategori';

        const isOwner = uploader === user.username;
        const isShared = sharedVal.includes(user.username) || sharedVal === 'Public';

        if (isOwner) myArsip++;
        if (isShared && !isOwner) sharedWithMe++;

        if (user.role === 'Admin' || isOwner || isShared) {
          userVisibleTotal++;
          categoryCounts[category] = (categoryCounts[category] || 0) + 1;
        }
      }
    }
  });

  let totalArsipDisplay = user.role === 'Admin' ? totalArsipSystem : userVisibleTotal;

  let recentLogs = [];
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    for (let i = logData.length - 1; i > 0; i--) {
      if (recentLogs.length >= 5) break;
      const logRow = logData[i];
      if (user.role === 'Admin' || logRow[1] === user.username) {
        recentLogs.push({
          waktu: formatDate(logRow[0]),
          user: logRow[1],
          aksi: logRow[2],
          detail: logRow[3]
        });
      }
    }
  }

  return {
    status: 'success',
    data: {
      totalArsip: totalArsipDisplay,
      totalUsers,
      myArsip,
      sharedWithMe,
      recentLogs,
      chartData: categoryCounts,
      role: user.role
    }
  };
}

// --- GET DATA (PARSING BPM DENGAN VARIABLE LINK) ---
function getData(ss, { token, type }) {
  const user = validateToken(ss, token);
  let result = [];

  if (type === 'users') {
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { status: 'success', data: [], list: [] };

    let rawData = sheet.getDataRange().getValues();
    rawData.shift(); // Hapus header
    
    rawData.forEach(row => {
      if (row[0]) {
        result.push({
          username: row[0],
          password: user.role === 'Admin' ? row[1] : '***',
          nama_lengkap: row[2],
          fullname: row[2],
          role: row[3],
          status: row[4],
          email: row[6] || ''
        });
      }
    });
  }
  else if (type === 'archives') {
    const lemariList = getLemariListFromMaster(ss);

    lemariList.forEach(lemariName => {
      const lemariSheet = ss.getSheetByName(lemariName);
      if (lemariSheet) {
        let rawData = lemariSheet.getDataRange().getValues();
        rawData.shift();
        rawData.forEach(row => {
          if (!row[0]) return;

          let isOwner = row[10] === user.username;
          let sharedVal = row[11] ? row[11].toString() : '';
          let isShared = sharedVal.includes(user.username) || sharedVal === 'Public';

          if (user.role === 'Admin' || isOwner || isShared) {
            let rawLinkString = row[8] ? row[8].toString() : '';
            let linkArray = rawLinkString ? rawLinkString.split(',').map(l => l.trim()).filter(Boolean) : [];

            result.push({
              id: row[0],
              nomor: row[1],
              nama: row[2],
              perihal: row[3],
              kategori: row[4],
              lemari: row[5] || lemariName,
              ordner: row[6] || '-',
              jenis: row[7],
              link: linkArray[0] || '#',
              links: linkArray,
              tanggal: formatDate(row[9]),
              uploader: row[10],
              shared: row[11]
            });
          }
        });
      }
    });
  }
  else if (type === 'bpm') {
    const bpmSheet = getOrCreateBpmSheet(ss);
    let rawData = bpmSheet.getDataRange().getValues();
    rawData.shift(); // Hapus header
    
    rawData.forEach(row => {
      if (row[0] || row[1] || row[2]) { // Cek kelayakan baris
        let safeIso = function(val) {
          if (!val) return null;
          let d = new Date(val);
          return isNaN(d.getTime()) ? null : d.toISOString();
        };

        // Membaca boolean IsPFK dengan aman dari string "TRUE"/"FALSE"
        let rawIsPfk = row[3];
        let isPfkVal = typeof rawIsPfk === 'string' 
          ? rawIsPfk.trim().toUpperCase() === 'TRUE' 
          : Boolean(rawIsPfk);

        // Membaca array link lampiran dari Kolom ke-14 (Kolom N)
        let rawLinkString = row[13] ? row[13].toString() : '';
        let linkArray = rawLinkString ? rawLinkString.split(',').map(l => l.trim()).filter(Boolean) : [];

        result.push({
          id: String(row[0] || Date.now()),
          kode: row[1] ? String(row[1]) : '-',
          judul: row[2] ? String(row[2]) : 'Permohonan Survey',
          isPFK: isPfkVal,
          step: parseInt(row[4]) || 1,
          statusDetail: row[5] ? String(row[5]) : 'Kirim Surat Permohonan Survey & RAB',
          tanggal: formatDate(row[6]),
          uploader: row[7] ? String(row[7]) : 'System',
          pemohon: row[7] ? String(row[7]) : 'System',
          lokasi: row[8] ? String(row[8]) : 'cemorokandang',
          tglPengajuan: safeIso(row[9]) || safeIso(row[6]) || new Date().toISOString(),
          tglSurvey: safeIso(row[10]),
          tglManajemen: safeIso(row[11]),
          tglSelesai: safeIso(row[12]),
          links: linkArray,
          link: linkArray[0] || '#'
        });
      }
    });
  }
  else if (type === 'logs') {
    const sheet = ss.getSheetByName('ActivityLog');
    if (sheet) {
      let rawData = sheet.getDataRange().getValues();
      rawData.shift();
      const logs = rawData.reverse().slice(0, 1000);
      logs.forEach(row => {
        if (user.role === 'Admin' || row[1] === user.username) {
          result.push({
            waktu: formatDate(row[0]),
            user: row[1],
            aksi: row[2],
            detail: row[3],
            timestamp: new Date(row[0]).getTime()
          });
        }
      });
    }
  }

  return {
    status: 'success',
    success: true,
    data: result,
    list: result,
    users: result
  };
}

// --- SAVE DATA (SUPPORT MULTIPLE FILES UPLOAD & DYNAMIC LINKS) ---
function saveData(ss, { token, type, data }) {
  const user = validateToken(ss, token);

  if (type === 'archives') {
    const id = data.id || Utilities.getUuid();
    const timestamp = new Date();
    const targetLemariName = (data.lemari || 'Lemari_A').trim().replace(/\s+/g, '_');
    const targetOrdner = (data.ordner || '-').trim();

    const targetSheet = getOrCreateLemariSheet(ss, targetLemariName);

    // 1. Ambil link yang sudah ada sebelumnya (apabila mode Edit)
    let finalLinks = Array.isArray(data.links) ? data.links : (Array.isArray(data.existingLinks) ? data.existingLinks : []);

    // 2. PENANGANAN MULTIPLE FILES (Diambil dari array `data.files` frontend)
    if (data.files && Array.isArray(data.files) && data.files.length > 0) {
      data.files.forEach(fileItem => {
        if (fileItem && fileItem.base64) {
          try {
            const uploadedUrl = uploadToDrive(fileItem, fileItem.name || data.nama, data.kategori, user.nama_lengkap);
            if (uploadedUrl && !finalLinks.includes(uploadedUrl)) {
              finalLinks.push(uploadedUrl);
            }
          } catch (e) {
            throw new Error("Gagal upload file ke Drive: " + e.message);
          }
        }
      });
    } 
    // 3. FALLBACK SINGLE FILE (Jika frontend hanya mengirim `data.fileObj`)
    else if (data.fileObj && data.fileObj.base64) {
      try {
        const uploadedUrl = uploadToDrive(data.fileObj, data.fileObj.name || data.nama, data.kategori, user.nama_lengkap);
        if (uploadedUrl && !finalLinks.includes(uploadedUrl)) {
          finalLinks.push(uploadedUrl);
        }
      } catch (e) {
        throw new Error("Gagal upload file ke Drive: " + e.message);
      }
    }

    // Validasi wajib memiliki minimal 1 lampiran
    if (finalLinks.length === 0) {
      throw new Error("Dokumen wajib memiliki setidaknya satu file lampiran.");
    }

    const combinedLinksStr = finalLinks.join(',');

    if (data.id) { // Mode Edit Arsip
      let found = false;
      const lemariList = getLemariListFromMaster(ss);

      for (let l = 0; l < lemariList.length; l++) {
        const currentSheet = ss.getSheetByName(lemariList[l]);
        if (!currentSheet) continue;

        const allData = currentSheet.getDataRange().getValues();
        for (let i = 1; i < allData.length; i++) {
          if (allData[i][0] == data.id) {
            const owner = allData[i][10];
            const sharedStatus = allData[i][11];
            const isOwner = owner === user.username;
            const isPublic = sharedStatus === 'Public';

            if (!isOwner) {
              if (user.role === 'Admin') {
                if (!isPublic) throw new Error("Akses Ditolak: File Private milik user lain.");
              } else {
                throw new Error("Anda tidak memiliki izin mengedit file ini.");
              }
            }

            const originalDate = allData[i][9];
            const originalUploader = allData[i][10];

            if (currentSheet.getName() !== targetSheet.getName()) {
              currentSheet.deleteRow(i + 1);
              targetSheet.appendRow([
                data.id, data.nomor, data.nama, data.perihal, data.kategori,
                targetLemariName, targetOrdner, data.jenis, combinedLinksStr, originalDate, originalUploader, data.shared || 'Public'
              ]);
            } else {
              currentSheet.getRange(i + 1, 1, 1, 12).setValues([[
                data.id, data.nomor, data.nama, data.perihal, data.kategori,
                targetLemariName, targetOrdner, data.jenis, combinedLinksStr, originalDate, originalUploader, data.shared || 'Public'
              ]]);
            }

            found = true;
            break;
          }
        }
        if (found) break;
      }
      if (!found) throw new Error("Data tidak ditemukan.");
      logActivity(ss, user.username, 'Edit Arsip', `Mengubah arsip: ${data.nama} (${finalLinks.length} lampiran)`);
    } else { // Mode Baru
      targetSheet.appendRow([
        id,
        data.nomor,
        data.nama,
        data.perihal,
        data.kategori,
        targetLemariName,
        targetOrdner,
        data.jenis,
        combinedLinksStr,
        timestamp,
        user.username,
        data.shared || 'Public'
      ]);
      logActivity(ss, user.username, 'Upload Arsip', `Menambah arsip: ${data.nama} dengan ${finalLinks.length} lampiran`);
    }

    if (data.shared && data.shared !== 'Public') {
      const recipients = data.shared.split(',').map(s => s.trim());
      recipients.forEach(targetUser => {
        if (targetUser && targetUser !== user.username) {
          createNotification(ss, {
            to: targetUser,
            from: user.username,
            msg: `membagikan dokumen "${data.nama}" kepada Anda.`,
            type: 'share',
            relId: id
          });
        }
      });
    }

  }
  else if (type === 'bpm') {
    const bpmSheet = getOrCreateBpmSheet(ss);
    const rows = bpmSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === (data.id || '').toString() || rows[i][1].toString() === (data.id || '').toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    const now = new Date();

    if (rowIndex > 0) { // Update status step BPM atau Upload Lampiran BPM
      let rowData = rows[rowIndex - 1];
      let stepVal = parseInt(data.step) || parseInt(rowData[4]) || 1;

      let tglPengajuan = rowData[9]  || rowData[6] || now;
      let tglSurvey    = rowData[10] || (stepVal >= 4 ? now : '');
      let tglManajemen = rowData[11] || (stepVal >= 7 ? now : '');
      let tglSelesai   = rowData[12] || (stepVal >= 10 ? now : '');

      // PENANGANAN UNTUK MULTIPLE / SINGLE FILE UPLOAD DARI MODAL DETAIL BPM
      const filesToUpload = (data.files && Array.isArray(data.files) && data.files.length > 0) ? data.files : (data.fileObj && data.fileObj.base64 ? [data.fileObj] : []);

      if (filesToUpload.length > 0) {
        try {
          let newUploadedUrls = [];

          filesToUpload.forEach(fObj => {
            // 1. Upload file fisik ke Drive
            const uploadedUrl = uploadToDrive(fObj, fObj.name || data.nama || rowData[2], 'BPM Survey', user.nama_lengkap);
            newUploadedUrls.push(uploadedUrl);

            // 2. Masukkan metadata berkas ke Sheet Kotak Arsip Digital (Lemari_A)
            const targetSheet = getOrCreateLemariSheet(ss, 'Lemari_A');
            const archiveId = Utilities.getUuid();
            const extName = (fObj.name || '').split('.').pop().toUpperCase() || 'PDF';

            targetSheet.appendRow([
              archiveId,
              rowData[1], // Menggunakan Kode BPM sebagai Nomor Arsip
              fObj.name || rowData[2],
              'Lampiran File Dokumen BPM untuk: ' + rowData[2],
              'BPM Survey',
              'Lemari_A',
              'Ordner_01',
              extName,
              uploadedUrl,
              now,
              user.username,
              'Public'
            ]);
          });

          // 3. GABUNGKAN LINK GOOGLE DRIVE KE SHEET BPM KOLOM N (KOLOM KE-14)
          let existingLink = bpmSheet.getRange(rowIndex, 14).getValue();
          let existingLinksArray = existingLink ? existingLink.toString().split(',').map(s => s.trim()) : [];
          let mergedLinks = existingLinksArray.concat(newUploadedUrls);

          bpmSheet.getRange(rowIndex, 14).setValue(mergedLinks.join(','));

          logActivity(ss, user.username, 'Upload Lampiran BPM', `Mengunggah ${newUploadedUrls.length} file lampiran untuk agenda BPM: ${rowData[2]}`);
        } catch (e) {
          throw new Error("Gagal upload file lampiran BPM ke Drive: " + e.message);
        }
      } else {
        // Update rutin status/step BPM
        bpmSheet.getRange(rowIndex, 1, 1, 13).setValues([[
          data.id,
          data.kode || rowData[1],
          data.judul || rowData[2],
          data.isPFK !== undefined ? Boolean(data.isPFK) : rowData[3],
          stepVal,
          data.statusDetail || rowData[5],
          rowData[6] || now,
          rowData[7] || user.username,
          data.lokasi || rowData[8] || 'cemorokandang',
          tglPengajuan,
          tglSurvey,
          tglManajemen,
          tglSelesai
        ]]);

        logActivity(ss, user.username, 'Update BPM', `Memperbarui alur BPM: ${data.judul || rowData[2]} ke Step ${stepVal}`);
      }
    } else { // Permohonan BPM Baru
      const newId = Date.now().toString();
      bpmSheet.appendRow([
        newId,
        data.kode || ('PLN-PFK-' + Math.floor(10000000 + Math.random() * 90000000)),
        data.judul || 'Permohonan Survey',
        data.isPFK !== undefined ? Boolean(data.isPFK) : false,
        data.step || 1,
        data.statusDetail || 'Kirim Surat Permohonan Survey & RAB',
        now,
        user.username,
        data.lokasi || 'cemorokandang',
        now, // Tgl_Pengajuan
        '',  // Tgl_Survey
        '',  // Tgl_Manajemen
        '',  // Tgl_Selesai
        ''   // Link
      ]);
      logActivity(ss, user.username, 'Pengajuan BPM', `Membuat permohonan survey baru: ${data.judul}`);
    }
  }
  else if (type === 'users' && user.role === 'Admin') {
    const sheet = ss.getSheetByName('Users');
    if (!data.isEdit) {
      const users = sheet.getDataRange().getValues();
      if (users.some(u => u[0] === data.username)) throw new Error("Username sudah ada.");
      sheet.appendRow([data.username, data.password, data.nama_lengkap, data.role, data.status, '', data.email || '', '']);
      logActivity(ss, user.username, 'Add User', `Menambah user: ${data.username}`);
    } else {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === data.username) {
          sheet.getRange(i + 1, 2, 1, 4).setValues([[data.password, data.nama_lengkap, data.role, data.status]]);
          sheet.getRange(i + 1, 7).setValue(data.email || '');
          break;
        }
      }
      logActivity(ss, user.username, 'Edit User', `Mengubah user: ${data.username}`);
    }
  }

  return { status: 'success', success: true };
}

// --- DELETE DATA ---
function deleteData(ss, { token, type, id }) {
  const user = validateToken(ss, token);
  let deleted = false;
  let deletedName = '';

  if (type === 'users') {
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        deletedName = data[i][0];
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }
  } else if (type === 'archives') {
    const lemariList = getLemariListFromMaster(ss);
    for (let l = 0; l < lemariList.length; l++) {
      const currentSheet = ss.getSheetByName(lemariList[l]);
      if (!currentSheet) continue;

      const data = currentSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == id) {
          if (user.role !== 'Admin' && data[i][10] !== user.username) {
            throw new Error("Anda tidak berhak menghapus file ini.");
          }
          deletedName = data[i][2];
          const rawLinks = data[i][8] ? data[i][8].toString().split(',') : [];
          rawLinks.forEach(link => deleteFileFromDrive(link));

          currentSheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
      if (deleted) break;
    }
  } else if (type === 'bpm') {
    const bpmSheet = getOrCreateBpmSheet(ss);
    const data = bpmSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        deletedName = data[i][2];
        bpmSheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }
  }

  if (deleted) {
    logActivity(ss, user.username, 'Hapus Data', `Menghapus ${type}: ${deletedName}`);
    return { status: 'success', success: true };
  } else {
    throw new Error("Data tidak ditemukan.");
  }
}

// --- SETTINGS & MASTER DATA ---
function getSettings(ss, { token }) {
  const sheet = ss.getSheetByName('Settings');
  const userSheet = ss.getSheetByName('Users');

  let categories = [];
  let extensions = [];
  let nomors = [];
  let lemaris = [];
  let ordners = [];

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Category') categories.push(data[i][1]);
      if (data[i][0] === 'Extension') extensions.push(data[i][1]);
      if (data[i][0] === 'Nomor') nomors.push(data[i][1]);
      if (data[i][0] === 'Lemari') lemaris.push(data[i][1]);
      if (data[i][0] === 'Ordner') ordners.push(data[i][1]);
    }
  }

  let userDirectory = [];
  if (userSheet) {
    const users = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (users[i][4] === 'Active') {
        userDirectory.push({
          username: users[i][0],
          fullname: users[i][2] || users[i][0]
        });
      }
    }
  }

  const cleanLemaris = lemaris.length > 0 ? lemaris : ['Lemari_A', 'Lemari_B', 'Lemari_C'];
  const cleanOrdners = ordners.length > 0 ? ordners : ['Ordner_01', 'Ordner_02', 'Ordner_03'];
  cleanLemaris.forEach(l => getOrCreateLemariSheet(ss, l));

  return {
    status: 'success',
    data: {
      categories: categories.length > 0 ? categories : ['PFK', 'ESTETIKA', 'Pelanggan TM', 'SPKLU'],
      extensions,
      nomors,
      lemaris: cleanLemaris,
      ordners: cleanOrdners,
      userDirectory
    }
  };
}

function saveSetting(ss, { token, type, value }) {
  const user = validateToken(ss, token);
  if (user.role !== 'Admin') throw new Error("Akses Ditolak.");

  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  let cleanVal = value.toString().trim();

  if (type === 'Lemari') {
    cleanVal = cleanVal.replace(/\s+/g, '_');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1].toString().toLowerCase() === cleanVal.toLowerCase()) {
      throw new Error("Data master sudah ada.");
    }
  }

  sheet.appendRow([type, cleanVal]);

  if (type === 'Lemari') {
    getOrCreateLemariSheet(ss, cleanVal);
  }

  logActivity(ss, user.username, 'Update Setting', `Menambah ${type}: ${cleanVal}`);
  return { status: 'success' };
}

function deleteSetting(ss, { token, type, value }) {
  const user = validateToken(ss, token);
  if (user.role !== 'Admin') throw new Error("Akses Ditolak.");

  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1] === value) {
      sheet.deleteRow(i + 1);
      logActivity(ss, user.username, 'Update Setting', `Menghapus ${type}: ${value}`);
      return { status: 'success' };
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// --- REPORTING ---
function generateReport(ss, payload) {
  var tempSs = SpreadsheetApp.create("Temp_Report_All");
  var tempSheet = tempSs.getActiveSheet();

  tempSheet.appendRow(['ID', 'Nomor Arsip', 'Nama Arsip', 'Perihal', 'Kategori', 'Lemari', 'Ordner', 'Jenis File', 'Link File', 'Tanggal Upload', 'Pengupload', 'SharedWith']);

  var lemariList = getLemariListFromMaster(ss);
  var count = 0;

  lemariList.forEach(lemariName => {
    var sheet = ss.getSheetByName(lemariName);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row[0] || row[1]) {
          tempSheet.appendRow(row);
          count++;
        }
      }
    }
  });

  if (count === 0) {
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    return { success: true, data: { bytes: null, filename: null } };
  }

  SpreadsheetApp.flush();

  var url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  var bytes = Utilities.base64Encode(response.getContent());
  var filename = "Rekap_Seluruh_Laporan_Arsip.xlsx";

  DriveApp.getFileById(tempSs.getId()).setTrashed(true);

  return { success: true, data: { bytes: bytes, filename: filename } };
}

// --- NOTIFICATION SYSTEM ---
function createNotification(ss, { to, from, msg, type, relId }) {
  let sheet = ss.getSheetByName('Notifications');
  if (!sheet) {
    sheet = ss.insertSheet('Notifications');
    sheet.appendRow(['ID', 'ToUser', 'Message', 'Type', 'IsRead', 'RelatedId', 'Timestamp', 'FromUser']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([Utilities.getUuid(), to, msg, type, 'FALSE', relId, new Date(), from]);
}

function getNotifications(ss, { token }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Notifications');
  if (!sheet) return { status: 'success', data: { list: [], unread: 0 } };

  const data = sheet.getDataRange().getValues();
  let notifs = [];
  let unreadCount = 0;
  let count = 0;

  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === user.username) {
      const isRead = data[i][4] === true || String(data[i][4]).toUpperCase() === 'TRUE';
      if (!isRead) unreadCount++;

      if (count < 20) {
        notifs.push({
          id: data[i][0],
          message: data[i][2],
          type: data[i][3],
          isRead: isRead,
          relatedId: data[i][5],
          timestamp: formatDate(data[i][6]),
          from: data[i][7]
        });
        count++;
      }
    }
  }

  return { status: 'success', data: { list: notifs, unread: unreadCount } };
}

function markRead(ss, { token, notifId, all }) {
  const user = validateToken(ss, token);
  const sheet = ss.getSheetByName('Notifications');
  if (!sheet) return { status: 'success' };

  const data = sheet.getDataRange().getValues();

  if (all || notifId === 'all') {
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === user.username && String(data[i][4]).toUpperCase() !== 'TRUE') {
        sheet.getRange(i + 1, 5).setValue('TRUE');
      }
    }
  } else {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == notifId && data[i][1] === user.username) {
        sheet.getRange(i + 1, 5).setValue('TRUE');
        break;
      }
    }
  }
  return { status: 'success' };
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function validateToken(ss, token) {
  if (!token) throw new Error("Sesi tidak valid.");
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === token) {
      if (data[i][4] === 'Active') {
        return {
          username: data[i][0],
          nama_lengkap: data[i][2],
          role: data[i][3],
          token: token,
          email: data[i][6] || ''
        };
      } else {
        throw new Error("Akun dinonaktifkan.");
      }
    }
  }
  throw new Error("Sesi kadaluarsa.");
}

function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function uploadToDrive(fileObj, fileName, category, uploaderFullName) {
  const rootIterator = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let rootFolder = rootIterator.hasNext() ? rootIterator.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  const now = new Date();
  const yearFolder = getOrCreateSubFolder(rootFolder, now.getFullYear().toString());

  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const monthIndex = now.getMonth();
  const monthPrefix = (monthIndex + 1).toString().padStart(2, '0');
  const monthFolder = getOrCreateSubFolder(yearFolder, `${monthPrefix} - ${months[monthIndex]}`);

  const categoryFolder = getOrCreateSubFolder(monthFolder, category ? category.trim() : "Umum");
  const finalFolder = getOrCreateSubFolder(categoryFolder, uploaderFullName ? uploaderFullName.trim() : "Anonim");

  const decoded = Utilities.base64Decode(fileObj.base64);
  const blob = Utilities.newBlob(decoded, fileObj.mimeType, fileName);
  const file = finalFolder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function logActivity(ss, user, action, detail) {
  const sheet = ss.getSheetByName('ActivityLog');
  if (sheet) {
    sheet.appendRow([new Date(), user, action, detail]);
  }
}

function formatDate(date) {
  if (!date) return '-';
  try {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  } catch (e) {
    return date.toString();
  }
}

function deleteFileFromDrive(fileUrl) {
  if (!fileUrl) return;
  try {
    const idMatch = fileUrl.match(/[-\w]{25,}/);
    if (idMatch) {
      const file = DriveApp.getFileById(idMatch[0]);
      file.setTrashed(true);
    }
  } catch (e) {
    console.warn("Gagal menghapus file fisik: " + e.message);
  }
}

function requestPermissions() {
  const tempSheet = SpreadsheetApp.create("Dummy_Permission_Trigger");
  DriveApp.getRootFolder();
  UrlFetchApp.fetch("https://www.google.com");
  DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
  console.log("Izin berhasil diperbarui!");
}
