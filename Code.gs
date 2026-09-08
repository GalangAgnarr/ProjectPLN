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

// --- KONFIGURASI KEAMANAN ---
const MAX_LOGIN_ATTEMPTS = 5;         // Percobaan login gagal sebelum akun dikunci sementara
const LOCK_DURATION_MINUTES = 15;     // Lama penguncian akun (menit)
const OTP_EXPIRY_MINUTES = 10;        // Lama kode OTP berlaku (menit)
const OTP_REQUEST_COOLDOWN_SECONDS = 60; // Jeda minimum antar permintaan OTP (detik)

// Urutan kolom lengkap sheet Users (ditambah kolom keamanan di akhir agar kompatibel mundur)
const USER_SHEET_HEADERS = [
  'Username', 'Password', 'Nama Lengkap', 'Role', 'Status', 'Token', 'Email', 'OTP',
  'Salt', 'OtpExpiry', 'FailedAttempts', 'LockUntil', 'LastOtpRequestAt'
];

// Pastikan sheet Users punya seluruh kolom keamanan (migrasi otomatis untuk sheet lama)
function ensureUserSheetSchema(ss) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const header = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (header.length < USER_SHEET_HEADERS.length) {
    sheet.getRange(1, 1, 1, USER_SHEET_HEADERS.length).setValues([USER_SHEET_HEADERS]);
  }
}

// Hash password dengan SHA-256 + salt unik per user
function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt), Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))).join('');
}

function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '');
}

// Deteksi apakah string sudah berupa hash SHA-256 (64 karakter hex) atau masih password polos (data lama)
function looksHashed(pw) {
  return typeof pw === 'string' && /^[a-f0-9]{64}$/i.test(pw);
}

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
    userSheet.appendRow(USER_SHEET_HEADERS);
    const adminSalt = generateSalt();
    userSheet.appendRow([
      ADMIN_DEFAULT.user,
      hashPassword(ADMIN_DEFAULT.pass, adminSalt),
      ADMIN_DEFAULT.nama,
      ADMIN_DEFAULT.role,
      'Active',
      '',
      ADMIN_DEFAULT.email,
      '',
      adminSalt,
      '', 0, '', ''
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
      ['Ordner', 'Ordner_02'],
      ['ULP', 'ULP Malang Kota'],
      ['ULP', 'ULP Blimbing'],
      ['ULP', 'ULP Dinoyo'],
      ['ULP', 'ULP Kebonagung'],
      ['ULP', 'ULP Singosari'],
      ['ULP', 'ULP Lawang'],
      ['ULP', 'ULP Batu'],
      ['ULP', 'ULP Tumpang'],
      ['ULP', 'ULP Bululawang'],
      ['ULP', 'ULP Gondanglegi'],
      ['ULP', 'ULP Kepanjen'],
      ['ULP', 'ULP Sumberpucung'],
      ['ULP', 'ULP Dampit']
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
 * HELPER: Dapatkan atau buat Sheet khusus BPM (TER-UPDATE DENGAN KOLOM FORM LENGKAP)
 */
function getOrCreateBpmSheet(ss) {
  let sheet = ss.getSheetByName('BPM');
  if (!sheet) {
    sheet = ss.insertSheet('BPM');
    sheet.appendRow([
      'ID', 'Kode', 'Judul', 'IsPFK', 'Step', 'StatusDetail', 'Tanggal', 'Pemohon',
      'Lokasi', 'Alamat', 'ULP', 'JumlahUnit', 'TarifDaya', 'NomorSurat', 'TanggalSurat', 'PIC',
      'Tgl_Pengajuan', 'Tgl_Survey', 'Tgl_Manajemen', 'Tgl_Selesai', 'StepTimestamps', 'Lampiran'
    ]);
    sheet.setFrozenRows(1);
  } else {
    // Pastikan header diperbarui jika sheet sudah ada sebelumnya
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headerRow.length < 22) {
      sheet.getRange(1, 1, 1, 22).setValues([[
        'ID', 'Kode', 'Judul', 'IsPFK', 'Step', 'StatusDetail', 'Tanggal', 'Pemohon',
        'Lokasi', 'Alamat', 'ULP', 'JumlahUnit', 'TarifDaya', 'NomorSurat', 'TanggalSurat', 'PIC',
        'Tgl_Pengajuan', 'Tgl_Survey', 'Tgl_Manajemen', 'Tgl_Selesai', 'StepTimestamps', 'Lampiran'
      ]]);
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

// --- STRUKTUR ARSIP BARU: TAHUN / BULAN (menggantikan Lemari sebagai pengelompokan utama) ---

const BULAN_INDONESIA = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function getBulanFromDate(dateVal) {
  try {
    const d = new Date(dateVal);
    if (isNaN(d.getTime())) return BULAN_INDONESIA[0];
    return BULAN_INDONESIA[d.getMonth()];
  } catch (e) {
    return BULAN_INDONESIA[0];
  }
}

function getOrCreateTahunSheet(ss, rawTahun) {
  let cleanName = (rawTahun || new Date().getFullYear().toString()).toString().trim().replace(/\s+/g, '_');
  let sheet = ss.getSheetByName(cleanName);

  if (!sheet) {
    sheet = ss.insertSheet(cleanName);
    sheet.appendRow(['ID', 'Nomor Arsip', 'Nama Arsip', 'Perihal', 'Kategori', 'Bulan', 'Ordner', 'Jenis File', 'Link File', 'Tanggal Upload', 'Pengupload', 'SharedWith']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getTahunListFromMaster(ss) {
  const setSheet = ss.getSheetByName('Settings');
  let tahuns = [];
  if (setSheet) {
    const data = setSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Tahun' && data[i][1]) {
        tahuns.push(data[i][1].toString().trim());
      }
    }
  }
  if (tahuns.length === 0) {
    const currentYear = new Date().getFullYear();
    tahuns = [(currentYear - 1).toString(), currentYear.toString(), (currentYear + 1).toString()];
  }
  // Urutkan dari tahun terbaru
  return tahuns.sort((a, b) => b.localeCompare(a));
}

/**
 * 3. API ROUTER
 */
// =============================================================================
// FUNGSI UTILITAS SEKALI-PAKAI — JANGAN dihubungkan ke apiHandler/web app.
// Fungsi ini SENGAJA tidak diberi 'case' di apiHandler supaya tidak bisa
// dipanggil dari luar (internet) oleh siapapun, demi keamanan.
//
// CARA PAKAI:
// 1. Buka project ini di Apps Script Editor (Extensions > Apps Script).
// 2. Di dropdown pilihan fungsi (atas, sebelah tombol Run/Debug), pilih:
//    resetAllPasswordsToDefault
// 3. Klik tombol "Run".
// 4. Setelah selesai, buka menu "Execution log" / "View > Logs" untuk
//    melihat konfirmasi berapa user yang berhasil direset.
// 5. Beri tahu semua user untuk login pakai password default di bawah,
//    lalu segera ganti sendiri lewat menu Profil setelah berhasil masuk.
// =============================================================================
function resetAllPasswordsToDefault() {
  const DEFAULT_PASSWORD = 'pln@123';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureUserSheetSchema(ss);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    Logger.log('Sheet "Users" tidak ditemukan.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // lewati baris kosong
    const salt = generateSalt();
    const hash = hashPassword(DEFAULT_PASSWORD, salt);
    const row = i + 1;

    sheet.getRange(row, 2).setValue(hash);   // Password (hash baru)
    sheet.getRange(row, 9).setValue(salt);   // Salt baru
    sheet.getRange(row, 11).setValue(0);     // FailedAttempts direset
    sheet.getRange(row, 12).setValue('');    // LockUntil dibuka
    sheet.getRange(row, 8).setValue('');     // OTP dibersihkan
    sheet.getRange(row, 10).setValue('');    // OtpExpiry dibersihkan

    count++;
  }

  Logger.log(`SELESAI. ${count} akun user telah direset ke password default: "${DEFAULT_PASSWORD}"`);
  Logger.log('Setiap user WAJIB login lalu segera ganti password sendiri lewat menu Profil.');
}
// =============================================================================

// =============================================================================
// MIGRASI STRUKTUR ARSIP: LEMARI -> TAHUN/BULAN
// SENGAJA tidak dihubungkan ke apiHandler — hanya bisa dijalankan manual
// oleh Admin lewat Apps Script Editor, karena ini operasi struktural sekali-pakai.
//
// CARA PAKAI:
// 1. Buka Apps Script Editor (Extensions > Apps Script).
// 2. Pilih fungsi 'migrateArchivesToYearStructure' di dropdown atas, klik Run.
// 3. Cek "Execution log" untuk konfirmasi jumlah dokumen yang dipindahkan.
//
// Sheet Lemari lama TIDAK dihapus — hanya diganti nama jadi "OLD_<nama>"
// sebagai cadangan, supaya data tidak pernah hilang meski migrasi diulang.
// =============================================================================
function migrateArchivesToYearStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lemariList = getLemariListFromMaster(ss);
  let migratedCount = 0;
  let skippedCount = 0;
  const yearsCreated = {};

  lemariList.forEach(lemariName => {
    const oldSheet = ss.getSheetByName(lemariName);
    if (!oldSheet) {
      Logger.log(`Lewati: sheet "${lemariName}" tidak ditemukan.`);
      return;
    }

    const data = oldSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // lewati baris kosong

      const uploadDateRaw = row[9];
      const d = new Date(uploadDateRaw);
      const tahun = isNaN(d.getTime()) ? new Date().getFullYear().toString() : d.getFullYear().toString();
      const bulan = getBulanFromDate(uploadDateRaw);

      const targetSheet = getOrCreateTahunSheet(ss, tahun);

      // Cek duplikasi (jaga-jaga kalau migrasi dijalankan dua kali)
      const existingIds = targetSheet.getDataRange().getValues().map(r => r[0]);
      if (existingIds.includes(row[0])) {
        skippedCount++;
        continue;
      }

      targetSheet.appendRow([
        row[0],           // ID
        row[1],           // Nomor Arsip
        row[2],           // Nama Arsip
        row[3],           // Perihal
        row[4],           // Kategori
        bulan,            // Bulan (baru, dari tanggal upload)
        row[6],           // Ordner (tetap dipertahankan)
        row[7],           // Jenis File
        row[8],           // Link File
        row[9],           // Tanggal Upload (asli)
        row[10],          // Pengupload
        row[11]           // SharedWith
      ]);

      yearsCreated[tahun] = true;
      migratedCount++;
    }

    // Ganti nama sheet lama jadi cadangan (TIDAK dihapus)
    try {
      if (!oldSheet.getName().startsWith('OLD_')) {
        oldSheet.setName('OLD_' + lemariName);
      }
    } catch (e) {
      Logger.log(`Gagal mengganti nama sheet "${lemariName}": ${e.message}`);
    }
  });

  // Daftarkan tahun-tahun baru ke Settings supaya langsung muncul di Konfigurasi Master Data
  const setSheet = ss.getSheetByName('Settings');
  if (setSheet) {
    const existingSettings = setSheet.getDataRange().getValues();
    Object.keys(yearsCreated).forEach(tahun => {
      const alreadyExists = existingSettings.some(r => r[0] === 'Tahun' && r[1].toString() === tahun);
      if (!alreadyExists) {
        setSheet.appendRow(['Tahun', tahun]);
      }
    });
  }

  Logger.log(`SELESAI. ${migratedCount} dokumen berhasil dipindahkan ke struktur Tahun/Bulan.`);
  if (skippedCount > 0) Logger.log(`${skippedCount} dokumen dilewati karena sudah pernah dimigrasi sebelumnya.`);
  Logger.log(`Tahun yang dibuat/terpakai: ${Object.keys(yearsCreated).join(', ') || '(tidak ada)'}`);
  Logger.log('Sheet Lemari lama telah diganti nama menjadi "OLD_<nama>" sebagai cadangan (tidak dihapus).');
}
// =============================================================================
function apiHandler(action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureUserSheetSchema(ss);
    switch (action) {
      case 'checkSession': return checkSession(ss, payload);
      case 'login': return loginUser(ss, payload);
      case 'logout': return logoutUser(ss, payload);
      case 'requestPasswordOTP': return requestOtp(ss, payload);
      case 'requestOtp': return requestOtp(ss, payload);
      case 'verifyOTPAndResetPassword': return resetPassword(ss, payload);
      case 'resetPassword': return resetPassword(ss, payload);
      case 'updateProfile': return updateUserProfile(ss, payload);
      case 'saveProfile': return saveProfile(ss, payload);
      case 'getStats': return getDashboardStats(ss, payload);
      case 'getDashboardStats': return getDashboardStats(ss, payload);
      case 'getNextNomorSurat': return getNextNomorSurat(ss);
      case 'uploadBpmAttachments': return uploadBpmAttachments(ss, payload);
      case 'getData': return getData(ss, payload);
      case 'saveData': return saveData(ss, payload);
      case 'deleteData': return deleteData(ss, payload);
      case 'getSettings': return getSettings(ss, payload);
      case 'saveSetting': return saveSetting(ss, payload);
      case 'saveSettings': return saveMasterSettings(ss, payload);
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
    if (data[i][0] == username) {
      const row = i + 1;
      const now = new Date();

      // 1. Cek apakah akun sedang terkunci akibat terlalu banyak percobaan gagal
      const lockUntilRaw = data[i][11];
      if (lockUntilRaw) {
        const lockUntil = new Date(lockUntilRaw);
        if (!isNaN(lockUntil) && now < lockUntil) {
          const minutesLeft = Math.ceil((lockUntil - now) / 60000);
          throw new Error(`Akun terkunci sementara akibat terlalu banyak percobaan gagal. Coba lagi dalam ${minutesLeft} menit.`);
        }
      }

      // 2. Verifikasi password (mendukung migrasi otomatis dari password lama yang belum di-hash)
      const storedPassword = data[i][1] ? data[i][1].toString() : '';
      const storedSalt = data[i][8] ? data[i][8].toString() : '';
      let passwordMatches = false;

      if (looksHashed(storedPassword) && storedSalt) {
        passwordMatches = (hashPassword(password, storedSalt) === storedPassword);
      } else {
        // Data lama: password masih polos di sheet. Cocokkan langsung, lalu upgrade ke hash.
        passwordMatches = (storedPassword === String(password));
        if (passwordMatches) {
          const newSalt = generateSalt();
          sheet.getRange(row, 2).setValue(hashPassword(password, newSalt));
          sheet.getRange(row, 9).setValue(newSalt);
        }
      }

      if (!passwordMatches) {
        const attempts = (parseInt(data[i][10], 10) || 0) + 1;
        sheet.getRange(row, 11).setValue(attempts);
        if (attempts >= MAX_LOGIN_ATTEMPTS) {
          const lockUntil = new Date(now.getTime() + LOCK_DURATION_MINUTES * 60000);
          sheet.getRange(row, 12).setValue(lockUntil);
          throw new Error(`Terlalu banyak percobaan gagal. Akun dikunci selama ${LOCK_DURATION_MINUTES} menit.`);
        }
        throw new Error("Username atau Password salah.");
      }

      if (data[i][4] !== 'Active') throw new Error("Akun Anda dinonaktifkan/suspend.");

      // 3. Login berhasil -> reset counter percobaan gagal
      sheet.getRange(row, 11).setValue(0);
      sheet.getRange(row, 12).setValue('');

      const token = Utilities.getUuid();
      sheet.getRange(row, 6).setValue(token);

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
      const row = i + 1;
      const now = new Date();

      // Cegah spam permintaan OTP beruntun
      const lastRequestRaw = data[i][12];
      if (lastRequestRaw) {
        const lastRequest = new Date(lastRequestRaw);
        const secondsSince = (now - lastRequest) / 1000;
        if (!isNaN(lastRequest) && secondsSince < OTP_REQUEST_COOLDOWN_SECONDS) {
          const waitSeconds = Math.ceil(OTP_REQUEST_COOLDOWN_SECONDS - secondsSince);
          throw new Error(`Mohon tunggu ${waitSeconds} detik sebelum meminta kode OTP baru.`);
        }
      }

      const otp = Math.floor(100000 + Math.random() * 900000).toString();
      const expiry = new Date(now.getTime() + OTP_EXPIRY_MINUTES * 60000);
      sheet.getRange(row, 8).setValue(otp);
      sheet.getRange(row, 10).setValue(expiry);
      sheet.getRange(row, 13).setValue(now);

      try {
        MailApp.sendEmail({
          to: email,
          subject: "[E-Arsip PLN] Kode Reset Password",
          htmlBody: `
            <h3>Permintaan Reset Password</h3>
            <p>Halo ${data[i][2]},</p>
            <p>Gunakan kode OTP berikut untuk me-reset kata sandi Anda. Kode berlaku selama ${OTP_EXPIRY_MINUTES} menit:</p>
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
      const row = i + 1;

      if (String(data[i][7]) !== String(otp) || !data[i][7]) {
        throw new Error("Kode OTP salah atau kadaluarsa.");
      }

      const expiryRaw = data[i][9];
      if (expiryRaw) {
        const expiry = new Date(expiryRaw);
        if (!isNaN(expiry) && new Date() > expiry) {
          throw new Error("Kode OTP sudah kadaluarsa. Silakan minta kode baru.");
        }
      }

      const newSalt = generateSalt();
      sheet.getRange(row, 2).setValue(hashPassword(newPass, newSalt));
      sheet.getRange(row, 9).setValue(newSalt);
      sheet.getRange(row, 8).setValue("");   // Hapus OTP setelah dipakai
      sheet.getRange(row, 10).setValue("");  // Hapus expiry
      sheet.getRange(row, 11).setValue(0);   // Reset percobaan gagal login
      sheet.getRange(row, 12).setValue("");  // Buka lock jika ada

      logActivity(ss, data[i][0], 'Reset Password', 'Sukses reset password via OTP');
      return { status: 'success', message: 'Password berhasil diubah.' };
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
      const row = i + 1;
      const storedPassword = data[i][1] ? data[i][1].toString() : '';
      const storedSalt = data[i][8] ? data[i][8].toString() : '';

      let oldPasswordMatches = false;
      if (looksHashed(storedPassword) && storedSalt) {
        oldPasswordMatches = (hashPassword(password_lama, storedSalt) === storedPassword);
      } else {
        oldPasswordMatches = (storedPassword === String(password_lama));
      }
      if (!oldPasswordMatches) throw new Error("Password lama salah.");

      sheet.getRange(row, 3).setValue(nama_lengkap);
      if (password_baru && password_baru.trim() !== "") {
        const newSalt = generateSalt();
        sheet.getRange(row, 2).setValue(hashPassword(password_baru, newSalt));
        sheet.getRange(row, 9).setValue(newSalt);
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

// Adapter: frontend mengirim payload ter-nested { token, profile: {...} },
// sedangkan updateUserProfile menerima field langsung.
function saveProfile(ss, { token, profile }) {
  profile = profile || {};
  return updateUserProfile(ss, {
    token: token,
    nama_lengkap: profile.nama_lengkap,
    password_lama: profile.password_lama,
    password_baru: profile.password_baru
  });
}

// Konversi angka bulan (1-12) ke angka romawi, untuk format penomoran surat resmi
function toRomanMonth(month) {
  const romans = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
  return romans[month - 1] || 'I';
}

// Hitung nomor urut berikutnya untuk format: 001/REN/VIII/2026 (reset tiap tahun)
function generateNomorSuratPreview(ss) {
  const bpmSheet = getOrCreateBpmSheet(ss);
  const data = bpmSheet.getDataRange().getValues();
  const now = new Date();
  const year = now.getFullYear();
  const pattern = /^(\d+)\/REN\/[IVXLCDM]+\/(\d{4})$/;
  let maxSeq = 0;

  for (let i = 1; i < data.length; i++) {
    const val = data[i][13] ? data[i][13].toString().trim() : '';
    const m = val.match(pattern);
    if (m && parseInt(m[2], 10) === year) {
      const seq = parseInt(m[1], 10);
      if (seq > maxSeq) maxSeq = seq;
    }
  }

  const nextSeq = maxSeq + 1;
  return `${String(nextSeq).padStart(3, '0')}/REN/${toRomanMonth(now.getMonth() + 1)}/${year}`;
}

function getNextNomorSurat(ss) {
  return { status: 'success', data: generateNomorSuratPreview(ss) };
}
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

// --- GET DATA (DI-UPDATE DENGAN BACAAN SELURUH FIELD BPM) ---
function getData(ss, { token, type }) {
  const user = validateToken(ss, token);
  let result = [];

  if (type === 'users') {
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { status: 'success', data: [], list: [] };

    let rawData = sheet.getDataRange().getValues();
    rawData.shift();
    
    rawData.forEach(row => {
      if (row[0]) {
        result.push({
          username: row[0],
          password: '',
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
    const tahunList = getTahunListFromMaster(ss);

    tahunList.forEach(tahunName => {
      const tahunSheet = ss.getSheetByName(tahunName.replace(/\s+/g, '_'));
      if (tahunSheet) {
        let rawData = tahunSheet.getDataRange().getValues();
        rawData.shift();
        rawData.forEach(row => {
          if (!row[0]) return;

          let isOwner = row[10] === user.username;
          let sharedVal = row[11] ? row[11].toString() : '';
          let isShared = sharedVal.includes(user.username) || sharedVal === 'Public';

          if (user.role === 'Admin' || isOwner || isShared) {
            let rawLinkString = row[8] ? row[8].toString() : '';
            let linkArray = rawLinkString ? rawLinkString.split(', ').map(l => l.trim()).filter(Boolean) : [];

            result.push({
              id: row[0],
              nomor: row[1],
              nama: row[2],
              perihal: row[3],
              kategori: row[4],
              tahun: tahunName,
              bulan: row[5] || 'Januari',
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
    rawData.shift();
    
    rawData.forEach(row => {
      if (row[0] || row[1] || row[2]) {
        let safeIso = function(val) {
          if (!val) return null;
          let d = new Date(val);
          return isNaN(d.getTime()) ? null : d.toISOString();
        };

        let rawIsPfk = row[3];
        let isPfkVal = typeof rawIsPfk === 'string' 
          ? rawIsPfk.trim().toUpperCase() === 'TRUE' 
          : Boolean(rawIsPfk);

        let stepTimestamps = {};
        if (row[20]) {
          try {
            stepTimestamps = JSON.parse(row[20]);
          } catch(e) {
            stepTimestamps = {};
          }
        }

        let rawLampiranStr = row[21] ? row[21].toString() : '';
        let lampiranArr = rawLampiranStr ? rawLampiranStr.split(', ').map(l => l.trim()).filter(Boolean) : [];

        result.push({
          id: String(row[0] || Date.now()),
          kode: row[1] ? String(row[1]) : '-',
          judul: row[2] ? String(row[2]) : 'Permohonan Survey',
          permohonan: row[2] ? String(row[2]) : 'Permohonan Survey',
          isPFK: isPfkVal,
          step: parseInt(row[4]) || 1,
          statusDetail: row[5] ? String(row[5]) : 'Kirim Surat Permohonan Survey & RAB',
          tanggal: formatDate(row[6]),
          uploader: row[7] ? String(row[7]) : 'System',
          pemohon: row[7] ? String(row[7]) : 'System',
          lokasi: row[8] ? String(row[8]) : '-',
          namaLokasi: row[8] ? String(row[8]) : '-',
          alamat: row[9] ? String(row[9]) : '-',
          ulp: row[10] ? String(row[10]) : '-',
          jumlahUnit: row[11] ? String(row[11]) : '-',
          tarifDaya: row[12] ? String(row[12]) : '-',
          nomorSurat: row[13] ? String(row[13]) : '-',
          tanggalSurat: row[14] ? String(row[14]) : '-',
          pic: row[15] ? String(row[15]) : '-',
          tglPengajuan: safeIso(row[16]) || safeIso(row[6]) || new Date().toISOString(),
          tglSurvey: safeIso(row[17]),
          tglManajemen: safeIso(row[18]),
          tglSelesai: safeIso(row[19]),
          stepTimestamps: stepTimestamps,
          links: lampiranArr,
          link: lampiranArr[0] || null
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

// --- SAVE DATA (DI-UPDATE DENGAN SIMPAN SELURUH FIELD PENGAJUAN SURVEY) ---
function saveData(ss, { token, type, data }) {
  const user = validateToken(ss, token);

  if (type === 'archives') {
    const id = data.id || Utilities.getUuid();
    const timestamp = new Date();
    const targetTahun = (data.tahun || new Date().getFullYear().toString()).toString().trim();
    const targetBulan = (data.bulan || getBulanFromDate(timestamp)).toString().trim();
    const targetOrdner = (data.ordner || '-').trim();

    const targetSheet = getOrCreateTahunSheet(ss, targetTahun);

    let finalLinks = Array.isArray(data.links) ? data.links : (Array.isArray(data.existingLinks) ? data.existingLinks : []);

    if (Array.isArray(data.files) && data.files.length > 0) {
      data.files.forEach(fObj => {
        if (fObj && fObj.base64) {
          try {
            const uploadedUrl = uploadArsipToDrive(fObj, fObj.name || data.nama, data.kategori, targetTahun, targetBulan);
            finalLinks.push(uploadedUrl);
          } catch (e) {
            console.warn("Gagal upload salah satu file: " + e.message);
          }
        }
      });
    } else if (data.fileObj && data.fileObj.base64) {
      try {
        const uploadedUrl = uploadArsipToDrive(data.fileObj, data.fileObj.name || data.nama, data.kategori, targetTahun, targetBulan);
        finalLinks.push(uploadedUrl);
      } catch (e) {
        throw new Error("Gagal upload file ke Drive: " + e.message);
      }
    }

    if (finalLinks.length === 0) {
      throw new Error("Dokumen wajib memiliki setidaknya satu file lampiran.");
    }

    const combinedLinksStr = finalLinks.join(', ');

    if (data.id) { // Mode Edit
      let found = false;
      const tahunList = getTahunListFromMaster(ss);

      for (let l = 0; l < tahunList.length; l++) {
        const currentSheet = ss.getSheetByName(tahunList[l].replace(/\s+/g, '_'));
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
                targetBulan, targetOrdner, data.jenis, combinedLinksStr, originalDate, originalUploader, data.shared || 'Public'
              ]);
            } else {
              currentSheet.getRange(i + 1, 1, 1, 12).setValues([[
                data.id, data.nomor, data.nama, data.perihal, data.kategori,
                targetBulan, targetOrdner, data.jenis, combinedLinksStr, originalDate, originalUploader, data.shared || 'Public'
              ]]);
            }

            found = true;
            break;
          }
        }
        if (found) break;
      }
      if (!found) throw new Error("Data tidak ditemukan.");
      logActivity(ss, user.username, 'Edit Arsip', `Mengubah arsip: ${data.nama} (${targetTahun}/${targetBulan}, Ordner: ${targetOrdner})`);
    } else { // Mode Baru
      targetSheet.appendRow([
        id,
        data.nomor,
        data.nama,
        data.perihal,
        data.kategori,
        targetBulan,
        targetOrdner,
        data.jenis,
        combinedLinksStr,
        timestamp,
        user.username,
        data.shared || 'Public'
      ]);
      logActivity(ss, user.username, 'Upload Arsip', `Menambah arsip: ${data.nama} di ${targetTahun}/${targetBulan} (${targetOrdner})`);
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
      if (rows[i][0].toString() === (data.id || '').toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    const now = new Date();

    if (rowIndex > 0) { // Update status step BPM
      let rowData = rows[rowIndex - 1];
      let stepVal = data.step !== undefined ? parseInt(data.step) : (parseInt(rowData[4]) || 1);

      let tglPengajuan = rowData[16] || rowData[6] || now;
      let tglSurvey    = rowData[17] || (stepVal >= 4 ? now : '');
      let tglManajemen = rowData[18] || (stepVal >= 7 ? now : '');
      let tglSelesai   = rowData[19] || (stepVal >= 9 ? now : '');

      let timestampsStr = data.stepTimestamps ? JSON.stringify(data.stepTimestamps) : (rowData[20] || '{}');

      bpmSheet.getRange(rowIndex, 1, 1, 21).setValues([[
        data.id,
        data.kode || rowData[1],
        data.permohonan || data.judul || rowData[2],
        data.isPFK !== undefined ? Boolean(data.isPFK) : rowData[3],
        stepVal,
        data.statusDetail || rowData[5],
        rowData[6] || now,
        rowData[7] || user.username,
        data.namaLokasi || data.lokasi || rowData[8] || '-',
        data.alamat || rowData[9] || '-',
        data.ulp || rowData[10] || '-',
        data.jumlahUnit || rowData[11] || '-',
        data.tarifDaya || rowData[12] || '-',
        data.nomorSurat || rowData[13] || '-',
        data.tanggalSurat || rowData[14] || '-',
        data.pic || rowData[15] || '-',
        tglPengajuan,
        tglSurvey,
        tglManajemen,
        tglSelesai,
        timestampsStr
      ]]);
      
      logActivity(ss, user.username, 'Update BPM', `Memperbarui alur BPM: ${data.permohonan || data.judul || rowData[2]} ke Step ${stepVal}`);
    } else { // Permohonan BPM Baru
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);
      let finalNomorSurat;
      try {
        finalNomorSurat = generateNomorSuratPreview(ss);

        const newId = Date.now().toString();
        const initialTimestampsStr = JSON.stringify(data.stepTimestamps || { 1: { start: now.toISOString(), end: null } });

        bpmSheet.appendRow([
          newId,
          data.kode || ('PLN-PFK-' + Math.floor(10000000 + Math.random() * 90000000)),
          data.permohonan || data.judul || 'Permohonan Survey',
          data.isPFK !== undefined ? Boolean(data.isPFK) : false,
          data.step || 1,
          data.statusDetail || 'Kirim Surat Permohonan Survey & RAB',
          now,
          user.username,
          data.namaLokasi || data.lokasi || '-',
          data.alamat || '-',
          data.ulp || '-',
          data.jumlahUnit || '-',
          data.tarifDaya || '-',
          finalNomorSurat,
          data.tanggalSurat || '-',
          data.pic || '-',
          data.tglPengajuan || now, // Tgl_Pengajuan
          '',  // Tgl_Survey
          '',  // Tgl_Manajemen
          '',  // Tgl_Selesai
          initialTimestampsStr,
          ''   // Lampiran
        ]);
      } finally {
        lock.releaseLock();
      }
      logActivity(ss, user.username, 'Pengajuan BPM', `Membuat permohonan survey baru: ${data.permohonan || data.judul} (No. Surat: ${finalNomorSurat})`);
    }
  }
  else if (type === 'users' && user.role === 'Admin') {
    const sheet = ss.getSheetByName('Users');
    if (!data.isEdit) {
      const users = sheet.getDataRange().getValues();
      if (users.some(u => u[0] === data.username)) throw new Error("Username sudah ada.");
      if (!data.password || data.password.trim() === '') throw new Error("Password wajib diisi untuk user baru.");
      const salt = generateSalt();
      sheet.appendRow([data.username, hashPassword(data.password, salt), data.nama_lengkap, data.role, data.status, '', data.email || '', '', salt, '', 0, '', '']);
      logActivity(ss, user.username, 'Add User', `Menambah user: ${data.username}`);
    } else {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === data.username) {
          const row = i + 1;
          sheet.getRange(row, 3, 1, 3).setValues([[data.nama_lengkap, data.role, data.status]]);
          sheet.getRange(row, 7).setValue(data.email || '');
          // Hanya ubah password jika admin mengisi field baru; kosongkan untuk tetap pakai password lama
          if (data.password && data.password.trim() !== '') {
            const newSalt = generateSalt();
            sheet.getRange(row, 2).setValue(hashPassword(data.password, newSalt));
            sheet.getRange(row, 9).setValue(newSalt);
          }
          break;
        }
      }
      logActivity(ss, user.username, 'Edit User', `Mengubah user: ${data.username}`);
    }
  }

  return { status: 'success', success: true };
}

// --- UPLOAD LAMPIRAN BPM ---
function uploadBpmAttachments(ss, { token, bpmId, files }) {
  const user = validateToken(ss, token);
  const bpmSheet = getOrCreateBpmSheet(ss);
  const data = bpmSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === (bpmId || '').toString()) {
      let existingStr = data[i][21] ? data[i][21].toString() : '';
      let linksArr = existingStr ? existingStr.split(', ').map(l => l.trim()).filter(Boolean) : [];

      if (Array.isArray(files) && files.length > 0) {
        files.forEach(fObj => {
          if (fObj && fObj.base64) {
            try {
              const uploadedUrl = uploadToDrive(fObj, fObj.name || 'Lampiran_BPM', 'BPM', user.nama_lengkap);
              linksArr.push(uploadedUrl);
            } catch (e) {
              console.warn("Gagal upload salah satu lampiran BPM: " + e.message);
            }
          }
        });
      } else {
        throw new Error("Tidak ada file yang diunggah.");
      }

      bpmSheet.getRange(i + 1, 22).setValue(linksArr.join(', '));
      logActivity(ss, user.username, 'Upload Lampiran BPM', `Menambah ${files.length} lampiran ke BPM: ${data[i][2]}`);

      return { status: 'success', success: true, data: linksArr };
    }
  }
  throw new Error("Data BPM tidak ditemukan.");
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
    const tahunList = getTahunListFromMaster(ss);
    for (let l = 0; l < tahunList.length; l++) {
      const currentSheet = ss.getSheetByName(tahunList[l].replace(/\s+/g, '_'));
      if (!currentSheet) continue;

      const data = currentSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == id) {
          if (user.role !== 'Admin' && data[i][10] !== user.username) {
            throw new Error("Anda tidak berhak menghapus file ini.");
          }
          deletedName = data[i][2];
          const rawLinks = data[i][8] ? data[i][8].toString().split(', ') : [];
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
        const rawLampiran = data[i][21] ? data[i][21].toString().split(', ') : [];
        rawLampiran.forEach(link => deleteFileFromDrive(link));

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
  let ulps = [];
  let tahuns = [];

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Category') categories.push(data[i][1]);
      if (data[i][0] === 'Extension') extensions.push(data[i][1]);
      if (data[i][0] === 'Nomor') nomors.push(data[i][1]);
      if (data[i][0] === 'Lemari') lemaris.push(data[i][1]);
      if (data[i][0] === 'Ordner') ordners.push(data[i][1]);
      if (data[i][0] === 'ULP') ulps.push(data[i][1]);
      if (data[i][0] === 'Tahun') tahuns.push(data[i][1].toString());
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
  const cleanUlps = ulps.length > 0 ? ulps : [
    'ULP Malang Kota', 'ULP Blimbing', 'ULP Dinoyo', 'ULP Kebonagung', 'ULP Singosari',
    'ULP Lawang', 'ULP Batu', 'ULP Tumpang', 'ULP Bululawang', 'ULP Gondanglegi',
    'ULP Kepanjen', 'ULP Sumberpucung', 'ULP Dampit'
  ];
  const currentYear = new Date().getFullYear();
  const cleanTahuns = (tahuns.length > 0 ? tahuns : [(currentYear - 1).toString(), currentYear.toString(), (currentYear + 1).toString()])
    .sort((a, b) => b.localeCompare(a));
  cleanLemaris.forEach(l => getOrCreateLemariSheet(ss, l));
  cleanTahuns.forEach(t => getOrCreateTahunSheet(ss, t));

  return {
    status: 'success',
    data: {
      categories: categories.length > 0 ? categories : ['PFK', 'ESTETIKA', 'Pelanggan TM', 'SPKLU'],
      extensions,
      nomors,
      lemaris: cleanLemaris,
      ordners: cleanOrdners,
      ulps: cleanUlps,
      tahuns: cleanTahuns,
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
  if (type === 'Tahun') {
    getOrCreateTahunSheet(ss, cleanVal);
  }

  logActivity(ss, user.username, 'Update Setting', `Menambah ${type}: ${cleanVal}`);
  return { status: 'success' };
}

function saveMasterSettings(ss, { token, data }) {
  const user = validateToken(ss, token);
  if (user.role !== 'Admin') throw new Error("Akses Ditolak.");

  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['Type', 'Value']);
  } else {
    sheet.clearContents();
    sheet.appendRow(['Type', 'Value']);
  }

  const mapKeys = {
    categories: 'Category',
    extensions: 'Extension',
    nomors: 'Nomor',
    lemaris: 'Lemari',
    ordners: 'Ordner',
    ulps: 'ULP',
    tahuns: 'Tahun'
  };

  Object.keys(mapKeys).forEach(key => {
    if (Array.isArray(data[key])) {
      data[key].forEach(val => {
        let clean = val.toString().trim();
        if (mapKeys[key] === 'Lemari') clean = clean.replace(/\s+/g, '_');
        sheet.appendRow([mapKeys[key], clean]);
        if (mapKeys[key] === 'Lemari') getOrCreateLemariSheet(ss, clean);
      });
    }
  });

  logActivity(ss, user.username, 'Update Master Data', 'Menyimpan konfigurasi master variabel baru.');
  return { status: 'success', success: true };
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

// Upload khusus untuk Kotak Arsip Digital: folder utama dikelompokkan per KATEGORI dulu
// (Folder PFK, Folder ESTETIKA, Folder Pelanggan TM, Folder SPKLU), baru di dalamnya per Tahun/Bulan.
// Terpisah dari uploadToDrive() supaya tidak mengubah struktur folder lampiran BPM yang sudah ada.
function uploadArsipToDrive(fileObj, fileName, kategori, tahun, bulan) {
  const rootIterator = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let rootFolder = rootIterator.hasNext() ? rootIterator.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  const kategoriFolder = getOrCreateSubFolder(rootFolder, 'Folder ' + (kategori ? kategori.trim() : 'Umum'));
  const tahunFolder = getOrCreateSubFolder(kategoriFolder, tahun ? tahun.toString().trim() : new Date().getFullYear().toString());
  const bulanFolder = getOrCreateSubFolder(tahunFolder, bulan ? bulan.toString().trim() : 'Januari');

  const decoded = Utilities.base64Decode(fileObj.base64);
  const blob = Utilities.newBlob(decoded, fileObj.mimeType, fileName);
  const file = bulanFolder.createFile(blob);

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
