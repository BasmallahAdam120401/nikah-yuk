// Global Configuration & Auto-Database Setup
function doGet() {
  setupDatabase(); // Otomatis buat sheet & header jika belum ada
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('NIKAH YUK! - Wedding Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Setup struktur database otomatis
function setupDatabase() {
  const ss = getDb();
  const sheets = [
    { name: 'Profile', headers: ['Key', 'Value'] },
    { name: 'Users', headers: ['Role', 'UserID', 'Password', 'DisplayName'] },
    { name: 'Tabungan', headers: ['ID', 'Tanggal', 'Role', 'Jumlah', 'Bukti', 'Status'] },
    { name: 'Budget', headers: ['ID', 'Kategori', 'Item', 'Estimasi', 'Porsi_Pria', 'Porsi_Wanita', 'Realisasi', 'Dibayar_Pria', 'Dibayar_Wanita'] },
    { name: 'Progress_Checklist', headers: ['ID', 'Kategori', 'Tugas', 'Target_Selesai', 'Status'] },
    { name: 'Tamu_Undangan', headers: ['ID', 'Nama', 'Kategori', 'Sesi', 'Status_Undang', 'Konfirmasi_Hadir', 'Jumlah_Hadir'] },
    { name: 'Dokumen_KUA', headers: ['ID', 'Nama_Dokumen', 'Pihak', 'Status', 'Catatan'] },
    { name: 'Perlengkapan', headers: ['ID', 'Item', 'Penanggung_Jawab', 'Jumlah', 'Status_Siap'] },
    { name: 'Moodboards', headers: ['ID', 'Judul', 'Kategori', 'URL_Foto', 'Catatan'] }
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#E2F0D9');
    }
  });

  // Seed default Profile jika kosong
  const profSheet = ss.getSheetByName('Profile');
  if (profSheet.getLastRow() <= 1) {
    const defaultProfiles = [
      ['nama_pria', 'Mempelai Pria'],
      ['foto_pria', 'https://via.placeholder.com/150/89CFF0/ffffff?text=Pria'],
      ['nama_wanita', 'Mempelai Wanita'],
      ['foto_wanita', 'https://via.placeholder.com/150/98FB98/ffffff?text=Wanita'],
      ['tanggal_nikah', '2026-12-31'],
      ['rekening_bersama', 'BCA 1234567890 a.n Nikah Yuk']
    ];
    defaultProfiles.forEach(row => profSheet.appendRow(row));
  }

  // Seed default akun login jika sheet Users masih kosong
  const userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() <= 1) {
    const defaultUsers = [
      ['Admin', 'admin', 'admin123', 'Administrator'],
      ['Pria', 'pria', 'pria123', 'Calon Mempelai Pria'],
      ['Wanita', 'wanita', 'wanita123', 'Calon Mempelai Wanita']
    ];
    defaultUsers.forEach(row => userSheet.appendRow(row));
  }
}

// === API BACKEND SERVICES ===

function getDataAll() {
  const ss = getDb();
  const getSheetData = (name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const headers = values[0];
    return values.slice(1).map((row, idx) => {
      let obj = { _rowIndex: idx + 2 };
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  };

  // Convert profile key-value array to object
  const rawProfile = getSheetData('Profile');
  let profile = {};
  rawProfile.forEach(item => profile[item.Key] = item.Value);

  return {
    profile: profile,
    accounts: getAccounts(),
    tabungan: getSheetData('Tabungan'),
    budget: getSheetData('Budget'),
    checklist: getSheetData('Progress_Checklist'),
    tamu: getSheetData('Tamu_Undangan'),
    kua: getSheetData('Dokumen_KUA'),
    perlengkapan: getSheetData('Perlengkapan'),
    moodboard: getSheetData('Moodboards')
    moodboard: getSheetData('Moodboard')
  };
}

function getAccounts() {
  setupDatabase();
  const sheet = getDb().getSheetByName('Users');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      obj[h] = h === 'Password' ? '' : row[i];
    });
    return obj;
  });
}

function validateLogin(role, userId, password) {
  setupDatabase();
  const sheet = getDb().getSheetByName('Users');
  const values = sheet.getDataRange().getValues();
  const normalize = value => String(value || '').trim();

  for (let i = 1; i < values.length; i++) {
    const rowRole = normalize(values[i][0]);
    const rowUserId = normalize(values[i][1]);
    const rowPassword = normalize(values[i][2]);
    const displayName = normalize(values[i][3]);

    if (rowRole === role && rowUserId === normalize(userId) && rowPassword === normalize(password)) {
      return {
        success: true,
        user: {
          role: rowRole,
          userId: rowUserId,
          displayName: displayName || rowRole
        }
      };
    }
  }

  return {
    success: false,
    message: 'ID, password, atau role tidak sesuai.'
  };
}

function updateAccount(role, userId, password, displayName) {
  setupDatabase();
  const sheet = getDb().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const normalize = value => String(value || '').trim();

  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][0]) === normalize(role)) {
      sheet.getRange(i + 1, 2).setValue(normalize(userId));
      if (normalize(password)) {
        sheet.getRange(i + 1, 3).setValue(normalize(password));
      }
      sheet.getRange(i + 1, 4).setValue(normalize(displayName));
      return { success: true };
    }
  }

  sheet.appendRow([normalize(role), normalize(userId), normalize(password), normalize(displayName)]);
  return { success: true };
}

function updateProfile(key, value) {
  const sheet = getDb().getSheetByName('Profile');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return { success: true };
    }
  }
  sheet.appendRow([key, value]);
  return { success: true };
}

function addDataRow(sheetName, rowData) {
  const sheet = getDb().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let row = [];
  rowData.ID = 'ID-' + new Date().getTime();
  headers.forEach(h => {
    row.push(rowData[h] !== undefined ? rowData[h] : '');
  });
  sheet.appendRow(row);
  return { success: true };
}

function updateDataRow(sheetName, id, rowData) {
  const sheet = getDb().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] == id) {
      headers.forEach((h, colIdx) => {
        if (rowData[h] !== undefined) {
          sheet.getRange(i + 1, colIdx + 1).setValue(rowData[h]);
        }
      });
      return { success: true };
    }
  }
  return { success: false, message: 'ID not found' };
}

function deleteDataRow(sheetName, id) {
  const sheet = getDb().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] == id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}
