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
    tabungan: getSheetData('Tabungan'),
    budget: getSheetData('Budget'),
    checklist: getSheetData('Progress_Checklist'),
    tamu: getSheetData('Tamu_Undangan'),
    kua: getSheetData('Dokumen_KUA'),
    perlengkapan: getSheetData('Perlengkapan'),
    moodboard: getSheetData('Moodboard')
  };
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
