// ============================================================
// SMK Avicena Rajeg - Portal Akademik
// Google Apps Script Backend (kode.gs)
// ============================================================

const SPREADSHEET_ID = 'MASUKKAN_SPREADSHEET_ID_ANDA_DI_SINI'; // Ganti dengan ID Spreadsheet
const SHEET_USERS = 'Users';
const SHEET_TASKS = 'Tasks';
const SHEET_INFO = 'Info';
const SHEET_ATTENDANCE = 'Attendance';
const SHEET_GRADES = 'Grades';
const SHEET_SUBMISSIONS = 'Submissions';

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SMK Avicena Rajeg - Portal Akademik')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// INITIALIZE SPREADSHEET & SHEETS
// ============================================================
function initializeSpreadsheet() {
  try {
    let ss;
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      // Jika spreadsheet belum ada, buat baru
      ss = SpreadsheetApp.create('SMK Avicena - Database Portal Akademik');
      const newId = ss.getId();
      // Simpan ID ke Script Properties
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', newId);
      Logger.log('Spreadsheet baru dibuat. ID: ' + newId);
      Logger.log('Silakan update SPREADSHEET_ID di kode dengan ID: ' + newId);
    }

    // Buat sheet jika belum ada
    const sheets = {
      [SHEET_USERS]: ['username', 'password', 'role', 'name', 'fullName', 'nis', 'nip', 'mapel', 'kelas', 'tempatLahir', 'tanggalLahir', 'jenisKelamin', 'alamat', 'noTelp', 'ekskul', 'namaOrtu', 'noTelpOrtu', 'anakUsername', 'anak', 'hubungan', 'jabatan', 'isCoAdmin', 'namaPanggilan'],
      [SHEET_TASKS]: ['id', 'judul', 'mapel', 'kelas', 'deadline', 'deskripsi', 'guru', 'createdAt', 'attachment'],
      [SHEET_INFO]: ['id', 'tanggal', 'judul', 'isi'],
      [SHEET_ATTENDANCE]: ['username', 'tanggal', 'status', 'recordedBy', 'recordedAt'],
      [SHEET_GRADES]: ['username', 'mapel', 'nilai', 'semester'],
      [SHEET_SUBMISSIONS]: ['username', 'taskId', 'status', 'submittedAt', 'jawaban', 'attachment', 'nilai', 'feedback', 'gradedBy', 'gradedAt']
    };

    for (const [sheetName, headers] of Object.entries(sheets)) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
        Logger.log('Sheet dibuat: ' + sheetName);
      }
    }

    // Seed data awal jika Users kosong
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (usersSheet.getLastRow() <= 1) {
      seedInitialData(ss);
    }

    return { success: true, spreadsheetId: ss.getId(), url: ss.getUrl() };

  } catch (error) {
    Logger.log('Error initializeSpreadsheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// SEED INITIAL DATA
// ============================================================
function seedInitialData(ss) {
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  
  const users = [
    ['admin', 'admin123', 'admin', 'Staff Admin', 'Administrator Sistem', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Super Admin', '', ''],
    ['budi', 'guru123', 'guru', 'Bpk. Budi Santoso, S.Kom.', 'Budi Santoso', '', '1987654321', 'Pemrograman Web & Basis Data', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['ahmad', 'siswa123', 'siswa', 'Ahmad Fauzi', 'Ahmad Fauzi', '2024001', '', '', 'X TKJ', 'Jakarta', '2006-05-15', 'Laki-laki', 'Jl. Raya Rajeg No. 123', '081234567890', 'Futsal', 'Siti Rahayu', '081298765432', '', '', '', '', '', 'Ahmad'],
    ['siti', 'ortu123', 'ortu', 'Ibu Siti Rahayu', 'Siti Rahayu', '', '', '', '', '', '', '', '', '', '', '', '', 'ahmad', 'Ahmad Fauzi', 'Ibu', '', '', '']
  ];

  usersSheet.getRange(2, 1, users.length, users[0].length).setValues(users);
  Logger.log('Seed data pengguna berhasil.');

  // Seed Info
  const infoSheet = ss.getSheetByName(SHEET_INFO);
  const today = new Date().toISOString().split('T')[0];
  const info = [
    ['I1', today, 'Selamat Datang di Portal Akademik', 'Sistem siap digunakan. Silakan cek jadwal dan tugas terbaru.'],
    ['I2', today, 'Ujian Tengah Semester', 'UTS akan dilaksanakan minggu depan. Persiapkan diri dengan baik.']
  ];
  infoSheet.getRange(2, 1, info.length, info[0].length).setValues(info);

  // Seed Grades
  const gradesSheet = ss.getSheetByName(SHEET_GRADES);
  const grades = [
    ['ahmad', 'K.K TKJ', 88, 'Ganjil 2025/2026'],
    ['ahmad', 'Matematika', 85, 'Ganjil 2025/2026'],
    ['ahmad', 'Bahasa Inggris', 78, 'Ganjil 2025/2026']
  ];
  gradesSheet.getRange(2, 1, grades.length, grades[0].length).setValues(grades);

  // Seed Attendance
  const attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const att = [
    ['ahmad', yesterday.toISOString().split('T')[0], 'Hadir', 'admin', new Date().toISOString()]
  ];
  attSheet.getRange(2, 1, att.length, att[0].length).setValues(att);
}

// ============================================================
// USER CRUD
// ============================================================
function getAllUsers() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const users = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const user = {};
      headers.forEach((h, idx) => {
        if (row[idx] !== '' && row[idx] !== null && row[idx] !== undefined) {
          user[h] = row[idx];
        }
      });
      if (user.username) {
        users[user.username] = user;
      }
    }
    return users;
  } catch (error) {
    Logger.log('Error getAllUsers: ' + error.toString());
    return {};
  }
}

function saveUser(userData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameIdx = headers.indexOf('username');

    // Cek apakah user sudah ada
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameIdx] === userData.username) {
        existingRow = i + 1;
        break;
      }
    }

    const rowData = headers.map(h => userData[h] !== undefined ? userData[h] : '');

    if (existingRow > 0) {
      // Update
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowData]);
      return { success: true, message: 'User updated' };
    } else {
      // Insert
      sheet.appendRow(rowData);
      return { success: true, message: 'User created' };
    }
  } catch (error) {
    Logger.log('Error saveUser: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function deleteUserFromSheet(username) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const usernameIdx = data[0].indexOf('username');

    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameIdx] === username) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'User not found' };
  } catch (error) {
    Logger.log('Error deleteUserFromSheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// TASKS CRUD
// ============================================================
function getAllTasks() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const tasks = [];
    for (let i = 1; i < data.length; i++) {
      const task = {};
      headers.forEach((h, idx) => {
        task[h] = data[i][idx];
      });
      // Parse attachment JSON jika ada
      if (task.attachment && typeof task.attachment === 'string') {
        try { task.attachment = JSON.parse(task.attachment); } catch(e) { task.attachment = null; }
      }
      tasks.push(task);
    }
    return tasks;
  } catch (error) {
    Logger.log('Error getAllTasks: ' + error.toString());
    return [];
  }
}

function saveTask(taskData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');

    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === taskData.id) {
        existingRow = i + 1;
        break;
      }
    }

    const rowData = headers.map(h => {
      let val = taskData[h];
      if (h === 'attachment' && val && typeof val === 'object') {
        val = JSON.stringify(val);
      }
      return val !== undefined ? val : '';
    });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch (error) {
    Logger.log('Error saveTask: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function deleteTaskFromSheet(taskId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    const data = sheet.getDataRange().getValues();
    const idIdx = data[0].indexOf('id');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === taskId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Task not found' };
  } catch (error) {
    Logger.log('Error deleteTaskFromSheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// INFO CRUD
// ============================================================
function getAllInfo() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INFO);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const info = [];
    for (let i = 1; i < data.length; i++) {
      const item = {};
      headers.forEach((h, idx) => { item[h] = data[i][idx]; });
      info.push(item);
    }
    return info.reverse(); // Terbaru di atas
  } catch (error) {
    Logger.log('Error getAllInfo: ' + error.toString());
    return [];
  }
}

function saveInfo(infoData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INFO);
    const headers = sheet.getDataRange().getValues()[0];
    const rowData = headers.map(h => infoData[h] !== undefined ? infoData[h] : '');
    sheet.appendRow(rowData);
    return { success: true };
  } catch (error) {
    Logger.log('Error saveInfo: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function deleteInfoFromSheet(infoId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INFO);
    const data = sheet.getDataRange().getValues();
    const idIdx = data[0].indexOf('id');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === infoId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false };
  } catch (error) {
    Logger.log('Error deleteInfoFromSheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// ATTENDANCE CRUD
// ============================================================
function getAllAttendance() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const attendance = [];
    for (let i = 1; i < data.length; i++) {
      const item = {};
      headers.forEach((h, idx) => { item[h] = data[i][idx]; });
      attendance.push(item);
    }
    return attendance;
  } catch (error) {
    Logger.log('Error getAllAttendance: ' + error.toString());
    return [];
  }
}

function saveAttendanceBatch(records) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    records.forEach(record => {
      // Hapus record lama untuk siswa+tanggal yang sama
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === record.username && data[i][1] === record.tanggal) {
          sheet.deleteRow(i + 1);
        }
      }
      // Tambah record baru
      const rowData = headers.map(h => record[h] !== undefined ? record[h] : '');
      sheet.appendRow(rowData);
    });

    return { success: true };
  } catch (error) {
    Logger.log('Error saveAttendanceBatch: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GRADES CRUD
// ============================================================
function getAllGrades() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_GRADES);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const grades = [];
    for (let i = 1; i < data.length; i++) {
      const item = {};
      headers.forEach((h, idx) => { item[h] = data[i][idx]; });
      grades.push(item);
    }
    return grades;
  } catch (error) {
    Logger.log('Error getAllGrades: ' + error.toString());
    return [];
  }
}

// ============================================================
// SUBMISSIONS CRUD
// ============================================================
function getAllSubmissions() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    const headers = data[0];
    const submissions = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const username = row[headers.indexOf('username')];
      const taskId = row[headers.indexOf('taskId')];
      
      if (!submissions[username]) submissions[username] = {};
      
      const sub = {};
      headers.forEach((h, idx) => {
        if (h === 'attachment' && row[idx]) {
          try { sub[h] = JSON.parse(row[idx]); } catch(e) { sub[h] = null; }
        } else {
          sub[h] = row[idx];
        }
      });
      submissions[username][taskId] = sub;
    }
    return submissions;
  } catch (error) {
    Logger.log('Error getAllSubmissions: ' + error.toString());
    return {};
  }
}

function saveSubmission(username, taskId, submissionData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Cari record yang ada
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === taskId) {
        existingRow = i + 1;
        break;
      }
    }

    const rowData = headers.map(h => {
      if (h === 'username') return username;
      if (h === 'taskId') return taskId;
      let val = submissionData[h];
      if (h === 'attachment' && val && typeof val === 'object') {
        val = JSON.stringify(val);
      }
      return val !== undefined ? val : '';
    });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch (error) {
    Logger.log('Error saveSubmission: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET ALL DATA (untuk inisialisasi frontend)
// ============================================================
function getAllData() {
  try {
    return {
      success: true,
      users: getAllUsers(),
      tasks: getAllTasks(),
      info: getAllInfo(),
      attendance: getAllAttendance(),
      grades: getAllGrades(),
      submissions: getAllSubmissions()
    };
  } catch (error) {
    Logger.log('Error getAllData: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// SAVE ALL STATE (dari frontend)
// ============================================================
function saveAllState(state) {
  try {
    const ss = getSpreadsheet();

    // Save Users
    if (state.users) {
      const sheet = ss.getSheetByName(SHEET_USERS);
      const headers = sheet.getDataRange().getValues()[0];
      // Clear data (kecuali header)
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      const rows = Object.values(state.users).map(user =>
        headers.map(h => {
          let val = user[h];
          if (h === 'attachment' && val && typeof val === 'object') val = JSON.stringify(val);
          return val !== undefined ? val : '';
        })
      );
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
    }

    // Save Tasks
    if (state.tasks) {
      const sheet = ss.getSheetByName(SHEET_TASKS);
      const headers = sheet.getDataRange().getValues()[0];
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const rows = state.tasks.map(task =>
        headers.map(h => {
          let val = task[h];
          if (h === 'attachment' && val && typeof val === 'object') val = JSON.stringify(val);
          return val !== undefined ? val : '';
        })
      );
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
    }

    // Save Info
    if (state.info) {
      const sheet = ss.getSheetByName(SHEET_INFO);
      const headers = sheet.getDataRange().getValues()[0];
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const rows = state.info.map(i =>
        headers.map(h => i[h] !== undefined ? i[h] : '')
      );
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
    }

    // Save Attendance
    if (state.attendance) {
      const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
      const headers = sheet.getDataRange().getValues()[0];
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const rows = state.attendance.map(a =>
        headers.map(h => a[h] !== undefined ? a[h] : '')
      );
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
    }

    // Save Grades
    if (state.grades) {
      const sheet = ss.getSheetByName(SHEET_GRADES);
      const headers = sheet.getDataRange().getValues()[0];
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      const rows = state.grades.map(g =>
        headers.map(h => g[h] !== undefined ? g[h] : '')
      );
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
    }

    // Save Submissions
    if (state.submissions) {
      const sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
      const headers = sheet.getDataRange().getValues()[0];
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      
      const allSubs = [];
      Object.entries(state.submissions).forEach(([username, tasks]) => {
        Object.entries(tasks).forEach(([taskId, sub]) => {
          const row = headers.map(h => {
            if (h === 'username') return username;
            if (h === 'taskId') return taskId;
            let val = sub[h];
            if (h === 'attachment' && val && typeof val === 'object') val = JSON.stringify(val);
            return val !== undefined ? val : '';
          });
          allSubs.push(row);
        });
      });
      
      if (allSubs.length > 0) {
        sheet.getRange(2, 1, allSubs.length, headers.length).setValues(allSubs);
      }
    }

    return { success: true };
  } catch (error) {
    Logger.log('Error saveAllState: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// HELPER: GET SPREADSHEET
// ============================================================
function getSpreadsheet() {
  let id = SPREADSHEET_ID;
  if (!id || id === 'MASUKKAN_SPREADSHEET_ID_ANDA_DI_SINI') {
    id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  }
  if (!id) {
    throw new Error('Spreadsheet ID belum diatur. Jalankan initializeSpreadsheet() terlebih dahulu.');
  }
  return SpreadsheetApp.openById(id);
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testInitialize() {
  const result = initializeSpreadsheet();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function testGetAllData() {
  const data = getAllData();
  Logger.log('Users: ' + Object.keys(data.users).length);
  Logger.log('Tasks: ' + data.tasks.length);
  Logger.log('Info: ' + data.info.length);
  return data;
}
