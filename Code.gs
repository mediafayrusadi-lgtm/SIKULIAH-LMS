// ================================================
// SIKULIAH - SISTEM INFORMASI PERKULIAHAN
// Backend Google Apps Script
// ================================================

const SPREADSHEET_ID = ''; // Akan diisi dengan ID spreadsheet setelah dibuat
const FOLDER_ID = ''; // Akan diisi dengan ID folder setelah dibuat

// ================================================
// 1. INISIALISASI DAN SETUP
// ================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(800)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function initializeDatabase() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let sheetNames = sheets.map(s => s.getName());

  const requiredSheets = [
    'USERS', 'MAHASISWA', 'DOSEN', 'MATA_KULIAH', 'KELAS',
    'MATERI', 'TUGAS', 'PENGUMPULAN_TUGAS', 'BANK_SOAL',
    'EVALUASI', 'EVALUASI_SOAL', 'HASIL_EVALUASI', 'GAME',
    'HASIL_GAME', 'NILAI', 'PROGRESS', 'LOG_AKTIVITAS', 'PENGATURAN'
  ];

  requiredSheets.forEach(name => {
    if (!sheetNames.includes(name)) {
      ss.insertSheet(name);
    }
  });

  // Inisialisasi headers
  initializeHeaders(ss);
}

function initializeHeaders(ss) {
  const headers = {
    'USERS': ['ID', 'USERNAME', 'PASSWORD_HASH', 'EMAIL', 'ROLE', 'STATUS', 'CREATED_AT', 'UPDATED_AT'],
    'MAHASISWA': ['ID', 'USER_ID', 'NIM', 'NAMA', 'EMAIL', 'NOMOR_HP', 'PRODI', 'KELAS', 'SEMESTER', 'STATUS', 'CREATED_AT'],
    'DOSEN': ['ID', 'NAMA', 'EMAIL', 'NOMOR_HP', 'NIDN', 'STATUS', 'CREATED_AT'],
    'MATA_KULIAH': ['ID', 'KODE_MK', 'NAMA_MK', 'PRODI', 'SEMESTER', 'SKS', 'DOSEN_ID', 'DESKRIPSI', 'STATUS', 'CREATED_AT'],
    'KELAS': ['ID', 'NAMA_KELAS', 'PRODI', 'SEMESTER', 'JUMLAH_MAHASISWA', 'STATUS', 'CREATED_AT'],
    'MATERI': ['ID', 'MATA_KULIAH_ID', 'PERTEMUAN', 'JUDUL', 'DESKRIPSI', 'ISI_MATERI', 'FILE_URL', 'VIDEO_URL', 'PUBLISHED_AT', 'CREATED_AT'],
    'TUGAS': ['ID', 'MATA_KULIAH_ID', 'PERTEMUAN', 'JUDUL', 'DESKRIPSI', 'INSTRUKSI', 'DEADLINE', 'FILE_PENDUKUNG', 'NILAI_MAKSIMAL', 'STATUS', 'CREATED_AT'],
    'PENGUMPULAN_TUGAS': ['ID', 'TUGAS_ID', 'MAHASISWA_ID', 'NIM', 'NAMA', 'FILE_URL', 'WAKTU_KUMPUL', 'STATUS', 'NILAI', 'CATATAN', 'CREATED_AT'],
    'BANK_SOAL': ['ID', 'MATA_KULIAH_ID', 'MATERI_ID', 'TIPE_SOAL', 'PERTANYAAN', 'PILIHAN_JAWABAN', 'JAWABAN_BENAR', 'PEMBAHASAN', 'TINGKAT_KESULITAN', 'BOBOT', 'CREATED_AT'],
    'EVALUASI': ['ID', 'MATA_KULIAH_ID', 'JUDUL', 'DURASI', 'JUMLAH_SOAL', 'NILAI_MAKSIMAL', 'BATAS_PERCOBAAN', 'STATUS', 'CREATED_AT'],
    'EVALUASI_SOAL': ['ID', 'EVALUASI_ID', 'SOAL_ID', 'URUTAN', 'CREATED_AT'],
    'HASIL_EVALUASI': ['ID', 'EVALUASI_ID', 'MAHASISWA_ID', 'SKOR', 'JUMLAH_BENAR', 'JUMLAH_SALAH', 'WAKTU_PENGERJAAN', 'TANGGAL_KERJAKAN', 'CREATED_AT'],
    'GAME': ['ID', 'MATA_KULIAH_ID', 'JUDUL', 'TIPE_GAME', 'JUMLAH_SOAL', 'DURASI', 'STATUS', 'CREATED_AT'],
    'HASIL_GAME': ['ID', 'GAME_ID', 'MAHASISWA_ID', 'SKOR', 'WAKTU_MAIN', 'TANGGAL_MAIN', 'CREATED_AT'],
    'NILAI': ['ID', 'MAHASISWA_ID', 'MATA_KULIAH_ID', 'NILAI_TUGAS', 'NILAI_QUIZ', 'NILAI_GAME', 'NILAI_UTS', 'NILAI_UAS', 'NILAI_AKHIR', 'PREDIKAT', 'CREATED_AT'],
    'PROGRESS': ['ID', 'MAHASISWA_ID', 'MATA_KULIAH_ID', 'PERTEMUAN', 'STATUS_MATERI', 'STATUS_TUGAS', 'STATUS_EVALUASI', 'PERSENTASE', 'UPDATED_AT'],
    'LOG_AKTIVITAS': ['ID', 'USER_ID', 'NAMA', 'AKTIVITAS', 'KETERANGAN', 'TIMESTAMP'],
    'PENGATURAN': ['KUNCI', 'NILAI', 'KETERANGAN']
  };

  for (let sheetName in headers) {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers[sheetName]);
    }
  }

  // Inisialisasi data default
  let settingSheet = ss.getSheetByName('PENGATURAN');
  settingSheet.appendRow(['BOBOT_TUGAS', '20', 'Bobot nilai tugas dalam persen']);
  settingSheet.appendRow(['BOBOT_QUIZ', '15', 'Bobot nilai quiz dalam persen']);
  settingSheet.appendRow(['BOBOT_GAME', '15', 'Bobot nilai game dalam persen']);
  settingSheet.appendRow(['BOBOT_UTS', '25', 'Bobot nilai UTS dalam persen']);
  settingSheet.appendRow(['BOBOT_UAS', '25', 'Bobot nilai UAS dalam persen']);
}

// ================================================
// 2. SISTEM AUTENTIKASI & USER
// ================================================

function hashPassword(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    .reduce((str, chr) => str + ('0' + (chr & 0xFF).toString(16)).slice(-2), '');
}

function registerMahasiswa(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mahasiswaSheet = ss.getSheetByName('MAHASISWA');
    const usersSheet = ss.getSheetByName('USERS');

    // Validasi
    if (!data.nim || !data.nama || !data.email || !data.password) {
      return { success: false, message: 'Semua field harus diisi' };
    }

    // Cek NIM duplikat
    let range = mahasiswaSheet.getDataRange();
    let values = range.getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][2] === data.nim) {
        return { success: false, message: 'NIM sudah terdaftar' };
      }
      if (values[i][3] === data.email) {
        return { success: false, message: 'Email sudah terdaftar' };
      }
    }

    // Cek password match
    if (data.password !== data.konfirmasi_password) {
      return { success: false, message: 'Password tidak cocok' };
    }

    // Buat user
    let userId = Utilities.getUuid();
    let passwordHash = hashPassword(data.password);
    let timestamp = new Date().toISOString();

    usersSheet.appendRow([
      userId,
      data.nim,
      passwordHash,
      data.email,
      'MAHASISWA',
      'AKTIF',
      timestamp,
      timestamp
    ]);

    // Simpan data mahasiswa
    mahasiswaSheet.appendRow([
      Utilities.getUuid(),
      userId,
      data.nim,
      data.nama,
      data.email,
      data.nomor_hp,
      data.prodi,
      data.kelas,
      data.semester,
      'AKTIF',
      timestamp
    ]);

    logActivity(userId, data.nama, 'REGISTRASI', 'Mahasiswa berhasil mendaftar');

    return { success: true, message: 'Registrasi berhasil' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName('USERS');
    const mahasiswaSheet = ss.getSheetByName('MAHASISWA');

    let range = usersSheet.getDataRange();
    let values = range.getValues();
    let passwordHash = hashPassword(password);

    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === username && values[i][2] === passwordHash) {
        let user = {
          id: values[i][0],
          username: values[i][1],
          email: values[i][3],
          role: values[i][4],
          status: values[i][5]
        };

        if (user.status !== 'AKTIF') {
          return { success: false, message: 'Akun tidak aktif' };
        }

        // Ambil data mahasiswa jika role MAHASISWA
        if (user.role === 'MAHASISWA') {
          let mahasiswaRange = mahasiswaSheet.getDataRange();
          let mahasiswaValues = mahasiswaRange.getValues();
          for (let j = 1; j < mahasiswaValues.length; j++) {
            if (mahasiswaValues[j][1] === user.id) {
              user.nim = mahasiswaValues[j][2];
              user.nama = mahasiswaValues[j][3];
              user.prodi = mahasiswaValues[j][6];
              user.kelas = mahasiswaValues[j][7];
              user.semester = mahasiswaValues[j][8];
              break;
            }
          }
        }

        logActivity(user.id, user.username, 'LOGIN', 'User berhasil login');
        return { success: true, user: user };
      }
    }

    return { success: false, message: 'Username atau password salah' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function verifyAdmin(password) {
  // Admin password disimpan di properties script
  let adminPassword = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!adminPassword) {
    return false;
  }
  let inputHash = hashPassword(password);
  return inputHash === hashPassword(adminPassword);
}

// ================================================
// 3. DATA MAHASISWA (ADMIN)
// ================================================

function getAllMahasiswa() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MAHASISWA');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        nim: values[i][2],
        nama: values[i][3],
        email: values[i][4],
        nomor_hp: values[i][5],
        prodi: values[i][6],
        kelas: values[i][7],
        semester: values[i][8],
        status: values[i][9]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function addMahasiswa(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MAHASISWA');
    const usersSheet = ss.getSheetByName('USERS');

    let timestamp = new Date().toISOString();
    let userId = Utilities.getUuid();

    // Buat user otomatis dengan password default
    let defaultPassword = data.nim;
    let passwordHash = hashPassword(defaultPassword);

    usersSheet.appendRow([
      userId,
      data.nim,
      passwordHash,
      data.email,
      'MAHASISWA',
      'AKTIF',
      timestamp,
      timestamp
    ]);

    sheet.appendRow([
      Utilities.getUuid(),
      userId,
      data.nim,
      data.nama,
      data.email,
      data.nomor_hp,
      data.prodi,
      data.kelas,
      data.semester,
      'AKTIF',
      timestamp
    ]);

    logActivity('ADMIN', 'ADMIN', 'TAMBAH_MAHASISWA', 'Tambah mahasiswa: ' + data.nama);
    return { success: true, message: 'Mahasiswa berhasil ditambahkan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function editMahasiswa(id, data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MAHASISWA');
    let range = sheet.getDataRange();
    let values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 3, 1, 8).setValues([[
          data.nim,
          data.nama,
          data.email,
          data.nomor_hp,
          data.prodi,
          data.kelas,
          data.semester,
          data.status
        ]]);
        logActivity('ADMIN', 'ADMIN', 'EDIT_MAHASISWA', 'Edit mahasiswa: ' + data.nama);
        return { success: true, message: 'Mahasiswa berhasil diperbarui' };
      }
    }
    return { success: false, message: 'Mahasiswa tidak ditemukan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function deleteMahasiswa(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MAHASISWA');
    let range = sheet.getDataRange();
    let values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.deleteRow(i + 1);
        logActivity('ADMIN', 'ADMIN', 'HAPUS_MAHASISWA', 'Hapus mahasiswa dengan ID: ' + id);
        return { success: true, message: 'Mahasiswa berhasil dihapus' };
      }
    }
    return { success: false, message: 'Mahasiswa tidak ditemukan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 4. DATA MATA KULIAH (ADMIN)
// ================================================

function getAllMataKuliah() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MATA_KULIAH');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        kode_mk: values[i][1],
        nama_mk: values[i][2],
        prodi: values[i][3],
        semester: values[i][4],
        sks: values[i][5],
        dosen_id: values[i][6],
        deskripsi: values[i][7],
        status: values[i][8]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function addMataKuliah(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MATA_KULIAH');
    let timestamp = new Date().toISOString();

    sheet.appendRow([
      Utilities.getUuid(),
      data.kode_mk,
      data.nama_mk,
      data.prodi,
      data.semester,
      data.sks,
      data.dosen_id || '',
      data.deskripsi || '',
      data.status || 'AKTIF'
    ]);

    logActivity('ADMIN', 'ADMIN', 'TAMBAH_MATAKULIAH', 'Tambah mata kuliah: ' + data.nama_mk);
    return { success: true, message: 'Mata kuliah berhasil ditambahkan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 5. MATERI PERKULIAHAN
// ================================================

function getAllMateri() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MATERI');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        mata_kuliah_id: values[i][1],
        pertemuan: values[i][2],
        judul: values[i][3],
        deskripsi: values[i][4],
        isi_materi: values[i][5],
        file_url: values[i][6],
        video_url: values[i][7],
        published_at: values[i][8]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function addMateri(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MATERI');
    let timestamp = new Date().toISOString();

    sheet.appendRow([
      Utilities.getUuid(),
      data.mata_kuliah_id,
      data.pertemuan,
      data.judul,
      data.deskripsi,
      data.isi_materi,
      data.file_url || '',
      data.video_url || '',
      data.published_at || timestamp
    ]);

    logActivity('ADMIN', 'ADMIN', 'TAMBAH_MATERI', 'Tambah materi: ' + data.judul);
    return { success: true, message: 'Materi berhasil ditambahkan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 6. TUGAS
// ================================================

function getAllTugas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('TUGAS');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        mata_kuliah_id: values[i][1],
        pertemuan: values[i][2],
        judul: values[i][3],
        deskripsi: values[i][4],
        instruksi: values[i][5],
        deadline: values[i][6],
        file_pendukung: values[i][7],
        nilai_maksimal: values[i][8],
        status: values[i][9]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function addTugas(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('TUGAS');
    let timestamp = new Date().toISOString();

    sheet.appendRow([
      Utilities.getUuid(),
      data.mata_kuliah_id,
      data.pertemuan,
      data.judul,
      data.deskripsi,
      data.instruksi,
      data.deadline,
      data.file_pendukung || '',
      data.nilai_maksimal || 100,
      data.status || 'AKTIF'
    ]);

    logActivity('ADMIN', 'ADMIN', 'TAMBAH_TUGAS', 'Tambah tugas: ' + data.judul);
    return { success: true, message: 'Tugas berhasil ditambahkan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 7. PENGUMPULAN TUGAS
// ================================================

function submitTugas(tugas_id, mahasiswa_id, file_url, catatan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENGUMPULAN_TUGAS');
    const tugasSheet = ss.getSheetByName('TUGAS');
    const mahasiswaSheet = ss.getSheetByName('MAHASISWA');

    let timestamp = new Date().toISOString();
    let status = 'DIKUMPULKAN';

    // Cek deadline
    let tugasData = getAllTugas().find(t => t.id === tugas_id);
    let deadline = new Date(tugasData.deadline);
    let now = new Date();

    if (now > deadline) {
      status = 'TERLAMBAT';
    }

    // Ambil data mahasiswa
    let mahasiswaData = getAllMahasiswa().find(m => m.id === mahasiswa_id);

    sheet.appendRow([
      Utilities.getUuid(),
      tugas_id,
      mahasiswa_id,
      mahasiswaData.nim,
      mahasiswaData.nama,
      file_url,
      timestamp,
      status,
      0,
      catatan || '',
      timestamp
    ]);

    logActivity(mahasiswa_id, mahasiswaData.nama, 'KUMPUL_TUGAS', 'Mengumpulkan tugas: ' + tugasData.judul);
    return { success: true, message: 'Tugas berhasil dikumpulkan', status: status };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getPengumpulanTugas(tugas_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENGUMPULAN_TUGAS');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === tugas_id) {
        result.push({
          id: values[i][0],
          mahasiswa_id: values[i][2],
          nim: values[i][3],
          nama: values[i][4],
          file_url: values[i][5],
          waktu_kumpul: values[i][6],
          status: values[i][7],
          nilai: values[i][8],
          catatan: values[i][9]
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function nilaiTugas(pengumpulan_id, nilai, catatan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENGUMPULAN_TUGAS');
    let range = sheet.getDataRange();
    let values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === pengumpulan_id) {
        sheet.getRange(i + 1, 9, 1, 2).setValues([[nilai, catatan]]);
        logActivity('ADMIN', 'ADMIN', 'NILAI_TUGAS', 'Memberikan nilai untuk pengumpulan tugas');
        return { success: true, message: 'Nilai berhasil disimpan' };
      }
    }
    return { success: false, message: 'Pengumpulan tidak ditemukan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 8. BANK SOAL
// ================================================

function getAllBankSoal() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('BANK_SOAL');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        mata_kuliah_id: values[i][1],
        tipe_soal: values[i][3],
        pertanyaan: values[i][4],
        pilihan_jawaban: values[i][5],
        jawaban_benar: values[i][6],
        pembahasan: values[i][7],
        tingkat_kesulitan: values[i][8],
        bobot: values[i][9]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function addBankSoal(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('BANK_SOAL');
    let timestamp = new Date().toISOString();

    sheet.appendRow([
      Utilities.getUuid(),
      data.mata_kuliah_id,
      data.materi_id || '',
      data.tipe_soal,
      data.pertanyaan,
      JSON.stringify(data.pilihan_jawaban) || '',
      data.jawaban_benar,
      data.pembahasan || '',
      data.tingkat_kesulitan || 'SEDANG',
      data.bobot || 1,
      timestamp
    ]);

    logActivity('ADMIN', 'ADMIN', 'TAMBAH_SOAL', 'Tambah soal: ' + data.pertanyaan.substring(0, 50));
    return { success: true, message: 'Soal berhasil ditambahkan' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 9. EVALUASI / QUIZ
// ================================================

function addEvaluasi(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('EVALUASI');
    let timestamp = new Date().toISOString();

    let evaluasiId = Utilities.getUuid();

    sheet.appendRow([
      evaluasiId,
      data.mata_kuliah_id,
      data.judul,
      data.durasi || 60,
      data.jumlah_soal || 10,
      data.nilai_maksimal || 100,
      data.batas_percobaan || 1,
      data.status || 'AKTIF',
      timestamp
    ]);

    // Tambahkan soal ke evaluasi
    if (data.soal_ids && data.soal_ids.length > 0) {
      addSoalToEvaluasi(evaluasiId, data.soal_ids);
    }

    logActivity('ADMIN', 'ADMIN', 'BUAT_EVALUASI', 'Buat evaluasi: ' + data.judul);
    return { success: true, message: 'Evaluasi berhasil dibuat', evaluasiId: evaluasiId };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function addSoalToEvaluasi(evaluasiId, soalIds) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('EVALUASI_SOAL');
    let timestamp = new Date().toISOString();

    soalIds.forEach((soalId, index) => {
      sheet.appendRow([
        Utilities.getUuid(),
        evaluasiId,
        soalId,
        index + 1,
        timestamp
      ]);
    });

    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getEvaluasiSoal(evaluasiId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('EVALUASI_SOAL');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    let bankSoal = getAllBankSoal();

    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === evaluasiId) {
        let soal = bankSoal.find(s => s.id === values[i][2]);
        result.push({
          evaluasi_soal_id: values[i][0],
          soal_id: values[i][2],
          urutan: values[i][3],
          ...soal
        });
      }
    }

    // Randomize soal jika diperlukan
    return result.sort((a, b) => a.urutan - b.urutan);
  } catch (e) {
    return [];
  }
}

function submitHasilEvaluasi(evaluasi_id, mahasiswa_id, hasil, waktu_pengerjaan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HASIL_EVALUASI');
    let timestamp = new Date().toISOString();

    // Hitung skor
    let benar = 0;
    let salah = 0;
    let skor = 0;

    hasil.forEach(h => {
      if (h.jawaban_user === h.jawaban_benar) {
        benar++;
        skor += h.bobot;
      } else {
        salah++;
      }
    });

    sheet.appendRow([
      Utilities.getUuid(),
      evaluasi_id,
      mahasiswa_id,
      skor,
      benar,
      salah,
      waktu_pengerjaan,
      timestamp,
      timestamp
    ]);

    logActivity(mahasiswa_id, mahasiswa_id, 'SELESAI_EVALUASI', 'Menyelesaikan evaluasi dengan skor: ' + skor);
    return { success: true, skor: skor, benar: benar, salah: salah };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 10. GAME EDUKASI
// ================================================

function addGame(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GAME');
    let timestamp = new Date().toISOString();

    let gameId = Utilities.getUuid();

    sheet.appendRow([
      gameId,
      data.mata_kuliah_id,
      data.judul,
      data.tipe_game,
      data.jumlah_soal || 10,
      data.durasi || 60,
      data.status || 'AKTIF',
      timestamp
    ]);

    logActivity('ADMIN', 'ADMIN', 'BUAT_GAME', 'Buat game: ' + data.judul);
    return { success: true, message: 'Game berhasil dibuat', gameId: gameId };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function submitHasilGame(game_id, mahasiswa_id, skor, waktu_main) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HASIL_GAME');
    let timestamp = new Date().toISOString();

    sheet.appendRow([
      Utilities.getUuid(),
      game_id,
      mahasiswa_id,
      skor,
      waktu_main,
      timestamp,
      timestamp
    ]);

    logActivity(mahasiswa_id, mahasiswa_id, 'SELESAI_GAME', 'Menyelesaikan game dengan skor: ' + skor);
    return { success: true, skor: skor };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 11. NILAI MAHASISWA
// ================================================

function getNilaiMahasiswa(mahasiswa_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('NILAI');
    let range = sheet.getDataRange();
    let values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === mahasiswa_id) {
        return {
          id: values[i][0],
          mata_kuliah_id: values[i][2],
          nilai_tugas: values[i][3],
          nilai_quiz: values[i][4],
          nilai_game: values[i][5],
          nilai_uts: values[i][6],
          nilai_uas: values[i][7],
          nilai_akhir: values[i][8],
          predikat: values[i][9]
        };
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function hitungNilaiAkhir(mahasiswa_id, mata_kuliah_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const pengaturanSheet = ss.getSheetByName('PENGATURAN');
    
    // Ambil bobot
    let bobotTugas = 20, bobotQuiz = 15, bobotGame = 15, bobotUts = 25, bobotUas = 25;
    let range = pengaturanSheet.getDataRange();
    let values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === 'BOBOT_TUGAS') bobotTugas = parseInt(values[i][1]);
      if (values[i][0] === 'BOBOT_QUIZ') bobotQuiz = parseInt(values[i][1]);
      if (values[i][0] === 'BOBOT_GAME') bobotGame = parseInt(values[i][1]);
      if (values[i][0] === 'BOBOT_UTS') bobotUts = parseInt(values[i][1]);
      if (values[i][0] === 'BOBOT_UAS') bobotUas = parseInt(values[i][1]);
    }

    // Hitung rata-rata
    let nilaiTugas = getRataRataNilaiTugas(mahasiswa_id, mata_kuliah_id);
    let nilaiQuiz = getRataRataNilaiQuiz(mahasiswa_id, mata_kuliah_id);
    let nilaiGame = getRataRataNilaiGame(mahasiswa_id, mata_kuliah_id);

    let nilaiAkhir = (nilaiTugas * bobotTugas / 100) + 
                     (nilaiQuiz * bobotQuiz / 100) + 
                     (nilaiGame * bobotGame / 100) + 
                     (0 * bobotUts / 100) + 
                     (0 * bobotUas / 100);

    let predikat = getPredikat(nilaiAkhir);

    return { nilaiAkhir: nilaiAkhir, predikat: predikat };
  } catch (e) {
    return { nilaiAkhir: 0, predikat: '' };
  }
}

function getRataRataNilaiTugas(mahasiswa_id, mata_kuliah_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENGUMPULAN_TUGAS');
    const tugasSheet = ss.getSheetByName('TUGAS');

    let tugasList = getAllTugas().filter(t => t.mata_kuliah_id === mata_kuliah_id);
    let totalNilai = 0;
    let jumlah = 0;

    tugasList.forEach(tugas => {
      let range = sheet.getDataRange();
      let values = range.getValues();

      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === tugas.id && values[i][2] === mahasiswa_id) {
          totalNilai += parseInt(values[i][8]) || 0;
          jumlah++;
        }
      }
    });

    return jumlah > 0 ? totalNilai / jumlah : 0;
  } catch (e) {
    return 0;
  }
}

function getRataRataNilaiQuiz(mahasiswa_id, mata_kuliah_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HASIL_EVALUASI');

    let evaluasiList = getAllEvaluasi().filter(e => e.mata_kuliah_id === mata_kuliah_id);
    let totalNilai = 0;
    let jumlah = 0;

    evaluasiList.forEach(evaluasi => {
      let range = sheet.getDataRange();
      let values = range.getValues();

      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === evaluasi.id && values[i][2] === mahasiswa_id) {
          totalNilai += parseInt(values[i][3]) || 0;
          jumlah++;
        }
      }
    });

    return jumlah > 0 ? totalNilai / jumlah : 0;
  } catch (e) {
    return 0;
  }
}

function getRataRataNilaiGame(mahasiswa_id, mata_kuliah_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HASIL_GAME');

    let gameList = getAllGame().filter(g => g.mata_kuliah_id === mata_kuliah_id);
    let totalNilai = 0;
    let jumlah = 0;

    gameList.forEach(game => {
      let range = sheet.getDataRange();
      let values = range.getValues();

      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === game.id && values[i][2] === mahasiswa_id) {
          totalNilai += parseInt(values[i][3]) || 0;
          jumlah++;
        }
      }
    });

    return jumlah > 0 ? totalNilai / jumlah : 0;
  } catch (e) {
    return 0;
  }
}

function getPredikat(nilai) {
  if (nilai >= 85) return 'A';
  if (nilai >= 75) return 'B';
  if (nilai >= 65) return 'C';
  if (nilai >= 55) return 'D';
  return 'E';
}

function getAllEvaluasi() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('EVALUASI');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        mata_kuliah_id: values[i][1],
        judul: values[i][2],
        durasi: values[i][3],
        jumlah_soal: values[i][4],
        nilai_maksimal: values[i][5],
        batas_percobaan: values[i][6],
        status: values[i][7]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function getAllGame() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GAME');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      result.push({
        id: values[i][0],
        mata_kuliah_id: values[i][1],
        judul: values[i][2],
        tipe_game: values[i][3],
        jumlah_soal: values[i][4],
        durasi: values[i][5],
        status: values[i][6]
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

// ================================================
// 12. PROGRESS PERKULIAHAN
// ================================================

function getProgress(mahasiswa_id, mata_kuliah_id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PROGRESS');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let result = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === mahasiswa_id && values[i][2] === mata_kuliah_id) {
        result.push({
          id: values[i][0],
          pertemuan: values[i][3],
          status_materi: values[i][4],
          status_tugas: values[i][5],
          status_evaluasi: values[i][6],
          persentase: values[i][7]
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function updateProgress(mahasiswa_id, mata_kuliah_id, pertemuan, status_materi, status_tugas, status_evaluasi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PROGRESS');
    let range = sheet.getDataRange();
    let values = range.getValues();

    let persentase = 0;
    if (status_materi === 'SELESAI') persentase += 33;
    if (status_tugas === 'SELESAI') persentase += 33;
    if (status_evaluasi === 'SELESAI') persentase += 34;

    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === mahasiswa_id && values[i][2] === mata_kuliah_id && values[i][3] === pertemuan) {
        sheet.getRange(i + 1, 5, 1, 3).setValues([[status_materi, status_tugas, status_evaluasi]]);
        sheet.getRange(i + 1, 8, 1, 1).setValues([[persentase]]);
        return { success: true };
      }
    }

    // Insert baru jika tidak ada
    sheet.appendRow([
      Utilities.getUuid(),
      mahasiswa_id,
      mata_kuliah_id,
      pertemuan,
      status_materi,
      status_tugas,
      status_evaluasi,
      persentase,
      new Date().toISOString()
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ================================================
// 13. LOG AKTIVITAS
// ================================================

function logActivity(user_id, nama, aktivitas, keterangan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('LOG_AKTIVITAS');

    sheet.appendRow([
      Utilities.getUuid(),
      user_id,
      nama,
      aktivitas,
      keterangan,
      new Date().toISOString()
    ]);
  } catch (e) {
    Logger.log('Error logging activity: ' + e.toString());
  }
}

// ================================================
// 14. HELPER & UTILITY
// ================================================

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
}

function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

function setFolderId(id) {
  PropertiesService.getScriptProperties().setProperty('FOLDER_ID', id);
}

function getFolderId() {
  return PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
}

function setAdminPassword(password) {
  let hash = hashPassword(password);
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', hash);
}

function getDashboardStats() {
  try {
    let mahasiswa = getAllMahasiswa();
    let mataKuliah = getAllMataKuliah();
    let materi = getAllMateri();
    let tugas = getAllTugas();
    let evaluasi = getAllEvaluasi();

    return {
      jumlah_mahasiswa: mahasiswa.length,
      jumlah_mata_kuliah: mataKuliah.length,
      jumlah_materi: materi.length,
      jumlah_tugas: tugas.length,
      jumlah_evaluasi: evaluasi.length
    };
  } catch (e) {
    return {};
  }
}
