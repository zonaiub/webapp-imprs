/**
 * =================================================================
 * SISTEM INPUT IMPRS (INDIKATOR MUTU PRIORITAS RS) - BERBASIS DATA MASTER
 * =================================================================
 */

const ARCHIVE_FOLDER_ID = '1AzWvvP_Lo6TBWR06TGHFDsFkT1J5ET5n';

function autoArchiveTahunan() {
  const prevYear = new Date().getFullYear() - 1;
  archiveData(prevYear);
}

function archiveData(yearToArchive) {
  if (!ARCHIVE_FOLDER_ID) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  const rowsToArchive = [];
  const rowsToKeep = [];

  data.forEach(row => {
    if (!row[1]) return;
    new Date(row[1]).getFullYear() === yearToArchive ? rowsToArchive.push(row) : rowsToKeep.push(row);
  });

  if (rowsToArchive.length === 0) return;

  const folder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
  const fileName = "Arsip_IMPRS_" + yearToArchive;
  let archiveSs;
  const existingFiles = folder.searchFiles("title = '" + fileName + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
  
  if (existingFiles.hasNext()) {
    archiveSs = SpreadsheetApp.openById(existingFiles.next().getId());
  } else {
    const newFile = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(newFile.getId()).moveTo(folder);
    archiveSs = SpreadsheetApp.openById(newFile.getId());
    archiveSs.getSheets()[0].appendRow(DATA_HEADERS);
    archiveSs.getSheets()[0].setFrozenRows(1);
  }

  const archiveSheet = archiveSs.getSheets()[0];
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, DATA_HEADERS.length).setValues(rowsToArchive);

  sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).clearContent();
  if (rowsToKeep.length > 0) sheet.getRange(2, 1, rowsToKeep.length, DATA_HEADERS.length).setValues(rowsToKeep);
}

function getRawDataByYears(startYear, endYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allData = [];
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    allData = allData.concat(sheet.getRange(2, 1, sheet.getLastRow() - 1, DATA_HEADERS.length).getValues());
  }
  if (ARCHIVE_FOLDER_ID && ARCHIVE_FOLDER_ID !== 'GANTI_DENGAN_ID_FOLDER_00_ARSIP_DATA_MUTU_DI_SINI') {
    for (let y = startYear; y <= endYear; y++) {
      try {
        const files = DriveApp.getFolderById(ARCHIVE_FOLDER_ID).searchFiles("title = 'Arsip_IMPRS_" + y + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
        if (files.hasNext()) {
          const archSheet = SpreadsheetApp.openById(files.next().getId()).getSheets()[0];
          if (archSheet && archSheet.getLastRow() > 1) {
            allData = allData.concat(archSheet.getRange(2, 1, archSheet.getLastRow() - 1, DATA_HEADERS.length).getValues());
          }
        }
      } catch (e) {}
    }
  }
  const filtered = allData.filter(row => {
    if(!row[1]) return false;
    const y = new Date(row[1]).getFullYear();
    return y >= startYear && y <= endYear;
  });
  const uniqueData = {};
  filtered.forEach(row => {
    const key = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '_' + normText(row[2]) + '_' + normText(row[3]);
    if (!uniqueData[key]) uniqueData[key] = row;
  });
  return Object.values(uniqueData);
}
const CONFIG_SHEET_NAME = 'Config';
const DATA_SHEET_NAME = 'Data Master';
const USERS_SHEET_NAME = 'Users';

const DATA_HEADERS = ['Timestamp', 'Tanggal', 'Ruangan', 'Indikator', 'Numerator', 'Denominator', 'Diisi Oleh', 'Keterangan', 'Status Validasi', 'Populasi', 'Sampel', 'Num Sampel', 'Link Bukti', 'Catatan Validator', 'Akurasi Validasi (%)', 'Analisa Validator', 'RTL Validator'];

const REVERSE_INDICATOR_NAMES = [];

function normText(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function titleCase(s) {
  return String(s || '').trim().toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
}

const INDICATOR_DETAILS = [
  {
    nama: 'Kepatuhan Identifikasi Pasien NAPZA oleh Petugas',
    num: 'Jumlah petugas NAPZA yang melakukan identifikasi pasien secara benar dalam periode observasi',
    den: 'Jumlah petugas NAPZA yang diobservasi dalam periode observasi',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Komunikasi Efektif Tenaga Kesehatan dalam Konseling Pasien NAPZA',
    num: 'Jumlah konseling pasien di ruang napza yang sesuai pedoman pelayanan NAPZA dan terdapat umpan balik',
    den: 'Total konseling pasien di ruang napza yang sesuai program modul dalam periode observasi',
    target: 85,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Penerapan Keamanan Obat High Alert Tergolong Psikotropika',
    num: 'Jumlah pasien yang mendapatkan obat high alert yang tergolong psikotropika sesuai prosedur keamanan',
    den: 'Jumlah total pasien yang mendapatkan obat high alert tergolong psikotropika',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Tepat Lokasi, Prosedur, Pasien pada Pelaksanaan ECT',
    num: 'Jumlah tindakan ECT yang dilakukan sesuai prosedur',
    den: 'Total tindakan ECT yang dilakukan',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Petugas Melakukan 5 Momen Cuci Tangan',
    num: 'Jumlah petugas yang melakukan 5 momen cuci tangan',
    den: 'Total tindakan yang diobservasi sesuai periode',
    target: 85,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Skrining dan Penanda Risiko Jatuh di Rawat Jalan',
    num: 'Jumlah pasien rawat jalan berisiko tinggi jatuh yang mendapatkan intervensi risiko jatuh',
    den: 'Jumlah pasien rawat jalan berisiko tinggi jatuh yang diobservasi',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Assessment IPWL ASI Minggu 1 Masa Rehabilitasi NAPZA',
    num: 'Jumlah pasien yang dilakukan asesmen IPWL ASI pada minggu 1 masa rehab',
    den: 'Total seluruh pasien yang ada di rehab napza',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Pasien Ketergantungan Stimulan dengan Perbaikan Kualitas Hidup',
    num: 'Jumlah pasien ketergantungan stimulan yang selesai program dengan peningkatan skor WHO QoL ≥ 20%',
    den: 'Total pasien ketergantungan stimulan yang selesai program',
    target: 70,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Kontrak Klinis Dokter Mitra',
    num: 'Jumlah dokter mitra yang memiliki dan mematuhi kontrak',
    den: 'Total dokter mitra yang memberikan pelayanan',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Kepatuhan Skrining dan Penandaan Fast Track di Rawat Jalan',
    num: 'Jumlah petugas yang patuh dalam melakukan skrining dan penandaan fast track',
    den: 'Total petugas fast track pada periode observasi',
    target: 100,
    tipe: 'persen'
  },
  {
    nama: 'Peningkatan Jejaring Rehabilitasi NAPZA',
    num: 'Jumlah MOU (kerja sama) jejaring rehabilitasi NAPZA baru dalam periode observasi',
    den: '',
    target: 2,
    tipe: 'hitungan'
  },
  {
    nama: 'Terselenggaranya Skrining Napza di Rawat Jalan',
    num: 'Jumlah pasien napza yang terskrining ASSIST selama rawat jalan',
    den: 'Jumlah pasien napza yang dirawat jalan',
    target: 80,
    tipe: 'persen'
  },
  {
    nama: 'Penelitian oleh Staf RS yang Dipublikasikan',
    num: 'Jumlah penelitian oleh staf RS yang dipublikasikan dalam periode observasi',
    den: '',
    target: 3,
    tipe: 'hitungan'
  },
  {
    nama: 'Artikel Kesehatan Jiwa yang Diterbitkan di Website RSJ',
    num: 'Jumlah artikel terkait kesehatan jiwa yang diterbitkan dalam website RSJ pada periode observasi',
    den: '',
    target: 2,
    tipe: 'hitungan'
  }
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Input Data IMPRS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET_NAME);
  }
  if (usersSheet.getLastRow() > 0) {
    Logger.log('Sheet Users sudah ada isinya, tidak ditimpa.');
    return;
  }

  usersSheet.appendRow(['Username', 'Password', 'Role', 'Ruangan']);
  usersSheet.setFrozenRows(1);

  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = configSheet.getLastRow();
  const rooms = configSheet.getRange('A2:A' + lastRow).getValues().flat().filter(String);

  const rows = rooms.map(function (room) {
    const username = room.toLowerCase().replace(/[^a-z0-9]/g, '');
    return [username, 'inm123', 'pic', room];
  });
  rows.push(['admin', 'admin123', 'admin', '']);

  usersSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function login(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Sistem login belum di-setup.' };

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputPass = String(password || '').trim();

  for (let i = 0; i < data.length; i++) {
    const rowUser = String(data[i][0] || '').trim().toLowerCase();
    const rowPass = String(data[i][1] || '').trim();
    if (rowUser === inputUser && rowPass === inputPass) {
      return { success: true, role: data[i][2], ruangan: data[i][3] };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

/**
 * Jalankan SEKALI SAJA (manual dari editor Apps Script) kalau sheet "Data Master"
 * dibuat sebelum kolom "Akurasi Validasi (%)" ditambahkan, supaya header di
 * baris 1 ikut terisi. Aman dijalankan berkali-kali.
 */
/**
 * Jalankan SEKALI SAJA (manual dari editor Apps Script) untuk menambahkan
 * kolom "Arah Target" di sheet Config. Setelah ini ada, admin bisa atur
 * arah target per indikator LANGSUNG dari spreadsheet (isi "Minimal" untuk
 * target ≥ atau "Maksimal" untuk target ≤), tanpa perlu edit kode lagi.
 * Kalau kolom ini kosong/belum ada, sistem tetap jalan seperti biasa
 * (pakai daftar REVERSE_INDICATOR_NAMES di kode sebagai default).
 */
function tambahKolomArahTarget() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headerRow.indexOf('Arah Target') > -1) return; // sudah ada, tidak perlu ditambah lagi
  sheet.getRange(1, lastCol + 1).setValue('Arah Target');
}

function tambahKolomAkurasi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  sheet.getRange(1, DATA_HEADERS.length - 2).setValue('Akurasi Validasi (%)');
  sheet.getRange(1, DATA_HEADERS.length - 1).setValue('Analisa Validator');
  sheet.getRange(1, DATA_HEADERS.length).setValue('RTL Validator');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(DATA_SHEET_NAME);
  }
  if (dataSheet.getLastRow() === 0) {
    dataSheet.appendRow(DATA_HEADERS);
    dataSheet.setFrozenRows(1);
  }

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['Daftar Ruangan', 'Daftar Indikator', 'Keterangan Numerator', 'Keterangan Denominator', 'Target', 'Tipe']);
    const rooms = [
      'ARIMBI', 'BROTOJOYO', 'CITROANGGODO', 'DEWARUCI', 'JANOKO', 'GATUTKACA',
      'HUDOWO', 'ECT', 'UPIP', 'RIPD', 'IGD', 'DIKLAT', 'NAPZA', 'BISMA',
      'ENDROTENOYO', 'IBS', 'ICU', 'SRIKANDI A', 'SRIKANDI B', 'FARMASI',
      'RAMSIN', 'RAJAL', 'YUDISTIRA', 'TUKEPEG'
    ];
    const maxLen = Math.max(rooms.length, INDICATOR_DETAILS.length);
    for (let i = 0; i < maxLen; i++) {
      configSheet.getRange(i + 2, 1).setValue(rooms[i] || '');
      if (INDICATOR_DETAILS[i]) {
        configSheet.getRange(i + 2, 2).setValue(INDICATOR_DETAILS[i].nama);
        configSheet.getRange(i + 2, 3).setValue(INDICATOR_DETAILS[i].num);
        configSheet.getRange(i + 2, 4).setValue(INDICATOR_DETAILS[i].den);
        configSheet.getRange(i + 2, 5).setValue(INDICATOR_DETAILS[i].target);
        configSheet.getRange(i + 2, 6).setValue(INDICATOR_DETAILS[i].tipe);
      }
    }
  }
}

function getRooms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const values = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(String);
  return values;
}

function getIndicatorDetails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 6);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxArah = headerRow.indexOf('Arah Target'); // -1 kalau kolom ini belum ditambahkan admin
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data
    .filter(function (row) { return row[1]; })
    .map(function (row) {
      const nama = row[1];
      const target = (row[4] === '' || row[4] === undefined) ? null : Number(row[4]);
      let reverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(nama)) > -1;
      // Kalau admin sudah isi kolom "Arah Target" di Config, itu yang dipakai
      // (lebih diprioritaskan daripada daftar hardcode REVERSE_INDICATOR_NAMES).
      if (idxArah > -1 && row[idxArah]) {
        const arah = normText(row[idxArah]);
        reverse = (arah.indexOf('maksimal') > -1 || arah.indexOf('turun') > -1 || arah.indexOf('kecil') > -1 || arah.indexOf('rendah') > -1);
      }
      return { nama: nama, num: row[2] || '', den: row[3] || '', target: target, tipe: row[5] || 'persen', reverse: reverse };
    });
}

function getExistingEntries(ruangan, tanggal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const targetYear = new Date(tanggal).getFullYear();
  const data = getRawDataByYears(targetYear, targetYear);
  const result = {};
  const targetRuangan = normText(ruangan);
  data.forEach(function (row) {
    const rowTanggal = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowTanggal === tanggal && normText(row[2]) === targetRuangan) {
      result[normText(row[3])] = { numerator: row[4], denominator: row[5], keterangan: row[7] || '' };
    }
  });
  return result;
}

function submitData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();

  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues()
    : [];

  const tanggal = payload.tanggal;
  const ruangan = payload.ruangan;
  const targetRuangan = normText(ruangan);
  const now = new Date();

 payload.entries.forEach(function (entry) {
    if (entry.numerator === '' && entry.denominator === '') return;

    const targetIndikator = normText(entry.indikator);
    let foundRowIndex = -1;
    for (let i = 0; i < existing.length; i++) {
      const rowTanggal = Utilities.formatDate(new Date(existing[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowTanggal === tanggal && normText(existing[i][2]) === targetRuangan && normText(existing[i][3]) === targetIndikator) {
        foundRowIndex = i + 2;
        break;
      }
    }

    const rowValues = [
      now, 
      new Date(tanggal), 
      ruangan, 
      entry.indikator, 
      Number(entry.numerator) || 0, 
      Number(entry.denominator) || 0, 
      payload.diisiOleh || '', 
      entry.keterangan || '',
      '⏳ Menunggu', '', '', '', '', '', '', '', ''
    ];
    
    if (foundRowIndex > -1) {
      sheet.getRange(foundRowIndex, 1, 1, DATA_HEADERS.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  });

  cekDanArsipOtomatis(); 

  return { status: 'ok', message: 'Data tersimpan.' };
}

function getRecapMatrix(mode, params) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const rooms = configSheet.getRange('A2:A' + lastConfigRow).getValues().flat().filter(String);
  const configValues = configSheet.getRange('B2:F' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const indicatorList = configValues.map(function (r) { return { nama: r[0], tipe: r[4] || 'persen' }; });

  let targetYear;
  if (mode === 'harian') targetYear = new Date(params.tanggal).getFullYear();
  else targetYear = Number(params.tahun);
  const data = getRawDataByYears(targetYear, targetYear);

  const totals = {};
  const roomPics = {};

  data.forEach(function (row) {
    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      const rowTanggal = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      periodMatch = rowTanggal === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun) &&
        minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun);
    }
    if (!periodMatch) return;

    const indKey = normText(row[3]);
    const roomKey = normText(row[2]);
    if (!totals[indKey]) totals[indKey] = {};
    if (!totals[indKey][roomKey]) totals[indKey][roomKey] = { numerator: 0, denominator: 0 };
    totals[indKey][roomKey].numerator += Number(row[4]) || 0;
    totals[indKey][roomKey].denominator += Number(row[5]) || 0;

    const picName = String(row[6] || '').trim();
    if (picName) {
      if (!roomPics[roomKey]) roomPics[roomKey] = {};
      roomPics[roomKey][picName.toUpperCase()] = picName;
    }
  });

  const rows = indicatorList.map(function (ind) {
    const indKey = normText(ind.nama);
    const values = rooms.map(function (room) {
      const roomKey = normText(room);
      const t = totals[indKey] && totals[indKey][roomKey];
      if (!t) return null;
      if (ind.tipe === 'hitungan') return t.numerator;
      return t.denominator > 0 ? Math.round((t.numerator / t.denominator) * 1000) / 10 : null;
    });
    let totalNum = 0, totalDen = 0;
    if (totals[indKey]) {
      Object.keys(totals[indKey]).forEach(function (roomKey) {
        totalNum += totals[indKey][roomKey].numerator;
        totalDen += totals[indKey][roomKey].denominator;
      });
    }
    return { indikator: ind.nama, tipe: ind.tipe, values: values, totalNum: totalNum, totalDen: totalDen };
  });

  const pics = rooms.map(function (room) {
    const roomKey = normText(room);
    const p = roomPics[roomKey];
    if (!p) return '';
    return Object.keys(p).sort().map(function (k) { return titleCase(p[k]); }).join(', ');
  });

  return { rooms: rooms, rows: rows, pics: pics };
}

function getRecap(ruangan, mode, params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const configValues = configSheet.getRange('B2:F' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const targetMap = {};
  const tipeMap = {};
  configValues.forEach(function (r) {
    targetMap[normText(r[0])] = (r[3] === '' || r[3] === undefined) ? null : Number(r[3]);
    tipeMap[normText(r[0])] = r[4] || 'persen';
  });
  const REVERSE_INDICATORS = REVERSE_INDICATOR_NAMES.map(normText);

  let targetYear;
  if (mode === 'harian') targetYear = new Date(params.tanggal).getFullYear();
  else targetYear = Number(params.tahun);
  const data = getRawDataByYears(targetYear, targetYear);
  const totals = {};
  const targetRuangan = normText(ruangan);

  data.forEach(function (row) {
    const roomMatch = (ruangan === 'SEMUA') || (normText(row[2]) === targetRuangan);
    if (!roomMatch) return;

    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      const rowTanggal = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      periodMatch = rowTanggal === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun) &&
        minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun);
    }
    if (!periodMatch) return;

    const ind = normText(row[3]);
    if (!totals[ind]) totals[ind] = { numerator: 0, denominator: 0, pics: {}, alasanList: [] };
    totals[ind].numerator += Number(row[4]) || 0;
    totals[ind].denominator += Number(row[5]) || 0;
    const picName = String(row[6] || '').trim();
    if (picName) totals[ind].pics[picName.toUpperCase()] = picName;
    const ket = String(row[7] || '').trim();
    if (ket) {
      const tglStr = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM');
      totals[ind].alasanList.push(tglStr + ': ' + ket);
    }
  });

  return configValues
    .map(function (r) { return r[0]; })
    .filter(function (ind) { return totals[normText(ind)]; })
    .map(function (ind) {
      const key = normText(ind);
      const t = totals[key];
      const tipe = tipeMap[key] || 'persen';
      const target = targetMap[key] !== undefined ? targetMap[key] : null;
      const picList = Object.keys(t.pics).sort().map(function (k) { return titleCase(t.pics[k]); }).join(', ');
      const alasan = t.alasanList.join(' | ');
      let hasil, memenuhi = null;

      if (tipe === 'hitungan') {
        hasil = t.numerator;
        if (hasil !== null && target !== null) {
          memenuhi = hasil >= target;
        }
      } else {
        hasil = t.denominator > 0 ? (t.numerator / t.denominator * 100) : null;
        if (hasil !== null && target !== null) {
          memenuhi = REVERSE_INDICATORS.indexOf(key) > -1 ? (hasil <= target) : (hasil >= target);
        }
      }

      return { indikator: ind, numerator: t.numerator, denominator: t.denominator, hasil: hasil, target: target, memenuhi: memenuhi, tipe: tipe, pic: picList, alasan: alasan };
    });
}

function computeReportData(startDateStr, endDateStr) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const configValues = configSheet.getRange('B2:F' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const indicatorList = configValues.map(function (r) {
    return {
      nama: r[0],
      target: r[3] === '' || r[3] === undefined ? null : Number(r[3]),
      tipe: r[4] || 'persen'
    };
  });

  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  endDate.setHours(23, 59, 59, 999);

  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];

  const isQ1Only = startDate.getMonth() === 0 && startDate.getDate() === 1 &&
    endDate.getMonth() === 2 && endDate.getDate() >= 28 &&
    startDate.getFullYear() === endDate.getFullYear();

  const buckets = [];

  if (isQ1Only) {
    let cursor = new Date(startDate);
    let weekNum = 1;
    while (cursor <= endDate) {
      const bStart = new Date(cursor);
      const bEnd = new Date(cursor);
      bEnd.setDate(bEnd.getDate() + 6);
      bEnd.setHours(23, 59, 59, 999);
      if (bEnd > endDate) bEnd.setTime(endDate.getTime());
      buckets.push({ label: 'Minggu ' + weekNum, start: bStart, end: bEnd });
      cursor.setDate(cursor.getDate() + 7);
      weekNum++;
    }
  } else {
    const rangeDays = (endDate - startDate) / (1000 * 60 * 60 * 24);
    if (rangeDays <= 35) {
      let cursor = new Date(startDate);
      let weekNum = 1;
      while (cursor <= endDate) {
        const bStart = new Date(cursor);
        const bEnd = new Date(cursor);
        bEnd.setDate(bEnd.getDate() + 6);
        bEnd.setHours(23, 59, 59, 999);
        if (bEnd > endDate) bEnd.setTime(endDate.getTime());
        buckets.push({ label: 'Minggu ' + weekNum, start: bStart, end: bEnd });
        cursor.setDate(cursor.getDate() + 7);
        weekNum++;
      }
    } else {
      let cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
      while (cursor <= endDate) {
        const bStart = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
        const bEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0, 23, 59, 59, 999);
        buckets.push({ label: monthNames[cursor.getMonth()] + ' ' + cursor.getFullYear(), start: bStart, end: bEnd });
        cursor.setMonth(cursor.getMonth() + 1);
      }
    }
  }

  const rawData = getRawDataByYears(startDate.getFullYear(), endDate.getFullYear());

  return indicatorList.map(function (ind) {
    const bucketResults = buckets.map(function (b) {
      let num = 0, den = 0;
      rawData.forEach(function (row) {
        if (normText(row[3]) !== normText(ind.nama)) return;
        const tgl = new Date(row[1]);
        if (tgl >= b.start && tgl <= b.end) {
          num += Number(row[4]) || 0;
          den += Number(row[5]) || 0;
        }
      });
      const hasil = ind.tipe === 'hitungan'
        ? num
        : (den > 0 ? Math.round((num / den) * 10000) / 100 : null);
      return { label: b.label, hasil: hasil };
    });

    const alasanList = [];
    rawData.forEach(function (row) {
      if (normText(row[3]) !== normText(ind.nama)) return;
      const tgl = new Date(row[1]);
      if (tgl < startDate || tgl > endDate) return;
      const ket = String(row[7] || '').trim();
      if (ket) {
        const tglStr = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM');
        alasanList.push(tglStr + ' (' + row[2] + '): ' + ket);
      }
    });

    return { nama: ind.nama, target: ind.target, tipe: ind.tipe, buckets: bucketResults, alasan: alasanList.join('\n') };
  });
}

/**
 * Palet warna & ikon untuk mempercantik slide PPT supaya tiap indikator
 * punya "identitas" visual sendiri (tapi tetap simple, grafik tetap fokus utama).
 * Dipilih otomatis & konsisten berdasarkan nama indikatornya (hash sederhana),
 * jadi indikator yang sama akan selalu dapat warna yang sama tiap kali export.
 */
const SLIDE_THEMES = [
  { accent: '#3d85c6', tint: '#eaf2fb' }, // biru
  { accent: '#6aa84f', tint: '#eef7ec' }, // hijau
  { accent: '#8e63ce', tint: '#f2ecfa' }, // ungu
  { accent: '#e69138', tint: '#fdf1e3' }, // oranye
  { accent: '#45818e', tint: '#e9f4f5' }, // teal
  { accent: '#c27ba0', tint: '#faeef4' }  // pink
];

function getSlideTheme(nama) {
  const s = String(nama || '');
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash += s.charCodeAt(i);
  return SLIDE_THEMES[hash % SLIDE_THEMES.length];
}

/**
 * Pilih ikon kecil berdasarkan kata kunci di nama indikator, biar slide
 * terasa "sesuai judul" tanpa perlu desain rumit. Kalau tidak ada kata kunci
 * yang cocok, pakai ikon umum.
 */
function getIndicatorIcon(nama) {
  const n = String(nama || '').toLowerCase();
  if (n.indexOf('cuci tangan') > -1 || n.indexOf('tangan') > -1) return '🧼';
  if (n.indexOf('napza') > -1 || n.indexOf('rehabilitasi') > -1) return '🧠';
  if (n.indexOf('jatuh') > -1) return '🛡️';
  if (n.indexOf('obat') > -1 || n.indexOf('psikotropika') > -1) return '💊';
  if (n.indexOf('identifikasi pasien') > -1) return '🪪';
  if (n.indexOf('ect') > -1) return '⚕️';
  if (n.indexOf('penelitian') > -1 || n.indexOf('dipublikasikan') > -1) return '🔬';
  if (n.indexOf('artikel') > -1 || n.indexOf('website') > -1) return '📰';
  if (n.indexOf('jejaring') > -1 || n.indexOf('mou') > -1 || n.indexOf('kontrak') > -1) return '🤝';
  if (n.indexOf('konseling') > -1 || n.indexOf('komunikasi') > -1) return '💬';
  if (n.indexOf('skrining') > -1 || n.indexOf('fast track') > -1) return '📋';
  return '📊';
}

function buildPptPresentation(startDateStr, endDateStr) {
  const reportData = computeReportData(startDateStr, endDateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChart_' + new Date().getTime());

  const presTitle = 'Laporan Indikator Mutu ' + startDateStr + ' sd ' + endDateStr;
  const pres = SlidesApp.create(presTitle);

    const titleSlide = pres.getSlides()[0];
    titleSlide.getShapes().forEach(function (sh) {
      try { sh.remove(); } catch (e) { /* ignore */ }
    });
    
    titleSlide.getBackground().setSolidFill('#0f172a'); 
    const pageWidth = pres.getPageWidth();
    
    try {
      const logoBlob = DriveApp.getFileById('1Ityk6FqSmCIm7ij-SEgtCju_pLgnGz1B').getBlob(); 
      const logoImg = titleSlide.insertImage(logoBlob);
      logoImg.setWidth(130).setHeight(130);
      logoImg.setLeft((pageWidth - 130) / 2); 
      logoImg.setTop(45);
    } catch(e) {
      Logger.log("Logo tidak ditemukan/ID salah: " + e.message);
    }

    const titleBox = titleSlide.insertTextBox('LAPORAN INDIKATOR MUTU PRIORITAS RS (IMPRS)', 40, 200, pageWidth - 80, 100);
    titleBox.getText().getTextStyle().setFontSize(32).setBold(true).setForegroundColor('#ffffff');
    titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    const dStart = new Date(startDateStr);
    const dEnd = new Date(endDateStr);
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const dateText = dStart.getDate() + ' ' + months[dStart.getMonth()] + ' ' + dStart.getFullYear() + ' - ' + 
                     dEnd.getDate() + ' ' + months[dEnd.getMonth()] + ' ' + dEnd.getFullYear();

    const subBox = titleSlide.insertTextBox(dateText, 40, 320, pageWidth - 80, 40);
    subBox.getText().getTextStyle().setFontSize(18).setForegroundColor('#94a3b8');
    subBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  let chartRow = 1;
  reportData.forEach(function (ind) {
    const headerRow = chartRow;
    tempSheet.getRange(headerRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
    const rows = ind.buckets.map(function (b) {
      return [b.label, b.hasil, ind.target];
    });
    tempSheet.getRange(headerRow + 1, 1, rows.length, 3).setValues(rows);

    const dataRange = tempSheet.getRange(headerRow, 1, rows.length + 1, 3);
    const chartBuilder = tempSheet.newChart()
      .asComboChart()
      .addRange(dataRange)
      .setNumHeaders(1)
      .setOption('title', '')
      .setOption('legend', { position: 'bottom' })
      .setOption('series', {
        0: { type: 'bars', color: '#f4b400', dataLabel: 'value' },
        1: { type: 'line', color: '#ea4335', lineWidth: 2, pointSize: 0 }
      })
      .setPosition(headerRow, 6, 0, 0);

    // KODE TAMBAHAN: Paksa sumbu Y mulai dari 0.
    // Jika indikator berupa persen, batas atas dikunci di 100.
    chartBuilder.setOption('vAxis.viewWindow.min', 0);
    if (ind.tipe === 'persen') {
      chartBuilder.setOption('vAxis.viewWindow.max', 100);
    }

    const chart = chartBuilder.build();
    tempSheet.insertChart(chart);
    SpreadsheetApp.flush();

    const chartsOnSheet = tempSheet.getCharts();
    const embeddedChart = chartsOnSheet[chartsOnSheet.length - 1];

    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    const theme = getSlideTheme(ind.nama);
    const icon = getIndicatorIcon(ind.nama);

    // Latar belakang slide diberi warna pastel sesuai tema indikator,
    // supaya tidak putih polos & kosong, tapi tetap ringan (tidak mengganggu grafik).
    slide.getBackground().setSolidFill(theme.tint);

    const pageWidth = pres.getPageWidth();
    const pageHeight = pres.getPageHeight();
    const margin = 20;
    const contentWidth = pageWidth - margin * 2;
    const bottomMargin = 20;
    const noteBoxHeight = 60;
    const gapBeforeNote = 15;
    const chartTop = 85; 
    const noteBoxTop = pageHeight - bottomMargin - noteBoxHeight;
    const chartHeight = noteBoxTop - gapBeforeNote - chartTop;

    // Aksen tipis di sisi kiri slide (tidak menimpa area chart/banner),
    // sekadar biar slide terasa punya identitas visual per indikator.
    const accentBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, margin * 0.6, pageHeight);
    accentBar.getFill().setSolidFill(theme.accent);
    accentBar.getBorder().setTransparent();

    const leftBannerWidth = 100;
    const targetBannerWidth = 150;
    const centerBannerWidth = contentWidth - leftBannerWidth - targetBannerWidth;
    const bannerHeight = 60;

    const bannerLeft = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, 15, leftBannerWidth, bannerHeight);
    bannerLeft.getFill().setSolidFill('#4a86c8');
    bannerLeft.getBorder().setTransparent();
    bannerLeft.getText().setText('IMPRS');
    bannerLeft.getText().getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerLeft.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerLeft.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    const bannerCenter = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBannerWidth, 15, centerBannerWidth, bannerHeight);
    bannerCenter.getFill().setSolidFill(theme.accent);
    bannerCenter.getBorder().setTransparent();
    bannerCenter.getText().setText(icon + ' ' + ind.nama);
    bannerCenter.getText().getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(13);
    bannerCenter.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerCenter.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    const isReverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(ind.nama)) > -1;
    let targetValText = '-';
    if (ind.target !== null && ind.target !== undefined && ind.target !== '') {
      targetValText = (ind.tipe === 'hitungan') ? ind.target : (isReverse ? '≤ ' : '≥ ') + ind.target + '%';
    }

    const bannerTarget = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBannerWidth + centerBannerWidth, 15, targetBannerWidth, bannerHeight);
    bannerTarget.getFill().setSolidFill('#c5221f');
    bannerTarget.getBorder().setTransparent();
    bannerTarget.getText().setText('Target: ' + targetValText);
    bannerTarget.getText().getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerTarget.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerTarget.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    slide.insertSheetsChart(embeddedChart, margin, chartTop, contentWidth, chartHeight);

    const noteBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, noteBoxTop, contentWidth, noteBoxHeight);
    noteBox.getFill().setSolidFill('#d9ead3');
    noteBox.getBorder().setTransparent();
    const analisaText = ind.alasan
      ? 'ANALISA (alasan dari ruangan):\n' + ind.alasan
      : 'ANALISA : [isi analisa di sini]';
    noteBox.getText().setText(analisaText + '\nRTL : [isi rencana tindak lanjut di sini]');
    noteBox.getText().getTextStyle().setFontSize(11).setBold(true);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    chartRow += rows.length + 3;
  });

  SpreadsheetApp.flush();

  return {
    presId: pres.getId(),
    filename: presTitle.replace(/[^a-zA-Z0-9]/g, '_') + '.pptx',
    tempSheetName: tempSheet.getName()
  };
}

function exportPresentationToPptx(presId, filename, tempSheetName) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + presId +
    '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const blob = response.getBlob();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  try { DriveApp.getFileById(presId).setTrashed(true); } catch (e) {}

  if (responseCode !== 200) {
    throw new Error('Gagal export PPT (kode ' + responseCode + '). Coba lagi.');
  }

  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: filename
  };
}

function getMissingRooms(bulan, tahun) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const allRooms = configSheet.getRange('A2:A' + lastConfigRow).getValues().flat().filter(String);

  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const roomsWithData = {};

  const rawData = getRawDataByYears(Number(tahun), Number(tahun));
  rawData.forEach(function (row) {
    const tgl = new Date(row[1]);
    if ((tgl.getMonth() + 1) === Number(bulan) && tgl.getFullYear() === Number(tahun)) {
      roomsWithData[normText(row[2])] = true;
    }
  });
  
  return allRooms.filter(function (room) { return !roomsWithData[normText(room)]; });
}

function buildFormattedSheet(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempExport_' + new Date().getTime());
  const numCols = headers.length;

  tempSheet.getRange(1, 1, 1, numCols).merge();
  tempSheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  tempSheet.getRange(2, 1, 1, numCols).setValues([headers])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#f0f2f5');

  if (rows.length > 0) {
    tempSheet.getRange(3, 1, rows.length, numCols).setValues(rows)
      .setHorizontalAlignment('center');
  }

  const totalRows = rows.length + 2;
  tempSheet.getRange(2, 1, totalRows - 1, numCols)
    .setBorder(true, true, true, true, true, true);

  tempSheet.autoResizeColumns(1, numCols);
  SpreadsheetApp.flush();
  return tempSheet;
}

function exportSheetAs(tempSheet, format) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = tempSheet.getSheetId();
  let url;
  if (format === 'pdf') {
    url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
      '/export?format=pdf&gid=' + sheetId +
      '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&attachment=true&horizontal_alignment=CENTER';
  } else {
    url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
      '/export?format=xlsx&gid=' + sheetId;
  }
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return response.getBlob();
}

function exportRecapPdf(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'pdf');
  ss.deleteSheet(tempSheet);
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf'
  };
}

function exportRecapExcel(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'xlsx');
  ss.deleteSheet(tempSheet);
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.xlsx'
  };
}

function resetPassword(username, ruangan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputRuangan = normText(ruangan);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === inputUser) {
      if (normText(data[i][3]) === inputRuangan) {
        sheet.getRange(i + 1, 2).setValue('inm123');
        return { success: true, message: 'Berhasil! Password direset menjadi: inm123' };
      } else {
        return { success: false, message: 'Gagal: Username dan Ruangan tidak cocok!' };
      }
    }
  }
  return { success: false, message: 'Gagal: Username tidak ditemukan.' };
}

function changePassword(username, oldPass, newPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === inputUser) {
      if (String(data[i][1]) === String(oldPass)) {
        sheet.getRange(i + 1, 2).setValue(newPass);
        return { success: true, message: "Berhasil! Password berhasil diubah." };
      } else {
        return { success: false, message: "Gagal: Password lama salah!" };
      }
    }
  }
  return { success: false, message: "Gagal: Username tidak ditemukan." };
}

function cekDanArsipOtomatis() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DATA_SHEET_NAME); 
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const dateValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const currentYear = new Date().getFullYear();
    let yearsToArchive = [];

    dateValues.forEach(row => {
      if (row[0]) {
        const y = new Date(row[0]).getFullYear();
        if (y !== currentYear && yearsToArchive.indexOf(y) === -1) {
          yearsToArchive.push(y);
        }
      }
    });

    yearsToArchive.forEach(y => {
      archiveData(y);
    });

  } catch (e) {
    Logger.log("Gagal auto-arsip: " + e.message);
  }
}

function getPendingValidations(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data Master');
  var data = sheet.getDataRange().getValues();
  
  var sheetConfig = ss.getSheetByName('Config');
  var configData = sheetConfig.getDataRange().getValues();
  var headerConfig = configData[0];
  var idxInd = headerConfig.indexOf('Daftar Indikator');
  var idxVal = headerConfig.indexOf('Username Validator');
  
  var valMap = {};
  if(idxInd > -1 && idxVal > -1) {
    for (var i = 1; i < configData.length; i++) {
      var indikator = configData[i][idxInd];
      var penilai = configData[i][idxVal];
      if (indikator) {
        valMap[indikator.toString().trim()] = penilai ? penilai.toString().trim().toLowerCase() : "";
      }
    }
  }

  var usr = username ? username.toString().trim().toLowerCase() : "";
  var pendingGroups = {};
  var pendingArray = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var statusStr = row[8] ? row[8].toString().trim() : "";
    
    // Tarik data yang masih nunggu atau kosong
    if (statusStr === '⏳ Menunggu' || statusStr === '') {
      var indName = row[3] ? row[3].toString().trim() : "";
      var hakVal = valMap[indName] || "";
      
      if (hakVal === usr || usr === 'admin') {
        var tglRaw = row[1];
        var monthYear = "";
        
        // Konversi tanggal jadi format "Bulan Tahun" (misal: Agustus 2026)
        if (tglRaw && tglRaw instanceof Date) {
          var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
          monthYear = months[tglRaw.getMonth()] + " " + tglRaw.getFullYear();
        } else if (tglRaw) {
          monthYear = tglRaw.toString();
        }

        var ruangan = row[2] ? row[2].toString().trim() : "";
        
        // Bikin kunci grup biar data di bulan, ruangan, dan indikator yang sama jadi 1 kartu
        var key = monthYear + "_" + ruangan + "_" + indName;
        
        if (!pendingGroups[key]) {
          pendingGroups[key] = {
            periode: monthYear,
            ruangan: ruangan,
            indikator: indName,
            numerator: 0,
            denominator: 0,
            rowIndexes: [] // Menyimpan semua baris data hariannya
          };
        }
        
        pendingGroups[key].numerator += (Number(row[4]) || 0);
        pendingGroups[key].denominator += (Number(row[5]) || 0);
        pendingGroups[key].rowIndexes.push(i + 1);
      }
    }
  }
  
  for (var k in pendingGroups) {
    pendingArray.push(pendingGroups[k]);
  }
  
  return pendingArray;
}

/**
 * Mengambil daftar Catatan Validator untuk 1 ruangan tertentu, supaya
 * bisa ditampilkan ke PIC ruangan itu saat login. Hanya catatan dengan
 * status "❌ Kembalikan" yang ditampilkan.
 * Catatan disimpan sama di semua baris harian dalam 1 bulan validasi,
 * jadi dikelompokkan per Indikator + Bulan + Tahun.
 */
function getCatatanUntukRuangan(ruangan) {
  var currentYear = new Date().getFullYear();
  var data = getRawDataByYears(currentYear - 1, currentYear);
  var targetRuangan = normText(ruangan);
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var groups = {};

  data.forEach(function (row) {
    if (normText(row[2]) !== targetRuangan) return;
    var catatan = row[13] ? String(row[13]).trim() : '';
    if (!catatan) return;
    var status = row[8] ? String(row[8]).trim() : '';
    if (status !== '❌ Kembalikan') return;

    var tgl = new Date(row[1]);
    var bulan = tgl.getMonth() + 1;
    var tahun = tgl.getFullYear();
    var indikator = row[3] ? String(row[3]).trim() : '';
    var key = indikator + '_' + bulan + '_' + tahun;

    if (!groups[key]) {
      groups[key] = {
        indikator: indikator,
        periode: months[bulan - 1] + ' ' + tahun,
        status: status,
        catatan: catatan,
        tglAwal: tgl,
        tglAkhir: tgl
      };
    } else {
      if (tgl < groups[key].tglAwal) groups[key].tglAwal = tgl;
      if (tgl > groups[key].tglAkhir) groups[key].tglAkhir = tgl;
    }
  });

  var tz = Session.getScriptTimeZone();
  var result = Object.keys(groups).map(function (k) {
    var g = groups[k];
    var fmtAwal = Utilities.formatDate(g.tglAwal, tz, 'dd/MM');
    var fmtAkhir = Utilities.formatDate(g.tglAkhir, tz, 'dd/MM');
    var tanggalLabel = fmtAwal === fmtAkhir ? fmtAwal : (fmtAwal + ' - ' + fmtAkhir);
    return {
      indikator: g.indikator,
      periode: g.periode,
      tanggal: tanggalLabel,
      status: g.status,
      catatan: g.catatan,
      urutan: g.tglAkhir.getTime()
    };
  });

  result.sort(function (a, b) { return b.urutan - a.urutan; });
  return result;
}

function saveValidation(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Master');
    const rowIndexes = payload.rowIndexes; // Array baris dari sebulan penuh

    // Hitung ulang Capaian Bulanan (Total Num/Den) langsung dari baris-baris
    // yang sedang divalidasi, supaya akurat walau ada pembulatan di client.
    let totalNum = 0, totalDen = 0;
    rowIndexes.forEach(function (rIndex) {
      const nd = sheet.getRange(rIndex, 5, 1, 2).getValues()[0];
      totalNum += Number(nd[0]) || 0;
      totalDen += Number(nd[1]) || 0;
    });
    const capaian = totalDen > 0 ? (totalNum / totalDen) * 100 : null;

    const sampel = Number(payload.sampel) || 0;
    const numSampel = Number(payload.numSampel) || 0;
    const temuan = sampel > 0 ? (numSampel / sampel) * 100 : null;

    let akurasi = '';
    if (capaian !== null && capaian > 0 && temuan !== null) {
      akurasi = Math.round((1 - Math.abs(capaian - temuan) / capaian) * 100 * 100) / 100;
    }

    // Update semua baris harian dalam 1 kali proses loop
    rowIndexes.forEach(function(rIndex) {
      sheet.getRange(rIndex, 9, 1, 9).setValues([[
        payload.status,
        payload.populasi || '',
        payload.sampel || '',
        payload.numSampel || '',
        payload.link || '',
        payload.catatan || '',
        akurasi,
        payload.analisa || '',
        payload.rtl || ''
      ]]);
    });

    return {
      success: true,
      message: "Validasi bulanan berhasil disimpan massal!",
      capaian: capaian !== null ? Math.round(capaian * 100) / 100 : null,
      temuan: temuan !== null ? Math.round(temuan * 100) / 100 : null,
      akurasi: akurasi
    };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan: " + e.message };
  }
}

/**
 * Mengambil grup data yang sudah divalidasi (Valid / Dikembalikan) milik
 * seorang validator (atau semua jika 'admin'). Dipakai bersama oleh
 * downloadValidationReport() (untuk export) dan getValidatedRecapForDisplay()
 * (untuk ditampilkan langsung di dashboard validator).
 */
function getValidatedGroups(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName('Data Master');
  var data = dataSheet.getDataRange().getValues();

  var configSheet = ss.getSheetByName('Config');
  var configData = configSheet.getDataRange().getValues();
  var headerConfig = configData[0];
  var idxInd = headerConfig.indexOf('Daftar Indikator');
  var idxVal = headerConfig.indexOf('Username Validator');

  var valMap = {};
  if (idxInd > -1 && idxVal > -1) {
    for (var i = 1; i < configData.length; i++) {
      var indikator = configData[i][idxInd];
      var penilai = configData[i][idxVal];
      if (indikator) {
        valMap[indikator.toString().trim()] = penilai ? penilai.toString().trim().toLowerCase() : "";
      }
    }
  }

  var usr = username ? username.toString().trim().toLowerCase() : "";
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  function getMonthYear(tglRaw) {
    if (tglRaw && tglRaw instanceof Date) return months[tglRaw.getMonth()] + " " + tglRaw.getFullYear();
    return tglRaw ? tglRaw.toString() : "";
  }

  // Pass 1: total SEMUA baris (berapapun statusnya) per ruangan+indikator+periode.
  // Ini dipakai untuk mendeteksi kalau ruangan mengedit data SETELAH divalidasi
  // (sehingga total di rekap validator jadi beda dengan total di rekap ruangan).
  var fullTotals = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var indName = row[3] ? row[3].toString().trim() : "";
    var hakVal = valMap[indName] || "";
    if (!(hakVal === usr || usr === 'admin')) continue;

    var ruangan = row[2] ? row[2].toString().trim() : "";
    var key = getMonthYear(row[1]) + "_" + ruangan + "_" + indName;
    if (!fullTotals[key]) fullTotals[key] = { num: 0, den: 0 };
    fullTotals[key].num += (Number(row[4]) || 0);
    fullTotals[key].den += (Number(row[5]) || 0);
  }

  // Pass 2: hanya baris yang statusnya sudah Valid / Dikembalikan.
  var validatedGroups = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var indName = row[3] ? row[3].toString().trim() : "";
    var hakVal = valMap[indName] || "";
    var statusStr = row[8] ? row[8].toString().trim() : "";

    if ((hakVal === usr || usr === 'admin') && (statusStr === '✅ Valid' || statusStr === '❌ Kembalikan')) {
      var monthYear = getMonthYear(row[1]);
      var ruangan = row[2] ? row[2].toString().trim() : "";
      var key = monthYear + "_" + ruangan + "_" + indName;

      if (!validatedGroups[key]) {
        validatedGroups[key] = {
          periode: monthYear,
          ruangan: ruangan,
          indikator: indName,
          num: 0, den: 0,
          status: statusStr,
          populasi: row[9] || '-',
          sampel: row[10] || '-',
          numSampel: row[11] || '-',
          link: row[12] || '-',
          catatan: row[13] || '-',
          akurasi: row[14] || '',
          analisa: row[15] || '-',
          rtl: row[16] || '-'
        };
      }
      validatedGroups[key].num += (Number(row[4]) || 0);
      validatedGroups[key].den += (Number(row[5]) || 0);
    }
  }

  Object.keys(validatedGroups).forEach(function (key) {
    var g = validatedGroups[key];
    var full = fullTotals[key] || { num: g.num, den: g.den };
    g.numSaatIni = full.num;
    g.denSaatIni = full.den;
    g.perluValidasiUlang = (full.num !== g.num) || (full.den !== g.den);
  });

  return Object.values(validatedGroups);
}

/**
 * Rekap hasil validasi untuk ditampilkan langsung di dashboard validator
 * (bukan diunduh). Diurutkan dari periode terbaru.
 */
function getValidatedRecapForDisplay(username) {
  var groups = getValidatedGroups(username);
  groups.sort(function (a, b) {
    if (a.periode === b.periode) return a.ruangan.localeCompare(b.ruangan);
    return a.periode < b.periode ? 1 : -1;
  });
  return groups;
}

/**
 * Hasil validasi (lengkap dengan Analisa & RTL dari validator) untuk
 * ditampilkan ke PIC ruangan, supaya ruangan bisa lihat feedback validator.
 * Bisa difilter per bulan & tahun biar gampang dicari (kalau dikosongkan,
 * tampil semua periode tahun ini & tahun lalu).
 */
function getHasilValidasiUntukRuangan(ruangan, bulan, tahun) {
  var currentYear = new Date().getFullYear();
  var startYear = tahun ? Number(tahun) : currentYear - 1;
  var endYear = tahun ? Number(tahun) : currentYear;
  var data = getRawDataByYears(startYear, endYear);
  var targetRuangan = (!ruangan || ruangan === 'SEMUA') ? null : normText(ruangan);
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var groups = {};

  data.forEach(function (row) {
    if (targetRuangan && normText(row[2]) !== targetRuangan) return;
    var statusStr = row[8] ? row[8].toString().trim() : '';
    if (statusStr !== '✅ Valid' && statusStr !== '❌ Kembalikan') return;

    var tgl = new Date(row[1]);
    var bln = tgl.getMonth() + 1, thn = tgl.getFullYear();
    if (bulan && bln !== Number(bulan)) return;
    if (tahun && thn !== Number(tahun)) return;

    var ruanganNama = row[2] ? row[2].toString().trim() : '';
    var indikator = row[3] ? row[3].toString().trim() : '';
    var key = ruanganNama + '_' + indikator + '_' + bln + '_' + thn;

    if (!groups[key]) {
      groups[key] = {
        ruangan: ruanganNama,
        indikator: indikator,
        periode: months[bln - 1] + ' ' + thn,
        num: 0, den: 0,
        status: statusStr,
        populasi: row[9] || '-',
        sampel: row[10] || '-',
        numSampel: row[11] || '-',
        catatan: row[13] || '-',
        akurasi: row[14] || '',
        analisa: row[15] || '-',
        rtl: row[16] || '-',
        tglAwal: tgl,
        tglAkhir: tgl
      };
    } else {
      if (tgl < groups[key].tglAwal) groups[key].tglAwal = tgl;
      if (tgl > groups[key].tglAkhir) groups[key].tglAkhir = tgl;
    }
    groups[key].num += (Number(row[4]) || 0);
    groups[key].den += (Number(row[5]) || 0);
  });

  var tz = Session.getScriptTimeZone();
  var result = Object.keys(groups).map(function (k) {
    var g = groups[k];
    var fmtAwal = Utilities.formatDate(g.tglAwal, tz, 'dd/MM');
    var fmtAkhir = Utilities.formatDate(g.tglAkhir, tz, 'dd/MM');
    g.tanggal = fmtAwal === fmtAkhir ? fmtAwal : (fmtAwal + ' - ' + fmtAkhir);
    g.urutan = g.tglAkhir.getTime();
    delete g.tglAwal;
    delete g.tglAkhir;
    return g;
  });

  result.sort(function (a, b) { return b.urutan - a.urutan; });
  return result;
}

/**
 * Total Numerator & Denominator (raw, bukan persen) per indikator untuk
 * bulan & tahun tertentu, dibatasi hanya pada indikator yang menjadi
 * tanggung jawab validator ybs (atau semua indikator kalau usernya 'admin').
 */
function getValidatorTotalNumDen(username, bulan, tahun) {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  var configData = configSheet.getDataRange().getValues();
  var headerConfig = configData[0];
  var idxInd = headerConfig.indexOf('Daftar Indikator');
  var idxVal = headerConfig.indexOf('Username Validator');

  var valMap = {};
  var indicatorOrder = [];
  if (idxInd > -1) {
    for (var i = 1; i < configData.length; i++) {
      var indikator = configData[i][idxInd];
      if (!indikator) continue;
      indikator = indikator.toString().trim();
      indicatorOrder.push(indikator);
      var penilai = idxVal > -1 ? configData[i][idxVal] : "";
      valMap[indikator] = penilai ? penilai.toString().trim().toLowerCase() : "";
    }
  }

  var usr = username ? username.toString().trim().toLowerCase() : "";
  var myIndicators = indicatorOrder.filter(function (ind) {
    return usr === 'admin' || valMap[ind] === usr;
  });

  var data = getRawDataByYears(Number(tahun), Number(tahun));
  var totals = {};

  data.forEach(function (row) {
    var tgl = new Date(row[1]);
    if ((tgl.getMonth() + 1) !== Number(bulan) || tgl.getFullYear() !== Number(tahun)) return;
    var ind = row[3] ? row[3].toString().trim() : "";
    if (myIndicators.indexOf(ind) === -1) return;
    if (!totals[ind]) totals[ind] = { numerator: 0, denominator: 0 };
    totals[ind].numerator += Number(row[4]) || 0;
    totals[ind].denominator += Number(row[5]) || 0;
  });

  return myIndicators.map(function (ind) {
    var t = totals[ind] || { numerator: 0, denominator: 0 };
    return { indikator: ind, totalNum: t.numerator, totalDen: t.denominator };
  });
}

function downloadValidationReport(username, format) {
  var rows = getValidatedGroups(username).map(function (vg) {
    return [vg.periode, vg.ruangan, vg.indikator, vg.num, vg.den, vg.status, vg.populasi, vg.sampel, vg.numSampel, vg.akurasi, (vg.perluValidasiUlang ? 'Perlu Validasi Ulang' : 'Sinkron'), vg.analisa, vg.rtl, vg.link, vg.catatan];
  });

  var usr = username ? username.toString().trim().toLowerCase() : "";
  var title = 'Riwayat_Validasi_Bulanan_' + (usr || 'Semua');
  var headers = ['Periode', 'Ruangan', 'Indikator', 'Total Num', 'Total Den', 'Status', 'Populasi', 'Sampel', 'Num Sampel', 'Akurasi (%)', 'Status Data', 'Analisa', 'RTL', 'Link Bukti', 'Catatan'];

  if (format === 'pdf') {
    return exportRecapPdf(title, headers, rows);
  } else {
    return exportRecapExcel(title, headers, rows);
  }
}
