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

const DATA_HEADERS = ['Timestamp', 'Tanggal', 'Ruangan', 'Indikator', 'Numerator', 'Denominator', 'Diisi Oleh', 'Keterangan', 'Status Validasi', 'Populasi', 'Sampel', 'Num Sampel', 'Link Bukti', 'Catatan Validator'];

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
  const data = sheet.getRange('B2:F' + lastRow).getValues();
  return data
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      const target = (row[3] === '' || row[3] === undefined) ? null : Number(row[3]);
      const reverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(row[0])) > -1;
      return { nama: row[0], num: row[1] || '', den: row[2] || '', target: target, tipe: row[4] || 'persen', reverse: reverse };
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
      '⏳ Menunggu', '', '', '', '', ''
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
    return { indikator: ind.nama, tipe: ind.tipe, values: values };
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
    bannerCenter.getFill().setSolidFill('#6fa8dc');
    bannerCenter.getBorder().setTransparent();
    bannerCenter.getText().setText(ind.nama);
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

function saveValidation(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Master');
    const rowIndexes = payload.rowIndexes; // Array baris dari sebulan penuh
    
    // Update semua baris harian dalam 1 kali proses loop
    rowIndexes.forEach(function(rIndex) {
      sheet.getRange(rIndex, 9, 1, 6).setValues([[
        payload.status,
        payload.populasi || '',
        payload.sampel || '',
        payload.numSampel || '',
        payload.link || '',
        payload.catatan || ''
      ]]);
    });
    
    return { success: true, message: "Validasi bulanan berhasil disimpan massal!" };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan: " + e.message };
  }
}

function downloadValidationReport(username, format) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName('Data Master');
  var data = dataSheet.getDataRange().getValues();
  
  var configSheet = ss.getSheetByName('Config');
  var configData = configSheet.getDataRange().getValues();
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
  var validatedGroups = {};
  var rows = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var indName = row[3] ? row[3].toString().trim() : "";
    var hakVal = valMap[indName] || "";
    var statusStr = row[8] ? row[8].toString().trim() : "";
    
    // Cek data yang sudah tervalidasi
    if ((hakVal === usr || usr === 'admin') && (statusStr === '✅ Valid' || statusStr === '❌ Kembalikan')) {
        var tglRaw = row[1];
        var monthYear = "";
        
        if (tglRaw && tglRaw instanceof Date) {
          var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
          monthYear = months[tglRaw.getMonth()] + " " + tglRaw.getFullYear();
        } else if (tglRaw) {
          monthYear = tglRaw.toString();
        }
        
        var ruangan = row[2] ? row[2].toString().trim() : "";
        var key = monthYear + "_" + ruangan + "_" + indName;

        // Grupkan laporannya jadi bulanan juga biar rapi
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
            catatan: row[13] || '-'
          };
        }
        validatedGroups[key].num += (Number(row[4]) || 0);
        validatedGroups[key].den += (Number(row[5]) || 0);
    }
  }
  
  for(var k in validatedGroups) {
     var vg = validatedGroups[k];
     rows.push([vg.periode, vg.ruangan, vg.indikator, vg.num, vg.den, vg.status, vg.populasi, vg.sampel, vg.numSampel, vg.link, vg.catatan]);
  }
  
  var title = 'Riwayat_Validasi_Bulanan_' + (usr || 'Semua');
  var headers = ['Periode', 'Ruangan', 'Indikator', 'Total Num', 'Total Den', 'Status', 'Populasi', 'Sampel', 'Num Sampel', 'Link Bukti', 'Catatan'];
  
  if (format === 'pdf') {
    return exportRecapPdf(title, headers, rows);
  } else {
    return exportRecapExcel(title, headers, rows);
  }
}
