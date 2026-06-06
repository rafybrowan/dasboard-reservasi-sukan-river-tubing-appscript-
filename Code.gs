// ============================================
// Sukan River Tubing - Dashboard Reservasi
// Google Apps Script Backend
// ============================================

var SPREADSHEET_ID = '1yiy5nlMCen9Y8qxD9M-5Zc5Ea5Xghk95zuaDX1mQYgk';
var SHEET_NAME = 'Sheet1';

/**
 * Serve the HTML dashboard
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sukan River Tubing - Dashboard Reservasi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Get the spreadsheet and sheet reference
 */
function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME);
}

/**
 * Helper: find column index by name (case-insensitive, trimmed)
 */
function normalizeColName(name) {
  return String(name).trim().toUpperCase().replace(/[\s_\-]+/g, '');
}

function findCol(headers, name) {
  var target = normalizeColName(name);
  // Pass 1: exact match (trimmed, uppercased)
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toUpperCase() === name.toUpperCase()) {
      return i;
    }
  }
  // Pass 2: normalized match (ignore spaces, underscores, hyphens)
  for (var i = 0; i < headers.length; i++) {
    if (normalizeColName(headers[i]) === target) {
      return i;
    }
  }
  return -1;
}

/**
 * Debug: return actual column headers from the sheet
 */
function getHeaders() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  return data[0].map(function(h) { return String(h).trim(); });
}

/**
 * Helper: safely get cell value
 */
function getVal(row, idx, fallback) {
  if (idx < 0 || idx >= row.length) return fallback;
  var v = row[idx];
  if (v === null || v === undefined || v === '') return fallback;
  // Handle Date objects from Sheets
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return v;
}

/**
 * Fetch all reservation data from the spreadsheet
 */
function getReservationData() {
  try {
    var sheet = getSheet();
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan.');

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var rows = data.slice(1);

    // Find column indices (case-insensitive)
    var colPengirim   = findCol(headers, 'PENGIRIM');
    var colPenerima   = findCol(headers, 'PENERIMA');
    var colNominal    = findCol(headers, 'NOMINAL');
    var colMetode     = findCol(headers, 'METODE');
    var colStatus     = findCol(headers, 'STATUS');
    var colTanggal    = findCol(headers, 'TANGGAL');
    var colWaktu      = findCol(headers, 'WAKTU');
    var colRefId      = findCol(headers, 'REF_ID');
    var colCatatan    = findCol(headers, 'CATATAN');
    var colKode       = findCol(headers, 'KODE RESERVASI');
    var colCheckin    = findCol(headers, 'STATUS_CHECKIN');

    var reservations = [];
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      // Skip completely empty rows
      var kode = getVal(row, colKode, '');
      if (kode === '' && getVal(row, colPengirim, '') === '') continue;

      reservations.push({
        rowIndex: i + 2,
        pengirim: String(getVal(row, colPengirim, '-')),
        penerima: String(getVal(row, colPenerima, '-')),
        nominal: getVal(row, colNominal, 0),
        metode: String(getVal(row, colMetode, '-')),
        status: String(getVal(row, colStatus, '-')),
        tanggal: String(getVal(row, colTanggal, '-')),
        waktu: String(getVal(row, colWaktu, '-')),
        refId: String(getVal(row, colRefId, '-')),
        catatan: String(getVal(row, colCatatan, '-')),
        kodeReservasi: String(kode),
        statusCheckin: String(getVal(row, colCheckin, 'Belum Check-In'))
      });
    }

    return reservations;
  } catch (e) {
    throw new Error('Gagal membaca data: ' + e.message);
  }
}

/**
 * Get dashboard summary statistics
 */
function getDashboardStats() {
  try {
    var reservations = getReservationData();

    var totalReservasi = reservations.length;
    var totalPendapatan = 0;
    var sudahCheckin = 0;

    for (var i = 0; i < reservations.length; i++) {
      var r = reservations[i];
      var nominal = r.nominal;
      if (typeof nominal === 'string') {
        nominal = parseFloat(nominal.replace(/[^0-9.\-]/g, ''));
      }
      if (!isNaN(nominal)) totalPendapatan += nominal;

      if (String(r.statusCheckin).toLowerCase().indexOf('sudah') !== -1) {
        sudahCheckin++;
      }
    }

    var belumCheckin = totalReservasi - sudahCheckin;

    return {
      totalReservasi: totalReservasi,
      totalPendapatan: totalPendapatan,
      sudahCheckin: sudahCheckin,
      belumCheckin: belumCheckin,
      reservations: reservations
    };
  } catch (e) {
    throw new Error('Gagal memuat statistik: ' + e.message);
  }
}

/**
 * Search reservation by KODE RESERVASI
 */
function searchReservation(kodeReservasi) {
  try {
    var reservations = getReservationData();
    var search = String(kodeReservasi).trim().toLowerCase();

    for (var i = 0; i < reservations.length; i++) {
      if (String(reservations[i].kodeReservasi).trim().toLowerCase() === search) {
        return reservations[i];
      }
    }
    return null;
  } catch (e) {
    throw new Error('Gagal mencari reservasi: ' + e.message);
  }
}

/**
 * Confirm check-in: update STATUS_CHECKIN to "Sudah Check-In"
 */
function confirmCheckin(kodeReservasi) {
  try {
    var sheet = getSheet();
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var kodeCol = findCol(headers, 'KODE RESERVASI');
    var checkinCol = findCol(headers, 'STATUS_CHECKIN');

    // Otomatis buat kolom STATUS_CHECKIN jika belum ada
    if (checkinCol === -1) {
      var newColIndex = headers.length + 1;
      sheet.getRange(1, newColIndex).setValue('STATUS_CHECKIN');
      // Set semua baris yang ada ke "Belum Check-In"
      if (data.length > 1) {
        for (var j = 2; j <= data.length; j++) {
          sheet.getRange(j, newColIndex).setValue('Belum Check-In');
        }
      }
      checkinCol = newColIndex - 1;
    }

    if (kodeCol === -1) {
      return { success: false, message: 'Kolom KODE RESERVASI tidak ditemukan di sheet.' };
    }

    var search = String(kodeReservasi).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][kodeCol]).trim().toLowerCase() === search) {
        sheet.getRange(i + 1, checkinCol + 1).setValue('Sudah Check-In');
        return { success: true, message: 'Check-in berhasil dikonfirmasi!' };
      }
    }

    return { success: false, message: 'Kode reservasi tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}
