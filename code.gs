/**
 * Google Apps Script Backend for Doorprize Registration
 * 
 * SETUP INSTRUCTIONS:
 * 1. Buka Google Sheets, buat spreadsheet baru
 * 2. Beri nama sheet pertama: "Registrasi"
 * 3. Tambahkan header di baris 1: No | Nama | No HP | Nomor Undian | Waktu Daftar | ID Perangkat
 * 4. Buka menu Extensions > Apps Script
 * 5. Paste kode ini ke editor Apps Script
 * 6. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copy URL deployment, paste ke file index.html (variabel APPS_SCRIPT_URL)
 */

// Nama sheet yang digunakan
const SHEET_NAME = "Registrasi";

// Prefix nomor undian (bisa diganti sesuai event)
const UNDIAN_PREFIX = "SE26-";

// Panjang digit nomor undian
const UNDIAN_DIGITS = 4;

/**
 * Handle GET requests - cek apakah Perangkat sudah terdaftar
 */
function doGet(e) {
  try {
    const action = e.parameter.action || "";
    const deviceId = e.parameter.ip || e.parameter.deviceId || "";

    if ((action === "checkIp" || action === "checkDevice") && deviceId) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet) {
        return createResponse(false, "Sheet tidak ditemukan", null);
      }

      const existingData = sheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        if (String(existingData[i][5]).trim() === deviceId.trim()) {
          return createResponse(true, "Perangkat sudah terdaftar", {
            nomorUndian: existingData[i][3],
            nama: existingData[i][1]
          });
        }
      }

      return createResponse(false, "Perangkat belum terdaftar", null);
    }

    return createResponse(true, "Doorprize API is running", null);
  } catch (error) {
    return createResponse(false, "Error: " + error.message, null);
  }
}

/**
 * Handle POST requests - registrasi peserta
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const nama = data.nama ? data.nama.trim() : "";
    const noHp = data.noHp ? data.noHp.trim() : "";
    const deviceId = (data.ip || data.deviceId || "").trim();

    // Validasi input
    if (!nama || !noHp) {
      return createResponse(false, "Nama dan No HP wajib diisi", null);
    }

    // Validasi format nomor HP
    if (!/^(\+62|62|08)[0-9]{8,13}$/.test(noHp.replace(/[\s\-]/g, ""))) {
      return createResponse(false, "Format nomor HP tidak valid", null);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return createResponse(false, "Sheet tidak ditemukan", null);
    }

    // Cek apakah nomor HP atau Perangkat sudah terdaftar
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      // Cek duplikat nomor HP
      const existingHp = String(existingData[i][2]).replace(/[\s\-]/g, "");
      const inputHp = noHp.replace(/[\s\-]/g, "");
      if (existingHp === inputHp) {
        return createResponse(false, "Nomor HP sudah terdaftar", {
          nomorUndian: existingData[i][3],
          nama: existingData[i][1]
        });
      }

      // Cek duplikat Device ID
      if (deviceId && String(existingData[i][5]).trim() === deviceId) {
        return createResponse(false, "Perangkat ini sudah terdaftar", {
          nomorUndian: existingData[i][3],
          nama: existingData[i][1]
        });
      }
    }

    // Generate nomor undian
    const nomorUrut = existingData.length; // karena baris 1 header
    const nomorUndian = UNDIAN_PREFIX + String(nomorUrut).padStart(UNDIAN_DIGITS, "0");

    // Timestamp
    const waktu = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    // Simpan ke spreadsheet (termasuk ID Perangkat di kolom 6)
    sheet.appendRow([nomorUrut, nama, noHp, nomorUndian, waktu, deviceId]);

    return createResponse(true, "Registrasi berhasil!", {
      nomorUndian: nomorUndian,
      nama: nama
    });

  } catch (error) {
    return createResponse(false, "Terjadi kesalahan: " + error.message, null);
  }
}

/**
 * Helper function untuk membuat response JSON
 */
function createResponse(success, message, data) {
  const response = {
    success: success,
    message: message,
    data: data
  };

  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fungsi untuk setup awal sheet (jalankan manual sekali)
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Set header (termasuk kolom ID Perangkat)
  const headers = ["No", "Nama", "No HP", "Nomor Undian", "Waktu Daftar", "ID Perangkat"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4a90d9");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");

  // Set column widths
  sheet.setColumnWidth(1, 60);   // No
  sheet.setColumnWidth(2, 200);  // Nama
  sheet.setColumnWidth(3, 160);  // No HP
  sheet.setColumnWidth(4, 140);  // Nomor Undian
  sheet.setColumnWidth(5, 180);  // Waktu Daftar
  sheet.setColumnWidth(6, 160);  // ID Perangkat

  // Freeze header row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("Setup selesai! Sheet 'Registrasi' sudah siap.");
}
