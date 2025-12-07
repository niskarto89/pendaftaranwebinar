// =====================================================
// Google Apps Script untuk Registrasi Webinar
// =====================================================
// PETUNJUK DEPLOYMENT:
// 1. Buka script.google.com -> Buat project baru
// 2. Paste seluruh kode ini
// 3. Klik "Deploy" -> "New deployment"
// 4. Pilih "Web app"
// 5. Execute as: "Me" 
// 6. Who has access: "Anyone" (PENTING!)
// 7. Klik "Deploy" dan copy URL-nya
// =====================================================

const SPREADSHEET_ID = '1uzlCF1sO0ozjpkHhyA2JhutD-FNeI8VbgAQiBMERucU';
const SHEET_NAME = 'Sheet1';
const UPLOAD_FOLDER_ID = '1b8chkM6Nj08IQwltXOQQ4A-p8uj0lwUl'; // Folder untuk bukti pembayaran

// =====================================================
// GET Request Handler - Untuk test, JSONP, checkNik, dan getLatestFile
// =====================================================
function doGet(e) {
  // Handle jika e undefined (dijalankan manual dari editor)
  e = e || { parameter: {} };
  
  var callback = (e.parameter && e.parameter.callback) ? e.parameter.callback : '';
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'test';
  var filename = (e.parameter && e.parameter.filename) ? e.parameter.filename : '';
  var nik = (e.parameter && e.parameter.nik) ? e.parameter.nik : '';
  
  var result;
  
  // Action: checkNik - cek apakah NIK sudah terdaftar
  if (action === 'checkNik' && nik) {
    result = checkNikExists(nik);
  }
  // Action: getLatestFile - cari file berdasarkan nama di folder upload
  else if (action === 'getLatestFile' && filename) {
    result = getFileByName(filename);
  } 
  // Action: getRecentFiles - ambil 5 file terbaru
  else if (action === 'getRecentFiles') {
    result = getRecentFiles(5);
  }
  else {
    result = { success: true, message: 'API is working!', timestamp: new Date().toISOString() };
  }
  
  // Jika ada callback (JSONP), return sebagai JavaScript
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// Check if NIK exists in spreadsheet
// =====================================================
function checkNikExists(nik) {
  try {
    if (!nik || nik.length !== 16) {
      return { success: false, error: 'NIK tidak valid (harus 16 digit)' };
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      // Coba sheet pertama
      sheet = ss.getSheets()[0];
    }
    
    var data = sheet.getDataRange().getValues();
    var nikColumnIndex = -1;
    
    // Cari kolom NIK di header (baris pertama)
    if (data.length > 0) {
      for (var col = 0; col < data[0].length; col++) {
        var header = String(data[0][col]).toLowerCase().trim();
        if (header === 'nik' || header === 'nik (ktp)' || header.indexOf('nik') >= 0) {
          nikColumnIndex = col;
          break;
        }
      }
    }
    
    // Jika tidak ditemukan kolom NIK, asumsikan kolom ke-3 (index 2) adalah NIK
    // berdasarkan urutan: Timestamp, Nama, Tgl Lahir, NIK, ...
    if (nikColumnIndex === -1) {
      nikColumnIndex = 3; // Index 3 = kolom D (NIK)
    }
    
    // Cek apakah NIK sudah ada
    for (var row = 1; row < data.length; row++) { // Skip header
      var cellValue = String(data[row][nikColumnIndex]).trim();
      if (cellValue === nik) {
        return { 
          success: true, 
          exists: true, 
          message: 'NIK sudah terdaftar' 
        };
      }
    }
    
    return { 
      success: true, 
      exists: false, 
      message: 'NIK belum terdaftar' 
    };
    
  } catch (err) {
    Logger.log('checkNikExists error: ' + err.message);
    return { success: false, error: err.message };
  }
}

// =====================================================
// Get file by name dari folder upload
// =====================================================
function getFileByName(filename) {
  try {
    if (!UPLOAD_FOLDER_ID) {
      return { success: false, error: 'Upload folder not configured' };
    }
    
    var folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    var files = folder.getFilesByName(filename);
    
    if (files.hasNext()) {
      var file = files.next();
      var fileId = file.getId();
      return {
        success: true,
        id: fileId,
        name: file.getName(),
        url: 'https://drive.google.com/file/d/' + fileId + '/view',
        directLink: 'https://drive.google.com/uc?id=' + fileId
      };
    }
    
    return { success: false, error: 'File not found: ' + filename };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =====================================================
// Get recent files dari folder upload
// =====================================================
function getRecentFiles(limit) {
  try {
    if (!UPLOAD_FOLDER_ID) {
      return { success: false, error: 'Upload folder not configured' };
    }
    
    var folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    var files = folder.getFiles();
    var fileList = [];
    
    while (files.hasNext() && fileList.length < limit) {
      var file = files.next();
      var fileId = file.getId();
      fileList.push({
        id: fileId,
        name: file.getName(),
        url: 'https://drive.google.com/file/d/' + fileId + '/view',
        created: file.getDateCreated().toISOString()
      });
    }
    
    // Sort by created date descending
    fileList.sort(function(a, b) {
      return new Date(b.created) - new Date(a.created);
    });
    
    return { success: true, files: fileList };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =====================================================
// POST Request Handler - Untuk simpan data
// =====================================================
function doPost(e) {
  try {
    // Log untuk debugging
    Logger.log('=== doPost called ===');
    Logger.log('e: ' + JSON.stringify(e));
    Logger.log('e.parameter: ' + JSON.stringify(e ? e.parameter : 'null'));
    Logger.log('e.postData: ' + JSON.stringify(e && e.postData ? e.postData.contents : 'null'));
    
    // Handle jika e undefined
    e = e || { parameter: {}, postData: null };
    
    var body = parseRequest(e);
    Logger.log('Parsed body: ' + JSON.stringify(body));
    
    // Test endpoint
    if (body.test) {
      return htmlResponseWithPostMessage({ success: true, test: true, timestamp: new Date().toISOString() });
    }
    
    // Upload file ke Drive
    if (body.action === 'upload') {
      return handleUpload(body);
    }
    
    // Simpan ke sheet
    return saveToSheet(body);
    
  } catch (err) {
    Logger.log('doPost ERROR: ' + err.message);
    return htmlResponseWithPostMessage({ success: false, error: err.message });
  }
}

// =====================================================
// Parse berbagai jenis request
// =====================================================
function parseRequest(e) {
  var body = {};
  
  // Handle jika e undefined atau tidak ada postData
  if (!e) {
    return body;
  }
  
  // Prioritas 1: parameter (dari form-data atau URL params)
  if (e.parameter && Object.keys(e.parameter).length > 0) {
    body = e.parameter;
    Logger.log('Parsed from e.parameter: ' + JSON.stringify(body));
    return body;
  }
  
  // Prioritas 2: postData (dari JSON body)
  if (e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
      Logger.log('Parsed from JSON postData: ' + JSON.stringify(body));
    } catch (parseErr) {
      Logger.log('postData is not JSON, using parameter instead');
      body = e.parameter || {};
    }
  }
  
  return body;
}

// =====================================================
// Simpan data ke Google Sheet
// =====================================================
function saveToSheet(body) {
  // Validasi NIK
  if (!body.nik) {
    return jsonResponse({ success: false, error: 'NIK wajib diisi' });
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // Buat sheet jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // Setup headers jika sheet kosong
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    var headers = ['Timestamp', 'Nama', 'TglLahir', 'NIK', 'Profesi', 'Email', 'SatuSehat', 'Alamat', 'Provinsi', 'FileName', 'FileUrl'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    lastRow = 1;
  }
  
  // Cek duplikat NIK
  var nikCol = 4; // Kolom D
  if (lastRow > 1) {
    var nikValues = sheet.getRange(2, nikCol, lastRow - 1, 1).getValues();
    var incomingNik = String(body.nik).trim();
    
    for (var i = 0; i < nikValues.length; i++) {
      if (String(nikValues[i][0]).trim() === incomingNik) {
        return jsonResponse({ 
          success: false, 
          duplicate: true, 
          message: 'NIK sudah terdaftar sebelumnya' 
        });
      }
    }
  }
  
  // Tambah baris baru
  var newRow = [
    new Date(),
    body.nama || '',
    body.tglLahir || '',
    body.nik || '',
    body.profesi || '',
    body.email || '',
    body.satuSehat || '',
    body.alamat || '',
    body.provinsi || '',
    body.fileName || '',
    body.fileUrl || ''
  ];
  
  sheet.appendRow(newRow);
  
  // Set FileUrl sebagai hyperlink yang bisa diklik
  var newRowNum = sheet.getLastRow();
  var fileUrlCol = 11; // Kolom K (FileUrl)
  
  if (body.fileUrl && body.fileUrl.startsWith('http')) {
    var cell = sheet.getRange(newRowNum, fileUrlCol);
    var fileName = body.fileName || 'Lihat File';
    // Buat formula hyperlink: =HYPERLINK("url", "text")
    cell.setFormula('=HYPERLINK("' + body.fileUrl + '", "' + fileName + '")');
    cell.setFontColor('#1155CC'); // Warna biru hyperlink
    cell.setFontUnderline(true); // Underline
  }
  
  return jsonResponse({ 
    success: true, 
    message: 'Data berhasil disimpan!',
    row: sheet.getLastRow()
  });
}

// =====================================================
// Upload file ke Google Drive
// =====================================================
function handleUpload(body) {
  try {
    var filename = body.filename || 'bukti_' + new Date().getTime();
    var base64 = body.base64 || body.b64 || '';
    var mimeType = body.mimeType || 'application/octet-stream';
    
    Logger.log('handleUpload - filename: ' + filename);
    
    // Extract base64 dari data URL jika ada
    var match = base64.match(/^data:(.+);base64,(.*)$/);
    if (match) {
      mimeType = match[1];
      base64 = match[2];
    }
    
    if (!base64) {
      return simpleHtmlResponse('Error: Tidak ada data file', false);
    }
    
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    
    var file;
    if (UPLOAD_FOLDER_ID) {
      var folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
      file = folder.createFile(blob);
    } else {
      file = DriveApp.createFile(blob);
    }
    
    // Set file bisa diakses siapa saja dengan link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // URL untuk view file di browser (bukan download)
    var fileId = file.getId();
    var url = 'https://drive.google.com/file/d/' + fileId + '/view';
    
    Logger.log('Upload success! URL: ' + url);
    
    // Return simple HTML response (file URL akan di-query via JSONP)
    return simpleHtmlResponse('Upload berhasil! File: ' + filename, true);
    
  } catch (err) {
    Logger.log('Upload error: ' + err.message);
    return simpleHtmlResponse('Error: ' + err.message, false);
  }
}

// =====================================================
// Helper: Buat JSON response
// =====================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// Helper: Simple HTML response (untuk iframe)
// =====================================================
function simpleHtmlResponse(message, success) {
  var color = success ? '#28a745' : '#dc3545';
  var icon = success ? '✅' : '❌';
  
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;padding:20px;text-align:center;">' +
    '<p style="color:' + color + ';font-size:14px;">' + icon + ' ' + message + '</p>' +
    '</body></html>';
  
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================
// Test Function - Jalankan di Apps Script Editor
// =====================================================
function testSaveToSheet() {
  Logger.log('=== TEST SAVE TO SHEET ===');
  Logger.log('SPREADSHEET_ID: ' + SPREADSHEET_ID);
  Logger.log('SHEET_NAME: ' + SHEET_NAME);
  
  var testData = {
    nama: 'Test User',
    tglLahir: '1990-01-01',
    nik: 'TEST' + new Date().getTime(),
    profesi: 'Ners',
    email: 'test@example.com',
    satuSehat: 'Ya',
    alamat: 'Jl. Test No. 123',
    provinsi: 'Sumatera Utara'
  };
  
  Logger.log('Test Data: ' + JSON.stringify(testData));
  
  try {
    // Coba buka spreadsheet langsung
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Spreadsheet opened: ' + ss.getName());
    
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log('Sheet not found, creating new sheet...');
      sheet = ss.insertSheet(SHEET_NAME);
    }
    Logger.log('Sheet: ' + sheet.getName());
    Logger.log('Last row before: ' + sheet.getLastRow());
    
    var result = saveToSheet(testData);
    var resultContent = result.getContent();
    Logger.log('Result: ' + resultContent);
    
    // Cek lagi last row
    Logger.log('Last row after: ' + sheet.getLastRow());
    
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    Logger.log('Stack: ' + err.stack);
  }
}

// Test langsung append tanpa fungsi saveToSheet
function testDirectAppend() {
  Logger.log('=== TEST DIRECT APPEND ===');
  
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Spreadsheet: ' + ss.getName());
    
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    
    // Append langsung
    var testRow = [new Date(), 'Direct Test', '2000-01-01', 'DIRECT123', 'Test', 'direct@test.com', 'Ya', 'Alamat Test', 'Test Province', '', ''];
    sheet.appendRow(testRow);
    
    Logger.log('Direct append success! Last row: ' + sheet.getLastRow());
    
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
  }
}
