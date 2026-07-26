// =========================================================================
// GOOGLE APPS SCRIPT FOR GKMI YESUS RAJA - JADWAL PELAYANAN MUSIK GEREJA
// =========================================================================
// Cara Menggunakan:
// 1. Buka Google Spreadsheet di Google Drive / sheets.google.com
// 2. Klik menu Ekstensi > Apps Script
// 3. Paste seluruh isi file ini ke dalam editor Apps Script
// 4. Klik Simpan (Ctrl + S)
// 5. Klik Deploy > New deployment > pilih Web app
// 6. Set "Execute as" = Me (Saya) dan "Who has access" = Anyone (Siapa saja)
// 7. Klik Deploy dan Copy Web App URL ke aplikasi website (tombol hijau 📊)
// =========================================================================

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JadwalMusik");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var val = row[j];
      if (key === 'daftarLagu' && typeof val === 'string' && val.startsWith('[')) {
        try { val = JSON.parse(val); } catch(err) {}
      }
      obj[key] = val;
    }
    result.push(obj);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JadwalMusik");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("JadwalMusik");
    sheet.appendRow(["id", "tanggalWaktu", "sesiIbadah", "worshipLeader", "singers", "keyboardist", "guitarist", "bassist", "drummer", "soundman", "daftarLagu"]);
  }
  
  var contents = JSON.parse(e.postData.contents);
  var action = contents.action;
  
  if (action === 'save') {
    var payload = contents.data;
    var data = sheet.getDataRange().getValues();
    var foundIndex = -1;
    
    if (typeof payload.daftarLagu === 'object') {
      payload.daftarLagu = JSON.stringify(payload.daftarLagu);
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        foundIndex = i + 1;
        break;
      }
    }
    
    var rowData = [
      payload.id || '',
      payload.tanggalWaktu || '',
      payload.sesiIbadah || '',
      payload.worshipLeader || '',
      payload.singers || '',
      payload.keyboardist || '',
      payload.guitarist || '',
      payload.bassist || '',
      payload.drummer || '',
      payload.soundman || '',
      payload.daftarLagu || ''
    ];
    
    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'delete') {
    var idToDelete = contents.id;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idToDelete) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: 'error' })).setMimeType(ContentService.MimeType.JSON);
}
