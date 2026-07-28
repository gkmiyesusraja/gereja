// =========================================================================
// GOOGLE APPS SCRIPT FOR GKMI YESUS RAJA - JADWAL PELAYANAN MUSIK GEREJA
// =========================================================================
// Mendukung 3 Tab Sheet:
// 1. "JadwalMusik" (id, tanggalWaktu, sesiIbadah, worshipLeader, singers, keyboardist, guitarist, bassist, drummer, soundman, daftarLagu, createdBy)
// 2. "Users" (username, password, nama, role)
// 3. "Jemaat" (id, nama, telepon, bidang, username)
// =========================================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getSchedules';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'getUsers') {
    var userSheet = ss.getSheetByName("Users");
    if (!userSheet) return jsonOutput([]);
    var data = userSheet.getDataRange().getValues();
    if (data.length <= 1) return jsonOutput([]);
    var headers = data[0];
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var u = {};
      for (var j = 0; j < headers.length; j++) u[headers[j]] = row[j];
      users.push(u);
    }
    return jsonOutput(users);
  }

  if (action === 'getJemaat') {
    var jemaatSheet = ss.getSheetByName("Jemaat");
    if (!jemaatSheet) return jsonOutput([]);
    var data = jemaatSheet.getDataRange().getValues();
    if (data.length <= 1) return jsonOutput([]);
    var headers = data[0];
    var jemaatList = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var jObj = {};
      for (var j = 0; j < headers.length; j++) jObj[headers[j]] = row[j];
      jemaatList.push(jObj);
    }
    return jsonOutput(jemaatList);
  }
  
  // Default: getSchedules
  var sheet = ss.getSheetByName("JadwalMusik");
  if (!sheet) return jsonOutput([]);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonOutput([]);
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
  return jsonOutput(result);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contents = JSON.parse(e.postData.contents);
  var action = contents.action;
  
  if (action === 'saveJemaat') {
    var jSheet = ss.getSheetByName("Jemaat");
    if (!jSheet) {
      jSheet = ss.insertSheet("Jemaat");
      jSheet.appendRow(["id", "nama", "telepon", "bidang", "username"]);
    }
    var jData = contents.data;
    var data = jSheet.getDataRange().getValues();
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === jData.id) {
        foundIndex = i + 1;
        break;
      }
    }
    var jRow = [jData.id || '', jData.nama || '', jData.telepon || '', jData.bidang || '', jData.username || ''];
    if (foundIndex > 0) jSheet.getRange(foundIndex, 1, 1, jRow.length).setValues([jRow]);
    else jSheet.appendRow(jRow);
    return jsonOutput({ status: 'success' });
  }

  if (action === 'deleteJemaat') {
    var jSheet = ss.getSheetByName("Jemaat");
    if (jSheet) {
      var idToDelete = contents.id;
      var data = jSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === idToDelete) {
          jSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    return jsonOutput({ status: 'success' });
  }

  if (action === 'saveUser') {
    var userSheet = ss.getSheetByName("Users");
    if (!userSheet) {
      userSheet = ss.insertSheet("Users");
      userSheet.appendRow(["username", "password", "nama", "role"]);
    }
    var uData = contents.data;
    var data = userSheet.getDataRange().getValues();
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === uData.username) {
        foundIndex = i + 1;
        break;
      }
    }
    var uRow = [uData.username, uData.password, uData.nama, uData.role];
    if (foundIndex > 0) userSheet.getRange(foundIndex, 1, 1, uRow.length).setValues([uRow]);
    else userSheet.appendRow(uRow);
    return jsonOutput({ status: 'success' });
  }

  if (action === 'deleteUser') {
    var userSheet = ss.getSheetByName("Users");
    if (userSheet) {
      var uName = contents.username;
      var data = userSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === uName) {
          userSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    return jsonOutput({ status: 'success' });
  }

  if (action === 'save') {
    var sheet = ss.getSheetByName("JadwalMusik");
    if (!sheet) {
      sheet = ss.insertSheet("JadwalMusik");
      sheet.appendRow(["id", "tanggalWaktu", "sesiIbadah", "worshipLeader", "singers", "keyboardist", "guitarist", "bassist", "drummer", "soundman", "gambarFlyer", "daftarLagu", "createdBy"]);
    }
    var payload = contents.data;
    var data = sheet.getDataRange().getValues();
    var foundIndex = -1;
    if (typeof payload.daftarLagu === 'object') payload.daftarLagu = JSON.stringify(payload.daftarLagu);
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
      payload.gambarFlyer || '',
      payload.daftarLagu || '',
      payload.createdBy || ''
    ];
    if (foundIndex > 0) sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
    else sheet.appendRow(rowData);
    return jsonOutput({ status: 'success' });
  }
  
  if (action === 'delete') {
    var sheet = ss.getSheetByName("JadwalMusik");
    if (sheet) {
      var idToDelete = contents.id;
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === idToDelete) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }
    return jsonOutput({ status: 'success' });
  }

  return jsonOutput({ status: 'error' });
}

function jsonOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
