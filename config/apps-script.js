// Google Apps Script for Swooper Logbook
// This script acts as a backend API for the Swooper Logbook web app

function doGet(e) {
  const action = e.parameter.action;
  
  try {
    let response;
    if (action === 'getJumps') {
      response = getJumps();
    } else if (action === 'getEquipment') {
      response = getEquipment();
    } else if (action === 'getBackupEquipment') {
      response = getBackupEquipment();
    } else {
      response = ContentService
        .createTextOutput(JSON.stringify({ error: 'Invalid action' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return response;
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    let response;
    if (action === 'addJump') {
      response = addJump(data.data);
    } else if (action === 'uploadJumps') {
      response = uploadJumps(data.data);
    } else if (action === 'saveEquipment') {
      response = saveEquipment(data.data);
    } else {
      response = ContentService
        .createTextOutput(JSON.stringify({ error: 'Invalid action' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return response;
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getJumps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jumps');
  const data = sheet.getDataRange().getValues();
  
  return ContentService
    .createTextOutput(JSON.stringify({ data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addJump(jumpData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jumps');
  
  // Add the jump data as a new row (col 7 = equipmentId preserves the rig ID)
  sheet.appendRow([
    jumpData.jumpNumber,
    jumpData.date,
    jumpData.location,
    jumpData.equipment,
    jumpData.notes,
    jumpData.timestamp,
    jumpData.equipmentId || '',
    jumpData.linesetNumber || 1
  ]);
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function uploadJumps(jumpsData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jumps');
  
  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  // Add all jumps
  jumpsData.forEach(jump => {
    sheet.appendRow([
      jump.jumpNumber,
      jump.date,
      jump.location,
      jump.equipment,
      jump.notes,
      jump.timestamp,
      jump.equipmentId || '',
      jump.linesetNumber || 1
    ]);
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, count: jumpsData.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Equipment tab ────────────────────────────────────────────────────────────
// The "Equipment" sheet stores all component / rig data as named rows:
//   Row 1:  key="harnesses"          value=<JSON array>
//   Row 2:  key="canopies"      value=<JSON array>
//   Row 3:  key="rigs"  value=<JSON array>
//   Row 4:  key="settings"      value=<JSON object>
//   Row 5:  key="locations"     value=<JSON array>
//   Row 6:  key="_syncMeta"     value=<JSON object>

function getOrCreateEquipmentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Equipment');
  if (!sheet) {
    sheet = ss.insertSheet('Equipment');
    sheet.getRange('A1').setValue('harnesses');
    sheet.getRange('A2').setValue('canopies');
    sheet.getRange('A3').setValue('rigs');
    sheet.getRange('A4').setValue('settings');
    sheet.getRange('A5').setValue('locations');
    sheet.getRange('A6').setValue('_syncMeta');
  } else {
    const row6Key = sheet.getRange('A6').getValue();
    if (!row6Key) {
      sheet.getRange('A6').setValue('_syncMeta');
    }
  }
  return sheet;
}

function getEquipment() {
  const sheet = getOrCreateEquipmentSheet();
  
  // Read each row: col A = key, col B = JSON value
  const rows = sheet.getRange(1, 1, 6, 2).getValues();
  const result = {};
  rows.forEach(row => {
    const key   = row[0];
    const value = row[1];
    if (key && value !== '') {
      try {
        result[key] = JSON.parse(value);
      } catch (e) {
        result[key] = null;
      }
    }
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getBackupEquipment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('backupRigs');
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, hasBackupRigsSheet: false, data: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Same layout as the Equipment sheet: col A = key, col B = JSON value
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const rows = sheet.getRange(1, 1, lastRow, 2).getValues();
  const result = {};
  rows.forEach(row => {
    const key   = row[0];
    const value = row[1];
    if (key && value !== '') {
      try {
        result[key] = JSON.parse(value);
      } catch (e) {
        result[key] = null;
      }
    }
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, hasBackupRigsSheet: true, data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveEquipment(equipmentData) {
  const sheet = getOrCreateEquipmentSheet();
  
  const keyOrder = ['harnesses', 'canopies', 'rigs', 'settings', 'locations', '_syncMeta'];
  keyOrder.forEach((key, index) => {
    sheet.getRange(index + 1, 1).setValue(key);
    if (equipmentData[key] !== undefined) {
      sheet.getRange(index + 1, 2).setValue(JSON.stringify(equipmentData[key]));
    }
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
