function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data;
    
    // Support JSON payload or Form data
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }

    // New field mappings: Center_Code, Reg No, Coordinates
    var centerCode = data.center_code || data.Center_Code || '';
    var regNo = data.reg_no || data.Reg_No || data['Reg No'] || '';
    var coordinates = data.coordinates || data.Coordinates || '';
    var timestamp = new Date().toLocaleString();

    // Append new row matching your updated columns: Center_Code | Reg No | Coordinates | Timestamp
    sheet.appendRow([centerCode, regNo, coordinates, timestamp]);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Row added successfully',
        data: { center_code: centerCode, reg_no: regNo, coordinates: coordinates, timestamp: timestamp }
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'active',
      message: 'Field DataSync Endpoint is active and ready to record Center_Code, Reg No, and Coordinates.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
