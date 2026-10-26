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

    // Field Mapping columns:
    // Center_Code | Reg No | Coordinates | Date Of Mappng | Province | District | Name of Staff | Timestamp
    var centerCode = data.center_code || data.Center_Code || '';
    var regNo = data.reg_no || data.Reg_No || data['Reg No'] || '';
    var coordinates = data.coordinates || data.Coordinates || '';
    var dateOfMapping = data.date_of_mapping || data['Date Of Mappng'] || data.dateOfMapping || data.Date_Of_Mapping || '';
    var province = data.province || data.Province || '';
    var district = data.district || data.District || '';
    var nameOfStaff = data.name_of_staff || data['Name of Staff'] || data.nameOfStaff || data.Name_Of_Staff || '';
    var timestamp = new Date().toLocaleString();

    // Append new row matching the columns
    sheet.appendRow([centerCode, regNo, coordinates, dateOfMapping, province, district, nameOfStaff, timestamp]);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Row added successfully',
        data: {
          center_code: centerCode,
          reg_no: regNo,
          coordinates: coordinates,
          date_of_mapping: dateOfMapping,
          province: province,
          district: district,
          name_of_staff: nameOfStaff,
          timestamp: timestamp
        }
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
      message: 'Field Mapping Endpoint is active and ready to record entries.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
