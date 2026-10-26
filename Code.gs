function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};
    
    // Support JSON payload or Form/URL-encoded parameters
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // Extract all 8 fields with generous aliases
    var centerCode = data.center_code || data.Center_Code || data['Center Code'] || '';
    var regNo = data.reg_no || data.Reg_No || data['Reg No'] || data.regNo || '';
    var coordinates = data.coordinates || data.Coordinates || data.GPS || '';
    var dateOfMapping = data.date_of_mapping || data['Date Of Mapping'] || data['Date Of Mappng'] || data.dateOfMapping || data.date || '';
    var province = data.province || data.Province || '';
    var district = data.district || data.District || '';
    var nameOfStaff = data.name_of_staff || data['Name of Staff'] || data.nameOfStaff || data.staff_name || data.staff || '';
    var timestamp = data.timestamp || new Date().toLocaleString();

    // Standard 8 Field Mapping Headers:
    var standardHeaders = [
      'Center_Code',
      'Reg No',
      'Coordinates',
      'Date Of Mapping',
      'Province',
      'District',
      'Name of Staff',
      'Timestamp'
    ];

    // Helper to normalize header string for reliable matching
    function normalize(str) {
      return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    var lastCol = sheet.getLastColumn();
    var existingHeaders = [];
    if (lastCol > 0) {
      existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // If sheet has no headers in Row 1, create all 8 standard headers
    if (existingHeaders.length === 0 || (existingHeaders.length === 1 && String(existingHeaders[0]).trim() === '')) {
      sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
      sheet.getRange(1, 1, 1, standardHeaders.length).setFontWeight('bold');
      existingHeaders = standardHeaders;
      lastCol = standardHeaders.length;
    }

    // Field value mapping table
    var fieldMap = {
      'centercode': centerCode,
      'center': centerCode,
      'regno': regNo,
      'registration': regNo,
      'registrationno': regNo,
      'coordinates': coordinates,
      'gps': coordinates,
      'location': coordinates,
      'dateofmapping': dateOfMapping,
      'dateofmappng': dateOfMapping,
      'date': dateOfMapping,
      'mappingdate': dateOfMapping,
      'province': province,
      'district': district,
      'nameofstaff': nameOfStaff,
      'staffname': nameOfStaff,
      'staff': nameOfStaff,
      'timestamp': timestamp,
      'time': timestamp,
      'datetime': timestamp
    };

    var rowData = new Array(existingHeaders.length);
    var matchedKeys = {};

    for (var i = 0; i < existingHeaders.length; i++) {
      var norm = normalize(existingHeaders[i]);
      if (fieldMap.hasOwnProperty(norm)) {
        rowData[i] = fieldMap[norm];
        matchedKeys[norm] = true;
      } else {
        rowData[i] = '';
      }
    }

    // Auto-expand headers in Row 1 if sheet was missing Province, District, Name of Staff, or Date
    var missingHeaders = [];
    var missingValues = [];

    if (!matchedKeys['dateofmapping'] && !matchedKeys['dateofmappng'] && !matchedKeys['date']) {
      missingHeaders.push('Date Of Mapping');
      missingValues.push(dateOfMapping);
    }
    if (!matchedKeys['province']) {
      missingHeaders.push('Province');
      missingValues.push(province);
    }
    if (!matchedKeys['district']) {
      missingHeaders.push('District');
      missingValues.push(district);
    }
    if (!matchedKeys['nameofstaff'] && !matchedKeys['staffname'] && !matchedKeys['staff']) {
      missingHeaders.push('Name of Staff');
      missingValues.push(nameOfStaff);
    }

    if (missingHeaders.length > 0) {
      var startNewCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, startNewCol, 1, missingHeaders.length).setValues([missingHeaders]);
      sheet.getRange(1, startNewCol, 1, missingHeaders.length).setFontWeight('bold');
      for (var m = 0; m < missingValues.length; m++) {
        rowData.push(missingValues[m]);
      }
    }

    // Append the row to sheet
    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Row added successfully',
        recorded: {
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
      message: 'Field Mapping Endpoint is active and ready to record all 8 fields.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
