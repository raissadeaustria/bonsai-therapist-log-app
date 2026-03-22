function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['VisitDate','Timestamp','Therapist','Service','Price','Payment','Discount','DiscountCode','AmountPaid','Tip','TipPayment','ClientName','ClientEmail','ClientPhone','Currency','Status']);
  }

  // Helper: normalize timestamp for comparison (remove T, .000Z, Z, trim)
  function normalizeTs(val) {
    return String(val).replace('T', ' ').replace('.000Z', '').replace('Z', '').trim();
  }

  // Handle delete (mark row red with DELETED status)
  if (data._action === 'delete') {
    var range = sheet.getDataRange();
    var rows = range.getDisplayValues();
    var headers = rows[0];
    var statusCol = headers.indexOf('Status');
    if (statusCol === -1) {
      statusCol = sheet.getLastColumn();
      sheet.getRange(1, statusCol + 1).setValue('Status');
    }
    var searchTs = normalizeTs(data.timestamp);
    for (var i = rows.length - 1; i >= 1; i--) {
      var cellTs = normalizeTs(rows[i][1]);
      if (cellTs === searchTs) {
        var lastCol = Math.max(sheet.getLastColumn(), statusCol + 1);
        sheet.getRange(i + 1, 1, 1, lastCol).setBackground('#FFCCCC');
        sheet.getRange(i + 1, statusCol + 1).setValue('DELETED');
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'deleted'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Handle update
  if (data._action === 'update') {
    var range = sheet.getDataRange();
    var rows = range.getDisplayValues();
    var headers = rows[0];
    var searchTs = normalizeTs(data.timestamp);
    for (var i = rows.length - 1; i >= 1; i--) {
      var cellTs = normalizeTs(rows[i][1]);
      if (cellTs === searchTs) {
        for (var key in data) {
          if (key === '_action' || key === 'timestamp') continue;
          var colName = key.charAt(0).toUpperCase() + key.slice(1);
          var col = headers.indexOf(colName);
          if (col === -1) col = headers.indexOf(key);
          if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(data[key]);
        }
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: 'updated'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Default: add new entry
  var newRow = sheet.getLastRow() + 1;
  var rowData = [
    String(data.visitDate != null ? data.visitDate : ''),
    String(data.timestamp != null ? data.timestamp : ''),
    String(data.therapist != null ? data.therapist : ''),
    String(data.service != null ? data.service : ''),
    String(data.price != null ? data.price : ''),
    String(data.payment != null ? data.payment : ''),
    String(data.discount != null ? data.discount : ''),
    String(data.discountCode != null ? data.discountCode : ''),
    String(data.amountPaid != null ? data.amountPaid : ''),
    String(data.tip != null ? data.tip : ''),
    String(data.tipPayment != null ? data.tipPayment : ''),
    String(data.clientName != null ? data.clientName : ''),
    String(data.clientEmail != null ? data.clientEmail : ''),
    String(data.clientPhone != null ? data.clientPhone : ''),
    String(data.currency != null ? data.currency : '')
  ];
  // Set entire row to plain text first, then write values to prevent auto-conversion
  var range = sheet.getRange(newRow, 1, 1, rowData.length);
  range.setNumberFormat('@');
  range.setValues([rowData]);

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getDisplayValues();
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
