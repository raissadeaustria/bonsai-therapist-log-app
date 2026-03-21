function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['VisitDate','Timestamp','Therapist','Service','Price','Payment','Discount','DiscountCode','AmountPaid','Tip','TipPayment','ClientName','ClientEmail','ClientPhone','Currency','Status']);
  }

  // Helper: compare timestamps (Sheets converts to Date objects, so convert both to string)
  function matchTimestamp(cellValue, tsString) {
    if (!cellValue || !tsString) return false;
    var cellStr = (cellValue instanceof Date) ? cellValue.toISOString() : String(cellValue);
    // Normalize both to remove milliseconds and timezone differences
    return cellStr.replace(/\.\d{3}Z$/, '').replace(/Z$/, '') === tsString.replace(/\.\d{3}Z$/, '').replace(/Z$/, '');
  }

  // Handle delete (mark row red with DELETED status)
  if (data._action === 'delete') {
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var statusCol = headers.indexOf('Status');
    if (statusCol === -1) {
      statusCol = sheet.getLastColumn();
      sheet.getRange(1, statusCol + 1).setValue('Status');
    }
    for (var i = rows.length - 1; i >= 1; i--) {
      if (matchTimestamp(rows[i][1], data.timestamp)) {
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
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    for (var i = rows.length - 1; i >= 1; i--) {
      if (matchTimestamp(rows[i][1], data.timestamp)) {
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
  sheet.appendRow([
    data.visitDate,
    data.timestamp,
    data.therapist,
    data.service,
    data.price,
    data.payment,
    data.discount,
    data.discountCode,
    data.amountPaid,
    data.tip,
    data.tipPayment,
    data.clientName,
    data.clientEmail,
    data.clientPhone,
    data.currency
  ]);

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
