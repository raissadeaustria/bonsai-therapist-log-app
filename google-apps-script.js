function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['VisitDate','Timestamp','Therapist','Service','Price','Payment','Discount','DiscountCode','AmountPaid','Tip','TipPayment','ClientName','ClientEmail','ClientPhone','Currency']);
  }

  // Handle delete
  if (data._action === 'delete') {
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][1] === data.timestamp) {
        sheet.deleteRow(i + 1);
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
      if (rows[i][1] === data.timestamp) {
        for (var key in data) {
          if (key === '_action' || key === 'timestamp') continue;
          var col = headers.indexOf(key.charAt(0).toUpperCase() + key.slice(1));
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
