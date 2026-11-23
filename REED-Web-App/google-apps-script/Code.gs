var SYNC_SECRET = '';

function checkAuth_(e, bodyToken) {
  var expected = String(SYNC_SECRET || '').trim();
  if (!expected) return true;
  var paramToken = (e && e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : '';
  var postToken = bodyToken ? String(bodyToken).trim() : '';
  return paramToken === expected || postToken === expected;
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function readChunkedJson_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return '';
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var parts = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) parts.push(String(values[i][0]));
  }
  return parts.join('');
}

function writeChunkedJson_(sheet, text) {
  sheet.clearContents();
  var chunkSize = 45000;
  var rows = [];
  for (var i = 0; i < text.length; i += chunkSize) {
    rows.push([text.substring(i, i + chunkSize)]);
  }
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  }
  sheet.getRange(1, 2).setValue(new Date().toISOString());
}

function doGet(e) {
  if (!checkAuth_(e)) {
    return jsonOut_({ error: 'Unauthorized: Invalid sync secret token.' });
  }
  var metaSheet = getOrCreateSheet_('_VIRSA_SYNC_JSON');
  var raw = readChunkedJson_(metaSheet);
  if (!raw) {
    return jsonOut_({ parcels: [], customCouriers: [], deletedIds: [], isSampleData: true });
  }
  try {
    return ContentService.createTextOutput(raw).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

function doPost(e) {
  try {
    var rawBody = (e && e.postData && e.postData.contents) || '{}';
    var payload = JSON.parse(rawBody);
    if (!checkAuth_(e, payload.token)) {
      return jsonOut_({ error: 'Unauthorized: Invalid sync secret token.' });
    }
    var parcels = Array.isArray(payload.parcels) ? payload.parcels : [];

    var metaSheet = getOrCreateSheet_('_VIRSA_SYNC_JSON');
    writeChunkedJson_(metaSheet, rawBody);

    var viewSheet = getOrCreateSheet_('Live Ledger');
    viewSheet.clearContents();
    var headers = [
      'Tracking Number', 'Order ID', 'Courier', 'Dispatch Date',
      'Status', 'Customer Name', 'Customer Phone', 'Return Verified At',
      'Live Checkpoint', 'Notes', 'Updated At'
    ];
    var rows = [headers];
    for (var i = 0; i < parcels.length; i++) {
      var p = parcels[i];
      rows.push([
        p.trackingNumber || '',
        p.orderId || '',
        p.courier || '',
        p.dispatchDate || '',
        p.status || '',
        p.customerName || '',
        p.customerPhone || '',
        p.returnedAt || '',
        p.liveCheckpoint || '',
        p.notes || '',
        p.updatedAt || ''
      ]);
    }
    viewSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

    return jsonOut_({ ok: true, count: parcels.length, syncedAt: new Date().toISOString() });
  } catch (err) {
    return jsonOut_({ error: 'Sync write failed: ' + String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
