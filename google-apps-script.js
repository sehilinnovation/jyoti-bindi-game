// ============================================================
// JYOTI BINDI — Google Apps Script Backend
// Paste into Apps Script editor
// Deploy: Execute as Me, Access: Anyone
// ============================================================

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var result;

  if (action === 'register') {
    result = registerPlayer(e.parameter.name, e.parameter.phone);
  } else if (action === 'getPrizes') {
    result = getPrizes();
  } else if (action === 'claimPrize') {
    result = claimPrize(e.parameter.prizeId, e.parameter.phone);
  } else if (action === 'getPlayers') {
    result = getPlayers();
  } else {
    result = { status: 'ok', message: 'Jyoti Bindi API running' };
  }

  var output = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function registerPlayer(name, phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Players') || ss.insertSheet('Players');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Phone']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][2] == phone) return { status: 'already_played', message: 'Already played.' };
    }
    sheet.appendRow([new Date().toLocaleString('en-IN'), name, phone]);
    return { status: 'ok', message: 'Registered!' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function getPrizes() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Prizes');
    if (!sheet) return { prizes: [] };
    var data = sheet.getDataRange().getValues();
    var prizes = [];
    for (var i = 1; i < data.length; i++) {
      prizes.push({ id: data[i][0], icon: data[i][1], name: data[i][2], quantity: Number(data[i][3]) });
    }
    return { prizes: prizes };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function claimPrize(prizeId, phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Prizes');
    if (!sheet) return { status: 'error', message: 'Prizes sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == prizeId) {
        var qty = Number(data[i][3]);
        if (qty <= 0) return { status: 'out', message: 'Prize out of stock.' };
        sheet.getRange(i + 1, 4).setValue(qty - 1);
        return { status: 'ok', prize: data[i][2] };
      }
    }
    return { status: 'error', message: 'Prize not found.' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function getPlayers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Players');
    if (!sheet) return { players: [] };
    return { players: sheet.getDataRange().getValues().slice(1) };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

// ============================================================
// RUN ONCE: Select setupPrizesSheet → Run
// ============================================================
function setupPrizesSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName('Prizes');
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet('Prizes');
  sheet.appendRow(['ID', 'Icon', 'Prize Name', 'Quantity']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#d4b800').setFontColor('#0a0800');

  var prizes = [
    ['P1', '🎫', '5% Discount', 999],
  ];

  prizes.forEach(function(row) { sheet.appendRow(row); });
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 100);

  Logger.log('Prizes sheet created with ' + prizes.length + ' prizes.');
}
