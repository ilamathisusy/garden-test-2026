/**
 * A Garden of Wishes — collects guests' wishes in this Google Sheet.
 * Only people you share this Sheet with can read the wishes.
 *
 * Setup: Extensions → Apps Script → paste this file → Save →
 *        Deploy → New deployment → Web app →
 *        Execute as: Me · Who has access: Anyone → Deploy → copy the Web app URL.
 */
var SHEET_NAME = 'Wishes';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var name = clean(data.name, 40);
    var wish = clean(data.wish, 280);
    var flower = clean(data.flower, 30);
    if (!name || wish.length < 3) return ContentService.createTextOutput('ignored');

    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Received', 'Name', 'Flower', 'Wish']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), name, flower, wish]);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  } finally {
    lock.releaseLock();
  }
}

// Trims, limits length and stops text being read as a spreadsheet formula.
function clean(value, max) {
  var s = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
