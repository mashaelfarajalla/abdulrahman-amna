/**
 * Wedding invitation — RSVP backend (Google Sheets)
 *
 * Setup:
 * 1. Create a new Google Sheet.
 * 2. Extensions → Apps Script, paste this whole file, save.
 * 3. Deploy → New deployment → type "Web app":
 *      Execute as: Me    |    Who has access: Anyone
 * 4. Copy the Web app URL and paste it into CONFIG.sheetUrl in index.html.
 *
 * Replies land in an "RSVP" tab that is created automatically.
 */

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Trim, limit length, and stop values from being read as spreadsheet formulas
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Returns the latest wishes so the invitation page can show them
function doGet(e) {
  const out = { wishes: [] };
  if (e && e.parameter && e.parameter.action === 'wishes') {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RSVP');
    if (sh && sh.getLastRow() > 1) {
      out.wishes = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
        .filter(r => String(r[4]).trim())
        .map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: String(r[1]).replace(/^'/, ''), message: String(r[4]).replace(/^'/, '') }))
        .reverse()
        .slice(0, 100);
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.type === 'rsvp') {
    sheet_('RSVP', ['Time', 'Name', 'Response', 'Contact', 'Message'])
      .appendRow([new Date(), clean_(d.name, 80), clean_(d.response, 10), clean_(d.contact, 120), clean_(d.message, 500)]);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
