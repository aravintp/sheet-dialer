// Sheet Dialer helper. Paste into Extensions → Apps Script in your Google Sheet,
// then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
// The web app link it gives you is what you paste into Sheet Dialer.
// Treat that link like a password: anyone who has it can read and update this sheet.

// GET ?gid=123&statusCol=5 → the tab's values (row 1 is the header row), the list of
// tabs, and the options of the status column's dropdown. statusCol is 1-based; without
// it, the first column whose header contains "status" is used.
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheet_(e.parameter.gid);
  const values = sh.getDataRange().getDisplayValues();
  let statusCol = Number(e.parameter.statusCol) || 0;
  if (!statusCol) statusCol = (values[0] || []).findIndex(h => /status/i.test(h)) + 1;
  return json_({
    name: sh.getName(),
    gid: sh.getSheetId(),
    tabs: ss.getSheets().map(s => ({ gid: s.getSheetId(), name: s.getName() })),
    values: values,
    statusOptions: statusCol ? dropdownOptions_(sh, statusCol) : [],
  });
}

// POST { gid, row, phoneCol, phone, cells: [{ col, value }] }
// Columns are 1-based. If the row no longer holds that phone number (the sheet was
// sorted or rows were inserted), the row is found again by searching the phone column.
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(d.gid);
    let row = d.row;
    if (d.phoneCol && d.phone && sh.getRange(row, d.phoneCol).getDisplayValue() !== d.phone) {
      const phones = sh.getRange(1, d.phoneCol, sh.getLastRow(), 1).getDisplayValues();
      const i = phones.findIndex(p => p[0] === d.phone);
      if (i < 0) return json_({ ok: false, error: "Couldn't find " + d.phone + " in the sheet" });
      row = i + 1;
    }
    d.cells.forEach(c => sh.getRange(row, c.col).setValue(c.value));
    return json_({ ok: true, row: row });
  } finally {
    lock.releaseLock();
  }
}

// Options of the dropdown on the first data cell of a column, from a list or a range.
function dropdownOptions_(sh, col) {
  const rule = sh.getRange(2, col).getDataValidation();
  if (!rule) return [];
  const type = rule.getCriteriaType();
  const args = rule.getCriteriaValues();
  const C = SpreadsheetApp.DataValidationCriteria;
  if (type === C.VALUE_IN_LIST) return args[0].filter(String);
  if (type === C.VALUE_IN_RANGE) return args[0].getDisplayValues().flat().filter(String);
  return [];
}

function getSheet_(gid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(s => String(s.getSheetId()) === String(gid)) || ss.getSheets()[0];
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
