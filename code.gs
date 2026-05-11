// ─────────────────────────────────────────────────────────────
//  Time.Log — Google Apps Script Backend
//  
//  SETUP INSTRUCTIONS
//  ─────────────────────────────────────────────────────────────
//  1. Open your Google Sheet (any name is fine).
//  2. Make sure the first sheet tab is named: Sheet1
//     (or update the SHEET constant below to match your tab name)
//  3. Click Extensions → Apps Script
//  4. Delete all existing code and paste this entire file
//  5. Click Save (floppy disk icon)
//  6. Click Deploy → New deployment
//     - Type:            Web app
//     - Execute as:      Me
//     - Who has access:  Anyone
//  7. Click Deploy and authorize when prompted
//  8. Copy the Web App URL and paste it into the app's Settings
//
//  SHEET STRUCTURE (auto-created on first entry)
//  ─────────────────────────────────────────────────────────────
//  A: id | B: className | C: date | D: clockIn | E: clockOut | F: createdAt
// ─────────────────────────────────────────────────────────────

const SHEET = 'Sheet1'; // Change this if your sheet tab has a different name

// ── GET ───────────────────────────────────────────────────────
function doGet(e) {
  if (e.parameter.action === 'getAll') return getAll();
  return out({ error: 'Unknown action' });
}

// ── POST ──────────────────────────────────────────────────────
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action === 'add')    return addRow(d.entry);
  if (d.action === 'update') return updateRow(d.entry);
  if (d.action === 'delete') return deleteRow(d.id);
  return out({ error: 'Unknown action' });
}

// ── READ ALL ENTRIES ──────────────────────────────────────────
function getAll() {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET);

  if (!sh) return out({ entries: [] });

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return out({ entries: [] }); // header only or empty

  const rows = vals.slice(1); // skip header row
  const entries = rows
    .map(r => ({
      id:        String(r[0]),
      className: String(r[1]),
      date:      String(r[2]),
      clockIn:   String(r[3]),
      clockOut:  String(r[4]),
      createdAt: String(r[5])
    }))
    .filter(r => r.id); // skip any blank rows

  return out({ entries });
}

// ── ADD ENTRY ─────────────────────────────────────────────────
function addRow(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);

  // Create header row if sheet is empty
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'className', 'date', 'clockIn', 'clockOut', 'createdAt']);
  }

  sh.appendRow([
    entry.id,
    entry.className,
    entry.date,
    entry.clockIn,
    entry.clockOut || '',
    entry.createdAt
  ]);

  return out({ ok: true });
}

// ── UPDATE ENTRY ──────────────────────────────────────────────
function updateRow(entry) {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET);

  if (!sh) return out({ error: 'Sheet not found' });

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entry.id)) {
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        entry.id,
        entry.className,
        entry.date,
        entry.clockIn,
        entry.clockOut || '',
        entry.createdAt
      ]]);
      return out({ ok: true });
    }
  }

  return out({ error: 'Entry not found' });
}

// ── DELETE ENTRY ──────────────────────────────────────────────
function deleteRow(id) {
  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET);

  if (!sh) return out({ error: 'Sheet not found' });

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return out({ ok: true });
    }
  }

  return out({ error: 'Entry not found' });
}

// ── RESPONSE HELPER ───────────────────────────────────────────
function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
