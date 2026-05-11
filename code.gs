// ─────────────────────────────────────────────────────────────
//  Time.Log — Google Apps Script Backend
//
//  SETUP
//  ─────────────────────────────────────────────────────────────
//  1. Open your Google Sheet (any file name is fine).
//  2. Two sheet tabs are created automatically on first use:
//       - Sheet1   (time entries)
//       - Classes  (class names)
//  3. Click Extensions → Apps Script
//  4. Delete all existing code and paste this entire file
//  5. Click Save (floppy disk icon)
//  6. Click Deploy → New deployment
//       Type:            Web app
//       Execute as:      Me
//       Who has access:  Anyone
//  7. Click Deploy and authorize when prompted
//  8. Copy the Web App URL and paste it into the app Settings
//
//  SHEET STRUCTURE
//  ─────────────────────────────────────────────────────────────
//  Sheet1  → A:id | B:className | C:date | D:clockIn | E:clockOut | F:createdAt
//  Classes → A:name
//
//  IMPORTANT: Format columns D and E (clockIn, clockOut) as
//  Plain Text to prevent Sheets auto-converting times to Date objects.
// ─────────────────────────────────────────────────────────────

const ENTRIES_SHEET = 'Sheet1';
const CLASSES_SHEET = 'Classes';

// ── GET ───────────────────────────────────────────────────────
function doGet(e) {
  if (e.parameter.action === 'getAll') return getAll();
  return out({ error: 'Unknown action' });
}

// ── POST ──────────────────────────────────────────────────────
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action === 'add')         return addRow(d.entry);
  if (d.action === 'update')      return updateRow(d.entry);
  if (d.action === 'delete')      return deleteRow(d.id);
  if (d.action === 'addClass')    return addClass(d.name);
  if (d.action === 'deleteClass') return deleteClass(d.name);
  if (d.action === 'renameClass') return renameClass(d.oldName, d.newName);
  return out({ error: 'Unknown action' });
}

// ── READ ALL ──────────────────────────────────────────────────
function getAll() {
  return out({
    entries: getEntries(),
    classes: getClasses()
  });
}

function getEntries() {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1)
    .map(function(r) {
      return {
        id:        String(r[0]),
        className: String(r[1]),
        date:      fmtDate(r[2]),
        clockIn:   fmtTime(r[3]),
        clockOut:  fmtTime(r[4]),
        createdAt: String(r[5])
      };
    })
    .filter(function(r) { return r.id && r.id !== 'undefined'; });
}

function getClasses() {
  var sh = getOrCreateSheet(CLASSES_SHEET);
  var vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1)
    .map(function(r) { return String(r[0]); })
    .filter(function(name) { return name && name !== 'undefined'; });
}

// ── ADD ENTRY ─────────────────────────────────────────────────
function addRow(entry) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'className', 'date', 'clockIn', 'clockOut', 'createdAt']);
    // Format clockIn/clockOut columns as plain text so Sheets doesn't convert times
    sh.getRange('D:E').setNumberFormat('@');
  }
  sh.appendRow([
    entry.id,
    entry.className,
    entry.date,
    entry.clockIn   || '',
    entry.clockOut  || '',
    entry.createdAt || ''
  ]);
  return out({ ok: true });
}

// ── UPDATE ENTRY ──────────────────────────────────────────────
function updateRow(entry) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entry.id)) {
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        entry.id,
        entry.className,
        entry.date,
        entry.clockIn   || '',
        entry.clockOut  || '',
        entry.createdAt || ''
      ]]);
      return out({ ok: true });
    }
  }
  return out({ error: 'Entry not found' });
}

// ── DELETE ENTRY ──────────────────────────────────────────────
function deleteRow(id) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return out({ ok: true });
    }
  }
  return out({ error: 'Entry not found' });
}

// ── ADD CLASS ─────────────────────────────────────────────────
function addClass(name) {
  var sh = getOrCreateSheet(CLASSES_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(['name']);
  var existing = sh.getDataRange().getValues().map(function(r) { return String(r[0]); });
  if (existing.indexOf(name) !== -1) return out({ ok: true, note: 'already exists' });
  sh.appendRow([name]);
  return out({ ok: true });
}

// ── DELETE CLASS ──────────────────────────────────────────────
function deleteClass(name) {
  var sh = getOrCreateSheet(CLASSES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      sh.deleteRow(i + 1);
      return out({ ok: true });
    }
  }
  return out({ error: 'Class not found' });
}

// ── FORMAT DATE → YYYY-MM-DD ──────────────────────────────────
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1).padStart('0', 2);
    var d = String(val.getDate()).padStart('0', 2);
    // padStart doesn't exist in GAS — use slice trick
    m = ('0' + (val.getMonth() + 1)).slice(-2);
    d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

// ── FORMAT TIME → HH:MM ───────────────────────────────────────
// Google Sheets stores times as decimal fractions of a day
// OR as Date objects (epoch 1899-12-30 + time fraction)
// This function handles all cases and returns clean HH:MM
function fmtTime(val) {
  if (val === '' || val === null || val === undefined) return '';

  // Date object (how Sheets internally represents time cells)
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }

  // Number — decimal fraction of a day (0.708333 = 17:00)
  if (typeof val === 'number') {
    var totalMins = Math.round(val * 24 * 60);
    var hh = Math.floor(totalMins / 60) % 24;
    var mm = totalMins % 60;
    return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
  }

  // String
  var s = String(val).trim();
  if (!s) return '';

  // Already HH:MM or HH:MM:SS
  if (/^\d{1,2}:\d{2}/.test(s)) return s.slice(0, 5);

  // Date string like "Sat Dec 30 1899 17:00:00 GMT+0000"
  // Extract HH:MM with regex
  var match = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (match) {
    return ('0' + match[1]).slice(-2) + ':' + match[2];
  }

  return '';
}

// ── RENAME CLASS ─────────────────────────────────────────────
function renameClass(oldName, newName) {
  // Update Classes sheet
  var sh = getOrCreateSheet(CLASSES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === oldName) {
      sh.getRange(i + 1, 1).setValue(newName);
      break;
    }
  }
  // Update className in all entries
  var esh = getOrCreateSheet(ENTRIES_SHEET);
  var edata = esh.getDataRange().getValues();
  for (var j = 1; j < edata.length; j++) {
    if (String(edata[j][1]) === oldName) {
      esh.getRange(j + 1, 2).setValue(newName);
    }
  }
  return out({ ok: true });
}

// ── HELPERS ───────────────────────────────────────────────────
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
