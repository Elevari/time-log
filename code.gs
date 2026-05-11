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
//  8. Copy the Web App URL and hardcode it in the app
//
//  All requests use GET:
//  action=getAll        -> returns all entries and classes
//  payload=<JSON>       -> write operation (add/update/delete/etc.)
// ─────────────────────────────────────────────────────────────

var ENTRIES_SHEET = 'Sheet1';
var CLASSES_SHEET = 'Classes';

// ── ROUTER ────────────────────────────────────────────────────
function doGet(e) {
  var action  = e.parameter.action;
  var payload = e.parameter.payload;

  if (action === 'getAll') return getAll();

  if (payload) {
    try {
      var d = JSON.parse(decodeURIComponent(payload));
      if (d.action === 'add')         return addRow(d.entry);
      if (d.action === 'update')      return updateRow(d.entry);
      if (d.action === 'delete')      return deleteRow(d.id);
      if (d.action === 'addClass')    return addClass(d.name);
      if (d.action === 'deleteClass') return deleteClass(d.name);
      if (d.action === 'renameClass') return renameClass(d.oldName, d.newName);
    } catch(err) {
      return out({ error: 'Invalid payload: ' + err.message });
    }
  }

  return out({ error: 'Unknown action' });
}

function doPost(e) {
  return out({ error: 'Use GET requests with payload parameter' });
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

  var seen  = {};
  var result = [];

  // Iterate in reverse so the LAST occurrence of a duplicate wins
  // (most recent update is most reliable)
  var rows = vals.slice(1);
  for (var i = rows.length - 1; i >= 0; i--) {
    var r  = rows[i];
    var id = String(r[0]);
    if (!id || id === 'undefined') continue;
    if (seen[id]) continue;   // skip duplicate — already have a newer one
    seen[id] = true;
    result.unshift({
      id:        id,
      className: String(r[1]),
      date:      fmtDate(r[2]),
      clockIn:   fmtTime(r[3]),
      clockOut:  fmtTime(r[4]),
      createdAt: String(r[5])
    });
  }
  return result;
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
// Upsert: if entry with same ID already exists, update it instead of duplicating
function addRow(entry) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'className', 'date', 'clockIn', 'clockOut', 'createdAt']);
    sh.getRange('D:E').setNumberFormat('@');
  }

  // Check for existing row with same ID
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entry.id)) {
      // Already exists — update it
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        entry.id, entry.className, entry.date,
        entry.clockIn, entry.clockOut || '', entry.createdAt || ''
      ]]);
      return out({ ok: true, action: 'updated' });
    }
  }

  // New entry
  sh.appendRow([
    entry.id, entry.className, entry.date,
    entry.clockIn, entry.clockOut || '', entry.createdAt || ''
  ]);
  return out({ ok: true, action: 'added' });
}

// ── UPDATE ENTRY ──────────────────────────────────────────────
// Falls back to insert if entry not found (handles race condition)
function updateRow(entry) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(entry.id)) {
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        entry.id, entry.className, entry.date,
        entry.clockIn, entry.clockOut || '', entry.createdAt || ''
      ]]);
      return out({ ok: true });
    }
  }
  // Not found — insert it (race condition fallback)
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'className', 'date', 'clockIn', 'clockOut', 'createdAt']);
    sh.getRange('D:E').setNumberFormat('@');
  }
  sh.appendRow([
    entry.id, entry.className, entry.date,
    entry.clockIn, entry.clockOut || '', entry.createdAt || ''
  ]);
  return out({ ok: true, action: 'inserted' });
}

// ── DELETE ENTRY ──────────────────────────────────────────────
// Deletes ALL rows with the given ID (cleans up any duplicates too)
function deleteRow(id) {
  var sh = getOrCreateSheet(ENTRIES_SHEET);
  var data = sh.getDataRange().getValues();
  // Iterate backwards so row deletions don't shift indexes
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
    }
  }
  return out({ ok: true });
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
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === name) {
      sh.deleteRow(i + 1);
    }
  }
  return out({ ok: true });
}

// ── RENAME CLASS ──────────────────────────────────────────────
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

// ── FORMAT DATE → YYYY-MM-DD ──────────────────────────────────
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return val.getFullYear() + '-' + m + '-' + d;
  }
  return String(val);
}

// ── FORMAT TIME → HH:MM ───────────────────────────────────────
function fmtTime(val) {
  if (val === '' || val === null || val === undefined) return '';
  if (val instanceof Date) {
    return ('0' + val.getHours()).slice(-2) + ':' + ('0' + val.getMinutes()).slice(-2);
  }
  if (typeof val === 'number') {
    var totalMins = Math.round(val * 24 * 60);
    var h = Math.floor(totalMins / 60) % 24;
    var m = totalMins % 60;
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }
  var s = String(val).trim();
  if (!s) return '';
  if (/^\d{1,2}:\d{2}/.test(s)) return s.slice(0, 5);
  var match = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (match) return ('0' + match[1]).slice(-2) + ':' + match[2];
  return '';
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
