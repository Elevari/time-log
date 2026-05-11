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
// ─────────────────────────────────────────────────────────────

const ENTRIES_SHEET = 'Sheet1';
const CLASSES_SHEET = 'Classes';

// ── GET — returns both entries and classes ────────────────────
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
  return out({ error: 'Unknown action' });
}

// ── READ ALL (entries + classes) ──────────────────────────────
function getAll() {
  return out({
    entries: getEntries(),
    classes: getClasses()
  });
}

function getEntries() {
  const sh = getOrCreateSheet(ENTRIES_SHEET);
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1)
    .map(r => ({
      id:        String(r[0]),
      className: String(r[1]),
      date:      fmtDate(r[2]),
      clockIn:   fmtTime(r[3]),
      clockOut:  fmtTime(r[4]),
      createdAt: String(r[5])
    }))
    .filter(r => r.id);
}

function getClasses() {
  const sh = getOrCreateSheet(CLASSES_SHEET);
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1)
    .map(r => String(r[0]))
    .filter(name => name);
}

// ── ADD ENTRY ─────────────────────────────────────────────────
function addRow(entry) {
  const sh = getOrCreateSheet(ENTRIES_SHEET);
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
  const sh = getOrCreateSheet(ENTRIES_SHEET);
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
  const sh = getOrCreateSheet(ENTRIES_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return out({ ok: true });
    }
  }
  return out({ error: 'Entry not found' });
}

// ── ADD CLASS ─────────────────────────────────────────────────
function addClass(name) {
  const sh = getOrCreateSheet(CLASSES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['name']);
  }
  const existing = sh.getDataRange().getValues().map(r => String(r[0]));
  if (existing.includes(name)) return out({ ok: true, note: 'already exists' });
  sh.appendRow([name]);
  return out({ ok: true });
}

// ── DELETE CLASS ──────────────────────────────────────────────
function deleteClass(name) {
  const sh = getOrCreateSheet(CLASSES_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      sh.deleteRow(i + 1);
      return out({ ok: true });
    }
  }
  return out({ error: 'Class not found' });
}

// ── HELPERS ───────────────────────────────────────────────────

// Format a cell value as YYYY-MM-DD regardless of how Sheets stored it
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth()+1).padStart(2,'0');
    const d = String(val.getDate()).padStart(2,'0');
    return y + '-' + m + '-' + d;
  }
  // Already a string — return as-is
  return String(val);
}

// Format a cell value as HH:MM
// Handles: Sheets decimal fraction, HH:MM:SS string, HH:MM string
function fmtTime(val) {
  if (val === '' || val === null || val === undefined) return '';
  // Sheets stores times as decimal fractions of a day (e.g. 0.708333 = 17:00)
  if (typeof val === 'number') {
    var totalMins = Math.round(val * 24 * 60);
    var h = Math.floor(totalMins / 60) % 24;
    var m = totalMins % 60;
    return ('0'+h).slice(-2) + ':' + ('0'+m).slice(-2);
  }
  var s = String(val).trim();
  if (!s) return '';
  // HH:MM:SS → HH:MM
  if (/^\d{1,2}:\d{2}:\d{2}$/.test(s)) return s.slice(0, 5);
  // Already HH:MM
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  return s;
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
