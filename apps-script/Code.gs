/**
 * Meal Map backend. Paste this into Extensions > Apps Script of your Google Sheet.
 * Then set Script Properties (Project Settings > Script properties):
 *   CLIENT_CODE  = the code your client types in
 *   COACH_CODE   = your own code (coach view)
 *   CLIENT_NAME  = your client's first name (shown in the app)
 * Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 */
var COLLECTIONS = ['favorites', 'notes', 'pulse', 'checks', 'extras'];

function role_(code) {
  var p = PropertiesService.getScriptProperties();
  if (code && code === p.getProperty('COACH_CODE')) return 'coach';
  if (code && code === p.getProperty('CLIENT_CODE')) return 'client';
  return null;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function dataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('data');
  if (!s) {
    s = ss.insertSheet('data');
    s.appendRow(['collection', 'id', 'json', 'updated']);
    s.setFrozenRows(1);
  }
  return s;
}

function doGet(e) {
  var r = role_(e.parameter.code);
  if (!r) return out_({ error: 'bad_code' });
  var v = dataSheet_().getDataRange().getValues();
  var docs = [];
  for (var i = 1; i < v.length; i++) {
    try { docs.push({ c: v[i][0], id: String(v[i][1]), data: JSON.parse(v[i][2]) }); } catch (err) {}
  }
  var name = PropertiesService.getScriptProperties().getProperty('CLIENT_NAME') || 'Your client';
  return out_({ role: r, clientName: name, docs: docs });
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad_json' }); }
  var r = role_(b.code);
  if (!r) return out_({ error: 'bad_code' });
  if (COLLECTIONS.indexOf(b.c) < 0) return out_({ error: 'bad_collection' });
  if (b.c === 'extras' && r !== 'coach') return out_({ error: 'coach_only' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = dataSheet_();
    var v = s.getDataRange().getValues();
    var row = -1;
    if (b.id) {
      for (var i = 1; i < v.length; i++) {
        if (v[i][0] === b.c && String(v[i][1]) === String(b.id)) { row = i + 1; break; }
      }
    }
    if (b.action === 'delete') {
      if (row > 0) s.deleteRow(row);
      return out_({ ok: true });
    }
    var id = b.id || Utilities.getUuid();
    var json = JSON.stringify(b.data || {});
    if (json.length > 40000) return out_({ error: 'too_big' });
    if (row > 0) s.getRange(row, 3, 1, 2).setValues([[json, new Date()]]);
    else s.appendRow([b.c, id, json, new Date()]);
    if (b.c === 'pulse') pulseRow_(id, b.data || {});
    if (b.c === 'notes') noteRow_(b.data || {});
    return out_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

// Easy-to-read tab of Daily Pulse entries (one row per day).
function pulseRow_(id, d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Daily Pulse');
  if (!s) {
    s = ss.insertSheet('Daily Pulse');
    s.appendRow(['Date', 'Energy', 'Pain', 'Hunger cues', 'Stress', 'Home load', 'Win']);
    s.setFrozenRows(1);
  }
  var vals = ["'" + id, d.energy || '', d.pain || '', d.hunger || '', d.stress || '', d.load || '', d.win || ''];
  var last = s.getLastRow();
  if (last > 1) {
    var dates = s.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < dates.length; i++) {
      if (dates[i][0] === id) { s.getRange(i + 2, 1, 1, 7).setValues([vals]); return; }
    }
  }
  s.appendRow(vals);
}

// Easy-to-read tab of messages and meal notes.
function noteRow_(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Messages + Notes');
  if (!s) {
    s = ss.insertSheet('Messages + Notes');
    s.appendRow(['When', 'From', 'Type', 'Meal', 'Text']);
    s.setFrozenRows(1);
  }
  s.appendRow([new Date(d.at || Date.now()), d.role === 'coach' ? 'Coach' : 'Client',
    d.kind === 'message' ? 'Message' : (d.mealId ? 'Meal note' : 'Note'), d.mealId || '', d.text || '']);
}
