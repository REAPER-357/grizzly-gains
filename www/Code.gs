/**
 * GRIZZLY GAINS — Google Sheets backup endpoint
 *
 * SETUP (2 minutes)
 *  1. Create a new Google Sheet. Open  Extensions > Apps Script.
 *  2. Delete everything in the editor, paste this whole file, click Save.
 *  3. Deploy > New deployment > type "Web app"
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy, approve the permissions, copy the Web app URL.
 *  4. In Grizzly Gains: Settings > Cloud Backup > paste the URL > Save & Test.
 *
 * OPTIONAL password: Project Settings (gear) > Script properties > add
 *   TOKEN = any-secret-word   and enter the same word in the app.
 *
 * Tabs created automatically:
 *   History     one row per workout session        (id is the key)
 *   Workouts    one row per routine
 *   BodyWeight  one row per weigh-in
 *   Exercises   one row per exercise, rebuilt from Workouts. Share THIS sheet
 *               ("Anyone with the link: Viewer") so friends can import it.
 *
 * After changing this script you must Deploy > Manage deployments > Edit >
 * Version: New version > Deploy.
 */

var TABS = {
  history:    { name: 'History',    headers: ['id', 'date', 'workout', 'group', 'summary', 'notes', 'updatedAt', 'json'] },
  workouts:   { name: 'Workouts',   headers: ['id', 'name', 'group', 'order', 'exercises', 'updatedAt', 'json'] },
  bodyweight: { name: 'BodyWeight', headers: ['id', 'date', 'weight', 'unit', 'updatedAt', 'json'] }
};
var EX_TAB = 'Exercises';
var EX_HEADERS = ['workout', 'group', 'exercise', 'type', 'sets', 'reps', 'distance', 'duration', 'level', 'notes'];

/* ---------- entry points ---------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var body = JSON.parse(e.postData.contents);
    if (!authOk_(body.token)) return out_({ ok: false, error: 'Wrong token' });
    var n = applyOps_(body.ops || []);
    return out_({ ok: true, applied: n });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (!authOk_(p.token)) return out_({ ok: false, error: 'Wrong token' });
    var action = p.action || 'ping';
    if (action === 'ping') return out_({ ok: true, name: SpreadsheetApp.getActiveSpreadsheet().getName() });
    if (action === 'all') return out_(readAll_());
    if (action === 'exercises') return out_(readExerciseRows_(p.sheet));
    return out_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/** Optional: run once from the editor to create all tabs up front. */
function setup() {
  Object.keys(TABS).forEach(function (k) { sheet_(k); });
  rebuildExercises_();
}

/* ---------- helpers ---------- */

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function authOk_(t) {
  var tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !tok || tok === t;
}

function sheet_(kind) {
  var def = TABS[kind];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.appendRow(def.headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, def.headers.length).setFontWeight('bold');
    if (kind === 'history' || kind === 'bodyweight') sh.getRange('B:B').setNumberFormat('@'); // keep dates as text
  }
  return sh;
}

function unitDist_(u) { return u === 'lbs' ? 'mi' : 'km'; }

function summary_(s) {
  var u = s.unit || 'kg';
  return (s.rawExercises || []).map(function (ex) {
    if (ex.type === 'cardio' && ex.cardio) {
      var c = ex.cardio, bits = [];
      if (c.distance) bits.push(c.distance + unitDist_(u));
      if (c.duration) bits.push(c.duration + 'min');
      if (c.level) bits.push('L' + c.level);
      return ex.name + ': ' + bits.join(' ');
    }
    return ex.name + ': ' + (ex.sets || []).map(function (st) { return st.weight + u + 'x' + st.reps; }).join(', ');
  }).join(' | ');
}

function row_(kind, d) {
  var json = JSON.stringify(d);
  if (kind === 'history')
    return [d.id, d.date || '', d.name || '', d.group || '', summary_(d), d.notes || '', d.updatedAt || 0, json];
  if (kind === 'workouts')
    return [d.id, d.name || '', d.group || '', d.order || 0, (d.exercises || []).length, d.updatedAt || 0, json];
  return [d.id, d.date || '', d.weight, d.unit || 'kg', d.updatedAt || 0, json];
}

function applyOps_(ops) {
  var byKind = {}, touched = {}, count = 0;
  ops.forEach(function (o) {
    if (o && TABS[o.kind] && o.id) (byKind[o.kind] = byKind[o.kind] || []).push(o);
  });

  Object.keys(byKind).forEach(function (kind) {
    var sh = sheet_(kind);
    var last = sh.getLastRow();
    var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    var index = {};
    ids.forEach(function (r, i) { index[String(r[0])] = i + 2; });
    var deletes = [];

    byKind[kind].forEach(function (o) {
      var rowNum = index[String(o.id)];
      if (o.op === 'delete') {
        if (rowNum) deletes.push(rowNum);
        count++;
        return;
      }
      if (!o.data) return;
      var vals = row_(kind, o.data);
      if (rowNum) {
        sh.getRange(rowNum, 1, 1, vals.length).setValues([vals]);
      } else {
        sh.appendRow(vals);
        index[String(o.id)] = sh.getLastRow();
      }
      count++;
    });

    deletes.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
    touched[kind] = true;
  });

  if (touched.workouts) rebuildExercises_();
  return count;
}

function readKind_(kind) {
  var sh = sheet_(kind);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var jsonCol = TABS[kind].headers.length;
  var vals = sh.getRange(2, jsonCol, last - 1, 1).getValues();
  var out = [];
  vals.forEach(function (r) {
    try { var o = JSON.parse(r[0]); if (o && o.id) out.push(o); } catch (e) {}
  });
  return out;
}

function readAll_() {
  return {
    ok: true,
    workouts: readKind_('workouts'),
    history: readKind_('history'),
    bodyweight: readKind_('bodyweight')
  };
}

/** Rebuild the shareable "Exercises" tab: one row per exercise. */
function rebuildExercises_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(EX_TAB) || ss.insertSheet(EX_TAB);
  var workouts = readKind_('workouts').sort(function (a, b) { return (a.order || 0) - (b.order || 0); });
  var rows = [EX_HEADERS];
  workouts.forEach(function (w) {
    (w.exercises || []).forEach(function (ex) {
      var cardio = ex.type === 'cardio';
      rows.push([
        w.name, w.group || 'gym', ex.name, cardio ? 'cardio' : 'strength',
        cardio ? '' : (ex.sets || ''), cardio ? '' : (ex.reps || ''),
        cardio ? (ex.targetDistance || '') : '', cardio ? (ex.targetDuration || '') : '', cardio ? (ex.targetLevel || '') : '',
        ex.notes || ''
      ]);
    });
  });
  sh.clearContents();
  sh.getRange(1, 1, rows.length, EX_HEADERS.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, EX_HEADERS.length).setFontWeight('bold');
}

/** Read the Exercises tab (or first tab) of ANY sheet this account can open. */
function readExerciseRows_(ref) {
  if (!ref) return { ok: false, error: 'No sheet given' };
  var m = String(ref).match(/\/d\/([a-zA-Z0-9-_]+)/);
  var ss = SpreadsheetApp.openById(m ? m[1] : String(ref).trim());
  var sh = ss.getSheetByName(EX_TAB) || ss.getSheets()[0];
  var vals = sh.getDataRange().getDisplayValues();
  return { ok: true, rows: vals };
}
