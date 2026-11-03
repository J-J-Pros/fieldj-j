/**
 * J&J Obra — sync endpoint and database.
 * Lives inside the J&J data Google Sheet (Extensions > Apps Script).
 *
 * One-time setup:
 *   1. Run setup() once. It builds the tabs, the private receipt folder, and the device key.
 *   2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *   3. Reload the Sheet, then J&J Obra menu > Phone setup link.
 *
 * Privacy: every request must carry the device key. GET never returns data.
 * Day rates live only on the Rates tab and are never sent to the phone.
 */

var TABS = {
  jobs:  { name: 'Jobs',  cols: ['id', 'name', 'active', 'createdAt', 'updatedAt'] },
  crew:  { name: 'Crew',  cols: ['id', 'name', 'active', 'createdAt', 'updatedAt'] },
  labor: { name: 'Labor', cols: ['id', 'date', 'jobId', 'jobName', 'step', 'stepLabel', 'workerId', 'workerName', 'days', 'allocatedDays', 'savedAt', 'device', 'syncedAt'] },
  costs: { name: 'Costs', cols: ['id', 'date', 'jobId', 'jobName', 'step', 'stepLabel', 'vendor', 'category', 'categoryLabel', 'amount', 'note', 'photoUrl', 'savedAt', 'device', 'syncedAt'] }
};
var RATES = { name: 'Rates', cols: ['workerName', 'dayRate', 'effectiveFrom'] };
var TEXT_COLS = ['id', 'date', 'jobId', 'workerId', 'createdAt', 'updatedAt', 'savedAt', 'syncedAt', 'name', 'vendor', 'note', 'jobName', 'workerName', 'photoUrl', 'step'];
var CAT_LABEL = { material: 'Material', rental: 'Equipment rental', fuel: 'Fuel', tool: 'Tool',
  dumpster: 'Dumpster', permit: 'Permit', sub: 'Outside labor', other: 'Other' };
var STEP_LABEL = { hang: 'Drywall hang', finish: 'Tape and finish', plaster: 'Lath and plaster', stucco: 'Stucco finish', '': 'General' };
var RECENT_DAYS = 62;

/* ---------------- setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('J&J Obra')
    .addItem('Phone setup link', 'setupLink')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var defs = Object.keys(TABS).map(function (t) { return TABS[t]; }).concat([RATES]);
  defs.forEach(function (def) {
    var sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    def.cols.forEach(function (c, i) {
      if (TEXT_COLS.indexOf(c) >= 0 || c === 'effectiveFrom') sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  receiptFolder();
  Logger.log('Setup complete. Now deploy as a Web app, then use J&J Obra > Phone setup link.');
}

function receiptFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('J&J Obra Receipts');
  props.setProperty('FOLDER', f.getId());
  return f;
}

function setupLink() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var r = ui.prompt('Phone setup link', 'Paste the Web app URL (ends in /exec):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var exec = r.getResponseText().trim();
  var appUrl = props.getProperty('APP_URL') || 'https://j-j-pros.github.io/fieldj-j/';
  var cfg = { u: exec, k: props.getProperty('KEY'), n: 'Jose', l: 'es' };
  var s = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(cfg)).getBytes()).replace(/=+$/, '');
  var link = appUrl + '?s=' + s;
  Logger.log(link);
  ui.alert('Jose\'s setup link (treat it like a password)', link, ui.ButtonSet.OK);
}

/** Issue a new device key. Every phone must then be re-linked with a new setup link. */
function rotateKey() {
  PropertiesService.getScriptProperties().setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('Key rotated. Make a new Phone setup link for each phone.');
}

/* ---------------- web app ---------------- */

function doGet() {
  return out({ ok: false });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, err: 'bad json' }); }
  if (!authorized(body.k)) return out({ ok: false, err: 'auth' });

  if (body.a === 'pull') {
    return out({ ok: true, jobs: readTab('jobs'), crew: readTab('crew'),
      labor: readTab('labor', RECENT_DAYS), costs: readTab('costs', RECENT_DAYS) });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (body.a === 'photo') {
      savePhoto(String(body.id || ''), body.data, body.mime || 'image/jpeg');
      SpreadsheetApp.flush();
      return out({ ok: true });
    }
    var now = new Date().toISOString();
    var touched = {};
    (Array.isArray(body.ops) ? body.ops : []).forEach(function (op) {
      if (!TABS[op.t] || !op.id) return;
      var id = String(op.id);
      if (op.t === 'labor') {
        var prev = readRow('labor', id);
        if (prev) touched[prev.date + '|' + prev.workerId] = 1;
      }
      if (op.a === 'del') removeRow(op.t, id);
      else if (op.a === 'put' && op.d) {
        var r = clean(op.t, op.d, body.dev, now);
        if (op.t === 'costs') { var old = readRow('costs', id); if (old && old.photoUrl) r.photoUrl = old.photoUrl; }
        upsertRow(op.t, id, r);
        if (op.t === 'labor') touched[r.date + '|' + r.workerId] = 1;
      }
    });
    Object.keys(touched).forEach(function (k) { var p = k.split('|'); reallocate(p[0], p[1]); });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return out({ ok: true });
}

/* ---------------- helpers ---------------- */

function authorized(k) {
  var key = PropertiesService.getScriptProperties().getProperty('KEY');
  return !!key && k === key;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet(t) { return SpreadsheetApp.getActive().getSheetByName(TABS[t].name); }

function clean(t, d, dev, now) {
  var r = {};
  TABS[t].cols.forEach(function (c) { r[c] = d[c] == null ? '' : d[c]; });
  if (t === 'labor') { r.days = 1; r.allocatedDays = 1; r.stepLabel = STEP_LABEL[d.step] || d.step || ''; }
  if (t === 'costs') {
    r.amount = Math.round(Number(d.amount) * 100) / 100;
    r.categoryLabel = CAT_LABEL[d.category] || d.category || '';
    r.stepLabel = STEP_LABEL[d.step || ''] || d.step;
    r.photoUrl = '';
  }
  if (t === 'jobs' || t === 'crew') r.active = d.active === false ? false : true;
  if ('device' in r) r.device = dev || '';
  if ('syncedAt' in r) r.syncedAt = now;
  return r;
}

/** One person, one paid day: if they appear on N jobs that date, each row carries 1/N of the day. */
function reallocate(date, workerId) {
  var sh = sheet('labor'), cols = TABS.labor.cols, last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  var di = cols.indexOf('date'), wi = cols.indexOf('workerId'), ai = cols.indexOf('allocatedDays');
  var rows = [];
  for (var i = 0; i < vals.length; i++) if (vals[i][di] === date && vals[i][wi] === workerId) rows.push(i + 2);
  var share = rows.length ? Math.round(10000 / rows.length) / 10000 : 0;
  rows.forEach(function (r) { sh.getRange(r, ai + 1).setValue(share); });
}

function savePhoto(id, b64, mime) {
  if (!id || !b64) return;
  var row = readRow('costs', id);
  var name = (row ? [row.date, row.jobName, row.vendor, row.amount].join(' ') : id).replace(/[\/\\:*?"<>|]/g, '-') + '.jpg';
  var file = receiptFolder().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  if (row) {
    var sh = sheet('costs'), at = findRow(sh, id), ci = TABS.costs.cols.indexOf('photoUrl');
    if (at > 0) sh.getRange(at, ci + 1).setValue(file.getUrl());
  }
}

function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function readRow(t, id) {
  var sh = sheet(t), at = findRow(sh, id), cols = TABS[t].cols;
  if (at < 0) return null;
  var v = sh.getRange(at, 1, 1, cols.length).getDisplayValues()[0], o = {};
  cols.forEach(function (c, j) { o[c] = v[j]; });
  return o;
}

function upsertRow(t, id, r) {
  var sh = sheet(t), cols = TABS[t].cols;
  var row = cols.map(function (c) { return r[c]; });
  var at = findRow(sh, id);
  if (at > 0) sh.getRange(at, 1, 1, cols.length).setValues([row]);
  else sh.appendRow(row);
}

function removeRow(t, id) {
  var sh = sheet(t), at = findRow(sh, id);
  if (at > 0) sh.deleteRow(at);
}

function readTab(t, days) {
  var sh = sheet(t), cols = TABS[t].cols, last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var disp = sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  var cutoff = days ? Utilities.formatDate(new Date(Date.now() - days * 864e5), 'America/Phoenix', 'yyyy-MM-dd') : '';
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var o = {};
    cols.forEach(function (c, j) { o[c] = TEXT_COLS.indexOf(c) >= 0 ? disp[i][j] : vals[i][j]; });
    if (!o.id) continue;
    if (cutoff && o.date < cutoff) continue;
    if ('active' in o) o.active = !(o.active === false || String(o.active).toUpperCase() === 'FALSE');
    if (t === 'costs') { o.hasPhoto = !!o.photoUrl; delete o.photoUrl; }
    rows.push(o);
  }
  return rows;
}
