/**
 * J&J Obra — sync endpoint and database.
 * Lives inside the J&J data Google Sheet (Extensions > Apps Script).
 *
 * One-time setup:
 *   1. Run setup() once. It builds the four tabs and creates the device key.
 *   2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *   3. Run setupLink() and paste the web app URL when it asks. It logs the
 *      phone setup link for Jose.
 */

var TABS = {
  jobs:  { name: 'Jobs',  cols: ['id', 'name', 'active', 'createdAt', 'updatedAt'] },
  crew:  { name: 'Crew',  cols: ['id', 'name', 'active', 'createdAt', 'updatedAt'] },
  hours: { name: 'Hours', cols: ['id', 'date', 'jobId', 'jobName', 'workerId', 'workerName', 'hours', 'savedAt', 'device', 'syncedAt'] },
  costs: { name: 'Costs', cols: ['id', 'date', 'jobId', 'jobName', 'vendor', 'category', 'categoryLabel', 'amount', 'note', 'savedAt', 'device', 'syncedAt'] }
};
var TEXT_COLS = ['id', 'date', 'jobId', 'workerId', 'createdAt', 'updatedAt', 'savedAt', 'syncedAt'];
var CAT_LABEL = { material: 'Material', rental: 'Equipment rental', fuel: 'Fuel', tool: 'Tool',
  dumpster: 'Dumpster', permit: 'Permit', sub: 'Outside labor', other: 'Other' };
var RECENT_DAYS = 62;

/* ---------------- setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('J&J Obra')
    .addItem('Phone setup link', 'setupLink')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(function (t) {
    var def = TABS[t];
    var sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    def.cols.forEach(function (c, i) {
      if (TEXT_COLS.indexOf(c) >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('Setup complete. Now deploy as a Web app, then run setupLink().');
}

function setupLink() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Phone setup link', 'Paste the Web app URL (ends in /exec):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var exec = r.getResponseText().trim();
  var appUrl = PropertiesService.getScriptProperties().getProperty('APP_URL');
  if (!appUrl) {
    var a = ui.prompt('App address', 'Paste the app address (the GitHub Pages URL):', ui.ButtonSet.OK_CANCEL);
    if (a.getSelectedButton() !== ui.Button.OK) return;
    appUrl = a.getResponseText().trim().replace(/\/?$/, '/');
    PropertiesService.getScriptProperties().setProperty('APP_URL', appUrl);
  }
  var cfg = { u: exec, k: PropertiesService.getScriptProperties().getProperty('KEY'), n: 'Jose', l: 'es' };
  var s = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(cfg)).getBytes()).replace(/=+$/, '');
  var link = appUrl + '?s=' + s;
  Logger.log(link);
  ui.alert('Jose\'s setup link', link, ui.ButtonSet.OK);
}

/** Issue a new device key. Every phone must then be re-linked with a new setupLink(). */
function rotateKey() {
  PropertiesService.getScriptProperties().setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('Key rotated. Run setupLink() again for each phone.');
}

/* ---------------- web app ---------------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!authorized(p.k)) return out({ ok: false, err: 'auth' });
  return out({
    ok: true,
    jobs: readTab('jobs'),
    crew: readTab('crew'),
    hours: readTab('hours', RECENT_DAYS),
    costs: readTab('costs', RECENT_DAYS)
  });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, err: 'bad json' }); }
  if (!authorized(body.k)) return out({ ok: false, err: 'auth' });
  var ops = Array.isArray(body.ops) ? body.ops : [];
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var now = new Date().toISOString();
    ops.forEach(function (op) {
      if (!TABS[op.t] || !op.id) return;
      if (op.a === 'del') removeRow(op.t, String(op.id));
      else if (op.a === 'put' && op.d) upsertRow(op.t, String(op.id), clean(op.t, op.d, body.dev, now));
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return out({ ok: true, n: ops.length });
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
  if (t === 'hours') r.hours = Math.round(Number(d.hours) * 2) / 2;
  if (t === 'costs') {
    r.amount = Math.round(Number(d.amount) * 100) / 100;
    r.categoryLabel = CAT_LABEL[d.category] || d.category || '';
  }
  if (t === 'jobs' || t === 'crew') r.active = d.active === false ? false : true;
  if ('device' in r) r.device = dev || '';
  if ('syncedAt' in r) r.syncedAt = now;
  return r;
}

function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
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
    cols.forEach(function (c, j) { o[c] = TEXT_COLS.indexOf(c) >= 0 || c === 'name' || c === 'vendor' || c === 'note' ? disp[i][j] : vals[i][j]; });
    if (!o.id) continue;
    if (cutoff && o.date < cutoff) continue;
    if ('active' in o) o.active = !(o.active === false || String(o.active).toUpperCase() === 'FALSE');
    rows.push(o);
  }
  return rows;
}
