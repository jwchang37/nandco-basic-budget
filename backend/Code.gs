/**
 * NANDCO estimator backend (Google Apps Script web app, bound to a Google Sheet).
 *
 *   GET  ?action=settings            -> { ok, settings }                 (public)
 *   GET  ?action=list&key=ADMIN_KEY  -> { ok, rows, settings }           (admin)
 *   POST {action:'submit', name, email, url, project, total}            (public)
 *        -> appends a row to the "Requests" sheet and emails NOTIFY_EMAIL
 *   POST {action:'saveSettings', key, multiplierPct, gcFeePct}          (admin)
 *        -> { ok, settings }
 *
 * The admin password lives in Project Settings -> Script Properties as ADMIN_KEY.
 * It is never stored in the website's code. See SETUP.md.
 */

var NOTIFY_EMAIL = 'hello@nandcoconstruction.com';
var SHEET_NAME = 'Requests';
var DEFAULTS = { multiplierPct: 20, gcFeePct: 15 };
var HEADERS = ['Date', 'Name', 'Email', 'Project', 'Total', 'Estimate URL'];

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'settings') return json_({ ok: true, settings: getSettings_() });
    if (p.action === 'list') {
      if (!isAdmin_(p.key)) return json_({ ok: false, error: 'unauthorized' });
      return json_({ ok: true, rows: listRequests_(), settings: getSettings_() });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad request' }); }
  try {
    if (body.action === 'submit') return json_(submitRequest_(body));
    if (body.action === 'saveSettings') {
      if (!isAdmin_(body.key)) return json_({ ok: false, error: 'unauthorized' });
      return json_({ ok: true, settings: saveSettings_(body.multiplierPct, body.gcFeePct) });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// ── Requests ──
function submitRequest_(b) {
  var name = clean_(b.name, 120);
  var email = clean_(b.email, 200);
  var url = String(b.url || '').slice(0, 8000);
  var project = clean_(b.project, 300);
  var total = clean_(b.total, 40);
  if (!name) return { ok: false, error: 'name required' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'invalid email' };
  if (!/^https:\/\/[^\s]+$/i.test(url)) return { ok: false, error: 'invalid link' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    sh.appendRow([new Date(), cell_(name), cell_(email), cell_(project), cell_(total), cell_(url)]);
  } finally {
    lock.releaseLock();
  }

  var subject = 'Detailed estimate request — ' + name;
  var text = 'A customer requested a detailed estimate breakdown.\n\n'
    + 'Name: ' + name + '\nEmail: ' + email + '\nProject: ' + project + '\nEstimated total: ' + total
    + '\n\nOpen their estimate: ' + url + '\n';
  var html = '<p>A customer requested a detailed estimate breakdown.</p>'
    + '<p><b>Name:</b> ' + esc_(name) + '<br><b>Email:</b> ' + esc_(email)
    + '<br><b>Project:</b> ' + esc_(project) + '<br><b>Estimated total:</b> ' + esc_(total) + '</p>'
    + '<p><a href="' + esc_(url) + '">Open their estimate</a> (opens the estimator with their exact selections; '
    + 'click "Open printable report" for the PDF).</p>';
  MailApp.sendEmail({ to: NOTIFY_EMAIL, replyTo: email, subject: subject, body: text, htmlBody: html, name: 'NANDCO Estimator' });
  return { ok: true };
}

function listRequests_() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var start = Math.max(2, sh.getLastRow() - 499); // newest 500
  var values = sh.getRange(start, 1, sh.getLastRow() - start + 1, HEADERS.length).getValues();
  return values.reverse().map(function (r) {
    return {
      date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: uncell_(r[1]), email: uncell_(r[2]), project: uncell_(r[3]), total: uncell_(r[4]), url: uncell_(r[5])
    };
  });
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Settings ──
function getSettings_() {
  var props = PropertiesService.getScriptProperties();
  var m = parseFloat(props.getProperty('MULTIPLIER_PCT'));
  var g = parseFloat(props.getProperty('GC_FEE_PCT'));
  return {
    multiplierPct: isFinite(m) ? m : DEFAULTS.multiplierPct,
    gcFeePct: isFinite(g) ? g : DEFAULTS.gcFeePct
  };
}

function saveSettings_(m, g) {
  m = Number(m); g = Number(g);
  if (!isFinite(m) || m < 0 || m > 200) throw new Error('multiplier must be 0 to 200');
  if (!isFinite(g) || g < 0 || g > 100) throw new Error('GC fee must be 0 to 100');
  PropertiesService.getScriptProperties().setProperties({ MULTIPLIER_PCT: String(m), GC_FEE_PCT: String(g) });
  return getSettings_();
}

// ── Helpers ──
function isAdmin_(key) {
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!expected && typeof key === 'string' && key === expected;
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }
// Stop spreadsheet formula injection: a leading = + - @ is stored as text.
function cell_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function uncell_(v) { v = String(v == null ? '' : v); return v.charAt(0) === "'" ? v.slice(1) : v; }
function esc_(v) {
  return String(v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Run this once from the editor to grant the email + sheet permissions.
function authorize() {
  sheet_();
  getSettings_();
  MailApp.getRemainingDailyQuota();
}
