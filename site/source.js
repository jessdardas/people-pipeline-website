/*
 * PEOPLE PIPELINE - the website's data connection (replaces Google Apps Script).
 * - the data:     data/people pipeline.xlsx on this IIS server, read in the browser with SheetJS and
 *                 turned into the page's data by server/Pipeline.js (the same code as in Apps Script)
 * - the settings: api/settings.ashx (saves the New DB date for everyone, checks the admin password);
 *                 if that handler is missing, data/settings.json is only read
 * Loaded after the app's code and before it starts, so the api() below replaces the Apps Script one.
 */
var SITE = {
  file: 'data/people pipeline.xlsx', // the Excel file (the copy task / export puts it here)
  settingsApi: 'api/settings.ashx',
  settingsFile: 'data/settings.json',
  last: null // hash of the last data read
};

// exports use the copies of the libraries inside the website (works without internet)
XLSX_URL = 'site/vendor/xlsx-style.bundle.js';
JSPDF_URL = 'site/vendor/jspdf.umd.min.js';
AUTOTABLE_URL = 'site/vendor/jspdf.plugin.autotable.min.js';

function api(fn, args) {
  args = args || [];
  if (fn === 'getData') return siteData();
  if (fn === 'getStamp') return siteStamp();
  if (fn === 'saveSettings') return sitePost({ password: args[0], newDbFrom: (args[1] || {}).newDbFrom });
  if (fn === 'checkSource')
    return sitePost({ password: args[0], check: true }).then(function () {
      return siteRead().then(function (d) {
        return { source: d.source, accounts: d.accounts, checks: d.checks };
      });
    });
  return Promise.reject(new Error('Unknown call ' + fn));
}

/** Reads the Excel file now (never from the browser cache). */
function siteRead() {
  return fetch(encodeURI(SITE.file) + '?t=' + Date.now(), { cache: 'no-store' }).then(function (r) {
    if (!r.ok) throw new Error('The Excel file "' + SITE.file + '" was not found on the server (' + r.status + ')');
    var lm = Date.parse(r.headers.get('Last-Modified') || '') || Date.now();
    return r.arrayBuffer().then(function (buf) {
      var wb = XLSX.read(buf, { type: 'array' });
      var sheets = wb.SheetNames.map(function (n) {
        return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) };
      });
      return buildFromSheets_(sheets, { name: 'people pipeline.xlsx', id: SITE.file, updated: lm, loaded: Date.now() });
    });
  });
}

function siteData() {
  return Promise.all([siteRead(), siteSettings()]).then(function (r) {
    var d = r[0];
    d.source.changed = SITE.last !== null && SITE.last !== d.source.hash;
    SITE.last = d.source.hash;
    delete d.checks;
    d.settings = r[1];
    return d;
  });
}

function siteStamp() {
  return fetch(encodeURI(SITE.file) + '?t=' + Date.now(), { method: 'HEAD', cache: 'no-store' }).then(function (r) {
    return { updated: Date.parse(r.headers.get('Last-Modified') || '') || 0 };
  });
}

function siteSettings() {
  var fallback = function () {
    return fetch(SITE.settingsFile + '?t=' + Date.now(), { cache: 'no-store' })
      .then(function (r) {
        return r.ok ? r.json() : {};
      })
      .catch(function () {
        return {};
      });
  };
  return fetch(SITE.settingsApi + '?t=' + Date.now(), { cache: 'no-store' })
    .then(function (r) {
      return r.ok ? r.json() : fallback();
    })
    .catch(fallback)
    .then(function (s) {
      return Object.assign({ newDbFrom: '2026-07-01' }, s);
    });
}

/** Sends the password (+ new date) to the IIS handler; it checks the password and saves for everyone. */
function sitePost(body) {
  return fetch(SITE.settingsApi, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }).then(function (r) {
    return r.text().then(function (t) {
      var j = {};
      try {
        j = JSON.parse(t);
      } catch (e) {}
      if (r.status === 404)
        throw new Error('The admin handler api/settings.ashx is not running on this server (see README: ASP.NET).');
      if (!r.ok) throw new Error(j.error || 'Server error ' + r.status);
      return j;
    });
  });
}
