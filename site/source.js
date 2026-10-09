/*
 * PEOPLE PIPELINE - the website's data connection (replaces Google Apps Script).
 * - the data:     data/people pipeline.xlsx on this IIS server, read in the browser with SheetJS and
 *                 turned into the page's data by server/Pipeline.js (the same code as in Apps Script)
 * - "Last updated at": the LAST date written in data/refresh-log.txt (C:\inetpub\people-pipeline\data\refresh-log.txt);
 *                 without that file, the time the Excel file last changed
 * - the settings: api/settings.ashx (saves the New DB date for everyone, checks the admin password);
 *                 if that handler is missing, data/settings.json is only read
 * Loaded after the app's code and before it starts, so the api() below replaces the Apps Script one.
 */
var SITE = {
  file: 'data/people pipeline.xlsx', // the Excel file (the copy task / export puts it here)
  refreshFile: 'data/refresh-log.txt', // the refresh log: its last date is shown as "Last updated at", as written
  settingsApi: 'api/settings.ashx',
  settingsFile: 'data/settings.json',
  last: null, // hash of the last data read
  // shown very small at the bottom left of the page, so you can check that the server runs the newest version
  // (change it with every update: the date + a word about what changed)
  version: 'version 2026-10-09c · current demand, map pairs',
  // Faster opening: the data read from the Excel file is kept in this browser (IndexedDB) together with the
  // file's date + size. Next time, if the file on the server has the same date + size, the kept copy is used
  // and the Excel file is not downloaded or decoded again. Change cacheVersion when server/Pipeline.js changes
  // how the file is read, so every browser reads the file again once.
  cacheVersion: '2026-10-02b'
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

/** The Excel file's "stamp" on the server: date + size (+ ETag), without downloading it. */
function siteFileStamp() {
  return fetch(encodeURI(SITE.file) + '?t=' + Date.now(), { method: 'HEAD', cache: 'no-store' }).then(function (r) {
    if (!r.ok) throw new Error('The Excel file "' + SITE.file + '" was not found on the server (' + r.status + ')');
    return [SITE.cacheVersion, r.headers.get('Last-Modified') || '', r.headers.get('Content-Length') || '', r.headers.get('ETag') || ''].join('|');
  });
}

/** Loads the Excel reader (SheetJS, ~900 KB) the first time it is needed - not on every opening. */
function siteXlsx() {
  if (window.XLSX && XLSX.read) return Promise.resolve();
  if (!SITE._xlsx)
    SITE._xlsx = new Promise(function (ok, fail) {
      var el = document.createElement('script');
      el.src = 'site/vendor/xlsx.full.min.js';
      el.onload = ok;
      el.onerror = function () {
        SITE._xlsx = null;
        fail(new Error('Could not load site/vendor/xlsx.full.min.js'));
      };
      document.head.appendChild(el);
    });
  return SITE._xlsx;
}

/** Downloads the Excel file (never from the browser's own HTTP cache). */
function siteDownload() {
  return fetch(encodeURI(SITE.file) + '?t=' + Date.now(), { cache: 'no-store' }).then(function (r) {
    if (!r.ok) throw new Error('The Excel file "' + SITE.file + '" was not found on the server (' + r.status + ')');
    var lm = Date.parse(r.headers.get('Last-Modified') || '') || Date.now();
    return r.arrayBuffer().then(function (buf) {
      return { buf: buf, lm: lm };
    });
  });
}

/** Reads the Excel file now: download + Excel reader in parallel, then decode and turn into the page's data. */
function siteRead() {
  return Promise.all([siteXlsx(), siteDownload()]).then(function (r) {
    var got = r[1];
    // only the cell values are needed: skipping formats, styles and formulas makes decoding ~40% faster
    var wb = XLSX.read(got.buf, {
      type: 'array',
      dense: true,
      cellText: false,
      cellHTML: false,
      cellFormula: false,
      cellStyles: false,
      cellNF: false,
      cellDates: false
    });
    var sheets = wb.SheetNames.map(function (n) {
      return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) };
    });
    return buildFromSheets_(sheets, { name: 'people pipeline.xlsx', id: SITE.file, updated: got.lm, loaded: Date.now() });
  });
}

/* ---- the copy kept in this browser (IndexedDB). Every step may fail (private window, blocked storage):
        then the file is simply read as normal. ---- */
function siteDb() {
  return new Promise(function (ok, fail) {
    var rq = indexedDB.open('people-pipeline', 1);
    rq.onupgradeneeded = function () {
      rq.result.createObjectStore('data');
    };
    rq.onsuccess = function () {
      ok(rq.result);
    };
    rq.onerror = function () {
      fail(rq.error);
    };
  });
}

function siteCacheGet() {
  return siteDb()
    .then(function (db) {
      return new Promise(function (ok) {
        var rq = db.transaction('data').objectStore('data').get('last');
        rq.onsuccess = function () {
          ok(rq.result || null);
        };
        rq.onerror = function () {
          ok(null);
        };
      });
    })
    .catch(function () {
      return null;
    });
}

function siteCachePut(stamp, d) {
  return siteDb()
    .then(function (db) {
      db.transaction('data', 'readwrite').objectStore('data').put({ stamp: stamp, d: d }, 'last');
    })
    .catch(function () {});
}

/** The data: the kept copy when the file did not change since it was kept, else the file read now. */
function siteReadFast() {
  return Promise.all([siteFileStamp(), siteCacheGet()]).then(function (r) {
    var stamp = r[0],
      kept = r[1];
    if (kept && kept.stamp === stamp && kept.d) {
      kept.d.source.loaded = Date.now();
      return kept.d;
    }
    return siteRead().then(function (d) {
      siteCachePut(stamp, d);
      return d;
    });
  });
}

// a date (+ time) as people / scripts write it: 28/09/2026 10:00, 2026-09-28T10:00:05, 28-09-2026 10:00 AM, 28 Sep 2026 10:00
var LOG_DATE = /\d{4}-\d{1,2}-\d{1,2}(?:[ T]+\d{1,2}:\d{2}(?::\d{2})?)?|\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}(?:[ ,]+\d{1,2}:\d{2}(?::\d{2})?(?:\s?[AaPp][Mm])?)?|\d{1,2} [A-Za-z]{3,9}\.? \d{4}(?:[ ,]+\d{1,2}:\d{2}(?::\d{2})?)?/g;

/** "Last updated at": the newest SUCCESSFUL refresh in the refresh log ('' when there is none).
 *  Only lines with a date AND the word OK count (e.g. "2026-10-08 09:00:12 OK", "08/10/2026 09:00 - OK").
 *  A line with ERROR is ignored completely, even if it also says OK - error times are never shown. */
var LOG_OK = /\bOK\b/i;
var LOG_ERROR = /\bERRORS?\b|\bFAIL(ED|URE)?\b/i;

function siteRefreshed() {
  return fetch(encodeURI(SITE.refreshFile) + '?t=' + Date.now(), { cache: 'no-store' })
    .then(function (r) {
      return r.ok ? r.arrayBuffer() : null;
    })
    .then(function (buf) {
      if (!buf) return '';
      var b = new Uint8Array(buf),
        enc = b[0] === 0xff && b[1] === 0xfe ? 'utf-16le' : b[0] === 0xfe && b[1] === 0xff ? 'utf-16be' : 'utf-8'; // Notepad "Unicode" = UTF-16
      var lines = new TextDecoder(enc)
        .decode(buf)
        .replace(/^\uFEFF/, '')
        .split(/\r?\n/)
        .map(function (x) {
          return x.trim();
        })
        .filter(Boolean);
      // every line with a date that says OK (and no ERROR); the newest time wins (toSerial_ in server/Pipeline.js
      // reads 2026-10-08 09:00, 08/10/2026 09:00, 10/8/2026 9:00 AM …); unreadable dates count by their position
      var best = null;
      lines.forEach(function (line, i) {
        if (!LOG_OK.test(line) || LOG_ERROR.test(line)) return;
        var m = line.match(LOG_DATE);
        if (!m) return;
        var raw = m[m.length - 1].trim(),
          t = toSerial_(raw);
        if (!best || (t != null && (best.t == null || t >= best.t)) || (t == null && best.t == null)) best = { raw: raw, t: t };
      });
      if (!best) return '';
      if (best.t == null) return best.raw;
      var d = new Date(Math.round((best.t - 25569) * 864e5)); // Excel date number → time (UTC = as written)
      return (
        d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }) +
        ' at ' +
        d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', timeZone: 'UTC' })
      );
      return '';
    })
    .catch(function () {
      return '';
    });
}

function siteData() {
  return Promise.all([siteReadFast(), siteSettings(), siteRefreshed()]).then(function (r) {
    var d = r[0];
    d.source.refreshed = r[2];
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

/* the version label (SITE.version) at the bottom left of the page */
document.addEventListener('DOMContentLoaded', function () {
  var v = document.createElement('div');
  v.className = 'site-version';
  v.textContent = SITE.version;
  v.style.cssText = 'position:fixed;left:8px;bottom:4px;font-size:10px;color:#b5b5b5;z-index:1;pointer-events:none';
  document.body.appendChild(v);
});
