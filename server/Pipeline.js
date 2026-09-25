/**
 * PEOPLE PIPELINE - turns the sheets of "people pipeline.xlsx" into the data the page uses.
 *
 * Plain JavaScript without any Google service, so the IIS website can use this exact file too.
 * Input: a list of { name, rows } (rows = arrays of cell values, first useful row = the column names).
 *
 * Every sheet is recognised by its COLUMN NAMES (sheet names / order don't matter) and linked on accountid:
 *   accounts   the sheet with accountid + name + classification ("accounts detailed")
 *   projects   has opportunityid                                  → list of projects per account
 *   contacts   has contactid / contact (or is called "contact …") → list of contact persons per account
 *   social     has a *platform* column                            → list of platforms per account
 *   merge      everything else (activity, account proj details, validation …): ONE row per account.
 *              If an account has several rows they are combined (see combineRows_), never dropped.
 */

/** "Phase_Group" → "phasegroup": column names are compared like this everywhere. */
function hKey_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

/** Account ids from different queries can differ in case or have {braces}: "{ab-12}" → "AB-12". */
function normId_(v) {
  return String(v == null ? '' : v)
    .trim()
    .replace(/^\{|\}$/g, '')
    .toUpperCase();
}

/** Excel date number for a cell that holds a date as text ("2026-07-10", "10/07/2026", "45848"). */
function toSerial_(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).trim();
  if (!s || s === 'NULL') return null;
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/); // 2026-07-10 (14:30)
  if (!m) {
    const d = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})(?:[ T](\d{1,2}):(\d{2}))?/); // 10/07/2026 = day/month/year
    if (d) m = [d[0], d[3], d[2], d[1], d[4], d[5]];
  }
  if (!m) return null;
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  return isNaN(ms) ? null : ms / 864e5 + 25569;
}

/** Columns that only exist once per account: a sheet with any of them is merged into the account. */
const ACCOUNT_LEVEL_COLS = [
  'lastcontact1date', 'lastmeetingdate', 'lastmeetingindate', 'lastmeetingoutdate', 'metin',
  'prep1count', 'postp1count', 'postp1amount', 'lastprojectdatein', 'phasegroup', 'pipeline', 'finaltype',
  'validationtaskenddate'
];

/** What a sheet is, from its name and column names (hk = column names as hKey_). */
function sheetRole_(name, hk) {
  const has = function (c) {
    return hk.indexOf(c) >= 0;
  };
  const n = hKey_(name);
  if (has('opportunityid')) return 'projects';
  if (has('contactid') || has('contact') || has('contactperson') || /^contact/.test(n) || /query5$/.test(n)) return 'contacts';
  if (hk.some(function (h) { return /platform/.test(h); })) return 'social'; // prettier-ignore
  return 'merge'; // decided per account later: several rows → combined, or kept as a list if nothing is account-level
}

/** The row of the column names: the first of the top 10 rows that has an "accountid" cell (else the first row). */
function headerRow_(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    if ((rows[i] || []).map(hKey_).indexOf('accountid') >= 0) return i;
  }
  return 0;
}

/** Big date number = later date; "P3" > "P1"; for pipelines the LOWEST number wins (Pipeline 1 = strongest). */
function phaseRank_(v) {
  const m = String(v == null ? '' : v).match(/^\s*p?\s*([0-5])\b/i);
  return m ? +m[1] : -1;
}

/**
 * Several rows of an account in a one-row-per-account sheet → one row:
 *  - validation sheets: the whole row with the latest validation_task_enddate (owner stays with its date)
 *  - dates: the latest · numbers (counts, amounts): the highest · Phase_Group: the most advanced
 *  - pipeline: the lowest number · "met in": yes if any row says yes · other text: the first filled in
 */
function combineRows_(rows, hk) {
  const vk = hk.indexOf('validationtaskenddate');
  if (vk >= 0) {
    return rows.slice().sort(function (a, b) {
      return (toSerial_(b[vk]) || 0) - (toSerial_(a[vk]) || 0);
    })[0];
  }
  return hk.map(function (h, j) {
    const vals = rows
      .map(function (r) {
        return r[j];
      })
      .filter(function (v) {
        return v !== '' && v != null && v !== 'NULL';
      });
    if (!vals.length) return '';
    if (/date|creation/.test(h)) {
      return vals.reduce(function (best, v) {
        return (toSerial_(v) || 0) > (toSerial_(best) || 0) ? v : best;
      });
    }
    if (h === 'phasegroup') {
      return vals.reduce(function (best, v) {
        return phaseRank_(v) > phaseRank_(best) ? v : best;
      });
    }
    if (h === 'pipeline') {
      return vals.reduce(function (best, v) {
        const n = parseInt(String(v).replace(/\D/g, ''), 10) || 9,
          b = parseInt(String(best).replace(/\D/g, ''), 10) || 9;
        return n < b ? v : best;
      });
    }
    if (vals.every(function (v) { return /^(yes|no)$/i.test(String(v).trim()); })) { // prettier-ignore
      return vals.some(function (v) { return /^yes$/i.test(String(v).trim()); }) ? 'Yes' : 'No'; // prettier-ignore
    }
    if (vals.every(function (v) { return typeof v === 'number'; })) return Math.max.apply(null, vals); // prettier-ignore
    return vals[0];
  });
}

/**
 * All sheets → { source, accounts, sheets, checks }.
 * checks = how every sheet was read (shown in the hidden "Data check" screen).
 */
function buildFromSheets_(sheetList, source) {
  const checks = [];
  const sheets = [];
  sheetList.forEach(function (s) {
    const rows = (s.rows || []).filter(function (r) {
      return r && r.some(function (v) { return v !== '' && v != null; }); // prettier-ignore
    });
    if (!rows.length) return checks.push({ sheet: s.name, role: 'skipped', note: 'empty sheet' });
    const h = headerRow_(rows);
    const header = rows[h].map(function (v) {
      return String(v == null ? '' : v).trim();
    });
    sheets.push({ name: s.name, header: header, hk: header.map(hKey_), rows: rows.slice(h + 1) });
  });

  // 1. the accounts sheet
  const hasCols = function (s, cols) {
    return cols.every(function (c) {
      return s.hk.indexOf(c) >= 0;
    });
  };
  const master =
    sheets.filter(function (s) { return hasCols(s, ['accountid', 'name', 'classification']); })[0] || // prettier-ignore
    sheets.filter(function (s) { return hasCols(s, ['accountid', 'name']); })[0]; // prettier-ignore
  if (!master) throw new Error('No sheet with the columns "accountid" and "name" found in ' + source.name);

  const mk = master.hk.indexOf('accountid');
  const accIndex = {};
  const mRows = [];
  let mDup = 0,
    mNoId = 0;
  master.rows.forEach(function (r) {
    const id = normId_(r[mk]);
    if (!id) return mNoId++;
    if (accIndex[id] !== undefined) return mDup++;
    accIndex[id] = mRows.length;
    mRows.push(r);
  });
  const out = [packSheet_(master.name, master.header, mRows, null, 'accounts')];
  checks.push({
    sheet: master.name, role: 'accounts', columns: master.header, rows: master.rows.length, linked: mRows.length,
    note: (mDup ? mDup + ' duplicate account rows ignored. ' : '') + (mNoId ? mNoId + ' rows without accountid.' : '')
  }); // prettier-ignore

  // 2. every other sheet with an accountid column
  sheets.forEach(function (s) {
    if (s === master) return;
    const k = s.hk.indexOf('accountid');
    if (k < 0) return checks.push({ sheet: s.name, role: 'skipped', columns: s.header, note: 'no accountid column' });
    let role = sheetRole_(s.name, s.hk);
    const byAcc = {},
      seen = {},
      order = [],
      unmatched = [];
    let empty = 0,
      dups = 0;
    s.rows.forEach(function (r) {
      const a = accIndex[normId_(r[k])];
      if (a === undefined) return unmatched.push(normId_(r[k]));
      if (!r.some(function (v, j) { return j !== k && v !== '' && v != null && v !== 'NULL'; })) return empty++; // prettier-ignore
      const sig = r.join('\u0001'); // the same row twice (joins in the query)
      if (seen[sig]) return dups++;
      seen[sig] = 1;
      if (!byAcc[a]) {
        byAcc[a] = [];
        order.push(a);
      }
      byAcc[a].push(r);
    });
    const multi = order.filter(function (a) {
      return byAcc[a].length > 1;
    }).length;
    let note = '';
    if (role === 'merge' && multi) {
      const accountLevel = s.hk.some(function (h) {
        return ACCOUNT_LEVEL_COLS.indexOf(h) >= 0;
      });
      if (accountLevel) note = multi + ' accounts had several rows: combined into one (latest dates, highest numbers).';
      else role = 'list'; // a real list (several things per account): shown as "Other linked rows"
    }
    const rows = [],
      link = [];
    order.forEach(function (a) {
      if (role === 'merge') {
        rows.push(byAcc[a].length > 1 ? combineRows_(byAcc[a], s.hk) : byAcc[a][0]);
        link.push(a);
      } else
        byAcc[a].forEach(function (r) {
          rows.push(r);
          link.push(a);
        });
    });
    const drop = function (r) {
      return r.filter(function (v, j) {
        return j !== k;
      });
    }; // accountid travels as "link"
    out.push(packSheet_(s.name, drop(s.header), rows.map(drop), link, role));
    const noMatch = unmatched.filter(Boolean);
    checks.push({
      sheet: s.name, role: role, columns: s.header, rows: s.rows.length, linked: rows.length, accounts: order.length,
      unmatched: noMatch.length, unmatchedSample: noMatch.slice(0, 3), empty: empty, duplicates: dups, note: note
    }); // prettier-ignore
  });

  const data = { source: source, accounts: mRows.length, sheets: out, checks: checks };
  data.source.hash = dataHash_(out);
  return data;
}

/** Column-oriented, dictionary-encoded sheet (keeps the payload small). Date columns become Excel date numbers. */
function packSheet_(name, header, rows, link, role) {
  const cols = header.map(function (h, j) {
    const isDate = /date|creation/.test(hKey_(h));
    const vals = new Array(rows.length);
    let isNum = true;
    for (let i = 0; i < rows.length; i++) {
      let v = rows[i][j];
      if (v === undefined || v === '' || v === 'NULL' || v === 'null') v = null;
      if (isDate && v !== null && typeof v !== 'number') {
        const d = toSerial_(v);
        if (d !== null) v = d;
      }
      if (typeof v === 'number') v = Math.round(v * 100000) / 100000;
      else if (v !== null) {
        isNum = false;
        v = String(v);
      }
      vals[i] = v;
    }
    const col = { h: String(h).trim() };
    if (!isNum) {
      for (let i = 0; i < vals.length; i++) if (vals[i] !== null) vals[i] = String(vals[i]); // text column: all text
      const dict = [],
        di = {};
      vals.forEach(function (v) {
        if (v !== null && di[v] === undefined) {
          di[v] = dict.length;
          dict.push(v);
        }
      });
      if (dict.length < rows.length * 0.5) {
        col.d = dict;
        col.v = vals.map(function (v) {
          return v === null ? -1 : di[v];
        });
        return col;
      }
    }
    col.v = vals;
    return col;
  });
  return { name: name, role: role, n: rows.length, link: link, cols: cols };
}

/** A short fingerprint of the data: the same file gives the same text (used to tell "no changes"). */
function dataHash_(obj) {
  const s = JSON.stringify(obj);
  let h1 = 0x811c9dc5,
    h2 = 0x01000193;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 16777619) >>> 0;
    h2 = Math.imul(h2 ^ c, 2246822519) >>> 0;
  }
  return h1.toString(16) + h2.toString(16) + ':' + s.length;
}
