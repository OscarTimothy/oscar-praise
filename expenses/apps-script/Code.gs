/**
 * Wedding expenses API for the GitHub Pages site in /expenses.
 *
 * Paste this into Extensions > Apps Script on the "Praise & Oscar Wedding
 * Expenses" spreadsheet, set the PASSWORD script property, and deploy it as a
 * web app (Execute as: Me, Who has access: Anyone). See ../README.md.
 *
 * The password lives only in Script Properties, never in the website code.
 */

const EXPENSES_TAB = 'Expenses';
const SUMMARY_TAB = 'Summary';
const TIMELINE_TAB = 'Timeline';
const HEADER_ROW = 3;
const FIRST_ROW = 4;
const LAST_COL = 13; // A..M

// Columns on the Expenses tab (1-based). J and K are formulas and are never written.
const COL = {
  date: 1, vendor: 2, item: 3, category: 4, event: 5, paidBy: 6, currency: 7,
  total: 8, paid: 9, paidUsd: 10, balance: 11, status: 12, notes: 13,
};

const MAX_FAILS = 10;          // wrong passwords allowed...
const FAIL_WINDOW_SEC = 600;   // ...within this many seconds before a lockout

function doGet() {
  return json({ ok: true, service: 'wedding-expenses' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json({ ok: false, error: 'bad_request', message: 'The request was not valid JSON.' });
  }

  const gate = checkPassword_(body.password);
  if (gate) return json(gate);

  try {
    switch (body.action) {
      case 'login':
        return json({ ok: true });
      case 'read':
        return json({ ok: true, data: readAll_() });
      case 'add':
        return json(withLock_(() => addExpense_(body.expense)));
      case 'update':
        return json(withLock_(() => updateExpense_(body.row, body.expect, body.expense)));
      default:
        return json({ ok: false, error: 'bad_request', message: 'Unknown action.' });
    }
  } catch (err) {
    return json({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}

/* ---------- password ---------- */

function checkPassword_(given) {
  const expected = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!expected) {
    return { ok: false, error: 'not_configured', message: 'Set the PASSWORD script property in Apps Script first.' };
  }
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) {
    return { ok: false, error: 'locked', message: 'Too many wrong passwords. Try again in 10 minutes.' };
  }
  if (typeof given !== 'string' || given !== expected) {
    cache.put('fails', String(fails + 1), FAIL_WINDOW_SEC);
    return { ok: false, error: 'wrong_password', message: 'That password is not right.' };
  }
  return null;
}

/* ---------- reading ---------- */

function readAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(EXPENSES_TAB);
  const last = sh.getLastRow();
  const rows = [];
  if (last >= FIRST_ROW) {
    const values = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, LAST_COL).getValues();
    values.forEach((r, i) => {
      const vendor = String(r[COL.vendor - 1] || '').trim();
      const item = String(r[COL.item - 1] || '').trim();
      if (!vendor && !item) return;
      const d = r[COL.date - 1];
      rows.push({
        row: FIRST_ROW + i,
        date: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || ''),
        vendor: vendor,
        item: item,
        category: String(r[COL.category - 1] || ''),
        event: String(r[COL.event - 1] || ''),
        paidBy: String(r[COL.paidBy - 1] || ''),
        currency: String(r[COL.currency - 1] || 'NGN').toUpperCase(),
        total: numOrNull_(r[COL.total - 1]),
        paid: numOrNull_(r[COL.paid - 1]),
        balance: numOrNull_(r[COL.balance - 1]),
        status: String(r[COL.status - 1] || ''),
        notes: String(r[COL.notes - 1] || ''),
      });
    });
  }

  const summary = labelValues_(ss.getSheetByName(SUMMARY_TAB));
  const timeline = labelValues_(ss.getSheetByName(TIMELINE_TAB));
  const wedding = timeline['wedding date'];

  return {
    title: ss.getName(),
    rows: rows,
    rate: numOrNull_(findStarting_(summary, 'exchange rate')),
    account: numOrNull_(findStarting_(summary, 'account balance (')),
    weddingDate: wedding instanceof Date ? Utilities.formatDate(wedding, tz, 'yyyy-MM-dd') : (wedding ? String(wedding) : ''),
    lists: dropdowns_(sh),
    readAt: new Date().toISOString(),
  };
}

// Reads a two-column "label | value" layout from columns B:C.
function labelValues_(sheet) {
  const out = {};
  if (!sheet) return out;
  const n = Math.min(sheet.getLastRow(), 80);
  if (n < 1) return out;
  sheet.getRange(1, 2, n, 2).getValues().forEach((r) => {
    const label = String(r[0] || '').trim().toLowerCase();
    if (label && r[1] !== '' && r[1] !== null && !(label in out)) out[label] = r[1];
  });
  return out;
}

function findStarting_(map, prefix) {
  const key = Object.keys(map).find((k) => k.indexOf(prefix) === 0);
  return key ? map[key] : null;
}

function dropdowns_(sh) {
  const read = (col) => {
    const rule = sh.getRange(FIRST_ROW, col).getDataValidation();
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return null;
    return rule.getCriteriaValues()[0].map(String);
  };
  return {
    category: read(COL.category),
    event: read(COL.event),
    paidBy: read(COL.paidBy),
    currency: read(COL.currency),
    status: read(COL.status),
  };
}

/* ---------- writing ---------- */

function addExpense_(input) {
  const clean = validate_(input);
  if (clean.error) return clean;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSES_TAB);
  const row = nextEmptyRow_(sh);
  writeRow_(sh, row, clean.value);
  return { ok: true, row: row };
}

function updateExpense_(row, expect, input) {
  row = Number(row);
  if (!Number.isInteger(row) || row < FIRST_ROW) return { ok: false, error: 'bad_request', message: 'That row number is not valid.' };
  const clean = validate_(input);
  if (clean.error) return clean;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPENSES_TAB);
  const current = sh.getRange(row, COL.vendor, 1, 2).getValues()[0];
  const now = { vendor: String(current[0] || '').trim(), item: String(current[1] || '').trim() };
  if (!expect || now.vendor !== String(expect.vendor || '') || now.item !== String(expect.item || '')) {
    return { ok: false, error: 'changed', message: 'This row changed in the sheet since you opened it. Reload and try again.' };
  }
  writeRow_(sh, row, clean.value);
  return { ok: true, row: row };
}

function writeRow_(sh, row, v) {
  // A:I in one write, L:M in another, so the J and K formulas are untouched.
  sh.getRange(row, COL.date, 1, 9).setValues([[
    v.date, v.vendor, v.item, v.category, v.event, v.paidBy, v.currency,
    v.total === null ? '' : v.total, v.paid === null ? '' : v.paid,
  ]]);
  sh.getRange(row, COL.status, 1, 2).setValues([[v.status, v.notes]]);
}

function nextEmptyRow_(sh) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return FIRST_ROW;
  const vals = sh.getRange(FIRST_ROW, COL.vendor, last - FIRST_ROW + 1, 2).getValues();
  let lastUsed = FIRST_ROW - 1;
  vals.forEach((r, i) => { if (String(r[0]).trim() || String(r[1]).trim()) lastUsed = FIRST_ROW + i; });
  return lastUsed + 1;
}

function validate_(x) {
  if (!x || typeof x !== 'object') return { ok: false, error: 'bad_request', message: 'No expense was sent.' };
  const text = (s, max) => String(s == null ? '' : s).trim().slice(0, max);
  const m = String(x.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const vendor = text(x.vendor, 120);
  const item = text(x.item, 200);
  const missing = [];
  if (!m) missing.push('a date');
  if (!vendor) missing.push('a vendor');
  if (!item) missing.push('an item');
  if (missing.length) return { ok: false, error: 'invalid', message: 'Add ' + missing.join(', ') + '.' };
  const amount = (a) => {
    if (a === null || a === undefined || a === '') return null;
    const n = Number(a);
    return isFinite(n) && n >= 0 && n < 1e12 ? n : NaN;
  };
  const total = amount(x.total);
  const paid = amount(x.paid);
  if (Number.isNaN(total) || Number.isNaN(paid)) return { ok: false, error: 'invalid', message: 'Amounts must be numbers of zero or more.' };
  const currency = text(x.currency, 3).toUpperCase() === 'USD' ? 'USD' : 'NGN';
  return {
    value: {
      date: new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])),
      vendor: vendor, item: item,
      category: text(x.category, 60), event: text(x.event, 60), paidBy: text(x.paidBy, 40),
      currency: currency, total: total, paid: paid,
      status: text(x.status, 40), notes: text(x.notes, 500),
    },
  };
}

/* ---------- helpers ---------- */

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'busy', message: 'Someone else is saving right now. Try again in a moment.' };
  try { return fn(); } finally { lock.releaseLock(); }
}

function numOrNull_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/[₦$,\s]/g, ''));
  return isFinite(n) ? n : null;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
