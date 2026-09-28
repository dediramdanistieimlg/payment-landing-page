/**
 * ============================================================
 * PaymentShopeeProductHunter.gs — v2.2
 * Google Apps Script bound to the "Payment" Google Sheet.
 * Counterpart to worker-payment-lynk-final.js (Cloudflare Worker).
 * ============================================================
 *
 * NOTE ON THIS VERSION: this file was rewritten from scratch to match the
 * exact HTTP contract the Worker already calls (see worker-payment-lynk-final.js
 * — reconcilePaymentFromSpreadsheet / writeLicenseToPaymentSheet / callAppsScript)
 * and the exact "Payment" sheet layout shown in the reference screenshot, since
 * no prior version of this file was available to edit in place. If your
 * deployed script has extra logic beyond what's described below (a different
 * email template, extra actions, etc.), port that over before replacing your
 * live deployment — everything here is written to be a safe drop-in for the
 * lookup_payment / fulfill_license contract, not a guess at unseen code.
 *
 * v2.0 CHANGES:
 *   - FIX (column misalignment — G "Lisensi Code" / H "Email Terkirim" landing
 *     on the wrong row, or orphaned with no matching A:F data, as seen in the
 *     reference screenshot): columns A:F on the Payment sheet are the output of
 *     a live array formula —
 *       =VSTACK(
 *         {"Judul Barang","Harga","Tanggal","Status","Buyer Email","Ref"};
 *         SORT(CHOOSECOLS(FILTER(Raw_all!A2:Z; Raw_all!A2:A<>""; Raw_all!P2:P="SUCCESS"); 1;3;15;16;17;26); 3; TRUE)
 *       )
 *     — so every time Raw_all gets a new row, ALL of A:F re-sorts and every
 *     row's position can shift. The previous approach (apparently) wrote the
 *     license code directly into a specific G/H cell once, keyed by row
 *     number — the instant the sheet re-sorted, that row number no longer
 *     matched the same Ref, leaving the license code stranded on whatever row
 *     happened to be there afterwards (exactly the orphaned rows visible in
 *     the screenshot).
 *     FIX: license fulfillment records are no longer written directly into
 *     Payment!G/H at all. They're appended (append-only, never reordered) to
 *     a separate "Fulfillment" sheet keyed by the stable Ref value. Payment!G
 *     and Payment!H are then set to a single ARRAYFORMULA (installed once by
 *     ensurePaymentColumnFormulas_) that VLOOKUPs each row's own F (Ref)
 *     against the Fulfillment sheet. Because that lookup is keyed by Ref and
 *     recalculates automatically along with A:F, G/H can never drift out of
 *     alignment again — no matter how many times the sheet re-sorts.
 *     A one-time migrateOrphanedLicenseColumns_() is included to fold any
 *     already-correct existing G/H values into the new Fulfillment sheet, and
 *     to flag any already-orphaned ones (blank F on that row) for manual
 *     review instead of silently discarding them.
 *   - FEATURE: checkPendingPaymentsAndReport_() — a time-driven trigger
 *     (installPaymentCheckerTrigger(), ~every 1 minute) that scans Payment
 *     for SUCCESS rows with no matching Fulfillment entry yet, and if any are
 *     found, pings the Worker's POST /api/v1/payment/reconcile-ping so it
 *     reconciles immediately instead of waiting for its own 5-minute cron.
 *     This script deliberately does NOT do the Supabase matching/fulfillment
 *     itself — it only detects + notifies — so the fulfillment business logic
 *     stays in one place (the Worker) instead of being duplicated here.
 *
 * v2.1 CHANGES (this pass):
 *   - CONFIG: script is now explicitly bound to a hardcoded SPREADSHEET_ID
 *     (getSpreadsheet_()) instead of relying only on
 *     SpreadsheetApp.getActiveSpreadsheet(). Behavior is identical when run
 *     as a normal container-bound script, but this removes any ambiguity if
 *     the project is ever duplicated, run as a standalone script, or the
 *     binding is otherwise unclear — it always operates on the intended
 *     spreadsheet or fails loudly, instead of silently working on whatever
 *     happens to be "active".
 *   - FEATURE: checkPaymentMatrix_() — a validation pass over Payment!A:F
 *     (Judul Barang, Harga, Tanggal, Status, Buyer Email, Ref) that flags
 *     missing fields, unexpected Status values, unparseable dates, unusual
 *     Ref formats, invalid emails, duplicate line items, and rows that are
 *     out of chronological order — plus a fulfilled/unfulfilled cross-check
 *     against the Fulfillment sheet. Reachable three ways: the "Payment
 *     Tools" sheet menu, a new doPost action check_payment_matrix, or
 *     doGet with ?secret=...&matrix=1 for quick remote diagnostics.
 *   - ORDERING GUARANTEE: Payment!G/H are matched to each row by Ref (via
 *     VLOOKUP), which was already immune to re-sorting since v2.0 — that
 *     part doesn't depend on row order at all. What v2.1 adds is that every
 *     place this script itself has to pick an order (which pending payment
 *     to report/process first, which of several ambiguous candidates is
 *     "the" match) now explicitly sorts by Tanggal — the real lynk.id
 *     payment-success timestamp — ascending, instead of assuming the
 *     Payment sheet's own SORT() formula direction. That way licenses are
 *     always reported/issued in true payment order even if that formula is
 *     ever edited.
 *   - HARDENING:
 *       • api_secret comparison is now constant-time (safeCompare_) instead
 *         of `!==`, and repeated failed auth attempts are throttled
 *         (CacheService counter + short forced delay past a threshold) to
 *         slow down brute-force guesses against the /exec endpoint.
 *       • doGet no longer leaks the spreadsheet ID / sheet layout to
 *         unauthenticated callers — it returns a bare "ok" unless called
 *         with the correct ?secret=.
 *       • All string inputs from the request body are trimmed and capped in
 *         length (clean_ / MAX_FIELD_LENGTH) before being written to a
 *         sheet cell or used in an email, and email subject text is
 *         stripped of CR/LF (sanitizeForEmailHeader_) to prevent header
 *         injection.
 *       • handleFulfillLicense_ now takes a script lock (LockService) around
 *         its read-check-then-append critical section, so two near-
 *         simultaneous fulfill_license calls (e.g. a Worker retry racing the
 *         original request) can't both pass the "not yet fulfilled" check
 *         and append two Fulfillment rows for the same Ref.
 *       • The Worker reconcile-ping now retries once on failure/non-2xx
 *         instead of giving up after a single UrlFetchApp call.
 *       • sendLicenseEmail_ checks MailApp's remaining daily quota first and
 *         fails loudly (caught by the existing try/catch) instead of
 *         silently hitting a quota error mid-send.
 *       • doPost validates that a body was actually sent and is a JSON
 *         object (not an array/primitive) before touching it.
 *
 * v2.2 CHANGES (this pass) — root-cause pass for "license delivered by
 * Telegram + email, but Payment!G / H stay empty":
 *   - DIAGNOSABILITY: every doPost response now carries gs_version, and
 *     fulfill_license additionally returns sheet_verified /
 *     sheet_verify_reason / sheet_rows. The Worker can therefore tell
 *     (a) whether it is really talking to THIS version of the script (an
 *     older pinned deployment, or a duplicate doPost in another .gs file,
 *     answers without gs_version) and (b) whether the license is actually
 *     visible on Payment!G:H — and page the admin bot when it is not.
 *     Opening the /exec URL in a browser now shows the deployed version too.
 *   - VERIFY-AFTER-WRITE: handleFulfillLicense_ re-reads Payment!G after
 *     writing to Fulfillment instead of assuming the formula picked it up.
 *   - FORMULAS: Payment!G/H headers are plain text again and the lookup is a
 *     single ARRAYFORMULA in row 2, e.g.
 *       =ARRAYFORMULA(IFERROR(VLOOKUP(TRIM(F2:F);Fulfillment!$A:$B;2;FALSE)))
 *     (";" shown for id-ID locale). It contains no "" literal: on a
 *     ";"-separator spreadsheet Sheets only half-translated the commas of a
 *     formula with an empty-string argument, giving #ERROR!. The installer
 *     now verifies the cell after writing it, retries with ";" separators,
 *     and fails loudly instead of leaving an error in G/H. The key is
 *     TRIM()med so a stray space in Payment!F can't break the match either.
 *   - FULFILLMENT SHEET: Ref (col A) and timestamp (col C) are forced to
 *     plain text so a hex Ref like "12345e6789" or a "28-09-2026 12:21"
 *     stamp can never be auto-converted to a number/date by Sheets.
 *   - EMAIL: license email rebuilt as a clean HTML layout (Paket / License Code
 *     table, numbered activation steps, "Penting" notes, support link) with a
 *     plain-text fallback, sender name "Shopee Product Hunter", and pure-ASCII
 *     content — no emoji, so no more "������" in mail clients.
 *   - LOOKUP: product titles are compared with all whitespace collapsed (a
 *     lynk.id title with a double space, "Paket Pro  Shopee Product Hunter",
 *     never matched the Worker's "Paket Pro Shopee Product Hunter"), several
 *     rows sharing one Ref no longer count as "ambiguous", and a failed lookup
 *     now returns a specific reason (email_not_found / title_mismatch /
 *     paid_before_order_created) plus a small diag object instead of a bare
 *     "no_match".
 *   - MATRIX: new issue type fulfilled_but_not_visible_in_G, plus a
 *     fulfillmentRows counter, so "Cek Matrix Pembayaran (A:F)" immediately
 *     separates "never written to Fulfillment" from "written but not shown".
 * ============================================================
 * ONE-TIME SETUP (run these from the Apps Script editor, each once):
 *   1. setup()                              — creates the Fulfillment sheet,
 *                                              installs the G/H formulas, and
 *                                              installs the 1-minute trigger.
 *   2. migrateOrphanedLicenseColumns_()      — OPTIONAL, run once if you have
 *                                              existing plain-value G/H data
 *                                              from before this version.
 * Then set these Script Properties (Project Settings → Script Properties):
 *   - API_SECRET                  — must match PAYMENT_APPS_SCRIPT_SECRET in the Worker
 *   - WORKER_RECONCILE_PING_URL   — e.g. https://<your-worker>.workers.dev/api/v1/payment/reconcile-ping
 *   - WORKER_RECONCILE_PING_SECRET — must also match PAYMENT_APPS_SCRIPT_SECRET in the Worker
 *   - SUPPORT_URL (optional)      — fallback support link used in the fulfillment email
 * Then deploy: Deploy → Manage deployments → (New deployment, or edit the
 * existing one and pick "New version") → Web app → Execute as: Me,
 * Who has access: Anyone. Copy the resulting /exec URL into the Worker's
 * PAYMENT_APPS_SCRIPT_URL secret. IMPORTANT: reusing "New version" on the
 * SAME deployment keeps the /exec URL stable; creating a brand new deployment
 * gives you a different URL and is exactly what causes the Worker's
 * "Apps Script gagal: HTTP 404" bug report if PAYMENT_APPS_SCRIPT_URL isn't
 * updated to match.
 *
 * Quick remote diagnostics (v2.1): GET the /exec URL with
 * ?secret=<API_SECRET> for basic health info, or
 * ?secret=<API_SECRET>&matrix=1 to also run checkPaymentMatrix_() and
 * include its report in the response.
 * ============================================================
 */

// ============================================================
// CONFIG
// ============================================================
// Hardcoded so this script always operates on the intended spreadsheet
// regardless of how/where it's executed from (see v2.1 changelog above).
var SPREADSHEET_ID = '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64';
var SCRIPT_VERSION = 'v2.2';
var SCRIPT_BUILD = '2026-09-28-lookup-diag'; // bump on every redeploy so /exec?secret=... proves which build is live

var PAYMENT_SHEET_NAME = 'Payment';
var FULFILLMENT_SHEET_NAME = 'Fulfillment';
var RAW_ALL_SHEET_NAME = 'Raw_all';
var DEFAULT_SUPPORT_URL = 'https://t.me/+tDtquz8U_fUzMTI9';

// Hardening knobs
var MAX_FIELD_LENGTH = 500;         // hard cap on any single string field from a request
var AUTH_FAIL_CACHE_KEY = 'payment_gs_auth_failures';
var AUTH_FAIL_THRESHOLD = 5;        // failed api_secret attempts allowed...
var AUTH_FAIL_WINDOW_SECONDS = 300; // ...within this rolling window before we start slowing responses down
var REF_PATTERN = /^[a-f0-9]{16,64}$/i;   // lynk.id refs observed as 32-char hex; informational only, never hard-blocks
var EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Payment sheet column indices (1-based), matching A:H exactly as laid out
// by the VSTACK formula plus the two script-managed columns.
var COL = {
  JUDUL: 1,   // A — Judul Barang
  HARGA: 2,   // B — Harga
  TANGGAL: 3, // C — Tanggal
  STATUS: 4,  // D — Status
  EMAIL: 5,   // E — Buyer Email
  REF: 6,     // F — Ref
  LISENSI: 7, // G — Lisensi Code (formula-driven, see ensurePaymentColumnFormulas_)
  TERKIRIM: 8 // H — Email Terkirim (formula-driven, see ensurePaymentColumnFormulas_)
};

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    apiSecret: props.getProperty('API_SECRET') || '',
    workerPingUrl: props.getProperty('WORKER_RECONCILE_PING_URL') || '',
    workerPingSecret: props.getProperty('WORKER_RECONCILE_PING_SECRET') || '',
    supportUrl: props.getProperty('SUPPORT_URL') || DEFAULT_SUPPORT_URL
  };
}

// ============================================================
// ENTRY POINTS
// ============================================================
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonOutput_({ success: false, error: 'Request body kosong / tidak valid' });
  }

  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ success: false, error: 'Invalid JSON body' });
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return jsonOutput_({ success: false, error: 'Body harus berupa JSON object' });
  }

  var cfg = getConfig_();
  if (!cfg.apiSecret || !safeCompare_(body.api_secret, cfg.apiSecret)) {
    var failCount = recordAuthFailureAndCheckLockout_();
    if (failCount > AUTH_FAIL_THRESHOLD) {
      console.error('Auth gagal berulang kali (' + failCount + 'x dalam ' + AUTH_FAIL_WINDOW_SECONDS +
        's) — kemungkinan percobaan brute-force pada endpoint ini.');
      Utilities.sleep(2000); // deliberate small delay to slow down brute-force attempts
    }
    return jsonOutput_({ success: false, error: 'Unauthorized: invalid api_secret' });
  }
  clearAuthFailures_();

  if (body.spreadsheet_id && SPREADSHEET_ID && body.spreadsheet_id !== SPREADSHEET_ID) {
    console.warn('spreadsheet_id pada request (' + body.spreadsheet_id + ') tidak sama dengan ' +
      'SPREADSHEET_ID script (' + SPREADSHEET_ID + ') — tetap memproses menggunakan spreadsheet ' +
      'yang di-bind ke script ini.');
  }

  try {
    if (body.action === 'lookup_payment') return jsonOutput_(withVersion_(handleLookupPayment_(body)));
    if (body.action === 'fulfill_license') return jsonOutput_(withVersion_(handleFulfillLicense_(body)));
    if (body.action === 'check_payment_matrix') return jsonOutput_(withVersion_({ success: true, report: checkPaymentMatrix_() }));
    return jsonOutput_({ success: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    console.error('doPost error: ' + (err && err.stack || err));
    return jsonOutput_({ success: false, error: String((err && err.message) || err) });
  }
}

// Browser-reachable health check — not called by the Worker, just useful for
// confirming the deployment URL is alive when debugging a 404. Unauthenticated
// callers get a bare "ok"; pass ?secret=<API_SECRET> for real diagnostics, and
// add &matrix=1 to also run the A:F payment matrix check (see v2.1 changelog).
function doGet(e) {
  var cfg = getConfig_();
  var providedSecret = e && e.parameter && e.parameter.secret;
  var authorized = cfg.apiSecret && providedSecret && safeCompare_(providedSecret, cfg.apiSecret);

  if (!authorized) {
    return jsonOutput_({ success: true, service: 'PaymentShopeeProductHunter.gs', version: SCRIPT_VERSION, status: 'ok' });
  }

  var ss;
  try {
    ss = getSpreadsheet_();
  } catch (err) {
    return jsonOutput_({ success: false, error: String((err && err.message) || err) });
  }

  var result = {
    success: true,
    service: 'PaymentShopeeProductHunter.gs',
    version: SCRIPT_VERSION,
    spreadsheet_id: ss.getId(),
    payment_sheet_exists: !!ss.getSheetByName(PAYMENT_SHEET_NAME),
    fulfillment_sheet_exists: !!ss.getSheetByName(FULFILLMENT_SHEET_NAME),
    time: new Date().toISOString()
  };

  try {
    var payForDiag = ss.getSheetByName(PAYMENT_SHEET_NAME);
    var fulForDiag = ss.getSheetByName(FULFILLMENT_SHEET_NAME);
    result.script_timezone = Session.getScriptTimeZone();
    result.fulfillment_rows = fulForDiag ? Math.max(0, fulForDiag.getLastRow() - 1) : 0;
    result.build = SCRIPT_BUILD;
    result.spreadsheet_locale = ss.getSpreadsheetLocale();
    result.payment_G2_lookup_healthy = !!payForDiag && isLookupHealthy_(payForDiag, COL.LISENSI, 'Lisensi Code');
    result.payment_H2_lookup_healthy = !!payForDiag && isLookupHealthy_(payForDiag, COL.TERKIRIM, 'Email Terkirim');
  } catch (err) {
    result.diagnostics_error = String((err && err.message) || err);
  }

  if (e && e.parameter && e.parameter.matrix === '1') {
    try {
      result.matrix = checkPaymentMatrix_();
    } catch (err) {
      result.matrix_error = String((err && err.message) || err);
    }
  }

  return jsonOutput_(result);
}

function withVersion_(obj) {
  obj.gs_version = SCRIPT_VERSION;
  return obj;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SECURITY / HARDENING HELPERS
// ============================================================

// Constant-time-ish string comparison so a mismatched api_secret doesn't
// leak "how many leading characters matched" via response timing. Still
// walks the full length of `a` on a length mismatch rather than returning
// immediately.
function safeCompare_(a, b) {
  a = String(a == null ? '' : a);
  b = String(b == null ? '' : b);
  if (a.length !== b.length) {
    var dummy = 0;
    for (var i = 0; i < a.length; i++) dummy |= a.charCodeAt(i) ^ a.charCodeAt(i);
    return false;
  }
  var diff = 0;
  for (var j = 0; j < a.length; j++) diff |= a.charCodeAt(j) ^ b.charCodeAt(j);
  return diff === 0;
}

function recordAuthFailureAndCheckLockout_() {
  var cache = CacheService.getScriptCache();
  var count = Number(cache.get(AUTH_FAIL_CACHE_KEY) || 0) + 1;
  cache.put(AUTH_FAIL_CACHE_KEY, String(count), AUTH_FAIL_WINDOW_SECONDS);
  return count;
}

function clearAuthFailures_() {
  CacheService.getScriptCache().remove(AUTH_FAIL_CACHE_KEY);
}

function truncate_(str, max) {
  return str.length > max ? str.substring(0, max) : str;
}

// Trim + length-cap any string coming from the request body before it's
// written to a sheet cell or used elsewhere.
function clean_(value) {
  return truncate_(String(value == null ? '' : value).trim(), MAX_FIELD_LENGTH);
}

function isValidEmail_(email) {
  return EMAIL_PATTERN.test(String(email || '').trim());
}

function isValidRef_(ref) {
  return REF_PATTERN.test(String(ref || '').trim());
}

// Strips CR/LF so a value that ends up in an email subject line can't be
// used for header injection (e.g. a crafted product_title containing
// "\r\nBcc: ...").
function sanitizeForEmailHeader_(str) {
  return String(str == null ? '' : str).replace(/[\r\n]+/g, ' ').trim();
}

function sortByTanggalAsc_(list) {
  list.sort(function (a, b) {
    var ta = a.tanggalDate ? a.tanggalDate.getTime() : 0;
    var tb = b.tanggalDate ? b.tanggalDate.getTime() : 0;
    return ta - tb;
  });
  return list;
}

// ============================================================
// SHEET HELPERS
// ============================================================
function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (err) {
      // Fail loudly rather than silently falling back to whatever happens to
      // be "active" — that would defeat the point of hardcoding the ID.
      throw new Error('Tidak bisa membuka spreadsheet dengan SPREADSHEET_ID=' + SPREADSHEET_ID + ': ' + err);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getPaymentSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(PAYMENT_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + PAYMENT_SHEET_NAME + '" tidak ditemukan');
  return sheet;
}

function getOrCreateFulfillmentSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(FULFILLMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FULFILLMENT_SHEET_NAME);
    sheet.getRange('A1:F1').setValues([[
      'Ref', 'Lisensi Code', 'Email Terkirim', 'Buyer Email', 'Judul Barang', 'Plan Name'
    ]]);
    sheet.setFrozenRows(1);
  }
  // Force Ref (A) and the timestamp (C) to plain text so Sheets can never
  // auto-convert a hex Ref such as "12345e6789" into a number, or a
  // "28-09-2026 12:21" stamp into a date serial — either would silently
  // break the VLOOKUP that feeds Payment!G:H.
  if (sheet.getRange('A2').getNumberFormat() !== '@') {
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('C:C').setNumberFormat('@');
  }
  return sheet;
}

function normalize_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

// Product titles come from lynk.id and can carry stray/double/non-breaking
// spaces ("Paket Pro  Shopee Product Hunter"). Collapse all whitespace before
// comparing so such a typo can't silently stop a paid row from matching.
function normalizeTitle_(value) {
  return String(value == null ? '' : value).replace(/[\s\u00A0]+/g, ' ').trim().toLowerCase();
}

// Accepts either a real Date (if Sheets auto-parsed the cell) or a
// "DD-MM-YYYY HH:MM" string (matching the Tanggal format used throughout
// Raw_all / Payment, e.g. "25-09-2026 15:33").
function parseSheetDateTime_(value) {
  if (value instanceof Date) return value;
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return null;
  var m = raw.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  var day = Number(m[1]), month = Number(m[2]), year = Number(m[3]);
  var hour = Number(m[4] || 0), minute = Number(m[5] || 0), second = Number(m[6] || 0);
  return new Date(year, month - 1, day, hour, minute, second);
}

function formatSheetDateTime_(date) {
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  return pad(date.getDate()) + '-' + pad(date.getMonth() + 1) + '-' + date.getFullYear() + ' ' +
    pad(date.getHours()) + ':' + pad(date.getMinutes());
}

// Reads Payment!A2:H<lastRow> once and returns plain row objects. Always a
// fresh read — nothing here is ever cached across calls, which is exactly
// what avoids the row-alignment bug: every lookup/fulfill re-reads current
// reality instead of trusting a remembered row number.
function readPaymentRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var ref = String(r[COL.REF - 1] || '').trim();
    if (!ref && !r[COL.JUDUL - 1]) continue; // fully blank row (past the live formula's spill range)
    rows.push({
      rowIndex: i + 2,
      judul: String(r[COL.JUDUL - 1] || '').trim(),
      harga: r[COL.HARGA - 1],
      tanggalRaw: r[COL.TANGGAL - 1],
      tanggalDate: parseSheetDateTime_(r[COL.TANGGAL - 1]),
      status: String(r[COL.STATUS - 1] || '').trim(),
      buyerEmail: String(r[COL.EMAIL - 1] || '').trim(),
      ref: ref,
      lisensiCode: String(r[COL.LISENSI - 1] || '').trim(),
      emailTerkirim: r[COL.TERKIRIM - 1]
    });
  }
  return rows;
}

function getFulfillmentMap_() {
  var sheet = getOrCreateFulfillmentSheet_();
  var lastRow = sheet.getLastRow();
  var map = {}; // ref -> { rowIndex, licenseCode, emailSentAt }
  if (lastRow < 2) return map;
  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    var ref = String(values[i][0] || '').trim();
    if (!ref) continue;
    map[ref] = { rowIndex: i + 2, licenseCode: String(values[i][1] || '').trim(), emailSentAt: values[i][2] };
  }
  return map;
}

// Installs the self-aligning lookup formulas for Payment!G (Lisensi Code) and
// Payment!H (Email Terkirim): plain-text header in row 1, ONE ARRAYFORMULA in
// row 2 that spills downward.
//
// LOCALE NOTE (learned the hard way): Range.setFormula() takes US syntax and
// Sheets then "translates" it to the spreadsheet locale. On this id-ID
// spreadsheet (";" as argument separator) that translation only converted
// SOME commas when the formula contained an empty-string argument ("",),
// leaving a half-comma/half-semicolon formula => #ERROR!. Two defences:
//   1. The formula below contains NO "" literals at all (IFERROR(x) with one
//      argument already yields blank, and VLOOKUP of a blank key is #N/A).
//   2. ensureLookupColumn_ verifies the cell after setting it and, if it still
//      shows #ERROR!, retries with ";" separators before giving up loudly.
// Manual fallback (type into G2 / H2 yourself, id-ID locale):
//   =ARRAYFORMULA(IFERROR(VLOOKUP(TRIM(F2:F);Fulfillment!$A:$B;2;FALSE)))
//   =ARRAYFORMULA(IFERROR(VLOOKUP(TRIM(F2:F);Fulfillment!$A:$C;3;FALSE)))
function ensurePaymentColumnFormulas_() {
  var payment = getPaymentSheet_();
  getOrCreateFulfillmentSheet_(); // must exist before the VLOOKUP formulas reference it

  ensureLookupColumn_(payment, COL.LISENSI, 'Lisensi Code',
    '=ARRAYFORMULA(IFERROR(VLOOKUP(TRIM(F2:F),' + FULFILLMENT_SHEET_NAME + '!$A:$B,2,FALSE)))');
  ensureLookupColumn_(payment, COL.TERKIRIM, 'Email Terkirim',
    '=ARRAYFORMULA(IFERROR(VLOOKUP(TRIM(F2:F),' + FULFILLMENT_SHEET_NAME + '!$A:$C,3,FALSE)))');
}

// Converts a US-syntax formula to ";"-separated syntax by swapping every
// comma that is outside a string literal.
function toSemicolonSyntax_(formula) {
  var out = '';
  var inString = false;
  for (var i = 0; i < formula.length; i++) {
    var ch = formula.charAt(i);
    if (ch === '"') inString = !inString;
    out += (ch === ',' && !inString) ? ';' : ch;
  }
  return out;
}

// True only when header + formula are the current ones AND the formula
// actually evaluates (no #ERROR!/#REF! in the anchor cell).
function isLookupHealthy_(sheet, col, header) {
  var headerCell = sheet.getRange(1, col);
  var bodyCell = sheet.getRange(2, col);
  var f = bodyCell.getFormula();
  return headerCell.getFormula() === '' &&
    String(headerCell.getValue()) === header &&
    f.indexOf(FULFILLMENT_SHEET_NAME + '!') !== -1 &&
    f.indexOf('""') === -1 &&                              // older builds used "" and broke on ";" locales
    String(bodyCell.getDisplayValue()).indexOf('#') !== 0;  // '#ERROR!', '#REF!', '#N/A' ...
}

function ensureLookupColumn_(sheet, col, header, usFormula) {
  if (isLookupHealthy_(sheet, col, header)) return;

  // First install, a broken/older formula, or leftover plain values: wipe the
  // column so the spill isn't blocked. (Run migrateOrphanedLicenseColumns_()
  // BEFORE this if old plain-value license codes in G/H must be preserved.)
  var headerCell = sheet.getRange(1, col);
  var bodyCell = sheet.getRange(2, col);
  sheet.getRange(1, col, sheet.getMaxRows(), 1).clearContent();
  headerCell.setValue(header);

  var attempts = [usFormula, toSemicolonSyntax_(usFormula)];
  for (var i = 0; i < attempts.length; i++) {
    bodyCell.setFormula(attempts[i]);
    SpreadsheetApp.flush();
    if (String(bodyCell.getDisplayValue()).indexOf('#ERROR') === -1) {
      console.log('Payment kolom ' + col + ': formula terpasang (percobaan ' + (i + 1) + ').');
      return;
    }
    console.warn('Payment kolom ' + col + ': percobaan ' + (i + 1) + ' menghasilkan #ERROR! -> ' + attempts[i]);
  }
  throw new Error('Gagal memasang formula lookup di Payment kolom ' + col + ' (locale spreadsheet: ' +
    getSpreadsheet_().getSpreadsheetLocale() + '). Pasang manual — lihat komentar di atas ensurePaymentColumnFormulas_.');
}

// Re-reads Payment!G for every row carrying this Ref and confirms the
// license is really visible there. Never trusts that the formula "must have"
// picked up the new Fulfillment row.
function verifyFulfillmentVisible_(ref, licenseCode) {
  var attempt = function () {
    SpreadsheetApp.flush();
    var rows = readPaymentRows_(getPaymentSheet_()).filter(function (r) { return r.ref === ref; });
    if (!rows.length) return { verified: false, reason: 'ref_not_found_in_payment_sheet', rows: [] };
    var wrong = rows.filter(function (r) { return r.lisensiCode !== licenseCode; });
    if (wrong.length) {
      return { verified: false, reason: 'payment_G_does_not_show_license',
        rows: wrong.map(function (r) { return r.rowIndex; }) };
    }
    return { verified: true, reason: '', rows: rows.map(function (r) { return r.rowIndex; }) };
  };
  var result = attempt();
  if (!result.verified) {
    Utilities.sleep(1500); // give the array formula a moment to recalculate, then look once more
    result = attempt();
  }
  return result;
}

// ============================================================
// ACTION: lookup_payment
// Body: { action, api_secret, spreadsheet_id, sheet_name, product_title,
//         buyer_email, created_at }
// Response: { success:true, matched:true, ref, completed_at }
//        or { success:true, matched:false, reason, candidates }
// ============================================================
function handleLookupPayment_(body) {
  var productTitle = normalizeTitle_(clean_(body.product_title));
  var buyerEmail = normalize_(clean_(body.buyer_email));
  if (!productTitle || !buyerEmail) {
    return { success: false, error: 'product_title and buyer_email are required' };
  }

  var createdAt = body.created_at ? new Date(body.created_at) : null;
  var hasCreatedAt = !!createdAt && !isNaN(createdAt.getTime());
  var payment = getPaymentSheet_();
  var rows = readPaymentRows_(payment);

  // Step-by-step narrowing so a "no match" can say WHICH criterion failed.
  var successRows = rows.filter(function (r) { return normalize_(r.status) === 'success' && r.ref; });
  var emailRows = successRows.filter(function (r) { return normalize_(r.buyerEmail) === buyerEmail; });
  var titleRows = emailRows.filter(function (r) { return normalizeTitle_(r.judul) === productTitle; });
  var candidates = titleRows.filter(function (r) {
    // The paid row's timestamp must be at/after the order was created
    // (with a small tolerance for clock skew) — stops an old, unrelated
    // successful payment for the same product+email from being matched to
    // a brand-new pending order. Needs the script time zone to be
    // Asia/Jakarta (sheet timestamps are WIB).
    if (hasCreatedAt && r.tanggalDate) {
      if (r.tanggalDate.getTime() + 60000 < createdAt.getTime()) return false;
    }
    return true;
  });

  if (candidates.length === 0) {
    var reason = emailRows.length === 0 ? 'email_not_found'
      : titleRows.length === 0 ? 'title_mismatch'
      : 'paid_before_order_created';
    var seenTitles = [];
    emailRows.forEach(function (r) { if (seenTitles.indexOf(r.judul) === -1 && seenTitles.length < 5) seenTitles.push(r.judul); });
    return {
      success: true, matched: false, reason: reason, candidates: 0,
      diag: { success_rows: successRows.length, email_rows: emailRows.length, title_rows: titleRows.length,
              seen_titles_for_email: seenTitles, script_timezone: Session.getScriptTimeZone() }
    };
  }

  // Always resolve in true payment order (lynk.id success time), never in
  // whatever order the sheet's own SORT() formula currently happens to use.
  sortByTanggalAsc_(candidates);

  // Several rows can legitimately share ONE Ref (duplicate line in Raw_all);
  // that is not ambiguity. Only different Refs for the same product+email are.
  var refs = [];
  candidates.forEach(function (r) { if (refs.indexOf(r.ref) === -1) refs.push(r.ref); });
  if (refs.length > 1) {
    // Ambiguous on purpose rather than guessing — e.g. the same buyer bought
    // the same package twice. The Worker's per-Ref dedupe (findPaymentByLynkRef)
    // still protects against double-fulfilling, but we'd rather report
    // "ambiguous" than silently pick the wrong one.
    return { success: true, matched: false, reason: 'ambiguous', candidates: refs.length };
  }

  var match = candidates[0];
  return {
    success: true,
    matched: true,
    ref: match.ref,
    completed_at: match.tanggalDate ? formatSheetDateTime_(match.tanggalDate) : String(match.tanggalRaw || '')
  };
}

// ============================================================
// ACTION: fulfill_license
// Body: { action, api_secret, spreadsheet_id, sheet_name, ref, license_code,
//         product_title, buyer_email, plan_name, support_url }
// Response: { success:true, status:'fulfilled', license_code, ref }
//        or { success:true, status:'pending' }   (ref's row isn't SUCCESS — shouldn't
//                                                  normally happen; defensive only)
//        or { success:false, error }
// ============================================================
function handleFulfillLicense_(body) {
  var ref = clean_(body.ref);
  var licenseCode = clean_(body.license_code);
  if (!ref) return { success: false, error: 'ref is required' };
  if (!licenseCode) return { success: false, error: 'license_code is required' };
  if (!isValidRef_(ref)) {
    console.warn('fulfill_license dipanggil dengan ref berformat tidak umum (bukan hex 16-64 char): ' + ref);
  }

  // Guards the read-check-then-append critical section below so two
  // near-simultaneous fulfill_license calls for the same Ref (e.g. a Worker
  // retry racing the original request) can't both pass the "not yet
  // fulfilled" check and append two Fulfillment rows for the same Ref.
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try {
    gotLock = lock.tryLock(30000);
    if (!gotLock) {
      return { success: false, error: 'Server sibuk (lock timeout), silakan coba lagi.' };
    }

    // Idempotency: if this Ref has already been fulfilled, echo back the
    // existing license instead of writing a duplicate row / sending a second
    // email. This is what makes it safe for the Worker to retry this call.
    var fulfillment = getFulfillmentMap_();
    var existing = fulfillment[ref];
    if (existing) {
      ensurePaymentColumnFormulas_();
      var again = verifyFulfillmentVisible_(ref, existing.licenseCode);
      return { success: true, status: 'fulfilled', license_code: existing.licenseCode, ref: ref,
        sheet_verified: again.verified, sheet_verify_reason: again.reason, sheet_rows: again.rows };
    }

    // Fresh scan for the row currently holding this Ref — never a remembered
    // row number from an earlier lookup_payment call (that's the whole fix
    // for the alignment bug: A:F can have re-sorted any number of times
    // between lookup_payment and fulfill_license).
    var payment = getPaymentSheet_();
    var rows = readPaymentRows_(payment);
    var row = null;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].ref === ref) { row = rows[i]; break; }
    }
    if (!row || normalize_(row.status) !== 'success') {
      return { success: true, status: 'pending' };
    }

    var buyerEmail = clean_(body.buyer_email || row.buyerEmail);
    var productTitle = clean_(body.product_title || row.judul);
    var planName = clean_(body.plan_name);

    var sheet = getOrCreateFulfillmentSheet_();
    var now = new Date();
    sheet.appendRow([
      ref,
      licenseCode,
      formatSheetDateTime_(now),
      buyerEmail,
      productTitle,
      planName
    ]);

    ensurePaymentColumnFormulas_();
    var check = verifyFulfillmentVisible_(ref, licenseCode);
    if (!check.verified) {
      console.error('fulfill_license ref=' + ref + ': tercatat di Fulfillment tapi TIDAK tampil di Payment!G — ' +
        check.reason + ' rows=' + JSON.stringify(check.rows));
    }

    if (buyerEmail) {
      if (!isValidEmail_(buyerEmail)) {
        console.error('Alamat email tidak valid, tidak mengirim email lisensi: ' + buyerEmail);
      } else {
        try {
          sendLicenseEmail_(buyerEmail, licenseCode, productTitle, planName, clean_(body.support_url));
        } catch (err) {
          // Don't fail the whole fulfillment over an email hiccup — the license
          // is already recorded and the Worker also notifies the buyer via
          // Telegram. Log it so it's visible in Apps Script's execution log.
          console.error('sendLicenseEmail_ failed for ref ' + ref + ': ' + err);
        }
      }
    }

    return { success: true, status: 'fulfilled', license_code: licenseCode, ref: ref,
      sheet_verified: check.verified, sheet_verify_reason: check.reason, sheet_rows: check.rows };
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

// ============================================================
// LICENSE EMAIL (v2.2): HTML + plain-text fallback, pure ASCII (no emoji —
// 4-byte emoji were what showed up as "������" in some mail clients).
// ============================================================
function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// "Starter" from plan_name, or derived from "Paket Starter Shopee Product Hunter".
function resolvePlanLabel_(planName, productTitle) {
  var plan = String(planName || '').trim();
  if (plan) return plan;
  var m = String(productTitle || '').match(/^Paket\s+(.+?)\s+Shopee Product Hunter$/i);
  return m ? m[1] : (String(productTitle || '').trim() || 'Shopee Product Hunter');
}

// Pure function (no Apps Script services) so it can be previewed/tested anywhere.
function buildLicenseEmail_(licenseCode, productTitle, planName, supportUrl) {
  var plan = resolvePlanLabel_(planName, productTitle);
  var safeUrl = /^https?:\/\//i.test(String(supportUrl || '')) ? String(supportUrl) : DEFAULT_SUPPORT_URL;
  var steps = [
    'Buka Google Chrome.',
    'Buka extension Shopee Product Hunter.',
    'Login/Register menggunakan akun Anda.',
    'Pilih menu Aktifkan License.',
    'Masukkan License Code di atas.',
    'Klik Aktifkan.',
    'Setelah aktif, extension siap digunakan.'
  ];
  var notes = [
    'Simpan License Code ini dengan aman.',
    'Jangan membagikan License Code kepada orang lain.',
    'Jika mengalami kendala aktivasi, hubungi support.'
  ];

  var subject = 'License Code Shopee Product Hunter - ' + sanitizeForEmailHeader_(plan);

  var text =
    'Pembayaran Berhasil\n\n' +
    'Halo,\n\n' +
    'Terima kasih telah melakukan pembelian Shopee Product Hunter.\n' +
    'Pembayaran Anda telah berhasil diverifikasi.\n\n' +
    'Paket: ' + plan + '\n' +
    'License Code: ' + licenseCode + '\n\n' +
    'Panduan Aktivasi Chrome Extension Shopee Product Hunter\n' +
    steps.map(function (t, i) { return (i + 1) + '. ' + t; }).join('\n') + '\n\n' +
    'Penting:\n' +
    notes.map(function (t) { return '- ' + t; }).join('\n') + '\n\n' +
    'Terima kasih telah menggunakan Shopee Product Hunter.\n\n' +
    'Support: ' + safeUrl + '\n';

  var font = 'font-family:Arial,Helvetica,sans-serif;';
  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f4f5f7;">' +
    '<div style="max-width:560px;margin:0 auto;padding:24px 16px;' + font + 'color:#1f2937;font-size:14px;line-height:1.6;">' +
      '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:8px;padding:28px;">' +
        '<h2 style="margin:0 0 16px 0;font-size:20px;color:#111827;">Pembayaran Berhasil</h2>' +
        '<p style="margin:0 0 12px 0;">Halo,</p>' +
        '<p style="margin:0 0 12px 0;">Terima kasih telah melakukan pembelian <b>Shopee Product Hunter</b>.</p>' +
        '<p style="margin:0 0 20px 0;">Pembayaran Anda telah berhasil diverifikasi.</p>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin:0 0 24px 0;">' +
          '<tr><td style="padding:10px 12px;border:1px solid #e5e7eb;background:#f9fafb;width:34%;"><b>Paket</b></td>' +
              '<td style="padding:10px 12px;border:1px solid #e5e7eb;">' + escapeHtml_(plan) + '</td></tr>' +
          '<tr><td style="padding:10px 12px;border:1px solid #e5e7eb;background:#f9fafb;"><b>License Code</b></td>' +
              '<td style="padding:10px 12px;border:1px solid #e5e7eb;font-family:Consolas,Menlo,monospace;font-size:15px;font-weight:bold;letter-spacing:0.5px;">' +
              escapeHtml_(licenseCode) + '</td></tr>' +
        '</table>' +
        '<h3 style="margin:0 0 8px 0;font-size:15px;color:#111827;">Panduan Aktivasi Chrome Extension Shopee Product Hunter</h3>' +
        '<ol style="margin:0 0 24px 0;padding-left:22px;">' +
          steps.map(function (t) { return '<li style="margin:0 0 4px 0;">' + escapeHtml_(t) + '</li>'; }).join('') +
        '</ol>' +
        '<h3 style="margin:0 0 8px 0;font-size:15px;color:#111827;">Penting:</h3>' +
        '<ul style="margin:0 0 24px 0;padding-left:22px;">' +
          notes.map(function (t) { return '<li style="margin:0 0 4px 0;">' + escapeHtml_(t) + '</li>'; }).join('') +
        '</ul>' +
        '<p style="margin:0 0 12px 0;">Terima kasih telah menggunakan Shopee Product Hunter.</p>' +
        '<p style="margin:0;">Support: <a href="' + escapeHtml_(safeUrl) + '" style="color:#2563eb;">' + escapeHtml_(safeUrl) + '</a></p>' +
      '</div>' +
    '</div></body></html>';

  return { subject: subject, text: text, html: html };
}

function sendLicenseEmail_(buyerEmail, licenseCode, productTitle, planName, supportUrlFromBody) {
  if (MailApp.getRemainingDailyQuota() <= 0) {
    throw new Error('MailApp daily quota exceeded — email lisensi untuk ' + buyerEmail + ' tidak terkirim otomatis, perlu dikirim manual.');
  }
  var cfg = getConfig_();
  var mail = buildLicenseEmail_(licenseCode, productTitle, planName, supportUrlFromBody || cfg.supportUrl);
  GmailApp.sendEmail(buyerEmail, mail.subject, mail.text, { htmlBody: mail.html, name: 'Shopee Product Hunter' });
}

// ============================================================
// FEATURE: A:F payment matrix check (v2.1)
// Validates Payment!A:F (Judul Barang, Harga, Tanggal, Status, Buyer Email,
// Ref) row by row and cross-checks against the Fulfillment sheet. Read-only
// — never modifies any sheet. Reachable from the "Payment Tools" menu, via
// doPost action "check_payment_matrix", or doGet ?secret=...&matrix=1.
// ============================================================
function checkPaymentMatrix_() {
  var payment = getPaymentSheet_();
  var rows = readPaymentRows_(payment);
  var fulfillment = getFulfillmentMap_();

  var issues = [];
  var seenRefJudul = {};
  var lastTime = null;
  var fulfilledCount = 0;
  var unfulfilledCount = 0;

  rows.forEach(function (r) {
    // 1. Required fields present
    var missing = [];
    if (!r.judul) missing.push('Judul Barang');
    if (r.harga === '' || r.harga === null || typeof r.harga === 'undefined') missing.push('Harga');
    if (!r.tanggalRaw) missing.push('Tanggal');
    if (!r.status) missing.push('Status');
    if (!r.buyerEmail) missing.push('Buyer Email');
    if (!r.ref) missing.push('Ref');
    if (missing.length) {
      issues.push({ rowIndex: r.rowIndex, type: 'missing_field', detail: missing.join(', ') });
    }

    // 2. Status sanity — this sheet is filtered to only SUCCESS rows by the
    // upstream VSTACK/FILTER formula, so anything else here means that
    // formula was edited or broken.
    if (r.status && normalize_(r.status) !== 'success') {
      issues.push({ rowIndex: r.rowIndex, type: 'unexpected_status', detail: r.status });
    }

    // 3. Tanggal must be parseable
    if (r.tanggalRaw && !r.tanggalDate) {
      issues.push({ rowIndex: r.rowIndex, type: 'unparseable_date', detail: String(r.tanggalRaw) });
    }

    // 4. Ref format sanity — informational only; lynk.id's format could
    // change, so this warns rather than blocking anything.
    if (r.ref && !isValidRef_(r.ref)) {
      issues.push({ rowIndex: r.rowIndex, type: 'unusual_ref_format', detail: r.ref });
    }

    // 5. Email format sanity
    if (r.buyerEmail && !isValidEmail_(r.buyerEmail)) {
      issues.push({ rowIndex: r.rowIndex, type: 'invalid_email', detail: r.buyerEmail });
    }

    // 6. Exact duplicate line item (same Ref + same Judul Barang appearing
    // twice) — points at a duplicated row in Raw_all rather than a genuine
    // second purchase.
    var key = r.ref + '||' + normalize_(r.judul);
    if (seenRefJudul[key]) {
      issues.push({ rowIndex: r.rowIndex, type: 'duplicate_line_item', detail: 'sama seperti baris ' + seenRefJudul[key] });
    } else {
      seenRefJudul[key] = r.rowIndex;
    }

    // 7. Chronological order — Payment!A:F is expected to be sorted
    // ascending by Tanggal (see the SORT(...,3,TRUE) referenced in the file
    // header). Flag it if that ever stops being true, since it's the
    // assumption several helpers use as a best-effort default (the actual
    // Ref-keyed VLOOKUP for G/H does not depend on this, only the ordering
    // used when reporting/picking among ambiguous rows does).
    if (r.tanggalDate) {
      if (lastTime !== null && r.tanggalDate.getTime() < lastTime) {
        issues.push({ rowIndex: r.rowIndex, type: 'out_of_chronological_order', detail: formatSheetDateTime_(r.tanggalDate) });
      }
      lastTime = r.tanggalDate.getTime();
    }

    // 8. Fulfillment cross-check
    if (r.ref && fulfillment[r.ref]) {
      fulfilledCount++;
      if (r.lisensiCode !== fulfillment[r.ref].licenseCode) {
        issues.push({ rowIndex: r.rowIndex, type: 'fulfilled_but_not_visible_in_G',
          detail: 'Fulfillment punya lisensi untuk Ref ini, tapi Payment!G baris ini ' + (r.lisensiCode ? 'berbeda' : 'kosong') });
      }
    } else {
      unfulfilledCount++;
    }
  });

  var report = {
    checkedAt: new Date().toISOString(),
    totalRows: rows.length,
    fulfilled: fulfilledCount,
    unfulfilled: unfulfilledCount,
    fulfillmentRows: Object.keys(fulfillment).length,
    issueCount: issues.length,
    issues: issues
  };

  console.log('Payment matrix check: ' + rows.length + ' baris, ' + issues.length + ' isu ditemukan.');
  if (issues.length) console.log(JSON.stringify(issues, null, 2));
  return report;
}

function runPaymentMatrixCheckFromMenu_() {
  var report = checkPaymentMatrix_();
  var ui = SpreadsheetApp.getUi();
  var msg = 'Total baris: ' + report.totalRows + '\n' +
    'Sudah fulfilled: ' + report.fulfilled + '\n' +
    'Belum fulfilled: ' + report.unfulfilled + '\n' +
    'Isu ditemukan: ' + report.issueCount + '\n\n';
  if (report.issueCount > 0) {
    msg += 'Isu (maks 20 ditampilkan di sini, selebihnya lihat Log Eksekusi):\n';
    report.issues.slice(0, 20).forEach(function (it) {
      msg += '  Baris ' + it.rowIndex + ' — ' + it.type + ': ' + it.detail + '\n';
    });
  } else {
    msg += 'Tidak ada isu ditemukan pada kolom A:F.';
  }
  ui.alert('Hasil Cek Matrix Pembayaran', msg, ui.ButtonSet.OK);
}

// ============================================================
// ~1-MINUTE CRON — detect unfulfilled SUCCESS rows, ping the Worker
// ============================================================
function pingWorkerWithRetry_(url, options, retries) {
  var lastErr = null;
  for (var attempt = 0; attempt <= retries; attempt++) {
    try {
      var res = UrlFetchApp.fetch(url, options);
      var code = res.getResponseCode();
      if (code >= 200 && code < 300) return res;
      lastErr = new Error('HTTP ' + code + ': ' + res.getContentText().substring(0, 300));
    } catch (err) {
      lastErr = err;
    }
    if (attempt < retries) Utilities.sleep(1000 * (attempt + 1));
  }
  throw lastErr;
}

function checkPendingPaymentsAndReport_() {
  var cfg = getConfig_();
  var payment = getPaymentSheet_();
  var rows = readPaymentRows_(payment);
  var fulfillment = getFulfillmentMap_();

  var pending = rows.filter(function (r) {
    return normalize_(r.status) === 'success' && r.ref && !fulfillment[r.ref];
  });

  if (pending.length === 0) return; // nothing to do — stay quiet, don't ping every minute for no reason

  // Always report/process oldest (by real lynk.id payment time) first,
  // regardless of the Payment sheet's current row order — see v2.1
  // changelog ("ORDERING GUARANTEE").
  sortByTanggalAsc_(pending);

  if (!cfg.workerPingUrl) {
    console.log('WORKER_RECONCILE_PING_URL belum diset di Script Properties — ' +
      pending.length + ' pembayaran belum di-fulfill tapi tidak bisa memberitahu Worker.');
    return;
  }

  var payload = {
    source: 'apps-script-cron',
    checked_at: new Date().toISOString(),
    count: pending.length,
    pending_refs: pending.slice(0, 50).map(function (r) {
      return { ref: r.ref, product_title: r.judul, buyer_email: r.buyerEmail, tanggal: String(r.tanggalRaw || '') };
    })
  };

  try {
    var res = pingWorkerWithRetry_(cfg.workerPingUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Apps-Script-Secret': cfg.workerPingSecret },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }, 1);
    console.log('Reconcile ping (' + pending.length + ' pending) -> HTTP ' + res.getResponseCode() +
      ': ' + res.getContentText().substring(0, 300));
  } catch (err) {
    console.error('Failed to ping Worker at ' + cfg.workerPingUrl + ' after retries: ' + err);
  }
}

function installPaymentCheckerTrigger() {
  removePaymentCheckerTrigger();
  ScriptApp.newTrigger('checkPendingPaymentsAndReport_')
    .timeBased()
    .everyMinutes(1)
    .create();
  console.log('Trigger installed: checkPendingPaymentsAndReport_ every 1 minute.');
}

function removePaymentCheckerTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkPendingPaymentsAndReport_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ============================================================
// ONE-TIME SETUP
// ============================================================
function setup() {
  getOrCreateFulfillmentSheet_();
  ensurePaymentColumnFormulas_();
  installPaymentCheckerTrigger();
  console.log('Setup complete: Fulfillment sheet ready, Payment!G:H formulas installed, 1-minute trigger installed.');
}

// ============================================================
// ONE-TIME MIGRATION — run manually if Payment!G:H already has plain-value
// data from before this version. Folds anything still correctly aligned
// (G has a code AND that same row's F/Ref is non-blank) into the new
// Fulfillment sheet, and separately lists anything already orphaned (G has
// a code but F on that row is blank — like the two stray SPH- codes in the
// reference screenshot) for manual review, since there's no way to
// automatically recover which Ref an orphaned code belonged to once its
// row's data has been overwritten by the live A:F formula.
// ============================================================
function migrateOrphanedLicenseColumns_() {
  var payment = getPaymentSheet_();
  var lastRow = payment.getLastRow();
  if (lastRow < 2) { console.log('Payment sheet has no data rows — nothing to migrate.'); return; }

  var values = payment.getRange(2, 1, lastRow - 1, 8).getValues();
  var fulfillmentSheet = getOrCreateFulfillmentSheet_();
  var existingRefs = getFulfillmentMap_();

  var migrated = 0;
  var orphaned = [];
  var toAppend = [];

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var ref = String(r[COL.REF - 1] || '').trim();
    var lisensi = String(r[COL.LISENSI - 1] || '').trim();
    if (!lisensi) continue; // nothing written for this row, skip

    if (!ref) {
      orphaned.push({ rowIndex: i + 2, lisensiCode: lisensi, emailTerkirim: r[COL.TERKIRIM - 1] });
      continue;
    }
    if (existingRefs[ref]) continue; // already migrated / already in Fulfillment

    toAppend.push([
      ref,
      lisensi,
      r[COL.TERKIRIM - 1] || formatSheetDateTime_(new Date()),
      r[COL.EMAIL - 1] || '',
      r[COL.JUDUL - 1] || '',
      ''
    ]);
    migrated++;
  }

  if (toAppend.length > 0) {
    fulfillmentSheet.getRange(fulfillmentSheet.getLastRow() + 1, 1, toAppend.length, 6).setValues(toAppend);
  }

  // Now that everything recoverable is safely copied into Fulfillment,
  // install the self-aligning formulas (this also clears the old plain
  // values in G:H, including the orphaned ones — they're logged below
  // first, so nothing is lost silently).
  ensurePaymentColumnFormulas_();

  console.log('Migration done. Migrated: ' + migrated + '. Orphaned (needs manual review): ' + orphaned.length);
  if (orphaned.length > 0) {
    console.log('Orphaned license codes (row had a code but blank Ref — cross-reference manually against buyer support requests):');
    orphaned.forEach(function (o) {
      console.log('  row ' + o.rowIndex + ': ' + o.lisensiCode + ' (Email Terkirim: ' + o.emailTerkirim + ')');
    });
  }
  return { migrated: migrated, orphaned: orphaned };
}

// ============================================================
// Optional: Sheets UI menu for manual/admin use
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Payment Tools')
    .addItem('Setup (create Fulfillment sheet + formulas + trigger)', 'setup')
    .addItem('Run migration for old G/H values', 'migrateOrphanedLicenseColumns_')
    .addItem('Cek Matrix Pembayaran (A:F)', 'runPaymentMatrixCheckFromMenu_')
    .addItem('Check pending payments now', 'checkPendingPaymentsAndReport_')
    .addToUi();
}
