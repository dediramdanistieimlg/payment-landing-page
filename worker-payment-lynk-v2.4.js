// ============================================================
// worker-payment.js — v2.4 — LYNK.ID + Google Spreadsheet Payment Gateway
// Shopee Product Hunter v2 — Lynk.id + 2 Telegram Bots + Supabase
// ============================================================
//
// v2.4 CHANGES (this pass):
//   - FEATURE: public WEB checkout flow for the landing page (no JWT, no
//     Telegram needed):
//       GET  /api/v1/web/plans               — price list (single source of truth = PRICING_TIERS)
//       POST /api/v1/web/checkout            — {plan_duration_days, email} → creates (or reuses) a
//                                              pending order and returns the Lynk.id checkout link
//       GET  /api/v1/web/status/:order_id?email=…
//                                            — reconciles against the Payment sheet and, once paid,
//                                              returns the license key (order_id + email must match)
//     Web orders are stored in payments.guest_chat_id as "web:<email>" so NO schema
//     change is needed (the column must be TEXT — it already holds Telegram chat ids as strings).
//     completePayment() skips the Telegram message for web orders; the license is shown on the
//     page and e-mailed by Apps Script exactly as before.
//   - RELIABILITY: reconcilePendingPayments() now scans only orders from the last 3 days, newest
//     first. Before, it read the 50 OLDEST pending rows forever — abandoned checkouts would
//     eventually fill that window and starve new orders (much more likely once a public web
//     checkout exists).
//   - SECURITY: web checkout has a best-effort per-IP rate limit and a honeypot field; web status
//     polling is throttled per order so it can't hammer Apps Script.
//
// v2.3 CHANGES:
//   - FEATURE: new POST /api/v1/payment/reconcile-ping endpoint (secured by
//     PAYMENT_APPS_SCRIPT_SECRET via the X-Apps-Script-Secret header). This
//     is the receiving end for PaymentShopeeProductHunter.gs's new ~1-minute
//     time trigger: when Apps Script notices a paid-but-unfulfilled row, it
//     pings this endpoint and the Worker runs its existing
//     reconcilePendingPayments() sweep immediately instead of waiting for
//     the next 5-minute cron tick — so buyers get their license within
//     roughly a minute of paying rather than up to five.
//   - PERFORMANCE: added a 1-hour, in-memory, per-isolate response cache for
//     every non-payment bot feature (menu, /paket, /fitur, /support, /help,
//     /akun, /license). Anything that shows live payment data — /status,
//     /riwayat, the "Cek Pembayaran" callback, payment creation, and the
//     HTTP payment endpoints — is deliberately left out of this cache and
//     always computed fresh. completePayment() also proactively clears a
//     user's cached /akun and /license entries the moment their purchase
//     completes, so they don't see stale "no license" info if they check
//     right after paying.
//
// v2.2 CHANGES:
//   - FIX (BUG REPORT "Apps Script gagal: HTTP 404"): reconcilePaymentFromSpreadsheet
//     and writeLicenseToPaymentSheet used to do `response.json().catch(() => ({}))`,
//     which silently swallowed the response body whenever Apps Script returned
//     something that wasn't valid JSON — exactly what happens on a Google-side
//     404/error page — leaving nothing but a bare "HTTP 404" to debug from.
//     Both calls now go through a shared callAppsScript() helper that reads the
//     raw response text first, includes it (or Apps Script's own `error` field)
//     in the thrown message, retries once on transport/HTTP-level failures, and
//     — specifically for HTTP 404 — appends a hint that PAYMENT_APPS_SCRIPT_URL
//     is likely stale (Apps Script deployment id changed/deleted) or wrong.
//     NOTE: a 404 from Google's side still means the deployment URL itself needs
//     fixing on the Apps Script side; this change makes that fully visible in the
//     next bug-report message instead of guessing from "HTTP 404" alone.
//   - CONTENT: removed the word "spreadsheet" from every buyer-facing Telegram
//     notification (payment-created, guest payment-created, /status, /help);
//     internal admin bug-report text is unaffected.
//   - CONTENT: /fitur feature list updated (Invite Team replaces Kompetitor
//     Analyzer; Filter Produk Lanjutan / Product Card Badge copy updated).
//   - PRICING: updated to 15k / 39k / 99k / 149k / 249k. The price list shown by
//     /paket (both the message text and the inline keyboard) is now generated
//     from PRICING_TIERS instead of being duplicated as hardcoded strings, so a
//     future price change can't leave one of the two out of sync.
//   - FEATURE: registers a Telegram command menu (the round "Menu" button next
//     to the attachment icon) via setMyCommands + setChatMenuButton. Apply it
//     immediately with POST /api/v1/telegram/setup-commands; it's also re-applied
//     automatically on every scheduled() cron tick as a self-healing safety net.
//     Bare "/status" (tapped from the menu, no args) now replies with usage
//     instructions instead of falling through to "perintah tidak dikenal".
//
// v2.1 HARDENING CHANGES:
//   - FIX: removed a column-mapping bug class shared with the Apps Script
//     side (see PaymentShopeeProductHunter.gs) that could make automatic
//     reconciliation silently never match a paid row.
//   - SECURITY: generateLicenseKey()/generateOrderId() now use
//     crypto.getRandomValues instead of Math.random (license keys are a
//     bearer credential and must not be guessable).
//   - SECURITY: verifyToken() now fails closed if JWT_SECRET is missing or
//     too short, instead of silently verifying against an empty key.
//   - SECURITY: Telegram webhook can now be pinned to Telegram's own
//     secret token (X-Telegram-Bot-Api-Secret-Token) via
//     TELEGRAM_WEBHOOK_SECRET, so third parties can't POST fake updates.
//   - SECURITY: user-controlled Telegram display names/usernames are HTML
//     escaped before being interpolated into parse_mode:"HTML" messages;
//     buyer email input is validated against a stricter pattern.
//   - RELIABILITY: completePayment() re-reads the payment row immediately
//     before writing, closing most of the race window between the cron
//     reconciliation and a manual "Cek Pembayaran" tap (avoids issuing a
//     duplicate license row for the same order).
//   - RELIABILITY: repeated "Bayar Sekarang" taps for the same package now
//     reuse the existing pending order instead of spawning a new one each
//     time (less clutter in `payments`, fewer duplicate admin pings).
//   - RELIABILITY: spreadsheet reconciliation failures now also trigger an
//     admin bug report (previously only logged to the Worker console,
//     where nobody would see them).
//   - CLEANUP: removed an unused/legacy CSV-based spreadsheet reader
//     (fetchSpreadsheetRows/parseCsv/sheetRowMatchesPayment and friends).
//     It was dead code — reconciliation goes exclusively through the Apps
//     Script endpoint — and it assumed the spreadsheet was publicly
//     viewable, which is an unnecessary data-exposure surface to keep
//     around unused.
//   - Internal error details are no longer echoed back to HTTP/Telegram
//     callers; they're logged and sent to the admin bot instead.
//
// v2.0 CHANGES:
//   - MIGRATION: Pakasir removed; payment uses fixed Lynk.id checkout links per package
//   - CONFIRMATION: no payment webhook; status reconciles against Google Spreadsheet
//   - PRICING: updated to 19k / 59k / 159k / 239k / 359k
//
// v1.4 CHANGES:
//   - FIX: Guest payment now INSERT license ke licenses table (owner=NULL, status='unused')
//     Supaya saat user activate di extension, worker extension bisa find + update owner
//     Sebelumnya: license hanya disimpan di payments.license_key → aktivasi gagal di extension
//
// v1.3: Supabase Auth API + stateful conversation + /akun + /license + /logout
// v1.2: Inline menu + guest payment + 10 commands + support group
// v1.1: 2 Telegram bots (User + Admin) + 5 pricing tiers + Telegram webhook handler
// v1.0: Initial release — Pakasir + Telegram + Supabase
//
// ENVIRONMENT VARIABLES / WRANGLER SETTINGS
// ============================================================
// REQUIRED — Supabase:
//   SUPABASE_URL
//   SUPABASE_SERVICE_ROLE_KEY
//   SUPABASE_ANON_KEY
//
// REQUIRED — Authentication:
//   JWT_SECRET
//     - minimum 16 characters; use a long random secret.
//
// REQUIRED — Google Apps Script / Payment reconciliation:
//   PAYMENT_APPS_SCRIPT_URL
//     - CURRENT Google Apps Script Web App /exec URL.
//     - Example: https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
//   PAYMENT_APPS_SCRIPT_SECRET
//     - Must match BOTH Apps Script Script Properties:
//         API_SECRET
//         WORKER_RECONCILE_PING_SECRET
//
// RECOMMENDED — Telegram webhook security:
//   TELEGRAM_WEBHOOK_SECRET
//     - Must match the secret_token configured in Telegram setWebhook.
//
// REQUIRED — Telegram user bot:
//   TELEGRAM_USER_BOT_TOKEN
//   TELEGRAM_USER_BOT_USERNAME
//
// REQUIRED — Telegram admin bot:
//   TELEGRAM_ADMIN_BOT_TOKEN
//   TELEGRAM_ADMIN_CHAT_ID
//
// OPTIONAL:
//   LICENSE_KEY_PREFIX   (default: SPH)
//   SUPPORT_GROUP_URL    (default: https://t.me/+tDtquz8U_fUzMTI9)
//
// CANONICAL PAYMENT SPREADSHEET:
//   PAYMENT_SPREADSHEET_ID is intentionally hardcoded below and must match
//   PaymentShopeeProductHunter.gs:
//   1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64
//
// IMPORTANT SECRET RELATIONSHIP:
//   Worker PAYMENT_APPS_SCRIPT_SECRET
//              = Apps Script API_SECRET
//              = Apps Script WORKER_RECONCILE_PING_SECRET
//
// WRANGLER COMMANDS (run locally; NEVER commit the real secret values):
//   npx wrangler secret put SUPABASE_URL
//   npx wrangler secret put SUPABASE_SERVICE_ROLE_KEY
//   npx wrangler secret put SUPABASE_ANON_KEY
//   npx wrangler secret put JWT_SECRET
//   npx wrangler secret put PAYMENT_APPS_SCRIPT_URL
//   npx wrangler secret put PAYMENT_APPS_SCRIPT_SECRET
//   npx wrangler secret put TELEGRAM_WEBHOOK_SECRET
//   npx wrangler secret put TELEGRAM_USER_BOT_TOKEN
//   npx wrangler secret put TELEGRAM_USER_BOT_USERNAME
//   npx wrangler secret put TELEGRAM_ADMIN_BOT_TOKEN
//   npx wrangler secret put TELEGRAM_ADMIN_CHAT_ID
//   npx wrangler secret put LICENSE_KEY_PREFIX
//   npx wrangler secret put SUPPORT_GROUP_URL
//
// Do not put real API keys/tokens/passwords directly in this .js file.
// ============================================================

// ENDPOINTS:
//   POST /api/v1/payment/create           — Create new payment (JWT auth)
//   GET  /api/v1/payment/status/:order_id — Cek status payment
//   GET  /api/v1/payments/history         — History payment user (JWT auth)
//   POST /api/v1/payment/reconcile-ping   — Apps Script cron ping (X-Apps-Script-Secret) → reconcile now
//   GET  /api/v1/web/plans                — Landing page: price list
//   POST /api/v1/web/checkout             — Landing page: create order {plan_duration_days, email}
//   GET  /api/v1/web/status/:order_id     — Landing page: poll status (?email=) → license when paid
//   POST /api/v1/telegram/user-bot        — Telegram webhook for user bot
//   POST /api/v1/telegram/setup-commands  — Register the Telegram command menu (run once after deploy)
//   GET  /health                          — Health check
// ============================================================

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Api-Key, X-Secret',
  'Access-Control-Max-Age': '86400'
};

const TELEGRAM_BASE_URL = 'https://api.telegram.org';
const SUPPORT_GROUP_URL = 'https://t.me/+tDtquz8U_fUzMTI9';

const PRICING_TIERS = {
  7:   { amount: 15000,  name: 'Basic',   productTitle: 'Paket Basic Shopee Product Hunter',   lynk: 'https://lynk.id/trendvision/0jgwrwlwrjlm/checkout', spreadsheetId: '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64' },
  30:  { amount: 39000,  name: 'Starter', productTitle: 'Paket Starter Shopee Product Hunter', lynk: 'https://lynk.id/trendvision/e80n5066nmdp/checkout', spreadsheetId: '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64' },
  90:  { amount: 99000,  name: 'Growth',  productTitle: 'Paket Growth Shopee Product Hunter',  lynk: 'https://lynk.id/trendvision/oyzkjz5y2pjg/checkout', spreadsheetId: '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64' },
  180: { amount: 149000, name: 'Pro',     productTitle: 'Paket Pro Shopee Product Hunter',     lynk: 'https://lynk.id/trendvision/1n2o0rx81zyx/checkout', spreadsheetId: '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64' },
  365: { amount: 249000, name: 'Annual',  productTitle: 'Paket Annual Shopee Product Hunter',  lynk: 'https://lynk.id/trendvision/xz7rxk93ye6x/checkout', spreadsheetId: '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64' }
};

const ALLOWED_PAYMENT_METHODS = ['lynk', 'payment_link']; // payment_link kept as backward-compatible alias
const PAYMENT_METHOD_LABELS = { lynk: '💳 Lynk.id', payment_link: '💳 Lynk.id' };
const PAYMENT_SPREADSHEET_ID = '1svvmhCfkSgOCgrCsvSFWkZsjOIsf6b2LahKhbWdWW64';

// ---- v2.4: web checkout helpers ------------------------------------------
const WEB_GUEST_PREFIX = 'web:';
const PENDING_RECONCILE_WINDOW_MS = 3 * 24 * 60 * 60 * 1000; // only sweep orders newer than 3 days
const WEB_CHECKOUT_LIMIT = { max: 10, windowMs: 10 * 60 * 1000 }; // per IP, per isolate (best-effort)
const WEB_STATUS_MIN_INTERVAL_MS = 6000;                          // per order: max 1 Apps Script lookup / 6s
const _webCheckoutHits = new Map(); // ip -> [timestamps]
const _webStatusLast = new Map();   // order_id -> last lookup ms

function webGuestId(email) { return WEB_GUEST_PREFIX + String(email || '').trim().toLowerCase(); }
function isWebGuest(id) { return String(id || '').startsWith(WEB_GUEST_PREFIX); }
function webGuestEmail(id) { return isWebGuest(id) ? String(id).slice(WEB_GUEST_PREFIX.length) : null; }
function clientIp(request) { return request.headers.get('CF-Connecting-IP') || 'unknown'; }

function rateLimited(map, key, { max, windowMs }) {
  const now = Date.now();
  const hits = (map.get(key) || []).filter(t => now - t < windowMs);
  hits.push(now);
  map.set(key, hits);
  if (map.size > 2000) { for (const [k, v] of map) if (!v.length || now - v[v.length - 1] > windowMs) map.delete(k); }
  return hits.length > max;
}
// Struktur spreadsheet Payment: A=Judul Barang, B=Harga, C=Tanggal, D=Status,
// E=Buyer Email, F=Ref, G=Lisensi Code, H=Email Terkirim.
// Apps Script (lookup_payment + fulfill_license) is the single source of
// truth for this layout — see PaymentShopeeProductHunter.gs. The Worker
// never reads the spreadsheet directly; it always goes through Apps Script.

// ============================================================
// HELPERS
// ============================================================

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS }
  });
}

function errorResponse(code, message, status = 400) {
  return jsonResponse({ success: false, error: { code, message } }, status);
}

// Internal exception text can contain things we don't want to hand to an
// HTTP caller or a Telegram user — Supabase URLs, stack fragments, etc.
// Callers should console.error()/sendBugReport() the real error and use
// this generic string for anything user-facing.
function safeErrorMessage() {
  return 'Terjadi kesalahan pada server. Tim kami sudah diberitahu, silakan coba lagi.';
}

// Minimal HTML escaping for values interpolated into Telegram
// parse_mode:"HTML" messages. Telegram display names/usernames are
// attacker-controlled (anyone can set their own name), so anything derived
// from them must be escaped before going into an HTML message.
function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function isValidEmail(email) {
  const value = String(email || '').trim();
  if (!value || value.length > 254) return false;
  return /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,24}$/.test(value);
}

// crypto.getRandomValues-backed random string — used anywhere the output
// doubles as a credential (license keys) or should be hard to predict
// (order ids). Math.random() is NOT a CSPRNG and shouldn't back either.
function secureRandomChars(length, alphabet) {
  const bytes = new Uint32Array(length);
  crypto.getRandomValues(bytes);
  let out = '';
  for (let i = 0; i < length; i++) out += alphabet[bytes[i] % alphabet.length];
  return out;
}

function generateOrderId() {
  const now = new Date();
  const dateStr = now.getUTCFullYear().toString() +
    String(now.getUTCMonth() + 1).padStart(2, '0') +
    String(now.getUTCDate()).padStart(2, '0');
  const random = secureRandomChars(6, '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ');
  return `INV-${dateStr}-${random}`;
}

function generateLicenseKey(prefix = 'SPH') {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  function block() { return secureRandomChars(4, chars); }
  return `${prefix}-${block()}-${block()}-${block()}-${block()}`;
}

async function verifyToken(env, accessToken) {
  if (!accessToken) return { valid: false, error: { code: 'SESSION_INVALID', message: 'No access token' } };

  const secret = env.JWT_SECRET || '';
  if (secret.length < 16) {
    // Fail closed: an unset/too-short JWT secret would mean anyone can
    // forge a token that verifies successfully. Refuse everything instead
    // of silently importing an empty-string HMAC key.
    console.error('[Auth] JWT_SECRET is missing or too short (<16 chars) — refusing all tokens.');
    return { valid: false, error: { code: 'CONFIG_ERROR', message: 'Server auth is not configured correctly' } };
  }

  try {
    const parts = accessToken.split('.');
    if (parts.length !== 3) return { valid: false, error: { code: 'INVALID_TOKEN', message: 'Malformed JWT' } };
    const [headerB64, payloadB64, signatureB64] = parts;

    let header = {};
    try { header = JSON.parse(atob(headerB64.replace(/-/g, '+').replace(/_/g, '/'))); } catch (_) {}
    if (header.alg !== 'HS256') return { valid: false, error: { code: 'INVALID_TOKEN', message: 'Unsupported algorithm' } };

    const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
    const signature = Uint8Array.from(atob(signatureB64.replace(/-/g, '+').replace(/_/g, '/')), c => c.charCodeAt(0));
    const data = new TextEncoder().encode(`${headerB64}.${payloadB64}`);
    const valid = await crypto.subtle.verify('HMAC', key, signature, data);
    if (!valid) return { valid: false, error: { code: 'INVALID_TOKEN', message: 'Signature failed' } };
    const payload = JSON.parse(atob(payloadB64.replace(/-/g, '+').replace(/_/g, '/')));
    if (payload.exp && Date.now() >= payload.exp * 1000) return { valid: false, error: { code: 'SESSION_EXPIRED', message: 'Token expired' } };
    return { valid: true, user: { id: payload.sub, email: payload.email } };
  } catch (e) {
    return { valid: false, error: { code: 'INVALID_TOKEN', message: String(e) } };
  }
}

async function verifyTokenFromRequest(env, request) {
  const authHeader = request.headers.get('Authorization') || '';
  return verifyToken(env, authHeader.replace('Bearer ', '').trim());
}

async function supabaseFetch(env, path, options = {}) {
  const url = (env.SUPABASE_URL || '') + path;
  const fetchOptions = {
    method: options.method || 'GET',
    headers: { 'Content-Type': 'application/json', 'apikey': env.SUPABASE_SERVICE_ROLE_KEY || '', 'Authorization': 'Bearer ' + (env.SUPABASE_SERVICE_ROLE_KEY || '') }
  };
  if (options.body) fetchOptions.body = JSON.stringify(options.body);
  const response = await fetch(url, fetchOptions);
  const data = await response.json().catch(() => ({}));
  return { ok: response.ok, status: response.status, data };
}

function parseSheetDate(value) {
  const raw = String(value ?? '').trim(); if (!raw) return null;
  const m = raw.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    const day = Number(m[1]), month = Number(m[2]), year = Number(m[3]);
    const hour = Number(m[4] || 0), minute = Number(m[5] || 0), second = Number(m[6] || 0);
    // Spreadsheet timestamp is Indonesian local time (WIB / UTC+7).
    return new Date(Date.UTC(year, month - 1, day, hour - 7, minute, second));
  }
  const parsed = Date.parse(raw); return Number.isFinite(parsed) ? new Date(parsed) : null;
}

// Shared caller for the Payment Apps Script endpoint. Both lookup_payment
// and fulfill_license go through here so they get the same diagnostics and
// retry behavior.
//
// IMPORTANT: previously each call site did
//   const data = await response.json().catch(() => ({}));
// which silently threw away the response body whenever it wasn't valid
// JSON — exactly the case for Google's own error pages (deployment not
// found, access denied, etc). That's why a failure only ever surfaced as
// a bare "HTTP 404" with zero context. We now always read the body as text
// first, so whatever Google actually said comes through in the bug report.
async function callAppsScript(env, payload, { retries = 1 } = {}) {
  const endpoint = String(env.PAYMENT_APPS_SCRIPT_URL || '').trim();
  if (!endpoint) return { ok: false, status: 0, data: {}, raw: '', error: 'PAYMENT_APPS_SCRIPT_URL belum dikonfigurasi' };

  let last;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const raw = await response.text();
      let data = {};
      try { data = JSON.parse(raw); } catch (_) {}
      if (response.ok && data.success === true) return { ok: true, status: response.status, data, raw };
      last = { ok: false, status: response.status, data, raw };
      // A clean "success:false" business response (e.g. invalid ref) won't
      // change on retry — only retry actual transport/HTTP-level failures.
      if (response.ok && data.success === false) return last;
    } catch (e) {
      last = { ok: false, status: 0, data: {}, raw: '', error: String(e) };
    }
    if (attempt < retries) await new Promise(r => setTimeout(r, 700));
  }
  return last;
}

function appsScriptErrorMessage(result, prefix) {
  const detail = result.data?.error
    || (result.raw ? result.raw.slice(0, 300).replace(/\s+/g, ' ').trim() : '')
    || result.error
    || `HTTP ${result.status || 0}`;
  const hint = result.status === 404
    ? ' — kemungkinan PAYMENT_APPS_SCRIPT_URL salah atau deployment Apps Script sudah tidak aktif (redeploy dan perbarui secret-nya)'
    : '';
  return `${prefix} (HTTP ${result.status || 0}): ${detail}${hint}`;
}
async function reconcilePaymentFromSpreadsheet(env, payment) {
  const tier = PRICING_TIERS[payment.plan_duration_days];
  if (!tier?.spreadsheetId) return { matched: false, reason: 'spreadsheet_not_configured' };
  if (!payment._buyer_email) return { matched: false, reason: 'buyer_email_required' };

  const result = await callAppsScript(env, {
    action: 'lookup_payment',
    api_secret: String(env.PAYMENT_APPS_SCRIPT_SECRET || ''),
    spreadsheet_id: PAYMENT_SPREADSHEET_ID,
    sheet_name: 'Payment',
    product_title: tier.productTitle,
    buyer_email: payment._buyer_email,
    created_at: payment.created_at || ''
  });
  if (!result.ok) throw new Error(appsScriptErrorMessage(result, 'Apps Script lookup gagal'));
  const data = result.data;

  if (!data.matched) {
    return {
      matched: false,
      reason: data.reason || 'no_match',
      candidates: data.candidates || 0
    };
  }

  const ref = String(data.ref || '').trim();
  if (!ref) return { matched: false, reason: 'missing_ref', candidates: 1 };

  const completedDate = parseSheetDate(data.completed_at);
  return {
    matched: true,
    row: data,
    lynkRef: ref,
    completed_at: completedDate?.toISOString() || new Date().toISOString()
  };
}

async function findPaymentByLynkRef(env, lynkRef) {
  if (!lynkRef) return null;
  const result = await supabaseFetch(env, `/rest/v1/payments?txn_id=eq.${encodeURIComponent(lynkRef)}&select=order_id,status,license_key`, { method: 'GET' });
  if (!result.ok || !Array.isArray(result.data) || !result.data.length) return null;
  return result.data[0];
}

async function writeLicenseToPaymentSheet(env, { ref, licenseKey, productTitle, buyerEmail, planName }) {
  const result = await callAppsScript(env, {
    action: 'fulfill_license',
    api_secret: String(env.PAYMENT_APPS_SCRIPT_SECRET || ''),
    spreadsheet_id: PAYMENT_SPREADSHEET_ID,
    sheet_name: 'Payment',
    ref: String(ref || '').trim(),
    license_code: String(licenseKey || '').trim(),
    product_title: productTitle || '',
    buyer_email: buyerEmail || '',
    plan_name: planName || '',
    support_url: getSupportGroupUrl(env)
  });
  if (!result.ok) throw new Error(appsScriptErrorMessage(result, 'Apps Script gagal'));
  const data = result.data;

  if (data.status === 'pending') {
    throw new Error('Pembayaran di Payment sheet masih PENDING');
  }

  if (!data.license_code) {
    throw new Error('Apps Script tidak mengembalikan license_code');
  }

  return data;
}

async function completePayment(env, payment, completedAt, source = 'lynk-spreadsheet', lynkRef = null) {
  if (payment.status === 'completed' && payment.license_key) return { license_key: payment.license_key, expires_at: payment.expires_at };

  // Idempotency guard: re-read the payment row immediately before writing
  // anything. This closes most of the race window between the cron
  // reconciliation and a manual "Cek Pembayaran" tap landing at the same
  // time — if another in-flight request already completed this order, we
  // just return its result instead of generating a second license or
  // inserting a duplicate `licenses` row.
  const fresh = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(payment.order_id)}&select=status,license_key,expires_at`, { method: 'GET' });
  if (fresh.ok && Array.isArray(fresh.data) && fresh.data[0]?.status === 'completed' && fresh.data[0]?.license_key) {
    return { license_key: fresh.data[0].license_key, expires_at: fresh.data[0].expires_at };
  }

  const generatedLicenseKey = generateLicenseKey(env.LICENSE_KEY_PREFIX || 'SPH');
  const now = new Date(completedAt || new Date().toISOString());
  const expiresAt = new Date(now.getTime() + payment.plan_duration_days * 86400000).toISOString();
  const buyerEmail = payment._buyer_email || '';

  // First fulfill the single Payment spreadsheet row. Apps Script:
  // - verifies Ref + SUCCESS
  // - refuses duplicate Ref
  // - writes Payment!G (Lisensi Code)
  // - sends the license email exactly once
  // If this fails, Supabase payment stays pending so the next cron can retry.
  const sheetResult = await writeLicenseToPaymentSheet(env, {
    ref: lynkRef,
    licenseKey: generatedLicenseKey,
    productTitle: PRICING_TIERS[payment.plan_duration_days]?.productTitle || payment.plan_name || '',
    buyerEmail,
    planName: payment.plan_name || ''
  });
  const licenseKey = sheetResult.license_code;

  await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(payment.order_id)}`, { method: 'PATCH', body: {
    status: 'completed',
    completed_at: now.toISOString(),
    license_key: licenseKey,
    ...(lynkRef ? { txn_id: lynkRef } : {})
  } });

  if (!payment.user_id && payment.guest_chat_id) {
    // Apps Script always echoes back the existing license for a Ref it has
    // already fulfilled, so licenseKey here is stable across racers — but
    // guard the `licenses` insert anyway in case this exact key already
    // exists (e.g. a retried request).
    const existingLicenseRow = await supabaseFetch(env, `/rest/v1/licenses?license_key=eq.${encodeURIComponent(licenseKey)}&select=id`, { method: 'GET' });
    if (!existingLicenseRow.ok || !Array.isArray(existingLicenseRow.data) || existingLicenseRow.data.length === 0) {
      await supabaseFetch(env, '/rest/v1/licenses', { method: 'POST', body: {
        license_key: licenseKey,
        plan: 'premium',
        duration_days: payment.plan_duration_days,
        status: 'unused',
        owner_user_id: null,
        max_seats: 5,
        created_by: 'payment-worker-v2.1-lynk',
        notes: `Guest Lynk payment ${payment.order_id}, chat_id: ${payment.guest_chat_id}`
      } });
    }

    if (!isWebGuest(payment.guest_chat_id)) await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, payment.guest_chat_id, `🎉 <b>Pembayaran Berhasil!</b>\n\nTerima kasih telah membeli <b>Shopee Product Hunter</b>.\n\n📦 <b>Paket:</b> ${payment.plan_name} (${payment.plan_duration_days} hari)\n💵 <b>Amount:</b> ${formatRupiah(payment.amount)}\n🔑 <b>License Code:</b>\n<code>${licenseKey}</code>\n\n📧 License Code juga telah dikirim ke email <code>${escapeHtml(buyerEmail || '-')}</code>.\n\n<b>Panduan aktivasi Chrome Extension Shopee Product Hunter:</b>\n1. Buka Google Chrome.\n2. Buka extension <b>Shopee Product Hunter</b>.\n3. Login/Register menggunakan akun Anda.\n4. Pilih menu <b>Aktifkan License</b>.\n5. Masukkan License Code di atas.\n6. Klik <b>Aktifkan</b> dan mulai gunakan fitur extension.\n\nSimpan License Code ini dengan aman dan jangan membagikannya kepada orang lain.\n\nButuh bantuan? ${getSupportGroupUrl(env)}`);
    await sendAdminNotification(env, `💰 <b>LYNK PAYMENT SUCCESS</b>\n\n💬 Chat ID: <code>${payment.guest_chat_id}</code>\n📦 ${payment.plan_name} (${payment.plan_duration_days}d)\n💵 ${formatRupiah(payment.amount)}\n🔑 <code>${licenseKey}</code>\n📧 ${escapeHtml(buyerEmail || '-')}\n🔗 Ref: <code>${escapeHtml(lynkRef || '-')}</code>\n🆔 <code>${payment.order_id}</code>`);
    return { license_key: licenseKey, expires_at: expiresAt };
  }

  const existing = await supabaseFetch(env, `/rest/v1/licenses?owner_user_id=eq.${payment.user_id}&status=eq.active&select=id`, { method: 'GET' });
  if (!existing.ok || !Array.isArray(existing.data) || !existing.data.length) await supabaseFetch(env, '/rest/v1/licenses', { method: 'POST', body: {
    license_key: licenseKey,
    plan: 'premium',
    duration_days: payment.plan_duration_days,
    status: 'active',
    owner_user_id: payment.user_id,
    activated_at: now.toISOString(),
    expires_at: expiresAt,
    max_seats: 5,
    created_by: 'payment-worker-v2.1-lynk',
    notes: `Auto from ${payment.order_id}`
  } });
  else await supabaseFetch(env, `/rest/v1/licenses?id=eq.${existing.data[0].id}`, { method: 'PATCH', body: { expires_at: expiresAt, status: 'active' } });

  await supabaseFetch(env, `/rest/v1/subscriptions?user_id=eq.${payment.user_id}`, { method: 'PATCH', body: { plan: 'premium', status: 'active', expires_at: expiresAt } });

  // The /akun and /license cache (see FEATURE-RESPONSE CACHE below) can
  // hold up to an hour of stale "no active license" data — clear it now so
  // a user who just paid doesn't see outdated info if they check right away.
  _featureCache.delete(`akun:${payment.user_id}`);
  _featureCache.delete(`license:${payment.user_id}`);

  await sendUserNotification(env, payment.user_id, `🎉 <b>Pembayaran Berhasil!</b>\n\nTerima kasih telah membeli <b>Shopee Product Hunter</b>.\n\n📦 <b>Paket:</b> ${payment.plan_name} (${payment.plan_duration_days} hari)\n💵 <b>Amount:</b> ${formatRupiah(payment.amount)}\n🔑 <b>License Code:</b>\n<code>${licenseKey}</code>\n\n📧 License Code juga telah dikirim ke email <code>${escapeHtml(buyerEmail || '-')}</code>.\n\n<b>Panduan aktivasi Chrome Extension Shopee Product Hunter:</b>\n1. Buka Google Chrome.\n2. Buka extension <b>Shopee Product Hunter</b>.\n3. Login/Register menggunakan akun Anda.\n4. Pilih menu <b>Aktifkan License</b>.\n5. Masukkan License Code di atas.\n6. Klik <b>Aktifkan</b> dan mulai gunakan fitur extension.\n\nSimpan License Code ini dengan aman dan jangan membagikannya kepada orang lain.\n\nButuh bantuan? ${getSupportGroupUrl(env)}`);

  await sendAdminNotification(env, `💰 <b>LYNK PAYMENT SUCCESS</b>\n\n👤 ${escapeHtml(buyerEmail || 'unknown')}\n📦 ${payment.plan_name} (${payment.plan_duration_days}d)\n💵 ${formatRupiah(payment.amount)}\n🔑 <code>${licenseKey}</code>\n📧 ${escapeHtml(buyerEmail || '-')}\n🔗 Ref: <code>${escapeHtml(lynkRef || '-')}</code>\n🆔 <code>${payment.order_id}</code>`);

  return { license_key: licenseKey, expires_at: expiresAt };
}

// ============================================================
// TELEGRAM HELPERS
// ============================================================

async function telegramSendMessage(botToken, chatId, text, options = {}) {
  if (!botToken || !chatId) return { ok: false, error: 'config_missing' };
  try {
    const body = { chat_id: chatId, text, parse_mode: 'HTML', disable_web_page_preview: true };
    if (options.reply_markup) body.reply_markup = options.reply_markup;
    const response = await fetch(`${TELEGRAM_BASE_URL}/bot${botToken}/sendMessage`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || data.ok === false) console.warn('[Telegram] sendMessage rejected:', data.description || response.status);
    return data;
  } catch (e) { console.warn('[Telegram] Error:', e); return { ok: false, error: String(e) }; }
}

async function telegramAnswerCallback(botToken, callbackQueryId, text) {
  try {
    await fetch(`${TELEGRAM_BASE_URL}/bot${botToken}/answerCallbackQuery`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ callback_query_id: callbackQueryId, text: text || '' })
    });
  } catch (e) {}
}

// The list behind the round "Menu" button next to the chat's attachment
// icon (Telegram only shows that button once commands are registered via
// setMyCommands + the menu button type is set to "commands").
const USER_BOT_COMMANDS = [
  { command: 'start', description: 'Mulai / tampilkan menu utama' },
  { command: 'menu', description: 'Tampilkan menu utama' },
  { command: 'paket', description: 'Lihat daftar paket license' },
  { command: 'fitur', description: 'Informasi fitur aplikasi' },
  { command: 'akun', description: 'Info akun Anda' },
  { command: 'license', description: 'Info license Anda' },
  { command: 'riwayat', description: 'Riwayat transaksi' },
  { command: 'status', description: 'Cek status pembayaran' },
  { command: 'support', description: 'Grup support' },
  { command: 'help', description: 'Bantuan & daftar perintah' },
  { command: 'logout', description: 'Keluar dari sesi login' }
];

async function telegramSetupCommandMenu(botToken) {
  if (!botToken) return { ok: false, error: 'config_missing' };
  try {
    const cmdRes = await fetch(`${TELEGRAM_BASE_URL}/bot${botToken}/setMyCommands`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ commands: USER_BOT_COMMANDS })
    });
    const cmdData = await cmdRes.json().catch(() => ({}));

    const menuRes = await fetch(`${TELEGRAM_BASE_URL}/bot${botToken}/setChatMenuButton`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ menu_button: { type: 'commands' } })
    });
    const menuData = await menuRes.json().catch(() => ({}));

    return { ok: cmdData.ok !== false && menuData.ok !== false, setMyCommands: cmdData, setChatMenuButton: menuData };
  } catch (e) {
    console.warn('[Telegram] setupCommandMenu error:', e);
    return { ok: false, error: String(e) };
  }
}

async function sendUserNotification(env, userId, text, options = {}) {
  try {
    const result = await supabaseFetch(env, `/rest/v1/user_telegram_chats?user_id=eq.${userId}&is_blocked=eq.false&select=chat_id`, { method: 'GET' });
    if (!result.ok || !Array.isArray(result.data) || result.data.length === 0) return { ok: false, error: 'no_chat_id' };
    return await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, result.data[0].chat_id, text, options);
  } catch (e) { return { ok: false, error: String(e) }; }
}

async function sendAdminNotification(env, text, options = {}) {
  return await telegramSendMessage(env.TELEGRAM_ADMIN_BOT_TOKEN, env.TELEGRAM_ADMIN_CHAT_ID, text, options);
}

async function sendBugReport(env, error, context = {}) {
  const bugText = `🐛 <b>BUG REPORT</b>\n\n<b>Error:</b> <code>${escapeHtml(String(error).substring(0, 500))}</code>\n<b>Context:</b> <code>${escapeHtml(JSON.stringify(context).substring(0, 500))}</code>\n<b>Time:</b> ${new Date().toISOString()}`;
  return await sendAdminNotification(env, bugText);
}

function formatRupiah(amount) {
  return 'Rp ' + amount.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function formatDate(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  return d.toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'UTC' }) + ' UTC';
}

function getSupportGroupUrl(env) {
  return env.SUPPORT_GROUP_URL || SUPPORT_GROUP_URL;
}

// ============================================================
// v1.3: SUPABASE AUTH API + SESSION MANAGEMENT
// ============================================================

const AUTH_FLOW_TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes for auth flow

/**
 * Verify email+password via Supabase Auth API (anon key)
 * Returns: { valid, user, session } or { valid: false, error }
 */
async function supabaseAuthVerify(env, email, password) {
  try {
    const url = (env.SUPABASE_URL || '') + '/auth/v1/token?grant_type=password';
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': env.SUPABASE_ANON_KEY || ''
      },
      body: JSON.stringify({ email, password })
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || !data.access_token) {
      return { valid: false, error: data.error_description || data.error?.message || 'Invalid credentials' };
    }
    return {
      valid: true,
      user: { id: data.user?.id, email: data.user?.email },
      session: { access_token: data.access_token, refresh_token: data.refresh_token }
    };
  } catch (e) {
    return { valid: false, error: String(e) };
  }
}

/**
 * Get auth session from bot_auth_sessions table
 */
async function getAuthSession(env, chatId) {
  const result = await supabaseFetch(env, `/rest/v1/bot_auth_sessions?chat_id=eq.${encodeURIComponent(chatId)}&select=*`, { method: 'GET' });
  if (!result.ok || !Array.isArray(result.data) || result.data.length === 0) return null;
  return result.data[0];
}

/**
 * Save auth session (upsert)
 */
async function saveAuthSession(env, chatId, sessionData) {
  // Try insert first
  const insertResult = await supabaseFetch(env, '/rest/v1/bot_auth_sessions', {
    method: 'POST',
    body: { chat_id: chatId, ...sessionData, updated_at: new Date().toISOString() }
  });
  if (!insertResult.ok) {
    // Conflict — update existing
    await supabaseFetch(env, `/rest/v1/bot_auth_sessions?chat_id=eq.${encodeURIComponent(chatId)}`, {
      method: 'PATCH',
      body: { ...sessionData, updated_at: new Date().toISOString() }
    });
  }
}

/**
 * Delete auth session
 */
async function deleteAuthSession(env, chatId) {
  await supabaseFetch(env, `/rest/v1/bot_auth_sessions?chat_id=eq.${encodeURIComponent(chatId)}`, { method: 'DELETE' });
}

/**
 * Check if auth flow session is expired (5 min for awaiting_email/awaiting_password)
 */
function isAuthFlowExpired(session) {
  if (!session) return true;
  if (session.step === 'authenticated' || session.step === 'payment_email_saved') return false; // Persistent until replaced/cleared
  const age = Date.now() - new Date(session.updated_at).getTime();
  return age > AUTH_FLOW_TIMEOUT_MS;
}

/**
 * Delete a Telegram message (for password security)
 */
async function telegramDeleteMessage(botToken, chatId, messageId) {
  try {
    await fetch(`${TELEGRAM_BASE_URL}/bot${botToken}/deleteMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: chatId, message_id: messageId })
    });
  } catch (e) {}
}

/**
 * Get team members info for a license owner
 */
async function getTeamMembersInfo(env, userId) {
  // Get user's active license
  const licResult = await supabaseFetch(env, `/rest/v1/licenses?owner_user_id=eq.${userId}&status=eq.active&select=*`, { method: 'GET' });
  if (!licResult.ok || !Array.isArray(licResult.data) || licResult.data.length === 0) {
    return { hasLicense: false, members: [], maxSeats: 0 };
  }
  const license = licResult.data[0];

  // Get team members
  const membersResult = await supabaseFetch(env, `/rest/v1/team_members?license_id=eq.${license.id}&select=*,member_user_id&order=created_at.asc`, { method: 'GET' });
  const members = (membersResult.ok && Array.isArray(membersResult.data)) ? membersResult.data : [];

  // Get emails for each member
  const memberEmails = [];
  for (const m of members) {
    if (m.status === 'revoked') continue;
    const profileResult = await supabaseFetch(env, `/rest/v1/profiles?id=eq.${m.member_user_id}&select=email`, { method: 'GET' });
    const email = (profileResult.ok && profileResult.data?.[0]?.email) || 'unknown';
    memberEmails.push({ email, status: m.status });
  }

  const activeCount = memberEmails.filter(m => m.status === 'active').length;
  const invitedCount = memberEmails.filter(m => m.status === 'invited').length;
  const usedSeats = activeCount + invitedCount;
  const unusedSeats = (license.max_seats - 1) - usedSeats; // max_seats includes owner

  return {
    hasLicense: true,
    license,
    members: memberEmails,
    maxSeats: license.max_seats,
    activeCount,
    invitedCount,
    unusedSeats: Math.max(0, unusedSeats)
  };
}

// ============================================================
// INLINE KEYBOARDS
// ============================================================

function getMainMenuKeyboard() {
  return JSON.stringify({
    inline_keyboard: [
      [{ text: '💎 Beli License', callback_data: 'beli' }, { text: '👤 Akun Saya', callback_data: 'akun' }],
      [{ text: '📄 Riwayat Transaksi', callback_data: 'riwayat' }, { text: '🔑 License Saya', callback_data: 'license' }],
      [{ text: '📋 Informasi Fitur', callback_data: 'fitur' }, { text: '💬 Grup Support', callback_data: 'support' }],
      [{ text: '❓ Bantuan', callback_data: 'bantuan' }]
    ]
  });
}

function getPaketKeyboard() {
  const rows = Object.entries(PRICING_TIERS).map(([days, t]) =>
    [{ text: `${t.name} — ${days} hari — ${formatRupiah(t.amount)}`, callback_data: `paket_${days}` }]
  );
  rows.push([{ text: '⬅️ Kembali ke Menu', callback_data: 'menu' }]);
  return JSON.stringify({ inline_keyboard: rows });
}

function getPaymentMethodKeyboard(paketDays) {
  return JSON.stringify({ inline_keyboard: [
    [{ text: '💳 Bayar via Lynk.id', callback_data: `pay_${paketDays}_lynk` }],
    [{ text: '⬅️ Kembali ke Paket', callback_data: 'beli' }]
  ] });
}

function getBayarKeyboard(paymentLink) {
  if (paymentLink) {
    return JSON.stringify({
      inline_keyboard: [
        [{ text: '💳 Bayar Sekarang', url: paymentLink }],
        [{ text: '⬅️ Kembali ke Menu', callback_data: 'menu' }]
      ]
    });
  }
  return JSON.stringify({
    inline_keyboard: [
      [{ text: '⬅️ Kembali ke Menu', callback_data: 'menu' }]
    ]
  });
}

// ============================================================
// HANDLER: POST /api/v1/payment/create (JWT auth — for extension)
// ============================================================
async function createLynkPaymentRecord(env, { userId = null, guestChatId = null, buyerEmail = null, planDurationDays, orderId }) {
  const tier = PRICING_TIERS[planDurationDays]; const txnId = `LYNK-${orderId}`;
  const insert = await supabaseFetch(env, '/rest/v1/payments', { method: 'POST', body: { user_id: userId, order_id: orderId, txn_id: txnId, amount: tier.amount, payment_method: 'lynk', status: 'pending', payment_link: tier.lynk, qr_string: null, va_number: null, fee: null, total_payment: tier.amount, expired_at: null, plan_duration_days: planDurationDays, plan_name: tier.name, is_sandbox: false, pakasir_project: null, telegram_notif_sent: false, guest_chat_id: guestChatId } });
  if (!insert.ok) throw new Error(`Gagal menyimpan payment: ${JSON.stringify(insert.data)}`); return { txnId };
}

// Returns an existing PENDING guest order for the same chat + package, if
// any, so repeated "Bayar Sekarang" taps don't spawn a new payments row
// (and a new admin ping) every single time.
async function findReusablePendingGuestPayment(env, chatId, planDurationDays) {
  const result = await supabaseFetch(
    env,
    `/rest/v1/payments?guest_chat_id=eq.${encodeURIComponent(chatId)}&plan_duration_days=eq.${planDurationDays}&payment_method=eq.lynk&status=eq.pending&order=created_at.desc&limit=1&select=order_id`,
    { method: 'GET' }
  );
  if (result.ok && Array.isArray(result.data) && result.data.length > 0) return result.data[0];
  return null;
}

async function handlePaymentCreate(env, request) {
  try {
    const verification = await verifyTokenFromRequest(env, request); if (!verification.valid) return errorResponse(verification.error.code, verification.error.message, 401);
    const user = verification.user;

    let body;
    try { body = await request.json(); } catch (_) { return errorResponse('VALIDATION_ERROR', 'Body harus berupa JSON yang valid', 400); }

    const planDurationDays = parseInt(body.plan_duration_days, 10); const requestedMethod = body.payment_method || 'lynk';
    if (!PRICING_TIERS[planDurationDays]) return errorResponse('INVALID_PLAN', `plan_duration_days harus: ${Object.keys(PRICING_TIERS).join(', ')}`, 400);
    if (!ALLOWED_PAYMENT_METHODS.includes(requestedMethod)) return errorResponse('INVALID_METHOD', 'payment_method tidak valid; gunakan lynk/payment_link', 400);

    const rawBuyerEmail = String(body.buyer_email || user.email || '').trim();
    if (rawBuyerEmail && !isValidEmail(rawBuyerEmail)) return errorResponse('INVALID_EMAIL', 'buyer_email tidak valid', 400);
    const buyerEmail = rawBuyerEmail || null;

    const tier = PRICING_TIERS[planDurationDays]; const orderId = generateOrderId();
    await createLynkPaymentRecord(env, { userId: user.id, buyerEmail, planDurationDays, orderId });
    const telegramDeepLink = `https://t.me/${env.TELEGRAM_USER_BOT_USERNAME || 'PaymentShopeeProductHunter_bot'}?start=pay_${orderId}`;
    await sendAdminNotification(env, `🆕 <b>NEW LYNK PAYMENT (Extension)</b>\n\n👤 ${escapeHtml(buyerEmail || 'unknown')}\n📦 ${tier.name} (${planDurationDays}d)\n💵 ${formatRupiah(tier.amount)}\n💳 Lynk.id\n🆔 <code>${orderId}</code>`);
    await sendUserNotification(env, user.id, `👋 <b>Pembayaran Dibuat</b>\n\n📦 ${tier.name} (${planDurationDays} hari)\n💵 ${formatRupiah(tier.amount)}\n💳 Lynk.id\n🆔 <code>${orderId}</code>\n\nSetelah membayar, status akan diverifikasi otomatis oleh sistem kami.`, { reply_markup: JSON.stringify({ inline_keyboard: [[{ text: '💳 Bayar Sekarang', url: tier.lynk }], [{ text: '🔄 Cek Pembayaran', callback_data: `check_${orderId}` }]] }) });
    return jsonResponse({ success: true, order_id: orderId, txn_id: `LYNK-${orderId}`, amount: tier.amount, plan: tier.name, plan_duration_days: planDurationDays, payment_method: 'lynk', payment_link: tier.lynk, qr_string: null, va_number: null, fee: 0, total_payment: tier.amount, expired_at: null, is_sandbox: false, telegram_bot_url: telegramDeepLink });
  } catch (e) { console.error('[PaymentCreate] Error:', e); await sendBugReport(env, e, { handler: 'payment_create' }); return errorResponse('INTERNAL_ERROR', safeErrorMessage(), 500); }
}

async function reconcilePaymentByOrder(env, orderId, buyerEmail = null) {
  const find = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&select=*`, { method: 'GET' });
  if (!find.ok || !Array.isArray(find.data) || !find.data.length) return { ok: false, code: 'PAYMENT_NOT_FOUND' };
  const payment = find.data[0];
  if (payment.status === 'completed') return { ok: true, payment, reconciled: true };
  const webEmail = webGuestEmail(payment.guest_chat_id);
  if (webEmail) payment._buyer_email = webEmail; // web orders: the email typed on the landing page
  else if (buyerEmail) payment._buyer_email = buyerEmail;
  try {
    const result = await reconcilePaymentFromSpreadsheet(env, payment);
    if (!result.matched) return { ok: true, payment, reconciled: false, reason: result.reason, candidates: result.candidates || 0 };
    const used = await findPaymentByLynkRef(env, result.lynkRef);
    if (used && used.order_id !== payment.order_id) return { ok: true, payment, reconciled: false, reason: 'ref_already_used' };
    const completed = await completePayment(env, payment, result.completed_at, 'lynk-spreadsheet', result.lynkRef);
    payment.status = 'completed'; payment.license_key = completed.license_key; payment.completed_at = result.completed_at; payment.expires_at = completed.expires_at; payment.txn_id = result.lynkRef; payment.lynk_ref = result.lynkRef;
    return { ok: true, payment, reconciled: true };
  } catch (e) {
    console.error('[SpreadsheetSync] Error:', orderId, e);
    // Previously this only went to the Worker console, where nobody would
    // ever see it — now the admin also gets pinged so a broken Apps
    // Script/Sheets integration doesn't fail silently for days.
    await sendBugReport(env, e, { handler: 'reconcile_payment', order_id: orderId });
    return { ok: false, code: 'SPREADSHEET_ERROR', message: safeErrorMessage() };
  }
}

// ============================================================
// HANDLER: GET /api/v1/payment/status/:order_id
// ============================================================
async function handlePaymentStatus(env, request, url) {
  const verification = await verifyTokenFromRequest(env, request); if (!verification.valid) return errorResponse(verification.error.code, verification.error.message, 401);
  const orderId = decodeURIComponent(url.pathname.split('/').pop() || '');
  const result = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&user_id=eq.${verification.user.id}&select=*`, { method: 'GET' });
  if (!result.ok || !Array.isArray(result.data) || !result.data.length) return errorResponse('NOT_FOUND', 'Not found', 404);
  const reconciled = await reconcilePaymentByOrder(env, orderId, verification.user.email || null);
  if (!reconciled.ok) return errorResponse(reconciled.code, reconciled.message || safeErrorMessage(), 502);
  return jsonResponse({ success: true, reconciled: !!reconciled.reconciled, payment: reconciled.payment });
}


// ============================================================
// HANDLER: POST /api/v1/payment/reconcile-ping
// Called by the Apps Script time-driven trigger (see
// PaymentShopeeProductHunter.gs, checkPendingPaymentsAndReport_(), which
// runs roughly every minute) whenever it notices a Payment sheet row that's
// SUCCESS but not yet fulfilled (empty Lisensi Code). Rather than have Apps
// Script replicate the Supabase-matching/fulfillment logic itself, it just
// pings the Worker so the *existing* reconcilePendingPayments() sweep — the
// same one the 5-minute cron runs — happens immediately. This is what turns
// "check every ~1 minute" into "buyer gets their license within ~1 minute"
// instead of waiting for the next scheduled tick.
// ============================================================
async function handlePaymentReconcilePing(env, request) {
  const secret = String(env.PAYMENT_APPS_SCRIPT_SECRET || '');
  const provided = request.headers.get('X-Apps-Script-Secret') || '';
  if (!secret || provided !== secret) return errorResponse('UNAUTHORIZED', 'Secret tidak valid', 401);

  let body = {};
  try { body = await request.json(); } catch (_) {}

  const result = await reconcilePendingPayments(env);
  console.log('[AppsScriptPing] reconcile triggered', { reported: body.pending_refs?.length ?? body.count ?? null, ...result });
  return jsonResponse({ success: true, ...result });
}

// ============================================================
// v2.4 — WEB (LANDING PAGE) CHECKOUT — public, no JWT
// Flow: plans → checkout (email + plan → pending order + Lynk link) →
//       buyer pays on Lynk.id with the SAME email → status polling
//       reconciles via Apps Script → license is returned + e-mailed.
// ============================================================
async function handleWebPlans() {
  return jsonResponse({
    success: true,
    plans: Object.entries(PRICING_TIERS).map(([days, t]) => ({
      duration_days: parseInt(days, 10),
      name: t.name,
      amount: t.amount,
      amount_formatted: formatRupiah(t.amount)
    }))
  });
}

async function handleWebCheckout(env, request) {
  try {
    if (rateLimited(_webCheckoutHits, clientIp(request), WEB_CHECKOUT_LIMIT)) {
      return errorResponse('RATE_LIMITED', 'Terlalu banyak percobaan. Coba lagi dalam beberapa menit.', 429);
    }
    let body;
    try { body = await request.json(); } catch (_) { return errorResponse('VALIDATION_ERROR', 'Body harus berupa JSON yang valid', 400); }
    if (body.website) return errorResponse('VALIDATION_ERROR', 'Permintaan ditolak', 400); // honeypot

    const planDurationDays = parseInt(body.plan_duration_days, 10);
    const tier = PRICING_TIERS[planDurationDays];
    if (!tier) return errorResponse('INVALID_PLAN', `plan_duration_days harus: ${Object.keys(PRICING_TIERS).join(', ')}`, 400);

    const email = String(body.email || '').trim().toLowerCase();
    if (!isValidEmail(email)) return errorResponse('INVALID_EMAIL', 'Format email tidak valid', 400);

    const guestId = webGuestId(email);
    // Same email + same package while an order is still pending → reuse it.
    const reusable = await findReusablePendingGuestPayment(env, guestId, planDurationDays);
    let orderId;
    if (reusable) {
      orderId = reusable.order_id;
    } else {
      orderId = generateOrderId();
      await createLynkPaymentRecord(env, { guestChatId: guestId, buyerEmail: email, planDurationDays, orderId });
      await sendAdminNotification(env, `🆕 <b>WEB LYNK PAYMENT CREATED</b>\n\n📧 ${escapeHtml(email)}\n📦 ${tier.name} (${planDurationDays}d)\n💵 ${formatRupiah(tier.amount)}\n💳 Lynk.id\n🆔 <code>${orderId}</code>`);
    }

    return jsonResponse({
      success: true,
      order_id: orderId,
      reused: !!reusable,
      buyer_email: email,
      plan: tier.name,
      plan_duration_days: planDurationDays,
      amount: tier.amount,
      amount_formatted: formatRupiah(tier.amount),
      payment_link: tier.lynk
    });
  } catch (e) {
    console.error('[WebCheckout] Error:', e);
    await sendBugReport(env, e, { handler: 'web_checkout' });
    return errorResponse('INTERNAL_ERROR', safeErrorMessage(), 500);
  }
}

async function handleWebStatus(env, request, url) {
  try {
    const orderId = decodeURIComponent(url.pathname.split('/').pop() || '').trim();
    const email = String(url.searchParams.get('email') || '').trim().toLowerCase();
    if (!/^INV-\d{8}-[0-9A-Z]{6}$/.test(orderId) || !isValidEmail(email)) {
      return errorResponse('VALIDATION_ERROR', 'order_id atau email tidak valid', 400);
    }

    // order_id AND the e-mail it was created with must both match — the license
    // key is a bearer credential, so an order_id alone must never be enough.
    const find = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&guest_chat_id=eq.${encodeURIComponent(webGuestId(email))}&select=*`, { method: 'GET' });
    if (!find.ok || !Array.isArray(find.data) || !find.data.length) return errorResponse('NOT_FOUND', 'Pesanan tidak ditemukan', 404);
    let payment = find.data[0];
    let checked = false;

    if (payment.status !== 'completed') {
      const now = Date.now();
      const last = _webStatusLast.get(orderId) || 0;
      if (now - last >= WEB_STATUS_MIN_INTERVAL_MS) {
        _webStatusLast.set(orderId, now);
        if (_webStatusLast.size > 2000) for (const [k, t] of _webStatusLast) if (now - t > 600000) _webStatusLast.delete(k);
        const reconciled = await reconcilePaymentByOrder(env, orderId, email);
        if (!reconciled.ok) return errorResponse(reconciled.code || 'SPREADSHEET_ERROR', reconciled.message || safeErrorMessage(), 502);
        payment = reconciled.payment;
        checked = true;
      }
    }

    const completed = payment.status === 'completed' && !!payment.license_key;
    return jsonResponse({
      success: true,
      order_id: orderId,
      status: completed ? 'completed' : 'pending',
      checked,
      plan: payment.plan_name,
      plan_duration_days: payment.plan_duration_days,
      amount: payment.amount,
      amount_formatted: formatRupiah(payment.amount),
      ...(completed ? { license_key: payment.license_key, completed_at: payment.completed_at || null } : {})
    });
  } catch (e) {
    console.error('[WebStatus] Error:', e);
    await sendBugReport(env, e, { handler: 'web_status' });
    return errorResponse('INTERNAL_ERROR', safeErrorMessage(), 500);
  }
}

// ============================================================
// HANDLER: GET /api/v1/payments/history
// ============================================================
async function handlePaymentsHistory(env, request) {
  const verification = await verifyTokenFromRequest(env, request);
  if (!verification.valid) return errorResponse(verification.error.code, verification.error.message, 401);
  const result = await supabaseFetch(env, `/rest/v1/payments?user_id=eq.${verification.user.id}&order=created_at.desc&limit=50&select=*`, { method: 'GET' });
  return jsonResponse({ success: true, count: (result.data || []).length, payments: result.data || [] });
}

// ============================================================
// HANDLER: POST /api/v1/telegram/user-bot — Telegram Webhook
// ============================================================
async function handleTelegramUserBotWebhook(env, request) {
  try {
    // If TELEGRAM_WEBHOOK_SECRET is configured, only accept requests that
    // carry Telegram's own secret token (set via setWebhook's
    // `secret_token` parameter). Without this, anyone who finds the
    // webhook URL can POST fake Telegram updates. Left optional (rather
    // than required) so an existing deployment that hasn't configured the
    // secret yet doesn't suddenly lose its webhook.
    const configuredSecret = env.TELEGRAM_WEBHOOK_SECRET || '';
    if (configuredSecret) {
      const providedSecret = request.headers.get('X-Telegram-Bot-Api-Secret-Token') || '';
      if (providedSecret !== configuredSecret) {
        console.warn('[TG Webhook] Rejected request with invalid/missing secret token');
        return new Response('Unauthorized', { status: 401 });
      }
    }

    const update = await request.json().catch(() => null);
    if (!update) return new Response('OK', { status: 200 });

    // Handle callback query (inline button clicks)
    if (update.callback_query) {
      return await handleCallbackQuery(env, update.callback_query);
    }

    // Handle message
    if (!update.message) return new Response('OK', { status: 200 });
    const msg = update.message;
    const chatId = String(msg.chat.id);
    const text = msg.text || '';
    const username = msg.from?.username || null;
    const firstName = msg.from?.first_name || '';
    const messageId = msg.message_id;

    // v1.3: Check auth session state FIRST (before command processing)
    const session = await getAuthSession(env, chatId);

    if (session && !isAuthFlowExpired(session) && /^awaiting_payment_email_\d+$/.test(String(session.step || ''))) {
      const days = parseInt(String(session.step).split('_').pop(), 10);
      return await handlePaymentEmailInput(env, chatId, text, messageId, days);
    }

    // If in auth flow, handle email/password input
    if (session && !isAuthFlowExpired(session)) {
      if (session.step === 'awaiting_email') {
        return await handleAuthEmailInput(env, chatId, text, messageId);
      }
      if (session.step === 'awaiting_password') {
        return await handleAuthPasswordInput(env, chatId, text, messageId, session);
      }
    }

    // If auth flow expired, clean up
    if (session && isAuthFlowExpired(session) && session.step !== 'authenticated') {
      await deleteAuthSession(env, chatId);
    }

    // /start pay_<order_id> — deep link dari extension
    if (text.startsWith('/start pay_')) {
      const orderId = text.substring('/start pay_'.length).trim();
      return await handleStartPay(env, chatId, orderId, username, firstName);
    }

    // /start — show main menu
    if (text === '/start' || text.startsWith('/start ')) {
      return await handleStartMenu(env, chatId, firstName);
    }

    // /menu
    if (text === '/menu') return await handleStartMenu(env, chatId, firstName);

    // /paket or /beli
    if (text === '/paket' || text === '/beli') return await handlePaketCommand(env, chatId);

    // /akun — v1.3: check authenticated session first
    if (text === '/akun') return await handleAkunCommand(env, chatId, session);

    // /license — v1.3: check authenticated session first
    if (text === '/license') return await handleLicenseCommand(env, chatId, session);

    // /logout — v1.3: clear session
    if (text === '/logout') return await handleLogoutCommand(env, chatId);

    // /cancel — v1.3: cancel auth flow
    if (text === '/cancel') {
      await deleteAuthSession(env, chatId);
      await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Dibatalkan.');
      return new Response('OK', { status: 200 });
    }

    // /riwayat
    if (text === '/riwayat') return await handleRiwayatCommand(env, chatId);

    // /fitur
    if (text === '/fitur') return await handleFiturCommand(env, chatId);

    // /support
    if (text === '/support') return await handleSupportCommand(env, chatId);

    // /help or /bantuan
    if (text === '/help' || text === '/bantuan') return await handleHelpCommand(env, chatId);

    // /status <order_id> — bare /status (from the command menu, no args yet)
    if (text === '/status') {
      await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
        `ℹ️ Format: <code>/status ORDER_ID [email]</code>\n\nContoh:\n<code>/status INV-20260926-TQDWBE email@anda.com</code>\n\nOrder ID bisa dilihat di pesan "Pembayaran Dibuat" atau di /riwayat.`);
      return new Response('OK', { status: 200 });
    }
    if (text.startsWith('/status ')) {
      const statusParts = text.substring('/status '.length).trim().split(/\s+/);
      const orderId = statusParts[0];
      const buyerEmail = statusParts[1] || null;
      return await handleStatusCommand(env, chatId, orderId, buyerEmail);
    }

    // Unknown
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
      `🤖 Maaf, saya tidak mengerti perintah itu.\n\nKetik /menu untuk melihat menu utama.`);
    return new Response('OK', { status: 200 });

  } catch (e) {
    console.error('[TG Webhook] Error:', e);
    await sendBugReport(env, e, { handler: 'telegram_webhook' });
    return new Response('OK', { status: 200 });
  }
}

// ============================================================
// TELEGRAM COMMAND HANDLERS
// ============================================================

// ============================================================
// FEATURE-RESPONSE CACHE — 1 hour TTL, non-payment bot features only.
// /status, /riwayat, the "Cek Pembayaran" callback, the payment-create /
// guest-payment flow, and the HTTP payment endpoints must always reflect
// live data and deliberately never go through this cache. Everything else
// (menu, /paket, /fitur, /support, /help, /akun, /license) is either fully
// static or safe to serve slightly stale for up to an hour, so it's cached
// to cut down repeat Supabase calls. Best-effort only: this is an
// in-memory per-isolate Map, so a recycled Worker isolate just recomputes
// on the next request — never a correctness issue, only a perf one.
// ============================================================
const FEATURE_CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour
const _featureCache = new Map(); // key -> { value, expiresAt }

function getFeatureCache(key) {
  const hit = _featureCache.get(key);
  if (!hit) return null;
  if (hit.expiresAt <= Date.now()) { _featureCache.delete(key); return null; }
  return hit.value;
}

function setFeatureCache(key, value, ttlMs = FEATURE_CACHE_TTL_MS) {
  _featureCache.set(key, { value, expiresAt: Date.now() + ttlMs });
  return value;
}

async function handleStartMenu(env, chatId, firstName) {
  const cacheKey = `menu:${chatId}`;
  const msg = getFeatureCache(cacheKey) || setFeatureCache(cacheKey, `👋 <b>Halo ${escapeHtml(firstName || '')}! Selamat datang di Shopee Product Hunter Bot</b>

🤖 Bot ini akan membantu Anda membeli license, cek status pembayaran, dan mendapatkan notifikasi otomatis.

Pilih menu di bawah:`);
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

async function handlePaketCommand(env, chatId) {
  const cacheKey = 'paket';
  const msg = getFeatureCache(cacheKey) || setFeatureCache(cacheKey, (() => {
    const list = Object.entries(PRICING_TIERS)
      .map(([days, t]) => `🔹 <b>${t.name}</b> — ${days} hari — ${formatRupiah(t.amount)}`)
      .join('\n');
    return `💎 <b>Pilih Paket License</b>

Berikut pilihan paket yang tersedia:

${list}

Pilih paket di bawah:`;
  })());
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getPaketKeyboard() });
  return new Response('OK', { status: 200 });
}

// v1.3: AUTH FLOW HANDLERS — stateful conversation

async function handlePaymentEmailInput(env, chatId, text, messageId, days) {
  const email = text.trim().toLowerCase();
  if (!isValidEmail(email)) {
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Format email tidak valid. Silakan masukkan email yang benar.');
    return new Response('OK', { status: 200 });
  }
  if (!PRICING_TIERS[days]) { await deleteAuthSession(env, chatId); await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Paket pembayaran tidak valid. Silakan ketik /beli lagi.'); return new Response('OK', { status: 200 }); }
  await saveAuthSession(env, chatId, { step: 'payment_email_saved', email });
  const tier = PRICING_TIERS[days];
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `✅ Email tersimpan: <code>${escapeHtml(email)}</code>\n\n📦 <b>${tier.name}</b> — ${days} hari\n💵 <b>${formatRupiah(tier.amount)}</b>\n\nKlik tombol berikut untuk melanjutkan ke pembayaran Lynk.id.`, { reply_markup: getPaymentMethodKeyboard(days) });
  return new Response('OK', { status: 200 });
}

async function handleAuthEmailInput(env, chatId, text, messageId) {
  const email = text.trim();
  // Basic email validation
  if (!isValidEmail(email)) {
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
      `❌ Email tidak valid. Silakan masukkan email yang benar:\n\nAtau ketik /cancel untuk batal.`);
    return new Response('OK', { status: 200 });
  }
  // Save email, move to awaiting_password
  await saveAuthSession(env, chatId, { step: 'awaiting_password', email });
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
    `📧 Email diterima: <code>${escapeHtml(email)}</code>\n\n🔑 Sekarang masukkan password Anda:\n\n⚠️ Password akan dihapus otomatis setelah verifikasi untuk keamanan.\n\nKetik /cancel untuk batal.`);
  return new Response('OK', { status: 200 });
}

async function handleAuthPasswordInput(env, chatId, text, messageId, session) {
  const password = text.trim();

  // Delete password message immediately (security)
  await telegramDeleteMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, messageId);

  if (!password) {
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
      `❌ Password tidak boleh kosong. Silakan masukkan password:\n\nKetik /cancel untuk batal.`);
    return new Response('OK', { status: 200 });
  }

  // Show "verifying" message
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `⏳ Memverifikasi...`);

  // Verify via Supabase Auth API
  const authResult = await supabaseAuthVerify(env, session.email, password);
  if (!authResult.valid) {
    await deleteAuthSession(env, chatId);
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
      `❌ Email atau password salah.\n\nCoba lagi dengan /akun atau /license.`);
    return new Response('OK', { status: 200 });
  }

  // Auth success — save authenticated session (cached forever until /logout)
  await saveAuthSession(env, chatId, {
    step: 'authenticated',
    email: session.email,
    user_id: authResult.user.id,
    access_token: authResult.session.access_token,
    refresh_token: authResult.session.refresh_token
  });

  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
    `✅ Login berhasil!\n\nSekarang Anda bisa mengakses /akun dan /license tanpa login ulang.\n\nKetik /akun untuk info akun, atau /license untuk info license.`);
  return new Response('OK', { status: 200 });
}

// v1.3: /akun — check authenticated session, if not authed → start auth flow
async function handleAkunCommand(env, chatId, session) {
  // Check if already authenticated
  if (session && session.step === 'authenticated' && !isAuthFlowExpired(session)) {
    return await displayAkunInfo(env, chatId, session.user_id);
  }

  // Not authenticated → start auth flow
  await saveAuthSession(env, chatId, { step: 'awaiting_email', email: null });
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
    `👤 <b>Akun Saya</b>\n\nUntuk melihat info akun, silakan login.\n\n📧 Masukkan email Anda:\n\nKetik /cancel untuk batal.`);
  return new Response('OK', { status: 200 });
}

async function displayAkunInfo(env, chatId, userId) {
  const cacheKey = `akun:${userId}`;
  let msg = getFeatureCache(cacheKey);
  if (!msg) {
    // Get subscription info
    const subResult = await supabaseFetch(env, `/rest/v1/subscriptions?user_id=eq.${userId}&select=*`, { method: 'GET' });
    let subInfo = 'Status: Tidak diketahui';
    if (subResult.ok && Array.isArray(subResult.data) && subResult.data.length > 0) {
      const sub = subResult.data[0];
      subInfo = `Plan: <b>${escapeHtml(sub.plan)}</b>\nStatus: ${escapeHtml(sub.status)}\nExpires: ${formatDate(sub.expires_at)}`;
    }

    // Get profile email
    const profileResult = await supabaseFetch(env, `/rest/v1/profiles?id=eq.${userId}&select=email`, { method: 'GET' });
    const email = (profileResult.ok && profileResult.data?.[0]?.email) || 'unknown';

    // Get team info
    const teamInfo = await getTeamMembersInfo(env, userId);

    let teamSection = '';
    if (teamInfo.hasLicense) {
      teamSection = `\n\n👥 <b>Tim</b>\nTotal Seats: ${teamInfo.maxSeats} (1 owner + ${teamInfo.maxSeats - 1} anggota)\n✅ Aktif: ${teamInfo.activeCount}\n⏳ Menunggu: ${teamInfo.invitedCount}\n🆓 Kosong (unused): ${teamInfo.unusedSeats}`;

      if (teamInfo.members.length > 0) {
        teamSection += '\n\n<b>Daftar Anggota:</b>';
        for (const m of teamInfo.members) {
          const statusIcon = m.status === 'active' ? '✅' : '⏳';
          teamSection += `\n${statusIcon} <code>${escapeHtml(m.email)}</code>`;
        }
      } else {
        teamSection += '\n\nBelum ada anggota tim. Invite via extension popup.';
      }
    } else {
      teamSection = '\n\n👥 <b>Tim</b>\nAnda belum punya license aktif.';
    }

    msg = setFeatureCache(cacheKey, `👤 <b>Akun Saya</b>

📧 Email: ${escapeHtml(email)}
📊 ${subInfo}${teamSection}

💡 Ketik /logout untuk keluar dari sesi ini.`);
  }
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

// v1.3: /license — check authenticated session, if not authed → start auth flow
async function handleLicenseCommand(env, chatId, session) {
  // Check if already authenticated
  if (session && session.step === 'authenticated' && !isAuthFlowExpired(session)) {
    return await displayLicenseInfo(env, chatId, session.user_id);
  }

  // Not authenticated → start auth flow
  await saveAuthSession(env, chatId, { step: 'awaiting_email', email: null });
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
    `🔑 <b>License Saya</b>\n\nUntuk melihat info license, silakan login.\n\n📧 Masukkan email Anda:\n\nKetik /cancel untuk batal.`);
  return new Response('OK', { status: 200 });
}

async function displayLicenseInfo(env, chatId, userId) {
  const cacheKey = `license:${userId}`;
  let msg = getFeatureCache(cacheKey);
  if (!msg) {
    const licResult = await supabaseFetch(env, `/rest/v1/licenses?owner_user_id=eq.${userId}&status=eq.active&select=*`, { method: 'GET' });
    if (!licResult.ok || !Array.isArray(licResult.data) || licResult.data.length === 0) {
      msg = setFeatureCache(cacheKey, `🔑 <b>License Saya</b>

⚠️ Anda belum punya license aktif.

Ketik /beli untuk membeli license baru.`);
    } else {
      const lic = licResult.data[0];
      msg = setFeatureCache(cacheKey, `🔑 <b>License Saya</b>

🔑 License Key: <code>${lic.license_key}</code>
📦 Plan: ${escapeHtml(lic.plan)}
📊 Status: ${escapeHtml(lic.status)}
📅 Expires: ${formatDate(lic.expires_at)}
👥 Max Seats: ${lic.max_seats}
📅 Activated: ${formatDate(lic.activated_at)}

💡 Ketik /logout untuk keluar dari sesi ini.`);
    }
  }
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

// v1.3: /logout — clear session
async function handleLogoutCommand(env, chatId) {
  await deleteAuthSession(env, chatId);
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId,
    `✅ Anda telah logout.\n\nSesi berhasil dihapus. Untuk mengakses /akun atau /license, silakan login kembali.`);
  return new Response('OK', { status: 200 });
}

async function handleRiwayatCommand(env, chatId) {
  // Get guest payments by chat_id
  const guestResult = await supabaseFetch(env, `/rest/v1/payments?guest_chat_id=eq.${encodeURIComponent(chatId)}&order=created_at.desc&limit=10&select=*`, { method: 'GET' });

  // Get linked user payments
  const linkResult = await supabaseFetch(env, `/rest/v1/user_telegram_chats?chat_id=eq.${encodeURIComponent(chatId)}&select=user_id`, { method: 'GET' });
  let linkedPayments = [];
  if (linkResult.ok && Array.isArray(linkResult.data) && linkResult.data.length > 0) {
    const userId = linkResult.data[0].user_id;
    const userPays = await supabaseFetch(env, `/rest/v1/payments?user_id=eq.${userId}&order=created_at.desc&limit=10&select=*`, { method: 'GET' });
    if (userPays.ok && Array.isArray(userPays.data)) linkedPayments = userPays.data;
  }

  const allPayments = [...(guestResult.data || []), ...linkedPayments]
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at))
    .slice(0, 10);

  if (allPayments.length === 0) {
    const msg = `📄 <b>Riwayat Transaksi</b>

⚠️ Belum ada transaksi.

Ketik /beli untuk mulai membeli license.`;
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
    return new Response('OK', { status: 200 });
  }

  let msg = `📄 <b>Riwayat Transaksi (10 terakhir)</b>\n`;
  for (const p of allPayments) {
    const emoji = p.status === 'completed' ? '✅' : p.status === 'canceled' ? '❌' : '⏳';
    msg += `\n${emoji} <code>${p.order_id}</code>\n   ${escapeHtml(p.plan_name)} • ${formatRupiah(p.amount)} • ${escapeHtml(p.status)}\n`;
  }
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

async function handleFiturCommand(env, chatId) {
  const cacheKey = 'fitur';
  const msg = getFeatureCache(cacheKey) || setFeatureCache(cacheKey, `📋 <b>Informasi Fitur Shopee Product Hunter</b>

🔍 <b>Analisa Pencarian</b>
Analisis kata kunci pencarian, jumlah produk, harga rata-rata, dan potensi market.

📦 <b>Analisis Variasi Terlaris</b>
Lihat variasi produk mana yang paling laris dengan data real-time.

🔎 <b>Invite Team</b>
Kamu bisa invite beberapa email untuk lisensi yang syarat dan ketentuan berlaku.

💎 <b>CTime Box</b>
Lihat tanggal upload produk dan umur produk secara akurat.

🔄 <b>View Similar Products</b>
Cari produk serupa dengan analisis harga dan sold count.

📊 <b>Filter Produk Lanjutan</b>
Filter produk berdasarkan harga, rating, penjualan.

💰 <b>Affiliate Intelligence</b>
Lihat komisi affiliate, buat link affiliate, dan analisis potensi earning.

🎯 <b>Product Card Badge</b>
Badge di kartu produk menampilkan info penting (tandai warna tanggal upload produk dan umur produk dll).

Semua fitur aktif setelah license diaktivasi! ✅`);
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

async function handleSupportCommand(env, chatId) {
  const cacheKey = 'support';
  const msg = getFeatureCache(cacheKey) || setFeatureCache(cacheKey, `💬 <b>Grup Support</b>

Butuh bantuan? Join grup support kami:

👉 <a href="${escapeHtml(getSupportGroupUrl(env))}">Klik untuk Join Grup Support</a>

Di grup support Anda bisa:
• Tanya jawab dengan admin
• Dapatkan update terbaru
• Berbagi tips dengan user lain
• Report bug atau masalah`);
  const keyboard = JSON.stringify({
    inline_keyboard: [
      [{ text: '💬 Join Grup Support', url: getSupportGroupUrl(env) }],
      [{ text: '⬅️ Kembali ke Menu', callback_data: 'menu' }]
    ]
  });
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: keyboard });
  return new Response('OK', { status: 200 });
}

async function handleHelpCommand(env, chatId) {
  const cacheKey = 'help';
  const msg = getFeatureCache(cacheKey) || setFeatureCache(cacheKey, `❓ <b>Bantuan — Shopee Product Hunter Bot</b>

<b>Perintah tersedia:</b>
/start — Tampilkan menu utama
/menu — Tampilkan menu utama
/beli — Beli license langsung dari Telegram
/paket — Lihat daftar paket
/akun — Info akun Anda (perlu login email+password)
/license — Info license Anda (perlu login email+password)
/logout — Keluar dari sesi login
/cancel — Batalkan proses login
/riwayat — Riwayat transaksi
/fitur — Informasi fitur aplikasi
/support — Grup support
/help — Bantuan ini
/status &lt;order_id&gt; [email] — Cek + sinkronisasi status pembayaran

<b>Cara beli license (Pemula):</b>
1. Ketik /beli
2. Pilih paket (Basic/Starter/Growth/Pro/Annual)
3. Pilih pembayaran via Lynk.id
4. Selesaikan pembayaran
5. Cek status pembayaran, lalu license key dikirim otomatis ✅
6. Install extension → Login → Aktifkan License

<b>Cara cek akun/license:</b>
1. Ketik /akun atau /license
2. Masukkan email akun extension Anda
3. Masukkan password (akan dihapus otomatis untuk keamanan)
4. Info ditampilkan ✅
5. Sekali login, tidak perlu login ulang sampai /logout

Butuh bantuan lain? Join: ${getSupportGroupUrl(env)}`);
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, { reply_markup: getMainMenuKeyboard() });
  return new Response('OK', { status: 200 });
}

async function handleStatusCommand(env, chatId, orderId, buyerEmail = null) {
  let result = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&guest_chat_id=eq.${encodeURIComponent(chatId)}&select=*`, { method: 'GET' });
  if (!result.ok || !Array.isArray(result.data) || !result.data.length) { const link = await supabaseFetch(env, `/rest/v1/user_telegram_chats?chat_id=eq.${encodeURIComponent(chatId)}&select=user_id`, { method: 'GET' }); if (link.ok && Array.isArray(link.data) && link.data.length) result = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&user_id=eq.${link.data[0].user_id}&select=*`, { method: 'GET' }); }
  if (!result.ok || !Array.isArray(result.data) || !result.data.length) { await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `❌ Order ID tidak ditemukan: <code>${escapeHtml(orderId)}</code>`); return new Response('OK', { status: 200 }); }
  const reconciled = await reconcilePaymentByOrder(env, orderId, buyerEmail); if (!reconciled.ok) { await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `❌ Gagal sinkronisasi pembayaran: ${escapeHtml(reconciled.message || reconciled.code)}`); return new Response('OK', { status: 200 }); }
  const p = reconciled.payment || result.data[0]; const emoji = p.status === 'completed' ? '✅' : '⏳'; let msg = `${emoji} <b>Status Pembayaran</b>\n\n🆔 <code>${p.order_id}</code>\n📦 ${escapeHtml(p.plan_name)} (${p.plan_duration_days} hari)\n💵 ${formatRupiah(p.amount)}\n💳 Lynk.id\n📊 Status: <b>${escapeHtml(p.status)}</b>`;
  if (p.status === 'completed' && p.license_key) msg += `\n\n🎉 <b>License Key:</b>\n<code>${p.license_key}</code>`; else msg += `\n\nBelum ditemukan transaksi sukses. Pastikan email checkout benar lalu coba lagi.`;
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg); return new Response('OK', { status: 200 });
}


async function handleStartPay(env, chatId, orderId, username, firstName) {
  const findResult = await supabaseFetch(env, `/rest/v1/payments?order_id=eq.${encodeURIComponent(orderId)}&select=*`, { method: 'GET' });
  if (!findResult.ok || !Array.isArray(findResult.data) || findResult.data.length === 0) {
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `❌ Order ID tidak ditemukan: <code>${escapeHtml(orderId)}</code>`);
    return new Response('OK', { status: 200 });
  }
  const payment = findResult.data[0];
  // Save mapping
  if (payment.user_id) {
    await supabaseFetch(env, '/rest/v1/user_telegram_chats', {
      method: 'POST',
      body: { user_id: payment.user_id, chat_id: chatId, username, first_name: firstName, is_blocked: false }
    });
    await supabaseFetch(env, `/rest/v1/user_telegram_chats?user_id=eq.${payment.user_id}`, {
      method: 'PATCH',
      body: { chat_id: chatId, username, first_name: firstName, is_blocked: false, updated_at: new Date().toISOString() }
    });
  }
  const tier = PRICING_TIERS[payment.plan_duration_days] || { name: payment.plan_name };
  const msg = `👋 <b>Halo ${escapeHtml(firstName || '')}! Selamat datang!</b>

✅ Akun Telegram Anda berhasil terhubung!

📦 <b>Detail Pembayaran:</b>
• Paket: <b>${escapeHtml(tier.name || payment.plan_name)}</b> (${payment.plan_duration_days} hari)
• Total: <b>${formatRupiah(payment.amount)}</b>
• Metode: ${PAYMENT_METHOD_LABELS[payment.payment_method] || escapeHtml(payment.payment_method)}
• Order ID: <code>${orderId}</code>
• Status: ${payment.status === 'pending' ? '⏳ Menunggu pembayaran' : escapeHtml(payment.status)}

⏰ <b>Bayar sebelum:</b> ${formatDate(payment.expired_at)}`;
  const opts = {};
  if ((payment.payment_method === 'payment_link' || payment.payment_method === 'lynk') && payment.payment_link) {
    opts.reply_markup = JSON.stringify({ inline_keyboard: [[{ text: '💳 Bayar Sekarang', url: payment.payment_link }]] });
  }
  await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg, opts);
  await sendAdminNotification(env, `🔔 <b>User Register Telegram</b>\n\n💬 Chat ID: <code>${chatId}</code>\n👤 @${escapeHtml(username || '-')}\n📦 <code>${orderId}</code>`);
  return new Response('OK', { status: 200 });
}

// ============================================================
// CALLBACK QUERY HANDLER — Inline Button Clicks
// ============================================================
async function handleCallbackQuery(env, cq) {
  const chatId = String(cq.message?.chat?.id || '');
  const data = cq.data || '';
  const cqId = cq.id;
  const firstName = cq.from?.first_name || '';
  const username = cq.from?.username || '';

  await telegramAnswerCallback(env.TELEGRAM_USER_BOT_TOKEN, cqId);

  // menu
  if (data === 'menu') return await handleStartMenu(env, chatId, firstName);

  // beli
  if (data === 'beli') return await handlePaketCommand(env, chatId);

  // akun
  if (data === 'akun') {
    const sess = await getAuthSession(env, chatId);
    return await handleAkunCommand(env, chatId, sess);
  }

  if (data === 'license') {
    const sess = await getAuthSession(env, chatId);
    return await handleLicenseCommand(env, chatId, sess);
  }

  // riwayat
  if (data === 'riwayat') return await handleRiwayatCommand(env, chatId);

  // fitur
  if (data === 'fitur') return await handleFiturCommand(env, chatId);

  // support
  if (data === 'support') return await handleSupportCommand(env, chatId);

  // bantuan
  if (data === 'bantuan') return await handleHelpCommand(env, chatId);

  // paket_<days> — user pilih paket, show payment methods
  if (data.startsWith('paket_')) {
    const days = parseInt(data.substring('paket_'.length), 10);
    const tier = PRICING_TIERS[days];
    if (!tier) return new Response('OK', { status: 200 });
    const currentSession = await getAuthSession(env, chatId);
    if (currentSession?.step === 'authenticated' && currentSession.email) {
      await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `✅ <b>Paket dipilih: ${tier.name}</b>

Duration: ${days} hari
Harga: ${formatRupiah(tier.amount)}

Pilih metode pembayaran:`, { reply_markup: getPaymentMethodKeyboard(days) });
    } else {
      await saveAuthSession(env, chatId, { step: `awaiting_payment_email_${days}`, email: null });
      const msg = `✅ <b>Paket dipilih: ${tier.name}</b>

Duration: ${days} hari
Harga: ${formatRupiah(tier.amount)}

📧 Masukkan email yang akan Anda gunakan saat checkout Lynk.id:`;
      await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, msg);
    }
    return new Response('OK', { status: 200 });
  }

  // check_<order_id> — reconcile against Google Spreadsheet
  if (data.startsWith('check_')) {
    const orderId = data.substring('check_'.length).trim();
    const session = await getAuthSession(env, chatId);
    const buyerEmail = (session?.email && (session.step === 'payment_email_saved' || session.step === 'authenticated')) ? session.email : null;
    const result = await reconcilePaymentByOrder(env, orderId, buyerEmail);
    if (!result.ok) { await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `❌ Gagal cek pembayaran: ${escapeHtml(result.message || result.code)}`); return new Response('OK', { status: 200 }); }
    if (result.reconciled) await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `🎉 <b>Pembayaran terverifikasi!</b>\n\n🆔 <code>${orderId}</code>\n🔑 License: <code>${result.payment.license_key}</code>`);
    else await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `⏳ <b>Belum terverifikasi</b>\n\n🆔 <code>${orderId}</code>\n\nJika email checkout berbeda, gunakan:\n<code>/status ${orderId} email@anda.com</code>`);
    return new Response('OK', { status: 200 });
  }

  // pay_<days>_lynk — create guest payment
  if (data.startsWith('pay_')) {
    const parts = data.split('_');
    if (parts.length < 3) return new Response('OK', { status: 200 });
    const days = parseInt(parts[1], 10);
    const method = parts.slice(2).join('_');
    return await handleGuestPayment(env, chatId, days, method, username, firstName);
  }

  return new Response('OK', { status: 200 });
}

// ============================================================
// GUEST PAYMENT — Create payment from Telegram (no JWT needed)
// ============================================================
async function handleGuestPayment(env, chatId, planDurationDays, paymentMethod, username, firstName) {
  try {
    if (!PRICING_TIERS[planDurationDays]) { await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Paket tidak valid.'); return new Response('OK', { status: 200 }); }
    if (!ALLOWED_PAYMENT_METHODS.includes(paymentMethod)) { await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Metode pembayaran tidak valid.'); return new Response('OK', { status: 200 }); }
    const session = await getAuthSession(env, chatId);
    const buyerEmail = (session && (session.step === 'payment_email_saved' || session.step === 'authenticated') && session.email) ? String(session.email).trim().toLowerCase() : null;
    if (!buyerEmail) {
      await saveAuthSession(env, chatId, { step: `awaiting_payment_email_${planDurationDays}`, email: null });
      await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, `📧 <b>Email Checkout Lynk.id</b>\n\nMasukkan email yang akan Anda gunakan saat checkout Lynk.id.\n\n⚠️ Gunakan email yang sama saat membayar agar pembayaran bisa diverifikasi otomatis.`);
      return new Response('OK', { status: 200 });
    }

    const tier = PRICING_TIERS[planDurationDays];

    // Reuse an existing pending order for this chat+package instead of
    // creating a new payments row (and a new admin ping) on every tap of
    // "Bayar Sekarang" / re-selection of the same package.
    const reusable = await findReusablePendingGuestPayment(env, chatId, planDurationDays);
    let orderId;
    if (reusable) {
      orderId = reusable.order_id;
    } else {
      orderId = generateOrderId();
      await createLynkPaymentRecord(env, { guestChatId: chatId, buyerEmail, planDurationDays, orderId });
      await sendAdminNotification(env, `🆕 <b>GUEST LYNK PAYMENT CREATED</b>\n\n💬 Chat ID: <code>${chatId}</code>\n👤 @${escapeHtml(username || '-')} (${escapeHtml(firstName || '')})\n📦 ${tier.name} (${planDurationDays}d)\n💵 ${formatRupiah(tier.amount)}\n💳 Lynk.id\n🆔 <code>${orderId}</code>`);
    }

    const payMsg = `✅ <b>Pembayaran Dibuat!</b>\n\n📦 <b>Paket:</b> ${tier.name} (${planDurationDays} hari)\n💵 <b>Total:</b> ${formatRupiah(tier.amount)}\n💳 <b>Metode:</b> Lynk.id\n🆔 <b>Order ID:</b> <code>${orderId}</code>\n\nKlik <b>Bayar Sekarang</b> untuk checkout di Lynk.id.\n\n⚠️ Setelah membayar, status akan diverifikasi otomatis oleh sistem kami.\n\nJika belum terdeteksi, gunakan:\n<code>/status ${orderId} email@anda.com</code>`;
    const opts = { reply_markup: JSON.stringify({ inline_keyboard: [[{ text: '💳 Bayar Sekarang', url: tier.lynk }], [{ text: '🔄 Cek Pembayaran', callback_data: `check_${orderId}` }], [{ text: '⬅️ Kembali ke Menu', callback_data: 'menu' }]] }) };
    await saveAuthSession(env, chatId, { step: 'payment_email_saved', email: buyerEmail });
    await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, payMsg, opts); return new Response('OK', { status: 200 });
  } catch (e) { console.error('[GuestPayment] Error:', e); await sendBugReport(env, e, { handler: 'guest_payment', chatId, planDurationDays, paymentMethod }); await telegramSendMessage(env.TELEGRAM_USER_BOT_TOKEN, chatId, '❌ Terjadi error. Coba lagi atau hubungi support.'); return new Response('OK', { status: 200 }); }
}


// ============================================================
// HANDLER: GET /health
// ============================================================
async function handleHealth(env) {
  const config = {
    supabase: !!env.SUPABASE_URL && !!env.SUPABASE_SERVICE_ROLE_KEY,
    supabase_anon: !!env.SUPABASE_ANON_KEY,
    lynk_checkout_links: Object.values(PRICING_TIERS).every(t => !!t.lynk),
    // Deliberately expose only a boolean. Never return the Spreadsheet ID.
    google_spreadsheet: true,
    payment_apps_script: !!env.PAYMENT_APPS_SCRIPT_URL && !!env.PAYMENT_APPS_SCRIPT_SECRET,
    jwt: !!env.JWT_SECRET && env.JWT_SECRET.length >= 16,
    telegram_user_bot: !!env.TELEGRAM_USER_BOT_TOKEN && !!env.TELEGRAM_USER_BOT_USERNAME,
    telegram_admin_bot: !!env.TELEGRAM_ADMIN_BOT_TOKEN && !!env.TELEGRAM_ADMIN_CHAT_ID,
    telegram_webhook_secret: !!env.TELEGRAM_WEBHOOK_SECRET
  };

  // Public health response intentionally excludes:
  // - spreadsheet_id / Spreadsheet ID
  // - secret values
  // - Apps Script URL
  // - Supabase URL/keys
  // - Telegram bot tokens
  //
  // The Spreadsheet ID remains internal configuration only.
  return jsonResponse({
    success: true,
    service: 'shopee-product-hunter-payment',
    version: 'v2.4',
    release: 'v2.4.0-lynk-web-checkout',
    timestamp: new Date().toISOString(),
    config_status: config,
    ready: [
      'supabase',
      'supabase_anon',
      'lynk_checkout_links',
      'google_spreadsheet',
      'payment_apps_script',
      'jwt',
      'telegram_user_bot',
      'telegram_admin_bot'
    ].every(k => config[k] === true),
    pricing_tiers: Object.entries(PRICING_TIERS).map(([days, t]) => ({
      duration_days: parseInt(days, 10),
      name: t.name,
      product_title: t.productTitle,
      amount: t.amount,
      amount_formatted: formatRupiah(t.amount),
      payment_link: t.lynk
      // spreadsheet_id intentionally omitted
    })),
    payment_methods: [{ code: 'lynk', label: '💳 Lynk.id' }],
    endpoints: [
      'POST /api/v1/payment/create',
      'GET /api/v1/payment/status/:order_id',
      'GET /api/v1/payments/history',
      'POST /api/v1/payment/reconcile-ping',
      'GET /api/v1/web/plans',
      'POST /api/v1/web/checkout',
      'GET /api/v1/web/status/:order_id',
      'POST /api/v1/telegram/user-bot',
      'POST /api/v1/telegram/setup-commands',
      'GET /health'
    ]
  });
}


// ============================================================
// OPTIONAL CRON RECONCILIATION — no webhook
// Configure a Cloudflare Cron trigger (e.g. */5 * * * *) to make
// pending payments reconcile automatically without user interaction.
// ============================================================
async function reconcilePendingPayments(env) {
  const cutoff = new Date(Date.now() - PENDING_RECONCILE_WINDOW_MS).toISOString();
  const result = await supabaseFetch(env, `/rest/v1/payments?status=eq.pending&payment_method=eq.lynk&created_at=gte.${encodeURIComponent(cutoff)}&order=created_at.desc&limit=50&select=*`, { method: 'GET' });
  if (!result.ok || !Array.isArray(result.data)) return { checked: 0, completed: 0 };
  let completed = 0;
  for (const payment of result.data) {
    try {
      let buyerEmail = null;
      if (payment.user_id) {
        const profile = await supabaseFetch(env, `/rest/v1/profiles?id=eq.${payment.user_id}&select=email`, { method: 'GET' });
        buyerEmail = profile.ok && Array.isArray(profile.data) && profile.data[0]?.email ? profile.data[0].email : null;
      } else if (isWebGuest(payment.guest_chat_id)) {
        buyerEmail = webGuestEmail(payment.guest_chat_id);
      } else if (payment.guest_chat_id) {
        const session = await getAuthSession(env, String(payment.guest_chat_id));
        if (session?.email && (session.step === 'payment_email_saved' || session.step === 'authenticated')) buyerEmail = session.email;
      }
      const reconciled = await reconcilePaymentByOrder(env, payment.order_id, buyerEmail);
      if (reconciled.ok && reconciled.reconciled) completed++;
    } catch (e) {
      console.error('[CronSync] Payment error:', payment.order_id, e);
    }
  }
  return { checked: result.data.length, completed };
}

// ============================================================
// MAIN WORKER
// ============================================================
// ============================================================
// HANDLER: POST /api/v1/telegram/setup-commands
// One-shot (idempotent, safe to re-run) endpoint to register the bot's
// command list + turn on the "Menu" button in the chat's compose bar.
// Also re-applied automatically from scheduled() so it self-heals if the
// command list is ever changed and the endpoint isn't called manually.
// ============================================================
async function handleTelegramSetupCommands(env, request) {
  const configuredSecret = env.TELEGRAM_WEBHOOK_SECRET || env.PAYMENT_APPS_SCRIPT_SECRET || '';
  if (configuredSecret) {
    const provided = request.headers.get('X-Setup-Secret') || '';
    if (provided !== configuredSecret) return errorResponse('UNAUTHORIZED', 'Setup secret tidak valid', 401);
  }
  const result = await telegramSetupCommandMenu(env.TELEGRAM_USER_BOT_TOKEN);
  return jsonResponse({ success: !!result.ok, result });
}

export default {
  async fetch(request, env, ctx) {
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS_HEADERS });
    const url = new URL(request.url);
    const path = url.pathname;
    const method = request.method;
    try {
      if (path === '/health' || path === '/') return await handleHealth(env);
      if (path === '/api/v1/payment/create' && method === 'POST') return await handlePaymentCreate(env, request);
      if (path.startsWith('/api/v1/payment/status/') && method === 'GET') return await handlePaymentStatus(env, request, url);
      if (path === '/api/v1/payments/history' && method === 'GET') return await handlePaymentsHistory(env, request);
      if (path === '/api/v1/web/plans' && method === 'GET') return await handleWebPlans();
      if (path === '/api/v1/web/checkout' && method === 'POST') return await handleWebCheckout(env, request);
      if (path.startsWith('/api/v1/web/status/') && method === 'GET') return await handleWebStatus(env, request, url);
      if (path === '/api/v1/payment/reconcile-ping' && method === 'POST') return await handlePaymentReconcilePing(env, request);
      if (path === '/api/v1/telegram/user-bot' && method === 'POST') return await handleTelegramUserBotWebhook(env, request);
      if (path === '/api/v1/telegram/setup-commands' && method === 'POST') return await handleTelegramSetupCommands(env, request);
      return errorResponse('ROUTE_NOT_FOUND', `Route not found: ${method} ${path}`, 404);
    } catch (e) {
      console.error('[Worker] Unhandled error:', e);
      await sendBugReport(env, e, { path, method });
      return errorResponse('INTERNAL_ERROR', safeErrorMessage(), 500);
    }
  },

  async scheduled(controller, env, ctx) {
    ctx.waitUntil(reconcilePendingPayments(env));
    // Cheap and idempotent — keeps the command menu in sync automatically
    // even if nobody remembers to call the setup endpoint after a change.
    ctx.waitUntil(telegramSetupCommandMenu(env.TELEGRAM_USER_BOT_TOKEN));
  }
};