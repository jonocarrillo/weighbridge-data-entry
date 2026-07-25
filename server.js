'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const net = require('node:net');
const tls = require('node:tls');
const crypto = require('node:crypto');
const { URL } = require('node:url');
const { getDb, resolveDatabasePath, sigBytes, runDbBackup, listBackups, dbStats } = require('./db');

/** Detect PNG/JPEG from magic bytes (signatures may be either after compression). */
function imageMimeFromBytes(buf) {
  if (!buf || buf.length < 3) return 'application/octet-stream';
  const b = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e) return 'image/png';
  return 'application/octet-stream';
}
function imageDataUrlFromBytes(buf) {
  const bytes = sigBytes(buf) || (Buffer.isBuffer(buf) ? buf : null);
  if (!bytes) return '';
  const mime = imageMimeFromBytes(bytes);
  return `data:${mime};base64,${bytes.toString('base64')}`;
}

const PORT = Number(process.env.PORT || process.env.AIS_DASHBOARD_PORT || 5000);
const HOST = process.env.AIS_DASHBOARD_HOST || '0.0.0.0';
const PROJECT_ROOT = __dirname;
const PUBLIC_DIR = path.join(PROJECT_ROOT, 'public');
const VIEWS_DIR = path.join(PROJECT_ROOT, 'views');
const SIGNATURE_LINE = 'Jesus Polanco';
const MAX_JSON_BYTES = 1_000_000;
const MAX_UPLOAD_BYTES = 5_000_000;
const MAX_SIGNATURE_BYTES = 250_000;

// Signature-only kiosk devices (LAN tablets). Matched by MAC (preferred) and/or IP.
// Example: AIS_SIGNATURE_ONLY_MACS=16:47:d8:b1:ef:e7
//          AIS_SIGNATURE_ONLY_IPS=192.168.1.102
function parseListEnv(name) {
  return String(process.env[name] || '')
    .split(/[\s,;]+/)
    .map(s => s.trim())
    .filter(Boolean);
}
function normalizeMac(mac) {
  const hex = String(mac || '').toLowerCase().replace(/[^0-9a-f]/g, '');
  if (hex.length !== 12) return '';
  return hex.match(/.{2}/g).join(':');
}
function normalizeIp(ip) {
  let s = String(ip || '').trim().toLowerCase();
  if (s.startsWith('::ffff:')) s = s.slice(7);
  return s;
}
const SIGNATURE_ONLY_MACS = new Set(parseListEnv('AIS_SIGNATURE_ONLY_MACS').map(normalizeMac).filter(Boolean));
const SIGNATURE_ONLY_IPS = new Set(parseListEnv('AIS_SIGNATURE_ONLY_IPS').map(normalizeIp).filter(Boolean));

function clientIp(req) {
  return normalizeIp(req.socket && req.socket.remoteAddress);
}

/** Best-effort MAC from kernel ARP/neighbor table (LAN only). Cached briefly. */
let arpCache = { at: 0, map: new Map() };
const ARP_CACHE_MS = 15_000;

function refreshArpCache(force = false) {
  const now = Date.now();
  if (!force && now - arpCache.at < ARP_CACHE_MS && arpCache.map.size >= 0) {
    return arpCache.map;
  }
  const map = new Map();
  try {
    const text = fs.readFileSync('/proc/net/arp', 'utf8');
    for (const line of text.split('\n').slice(1)) {
      const parts = line.trim().split(/\s+/);
      if (parts.length < 4) continue;
      const ip = normalizeIp(parts[0]);
      const mac = normalizeMac(parts[3]);
      if (ip && mac && mac !== '00:00:00:00:00:00') map.set(ip, mac);
    }
  } catch { /* non-Linux or unreadable */ }
  arpCache = { at: now, map };
  return map;
}

function macForIp(ip) {
  const target = normalizeIp(ip);
  if (!target || target === '127.0.0.1') return '';
  // Skip ARP entirely when no MAC-based kiosk rules are configured.
  if (!SIGNATURE_ONLY_MACS.size) return '';
  return refreshArpCache().get(target) || '';
}

function clientRole(req) {
  const ip = clientIp(req);
  if (SIGNATURE_ONLY_IPS.has(ip)) return 'signature';
  if (!SIGNATURE_ONLY_MACS.size) return 'full';
  const mac = macForIp(ip);
  if (mac && SIGNATURE_ONLY_MACS.has(mac)) return 'signature';
  return 'full';
}

function isSignatureAllowedPath(pathname, method) {
  const m = (method || 'GET').toUpperCase();
  if (pathname === '/api/status' && m === 'GET') return true;
  if (pathname === '/api/whoami' && m === 'GET') return true;
  if (pathname === '/api/sign/queue' && m === 'GET') return true;
  if (/^\/api\/ticket\/\d+$/.test(pathname) && m === 'GET') return true;
  if (/^\/api\/ticket\/\d+\/signature\.png$/.test(pathname) && m === 'GET') return true;
  if (/^\/api\/ticket\/\d+\/signature$/.test(pathname) && m === 'POST') return true;
  // Pages + static assets for the kiosk UI
  if (m === 'GET') {
    if (pathname === '/' || pathname === '/sign' || pathname === '/sign.html') return true;
    if (pathname === '/auth.js' || pathname === '/sig-compress.js' || pathname === '/style.css' || pathname === '/ais.png') return true;
    if (pathname.startsWith('/public/')) return true;
  }
  return false;
}

function denySignatureOnly(res, pathname) {
  return json(res, 403, {
    error: 'signature-only device: access denied',
    path: pathname,
    hint: 'This tablet may only use /sign.html (driver signature).'
  });
}

function ticketForSignPad(detail) {
  if (!detail) return null;
  return {
    id: detail.id,
    weighmasterNumber: detail.weighmasterNumber || detail.weighmaster || '',
    invoiceNumber: detail.invoiceNumber || '',
    transactionDate: detail.transactionDate || '',
    customer: detail.customer || '',
    truckId: detail.truckId || '',
    productService: detail.productService || '',
    city: detail.city || '',
    t1: detail.t1,
    t2: detail.t2,
    tonage: detail.tonage,
    byPerson: detail.byPerson || '',
    hasDriverSignature: !!detail.hasDriverSignature
  };
}



function json(res, status, value, extraHeaders) {
  const body = JSON.stringify(value);
  const headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body)
  };
  if (extraHeaders && typeof extraHeaders === 'object') {
    for (const [k, v] of Object.entries(extraHeaders)) {
      if (v != null) headers[k] = v;
    }
  }
  res.writeHead(status, headers);
  res.end(body);
}

function text(res, status, body, contentType = 'text/plain; charset=utf-8') {
  res.writeHead(status, {
    'Content-Type': contentType,
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

function fileBytes(res, status, bytes, contentType, fileName) {
  const headers = {
    'Content-Type': contentType,
    'Content-Length': bytes.length
  };
  if (fileName) headers['Content-Disposition'] = `attachment; filename="${fileName}"`;
  res.writeHead(status, headers);
  res.end(bytes);
}

function errorJson(res, status, message) {
  json(res, status, { error: message });
}

// ---------- Simple login: password once → cookie session for everything ----------
// Flow: POST /api/login → Set-Cookie ais_session → browser sends cookie until Logout
//   or expiry. No re-prompt. Signature tablet still bypasses via MAC/IP allowlist.
//
// Env:
//   AIS_DASHBOARD_USER          default admin
//   AIS_DASHBOARD_PASSWORD      required for login (default ais)
//   AIS_SESSION_HOURS          default 168 (7 days) — stay signed in
//   AIS_AUTH_DISABLED=1        open access, no login
const AUTH_DISABLED = /^(1|true|yes)$/i.test(String(process.env.AIS_AUTH_DISABLED || ''));
const AUTH_USER = String(process.env.AIS_DASHBOARD_USER || 'admin').trim() || 'admin';
const AUTH_PASSWORD = String(
  process.env.AIS_DASHBOARD_PASSWORD ||
  process.env.AIS_DASHBOARD_TOKEN ||
  (AUTH_DISABLED ? '' : 'ais')
);
const AUTH_ENABLED = !AUTH_DISABLED && AUTH_PASSWORD.length > 0;
const SESSION_TTL_MS = Math.max(
  60 * 60 * 1000,
  Number(process.env.AIS_SESSION_HOURS || 168) * 60 * 60 * 1000
);
const SESSION_COOKIE = 'ais_session';
const SESSION_SECRET = crypto
  .createHash('sha256')
  .update(`ais-session|${AUTH_USER}|${AUTH_PASSWORD}|v2`)
  .digest();

function safeEqualString(a, b) {
  const sa = String(a);
  const sb = String(b);
  if (sa.length !== sb.length) return false;
  try {
    return crypto.timingSafeEqual(Buffer.from(sa, 'utf8'), Buffer.from(sb, 'utf8'));
  } catch {
    return false;
  }
}

function readSessionCookie(req) {
  const header = req.headers.cookie;
  if (!header) return '';
  const key = SESSION_COOKIE + '=';
  let start = header.indexOf('; ' + key);
  if (start >= 0) start += 2;
  else if (header.startsWith(key)) start = 0;
  else {
    start = header.indexOf(key);
    if (start < 0) return '';
  }
  start += key.length;
  let end = header.indexOf(';', start);
  if (end < 0) end = header.length;
  try {
    return decodeURIComponent(header.slice(start, end).trim());
  } catch {
    return header.slice(start, end).trim();
  }
}

function createSessionToken(user) {
  const exp = Date.now() + SESSION_TTL_MS;
  const payload = `${String(user || AUTH_USER)}|${exp}`;
  const sig = crypto.createHmac('sha256', SESSION_SECRET).update(payload).digest('base64url');
  return Buffer.from(`${payload}|${sig}`, 'utf8').toString('base64url');
}

function verifySessionToken(token) {
  if (!token) return null;
  let raw;
  try {
    raw = Buffer.from(token, 'base64url').toString('utf8');
  } catch {
    return null;
  }
  const parts = raw.split('|');
  if (parts.length !== 3) return null;
  const [user, expStr, sig] = parts;
  const exp = Number(expStr);
  if (!user || !Number.isFinite(exp) || exp <= Date.now()) return null;
  const payload = `${user}|${exp}`;
  const expect = crypto.createHmac('sha256', SESSION_SECRET).update(payload).digest('base64url');
  if (expect.length !== String(sig).length) return null;
  try {
    if (!crypto.timingSafeEqual(Buffer.from(expect), Buffer.from(sig))) return null;
  } catch {
    return null;
  }
  return { user, expires: exp, auth: true };
}

/** Logged-in staff session, or null. Memoized on req. */
function getSession(req) {
  if (req._aisSession !== undefined) return req._aisSession;
  if (!AUTH_ENABLED) {
    req._aisSession = { user: 'open', auth: true }; // open mode = always "in"
    return req._aisSession;
  }
  req._aisSession = verifySessionToken(readSessionCookie(req));
  return req._aisSession;
}

function isLoggedIn(req) {
  const s = getSession(req);
  return !!(s && s.auth);
}

function sessionCookieHeader(token) {
  const maxAge = Math.floor(SESSION_TTL_MS / 1000);
  return [
    `${SESSION_COOKIE}=${encodeURIComponent(token)}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    `Max-Age=${maxAge}`
  ].join('; ');
}

function clearSessionCookieHeader() {
  return `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`;
}

function isPublicPath(pathname, method) {
  const m = (method || 'GET').toUpperCase();
  if (pathname === '/api/login' && m === 'POST') return true;
  if (pathname === '/api/auth/status' && m === 'GET') return true;
  if (pathname === '/api/status' && m === 'GET') return true;
  if (m === 'GET') {
    if (pathname === '/login' || pathname === '/login.html') return true;
    if (pathname === '/auth.js' || pathname === '/sig-compress.js' || pathname === '/style.css' || pathname === '/ais.png') return true;
    if (pathname.startsWith('/public/')) return true;
  }
  return false;
}

/**
 * One gate for pages + APIs.
 * After login, cookie is enough for the whole session (reads and writes).
 */
function requireLogin(req, res, url, { htmlRedirect = false } = {}) {
  if (!AUTH_ENABLED) return true;
  const pathname = url.pathname === '/sign' ? '/sign.html' : (url.pathname || '/');
  if (isPublicPath(pathname, req.method)) return true;
  if (isLoggedIn(req)) return true;
  // Signature tablet (MAC/IP) — no staff login
  if (clientRole(req) === 'signature' && isSignatureAllowedPath(pathname, req.method)) {
    return true;
  }
  if (htmlRedirect && req.method === 'GET') {
    const next = pathname + (url.search || '');
    res.writeHead(302, { Location: '/login.html?next=' + encodeURIComponent(next || '/') });
    res.end();
    return false;
  }
  errorJson(res, 401, 'login required');
  return false;
}

/** Writes use the same session as pages — no second password/token. */
function requireToken(req, res) {
  if (!AUTH_ENABLED) return true;
  if (isLoggedIn(req)) return true;
  errorJson(res, 401, 'login required');
  return false;
}

function readBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on('data', chunk => {
      total += chunk.length;
      if (total > maxBytes) {
        reject(new Error('request body too large'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

async function readJson(req) {
  const body = await readBody(req, MAX_JSON_BYTES);
  if (body.length === 0) return {};
  return JSON.parse(body.toString('utf8'));
}

function todayIso() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function parseDate(value) {
  if (!value || !String(value).trim()) return null;
  const textValue = String(value).trim();
  const match = textValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    return {
      year: Number(match[1]),
      month: Number(match[2]),
      day: Number(match[3]),
      iso: `${match[1]}-${String(Number(match[2])).padStart(2, '0')}-${String(Number(match[3])).padStart(2, '0')}`
    };
  }
  const parsed = new Date(textValue);
  if (Number.isNaN(parsed.getTime())) return null;
  return {
    year: parsed.getFullYear(),
    month: parsed.getMonth() + 1,
    day: parsed.getDate(),
    iso: `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`
  };
}

function dateToTime(date) {
  return Date.UTC(date.year, date.month - 1, date.day);
}

function monthOf(date) {
  return date ? `${date.year}-${String(date.month).padStart(2, '0')}` : null;
}

function isoWeek(date) {
  if (!date) return null;
  const d = new Date(Date.UTC(date.year, date.month - 1, date.day));
  const day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return week;
}

function weekOf(date) {
  const week = isoWeek(date);
  return week ? `${date.year}-W${String(week).padStart(2, '0')}` : null;
}

function biweeklyOf(date) {
  const week = isoWeek(date);
  return week ? `${date.year}-B${String(Math.floor((week + 1) / 2)).padStart(2, '0')}` : null;
}

function cleanText(value) {
  if (value === undefined || value === null) return null;
  const trimmed = String(value).trim();
  return trimmed ? trimmed : null;
}

function cleanNumber(value) {
  if (value === undefined || value === null || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const cleaned = String(value).replace(/[$,%]/g, '').replace(/,/g, '').trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function nullToEmpty(value) {
  return value === null || value === undefined ? '' : value;
}

function distinctValues(db, column) {
  const allowed = new Set([
    'customer', 'product_service', 'weighmaster', 'truck_id', 'class',
    'location', 'deposit_to', 'payment', 'po_origin', 'origin', 'city'
  ]);
  if (!allowed.has(column)) throw new Error(`Column not allowed: ${column}`);
  return db.prepare(`
    SELECT DISTINCT ${column} AS value
    FROM quickbooks_transactions
    WHERE ${column} IS NOT NULL AND TRIM(${column}) <> ''
    ORDER BY ${column} COLLATE NOCASE;
  `).all().map(r => r.value);
}

/**
 * Numbering model (QuickBooks-style):
 *   - Weighmaster ticket #  — always sequential (scale ticket)
 *   - AR           → document_type "Invoice"        + own Invoice # series
 *   - CASH/CREDIT/DEBIT → document_type "Sales Receipt" + own Sales Receipt # series
 *   - EXPENSE      → document_type "Expense"        + own Expense # series
 *
 * Series starts (first issued = start + 1), overridable via env:
 *   AIS_TICKET_START, AIS_INVOICE_START, AIS_SALES_RECEIPT_START, AIS_EXPENSE_START
 */
const TICKET_NUMBER_START = Number(process.env.AIS_TICKET_START || 117964);
// Document number series by payment:
//   AR      → Invoice #     (numeric 910937+)
//   CASH/…  → Sales Receipt #  as SR1, SR2, SR3, …
//   EXPENSE → Expense #     (numeric 700001+)
const DOC_SERIES = {
  Invoice: {
    type: 'Invoice',
    label: 'Invoice #',
    format: 'numeric',
    startFrom: Number(process.env.AIS_INVOICE_START || 910936),
    min: Number(process.env.AIS_INVOICE_MIN || 900000),
    max: Number(process.env.AIS_INVOICE_MAX || 999999999)
  },
  'Sales Receipt': {
    type: 'Sales Receipt',
    label: 'Sales Receipt #',
    format: 'sr', // SR1, SR2, …
    prefix: String(process.env.AIS_SALES_RECEIPT_PREFIX || 'SR').toUpperCase(),
    startFrom: Number(process.env.AIS_SALES_RECEIPT_START || 0) // first = SR1
  },
  Expense: {
    type: 'Expense',
    label: 'Expense #',
    format: 'numeric',
    startFrom: Number(process.env.AIS_EXPENSE_START || 700000),
    min: Number(process.env.AIS_EXPENSE_MIN || 700000),
    max: Number(process.env.AIS_EXPENSE_MAX || 799999)
  }
};

function normalizePaymentMethod(value) {
  const p = String(value || 'CASH').trim().toUpperCase();
  if (p === 'AR' || p === 'ACCOUNT') return 'AR';
  if (p === 'EXPENSE') return 'EXPENSE';
  if (p === 'CREDIT') return 'CREDIT';
  if (p === 'DEBIT') return 'DEBIT';
  if (p === 'CASH') return 'CASH';
  return p || 'CASH';
}

/** Map payment → QB document series. */
function documentSeriesForPayment(payment) {
  const p = normalizePaymentMethod(payment);
  if (p === 'AR') return DOC_SERIES.Invoice;
  if (p === 'EXPENSE') return DOC_SERIES.Expense;
  return DOC_SERIES['Sales Receipt']; // CASH, CREDIT, DEBIT, default
}

/**
 * Gap-free sequence: fill the lowest hole in [minUsed … maxUsed], else max+1.
 * If no numbers yet → startFrom + 1.
 */
function nextInSequence(usedNumbers, startFrom) {
  const used = new Set();
  let min = Infinity;
  let max = -Infinity;
  for (const raw of usedNumbers) {
    const n = Number(raw);
    if (!Number.isFinite(n)) continue;
    used.add(n);
    if (n < min) min = n;
    if (n > max) max = n;
  }
  if (used.size === 0) return startFrom + 1;
  for (let n = min; n <= max; n++) {
    if (!used.has(n)) return n;
  }
  return max + 1;
}

function collectWeighmasterNumbers(db) {
  return db.prepare(`
    SELECT CAST(weighmaster_number AS INTEGER) AS value
    FROM quickbooks_transactions
    WHERE weighmaster_number IS NOT NULL
      AND TRIM(weighmaster_number) <> ''
      AND TRIM(weighmaster_number) GLOB '[0-9]*'
    UNION
    SELECT CAST(weighmaster AS INTEGER) AS value
    FROM quickbooks_transactions
    WHERE weighmaster IS NOT NULL
      AND TRIM(weighmaster) <> ''
      AND TRIM(weighmaster) GLOB '[0-9]*'
      AND CAST(weighmaster AS INTEGER) >= 100000
  `).all().map(r => r.value);
}

/** All document numbers as stored strings (for global uniqueness checks). */
function collectAllDocumentNumberStrings(db) {
  const rows = db.prepare(`
    SELECT TRIM(invoice_number) AS value FROM quickbooks_transactions
    WHERE invoice_number IS NOT NULL AND TRIM(invoice_number) <> ''
    UNION
    SELECT TRIM(document_number) AS value FROM quickbooks_transactions
    WHERE document_number IS NOT NULL AND TRIM(document_number) <> ''
  `).all();
  return rows.map(r => String(r.value));
}

/** Payment filter SQL for a document series (Invoice = AR only). */
function seriesPaymentSql(seriesType) {
  if (seriesType === 'Invoice') {
    return `UPPER(TRIM(COALESCE(payment, ''))) IN ('AR', 'ACCOUNT')`;
  }
  if (seriesType === 'Expense') {
    return `UPPER(TRIM(COALESCE(payment, ''))) = 'EXPENSE'`;
  }
  return `UPPER(TRIM(COALESCE(payment, ''))) IN ('CASH', 'CREDIT', 'DEBIT')
    OR TRIM(COALESCE(payment, '')) = ''`;
}

/**
 * Numeric sequence values already used by a series (for gap-free next).
 * Sales Receipt parses SR1 → 1, SR2 → 2, etc.
 */
function collectSeriesSequenceValues(db, series) {
  const paySql = seriesPaymentSql(series.type);
  const rows = db.prepare(`
    SELECT TRIM(invoice_number) AS value FROM quickbooks_transactions
    WHERE invoice_number IS NOT NULL AND TRIM(invoice_number) <> '' AND (${paySql})
    UNION
    SELECT TRIM(document_number) AS value FROM quickbooks_transactions
    WHERE document_number IS NOT NULL AND TRIM(document_number) <> '' AND (${paySql})
  `).all().map(r => String(r.value));

  if (series.format === 'sr') {
    const prefix = (series.prefix || 'SR').toUpperCase();
    const re = new RegExp(`^${prefix}(\\d+)$`, 'i');
    return rows.map(v => {
      const m = re.exec(String(v).trim());
      return m ? Number(m[1]) : null;
    }).filter(n => Number.isFinite(n));
  }

  // Numeric series: only pure digits inside [min, max]
  return rows.map(v => {
    if (!/^\d+$/.test(v)) return null;
    const n = Number(v);
    if (series.min != null && n < series.min) return null;
    if (series.max != null && n > series.max) return null;
    return n;
  }).filter(n => Number.isFinite(n));
}

function formatSeriesNumber(series, n) {
  if (series.format === 'sr') {
    return `${series.prefix || 'SR'}${n}`;
  }
  return String(n);
}

/** Next weighmaster ticket # (gap-free in ticket series). */
function nextWeighmasterNumber(db, startFrom = TICKET_NUMBER_START) {
  return String(nextInSequence(collectWeighmasterNumbers(db), startFrom));
}

/**
 * Next document # for payment:
 *   AR → Invoice 910943…
 *   CASH/CREDIT/DEBIT → SR1, SR2, …
 *   EXPENSE → 700001…
 */
function nextDocumentNumber(db, payment) {
  const series = documentSeriesForPayment(payment);
  const usedSeq = collectSeriesSequenceValues(db, series);
  const allUsed = new Set(
    collectAllDocumentNumberStrings(db).map(s => s.toUpperCase())
  );

  let n = nextInSequence(usedSeq, series.startFrom);
  if (series.format !== 'sr') {
    if (series.min != null && n < series.min) n = series.startFrom + 1;
    if (n <= series.startFrom) n = series.startFrom + 1;
  } else if (n < 1) {
    n = 1;
  }

  for (let guard = 0; guard < 1000; guard++) {
    if (series.max != null && n > series.max) {
      throw new Error(`${series.type} number range exhausted`);
    }
    const formatted = formatSeriesNumber(series, n);
    if (!allUsed.has(formatted.toUpperCase())
      && !identifierExists(db, 'invoice_number', formatted, null)) {
      return { number: formatted, ...series };
    }
    n += 1;
  }
  throw new Error(`Could not allocate ${series.type} number`);
}

/** Next AR Invoice # only. */
function nextInvoiceNumber(db) {
  return nextDocumentNumber(db, 'AR').number;
}

/**
 * Atomically reserve next free ticket # + document # for this payment.
 * AR → Invoice #; CASH/CREDIT/DEBIT → SR1…; EXPENSE → Expense #.
 */
function allocateTicketNumbers(db, payment = 'CASH') {
  const pay = normalizePaymentMethod(payment);
  const series = documentSeriesForPayment(pay);

  let wmNum = Number(nextWeighmasterNumber(db));
  if (!Number.isFinite(wmNum)) wmNum = TICKET_NUMBER_START + 1;
  for (let i = 0; i < 100 && identifierExists(db, 'weighmaster_number', String(wmNum), null); i++) {
    wmNum += 1;
  }

  const doc = nextDocumentNumber(db, pay);

  return {
    weighmasterNumber: String(wmNum),
    invoiceNumber: doc.number,
    documentNumber: doc.number,
    documentType: series.type,
    documentLabel: series.label,
    payment: pay
  };
}

function withImmediateTransaction(db, fn) {
  db.exec('BEGIN IMMEDIATE');
  try {
    const result = fn();
    db.exec('COMMIT');
    return result;
  } catch (err) {
    try { db.exec('ROLLBACK'); } catch { /* ignore */ }
    throw err;
  }
}

function identifierExists(db, column, value, excludeId) {
  const cleaned = cleanText(value);
  if (!cleaned) return false;
  if (column !== 'weighmaster_number' && column !== 'invoice_number') {
    throw new Error(`Column not allowed: ${column}`);
  }

  const params = [cleaned];
  const idClause = excludeId ? 'AND id <> ?' : '';
  if (excludeId) params.push(Number(excludeId));

  const sql = column === 'invoice_number'
    ? `SELECT 1 AS found FROM quickbooks_transactions
       WHERE (invoice_number = ? OR document_number = ?)
       ${excludeId ? 'AND id <> ?' : ''}
       LIMIT 1;`
    : `SELECT 1 AS found FROM quickbooks_transactions
       WHERE (
         weighmaster_number = ?
         OR (
           (weighmaster_number IS NULL OR TRIM(weighmaster_number) = '')
           AND weighmaster IS NOT NULL
           AND TRIM(weighmaster) = ?
           AND TRIM(weighmaster) NOT GLOB '*[^0-9]*'
         )
       )
       ${idClause}
       LIMIT 1;`;

  const finalParams = column === 'invoice_number'
    ? (excludeId ? [cleaned, cleaned, Number(excludeId)] : [cleaned, cleaned])
    : (excludeId ? [cleaned, cleaned, Number(excludeId)] : [cleaned, cleaned]);
  return Boolean(db.prepare(sql).get(...finalParams));
}

function nextManualRowNumber(db, sourceFile) {
  const row = db.prepare(`
    SELECT COALESCE(MAX(source_row_number), 0) + 1 AS next
    FROM quickbooks_transactions
    WHERE source_file = ?;
  `).get(sourceFile);
  return row.next;
}

function loadWeighmasterTickets(db, limit = 500) {
  // Never SELECT * here — driver_signature BLOBs would be read on every list/refresh.
  return db.prepare(`
    WITH ticket_rows AS (
      SELECT
        id, weighmaster_number, weighmaster, invoice_number, document_number,
        transaction_date, customer, truck_id, product_service, qty, tonage,
        rate, amount, cash, card, ar, payment, time_in, time_out, t1, t2,
        origin, po_origin, city, waste_data, memo, ticket_notes, by_person,
        weight_certificate, recycling_cd_certificate,
        COALESCE(
          NULLIF(TRIM(weighmaster_number), ''),
          CASE
            WHEN TRIM(COALESCE(weighmaster, '')) <> ''
                 AND TRIM(weighmaster) NOT GLOB '*[^0-9]*'
            THEN TRIM(weighmaster)
          END
        ) AS ticket_number
      FROM quickbooks_transactions
    )
    SELECT CAST(id AS INTEGER) AS id,
           CAST(ticket_number AS TEXT) AS ticketNo,
           CAST(COALESCE(invoice_number, document_number) AS TEXT) AS invoiceNo,
           CAST(transaction_date AS TEXT) AS date,
           CAST(customer AS TEXT) AS customer,
           CAST(truck_id AS TEXT) AS truck,
           CAST(product_service AS TEXT) AS product,
           CAST(COALESCE(qty, tonage) AS REAL) AS tons,
           CAST(rate AS REAL) AS rate,
           CAST(amount AS REAL) AS amount,
           CAST(cash AS REAL) AS cash,
           CAST(card AS REAL) AS card,
           CAST(ar AS REAL) AS ar,
           CAST(payment AS TEXT) AS payment,
           CAST(time_in AS TEXT) AS timeIn,
           CAST(time_out AS TEXT) AS timeOut,
           CAST(t1 AS REAL) AS gross,
           CAST(t2 AS REAL) AS tare,
           CAST(origin AS TEXT) AS origin,
           CAST(po_origin AS TEXT) AS wasteOrigin,
           CAST(city AS TEXT) AS city,
           CAST(waste_data AS TEXT) AS wasteData,
           CAST(memo AS TEXT) AS memo,
           CAST(ticket_notes AS TEXT) AS ticketNotes,
           CAST(weighmaster AS TEXT) AS weighmaster,
           CAST(by_person AS TEXT) AS byPerson,
           CASE WHEN COALESCE(weight_certificate, 0) <> 0 THEN 'Yes' ELSE '' END AS weightCertificate,
           CASE WHEN COALESCE(recycling_cd_certificate, 0) <> 0 THEN 'Yes' ELSE '' END AS recyclingCdCertificate
    FROM ticket_rows
    WHERE ticket_number IS NOT NULL
      AND TRIM(ticket_number) <> ''
      AND transaction_date IS NOT NULL
      AND TRIM(transaction_date) <> ''
    ORDER BY id DESC
    LIMIT ?;
  `).all(limit);
}

function loadTickets(db) {
  return loadWeighmasterTickets(db, 2147483647).map(row => {
    const amount = cleanNumber(row.amount);
    const ar = cleanNumber(row.ar);
    const payment = row.payment || '';
    const isAr = payment.toUpperCase() === 'AR' || (ar || 0) !== 0;
    return {
      id: row.id,
      ticketNo: row.ticketNo || '',
      invoiceNo: row.invoiceNo || '',
      date: row.date || '',
      customer: row.customer || '',
      truck: row.truck || '',
      product: row.product || '',
      tons: cleanNumber(row.tons),
      rate: cleanNumber(row.rate),
      amount,
      cash: cleanNumber(row.cash),
      card: cleanNumber(row.card),
      ar,
      payment,
      city: row.city || '',
      origin: row.origin || row.wasteOrigin || '',
      memo: row.memo || '',
      isAr,
      arAmount: ar ?? (isAr ? amount ?? 0 : 0)
    };
  });
}

function buildSummary(tickets) {
  return {
    ticketCount: tickets.length,
    customerCount: new Set(tickets.map(t => (t.customer || '').toLowerCase()).filter(Boolean)).size,
    totalTons: sum(tickets, 'tons'),
    totalAmount: sum(tickets, 'amount'),
    arTotal: tickets.reduce((s, t) => s + (t.arAmount || 0), 0),
    cashTotal: sum(tickets, 'cash'),
    cardTotal: sum(tickets, 'card'),
    generatedAt: new Date().toISOString()
  };
}

function buildDaySummary(tickets, targetDate) {
  const dayRows = tickets.filter(t => (t.date || '').slice(0, 10) === targetDate);
  const runSplit = matcher => {
    const hits = dayRows.filter(matcher);
    return {
      in: hits.filter(r => (r.amount || 0) > 0).length,
      out: hits.filter(r => (r.amount || 0) < 0).length
    };
  };
  const isCustomer = (ticket, tokens) => {
    const upper = (ticket.customer || '').toUpperCase();
    return tokens.some(token => upper.includes(token));
  };
  const edco = runSplit(t => isCustomer(t, ['EDCO']));
  const republic = runSplit(t => isCustomer(t, ['REPUBLIC']));
  const wcsg = runSplit(t => isCustomer(t, ['WEST COAST SAND', 'WEST COAST SAND&GRAVEL']));
  const tonsIn = dayRows.filter(t => (t.amount || 0) > 0).reduce((s, t) => s + (t.tons || 0), 0);
  const tonsOut = -Math.abs(dayRows.filter(t => (t.amount || 0) < 0).reduce((s, t) => s + (t.tons || 0), 0));
  const amountIn = dayRows.filter(t => (t.amount || 0) > 0).reduce((s, t) => s + (t.amount || 0), 0);
  const amountOut = dayRows.filter(t => (t.amount || 0) < 0).reduce((s, t) => s + (t.amount || 0), 0);
  return {
    date: targetDate,
    ticketCount: dayRows.length,
    tons: { in: round2(tonsIn), out: round2(tonsOut), net: round2(tonsIn + tonsOut) },
    amount: { in: round2(amountIn), out: round2(amountOut), net: round2(amountIn + amountOut) },
    runs: {
      edco: { in: edco.in, out: edco.out, net: edco.in - edco.out },
      republic: { in: republic.in, out: republic.out, net: republic.in - republic.out },
      wcsg: { in: wcsg.in, out: wcsg.out, net: wcsg.in - wcsg.out }
    },
    wcsgCount: wcsg.in + wcsg.out,
    tickets: dayRows
  };
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (row[key] || 0), 0);
}

function round2(value) {
  return Math.round(value * 100) / 100;
}

function loadRecent(db, limit = 100) {
  return db.prepare(`
    SELECT id,
           transaction_date AS date,
           document_type AS type,
           COALESCE(invoice_number, document_number) AS docNo,
           customer AS customer,
           product_service AS product,
           qty AS qty,
           rate AS rate,
           amount AS amount,
           truck_id AS truck,
           weighmaster AS weighmaster,
           memo AS memo
    FROM quickbooks_transactions
    ORDER BY id DESC
    LIMIT ?;
  `).all(limit);
}

function loadTicketById(db, id) {
  // Exclude signature BLOB from normal loads (use signature.png / dedicated queries).
  return db.prepare(`
    SELECT id, source_file, source_row_number, document_type, document_number,
           date_copy, waste_data, memo, truck_id, po_origin, transaction_date,
           customer, weighmaster, product_service, qty, rate, amount, payment, ar,
           week, biweekly, class, scalemaster, month, invoice_date, description,
           deposit_to, location, round_amount, time_in, time_out, t1, t2, tonage,
           origin, cash, card, imported_at_utc, weighmaster_number, city,
           invoice_number, ticket_notes, weight_certificate, recycling_cd_certificate,
           by_person, payment_terms, due_date, project_address_id,
           CASE WHEN driver_signature IS NOT NULL AND length(driver_signature) > 0
                THEN 1 ELSE 0 END AS has_driver_signature
    FROM quickbooks_transactions
    WHERE id = ?;
  `).get(Number(id));
}

/** Parse "Net 30" / "NET10" / "30" → { label, days }. */
function parsePaymentTerms(value) {
  const t = cleanText(value);
  if (!t) return null;
  const m = /net\s*[-]?\s*(\d+)/i.exec(t) || /^(\d+)\s*days?$/i.exec(t) || /^(\d+)$/.exec(t);
  if (!m) return { label: t, days: null };
  const days = Number(m[1]);
  if (!Number.isFinite(days) || days < 0) return { label: t, days: null };
  return { label: `Net ${days}`, days };
}

/** transactionDate ISO yyyy-mm-dd + days → due date ISO (local calendar). */
function computeDueDateIso(transactionDateIso, days) {
  if (days == null || !Number.isFinite(Number(days))) return null;
  const d = parseDate(transactionDateIso);
  if (!d) return null;
  const dt = new Date(d.year, d.month - 1, d.day);
  dt.setDate(dt.getDate() + Number(days));
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}


function rowToTicketDetail(row) {
  const result = {};
  for (const [key, value] of Object.entries(row)) {
    if (key === 'driver_signature' || key === 'has_driver_signature') continue;
    result[toCamelCase(key)] = value;
  }
  result.hasDriverSignature = !!(row.has_driver_signature || sigBytes(row.driver_signature));
  return result;
}

function toCamelCase(value) {
  return value.replace(/_([a-z])/g, (_, c) => c.toUpperCase());
}

function normalizeAccountType(value) {
  const v = cleanText(value) || '';
  if (/^(account|ar)$/i.test(v)) return 'AR';
  if (/^cash$/i.test(v)) return 'Cash';
  if (/^credit$/i.test(v)) return 'Credit';
  if (/^debit$/i.test(v)) return 'Debit';
  if (/^expense$/i.test(v)) return 'Expense';
  return 'Regular';
}

function paymentMethodFor(accountType) {
  switch (normalizeAccountType(accountType)) {
    case 'AR': return 'AR';
    case 'Cash': return 'CASH';
    case 'Credit': return 'CREDIT';
    case 'Debit': return 'DEBIT';
    case 'Expense': return 'EXPENSE';
    default: return null;
  }
}

function normalizePricingType(value) {
  if (/^hauling$/i.test(cleanText(value) || '')) return 'Hauling';
  if (/^run$/i.test(cleanText(value) || '')) return 'Run';
  return 'Flat';
}

function normalizeRuleDate(raw, onDate) {
  const parsed = parseDate(raw);
  if (!parsed) return null;
  if (parsed.year < 2020) {
    const anchorYear = onDate?.year || new Date().getFullYear();
    return { ...parsed, year: anchorYear, iso: `${anchorYear}-${String(parsed.month).padStart(2, '0')}-${String(parsed.day).padStart(2, '0')}` };
  }
  return parsed;
}

function findPricingRule(db, customer, product, onDate, className, projectAddressId) {
  const productValue = cleanText(product);
  if (!productValue) return null;

  const customerValue = cleanText(customer);
  const classValue = cleanText(className);
  const projectId = Number(projectAddressId) > 0 ? Number(projectAddressId) : null;
  const rows = db.prepare(`
    SELECT id, customer, class, rate, account_type, pricing_type,
           base_rate, allowed_tons, overage_rate, miles, mile_rate, surcharge_percent,
           effective_date, end_date, payment_terms, project_address_id
    FROM pricing
    WHERE TRIM(product_service) = TRIM(?) COLLATE NOCASE
      AND (customer IS NULL OR ? IS NULL OR TRIM(customer) = TRIM(?) COLLATE NOCASE)
      AND (class IS NULL OR TRIM(class) = '' OR ? IS NULL OR TRIM(class) = TRIM(?) COLLATE NOCASE)
      AND (
        -- Generic rules always match; project-scoped only when that project is selected.
        project_address_id IS NULL
        OR (? IS NOT NULL AND project_address_id = ?)
      );
  `).all(
    productValue,
    customerValue, customerValue,
    classValue, classValue,
    projectId, projectId
  );

  const onTime = onDate ? dateToTime(onDate) : null;
  const candidates = rows.map(row => {
    const start = normalizeRuleDate(row.effective_date, onDate);
    const end = normalizeRuleDate(row.end_date, onDate);
    return { row, start, end };
  }).filter(c => {
    if (onTime === null) return true;
    if (c.start && onTime < dateToTime(c.start)) return false;
    if (c.end && onTime > dateToTime(c.end)) return false;
    return true;
  });

  if (candidates.length === 0) return null;
  candidates.sort((a, b) => {
    // Prefer project-specific rules when a project is selected.
    const aProj = projectId && Number(a.row.project_address_id) === projectId ? 0 : 1;
    const bProj = projectId && Number(b.row.project_address_id) === projectId ? 0 : 1;
    if (aProj !== bProj) return aProj - bProj;
    // When no project selected, prefer generic (null project) over project-scoped.
    if (!projectId) {
      const aHas = a.row.project_address_id != null ? 1 : 0;
      const bHas = b.row.project_address_id != null ? 1 : 0;
      if (aHas !== bHas) return aHas - bHas;
    }
    const aCustomer = eq(a.row.customer, customerValue) ? 0 : 1;
    const bCustomer = eq(b.row.customer, customerValue) ? 0 : 1;
    if (aCustomer !== bCustomer) return aCustomer - bCustomer;
    const aClass = eq(a.row.class, classValue) ? 0 : 1;
    const bClass = eq(b.row.class, classValue) ? 0 : 1;
    if (aClass !== bClass) return aClass - bClass;
    const aStart = a.start ? dateToTime(a.start) : -8640000000000000;
    const bStart = b.start ? dateToTime(b.start) : -8640000000000000;
    if (aStart !== bStart) return bStart - aStart;
    return Number(b.row.id) - Number(a.row.id);
  });

  const best = candidates[0].row;
  return {
    rate: cleanNumber(best.rate) || 0,
    accountType: normalizeAccountType(best.account_type),
    pricingType: normalizePricingType(best.pricing_type),
    baseRate: cleanNumber(best.base_rate),
    allowedTons: cleanNumber(best.allowed_tons),
    overageRate: cleanNumber(best.overage_rate),
    miles: cleanNumber(best.miles),
    mileRate: cleanNumber(best.mile_rate),
    surchargePercent: cleanNumber(best.surcharge_percent),
    paymentTerms: cleanText(best.payment_terms),
    projectAddressId: best.project_address_id != null ? Number(best.project_address_id) : null
  };
}

function eq(a, b) {
  const aa = cleanText(a);
  const bb = cleanText(b);
  return aa !== null && bb !== null && aa.toLowerCase() === bb.toLowerCase();
}

function calculateAmount(rule, tons) {
  // % surcharge applies to every pricing type (fuel/environmental etc.)
  const surchargeMult = 1 + Math.max(rule.surchargePercent || 0, 0) / 100;
  if (rule.pricingType === 'Run') return rule.rate * surchargeMult;
  if (rule.pricingType !== 'Hauling') return rule.rate * tons * surchargeMult;
  const baseAmount = rule.baseRate || 0;
  const allowed = rule.allowedTons || 0;
  const overageTons = Math.max(tons - allowed, 0);
  const overageAmount = overageTons * (rule.overageRate || 0);
  const distanceAmount = (rule.miles || 0) * (rule.mileRate || 0);
  const subtotal = baseAmount + overageAmount + distanceAmount;
  return subtotal * surchargeMult;
}

function describePricing(rule, tons) {
  const sc = rule.surchargePercent || 0;
  const scText = sc > 0 ? ` + ${compact(sc)}% surcharge` : '';
  if (rule.pricingType === 'Run') return `$${money(rule.rate)} flat${scText}`;
  if (rule.pricingType !== 'Hauling') return `$${money(rule.rate)} x ${compact(tons)} tons${scText}`;
  const allowed = rule.allowedTons || 0;
  const overageTons = Math.max(tons - allowed, 0);
  const baseAmount = rule.baseRate || 0;
  const overageRate = rule.overageRate || 0;
  const distanceAmount = (rule.miles || 0) * (rule.mileRate || 0);
  const surcharge = rule.surchargePercent || 0;
  return `base $${money(baseAmount)} + overage ${compact(overageTons)} tons x $${money(overageRate)} + distance $${money(distanceAmount)}${surcharge > 0 ? ` + ${compact(surcharge)}% surcharge` : ''}`;
}

function compact(value) {
  return Number(value).toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function money(value) {
  return Number(value || 0).toFixed(2);
}

function loadPricingRules(db) {
  return db.prepare(`
    SELECT p.id AS id,
           p.customer AS customer,
           p.class AS class,
           p.product_service AS product,
           p.pricing_type AS pricingType,
           p.rate AS rate,
           p.base_rate AS baseRate,
           p.allowed_tons AS allowedTons,
           p.overage_rate AS overageRate,
           p.miles AS miles,
           p.mile_rate AS mileRate,
           p.surcharge_percent AS surchargePercent,
           p.account_type AS accountType,
           p.payment_terms AS paymentTerms,
           p.effective_date AS effectiveDate,
           p.end_date AS endDate,
           p.notes AS notes,
           p.project_address_id AS projectAddressId,
           pa.label AS projectLabel,
           pa.address AS projectAddress,
           pa.city AS projectCity
    FROM pricing p
    LEFT JOIN project_addresses pa ON pa.id = p.project_address_id
    ORDER BY p.customer COLLATE NOCASE, p.product_service COLLATE NOCASE, p.effective_date DESC;
  `).all();
}

function upsertPricingRule(db, id, body) {
  const accountType = normalizeAccountType(body.accountType);
  let paymentTerms = cleanText(body.paymentTerms) || cleanText(body.terms);
  if (accountType === 'AR') {
    if (!paymentTerms) paymentTerms = 'Net 30';
    else {
      const parsed = parsePaymentTerms(paymentTerms);
      paymentTerms = parsed ? parsed.label : paymentTerms;
    }
  } else {
    paymentTerms = null;
  }

  let projectAddressId = cleanNumber(body.projectAddressId);
  if (projectAddressId != null && projectAddressId > 0) {
    projectAddressId = Math.trunc(projectAddressId);
    const pa = findProjectAddressById(db, projectAddressId);
    if (!pa) throw new Error(`Project address id ${projectAddressId} not found.`);
  } else {
    projectAddressId = null;
  }

  const row = {
    id,
    customer: cleanText(body.customer),
    className: cleanText(body.className),
    product: cleanText(body.productService),
    pricingType: normalizePricingType(body.pricingType),
    rate: cleanNumber(body.rate) || 0,
    baseRate: cleanNumber(body.baseRate),
    allowedTons: cleanNumber(body.allowedTons),
    overageRate: cleanNumber(body.overageRate),
    miles: cleanNumber(body.miles),
    mileRate: cleanNumber(body.mileRate),
    surchargePercent: cleanNumber(body.surchargePercent),
    accountType,
    paymentTerms,
    projectAddressId,
    effectiveDate: parseDate(body.effectiveDate)?.iso || null,
    endDate: parseDate(body.endDate)?.iso || null,
    notes: cleanText(body.notes)
  };

  if (!row.product) throw new Error('productService is required');
  if (row.id) {
    db.prepare(`
      UPDATE pricing
      SET customer = ?, class = ?, product_service = ?, rate = ?, pricing_type = ?,
          base_rate = ?, allowed_tons = ?, overage_rate = ?, miles = ?, mile_rate = ?,
          surcharge_percent = ?, account_type = ?, payment_terms = ?,
          project_address_id = ?,
          effective_date = ?, end_date = ?,
          notes = ?, updated_at_utc = CURRENT_TIMESTAMP
      WHERE id = ?;
    `).run(
      row.customer, row.className, row.product, row.rate, row.pricingType,
      row.baseRate, row.allowedTons, row.overageRate, row.miles, row.mileRate,
      row.surchargePercent, row.accountType, row.paymentTerms,
      row.projectAddressId,
      row.effectiveDate, row.endDate,
      row.notes, row.id
    );
    return row.id;
  }

  const result = db.prepare(`
    INSERT INTO pricing (
      customer, class, product_service, rate, pricing_type, base_rate, allowed_tons,
      overage_rate, miles, mile_rate, surcharge_percent, account_type, payment_terms,
      project_address_id, effective_date, end_date, notes
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
  `).run(
    row.customer, row.className, row.product, row.rate, row.pricingType,
    row.baseRate, row.allowedTons, row.overageRate, row.miles, row.mileRate,
    row.surchargePercent, row.accountType, row.paymentTerms,
    row.projectAddressId, row.effectiveDate, row.endDate, row.notes
  );
  return result.lastInsertRowid;
}

function deletePricingRule(db, id) {
  return db.prepare('DELETE FROM pricing WHERE id = ?;').run(Number(id)).changes;
}

function findExistingPricingId(db, customer, className, product, pricingType, effectiveDate) {
  const row = db.prepare(`
    SELECT id
    FROM pricing
    WHERE product_service = ?
      AND pricing_type = ?
      AND ((? IS NULL AND customer IS NULL) OR customer = ?)
      AND ((? IS NULL AND class IS NULL) OR class = ?)
      AND ((? IS NULL AND effective_date IS NULL) OR effective_date = ?)
    ORDER BY id
    LIMIT 1;
  `).get(product, pricingType, customer, customer, className, className, effectiveDate, effectiveDate);
  return row?.id || null;
}

function importPricingCsv(db, content) {
  const rows = parseCsv(content);
  if (rows.length === 0) return { imported: 0, skipped: 0, errors: [] };
  const headers = rows[0].map(normalizeHeader);
  const index = new Map();
  headers.forEach((header, i) => {
    if (!index.has(header)) index.set(header, i);
  });

  let imported = 0;
  let skipped = 0;
  const errors = [];
  db.exec('BEGIN TRANSACTION;');
  try {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row.some(cell => cleanText(cell))) continue;
      const rowNumber = i + 1;
      const product = getAny(row, index, 'Product/Service', 'Product Service', 'Product', 'Service');
      if (!product) {
        skipped++;
        errors.push(`Row ${rowNumber}: missing Product/Service.`);
        continue;
      }
      const pricingType = normalizePricingType(getAny(row, index, 'Pricing Type', 'PricingType', 'Type'));
      const rate = cleanNumber(getAny(row, index, 'Rate', 'Flat Rate', 'RATE'));
      const overageRate = cleanNumber(getAny(row, index, 'Overage Rate', 'OverageRate'));
      if ((pricingType === 'Flat' || pricingType === 'Run') && rate === null) {
        skipped++;
        errors.push(`Row ${rowNumber}: ${pricingType} pricing requires Rate.`);
        continue;
      }
      if (pricingType === 'Hauling' && overageRate === null) {
        skipped++;
        errors.push(`Row ${rowNumber}: Hauling pricing requires Overage Rate.`);
        continue;
      }

      const effectiveDate = parseDate(getAny(row, index, 'Effective Date', 'EffectiveDate', 'Effective', 'Start', 'Start Date'))?.iso || null;
      const customer = cleanText(getAny(row, index, 'Customer', 'Account'));
      const className = cleanText(getAny(row, index, 'Class'));
      const id = findExistingPricingId(db, customer, className, product, pricingType, effectiveDate);
      upsertPricingRule(db, id, {
        customer,
        className,
        productService: product,
        pricingType,
        rate: pricingType === 'Hauling' ? 0 : rate,
        baseRate: cleanNumber(getAny(row, index, 'Base Rate', 'BaseRate', 'Base')),
        allowedTons: cleanNumber(getAny(row, index, 'Allowed Tons', 'AllowedTons')),
        overageRate,
        miles: cleanNumber(getAny(row, index, 'Miles')),
        mileRate: cleanNumber(getAny(row, index, 'Mile Rate', 'MileRate')),
        surchargePercent: cleanNumber(getAny(row, index, 'Surcharge %', 'Surcharge Percent', 'SurchargePercent')),
        accountType: getAny(row, index, 'Account Type', 'AccountType') || 'Regular',
        effectiveDate,
        endDate: parseDate(getAny(row, index, 'End Date', 'EndDate', 'End'))?.iso || null,
        notes: getAny(row, index, 'Notes', 'Memo')
      });
      imported++;
    }
    db.exec('COMMIT;');
  } catch (err) {
    db.exec('ROLLBACK;');
    throw err;
  }
  return { imported, skipped, errors };
}

function normalizeHeader(value) {
  return String(value || '').trim().replace('\ufeff', '').toUpperCase();
}

function getAny(row, index, ...headers) {
  for (const header of headers) {
    const i = index.get(normalizeHeader(header));
    const value = i !== undefined && i < row.length ? cleanText(row[i]) : null;
    if (value) return value;
  }
  return null;
}

function parseCsv(content) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  const textValue = String(content || '').replace(/^\ufeff/, '');
  for (let i = 0; i < textValue.length; i++) {
    const c = textValue[i];
    if (c === '"') {
      if (inQuotes && textValue[i + 1] === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      row.push(cell);
      cell = '';
    } else if ((c === '\n' || c === '\r') && !inQuotes) {
      if (c === '\r' && textValue[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += c;
    }
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function loadTrucks(db) {
  return db.prepare(`
    SELECT id AS id,
           truck_id AS truckId,
           tare_weight AS tareWeight,
           company AS company,
           notes AS notes
    FROM trucks
    ORDER BY truck_id COLLATE NOCASE;
  `).all();
}

function upsertTruck(db, id, body) {
  const truckId = cleanText(body.truckId);
  if (!truckId) throw new Error('Truck ID is required.');
  const tareWeight = cleanNumber(body.tareWeight);
  const company = cleanText(body.company);
  const notes = cleanText(body.notes);
  if (id) {
    db.prepare(`
      UPDATE trucks
      SET truck_id = ?, tare_weight = ?, company = ?, notes = ?, updated_at_utc = CURRENT_TIMESTAMP
      WHERE id = ?;
    `).run(truckId, tareWeight, company, notes, Number(id));
    return Number(id);
  }
  db.prepare(`
    INSERT INTO trucks (truck_id, tare_weight, company, notes)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(truck_id) DO UPDATE SET
      tare_weight = excluded.tare_weight,
      company = excluded.company,
      notes = excluded.notes,
      updated_at_utc = CURRENT_TIMESTAMP;
  `).run(truckId, tareWeight, company, notes);
  return db.prepare('SELECT id FROM trucks WHERE truck_id = ? COLLATE NOCASE;').get(truckId).id;
}

function deleteTruck(db, id) {
  return db.prepare('DELETE FROM trucks WHERE id = ?;').run(Number(id)).changes;
}

function findTruckTare(db, truckId) {
  const value = cleanText(truckId);
  if (!value) return null;
  const row = db.prepare(`
    SELECT tare_weight AS tareWeight
    FROM trucks
    WHERE TRIM(truck_id) = TRIM(?) COLLATE NOCASE
    LIMIT 1;
  `).get(value);
  return cleanNumber(row?.tareWeight);
}

// ---------- Companies (haulers) — editable AR terms / default payment ----------
function loadCompanies(db) {
  return db.prepare(`
    SELECT id AS id,
           name AS name,
           account_type AS accountType,
           payment_terms AS paymentTerms,
           notes AS notes,
           updated_at_utc AS updatedAt
    FROM companies
    ORDER BY name COLLATE NOCASE;
  `).all();
}

function findCompanyByName(db, name) {
  const value = cleanText(name);
  if (!value) return null;
  return db.prepare(`
    SELECT id AS id,
           name AS name,
           account_type AS accountType,
           payment_terms AS paymentTerms,
           notes AS notes
    FROM companies
    WHERE TRIM(name) = TRIM(?) COLLATE NOCASE
    LIMIT 1;
  `).get(value) || null;
}

function upsertCompany(db, id, body) {
  const name = cleanText(body.name) || cleanText(body.customer);
  if (!name) throw new Error('Company name is required.');
  const accountType = normalizeAccountType(body.accountType || body.defaultPayment || 'Regular');
  let paymentTerms = cleanText(body.paymentTerms) || cleanText(body.terms);
  // Only store terms for AR-style companies; clear otherwise.
  if (accountType !== 'AR') {
    paymentTerms = paymentTerms || null;
  } else if (!paymentTerms) {
    paymentTerms = 'Net 30';
  } else {
    const parsed = parsePaymentTerms(paymentTerms);
    paymentTerms = parsed ? parsed.label : paymentTerms;
  }
  const notes = cleanText(body.notes);

  if (id) {
    db.prepare(`
      UPDATE companies
      SET name = ?, account_type = ?, payment_terms = ?, notes = ?,
          updated_at_utc = CURRENT_TIMESTAMP
      WHERE id = ?;
    `).run(name, accountType, paymentTerms, notes, Number(id));
    return Number(id);
  }
  db.prepare(`
    INSERT INTO companies (name, account_type, payment_terms, notes)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(name) DO UPDATE SET
      account_type = excluded.account_type,
      payment_terms = excluded.payment_terms,
      notes = excluded.notes,
      updated_at_utc = CURRENT_TIMESTAMP;
  `).run(name, accountType, paymentTerms, notes);
  const row = db.prepare(
    'SELECT id FROM companies WHERE TRIM(name) = TRIM(?) COLLATE NOCASE;'
  ).get(name);
  return row.id;
}

function deleteCompany(db, id) {
  return db.prepare('DELETE FROM companies WHERE id = ?;').run(Number(id)).changes;
}

// ---------- Project addresses (stapled job sites per customer) ----------
function mapProjectAddressRow(row) {
  if (!row) return null;
  return {
    id: row.id,
    customer: row.customer,
    label: row.label,
    address: row.address,
    city: row.city,
    accountType: row.account_type ?? row.accountType ?? null,
    paymentTerms: row.payment_terms ?? row.paymentTerms ?? null,
    notes: row.notes,
    active: row.active == null ? 1 : Number(row.active),
    displayName: formatProjectDisplay(row)
  };
}

function formatProjectDisplay(row) {
  if (!row) return '';
  const label = cleanText(row.label);
  const address = cleanText(row.address);
  const city = cleanText(row.city);
  const base = label || address || 'Project';
  const tail = [!label && address ? null : address, city].filter(Boolean).join(', ');
  return tail ? `${base} — ${tail}` : base;
}

function loadProjectAddresses(db, customer, activeOnly) {
  const customerValue = cleanText(customer);
  if (customerValue) {
    return db.prepare(`
      SELECT id, customer, label, address, city, account_type, payment_terms, notes, active
      FROM project_addresses
      WHERE TRIM(customer) = TRIM(?) COLLATE NOCASE
        AND (? = 0 OR active = 1)
      ORDER BY COALESCE(label, address) COLLATE NOCASE;
    `).all(customerValue, activeOnly ? 1 : 0).map(mapProjectAddressRow);
  }
  return db.prepare(`
    SELECT id, customer, label, address, city, account_type, payment_terms, notes, active
    FROM project_addresses
    WHERE (? = 0 OR active = 1)
    ORDER BY customer COLLATE NOCASE, COALESCE(label, address) COLLATE NOCASE;
  `).all(activeOnly ? 1 : 0).map(mapProjectAddressRow);
}

function findProjectAddressById(db, id) {
  if (!id) return null;
  const row = db.prepare(`
    SELECT id, customer, label, address, city, account_type, payment_terms, notes, active
    FROM project_addresses
    WHERE id = ?;
  `).get(Number(id));
  return mapProjectAddressRow(row);
}

function upsertProjectAddress(db, id, body) {
  const customer = cleanText(body.customer);
  const address = cleanText(body.address);
  if (!customer) throw new Error('Customer is required for a project address.');
  if (!address) throw new Error('Street address is required.');
  const label = cleanText(body.label);
  const city = cleanText(body.city);
  let accountType = cleanText(body.accountType);
  if (accountType) accountType = normalizeAccountType(accountType);
  else accountType = null;

  let paymentTerms = cleanText(body.paymentTerms) || cleanText(body.terms);
  if (accountType === 'AR' || (!accountType && paymentTerms)) {
    if (accountType === 'AR' && !paymentTerms) paymentTerms = 'Net 30';
    if (paymentTerms) {
      const parsed = parsePaymentTerms(paymentTerms);
      paymentTerms = parsed ? parsed.label : paymentTerms;
    }
  } else if (accountType && accountType !== 'AR') {
    paymentTerms = null;
  }

  const notes = cleanText(body.notes);
  const active = body.active === false || body.active === 0 || body.active === '0' ? 0 : 1;

  // Ensure company row exists so ticket lookup still works (do not overwrite).
  try {
    if (!findCompanyByName(db, customer)) {
      upsertCompany(db, null, {
        name: customer,
        accountType: accountType || 'Regular',
        paymentTerms: accountType === 'AR' ? paymentTerms : null
      });
    }
  } catch { /* non-fatal */ }

  if (id) {
    db.prepare(`
      UPDATE project_addresses
      SET customer = ?, label = ?, address = ?, city = ?,
          account_type = ?, payment_terms = ?, notes = ?, active = ?,
          updated_at_utc = CURRENT_TIMESTAMP
      WHERE id = ?;
    `).run(customer, label, address, city, accountType, paymentTerms, notes, active, Number(id));
    return Number(id);
  }
  const result = db.prepare(`
    INSERT INTO project_addresses (
      customer, label, address, city, account_type, payment_terms, notes, active
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
  `).run(customer, label, address, city, accountType, paymentTerms, notes, active);
  return Number(result.lastInsertRowid);
}

function deleteProjectAddress(db, id) {
  const pid = Number(id);
  // Detach pricing rules that pointed at this project (keep the rates as customer-level).
  try {
    db.prepare('UPDATE pricing SET project_address_id = NULL WHERE project_address_id = ?;').run(pid);
  } catch { /* column may not exist on ancient DBs */ }
  try {
    db.prepare('UPDATE quickbooks_transactions SET project_address_id = NULL WHERE project_address_id = ?;').run(pid);
  } catch { /* ignore */ }
  return db.prepare('DELETE FROM project_addresses WHERE id = ?;').run(pid).changes;
}

function validateWasteData(req) {
  const product = (cleanText(req.product) || '').toUpperCase();
  if ((product === '40YD' || product === '10YD') && !cleanText(req.wasteData)) {
    return `Waste Data is required for ${product} material type.`;
  }
  return null;
}

/**
 * Two-stage weigh:
 *  1) INBOUND only (gross) — save when truck arrives / before dump
 *  2) Later add OUTBOUND (tare) — complete net tons + pricing
 * Outbound is optional until the second weigh.
 */
function validateTicketWeights(req, existing = null) {
  let gross = cleanNumber(req.gross);
  let tare = cleanNumber(req.tare);
  if (existing) {
    if (gross === null) gross = cleanNumber(existing.t1);
    if (tare === null) tare = cleanNumber(existing.t2);
  }
  if (gross === null) {
    return 'Inbound weight (GROSS, lb) is required to save (first weigh).';
  }
  if (gross < 0) return 'Inbound weight cannot be negative.';
  if (tare !== null && tare < 0) return 'Outbound weight cannot be negative.';
  if (tare !== null && gross === tare) {
    return 'Inbound and outbound weights cannot be the same (net would be zero).';
  }
  return null;
}

/**
 * Build a ticket row for insert/update.
 * options:
 *   mode: 'create' | 'update'
 *   existing: row from DB when updating (preserves numbers if client omits them)
 *   numbers: { weighmasterNumber, invoiceNumber } forced allocation (create)
 */
function buildEntry(db, req, options = {}) {
  const date = parseDate(req.date);
  const mode = options.mode || 'create';
  const existing = options.existing || null;
  const forced = options.numbers || null;

  let weighmasterNumber;
  let invoiceNumber;
  if (forced) {
    weighmasterNumber = String(forced.weighmasterNumber);
    invoiceNumber = String(forced.invoiceNumber);
  } else if (mode === 'update' && existing) {
    // Never auto-mint new numbers on edit — that caused gaps/duplicates.
    weighmasterNumber = cleanText(req.weighmasterNumber)
      || cleanText(existing.weighmaster_number)
      || cleanText(existing.weighmaster)
      || nextWeighmasterNumber(db);
    invoiceNumber = cleanText(req.invoiceNumber)
      || cleanText(existing.invoice_number)
      || cleanText(existing.document_number)
      || nextInvoiceNumber(db);
  } else {
    // Preview / legacy path — prefer client values, else next.
    weighmasterNumber = cleanText(req.weighmasterNumber) || nextWeighmasterNumber(db);
    invoiceNumber = cleanText(req.invoiceNumber) || nextInvoiceNumber(db);
  }

  // Weights: prefer request body; on update never wipe existing t1/t2 with null.
  let gross = cleanNumber(req.gross);
  let tare = cleanNumber(req.tare);
  if (mode === 'update' && existing) {
    if (gross === null) gross = cleanNumber(existing.t1);
    if (tare === null) tare = cleanNumber(existing.t2);
  }

  let tonage = null;
  if (gross !== null && tare !== null) {
    tonage = Math.abs(gross - tare) / 2000.0;
    // Legacy scale rule: bill at least 1 ton when there is any net weight.
    if (tonage > 0 && tonage < 1) tonage = 1;
    tonage = Math.round(tonage * 100) / 100;
  }

  const product = cleanText(req.product)
    || (mode === 'update' && existing ? cleanText(existing.product_service) : null);

  // Stapled project: request id wins; else keep existing on update.
  let projectAddressId = cleanNumber(req.projectAddressId);
  if (projectAddressId != null && projectAddressId > 0) {
    projectAddressId = Math.trunc(projectAddressId);
  } else if (mode === 'update' && existing && existing.project_address_id != null) {
    // Only preserve if client omitted the field (null/undefined), not cleared.
    if (req.projectAddressId === undefined) {
      projectAddressId = Number(existing.project_address_id) || null;
    } else {
      projectAddressId = null;
    }
  } else {
    projectAddressId = null;
  }
  const project = projectAddressId ? findProjectAddressById(db, projectAddressId) : null;
  if (projectAddressId && !project) projectAddressId = null;

  const rule = findPricingRule(db, req.customer, product, date, null, projectAddressId);
  const rawPayment = (cleanText(req.payment) || '').toUpperCase();
  let payment = rawPayment
    || paymentMethodFor(project?.accountType)
    || paymentMethodFor(rule?.accountType)
    || 'CASH';
  let amount = cleanNumber(req.amount);
  // On update, keep previous amount if client sent blank (avoid wiping EDCO $).
  if (amount === null && mode === 'update' && existing) {
    amount = cleanNumber(existing.amount);
  }
  let amountAutoComputed = false;
  if (amount === null && rule && tonage !== null && tonage > 0) {
    amount = calculateAmount(rule, tonage);
    amountAutoComputed = true;
  }
  if (amountAutoComputed && amount !== null && payment === 'CREDIT') amount *= 1.03;
  if (amount !== null && amount < 0) payment = 'EXPENSE';

  let rateResolved = null;
  if (/^weight$/i.test(product || '')) {
    rateResolved = amount;
  } else if (amount !== null && tonage !== null && tonage > 0) {
    rateResolved = Math.round((amount / tonage) * 100) / 100;
  }

  // Times: preserve on update if omitted
  let timeIn = cleanText(req.timeIn);
  let timeOut = cleanText(req.timeOut);
  if (mode === 'update' && existing) {
    if (!timeIn) timeIn = cleanText(existing.time_in);
    if (!timeOut) timeOut = cleanText(existing.time_out);
  }

  // Origin / city: request wins; project stapled address fills gaps.
  let origin = cleanText(req.origin);
  let city = cleanText(req.city);
  if (project) {
    if (!origin) origin = cleanText(project.address);
    if (!city) city = cleanText(project.city);
  }
  if (mode === 'update' && existing) {
    if (!origin) origin = cleanText(existing.origin);
    if (!city) city = cleanText(existing.city);
  }

  // Document type follows payment (Invoice / Sales Receipt / Expense).
  const series = documentSeriesForPayment(payment);
  let documentType = series.type;
  if (mode === 'update' && existing && cleanText(existing.document_type)) {
    // Keep type unless payment changed to a different series.
    const prevSeries = documentSeriesForPayment(existing.payment);
    if (prevSeries.type === series.type) {
      documentType = cleanText(existing.document_type) || series.type;
    }
  }
  if (forced && forced.documentType) documentType = forced.documentType;

  const payNorm = normalizePaymentMethod(payment);
  // AR terms (Net 10 / Net 30) → due date; other payments clear terms/due.
  // Priority: explicit request → project terms → pricing rule → company default → Net 30.
  let paymentTerms = null;
  let dueDate = null;
  if (payNorm === 'AR') {
    let termsRaw = cleanText(req.paymentTerms) || cleanText(req.terms);
    if (!termsRaw && project?.paymentTerms) termsRaw = project.paymentTerms;
    if (!termsRaw && rule?.paymentTerms) termsRaw = rule.paymentTerms;
    if (!termsRaw) {
      const co = findCompanyByName(db, req.customer);
      if (co?.paymentTerms) termsRaw = co.paymentTerms;
    }
    if (!termsRaw && mode === 'update' && existing) {
      termsRaw = cleanText(existing.payment_terms);
    }
    if (!termsRaw) termsRaw = 'Net 30'; // default AR terms
    const parsed = parsePaymentTerms(termsRaw);
    paymentTerms = parsed ? parsed.label : termsRaw;
    const days = parsed && parsed.days != null ? parsed.days : 30;
    // Explicit dueDate from client wins; else compute from ticket date.
    dueDate = cleanText(req.dueDate)
      || computeDueDateIso(date?.iso, days);
    if (!dueDate && mode === 'update' && existing) {
      dueDate = cleanText(existing.due_date);
    }
  }

  return {
    documentType,
    documentNumber: invoiceNumber,
    transactionDate: date?.iso || null,
    customer: cleanText(req.customer),
    weighmaster: cleanText(req.weighmaster) || cleanText(req.byPerson),
    productService: product,
    qty: tonage,
    rate: rateResolved,
    amount,
    truckId: cleanText(req.truck)
      || (mode === 'update' && existing ? cleanText(existing.truck_id) : null),
    poOrigin: null,
    memo: cleanText(req.memo),
    payment: payNorm,
    ar: payNorm === 'AR' ? amount : null,
    className: null,
    scalemaster: null,
    invoiceDate: null,
    description: null,
    depositTo: null,
    location: null,
    roundAmount: null,
    dateCopy: date?.iso || null,
    wasteData: cleanText(req.wasteData),
    timeIn,
    timeOut,
    t1: gross,
    t2: tare,
    tonage,
    origin,
    cash: payNorm === 'CASH' ? amount : null,
    card: payNorm === 'CREDIT' || payNorm === 'DEBIT' ? amount : null,
    weighmasterNumber,
    invoiceNumber,
    city,
    ticketNotes: cleanText(req.notes),
    weightCertificate: req.weightCertificate ? 1 : 0,
    recyclingCdCertificate: req.recyclingCdCertificate ? 1 : 0,
    byPerson: cleanText(req.byPerson),
    paymentTerms,
    dueDate,
    projectAddressId,
    // undefined = client did not send a signature field (keep existing on UPDATE).
    // null / empty = explicit clear. Buffer = replace.
    driverSignature: Object.prototype.hasOwnProperty.call(req, 'driverSignaturePngBase64')
      ? decodeSignature(req.driverSignaturePngBase64, false)
      : undefined,
    rateResolved
  };
}

function decodeSignature(value, throwOnInvalid) {
  const textValue = cleanText(value);
  if (!textValue) return null;
  let raw = textValue;
  const comma = raw.indexOf(',');
  if (comma >= 0 && /^data:/i.test(raw)) raw = raw.slice(comma + 1);
  if (raw.length > (MAX_SIGNATURE_BYTES * 4 / 3) + 16) {
    if (throwOnInvalid) throw new Error('signature exceeds maximum size');
    return null;
  }
  if (!/^[a-z0-9+/=\s]+$/i.test(raw)) {
    if (throwOnInvalid) throw new Error('pngBase64 is not valid base64');
    return null;
  }
  const bytes = Buffer.from(raw, 'base64');
  if (bytes.length === 0 || bytes.length > MAX_SIGNATURE_BYTES) {
    if (throwOnInvalid) throw new Error(bytes.length === 0 ? 'signature is empty' : 'signature exceeds maximum size');
    return null;
  }
  return bytes;
}

function insertEntry(db, entry) {
  const sourceFile = `manual:${os.hostname()}`;
  const sourceRowNumber = nextManualRowNumber(db, sourceFile);
  const result = db.prepare(`
    INSERT INTO quickbooks_transactions (
      source_file, source_row_number,
      document_type, document_number, transaction_date, customer,
      weighmaster, product_service, qty, rate, amount,
      truck_id, po_origin, memo, payment, ar,
      week, biweekly, class, scalemaster, month,
      invoice_date, description, deposit_to, location, round_amount,
      date_copy, waste_data,
      time_in, time_out, t1, t2, tonage, origin, cash, card,
      weighmaster_number, invoice_number, city, ticket_notes,
      weight_certificate, recycling_cd_certificate, by_person,
      payment_terms, due_date, project_address_id,
      driver_signature
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
  `).run(...entryParams(entry, sourceFile, sourceRowNumber));
  return result.lastInsertRowid;
}

function updateEntry(db, id, entry) {
  // Preserve existing BLOB unless the client explicitly sent a signature field.
  // Previously every Save without a new pad stroke set driver_signature = NULL.
  const touchSignature = entry.driverSignature !== undefined;
  if (touchSignature) {
    return db.prepare(`
      UPDATE quickbooks_transactions SET
        document_type = ?,
        document_number = ?,
        transaction_date = ?,
        customer = ?,
        weighmaster = ?,
        product_service = ?,
        qty = ?,
        rate = ?,
        amount = ?,
        truck_id = ?,
        po_origin = ?,
        memo = ?,
        payment = ?,
        ar = ?,
        week = ?,
        biweekly = ?,
        class = ?,
        scalemaster = ?,
        month = ?,
        invoice_date = ?,
        description = ?,
        deposit_to = ?,
        location = ?,
        round_amount = ?,
        date_copy = ?,
        waste_data = ?,
        time_in = ?,
        time_out = ?,
        t1 = ?,
        t2 = ?,
        tonage = ?,
        origin = ?,
        cash = ?,
        card = ?,
        weighmaster_number = ?,
        invoice_number = ?,
        city = ?,
        ticket_notes = ?,
        weight_certificate = ?,
        recycling_cd_certificate = ?,
        by_person = ?,
        payment_terms = ?,
        due_date = ?,
        project_address_id = ?,
        driver_signature = ?
      WHERE id = ?;
    `).run(...entryParams(entry), Number(id)).changes;
  }
  return db.prepare(`
    UPDATE quickbooks_transactions SET
      document_type = ?,
      document_number = ?,
      transaction_date = ?,
      customer = ?,
      weighmaster = ?,
      product_service = ?,
      qty = ?,
      rate = ?,
      amount = ?,
      truck_id = ?,
      po_origin = ?,
      memo = ?,
      payment = ?,
      ar = ?,
      week = ?,
      biweekly = ?,
      class = ?,
      scalemaster = ?,
      month = ?,
      invoice_date = ?,
      description = ?,
      deposit_to = ?,
      location = ?,
      round_amount = ?,
      date_copy = ?,
      waste_data = ?,
      time_in = ?,
      time_out = ?,
      t1 = ?,
      t2 = ?,
      tonage = ?,
      origin = ?,
      cash = ?,
      card = ?,
      weighmaster_number = ?,
      invoice_number = ?,
      city = ?,
      ticket_notes = ?,
      weight_certificate = ?,
      recycling_cd_certificate = ?,
      by_person = ?,
      payment_terms = ?,
      due_date = ?,
      project_address_id = ?
    WHERE id = ?;
  `).run(...entryParamsWithoutSignature(entry), Number(id)).changes;
}

function entryParamsWithoutSignature(entry) {
  // Same as entryParams but omit trailing driverSignature.
  const all = entryParams(entry);
  return all.slice(0, -1);
}

function entryParams(entry, sourceFile = null, sourceRowNumber = null) {
  const date = parseDate(entry.transactionDate);
  const params = [
    entry.documentType,
    entry.documentNumber,
    entry.transactionDate,
    entry.customer,
    entry.weighmaster,
    entry.productService,
    entry.qty,
    entry.rate,
    entry.amount,
    entry.truckId,
    entry.poOrigin,
    entry.memo,
    entry.payment,
    entry.ar,
    weekOf(date),
    biweeklyOf(date),
    entry.className,
    entry.scalemaster,
    monthOf(date),
    entry.invoiceDate,
    entry.description,
    entry.depositTo,
    entry.location,
    entry.roundAmount,
    entry.dateCopy,
    entry.wasteData,
    entry.timeIn,
    entry.timeOut,
    entry.t1,
    entry.t2,
    entry.tonage,
    entry.origin,
    entry.cash,
    entry.card,
    entry.weighmasterNumber,
    entry.invoiceNumber,
    entry.city,
    entry.ticketNotes,
    entry.weightCertificate,
    entry.recyclingCdCertificate,
    entry.byPerson,
    entry.paymentTerms ?? null,
    entry.dueDate ?? null,
    entry.projectAddressId ?? null,
    // node:sqlite rejects undefined — use null when signature not sent.
    entry.driverSignature === undefined ? null : entry.driverSignature
  ];
  return sourceFile === null ? params : [sourceFile, sourceRowNumber, ...params];
}

function loadQuickBooksExtract(db, limit = 2147483647) {
  return db.prepare(`
    SELECT CAST(COALESCE(invoice_number, document_number) AS TEXT) AS "Invoice Number",
           CAST(waste_data AS TEXT) AS "WASTE DATA",
           CAST(time_in AS TEXT) AS "TIME IN",
           CAST(time_out AS TEXT) AS "TIME OUT",
           CAST(memo AS TEXT) AS "Memo",
           t1 AS "T1",
           t2 AS "T2",
           CAST(truck_id AS TEXT) AS "TRUCKID",
           CAST(po_origin AS TEXT) AS "PO# / ORIGIN",
           tonage AS "TONAGE",
           CAST(transaction_date AS TEXT) AS "DATE",
           CAST(customer AS TEXT) AS "Account",
           CAST(COALESCE(NULLIF(TRIM(weighmaster_number), ''), weighmaster) AS TEXT) AS "Weighmaster",
           CAST(product_service AS TEXT) AS "Product/Service",
           CAST(origin AS TEXT) AS "Origin",
           qty AS "Quantity",
           rate AS "RATE",
           amount AS "AMOUNT",
           CAST(payment AS TEXT) AS "PAYMENT",
           cash AS "CASH",
           ar AS "AR",
           card AS "CARD",
           CAST(week AS TEXT) AS "Week",
           CAST(biweekly AS TEXT) AS "Biweekly",
           CAST(class AS TEXT) AS "Class",
           CAST(month AS TEXT) AS "Month"
    FROM quickbooks_transactions
    ORDER BY transaction_date DESC, id DESC
    LIMIT ?;
  `).all(limit);
}

function loadArBillingTickets(db, from, to) {
  return db.prepare(`
    SELECT id AS id,
           CAST(customer AS TEXT) AS customer,
           CAST(transaction_date AS TEXT) AS date,
           CAST(COALESCE(NULLIF(TRIM(weighmaster_number), ''), weighmaster) AS TEXT) AS ticketNo,
           CAST(COALESCE(invoice_number, document_number) AS TEXT) AS invoiceNo,
           CAST(truck_id AS TEXT) AS truck,
           CAST(product_service AS TEXT) AS product,
           COALESCE(qty, tonage) AS tons,
           rate AS rate,
           amount AS amount,
           COALESCE(ar, CASE WHEN UPPER(COALESCE(payment, '')) = 'AR' THEN amount END, amount) AS arAmount,
           CAST(payment AS TEXT) AS payment,
           CAST(city AS TEXT) AS city,
           CAST(origin AS TEXT) AS origin,
           CAST(po_origin AS TEXT) AS wasteOrigin,
           CAST(memo AS TEXT) AS memo
    FROM quickbooks_transactions
    WHERE transaction_date >= ?
      AND transaction_date <= ?
      AND customer IS NOT NULL
      AND TRIM(customer) <> ''
      AND (COALESCE(ar, 0) <> 0 OR UPPER(COALESCE(payment, '')) = 'AR')
    ORDER BY customer COLLATE NOCASE, transaction_date, id;
  `).all(from, to);
}

function loadArBillingCustomerTotals(db, from, to) {
  return db.prepare(`
    SELECT CAST(customer AS TEXT) AS customer,
           COUNT(*) AS ticketCount,
           SUM(COALESCE(qty, tonage, 0)) AS totalTons,
           SUM(COALESCE(ar, CASE WHEN UPPER(COALESCE(payment, '')) = 'AR' THEN amount END, amount, 0)) AS totalAmount,
           CAST(MIN(transaction_date) AS TEXT) AS firstDate,
           CAST(MAX(transaction_date) AS TEXT) AS lastDate
    FROM quickbooks_transactions
    WHERE transaction_date >= ?
      AND transaction_date <= ?
      AND customer IS NOT NULL
      AND TRIM(customer) <> ''
      AND (COALESCE(ar, 0) <> 0 OR UPPER(COALESCE(payment, '')) = 'AR')
    GROUP BY customer
    ORDER BY customer COLLATE NOCASE;
  `).all(from, to);
}

function parseBillingRange(from, to) {
  const today = parseDate(todayIso());
  let fromDate = { year: today.year, month: today.month, day: 1, iso: `${today.year}-${String(today.month).padStart(2, '0')}-01` };
  let toDate = today;
  if (cleanText(from)) {
    fromDate = parseDate(from);
    if (!fromDate) return { error: 'from is not a valid date (expected yyyy-MM-dd)' };
  }
  if (cleanText(to)) {
    toDate = parseDate(to);
    if (!toDate) return { error: 'to is not a valid date (expected yyyy-MM-dd)' };
  }
  if (dateToTime(fromDate) > dateToTime(toDate)) {
    return { error: 'from must be on or before to' };
  }
  return { from: fromDate.iso, to: toDate.iso };
}

const POWER_BI_COLUMNS = [
  ['id', 'ID'],
  ['ticketNo', 'Ticket #'],
  ['invoiceNo', 'Invoice #'],
  ['date', 'Date'],
  ['customer', 'Account'],
  ['truck', 'Truck'],
  ['product', 'Product / Service'],
  ['tons', 'Tons'],
  ['rate', 'Rate'],
  ['amount', 'Amount $'],
  ['cash', 'Cash $'],
  ['card', 'Card $'],
  ['ar', 'AR $'],
  ['payment', 'Paid'],
  ['timeIn', 'Time In'],
  ['timeOut', 'Time Out'],
  ['gross', 'Gross (lb)'],
  ['tare', 'Tare (lb)'],
  ['origin', 'Address (Origin)'],
  ['wasteOrigin', 'PO / Origin'],
  ['city', 'City'],
  ['wasteData', 'Waste Data'],
  ['memo', 'Memo'],
  ['ticketNotes', 'Notes'],
  ['weighmaster', 'Weighmaster'],
  ['byPerson', 'By'],
  ['weightCertificate', 'Weight Cert.'],
  ['recyclingCdCertificate', 'Recycling C&D']
];

function renderCsv(rows, columns = null) {
  const cols = columns || inferColumns(rows);
  const header = cols.map(c => csvEscape(Array.isArray(c) ? c[1] : c)).join(',');
  const lines = [header];
  for (const row of rows) {
    lines.push(cols.map(c => csvEscape(formatCsvValue(row[Array.isArray(c) ? c[0] : c]))).join(','));
  }
  return lines.join('\r\n') + '\r\n';
}

function inferColumns(rows) {
  return rows.length ? Object.keys(rows[0]) : [];
}

function formatCsvValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return Number.isInteger(value) ? String(value) : compact(value);
  if (Buffer.isBuffer(value)) return value.toString('utf8').replace(/\0+$/, '');
  return String(value);
}

function csvEscape(value) {
  const textValue = String(value ?? '');
  return /[",\r\n]/.test(textValue) ? `"${textValue.replace(/"/g, '""')}"` : textValue;
}

function parseMultipartFile(req, body) {
  const contentType = req.headers['content-type'] || '';
  const match = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/i);
  if (!match) throw new Error("expected multipart/form-data with a 'file' field");
  const boundary = `--${match[1] || match[2]}`;
  const raw = body.toString('latin1');
  const parts = raw.split(boundary).slice(1, -1);
  for (const part of parts) {
    const clean = part.replace(/^\r\n/, '');
    const splitAt = clean.indexOf('\r\n\r\n');
    if (splitAt < 0) continue;
    const headers = clean.slice(0, splitAt);
    let content = clean.slice(splitAt + 4);
    if (content.endsWith('\r\n')) content = content.slice(0, -2);
    if (/name="file"|filename="/i.test(headers)) {
      return Buffer.from(content, 'latin1');
    }
  }
  throw new Error('no file uploaded');
}

// ---------- settings (key/value in SQLite) ----------

function getSetting(db, key) {
  const row = db.prepare('SELECT value FROM app_settings WHERE key = ?;').get(key);
  return row ? row.value : null;
}

function setSetting(db, key, value) {
  db.prepare('INSERT INTO app_settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value;')
    .run(key, value);
}

function getEmailConfig(db) {
  const raw = getSetting(db, 'email_config');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch { return {}; }
}

// The actual SMTP password used at send time. An env var wins over the stored
// value so the secret can stay off disk on the production box if preferred.
function effectiveEmailPassword(cfg) {
  return process.env.AIS_SMTP_PASSWORD || cfg.password || '';
}

function validateEmailConfig(cfg) {
  if (!cleanText(cfg.host)) return 'SMTP host is required';
  if (!cfg.port) return 'SMTP port is required';
  if (!cleanText(cfg.from)) return 'From address is required';
  if (splitRecipients(cfg.to).length === 0) return 'At least one recipient (To) is required';
  if (cleanText(cfg.username) && !effectiveEmailPassword(cfg)) return 'Password is required when a username is set';
  return null;
}

function splitRecipients(value) {
  return String(value || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
}

// Tickets for a single day (yyyy-MM-dd), newest first. Defaults to today.
function dayTickets(db, date) {
  const target = cleanText(date) || todayIso();
  const rows = loadWeighmasterTickets(db, 2147483647).filter(r => String(r.date || '').slice(0, 10) === target);
  return { date: target, rows };
}

function isoOf(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function yesterdayIso() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return isoOf(d);
}

function parseHhMm(value) {
  const m = String(value || '').match(/^(\d{1,2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : 18 * 60;
}

// Builds and sends one day's CSV email. Shared by the manual endpoint and the
// scheduler. Returns { count, date }. Throws on SMTP/validation failure.
async function sendDayEmail(db, cfg, dataDate) {
  const { date, rows } = dayTickets(db, dataDate);
  const csv = renderCsv(rows, POWER_BI_COLUMNS);
  const totalAmount = rows.reduce((s, r) => s + (cleanNumber(r.amount) || 0), 0);
  const totalTons = rows.reduce((s, r) => s + (cleanNumber(r.tons) || 0), 0);
  await smtpSend({
    host: cfg.host, port: cfg.port, secure: cfg.secure,
    user: cfg.username, pass: effectiveEmailPassword(cfg),
    from: cfg.from, to: splitRecipients(cfg.to),
    subject: `${cfg.subjectPrefix || 'AIS Daily Tickets'} — ${date} (${rows.length} ticket${rows.length === 1 ? '' : 's'})`,
    text: `Attached: ${rows.length} ticket(s) for ${date}.\n\n` +
          `Total amount: $${money(totalAmount)}\n` +
          `Total tons: ${compact(totalTons)}\n\n` +
          `Generated by the AIS LAN dashboard.`,
    attachments: [{
      filename: `tickets_${date.replace(/-/g, '')}.csv`,
      content: Buffer.from(csv, 'utf8'),
      contentType: 'text/csv; charset=utf-8'
    }],
    rejectUnauthorized: !cfg.tlsInsecure
  });
  return { count: rows.length, date };
}

// In-process daily scheduler. Runs every 60s; sends once per local day when the
// clock reaches the configured time. Survives restarts (catches up if the server
// was down at the target time) and retries transient failures up to a daily cap
// without re-sending on a day that already succeeded.
const AUTO_MAX_ATTEMPTS = 10;

async function autoSendTick() {
  let db;
  try {
    db = getDb();
  } catch {
    return; // DB not ready yet
  }
  const cfg = getEmailConfig(db);
  if (!cfg.autoSendEnabled) return;

  const now = new Date();
  const runDate = todayIso();                       // one send per local calendar day
  if (getSetting(db, 'auto_last_sent') === runDate) return;
  if (minutesOfDayNow(now) < parseHhMm(cfg.autoSendTime)) return; // not time yet

  // Daily attempt cap so a hard misconfig doesn't hammer the SMTP server all day.
  const attemptsKey = getSetting(db, 'auto_attempts_date') === runDate
    ? Number(getSetting(db, 'auto_attempts_count') || 0)
    : 0;
  if (attemptsKey >= AUTO_MAX_ATTEMPTS) {
    setSetting(db, 'auto_last_sent', runDate); // give up for today
    return;
  }
  setSetting(db, 'auto_attempts_date', runDate);
  setSetting(db, 'auto_attempts_count', String(attemptsKey + 1));

  const dataDate = cfg.autoSendWhich === 'yesterday' ? yesterdayIso() : todayIso();

  // Config errors are permanent for the day — record and stop retrying.
  const cfgErr = validateEmailConfig(cfg);
  if (cfgErr) {
    setSetting(db, 'auto_last_sent', runDate);
    setSetting(db, 'auto_last_result', `ERROR (config): ${cfgErr}`);
    console.error(`[auto-email] config error: ${cfgErr}`);
    return;
  }

  if (cfg.skipEmptyDays && dayTickets(db, dataDate).rows.length === 0) {
    setSetting(db, 'auto_last_sent', runDate);
    setSetting(db, 'auto_last_result', `Skipped ${dataDate}: no tickets`);
    console.log(`[auto-email] skipped ${dataDate} (no tickets)`);
    return;
  }

  try {
    const { count } = await sendDayEmail(db, cfg, dataDate);
    setSetting(db, 'auto_last_sent', runDate);     // success: no more sends today
    setSetting(db, 'auto_last_result', `OK: ${count} ticket(s) for ${dataDate} to ${cfg.to} at ${runDate} ${hhmmNow(now)}`);
    console.log(`[auto-email] sent ${count} ticket(s) for ${dataDate} to ${cfg.to}`);
  } catch (err) {
    // Transient (network/auth) failure: do NOT mark sent, so the next tick retries
    // (e.g., when the internet comes back), up to the daily attempt cap.
    setSetting(db, 'auto_last_result', `ERROR (attempt ${attemptsKey + 1}/${AUTO_MAX_ATTEMPTS}): ${err.message}`);
    console.error(`[auto-email] send failed (attempt ${attemptsKey + 1}): ${err.message}`);
  }
}

function minutesOfDayNow(d) { return d.getHours() * 60 + d.getMinutes(); }
function hhmmNow(d) { return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; }

// ---------- zero-dependency SMTP client ----------
// Speaks just enough SMTP to authenticate (AUTH LOGIN) and send one multipart
// message with a CSV attachment. Supports implicit TLS (secure=true, port 465)
// and STARTTLS upgrade (secure=false, port 587). Built on node:net / node:tls so
// the dashboard keeps its zero-dependency, copy-and-run deployment.
function smtpSend(options) {
  return new Promise((resolve, reject) => {
    const timeoutMs = options.timeoutMs || 20000;
    const rejectUnauthorized = options.rejectUnauthorized !== false;
    let socket;
    let buffer = '';
    const waiters = [];
    let settled = false;

    function fail(err) {
      if (settled) return;
      settled = true;
      try { socket && socket.destroy(); } catch {}
      reject(err instanceof Error ? err : new Error(String(err)));
    }

    function onData(chunk) {
      buffer += chunk.toString('utf8');
      // A complete SMTP reply is zero+ continuation lines ("250-...") followed by
      // a final line with a space after the code ("250 ...").
      let m;
      while ((m = buffer.match(/^(?:\d{3}-[^\r\n]*\r\n)*\d{3} [^\r\n]*\r\n/))) {
        const block = m[0];
        buffer = buffer.slice(block.length);
        const code = parseInt(block.slice(0, 3), 10);
        const waiter = waiters.shift();
        if (waiter) {
          if (waiter.codes.includes(code)) waiter.resolve({ code, text: block.trim() });
          else waiter.reject(new Error(`SMTP ${code}: ${block.trim().replace(/\r?\n/g, ' ')}`));
        }
      }
    }

    function expect(codes) {
      return new Promise((res, rej) => waiters.push({ codes, resolve: res, reject: rej }));
    }
    function write(line) { socket.write(line + '\r\n'); }

    function attach(sock) {
      sock.on('data', onData);
      sock.on('error', fail);
      sock.setTimeout(timeoutMs, () => fail(new Error(`SMTP timeout after ${timeoutMs}ms`)));
    }

    async function converse() {
      try {
        await expect([220]);
        write('EHLO ' + (os.hostname() || 'localhost'));
        await expect([250]);

        if (!options.secure) {
          write('STARTTLS');
          await expect([220]);
          const plain = socket;
          plain.removeListener('data', onData);
          socket = tls.connect({ socket: plain, servername: options.host, rejectUnauthorized });
          await new Promise((res, rej) => {
            socket.once('secureConnect', res);
            socket.once('error', rej);
          });
          attach(socket);
          write('EHLO ' + (os.hostname() || 'localhost'));
          await expect([250]);
        }

        if (cleanText(options.user)) {
          write('AUTH LOGIN');
          await expect([334]);
          write(Buffer.from(String(options.user), 'utf8').toString('base64'));
          await expect([334]);
          write(Buffer.from(String(options.pass || ''), 'utf8').toString('base64'));
          await expect([235]);
        }

        write(`MAIL FROM:<${options.from}>`);
        await expect([250]);
        for (const rcpt of options.to) {
          write(`RCPT TO:<${rcpt}>`);
          await expect([250, 251]);
        }
        write('DATA');
        await expect([354]);
        socket.write(dotStuff(buildMime(options)) + '\r\n.\r\n');
        await expect([250]);
        write('QUIT');
        settled = true;
        try { socket.end(); } catch {}
        resolve({ ok: true });
      } catch (err) {
        fail(err);
      }
    }

    try {
      if (options.secure) {
        socket = tls.connect({ host: options.host, port: options.port, servername: options.host, rejectUnauthorized });
      } else {
        socket = net.connect({ host: options.host, port: options.port });
      }
      attach(socket);
      converse();
    } catch (err) {
      fail(err);
    }
  });
}

// Doubles a leading dot on any line so message content can't terminate DATA early.
function dotStuff(message) {
  return message.replace(/\r\n\./g, '\r\n..').replace(/^\./, '..');
}

function wrap76(b64) {
  return b64.replace(/.{1,76}/g, '$&\r\n').replace(/\r\n$/, '');
}

function mimeDate(d = new Date()) {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const mons = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const pad = n => String(n).padStart(2, '0');
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? '+' : '-';
  const abs = Math.abs(off);
  return `${days[d.getDay()]}, ${pad(d.getDate())} ${mons[d.getMonth()]} ${d.getFullYear()} ` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())} ${sign}${pad(Math.floor(abs / 60))}${pad(abs % 60)}`;
}

function buildMime(o) {
  const boundary = 'aisbnd_' + crypto.randomBytes(12).toString('hex');
  const toHeader = Array.isArray(o.to) ? o.to.join(', ') : String(o.to);
  const headers = [
    `From: ${o.from}`,
    `To: ${toHeader}`,
    `Subject: ${o.subject}`,
    `Date: ${mimeDate()}`,
    `Message-ID: <${crypto.randomBytes(12).toString('hex')}@ais.local>`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${boundary}"`
  ].join('\r\n');

  const bodyText = String(o.text || '').replace(/\r?\n/g, '\r\n');
  const parts = [
    `--${boundary}`,
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: 8bit',
    '',
    bodyText,
    ''
  ];
  for (const att of (o.attachments || [])) {
    const b64 = wrap76(Buffer.from(att.content).toString('base64'));
    parts.push(
      `--${boundary}`,
      `Content-Type: ${att.contentType || 'application/octet-stream'}; name="${att.filename}"`,
      'Content-Transfer-Encoding: base64',
      `Content-Disposition: attachment; filename="${att.filename}"`,
      '',
      b64,
      ''
    );
  }
  parts.push(`--${boundary}--`, '');
  return headers + '\r\n\r\n' + parts.join('\r\n');
}

async function handleApi(req, res, url) {
  const db = getDb();
  const pathname = url.pathname;
  const query = url.searchParams;

  // ----- auth endpoints (public where noted) -----
  if (req.method === 'GET' && pathname === '/api/auth/status') {
    const session = getSession(req);
    // role only when needed — avoid ARP on every status poll
    const role = session ? 'full' : clientRole(req);
    return json(res, 200, {
      enabled: AUTH_ENABLED,
      authenticated: !!(session && (session.auth || !AUTH_ENABLED)),
      user: session && session.user ? session.user : null,
      role
    });
  }

  if (req.method === 'POST' && pathname === '/api/login') {
    if (!AUTH_ENABLED) {
      return json(res, 200, { ok: true, enabled: false, user: 'open' });
    }
    let body = {};
    try { body = await readJson(req); } catch {
      return errorJson(res, 400, 'invalid JSON body');
    }
    // Username optional — defaults to configured admin user.
    const username = String(body.username || body.user || AUTH_USER).trim() || AUTH_USER;
    const password = String(body.password || body.token || '');
    if (!password) {
      return errorJson(res, 400, 'password required');
    }
    if (!safeEqualString(username, AUTH_USER) || !safeEqualString(password, AUTH_PASSWORD)) {
      return errorJson(res, 401, 'invalid username or password');
    }
    const token = createSessionToken(AUTH_USER);
    const maxAge = Math.floor(SESSION_TTL_MS / 1000);
    return json(res, 200, {
      ok: true,
      user: AUTH_USER,
      expiresInSec: maxAge
    }, { 'Set-Cookie': sessionCookieHeader(token) });
  }

  if (req.method === 'POST' && pathname === '/api/logout') {
    return json(res, 200, { ok: true }, { 'Set-Cookie': clearSessionCookieHeader() });
  }

  // Gate all other APIs when login is enabled (signature allowlist still passes requireLogin).
  if (!requireLogin(req, res, url, { htmlRedirect: false })) return;

  // Staff with a valid session are never treated as signature-only (skip ARP).
  const sessionEarly = getSession(req);
  if (!sessionEarly || !sessionEarly.auth) {
    if (clientRole(req) === 'signature' && !isSignatureAllowedPath(pathname, req.method)) {
      return denySignatureOnly(res, pathname);
    }
  }

  if (req.method === 'GET' && pathname === '/api/whoami') {
    const ip = clientIp(req);
    const session = getSession(req);
    const role = (session && session.auth) ? 'full' : clientRole(req);
    return json(res, 200, {
      role,
      ip,
      mac: (session && session.auth) ? null : (macForIp(ip) || null),
      signatureOnlyConfigured: SIGNATURE_ONLY_MACS.size + SIGNATURE_ONLY_IPS.size > 0,
      user: session && session.user ? session.user : null,
      authEnabled: AUTH_ENABLED
    });
  }

  // Minimal ticket queue for signature tablets (no rates / amounts).
  if (req.method === 'GET' && pathname === '/api/sign/queue') {
    const target = cleanText(query.get('date')) || todayIso();
    const tickets = db.prepare(`
      SELECT CAST(id AS INTEGER) AS id,
             CAST(COALESCE(NULLIF(TRIM(weighmaster_number), ''), weighmaster) AS TEXT) AS ticketNo,
             CAST(transaction_date AS TEXT) AS date,
             CAST(customer AS TEXT) AS customer,
             CAST(truck_id AS TEXT) AS truck,
             CAST(product_service AS TEXT) AS product,
             CASE WHEN driver_signature IS NOT NULL AND length(driver_signature) > 0
                  THEN 1 ELSE 0 END AS hasSig
      FROM quickbooks_transactions
      WHERE transaction_date IS NOT NULL
        AND substr(transaction_date, 1, 10) = ?
      ORDER BY id DESC;
    `).all(target).map(t => ({
      id: t.id,
      ticketNo: t.ticketNo || '',
      date: t.date || '',
      customer: t.customer || '',
      truck: t.truck || '',
      product: t.product || '',
      hasDriverSignature: !!t.hasSig
    }));
    return json(res, 200, { date: target, tickets });
  }

  if (req.method === 'GET' && pathname === '/api/status') {
    let stats = null;
    try { stats = dbStats(db); } catch { /* ignore */ }
    return json(res, 200, {
      status: 'ok',
      database: resolveDatabasePath(),
      stats,
      generatedAt: new Date().toISOString()
    });
  }


  // Backup inventory + integrity (read). Manual backup requires write token when set.
  if (req.method === 'GET' && pathname === '/api/backups') {
    // verify=1 runs integrity_check on each file (slower; optional).
    const inventory = listBackups({ verify: query.get('verify') === '1' });
    return json(res, 200, {
      live: { path: resolveDatabasePath(), stats: dbStats(db) },
      dir: inventory.dir,
      backups: inventory.backups,
      generatedAt: new Date().toISOString()
    });
  }
  if (req.method === 'POST' && pathname === '/api/backups/run') {
    if (!requireToken(req, res)) return;
    const result = runDbBackup({ force: true });
    return result.ok
      ? json(res, 200, result)
      : errorJson(res, 500, result.error || 'backup failed');
  }
  if (req.method === 'GET' && pathname === '/api/tickets') {
    return json(res, 200, loadTickets(db));
  }
  if (req.method === 'GET' && pathname === '/api/summary') {
    return json(res, 200, buildSummary(loadTickets(db)));
  }
  if (req.method === 'GET' && pathname === '/api/customers') {
    const rows = groupBy(loadTickets(db), t => t.customer || 'Unknown').map(([customer, tickets]) => ({
      customer,
      ticketCount: tickets.length,
      totalTons: sum(tickets, 'tons'),
      totalAmount: sum(tickets, 'amount'),
      arTotal: tickets.reduce((s, t) => s + (t.arAmount || 0), 0)
    })).sort((a, b) => b.totalAmount - a.totalAmount);
    return json(res, 200, rows);
  }
  if (req.method === 'GET' && pathname === '/api/products') {
    const rows = groupBy(loadTickets(db), t => t.product || 'Unknown').map(([product, tickets]) => ({
      product,
      ticketCount: tickets.length,
      totalTons: sum(tickets, 'tons'),
      totalAmount: sum(tickets, 'amount')
    })).sort((a, b) => b.totalAmount - a.totalAmount);
    return json(res, 200, rows);
  }
  if (req.method === 'GET' && pathname === '/api/ar') {
    return json(res, 200, loadTickets(db).filter(t => t.isAr));
  }
  if (req.method === 'GET' && pathname === '/api/day') {
    const target = cleanText(query.get('date')) || todayIso();
    return json(res, 200, buildDaySummary(loadTickets(db), target));
  }
  // User-editable dashboard config (KPI strip, tracker cards, …), shared on LAN.
  if (req.method === 'GET' && pathname === '/api/dashboard/cards') {
    let cfg = {};
    try { cfg = JSON.parse(getSetting(db, 'dashboard_cards') || '{}'); } catch {}
    if (Array.isArray(cfg)) cfg = { cards: cfg };           // migrate old shape
    if (!cfg || typeof cfg !== 'object') cfg = {};
    return json(res, 200, {
      cards: Array.isArray(cfg.cards) ? cfg.cards : [],
      kpis: Array.isArray(cfg.kpis) ? cfg.kpis : null,      // null = never configured → seed defaults client-side
      dayCustomers: Array.isArray(cfg.dayCustomers) ? cfg.dayCustomers : null,
      highlight: cfg.highlight && typeof cfg.highlight === 'object' ? cfg.highlight : null
    });
  }
  if (req.method === 'POST' && pathname === '/api/dashboard/cards') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    // Preserve fields not sent in this request so KPI saves don't wipe trackers and vice versa.
    let prev = {};
    try { prev = JSON.parse(getSetting(db, 'dashboard_cards') || '{}'); } catch {}
    if (Array.isArray(prev)) prev = { cards: prev };
    if (!prev || typeof prev !== 'object') prev = {};
    const cfg = {
      cards: Array.isArray(body.cards) ? body.cards.slice(0, 50)
        : (Array.isArray(prev.cards) ? prev.cards : []),
      kpis: Array.isArray(body.kpis) ? body.kpis.slice(0, 20)
        : (Array.isArray(prev.kpis) ? prev.kpis : null),
      dayCustomers: body.dayCustomers !== undefined
        ? (Array.isArray(body.dayCustomers) ? body.dayCustomers.slice(0, 20) : null)
        : (Array.isArray(prev.dayCustomers) ? prev.dayCustomers : null),
      highlight: body.highlight !== undefined
        ? (body.highlight && typeof body.highlight === 'object' ? body.highlight : null)
        : (prev.highlight && typeof prev.highlight === 'object' ? prev.highlight : null)
    };
    setSetting(db, 'dashboard_cards', JSON.stringify(cfg));
    return json(res, 200, {
      saved: cfg.cards.length,
      kpis: Array.isArray(cfg.kpis) ? cfg.kpis.length : 0
    });
  }
  if (req.method === 'GET' && pathname === '/api/ticket/defaults') {
    const pay = normalizePaymentMethod(query.get('payment') || 'CASH');
    const doc = nextDocumentNumber(db, pay);
    return json(res, 200, {
      weighmasterNumber: nextWeighmasterNumber(db),
      invoiceNumber: doc.number, // preview of document # for selected payment
      documentNumber: doc.number,
      documentType: doc.type,
      documentLabel: doc.label,
      payment: pay,
      series: {
        Invoice: nextDocumentNumber(db, 'AR'),
        'Sales Receipt': nextDocumentNumber(db, 'CASH'),
        Expense: nextDocumentNumber(db, 'EXPENSE')
      },
      date: todayIso(),
      signatureLine: SIGNATURE_LINE
    });
  }
  if (req.method === 'GET' && pathname === '/api/dropdowns') {
    // Products: merge historical ticket products + pricing-rule products so the
    // clickable list stays useful even before many tickets exist.
    const productSet = new Map();
    for (const v of distinctValues(db, 'product_service')) {
      if (v) productSet.set(String(v).toLowerCase(), v);
    }
    try {
      db.prepare(`
        SELECT DISTINCT product_service AS value
        FROM pricing
        WHERE product_service IS NOT NULL AND TRIM(product_service) <> ''
        ORDER BY product_service COLLATE NOCASE;
      `).all().forEach(r => {
        if (r.value) productSet.set(String(r.value).toLowerCase(), r.value);
      });
    } catch { /* pricing table empty / missing */ }
    const products = [...productSet.values()].sort((a, b) =>
      String(a).localeCompare(String(b), undefined, { sensitivity: 'base' })
    );
    // Origins: ticket origin field + stapled project addresses.
    const originSet = new Map();
    for (const v of distinctValues(db, 'origin')) {
      if (v) originSet.set(String(v).toLowerCase(), v);
    }
    try {
      loadProjectAddresses(db, null, true).forEach(p => {
        if (p.address) originSet.set(String(p.address).toLowerCase(), p.address);
      });
    } catch { /* ignore */ }
    const origins = [...originSet.values()].sort((a, b) =>
      String(a).localeCompare(String(b), undefined, { sensitivity: 'base' })
    );
    return json(res, 200, {
      customers: distinctValues(db, 'customer'),
      products,
      cities: distinctValues(db, 'city'),
      origins
    });
  }
  if (req.method === 'GET' && pathname === '/api/pricing') {
    const product = query.get('product');
    if (!cleanText(product)) {
      return json(res, 200, { status: 'Select a product', rate: null, expected: null });
    }
    const tons = cleanNumber(query.get('tons'));
    const projectAddressId = cleanNumber(query.get('projectAddressId') || query.get('projectId'));
    const rule = findPricingRule(
      db,
      query.get('customer'),
      product,
      parseDate(query.get('date')),
      query.get('class'),
      projectAddressId
    );
    if (!rule) return json(res, 200, { status: 'No price found', rate: null, expected: null });
    if (tons === null || tons <= 0) {
      return json(res, 200, {
        status: rule.pricingType === 'Hauling' ? 'Hauling price x tons' : `$${money(rule.rate)} x tons`,
        rate: rule.rate,
        accountType: rule.accountType,
        pricingType: rule.pricingType,
        paymentTerms: rule.paymentTerms,
        projectAddressId: rule.projectAddressId,
        expected: null
      });
    }
    const expected = calculateAmount(rule, tons);
    return json(res, 200, {
      status: `${describePricing(rule, tons)} = $${money(expected)}`,
      rate: rule.rate,
      accountType: rule.accountType,
      pricingType: rule.pricingType,
      paymentTerms: rule.paymentTerms,
      projectAddressId: rule.projectAddressId,
      expected
    });
  }
  if (req.method === 'GET' && pathname === '/api/tickets/recent') {
    return json(res, 200, loadRecent(db, Number(query.get('limit') || 100)));
  }
  if (req.method === 'GET' && pathname === '/api/check-duplicate') {
    const field = query.get('field');
    if (field !== 'weighmaster_number' && field !== 'invoice_number') {
      return errorJson(res, 400, 'field must be weighmaster_number or invoice_number');
    }
    return json(res, 200, { exists: identifierExists(db, field, query.get('value'), cleanNumber(query.get('excludeId'))) });
  }

  const ticketPrint = pathname.match(/^\/api\/ticket\/(\d+)\/print$/);
  if (req.method === 'GET' && ticketPrint) {
    const row = loadTicketById(db, ticketPrint[1]);
    if (!row) return errorJson(res, 404, 'ticket not found');
    // Load signature BLOB only for print (kept out of normal ticket GET).
    const sigRow = db.prepare(
      'SELECT driver_signature AS sig FROM quickbooks_transactions WHERE id = ?;'
    ).get(Number(ticketPrint[1]));
    row.driver_signature = sigRow?.sig || null;
    return text(res, 200, renderTicketPrint(row), 'text/html; charset=utf-8');
  }
  const ticketSignaturePng = pathname.match(/^\/api\/ticket\/(\d+)\/signature\.png$/);
  if (req.method === 'GET' && ticketSignaturePng) {
    const row = db.prepare('SELECT driver_signature AS sig FROM quickbooks_transactions WHERE id = ?;').get(Number(ticketSignaturePng[1]));
    const sig = sigBytes(row?.sig);
    if (!sig) return errorJson(res, 404, 'signature not found');
    // Short private cache — same ticket signature rarely changes mid-session.
    res.writeHead(200, {
      'Content-Type': imageMimeFromBytes(sig),
      'Content-Length': sig.length,
      'Cache-Control': 'private, max-age=120'
    });
    return res.end(sig);
  }
  const ticketDetail = pathname.match(/^\/api\/ticket\/(\d+)$/);
  if (req.method === 'GET' && ticketDetail) {
    const id = Number(ticketDetail[1]);
    const row = loadTicketById(db, id);
    if (!row) return errorJson(res, 404, 'ticket not found');
    const detail = rowToTicketDetail(row);
    // Signature kiosk gets a redacted ticket; staff session gets full detail.
    const redact = !isLoggedIn(req) && clientRole(req) === 'signature';
    const payload = redact ? ticketForSignPad(detail) : detail;
    // ?sig=1 embeds PNG as data-URL so the UI needs only ONE round-trip
    // (avoids slow second request for /signature.png on Wi‑Fi / tablets).
    if (query.get('sig') === '1' && payload.hasDriverSignature) {
      const sigRow = db.prepare(
        'SELECT driver_signature AS sig FROM quickbooks_transactions WHERE id = ?;'
      ).get(id);
      const dataUrl = imageDataUrlFromBytes(sigRow?.sig);
      if (dataUrl) payload.driverSignaturePngBase64 = dataUrl; // name kept for API compat
    }
    return json(res, 200, payload);
  }
  if (req.method === 'POST' && pathname === '/api/tickets') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    const validationError = validateWasteData(body) || validateTicketWeights(body, null);
    if (validationError) return errorJson(res, 400, validationError);
    if (!cleanText(body.customer)) return errorJson(res, 400, 'Customer / hauler is required.');
    // Product optional on inbound-only first save; required later for full pricing if desired.

    // Server is the sole authority for NEW ticket/invoice numbers (prevents
    // stale form defaults and concurrent tabs from colliding / leaving gaps).
    for (let attempt = 0; attempt < 6; attempt++) {
      try {
        const result = withImmediateTransaction(db, () => {
          const numbers = allocateTicketNumbers(db, body.payment);
          const entry = buildEntry(db, body, { mode: 'create', numbers });
          const id = insertEntry(db, entry);
          return {
            id,
            weighmasterNumber: entry.weighmasterNumber,
            invoiceNumber: entry.invoiceNumber,
            documentNumber: entry.documentNumber,
            documentType: entry.documentType,
            documentLabel: documentSeriesForPayment(entry.payment).label,
            rate: entry.rateResolved,
            amount: entry.amount,
            payment: entry.payment
          };
        });
        return json(res, 200, result);
      } catch (err) {
        // Unique index race — retry with a fresh MAX+1 allocation.
        if (/unique|constraint/i.test(String(err && err.message))) {
          continue;
        }
        console.error('[ticket create]', err);
        return errorJson(res, 500, err.message || 'Could not create ticket');
      }
    }
    return errorJson(res, 409, 'Could not allocate a unique ticket/invoice number — try again.');
  }
  if (req.method === 'PUT' && ticketDetail) {
    if (!requireToken(req, res)) return;
    const id = Number(ticketDetail[1]);
    const body = await readJson(req);
    const existing = loadTicketById(db, id);
    if (!existing) return errorJson(res, 404, `Ticket id ${id} not found.`);
    const validationError = validateWasteData(body) || validateTicketWeights(body, existing);
    if (validationError) return errorJson(res, 400, validationError);
    if (!cleanText(body.customer) && !cleanText(existing.customer)) {
      return errorJson(res, 400, 'Customer / hauler is required.');
    }
    const entry = buildEntry(db, body, { mode: 'update', existing });
    if (identifierExists(db, 'weighmaster_number', entry.weighmasterNumber, id)) {
      return errorJson(res, 409, `Weighmaster # ${entry.weighmasterNumber} already exists.`);
    }
    if (identifierExists(db, 'invoice_number', entry.invoiceNumber, id)) {
      return errorJson(res, 409, `Invoice # ${entry.invoiceNumber} already exists.`);
    }
    try {
      const rows = updateEntry(db, id, entry);
      if (rows === 0) return errorJson(res, 404, `Ticket id ${id} not found.`);
      return json(res, 200, {
        id,
        weighmasterNumber: entry.weighmasterNumber,
        invoiceNumber: entry.invoiceNumber,
        documentNumber: entry.documentNumber,
        documentType: entry.documentType,
        documentLabel: documentSeriesForPayment(entry.payment).label,
        rate: entry.rateResolved,
        amount: entry.amount,
        payment: entry.payment
      });
    } catch {
      return errorJson(res, 409, 'Weighmaster or document number already exists.');
    }
  }
  const ticketSignature = pathname.match(/^\/api\/ticket\/(\d+)\/signature$/);
  if (req.method === 'POST' && ticketSignature) {
    // Staff session OR signature-only tablet may save a signature.
    const canSign = isLoggedIn(req) || clientRole(req) === 'signature';
    if (!canSign) {
      errorJson(res, 401, 'login required');
      return;
    }
    const body = await readJson(req);
    let bytes;
    try {
      bytes = decodeSignature(body.pngBase64, true);
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
    const result = db.prepare('UPDATE quickbooks_transactions SET driver_signature = ? WHERE id = ?;')
      .run(bytes, Number(ticketSignature[1]));
    return result.changes === 0
      ? errorJson(res, 404, `Ticket id ${ticketSignature[1]} not found.`)
      : json(res, 200, { id: Number(ticketSignature[1]), bytes: bytes.length });
  }

  if (req.method === 'GET' && pathname === '/api/export/powerbi.csv') {
    if (!requireToken(req, res)) return;
    const csv = renderCsv(loadWeighmasterTickets(db, 2147483647), POWER_BI_COLUMNS);
    return fileBytes(res, 200, Buffer.from(csv, 'utf8'), 'text/csv; charset=utf-8', `transactions_${stamp()}.csv`);
  }
  if (req.method === 'GET' && pathname === '/api/export/quickbooks-extract.csv') {
    if (!requireToken(req, res)) return;
    const csv = renderCsv(loadQuickBooksExtract(db));
    return fileBytes(res, 200, Buffer.from(csv, 'utf8'), 'text/csv; charset=utf-8', `quickbooks_extract_${stamp()}.csv`);
  }

  // Download a single day's tickets as CSV (offline; no email). Defaults to today.
  if (req.method === 'GET' && pathname === '/api/export/day.csv') {
    if (!requireToken(req, res)) return;
    const { date, rows } = dayTickets(db, query.get('date'));
    const csv = renderCsv(rows, POWER_BI_COLUMNS);
    return fileBytes(res, 200, Buffer.from(csv, 'utf8'), 'text/csv; charset=utf-8', `tickets_${date.replace(/-/g, '')}.csv`);
  }

  // ---------- email export (the one online-only feature) ----------

  // Current config, password never returned (only whether one is set).
  if (req.method === 'GET' && pathname === '/api/email/settings') {
    const cfg = getEmailConfig(db);
    return json(res, 200, {
      host: cfg.host || '',
      port: cfg.port || 587,
      secure: !!cfg.secure,
      username: cfg.username || '',
      from: cfg.from || '',
      to: cfg.to || '',
      subjectPrefix: cfg.subjectPrefix || 'AIS Daily Tickets',
      tlsInsecure: !!cfg.tlsInsecure,
      hasPassword: !!(cfg.password || process.env.AIS_SMTP_PASSWORD),
      passwordFromEnv: !!process.env.AIS_SMTP_PASSWORD,
      autoSendEnabled: !!cfg.autoSendEnabled,
      autoSendTime: cfg.autoSendTime || '18:00',
      autoSendWhich: cfg.autoSendWhich || 'today',
      skipEmptyDays: !!cfg.skipEmptyDays,
      autoLastSent: getSetting(db, 'auto_last_sent') || null,
      autoLastResult: getSetting(db, 'auto_last_result') || null
    });
  }

  // Save config. A blank password keeps the previously stored one (so the UI can
  // re-save other fields without forcing a password re-entry).
  if (req.method === 'POST' && pathname === '/api/email/settings') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    const existing = getEmailConfig(db);
    const cfg = {
      host: cleanText(body.host) || '',
      port: cleanNumber(body.port) || 587,
      secure: !!body.secure,
      username: cleanText(body.username) || '',
      from: cleanText(body.from) || '',
      to: cleanText(body.to) || '',
      subjectPrefix: cleanText(body.subjectPrefix) || 'AIS Daily Tickets',
      tlsInsecure: !!body.tlsInsecure,
      autoSendEnabled: !!body.autoSendEnabled,
      autoSendTime: /^\d{1,2}:\d{2}$/.test(String(body.autoSendTime || '')) ? String(body.autoSendTime) : '18:00',
      autoSendWhich: body.autoSendWhich === 'yesterday' ? 'yesterday' : 'today',
      skipEmptyDays: !!body.skipEmptyDays,
      password: (body.password !== undefined && body.password !== '')
        ? String(body.password)
        : (existing.password || '')
    };
    setSetting(db, 'email_config', JSON.stringify(cfg));
    return json(res, 200, { ok: true });
  }

  // Send a small test email to confirm the SMTP settings work.
  if (req.method === 'POST' && pathname === '/api/email/test') {
    if (!requireToken(req, res)) return;
    const cfg = getEmailConfig(db);
    const err = validateEmailConfig(cfg);
    if (err) return errorJson(res, 400, err);
    try {
      await smtpSend({
        host: cfg.host, port: cfg.port, secure: cfg.secure,
        user: cfg.username, pass: effectiveEmailPassword(cfg),
        from: cfg.from, to: splitRecipients(cfg.to),
        subject: `${cfg.subjectPrefix} — test message`,
        text: `This is a test from the AIS LAN dashboard at ${new Date().toLocaleString()}.\nIf you received this, email export is configured correctly.`,
        rejectUnauthorized: !cfg.tlsInsecure
      });
      return json(res, 200, { ok: true, to: cfg.to });
    } catch (e) {
      return errorJson(res, 502, 'Email failed: ' + e.message);
    }
  }

  // Email a single day's CSV as an attachment. Body: { date? }.
  if (req.method === 'POST' && pathname === '/api/email/send-day') {
    if (!requireToken(req, res)) return;
    const cfg = getEmailConfig(db);
    const err = validateEmailConfig(cfg);
    if (err) return errorJson(res, 400, err);
    const body = await readJson(req);
    try {
      const { date, count } = await sendDayEmail(db, cfg, body.date);
      return json(res, 200, { ok: true, date, count, to: cfg.to });
    } catch (e) {
      return errorJson(res, 502, 'Email failed: ' + e.message);
    }
  }

  // Companies master list (editable AR terms / default payment per hauler).
  if (req.method === 'GET' && pathname === '/api/companies') {
    return json(res, 200, loadCompanies(db));
  }
  if (req.method === 'GET' && pathname === '/api/companies/lookup') {
    const name = cleanText(query.get('name') || query.get('customer'));
    if (!name) return errorJson(res, 400, 'name is required');
    const row = findCompanyByName(db, name);
    if (!row) return errorJson(res, 404, `No company profile for '${name}'.`);
    const projects = loadProjectAddresses(db, name, true);
    return json(res, 200, { ...row, projects });
  }

  // Stapled project addresses per customer (job sites with optional terms).
  if (req.method === 'GET' && pathname === '/api/projects') {
    const customer = query.get('customer') || query.get('name');
    const activeOnly = query.get('all') !== '1' && query.get('all') !== 'true';
    return json(res, 200, loadProjectAddresses(db, customer, activeOnly));
  }
  if (req.method === 'POST' && pathname === '/api/projects') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      const id = upsertProjectAddress(db, null, body);
      return json(res, 200, { id, project: findProjectAddressById(db, id) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  const projectRoute = pathname.match(/^\/api\/projects\/(\d+)$/);
  if (projectRoute && req.method === 'GET') {
    const row = findProjectAddressById(db, projectRoute[1]);
    return row
      ? json(res, 200, row)
      : errorJson(res, 404, `Project id ${projectRoute[1]} not found.`);
  }
  if (projectRoute && req.method === 'PUT') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      const id = upsertProjectAddress(db, Number(projectRoute[1]), body);
      return json(res, 200, { id, project: findProjectAddressById(db, id) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  if (projectRoute && req.method === 'DELETE') {
    if (!requireToken(req, res)) return;
    const rows = deleteProjectAddress(db, projectRoute[1]);
    return rows === 0
      ? errorJson(res, 404, `Project id ${projectRoute[1]} not found.`)
      : json(res, 200, { id: Number(projectRoute[1]) });
  }
  if (req.method === 'POST' && pathname === '/api/companies') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      const id = upsertCompany(db, null, body);
      return json(res, 200, { id, company: findCompanyByName(db, body.name || body.customer) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  const companyRoute = pathname.match(/^\/api\/companies\/(\d+)$/);
  if (companyRoute && req.method === 'PUT') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      const id = upsertCompany(db, Number(companyRoute[1]), body);
      return json(res, 200, { id });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  if (companyRoute && req.method === 'DELETE') {
    if (!requireToken(req, res)) return;
    const rows = deleteCompany(db, companyRoute[1]);
    return rows === 0
      ? errorJson(res, 404, `Company id ${companyRoute[1]} not found.`)
      : json(res, 200, { id: Number(companyRoute[1]) });
  }

  if (req.method === 'GET' && pathname === '/api/pricing/rules') {
    return json(res, 200, loadPricingRules(db));
  }
  if (req.method === 'GET' && pathname === '/api/pricing/dropdowns') {
    const companyNames = loadCompanies(db).map(c => c.name);
    const ticketCustomers = distinctValues(db, 'customer');
    const projectCustomers = loadProjectAddresses(db, null, false).map(p => p.customer);
    const merged = [...new Set([...companyNames, ...ticketCustomers, ...projectCustomers])]
      .filter(Boolean)
      .sort((a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: 'base' }));
    return json(res, 200, {
      customers: merged,
      classes: distinctValues(db, 'class'),
      products: distinctValues(db, 'product_service'),
      companies: loadCompanies(db),
      projects: loadProjectAddresses(db, null, false)
    });
  }
  if (req.method === 'POST' && pathname === '/api/pricing/rules') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      return json(res, 200, { id: upsertPricingRule(db, null, body) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  const pricingRule = pathname.match(/^\/api\/pricing\/rules\/(\d+)$/);
  if (pricingRule && req.method === 'PUT') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      return json(res, 200, { id: upsertPricingRule(db, Number(pricingRule[1]), body) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  if (pricingRule && req.method === 'DELETE') {
    if (!requireToken(req, res)) return;
    const rows = deletePricingRule(db, pricingRule[1]);
    return rows === 0 ? errorJson(res, 404, `Pricing id ${pricingRule[1]} not found.`) : json(res, 200, { id: Number(pricingRule[1]) });
  }
  if (req.method === 'POST' && pathname === '/api/pricing/import') {
    if (!requireToken(req, res)) return;
    try {
      const body = await readBody(req, MAX_UPLOAD_BYTES);
      const file = parseMultipartFile(req, body);
      if (file.length === 0) return errorJson(res, 400, 'no file uploaded');
      return json(res, 200, importPricingCsv(db, file.toString('utf8')));
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }

  if (req.method === 'GET' && pathname === '/api/trucks') {
    return json(res, 200, loadTrucks(db));
  }
  if (req.method === 'GET' && pathname === '/api/trucks/lookup') {
    const truckId = cleanText(query.get('truckId'));
    if (!truckId) return errorJson(res, 400, 'truckId is required');
    const tareWeight = findTruckTare(db, truckId);
    return tareWeight === null
      ? errorJson(res, 404, `No tare on file for truck '${truckId}'.`)
      : json(res, 200, { truckId, tareWeight });
  }
  if (req.method === 'POST' && pathname === '/api/trucks') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      return json(res, 200, { id: upsertTruck(db, null, body) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  const truck = pathname.match(/^\/api\/trucks\/(\d+)$/);
  if (truck && req.method === 'PUT') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    try {
      return json(res, 200, { id: upsertTruck(db, Number(truck[1]), body) });
    } catch (err) {
      return errorJson(res, 400, err.message);
    }
  }
  if (truck && req.method === 'DELETE') {
    if (!requireToken(req, res)) return;
    const rows = deleteTruck(db, truck[1]);
    return rows === 0 ? errorJson(res, 404, `Truck id ${truck[1]} not found.`) : json(res, 200, { id: Number(truck[1]) });
  }

  if (req.method === 'GET' && pathname === '/api/billing/ar') {
    const range = parseBillingRange(query.get('from'), query.get('to'));
    if (range.error) return errorJson(res, 400, range.error);
    const detail = loadArBillingTickets(db, range.from, range.to);
    const totals = loadArBillingCustomerTotals(db, range.from, range.to);
    const grandTotal = totals.reduce((s, row) => s + (cleanNumber(row.totalAmount) || 0), 0);
    return json(res, 200, { from: range.from, to: range.to, detail, totals, grandTotal });
  }
  if (req.method === 'GET' && pathname === '/api/billing/ar.csv') {
    if (!requireToken(req, res)) return;
    const range = parseBillingRange(query.get('from'), query.get('to'));
    if (range.error) return errorJson(res, 400, range.error);
    const kind = (query.get('kind') || 'detail').toLowerCase();
    if (kind !== 'detail' && kind !== 'totals') return errorJson(res, 400, "kind must be 'detail' or 'totals'");
    const rows = kind === 'detail'
      ? loadArBillingTickets(db, range.from, range.to)
      : loadArBillingCustomerTotals(db, range.from, range.to);
    const columns = inferColumns(rows).filter(c => c.toLowerCase() !== 'id');
    const fileBase = kind === 'detail' ? 'ar_billing_detail' : 'ar_billing_totals';
    return fileBytes(res, 200, Buffer.from(renderCsv(rows, columns), 'utf8'), 'text/csv; charset=utf-8', `${fileBase}_${range.from.replace(/-/g, '')}_${range.to.replace(/-/g, '')}.csv`);
  }

  return errorJson(res, 404, 'not found');
}

function groupBy(items, keyFn) {
  const map = new Map();
  for (const item of items) {
    const key = keyFn(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  }
  return [...map.entries()];
}

function stamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}${String(d.getSeconds()).padStart(2, '0')}`;
}

// Legal text — verbatim from the Excel/VBA ticket (samples.pdf).
const TICKET_TERMS = 'This is to certify that the following described commodity was weighed, measured, or counted by a weighmaster, whose signature is on this certificate, who is a recognized authority of accuracy, as prescribed by Chapter 7 (commencing with Section 12700) of Division 5 of the California Business and Professions Code, administered by the Division of Measurement Standards of the California Department of Food and Agriculture. The undersigned individual signing this document on behalf of Customer acknowledges that he or she has read and understands the terms and conditions on the reverse side and that he or she has the authority to sign this document on behalf of the customer.';
const TICKET_WMCERT = '1. Unacceptable Waste. "Unacceptable Waste" means: (a) any material that is defined as hazardous material or hazardous waste, untreated medical or infectious waste, radioactive waste, or unprofiled sewage sludge under any applicable law, regulation or permit ("Applicable Law"); or (b) any other material that may present an endangerment to the facility, to its employees or to public health or safety. Customer shall not deliver any Unacceptable Waste. If Customer delivers Unacceptable Waste, Operator may, in its sole discretion: (x) reject the Unacceptable Waste; or, as Customer\'s agent, dispose of the Unacceptable Waste at a location authorized to accept it in accordance with all Applicable Laws and charge Customer all direct and indirect costs incurred for its handling, transportation and disposal.';

function formatTicketDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? `${m[2]}/${m[3]}/${m[1]}` : String(iso || '');
}

// Replicates the printed ticket in weighticket2026/samples.pdf (the Excel/VBA output).
function renderTicketPrint(row) {
  const textValue = key => String(row[key] ?? '');
  const numValue = key => cleanNumber(row[key]);
  let ticketNo = textValue('weighmaster_number');
  if (!cleanText(ticketNo)) ticketNo = textValue('weighmaster');
  const gross = numValue('t1');
  const tare = numValue('t2');
  const netLb = gross !== null && tare !== null ? Math.abs(gross - tare) : null;
  const tonsValue = numValue('tonage') ?? numValue('qty');
  const payment = textValue('payment');
  const isAr = payment.toUpperCase() === 'AR';
  const rawProduct = textValue('product_service');
  const productDisplay = rawProduct.trim().toUpperCase() === 'MIXED' ? 'Mixed C&D' : rawProduct;
  const sigDataUrl = imageDataUrlFromBytes(row.driver_signature);
  // VBA: $ column stays empty; Amount column and Total Amount carry the paid
  // amount. AR tickets hide prices entirely.
  const amountNum = numValue('amount');
  const amountText = isAr || amountNum === null ? '' : '$' + money(amountNum);
  // Cert header only when a checkbox was ticked (sample prints none by default).
  const weightCert = Number(row.weight_certificate) === 1;
  const recyclingCert = Number(row.recycling_cd_certificate) === 1;
  let certHeader = '';
  if (weightCert) certHeader = 'WEIGHMASTER CERTIFICATE';
  if (recyclingCert) certHeader = certHeader ? certHeader + ' / RECYCLING C&D CERTIFICATE' : 'RECYCLING C&D CERTIFICATE';
  const wasteData = cleanText(textValue('waste_data')) || '';
  const timeIn = textValue('time_in');
  const timeOut = textValue('time_out');
  const data = {
    ticketNo,
    date: formatTicketDate(textValue('transaction_date')),
    timeText: timeIn + (timeOut ? ' - ' + timeOut : ''),
    truck: textValue('truck_id'),
    customer: textValue('customer'),
    wasteData,
    memo: textValue('memo'),
    product: productDisplay,
    origin: textValue('origin'),
    gross: gross === null ? '' : Math.round(gross).toLocaleString('en-US'),
    tare: tare === null ? '' : Math.round(tare).toLocaleString('en-US'),
    netLb: netLb === null ? '' : Math.round(netLb).toLocaleString('en-US'),
    netTons: tonsValue === null ? '' : compact(tonsValue),
    amountText,
    payment,
    deputyName: cleanText(row.by_person) || SIGNATURE_LINE,
    certHeader,
    signature: sigDataUrl
  };

  const sigHtml = data.signature
    ? `<img class="sig-img" src="${data.signature}" alt="signature">`
    : '<span class="sig-line">____________________</span>';
  const custText = h(data.customer) + (data.wasteData ? `&nbsp;&nbsp;&nbsp;<b>Waste Data:</b> ${h(data.wasteData)}` : '');
  const copy = `
<section class="ticket">
  <div class="hdr">
    <div class="brand">
      <div class="bizname">AMERICAN INDUSTRIAL SERVICES</div>
      <div>5626 CHERRY AVE.</div>
      <div>LONG BEACH, CA 90805</div>
      <div>(562) 272-8060</div>
    </div>
    <div class="logo"><img src="/ais.png" alt="AIS"></div>
    <div class="tno">
      <div class="tk-line">Ticket: ${h(data.ticketNo)}</div>
      <div class="dt-line">Date: ${h(data.date)}</div>
      <div class="dt-line">Time: ${h(data.timeText)}</div>
      <div class="scale-u">Scale</div>
    </div>
  </div>
  <div class="mid">
    <div class="mid-left">
      <div class="truck"><b>Truck:</b> ${h(data.truck)}</div>
      <div class="cust"><b>Customer:</b> ${custText}</div>
      <div class="memo"><b>Memo:</b> ${h(data.memo)}</div>
    </div>
    <div class="mid-right">
      <div class="wrow"><span class="wl">Gross:</span><span class="wv">${h(data.gross)}</span><span class="io">In</span><span class="sc">Scale</span></div>
      <div class="wrow"><span class="wl">Tare:</span><span class="wv">${h(data.tare)}</span><span class="io">Out</span><span class="sc">Scale</span></div>
      <div class="wrow net"><span class="wl">Net:</span><span class="wv">${h(data.netLb)}</span><span class="io"></span><span class="sc"></span></div>
    </div>
  </div>
  <table class="mat">
    <thead><tr><th class="cAddr">Address</th><th class="cMat">Materials &amp; Services Quantity</th><th class="cUnit">Unit</th><th class="cDol num">$</th><th class="cAmt num">Amount</th></tr></thead>
    <tbody><tr>
      <td>${h(data.origin)}</td>
      <td><span class="prod">${h(data.product)}</span><span class="qty">${h(data.netTons)}</span></td>
      <td>Ton</td>
      <td class="num"></td>
      <td class="num">${h(data.amountText)}</td>
    </tr></tbody>
  </table>
  <div class="pay"><div><b>Payment Method:</b> ${h(data.payment)}</div><div><b>Total Amount:</b> ${h(data.amountText)}</div></div>
  <p class="terms">${h(TICKET_TERMS)}</p>
  <div class="siggap"></div>
  <div class="sigs">
    <div class="drv"><b>Driver:</b> ${sigHtml}</div>
    <div class="dep"><b>Deputy Weighmaster:</b> <span class="italic">${h(data.deputyName)}</span></div>
  </div>
  ${data.certHeader ? `<div class="certhdr">${h(data.certHeader)}</div>` : ''}
  <p class="wmcert">${h(TICKET_WMCERT)}</p>
</section>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Ticket ${h(data.ticketNo)} - Weighmaster Certificate</title>
<style>
/* Perforated stock: the sheet tears at exactly 5.5in. Each copy is locked to
   a 5.5in band (top/bottom padding inside the band) so content never crosses
   the perforation — mirrors the VBA's forced row-height symmetry. */
@page { size: letter; margin: 0 0.65in; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; font-family: Calibri, "Segoe UI", Arial, sans-serif; font-size: 10pt; line-height: 1.25; color: #000; background: #fff; }
.sheet { width: 7.2in; margin: 0 auto; }
/* Measured from samples.pdf: copy 1 header at 0.56in from paper top; copy 2
   exactly 5.5in below it. Side margins ~0.65in (Excel CenterHorizontally). */
.ticket { height: 5.5in; padding: 0.55in 0 0.25in; overflow: hidden; page-break-inside: avoid; display: flex; flex-direction: column; }
.ticket + .ticket { margin-top: 0; }
b { font-weight: 700; }
.italic { font-style: italic; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
/* Header: brand | logo | ticket block */
.hdr { display: grid; grid-template-columns: 2.6fr 2.6fr 2.1fr; align-items: start; }
.brand .bizname { font-weight: 700; font-size: 9pt; }
.brand div { font-size: 10pt; }
.brand .bizname + div, .brand div { line-height: 1.3; }
.logo { text-align: center; }
.logo img { max-width: 2.1in; max-height: 0.55in; }
.tno .tk-line { font-weight: 700; font-size: 12pt; }
.tno .dt-line { font-size: 9pt; }
.tno .scale-u { text-decoration: underline; font-size: 10pt; margin-top: 1px; }
/* Middle: truck/customer/memo | weights */
.mid { display: grid; grid-template-columns: 1.45fr 1fr; margin-top: 2px; }
.mid-left .truck { font-weight: 400; margin-bottom: 8px; }
.mid-left .cust { margin: 6px 0 10px; }
.mid-left .memo { margin: 4px 0 8px; }
.wrow { display: grid; grid-template-columns: 1.1fr 1.2fr 0.55fr 0.8fr; column-gap: 6px; align-items: baseline; }
.wrow .wl { text-align: right; }
.wrow .wv { text-align: right; font-variant-numeric: tabular-nums; }
.wrow.net { font-weight: 700; margin-top: 4px; }
/* Materials table */
table.mat { width: 100%; border-collapse: collapse; margin-top: 6px; }
table.mat th { border-top: 1.5px solid #000; border-bottom: 1.5px solid #000; font-weight: 700; padding: 2px 4px; text-align: left; font-size: 10pt; }
table.mat th.num { text-align: right; }
table.mat td { padding: 3px 4px; font-size: 10pt; vertical-align: top; }
th.cAddr { width: 30%; } th.cMat { width: 33%; } th.cUnit { width: 9%; } th.cDol { width: 12%; } th.cAmt { width: 16%; }
td .prod { }
td .qty { float: right; font-variant-numeric: tabular-nums; }
/* Payment / total */
.pay { display: flex; justify-content: space-between; align-items: baseline; margin: 8px 0 2px; }
/* Legal texts */
.terms { font-size: 6pt; line-height: 1.25; margin: 2px 0 0; text-align: justify; }
.siggap { flex: 1; min-height: 0.5in; }
/* Signatures */
.sigs { display: flex; justify-content: space-between; align-items: baseline; margin: 2px 0 6px; }
.sig-line { letter-spacing: 1px; }
.sig-img { max-height: 34px; vertical-align: bottom; }
.certhdr { text-align: center; font-weight: 700; font-size: 10pt; text-decoration: underline; margin: 2px 0; }
.wmcert { font-size: 5.5pt; line-height: 1.3; text-align: center; margin: 4px auto 0; width: 88%; }
@media screen {
  body { background: #ececec; padding: 24px 0; }
  .sheet { background: #fff; padding: 0 0.5in; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
  /* show where the perforation falls — screen only, never prints */
  .ticket + .ticket { border-top: 1px dashed #b36; }
}
@media print { body { background: #fff; padding: 0; } .sheet { box-shadow: none; padding: 0; } }
</style>
</head>
<body><div class="sheet">${copy}${copy}</div><script>window.addEventListener('load', () => setTimeout(() => window.print(), 50));</script></body>
</html>`;
}

function h(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Small mtime cache so each page navigation is not a disk hit. */
const viewCache = new Map(); // name -> { mtimeMs, bytes }
function readViewCached(fileName) {
  const full = path.join(VIEWS_DIR, fileName);
  const st = fs.statSync(full);
  const hit = viewCache.get(fileName);
  if (hit && hit.mtimeMs === st.mtimeMs) return hit.bytes;
  const bytes = fs.readFileSync(full);
  viewCache.set(fileName, { mtimeMs: st.mtimeMs, bytes });
  return bytes;
}

function serveStatic(req, res, url) {
  let pathname = url.pathname;
  if (pathname === '/sign') pathname = '/sign.html';
  if (pathname === '/login') pathname = '/login.html';

  // Staff login gate first (signed cookie — no ARP, no DB). Public paths pass through.
  if (!requireLogin(req, res, { pathname, search: url.search || '' }, { htmlRedirect: true })) {
    return;
  }

  // Signature-only tablets only when not logged in as staff.
  const staff = getSession(req);
  if (!staff || !staff.auth) {
    if (clientRole(req) === 'signature') {
      if (!isSignatureAllowedPath(pathname, req.method)) {
        if (req.method === 'GET' && !pathname.startsWith('/api/')) {
          res.writeHead(302, { Location: '/sign.html' });
          return res.end();
        }
        return denySignatureOnly(res, pathname);
      }
      if (pathname === '/') {
        res.writeHead(302, { Location: '/sign.html' });
        return res.end();
      }
    }
  }

  const pageMatch = pathname === '/' ? 'index.html'
    : /^\/([a-z-]+\.html)$/.exec(pathname)?.[1];
  if (pageMatch && fs.existsSync(path.join(VIEWS_DIR, pageMatch))) {
    // Logged-in users hitting login page → go home
    if (pageMatch === 'login.html' && AUTH_ENABLED && getSession(req)) {
      res.writeHead(302, { Location: '/' });
      return res.end();
    }
    const html = readViewCached(pageMatch);
    return fileBytes(res, 200, html, 'text/html; charset=utf-8');
  }
  const requestPath = decodeURIComponent(pathname);
  // Support both /auth.js and /public/auth.js style paths.
  let fullPath = path.resolve(PUBLIC_DIR, `.${requestPath}`);
  if (!fullPath.startsWith(path.resolve(PUBLIC_DIR))) {
    return errorJson(res, 403, 'forbidden');
  }
  if (!fs.existsSync(fullPath) && requestPath.startsWith('/public/')) {
    fullPath = path.resolve(PUBLIC_DIR, `.${requestPath.slice('/public'.length)}`);
  }
  fs.stat(fullPath, (err, stat) => {
    if (err || !stat.isFile()) return errorJson(res, 404, 'not found');
    res.writeHead(200, { 'Content-Type': mimeType(fullPath), 'Content-Length': stat.size });
    fs.createReadStream(fullPath).pipe(res);
  });
}

function mimeType(file) {
  switch (path.extname(file).toLowerCase()) {
    case '.html': return 'text/html; charset=utf-8';
    case '.js': return 'text/javascript; charset=utf-8';
    case '.css': return 'text/css; charset=utf-8';
    case '.png': return 'image/png';
    case '.jpg':
    case '.jpeg': return 'image/jpeg';
    case '.svg': return 'image/svg+xml';
    case '.ico': return 'image/x-icon';
    default: return 'application/octet-stream';
  }
}

// ---------- Access log (terminal only — not exposed in the dashboard UI) ----------
// AIS_ACCESS_LOG=0 off | api (default) | all
const ACCESS_LOG_MODE = String(process.env.AIS_ACCESS_LOG || 'api').toLowerCase();

function pushAccessLog(entry) {
  const show =
    ACCESS_LOG_MODE === 'all' ||
    (ACCESS_LOG_MODE !== '0' && ACCESS_LOG_MODE !== 'off' && ACCESS_LOG_MODE !== 'false' && entry.api);
  if (show) {
    console.log(`[http] ${entry.method} ${entry.path} → ${entry.status} ${entry.ms}ms  ip=${entry.ip}`);
  }
}

function attachAccessLog(req, res) {
  const started = process.hrtime.bigint();
  let statusCode = 0;
  const origWriteHead = res.writeHead;
  res.writeHead = function (code, ...rest) {
    if (typeof code === 'number') statusCode = code;
    return origWriteHead.call(this, code, ...rest);
  };
  res.on('finish', () => {
    try {
      const ms = Number(process.hrtime.bigint() - started) / 1e6;
      let path = req.url || '/';
      try {
        const u = new URL(req.url || '/', 'http://local');
        path = u.pathname + (u.search || '');
      } catch { /* keep raw */ }
      pushAccessLog({
        at: new Date().toISOString(),
        method: req.method || 'GET',
        path,
        status: statusCode || res.statusCode || 0,
        ms: Math.round(ms * 10) / 10,
        ip: clientIp(req) || '?',
        api: path.startsWith('/api/')
      });
    } catch { /* never break responses */ }
  });
}

const server = http.createServer(async (req, res) => {
  attachAccessLog(req, res);
  try {
    const url = new URL(req.url || '/', `http://${req.headers.host || 'localhost'}`);
    if (url.pathname.startsWith('/api/')) {
      await handleApi(req, res, url);
    } else {
      serveStatic(req, res, url);
    }
  } catch (err) {
    console.error(err);
    if (!res.headersSent) errorJson(res, 500, err.message || 'server error');
  }
});


function lanIPv4Addresses() {
  const out = [];
  const ifaces = os.networkInterfaces();
  for (const name of Object.keys(ifaces)) {
    for (const info of ifaces[name] || []) {
      if (info.family !== 'IPv4' && info.family !== 4) continue;
      if (info.internal) continue;
      out.push({ name, address: info.address });
    }
  }
  return out;
}

server.on('error', (err) => {
  if (err && err.code === 'EADDRINUSE') {
    console.error(`ERROR: port ${PORT} is already in use (another AIS/LAN process?).`);
    console.error('  Stop the other process, then run:  ./start-lan.sh');
    console.error(`  Hint:  ss -tlnp | grep ${PORT}`);
  } else {
    console.error('ERROR: server failed to start:', err && err.message ? err.message : err);
  }
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  const tokenState = process.env.AIS_DASHBOARD_TOKEN ? 'enabled' : 'open';
  const bindLabel = HOST === '0.0.0.0' ? 'all interfaces (LAN)' : HOST;
  console.log(`AIS LAN dashboard listening on ${bindLabel}, port ${PORT}`);
  console.log(`  This computer:  http://127.0.0.1:${PORT}`);
  const lans = lanIPv4Addresses();
  if (lans.length) {
    for (const { name, address } of lans) {
      console.log(`  LAN (${name}):   http://${address}:${PORT}`);
    }
    console.log('  Other phones/PCs on the SAME Wi-Fi must use a LAN URL above (not 127.0.0.1).');
  } else {
    console.log('  LAN: (no IPv4 address found — connect Wi-Fi/Ethernet, then restart)');
  }
  if (SIGNATURE_ONLY_MACS.size || SIGNATURE_ONLY_IPS.size) {
    console.log('Signature-only kiosk devices:');
    for (const mac of SIGNATURE_ONLY_MACS) console.log(`  MAC ${mac} → /sign.html only`);
    for (const ip of SIGNATURE_ONLY_IPS) console.log(`  IP  ${ip} → /sign.html only`);
  } else {
    console.log('Signature-only kiosk: (none configured — set AIS_SIGNATURE_ONLY_MACS / AIS_SIGNATURE_ONLY_IPS)');
  }
  console.log(`Database: ${resolveDatabasePath()}`);
  if (AUTH_ENABLED) {
    console.log(`Login: enabled (user=${AUTH_USER}, session=${Math.floor(SESSION_TTL_MS / 3600000)}h)`);
    if (!process.env.AIS_DASHBOARD_PASSWORD && !process.env.AIS_DASHBOARD_TOKEN) {
      console.log('  Using default password "ais" — set AIS_DASHBOARD_PASSWORD to change it.');
    }
  } else {
    console.log('Login: disabled (open LAN access)');
  }
  console.log(`API token header gate: ${tokenState}`);

  setTimeout(() => { autoSendTick().catch(e => console.error('[auto-email]', e.message)); }, 5000);
  setInterval(() => { autoSendTick().catch(e => console.error('[auto-email]', e.message)); }, 60000);

  // Deferred, non-blocking-ish backup (fast copy; skips if recent). Never on critical path.
  setTimeout(() => {
    try { runDbBackup({ force: false }); }
    catch (e) { console.error('[backup]', e.message || e); }
  }, 60_000);
  setInterval(() => {
    try { runDbBackup({ force: false }); }
    catch (e) { console.error('[backup]', e.message || e); }
  }, 24 * 60 * 60 * 1000);
});
