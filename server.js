'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const net = require('node:net');
const tls = require('node:tls');
const crypto = require('node:crypto');
const { URL } = require('node:url');
const { getDb, resolveDatabasePath, sigBytes, runDbBackup } = require('./db');

const PORT = Number(process.env.PORT || process.env.AIS_DASHBOARD_PORT || 5000);
const HOST = process.env.AIS_DASHBOARD_HOST || '0.0.0.0';
const PROJECT_ROOT = __dirname;
const PUBLIC_DIR = path.join(PROJECT_ROOT, 'public');
const VIEWS_DIR = path.join(PROJECT_ROOT, 'views');
const SIGNATURE_LINE = 'Jesus Polanco';
const MAX_JSON_BYTES = 1_000_000;
const MAX_UPLOAD_BYTES = 5_000_000;
const MAX_SIGNATURE_BYTES = 250_000;



function json(res, status, value) {
  const body = JSON.stringify(value);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body)
  });
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

function requireToken(req, res) {
  const token = process.env.AIS_DASHBOARD_TOKEN || '';
  if (!token.trim()) return true;
  if (req.headers['x-ais-token'] === token) return true;
  errorJson(res, 401, 'missing or invalid X-AIS-Token');
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

function nextWeighmasterNumber(db, startFrom = 117964) {
  const row = db.prepare(`
    WITH ticket_ids(value) AS (
      SELECT CAST(weighmaster_number AS INTEGER)
      FROM quickbooks_transactions
      WHERE weighmaster_number IS NOT NULL
        AND TRIM(weighmaster_number) <> ''
        AND TRIM(weighmaster_number) NOT GLOB '*[^0-9]*'
      UNION ALL
      SELECT CAST(weighmaster AS INTEGER)
      FROM quickbooks_transactions
      WHERE (weighmaster_number IS NULL OR TRIM(weighmaster_number) = '')
        AND weighmaster IS NOT NULL
        AND TRIM(weighmaster) <> ''
        AND TRIM(weighmaster) NOT GLOB '*[^0-9]*'
    )
    SELECT COALESCE(MAX(value), ?) + 1 AS next FROM ticket_ids;
  `).get(startFrom);
  return String(row.next);
}

function nextInvoiceNumber(db, startFrom = 910936) {
  const row = db.prepare(`
    WITH invoice_ids(value) AS (
      SELECT CAST(invoice_number AS INTEGER)
      FROM quickbooks_transactions
      WHERE invoice_number IS NOT NULL
        AND invoice_number GLOB '[0-9]*'
      UNION ALL
      SELECT CAST(document_number AS INTEGER)
      FROM quickbooks_transactions
      WHERE document_number IS NOT NULL
        AND document_number GLOB '[0-9]*'
        AND CAST(document_number AS INTEGER) >= 900000
    )
    SELECT COALESCE(MAX(value), ?) + 1 AS next FROM invoice_ids;
  `).get(startFrom);
  return String(row.next);
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
  return db.prepare(`
    WITH ticket_rows AS (
      SELECT *,
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
  return db.prepare('SELECT * FROM quickbooks_transactions WHERE id = ?;').get(Number(id));
}


function rowToTicketDetail(row) {
  const result = {};
  for (const [key, value] of Object.entries(row)) {
    if (key === 'driver_signature') continue;
    result[toCamelCase(key)] = value;
  }
  result.hasDriverSignature = sigBytes(row.driver_signature) !== null;
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

function findPricingRule(db, customer, product, onDate, className) {
  const productValue = cleanText(product);
  if (!productValue) return null;

  const customerValue = cleanText(customer);
  const classValue = cleanText(className);
  const rows = db.prepare(`
    SELECT id, customer, class, rate, account_type, pricing_type,
           base_rate, allowed_tons, overage_rate, miles, mile_rate, surcharge_percent,
           effective_date, end_date
    FROM pricing
    WHERE TRIM(product_service) = TRIM(?) COLLATE NOCASE
      AND (customer IS NULL OR ? IS NULL OR TRIM(customer) = TRIM(?) COLLATE NOCASE)
      AND (class IS NULL OR TRIM(class) = '' OR ? IS NULL OR TRIM(class) = TRIM(?) COLLATE NOCASE);
  `).all(productValue, customerValue, customerValue, classValue, classValue);

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
    surchargePercent: cleanNumber(best.surcharge_percent)
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
    SELECT id AS id,
           customer AS customer,
           class AS class,
           product_service AS product,
           pricing_type AS pricingType,
           rate AS rate,
           base_rate AS baseRate,
           allowed_tons AS allowedTons,
           overage_rate AS overageRate,
           miles AS miles,
           mile_rate AS mileRate,
           surcharge_percent AS surchargePercent,
           account_type AS accountType,
           effective_date AS effectiveDate,
           end_date AS endDate,
           notes AS notes
    FROM pricing
    ORDER BY customer COLLATE NOCASE, product_service COLLATE NOCASE, effective_date DESC;
  `).all();
}

function upsertPricingRule(db, id, body) {
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
    accountType: normalizeAccountType(body.accountType),
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
          surcharge_percent = ?, account_type = ?, effective_date = ?, end_date = ?,
          notes = ?, updated_at_utc = CURRENT_TIMESTAMP
      WHERE id = ?;
    `).run(
      row.customer, row.className, row.product, row.rate, row.pricingType,
      row.baseRate, row.allowedTons, row.overageRate, row.miles, row.mileRate,
      row.surchargePercent, row.accountType, row.effectiveDate, row.endDate,
      row.notes, row.id
    );
    return row.id;
  }

  const result = db.prepare(`
    INSERT INTO pricing (
      customer, class, product_service, rate, pricing_type, base_rate, allowed_tons,
      overage_rate, miles, mile_rate, surcharge_percent, account_type, effective_date,
      end_date, notes
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
  `).run(
    row.customer, row.className, row.product, row.rate, row.pricingType,
    row.baseRate, row.allowedTons, row.overageRate, row.miles, row.mileRate,
    row.surchargePercent, row.accountType, row.effectiveDate, row.endDate, row.notes
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

function validateWasteData(req) {
  const product = (cleanText(req.product) || '').toUpperCase();
  if ((product === '40YD' || product === '10YD') && !cleanText(req.wasteData)) {
    return `Waste Data is required for ${product} material type.`;
  }
  return null;
}

function buildEntry(db, req) {
  const date = parseDate(req.date);
  const weighmasterNumber = cleanText(req.weighmasterNumber) || nextWeighmasterNumber(db);
  const invoiceNumber = cleanText(req.invoiceNumber) || nextInvoiceNumber(db);

  let tonage = null;
  const gross = cleanNumber(req.gross);
  const tare = cleanNumber(req.tare);
  if (gross !== null && tare !== null) {
    tonage = Math.abs(gross - tare) / 2000.0;
    if (tonage > 0 && tonage < 1) tonage = 1;
  }

  const rule = findPricingRule(db, req.customer, req.product, date, null);
  const rawPayment = (cleanText(req.payment) || '').toUpperCase();
  let payment = rawPayment || paymentMethodFor(rule?.accountType) || 'CASH';
  let amount = cleanNumber(req.amount);
  let amountAutoComputed = false;
  if (amount === null && rule && tonage !== null) {
    amount = calculateAmount(rule, tonage);
    amountAutoComputed = true;
  }
  if (amountAutoComputed && amount !== null && payment === 'CREDIT') amount *= 1.03;
  if (amount !== null && amount < 0) payment = 'EXPENSE';

  let rateResolved = null;
  if (/^weight$/i.test(cleanText(req.product) || '')) {
    rateResolved = amount;
  } else if (amount !== null && tonage !== null && tonage > 0) {
    rateResolved = Math.round((amount / tonage) * 100) / 100;
  }

  return {
    documentType: 'Invoice',
    documentNumber: invoiceNumber,
    transactionDate: date?.iso || null,
    customer: cleanText(req.customer),
    weighmaster: cleanText(req.weighmaster) || cleanText(req.byPerson),
    productService: cleanText(req.product),
    qty: tonage,
    rate: rateResolved,
    amount,
    truckId: cleanText(req.truck),
    poOrigin: null,
    memo: cleanText(req.memo),
    payment,
    ar: payment === 'AR' ? amount : null,
    className: null,
    scalemaster: null,
    invoiceDate: null,
    description: null,
    depositTo: null,
    location: null,
    roundAmount: null,
    dateCopy: date?.iso || null,
    wasteData: cleanText(req.wasteData),
    timeIn: cleanText(req.timeIn),
    timeOut: cleanText(req.timeOut),
    t1: gross,
    t2: tare,
    tonage,
    origin: cleanText(req.origin),
    cash: payment === 'CASH' ? amount : null,
    card: payment === 'CREDIT' || payment === 'DEBIT' ? amount : null,
    weighmasterNumber,
    invoiceNumber,
    city: cleanText(req.city),
    ticketNotes: cleanText(req.notes),
    weightCertificate: req.weightCertificate ? 1 : 0,
    recyclingCdCertificate: req.recyclingCdCertificate ? 1 : 0,
    byPerson: cleanText(req.byPerson),
    driverSignature: decodeSignature(req.driverSignaturePngBase64, false),
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
      driver_signature
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
  `).run(...entryParams(entry, sourceFile, sourceRowNumber));
  return result.lastInsertRowid;
}

function updateEntry(db, id, entry) {
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
      driver_signature = ?
    WHERE id = ?;
  `).run(...entryParams(entry), Number(id)).changes;
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
    entry.driverSignature
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

  if (req.method === 'GET' && pathname === '/api/status') {
    return json(res, 200, { status: 'ok', database: databasePath, generatedAt: new Date().toISOString() });
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
  // User-editable dashboard config (tracker cards, day-panel rows, highlight
  // card customer), shared across all LAN devices.
  if (req.method === 'GET' && pathname === '/api/dashboard/cards') {
    let cfg = {};
    try { cfg = JSON.parse(getSetting(db, 'dashboard_cards') || '{}'); } catch {}
    if (Array.isArray(cfg)) cfg = { cards: cfg };           // migrate old shape
    if (!cfg || typeof cfg !== 'object') cfg = {};
    return json(res, 200, {
      cards: Array.isArray(cfg.cards) ? cfg.cards : [],
      dayCustomers: Array.isArray(cfg.dayCustomers) ? cfg.dayCustomers : null,
      highlight: cfg.highlight && typeof cfg.highlight === 'object' ? cfg.highlight : null
    });
  }
  if (req.method === 'POST' && pathname === '/api/dashboard/cards') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    const cfg = {
      cards: Array.isArray(body.cards) ? body.cards.slice(0, 50) : [],
      dayCustomers: Array.isArray(body.dayCustomers) ? body.dayCustomers.slice(0, 20) : null,
      highlight: body.highlight && typeof body.highlight === 'object' ? body.highlight : null
    };
    setSetting(db, 'dashboard_cards', JSON.stringify(cfg));
    return json(res, 200, { saved: cfg.cards.length });
  }
  if (req.method === 'GET' && pathname === '/api/ticket/defaults') {
    return json(res, 200, {
      weighmasterNumber: nextWeighmasterNumber(db),
      invoiceNumber: nextInvoiceNumber(db),
      date: todayIso(),
      signatureLine: SIGNATURE_LINE
    });
  }
  if (req.method === 'GET' && pathname === '/api/dropdowns') {
    return json(res, 200, {
      customers: distinctValues(db, 'customer'),
      products: distinctValues(db, 'product_service'),
      cities: distinctValues(db, 'city'),
      origins: distinctValues(db, 'po_origin')
    });
  }
  if (req.method === 'GET' && pathname === '/api/pricing') {
    const product = query.get('product');
    if (!cleanText(product)) {
      return json(res, 200, { status: 'Select a product', rate: null, expected: null });
    }
    const tons = cleanNumber(query.get('tons'));
    const rule = findPricingRule(db, query.get('customer'), product, parseDate(query.get('date')), query.get('class'));
    if (!rule) return json(res, 200, { status: 'No price found', rate: null, expected: null });
    if (tons === null || tons <= 0) {
      return json(res, 200, {
        status: rule.pricingType === 'Hauling' ? 'Hauling price x tons' : `$${money(rule.rate)} x tons`,
        rate: rule.rate,
        accountType: rule.accountType,
        pricingType: rule.pricingType,
        expected: null
      });
    }
    const expected = calculateAmount(rule, tons);
    return json(res, 200, {
      status: `${describePricing(rule, tons)} = $${money(expected)}`,
      rate: rule.rate,
      accountType: rule.accountType,
      pricingType: rule.pricingType,
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
    return row ? text(res, 200, renderTicketPrint(row), 'text/html; charset=utf-8') : errorJson(res, 404, 'ticket not found');
  }
  const ticketSignaturePng = pathname.match(/^\/api\/ticket\/(\d+)\/signature\.png$/);
  if (req.method === 'GET' && ticketSignaturePng) {
    const row = db.prepare('SELECT driver_signature AS sig FROM quickbooks_transactions WHERE id = ?;').get(Number(ticketSignaturePng[1]));
    const sig = sigBytes(row?.sig);
    return sig
      ? fileBytes(res, 200, sig, 'image/png')
      : errorJson(res, 404, 'signature not found');
  }
  const ticketDetail = pathname.match(/^\/api\/ticket\/(\d+)$/);
  if (req.method === 'GET' && ticketDetail) {
    const row = loadTicketById(db, ticketDetail[1]);
    return row ? json(res, 200, rowToTicketDetail(row)) : errorJson(res, 404, 'ticket not found');
  }
  if (req.method === 'POST' && pathname === '/api/tickets') {
    if (!requireToken(req, res)) return;
    const body = await readJson(req);
    const validationError = validateWasteData(body);
    if (validationError) return errorJson(res, 400, validationError);

    for (let attempt = 0; attempt < 4; attempt++) {
      const entry = buildEntry(db, body);
      if (identifierExists(db, 'weighmaster_number', entry.weighmasterNumber, null)) {
        return errorJson(res, 409, `Weighmaster # ${entry.weighmasterNumber} already exists.`);
      }
      if (identifierExists(db, 'invoice_number', entry.invoiceNumber, null)) {
        return errorJson(res, 409, `Invoice # ${entry.invoiceNumber} already exists.`);
      }
      try {
        const id = insertEntry(db, entry);
        return json(res, 200, {
          id,
          weighmasterNumber: entry.weighmasterNumber,
          invoiceNumber: entry.invoiceNumber,
          rate: entry.rateResolved,
          amount: entry.amount,
          payment: entry.payment
        });
      } catch (err) {
        if (!/constraint/i.test(err.message) || cleanText(body.weighmasterNumber) || cleanText(body.invoiceNumber)) {
          return errorJson(res, 409, 'Weighmaster or invoice number already exists.');
        }
      }
    }
    return errorJson(res, 409, 'Could not allocate a unique ticket number after retries - try again.');
  }
  if (req.method === 'PUT' && ticketDetail) {
    if (!requireToken(req, res)) return;
    const id = Number(ticketDetail[1]);
    const body = await readJson(req);
    const validationError = validateWasteData(body);
    if (validationError) return errorJson(res, 400, validationError);
    const entry = buildEntry(db, body);
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
        rate: entry.rateResolved,
        amount: entry.amount,
        payment: entry.payment
      });
    } catch {
      return errorJson(res, 409, 'Weighmaster or invoice number already exists.');
    }
  }
  const ticketSignature = pathname.match(/^\/api\/ticket\/(\d+)\/signature$/);
  if (req.method === 'POST' && ticketSignature) {
    if (!requireToken(req, res)) return;
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

  if (req.method === 'GET' && pathname === '/api/pricing/rules') {
    return json(res, 200, loadPricingRules(db));
  }
  if (req.method === 'GET' && pathname === '/api/pricing/dropdowns') {
    return json(res, 200, {
      customers: distinctValues(db, 'customer'),
      classes: distinctValues(db, 'class'),
      products: distinctValues(db, 'product_service')
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
  const printSig = sigBytes(row.driver_signature);
  const sigDataUrl = printSig
    ? `data:image/png;base64,${printSig.toString('base64')}`
    : '';
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

function serveStatic(req, res, url) {
  const pageMatch = url.pathname === '/' ? 'index.html'
    : /^\/([a-z-]+\.html)$/.exec(url.pathname)?.[1];
  if (pageMatch && fs.existsSync(path.join(VIEWS_DIR, pageMatch))) {
    const html = fs.readFileSync(path.join(VIEWS_DIR, pageMatch));
    return fileBytes(res, 200, html, 'text/html; charset=utf-8');
  }
  const requestPath = decodeURIComponent(url.pathname);
  const fullPath = path.resolve(PUBLIC_DIR, `.${requestPath}`);
  if (!fullPath.startsWith(path.resolve(PUBLIC_DIR))) {
    return errorJson(res, 403, 'forbidden');
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

const server = http.createServer(async (req, res) => {
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


server.listen(PORT, HOST, () => {
  const tokenState = process.env.AIS_DASHBOARD_TOKEN ? 'enabled' : 'open';
  console.log(`AIS LAN dashboard listening on http://${HOST}:${PORT}`);
  console.log(`Database: ${resolveDatabasePath()}`);
  console.log(`Write/export token gate: ${tokenState}`);

  setTimeout(() => { autoSendTick().catch(e => console.error('[auto-email]', e.message)); }, 5000);
  setInterval(() => { autoSendTick().catch(e => console.error('[auto-email]', e.message)); }, 60000);

  setTimeout(runDbBackup, 10000);
  setInterval(runDbBackup, 24 * 60 * 60 * 1000);
});
