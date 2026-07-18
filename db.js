'use strict';
// db.js - SQLite database module (node:sqlite, zero dependencies).
// Owns the connection, schema, blob handling, and daily backups.

const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');

const PROJECT_ROOT = __dirname;
const DATA_DIR = path.join(PROJECT_ROOT, 'data');

let database;
let databasePath;

function getDb() {
  const resolved = resolveDatabasePath();
  if (!database || databasePath !== resolved) {
    fs.mkdirSync(path.dirname(resolved), { recursive: true });
    database = new DatabaseSync(resolved);
    databasePath = resolved;
    database.exec('PRAGMA journal_mode = WAL;');
    ensureSchema(database);
  }
  return database;
}

function resolveDatabasePath() {
  if (process.env.AIS_DASHBOARD_DB && process.env.AIS_DASHBOARD_DB.trim()) {
    return path.resolve(process.env.AIS_DASHBOARD_DB.trim());
  }
  return path.join(DATA_DIR, 'quickbooks.db');
}

function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS quickbooks_transactions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      source_file TEXT NOT NULL,
      source_row_number INTEGER NOT NULL,
      document_type TEXT,
      document_number TEXT,
      date_copy TEXT,
      waste_data TEXT,
      memo TEXT,
      truck_id TEXT,
      po_origin TEXT,
      transaction_date TEXT,
      customer TEXT,
      weighmaster TEXT,
      product_service TEXT,
      qty REAL,
      rate REAL,
      amount REAL,
      payment TEXT,
      ar REAL,
      week TEXT,
      biweekly TEXT,
      class TEXT,
      scalemaster TEXT,
      month TEXT,
      invoice_date TEXT,
      description TEXT,
      deposit_to TEXT,
      location TEXT,
      round_amount REAL,
      time_in TEXT,
      time_out TEXT,
      t1 REAL,
      t2 REAL,
      tonage REAL,
      origin TEXT,
      cash REAL,
      card REAL,
      imported_at_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(source_file, source_row_number)
    );

    CREATE INDEX IF NOT EXISTS ix_quickbooks_transactions_date
      ON quickbooks_transactions(transaction_date);

    CREATE INDEX IF NOT EXISTS ix_quickbooks_transactions_customer
      ON quickbooks_transactions(customer);

    CREATE TABLE IF NOT EXISTS pricing (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer TEXT,
      class TEXT,
      product_service TEXT NOT NULL,
      rate REAL NOT NULL,
      pricing_type TEXT NOT NULL DEFAULT 'Flat',
      base_rate REAL,
      allowed_tons REAL,
      overage_rate REAL,
      miles REAL,
      mile_rate REAL,
      surcharge_percent REAL,
      account_type TEXT NOT NULL DEFAULT 'Regular',
      effective_date TEXT,
      notes TEXT,
      updated_at_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE INDEX IF NOT EXISTS ix_pricing_lookup
      ON pricing(customer, product_service);

    CREATE INDEX IF NOT EXISTS ix_pricing_product
      ON pricing(product_service);

    CREATE TABLE IF NOT EXISTS trucks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      truck_id TEXT NOT NULL,
      tare_weight REAL,
      company TEXT,
      notes TEXT,
      updated_at_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE UNIQUE INDEX IF NOT EXISTS ux_trucks_truck_id
      ON trucks(truck_id COLLATE NOCASE);

    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `);

  [
    ['quickbooks_transactions', 'time_in', 'TEXT'],
    ['quickbooks_transactions', 'time_out', 'TEXT'],
    ['quickbooks_transactions', 't1', 'REAL'],
    ['quickbooks_transactions', 't2', 'REAL'],
    ['quickbooks_transactions', 'tonage', 'REAL'],
    ['quickbooks_transactions', 'origin', 'TEXT'],
    ['quickbooks_transactions', 'cash', 'REAL'],
    ['quickbooks_transactions', 'card', 'REAL'],
    ['quickbooks_transactions', 'weighmaster_number', 'TEXT'],
    ['quickbooks_transactions', 'city', 'TEXT'],
    ['quickbooks_transactions', 'invoice_number', 'TEXT'],
    ['quickbooks_transactions', 'ticket_notes', 'TEXT'],
    ['quickbooks_transactions', 'weight_certificate', 'INTEGER'],
    ['quickbooks_transactions', 'recycling_cd_certificate', 'INTEGER'],
    ['quickbooks_transactions', 'by_person', 'TEXT'],
    ['quickbooks_transactions', 'driver_signature', 'BLOB'],
    ['pricing', 'class', 'TEXT'],
    ['pricing', 'account_type', "TEXT NOT NULL DEFAULT 'Regular'"],
    ['pricing', 'pricing_type', "TEXT NOT NULL DEFAULT 'Flat'"],
    ['pricing', 'base_rate', 'REAL'],
    ['pricing', 'allowed_tons', 'REAL'],
    ['pricing', 'overage_rate', 'REAL'],
    ['pricing', 'miles', 'REAL'],
    ['pricing', 'mile_rate', 'REAL'],
    ['pricing', 'surcharge_percent', 'REAL'],
    ['pricing', 'end_date', 'TEXT']
  ].forEach(([table, column, type]) => addColumnIfMissing(db, table, column, type));

  try {
    db.exec(`
      CREATE UNIQUE INDEX IF NOT EXISTS ux_quickbooks_transactions_weighmaster_number
        ON quickbooks_transactions(weighmaster_number)
        WHERE weighmaster_number IS NOT NULL AND TRIM(weighmaster_number) <> '';
    `);
  } catch {}

  try {
    db.exec(`
      CREATE UNIQUE INDEX IF NOT EXISTS ux_quickbooks_transactions_invoice_number
        ON quickbooks_transactions(invoice_number)
        WHERE invoice_number IS NOT NULL AND TRIM(invoice_number) <> '';
    `);
  } catch {}
}

function addColumnIfMissing(db, table, column, type) {
  try {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${type};`);
  } catch {}
}
// node:sqlite returns BLOB columns as Uint8Array (not Buffer), so accept both.
function sigBytes(value) {
  if (Buffer.isBuffer(value)) return value.length > 0 ? value : null;
  if (value instanceof Uint8Array) return value.length > 0 ? Buffer.from(value) : null;
  return null;
}
function runDbBackup() {
  try {
    const db = getDb();
    const dbFile = resolveDatabasePath();
    const backupDir = path.join(path.dirname(dbFile), 'backups');
    fs.mkdirSync(backupDir, { recursive: true });
    db.exec('PRAGMA wal_checkpoint(TRUNCATE);');
    const stamp = new Date().toISOString().slice(0, 10);
    const dest = path.join(backupDir, `quickbooks-${stamp}.db`);
    fs.copyFileSync(dbFile, dest);
    const old = fs.readdirSync(backupDir)
      .filter(f => /^quickbooks-\d{4}-\d{2}-\d{2}\.db$/.test(f))
      .sort().reverse().slice(14);
    for (const f of old) fs.unlinkSync(path.join(backupDir, f));
    console.log(`[backup] written: ${dest}`);
  } catch (err) {
    console.error('[backup] failed:', err.message || err);
  }
}

module.exports = { getDb, resolveDatabasePath, sigBytes, runDbBackup };
