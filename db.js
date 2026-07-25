'use strict';
// db.js - SQLite database module (node:sqlite, zero dependencies).
// Owns the connection, schema, blob handling, and daily backups.

const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');

const PROJECT_ROOT = __dirname;
const DATA_DIR = path.join(PROJECT_ROOT, 'data');
/** Keep this many dated daily backups (plus backups/latest.db). */
const BACKUP_KEEP_DAYS = Number(process.env.AIS_BACKUP_KEEP_DAYS || 14);

let database;
let databasePath;

let walCheckpointTimer = null;

function getDb() {
  const resolved = resolveDatabasePath();
  if (!database || databasePath !== resolved) {
    fs.mkdirSync(path.dirname(resolved), { recursive: true });
    database = new DatabaseSync(resolved);
    databasePath = resolved;
    // Durability / concurrency defaults for a LAN multi-device app.
    database.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA synchronous = NORMAL;
      PRAGMA foreign_keys = ON;
      PRAGMA busy_timeout = 500;
      PRAGMA temp_store = MEMORY;
      PRAGMA cache_size = -8000;
    `);
    ensureSchema(database);
    // Keep WAL from growing forever (large WAL = slower reads). PASSIVE = non-blocking.
    if (walCheckpointTimer) clearInterval(walCheckpointTimer);
    walCheckpointTimer = setInterval(() => {
      try {
        if (database) database.exec('PRAGMA wal_checkpoint(PASSIVE);');
      } catch { /* ignore */ }
    }, 5 * 60 * 1000);
    if (typeof walCheckpointTimer.unref === 'function') walCheckpointTimer.unref();
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

    -- Company / hauler master: default payment + editable AR terms (Net 10 / Net 30).
    CREATE TABLE IF NOT EXISTS companies (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      account_type TEXT NOT NULL DEFAULT 'Regular',
      payment_terms TEXT,
      notes TEXT,
      updated_at_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE UNIQUE INDEX IF NOT EXISTS ux_companies_name
      ON companies(name COLLATE NOCASE);

    -- Stapled project / job-site addresses per customer.
    -- Each project can carry its own AR terms and (via pricing.project_address_id) rates.
    CREATE TABLE IF NOT EXISTS project_addresses (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer TEXT NOT NULL,
      label TEXT,
      address TEXT NOT NULL,
      city TEXT,
      account_type TEXT,
      payment_terms TEXT,
      notes TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      updated_at_utc TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE INDEX IF NOT EXISTS ix_project_addresses_customer
      ON project_addresses(customer COLLATE NOCASE);
  `);

  // Seed companies from ticket + pricing customers (idempotent).
  try {
    db.exec(`
      INSERT OR IGNORE INTO companies (name, account_type, payment_terms)
      SELECT DISTINCT TRIM(customer), 'Regular', NULL
      FROM quickbooks_transactions
      WHERE customer IS NOT NULL AND TRIM(customer) <> '';
    `);
    db.exec(`
      INSERT OR IGNORE INTO companies (name, account_type, payment_terms)
      SELECT DISTINCT TRIM(customer), 'Regular', NULL
      FROM pricing
      WHERE customer IS NOT NULL AND TRIM(customer) <> '';
    `);
  } catch { /* first boot / empty */ }

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
    ['quickbooks_transactions', 'payment_terms', 'TEXT'],
    ['quickbooks_transactions', 'due_date', 'TEXT'],
    ['pricing', 'class', 'TEXT'],
    ['pricing', 'account_type', "TEXT NOT NULL DEFAULT 'Regular'"],
    ['pricing', 'pricing_type', "TEXT NOT NULL DEFAULT 'Flat'"],
    ['pricing', 'base_rate', 'REAL'],
    ['pricing', 'allowed_tons', 'REAL'],
    ['pricing', 'overage_rate', 'REAL'],
    ['pricing', 'miles', 'REAL'],
    ['pricing', 'mile_rate', 'REAL'],
    ['pricing', 'surcharge_percent', 'REAL'],
    ['pricing', 'end_date', 'TEXT'],
    ['pricing', 'payment_terms', 'TEXT'],
    ['pricing', 'project_address_id', 'INTEGER'],
    ['quickbooks_transactions', 'project_address_id', 'INTEGER']
  ].forEach(([table, column, type]) => addColumnIfMissing(db, table, column, type));

  try {
    db.exec(`
      CREATE INDEX IF NOT EXISTS ix_pricing_project
        ON pricing(project_address_id);
    `);
  } catch {}

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

function backupDirFor(dbFile) {
  return path.join(path.dirname(dbFile), 'backups');
}

function sqlQuotePath(filePath) {
  // SQLite string literal: escape single quotes.
  return `'${String(filePath).replace(/'/g, "''")}'`;
}

function dbStats(db) {
  const count = (sql) => {
    try { return db.prepare(sql).get().n; } catch { return null; }
  };
  return {
    tickets: count('SELECT COUNT(*) AS n FROM quickbooks_transactions'),
    signatures: count(`SELECT COUNT(*) AS n FROM quickbooks_transactions
      WHERE driver_signature IS NOT NULL AND length(driver_signature) > 0`),
    pricing: count('SELECT COUNT(*) AS n FROM pricing'),
    trucks: count('SELECT COUNT(*) AS n FROM trucks'),
    settings: count('SELECT COUNT(*) AS n FROM app_settings'),
    maxTicketId: (() => {
      try { return db.prepare('SELECT MAX(id) AS id FROM quickbooks_transactions').get().id; }
      catch { return null; }
    })()
  };
}

function verifyBackupFile(filePath) {
  const db = new DatabaseSync(filePath, { readOnly: true });
  try {
    const integrity = db.prepare('PRAGMA integrity_check').get();
    const ok = integrity && String(Object.values(integrity)[0]).toLowerCase() === 'ok';
    if (!ok) {
      throw new Error(`integrity_check failed: ${JSON.stringify(integrity)}`);
    }
    return dbStats(db);
  } finally {
    db.close();
  }
}

function pruneOldBackups(backupDir) {
  const keep = Number.isFinite(BACKUP_KEEP_DAYS) && BACKUP_KEEP_DAYS > 0 ? BACKUP_KEEP_DAYS : 14;
  const old = fs.readdirSync(backupDir)
    .filter(f => /^quickbooks-\d{4}-\d{2}-\d{2}\.db$/.test(f))
    .sort()
    .reverse()
    .slice(keep);
  for (const f of old) {
    try { fs.unlinkSync(path.join(backupDir, f)); }
    catch (err) { console.warn('[backup] prune failed:', f, err.message || err); }
  }
  return { kept: keep, removed: old };
}

/**
 * Fast on-disk backup (does not freeze the HTTP server).
 * Default: WAL checkpoint + file copy (milliseconds on small DBs).
 * Optional: AIS_BACKUP_VACUUM=1 for slower VACUUM INTO + integrity check.
 * Auto-backups skip if latest.db is newer than 6 hours (unless force=true).
 */
function runDbBackup(options = {}) {
  const force = !!options.force;
  const useVacuum = !!options.vacuum || /^(1|true|yes)$/i.test(String(process.env.AIS_BACKUP_VACUUM || ''));
  const verify = options.verify !== false && useVacuum;
  try {
    const db = getDb();
    const dbFile = resolveDatabasePath();
    const backupDir = backupDirFor(dbFile);
    fs.mkdirSync(backupDir, { recursive: true });

    const latest = path.join(backupDir, 'latest.db');
    if (!force && fs.existsSync(latest)) {
      const ageMs = Date.now() - fs.statSync(latest).mtimeMs;
      if (ageMs < 6 * 60 * 60 * 1000) {
        console.log('[backup] skip — latest.db is fresh (<6h); pass force to override');
        return { ok: true, skipped: true, reason: 'fresh', latest };
      }
    }

    const stamp = new Date().toISOString().slice(0, 10); // YYYY-MM-DD (UTC)
    const dest = path.join(backupDir, `quickbooks-${stamp}.db`);
    const tmp = path.join(backupDir, `.tmp-backup-${process.pid}.db`);

    try { fs.unlinkSync(tmp); } catch { /* ok */ }

    const liveStats = dbStats(db);
    let method = 'checkpoint_copy';
    if (useVacuum) {
      method = 'vacuum_into';
      try {
        db.exec(`VACUUM INTO ${sqlQuotePath(tmp)}`);
      } catch (vacuumErr) {
        method = 'checkpoint_copy';
        console.warn('[backup] VACUUM INTO failed, falling back to copy:', vacuumErr.message || vacuumErr);
        db.exec('PRAGMA wal_checkpoint(PASSIVE);');
        fs.copyFileSync(dbFile, tmp);
      }
    } else {
      // PASSIVE checkpoint does not block writers long; then copy main file.
      try { db.exec('PRAGMA wal_checkpoint(PASSIVE);'); } catch { /* ok */ }
      fs.copyFileSync(dbFile, tmp);
    }

    let backupStats = null;
    if (verify) {
      backupStats = verifyBackupFile(tmp);
    } else {
      backupStats = liveStats;
    }

    fs.renameSync(tmp, dest);
    try { fs.copyFileSync(dest, latest); } catch (err) {
      console.warn('[backup] could not update latest.db:', err.message || err);
    }

    for (const side of [`${dest}-wal`, `${dest}-shm`, `${latest}-wal`, `${latest}-shm`]) {
      try { fs.unlinkSync(side); } catch { /* ok */ }
    }

    const pruned = pruneOldBackups(backupDir);
    const size = fs.statSync(dest).size;
    console.log(
      `[backup] ok method=${method} file=${dest} bytes=${size} ` +
      `tickets=${backupStats.tickets} signatures=${backupStats.signatures} ` +
      `maxId=${backupStats.maxTicketId} liveTickets=${liveStats.tickets} ` +
      `pruned=${pruned.removed.length}`
    );
    return { ok: true, method, dest, latest, size, liveStats, backupStats, pruned };
  } catch (err) {
    console.error('[backup] failed:', err.message || err);
    return { ok: false, error: String(err.message || err) };
  }
}

/** List dated backups + latest with size/mtime. Integrity is optional (can be slow). */
function listBackups({ verify = false } = {}) {
  const dbFile = resolveDatabasePath();
  const backupDir = backupDirFor(dbFile);
  if (!fs.existsSync(backupDir)) return { dir: backupDir, backups: [] };
  const names = fs.readdirSync(backupDir)
    .filter(f => f === 'latest.db' || /^quickbooks-\d{4}-\d{2}-\d{2}\.db$/.test(f));
  const backups = names.map(name => {
    const filePath = path.join(backupDir, name);
    const st = fs.statSync(filePath);
    let integrity = null;
    let stats = null;
    if (verify) {
      try {
        stats = verifyBackupFile(filePath);
        integrity = 'ok';
      } catch (err) {
        integrity = err.message || 'failed';
      }
    }
    return {
      name,
      path: filePath,
      bytes: st.size,
      mtime: st.mtime.toISOString(),
      integrity,
      stats
    };
  }).sort((a, b) => String(b.mtime).localeCompare(String(a.mtime)));
  return { dir: backupDir, live: resolveDatabasePath(), backups };
}

module.exports = {
  getDb,
  resolveDatabasePath,
  sigBytes,
  runDbBackup,
  listBackups,
  dbStats,
  verifyBackupFile
};
