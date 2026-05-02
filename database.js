const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function nowISO() {
  return new Date().toISOString();
}

function todayISO() {
  return nowISO().slice(0, 10);
}

function normalizeInstallmentTotal(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 1;
  const whole = Math.floor(numeric);
  return Math.min(360, Math.max(1, whole));
}

function addMonthsIso(isoDate, monthOffset) {
  const text = String(isoDate || "").trim();
  const [yearText, monthText, dayText] = text.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);

  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return todayISO();
  }

  const base = new Date(Date.UTC(year, month - 1, 1));
  base.setUTCMonth(base.getUTCMonth() + Number(monthOffset || 0));
  const lastDay = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth() + 1, 0)).getUTCDate();
  const safeDay = Math.min(Math.max(day, 1), lastDay);
  return `${base.getUTCFullYear()}-${String(base.getUTCMonth() + 1).padStart(2, "0")}-${String(safeDay).padStart(2, "0")}`;
}

function normalizeRecurringFlag(value) {
  return Number(value) > 0 ? 1 : 0;
}

function normalizeRecurrenceInterval(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 1;
  const whole = Math.floor(numeric);
  return Math.min(120, Math.max(1, whole));
}

function parseIsoDateParts(isoDate) {
  const text = String(isoDate || "");
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  return { year, month, day };
}

function monthDiff(fromIsoDate, toIsoDate) {
  const from = parseIsoDateParts(fromIsoDate);
  const to = parseIsoDateParts(toIsoDate);
  if (!from || !to) return null;
  return (to.year - from.year) * 12 + (to.month - from.month);
}

function startOfMonthIso(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-01`;
}

function ensurePersonalExpenseColumns(db) {
  const columns = db.prepare("PRAGMA table_info(personal_expenses)").all();
  const names = new Set(columns.map((col) => String(col.name)));

  if (!names.has("installment_total")) {
    db.exec("ALTER TABLE personal_expenses ADD COLUMN installment_total INTEGER NOT NULL DEFAULT 1");
  }
  if (!names.has("installment_index")) {
    db.exec("ALTER TABLE personal_expenses ADD COLUMN installment_index INTEGER NOT NULL DEFAULT 1");
  }
  if (!names.has("installment_group_id")) {
    db.exec("ALTER TABLE personal_expenses ADD COLUMN installment_group_id TEXT");
  }
  if (!names.has("is_recurring")) {
    db.exec("ALTER TABLE personal_expenses ADD COLUMN is_recurring INTEGER NOT NULL DEFAULT 0");
  }
  if (!names.has("recurrence_interval_months")) {
    db.exec("ALTER TABLE personal_expenses ADD COLUMN recurrence_interval_months INTEGER NOT NULL DEFAULT 1");
  }

  db.exec(`
    UPDATE personal_expenses
    SET
      installment_total = COALESCE(NULLIF(installment_total, 0), 1),
      installment_index = COALESCE(NULLIF(installment_index, 0), 1),
      is_recurring = CASE
        WHEN COALESCE(is_recurring, 0) > 0 THEN 1
        ELSE 0
      END,
      recurrence_interval_months = CASE
        WHEN COALESCE(recurrence_interval_months, 1) < 1 THEN 1
        ELSE recurrence_interval_months
      END
  `);
}

function ensureBusinessEntryColumns(db) {
  const columns = db.prepare("PRAGMA table_info(business_entries)").all();
  const names = new Set(columns.map((col) => String(col.name)));

  if (!names.has("installment_total")) {
    db.exec("ALTER TABLE business_entries ADD COLUMN installment_total INTEGER NOT NULL DEFAULT 1");
  }
  if (!names.has("installment_index")) {
    db.exec("ALTER TABLE business_entries ADD COLUMN installment_index INTEGER NOT NULL DEFAULT 1");
  }
  if (!names.has("installment_group_id")) {
    db.exec("ALTER TABLE business_entries ADD COLUMN installment_group_id TEXT");
  }
  if (!names.has("is_recurring")) {
    db.exec("ALTER TABLE business_entries ADD COLUMN is_recurring INTEGER NOT NULL DEFAULT 0");
  }
  if (!names.has("recurrence_interval_months")) {
    db.exec("ALTER TABLE business_entries ADD COLUMN recurrence_interval_months INTEGER NOT NULL DEFAULT 1");
  }

  db.exec(`
    UPDATE business_entries
    SET
      installment_total = COALESCE(NULLIF(installment_total, 0), 1),
      installment_index = COALESCE(NULLIF(installment_index, 0), 1),
      is_recurring = CASE
        WHEN COALESCE(is_recurring, 0) > 0 THEN 1
        ELSE 0
      END,
      recurrence_interval_months = CASE
        WHEN COALESCE(recurrence_interval_months, 1) < 1 THEN 1
        ELSE recurrence_interval_months
      END
  `);
}

function tableExists(db, tableName) {
  const row = db.prepare(`
    SELECT 1
    FROM sqlite_master
    WHERE type = 'table' AND name = ?
  `).get(String(tableName || ""));

  return Boolean(row);
}

function normalizeBusinessTypeToken(value) {
  return String(value || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z]/g, "");
}

function looksLikeEntryType(value) {
  const token = normalizeBusinessTypeToken(value);
  return token.includes("entrada");
}

function looksLikeExitType(value) {
  const token = normalizeBusinessTypeToken(value);
  return token.startsWith("s") || token.includes("saida") || token.includes("sada");
}

function getBusinessExitLabel(db) {
  if (db.__businessExitLabel) return db.__businessExitLabel;

  let exitLabel = "Saida";
  const row = db.prepare(`
    SELECT sql
    FROM sqlite_master
    WHERE type = 'table' AND name = 'business_entries'
  `).get();

  const createSql = String(row?.sql || "");
  const inMatch = createSql.match(/entry_type\s+IN\s*\(([^)]+)\)/i);
  if (inMatch) {
    const quoted = String(inMatch[1])
      .match(/'([^']+)'/g)
      ?.map((item) => item.slice(1, -1))
      .filter(Boolean) || [];

    const nonEntry = quoted.find((value) => !looksLikeEntryType(value));
    if (nonEntry && looksLikeExitType(nonEntry)) {
      exitLabel = nonEntry;
    }
  }

  if (exitLabel === "Saida" && tableExists(db, "business_entries")) {
    const distinct = db.prepare(`
      SELECT DISTINCT entry_type
      FROM business_entries
      WHERE entry_type IS NOT NULL AND TRIM(entry_type) <> ''
      ORDER BY id DESC
      LIMIT 20
    `).all();

    const nonEntryFromData = distinct
      .map((rowValue) => String(rowValue.entry_type || ""))
      .find((value) => !looksLikeEntryType(value) && looksLikeExitType(value));

    if (nonEntryFromData) {
      exitLabel = nonEntryFromData;
    }
  }

  db.__businessExitLabel = exitLabel;
  return exitLabel;
}

function normalizeBusinessEntryTypeValue(value, exitLabel) {
  const raw = String(value || "").trim();
  if (!raw) return "Entrada";
  if (looksLikeEntryType(raw)) return "Entrada";
  if (looksLikeExitType(raw)) return String(exitLabel || "Saida");
  return "Entrada";
}

function toDbBusinessEntryType(value, db) {
  return normalizeBusinessEntryTypeValue(value, getBusinessExitLabel(db));
}

function ensureBusinessEntryTypeValues(db) {
  if (!tableExists(db, "business_entries")) return;

  const exitLabel = getBusinessExitLabel(db);
  const rows = db.prepare(`
    SELECT id, entry_type
    FROM business_entries
  `).all();

  if (!rows.length) return;

  const updateStmt = db.prepare(`
    UPDATE business_entries
    SET entry_type = ?, updated_at = ?
    WHERE id = ?
  `);

  const stamp = nowISO();
  const apply = db.transaction(() => {
    for (const row of rows) {
      const normalized = normalizeBusinessEntryTypeValue(row.entry_type, exitLabel);
      if (normalized !== row.entry_type) {
        updateStmt.run(normalized, stamp, row.id);
      }
    }
  });

  try {
    apply();
  } catch (error) {
    console.warn("[DB] Falha ao normalizar entry_type legado:", error.message);
  }
}

function ensureQuoteHistoryColumns(db) {
  if (!tableExists(db, "quote_history")) return;

  const columns = db.prepare("PRAGMA table_info(quote_history)").all();
  const names = new Set(columns.map((col) => String(col.name)));
  const required = ["portfolio_position_id", "ticker", "unit_price", "total_value", "source", "quoted_at"];
  const missingRequired = required.some((column) => !names.has(column));

  if (!missingRequired) {
    db.exec(`
      CREATE INDEX IF NOT EXISTS idx_quote_history_position_date
      ON quote_history(portfolio_position_id, quoted_at DESC)
    `);
    return;
  }

  const rawPositionExpr = names.has("portfolio_position_id")
    ? "portfolio_position_id"
    : names.has("position_id")
      ? "position_id"
      : names.has("portfolio_id")
        ? "portfolio_id"
        : "NULL";

  const safePositionExpr = rawPositionExpr === "NULL"
    ? "NULL"
    : `CASE WHEN ${rawPositionExpr} IN (SELECT id FROM portfolio_positions) THEN ${rawPositionExpr} ELSE NULL END`;

  const selectMap = {
    id: names.has("id") ? "id" : "NULL",
    portfolio_position_id: safePositionExpr,
    ticker: names.has("ticker") ? "ticker" : "NULL",
    unit_price: names.has("unit_price")
      ? "unit_price"
      : names.has("current_unit_price")
        ? "current_unit_price"
        : names.has("price")
          ? "price"
          : "0",
    total_value: names.has("total_value")
      ? "total_value"
      : names.has("current_value")
        ? "current_value"
        : names.has("market_value")
          ? "market_value"
          : names.has("unit_price")
            ? "unit_price"
            : "0",
    source: names.has("source") ? "source" : "'legacy-migrated'",
    quoted_at: names.has("quoted_at")
      ? "quoted_at"
      : names.has("created_at")
        ? "created_at"
        : "CURRENT_TIMESTAMP"
  };

  const migrate = db.transaction(() => {
    db.exec(`
      CREATE TABLE quote_history_new (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        portfolio_position_id INTEGER,
        ticker TEXT,
        unit_price REAL NOT NULL,
        total_value REAL NOT NULL,
        source TEXT NOT NULL DEFAULT 'polling',
        quoted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (portfolio_position_id) REFERENCES portfolio_positions(id) ON DELETE CASCADE
      )
    `);

    db.exec(`
      INSERT INTO quote_history_new (
        id, portfolio_position_id, ticker, unit_price, total_value, source, quoted_at
      )
      SELECT
        ${selectMap.id} AS id,
        ${selectMap.portfolio_position_id} AS portfolio_position_id,
        ${selectMap.ticker} AS ticker,
        ${selectMap.unit_price} AS unit_price,
        ${selectMap.total_value} AS total_value,
        ${selectMap.source} AS source,
        ${selectMap.quoted_at} AS quoted_at
      FROM quote_history
    `);

    db.exec("DROP TABLE quote_history");
    db.exec("ALTER TABLE quote_history_new RENAME TO quote_history");
    db.exec(`
      CREATE INDEX IF NOT EXISTS idx_quote_history_position_date
      ON quote_history(portfolio_position_id, quoted_at DESC)
    `);
  });

  migrate();
}

function calculateCdbProRata(principal, percentCdi, cdiAnnualRate, applicationDate, refDate = new Date()) {
  const amount = Number(principal) || 0;
  const pct = (Number(percentCdi) || 0) / 100;
  const annual = (Number(cdiAnnualRate) || 0) / 100;
  const start = new Date(`${applicationDate}T12:00:00`);

  if (!amount || Number.isNaN(start.getTime())) {
    return amount;
  }

  const elapsedMs = Math.max(0, refDate.getTime() - start.getTime());
  const days = Math.floor(elapsedMs / (24 * 60 * 60 * 1000));
  const effectiveAnnual = annual * pct;
  return amount * Math.pow(1 + effectiveAnnual, days / 365);
}

function initDatabase(dbFilePath) {
  ensureDir(path.dirname(dbFilePath));
  const db = new Database(dbFilePath);

  db.pragma("journal_mode = WAL");
  db.pragma("foreign_keys = ON");

  db.exec(`
    CREATE TABLE IF NOT EXISTS business_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entry_date TEXT NOT NULL,
      description TEXT NOT NULL,
      entry_type TEXT NOT NULL CHECK (entry_type IN ('Entrada', 'Saida')),
      category TEXT NOT NULL,
      amount REAL NOT NULL,
      installment_total INTEGER NOT NULL DEFAULT 1,
      installment_index INTEGER NOT NULL DEFAULT 1,
      installment_group_id TEXT,
      is_recurring INTEGER NOT NULL DEFAULT 0,
      recurrence_interval_months INTEGER NOT NULL DEFAULT 1,
      notes TEXT DEFAULT '',
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS personal_incomes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      income_date TEXT NOT NULL,
      description TEXT NOT NULL,
      source TEXT NOT NULL DEFAULT 'SALARIO',
      amount REAL NOT NULL,
      notes TEXT DEFAULT '',
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS user_settings (
      setting_key TEXT PRIMARY KEY,
      setting_value TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS personal_expenses (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      due_date TEXT NOT NULL,
      description TEXT NOT NULL,
      category TEXT NOT NULL,
      amount REAL NOT NULL,
      status TEXT NOT NULL CHECK (status IN ('Pendente', 'Pago')),
      paid_date TEXT,
      installment_total INTEGER NOT NULL DEFAULT 1,
      installment_index INTEGER NOT NULL DEFAULT 1,
      installment_group_id TEXT,
      is_recurring INTEGER NOT NULL DEFAULT 0,
      recurrence_interval_months INTEGER NOT NULL DEFAULT 1,
      notes TEXT DEFAULT '',
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS portfolio_positions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_label TEXT NOT NULL,
      asset_type TEXT NOT NULL CHECK (asset_type IN ('ACAO', 'FII', 'CDB')),
      ticker TEXT,
      quantity REAL NOT NULL DEFAULT 0,
      avg_price REAL NOT NULL DEFAULT 0,
      application_amount REAL NOT NULL DEFAULT 0,
      rate_percent REAL NOT NULL DEFAULT 0,
      cdi_annual_rate REAL NOT NULL DEFAULT 13.65,
      application_date TEXT,
      current_unit_price REAL NOT NULL DEFAULT 0,
      current_value REAL NOT NULL DEFAULT 0,
      last_dividend REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS quote_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      portfolio_position_id INTEGER,
      ticker TEXT,
      unit_price REAL NOT NULL,
      total_value REAL NOT NULL,
      source TEXT NOT NULL DEFAULT 'polling',
      quoted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (portfolio_position_id) REFERENCES portfolio_positions(id) ON DELETE CASCADE
    );
  `);

  ensureBusinessEntryColumns(db);
  ensurePersonalExpenseColumns(db);
  ensureQuoteHistoryColumns(db);
  ensureBusinessEntryTypeValues(db);
  ensureRecurringRowsForCurrentMonth(db, 1);

  seedDatabase(db);
  return db;
}

function seedDatabase(db) {
  // Mantem apenas configuracao base. Sem seed de dados ficticios.
  seedUserSettings(db);
}

function seedUserSettings(db) {
  const insert = db.prepare(`
    INSERT INTO user_settings (setting_key, setting_value, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(setting_key) DO NOTHING
  `);

  insert.run("meta_patrimonio", "100000", nowISO());
}

function getBusinessEntries(db) {
  return db.prepare(`
    SELECT
      id, entry_date, description, entry_type, category, amount,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes
    FROM business_entries
    ORDER BY entry_date DESC, id DESC
  `).all();
}

function getPersonalIncomes(db) {
  return db.prepare(`
    SELECT id, income_date, description, source, amount, notes
    FROM personal_incomes
    ORDER BY income_date DESC, id DESC
  `).all();
}

function getUserSettings(db) {
  const rows = db.prepare(`
    SELECT setting_key, setting_value
    FROM user_settings
    ORDER BY setting_key ASC
  `).all();

  const out = {};
  for (const row of rows) {
    out[row.setting_key] = row.setting_value;
  }
  return out;
}

function getUserSetting(db, key, fallback = null) {
  const row = db.prepare(`
    SELECT setting_value
    FROM user_settings
    WHERE setting_key = ?
  `).get(key);
  return row ? row.setting_value : fallback;
}

function setUserSetting(db, key, value) {
  db.prepare(`
    INSERT INTO user_settings (setting_key, setting_value, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(setting_key) DO UPDATE SET
      setting_value = excluded.setting_value,
      updated_at = excluded.updated_at
  `).run(String(key), String(value), nowISO());
}

function getPersonalExpenses(db) {
  return db.prepare(`
    SELECT
      id, due_date, description, category, amount, status, paid_date,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes
    FROM personal_expenses
    ORDER BY due_date ASC, id DESC
  `).all();
}

function getPortfolioPositions(db) {
  return db.prepare(`
    SELECT
      id, asset_label, asset_type, ticker, quantity, avg_price,
      application_amount, rate_percent, cdi_annual_rate, application_date,
      current_unit_price, current_value, last_dividend, updated_at
    FROM portfolio_positions
    ORDER BY
      CASE asset_type WHEN 'CDB' THEN 2 ELSE 1 END,
      asset_label ASC
  `).all();
}

function getQuoteTargets(db) {
  return db.prepare(`
    SELECT id, ticker, quantity, current_unit_price, current_value
    FROM portfolio_positions
    WHERE asset_type IN ('ACAO', 'FII') AND ticker IS NOT NULL AND ticker <> ''
  `).all();
}

function getCdbTargets(db) {
  return db.prepare(`
    SELECT id, application_amount, rate_percent, cdi_annual_rate, application_date, current_value
    FROM portfolio_positions
    WHERE asset_type = 'CDB'
  `).all();
}

function normalizeReferenceIsoDate(value) {
  const parsed = parseIsoDateParts(value);
  if (!parsed) return todayISO();
  return `${String(parsed.year).padStart(4, "0")}-${String(parsed.month).padStart(2, "0")}-${String(parsed.day).padStart(2, "0")}`;
}

function ensureRecurringBusinessEntries(db, referenceDate = todayISO(), monthsAhead = 1) {
  const normalizedReference = normalizeReferenceIsoDate(referenceDate);
  const horizon = Math.max(0, Math.min(12, Math.floor(Number(monthsAhead) || 0)));

  const baseRows = db.prepare(`
    SELECT
      id, entry_date, description, entry_type, category, amount,
      notes, recurrence_interval_months, is_recurring
    FROM business_entries
    WHERE is_recurring = 1 AND installment_index = 1
  `).all();

  if (!baseRows.length) return { created: 0 };

  const monthStart = startOfMonthIso(new Date(`${normalizedReference}T12:00:00`));
  const stamp = nowISO();
  let created = 0;

  const existsStmt = db.prepare(`
    SELECT id
    FROM business_entries
    WHERE
      entry_date = @entry_date AND
      description = @description AND
      entry_type = @entry_type AND
      category = @category AND
      ABS(amount - @amount) < 0.00001 AND
      installment_total = 1 AND
      installment_index = 1
    LIMIT 1
  `);

  const insertStmt = db.prepare(`
    INSERT INTO business_entries (
      entry_date, description, entry_type, category, amount,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes, created_at, updated_at
    ) VALUES (
      @entry_date, @description, @entry_type, @category, @amount,
      1, 1, NULL,
      @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
    )
  `);

  const run = db.transaction(() => {
    for (const row of baseRows) {
      const interval = normalizeRecurrenceInterval(row.recurrence_interval_months);
      for (let monthOffset = 0; monthOffset <= horizon; monthOffset += 1) {
        const targetMonth = addMonthsIso(monthStart, monthOffset);
        const diff = monthDiff(row.entry_date, targetMonth);
        if (!Number.isFinite(diff) || diff < 0) continue;
        if (diff % interval !== 0) continue;

        const scheduledDate = addMonthsIso(row.entry_date, diff);
        const params = {
          entry_date: scheduledDate,
          description: row.description || "",
          entry_type: row.entry_type || "Entrada",
          category: row.category || "Outros",
          amount: Number(row.amount) || 0,
          is_recurring: 1,
          recurrence_interval_months: interval,
          notes: row.notes || "",
          created_at: stamp,
          updated_at: stamp
        };

        const exists = existsStmt.get(params);
        if (exists) continue;

        insertStmt.run(params);
        created += 1;
      }
    }
  });

  run();
  return { created };
}

function ensureRecurringPersonalExpenses(db, referenceDate = todayISO(), monthsAhead = 1) {
  const normalizedReference = normalizeReferenceIsoDate(referenceDate);
  const horizon = Math.max(0, Math.min(12, Math.floor(Number(monthsAhead) || 0)));

  const baseRows = db.prepare(`
    SELECT
      id, due_date, description, category, amount,
      notes, recurrence_interval_months, is_recurring
    FROM personal_expenses
    WHERE is_recurring = 1 AND installment_index = 1
  `).all();

  if (!baseRows.length) return { created: 0 };

  const monthStart = startOfMonthIso(new Date(`${normalizedReference}T12:00:00`));
  const stamp = nowISO();
  let created = 0;

  const existsStmt = db.prepare(`
    SELECT id
    FROM personal_expenses
    WHERE
      due_date = @due_date AND
      description = @description AND
      category = @category AND
      ABS(amount - @amount) < 0.00001 AND
      installment_total = 1 AND
      installment_index = 1
    LIMIT 1
  `);

  const insertStmt = db.prepare(`
    INSERT INTO personal_expenses (
      due_date, description, category, amount, status, paid_date,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes, created_at, updated_at
    ) VALUES (
      @due_date, @description, @category, @amount, 'Pendente', NULL,
      1, 1, NULL,
      @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
    )
  `);

  const run = db.transaction(() => {
    for (const row of baseRows) {
      const interval = normalizeRecurrenceInterval(row.recurrence_interval_months);
      for (let monthOffset = 0; monthOffset <= horizon; monthOffset += 1) {
        const targetMonth = addMonthsIso(monthStart, monthOffset);
        const diff = monthDiff(row.due_date, targetMonth);
        if (!Number.isFinite(diff) || diff < 0) continue;
        if (diff % interval !== 0) continue;

        const scheduledDate = addMonthsIso(row.due_date, diff);
        const params = {
          due_date: scheduledDate,
          description: row.description || "",
          category: row.category || "Outros",
          amount: Number(row.amount) || 0,
          is_recurring: 1,
          recurrence_interval_months: interval,
          notes: row.notes || "",
          created_at: stamp,
          updated_at: stamp
        };

        const exists = existsStmt.get(params);
        if (exists) continue;

        insertStmt.run(params);
        created += 1;
      }
    }
  });

  run();
  return { created };
}

function ensureRecurringRowsForCurrentMonth(db, monthsAhead = 1) {
  const today = todayISO();
  const business = ensureRecurringBusinessEntries(db, today, monthsAhead);
  const personal = ensureRecurringPersonalExpenses(db, today, monthsAhead);

  return {
    businessCreated: Number(business.created || 0),
    personalCreated: Number(personal.created || 0),
    totalCreated: Number(business.created || 0) + Number(personal.created || 0)
  };
}

function upsertBusinessEntry(db, row) {
  const stamp = nowISO();
  const installmentTotal = normalizeInstallmentTotal(row.installment_total);
  const installmentIndex = Math.min(
    installmentTotal,
    Math.max(1, Number.isFinite(Number(row.installment_index)) ? Math.floor(Number(row.installment_index)) : 1)
  );

  const payload = {
    id: row.id ? Number(row.id) : null,
    entry_date: row.entry_date,
    description: row.description || "",
    entry_type: toDbBusinessEntryType(row.entry_type, db),
    category: row.category || "Outros",
    amount: Number(row.amount) || 0,
    installment_total: installmentTotal,
    installment_index: installmentIndex,
    installment_group_id: row.installment_group_id || null,
    is_recurring: installmentTotal > 1 ? 0 : normalizeRecurringFlag(row.is_recurring),
    recurrence_interval_months: installmentTotal > 1
      ? 1
      : normalizeRecurrenceInterval(row.recurrence_interval_months),
    notes: row.notes || ""
  };

  if (payload.id) {
    const current = db.prepare(`
      SELECT
        id, entry_date, description, entry_type, category, amount,
        installment_total, installment_index, installment_group_id,
        is_recurring, recurrence_interval_months, notes
      FROM business_entries
      WHERE id = ?
    `).get(payload.id);

    if (!current) return null;

    const canExpandInstallments =
      payload.installment_total > 1 &&
      Number(current.installment_total || 1) <= 1 &&
      Number(current.installment_index || 1) <= 1;

    if (canExpandInstallments) {
      const groupId = payload.installment_group_id || current.installment_group_id || `inst-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
      const generatedIds = [];

      const updateAndInsert = db.transaction(() => {
        db.prepare(`
          UPDATE business_entries
          SET
            entry_date = @entry_date,
            description = @description,
            entry_type = @entry_type,
            category = @category,
            amount = @amount,
            installment_total = @installment_total,
            installment_index = 1,
            installment_group_id = @installment_group_id,
            is_recurring = @is_recurring,
            recurrence_interval_months = @recurrence_interval_months,
            notes = @notes,
            updated_at = @updated_at
          WHERE id = @id
        `).run({
          ...payload,
          installment_index: 1,
          installment_group_id: groupId,
          is_recurring: 0,
          recurrence_interval_months: 1,
          updated_at: stamp
        });

        const insertStmt = db.prepare(`
          INSERT INTO business_entries (
            entry_date, description, entry_type, category, amount,
            installment_total, installment_index, installment_group_id,
            is_recurring, recurrence_interval_months, notes, created_at, updated_at
          ) VALUES (
            @entry_date, @description, @entry_type, @category, @amount,
            @installment_total, @installment_index, @installment_group_id,
            @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
          )
        `);

        for (let index = 2; index <= payload.installment_total; index += 1) {
          const installmentDate = addMonthsIso(payload.entry_date, index - 1);
          const inserted = insertStmt.run({
            entry_date: installmentDate,
            description: payload.description,
            entry_type: payload.entry_type,
            category: payload.category,
            amount: payload.amount,
            installment_total: payload.installment_total,
            installment_index: index,
            installment_group_id: groupId,
            is_recurring: 0,
            recurrence_interval_months: 1,
            notes: payload.notes,
            created_at: stamp,
            updated_at: stamp
          });
          generatedIds.push(Number(inserted.lastInsertRowid));
        }
      });

      updateAndInsert();

      const allIds = [Number(payload.id), ...generatedIds];
      const rows = db.prepare(`
        SELECT
          id, entry_date, description, entry_type, category, amount,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes
        FROM business_entries
        WHERE id IN (${allIds.map(() => "?").join(",")})
        ORDER BY entry_date ASC, installment_index ASC, id ASC
      `).all(...allIds);

      const primary = rows.find((item) => Number(item.id) === Number(payload.id)) || rows[0] || null;
      if (!primary) return null;

      return {
        ...primary,
        generated_rows: rows
      };
    }

    db.prepare(`
      UPDATE business_entries
      SET
        entry_date = @entry_date,
        description = @description,
        entry_type = @entry_type,
        category = @category,
        amount = @amount,
        installment_total = @installment_total,
        installment_index = @installment_index,
        installment_group_id = @installment_group_id,
        is_recurring = @is_recurring,
        recurrence_interval_months = @recurrence_interval_months,
        notes = @notes,
        updated_at = @updated_at
      WHERE id = @id
    `).run({ ...payload, updated_at: stamp });
  } else {
    if (payload.installment_total > 1) {
      const groupId = payload.installment_group_id || `inst-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
      const insertStmt = db.prepare(`
        INSERT INTO business_entries (
          entry_date, description, entry_type, category, amount,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes, created_at, updated_at
        ) VALUES (
          @entry_date, @description, @entry_type, @category, @amount,
          @installment_total, @installment_index, @installment_group_id,
          @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
        )
      `);

      const generatedIds = [];
      const insertMany = db.transaction(() => {
        for (let index = 1; index <= payload.installment_total; index += 1) {
          const installmentDate = addMonthsIso(payload.entry_date, index - 1);
          const inserted = insertStmt.run({
            entry_date: installmentDate,
            description: payload.description,
            entry_type: payload.entry_type,
            category: payload.category,
            amount: payload.amount,
            installment_total: payload.installment_total,
            installment_index: index,
            installment_group_id: groupId,
            is_recurring: 0,
            recurrence_interval_months: 1,
            notes: payload.notes,
            created_at: stamp,
            updated_at: stamp
          });
          generatedIds.push(Number(inserted.lastInsertRowid));
        }
      });

      insertMany();

      const rows = db.prepare(`
        SELECT
          id, entry_date, description, entry_type, category, amount,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes
        FROM business_entries
        WHERE id IN (${generatedIds.map(() => "?").join(",")})
        ORDER BY entry_date ASC, installment_index ASC, id ASC
      `).all(...generatedIds);

      const primary = rows[0] || null;
      if (!primary) return null;

      return {
        ...primary,
        generated_rows: rows
      };
    }

    const inserted = db.prepare(`
      INSERT INTO business_entries (
        entry_date, description, entry_type, category, amount,
        installment_total, installment_index, installment_group_id,
        is_recurring, recurrence_interval_months, notes, created_at, updated_at
      ) VALUES (
        @entry_date, @description, @entry_type, @category, @amount,
        @installment_total, @installment_index, @installment_group_id,
        @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
      )
    `).run({
      ...payload,
      installment_group_id: payload.installment_group_id || null,
      created_at: stamp,
      updated_at: stamp
    });
    payload.id = inserted.lastInsertRowid;
  }

  return db.prepare(`
    SELECT
      id, entry_date, description, entry_type, category, amount,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes
    FROM business_entries
    WHERE id = ?
  `).get(payload.id);
}

function addPersonalIncome(db, row) {
  const stamp = nowISO();
  const payload = {
    income_date: row.income_date || todayISO(),
    description: row.description || "Salario/Receita",
    source: row.source || "SALARIO",
    amount: Number(row.amount) || 0,
    notes: row.notes || ""
  };

  const inserted = db.prepare(`
    INSERT INTO personal_incomes (
      income_date, description, source, amount, notes, created_at, updated_at
    ) VALUES (
      @income_date, @description, @source, @amount, @notes, @created_at, @updated_at
    )
  `).run({
    ...payload,
    created_at: stamp,
    updated_at: stamp
  });

  return db.prepare(`
    SELECT id, income_date, description, source, amount, notes
    FROM personal_incomes
    WHERE id = ?
  `).get(inserted.lastInsertRowid);
}

function getPendingExpensesDueBetween(db, startIso, endIso) {
  return db.prepare(`
    SELECT id, due_date, description, category, amount, status
    FROM personal_expenses
    WHERE status = 'Pendente'
      AND due_date >= ?
      AND due_date <= ?
    ORDER BY due_date ASC
  `).all(startIso, endIso);
}

function upsertPersonalExpense(db, row) {
  const stamp = nowISO();
  const status = row.status === "Pago" || row.status === "PAGO" ? "Pago" : "Pendente";
  const dueDate = row.due_date || todayISO();
  const installmentTotal = normalizeInstallmentTotal(row.installment_total);
  const installmentIndex = Math.min(
    installmentTotal,
    Math.max(1, Number.isFinite(Number(row.installment_index)) ? Math.floor(Number(row.installment_index)) : 1)
  );

  const payload = {
    id: row.id ? Number(row.id) : null,
    due_date: dueDate,
    description: row.description || "",
    category: row.category || "Outros",
    amount: Number(row.amount) || 0,
    status,
    paid_date: status === "Pago" ? (row.paid_date || dueDate) : null,
    installment_total: installmentTotal,
    installment_index: installmentIndex,
    installment_group_id: row.installment_group_id || null,
    is_recurring: installmentTotal > 1 ? 0 : normalizeRecurringFlag(row.is_recurring),
    recurrence_interval_months: installmentTotal > 1
      ? 1
      : normalizeRecurrenceInterval(row.recurrence_interval_months),
    notes: row.notes || ""
  };

  if (payload.id) {
    const current = db.prepare(`
      SELECT
        id, due_date, description, category, amount, status, paid_date,
        installment_total, installment_index, installment_group_id,
        is_recurring, recurrence_interval_months, notes
      FROM personal_expenses
      WHERE id = ?
    `).get(payload.id);

    if (!current) return null;

    const canExpandInstallments =
      payload.installment_total > 1 &&
      Number(current.installment_total || 1) <= 1 &&
      Number(current.installment_index || 1) <= 1;

    if (canExpandInstallments) {
      const groupId = payload.installment_group_id || current.installment_group_id || `inst-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
      const generatedIds = [];

      const updateAndInsert = db.transaction(() => {
        db.prepare(`
          UPDATE personal_expenses
          SET
            due_date = @due_date,
            description = @description,
            category = @category,
            amount = @amount,
            status = @status,
            paid_date = @paid_date,
            installment_total = @installment_total,
            installment_index = 1,
            installment_group_id = @installment_group_id,
            is_recurring = @is_recurring,
            recurrence_interval_months = @recurrence_interval_months,
            notes = @notes,
            updated_at = @updated_at
          WHERE id = @id
        `).run({
          ...payload,
          installment_index: 1,
          installment_group_id: groupId,
          is_recurring: 0,
          recurrence_interval_months: 1,
          updated_at: stamp
        });

        const insertStmt = db.prepare(`
          INSERT INTO personal_expenses (
            due_date, description, category, amount, status, paid_date,
            installment_total, installment_index, installment_group_id,
            is_recurring, recurrence_interval_months, notes, created_at, updated_at
          ) VALUES (
            @due_date, @description, @category, @amount, @status, @paid_date,
            @installment_total, @installment_index, @installment_group_id,
            @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
          )
        `);

        for (let index = 2; index <= payload.installment_total; index += 1) {
          const installmentDueDate = addMonthsIso(payload.due_date, index - 1);
          const inserted = insertStmt.run({
            due_date: installmentDueDate,
            description: payload.description,
            category: payload.category,
            amount: payload.amount,
            status: "Pendente",
            paid_date: null,
            installment_total: payload.installment_total,
            installment_index: index,
            installment_group_id: groupId,
            is_recurring: 0,
            recurrence_interval_months: 1,
            notes: payload.notes,
            created_at: stamp,
            updated_at: stamp
          });
          generatedIds.push(Number(inserted.lastInsertRowid));
        }
      });

      updateAndInsert();

      const allIds = [Number(payload.id), ...generatedIds];
      const rows = db.prepare(`
        SELECT
          id, due_date, description, category, amount, status, paid_date,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes
        FROM personal_expenses
        WHERE id IN (${allIds.map(() => "?").join(",")})
        ORDER BY due_date ASC, installment_index ASC, id ASC
      `).all(...allIds);

      const primary = rows.find((item) => Number(item.id) === Number(payload.id)) || rows[0] || null;
      if (!primary) return null;

      return {
        ...primary,
        generated_rows: rows
      };
    }

    db.prepare(`
      UPDATE personal_expenses
      SET
        due_date = @due_date,
        description = @description,
        category = @category,
        amount = @amount,
        status = @status,
        paid_date = @paid_date,
        installment_total = @installment_total,
        installment_index = @installment_index,
        installment_group_id = @installment_group_id,
        is_recurring = @is_recurring,
        recurrence_interval_months = @recurrence_interval_months,
        notes = @notes,
        updated_at = @updated_at
      WHERE id = @id
    `).run({ ...payload, updated_at: stamp });
  } else {
    if (payload.installment_total > 1) {
      const groupId = payload.installment_group_id || `inst-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
      const insertStmt = db.prepare(`
        INSERT INTO personal_expenses (
          due_date, description, category, amount, status, paid_date,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes, created_at, updated_at
        ) VALUES (
          @due_date, @description, @category, @amount, @status, @paid_date,
          @installment_total, @installment_index, @installment_group_id,
          @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
        )
      `);

      const generatedIds = [];
      const insertMany = db.transaction(() => {
        for (let index = 1; index <= payload.installment_total; index += 1) {
          const installmentDueDate = addMonthsIso(payload.due_date, index - 1);
          const installmentStatus = index === 1 ? payload.status : "Pendente";
          const installmentPaidDate = installmentStatus === "Pago"
            ? (index === 1 ? payload.paid_date : installmentDueDate)
            : null;

          const inserted = insertStmt.run({
            due_date: installmentDueDate,
            description: payload.description,
            category: payload.category,
            amount: payload.amount,
            status: installmentStatus,
            paid_date: installmentPaidDate,
            installment_total: payload.installment_total,
            installment_index: index,
            installment_group_id: groupId,
            is_recurring: 0,
            recurrence_interval_months: 1,
            notes: payload.notes,
            created_at: stamp,
            updated_at: stamp
          });
          generatedIds.push(Number(inserted.lastInsertRowid));
        }
      });

      insertMany();

      const rows = db.prepare(`
        SELECT
          id, due_date, description, category, amount, status, paid_date,
          installment_total, installment_index, installment_group_id,
          is_recurring, recurrence_interval_months, notes
        FROM personal_expenses
        WHERE id IN (${generatedIds.map(() => "?").join(",")})
        ORDER BY due_date ASC, installment_index ASC, id ASC
      `).all(...generatedIds);

      const primary = rows[0] || null;
      if (!primary) {
        return null;
      }

      return {
        ...primary,
        generated_rows: rows
      };
    }

    const inserted = db.prepare(`
      INSERT INTO personal_expenses (
        due_date, description, category, amount, status, paid_date,
        installment_total, installment_index, installment_group_id,
        is_recurring, recurrence_interval_months, notes, created_at, updated_at
      ) VALUES (
        @due_date, @description, @category, @amount, @status, @paid_date,
        @installment_total, @installment_index, @installment_group_id,
        @is_recurring, @recurrence_interval_months, @notes, @created_at, @updated_at
      )
    `).run({
      ...payload,
      installment_group_id: payload.installment_group_id || null,
      created_at: stamp,
      updated_at: stamp
    });
    payload.id = inserted.lastInsertRowid;
  }

  return db.prepare(`
    SELECT
      id, due_date, description, category, amount, status, paid_date,
      installment_total, installment_index, installment_group_id,
      is_recurring, recurrence_interval_months, notes
    FROM personal_expenses
    WHERE id = ?
  `).get(payload.id);
}

function upsertPortfolioPosition(db, row) {
  const stamp = nowISO();
  const type = row.asset_type === "CDB" ? "CDB" : row.asset_type === "FII" ? "FII" : "ACAO";
  const fallbackApplicationDate = String(row.application_date || row.updated_at || todayISO()).slice(0, 10);

  const payload = {
    id: row.id ? Number(row.id) : null,
    asset_label: row.asset_label || "",
    asset_type: type,
    ticker: row.ticker ? String(row.ticker).toUpperCase() : null,
    quantity: type === "CDB" ? 0 : Number(row.quantity) || 0,
    avg_price: type === "CDB" ? 0 : Number(row.avg_price) || 0,
    application_amount: type === "CDB" ? Number(row.application_amount) || 0 : 0,
    rate_percent: type === "CDB" ? Number(row.rate_percent) || 0 : 0,
    cdi_annual_rate: type === "CDB" ? Number(row.cdi_annual_rate) || 13.65 : 13.65,
    application_date: fallbackApplicationDate || todayISO(),
    current_unit_price: type === "CDB" ? 0 : Number(row.current_unit_price) || 0,
    current_value: Number(row.current_value) || 0,
    last_dividend: type === "CDB" ? 0 : Number(row.last_dividend) || 0
  };

  if (type === "CDB") {
    payload.current_value = calculateCdbProRata(
      payload.application_amount,
      payload.rate_percent,
      payload.cdi_annual_rate,
      payload.application_date
    );
  } else {
    payload.current_value = payload.current_unit_price;
  }

  if (payload.id) {
    db.prepare(`
      UPDATE portfolio_positions
      SET
        asset_label = @asset_label,
        asset_type = @asset_type,
        ticker = @ticker,
        quantity = @quantity,
        avg_price = @avg_price,
        application_amount = @application_amount,
        rate_percent = @rate_percent,
        cdi_annual_rate = @cdi_annual_rate,
        application_date = @application_date,
        current_unit_price = @current_unit_price,
        current_value = @current_value,
        last_dividend = @last_dividend,
        updated_at = @updated_at
      WHERE id = @id
    `).run({ ...payload, updated_at: stamp });
  } else {
    const inserted = db.prepare(`
      INSERT INTO portfolio_positions (
        asset_label, asset_type, ticker, quantity, avg_price,
        application_amount, rate_percent, cdi_annual_rate, application_date,
        current_unit_price, current_value, last_dividend, updated_at
      ) VALUES (
        @asset_label, @asset_type, @ticker, @quantity, @avg_price,
        @application_amount, @rate_percent, @cdi_annual_rate, @application_date,
        @current_unit_price, @current_value, @last_dividend, @updated_at
      )
    `).run({ ...payload, updated_at: stamp });
    payload.id = inserted.lastInsertRowid;
  }

  return db.prepare(`
    SELECT
      id, asset_label, asset_type, ticker, quantity, avg_price,
      application_amount, rate_percent, cdi_annual_rate, application_date,
      current_unit_price, current_value, last_dividend, updated_at
    FROM portfolio_positions
    WHERE id = ?
  `).get(payload.id);
}

function deleteBusinessEntry(db, id) {
  const parsed = Number(id);
  if (!Number.isFinite(parsed) || parsed <= 0) return false;
  const result = db.prepare("DELETE FROM business_entries WHERE id = ?").run(parsed);
  return result.changes > 0;
}

function deletePersonalIncome(db, id) {
  const parsed = Number(id);
  if (!Number.isFinite(parsed) || parsed <= 0) return false;
  const result = db.prepare("DELETE FROM personal_incomes WHERE id = ?").run(parsed);
  return result.changes > 0;
}

function deletePersonalExpense(db, id) {
  const parsed = Number(id);
  if (!Number.isFinite(parsed) || parsed <= 0) return false;
  const result = db.prepare("DELETE FROM personal_expenses WHERE id = ?").run(parsed);
  return result.changes > 0;
}

function deletePortfolioPosition(db, id) {
  const parsed = Number(id);
  if (!Number.isFinite(parsed) || parsed <= 0) return false;
  const result = db.prepare("DELETE FROM portfolio_positions WHERE id = ?").run(parsed);
  return result.changes > 0;
}

function updatePortfolioLiveValues(db, update) {
  const payload = {
    id: Number(update.id),
    current_unit_price: Number(update.current_unit_price) || 0,
    current_value: Number(update.current_value) || 0,
    total_market_value: Number(update.total_market_value) || Number(update.current_value) || 0,
    quoted_at: update.quoted_at || nowISO(),
    source: update.source || "polling",
    ticker: update.ticker || null
  };

  db.prepare(`
    UPDATE portfolio_positions
    SET
      current_unit_price = @current_unit_price,
      current_value = @current_value,
      updated_at = @quoted_at
    WHERE id = @id
  `).run(payload);

  db.prepare(`
    INSERT INTO quote_history (
      portfolio_position_id, ticker, unit_price, total_value, source, quoted_at
    ) VALUES (
      @id, @ticker, @current_unit_price, @total_market_value, @source, @quoted_at
    )
  `).run(payload);
}

module.exports = {
  initDatabase,
  calculateCdbProRata,
  getBusinessEntries,
  getPersonalIncomes,
  getPersonalExpenses,
  getPendingExpensesDueBetween,
  getPortfolioPositions,
  getQuoteTargets,
  getCdbTargets,
  getUserSettings,
  getUserSetting,
  setUserSetting,
  upsertBusinessEntry,
  addPersonalIncome,
  deletePersonalIncome,
  upsertPersonalExpense,
  upsertPortfolioPosition,
  deleteBusinessEntry,
  deletePersonalExpense,
  deletePortfolioPosition,
  updatePortfolioLiveValues,
  ensureRecurringRowsForCurrentMonth
};
