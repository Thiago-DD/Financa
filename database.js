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
      entry_type TEXT NOT NULL CHECK (entry_type IN ('Entrada', 'Saída', 'Saida')),
      category TEXT NOT NULL,
      amount REAL NOT NULL,
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
      portfolio_position_id INTEGER NOT NULL,
      ticker TEXT,
      unit_price REAL NOT NULL,
      total_value REAL NOT NULL,
      source TEXT NOT NULL DEFAULT 'polling',
      quoted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (portfolio_position_id) REFERENCES portfolio_positions(id) ON DELETE CASCADE
    );
  `);

  seedDatabase(db);
  return db;
}

function seedDatabase(db) {
  seedBusiness(db);
  seedPersonalIncomes(db);
  seedUserSettings(db);
  seedPersonalExpenses(db);
  seedPortfolio(db);
}

function seedUserSettings(db) {
  const insert = db.prepare(`
    INSERT INTO user_settings (setting_key, setting_value, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(setting_key) DO NOTHING
  `);

  insert.run("meta_patrimonio", "100000", nowISO());
}

function seedBusiness(db) {
  const count = db.prepare("SELECT COUNT(*) AS total FROM business_entries").get().total;
  if (count > 0) return;

  const insert = db.prepare(`
    INSERT INTO business_entries (
      entry_date, description, entry_type, category, amount, notes, created_at, updated_at
    ) VALUES (
      @entry_date, @description, @entry_type, @category, @amount, @notes, @created_at, @updated_at
    )
  `);

  const stamp = nowISO();
  const rows = [
    {
      entry_date: "2026-04-29",
      description: "Vendas balcao - cartoes",
      entry_type: "Entrada",
      category: "Venda Balcao",
      amount: 3124.8,
      notes: ""
    },
    {
      entry_date: "2026-04-29",
      description: "Vendas balcao - PIX",
      entry_type: "Entrada",
      category: "Venda Balcao",
      amount: 1460.5,
      notes: ""
    },
    {
      entry_date: "2026-04-29",
      description: "Sangria para cofre",
      entry_type: "Saída",
      category: "Sangria",
      amount: 1800,
      notes: ""
    },
    {
      entry_date: "2026-04-29",
      description: "Fornecedor hortifruti",
      entry_type: "Saída",
      category: "Fornecedor",
      amount: 890.7,
      notes: ""
    }
  ];

  const tx = db.transaction((payload) => {
    for (const row of payload) {
      insert.run({ ...row, created_at: stamp, updated_at: stamp });
    }
  });
  tx(rows);
}

function seedPersonalIncomes(db) {
  const count = db.prepare("SELECT COUNT(*) AS total FROM personal_incomes").get().total;
  if (count > 0) return;

  const insert = db.prepare(`
    INSERT INTO personal_incomes (
      income_date, description, source, amount, notes, created_at, updated_at
    ) VALUES (
      @income_date, @description, @source, @amount, @notes, @created_at, @updated_at
    )
  `);

  const stamp = nowISO();
  const rows = [
    {
      income_date: "2026-04-05",
      description: "Salario principal",
      source: "SALARIO",
      amount: 5500,
      notes: ""
    },
    {
      income_date: "2026-04-20",
      description: "Renda extra",
      source: "RENDA_EXTRA",
      amount: 900,
      notes: "Freelance"
    }
  ];

  const tx = db.transaction((payload) => {
    for (const row of payload) {
      insert.run({ ...row, created_at: stamp, updated_at: stamp });
    }
  });
  tx(rows);
}

function seedPersonalExpenses(db) {
  const count = db.prepare("SELECT COUNT(*) AS total FROM personal_expenses").get().total;
  if (count > 0) return;

  const insert = db.prepare(`
    INSERT INTO personal_expenses (
      due_date, description, category, amount, status, paid_date, notes, created_at, updated_at
    ) VALUES (
      @due_date, @description, @category, @amount, @status, @paid_date, @notes, @created_at, @updated_at
    )
  `);

  const stamp = nowISO();
  const rows = [
    {
      due_date: "2026-04-15",
      description: "Conta de Luz",
      category: "Moradia",
      amount: 278.45,
      status: "Pago",
      paid_date: "2026-04-14",
      notes: ""
    },
    {
      due_date: "2026-04-18",
      description: "Internet",
      category: "Moradia",
      amount: 129.9,
      status: "Pendente",
      paid_date: null,
      notes: ""
    },
    {
      due_date: "2026-04-22",
      description: "Mercado",
      category: "Alimentacao",
      amount: 530.2,
      status: "Pago",
      paid_date: "2026-04-22",
      notes: ""
    },
    {
      due_date: "2026-04-26",
      description: "Combustivel",
      category: "Transporte",
      amount: 250,
      status: "Pendente",
      paid_date: null,
      notes: ""
    },
    {
      due_date: "2026-03-28",
      description: "Academia",
      category: "Lazer",
      amount: 89.9,
      status: "Pendente",
      paid_date: null,
      notes: "Exemplo de vencida"
    }
  ];

  const tx = db.transaction((payload) => {
    for (const row of payload) {
      insert.run({ ...row, created_at: stamp, updated_at: stamp });
    }
  });
  tx(rows);
}

function seedPortfolio(db) {
  const count = db.prepare("SELECT COUNT(*) AS total FROM portfolio_positions").get().total;
  if (count > 0) return;

  const insert = db.prepare(`
    INSERT INTO portfolio_positions (
      asset_label, asset_type, ticker, quantity, avg_price,
      application_amount, rate_percent, cdi_annual_rate, application_date,
      current_unit_price, current_value, last_dividend, updated_at
    ) VALUES (
      @asset_label, @asset_type, @ticker, @quantity, @avg_price,
      @application_amount, @rate_percent, @cdi_annual_rate, @application_date,
      @current_unit_price, @current_value, @last_dividend, @updated_at
    )
  `);

  const stamp = nowISO();

  const stockRows = [
    { asset_label: "MXRF11", asset_type: "FII", ticker: "MXRF11", quantity: 820, avg_price: 9.8, current_unit_price: 9.8, last_dividend: 0.1 },
    { asset_label: "BTLG11", asset_type: "FII", ticker: "BTLG11", quantity: 42, avg_price: 98.75, current_unit_price: 101.3, last_dividend: 0.76 },
    { asset_label: "BBAS3", asset_type: "ACAO", ticker: "BBAS3", quantity: 110, avg_price: 26.4, current_unit_price: 29.2, last_dividend: 0.44 },
    { asset_label: "TAEE4", asset_type: "ACAO", ticker: "TAEE4", quantity: 95, avg_price: 34.1, current_unit_price: 35.45, last_dividend: 0.69 },
    { asset_label: "CXSE3", asset_type: "ACAO", ticker: "CXSE3", quantity: 210, avg_price: 12.8, current_unit_price: 13.05, last_dividend: 0.08 },
    { asset_label: "CMIG4", asset_type: "ACAO", ticker: "CMIG4", quantity: 160, avg_price: 11.75, current_unit_price: 12.31, last_dividend: 0.17 }
  ].map((row) => ({
    ...row,
    application_amount: 0,
    rate_percent: 0,
    cdi_annual_rate: 13.65,
    application_date: null,
    current_value: row.current_unit_price,
    updated_at: stamp
  }));

  const cdbPrincipal = 12000;
  const cdbRate = 110;
  const cdbAnnual = 13.65;
  const cdbDate = "2026-02-10";

  const cdbRow = {
    asset_label: "CDB Banco Inter",
    asset_type: "CDB",
    ticker: null,
    quantity: 0,
    avg_price: 0,
    application_amount: cdbPrincipal,
    rate_percent: cdbRate,
    cdi_annual_rate: cdbAnnual,
    application_date: cdbDate,
    current_unit_price: 0,
    current_value: calculateCdbProRata(cdbPrincipal, cdbRate, cdbAnnual, cdbDate),
    last_dividend: 0,
    updated_at: stamp
  };

  const tx = db.transaction((payload) => {
    for (const row of payload) {
      insert.run(row);
    }
  });
  tx([...stockRows, cdbRow]);
}

function getBusinessEntries(db) {
  return db.prepare(`
    SELECT id, entry_date, description, entry_type, category, amount, notes
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
    SELECT id, due_date, description, category, amount, status, paid_date, notes
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

function upsertBusinessEntry(db, row) {
  const stamp = nowISO();
  const payload = {
    id: row.id ? Number(row.id) : null,
    entry_date: row.entry_date,
    description: row.description || "",
    entry_type:
      row.entry_type === "Saída" || row.entry_type === "Saida" || row.entry_type === "SAIDA"
        ? "Saída"
        : "Entrada",
    category: row.category || "Outros",
    amount: Number(row.amount) || 0,
    notes: row.notes || ""
  };

  if (payload.id) {
    db.prepare(`
      UPDATE business_entries
      SET
        entry_date = @entry_date,
        description = @description,
        entry_type = @entry_type,
        category = @category,
        amount = @amount,
        notes = @notes,
        updated_at = @updated_at
      WHERE id = @id
    `).run({ ...payload, updated_at: stamp });
  } else {
    const inserted = db.prepare(`
      INSERT INTO business_entries (
        entry_date, description, entry_type, category, amount, notes, created_at, updated_at
      ) VALUES (
        @entry_date, @description, @entry_type, @category, @amount, @notes, @created_at, @updated_at
      )
    `).run({ ...payload, created_at: stamp, updated_at: stamp });
    payload.id = inserted.lastInsertRowid;
  }

  return db.prepare(`
    SELECT id, entry_date, description, entry_type, category, amount, notes
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
  const status =
    row.status === "Pago" || row.status === "PAGO"
      ? "Pago"
      : "Pendente";
  const dueDate = row.due_date || todayISO();

  const payload = {
    id: row.id ? Number(row.id) : null,
    due_date: dueDate,
    description: row.description || "",
    category: row.category || "Outros",
    amount: Number(row.amount) || 0,
    status,
    paid_date: status === "Pago" ? (row.paid_date || dueDate) : null,
    notes: row.notes || ""
  };

  if (payload.id) {
    db.prepare(`
      UPDATE personal_expenses
      SET
        due_date = @due_date,
        description = @description,
        category = @category,
        amount = @amount,
        status = @status,
        paid_date = @paid_date,
        notes = @notes,
        updated_at = @updated_at
      WHERE id = @id
    `).run({ ...payload, updated_at: stamp });
  } else {
    const inserted = db.prepare(`
      INSERT INTO personal_expenses (
        due_date, description, category, amount, status, paid_date, notes, created_at, updated_at
      ) VALUES (
        @due_date, @description, @category, @amount, @status, @paid_date, @notes, @created_at, @updated_at
      )
    `).run({ ...payload, created_at: stamp, updated_at: stamp });
    payload.id = inserted.lastInsertRowid;
  }

  return db.prepare(`
    SELECT id, due_date, description, category, amount, status, paid_date, notes
    FROM personal_expenses
    WHERE id = ?
  `).get(payload.id);
}

function upsertPortfolioPosition(db, row) {
  const stamp = nowISO();
  const type = row.asset_type === "CDB" ? "CDB" : row.asset_type === "FII" ? "FII" : "ACAO";

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
    application_date: type === "CDB" ? (row.application_date || todayISO()) : null,
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
  upsertPersonalExpense,
  upsertPortfolioPosition,
  updatePortfolioLiveValues
};
