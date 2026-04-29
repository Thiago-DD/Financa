const fs = require("fs");
const path = require("path");
const { app, BrowserWindow, ipcMain, Notification } = require("electron");
const YahooFinance = require("yahoo-finance2").default;
const {
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
} = require("./database");

const POLLING_INTERVAL_MS = 3 * 60 * 1000;
const BILL_NOTIFICATION_INTERVAL_MS = 6 * 60 * 60 * 1000;
const MAX_AUTO_BACKUPS = 300;

let db = null;
let mainWindow = null;
let pollTimer = null;
let billNotifyTimer = null;
let backupsDirPath = "";
let backupQueue = Promise.resolve();

const yahooFinance = new YahooFinance();

function mapTickerToYahoo(symbol) {
  const normalized = String(symbol || "").toUpperCase().trim();
  if (!normalized) return normalized;

  if (
    normalized.includes(".") ||
    normalized.includes("=") ||
    normalized.includes("-") ||
    normalized.startsWith("^")
  ) {
    return normalized;
  }

  return `${normalized}.SA`;
}

function formatISODate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function dayLabel(targetIso, todayIso, tomorrowIso) {
  if (targetIso === todayIso) return "hoje";
  if (targetIso === tomorrowIso) return "amanha";
  return `em ${targetIso.split("-").reverse().join("/")}`;
}

function safeBackupReason(reason) {
  const normalized = String(reason || "alteracao")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();

  return normalized.slice(0, 40) || "alteracao";
}

function backupStamp(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  const ms = String(date.getMilliseconds()).padStart(3, "0");
  return `${year}${month}${day}-${hour}${minute}${second}-${ms}`;
}

function ensureBackupsDir() {
  if (!backupsDirPath) return;
  fs.mkdirSync(backupsDirPath, { recursive: true });
}

function pruneOldBackups() {
  if (!backupsDirPath) return;

  const files = fs.readdirSync(backupsDirPath)
    .filter((file) => file.startsWith("backup-") && file.endsWith(".db"))
    .map((file) => {
      const fullPath = path.join(backupsDirPath, file);
      const stats = fs.statSync(fullPath);
      return { file, fullPath, mtimeMs: stats.mtimeMs };
    })
    .sort((a, b) => b.mtimeMs - a.mtimeMs);

  if (files.length <= MAX_AUTO_BACKUPS) return;

  for (const stale of files.slice(MAX_AUTO_BACKUPS)) {
    try {
      fs.unlinkSync(stale.fullPath);
    } catch (error) {
      console.error("[Backup] Falha ao remover snapshot antigo:", stale.file, error.message);
    }
  }
}

async function createAutoBackup(reason) {
  if (!db || !backupsDirPath) return null;

  ensureBackupsDir();
  const suffix = safeBackupReason(reason);
  const filename = `backup-${backupStamp()}-${suffix}.db`;
  const targetPath = path.join(backupsDirPath, filename);

  await db.backup(targetPath);
  pruneOldBackups();

  return targetPath;
}

function queueAutoBackup(reason) {
  backupQueue = backupQueue
    .catch(() => null)
    .then(() => createAutoBackup(reason));

  return backupQueue;
}

async function queueAutoBackupSafe(reason) {
  try {
    await queueAutoBackup(reason);
  } catch (error) {
    console.error("[Backup] Falha ao gerar backup automatico:", error.message);
  }
}

function notifyUpcomingBills() {
  if (!db || !Notification.isSupported()) return;

  const now = new Date();
  const todayIso = formatISODate(now);
  const tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  const tomorrowIso = formatISODate(tomorrow);
  const plus2 = new Date(now);
  plus2.setDate(now.getDate() + 2);
  const plus2Iso = formatISODate(plus2);

  const bills = getPendingExpensesDueBetween(db, todayIso, plus2Iso);
  if (!bills.length) return;

  for (const bill of bills) {
    const value = Number(bill.amount || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
    const label = dayLabel(bill.due_date, todayIso, tomorrowIso);

    new Notification({
      title: "Lembrete Financeiro",
      body: `${bill.description} vence ${label} (${value}).`,
      timeoutType: "default"
    }).show();
  }
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1200,
    minHeight: 720,
    backgroundColor: "#090909",
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.js")
    }
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  mainWindow.once("ready-to-show", () => mainWindow.show());
  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

async function runQuotePollingAndNotify() {
  if (!db) return;

  const quotedAt = new Date().toISOString();
  const updates = [];

  const quoteTargets = getQuoteTargets(db);
  for (const row of quoteTargets) {
    const yahooTicker = mapTickerToYahoo(row.ticker);
    try {
      const quote = await yahooFinance.quote(yahooTicker, {
        fields: ["symbol", "regularMarketPrice"]
      });
      const newUnitPrice = Number(quote?.regularMarketPrice);
      if (!Number.isFinite(newUnitPrice) || newUnitPrice <= 0) continue;

      const oldUnit = Number(row.current_value) || 0;
      const totalValue = (Number(row.quantity) || 0) * newUnitPrice;

      updatePortfolioLiveValues(db, {
        id: row.id,
        ticker: row.ticker,
        current_unit_price: newUnitPrice,
        current_value: newUnitPrice,
        total_market_value: totalValue,
        quoted_at: quotedAt,
        source: "yahoo-finance2"
      });

      updates.push({
        id: row.id,
        oldValue: oldUnit,
        newValue: newUnitPrice,
        newUnitPrice,
        direction: newUnitPrice >= oldUnit ? "up" : "down",
        quotedAt
      });
    } catch (error) {
      console.error(`[Polling] ticker ${yahooTicker} falhou:`, error.message);
    }
  }

  const cdbTargets = getCdbTargets(db);
  for (const row of cdbTargets) {
    const oldTotal = Number(row.current_value) || 0;
    const totalValue = calculateCdbProRata(
      row.application_amount,
      row.rate_percent,
      row.cdi_annual_rate,
      row.application_date,
      new Date(quotedAt)
    );

    updatePortfolioLiveValues(db, {
      id: row.id,
      ticker: row.ticker,
      current_unit_price: 0,
      current_value: totalValue,
      total_market_value: totalValue,
      quoted_at: quotedAt,
      source: "cdb-pro-rata"
    });

    updates.push({
      id: row.id,
      oldValue: oldTotal,
      newValue: totalValue,
      newUnitPrice: 0,
      direction: totalValue >= oldTotal ? "up" : "down",
      quotedAt
    });
  }

  if (updates.length && mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("portfolio:updated", {
      quotedAt,
      updates
    });
  }
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);

  runQuotePollingAndNotify().catch((error) => {
    console.error("[Polling inicial] erro:", error.message);
  });

  pollTimer = setInterval(() => {
    runQuotePollingAndNotify().catch((error) => {
      console.error("[Polling interval] erro:", error.message);
    });
  }, POLLING_INTERVAL_MS);
}

function startBillNotificationScheduler() {
  if (billNotifyTimer) clearInterval(billNotifyTimer);

  notifyUpcomingBills();
  billNotifyTimer = setInterval(() => {
    notifyUpcomingBills();
  }, BILL_NOTIFICATION_INTERVAL_MS);
}

function registerIpcHandlers() {
  const getGoal = () => {
    const raw = getUserSetting(db, "meta_patrimonio", "100000");
    const value = Number(raw);
    return Number.isFinite(value) && value > 0 ? value : 100000;
  };

  ipcMain.handle("app:getInitialData", () => ({
    businessEntries: getBusinessEntries(db),
    personalIncomes: getPersonalIncomes(db),
    personalExpenses: getPersonalExpenses(db),
    portfolioPositions: getPortfolioPositions(db),
    userSettings: getUserSettings(db),
    meta: {
      pollingIntervalMs: POLLING_INTERVAL_MS,
      portfolioGoal: getGoal(),
      dbPath: db.name,
      backupsPath: backupsDirPath
    }
  }));

  ipcMain.handle("db:upsertBusinessEntry", async (_event, row) => {
    const saved = upsertBusinessEntry(db, row);
    await queueAutoBackupSafe("business-entry");
    return saved;
  });

  ipcMain.handle("db:addPersonalIncome", async (_event, row) => {
    const saved = addPersonalIncome(db, row);
    await queueAutoBackupSafe("personal-income");
    return saved;
  });

  ipcMain.handle("db:upsertPersonalExpense", async (_event, row) => {
    const saved = upsertPersonalExpense(db, row);
    await queueAutoBackupSafe("personal-expense");
    return saved;
  });

  ipcMain.handle("db:upsertPortfolioPosition", async (_event, row) => {
    const saved = upsertPortfolioPosition(db, row);
    await queueAutoBackupSafe("portfolio-position");
    return saved;
  });

  ipcMain.handle("db:setUserSetting", async (_event, payload) => {
    setUserSetting(db, payload.key, payload.value);
    await queueAutoBackupSafe("user-settings");
    return { key: payload.key, value: String(payload.value) };
  });
}

app.whenReady()
  .then(() => {
    const dataDir = path.join(app.getPath("userData"), "FinanceSpreadsheetData");
    const dbPath = path.join(dataDir, "financeiro.db");
    backupsDirPath = path.join(dataDir, "backups");

    db = initDatabase(dbPath);
    ensureBackupsDir();
    registerIpcHandlers();
    createMainWindow();
    startPolling();
    startBillNotificationScheduler();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
    });
  })
  .catch((error) => {
    console.error("[Startup] erro fatal:", error);
    app.quit();
  });

process.on("unhandledRejection", (reason) => {
  console.error("[UnhandledRejection]", reason);
});

app.on("before-quit", () => {
  if (pollTimer) clearInterval(pollTimer);
  if (billNotifyTimer) clearInterval(billNotifyTimer);
  if (db) db.close();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
