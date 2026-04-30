const fs = require("fs");
const path = require("path");
const { app, BrowserWindow, Menu, ipcMain, Notification, shell } = require("electron");
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
  deletePersonalIncome,
  upsertPersonalExpense,
  upsertPortfolioPosition,
  deleteBusinessEntry,
  deletePersonalExpense,
  deletePortfolioPosition,
  updatePortfolioLiveValues
} = require("./database");

const POLLING_INTERVAL_MS = 3 * 60 * 1000;
const BILL_NOTIFICATION_INTERVAL_MS = 6 * 60 * 60 * 1000;
const UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000;
const MAX_AUTO_BACKUPS = 300;
const UPDATE_REPO_OWNER = process.env.FINANCA_UPDATE_OWNER || "Thiago-DD";
const UPDATE_REPO_NAME = process.env.FINANCA_UPDATE_REPO || "Financa";
const UPDATE_ASSET_NAME = process.env.FINANCA_UPDATE_ASSET || "FinanceiroPessoal-Setup.exe";

let db = null;
let mainWindow = null;
let pollTimer = null;
let billNotifyTimer = null;
let updateTimer = null;
let backupsDirPath = "";
let backupQueue = Promise.resolve();
let lastUpdateTagNotified = null;
let lastBillToastDay = "";
const notifiedBillKeys = new Set();

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

function collectUpcomingBills(windowDays = 2) {
  if (!db) return { generatedAt: new Date().toISOString(), windowDays, bills: [] };

  const now = new Date();
  const todayIso = formatISODate(now);
  const tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  const tomorrowIso = formatISODate(tomorrow);
  const plusN = new Date(now);
  plusN.setDate(now.getDate() + Number(windowDays || 2));
  const endIso = formatISODate(plusN);

  const rows = getPendingExpensesDueBetween(db, todayIso, endIso);
  const bills = rows.map((bill) => ({
    id: Number(bill.id),
    due_date: String(bill.due_date || ""),
    description: String(bill.description || "Conta"),
    category: String(bill.category || ""),
    amount: Number(bill.amount || 0),
    status: String(bill.status || "Pendente"),
    label: dayLabel(String(bill.due_date || ""), todayIso, tomorrowIso)
  }));

  return {
    generatedAt: new Date().toISOString(),
    windowDays: Number(windowDays || 2),
    bills
  };
}

function normalizeSemver(value) {
  const match = String(value || "").trim().match(/v?(\d+)\.(\d+)\.(\d+)/i);
  if (!match) return null;
  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

function compareSemver(a, b) {
  const va = normalizeSemver(a);
  const vb = normalizeSemver(b);
  if (!va || !vb) return 0;

  for (let i = 0; i < 3; i += 1) {
    if (va[i] > vb[i]) return 1;
    if (va[i] < vb[i]) return -1;
  }

  return 0;
}

function resolveReleaseDownloadUrl(release) {
  const assets = Array.isArray(release?.assets) ? release.assets : [];
  const preferred = assets.find((item) => item?.name === UPDATE_ASSET_NAME);
  if (preferred?.browser_download_url) return preferred.browser_download_url;
  return release?.html_url || "";
}

async function checkForRepositoryUpdate() {
  if (!mainWindow || mainWindow.isDestroyed()) return;

  const endpoint = `https://api.github.com/repos/${UPDATE_REPO_OWNER}/${UPDATE_REPO_NAME}/releases/latest`;
  const headers = {
    Accept: "application/vnd.github+json",
    "User-Agent": "FinancaDesktopUpdater"
  };

  const token = process.env.FINANCA_GH_TOKEN || process.env.GH_TOKEN;
  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }

  try {
    const response = await fetch(endpoint, { headers, cache: "no-store" });

    if (response.status === 404) {
      return;
    }

    if (!response.ok) {
      throw new Error(`GitHub update check falhou com status ${response.status}`);
    }

    const release = await response.json();
    const latestTag = String(release?.tag_name || "").trim();
    if (!latestTag) return;

    const latestVersion = latestTag.replace(/^v/i, "");
    const currentVersion = app.getVersion();
    const isNewer = compareSemver(latestVersion, currentVersion) > 0;

    if (!isNewer) return;

    if (lastUpdateTagNotified === latestTag) return;
    lastUpdateTagNotified = latestTag;

    const payload = {
      tag: latestTag,
      version: latestVersion,
      currentVersion,
      releaseUrl: release?.html_url || "",
      downloadUrl: resolveReleaseDownloadUrl(release)
    };

    mainWindow.webContents.send("app:updateAvailable", payload);

    if (Notification.isSupported()) {
      new Notification({
        title: "Atualizacao disponivel",
        body: `Nova versao ${latestVersion} encontrada. Abra o app para atualizar.`,
        timeoutType: "default"
      }).show();
    }

  } catch (error) {
    console.error("[Updater] erro ao verificar atualizacao:", error.message);
  }
}

function startUpdateScheduler() {
  if (updateTimer) clearInterval(updateTimer);

  checkForRepositoryUpdate().catch((error) => {
    console.error("[Updater inicial] erro:", error.message);
  });

  updateTimer = setInterval(() => {
    checkForRepositoryUpdate().catch((error) => {
      console.error("[Updater interval] erro:", error.message);
    });
  }, UPDATE_CHECK_INTERVAL_MS);
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
  const payload = collectUpcomingBills(2);

  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("app:billAlerts", payload);
  }

  if (!Notification.isSupported()) return payload;

  const todayIso = formatISODate(new Date());
  if (lastBillToastDay !== todayIso) {
    lastBillToastDay = todayIso;
    notifiedBillKeys.clear();
  }

  for (const bill of payload.bills) {
    const uniqueKey = `${todayIso}|${bill.id}|${bill.due_date}|${bill.amount}`;
    if (notifiedBillKeys.has(uniqueKey)) continue;
    notifiedBillKeys.add(uniqueKey);

    const value = Number(bill.amount || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });

    new Notification({
      title: "Lembrete Financeiro",
      body: `${bill.description} vence ${bill.label} (${value}).`,
      timeoutType: "default"
    }).show();
  }

  return payload;
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1200,
    minHeight: 720,
    backgroundColor: "#090909",
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.js")
    }
  });

  Menu.setApplicationMenu(null);
  mainWindow.setMenuBarVisibility(false);
  if (typeof mainWindow.removeMenu === "function") {
    mainWindow.removeMenu();
  }

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
    billAlerts: collectUpcomingBills(2),
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

  ipcMain.handle("db:deletePersonalIncome", async (_event, id) => {
    const ok = deletePersonalIncome(db, id);
    if (ok) await queueAutoBackupSafe("personal-income-delete");
    return { ok };
  });

  ipcMain.handle("db:upsertPersonalExpense", async (_event, row) => {
    const saved = upsertPersonalExpense(db, row);
    await queueAutoBackupSafe("personal-expense");
    notifyUpcomingBills();
    return saved;
  });

  ipcMain.handle("db:upsertPortfolioPosition", async (_event, row) => {
    const saved = upsertPortfolioPosition(db, row);
    await queueAutoBackupSafe("portfolio-position");
    return saved;
  });

  ipcMain.handle("db:deleteBusinessEntry", async (_event, id) => {
    const ok = deleteBusinessEntry(db, id);
    if (ok) await queueAutoBackupSafe("business-delete");
    return { ok };
  });

  ipcMain.handle("db:deletePersonalExpense", async (_event, id) => {
    const ok = deletePersonalExpense(db, id);
    if (ok) await queueAutoBackupSafe("personal-delete");
    notifyUpcomingBills();
    return { ok };
  });

  ipcMain.handle("db:deletePortfolioPosition", async (_event, id) => {
    const ok = deletePortfolioPosition(db, id);
    if (ok) await queueAutoBackupSafe("portfolio-delete");
    return { ok };
  });

  ipcMain.handle("db:setUserSetting", async (_event, payload) => {
    setUserSetting(db, payload.key, payload.value);
    await queueAutoBackupSafe("user-settings");
    return { key: payload.key, value: String(payload.value) };
  });

  ipcMain.handle("app:checkForUpdatesNow", async () => {
    await checkForRepositoryUpdate();
    return { ok: true };
  });

  ipcMain.handle("app:getBillAlerts", () => notifyUpcomingBills());

  ipcMain.handle("app:openExternalLink", async (_event, rawUrl) => {
    const parsed = String(rawUrl || "").trim();
    if (!parsed) return { ok: false, reason: "url-vazia" };

    let url;
    try {
      url = new URL(parsed);
    } catch (_error) {
      return { ok: false, reason: "url-invalida" };
    }

    if (url.protocol !== "https:") {
      return { ok: false, reason: "protocolo-nao-permitido" };
    }

    await shell.openExternal(url.toString());
    return { ok: true };
  });
}

app.whenReady()
  .then(() => {
    if (process.platform === "win32") {
      app.setAppUserModelId("com.thiagodd.financa.desktop");
    }

    const dataDir = path.join(app.getPath("userData"), "FinanceSpreadsheetData");
    const dbPath = path.join(dataDir, "financeiro.db");
    backupsDirPath = path.join(dataDir, "backups");

    db = initDatabase(dbPath);
    ensureBackupsDir();
    registerIpcHandlers();
    createMainWindow();
    startPolling();
    startBillNotificationScheduler();
    startUpdateScheduler();

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
  if (updateTimer) clearInterval(updateTimer);
  if (db) db.close();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
