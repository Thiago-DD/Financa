const BRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const NUMBER_BR = new Intl.NumberFormat("pt-BR");

const VIEW_KEYS = ["business", "personal", "investments"];
const DEFAULT_BUSINESS_CATEGORIES = ["Sangria", "Fornecedor", "Venda Balcao"];
const DEFAULT_PERSONAL_CATEGORIES = ["Moradia", "Alimentacao", "Transporte", "Lazer"];
const DELETE_SELECTOR_COL_ID = "__delete_selector__";
const PIE_COLORS_INCOME = ["#22c55e", "#16a34a", "#4ade80", "#65a30d", "#10b981", "#84cc16"];
const PIE_COLORS_EXPENSE = ["#ef4444", "#dc2626", "#f97316", "#fb7185", "#b91c1c", "#ea580c"];
const MONTH_NAMES_SHORT_PTBR = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];

const appState = {
  activeView: "business",
  darkMode: true,
  portfolioGoal: 100000,
  businessRows: [],
  personalIncomes: [],
  personalExpenses: [],
  portfolioRows: [],
  businessMonthFilter: "",
  personalMonthFilter: "",
  investmentMonthFilter: "",
  businessSortMode: "date_desc",
  personalSortMode: "due_asc",
  investmentSortMode: "result_desc",
  personalFilter: "none",
  businessCategories: [...DEFAULT_BUSINESS_CATEGORIES],
  personalCategories: [...DEFAULT_PERSONAL_CATEGORIES],
  businessDeleteMode: false,
  personalDeleteMode: false,
  investmentDeleteMode: false,
  pendingUpdate: null,
  updateStatus: "idle",
  updateProgressPercent: 0,
  riskBannerDismissed: false,
  notifications: [],
  notificationsOpen: false,
  unreadNotifications: 0,
  lastBillNotificationKey: "",
  lastUpdateNotificationKey: ""
};

const gridState = {
  businessApi: null,
  personalApi: null,
  portfolioApi: null,
  savingBusiness: false,
  savingPersonal: false,
  savingPortfolio: false
};

const refs = {};
let tempSeq = 1;

function byId(id) {
  return document.getElementById(id);
}

class CategorySelectEditor {
  init(params) {
    this.params = params;
    this.eSelect = document.createElement("select");
    this.eSelect.className = "ag-input-field-input ag-text-field-input";
    this.eSelect.style.width = "100%";
    this.eSelect.style.height = "100%";
    this.eSelect.style.border = "0";
    this.eSelect.style.outline = "none";
    this.eSelect.style.backgroundColor = "transparent";
    this.eSelect.style.color = "inherit";
    this.eSelect.style.padding = "0 4px";

    const values = Array.isArray(params?.values)
      ? params.values
      : Array.isArray(params?.colDef?.cellEditorParams?.values)
        ? params.colDef.cellEditorParams.values
        : [];
    const safeValues = values.map((value) => String(value));
    const currentValue = String(params?.value ?? "");

    if (currentValue && !safeValues.includes(currentValue)) {
      safeValues.unshift(currentValue);
    }

    for (const value of safeValues) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      this.eSelect.appendChild(option);
    }

    this.eSelect.value = currentValue;
    if (params?.center) {
      this.eSelect.style.textAlign = "center";
      this.eSelect.style.textAlignLast = "center";
    }

    this.onChange = () => {
      this.params?.stopEditing?.();
    };
    this.onKeyDown = (event) => {
      if (event.key === "Enter" || event.key === "Tab") {
        this.params?.stopEditing?.();
      }
      if (event.key === "Escape") {
        this.params?.stopEditing?.(true);
      }
    };

    this.eSelect.addEventListener("change", this.onChange);
    this.eSelect.addEventListener("keydown", this.onKeyDown);
  }

  getGui() {
    return this.eSelect;
  }

  afterGuiAttached() {
    if (!this.eSelect) return;
    this.eSelect.focus();

    try {
      if (typeof this.eSelect.showPicker === "function") {
        this.eSelect.showPicker();
      } else {
        this.eSelect.click();
      }
    } catch (_error) {
      try {
        this.eSelect.click();
      } catch (_ignored) {
        // noop
      }
    }
  }

  getValue() {
    return this.eSelect?.value ?? "";
  }

  destroy() {
    if (this.eSelect && this.onChange) {
      this.eSelect.removeEventListener("change", this.onChange);
    }
    if (this.eSelect && this.onKeyDown) {
      this.eSelect.removeEventListener("keydown", this.onKeyDown);
    }
  }
}

function toNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value !== "string") return Number(value) || 0;
  const normalized = value.replace(/\./g, "").replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toInstallmentTotal(value) {
  const raw = String(value ?? "").trim();
  const numeric = Number(raw);
  if (Number.isFinite(numeric)) {
    const whole = Math.floor(numeric);
    return Math.min(360, Math.max(1, whole));
  }

  const digits = raw.match(/\d+/);
  if (digits) {
    const parsed = Number(digits[0]);
    if (Number.isFinite(parsed)) {
      const whole = Math.floor(parsed);
      return Math.min(360, Math.max(1, whole));
    }
  }

  return 1;
}

function isValidIsoDate(value) {
  const text = String(value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return false;
  const [year, month, day] = text.split("-").map(Number);
  const probe = new Date(Date.UTC(year, month - 1, day));
  return (
    probe.getUTCFullYear() === year &&
    probe.getUTCMonth() + 1 === month &&
    probe.getUTCDate() === day
  );
}

function lastDayOfMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function normalizeIsoDate(value, fallbackIso = todayISO(), baseIso = fallbackIso) {
  const fallback = isValidIsoDate(fallbackIso) ? fallbackIso : todayISO();
  const base = isValidIsoDate(baseIso) ? baseIso : fallback;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  const text = String(value || "").trim();
  if (!text) return fallback;
  if (isValidIsoDate(text)) return text;

  const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) {
    const day = Number(br[1]);
    const month = Number(br[2]);
    const year = Number(br[3]);
    const iso = `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    return isValidIsoDate(iso) ? iso : fallback;
  }

  const dayOnly = text.match(/^(\d{1,2})$/);
  if (dayOnly) {
    const day = Number(dayOnly[1]);
    const [yearText, monthText] = base.split("-");
    const year = Number(yearText);
    const month = Number(monthText);
    const clampedDay = Math.min(Math.max(day, 1), lastDayOfMonth(year, month));
    return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(clampedDay).padStart(2, "0")}`;
  }

  return fallback;
}

function todayISO() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function currentMonthPrefix() {
  return todayISO().slice(0, 7);
}

function parseMonthFromInput(value) {
  const text = String(value || "").trim();
  if (!text) return "";

  const monthMatch = text.match(/^(\d{4}-\d{2})$/);
  if (monthMatch) return monthMatch[1];

  const dateMatch = text.match(/^(\d{4}-\d{2})-\d{2}$/);
  if (dateMatch) return dateMatch[1];

  const brDateMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (brDateMatch) {
    const month = Number(brDateMatch[2]);
    const year = Number(brDateMatch[3]);
    if (Number.isFinite(month) && Number.isFinite(year) && month >= 1 && month <= 12) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}`;
    }
  }

  const brMonthMatch = text.match(/^(\d{1,2})\/(\d{4})$/);
  if (brMonthMatch) {
    const month = Number(brMonthMatch[1]);
    const year = Number(brMonthMatch[2]);
    if (Number.isFinite(month) && Number.isFinite(year) && month >= 1 && month <= 12) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}`;
    }
  }

  return "";
}

function monthToDateInputValue(monthValue) {
  const month = extractMonth(monthValue);
  return month || "";
}

function sanitizeCategory(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .slice(0, 60);
}

function uniqueCategories(values) {
  const out = [];
  const seen = new Set();

  for (const raw of values || []) {
    const clean = sanitizeCategory(raw);
    if (!clean) continue;
    const key = clean.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(clean);
  }

  return out;
}

function parseSettingArray(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(String(raw));
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function addCategoryToList(list, value) {
  const clean = sanitizeCategory(value);
  if (!clean) return { added: false, value: "" };

  const exists = list.some((item) => String(item).toLowerCase() === clean.toLowerCase());
  if (!exists) list.push(clean);

  return { added: !exists, value: clean };
}

function persistCategorySetting(settingKey, categories) {
  window.financeAPI
    .setUserSetting(settingKey, JSON.stringify(categories))
    .catch((error) => console.error("Falha ao salvar categorias:", error));
}

function initCategoryState(userSettings = {}) {
  const businessFromSetting = parseSettingArray(userSettings.business_categories);
  const personalFromSetting = parseSettingArray(userSettings.personal_categories);

  const businessFromRows = appState.businessRows.map((row) => row.category);
  const personalFromRows = appState.personalExpenses.map((row) => row.category);

  appState.businessCategories = uniqueCategories([
    ...DEFAULT_BUSINESS_CATEGORIES,
    ...businessFromSetting,
    ...businessFromRows
  ]);

  appState.personalCategories = uniqueCategories([
    ...DEFAULT_PERSONAL_CATEGORIES,
    ...personalFromSetting,
    ...personalFromRows
  ]);
}

function ensureTempId(row) {
  if (!row.id && !row.__tmpId) {
    row.__tmpId = `tmp-${tempSeq++}`;
  }
  return row;
}

function formatDateBR(value) {
  if (!value) return "";
  const date = new Date(`${value}T12:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("pt-BR").format(date);
}

function normalizeBusinessTypeToken(value) {
  return String(value || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z]/g, "");
}

function dbToUiBusinessType(value) {
  const token = normalizeBusinessTypeToken(value);
  if (token.includes("entrada")) return "Entrada";
  if (token.startsWith("s") || token.includes("saida") || token.includes("sada")) return "Saida";
  return "Entrada";
}

function uiToDbBusinessType(value) {
  return String(value || "").toLowerCase().startsWith("s") ? "Saida" : "Entrada";
}

function dbToUiStatus(value) {
  return value === "Pago" || value === "PAGO" ? "Pago" : "Pendente";
}

function uiToDbStatus(value) {
  return String(value || "").toLowerCase().startsWith("pago") ? "Pago" : "Pendente";
}

function dbToUiRecurring(value) {
  return Number(value) > 0 ? "Sim" : "Nao";
}

function uiToDbRecurring(value) {
  if (Number(value) > 0) return 1;
  const text = String(value || "").trim().toLowerCase();
  if (!text) return 0;
  return text.startsWith("s") || text === "true" ? 1 : 0;
}

function dbToUiAssetType(value) {
  if (value === "CDB") return "CDB";
  if (value === "FII") return "FII";
  return "Acao";
}

function uiToDbAssetType(value) {
  const t = String(value || "").toLowerCase();
  if (t.includes("cdb")) return "CDB";
  if (t.includes("fii")) return "FII";
  return "ACAO";
}

function normalizeBusinessRow(row) {
  const entryDate = normalizeIsoDate(row.entry_date, todayISO(), row.entry_date || todayISO());
  const installmentTotal = toInstallmentTotal(row.installment_total);
  const installmentIndex = Math.min(
    installmentTotal,
    Math.max(1, Number.isFinite(Number(row.installment_index)) ? Math.floor(Number(row.installment_index)) : 1)
  );
  return ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    __isNew: Boolean(row.__isNew),
    entry_date: entryDate,
    description: row.description || "",
    entry_type: uiToDbBusinessType(row.entry_type),
    category: row.category || "Venda Balcao",
    amount: toNumber(row.amount),
    installment_total: installmentTotal,
    installment_index: installmentIndex,
    installment_group_id: row.installment_group_id || null,
    is_recurring: installmentTotal > 1 ? 0 : uiToDbRecurring(row.is_recurring),
    recurrence_interval_months: installmentTotal > 1
      ? 1
      : toInstallmentTotal(row.recurrence_interval_months || 1),
    notes: row.notes || ""
  });
}

function normalizePersonalExpenseRow(row) {
  const status = uiToDbStatus(row.status);
  const dueDate = normalizeIsoDate(row.due_date, todayISO(), row.due_date || todayISO());
  const installmentTotal = toInstallmentTotal(row.installment_total);
  const installmentIndex = Math.min(
    installmentTotal,
    Math.max(1, Number.isFinite(Number(row.installment_index)) ? Math.floor(Number(row.installment_index)) : 1)
  );
  const paidDate = status === "Pago"
    ? normalizeIsoDate(row.paid_date || dueDate, dueDate, dueDate)
    : null;

  return ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    __isNew: Boolean(row.__isNew),
    due_date: dueDate,
    description: row.description || "",
    category: row.category || "Moradia",
    amount: toNumber(row.amount),
    status,
    paid_date: paidDate,
    installment_total: installmentTotal,
    installment_index: installmentIndex,
    installment_group_id: row.installment_group_id || null,
    is_recurring: installmentTotal > 1 ? 0 : uiToDbRecurring(row.is_recurring),
    recurrence_interval_months: installmentTotal > 1
      ? 1
      : toInstallmentTotal(row.recurrence_interval_months || 1),
    notes: row.notes || ""
  });
}

function normalizePersonalIncomeRow(row) {
  return {
    id: row.id ?? null,
    income_date: normalizeIsoDate(row.income_date, todayISO(), row.income_date || todayISO()),
    description: row.description || "",
    source: row.source || "SALARIO",
    amount: toNumber(row.amount),
    notes: row.notes || ""
  };
}

function normalizePortfolioRow(row) {
  const type = uiToDbAssetType(row.asset_type);
  const rawApplicationDate = row.application_date || row.updated_at || todayISO();
  const applicationDate = normalizeIsoDate(rawApplicationDate, todayISO(), rawApplicationDate);

  const clean = ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    __isNew: Boolean(row.__isNew),
    asset_label: row.asset_label || "",
    asset_type: type,
    ticker: row.ticker ? String(row.ticker).toUpperCase() : "",
    quantity: type === "CDB" ? 0 : toNumber(row.quantity),
    avg_price: type === "CDB" ? 0 : toNumber(row.avg_price),
    application_amount: type === "CDB" ? toNumber(row.application_amount) : 0,
    rate_percent: type === "CDB" ? toNumber(row.rate_percent) : 0,
    cdi_annual_rate: type === "CDB" ? toNumber(row.cdi_annual_rate || 13.65) : 13.65,
    application_date: applicationDate,
    current_unit_price: type === "CDB" ? 0 : toNumber(row.current_unit_price),
    current_value: toNumber(row.current_value),
    last_dividend: type === "CDB" ? 0 : toNumber(row.last_dividend),
    updated_at: row.updated_at || "",
    flash_direction: ""
  });

  return clean;
}

function indexByIdentity(list, row) {
  if (row.id) {
    const idxById = list.findIndex((item) => Number(item.id) === Number(row.id));
    if (idxById >= 0) return idxById;
  }
  if (row.__tmpId) return list.findIndex((item) => item.__tmpId && item.__tmpId === row.__tmpId);
  return -1;
}

function upsertIntoList(list, row) {
  const idx = indexByIdentity(list, row);
  if (idx >= 0) list[idx] = { ...list[idx], ...row };
  else list.unshift(row);
}

function initRefs() {
  refs.body = document.body;
  refs.sidebar = byId("sidebar");
  refs.sidebarToggle = byId("sidebarToggle");
  refs.themeToggle = byId("themeToggle");
  refs.checkUpdatesButton = byId("checkUpdatesButton");
  refs.globalSearch = byId("globalSearch");
  refs.updateBanner = byId("updateBanner");
  refs.updateText = byId("updateText");
  refs.openUpdateButton = byId("openUpdateButton");
  refs.dismissUpdateButton = byId("dismissUpdateButton");
  refs.notificationToggle = byId("notificationToggle");
  refs.notificationCount = byId("notificationCount");
  refs.notificationPanel = byId("notificationPanel");
  refs.notificationMeta = byId("notificationMeta");
  refs.notificationList = byId("notificationList");
  refs.notificationMarkRead = byId("notificationMarkRead");
  refs.riskBanner = byId("riskBanner");
  refs.riskBannerTitle = byId("riskBannerTitle");
  refs.riskBannerMeta = byId("riskBannerMeta");
  refs.riskBannerList = byId("riskBannerList");
  refs.riskBannerDismiss = byId("riskBannerDismiss");

  refs.nav = {
    business: byId("navBusiness"),
    personal: byId("navPersonal"),
    investments: byId("navInvestments")
  };
  refs.views = {
    business: byId("viewBusiness"),
    personal: byId("viewPersonal"),
    investments: byId("viewInvestments")
  };

  refs.sidebarBalance = byId("sidebarBalance");
  refs.topbarBalance = byId("topbarBalance");

  refs.addBusinessRow = byId("addBusinessRow");
  refs.deleteBusinessRow = byId("deleteBusinessRow");
  refs.addPersonalRow = byId("addPersonalRow");
  refs.deletePersonalRow = byId("deletePersonalRow");
  refs.addPortfolioRow = byId("addPortfolioRow");
  refs.deletePortfolioRow = byId("deletePortfolioRow");
  refs.businessSortMode = byId("businessSortMode");
  refs.personalSortMode = byId("personalSortMode");
  refs.investmentSortMode = byId("investmentSortMode");
  refs.businessCategoryInput = byId("businessCategoryInput");
  refs.addBusinessCategoryButton = byId("addBusinessCategoryButton");
  refs.businessMonthInput = byId("businessMonthInput");
  refs.businessMonthClear = byId("businessMonthClear");
  refs.businessMonthTabs = byId("businessMonthTabs");
  refs.businessPrevMonth = byId("businessPrevMonth");
  refs.businessNextMonth = byId("businessNextMonth");
  refs.personalCategoryInput = byId("personalCategoryInput");
  refs.addPersonalCategoryButton = byId("addPersonalCategoryButton");
  refs.personalMonthInput = byId("personalMonthInput");
  refs.personalMonthClear = byId("personalMonthClear");
  refs.personalMonthTabs = byId("personalMonthTabs");
  refs.personalPrevMonth = byId("personalPrevMonth");
  refs.personalNextMonth = byId("personalNextMonth");
  refs.investmentMonthInput = byId("investmentMonthInput");
  refs.investmentMonthClear = byId("investmentMonthClear");
  refs.investmentMonthTabs = byId("investmentMonthTabs");
  refs.investmentPrevMonth = byId("investmentPrevMonth");
  refs.investmentNextMonth = byId("investmentNextMonth");
  refs.addIncomeButton = byId("addIncomeButton");
  refs.incomeAmountInput = byId("incomeAmountInput");
  refs.incomeDescriptionInput = byId("incomeDescriptionInput");
  refs.personalIncomeList = byId("personalIncomeList");
  refs.personalIncomeScope = byId("personalIncomeScope");
  refs.goalInput = byId("goalInput");
  refs.saveGoalButton = byId("saveGoalButton");

  refs.businessSalesCard = byId("businessSalesCard");
  refs.businessExpensesCard = byId("businessExpensesCard");
  refs.businessBalanceCard = byId("businessBalanceCard");
  refs.businessPieChart = byId("businessPieChart");
  refs.businessPieLegend = byId("businessPieLegend");
  refs.businessPieTotal = byId("businessPieTotal");
  refs.businessPieComparison = byId("businessPieComparison");

  refs.personalIncomeCard = byId("personalIncomeCard");
  refs.personalExpenseCard = byId("personalExpenseCard");
  refs.personalOpenCard = byId("personalOpenCard");
  refs.personalOpenMeta = byId("personalOpenMeta");
  refs.personalPieChart = byId("personalPieChart");
  refs.personalPieLegend = byId("personalPieLegend");
  refs.personalPieTotal = byId("personalPieTotal");
  refs.personalPieComparison = byId("personalPieComparison");

  refs.filterPending = byId("filterPending");
  refs.filterOverdue = byId("filterOverdue");
  refs.filterMonth = byId("filterMonth");
  refs.filterClear = byId("filterClear");

  refs.investPatrimonyCard = byId("investPatrimonyCard");
  refs.investedCard = byId("investedCard");
  refs.investResultCard = byId("investResultCard");
  refs.goalProgressBar = byId("goalProgressBar");
  refs.goalProgressText = byId("goalProgressText");
  refs.portfolioUpdatedAt = byId("portfolioUpdatedAt");
}

function buildDeleteSelectorColumnDef() {
  return {
    colId: DELETE_SELECTOR_COL_ID,
    headerName: "",
    checkboxSelection: true,
    headerCheckboxSelection: true,
    headerCheckboxSelectionFilteredOnly: true,
    pinned: "left",
    lockPinned: true,
    editable: false,
    sortable: false,
    filter: false,
    resizable: false,
    suppressMovable: true,
    maxWidth: 52,
    minWidth: 52,
    width: 52,
    hide: true
  };
}

function getDeleteMode(viewKey) {
  if (viewKey === "business") return appState.businessDeleteMode;
  if (viewKey === "personal") return appState.personalDeleteMode;
  return appState.investmentDeleteMode;
}

function setDeleteModeState(viewKey, enabled) {
  if (viewKey === "business") appState.businessDeleteMode = enabled;
  else if (viewKey === "personal") appState.personalDeleteMode = enabled;
  else appState.investmentDeleteMode = enabled;
}

function getDeleteIdleLabel(viewKey) {
  if (viewKey === "personal") return "Excluir registro";
  return "Excluir linha";
}

function getDeleteButtonRef(viewKey) {
  if (viewKey === "business") return refs.deleteBusinessRow;
  if (viewKey === "personal") return refs.deletePersonalRow;
  return refs.deletePortfolioRow;
}

function getGridApiByView(viewKey) {
  if (viewKey === "business") return gridState.businessApi;
  if (viewKey === "personal") return gridState.personalApi;
  return gridState.portfolioApi;
}

function setDeleteSelectorVisible(api, visible) {
  if (!api) return;
  if (typeof api.setColumnsVisible === "function") {
    api.setColumnsVisible([DELETE_SELECTOR_COL_ID], visible);
  } else if (typeof api.setColumnVisible === "function") {
    api.setColumnVisible(DELETE_SELECTOR_COL_ID, visible);
  }

  if (!visible && typeof api.deselectAll === "function") {
    api.deselectAll();
  }
}

function setDeleteMode(viewKey, enabled) {
  setDeleteModeState(viewKey, enabled);
  const button = getDeleteButtonRef(viewKey);
  if (button) {
    button.textContent = enabled ? "Excluir selecionadas" : getDeleteIdleLabel(viewKey);
  }

  const api = getGridApiByView(viewKey);
  setDeleteSelectorVisible(api, enabled);
}

function disableDeleteModesExcept(activeViewKey) {
  for (const key of VIEW_KEYS) {
    if (key === activeViewKey) continue;
    if (getDeleteMode(key)) {
      setDeleteMode(key, false);
    }
  }
}

function applyGlobalSearch(value) {
  const text = String(value || "");
  const apis = [gridState.businessApi, gridState.personalApi, gridState.portfolioApi];
  for (const api of apis) {
    if (!api) continue;
    api.setGridOption("quickFilterText", text);
  }
}

function updateThemeToggleText() {
  refs.themeToggle.textContent = appState.darkMode ? "Light" : "Dark";
}

function resolveUpdateUrl(payload) {
  const downloadUrl = String(payload?.downloadUrl || "").trim();
  if (downloadUrl) return downloadUrl;
  return String(payload?.releaseUrl || "").trim();
}

function setUpdateActionButton(label, disabled = false) {
  if (!refs.openUpdateButton) return;
  refs.openUpdateButton.textContent = label;
  refs.openUpdateButton.disabled = Boolean(disabled);
  refs.openUpdateButton.classList.toggle("opacity-60", Boolean(disabled));
  refs.openUpdateButton.classList.toggle("cursor-not-allowed", Boolean(disabled));
}

function hideUpdateBanner() {
  appState.updateStatus = "idle";
  appState.updateProgressPercent = 0;
  if (!refs.updateBanner) return;
  refs.updateBanner.classList.add("hidden");
}

function nowPtBr() {
  return new Date().toLocaleString("pt-BR");
}

function setNotificationPanelOpen(open) {
  appState.notificationsOpen = Boolean(open);
  if (!refs.notificationPanel) return;
  refs.notificationPanel.classList.toggle("hidden", !appState.notificationsOpen);
  if (appState.notificationsOpen) {
    appState.unreadNotifications = 0;
    renderNotificationCenter();
  }
}

function renderNotificationCenter() {
  if (!refs.notificationCount || !refs.notificationMeta || !refs.notificationList) return;

  const unread = Math.max(0, Number(appState.unreadNotifications || 0));
  refs.notificationCount.textContent = String(unread > 99 ? "99+" : unread);
  refs.notificationCount.classList.toggle("hidden", unread <= 0);

  const total = appState.notifications.length;
  refs.notificationMeta.textContent = total > 0 ? `${total} no historico` : "Sem novidades";

  if (!total) {
    refs.notificationList.innerHTML = '<p class="notification-empty">Nenhuma notificacao no momento.</p>';
    return;
  }

  refs.notificationList.innerHTML = appState.notifications
    .slice(0, 50)
    .map((item) => {
      const title = escapeHtml(item.title || "Notificacao");
      const text = escapeHtml(item.text || "");
      const time = escapeHtml(item.time || "");
      return `
        <div class="notification-item">
          <p class="notification-item-title">${title}</p>
          <p class="notification-item-text">${text}</p>
          <p class="notification-item-time">${time}</p>
        </div>
      `;
    })
    .join("");
}

function pushNotification(title, text, dedupeKey = "") {
  const key = String(dedupeKey || "").trim();
  if (key) {
    const existingIdx = appState.notifications.findIndex((item) => item.key === key);
    if (existingIdx >= 0) {
      appState.notifications.splice(existingIdx, 1);
    }
  }

  appState.notifications.unshift({
    key: key || `notif-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    title: String(title || "Notificacao"),
    text: String(text || ""),
    time: nowPtBr()
  });

  if (appState.notifications.length > 200) {
    appState.notifications = appState.notifications.slice(0, 200);
  }

  if (!appState.notificationsOpen) {
    appState.unreadNotifications += 1;
  }

  renderNotificationCenter();
}

function applyUpdateState(payload) {
  const status = String(payload?.status || "").trim() || "available";
  const current = String(payload?.currentVersion || "").trim() || "atual";
  const latest = String(payload?.version || payload?.tag || "").trim() || "nova";
  const autoEnabled = Boolean(payload?.autoUpdateEnabled);
  const downloaded = Boolean(payload?.downloaded);
  const canAutoInstall = Boolean(payload?.canAutoInstall);
  const progress = Number(payload?.progressPercent || 0);

  appState.updateStatus = status;
  if (Number.isFinite(progress) && progress > 0) {
    appState.updateProgressPercent = Math.max(0, Math.min(100, progress));
  }

  if (status === "not-available") {
    appState.pendingUpdate = null;
    hideUpdateBanner();
    return;
  }

  if (status === "error") {
    if (refs.updateText) {
      refs.updateText.textContent = `Falha no update: ${String(payload?.message || "erro desconhecido")}`;
    }
    if (refs.updateBanner) {
      refs.updateBanner.classList.remove("hidden");
    }
    setUpdateActionButton("Abrir release", false);
    return;
  }

  if (status === "checking") {
    if (refs.updateText) {
      refs.updateText.textContent = "Verificando atualizacao...";
    }
    if (refs.updateBanner) refs.updateBanner.classList.remove("hidden");
    setUpdateActionButton("Aguarde...", true);
    return;
  }

  if (status === "downloading") {
    appState.pendingUpdate = { ...(appState.pendingUpdate || {}), ...payload };
    if (refs.updateText) {
      refs.updateText.textContent = `Baixando update ${latest}: ${appState.updateProgressPercent.toFixed(1)}%`;
    }
    if (refs.updateBanner) refs.updateBanner.classList.remove("hidden");
    setUpdateActionButton("Baixando...", true);
    return;
  }

  appState.pendingUpdate = payload || null;
  if (refs.updateBanner) refs.updateBanner.classList.remove("hidden");

  if (downloaded && canAutoInstall) {
    if (refs.updateText) {
      refs.updateText.textContent = `Update pronto: ${current} -> ${latest}. Reinicie para instalar.`;
    }
    setUpdateActionButton("Reiniciar e instalar", false);
  } else if (autoEnabled) {
    if (refs.updateText) {
      refs.updateText.textContent = `Atualizacao disponivel: ${current} -> ${latest}. Download automatico iniciado.`;
    }
    setUpdateActionButton("Baixando...", true);
  } else {
    if (refs.updateText) {
      refs.updateText.textContent = `Atualizacao disponivel: ${current} -> ${latest}`;
    }
    setUpdateActionButton("Baixar update", false);
  }

  const updateKey = String(payload?.tag || payload?.version || "").trim();
  if (updateKey && appState.lastUpdateNotificationKey !== updateKey) {
    appState.lastUpdateNotificationKey = updateKey;
    const notifMessage = downloaded && canAutoInstall
      ? `Nova versao ${latest} pronta para instalar.`
      : autoEnabled
        ? `Nova versao ${latest} detectada. Download automatico iniciado.`
        : `Nova versao ${latest} detectada. Clique em Atualizar para baixar.`;
    pushNotification("Atualizacao disponivel", notifMessage, `update:${updateKey}`);
  }
}

function showUpdateBanner(payload) {
  applyUpdateState(payload);
}

function hideRiskBanner() {
  appState.riskBannerDismissed = true;
  if (refs.riskBanner) {
    refs.riskBanner.classList.add("hidden");
  }
}

function showRiskBanner() {
  if (refs.riskBanner) {
    refs.riskBanner.classList.remove("hidden");
  }
}

function renderBillAlerts(payload) {
  if (!refs.riskBanner || !refs.riskBannerTitle || !refs.riskBannerMeta || !refs.riskBannerList) return;

  const bills = Array.isArray(payload?.bills) ? payload.bills : [];
  const windowDays = Number(payload?.windowDays || 2);
  const billKey = bills
    .map((bill) => `${bill.id}|${bill.due_date}|${toNumber(bill.amount).toFixed(2)}`)
    .sort()
    .join(";");

  if (!bills.length) {
    refs.riskBannerTitle.textContent = "Sem alertas de vencimento";
    refs.riskBannerMeta.textContent = `Nenhuma conta pendente vencendo nos proximos ${windowDays} dias.`;
    refs.riskBannerList.innerHTML = "";
    refs.riskBanner.classList.add("hidden");
    appState.riskBannerDismissed = false;
    appState.lastBillNotificationKey = "";
    return;
  }

  refs.riskBannerTitle.textContent = `${bills.length} conta(s) com vencimento proximo`;
  refs.riskBannerMeta.textContent = `Contas pendentes vencendo em ate ${windowDays} dias.`;

  refs.riskBannerList.innerHTML = bills
    .slice(0, 6)
    .map((bill) => {
      const label = String(bill.label || "em breve");
      const text = `${String(bill.description || "Conta")} (${String(bill.category || "Sem categoria")}) vence ${label}`;
      const amount = BRL.format(toNumber(bill.amount));
      return `
        <div class="risk-item">
          <span class="risk-item-label">${escapeHtml(text)}</span>
          <span class="risk-item-value">${escapeHtml(amount)}</span>
        </div>
      `;
    })
    .join("");

  if (billKey && appState.lastBillNotificationKey !== billKey) {
    appState.lastBillNotificationKey = billKey;
    const preview = bills.slice(0, 2).map((bill) => bill.description).join(", ");
    const suffix = bills.length > 2 ? ` e mais ${bills.length - 2}` : "";
    pushNotification(
      "Risco de vencimento",
      `${bills.length} conta(s) vencem em ate ${windowDays} dias: ${preview}${suffix}.`,
      `bills:${billKey}`
    );
  }

  if (!appState.riskBannerDismissed) {
    showRiskBanner();
  }
}

function applyThemeMode() {
  refs.body.classList.toggle("dark", appState.darkMode);
  refs.body.classList.toggle("theme-dark", appState.darkMode);
  refs.body.classList.toggle("theme-light", !appState.darkMode);

  const darkClass = "ag-theme-quartz-dark";
  const lightClass = "ag-theme-quartz";
  const gridElements = document.querySelectorAll(".fin-grid");
  for (const gridEl of gridElements) {
    gridEl.classList.remove(appState.darkMode ? lightClass : darkClass);
    gridEl.classList.add(appState.darkMode ? darkClass : lightClass);
  }

  if (gridState.businessApi) {
    gridState.businessApi.refreshHeader();
    gridState.businessApi.refreshCells({ force: true });
  }
  if (gridState.personalApi) {
    gridState.personalApi.refreshHeader();
    gridState.personalApi.refreshCells({ force: true });
  }
  if (gridState.portfolioApi) {
    gridState.portfolioApi.refreshHeader();
    gridState.portfolioApi.refreshCells({ force: true });
  }
}

function toggleTheme() {
  appState.darkMode = !appState.darkMode;
  applyThemeMode();
  updateThemeToggleText();
}

function setActiveView(viewKey) {
  appState.activeView = viewKey;
  for (const key of VIEW_KEYS) {
    const isActive = key === viewKey;
    refs.views[key].classList.toggle("hidden", !isActive);
    refs.nav[key].classList.toggle("menu-item-active", isActive);
    refs.nav[key].setAttribute("aria-current", isActive ? "page" : "false");
  }
  disableDeleteModesExcept(viewKey);
  updateGlobalBalance();
}

function updateSidebarToggleLabel() {
  const collapsed = refs.sidebar.classList.contains("sidebar-collapsed");
  refs.sidebarToggle.textContent = collapsed ? ">" : "<";
}

function toggleSidebar() {
  refs.sidebar.classList.toggle("sidebar-collapsed");
  updateSidebarToggleLabel();
}

function setupShellEvents() {
  refs.sidebarToggle.addEventListener("click", toggleSidebar);
  refs.themeToggle.addEventListener("click", toggleTheme);
  refs.globalSearch.addEventListener("input", (event) => applyGlobalSearch(event.target.value));
  const bindMonthInput = (inputEl, applyFilter) => {
    if (!inputEl || typeof applyFilter !== "function") return;
    const handle = () => {
      const rawValue = String(inputEl.value || "");
      const parsedMonth = parseMonthFromInput(rawValue);
      applyFilter(parsedMonth, { inputMonth: parsedMonth });
    };

    // Mantem leitura consistente para input type="month" e fallback text.
    inputEl.addEventListener("change", handle);
    inputEl.addEventListener("input", handle);
    inputEl.addEventListener("blur", handle);
  };

  bindMonthInput(refs.businessMonthInput, setBusinessMonthFilter);
  refs.businessMonthClear.addEventListener("click", () => {
    setBusinessMonthFilter("");
  });
  refs.businessPrevMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.businessMonthFilter, -1);
    setBusinessMonthFilter(next, { inputMonth: next });
  });
  refs.businessNextMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.businessMonthFilter, 1);
    setBusinessMonthFilter(next, { inputMonth: next });
  });

  bindMonthInput(refs.personalMonthInput, setPersonalMonthFilter);
  refs.personalMonthClear.addEventListener("click", () => {
    setPersonalMonthFilter("");
  });
  refs.personalPrevMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.personalMonthFilter, -1);
    setPersonalMonthFilter(next, { inputMonth: next });
  });
  refs.personalNextMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.personalMonthFilter, 1);
    setPersonalMonthFilter(next, { inputMonth: next });
  });

  bindMonthInput(refs.investmentMonthInput, setInvestmentMonthFilter);
  refs.investmentMonthClear.addEventListener("click", () => {
    setInvestmentMonthFilter("");
  });
  refs.investmentPrevMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.investmentMonthFilter, -1);
    setInvestmentMonthFilter(next, { inputMonth: next });
  });
  refs.investmentNextMonth.addEventListener("click", () => {
    const next = shiftMonthValue(appState.investmentMonthFilter, 1);
    setInvestmentMonthFilter(next, { inputMonth: next });
  });

  refs.checkUpdatesButton.addEventListener("click", async () => {
    const originalLabel = "Verificar update";
    refs.checkUpdatesButton.disabled = true;
    refs.checkUpdatesButton.textContent = "Verificando...";

    try {
      const result = await window.financeAPI.checkForUpdatesNow();
      if (result?.autoUpdateEnabled === false) {
        refs.checkUpdatesButton.textContent = "Modo manual";
      } else {
        refs.checkUpdatesButton.textContent = "Verificando...";
      }
    } catch (error) {
      console.error("Falha ao verificar update:", error);
      refs.checkUpdatesButton.textContent = "Falha";
    } finally {
      setTimeout(() => {
        refs.checkUpdatesButton.textContent = originalLabel;
        refs.checkUpdatesButton.disabled = false;
      }, 1800);
    }
  });

  refs.openUpdateButton.addEventListener("click", async () => {
    const pending = appState.pendingUpdate || {};
    const shouldInstallNow = appState.updateStatus === "downloaded" &&
      Boolean(pending.downloaded) &&
      Boolean(pending.canAutoInstall);

    if (shouldInstallNow) {
      refs.openUpdateButton.disabled = true;
      refs.openUpdateButton.textContent = "Reiniciando...";
      try {
        const result = await window.financeAPI.installDownloadedUpdateNow();
        if (!result?.ok) {
          refs.openUpdateButton.disabled = false;
          refs.openUpdateButton.textContent = "Instalar update";
        }
      } catch (error) {
        console.error("Falha ao instalar update:", error);
        refs.openUpdateButton.disabled = false;
        refs.openUpdateButton.textContent = "Instalar update";
      }
      return;
    }

    const url = resolveUpdateUrl(pending);
    if (!url) return;

    try {
      await window.financeAPI.openExternalLink(url);
    } catch (error) {
      console.error("Falha ao abrir link de update:", error);
    }
  });

  refs.dismissUpdateButton.addEventListener("click", () => {
    hideUpdateBanner();
  });

  refs.notificationToggle.addEventListener("click", (event) => {
    event.stopPropagation();
    setNotificationPanelOpen(!appState.notificationsOpen);
  });

  refs.notificationMarkRead.addEventListener("click", () => {
    appState.unreadNotifications = 0;
    renderNotificationCenter();
  });

  refs.riskBannerDismiss.addEventListener("click", () => {
    hideRiskBanner();
  });

  refs.nav.business.addEventListener("click", () => setActiveView("business"));
  refs.nav.personal.addEventListener("click", () => setActiveView("personal"));
  refs.nav.investments.addEventListener("click", () => setActiveView("investments"));

  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && appState.notificationsOpen) {
      setNotificationPanelOpen(false);
      return;
    }

    if (event.key !== "Delete") return;

    const target = event.target;
    const tag = String(target?.tagName || "").toLowerCase();
    const isTypingField = tag === "input" || tag === "textarea" || tag === "select";
    if (isTypingField) return;

    if (appState.activeView === "business") {
      deleteFocusedBusinessRow().catch((error) => console.error("Falha ao excluir linha empresarial:", error));
    } else if (appState.activeView === "personal") {
      deleteFocusedPersonalRow().catch((error) => console.error("Falha ao excluir linha pessoal:", error));
    } else if (appState.activeView === "investments") {
      deleteFocusedPortfolioRow().catch((error) => console.error("Falha ao excluir linha de investimentos:", error));
    }
  });

  document.addEventListener("click", (event) => {
    if (!appState.notificationsOpen) return;
    const target = event.target;
    if (!target || typeof target.closest !== "function") return;
    if (target.closest("#notificationPanel") || target.closest("#notificationToggle")) return;
    setNotificationPanelOpen(false);
  });
}

function investmentBaseValue(row) {
  return row.asset_type === "CDB"
    ? toNumber(row.application_amount)
    : toNumber(row.quantity) * toNumber(row.avg_price);
}

function investmentMarketValue(row) {
  return row.asset_type === "CDB"
    ? toNumber(row.current_value)
    : toNumber(row.quantity) * toNumber(row.current_value);
}

function investmentResultValue(row) {
  return investmentMarketValue(row) - investmentBaseValue(row);
}

function isSameMonth(dateIso) {
  return String(dateIso || "").startsWith(currentMonthPrefix());
}

function matchesMonth(dateIso, monthValue) {
  if (!monthValue) return true;
  return String(dateIso || "").startsWith(String(monthValue));
}

function getPersonalExpenseReferenceDate(row) {
  if (!row) return "";
  const status = uiToDbStatus(row.status);
  if (status === "Pago" && isValidIsoDate(row.paid_date)) {
    return String(row.paid_date);
  }
  return String(row.due_date || "");
}

function monthLabel(monthValue) {
  if (!monthValue) return "Todos";
  const normalized = extractMonth(monthValue);
  if (!normalized) return "Todos";
  const [yearText, monthText] = String(normalized).split("-");
  const year = Number(yearText);
  const monthNumber = Number(monthText);
  if (!Number.isFinite(year) || !Number.isFinite(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    return String(monthValue);
  }
  return `${MONTH_NAMES_SHORT_PTBR[monthNumber - 1]}/${year}`;
}

function monthChipLabel(monthValue) {
  return monthValue
    ? `Referencia: ${monthLabel(monthValue)}`
    : "Referencia: Todos os meses";
}

function shiftMonthValue(monthValue, offset) {
  const base = extractMonth(monthValue) || currentMonthPrefix();
  const [yearText, monthText] = String(base).split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (!Number.isFinite(year) || !Number.isFinite(month)) {
    return currentMonthPrefix();
  }

  const probe = new Date(year, month - 1, 1);
  probe.setMonth(probe.getMonth() + Number(offset || 0));
  return `${probe.getFullYear()}-${String(probe.getMonth() + 1).padStart(2, "0")}`;
}

function defaultIsoDateForMonthFilter(monthFilter) {
  const month = extractMonth(monthFilter);
  if (!month) return todayISO();

  const today = todayISO();
  if (extractMonth(today) === month) return today;
  return `${month}-01`;
}

function updateMonthFilterLabels() {
  if (refs.businessMonthLabel) {
    refs.businessMonthLabel.textContent = monthChipLabel(appState.businessMonthFilter);
  }
  if (refs.personalMonthLabel) {
    refs.personalMonthLabel.textContent = monthChipLabel(appState.personalMonthFilter);
  }
  if (refs.investmentMonthLabel) {
    refs.investmentMonthLabel.textContent = monthChipLabel(appState.investmentMonthFilter);
  }
}

function extractMonth(value) {
  const text = String(value || "");
  const match = text.match(/^(\d{4}-\d{2})/);
  return match ? match[1] : "";
}

function getInvestmentMonthKey(row) {
  if (!row) return "";
  return extractMonth(row.application_date || row.updated_at);
}

function sortMonthsDesc(list) {
  return [...list].sort((a, b) => b.localeCompare(a));
}

function renderMonthTabs(container, months, activeValue, onSelect) {
  if (!container) return;
  const values = ["", ...sortMonthsDesc(months)];
  container.innerHTML = "";

  for (const month of values) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "month-tab";
    if (month === activeValue) {
      button.classList.add("month-tab-active");
    }
    button.textContent = monthLabel(month);
    button.addEventListener("click", () => onSelect(month));
    container.appendChild(button);
  }
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function decoratePieRows(rows) {
  let incomeIdx = 0;
  let expenseIdx = 0;

  const sorted = [...rows].sort((a, b) => {
    if (a.kind !== b.kind) return a.kind === "income" ? -1 : 1;
    return b.value - a.value;
  });

  return sorted.map((row) => {
    if (row.kind === "income") {
      const color = PIE_COLORS_INCOME[incomeIdx % PIE_COLORS_INCOME.length];
      incomeIdx += 1;
      return { ...row, color };
    }

    const color = PIE_COLORS_EXPENSE[expenseIdx % PIE_COLORS_EXPENSE.length];
    expenseIdx += 1;
    return { ...row, color };
  });
}

function looksNumericLabel(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  return /^[-+\d\s.,/:%R$r$]+$/.test(text);
}

function prettifySourceLabel(value) {
  const text = String(value || "").trim();
  if (!text) return "Receita";

  const normalized = text.toUpperCase();
  const sourceMap = {
    SALARIO: "Salario",
    RENDA_EXTRA: "Renda extra",
    DIVIDENDOS: "Dividendos",
    JUROS: "Juros",
    ALUGUEIS: "Alugueis",
    OUTROS: "Outros"
  };

  if (sourceMap[normalized]) return sourceMap[normalized];
  return text.replace(/_/g, " ");
}

function resolvePersonalIncomeLabel(row) {
  const description = sanitizeCategory(row?.description);
  if (description && !looksNumericLabel(description)) {
    return description;
  }

  const source = sanitizeCategory(row?.source);
  return prettifySourceLabel(source);
}

function disambiguateRepeatedLabels(rows) {
  const counter = new Map();

  for (const row of rows) {
    const key = String(row.label || "").toLowerCase();
    if (!key) continue;
    counter.set(key, (counter.get(key) || 0) + 1);
  }

  return rows.map((row) => {
    const key = String(row.label || "").toLowerCase();
    const repeats = counter.get(key) || 0;
    if (repeats <= 1) return row;

    const suffix = row.kind === "income" ? " (ganho)" : " (gasto)";
    return { ...row, label: `${row.label}${suffix}` };
  });
}

function buildBusinessPieRows() {
  const rows = filteredBusinessRows();
  const grouped = new Map();

  for (const row of rows) {
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;

    const isIncome = row.entry_type === "Entrada";
    const category = sanitizeCategory(row.category) || "Sem categoria";
    const key = `${isIncome ? "income" : "expense"}|${category.toLowerCase()}`;

    const existing = grouped.get(key);
    if (existing) {
      existing.value += amount;
      continue;
    }

    grouped.set(key, {
      kind: isIncome ? "income" : "expense",
      label: category,
      value: amount
    });
  }

  const clean = Array.from(grouped.values()).filter((item) => item.value > 0);
  return decoratePieRows(disambiguateRepeatedLabels(clean));
}

function buildPersonalPieRows() {
  const grouped = new Map();

  const incomes = appState.personalIncomes.filter((row) => {
    if (!appState.personalMonthFilter) return true;
    return matchesMonth(row.income_date, appState.personalMonthFilter);
  });

  for (const row of incomes) {
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;

    const label = resolvePersonalIncomeLabel(row);
    const key = `income|${label.toLowerCase()}`;
    const existing = grouped.get(key);
    if (existing) {
      existing.value += amount;
      continue;
    }

    grouped.set(key, {
      kind: "income",
      label,
      value: amount
    });
  }

  const expenses = appState.personalExpenses.filter((row) => {
    if (!appState.personalMonthFilter) return true;
    return matchesMonth(getPersonalExpenseReferenceDate(row), appState.personalMonthFilter);
  });

  for (const row of expenses) {
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;

    const category = sanitizeCategory(row.category) || "Sem categoria";
    const key = `expense|${category.toLowerCase()}`;
    const existing = grouped.get(key);
    if (existing) {
      existing.value += amount;
      continue;
    }

    grouped.set(key, {
      kind: "expense",
      label: category,
      value: amount
    });
  }

  const clean = Array.from(grouped.values()).filter((item) => item.value > 0);
  return decoratePieRows(disambiguateRepeatedLabels(clean));
}

function renderPieChart(chartEl, legendEl, totalEl, rows) {
  if (!chartEl || !legendEl || !totalEl) return;

  const data = (rows || []).filter((item) => toNumber(item.value) > 0);
  const incomeTotal = data
    .filter((item) => item.kind === "income")
    .reduce((sum, item) => sum + toNumber(item.value), 0);
  const expenseTotal = data
    .filter((item) => item.kind === "expense")
    .reduce((sum, item) => sum + toNumber(item.value), 0);
  const total = incomeTotal + expenseTotal;
  const balance = incomeTotal - expenseTotal;

  if (total <= 0 || !data.length) {
    chartEl.style.background = "conic-gradient(#27272a 0 100%)";
    totalEl.innerHTML = '<span class="pie-total-label">Sem dados</span>';
    legendEl.innerHTML = '<p class="pie-empty">Sem movimentacoes para exibir.</p>';
    return;
  }

  const segments = [];
  let accumulated = 0;

  for (const item of data) {
    const start = accumulated;
    const delta = (toNumber(item.value) / total) * 100;
    accumulated += delta;
    segments.push(`${item.color} ${start.toFixed(3)}% ${accumulated.toFixed(3)}%`);
  }

  chartEl.style.background = `conic-gradient(${segments.join(", ")})`;
  const balanceLabel = "Saldo";
  const balanceValue = `${balance >= 0 ? "+" : "-"} ${BRL.format(Math.abs(balance))}`;
  totalEl.innerHTML = `
    <span class="pie-total-label">${escapeHtml(balanceLabel)}</span>
    <span class="pie-total-value">${escapeHtml(balanceValue)}</span>
  `;
  fitPieTotalValue(totalEl);

  legendEl.innerHTML = data
    .map((item) => {
      const amount = toNumber(item.value);
      const pct = total > 0 ? (amount / total) * 100 : 0;
      return `
        <div class="pie-legend-item">
          <span class="pie-legend-dot" style="background:${escapeHtml(item.color)}"></span>
          <span class="pie-legend-label" title="${escapeHtml(item.label)}">${escapeHtml(item.label)}</span>
          <span class="pie-legend-value">${escapeHtml(BRL.format(amount))} (${pct.toFixed(1)}%)</span>
        </div>
      `;
    })
    .join("");
}

function fitPieTotalValue(totalEl) {
  if (!totalEl) return;
  const valueEl = totalEl.querySelector(".pie-total-value");
  if (!valueEl) return;

  const containerWidth = Math.max(42, Math.floor(totalEl.clientWidth * 0.54));
  let fontSize = 14;
  const minFontSize = 10;

  valueEl.style.fontSize = `${fontSize}px`;
  valueEl.style.whiteSpace = "nowrap";

  while (fontSize > minFontSize && valueEl.scrollWidth > containerWidth) {
    fontSize -= 0.5;
    valueEl.style.fontSize = `${fontSize}px`;
  }
}

function getPreviousMonthKey(monthValue) {
  const month = extractMonth(monthValue);
  if (!month) return "";
  const [yearText, monthText] = month.split("-");
  const year = Number(yearText);
  const monthNumber = Number(monthText);
  if (!Number.isFinite(year) || !Number.isFinite(monthNumber)) return "";

  const date = new Date(year, monthNumber - 1, 1);
  date.setMonth(date.getMonth() - 1);
  const prevYear = date.getFullYear();
  const prevMonth = String(date.getMonth() + 1).padStart(2, "0");
  return `${prevYear}-${prevMonth}`;
}

function resolveComparisonMonth(monthFilter, availableMonths) {
  const normalizedFilter = extractMonth(monthFilter);
  if (normalizedFilter) return normalizedFilter;
  const valid = uniqueCategories(availableMonths || []);
  const sorted = sortMonthsDesc(valid.filter((item) => /^\d{4}-\d{2}$/.test(item)));
  return sorted[0] || currentMonthPrefix();
}

function computeBusinessMonthTotals(monthKey) {
  let income = 0;
  let expense = 0;
  let count = 0;

  for (const row of appState.businessRows) {
    if (!matchesMonth(row.entry_date, monthKey)) continue;
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;
    count += 1;
    if (row.entry_type === "Entrada") income += amount;
    else expense += amount;
  }

  return {
    income,
    expense,
    balance: income - expense,
    count
  };
}

function computePersonalMonthTotals(monthKey) {
  let income = 0;
  let expense = 0;
  let count = 0;

  for (const row of appState.personalIncomes) {
    if (!matchesMonth(row.income_date, monthKey)) continue;
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;
    income += amount;
    count += 1;
  }

  for (const row of appState.personalExpenses) {
    if (!matchesMonth(getPersonalExpenseReferenceDate(row), monthKey)) continue;
    const amount = toNumber(row.amount);
    if (amount <= 0) continue;
    expense += amount;
    count += 1;
  }

  return {
    income,
    expense,
    balance: income - expense,
    count
  };
}

function formatPercentPtBr(value) {
  return Number(value).toLocaleString("pt-BR", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  });
}

function compareMetric(current, previous, preferLower) {
  const delta = current - previous;
  const almostZero = Math.abs(delta) < 0.005;
  const percentage = Math.abs(previous) > 0 ? (delta / Math.abs(previous)) * 100 : null;

  if (almostZero) {
    return {
      className: "compare-neutral",
      text: `${BRL.format(0)} (0,0%) estavel`
    };
  }

  const isBetter = preferLower ? delta < 0 : delta > 0;
  const className = isBetter ? "compare-better" : "compare-worse";
  const direction = delta > 0 ? "+" : "-";
  const pctText = percentage === null ? "base 0" : `${direction}${formatPercentPtBr(Math.abs(percentage))}%`;
  const verdict = isBetter ? "melhor" : "pior";
  const deltaMoney = `${direction} ${BRL.format(Math.abs(delta))}`;

  return {
    className,
    text: `${deltaMoney} (${pctText}) ${verdict}`
  };
}

function renderMonthlyComparison(container, options) {
  if (!container) return;

  const {
    title = "Comparativo mensal",
    currentMonth,
    previousMonth,
    currentTotals,
    previousTotals,
    labels
  } = options;

  if (!currentMonth) {
    container.innerHTML = `
      <p class="pie-comparison-title">${escapeHtml(title)}</p>
      <p class="pie-comparison-empty">Sem dados suficientes para comparar.</p>
    `;
    return;
  }

  if (!previousMonth) {
    container.innerHTML = `
      <p class="pie-comparison-title">${escapeHtml(title)}</p>
      <p class="pie-comparison-month">${escapeHtml(monthLabel(currentMonth))}</p>
      <p class="pie-comparison-empty">Ainda nao existe mes anterior para comparacao.</p>
    `;
    return;
  }

  const currentHasData = currentTotals.count > 0;
  const previousHasData = previousTotals.count > 0;

  if (!previousHasData) {
    container.innerHTML = `
      <p class="pie-comparison-title">${escapeHtml(title)}</p>
      <p class="pie-comparison-month">${escapeHtml(`${monthLabel(currentMonth)} vs ${monthLabel(previousMonth)}`)}</p>
      <p class="pie-comparison-empty">Sem lancamentos no mes anterior.</p>
    `;
    return;
  }

  if (!currentHasData) {
    container.innerHTML = `
      <p class="pie-comparison-title">${escapeHtml(title)}</p>
      <p class="pie-comparison-month">${escapeHtml(`${monthLabel(currentMonth)} vs ${monthLabel(previousMonth)}`)}</p>
      <p class="pie-comparison-empty">Sem lancamentos no mes selecionado.</p>
    `;
    return;
  }

  const income = compareMetric(currentTotals.income, previousTotals.income, false);
  const expense = compareMetric(currentTotals.expense, previousTotals.expense, true);
  const balance = compareMetric(currentTotals.balance, previousTotals.balance, false);

  container.innerHTML = `
    <p class="pie-comparison-title">${escapeHtml(title)}</p>
    <p class="pie-comparison-month">${escapeHtml(`${monthLabel(currentMonth)} vs ${monthLabel(previousMonth)}`)}</p>
    <div class="pie-comparison-list">
      <div class="pie-comparison-item">
        <span class="pie-comparison-label">${escapeHtml(labels.income)}</span>
        <span class="pie-comparison-value ${escapeHtml(income.className)}">${escapeHtml(income.text)}</span>
      </div>
      <div class="pie-comparison-item">
        <span class="pie-comparison-label">${escapeHtml(labels.expense)}</span>
        <span class="pie-comparison-value ${escapeHtml(expense.className)}">${escapeHtml(expense.text)}</span>
      </div>
      <div class="pie-comparison-item">
        <span class="pie-comparison-label">${escapeHtml(labels.balance)}</span>
        <span class="pie-comparison-value ${escapeHtml(balance.className)}">${escapeHtml(balance.text)}</span>
      </div>
    </div>
  `;
}

function refreshBusinessMonthlyComparison() {
  const months = new Set();
  for (const row of appState.businessRows) {
    const month = extractMonth(row.entry_date);
    if (month) months.add(month);
  }
  if (appState.businessMonthFilter) {
    months.add(appState.businessMonthFilter);
  }

  const currentMonth = resolveComparisonMonth(appState.businessMonthFilter, Array.from(months));
  const previousMonth = getPreviousMonthKey(currentMonth);
  const currentTotals = computeBusinessMonthTotals(currentMonth);
  const previousTotals = previousMonth ? computeBusinessMonthTotals(previousMonth) : { income: 0, expense: 0, balance: 0, count: 0 };

  renderMonthlyComparison(refs.businessPieComparison, {
    title: "Comparativo mensal",
    currentMonth,
    previousMonth,
    currentTotals,
    previousTotals,
    labels: {
      income: "Entradas",
      expense: "Saidas",
      balance: "Saldo"
    }
  });
}

function refreshPersonalMonthlyComparison() {
  const months = new Set();
  for (const row of appState.personalIncomes) {
    const month = extractMonth(row.income_date);
    if (month) months.add(month);
  }
  for (const row of appState.personalExpenses) {
    const month = extractMonth(getPersonalExpenseReferenceDate(row));
    if (month) months.add(month);
  }
  if (appState.personalMonthFilter) {
    months.add(appState.personalMonthFilter);
  }

  const currentMonth = resolveComparisonMonth(appState.personalMonthFilter, Array.from(months));
  const previousMonth = getPreviousMonthKey(currentMonth);
  const currentTotals = computePersonalMonthTotals(currentMonth);
  const previousTotals = previousMonth ? computePersonalMonthTotals(previousMonth) : { income: 0, expense: 0, balance: 0, count: 0 };

  renderMonthlyComparison(refs.personalPieComparison, {
    title: "Comparativo mensal",
    currentMonth,
    previousMonth,
    currentTotals,
    previousTotals,
    labels: {
      income: "Receitas",
      expense: "Gastos",
      balance: "Saldo"
    }
  });
}

function refreshBusinessPieChart() {
  renderPieChart(
    refs.businessPieChart,
    refs.businessPieLegend,
    refs.businessPieTotal,
    buildBusinessPieRows()
  );
  refreshBusinessMonthlyComparison();
}

function refreshPersonalPieChart() {
  renderPieChart(
    refs.personalPieChart,
    refs.personalPieLegend,
    refs.personalPieTotal,
    buildPersonalPieRows()
  );
  refreshPersonalMonthlyComparison();
}

function getPersonalIncomeScopeMonth() {
  return appState.personalMonthFilter || currentMonthPrefix();
}

function getPersonalReferenceMonth() {
  return extractMonth(appState.personalMonthFilter) || currentMonthPrefix();
}

function formatMonthScopeLabel(monthValue) {
  const month = extractMonth(monthValue);
  if (!month) return "Mes atual";
  const [year, monthPart] = month.split("-");
  return `${monthPart}/${year}`;
}

function getPersonalIncomesByScopeMonth() {
  const month = getPersonalIncomeScopeMonth();
  return appState.personalIncomes.filter((row) => matchesMonth(row.income_date, month));
}

function renderPersonalIncomeList() {
  if (!refs.personalIncomeList || !refs.personalIncomeScope) return;

  const month = getPersonalIncomeScopeMonth();
  refs.personalIncomeScope.textContent = `Receitas ${formatMonthScopeLabel(month)}`;

  const rows = getPersonalIncomesByScopeMonth();
  if (!rows.length) {
    refs.personalIncomeList.innerHTML = '<p class="income-empty">Nenhuma receita lancada.</p>';
    return;
  }

  refs.personalIncomeList.innerHTML = rows
    .map((row) => {
      const id = Number(row.id);
      const label = resolvePersonalIncomeLabel(row);
      return `
        <span class="income-chip">
          <span class="income-chip-label">${escapeHtml(label)} - ${escapeHtml(BRL.format(toNumber(row.amount)))}</span>
          <button
            type="button"
            class="income-chip-remove"
            data-income-delete-id="${Number.isFinite(id) ? id : ""}"
            title="Excluir receita"
            aria-label="Excluir receita"
          >
            x
          </button>
        </span>
      `;
    })
    .join("");
}

function computeBusinessSummary() {
  const today = todayISO();
  let sales = 0;
  let expenses = 0;
  let balance = 0;

  for (const row of appState.businessRows) {
    const amount = toNumber(row.amount);
    const isIncome = row.entry_type === "Entrada";
    balance += isIncome ? amount : -amount;
    if (row.entry_date === today) {
      if (isIncome) sales += amount;
      else expenses += amount;
    }
  }

  return { sales, expenses, balance };
}

function computePersonalSummary() {
  const today = todayISO();
  const scopeMonth = getPersonalReferenceMonth();
  const incomeMonth = appState.personalIncomes
    .filter((row) => matchesMonth(row.income_date, scopeMonth))
    .reduce((acc, row) => acc + toNumber(row.amount), 0);

  const expenseMonth = appState.personalExpenses
    .filter((row) => matchesMonth(getPersonalExpenseReferenceDate(row), scopeMonth))
    .reduce((acc, row) => acc + toNumber(row.amount), 0);

  const openRows = appState.personalExpenses.filter((row) =>
    row.status === "Pendente" &&
    matchesMonth(row.due_date, scopeMonth)
  );
  const overdueRows = openRows.filter((row) => row.due_date < today);
  const openAmount = openRows.reduce((acc, row) => acc + toNumber(row.amount), 0);

  return {
    scopeMonth,
    incomeMonth,
    expenseMonth,
    openAmount,
    openCount: openRows.length,
    overdueCount: overdueRows.length
  };
}

function computeInvestmentSummary() {
  let invested = 0;
  let patrimony = 0;
  for (const row of appState.portfolioRows) {
    invested += investmentBaseValue(row);
    patrimony += investmentMarketValue(row);
  }
  const result = patrimony - invested;
  const progress = appState.portfolioGoal > 0 ? (patrimony / appState.portfolioGoal) * 100 : 0;
  return {
    invested,
    patrimony,
    result,
    progress: Math.max(0, Math.min(100, progress))
  };
}

function refreshBusinessCards() {
  const summary = computeBusinessSummary();
  refs.businessSalesCard.textContent = BRL.format(summary.sales);
  refs.businessExpensesCard.textContent = BRL.format(summary.expenses);
  refs.businessBalanceCard.textContent = BRL.format(summary.balance);
  refs.businessBalanceCard.classList.toggle("money-positive", summary.balance >= 0);
  refs.businessBalanceCard.classList.toggle("money-negative", summary.balance < 0);
  refreshBusinessPieChart();
  updateGlobalBalance();
}

function refreshPersonalCards() {
  const summary = computePersonalSummary();
  refs.personalIncomeCard.textContent = BRL.format(summary.incomeMonth);
  refs.personalExpenseCard.textContent = BRL.format(summary.expenseMonth);
  refs.personalOpenCard.textContent = BRL.format(summary.openAmount);
  refs.personalOpenMeta.textContent = `${monthLabel(summary.scopeMonth)} - ${summary.openCount} abertas / ${summary.overdueCount} vencidas`;
  refreshPersonalPieChart();
  renderPersonalIncomeList();
  updateGlobalBalance();
}

function refreshInvestmentCards() {
  const summary = computeInvestmentSummary();
  refs.investPatrimonyCard.textContent = BRL.format(summary.patrimony);
  refs.investedCard.textContent = BRL.format(summary.invested);
  refs.investResultCard.textContent = BRL.format(summary.result);
  refs.investResultCard.classList.toggle("money-positive", summary.result >= 0);
  refs.investResultCard.classList.toggle("money-negative", summary.result < 0);
  refs.goalProgressBar.style.width = `${summary.progress}%`;
  refs.goalProgressText.textContent = `${summary.progress.toFixed(1)}% da meta (${BRL.format(appState.portfolioGoal)})`;
  updateGlobalBalance();
}

function updateGlobalBalance() {
  const business = computeBusinessSummary().balance;
  const personalSummary = computePersonalSummary();
  const personal = personalSummary.incomeMonth - personalSummary.expenseMonth;
  const investments = computeInvestmentSummary().patrimony;

  let current = 0;
  if (appState.activeView === "business") current = business;
  else if (appState.activeView === "personal") current = personal;
  else current = investments;

  refs.sidebarBalance.textContent = BRL.format(current);
  refs.topbarBalance.textContent = BRL.format(current);
}

function setBusinessMonthFilter(monthValue, options = {}) {
  const normalized = extractMonth(monthValue);
  appState.businessMonthFilter = normalized;
  if (refs.businessMonthInput) {
    const inputMonth = extractMonth(options.inputMonth || options.inputDate || "");
    const currentInputMonth = parseMonthFromInput(refs.businessMonthInput.value);
    const keepInputMonthFromOption =
      !!inputMonth &&
      (!normalized || inputMonth === normalized);
    const keepCurrentInputMonth =
      !!normalized &&
      !!currentInputMonth &&
      currentInputMonth === normalized;

    if (!normalized) {
      refs.businessMonthInput.value = "";
    } else if (keepInputMonthFromOption) {
      refs.businessMonthInput.value = inputMonth;
    } else if (keepCurrentInputMonth) {
      refs.businessMonthInput.value = currentInputMonth;
    } else {
      refs.businessMonthInput.value = monthToDateInputValue(normalized);
    }
  }
  updateMonthFilterLabels();
  if (options.render !== false) renderBusinessRows();
}

function setPersonalMonthFilter(monthValue, options = {}) {
  const normalized = extractMonth(monthValue);
  appState.personalMonthFilter = normalized;
  if (refs.personalMonthInput) {
    const inputMonth = extractMonth(options.inputMonth || options.inputDate || "");
    const currentInputMonth = parseMonthFromInput(refs.personalMonthInput.value);
    const keepInputMonthFromOption =
      !!inputMonth &&
      (!normalized || inputMonth === normalized);
    const keepCurrentInputMonth =
      !!normalized &&
      !!currentInputMonth &&
      currentInputMonth === normalized;

    if (!normalized) {
      refs.personalMonthInput.value = "";
    } else if (keepInputMonthFromOption) {
      refs.personalMonthInput.value = inputMonth;
    } else if (keepCurrentInputMonth) {
      refs.personalMonthInput.value = currentInputMonth;
    } else {
      refs.personalMonthInput.value = monthToDateInputValue(normalized);
    }
  }
  updateMonthFilterLabels();
  if (options.render !== false) renderPersonalRows();
  renderPersonalIncomeList();
}

function setInvestmentMonthFilter(monthValue, options = {}) {
  const normalized = extractMonth(monthValue);
  appState.investmentMonthFilter = normalized;
  if (refs.investmentMonthInput) {
    const inputMonth = extractMonth(options.inputMonth || options.inputDate || "");
    const currentInputMonth = parseMonthFromInput(refs.investmentMonthInput.value);
    const keepInputMonthFromOption =
      !!inputMonth &&
      (!normalized || inputMonth === normalized);
    const keepCurrentInputMonth =
      !!normalized &&
      !!currentInputMonth &&
      currentInputMonth === normalized;

    if (!normalized) {
      refs.investmentMonthInput.value = "";
    } else if (keepInputMonthFromOption) {
      refs.investmentMonthInput.value = inputMonth;
    } else if (keepCurrentInputMonth) {
      refs.investmentMonthInput.value = currentInputMonth;
    } else {
      refs.investmentMonthInput.value = monthToDateInputValue(normalized);
    }
  }
  updateMonthFilterLabels();
  if (options.render !== false) renderPortfolioRows();
}

function compareTextPtBr(a, b) {
  return String(a || "").localeCompare(String(b || ""), "pt-BR", {
    sensitivity: "base",
    ignorePunctuation: true
  });
}

function compareNumber(a, b) {
  const nA = toNumber(a);
  const nB = toNumber(b);
  if (nA === nB) return 0;
  return nA < nB ? -1 : 1;
}

function compareIsoDate(a, b) {
  return String(a || "").localeCompare(String(b || ""));
}

function stableIdentityCompare(a, b) {
  const idA = Number.isFinite(Number(a?.id)) ? Number(a.id) : Number.MAX_SAFE_INTEGER;
  const idB = Number.isFinite(Number(b?.id)) ? Number(b.id) : Number.MAX_SAFE_INTEGER;
  if (idA !== idB) return idA - idB;
  return compareTextPtBr(a?.__tmpId, b?.__tmpId);
}

function installmentIndexCompare(a, b) {
  const idxA = Math.max(1, Number.isFinite(Number(a?.installment_index)) ? Math.floor(Number(a.installment_index)) : 1);
  const idxB = Math.max(1, Number.isFinite(Number(b?.installment_index)) ? Math.floor(Number(b.installment_index)) : 1);
  if (idxA !== idxB) return idxA - idxB;
  return stableIdentityCompare(a, b);
}

function businessStableCompare(a, b) {
  const dateCompare = compareIsoDate(a?.entry_date, b?.entry_date);
  if (dateCompare !== 0) return dateCompare;
  return installmentIndexCompare(a, b);
}

function personalStableCompare(a, b) {
  const dueCompare = compareIsoDate(getPersonalExpenseReferenceDate(a), getPersonalExpenseReferenceDate(b));
  if (dueCompare !== 0) return dueCompare;
  return installmentIndexCompare(a, b);
}

function portfolioStableCompare(a, b) {
  const monthCompare = compareIsoDate(getInvestmentMonthKey(a), getInvestmentMonthKey(b));
  if (monthCompare !== 0) return monthCompare;
  return stableIdentityCompare(a, b);
}

function sortBusinessRows(rows) {
  const mode = appState.businessSortMode;
  return [...rows].sort((a, b) => {
    let cmp = 0;

    if (mode === "date_desc") cmp = compareIsoDate(b.entry_date, a.entry_date);
    else if (mode === "date_asc") cmp = compareIsoDate(a.entry_date, b.entry_date);
    else if (mode === "amount_desc") cmp = compareNumber(b.amount, a.amount);
    else if (mode === "amount_asc") cmp = compareNumber(a.amount, b.amount);
    else if (mode === "category_asc") cmp = compareTextPtBr(a.category, b.category);
    else if (mode === "type_group") cmp = compareTextPtBr(dbToUiBusinessType(a.entry_type), dbToUiBusinessType(b.entry_type));

    if (cmp !== 0) return cmp;
    return businessStableCompare(a, b);
  });
}

function sortPersonalRows(rows) {
  const mode = appState.personalSortMode;
  return [...rows].sort((a, b) => {
    let cmp = 0;
    if (mode === "due_desc") cmp = compareIsoDate(getPersonalExpenseReferenceDate(b), getPersonalExpenseReferenceDate(a));
    else if (mode === "due_asc") cmp = compareIsoDate(getPersonalExpenseReferenceDate(a), getPersonalExpenseReferenceDate(b));
    else if (mode === "amount_desc") cmp = compareNumber(b.amount, a.amount);
    else if (mode === "amount_asc") cmp = compareNumber(a.amount, b.amount);
    else if (mode === "category_asc") cmp = compareTextPtBr(a.category, b.category);
    else if (mode === "status_group") {
      const statusA = dbToUiStatus(a.status) === "Pendente" ? 0 : 1;
      const statusB = dbToUiStatus(b.status) === "Pendente" ? 0 : 1;
      cmp = statusA - statusB;
    }

    if (cmp !== 0) return cmp;
    return personalStableCompare(a, b);
  });
}

function sortPortfolioRows(rows) {
  const mode = appState.investmentSortMode;
  return [...rows].sort((a, b) => {
    let cmp = 0;

    if (mode === "result_desc") cmp = compareNumber(investmentResultValue(b), investmentResultValue(a));
    else if (mode === "result_asc") cmp = compareNumber(investmentResultValue(a), investmentResultValue(b));
    else if (mode === "market_desc") cmp = compareNumber(investmentMarketValue(b), investmentMarketValue(a));
    else if (mode === "market_asc") cmp = compareNumber(investmentMarketValue(a), investmentMarketValue(b));
    else if (mode === "asset_asc") cmp = compareTextPtBr(a.asset_label, b.asset_label);
    else if (mode === "type_asc") cmp = compareTextPtBr(dbToUiAssetType(a.asset_type), dbToUiAssetType(b.asset_type));

    if (cmp !== 0) return cmp;
    return portfolioStableCompare(a, b);
  });
}

function getRowListByView(viewKey) {
  if (viewKey === "business") return appState.businessRows;
  if (viewKey === "personal") return appState.personalExpenses;
  return appState.portfolioRows;
}

function scheduleNewRowFlashReset(viewKey, rowIdentity, timeoutMs = 1100) {
  setTimeout(() => {
    const list = getRowListByView(viewKey);
    const idx = indexByIdentity(list, rowIdentity);
    if (idx < 0) return;
    if (!list[idx].__isNew) return;
    list[idx] = { ...list[idx], __isNew: false };
    const api = getGridApiByView(viewKey);
    api?.redrawRows?.();
  }, timeoutMs);
}

function markRowAsNew(viewKey, row) {
  if (!row) return;
  row.__isNew = true;
  scheduleNewRowFlashReset(viewKey, row);
}

function filteredBusinessRows() {
  let rows = appState.businessRows;
  if (appState.businessMonthFilter) {
    rows = rows.filter((row) => matchesMonth(row.entry_date, appState.businessMonthFilter));
  }

  return sortBusinessRows(rows);
}

function filteredPersonalRows() {
  const key = appState.personalFilter;
  let rows = appState.personalExpenses;
  const today = todayISO();

  if (appState.personalMonthFilter) {
    rows = rows.filter((row) => matchesMonth(getPersonalExpenseReferenceDate(row), appState.personalMonthFilter));
  }

  if (key === "pending") rows = rows.filter((row) => row.status === "Pendente");
  else if (key === "overdue") rows = rows.filter((row) => row.status === "Pendente" && row.due_date < today);
  else if (key === "month") rows = rows.filter((row) => isSameMonth(getPersonalExpenseReferenceDate(row)));

  return sortPersonalRows(rows);
}

function filteredPortfolioRows() {
  let rows = appState.portfolioRows;
  if (appState.investmentMonthFilter) {
    rows = rows.filter((row) => {
      const month = getInvestmentMonthKey(row);
      return matchesMonth(month, appState.investmentMonthFilter);
    });
  }

  return sortPortfolioRows(rows);
}

function renderBusinessMonthTabs() {
  const months = new Set();
  for (const row of appState.businessRows) {
    const month = extractMonth(row.entry_date);
    if (month) months.add(month);
  }
  if (appState.businessMonthFilter) {
    months.add(appState.businessMonthFilter);
  }
  renderMonthTabs(
    refs.businessMonthTabs,
    Array.from(months),
    appState.businessMonthFilter,
    (month) => {
      setBusinessMonthFilter(month);
    }
  );
}

function renderBusinessRows() {
  if (!gridState.businessApi) return;
  gridState.businessApi.setGridOption("rowData", filteredBusinessRows());
  renderBusinessMonthTabs();
  refreshBusinessPieChart();
}

function renderPersonalMonthTabs() {
  const months = new Set();
  for (const row of appState.personalExpenses) {
    const month = extractMonth(getPersonalExpenseReferenceDate(row));
    if (month) months.add(month);
  }
  if (appState.personalMonthFilter) {
    months.add(appState.personalMonthFilter);
  }
  renderMonthTabs(
    refs.personalMonthTabs,
    Array.from(months),
    appState.personalMonthFilter,
    (month) => {
      setPersonalMonthFilter(month);
    }
  );
}

function renderPersonalRows() {
  if (!gridState.personalApi) return;
  gridState.personalApi.setGridOption("rowData", filteredPersonalRows());
  renderPersonalMonthTabs();
  refreshPersonalPieChart();
  refreshPersonalCards();
}

function renderInvestmentMonthTabs() {
  const months = new Set();
  for (const row of appState.portfolioRows) {
    const month = getInvestmentMonthKey(row);
    if (month) months.add(month);
  }
  if (appState.investmentMonthFilter) {
    months.add(appState.investmentMonthFilter);
  }
  renderMonthTabs(
    refs.investmentMonthTabs,
    Array.from(months),
    appState.investmentMonthFilter,
    (month) => {
      setInvestmentMonthFilter(month);
    }
  );
}

function renderPortfolioRows() {
  if (!gridState.portfolioApi) return;
  gridState.portfolioApi.setGridOption("rowData", filteredPortfolioRows());
  renderInvestmentMonthTabs();
}

function setupSortEvents() {
  refs.businessSortMode?.addEventListener("change", (event) => {
    appState.businessSortMode = String(event?.target?.value || "date_desc");
    renderBusinessRows();
  });

  refs.personalSortMode?.addEventListener("change", (event) => {
    appState.personalSortMode = String(event?.target?.value || "due_asc");
    renderPersonalRows();
  });

  refs.investmentSortMode?.addEventListener("change", (event) => {
    appState.investmentSortMode = String(event?.target?.value || "result_desc");
    renderPortfolioRows();
  });
}

function syncSortControls() {
  if (refs.businessSortMode) refs.businessSortMode.value = appState.businessSortMode;
  if (refs.personalSortMode) refs.personalSortMode.value = appState.personalSortMode;
  if (refs.investmentSortMode) refs.investmentSortMode.value = appState.investmentSortMode;
}

function updateFilterButtons() {
  const map = {
    pending: refs.filterPending,
    overdue: refs.filterOverdue,
    month: refs.filterMonth,
    none: refs.filterClear
  };

  const buttons = [refs.filterPending, refs.filterOverdue, refs.filterMonth, refs.filterClear];
  for (const btn of buttons) btn.classList.remove("filter-btn-active");
  map[appState.personalFilter].classList.add("filter-btn-active");
}

function setPersonalFilter(key) {
  appState.personalFilter = key;
  updateFilterButtons();
  renderPersonalRows();
}

function setupPersonalFilterEvents() {
  refs.filterPending.addEventListener("click", () => setPersonalFilter("pending"));
  refs.filterOverdue.addEventListener("click", () => setPersonalFilter("overdue"));
  refs.filterMonth.addEventListener("click", () => setPersonalFilter("month"));
  refs.filterClear.addEventListener("click", () => setPersonalFilter("none"));
}

function initBusinessGrid() {
  const columnDefs = [
    buildDeleteSelectorColumnDef(),
    {
      headerName: "Data",
      field: "entry_date",
      editable: true,
      minWidth: 120,
      maxWidth: 140,
      cellEditor: "agDateStringCellEditor",
      cellEditorParams: { min: "1900-01-01", max: "2100-12-31" },
      valueParser: (params) =>
        normalizeIsoDate(
          params.newValue,
          params.oldValue || todayISO(),
          params.oldValue || todayISO()
        ),
      valueFormatter: (params) => formatDateBR(params.value)
    },
    {
      headerName: "Descricao",
      field: "description",
      editable: true,
      minWidth: 260,
      flex: 1.2
    },
    {
      headerName: "Tipo",
      field: "entry_type",
      editable: true,
      minWidth: 120,
      maxWidth: 140,
      cellEditor: CategorySelectEditor,
      cellEditorParams: { values: ["Entrada", "Saida"], center: true },
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => dbToUiBusinessType(params.data.entry_type),
      valueSetter: (params) => {
        params.data.entry_type = uiToDbBusinessType(params.newValue);
        return true;
      }
    },
    {
      headerName: "Categoria",
      field: "category",
      editable: true,
      minWidth: 180,
      headerClass: "header-center",
      cellClass: "cell-center",
      cellEditor: CategorySelectEditor,
      cellEditorParams: () => ({ values: appState.businessCategories, center: true })
    },
    {
      headerName: "Valor",
      field: "amount",
      editable: true,
      minWidth: 130,
      maxWidth: 170,
      valueParser: (params) => toNumber(params.newValue),
      valueFormatter: (params) => BRL.format(toNumber(params.value)),
      cellClassRules: {
        "money-positive": (params) => params.data?.entry_type === "Entrada",
        "money-negative": (params) => params.data?.entry_type === "Saida"
      }
    },
    {
      headerName: "Parcela",
      field: "installment_progress",
      editable: false,
      minWidth: 110,
      maxWidth: 120,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => {
        const total = toInstallmentTotal(params.data.installment_total);
        const index = Math.min(
          total,
          Math.max(1, Number.isFinite(Number(params.data.installment_index)) ? Math.floor(Number(params.data.installment_index)) : 1)
        );
        return `${index}/${total}`;
      }
    },
    {
      headerName: "Meses",
      field: "installment_total",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1;
      },
      minWidth: 95,
      maxWidth: 105,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueSetter: (params) => {
        const parsed = toInstallmentTotal(params.newValue);
        params.data.installment_total = parsed;
        if (parsed > 1) {
          params.data.is_recurring = 0;
          params.data.recurrence_interval_months = 1;
        }
        return true;
      },
      valueFormatter: (params) => String(toInstallmentTotal(params.value))
    },
    {
      headerName: "Recorrente",
      field: "is_recurring",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1;
      },
      minWidth: 120,
      maxWidth: 130,
      cellEditor: CategorySelectEditor,
      cellEditorParams: { values: ["Nao", "Sim"], center: true },
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => dbToUiRecurring(params.data.is_recurring),
      valueSetter: (params) => {
        params.data.is_recurring = uiToDbRecurring(params.newValue);
        if (!params.data.is_recurring) {
          params.data.recurrence_interval_months = 1;
        }
        return true;
      }
    },
    {
      headerName: "Intervalo",
      field: "recurrence_interval_months",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1 && uiToDbRecurring(params.data.is_recurring) === 1;
      },
      minWidth: 105,
      maxWidth: 115,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueSetter: (params) => {
        params.data.recurrence_interval_months = toInstallmentTotal(params.newValue || 1);
        return true;
      },
      valueFormatter: (params) => `${toInstallmentTotal(params.value || 1)}m`
    }
  ];

  const gridOptions = {
    rowData: filteredBusinessRows(),
    columnDefs,
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
    rowClassRules: {
      "row-new-flash": (params) => Boolean(params.data?.__isNew)
    },
    defaultColDef: {
      editable: true,
      sortable: true,
      filter: true,
      resizable: true
    },
    singleClickEdit: true,
    stopEditingWhenCellsLoseFocus: true,
    enterNavigatesVertically: true,
    enterNavigatesVerticallyAfterEdit: true,
    undoRedoCellEditing: true,
    undoRedoCellEditingLimit: 30,
    onCellValueChanged: async (params) => {
      if (gridState.savingBusiness || !params.data) return;

      const changed = normalizeBusinessRow(params.data);
      const businessCategoryResult = addCategoryToList(appState.businessCategories, changed.category);
      if (businessCategoryResult.value) {
        changed.category = businessCategoryResult.value;
      }
      if (businessCategoryResult.added) {
        persistCategorySetting("business_categories", appState.businessCategories);
      }
      upsertIntoList(appState.businessRows, changed);
      renderBusinessRows();
      refreshBusinessCards();

      gridState.savingBusiness = true;
      try {
        const savedRaw = await window.financeAPI.saveBusinessEntry(changed);
        const generatedRowsRaw = Array.isArray(savedRaw?.generated_rows) ? savedRaw.generated_rows : [];

        if (generatedRowsRaw.length > 0) {
          removeFromListByIdentity(appState.businessRows, changed);
          for (const generatedRaw of generatedRowsRaw) {
            const generated = normalizeBusinessRow(generatedRaw);
            upsertIntoList(appState.businessRows, generated);
          }

          if (!appState.businessMonthFilter) {
            const firstRow = generatedRowsRaw[0];
            const firstMonth = extractMonth(firstRow?.entry_date);
            if (firstMonth) {
              setBusinessMonthFilter(firstMonth, { render: false });
            }
          }
        } else {
          if (changed.__tmpId && !savedRaw.__tmpId) {
            savedRaw.__tmpId = changed.__tmpId;
          }
          const saved = normalizeBusinessRow(savedRaw);
          upsertIntoList(appState.businessRows, saved);
        }
        renderBusinessRows();
        refreshBusinessCards();
      } finally {
        gridState.savingBusiness = false;
      }
    }
  };

  gridState.businessApi = agGrid.createGrid(byId("businessGrid"), gridOptions);
}

function initPersonalGrid() {
  const columnDefs = [
    buildDeleteSelectorColumnDef(),
    {
      headerName: "Vencimento",
      field: "due_date",
      editable: true,
      minWidth: 130,
      maxWidth: 150,
      cellEditor: "agDateStringCellEditor",
      cellEditorParams: { min: "1900-01-01", max: "2100-12-31" },
      valueParser: (params) =>
        normalizeIsoDate(
          params.newValue,
          params.oldValue || todayISO(),
          params.oldValue || todayISO()
        ),
      valueFormatter: (params) => formatDateBR(params.value)
    },
    {
      headerName: "Descricao",
      field: "description",
      editable: true,
      minWidth: 250,
      flex: 1.2
    },
    {
      headerName: "Categoria",
      field: "category",
      editable: true,
      minWidth: 160,
      headerClass: "header-center",
      cellClass: "cell-center",
      cellEditor: CategorySelectEditor,
      cellEditorParams: () => ({ values: appState.personalCategories, center: true })
    },
    {
      headerName: "Valor",
      field: "amount",
      editable: true,
      minWidth: 130,
      valueParser: (params) => toNumber(params.newValue),
      valueFormatter: (params) => BRL.format(toNumber(params.value)),
      cellClass: "money-negative"
    },
    {
      headerName: "Parcela",
      field: "installment_progress",
      editable: false,
      minWidth: 110,
      maxWidth: 120,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => {
        const total = toInstallmentTotal(params.data.installment_total);
        const index = Math.min(
          total,
          Math.max(1, Number.isFinite(Number(params.data.installment_index)) ? Math.floor(Number(params.data.installment_index)) : 1)
        );
        return `${index}/${total}`;
      }
    },
    {
      headerName: "Meses",
      field: "installment_total",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1;
      },
      minWidth: 95,
      maxWidth: 105,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueSetter: (params) => {
        const parsed = toInstallmentTotal(params.newValue);
        params.data.installment_total = parsed;
        if (parsed > 1) {
          params.data.is_recurring = 0;
          params.data.recurrence_interval_months = 1;
        }
        return true;
      },
      valueFormatter: (params) => String(toInstallmentTotal(params.value))
    },
    {
      headerName: "Recorrente",
      field: "is_recurring",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1;
      },
      minWidth: 120,
      maxWidth: 130,
      cellEditor: CategorySelectEditor,
      cellEditorParams: { values: ["Nao", "Sim"], center: true },
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => dbToUiRecurring(params.data.is_recurring),
      valueSetter: (params) => {
        params.data.is_recurring = uiToDbRecurring(params.newValue);
        if (!params.data.is_recurring) {
          params.data.recurrence_interval_months = 1;
        }
        return true;
      }
    },
    {
      headerName: "Intervalo",
      field: "recurrence_interval_months",
      editable: (params) => {
        const idx = Number(params.data.installment_index || 1);
        const total = toInstallmentTotal(params.data.installment_total);
        return idx <= 1 && total <= 1 && uiToDbRecurring(params.data.is_recurring) === 1;
      },
      minWidth: 105,
      maxWidth: 115,
      headerClass: "header-center",
      cellClass: "cell-center",
      valueSetter: (params) => {
        params.data.recurrence_interval_months = toInstallmentTotal(params.newValue || 1);
        return true;
      },
      valueFormatter: (params) => `${toInstallmentTotal(params.value || 1)}m`
    },
    {
      headerName: "Status",
      field: "status",
      editable: true,
      minWidth: 120,
      maxWidth: 140,
      cellEditor: CategorySelectEditor,
      cellEditorParams: { values: ["Pendente", "Pago"], center: true },
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => dbToUiStatus(params.data.status),
      valueSetter: (params) => {
        params.data.status = uiToDbStatus(params.newValue);
        if (params.data.status === "Pago" && !params.data.paid_date) {
          params.data.paid_date = todayISO();
        }
        if (params.data.status === "Pendente") {
          params.data.paid_date = null;
        }
        return true;
      }
    }
  ];

  const gridOptions = {
    rowData: filteredPersonalRows(),
    columnDefs,
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
    rowClassRules: {
      "row-new-flash": (params) => Boolean(params.data?.__isNew)
    },
    defaultColDef: {
      editable: true,
      sortable: true,
      filter: true,
      resizable: true
    },
    singleClickEdit: true,
    stopEditingWhenCellsLoseFocus: true,
    enterNavigatesVertically: true,
    enterNavigatesVerticallyAfterEdit: true,
    undoRedoCellEditing: true,
    undoRedoCellEditingLimit: 30,
    onCellValueChanged: async (params) => {
      if (gridState.savingPersonal || !params.data) return;

      const changed = normalizePersonalExpenseRow(params.data);
      const personalCategoryResult = addCategoryToList(appState.personalCategories, changed.category);
      if (personalCategoryResult.value) {
        changed.category = personalCategoryResult.value;
      }
      if (personalCategoryResult.added) {
        persistCategorySetting("personal_categories", appState.personalCategories);
      }
      upsertIntoList(appState.personalExpenses, changed);
      refreshPersonalCards();

      gridState.savingPersonal = true;
      try {
        const savedRaw = await window.financeAPI.savePersonalExpense(changed);
        const generatedRowsRaw = Array.isArray(savedRaw?.generated_rows) ? savedRaw.generated_rows : [];

        if (generatedRowsRaw.length > 0) {
          removeFromListByIdentity(appState.personalExpenses, changed);
          for (const generatedRaw of generatedRowsRaw) {
            const generated = normalizePersonalExpenseRow(generatedRaw);
            upsertIntoList(appState.personalExpenses, generated);
          }

          if (!appState.personalMonthFilter) {
            const firstRow = generatedRowsRaw[0];
            const firstMonth = extractMonth(firstRow?.due_date);
            if (firstMonth) {
              setPersonalMonthFilter(firstMonth, { render: false });
            }
          }
        } else {
          if (changed.__tmpId && !savedRaw.__tmpId) {
            savedRaw.__tmpId = changed.__tmpId;
          }
          const saved = normalizePersonalExpenseRow(savedRaw);
          upsertIntoList(appState.personalExpenses, saved);
        }
        renderPersonalRows();
        refreshPersonalCards();
      } finally {
        gridState.savingPersonal = false;
      }
    }
  };

  gridState.personalApi = agGrid.createGrid(byId("personalGrid"), gridOptions);
}

function initPortfolioGrid() {
  const columnDefs = [
    buildDeleteSelectorColumnDef(),
    {
      headerName: "Ativo/Banco",
      field: "asset_label",
      editable: true,
      minWidth: 180,
      flex: 1.2
    },
    {
      headerName: "Tipo",
      field: "asset_type",
      editable: true,
      minWidth: 110,
      maxWidth: 130,
      cellEditor: CategorySelectEditor,
      cellEditorParams: { values: ["Acao", "FII", "CDB"], center: true },
      headerClass: "header-center",
      cellClass: "cell-center",
      valueGetter: (params) => dbToUiAssetType(params.data.asset_type),
      valueSetter: (params) => {
        params.data.asset_type = uiToDbAssetType(params.newValue);
        if (params.data.asset_type === "CDB") {
          params.data.ticker = "";
          params.data.quantity = 0;
          params.data.avg_price = 0;
          params.data.current_unit_price = 0;
          params.data.application_amount = params.data.application_amount || 0;
          params.data.rate_percent = params.data.rate_percent || 100;
          params.data.application_date = params.data.application_date || todayISO();
        }
        return true;
      }
    },
    {
      headerName: "Qtd/Aplicacao",
      field: "qty_application",
      editable: true,
      minWidth: 150,
      valueGetter: (params) =>
        params.data.asset_type === "CDB"
          ? toNumber(params.data.application_amount)
          : toNumber(params.data.quantity),
      valueSetter: (params) => {
        const parsed = toNumber(params.newValue);
        if (params.data.asset_type === "CDB") params.data.application_amount = parsed;
        else params.data.quantity = parsed;
        return true;
      },
      valueFormatter: (params) =>
        params.data.asset_type === "CDB"
          ? BRL.format(toNumber(params.value))
          : NUMBER_BR.format(toNumber(params.value))
    },
    {
      headerName: "Preco Medio/Taxa",
      field: "avg_rate",
      editable: true,
      minWidth: 170,
      valueGetter: (params) =>
        params.data.asset_type === "CDB"
          ? toNumber(params.data.rate_percent)
          : toNumber(params.data.avg_price),
      valueSetter: (params) => {
        const parsed = toNumber(params.newValue);
        if (params.data.asset_type === "CDB") params.data.rate_percent = parsed;
        else params.data.avg_price = parsed;
        return true;
      },
      valueFormatter: (params) =>
        params.data.asset_type === "CDB"
          ? `${toNumber(params.value).toFixed(2)}% CDI`
          : BRL.format(toNumber(params.value))
    },
    {
      headerName: "Valor Atual",
      field: "current_value",
      editable: true,
      minWidth: 160,
      valueGetter: (params) => toNumber(params.data.current_value),
      valueSetter: (params) => {
        const parsed = toNumber(params.newValue);
        params.data.current_value = parsed;
        if (params.data.asset_type !== "CDB") {
          params.data.current_unit_price = parsed;
        }
        return true;
      },
      valueFormatter: (params) =>
        params.data.asset_type === "CDB"
          ? BRL.format(toNumber(params.value))
          : BRL.format(toNumber(params.value)),
      cellClassRules: {
        "flash-up": (params) => params.data.flash_direction === "up",
        "flash-down": (params) => params.data.flash_direction === "down"
      }
    },
    {
      headerName: "Resultado",
      field: "result_value",
      editable: false,
      minWidth: 140,
      valueGetter: (params) => investmentResultValue(params.data),
      valueFormatter: (params) => BRL.format(toNumber(params.value)),
      cellClassRules: {
        "money-positive": (params) => toNumber(params.value) >= 0,
        "money-negative": (params) => toNumber(params.value) < 0
      }
    }
  ];

  const gridOptions = {
    rowData: filteredPortfolioRows(),
    columnDefs,
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
    rowClassRules: {
      "row-new-flash": (params) => Boolean(params.data?.__isNew)
    },
    defaultColDef: {
      editable: true,
      sortable: true,
      filter: true,
      resizable: true
    },
    singleClickEdit: true,
    stopEditingWhenCellsLoseFocus: true,
    enterNavigatesVertically: true,
    enterNavigatesVerticallyAfterEdit: true,
    undoRedoCellEditing: true,
    undoRedoCellEditingLimit: 30,
    onCellValueChanged: async (params) => {
      if (gridState.savingPortfolio || !params.data) return;

      const changed = normalizePortfolioRow(params.data);
      upsertIntoList(appState.portfolioRows, changed);
      renderPortfolioRows();
      refreshInvestmentCards();

      gridState.savingPortfolio = true;
      try {
        const savedRaw = await window.financeAPI.savePortfolioPosition(changed);
        if (changed.__tmpId && !savedRaw.__tmpId) {
          savedRaw.__tmpId = changed.__tmpId;
        }
        const saved = normalizePortfolioRow(savedRaw);
        upsertIntoList(appState.portfolioRows, saved);
        renderPortfolioRows();
        refreshInvestmentCards();
      } finally {
        gridState.savingPortfolio = false;
      }
    }
  };

  gridState.portfolioApi = agGrid.createGrid(byId("portfolioGrid"), gridOptions);
}

function getFocusedRowData(api) {
  if (!api) return null;
  const focused = api.getFocusedCell?.();
  if (!focused || typeof focused.rowIndex !== "number") return null;
  const rowNode = api.getDisplayedRowAtIndex(focused.rowIndex);
  return rowNode?.data || null;
}

function focusFirstRowForEditing(api, colKey) {
  if (!api) return;
  const firstRow = api.getDisplayedRowAtIndex(0);
  if (!firstRow) return;
  if (typeof api.deselectAll === "function") api.deselectAll();
  firstRow.setSelected?.(true, true);
  api.ensureIndexVisible(0, "top");
  api.setFocusedCell(0, colKey);
  api.startEditingCell({ rowIndex: 0, colKey });
}

function focusRowForEditingByIdentity(api, rowIdentity, colKey) {
  if (!api || !rowIdentity) return;

  let targetNode = null;
  api.forEachNode((node) => {
    if (targetNode) return;
    const data = node?.data || {};

    if (rowIdentity.id && Number(data.id) === Number(rowIdentity.id)) {
      targetNode = node;
      return;
    }

    if (rowIdentity.__tmpId && data.__tmpId && data.__tmpId === rowIdentity.__tmpId) {
      targetNode = node;
    }
  });

  if (!targetNode || typeof targetNode.rowIndex !== "number" || targetNode.rowIndex < 0) {
    focusFirstRowForEditing(api, colKey);
    return;
  }

  if (typeof api.deselectAll === "function") api.deselectAll();
  targetNode.setSelected?.(true, true);
  api.ensureIndexVisible(targetNode.rowIndex, "middle");
  api.setFocusedCell(targetNode.rowIndex, colKey);
  api.startEditingCell({ rowIndex: targetNode.rowIndex, colKey });
}

function removeFromListByIdentity(list, row) {
  const idx = indexByIdentity(list, row);
  if (idx >= 0) list.splice(idx, 1);
}

async function deleteFocusedBusinessRow() {
  const row = getFocusedRowData(gridState.businessApi);
  if (!row) return;

  if (row.id) {
    const result = await window.financeAPI.deleteBusinessEntry(row.id);
    if (!result?.ok) return;
  }

  removeFromListByIdentity(appState.businessRows, row);
  renderBusinessRows();
  refreshBusinessCards();
}

async function deleteFocusedPersonalRow() {
  const row = getFocusedRowData(gridState.personalApi);
  if (!row) return;

  if (row.id) {
    const result = await window.financeAPI.deletePersonalExpense(row.id);
    if (!result?.ok) return;
  }

  removeFromListByIdentity(appState.personalExpenses, row);
  renderPersonalRows();
  refreshPersonalCards();
}

async function deleteFocusedPortfolioRow() {
  const row = getFocusedRowData(gridState.portfolioApi);
  if (!row) return;

  if (row.id) {
    const result = await window.financeAPI.deletePortfolioPosition(row.id);
    if (!result?.ok) return;
  }

  removeFromListByIdentity(appState.portfolioRows, row);
  renderPortfolioRows();
  refreshInvestmentCards();
}

async function deletePersonalIncomeById(incomeId) {
  const parsed = Number(incomeId);
  if (!Number.isFinite(parsed) || parsed <= 0) return false;

  const result = await window.financeAPI.deletePersonalIncome(parsed);
  if (!result?.ok) return false;

  const idx = appState.personalIncomes.findIndex((row) => Number(row.id) === parsed);
  if (idx >= 0) appState.personalIncomes.splice(idx, 1);

  refreshPersonalCards();
  return true;
}

function getSelectedRows(api) {
  if (!api || typeof api.getSelectedRows !== "function") return [];
  const selected = api.getSelectedRows();
  return Array.isArray(selected) ? selected.filter(Boolean) : [];
}

async function deleteRowsBatch(rows, remoteDeleteFn, list) {
  for (const row of rows) {
    if (!row) continue;

    let canRemoveLocal = true;
    if (row.id) {
      try {
        const result = await remoteDeleteFn(row.id);
        canRemoveLocal = Boolean(result?.ok);
      } catch (error) {
        canRemoveLocal = false;
        console.error("Falha ao excluir linha:", error);
      }
    }

    if (canRemoveLocal) {
      removeFromListByIdentity(list, row);
    }
  }
}

async function deleteSelectedRowsFromView(viewKey, rows) {
  if (!rows.length) return;

  if (viewKey === "business") {
    await deleteRowsBatch(rows, window.financeAPI.deleteBusinessEntry, appState.businessRows);
    renderBusinessRows();
    refreshBusinessCards();
    return;
  }

  if (viewKey === "personal") {
    await deleteRowsBatch(rows, window.financeAPI.deletePersonalExpense, appState.personalExpenses);
    renderPersonalRows();
    refreshPersonalCards();
    return;
  }

  await deleteRowsBatch(rows, window.financeAPI.deletePortfolioPosition, appState.portfolioRows);
  renderPortfolioRows();
  refreshInvestmentCards();
}

async function handleDeleteButtonClick(viewKey) {
  const isMode = getDeleteMode(viewKey);
  if (!isMode) {
    disableDeleteModesExcept(viewKey);
    setDeleteMode(viewKey, true);
    return;
  }

  const api = getGridApiByView(viewKey);
  const selectedRows = getSelectedRows(api);

  if (!selectedRows.length) {
    setDeleteMode(viewKey, false);
    return;
  }

  await deleteSelectedRowsFromView(viewKey, selectedRows);
  setDeleteMode(viewKey, false);
}

function registerAddButtons() {
  refs.deleteBusinessRow.addEventListener("click", () => {
    handleDeleteButtonClick("business").catch((error) => console.error("Falha ao excluir linhas empresariais:", error));
  });

  refs.addBusinessRow.addEventListener("click", () => {
    const row = normalizeBusinessRow({
      entry_date: defaultIsoDateForMonthFilter(appState.businessMonthFilter),
      description: "Nova movimentacao",
      entry_type: "Entrada",
      category: appState.businessCategories[0] || "Venda Balcao",
      amount: 0,
      installment_total: 1,
      installment_index: 1,
      installment_group_id: null,
      is_recurring: 0,
      recurrence_interval_months: 1
    });
    markRowAsNew("business", row);

    upsertIntoList(appState.businessRows, row);
    renderBusinessRows();
    refreshBusinessCards();
    focusRowForEditingByIdentity(gridState.businessApi, row, "entry_date");
  });

  refs.deletePersonalRow.addEventListener("click", () => {
    handleDeleteButtonClick("personal").catch((error) => console.error("Falha ao excluir linhas pessoais:", error));
  });

  refs.addPersonalRow.addEventListener("click", () => {
    const row = normalizePersonalExpenseRow({
      due_date: defaultIsoDateForMonthFilter(appState.personalMonthFilter),
      description: "Novo registro",
      category: appState.personalCategories[0] || "Moradia",
      amount: 0,
      status: "Pendente",
      installment_total: 1,
      installment_index: 1,
      installment_group_id: null,
      is_recurring: 0,
      recurrence_interval_months: 1
    });
    markRowAsNew("personal", row);
    if (appState.personalFilter !== "none") {
      appState.personalFilter = "none";
      updateFilterButtons();
    }

    upsertIntoList(appState.personalExpenses, row);
    renderPersonalRows();
    refreshPersonalCards();
    focusRowForEditingByIdentity(gridState.personalApi, row, "due_date");
  });

  refs.deletePortfolioRow.addEventListener("click", () => {
    handleDeleteButtonClick("investments").catch((error) => console.error("Falha ao excluir linhas de investimentos:", error));
  });

  refs.addPortfolioRow.addEventListener("click", () => {
    const row = normalizePortfolioRow({
      asset_label: "Nova Posicao",
      asset_type: "ACAO",
      ticker: "NOVO",
      quantity: 0,
      avg_price: 0,
      current_unit_price: 0,
      current_value: 0,
      updated_at: defaultIsoDateForMonthFilter(appState.investmentMonthFilter)
    });
    markRowAsNew("investments", row);

    upsertIntoList(appState.portfolioRows, row);
    renderPortfolioRows();
    refreshInvestmentCards();
    focusRowForEditingByIdentity(gridState.portfolioApi, row, "asset_label");
  });

  refs.addBusinessCategoryButton.addEventListener("click", () => {
    const result = addCategoryToList(appState.businessCategories, refs.businessCategoryInput.value);
    if (!result.added) return;

    refs.businessCategoryInput.value = "";
    persistCategorySetting("business_categories", appState.businessCategories);
    if (gridState.businessApi) {
      gridState.businessApi.refreshCells({ force: true });
    }
  });

  refs.businessCategoryInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      refs.addBusinessCategoryButton.click();
    }
  });

  refs.addPersonalCategoryButton.addEventListener("click", () => {
    const result = addCategoryToList(appState.personalCategories, refs.personalCategoryInput.value);
    if (!result.added) return;

    refs.personalCategoryInput.value = "";
    persistCategorySetting("personal_categories", appState.personalCategories);
    if (gridState.personalApi) {
      gridState.personalApi.refreshCells({ force: true });
    }
  });

  refs.personalCategoryInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      refs.addPersonalCategoryButton.click();
    }
  });

  refs.addIncomeButton.addEventListener("click", async () => {
    const amount = toNumber(refs.incomeAmountInput.value);
    if (amount <= 0) return;

    const description = String(refs.incomeDescriptionInput.value || "").trim() || "Salario/Receita";
    const payload = {
      income_date: todayISO(),
      description,
      source: "SALARIO",
      amount,
      notes: ""
    };

    try {
      const saved = normalizePersonalIncomeRow(await window.financeAPI.addPersonalIncome(payload));
      appState.personalIncomes.unshift(saved);
      refs.incomeAmountInput.value = "";
      refs.incomeDescriptionInput.value = "";
      refreshPersonalCards();
    } catch (error) {
      console.error("Falha ao salvar receita:", error);
    }
  });

  refs.personalIncomeList.addEventListener("click", (event) => {
    const target = event.target;
    if (!target || typeof target.closest !== "function") return;

    const button = target.closest("[data-income-delete-id]");
    if (!button) return;

    const incomeId = Number(button.getAttribute("data-income-delete-id"));
    if (!Number.isFinite(incomeId) || incomeId <= 0) return;

    deletePersonalIncomeById(incomeId).catch((error) => {
      console.error("Falha ao excluir receita:", error);
    });
  });

  refs.saveGoalButton.addEventListener("click", async () => {
    const goal = toNumber(refs.goalInput.value);
    if (goal <= 0) return;

    try {
      await window.financeAPI.setUserSetting("meta_patrimonio", String(goal));
      appState.portfolioGoal = goal;
      refreshInvestmentCards();
    } catch (error) {
      console.error("Falha ao salvar meta:", error);
    }
  });
}

function applyPortfolioPolling(payload) {
  if (!payload?.updates || !gridState.portfolioApi) return;

  for (const update of payload.updates) {
    const id = Number(update.id);
    if (!id) continue;

    const local = appState.portfolioRows.find((row) => Number(row.id) === id);
    if (local) {
      local.current_value = toNumber(update.newValue);
      local.current_unit_price = toNumber(update.newUnitPrice);
      local.updated_at = update.quotedAt || "";
      local.flash_direction = update.direction === "down" ? "down" : "up";
    }

    let targetNode = null;
    gridState.portfolioApi.forEachNode((node) => {
      if (targetNode) return;
      if (Number(node.data?.id) === id) targetNode = node;
    });

    if (targetNode) {
      targetNode.setDataValue("current_value", toNumber(update.newValue));
      targetNode.setDataValue("current_unit_price", toNumber(update.newUnitPrice));
      targetNode.setDataValue("flash_direction", update.direction === "down" ? "down" : "up");
      gridState.portfolioApi.refreshCells({
        rowNodes: [targetNode],
        columns: ["current_value", "result_value"],
        force: true
      });

      setTimeout(() => {
        targetNode.setDataValue("flash_direction", "");
        gridState.portfolioApi.refreshCells({
          rowNodes: [targetNode],
          columns: ["current_value"],
          force: true
        });
      }, 1000);
    }
  }

  if (payload.quotedAt) {
    refs.portfolioUpdatedAt.textContent = new Date(payload.quotedAt).toLocaleString("pt-BR");
  }

  renderPortfolioRows();
  refreshInvestmentCards();
}

function applyInitialDataToState(initialData) {
  const goalFromSettings = toNumber(initialData?.userSettings?.meta_patrimonio);
  appState.portfolioGoal = goalFromSettings > 0
    ? goalFromSettings
    : toNumber(initialData?.meta?.portfolioGoal || 100000);

  appState.businessRows = (initialData?.businessEntries || []).map(normalizeBusinessRow);
  appState.personalIncomes = (initialData?.personalIncomes || []).map(normalizePersonalIncomeRow);
  appState.personalExpenses = (initialData?.personalExpenses || []).map(normalizePersonalExpenseRow);
  appState.portfolioRows = (initialData?.portfolioPositions || []).map(normalizePortfolioRow);
  initCategoryState(initialData?.userSettings || {});
  renderBillAlerts(initialData?.billAlerts || { bills: [], windowDays: 2 });
}

function applyDefaultMonthFiltersIfNeeded() {
  const month = currentMonthPrefix();

  if (!extractMonth(appState.businessMonthFilter)) {
    setBusinessMonthFilter(month, { render: false, inputMonth: month });
  }
  if (!extractMonth(appState.personalMonthFilter)) {
    setPersonalMonthFilter(month, { render: false, inputMonth: month });
  }
  if (!extractMonth(appState.investmentMonthFilter)) {
    setInvestmentMonthFilter(month, { render: false, inputMonth: month });
  }
}

function renderAllViews() {
  updateMonthFilterLabels();
  renderBusinessRows();
  setPersonalFilter(appState.personalFilter || "none");
  renderPortfolioRows();
  refreshBusinessCards();
  refreshPersonalCards();
  refreshInvestmentCards();
}

async function reloadDataFromDatabase(options = {}) {
  const { silent = false } = options;

  try {
    const initialData = await window.financeAPI.getInitialData();
    applyInitialDataToState(initialData);

    if (refs.goalInput) {
      refs.goalInput.value = appState.portfolioGoal;
    }

    renderAllViews();

    if (!silent) {
      pushNotification(
        "Dados atualizados",
        "Novos lancamentos recorrentes foram adicionados ao mes atual.",
        `reload-${Date.now()}`
      );
    }
  } catch (error) {
    console.error("Falha ao recarregar dados:", error);
  }
}

async function bootstrap() {
  initRefs();
  renderNotificationCenter();
  updateSidebarToggleLabel();
  setupShellEvents();
  setupPersonalFilterEvents();
  setupSortEvents();
  applyThemeMode();
  updateThemeToggleText();

  const initialData = await window.financeAPI.getInitialData();
  applyInitialDataToState(initialData);

  initBusinessGrid();
  initPersonalGrid();
  initPortfolioGrid();
  registerAddButtons();
  applyThemeMode();

  applyDefaultMonthFiltersIfNeeded();
  syncSortControls();
  refs.goalInput.value = appState.portfolioGoal;

  renderAllViews();
  setActiveView("business");

  if (typeof window.financeAPI.onUpdateState === "function") {
    window.financeAPI.onUpdateState((payload) => applyUpdateState(payload));
  } else {
    window.financeAPI.onUpdateAvailable((payload) => showUpdateBanner(payload));
  }
  window.financeAPI.onPortfolioUpdated((payload) => applyPortfolioPolling(payload));
  window.financeAPI.onBillAlerts((payload) => renderBillAlerts(payload));
  window.financeAPI.onRecurringRowsGenerated((payload) => {
    reloadDataFromDatabase({ silent: true });
    if (payload?.totalCreated) {
      pushNotification(
        "Recorrencias sincronizadas",
        `${payload.totalCreated} lancamento(s) recorrente(s) adicionado(s) automaticamente.`,
        `recurring-${payload.businessCreated || 0}-${payload.personalCreated || 0}-${Date.now()}`
      );
    }
  });
  window.financeAPI.getBillAlerts().then((payload) => renderBillAlerts(payload)).catch((error) => {
    console.error("Falha ao atualizar alertas de vencimento:", error);
  });
  window.financeAPI.checkForUpdatesNow().catch((error) => {
    console.error("Falha na verificacao inicial de update:", error);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  bootstrap().catch((error) => {
    console.error("Falha ao iniciar renderer:", error);
  });
});
