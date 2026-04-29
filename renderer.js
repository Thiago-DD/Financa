const BRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const NUMBER_BR = new Intl.NumberFormat("pt-BR");

const VIEW_KEYS = ["business", "personal", "investments"];
const DEFAULT_BUSINESS_CATEGORIES = ["Sangria", "Fornecedor", "Venda Balcao"];
const DEFAULT_PERSONAL_CATEGORIES = ["Moradia", "Alimentacao", "Transporte", "Lazer"];

const appState = {
  activeView: "business",
  darkMode: true,
  portfolioGoal: 100000,
  businessRows: [],
  personalIncomes: [],
  personalExpenses: [],
  portfolioRows: [],
  personalFilter: "none",
  businessCategories: [...DEFAULT_BUSINESS_CATEGORIES],
  personalCategories: [...DEFAULT_PERSONAL_CATEGORIES]
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

function toNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value !== "string") return Number(value) || 0;
  const normalized = value.replace(/\./g, "").replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
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

function dbToUiBusinessType(value) {
  return value === "Saída" || value === "Saida" || value === "SAIDA"
    ? "Saida"
    : "Entrada";
}

function uiToDbBusinessType(value) {
  return String(value || "").toLowerCase().startsWith("s") ? "Saída" : "Entrada";
}

function dbToUiStatus(value) {
  return value === "Pago" || value === "PAGO" ? "Pago" : "Pendente";
}

function uiToDbStatus(value) {
  return String(value || "").toLowerCase().startsWith("pago") ? "Pago" : "Pendente";
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
  return ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    entry_date: row.entry_date || todayISO(),
    description: row.description || "",
    entry_type: uiToDbBusinessType(row.entry_type),
    category: row.category || "Venda Balcao",
    amount: toNumber(row.amount),
    notes: row.notes || ""
  });
}

function normalizePersonalExpenseRow(row) {
  const status = uiToDbStatus(row.status);
  return ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    due_date: row.due_date || todayISO(),
    description: row.description || "",
    category: row.category || "Moradia",
    amount: toNumber(row.amount),
    status,
    paid_date: status === "Pago" ? (row.paid_date || row.due_date || todayISO()) : null,
    notes: row.notes || ""
  });
}

function normalizePersonalIncomeRow(row) {
  return {
    id: row.id ?? null,
    income_date: row.income_date || todayISO(),
    description: row.description || "",
    source: row.source || "SALARIO",
    amount: toNumber(row.amount),
    notes: row.notes || ""
  };
}

function normalizePortfolioRow(row) {
  const type = uiToDbAssetType(row.asset_type);
  const clean = ensureTempId({
    id: row.id ?? null,
    __tmpId: row.__tmpId || null,
    asset_label: row.asset_label || "",
    asset_type: type,
    ticker: row.ticker ? String(row.ticker).toUpperCase() : "",
    quantity: type === "CDB" ? 0 : toNumber(row.quantity),
    avg_price: type === "CDB" ? 0 : toNumber(row.avg_price),
    application_amount: type === "CDB" ? toNumber(row.application_amount) : 0,
    rate_percent: type === "CDB" ? toNumber(row.rate_percent) : 0,
    cdi_annual_rate: type === "CDB" ? toNumber(row.cdi_annual_rate || 13.65) : 13.65,
    application_date: type === "CDB" ? (row.application_date || todayISO()) : null,
    current_unit_price: type === "CDB" ? 0 : toNumber(row.current_unit_price),
    current_value: toNumber(row.current_value),
    last_dividend: type === "CDB" ? 0 : toNumber(row.last_dividend),
    updated_at: row.updated_at || "",
    flash_direction: ""
  });

  return clean;
}

function indexByIdentity(list, row) {
  if (row.id) return list.findIndex((item) => Number(item.id) === Number(row.id));
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
  refs.globalSearch = byId("globalSearch");

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
  refs.addPersonalRow = byId("addPersonalRow");
  refs.addPortfolioRow = byId("addPortfolioRow");
  refs.businessCategoryInput = byId("businessCategoryInput");
  refs.addBusinessCategoryButton = byId("addBusinessCategoryButton");
  refs.personalCategoryInput = byId("personalCategoryInput");
  refs.addPersonalCategoryButton = byId("addPersonalCategoryButton");
  refs.addIncomeButton = byId("addIncomeButton");
  refs.incomeAmountInput = byId("incomeAmountInput");
  refs.incomeDescriptionInput = byId("incomeDescriptionInput");
  refs.goalInput = byId("goalInput");
  refs.saveGoalButton = byId("saveGoalButton");

  refs.businessSalesCard = byId("businessSalesCard");
  refs.businessExpensesCard = byId("businessExpensesCard");
  refs.businessBalanceCard = byId("businessBalanceCard");

  refs.personalIncomeCard = byId("personalIncomeCard");
  refs.personalExpenseCard = byId("personalExpenseCard");
  refs.personalOpenCard = byId("personalOpenCard");
  refs.personalOpenMeta = byId("personalOpenMeta");

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
    refs.views[key].classList.toggle("hidden", key !== viewKey);
    refs.nav[key].classList.toggle("menu-item-active", key === viewKey);
  }
  updateGlobalBalance();
}

function toggleSidebar() {
  refs.sidebar.classList.toggle("sidebar-collapsed");
}

function setupShellEvents() {
  refs.sidebarToggle.addEventListener("click", toggleSidebar);
  refs.themeToggle.addEventListener("click", toggleTheme);
  refs.globalSearch.addEventListener("input", (event) => applyGlobalSearch(event.target.value));

  refs.nav.business.addEventListener("click", () => setActiveView("business"));
  refs.nav.personal.addEventListener("click", () => setActiveView("personal"));
  refs.nav.investments.addEventListener("click", () => setActiveView("investments"));
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
  const incomeMonth = appState.personalIncomes
    .filter((row) => isSameMonth(row.income_date))
    .reduce((acc, row) => acc + toNumber(row.amount), 0);

  const expenseMonth = appState.personalExpenses
    .filter((row) => isSameMonth(row.due_date))
    .reduce((acc, row) => acc + toNumber(row.amount), 0);

  const openRows = appState.personalExpenses.filter((row) => row.status === "Pendente");
  const overdueRows = openRows.filter((row) => row.due_date < today);
  const openAmount = openRows.reduce((acc, row) => acc + toNumber(row.amount), 0);

  return {
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
  updateGlobalBalance();
}

function refreshPersonalCards() {
  const summary = computePersonalSummary();
  refs.personalIncomeCard.textContent = BRL.format(summary.incomeMonth);
  refs.personalExpenseCard.textContent = BRL.format(summary.expenseMonth);
  refs.personalOpenCard.textContent = BRL.format(summary.openAmount);
  refs.personalOpenMeta.textContent = `${summary.openCount} abertas / ${summary.overdueCount} vencidas`;
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
  const personal = computePersonalSummary().incomeMonth - computePersonalSummary().expenseMonth;
  const investments = computeInvestmentSummary().patrimony;

  let current = 0;
  if (appState.activeView === "business") current = business;
  else if (appState.activeView === "personal") current = personal;
  else current = investments;

  refs.sidebarBalance.textContent = BRL.format(current);
  refs.topbarBalance.textContent = BRL.format(current);
}

function filteredPersonalRows() {
  const key = appState.personalFilter;
  const allRows = appState.personalExpenses;
  const today = todayISO();

  if (key === "pending") return allRows.filter((row) => row.status === "Pendente");
  if (key === "overdue") return allRows.filter((row) => row.status === "Pendente" && row.due_date < today);
  if (key === "month") return allRows.filter((row) => isSameMonth(row.due_date));
  return allRows;
}

function renderPersonalRows() {
  if (!gridState.personalApi) return;
  gridState.personalApi.setGridOption("rowData", filteredPersonalRows());
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
    {
      headerName: "Data",
      field: "entry_date",
      editable: true,
      minWidth: 120,
      maxWidth: 140,
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
      cellEditor: "agSelectCellEditor",
      cellEditorParams: { values: ["Entrada", "Saida"] },
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
      cellEditor: "agSelectCellEditor",
      cellEditorParams: () => ({ values: appState.businessCategories })
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
        "money-negative": (params) => params.data?.entry_type === "Saída"
      }
    }
  ];

  const gridOptions = {
    rowData: appState.businessRows,
    columnDefs,
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
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
      refreshBusinessCards();

      gridState.savingBusiness = true;
      try {
        const saved = normalizeBusinessRow(await window.financeAPI.saveBusinessEntry(changed));
        params.node.setData(saved);
        upsertIntoList(appState.businessRows, saved);
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
    {
      headerName: "Vencimento",
      field: "due_date",
      editable: true,
      minWidth: 130,
      maxWidth: 150,
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
      cellEditor: "agSelectCellEditor",
      cellEditorParams: () => ({ values: appState.personalCategories })
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
      headerName: "Status",
      field: "status",
      editable: true,
      minWidth: 120,
      maxWidth: 140,
      cellEditor: "agSelectCellEditor",
      cellEditorParams: { values: ["Pendente", "Pago"] },
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
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
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
        const saved = normalizePersonalExpenseRow(await window.financeAPI.savePersonalExpense(changed));
        upsertIntoList(appState.personalExpenses, saved);
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
      cellEditor: "agSelectCellEditor",
      cellEditorParams: { values: ["Acao", "FII", "CDB"] },
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
    rowData: appState.portfolioRows,
    columnDefs,
    getRowId: (params) => String(params.data.id || params.data.__tmpId),
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
      refreshInvestmentCards();

      gridState.savingPortfolio = true;
      try {
      const saved = normalizePortfolioRow(await window.financeAPI.savePortfolioPosition(changed));
        params.node.setData(saved);
        upsertIntoList(appState.portfolioRows, saved);
        refreshInvestmentCards();
      } finally {
        gridState.savingPortfolio = false;
      }
    }
  };

  gridState.portfolioApi = agGrid.createGrid(byId("portfolioGrid"), gridOptions);
}

function registerAddButtons() {
  refs.addBusinessRow.addEventListener("click", () => {
    const row = normalizeBusinessRow({
      entry_date: todayISO(),
      description: "Nova movimentacao",
      entry_type: "Entrada",
      category: appState.businessCategories[0] || "Venda Balcao",
      amount: 0
    });
    gridState.businessApi.applyTransaction({ add: [row], addIndex: 0 });
    upsertIntoList(appState.businessRows, row);
    refreshBusinessCards();
  });

  refs.addPersonalRow.addEventListener("click", () => {
    const row = normalizePersonalExpenseRow({
      due_date: todayISO(),
      description: "Nova conta",
      category: appState.personalCategories[0] || "Moradia",
      amount: 0,
      status: "Pendente"
    });
    upsertIntoList(appState.personalExpenses, row);
    renderPersonalRows();
    refreshPersonalCards();
  });

  refs.addPortfolioRow.addEventListener("click", () => {
    const row = normalizePortfolioRow({
      asset_label: "Nova Posicao",
      asset_type: "ACAO",
      ticker: "NOVO",
      quantity: 0,
      avg_price: 0,
      current_unit_price: 0,
      current_value: 0
    });
    gridState.portfolioApi.applyTransaction({ add: [row], addIndex: 0 });
    upsertIntoList(appState.portfolioRows, row);
    refreshInvestmentCards();
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

  refreshInvestmentCards();
}

async function bootstrap() {
  initRefs();
  setupShellEvents();
  setupPersonalFilterEvents();
  applyThemeMode();
  updateThemeToggleText();

  const initialData = await window.financeAPI.getInitialData();
  const goalFromSettings = toNumber(initialData.userSettings?.meta_patrimonio);
  appState.portfolioGoal = goalFromSettings > 0
    ? goalFromSettings
    : toNumber(initialData.meta?.portfolioGoal || 100000);

  appState.businessRows = (initialData.businessEntries || []).map(normalizeBusinessRow);
  appState.personalIncomes = (initialData.personalIncomes || []).map(normalizePersonalIncomeRow);
  appState.personalExpenses = (initialData.personalExpenses || []).map(normalizePersonalExpenseRow);
  appState.portfolioRows = (initialData.portfolioPositions || []).map(normalizePortfolioRow);
  initCategoryState(initialData.userSettings || {});

  initBusinessGrid();
  initPersonalGrid();
  initPortfolioGrid();
  registerAddButtons();
  applyThemeMode();

  refs.goalInput.value = appState.portfolioGoal;

  setPersonalFilter("none");
  refreshBusinessCards();
  refreshPersonalCards();
  refreshInvestmentCards();
  setActiveView("business");

  window.financeAPI.onPortfolioUpdated((payload) => applyPortfolioPolling(payload));
}

document.addEventListener("DOMContentLoaded", () => {
  bootstrap().catch((error) => {
    console.error("Falha ao iniciar renderer:", error);
  });
});
