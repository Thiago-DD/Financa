const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("financeAPI", {
  getInitialData: () => ipcRenderer.invoke("app:getInitialData"),
  saveBusinessEntry: (row) => ipcRenderer.invoke("db:upsertBusinessEntry", row),
  addPersonalIncome: (row) => ipcRenderer.invoke("db:addPersonalIncome", row),
  savePersonalExpense: (row) => ipcRenderer.invoke("db:upsertPersonalExpense", row),
  savePortfolioPosition: (row) => ipcRenderer.invoke("db:upsertPortfolioPosition", row),
  setUserSetting: (key, value) => ipcRenderer.invoke("db:setUserSetting", { key, value }),
  onPortfolioUpdated: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("portfolio:updated", listener);
    return () => ipcRenderer.removeListener("portfolio:updated", listener);
  }
});
