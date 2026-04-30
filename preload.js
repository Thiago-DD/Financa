const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("financeAPI", {
  getInitialData: () => ipcRenderer.invoke("app:getInitialData"),
  saveBusinessEntry: (row) => ipcRenderer.invoke("db:upsertBusinessEntry", row),
  deleteBusinessEntry: (id) => ipcRenderer.invoke("db:deleteBusinessEntry", id),
  addPersonalIncome: (row) => ipcRenderer.invoke("db:addPersonalIncome", row),
  savePersonalExpense: (row) => ipcRenderer.invoke("db:upsertPersonalExpense", row),
  deletePersonalExpense: (id) => ipcRenderer.invoke("db:deletePersonalExpense", id),
  savePortfolioPosition: (row) => ipcRenderer.invoke("db:upsertPortfolioPosition", row),
  deletePortfolioPosition: (id) => ipcRenderer.invoke("db:deletePortfolioPosition", id),
  setUserSetting: (key, value) => ipcRenderer.invoke("db:setUserSetting", { key, value }),
  checkForUpdatesNow: () => ipcRenderer.invoke("app:checkForUpdatesNow"),
  openExternalLink: (url) => ipcRenderer.invoke("app:openExternalLink", url),
  onUpdateAvailable: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("app:updateAvailable", listener);
    return () => ipcRenderer.removeListener("app:updateAvailable", listener);
  },
  onPortfolioUpdated: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("portfolio:updated", listener);
    return () => ipcRenderer.removeListener("portfolio:updated", listener);
  }
});
