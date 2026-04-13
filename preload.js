const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("lookupAPI", {
  openDocumentDialog: () => ipcRenderer.invoke("dialog:open-document"),
  openPdfDialog: () => ipcRenderer.invoke("dialog:open-document"),
  openDocument: (filePath) => ipcRenderer.invoke("document:open", filePath),
  convertDocument: (filePath) => ipcRenderer.invoke("document:convert", filePath),
  isSupportedDocument: (filePath) => ipcRenderer.invoke("document:is-supported", filePath),
  savePdfDialog: (options) => ipcRenderer.invoke("dialog:save-pdf", options),
  confirmOverwrite: (options) => ipcRenderer.invoke("dialog:confirm-overwrite", options),
  readPdfFile: (filePath) => ipcRenderer.invoke("pdf:read-file", filePath),
  writePdfFile: (filePath, data) => ipcRenderer.invoke("pdf:write-file", { filePath, data }),
  printDocument: () => ipcRenderer.invoke("window:print"),
  copyText: (text) => ipcRenderer.invoke("clipboard:copy-text", text),
  getSettings: () => ipcRenderer.invoke("settings:get"),
  setLanguage: (language) => ipcRenderer.invoke("settings:set-language", language),

  toggleFullScreen: () => ipcRenderer.invoke("window:toggle-fullscreen"),
  setFullScreen: (enabled) => ipcRenderer.invoke("window:set-fullscreen", enabled),
  isFullScreen: () => ipcRenderer.invoke("window:is-fullscreen"),
  getAppVersion: () => ipcRenderer.invoke("app:get-version"),

  checkForUpdates: () => ipcRenderer.invoke("update:check"),
  installUpdateNow: () => ipcRenderer.invoke("update:install-now"),
  getUpdateConfig: () => ipcRenderer.invoke("update:get-config"),

  onSystemOpenFile: (callback) => {
    const listener = (_event, filePath) => callback(filePath);
    ipcRenderer.on("system-open-file", listener);
    return () => ipcRenderer.removeListener("system-open-file", listener);
  },
  onMenuAction: (callback) => {
    const listener = (_event, action) => callback(action);
    ipcRenderer.on("menu-action", listener);
    return () => ipcRenderer.removeListener("menu-action", listener);
  },
  onFullScreenChanged: (callback) => {
    const listener = (_event, isFullScreen) => callback(Boolean(isFullScreen));
    ipcRenderer.on("window-fullscreen-changed", listener);
    return () => ipcRenderer.removeListener("window-fullscreen-changed", listener);
  },
  onUpdateStatus: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("update-status", listener);
    return () => ipcRenderer.removeListener("update-status", listener);
  }
});
