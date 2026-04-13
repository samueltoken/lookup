const { app, BrowserWindow, Menu, dialog, ipcMain } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");

let mainWindow = null;
let pendingPdfToOpen = extractPdfPath(process.argv);
let updateConfig = null;
let updateDownloaded = false;

const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
}

function extractPdfPath(argv) {
  if (!Array.isArray(argv)) {
    return null;
  }

  for (const arg of argv.slice(1)) {
    if (!arg || typeof arg !== "string" || arg.startsWith("--")) {
      continue;
    }
    const candidate = arg.replace(/^"+|"+$/g, "");
    if (candidate.toLowerCase().endsWith(".pdf") && fs.existsSync(candidate)) {
      return path.resolve(candidate);
    }
  }
  return null;
}

function sendToRenderer(channel, payload) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }
  mainWindow.webContents.send(channel, payload);
}

function sendPdfToRenderer(filePath) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    pendingPdfToOpen = filePath;
    return;
  }
  sendToRenderer("system-open-file", filePath);
}

function toBuffer(data) {
  if (Buffer.isBuffer(data)) {
    return data;
  }
  if (data instanceof Uint8Array) {
    return Buffer.from(data);
  }
  if (Array.isArray(data)) {
    return Buffer.from(data);
  }
  if (data && data.type === "Buffer" && Array.isArray(data.data)) {
    return Buffer.from(data.data);
  }
  throw new Error("저장할 PDF 데이터 형식이 올바르지 않습니다.");
}

function createMenu() {
  const template = [
    {
      label: "파일",
      submenu: [
        {
          label: "열기...",
          accelerator: "Ctrl+O",
          click: () => sendToRenderer("menu-action", "open-file")
        },
        {
          label: "다른 이름으로 저장...",
          accelerator: "Ctrl+Shift+S",
          click: () => sendToRenderer("menu-action", "save-as")
        },
        {
          label: "원본 덮어쓰기 저장",
          accelerator: "Ctrl+S",
          click: () => sendToRenderer("menu-action", "save-overwrite")
        },
        { type: "separator" },
        {
          label: "인쇄...",
          accelerator: "Ctrl+P",
          click: () => sendToRenderer("menu-action", "print")
        },
        { type: "separator" },
        {
          label: "종료",
          accelerator: "Alt+F4",
          click: () => app.quit()
        }
      ]
    },
    {
      label: "보기",
      submenu: [
        {
          label: "이전 페이지",
          accelerator: "PageUp",
          click: () => sendToRenderer("menu-action", "prev-page")
        },
        {
          label: "다음 페이지",
          accelerator: "PageDown",
          click: () => sendToRenderer("menu-action", "next-page")
        },
        { type: "separator" },
        {
          label: "확대",
          accelerator: "Ctrl+=",
          click: () => sendToRenderer("menu-action", "zoom-in")
        },
        {
          label: "축소",
          accelerator: "Ctrl+-",
          click: () => sendToRenderer("menu-action", "zoom-out")
        },
        {
          label: "원래 크기",
          accelerator: "Ctrl+0",
          click: () => sendToRenderer("menu-action", "zoom-reset")
        },
        { type: "separator" },
        {
          label: "전체화면",
          accelerator: "F11",
          click: () => {
            if (!mainWindow) {
              return;
            }
            mainWindow.setFullScreen(!mainWindow.isFullScreen());
          }
        },
        {
          label: "전체화면 보기 모드 전환",
          accelerator: "Ctrl+Shift+F",
          click: () => sendToRenderer("menu-action", "toggle-fullscreen-view-mode")
        },
        {
          label: "미리보기 패널 표시/숨김",
          accelerator: "Ctrl+B",
          click: () => sendToRenderer("menu-action", "toggle-thumb-panel")
        },
        { type: "separator" },
        {
          label: "다크모드 전환",
          accelerator: "Ctrl+D",
          click: () => sendToRenderer("menu-action", "toggle-dark")
        }
      ]
    },
    {
      label: "업데이트",
      submenu: [
        {
          label: "업데이트 확인",
          click: () => sendToRenderer("menu-action", "check-update")
        },
        {
          label: "다운로드된 업데이트 설치",
          click: () => sendToRenderer("menu-action", "install-update")
        }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 940,
    minWidth: 1080,
    minHeight: 700,
    title: "lookup",
    show: false,
    autoHideMenuBar: false,
    backgroundColor: "#e8edf3",
    icon: path.join(__dirname, "build", "icon.png"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, "src", "renderer", "index.html"));

  mainWindow.once("ready-to-show", () => {
    mainWindow?.show();
  });

  mainWindow.webContents.on("did-finish-load", () => {
    if (pendingPdfToOpen) {
      sendToRenderer("system-open-file", pendingPdfToOpen);
      pendingPdfToOpen = null;
    }
    sendToRenderer("window-fullscreen-changed", mainWindow?.isFullScreen() ?? false);
  });

  mainWindow.on("enter-full-screen", () => {
    sendToRenderer("window-fullscreen-changed", true);
  });

  mainWindow.on("leave-full-screen", () => {
    sendToRenderer("window-fullscreen-changed", false);
  });
}

function loadUpdateConfig() {
  const envOwner = process.env.LOOKUP_UPDATE_OWNER?.trim();
  const envRepo = process.env.LOOKUP_UPDATE_REPO?.trim();
  if (envOwner && envRepo) {
    return { owner: envOwner, repo: envRepo };
  }

  const candidateFiles = [
    path.join(__dirname, "update-config.json"),
    path.join(process.resourcesPath || "", "update-config.json")
  ];

  for (const filePath of candidateFiles) {
    if (!filePath || !fs.existsSync(filePath)) {
      continue;
    }
    try {
      const raw = fs.readFileSync(filePath, "utf8");
      const parsed = JSON.parse(raw);
      const owner = String(parsed.owner || "").trim();
      const repo = String(parsed.repo || "").trim();
      if (owner && repo) {
        return { owner, repo };
      }
    } catch (_error) {
      // ignore and try next file
    }
  }

  return null;
}

function sendUpdateStatus(status, extra = {}) {
  sendToRenderer("update-status", { status, ...extra });
}

function setupAutoUpdater() {
  updateConfig = loadUpdateConfig();
  if (!updateConfig) {
    sendUpdateStatus("disabled", {
      message: "업데이트 설정이 없습니다. update-config.json에 owner/repo를 입력해 주세요."
    });
    return;
  }

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  if (!app.isPackaged) {
    autoUpdater.forceDevUpdateConfig = true;
  }

  autoUpdater.setFeedURL({
    provider: "github",
    owner: updateConfig.owner,
    repo: updateConfig.repo
  });

  autoUpdater.on("checking-for-update", () => {
    sendUpdateStatus("checking", { message: "업데이트를 확인하고 있습니다." });
  });

  autoUpdater.on("update-available", (info) => {
    sendUpdateStatus("available", {
      message: `새 버전 ${info.version}을 찾았습니다. 다운로드를 시작합니다.`,
      version: info.version
    });
  });

  autoUpdater.on("update-not-available", () => {
    sendUpdateStatus("not-available", { message: "현재 최신 버전입니다." });
  });

  autoUpdater.on("download-progress", (progress) => {
    sendUpdateStatus("downloading", {
      message: `업데이트 다운로드 중... ${Math.round(progress.percent || 0)}%`,
      percent: progress.percent || 0
    });
  });

  autoUpdater.on("update-downloaded", (info) => {
    updateDownloaded = true;
    sendUpdateStatus("downloaded", {
      message: `업데이트 ${info.version} 다운로드 완료. 다시 시작하면 설치됩니다.`,
      version: info.version
    });
  });

  autoUpdater.on("error", (error) => {
    sendUpdateStatus("error", {
      message: `업데이트 오류: ${error?.message || "알 수 없는 오류"}`
    });
  });
}

ipcMain.handle("dialog:open-pdf", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "PDF 파일 열기",
    properties: ["openFile"],
    filters: [{ name: "PDF 문서", extensions: ["pdf"] }]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

ipcMain.handle("dialog:save-pdf", async (_event, options = {}) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "PDF 저장",
    defaultPath: options.defaultPath || "edited.pdf",
    filters: [{ name: "PDF 문서", extensions: ["pdf"] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }
  return result.filePath;
});

ipcMain.handle("dialog:confirm-overwrite", async (_event, options = {}) => {
  const result = await dialog.showMessageBox(mainWindow, {
    type: "warning",
    buttons: ["덮어쓰기", "취소"],
    defaultId: 1,
    cancelId: 1,
    title: "원본 덮어쓰기 확인",
    message: options.message || "원본 파일을 덮어쓰시겠습니까?",
    detail: options.detail || "이 작업은 되돌릴 수 없습니다."
  });
  return result.response === 0;
});

ipcMain.handle("pdf:read-file", async (_event, filePath) => {
  if (typeof filePath !== "string") {
    throw new Error("파일 경로가 올바르지 않습니다.");
  }

  const resolved = path.resolve(filePath);
  if (path.extname(resolved).toLowerCase() !== ".pdf") {
    throw new Error("PDF 파일만 열 수 있습니다.");
  }

  const stat = await fsp.stat(resolved);
  if (!stat.isFile()) {
    throw new Error("파일이 존재하지 않습니다.");
  }

  return fsp.readFile(resolved);
});

ipcMain.handle("pdf:write-file", async (_event, payload) => {
  const filePath = payload?.filePath;
  if (typeof filePath !== "string" || !filePath.trim()) {
    throw new Error("저장 경로가 비어 있습니다.");
  }
  const buffer = toBuffer(payload?.data);
  await fsp.writeFile(path.resolve(filePath), buffer);
  return true;
});

ipcMain.handle("window:toggle-fullscreen", () => {
  if (!mainWindow) {
    return false;
  }
  mainWindow.setFullScreen(!mainWindow.isFullScreen());
  return mainWindow.isFullScreen();
});

ipcMain.handle("window:set-fullscreen", (_event, enabled) => {
  if (!mainWindow) {
    return false;
  }
  mainWindow.setFullScreen(Boolean(enabled));
  return mainWindow.isFullScreen();
});

ipcMain.handle("window:is-fullscreen", () => {
  return mainWindow?.isFullScreen() ?? false;
});

ipcMain.handle("window:print", async () => {
  if (!mainWindow) {
    return false;
  }

  return new Promise((resolve) => {
    mainWindow.webContents.print(
      {
        silent: false,
        printBackground: true
      },
      (success) => resolve(success)
    );
  });
});

ipcMain.handle("update:check", async () => {
  if (!updateConfig) {
    return { ok: false, message: "업데이트 설정이 없습니다." };
  }
  try {
    await autoUpdater.checkForUpdates();
    return { ok: true };
  } catch (error) {
    return { ok: false, message: error?.message || "업데이트 확인 실패" };
  }
});

ipcMain.handle("update:install-now", async () => {
  if (!updateDownloaded) {
    return { ok: false, message: "다운로드된 업데이트가 없습니다." };
  }
  setImmediate(() => {
    autoUpdater.quitAndInstall(false, true);
  });
  return { ok: true };
});

ipcMain.handle("update:get-config", async () => {
  return {
    enabled: Boolean(updateConfig),
    owner: updateConfig?.owner || "",
    repo: updateConfig?.repo || ""
  };
});

if (gotSingleInstanceLock) {
  app.whenReady().then(() => {
    createMenu();
    createWindow();
    setupAutoUpdater();

    if (updateConfig && app.isPackaged) {
      setTimeout(() => {
        autoUpdater.checkForUpdates().catch(() => {
          // status events already show the error
        });
      }, 6000);
    }

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  });

  app.on("second-instance", (_event, argv) => {
    const incomingFile = extractPdfPath(argv);
    if (mainWindow) {
      if (mainWindow.isMinimized()) {
        mainWindow.restore();
      }
      mainWindow.focus();
    }
    if (incomingFile) {
      sendPdfToRenderer(incomingFile);
    }
  });

  app.on("open-file", (event, filePath) => {
    event.preventDefault();
    if (filePath.toLowerCase().endsWith(".pdf")) {
      sendPdfToRenderer(path.resolve(filePath));
    }
  });

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
      app.quit();
    }
  });
}
