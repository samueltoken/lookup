const { app, BrowserWindow, Menu, dialog, ipcMain, clipboard, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const crypto = require("node:crypto");
const https = require("node:https");
const { spawnSync, spawn } = require("node:child_process");

let mainWindow = null;
let pendingDocumentToOpen = null;
let updateConfig = null;
let updateDownloaded = false;
let updateTargetVersion = "";
let updateInstallTriggered = false;
let updateBusy = false;
let appSettings = { language: "ko" };
let pendingInstalledUpdate = null;
let updateReleaseNotes = [];
let updateReleaseNotesRaw = "";
let updateReleaseNotesFetchPromise = null;

let officeParserFn = null;
let WordExtractorCtor = null;
let XLSXMod = null;
let hwpJsMod = null;
let readHwpxDocumentFn = null;
let pdfToPrinterMod = null;

const menuText = {
  ko: {
    file: "파일",
    open: "열기...",
    saveAs: "다른 이름 저장...",
    saveOverwrite: "원본 덮어쓰기 저장",
    print: "인쇄...",
    quit: "종료",
    view: "보기",
    prevPage: "이전 페이지",
    nextPage: "다음 페이지",
    zoomIn: "확대",
    zoomOut: "축소",
    zoomReset: "원래 크기",
    fullscreen: "전체화면",
    fullscreenMode: "전체화면 보기 모드 전환",
    thumbToggle: "미리보기 패널 표시/숨김",
    searchToggle: "검색 패널 표시/숨김",
    darkToggle: "다크모드 전환",
    settings: "설정",
    language: "언어",
    languageKo: "한국어",
    languageEn: "English",
    updateCheck: "업데이트 확인",
    copyDeveloperContact: "개발자 문의 이메일 복사"
  },
  en: {
    file: "File",
    open: "Open...",
    saveAs: "Save As...",
    saveOverwrite: "Overwrite Original",
    print: "Print...",
    quit: "Quit",
    view: "View",
    prevPage: "Previous Page",
    nextPage: "Next Page",
    zoomIn: "Zoom In",
    zoomOut: "Zoom Out",
    zoomReset: "Actual Size",
    fullscreen: "Fullscreen",
    fullscreenMode: "Toggle Fullscreen View Mode",
    thumbToggle: "Toggle Thumbnails",
    searchToggle: "Toggle Search Panel",
    darkToggle: "Toggle Dark Mode",
    settings: "Settings",
    language: "Language",
    languageKo: "Korean",
    languageEn: "English",
    updateCheck: "Check for Updates",
    copyDeveloperContact: "Copy Developer Email"
  }
};

function mt(key) {
  const lang = appSettings.language === "en" ? "en" : "ko";
  return menuText[lang][key] || menuText.ko[key] || key;
}

const SUPPORTED_DOCUMENT_EXTENSIONS = new Set([".pdf", ".hwp", ".hwpx", ".doc", ".docx", ".xls", ".xlsx"]);
const OFFICE_EXTENSIONS = new Set([".doc", ".docx", ".xls", ".xlsx"]);
const HWP_EXTENSIONS = new Set([".hwp", ".hwpx"]);
pendingDocumentToOpen = extractDocumentPath(process.argv);

const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
}
app.setName("lookup");
app.setAppUserModelId("com.lookup.pdfviewer");

function getFileExtension(filePath) {
  return path.extname(String(filePath || "")).toLowerCase();
}

function isSupportedDocumentPath(filePath) {
  if (!filePath || typeof filePath !== "string") {
    return false;
  }
  const ext = getFileExtension(filePath);
  return SUPPORTED_DOCUMENT_EXTENSIONS.has(ext) && fs.existsSync(filePath);
}

function extractDocumentPath(argv) {
  if (!Array.isArray(argv)) {
    return null;
  }

  for (const arg of argv.slice(1)) {
    if (!arg || typeof arg !== "string" || arg.startsWith("--")) {
      continue;
    }
    const candidate = arg.replace(/^"+|"+$/g, "");
    if (isSupportedDocumentPath(candidate)) {
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

function sendDocumentToRenderer(filePath) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    pendingDocumentToOpen = filePath;
    return;
  }
  sendToRenderer("system-open-file", filePath);
}

function ensureOfficeParser() {
  if (!officeParserFn) {
    ({ parseOffice: officeParserFn } = require("officeparser"));
  }
  return officeParserFn;
}

function ensureWordExtractorCtor() {
  if (!WordExtractorCtor) {
    WordExtractorCtor = require("word-extractor");
  }
  return WordExtractorCtor;
}

function ensureXlsxModule() {
  if (!XLSXMod) {
    XLSXMod = require("xlsx");
  }
  return XLSXMod;
}

function ensureHwpJsModule() {
  if (!hwpJsMod) {
    hwpJsMod = require("@ohah/hwpjs");
  }
  return hwpJsMod;
}

function ensureHwpxReader() {
  if (!readHwpxDocumentFn) {
    ({ read: readHwpxDocumentFn } = require("hwpx-js"));
  }
  return readHwpxDocumentFn;
}

function ensurePdfToPrinter() {
  if (!pdfToPrinterMod) {
    pdfToPrinterMod = require("pdf-to-printer");
  }
  return pdfToPrinterMod;
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

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeForPowerShell(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}

function readJsonSafe(filePath) {
  try {
    const raw = fs.readFileSync(filePath, "utf8");
    return JSON.parse(raw);
  } catch (_error) {
    return null;
  }
}

function getSettingsFilePath() {
  return path.join(app.getPath("userData"), "settings.json");
}

function loadSettings() {
  const loaded = readJsonSafe(getSettingsFilePath());
  if (!loaded || typeof loaded !== "object") {
    appSettings = { language: "ko" };
    return;
  }
  appSettings = {
    language: loaded.language === "en" ? "en" : "ko"
  };
}

function saveSettings() {
  try {
    const settingsPath = getSettingsFilePath();
    fs.mkdirSync(path.dirname(settingsPath), { recursive: true });
    fs.writeFileSync(settingsPath, JSON.stringify(appSettings, null, 2), "utf8");
  } catch (_error) {
    // settings save failure is non-fatal
  }
}

function getWindowIconPath() {
  const packagedIconPath = path.join(process.resourcesPath || "", "icon.ico");
  if (app.isPackaged && fs.existsSync(packagedIconPath)) {
    return packagedIconPath;
  }
  const devIconPath = path.join(__dirname, "build", "icon.ico");
  if (fs.existsSync(devIconPath)) {
    return devIconPath;
  }
  return undefined;
}

function ensureConvertedDir() {
  const dir = path.join(app.getPath("userData"), "converted-documents");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function ensurePrintPreviewDir() {
  const dir = path.join(app.getPath("userData"), "print-preview");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

async function openSystemPrintDialog(pdfBuffer, suggestedName = "lookup-print.pdf") {
  const rawName = String(suggestedName || "").trim();
  const baseName = rawName ? rawName.replace(/[\\/:*?"<>|]+/g, "_") : "lookup-print.pdf";
  const fileName = baseName.toLowerCase().endsWith(".pdf") ? baseName : `${baseName}.pdf`;
  const uniqueName = `${Date.now()}-${fileName}`;
  const outPath = path.join(ensurePrintPreviewDir(), uniqueName);
  await fsp.writeFile(outPath, pdfBuffer);

  try {
    const { print } = ensurePdfToPrinter();
    await print(outPath, { printDialog: true });
    return { ok: true, path: outPath };
  } catch (_error) {
    return await openSystemPrintDialogFallback(outPath);
  }
}

async function openSystemPrintDialogFallback(pdfPath) {
  const printWindow = new BrowserWindow({
    show: false,
    width: 1100,
    height: 1400,
    autoHideMenuBar: true,
    webPreferences: {
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  return await new Promise((resolve) => {
    let settled = false;
    const finish = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      if (!printWindow.isDestroyed()) {
        printWindow.close();
      }
      resolve(result);
    };

    const timeoutId = setTimeout(() => {
      finish({ ok: false, message: "인쇄 대화상자를 열지 못했습니다. 잠시 후 다시 시도해 주세요." });
    }, 20000);

    printWindow.webContents.once("did-fail-load", (_event, errorCode, errorDescription) => {
      clearTimeout(timeoutId);
      finish({
        ok: false,
        message: `인쇄 문서를 불러오지 못했습니다. (${errorDescription || errorCode || "load-failed"})`
      });
    });

    printWindow.webContents.once("did-finish-load", () => {
      setTimeout(() => {
        printWindow.webContents.print(
          {
            silent: false,
            printBackground: true
          },
          (success, failureReason) => {
            clearTimeout(timeoutId);
            if (!success) {
              finish({ ok: false, message: failureReason || "인쇄 대화상자를 열지 못했습니다." });
              return;
            }
            finish({ ok: true, path: pdfPath });
          }
        );
      }, 220);
    });

    printWindow.loadURL(pathToFileURL(pdfPath).toString()).catch((error) => {
      clearTimeout(timeoutId);
      finish({ ok: false, message: error?.message || "인쇄 준비 중 오류가 발생했습니다." });
    });
  });
}

function getUpdateMarkerFilePath() {
  return path.join(app.getPath("userData"), "last-installed-update.json");
}

function normalizeReleaseNotesLines(releaseNotes) {
  const raw = Array.isArray(releaseNotes) ? releaseNotes.join("\n") : String(releaseNotes || "");
  const plainText = raw
    .replace(/\uFEFF/g, "")
    .replace(/```[\s\S]*?```/g, "\n")
    .replace(/<\/?[^>]+>/g, " ")
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1");
  const lines = plainText
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+$/g, ""))
    .filter(Boolean);
  const result = [];
  const seen = new Set();
  for (const line of lines) {
    const normalized = line.replace(/\s+/g, " ").trim();
    if (!normalized) {
      continue;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(normalized);
    if (result.length >= 80) {
      break;
    }
  }
  return result;
}

function extractReleaseNotesRawText(releaseNotes) {
  if (Array.isArray(releaseNotes)) {
    return releaseNotes
      .map((entry) => {
        if (typeof entry === "string") {
          return entry;
        }
        if (!entry || typeof entry !== "object") {
          return "";
        }
        return String(entry.note || entry.releaseNotes || "").trim();
      })
      .filter(Boolean)
      .join("\n\n");
  }
  if (typeof releaseNotes === "string") {
    return releaseNotes;
  }
  if (releaseNotes && typeof releaseNotes === "object") {
    return String(releaseNotes.note || releaseNotes.releaseNotes || "");
  }
  return "";
}

function setUpdateReleaseNotes(rawReleaseNotes) {
  const rawText = String(rawReleaseNotes || "").replace(/\uFEFF/g, "").trim();
  if (!rawText) {
    return false;
  }
  updateReleaseNotesRaw = rawText;
  updateReleaseNotes = normalizeReleaseNotesLines(rawText);
  return true;
}

function extractReleaseNotes(info) {
  const releaseNotes = info?.releaseNotes;
  return extractReleaseNotesRawText(releaseNotes);
}

function rememberReleaseNotes(info) {
  const rawNotes = extractReleaseNotes(info);
  if (rawNotes.trim()) {
    setUpdateReleaseNotes(rawNotes);
  }
}

function requestJsonFromGitHubApi(apiPath) {
  return new Promise((resolve) => {
    if (!updateConfig?.owner || !updateConfig?.repo) {
      resolve(null);
      return;
    }
    const headers = {
      Accept: "application/vnd.github+json",
      "User-Agent": "lookup-updater"
    };
    const token = (process.env.GITHUB_TOKEN || process.env.LOOKUP_GITHUB_TOKEN || "").trim();
    if (token) {
      headers.Authorization = `Bearer ${token}`;
    }
    const request = https.request(
      {
        protocol: "https:",
        hostname: "api.github.com",
        path: apiPath,
        method: "GET",
        headers,
        timeout: 4800
      },
      (response) => {
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => {
          const statusCode = Number(response.statusCode || 0);
          if (statusCode < 200 || statusCode >= 300) {
            resolve(null);
            return;
          }
          try {
            const raw = Buffer.concat(chunks).toString("utf8");
            resolve(JSON.parse(raw));
          } catch (_error) {
            resolve(null);
          }
        });
      }
    );
    request.on("timeout", () => {
      request.destroy(new Error("timeout"));
    });
    request.on("error", () => resolve(null));
    request.end();
  });
}

async function fetchReleaseNotesFromGitHub(versionText = "") {
  if (!updateConfig?.owner || !updateConfig?.repo) {
    return "";
  }
  const owner = encodeURIComponent(updateConfig.owner);
  const repo = encodeURIComponent(updateConfig.repo);
  const version = String(versionText || "").trim();
  const tagCandidates = [];
  if (version) {
    if (version.toLowerCase().startsWith("v")) {
      tagCandidates.push(version);
    } else {
      tagCandidates.push(`v${version}`);
      tagCandidates.push(version);
    }
  }

  for (const tag of tagCandidates) {
    const release = await requestJsonFromGitHubApi(`/repos/${owner}/${repo}/releases/tags/${encodeURIComponent(tag)}`);
    const rawBody = String(release?.body || "");
    const lines = normalizeReleaseNotesLines(rawBody);
    if (lines.length > 0) {
      return rawBody;
    }
  }

  const latestRelease = await requestJsonFromGitHubApi(`/repos/${owner}/${repo}/releases/latest`);
  return String(latestRelease?.body || "");
}

function ensureUpdateReleaseNotes(versionText = "") {
  if (updateReleaseNotesRaw.trim()) {
    return Promise.resolve({
      raw: updateReleaseNotesRaw,
      lines: updateReleaseNotes
    });
  }
  if (updateReleaseNotesFetchPromise) {
    return updateReleaseNotesFetchPromise;
  }

  updateReleaseNotesFetchPromise = (async () => {
    const fetchedRaw = await fetchReleaseNotesFromGitHub(versionText);
    if (String(fetchedRaw || "").trim()) {
      setUpdateReleaseNotes(fetchedRaw);
    }
    return {
      raw: updateReleaseNotesRaw,
      lines: updateReleaseNotes
    };
  })().finally(() => {
    updateReleaseNotesFetchPromise = null;
  });

  return updateReleaseNotesFetchPromise;
}

function markInstalledVersion(version, releaseNotes = "") {
  try {
    const rawNotes = extractReleaseNotesRawText(releaseNotes) || String(releaseNotes || "");
    const plainNotes = normalizeReleaseNotesLines(rawNotes);
    const marker = {
      version: String(version || "").trim(),
      releaseNotesRaw: rawNotes,
      releaseNotesPlain: plainNotes,
      time: new Date().toISOString()
    };
    fs.mkdirSync(path.dirname(getUpdateMarkerFilePath()), { recursive: true });
    fs.writeFileSync(getUpdateMarkerFilePath(), JSON.stringify(marker, null, 2), "utf8");
  } catch (_error) {
    // ignore marker write errors
  }
}

function consumeInstalledVersionMarker() {
  try {
    const markerPath = getUpdateMarkerFilePath();
    if (!fs.existsSync(markerPath)) {
      return null;
    }
    const parsed = JSON.parse(fs.readFileSync(markerPath, "utf8"));
    fs.unlinkSync(markerPath);
    const version = String(parsed?.version || "").trim();
    if (!version) {
      return null;
    }
    return {
      version,
      releaseNotesRaw: extractReleaseNotesRawText(parsed?.releaseNotesRaw || parsed?.releaseNotes || ""),
      releaseNotes: Array.isArray(parsed?.releaseNotesPlain) && parsed.releaseNotesPlain.length
        ? normalizeReleaseNotesLines(parsed.releaseNotesPlain)
        : normalizeReleaseNotesLines(parsed?.releaseNotesRaw || parsed?.releaseNotes || "")
    };
  } catch (_error) {
    return null;
  }
}

function setUpdateBusy(nextBusy) {
  const boolValue = Boolean(nextBusy);
  if (updateBusy === boolValue) {
    return;
  }
  updateBusy = boolValue;
  if (app.isReady()) {
    createMenu();
  }
}

function createConvertedPdfPath(sourcePath) {
  const resolved = path.resolve(sourcePath);
  const stat = fs.statSync(resolved);
  const ext = getFileExtension(resolved).replace(".", "");
  const baseName = path.basename(resolved, path.extname(resolved)).replace(/[^\\w.\\-가-힣]/g, "_").slice(0, 56);
  const signature = `${resolved}|${stat.mtimeMs}|${stat.size}`;
  const digest = crypto.createHash("sha1").update(signature).digest("hex").slice(0, 12);
  return path.join(ensureConvertedDir(), `${baseName}-${ext}-${digest}.pdf`);
}

function classifyDocumentOpenError(error) {
  const message = String(error?.message || "");
  const code = String(error?.code || "");
  if (code === "ENOENT" || /존재하지 않습니다|no such file/i.test(message)) {
    return "NOT_FOUND";
  }
  if (code === "EACCES" || code === "EPERM" || /권한|permission/i.test(message)) {
    return "NO_PERMISSION";
  }
  if (/지원하지 않는 파일 형식|unsupported/i.test(message)) {
    return "UNSUPPORTED_FORMAT";
  }
  if (/timed out|timeout|시간.*초과|응답.*지연/i.test(message)) {
    return "CONVERT_TIMEOUT";
  }
  if (/libreoffice_missing|engine.*missing|변환 엔진.*찾/i.test(message)) {
    return "ENGINE_MISSING";
  }
  if (/변환|convert|hwp|office/i.test(message)) {
    return "CONVERT_FAILED";
  }
  return "OPEN_FAILED";
}

function runPowerShell(script, timeoutMs = 180000) {
  const result = spawnSync(
    "powershell",
    ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
    {
      encoding: "utf8",
      windowsHide: true,
      timeout: timeoutMs
    }
  );
  return {
    ok: result.status === 0 && !result.error,
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
    error: result.error ? String(result.error.message || result.error) : ""
  };
}

function runProcess(command, args, timeoutMs = 180000) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";
    let settled = false;
    let timedOut = false;

    const finish = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      resolve(result);
    };

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", (error) => {
      finish({
        ok: false,
        status: null,
        timedOut,
        stdout: stdout.trim(),
        stderr: stderr.trim(),
        error: String(error?.message || error || "process_error")
      });
    });
    child.on("close", (code) => {
      finish({
        ok: code === 0 && !timedOut,
        status: code,
        timedOut,
        stdout: stdout.trim(),
        stderr: stderr.trim(),
        error: ""
      });
    });

    const timer = setTimeout(() => {
      timedOut = true;
      try {
        child.kill("SIGTERM");
      } catch (_error) {
        // ignore kill errors
      }
    }, timeoutMs);
  });
}

function findLibreOfficeExecutable() {
  const knownPaths = [
    path.join(process.env["ProgramFiles"] || "C:\\Program Files", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)", "LibreOffice", "program", "soffice.exe")
  ];
  for (const candidate of knownPaths) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  const whereResult = spawnSync("where", ["soffice"], {
    encoding: "utf8",
    windowsHide: true
  });
  if (whereResult.status === 0) {
    const first = String(whereResult.stdout || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find(Boolean);
    if (first && fs.existsSync(first)) {
      return first;
    }
  }

  return "";
}

function isTimeoutResult(result) {
  if (!result) {
    return false;
  }
  return Boolean(result.timedOut) || /timed out|timeout/i.test(String(result.error || result.stderr || ""));
}

function buildConversionWarning(sourceExt, engineMessage, hint) {
  const extLabel = String(sourceExt || "").replace(".", "").toUpperCase();
  const details = String(engineMessage || "").trim();
  const reason = hint || "변환 엔진을 사용하지 못해 호환 보기로 전환했습니다.";
  if (!details) {
    return `${extLabel} 문서: ${reason}`;
  }
  return `${extLabel} 문서: ${reason} (${details.slice(0, 120)})`;
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .trim();
}

function convertWithWordCom(sourcePath, outputPdfPath) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$word = $null",
    "$doc = $null",
    "try {",
    "  $word = New-Object -ComObject Word.Application",
    "  $word.Visible = $false",
    "  $doc = $word.Documents.Open(" + source + ", $false, $true)",
    "  $doc.ExportAsFixedFormat(" + output + ", 17)",
    "} finally {",
    "  if ($doc -ne $null) { $doc.Close($false) | Out-Null }",
    "  if ($word -ne $null) { $word.Quit() | Out-Null }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'word_export_failed' }"
  ].join("; ");

  return runPowerShell(script, 180000);
}

function convertWithExcelCom(sourcePath, outputPdfPath) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$excel = $null",
    "$workbook = $null",
    "try {",
    "  $excel = New-Object -ComObject Excel.Application",
    "  $excel.Visible = $false",
    "  $excel.DisplayAlerts = $false",
    "  $workbook = $excel.Workbooks.Open(" + source + ", $null, $true)",
    "  $workbook.ExportAsFixedFormat(0, " + output + ")",
    "} finally {",
    "  if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }",
    "  if ($excel -ne $null) { $excel.Quit() | Out-Null }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'excel_export_failed' }"
  ].join("; ");

  return runPowerShell(script, 180000);
}

function convertWithHancomCom(sourcePath, outputPdfPath, ext = ".hwp") {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "try {",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null } catch {}",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { throw 'hwp_open_failed' }",
    "  $saved = $hwp.SaveAs(" + output + ", 'PDF', '')",
    "  if (-not $saved) { throw 'hwp_save_failed' }",
    "} finally {",
    "  if ($hwp -ne $null) { try { $hwp.Quit() | Out-Null } catch {} }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("; ");
  return runPowerShell(script, 90000);
}

async function convertWithLibreOffice(sourcePath, outputPdfPath) {
  const soffice = findLibreOfficeExecutable();
  if (!soffice) {
    return {
      ok: false,
      missingEngine: true,
      stderr: "libreoffice_missing"
    };
  }

  const outDir = path.dirname(outputPdfPath);
  const sourceBaseName = path.basename(sourcePath, path.extname(sourcePath));
  const producedPdfPath = path.join(outDir, `${sourceBaseName}.pdf`);
  try {
    if (fs.existsSync(producedPdfPath)) {
      fs.unlinkSync(producedPdfPath);
    }
  } catch (_error) {
    // ignore cleanup errors
  }

  const result = await runProcess(
    soffice,
    ["--headless", "--invisible", "--nologo", "--nolockcheck", "--norestore", "--convert-to", "pdf", "--outdir", outDir, sourcePath],
    120000
  );

  if (!result.ok) {
    return result;
  }

  if (!fs.existsSync(producedPdfPath)) {
    return {
      ok: false,
      stderr: "libreoffice_pdf_not_created"
    };
  }

  if (path.resolve(producedPdfPath) !== path.resolve(outputPdfPath)) {
    await fsp.copyFile(producedPdfPath, outputPdfPath);
  }

  return {
    ok: true,
    stdout: result.stdout,
    stderr: result.stderr,
    status: result.status
  };
}

async function extractWordLegacyText(filePath) {
  const Ctor = ensureWordExtractorCtor();
  const extractor = new Ctor();
  const document = await extractor.extract(filePath);
  return normalizeText(document?.getBody?.() || "");
}

async function extractOfficeAstText(filePath) {
  try {
    const parseOffice = ensureOfficeParser();
    const ast = await parseOffice(filePath, { outputErrorToConsole: false });
    if (!ast) {
      return "";
    }
    if (typeof ast.toText === "function") {
      return normalizeText(ast.toText());
    }
    if (typeof ast === "string") {
      return normalizeText(ast);
    }
    return normalizeText(ast.text || "");
  } catch (_error) {
    return "";
  }
}

function extractSpreadsheetText(filePath) {
  const XLSX = ensureXlsxModule();
  const workbook = XLSX.readFile(filePath, { cellDates: false, dense: true });
  const sections = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      continue;
    }
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    const rowText = rows
      .map((row) =>
        Array.isArray(row)
          ? row
              .map((cell) => String(cell ?? "").trim())
              .filter(Boolean)
              .join(" | ")
          : ""
      )
      .filter(Boolean)
      .join("\n");
    if (rowText) {
      sections.push(`[${sheetName}]\n${rowText}`);
    }
  }
  return normalizeText(sections.join("\n\n"));
}

function extractHwpxRunsText(runs) {
  if (!Array.isArray(runs)) {
    return "";
  }
  const parts = [];
  for (const run of runs) {
    if (!run || typeof run !== "object") {
      continue;
    }
    if (run.t === "text" && run.text) {
      parts.push(String(run.text));
      continue;
    }
    if (Array.isArray(run.rows)) {
      for (const row of run.rows) {
        if (!Array.isArray(row?.cells)) {
          continue;
        }
        const cellText = row.cells
          .map((cell) => extractHwpxRunsText(cell?.runs || []))
          .filter(Boolean)
          .join(" | ");
        if (cellText) {
          parts.push(cellText);
        }
      }
      continue;
    }
    if (run.caption) {
      parts.push(String(run.caption));
    }
  }
  return parts.join(" ");
}

function extractHwpFamilyText(filePath) {
  try {
    const readHwpxDocument = ensureHwpxReader();
    const raw = fs.readFileSync(filePath);
    const document = readHwpxDocument(new Uint8Array(raw));
    const blocks = [];
    const sections = Array.isArray(document?.sections) ? document.sections : [];
    for (const section of sections) {
      const paragraphs = Array.isArray(section?.paragraphs) ? section.paragraphs : [];
      for (const para of paragraphs) {
        const line = extractHwpxRunsText(para?.runs || []);
        if (line) {
          blocks.push(line);
        }
      }
    }
    return normalizeText(blocks.join("\n"));
  } catch (_error) {
    return "";
  }
}

function extractHwpLegacyText(filePath) {
  try {
    const hwpjs = ensureHwpJsModule();
    const raw = fs.readFileSync(filePath);
    const markdownResult = hwpjs.toMarkdown(raw);
    if (typeof markdownResult === "string") {
      return normalizeText(markdownResult);
    }
    if (markdownResult && typeof markdownResult.markdown === "string") {
      return normalizeText(markdownResult.markdown);
    }
    return "";
  } catch (_error) {
    return "";
  }
}

function normalizeHwpHtmlDocument(html, title) {
  const raw = String(html || "").trim();
  const safeTitle = escapeHtml(title || "lookup HWP");
  if (!raw) {
    return "";
  }
  if (/<!doctype/i.test(raw)) {
    return raw;
  }
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
</head>
<body>
${raw}
</body>
</html>`;
}

async function convertHwpToPdfWithHwpJs(sourcePath, outputPdfPath) {
  try {
    const hwpjs = ensureHwpJsModule();
    const raw = fs.readFileSync(sourcePath);
    const htmlResult = hwpjs.toHtml(raw);
    const htmlRaw =
      typeof htmlResult === "string"
        ? htmlResult
        : typeof htmlResult?.html === "string"
          ? htmlResult.html
          : "";
    const html = normalizeHwpHtmlDocument(htmlRaw, path.basename(sourcePath));
    if (!html) {
      return { ok: false, reason: "hwpjs_html_empty" };
    }
    const pdfBuffer = await htmlToPdfBuffer(html);
    await fsp.writeFile(outputPdfPath, pdfBuffer);
    return { ok: true };
  } catch (error) {
    return { ok: false, reason: error?.message || "hwpjs_convert_failed" };
  }
}

function buildFallbackHtml({ sourcePath, sourceExt, contentText, warningMessage }) {
  const safeTitle = escapeHtml(path.basename(sourcePath));
  const safeExt = escapeHtml(sourceExt.replace(".", "").toUpperCase());
  const safeBody = escapeHtml(contentText || "(문서 내용을 추출하지 못했습니다.)");
  const warning = warningMessage
    ? `<div class="warning">${escapeHtml(warningMessage)}</div>`
    : "";

  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <style>
    @page { size: A4; margin: 18mm; }
    html, body { margin: 0; padding: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif; color: #111827; }
    .head { margin-bottom: 14px; border-bottom: 1px solid #cbd5e1; padding-bottom: 10px; }
    .title { font-size: 18px; font-weight: 700; margin: 0 0 4px; }
    .meta { font-size: 12px; color: #475569; }
    .warning { margin: 10px 0 14px; padding: 8px 10px; border: 1px solid #f2d6a8; background: #fff7ea; color: #9a5f08; font-size: 12px; border-radius: 6px; }
    pre { white-space: pre-wrap; word-break: break-word; line-height: 1.45; font-size: 12px; margin: 0; }
  </style>
</head>
<body>
  <div class="head">
    <p class="title">${safeTitle}</p>
    <div class="meta">원본 형식: ${safeExt} / lookup 임시 변환 문서</div>
  </div>
  ${warning}
  <pre>${safeBody}</pre>
</body>
</html>`;
}

async function htmlToPdfBuffer(html) {
  const tempWindow = new BrowserWindow({
    show: false,
    width: 1040,
    height: 1440,
    autoHideMenuBar: true,
    webPreferences: {
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  try {
    await tempWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    await new Promise((resolve) => setTimeout(resolve, 100));
    return await tempWindow.webContents.printToPDF({
      printBackground: true,
      pageSize: "A4",
      preferCSSPageSize: true
    });
  } finally {
    if (!tempWindow.isDestroyed()) {
      tempWindow.close();
    }
  }
}

async function extractDocumentTextForFallback(filePath, ext) {
  if (ext === ".doc") {
    const legacyText = await extractWordLegacyText(filePath);
    if (legacyText) {
      return legacyText;
    }
    return extractOfficeAstText(filePath);
  }
  if (ext === ".docx") {
    return extractOfficeAstText(filePath);
  }
  if (ext === ".xls" || ext === ".xlsx") {
    const fromSheet = extractSpreadsheetText(filePath);
    if (fromSheet) {
      return fromSheet;
    }
    return extractOfficeAstText(filePath);
  }
  if (ext === ".hwp") {
    const legacy = extractHwpLegacyText(filePath);
    if (legacy) {
      return legacy;
    }
    return extractHwpFamilyText(filePath);
  }
  if (ext === ".hwpx") {
    return extractHwpFamilyText(filePath);
  }
  return "";
}

async function convertOfficeFileToPdf(sourcePath, outputPdfPath, ext) {
  let comResult = { ok: false, stderr: "unsupported_office_ext" };
  if (ext === ".doc" || ext === ".docx") {
    comResult = convertWithWordCom(sourcePath, outputPdfPath);
  } else if (ext === ".xls" || ext === ".xlsx") {
    comResult = convertWithExcelCom(sourcePath, outputPdfPath);
  }

  if (comResult.ok && fs.existsSync(outputPdfPath)) {
    return {
      ok: true,
      engine: "office-com",
      warningMessage: ""
    };
  }

  const libreResult = await convertWithLibreOffice(sourcePath, outputPdfPath);
  if (libreResult.ok && fs.existsSync(outputPdfPath)) {
    const warningMessage = buildConversionWarning(ext, "", "Office 자동화 변환에 실패해 LibreOffice 엔진으로 열었습니다.");
    return {
      ok: true,
      engine: "libreoffice",
      warningMessage
    };
  }

  const reasonHint = isTimeoutResult(comResult) || isTimeoutResult(libreResult)
    ? "변환 시간이 초과되어 호환 보기로 전환했습니다."
    : libreResult.missingEngine
      ? "Word/Excel 또는 LibreOffice 변환 엔진을 찾지 못해 호환 보기로 전환했습니다."
      : "고품질 변환에 실패해 호환 보기로 전환했습니다.";

  const details = [comResult?.stderr, comResult?.error, libreResult?.stderr, libreResult?.error]
    .map((value) => String(value || "").trim())
    .find(Boolean);

  return {
    ok: false,
    engine: "fallback",
    warningMessage: buildConversionWarning(ext, details, reasonHint)
  };
}

async function convertDocumentToPdf(sourcePath) {
  if (typeof sourcePath !== "string" || !sourcePath.trim()) {
    throw new Error("파일 경로가 올바르지 않습니다.");
  }

  const resolved = path.resolve(sourcePath);
  const ext = getFileExtension(resolved);
  if (!SUPPORTED_DOCUMENT_EXTENSIONS.has(ext)) {
    throw new Error("지원하지 않는 파일 형식입니다.");
  }

  const stat = await fsp.stat(resolved);
  if (!stat.isFile()) {
    throw new Error("파일이 존재하지 않습니다.");
  }

  if (ext === ".pdf") {
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: false,
      convertMode: "native",
      warningMessage: "",
      pdfPath: resolved,
      pdfBuffer: await fsp.readFile(resolved)
    };
  }

  const targetPdfPath = createConvertedPdfPath(resolved);
  if (fs.existsSync(targetPdfPath)) {
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: true,
      convertMode: "cache",
      warningMessage: "",
      pdfPath: targetPdfPath,
      pdfBuffer: await fsp.readFile(targetPdfPath)
    };
  }

  let warningMessage = "";
  if (OFFICE_EXTENSIONS.has(ext)) {
    const officeResult = await convertOfficeFileToPdf(resolved, targetPdfPath, ext);
    if (officeResult.ok && fs.existsSync(targetPdfPath)) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: officeResult.engine || "office",
        warningMessage: officeResult.warningMessage || "",
        pdfPath: targetPdfPath,
        pdfBuffer: await fsp.readFile(targetPdfPath)
      };
    }
    warningMessage = officeResult.warningMessage || "고품질 변환 엔진을 찾지 못해 텍스트 기반 보기 모드로 변환했습니다.";
  } else if (HWP_EXTENSIONS.has(ext)) {
    const hwpResult = convertWithHancomCom(resolved, targetPdfPath, ext);
    if (hwpResult.ok && fs.existsSync(targetPdfPath)) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: "hwp-layout",
        warningMessage: "",
        pdfPath: targetPdfPath,
        pdfBuffer: await fsp.readFile(targetPdfPath)
      };
    }
    const hwpJsResult = await convertHwpToPdfWithHwpJs(resolved, targetPdfPath);
    if (hwpJsResult.ok && fs.existsSync(targetPdfPath)) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: "hwpjs-html",
        warningMessage: "원본 변환 엔진을 찾지 못해 호환 보기 모드로 열었습니다.",
        pdfPath: targetPdfPath,
        pdfBuffer: await fsp.readFile(targetPdfPath)
      };
    }
    const hwpErrorDetail = [hwpResult?.stderr, hwpResult?.error, hwpJsResult?.reason]
      .map((value) => String(value || "").trim())
      .find(Boolean);
    const reasonHint = isTimeoutResult(hwpResult)
      ? "한컴 변환 엔진 응답이 지연되어 텍스트 기반 보기로 전환했습니다."
      : "원본 레이아웃 변환에 실패해 텍스트 기반 보기로 전환했습니다.";
    warningMessage = buildConversionWarning(ext, hwpErrorDetail, reasonHint);
  }

  const extractedText = await extractDocumentTextForFallback(resolved, ext);
  const html = buildFallbackHtml({
    sourcePath: resolved,
    sourceExt: ext,
    contentText: extractedText,
    warningMessage
  });
  const fallbackPdf = await htmlToPdfBuffer(html);
  await fsp.writeFile(targetPdfPath, fallbackPdf);

  return {
    sourcePath: resolved,
    sourceExt: ext,
    converted: true,
    convertMode: "fallback",
    warningMessage,
    pdfPath: targetPdfPath,
    pdfBuffer: fallbackPdf
  };
}

function createMenu() {
  const template = [
    {
      label: mt("file"),
      submenu: [
        {
          label: mt("open"),
          accelerator: "Ctrl+O",
          click: () => sendToRenderer("menu-action", "open-file")
        },
        {
          label: mt("saveAs"),
          accelerator: "Ctrl+Shift+S",
          click: () => sendToRenderer("menu-action", "save-as")
        },
        {
          label: mt("saveOverwrite"),
          accelerator: "Ctrl+S",
          click: () => sendToRenderer("menu-action", "save-overwrite")
        },
        { type: "separator" },
        {
          label: mt("print"),
          accelerator: "Ctrl+P",
          click: () => sendToRenderer("menu-action", "print")
        },
        { type: "separator" },
        {
          label: mt("quit"),
          accelerator: "Alt+F4",
          click: () => app.quit()
        }
      ]
    },
    {
      label: mt("view"),
      submenu: [
        {
          label: mt("prevPage"),
          accelerator: "PageUp",
          click: () => sendToRenderer("menu-action", "prev-page")
        },
        {
          label: mt("nextPage"),
          accelerator: "PageDown",
          click: () => sendToRenderer("menu-action", "next-page")
        },
        { type: "separator" },
        {
          label: mt("zoomIn"),
          accelerator: "Ctrl+=",
          click: () => sendToRenderer("menu-action", "zoom-in")
        },
        {
          label: mt("zoomOut"),
          accelerator: "Ctrl+-",
          click: () => sendToRenderer("menu-action", "zoom-out")
        },
        {
          label: mt("zoomReset"),
          accelerator: "Ctrl+0",
          click: () => sendToRenderer("menu-action", "zoom-reset")
        },
        { type: "separator" },
        {
          label: mt("fullscreen"),
          accelerator: "F11",
          click: () => {
            if (!mainWindow) {
              return;
            }
            mainWindow.setFullScreen(!mainWindow.isFullScreen());
          }
        },
        {
          label: mt("fullscreenMode"),
          accelerator: "Ctrl+Shift+F",
          click: () => sendToRenderer("menu-action", "toggle-fullscreen-view-mode")
        },
        {
          label: mt("thumbToggle"),
          accelerator: "Ctrl+B",
          click: () => sendToRenderer("menu-action", "toggle-thumb-panel")
        },
        {
          label: mt("searchToggle"),
          accelerator: "Ctrl+Shift+B",
          click: () => sendToRenderer("menu-action", "toggle-search-panel")
        },
        { type: "separator" },
        {
          label: mt("darkToggle"),
          accelerator: "Ctrl+D",
          click: () => sendToRenderer("menu-action", "toggle-dark")
        }
      ]
    },
    {
      label: mt("settings"),
      submenu: [
        {
          label: mt("language"),
          submenu: [
            {
              label: mt("languageKo"),
              type: "radio",
              checked: appSettings.language !== "en",
              click: () => sendToRenderer("menu-action", "set-language-ko")
            },
            {
              label: mt("languageEn"),
              type: "radio",
              checked: appSettings.language === "en",
              click: () => sendToRenderer("menu-action", "set-language-en")
            }
          ]
        },
        { type: "separator" },
        {
          label: mt("updateCheck"),
          enabled: !updateBusy,
          click: () => sendToRenderer("menu-action", "check-update")
        },
        {
          label: mt("copyDeveloperContact"),
          click: () => sendToRenderer("menu-action", "copy-developer-email")
        }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  const iconPath = getWindowIconPath();
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 940,
    minWidth: 1080,
    minHeight: 700,
    title: "lookup",
    show: false,
    autoHideMenuBar: false,
    backgroundColor: "#e8edf3",
    icon: iconPath,
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

  mainWindow.webContents.on("did-finish-load", async () => {
    if (pendingDocumentToOpen) {
      sendToRenderer("system-open-file", pendingDocumentToOpen);
      pendingDocumentToOpen = null;
    }
    sendToRenderer("window-fullscreen-changed", mainWindow?.isFullScreen() ?? false);
    if (pendingInstalledUpdate?.version) {
      if (
        (!Array.isArray(pendingInstalledUpdate.releaseNotes) || pendingInstalledUpdate.releaseNotes.length === 0) &&
        !String(pendingInstalledUpdate.releaseNotesRaw || "").trim()
      ) {
        updateConfig = updateConfig || loadUpdateConfig();
        const fetchedNotesRaw = await fetchReleaseNotesFromGitHub(pendingInstalledUpdate.version);
        if (String(fetchedNotesRaw || "").trim()) {
          pendingInstalledUpdate.releaseNotesRaw = fetchedNotesRaw;
          pendingInstalledUpdate.releaseNotes = normalizeReleaseNotesLines(fetchedNotesRaw);
        }
      }
      sendUpdateStatus("installed", {
        stage: "installed",
        targetVersion: pendingInstalledUpdate.version,
        releaseNotes: Array.isArray(pendingInstalledUpdate.releaseNotes) ? pendingInstalledUpdate.releaseNotes : [],
        releaseNotesRaw: String(pendingInstalledUpdate.releaseNotesRaw || ""),
        message: `업데이트 완료(v${pendingInstalledUpdate.version})`
      });
      pendingInstalledUpdate = null;
    }
  });

  mainWindow.on("enter-full-screen", () => {
    sendToRenderer("window-fullscreen-changed", true);
  });

  mainWindow.on("leave-full-screen", () => {
    sendToRenderer("window-fullscreen-changed", false);
  });
}

function parseGitHubRepoFromUrl(url) {
  const text = String(url || "").trim();
  if (!text) {
    return null;
  }
  const match = text.match(/github\.com[/:]([^/]+)\/([^/.]+)(?:\.git)?$/i);
  if (!match) {
    return null;
  }
  return { owner: match[1], repo: match[2] };
}

function extractRepoFromPackageJson() {
  const pkgPath = path.join(__dirname, "package.json");
  const pkg = readJsonSafe(pkgPath);
  if (!pkg) {
    return null;
  }
  if (typeof pkg.repository === "string") {
    return parseGitHubRepoFromUrl(pkg.repository);
  }
  if (pkg.repository && typeof pkg.repository.url === "string") {
    return parseGitHubRepoFromUrl(pkg.repository.url);
  }
  return null;
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
      const fromUrl = parseGitHubRepoFromUrl(parsed.repository || parsed.repoUrl || "");
      if (fromUrl) {
        return fromUrl;
      }
    } catch (_error) {
      // ignore and try next file
    }
  }

  return extractRepoFromPackageJson();
}

function sendUpdateStatus(status, extra = {}) {
  sendToRenderer("update-status", {
    status,
    stage: extra.stage || status,
    currentVersion: app.getVersion(),
    targetVersion: extra.targetVersion || updateTargetVersion || "",
    ...extra
  });
}

function setupAutoUpdater() {
  updateConfig = loadUpdateConfig();
  if (!updateConfig) {
    setUpdateBusy(false);
    sendUpdateStatus("disabled", {
      stage: "disabled",
      message: "업데이트 설정이 없습니다. update-config.json에 owner/repo를 입력해 주세요."
    });
    return;
  }

  updateDownloaded = false;
  updateTargetVersion = "";
  updateInstallTriggered = false;
  updateReleaseNotes = [];
  updateReleaseNotesRaw = "";
  updateReleaseNotesFetchPromise = null;
  setUpdateBusy(false);
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = false;
  autoUpdater.autoRunAppAfterInstall = true;
  if (!app.isPackaged) {
    autoUpdater.forceDevUpdateConfig = true;
  }

  autoUpdater.setFeedURL({
    provider: "github",
    owner: updateConfig.owner,
    repo: updateConfig.repo
  });

  autoUpdater.on("checking-for-update", () => {
    setUpdateBusy(true);
    sendUpdateStatus("checking", { stage: "checking", percent: 0, message: "업데이트 확인중..." });
  });

  autoUpdater.on("update-available", (info) => {
    updateTargetVersion = String(info?.version || "");
    rememberReleaseNotes(info);
    setUpdateBusy(true);
    sendUpdateStatus("available", {
      stage: "downloading",
      targetVersion: updateTargetVersion,
      percent: 0,
      releaseNotes: updateReleaseNotes,
      releaseNotesRaw: updateReleaseNotesRaw,
      message: `새 버전 v${updateTargetVersion}을 찾았습니다. 업데이트 중입니다.`
    });
    void ensureUpdateReleaseNotes(updateTargetVersion).then((notesPayload) => {
      if (!notesPayload.lines.length) {
        return;
      }
      sendUpdateStatus("available", {
        stage: "downloading",
        targetVersion: updateTargetVersion,
        percent: 0,
        releaseNotes: notesPayload.lines,
        releaseNotesRaw: notesPayload.raw,
        message: `새 버전 v${updateTargetVersion}을 찾았습니다. 업데이트 중입니다.`
      });
    });
  });

  autoUpdater.on("update-not-available", () => {
    updateTargetVersion = "";
    updateReleaseNotes = [];
    updateReleaseNotesRaw = "";
    updateReleaseNotesFetchPromise = null;
    setUpdateBusy(false);
    sendUpdateStatus("not-available", { stage: "idle", percent: 100, message: "현재 최신 버전입니다." });
  });

  autoUpdater.on("download-progress", (progress) => {
    const percent = Math.max(0, Math.min(100, Number(progress?.percent || 0)));
    setUpdateBusy(true);
    sendUpdateStatus("downloading", {
      stage: "downloading",
      targetVersion: updateTargetVersion,
      releaseNotes: updateReleaseNotes,
      releaseNotesRaw: updateReleaseNotesRaw,
      message: `업데이트 중입니다... ${Math.round(percent)}%`,
      percent
    });
  });

  autoUpdater.on("update-downloaded", (info) => {
    updateDownloaded = true;
    updateInstallTriggered = false;
    updateTargetVersion = String(info?.version || updateTargetVersion || "");
    rememberReleaseNotes(info);
    setUpdateBusy(true);
    sendUpdateStatus("downloaded", {
      stage: "ready-to-install",
      targetVersion: updateTargetVersion,
      releaseNotes: updateReleaseNotes,
      releaseNotesRaw: updateReleaseNotesRaw,
      percent: 100,
      message: `업데이트 중입니다. v${updateTargetVersion} 설치를 준비합니다.`
    });

    if (app.isPackaged) {
      const triggerInstall = async () => {
        if (updateInstallTriggered) {
          return;
        }
        try {
          updateInstallTriggered = true;
          if (!updateReleaseNotes.length) {
            await Promise.race([
              ensureUpdateReleaseNotes(updateTargetVersion),
              new Promise((resolve) => setTimeout(resolve, 1800))
            ]);
          }
          markInstalledVersion(updateTargetVersion || app.getVersion(), updateReleaseNotesRaw || updateReleaseNotes);
          sendUpdateStatus("installing", {
            stage: "restarting",
            targetVersion: updateTargetVersion,
            releaseNotes: updateReleaseNotes,
            releaseNotesRaw: updateReleaseNotesRaw,
            percent: 100,
            message: "업데이트 중입니다. 앱을 다시 시작합니다..."
          });
          setUpdateBusy(true);
          autoUpdater.quitAndInstall(false, true);
        } catch (error) {
          updateInstallTriggered = false;
          setUpdateBusy(false);
          sendUpdateStatus("error", {
            stage: "error",
            message: `업데이트 재시작 설치 실패: ${error?.message || "알 수 없는 오류"}`
          });
        }
      };
      setTimeout(() => {
        void triggerInstall();
      }, 280);
      setTimeout(() => {
        if (!updateInstallTriggered) {
          void triggerInstall();
        }
      }, 2200);
    }
  });

  autoUpdater.on("before-quit-for-update", () => {
    setUpdateBusy(true);
    sendUpdateStatus("installing", {
      stage: "restarting",
      targetVersion: updateTargetVersion,
      releaseNotes: updateReleaseNotes,
      releaseNotesRaw: updateReleaseNotesRaw,
      percent: 100,
      message: "업데이트 중입니다. 설치를 계속합니다..."
    });
  });

  autoUpdater.on("error", (error) => {
    setUpdateBusy(false);
    updateReleaseNotesFetchPromise = null;
    sendUpdateStatus("error", {
      stage: "error",
      message: `업데이트 오류: ${error?.message || "알 수 없는 오류"}`
    });
  });
}

async function showOpenDocumentDialog() {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "문서 열기",
    properties: ["openFile"],
    filters: [
      { name: "지원 문서", extensions: ["pdf", "hwp", "hwpx", "doc", "docx", "xls", "xlsx"] },
      { name: "PDF 문서", extensions: ["pdf"] },
      { name: "한글 문서", extensions: ["hwp", "hwpx"] },
      { name: "Word 문서", extensions: ["doc", "docx"] },
      { name: "Excel 문서", extensions: ["xls", "xlsx"] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
}

ipcMain.handle("dialog:open-document", async () => {
  return showOpenDocumentDialog();
});

ipcMain.handle("dialog:open-pdf", async () => {
  return showOpenDocumentDialog();
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
  if (getFileExtension(resolved) !== ".pdf") {
    throw new Error("PDF 파일만 직접 읽을 수 있습니다.");
  }
  return await fsp.readFile(resolved);
});

ipcMain.handle("document:open", async (_event, filePath) => {
  try {
    const converted = await convertDocumentToPdf(filePath);
    return {
      ok: true,
      sourcePath: converted.sourcePath,
      sourceExt: converted.sourceExt,
      converted: converted.converted,
      convertMode: converted.convertMode,
      warningMessage: converted.warningMessage,
      pdfPath: converted.pdfPath,
      data: converted.pdfBuffer
    };
  } catch (error) {
    return {
      ok: false,
      errorCode: classifyDocumentOpenError(error),
      message: String(error?.message || "문서를 열지 못했습니다.")
    };
  }
});

ipcMain.handle("document:convert", async (_event, filePath) => {
  try {
    const converted = await convertDocumentToPdf(filePath);
    return {
      ok: true,
      sourcePath: converted.sourcePath,
      sourceExt: converted.sourceExt,
      converted: converted.converted,
      convertMode: converted.convertMode,
      warningMessage: converted.warningMessage,
      pdfPath: converted.pdfPath
    };
  } catch (error) {
    return {
      ok: false,
      errorCode: classifyDocumentOpenError(error),
      message: String(error?.message || "문서를 변환하지 못했습니다.")
    };
  }
});

ipcMain.handle("document:is-supported", async (_event, filePath) => {
  if (!filePath || typeof filePath !== "string") {
    return false;
  }
  return SUPPORTED_DOCUMENT_EXTENSIONS.has(getFileExtension(filePath));
});

ipcMain.handle("clipboard:copy-text", async (_event, text) => {
  clipboard.writeText(String(text || ""));
  return true;
});

ipcMain.handle("settings:get", async () => {
  return { language: appSettings.language };
});

ipcMain.handle("settings:set-language", async (_event, language) => {
  appSettings.language = language === "en" ? "en" : "ko";
  saveSettings();
  createMenu();
  return { language: appSettings.language };
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

ipcMain.handle("window:set-title", async (_event, title) => {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return false;
  }
  const nextTitle = typeof title === "string" && title.trim() ? title.trim() : "lookup";
  mainWindow.setTitle(nextTitle);
  return true;
});

ipcMain.handle("window:print-preview", async (_event, payload = {}) => {
  try {
    const buffer = toBuffer(payload?.data);
    const rawName = typeof payload?.fileName === "string" ? payload.fileName.trim() : "";
    const baseName = rawName ? rawName.replace(/[\\/:*?\"<>|]+/g, "_") : "lookup-print-preview.pdf";
    const fileName = baseName.toLowerCase().endsWith(".pdf") ? baseName : `${baseName}.pdf`;
    const uniqueName = `${Date.now()}-${fileName}`;
    const outPath = path.join(ensurePrintPreviewDir(), uniqueName);
    await fsp.writeFile(outPath, buffer);
    const openResult = await shell.openPath(outPath);
    if (openResult) {
      return { ok: false, message: openResult };
    }
    return { ok: true, path: outPath };
  } catch (error) {
    return { ok: false, message: error?.message || "인쇄 미리보기를 열지 못했습니다." };
  }
});

ipcMain.handle("update:check", async () => {
  if (!updateConfig) {
    return { ok: false, message: "업데이트 설정이 없습니다." };
  }
  if (updateBusy) {
    return { ok: false, message: "업데이트가 이미 진행 중입니다.", code: "UPDATE_BUSY" };
  }
  try {
    setUpdateBusy(true);
    updateReleaseNotes = [];
    updateReleaseNotesRaw = "";
    updateTargetVersion = "";
    await autoUpdater.checkForUpdates();
    return { ok: true };
  } catch (error) {
    setUpdateBusy(false);
    return { ok: false, message: error?.message || "업데이트 확인 실패" };
  }
});

ipcMain.handle("window:print-document", async (_event, payload = {}) => {
  try {
    const buffer = toBuffer(payload?.data);
    const rawName = typeof payload?.fileName === "string" ? payload.fileName.trim() : "";
    const baseName = rawName ? rawName.replace(/[\\/:*?\"<>|]+/g, "_") : "lookup-print.pdf";
    return await openSystemPrintDialog(buffer, baseName);
  } catch (error) {
    return { ok: false, message: error?.message || "인쇄를 시작하지 못했습니다." };
  }
});

ipcMain.handle("update:get-config", async () => {
  return {
    enabled: Boolean(updateConfig),
    busy: updateBusy,
    owner: updateConfig?.owner || "",
    repo: updateConfig?.repo || "",
    currentVersion: app.getVersion(),
    targetVersion: updateTargetVersion || ""
  };
});

ipcMain.handle("app:get-version", async () => app.getVersion());

if (gotSingleInstanceLock) {
  app.whenReady().then(() => {
    loadSettings();
    pendingInstalledUpdate = consumeInstalledVersionMarker();
    createMenu();
    createWindow();
    setupAutoUpdater();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  });

  app.on("second-instance", (_event, argv) => {
    const incomingFile = extractDocumentPath(argv);
    if (mainWindow) {
      if (mainWindow.isMinimized()) {
        mainWindow.restore();
      }
      mainWindow.focus();
    }
    if (incomingFile) {
      sendDocumentToRenderer(incomingFile);
    }
  });

  app.on("open-file", (event, filePath) => {
    event.preventDefault();
    if (isSupportedDocumentPath(filePath)) {
      sendDocumentToRenderer(path.resolve(filePath));
    }
  });

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
      app.quit();
    }
  });
}








