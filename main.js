const { app, BrowserWindow, Menu, dialog, ipcMain, clipboard, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const crypto = require("node:crypto");
const https = require("node:https");
const { spawnSync, spawn } = require("node:child_process");
const { PDFDocument } = require("pdf-lib");
const packageMeta = require("./package.json");

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
let nodePdfJsModPromise = null;

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
    copyDeveloperContact: "개발자 문의 이메일 복사",
    versionInfo: "버전 정보"
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
    copyDeveloperContact: "Copy Developer Email",
    versionInfo: "Version Info"
  }
};

function mt(key) {
  const lang = appSettings.language === "en" ? "en" : "ko";
  return menuText[lang][key] || menuText.ko[key] || key;
}

const SUPPORTED_DOCUMENT_EXTENSIONS = new Set([".pdf", ".hwp", ".hwpx", ".doc", ".docx", ".xls", ".xlsx"]);
const OFFICE_EXTENSIONS = new Set([".doc", ".docx", ".xls", ".xlsx"]);
const HWP_EXTENSIONS = new Set([".hwp", ".hwpx"]);
const CONVERT_CACHE_SCHEMA_VERSION = "v5";
const HWP_QUALITY_SELECTOR_VERSION = "v2";
const HWP_SELECTION_META_VERSION = "v2";
const FALLBACK_CACHE_TTL_MS = 45 * 60 * 1000;
const APP_RELEASE_VERSION = String(packageMeta?.version || app.getVersion() || "0.0.0");
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

async function ensureNodePdfJs() {
  if (!nodePdfJsModPromise) {
    nodePdfJsModPromise = import("pdfjs-dist/legacy/build/pdf.mjs")
      .then((mod) => {
        if (mod?.GlobalWorkerOptions) {
          mod.GlobalWorkerOptions.workerSrc = "";
        }
        return mod;
      })
      .catch((error) => {
        nodePdfJsModPromise = null;
        throw error;
      });
  }
  return nodePdfJsModPromise;
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

function getConversionEngineToken(convertMode = "fallback") {
  const deps = packageMeta?.dependencies || {};
  const hwpJsVer = String(deps["@ohah/hwpjs"] || "na");
  const hwpxVer = String(deps["hwpx-js"] || "na");
  const xlsxVer = String(deps.xlsx || "na");
  switch (String(convertMode || "").toLowerCase()) {
    case "office-com":
      return "officecom-v1";
    case "libreoffice":
      return "libreoffice-v1";
    case "hwp-print-pdf":
      return `hancomcom-print-v2-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-saveas":
      return `hancomcom-saveas-v2-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-layout":
      return "hancomcom-legacy-v1";
    case "hwpjs-html":
      return `hwpjs-${hwpJsVer}-hwpx-${hwpxVer}`;
    case "xlsx-html":
      return `xlsx-html-${xlsxVer}`;
    case "fallback":
      return `fallback-xlsx-${xlsxVer}`;
    default:
      return "generic-v1";
  }
}

function createConvertedPdfPath(sourcePath, convertMode = "fallback") {
  const resolved = path.resolve(sourcePath);
  const stat = fs.statSync(resolved);
  const ext = getFileExtension(resolved).replace(".", "");
  const mode = String(convertMode || "fallback").toLowerCase();
  const engineToken = getConversionEngineToken(mode);
  const baseName = path.basename(resolved, path.extname(resolved)).replace(/[^\\w.\\-가-힣]/g, "_").slice(0, 56);
  const signature = `${resolved}|${stat.mtimeMs}|${stat.size}|${APP_RELEASE_VERSION}|${CONVERT_CACHE_SCHEMA_VERSION}|${mode}|${engineToken}`;
  const digest = crypto.createHash("sha1").update(signature).digest("hex").slice(0, 12);
  return path.join(ensureConvertedDir(), `${baseName}-${ext}-${mode}-${digest}.pdf`);
}

function buildHwpSelectionSourceSignature(sourcePath) {
  const resolved = path.resolve(sourcePath);
  const stat = fs.statSync(resolved);
  const signature = `${resolved}|${stat.mtimeMs}|${stat.size}|${APP_RELEASE_VERSION}|${CONVERT_CACHE_SCHEMA_VERSION}|${HWP_SELECTION_META_VERSION}|${HWP_QUALITY_SELECTOR_VERSION}`;
  return crypto.createHash("sha1").update(signature).digest("hex");
}

function createHwpSelectionMetaPath(sourcePath) {
  const resolved = path.resolve(sourcePath);
  const ext = getFileExtension(resolved).replace(".", "");
  const baseName = path.basename(resolved, path.extname(resolved)).replace(/[^\w.\-가-힣]/g, "_").slice(0, 56);
  const digest = buildHwpSelectionSourceSignature(resolved).slice(0, 12);
  return path.join(ensureConvertedDir(), `${baseName}-${ext}-hwp-selector-${digest}.json`);
}

function readHwpSelectionMeta(sourcePath) {
  try {
    const metaPath = createHwpSelectionMetaPath(sourcePath);
    if (!fs.existsSync(metaPath)) {
      return null;
    }
    const parsed = JSON.parse(fs.readFileSync(metaPath, "utf8"));
    const sourceSignature = buildHwpSelectionSourceSignature(sourcePath);
    if (String(parsed?.selectorVersion || "") !== HWP_SELECTION_META_VERSION) {
      return null;
    }
    if (String(parsed?.sourceSignature || "") !== sourceSignature) {
      return null;
    }
    const selectedPdfPath = String(parsed?.selectedPdfPath || "");
    if (!selectedPdfPath || !fs.existsSync(selectedPdfPath)) {
      return null;
    }
    return {
      selectedMode: String(parsed?.selectedMode || ""),
      selectedPdfPath,
      warningMessage: String(parsed?.warningMessage || ""),
      fallbackReason: String(parsed?.fallbackReason || ""),
      analysisMode: String(parsed?.analysisMode || "quick"),
      candidateScores: Array.isArray(parsed?.candidateScores) ? parsed.candidateScores : []
    };
  } catch (_error) {
    return null;
  }
}

function writeHwpSelectionMeta(sourcePath, payload = {}) {
  try {
    const metaPath = createHwpSelectionMetaPath(sourcePath);
    const entry = {
      selectorVersion: HWP_SELECTION_META_VERSION,
      sourceSignature: buildHwpSelectionSourceSignature(sourcePath),
      selectedMode: String(payload.selectedMode || ""),
      selectedPdfPath: String(payload.selectedPdfPath || ""),
      warningMessage: String(payload.warningMessage || ""),
      fallbackReason: String(payload.fallbackReason || ""),
      analysisMode: String(payload.analysisMode || "quick"),
      candidateScores: Array.isArray(payload.candidateScores) ? payload.candidateScores : [],
      updatedAt: new Date().toISOString()
    };
    fs.writeFileSync(metaPath, JSON.stringify(entry, null, 2), "utf8");
  } catch (_error) {
    // ignore selector meta write failures
  }
}

function readCachedPdf(cachePath, options = {}) {
  if (!cachePath || !fs.existsSync(cachePath)) {
    return null;
  }
  if (Number.isFinite(options.maxAgeMs) && options.maxAgeMs > 0) {
    try {
      const stat = fs.statSync(cachePath);
      const ageMs = Date.now() - stat.mtimeMs;
      if (ageMs > options.maxAgeMs) {
        return null;
      }
    } catch (_error) {
      return null;
    }
  }
  return cachePath;
}

function detectSpreadsheetVisualAssets(filePath) {
  const source = escapeForPowerShell(filePath);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "Add-Type -AssemblyName 'System.IO.Compression.FileSystem'",
    `$zip = [System.IO.Compression.ZipFile]::OpenRead(${source})`,
    "try {",
    "  $hasVisual = $false",
    "  foreach ($entry in $zip.Entries) {",
    "    $name = $entry.FullName.ToLowerInvariant()",
    "    if ($name.StartsWith('xl/media/') -or $name.StartsWith('xl/charts/') -or $name.StartsWith('xl/drawings/')) {",
    "      $hasVisual = $true",
    "      break",
    "    }",
    "  }",
    "  if ($hasVisual) { '1' } else { '0' }",
    "} finally {",
    "  if ($zip -ne $null) { $zip.Dispose() }",
    "}"
  ].join("; ");
  const result = runPowerShell(script, 20000);
  if (!result.ok) {
    return { inspected: false, hasVisual: false };
  }
  return { inspected: true, hasVisual: String(result.stdout || "").trim() === "1" };
}

function inspectSpreadsheetContent(filePath) {
  try {
    const XLSX = ensureXlsxModule();
    const workbook = XLSX.readFile(filePath, { cellDates: false, dense: true });
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        continue;
      }
      for (const cellKey of Object.keys(sheet)) {
        if (!cellKey || cellKey.startsWith("!")) {
          continue;
        }
        const cell = sheet[cellKey];
        if (!cell) {
          continue;
        }
        const rendered = String(cell.w ?? cell.v ?? "").trim();
        if (rendered || cell.f) {
          return { hasContent: true, reason: "cell-data" };
        }
      }
    }
    const visual = detectSpreadsheetVisualAssets(filePath);
    if (visual.hasVisual) {
      return { hasContent: true, reason: "visual-assets" };
    }
    if (visual.inspected) {
      return { hasContent: false, reason: "empty-document" };
    }
    return { hasContent: true, reason: "inspection-unavailable" };
  } catch (_error) {
    return { hasContent: true, reason: "parse-error" };
  }
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
  if (/empty_document|표시할 셀 내용이 없습니다|empty document/i.test(message)) {
    return "EMPTY_DOCUMENT";
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

async function runPowerShellAsync(script, timeoutMs = 180000) {
  const result = await runProcess(
    "powershell",
    ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
    timeoutMs
  );
  return {
    ok: Boolean(result.ok),
    status: result.status,
    timedOut: Boolean(result.timedOut),
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
    error: String(result.error || "")
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

function logConvertTrace(stage, payload = {}) {
  try {
    const record = {
      stage: String(stage || ""),
      ts: new Date().toISOString(),
      ...payload
    };
    console.info(`[lookup-convert] ${JSON.stringify(record)}`);
  } catch (_error) {
    // ignore logging errors
  }
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .trim();
}

async function isLikelyBookletPdf(pdfPath) {
  try {
    if (!pdfPath || !fs.existsSync(pdfPath)) {
      return false;
    }
    const bytes = await fsp.readFile(pdfPath);
    const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const pages = pdfDoc.getPages();
    if (!pages.length) {
      return false;
    }
    const sample = pages.slice(0, Math.min(3, pages.length));
    let wideCount = 0;
    for (const page of sample) {
      const size = page.getSize();
      const ratio = Number(size.width || 0) / Math.max(1, Number(size.height || 0));
      if (ratio >= 1.45) {
        wideCount += 1;
      }
    }
    return wideCount >= 1;
  } catch (_error) {
    return false;
  }
}

function clampNumber(value, min, max) {
  return Math.min(max, Math.max(min, Number(value) || 0));
}

function normalizeTextRect(item) {
  const transform = Array.isArray(item?.transform) ? item.transform : [];
  const x = Number(transform[4] || 0);
  const y = Number(transform[5] || 0);
  const height = Math.max(1, Math.abs(Number(transform[3] || item?.height || 0)));
  const width = Math.max(1, Math.abs(Number(item?.width || 0)));
  return {
    xMin: x,
    xMax: x + width,
    yMin: y - height,
    yMax: y,
    width,
    height
  };
}

function horizontalOverlapWidth(aMin, aMax, bMin, bMax) {
  return Math.max(0, Math.min(aMax, bMax) - Math.max(aMin, bMin));
}

function countLineTextOverlaps(segment, textRects) {
  let matches = 0;
  for (const rect of textRects) {
    const overlap = horizontalOverlapWidth(segment.xMin, segment.xMax, rect.xMin, rect.xMax);
    if (overlap <= 0) {
      continue;
    }
    const overlapRatio = overlap / Math.max(1, rect.width);
    if (overlapRatio < 0.3) {
      continue;
    }
    const yTolerance = Math.max(1.5, rect.height * 0.45);
    if (segment.y >= rect.yMin - yTolerance && segment.y <= rect.yMax + yTolerance) {
      matches += 1;
    }
  }
  return matches;
}

function estimateRasterLineArtifacts(segments, textRects, pageWidth) {
  let rasterLineHitCount = 0;
  let textPierceRatio = 0;
  const safePageWidth = Math.max(1, Number(pageWidth || 0));
  for (const segment of segments) {
    const coverage = Number(segment.length || 0) / safePageWidth;
    if (coverage < 0.35) {
      continue;
    }
    let piercedRects = 0;
    let piercedWidth = 0;
    for (const rect of textRects) {
      const overlap = horizontalOverlapWidth(segment.xMin, segment.xMax, rect.xMin, rect.xMax);
      if (overlap <= 0) {
        continue;
      }
      const overlapRatio = overlap / Math.max(1, rect.width);
      if (overlapRatio < 0.35) {
        continue;
      }
      const centerY = (rect.yMin + rect.yMax) / 2;
      const centerTolerance = Math.max(1.2, rect.height * 0.28);
      if (Math.abs(segment.y - centerY) <= centerTolerance) {
        piercedRects += 1;
        piercedWidth += overlap;
      }
    }
    if (piercedRects >= 2 && coverage >= 0.42) {
      rasterLineHitCount += 1;
      textPierceRatio += clampNumber((piercedWidth / safePageWidth) * 0.55 + coverage * 0.45, 0, 1);
    }
  }
  return {
    rasterLineHitCount,
    textLinePierceRatio: clampNumber(textPierceRatio, 0, 1)
  };
}

function extractHorizontalSegmentsFromOperatorList(pdfjsLib, operatorList, pageWidth) {
  const OPS = pdfjsLib.OPS || {};
  const segments = [];
  const minLength = Math.max(80, Number(pageWidth || 0) * 0.22);
  const fnArray = Array.isArray(operatorList?.fnArray) ? operatorList.fnArray : [];
  const argsArray = Array.isArray(operatorList?.argsArray) ? operatorList.argsArray : [];
  for (let i = 0; i < fnArray.length; i += 1) {
    if (fnArray[i] !== OPS.constructPath) {
      continue;
    }
    const entry = argsArray[i] || [];
    const pathOps = Array.isArray(entry[0]) ? entry[0] : [];
    const pathArgs = Array.isArray(entry[1]) ? entry[1] : [];
    let argIdx = 0;
    let currentX = 0;
    let currentY = 0;
    let startX = 0;
    let startY = 0;
    for (const pathOp of pathOps) {
      if (pathOp === OPS.moveTo) {
        currentX = Number(pathArgs[argIdx++] || 0);
        currentY = Number(pathArgs[argIdx++] || 0);
        startX = currentX;
        startY = currentY;
        continue;
      }
      if (pathOp === OPS.lineTo) {
        const nextX = Number(pathArgs[argIdx++] || 0);
        const nextY = Number(pathArgs[argIdx++] || 0);
        const dx = Math.abs(nextX - currentX);
        const dy = Math.abs(nextY - currentY);
        if (dx >= minLength && dy <= 1.2) {
          segments.push({
            xMin: Math.min(currentX, nextX),
            xMax: Math.max(currentX, nextX),
            y: (currentY + nextY) / 2,
            length: dx
          });
        }
        currentX = nextX;
        currentY = nextY;
        continue;
      }
      if (pathOp === OPS.closePath) {
        const dx = Math.abs(startX - currentX);
        const dy = Math.abs(startY - currentY);
        if (dx >= minLength && dy <= 1.2) {
          segments.push({
            xMin: Math.min(currentX, startX),
            xMax: Math.max(currentX, startX),
            y: (currentY + startY) / 2,
            length: dx
          });
        }
        currentX = startX;
        currentY = startY;
        continue;
      }
      if (pathOp === OPS.rectangle) {
        argIdx += 4;
        continue;
      }
      if (pathOp === OPS.curveTo) {
        argIdx += 6;
        continue;
      }
      if (pathOp === OPS.curveTo2 || pathOp === OPS.curveTo3) {
        argIdx += 4;
        continue;
      }
    }
  }
  return segments;
}

async function analyzeHwpPdfQuality(pdfPath, options = {}) {
  const analysisMode = String(options.analysisMode || "full").toLowerCase() === "quick" ? "quick" : "full";
  const maxPages = Number(options.maxPages || (analysisMode === "quick" ? 1 : 2));
  if (!pdfPath || !fs.existsSync(pdfPath)) {
    return {
      inspected: false,
      analysisMode,
      score: 0,
      metrics: {
        longLineCount: 0,
        overlapLineCount: 0,
        rasterLineHitCount: 0,
        textLinePierceRatio: 0,
        readableChars: 0,
        bookletLikePages: 0,
        sampledPages: 0,
        vectorLineScore: 0,
        rasterLineScore: 0,
        finalScore: 0
      }
    };
  }
  try {
    const pdfjsLib = await ensureNodePdfJs();
    const loadingTask = pdfjsLib.getDocument({
      url: pathToFileURL(pdfPath).href,
      disableWorker: true,
      stopAtErrors: false,
      useSystemFonts: true
    });
    const doc = await loadingTask.promise;
    const sampleCount = Math.min(maxPages, Number(doc.numPages || 0));
    let longLineCount = 0;
    let overlapLineCount = 0;
    let rasterLineHitCount = 0;
    let textLinePierceRatioAccum = 0;
    let readableChars = 0;
    let bookletLikePages = 0;
    for (let pageNumber = 1; pageNumber <= sampleCount; pageNumber += 1) {
      const page = await doc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      const pageRatio = Number(viewport.width || 0) / Math.max(1, Number(viewport.height || 0));
      if (pageRatio >= 1.45) {
        bookletLikePages += 1;
      }
      const textContent = await page.getTextContent({ disableCombineTextItems: false });
      const textRects = [];
      for (const item of textContent.items || []) {
        const str = String(item?.str || "");
        if (!str.trim()) {
          continue;
        }
        readableChars += str.replace(/\s+/g, "").length;
        textRects.push(normalizeTextRect(item));
      }
      const operatorList = await page.getOperatorList();
      const segments = extractHorizontalSegmentsFromOperatorList(pdfjsLib, operatorList, viewport.width);
      longLineCount += segments.length;
      for (const segment of segments) {
        const overlaps = countLineTextOverlaps(segment, textRects);
        if (overlaps >= 2) {
          overlapLineCount += 1;
        }
      }
      const rasterEstimate = estimateRasterLineArtifacts(segments, textRects, viewport.width);
      rasterLineHitCount += rasterEstimate.rasterLineHitCount;
      textLinePierceRatioAccum += rasterEstimate.textLinePierceRatio;
      try {
        page.cleanup();
      } catch (_error) {
        // ignore page cleanup failure
      }
    }
    await doc.destroy();
    try {
      await loadingTask.destroy();
    } catch (_error) {
      // ignore task cleanup failure
    }
    const textLinePierceRatio = sampleCount > 0 ? clampNumber(textLinePierceRatioAccum / sampleCount, 0, 1) : 0;
    let vectorLineScore = 100;
    vectorLineScore -= clampNumber(overlapLineCount * 10, 0, 60);
    vectorLineScore -= clampNumber((longLineCount - overlapLineCount) * 1.4, 0, 20);
    vectorLineScore -= clampNumber(bookletLikePages * 22, 0, 44);
    if (readableChars < 20) {
      vectorLineScore -= 8;
    }
    vectorLineScore = clampNumber(vectorLineScore, 0, 100);
    let rasterLineScore = 100;
    rasterLineScore -= clampNumber(rasterLineHitCount * 14, 0, 62);
    rasterLineScore -= clampNumber(textLinePierceRatio * 48, 0, 42);
    rasterLineScore -= clampNumber(bookletLikePages * 12, 0, 36);
    rasterLineScore = clampNumber(rasterLineScore, 0, 100);
    const finalScore = clampNumber(Math.round(vectorLineScore * 0.62 + rasterLineScore * 0.38), 0, 100);
    return {
      inspected: true,
      analysisMode,
      score: finalScore,
      metrics: {
        longLineCount,
        overlapLineCount,
        rasterLineHitCount,
        textLinePierceRatio,
        readableChars,
        bookletLikePages,
        sampledPages: sampleCount,
        vectorLineScore,
        rasterLineScore,
        finalScore
      }
    };
  } catch (_error) {
    return {
      inspected: false,
      analysisMode,
      score: 0,
      metrics: {
        longLineCount: 0,
        overlapLineCount: 0,
        rasterLineHitCount: 0,
        textLinePierceRatio: 0,
        readableChars: 0,
        bookletLikePages: 0,
        sampledPages: 0,
        vectorLineScore: 0,
        rasterLineScore: 0,
        finalScore: 0
      }
    };
  }
}

function chooseBestHwpCandidate(candidates) {
  const valid = candidates.filter((candidate) => candidate.available);
  if (!valid.length) {
    return null;
  }
  valid.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    if (a.mode === "hwp-saveas" && b.mode !== "hwp-saveas") {
      return -1;
    }
    if (b.mode === "hwp-saveas" && a.mode !== "hwp-saveas") {
      return 1;
    }
    return 0;
  });
  return valid[0];
}

async function convertWithWordCom(sourcePath, outputPdfPath) {
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

  return await runPowerShellAsync(script, 180000);
}

async function convertWithExcelCom(sourcePath, outputPdfPath) {
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

  return await runPowerShellAsync(script, 180000);
}

async function convertWithHancomComPrintPdf(sourcePath, outputPdfPath, ext = ".hwp") {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "$moduleRoot = 'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules'",
    "$moduleName = 'lookupFilePathChecker'",
    "$modulePath = ''",
    "$registered = $false",
    "try {",
    "  if (-not (Test-Path $moduleRoot)) { New-Item -Path $moduleRoot -Force | Out-Null }",
    "  $regValue = Get-ItemProperty -Path $moduleRoot -Name $moduleName -ErrorAction SilentlyContinue",
    "  if ($regValue -and $regValue.$moduleName) { $modulePath = [string]$regValue.$moduleName }",
    "  if (-not $modulePath -or -not (Test-Path $modulePath)) {",
    "    $candidates = @(",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\HOffice\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hnc\\Office\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\HOffice\\Bin\\FilePathCheckerModule.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\HOffice\\Bin\\FilePathCheckerModule.dll'",
    "    )",
    "    foreach ($candidate in $candidates) {",
    "      if ($candidate -and (Test-Path $candidate)) {",
    "        $modulePath = $candidate",
    "        break",
    "      }",
    "    }",
    "    if ($modulePath -and (Test-Path $modulePath)) {",
    "      Set-ItemProperty -Path $moduleRoot -Name $moduleName -Value $modulePath -Type String -Force",
    "    }",
    "  }",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  if ($modulePath -and (Test-Path $modulePath)) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', $moduleName) | Out-Null; $registered = $true } catch {}",
    "  }",
    "  if (-not $registered) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null; $registered = $true } catch {}",
    "  }",
    "  if (-not $registered) { throw 'hwp_security_module_missing' }",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { throw 'hwp_open_failed' }",
    "  try { $hwp.Run('Cancel') | Out-Null } catch {}",
    "  try { $hwp.HAction.GetDefault('Print', $hwp.HParameterSet.HPrint.HSet) | Out-Null } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintToFile = 1 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.FileName = " + output + " } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrnName = 'Microsoft Print to PDF' } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrinterName = 'Microsoft Print to PDF' } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintMethod = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintImage = 1 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.Collate = 1 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.NumCopy = 1 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.ReverseOrder = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.Pause = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.MarkPen = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.Memo = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintMemo = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintRevision = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.PrintGuideLine = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.ShowMarkPen = 0 } catch {}",
    "  try { $hwp.HParameterSet.HPrint.DrawObj = 1 } catch {}",
    "  try { Write-Output ('lookup_hprint_snapshot=' + $hwp.HParameterSet.HPrint.PrintToFile + '|' + $hwp.HParameterSet.HPrint.PrintRevision + '|' + $hwp.HParameterSet.HPrint.PrintGuideLine + '|' + $hwp.HParameterSet.HPrint.ShowMarkPen + '|' + $hwp.HParameterSet.HPrint.Memo) } catch {}",
    "  $printed = $false",
    "  try { $printed = $hwp.HAction.Execute('Print', $hwp.HParameterSet.HPrint.HSet) } catch { $printed = $false }",
    "  if (-not $printed) { throw 'hwp_print_failed' }",
    "  Write-Output 'lookup_hwp_registered=1'",
    "} finally {",
    "  if ($hwp -ne $null) { try { $hwp.Quit() | Out-Null } catch {} }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("; ");
  return await runPowerShellAsync(script, 120000);
}

async function convertWithHancomComSaveAs(sourcePath, outputPdfPath, ext = ".hwp") {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "$moduleRoot = 'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules'",
    "$moduleName = 'lookupFilePathChecker'",
    "$modulePath = ''",
    "$registered = $false",
    "try {",
    "  if (-not (Test-Path $moduleRoot)) { New-Item -Path $moduleRoot -Force | Out-Null }",
    "  $regValue = Get-ItemProperty -Path $moduleRoot -Name $moduleName -ErrorAction SilentlyContinue",
    "  if ($regValue -and $regValue.$moduleName) { $modulePath = [string]$regValue.$moduleName }",
    "  if (-not $modulePath -or -not (Test-Path $modulePath)) {",
    "    $candidates = @(",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\HOffice\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hnc\\Office\\Bin\\FilePathCheckerModuleExample.dll',",
    "      Join-Path $env:ProgramFiles 'Hancom\\HOffice\\Bin\\FilePathCheckerModule.dll',",
    "      Join-Path $env:ProgramFiles 'HNC\\HOffice\\Bin\\FilePathCheckerModule.dll'",
    "    )",
    "    foreach ($candidate in $candidates) {",
    "      if ($candidate -and (Test-Path $candidate)) {",
    "        $modulePath = $candidate",
    "        break",
    "      }",
    "    }",
    "    if ($modulePath -and (Test-Path $modulePath)) {",
    "      Set-ItemProperty -Path $moduleRoot -Name $moduleName -Value $modulePath -Type String -Force",
    "    }",
    "  }",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  if ($modulePath -and (Test-Path $modulePath)) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', $moduleName) | Out-Null; $registered = $true } catch {}",
    "  }",
    "  if (-not $registered) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null; $registered = $true } catch {}",
    "  }",
    "  if (-not $registered) { throw 'hwp_security_module_missing' }",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { throw 'hwp_open_failed' }",
    "  $saved = $hwp.SaveAs(" + output + ", 'PDF', 'embedfont:true;bookmark:true;compress:false')",
    "  if (-not $saved) { throw 'hwp_save_failed' }",
    "  Write-Output 'lookup_hwp_registered=1'",
    "} finally {",
    "  if ($hwp -ne $null) { try { $hwp.Quit() | Out-Null } catch {} }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("; ");
  return await runPowerShellAsync(script, 90000);
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

function readSpreadsheetSheetNames(sourcePath) {
  try {
    const XLSX = ensureXlsxModule();
    const workbook = XLSX.readFile(sourcePath, { cellDates: false, dense: true });
    return (workbook.SheetNames || []).map((name) => String(name || "").trim()).filter(Boolean);
  } catch (_error) {
    return [];
  }
}

function buildSpreadsheetFallbackHtml({ sourcePath, sourceExt, warningMessage }) {
  const XLSX = ensureXlsxModule();
  const workbook = XLSX.readFile(sourcePath, { cellDates: false, dense: false });
  const sections = [];
  const sheetMap = [];
  const maxRows = 300;
  const maxCols = 32;
  const chunkRows = 42;
  let nextPageNumber = 1;
  for (const sheetName of workbook.SheetNames || []) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet["!ref"]) {
      continue;
    }
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const startRow = range.s.r;
    const endRow = Math.min(range.e.r, startRow + maxRows - 1);
    const startCol = range.s.c;
    const endCol = Math.min(range.e.c, startCol + maxCols - 1);

    const htmlRows = [];
    let hasAnyValue = false;
    for (let row = startRow; row <= endRow; row += 1) {
      let rowCells = "";
      let rowHasValue = false;
      for (let col = startCol; col <= endCol; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = sheet[address];
        const cellText = cell ? String(cell.w ?? XLSX.utils.format_cell(cell) ?? cell.v ?? "").trim() : "";
        if (cellText) {
          rowHasValue = true;
          hasAnyValue = true;
        }
        rowCells += `<td>${escapeHtml(cellText)}</td>`;
      }
      if (rowHasValue || row === startRow) {
        htmlRows.push(`<tr>${rowCells}</tr>`);
      }
    }

    if (!htmlRows.length || !hasAnyValue) {
      continue;
    }
    const pageCountForSheet = Math.max(1, Math.ceil(htmlRows.length / chunkRows));
    const sheetStartPage = nextPageNumber;
    for (let chunkIndex = 0; chunkIndex < pageCountForSheet; chunkIndex += 1) {
      const chunkStart = chunkIndex * chunkRows;
      const chunkEnd = chunkStart + chunkRows;
      const rowsChunk = htmlRows.slice(chunkStart, chunkEnd).join("");
      const chunkTitle =
        pageCountForSheet > 1 ? `${escapeHtml(sheetName)} (${chunkIndex + 1}/${pageCountForSheet})` : escapeHtml(sheetName);
      sections.push(`<section class="sheet"><h2>${chunkTitle}</h2><table>${rowsChunk}</table></section>`);
      nextPageNumber += 1;
    }
    sheetMap.push({
      sheetName,
      startPage: sheetStartPage,
      endPage: nextPageNumber - 1
    });
  }

  if (!sections.length) {
    return {
      html: "",
      sheetMap: []
    };
  }
  const safeTitle = escapeHtml(path.basename(sourcePath));
  const safeExt = escapeHtml(sourceExt.replace(".", "").toUpperCase());
  const warning = warningMessage ? `<div class="warning">${escapeHtml(warningMessage)}</div>` : "";
  return {
    html: `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <style>
    @page { size: A4; margin: 14mm; }
    html, body { margin: 0; padding: 0; font-family: 'Malgun Gothic', 'Segoe UI', sans-serif; color: #1f2937; }
    .head { margin-bottom: 12px; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px; }
    .title { font-size: 16px; font-weight: 700; margin: 0 0 4px; }
    .meta { font-size: 11px; color: #475569; }
    .warning { margin: 10px 0 12px; padding: 8px 10px; border: 1px solid #f2d6a8; background: #fff7ea; color: #9a5f08; font-size: 12px; border-radius: 6px; }
    .sheet { page-break-after: always; }
    .sheet:last-child { page-break-after: auto; }
    .sheet h2 { margin: 0 0 8px; font-size: 14px; color: #0f172a; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10.5px; }
    td { border: 1px solid #d6deea; padding: 4px 5px; vertical-align: top; word-break: break-word; white-space: pre-wrap; }
  </style>
</head>
<body>
  <div class="head">
    <p class="title">${safeTitle}</p>
    <div class="meta">원본 형식: ${safeExt} / lookup 표 변환 보기</div>
  </div>
  ${warning}
  ${sections.join("\n")}
</body>
</html>`,
    sheetMap
  };
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

async function convertOfficeFileToPdf(sourcePath, ext) {
  const spreadsheetSheetNames = ext === ".xls" || ext === ".xlsx" ? readSpreadsheetSheetNames(sourcePath) : [];
  const spreadsheetUnknownSheetMap = spreadsheetSheetNames.map((sheetName) => ({
    sheetName,
    startPage: null,
    endPage: null
  }));
  if (ext === ".xls" || ext === ".xlsx") {
    const spreadsheetInspection = inspectSpreadsheetContent(sourcePath);
    if (!spreadsheetInspection.hasContent) {
      return {
        ok: false,
        engine: "empty",
        errorCode: "EMPTY_DOCUMENT",
        warningMessage: "문서에 표시할 셀 내용이 없습니다.",
        emptyDocument: true,
        sheetMap: spreadsheetUnknownSheetMap
      };
    }
  }

  const officePdfPath = createConvertedPdfPath(sourcePath, "office-com");
  const cachedOffice = readCachedPdf(officePdfPath);
  if (cachedOffice) {
    return {
      ok: true,
      engine: "office-com",
      warningMessage: "",
      pdfPath: cachedOffice,
      fromCache: true,
      errorCode: "",
      sheetMap: spreadsheetUnknownSheetMap
    };
  }

  let comResult = { ok: false, stderr: "unsupported_office_ext" };
  if (ext === ".doc" || ext === ".docx") {
    comResult = await convertWithWordCom(sourcePath, officePdfPath);
  } else if (ext === ".xls" || ext === ".xlsx") {
    comResult = await convertWithExcelCom(sourcePath, officePdfPath);
  }

  if (comResult.ok && fs.existsSync(officePdfPath)) {
    return {
      ok: true,
      engine: "office-com",
      warningMessage: "",
      pdfPath: officePdfPath,
      fromCache: false,
      errorCode: "",
      sheetMap: spreadsheetUnknownSheetMap
    };
  }

  const librePdfPath = createConvertedPdfPath(sourcePath, "libreoffice");
  const cachedLibre = readCachedPdf(librePdfPath);
  if (cachedLibre) {
    return {
      ok: true,
      engine: "libreoffice",
      warningMessage: "Office 기본 변환에 실패해 LibreOffice 호환 엔진으로 열었습니다.",
      pdfPath: cachedLibre,
      fromCache: true,
      errorCode: "",
      sheetMap: spreadsheetUnknownSheetMap
    };
  }

  const libreResult = await convertWithLibreOffice(sourcePath, librePdfPath);
  if (libreResult.ok && fs.existsSync(librePdfPath)) {
    return {
      ok: true,
      engine: "libreoffice",
      warningMessage: "Office 기본 변환에 실패해 LibreOffice 호환 엔진으로 열었습니다.",
      pdfPath: librePdfPath,
      fromCache: false,
      errorCode: "",
      sheetMap: spreadsheetUnknownSheetMap
    };
  }

  const timedOut = isTimeoutResult(comResult) || isTimeoutResult(libreResult);
  const reasonHint = timedOut
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
    warningMessage: buildConversionWarning(ext, details, reasonHint),
    errorCode: timedOut ? "ENGINE_TIMEOUT" : libreResult.missingEngine ? "ENGINE_MISSING" : "CONVERT_FAILED",
    emptyDocument: false,
    sheetMap: spreadsheetUnknownSheetMap
  };
}

async function collectHwpCandidate({ mode, pdfPath, fromCache, runResult, analysisMode = "full" }) {
  const exists = Boolean(pdfPath && fs.existsSync(pdfPath));
  if (!exists) {
    return {
      mode,
      path: pdfPath,
      available: false,
      fromCache: Boolean(fromCache),
      runResult: runResult || null,
      analysis: {
        inspected: false,
        analysisMode,
        score: 0,
        metrics: {
          longLineCount: 0,
          overlapLineCount: 0,
          rasterLineHitCount: 0,
          textLinePierceRatio: 0,
          readableChars: 0,
          bookletLikePages: 0,
          sampledPages: 0,
          vectorLineScore: 0,
          rasterLineScore: 0,
          finalScore: 0
        }
      },
      bookletLike: false,
      score: 0
    };
  }
  const normalizedMode = String(analysisMode || "full").toLowerCase() === "quick" ? "quick" : "full";
  const analysis = await analyzeHwpPdfQuality(pdfPath, {
    analysisMode: normalizedMode,
    maxPages: normalizedMode === "quick" ? 1 : 2
  });
  const bookletLike = Boolean(analysis?.metrics?.bookletLikePages > 0 || (await isLikelyBookletPdf(pdfPath)));
  let score = Number(analysis?.score || 0);
  if (bookletLike) {
    score = Math.max(0, score - 30);
  }
  if (mode === "hwp-saveas") {
    score += 1;
  }
  return {
    mode,
    path: pdfPath,
    available: true,
    fromCache: Boolean(fromCache),
    runResult: runResult || null,
    analysis,
    bookletLike,
    score: clampNumber(score, 0, 100)
  };
}

function buildHwpSelectionWarning(selected, candidates) {
  if (!selected) {
    return "";
  }
  if (selected.mode === "hwp-print-pdf" && selected.score >= 70) {
    return "";
  }
  const printCandidate = candidates.find((candidate) => candidate.mode === "hwp-print-pdf");
  if (selected.mode === "hwp-saveas" && printCandidate?.available) {
    if (printCandidate.bookletLike || Number(printCandidate.analysis?.metrics?.overlapLineCount || 0) >= 2) {
      return "변환 결과에 표시 오류가 감지되어 다른 변환 방식으로 자동 전환했습니다.";
    }
  }
  if (selected.mode === "hwp-saveas" && !printCandidate?.available) {
    return "한컴 변환 설정 확인이 필요합니다. 호환 보기로 열었습니다.";
  }
  if (selected.mode === "hwpjs-html") {
    return "한컴 변환 설정 확인이 필요합니다. 호환 보기로 열었습니다.";
  }
  return "";
}

function buildHwpCandidateScoreRows(candidates) {
  return candidates.map((candidate) => ({
    mode: candidate.mode,
    score: candidate.score,
    fromCache: candidate.fromCache,
    bookletLike: candidate.bookletLike,
    inspected: candidate.analysis.inspected,
    analysisMode: candidate.analysis.analysisMode || "full",
    longLineCount: candidate.analysis.metrics.longLineCount,
    overlapLineCount: candidate.analysis.metrics.overlapLineCount,
    rasterLineHitCount: candidate.analysis.metrics.rasterLineHitCount,
    textLinePierceRatio: Number(candidate.analysis.metrics.textLinePierceRatio || 0),
    readableChars: candidate.analysis.metrics.readableChars,
    vectorLineScore: Number(candidate.analysis.metrics.vectorLineScore || 0),
    rasterLineScore: Number(candidate.analysis.metrics.rasterLineScore || 0),
    finalScore: Number(candidate.analysis.metrics.finalScore || candidate.score || 0)
  }));
}

function isHwpQuickCandidateStable(candidate) {
  if (!candidate?.available) {
    return false;
  }
  const metrics = candidate.analysis?.metrics || {};
  if (candidate.bookletLike) {
    return false;
  }
  if (Number(metrics.overlapLineCount || 0) >= 2) {
    return false;
  }
  if (Number(metrics.rasterLineHitCount || 0) >= 2) {
    return false;
  }
  if (Number(metrics.textLinePierceRatio || 0) >= 0.26) {
    return false;
  }
  return Number(candidate.score || 0) >= 82;
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
      errorCode: "",
      convertDiagnostics: {
        selectedMode: "native",
        candidateScores: [],
        fallbackReason: ""
      },
      sheetMap: [],
      pdfPath: resolved,
      pdfBuffer: await fsp.readFile(resolved)
    };
  }

  let warningMessage = "";
  let errorCode = "";

  if (OFFICE_EXTENSIONS.has(ext)) {
    const officeResult = await convertOfficeFileToPdf(resolved, ext);
    if (officeResult.ok && officeResult.pdfPath && fs.existsSync(officeResult.pdfPath)) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: officeResult.engine || "office-com",
        warningMessage: officeResult.warningMessage || "",
        errorCode: "",
        convertDiagnostics: {
          selectedMode: officeResult.engine || "office-com",
          candidateScores: [],
          fallbackReason: ""
        },
        sheetMap: Array.isArray(officeResult.sheetMap) ? officeResult.sheetMap : [],
        pdfPath: officeResult.pdfPath,
        pdfBuffer: await fsp.readFile(officeResult.pdfPath)
      };
    }
    warningMessage = officeResult.warningMessage || "고품질 변환 엔진을 찾지 못해 텍스트 기반 보기 모드로 변환했습니다.";
    errorCode = officeResult.errorCode || "";
  } else if (HWP_EXTENSIONS.has(ext)) {
    logConvertTrace("hwp-open-start", { source: path.basename(resolved), ext });
    const hwpPrintPdfPath = createConvertedPdfPath(resolved, "hwp-print-pdf");
    const hwpSaveAsPdfPath = createConvertedPdfPath(resolved, "hwp-saveas");
    let hwpPrintResult = null;
    let hwpSaveAsResult = null;
    let hwpJsResult = null;
    const hwpJsPdfPath = createConvertedPdfPath(resolved, "hwpjs-html");
    const hwpCandidates = [];

    const selectorMeta = readHwpSelectionMeta(resolved);
    if (selectorMeta?.selectedMode && selectorMeta?.selectedPdfPath && fs.existsSync(selectorMeta.selectedPdfPath)) {
      logConvertTrace("hwp-meta-hit", {
        selectedMode: selectorMeta.selectedMode,
        analysisMode: selectorMeta.analysisMode || "quick",
        source: path.basename(resolved)
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: selectorMeta.selectedMode,
        warningMessage: selectorMeta.warningMessage || "",
        errorCode: selectorMeta.warningMessage ? "RENDER_ARTIFACT_DETECTED" : "",
        convertDiagnostics: {
          selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
          selectedMode: selectorMeta.selectedMode,
          analysisMode: selectorMeta.analysisMode || "quick",
          candidateScores: Array.isArray(selectorMeta.candidateScores) ? selectorMeta.candidateScores : [],
          fallbackReason: selectorMeta.fallbackReason || selectorMeta.warningMessage || ""
        },
        sheetMap: [],
        pdfPath: selectorMeta.selectedPdfPath,
        pdfBuffer: await fsp.readFile(selectorMeta.selectedPdfPath)
      };
    }

    const hwpPrintCached = readCachedPdf(hwpPrintPdfPath);
    let printCandidate = hwpPrintCached
      ? await collectHwpCandidate({
        mode: "hwp-print-pdf",
        pdfPath: hwpPrintCached,
        fromCache: true,
        analysisMode: "quick"
      })
      : await (async () => {
        hwpPrintResult = await convertWithHancomComPrintPdf(resolved, hwpPrintPdfPath, ext);
        return collectHwpCandidate({
          mode: "hwp-print-pdf",
          pdfPath: hwpPrintPdfPath,
          fromCache: false,
          runResult: hwpPrintResult,
          analysisMode: "quick"
        });
      })();
    hwpCandidates.push(printCandidate);

    if (isHwpQuickCandidateStable(printCandidate)) {
      logConvertTrace("hwp-quick-pass", {
        selectedMode: printCandidate.mode,
        score: printCandidate.score,
        source: path.basename(resolved)
      });
      const selectionWarning = buildHwpSelectionWarning(printCandidate, hwpCandidates);
      const candidateScores = buildHwpCandidateScoreRows(hwpCandidates);
      writeHwpSelectionMeta(resolved, {
        selectedMode: printCandidate.mode,
        selectedPdfPath: printCandidate.path,
        warningMessage: selectionWarning,
        fallbackReason: selectionWarning || "",
        analysisMode: "quick",
        candidateScores
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: printCandidate.mode,
        warningMessage: selectionWarning,
        errorCode: selectionWarning ? "RENDER_ARTIFACT_DETECTED" : "",
        convertDiagnostics: {
          selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
          selectedMode: printCandidate.mode,
          analysisMode: "quick",
          candidateScores,
          fallbackReason: selectionWarning || ""
        },
        sheetMap: [],
        pdfPath: printCandidate.path,
        pdfBuffer: await fsp.readFile(printCandidate.path)
      };
    }

    if (printCandidate?.available && String(printCandidate.analysis?.analysisMode || "").toLowerCase() !== "full") {
      printCandidate = await collectHwpCandidate({
        mode: "hwp-print-pdf",
        pdfPath: printCandidate.path,
        fromCache: true,
        runResult: hwpPrintResult,
        analysisMode: "full"
      });
      hwpCandidates[0] = printCandidate;
    }

    const hwpSaveAsCached = readCachedPdf(hwpSaveAsPdfPath);
    const saveAsCandidate = hwpSaveAsCached
      ? await collectHwpCandidate({
        mode: "hwp-saveas",
        pdfPath: hwpSaveAsCached,
        fromCache: true,
        analysisMode: "full"
      })
      : await (async () => {
        hwpSaveAsResult = await convertWithHancomComSaveAs(resolved, hwpSaveAsPdfPath, ext);
        return collectHwpCandidate({
          mode: "hwp-saveas",
          pdfPath: hwpSaveAsPdfPath,
          fromCache: false,
          runResult: hwpSaveAsResult,
          analysisMode: "full"
        });
      })();
    hwpCandidates.push(saveAsCandidate);

    let selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    const needHwpJsCompare = !selectedCandidate || Number(selectedCandidate.score || 0) < 68;
    if (needHwpJsCompare) {
      const cachedHwpJs = readCachedPdf(hwpJsPdfPath);
      let hwpJsCandidate = cachedHwpJs
        ? await collectHwpCandidate({
          mode: "hwpjs-html",
          pdfPath: cachedHwpJs,
          fromCache: true,
          analysisMode: "full"
        })
        : null;
      if (!hwpJsCandidate) {
        hwpJsResult = await convertHwpToPdfWithHwpJs(resolved, hwpJsPdfPath);
        if (hwpJsResult.ok && fs.existsSync(hwpJsPdfPath)) {
          hwpJsCandidate = await collectHwpCandidate({
            mode: "hwpjs-html",
            pdfPath: hwpJsPdfPath,
            fromCache: false,
            runResult: {
              ok: true,
              stdout: "",
              stderr: "",
              error: ""
            },
            analysisMode: "full"
          });
        }
      }
      if (hwpJsCandidate) {
        hwpCandidates.push(hwpJsCandidate);
      }
      selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    }

    if (selectedCandidate?.available && selectedCandidate.path && fs.existsSync(selectedCandidate.path)) {
      const selectedMode = selectedCandidate.mode;
      const selectionWarning = buildHwpSelectionWarning(selectedCandidate, hwpCandidates);
      logConvertTrace("hwp-select-final", {
        selectedMode,
        score: selectedCandidate.score,
        analysisMode: selectedCandidate.analysis?.analysisMode || "full",
        source: path.basename(resolved)
      });
      const candidateScores = buildHwpCandidateScoreRows(hwpCandidates);
      writeHwpSelectionMeta(resolved, {
        selectedMode,
        selectedPdfPath: selectedCandidate.path,
        warningMessage: selectionWarning,
        fallbackReason: selectionWarning || "",
        analysisMode: selectedCandidate.analysis?.analysisMode || "full",
        candidateScores
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: selectedMode,
        warningMessage: selectionWarning,
        errorCode: selectionWarning ? "RENDER_ARTIFACT_DETECTED" : "",
        convertDiagnostics: {
          selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
          selectedMode,
          analysisMode: selectedCandidate.analysis?.analysisMode || "full",
          candidateScores,
          fallbackReason: selectionWarning || ""
        },
        sheetMap: [],
        pdfPath: selectedCandidate.path,
        pdfBuffer: await fsp.readFile(selectedCandidate.path)
      };
    }

    const hwpErrorDetail = [
      hwpPrintResult?.stderr,
      hwpPrintResult?.error,
      hwpCandidates.find((candidate) => candidate.mode === "hwp-print-pdf")?.runResult?.stderr,
      hwpSaveAsResult?.stderr,
      hwpSaveAsResult?.error,
      hwpCandidates.find((candidate) => candidate.mode === "hwp-saveas")?.runResult?.stderr,
      hwpJsResult?.reason
    ]
      .map((value) => String(value || "").trim())
      .find(Boolean);
    const hwpTimedOut = isTimeoutResult(hwpPrintResult) || isTimeoutResult(hwpSaveAsResult);
    const hwpSecurityMissing = /hwp_security_module_missing/i.test(
      `${String(hwpPrintResult?.stderr || "")} ${String(hwpPrintResult?.error || "")} ${String(hwpSaveAsResult?.stderr || "")} ${String(hwpSaveAsResult?.error || "")}`
    );
    const hwpRenderArtifact = hwpCandidates.some(
      (candidate) => candidate.bookletLike
        || Number(candidate.analysis?.metrics?.overlapLineCount || 0) >= 2
        || Number(candidate.analysis?.metrics?.rasterLineHitCount || 0) >= 2
        || Number(candidate.analysis?.metrics?.textLinePierceRatio || 0) >= 0.26
    );
    const reasonHint = hwpSecurityMissing
      ? "한컴 보안 모듈이 준비되지 않아 텍스트 기반 보기로 전환했습니다."
      : hwpTimedOut
        ? "한컴 변환 엔진 응답이 지연되어 텍스트 기반 보기로 전환했습니다."
        : hwpRenderArtifact
          ? "변환 결과에서 비정상 페이지 패턴이 감지되어 텍스트 기반 보기로 전환했습니다."
          : "원본 레이아웃 변환에 실패해 텍스트 기반 보기로 전환했습니다.";
    warningMessage = buildConversionWarning(ext, hwpErrorDetail, reasonHint);
    errorCode = hwpSecurityMissing
      ? "ENGINE_MISSING"
      : hwpTimedOut
        ? "ENGINE_TIMEOUT"
        : hwpRenderArtifact
          ? "RENDER_ARTIFACT_DETECTED"
          : "CONVERT_FAILED";
    const hwpDiagnosticPayload = {
      selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
      selectedMode: "fallback",
      analysisMode: "full",
      candidateScores: buildHwpCandidateScoreRows(hwpCandidates),
      fallbackReason: reasonHint
    };
    logConvertTrace("hwp-fallback", {
      errorCode,
      fallbackReason: reasonHint,
      source: path.basename(resolved)
    });
    const fallbackPdfPath = createConvertedPdfPath(resolved, "fallback");
    const fallbackCache = readCachedPdf(fallbackPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
    if (fallbackCache) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: "fallback",
        warningMessage,
        errorCode,
        convertDiagnostics: hwpDiagnosticPayload,
        sheetMap: [],
        pdfPath: fallbackCache,
        pdfBuffer: await fsp.readFile(fallbackCache)
      };
    }
    const extractedText = await extractDocumentTextForFallback(resolved, ext);
    const html = buildFallbackHtml({
      sourcePath: resolved,
      sourceExt: ext,
      contentText: extractedText || "",
      warningMessage
    });
    const fallbackPdf = await htmlToPdfBuffer(html);
    await fsp.writeFile(fallbackPdfPath, fallbackPdf);
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: true,
      convertMode: "fallback",
      warningMessage,
      errorCode,
      convertDiagnostics: hwpDiagnosticPayload,
      sheetMap: [],
      pdfPath: fallbackPdfPath,
      pdfBuffer: fallbackPdf
    };
  }

  if (ext === ".xls" || ext === ".xlsx") {
    const xlsxHtmlPdfPath = createConvertedPdfPath(resolved, "xlsx-html");
    const xlsxHtmlCache = readCachedPdf(xlsxHtmlPdfPath);
    if (xlsxHtmlCache) {
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: "xlsx-html",
        warningMessage: warningMessage || "표 보기 모드로 열었습니다.",
        errorCode: "",
        convertDiagnostics: {
          selectedMode: "xlsx-html",
          candidateScores: [],
          fallbackReason: warningMessage || ""
        },
        sheetMap: readSpreadsheetSheetNames(resolved).map((sheetName) => ({ sheetName, startPage: null, endPage: null })),
        pdfPath: xlsxHtmlCache,
        pdfBuffer: await fsp.readFile(xlsxHtmlCache)
      };
    }
    if (errorCode !== "EMPTY_DOCUMENT") {
      try {
        const xlsxFallback = buildSpreadsheetFallbackHtml({
          sourcePath: resolved,
          sourceExt: ext,
          warningMessage: warningMessage || "고품질 변환 엔진을 찾지 못해 표 보기 모드로 전환했습니다."
        });
        if (xlsxFallback?.html) {
          const xlsxPdf = await htmlToPdfBuffer(xlsxFallback.html);
          await fsp.writeFile(xlsxHtmlPdfPath, xlsxPdf);
          return {
            sourcePath: resolved,
            sourceExt: ext,
            converted: true,
            convertMode: "xlsx-html",
            warningMessage: warningMessage || "고품질 변환 엔진을 찾지 못해 표 보기 모드로 전환했습니다.",
            errorCode: "",
            convertDiagnostics: {
              selectedMode: "xlsx-html",
              candidateScores: [],
              fallbackReason: warningMessage || "고품질 변환 엔진을 찾지 못해 표 보기 모드로 전환했습니다."
            },
            sheetMap: Array.isArray(xlsxFallback.sheetMap) ? xlsxFallback.sheetMap : [],
            pdfPath: xlsxHtmlPdfPath,
            pdfBuffer: xlsxPdf
          };
        }
      } catch (_error) {
        // continue to text fallback
      }
    }
  }

  const fallbackPdfPath = createConvertedPdfPath(resolved, "fallback");
  const fallbackCache = readCachedPdf(fallbackPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
  if (fallbackCache) {
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: true,
      convertMode: "fallback",
      warningMessage,
      errorCode,
      convertDiagnostics: {
        selectedMode: "fallback",
        candidateScores: [],
        fallbackReason: warningMessage || ""
      },
      sheetMap: [],
      pdfPath: fallbackCache,
      pdfBuffer: await fsp.readFile(fallbackCache)
    };
  }

  const extractedText = await extractDocumentTextForFallback(resolved, ext);
  const fallbackText = extractedText || (errorCode === "EMPTY_DOCUMENT" ? "문서에 표시할 셀 내용이 없습니다." : "");
  const html = buildFallbackHtml({
    sourcePath: resolved,
    sourceExt: ext,
    contentText: fallbackText,
    warningMessage
  });
  const fallbackPdf = await htmlToPdfBuffer(html);
  await fsp.writeFile(fallbackPdfPath, fallbackPdf);

  return {
    sourcePath: resolved,
    sourceExt: ext,
    converted: true,
    convertMode: "fallback",
    warningMessage,
    errorCode,
    convertDiagnostics: {
      selectedMode: "fallback",
      candidateScores: [],
      fallbackReason: warningMessage || ""
    },
    sheetMap: [],
    pdfPath: fallbackPdfPath,
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
        },
        { type: "separator" },
        {
          label: mt("versionInfo"),
          click: () => sendToRenderer("menu-action", "show-version-info")
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
      errorCode: converted.errorCode || "",
      convertDiagnostics: converted.convertDiagnostics || {
        selectedMode: converted.convertMode || "fallback",
        candidateScores: [],
        fallbackReason: converted.warningMessage || ""
      },
      sheetMap: Array.isArray(converted.sheetMap) ? converted.sheetMap : [],
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
      errorCode: converted.errorCode || "",
      convertDiagnostics: converted.convertDiagnostics || {
        selectedMode: converted.convertMode || "fallback",
        candidateScores: [],
        fallbackReason: converted.warningMessage || ""
      },
      sheetMap: Array.isArray(converted.sheetMap) ? converted.sheetMap : [],
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








