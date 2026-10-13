const { app, BrowserWindow, Menu, dialog, ipcMain, clipboard, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const crypto = require("node:crypto");
const https = require("node:https");
const { spawnSync, spawn } = require("node:child_process");
const APP_ROOT = path.resolve(__dirname, "../..");
const { PDFDocument } = require("pdf-lib");
const {
  SPREADSHEET_RENDER_VERSION,
  readSpreadsheetWorkbook,
  inspectSpreadsheetWorkbook,
  buildSpreadsheetCompatibility,
  renderSpreadsheetCompatibilityPdf
} = require("./spreadsheet/compatibility");
const {
  buildExcelReadOnlyScript,
  buildOwnedExcelCleanupScript,
  convertSpreadsheetWithLibreOffice,
  convertOfficeWithLibreOffice
} = require("./spreadsheet/native-engine");
const { normalizeWebSourceUrl, getWebPdfName, getPublicProductDetailUrl } = require("./web/web-source");
const { inspectPdfCaptureContent } = require("./web/pdf-content-check");
const packageMeta = require("../../package.json");
const { createFeedbackLogger } = require("./hwp/feedback-log");
const { buildHwpBlocksFromHwpJson } = require("./hwp/hwp-json-compat");
const { buildHwpPageModelCompatibilityHtmlFromJson } = require("./hwp/page-html-renderer");
const { inspectHwpFamilyBytes, readHwpDocumentVariants } = require("./hwp/document-readiness");
const { renderHwpWithRhwp } = require("./hwp/rhwp-compat");
const { buildHwpxSectionLayout, buildHwpxPictureHtml, normalizeHwpxDocumentStyles } = require("./hwp/hwpx-layout");
const {
  WEB_PRINT_ENGINE_VERSION,
  inspectWebPage,
  waitForWebPageReadiness,
  waitForWebAccess,
  prepareWebPageForPdf,
  useScreenMediaForPdf,
  assessWebPageContent,
  inspectCapturedImage
} = require("./web/page-preparation");
const { normalizeSpecialSymbols } = require("./hwp/symbol-normalizer");
const {
  hwpUnitToPx,
  inlineStyleToCss,
  paragraphStyleToCss,
  createStructuredCharStyleResolver,
  createStructuredParaStyleResolver
} = require("./hwp/style-profile");
const {
  createBorderFillResolver,
  buildTableCellStyle,
  shouldTreatFirstRowAsHeader
} = require("./hwp/table-style");
const {
  resolveForcedPageTarget,
  calibratePdfByPageTarget
} = require("./hwp/page-density-calibrator");

let mainWindow = null;
let pendingDocumentToOpen = null;
let updateConfig = null;
let updateDownloaded = false;
let updateTargetVersion = "";
let updateInstallTriggered = false;
let updateBusy = false;
let appSettings = {
  language: "ko",
  allowPrintStyleHwpFallback: false,
  hwpPrintFallbackExplicitlySet: false,
  hwpNativeRetryPolicy: "quality-first-single-native-try",
  printFallbackUserApprovedAtV132: false
};
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
let hancomComAvailabilityCache = {
  checkedAt: 0,
  available: false
};
let hwpNativeSessionState = {
  skipUntil: 0,
  reasonCode: "",
  reasonMessage: ""
};
let hwpSecurityModuleStateCache = {
  checkedAt: 0,
  lookupModulePath: "",
  lookupModuleTrusted: false,
  installedModuleCandidates: [],
  lookupHash: "",
  exampleHash: ""
};

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
    hwpPrintFallback: "HWP 출력 방식(PDF 변환용) 보조 시도 허용",
    updateCheck: "업데이트 확인",
    copySupportLink: "문의 링크 복사",
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
    hwpPrintFallback: "Allow HWP print-style fallback (PDF conversion)",
    updateCheck: "Check for Updates",
    copySupportLink: "Copy Support Link",
    versionInfo: "Version Info"
  }
};

function mt(key) {
  const lang = appSettings.language === "en" ? "en" : "ko";
  return menuText[lang][key] || menuText.ko[key] || key;
}

const SPREADSHEET_EXTENSIONS = new Set([".xls", ".xlsx", ".xlsm"]);
const ACCESS_EXTENSIONS = new Set([".accdb"]);
const WEB_DOCUMENT_EXTENSIONS = new Set([".html", ".htm"]);
const SUPPORTED_DOCUMENT_EXTENSIONS = new Set([
  ".pdf",
  ".hwp",
  ".hwpx",
  ".doc",
  ".docx",
  ".xls",
  ".xlsx",
  ".xlsm",
  ".ppt",
  ".pptx",
  ".accdb",
  ".html",
  ".htm"
]);
const OFFICE_EXTENSIONS = new Set([".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx"]);
const HWP_EXTENSIONS = new Set([".hwp", ".hwpx"]);
const CONVERT_CACHE_SCHEMA_VERSION = "v10";
const HWP_QUALITY_SELECTOR_VERSION = "v10";
const HWP_SELECTION_META_VERSION = "v10";
const TRUST_CUSTOM_HWP_SECURITY_MODULE = process.env.LOOKUP_TRUST_CUSTOM_HWP_SECURITY_MODULE === "1";
const FALLBACK_CACHE_TTL_MS = 45 * 60 * 1000;
const HWPJS_SELECTOR_META_TTL_MS = 20 * 60 * 1000;
const COMPATIBILITY_SELECTION_META_TTL_MS = 20 * 60 * 1000;
const MIN_VALID_PDF_BYTES = 1024;
const HWP_NATIVE_SKIP_TTL_MS = 3 * 60 * 1000;
const HWP_NATIVE_TIMEOUT_SKIP_TTL_MS = 90 * 1000;
const HWP_NATIVE_ENGINE_MISSING_SKIP_TTL_MS = 8 * 60 * 60 * 1000;
const ACCDB_PREVIEW_MAX_TABLES = 40;
const ACCDB_PREVIEW_MAX_ROWS = 200;
const ACCDB_PREVIEW_MAX_COLS = 40;
const ACCDB_DUMP_TIMEOUT_MS = 2 * 60 * 1000;
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

function normalizeConvertStageCode(stageText) {
  const raw = String(stageText || "").trim().toLowerCase();
  if (!raw) {
    return "convert";
  }
  if (raw.includes("prepare-open")) {
    return "prepare-open";
  }
  if (raw.includes("prepare-convert")) {
    return "prepare-convert";
  }
  if (raw.includes("engine")) {
    return "engine";
  }
  if (raw.includes("open") || raw.includes("열기")) {
    return "open";
  }
  if (raw.includes("validate") || raw.includes("검사")) {
    return "validate";
  }
  if (raw.includes("display") || raw.includes("표시")) {
    return "display";
  }
  if (raw.includes("compat")) {
    return "compatibility";
  }
  if (raw.includes("save") || raw.includes("export") || raw.includes("변환")) {
    return "convert";
  }
  if (raw.includes("엔진")) {
    return "engine";
  }
  return "convert";
}

function resolveDefaultConvertMessageKey(stageCode) {
  switch (String(stageCode || "").trim().toLowerCase()) {
    case "prepare-open":
      return "convert.step.prepare-open";
    case "prepare-convert":
      return "convert.step.prepare-convert";
    case "engine":
      return "convert.step.engine";
    case "open":
      return "convert.step.open";
    case "convert":
      return "convert.step.save";
    case "validate":
      return "convert.step.validate";
    case "display":
      return "convert.step.display";
    case "compatibility":
      return "convert.step.compatibility";
    default:
      return "";
  }
}

const feedbackLogger = createFeedbackLogger({ appRoot: APP_ROOT });
let feedbackWarmupPromise = null;
let feedbackGateFailed = false;

const convertProgressState = {
  activeOperationId: "",
  byOperation: new Map()
};

function createConvertOperationId(prefix = "open") {
  const safePrefix = String(prefix || "open").trim() || "open";
  return `${safePrefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function beginConvertOperation(operationId) {
  const safeOperationId = String(operationId || "").trim();
  if (!safeOperationId) {
    return;
  }
  convertProgressState.activeOperationId = safeOperationId;
  convertProgressState.byOperation.set(safeOperationId, 0);
}

function endConvertOperation(operationId) {
  const safeOperationId = String(operationId || "").trim();
  if (!safeOperationId) {
    return;
  }
  convertProgressState.byOperation.delete(safeOperationId);
  if (convertProgressState.activeOperationId === safeOperationId) {
    convertProgressState.activeOperationId = "";
  }
}

async function ensureFeedbackGate() {
  if (feedbackGateFailed) {
    return;
  }
  if (!feedbackWarmupPromise) {
    feedbackWarmupPromise = feedbackLogger.warmup().catch((error) => {
      void error;
      feedbackGateFailed = true;
    });
  }
  await feedbackWarmupPromise.catch(() => {});
}

function normalizeConvertPercent(stageCode, percent, operationId = "") {
  if (!Number.isFinite(percent)) {
    return null;
  }
  const normalizedStage = String(stageCode || "").trim().toLowerCase();
  const safeOperationId = String(operationId || "").trim();
  const bucketKey = safeOperationId || "__global__";
  const previous = Number(convertProgressState.byOperation.get(bucketKey) || 0);
  if (normalizedStage === "prepare-open" || normalizedStage === "prepare-convert") {
    convertProgressState.byOperation.set(bucketKey, 0);
  }
  const incoming = clampNumber(percent, 0, 100);
  const monotonic = Math.max(incoming, previous);
  convertProgressState.byOperation.set(bucketKey, monotonic);
  return monotonic;
}

function sendConvertStatus(stage, percent, message, extra = {}) {
  const extraPayload = extra && typeof extra === "object" ? extra : {};
  const operationId = String(extraPayload.operationId || convertProgressState.activeOperationId || "").trim();
  const stageCode = normalizeConvertStageCode(extraPayload.stageCode || stage);
  const safeMessage = typeof message === "string" ? message : "";
  const messageKeyRaw = typeof extraPayload.messageKey === "string" ? extraPayload.messageKey.trim() : "";
  const messageKey = messageKeyRaw || resolveDefaultConvertMessageKey(stageCode);
  const normalizedPercent = normalizeConvertPercent(stageCode, percent, operationId);
  sendToRenderer("document-convert-status", {
    ...extraPayload,
    operationId,
    stage: stageCode,
    stageLabel: String(stage || ""),
    stageCode,
    percent: normalizedPercent,
    message: safeMessage,
    messageKey
  });
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
    const packageJsonPath = require.resolve("@ohah/hwpjs/package.json");
    const nodeEntryPath = path.join(path.dirname(packageJsonPath), "dist", "index.js");
    hwpJsMod = require(nodeEntryPath);
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
          mod.GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(APP_ROOT, "node_modules", "pdfjs-dist", "legacy", "build", "pdf.worker.mjs")).href;
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
    appSettings = {
      language: "ko",
      allowPrintStyleHwpFallback: false,
      hwpPrintFallbackExplicitlySet: false,
      hwpNativeRetryPolicy: "quality-first-single-native-try",
      printFallbackUserApprovedAtV132: false
    };
    return;
  }
  const loadedAllowPrintFallback = loaded.allowPrintStyleHwpFallback === true;
  const loadedPrintFallbackExplicit = loaded.hwpPrintFallbackExplicitlySet === true;
  const loadedApprovedPrintFallback = loaded.printFallbackUserApprovedAtV132 === true;
  // HWP 설치 프로그램 경유를 원천 차단하기 위해 관련 스위치는 항상 false로 정규화한다.
  const safeAllowPrintFallback = false;
  appSettings = {
    language: loaded.language === "en" ? "en" : "ko",
    allowPrintStyleHwpFallback: safeAllowPrintFallback,
    hwpPrintFallbackExplicitlySet: safeAllowPrintFallback,
    hwpNativeRetryPolicy: "quality-first-single-native-try",
    printFallbackUserApprovedAtV132: safeAllowPrintFallback
  };
  if (loadedAllowPrintFallback || loadedPrintFallbackExplicit || loadedApprovedPrintFallback) {
    saveSettings();
  }
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
  const devIconPath = path.join(APP_ROOT, "build", "icon.ico");
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
    case "libreoffice-bundled":
      return "libreoffice-bundled-v1";
    case "libreoffice-system":
      return "libreoffice-system-v1";
    case "libreoffice":
      return "libreoffice-system-v1";
    case "hwp-print-pdf":
      return `hancomcom-print-v3-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-saveas":
      return `hancomcom-saveas-v3-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-export":
      return `hancomcom-export-v1-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-converter-exe":
      return `hancom-converter-exe-v1-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwp-layout":
      return "hancomcom-legacy-v1";
    case "hwpjs-html":
      return `hwpjs-${hwpJsVer}-hwpx-${hwpxVer}-${HWP_QUALITY_SELECTOR_VERSION}`;
    case "hwpx-reader":
      return `hwpx-reader-${hwpxVer}`;
    case "xlsx-html":
      return `xlsx-html-${xlsxVer}-${SPREADSHEET_RENDER_VERSION}`;
    case "accdb-ucanaccess":
      return "accdb-ucanaccess-v1";
    case "notion-web":
      return `notion-web-${WEB_PRINT_ENGINE_VERSION}`;
    case "web-html":
      return `web-html-${WEB_PRINT_ENGINE_VERSION}`;
    case "fallback":
      return `fallback-xlsx-${xlsxVer}`;
    default:
      return "generic-v1";
  }
}

function uniquePaths(paths = []) {
  const seen = new Set();
  const result = [];
  for (const entry of paths) {
    const normalized = String(entry || "").trim();
    if (!normalized) {
      continue;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(normalized);
  }
  return result;
}

function getSafeSha256(filePath) {
  try {
    if (!filePath || !fs.existsSync(filePath)) {
      return "";
    }
    const buf = fs.readFileSync(filePath);
    return crypto.createHash("sha256").update(buf).digest("hex");
  } catch (_error) {
    return "";
  }
}

function resolveHwpSecurityModuleState(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && now - Number(hwpSecurityModuleStateCache.checkedAt || 0) < 60 * 1000) {
    return hwpSecurityModuleStateCache;
  }
  const lookupCandidates = uniquePaths([
    path.join(process.resourcesPath || "", "security", "lookupFilePathChecker.dll"),
    path.join(process.resourcesPath || "", "security", "lookup-filepath-checker.dll"),
    path.join(APP_ROOT, "build", "security", "lookupFilePathChecker.dll"),
    path.join(APP_ROOT, "build", "security", "lookup-filepath-checker.dll")
  ]);
  const exampleCandidates = uniquePaths([
    path.join(process.resourcesPath || "", "security", "FilePathCheckerModuleExample.dll"),
    path.join(APP_ROOT, "build", "security", "FilePathCheckerModuleExample.dll"),
    path.join(APP_ROOT, "build", "security-src", "FilePathCheckerModuleExample.dll")
  ]);

  const lookupModulePath = lookupCandidates.find((candidate) => candidate && fs.existsSync(candidate)) || "";
  const exampleModulePath = exampleCandidates.find((candidate) => candidate && fs.existsSync(candidate)) || "";
  const lookupHash = getSafeSha256(lookupModulePath);
  const exampleHash = getSafeSha256(exampleModulePath);
  const lookupModuleTrusted = Boolean(
    TRUST_CUSTOM_HWP_SECURITY_MODULE
      && lookupModulePath
      && lookupHash
      && (!exampleHash || lookupHash !== exampleHash)
  );

  const roots = [
    process.env.ProgramFiles || "C:\\Program Files",
    process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)"
  ];
  const relativeCandidates = [
    "HNC\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll",
    "HNC\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll",
    "HNC\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll",
    "Hancom\\Office 2024\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll",
    "Hancom\\Office 2022\\HOffice120\\Bin\\FilePathCheckerModuleExample.dll",
    "Hancom\\Office 2020\\HOffice110\\Bin\\FilePathCheckerModuleExample.dll",
    "Hancom\\HOffice\\Bin\\FilePathCheckerModuleExample.dll",
    "Hnc\\Office\\Bin\\FilePathCheckerModuleExample.dll",
    "Hancom\\HOffice\\Bin\\FilePathCheckerModule.dll",
    "HNC\\HOffice\\Bin\\FilePathCheckerModule.dll"
  ];
  const installedCandidates = [];
  for (const root of roots) {
    for (const relative of relativeCandidates) {
      const candidate = path.join(root, relative);
      if (fs.existsSync(candidate)) {
        installedCandidates.push(candidate);
      }
    }
  }
  try {
    const registryProbe = runPowerShell(
      [
        "$roots = @(",
        "  'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules',",
        "  'HKLM:\\SOFTWARE\\HNC\\HwpAutomation\\Modules',",
        "  'HKLM:\\SOFTWARE\\WOW6432Node\\HNC\\HwpAutomation\\Modules'",
        ")",
        "$skip = @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')",
        "$found = @()",
        "foreach ($root in $roots) {",
        "  if (-not (Test-Path $root)) { continue }",
        "  try {",
        "    $props = Get-ItemProperty -Path $root -ErrorAction Stop",
        "    foreach ($prop in $props.PSObject.Properties) {",
        "      if ($skip -contains $prop.Name) { continue }",
        "      $value = [string]$prop.Value",
        "      if ($value -and (Test-Path $value)) { $found += $value }",
        "    }",
        "  } catch {}",
        "}",
        "$found | Select-Object -Unique | ForEach-Object { Write-Output $_ }"
      ].join("\n"),
      12000
    );
    if (registryProbe.ok && registryProbe.stdout) {
      for (const line of String(registryProbe.stdout).split(/\r?\n/)) {
        const candidate = String(line || "").trim();
        if (candidate && fs.existsSync(candidate)) {
          installedCandidates.push(candidate);
        }
      }
    }
  } catch (_error) {
    // registry probe failures are non-fatal
  }
  hwpSecurityModuleStateCache = {
    checkedAt: now,
    lookupModulePath,
    lookupModuleTrusted,
    installedModuleCandidates: uniquePaths(installedCandidates),
    lookupHash,
    exampleHash
  };
  return hwpSecurityModuleStateCache;
}

function getHancomSecurityModuleCandidates() {
  const resolved = resolveHwpSecurityModuleState();
  const candidates = [];
  if (resolved.lookupModuleTrusted && resolved.lookupModulePath) {
    candidates.push(resolved.lookupModulePath);
  }
  candidates.push(...(resolved.installedModuleCandidates || []));
  return uniquePaths(candidates);
}

function getHancomSecurityRegistrationPlan() {
  const resolved = resolveHwpSecurityModuleState();
  return {
    lookupModuleName: "lookupFilePathChecker",
    lookupModulePath: resolved.lookupModuleTrusted ? resolved.lookupModulePath : "",
    lookupModuleTrusted: resolved.lookupModuleTrusted === true,
    installedModuleCandidates: Array.isArray(resolved.installedModuleCandidates) ? resolved.installedModuleCandidates : []
  };
}

function buildPowerShellStringArray(values = []) {
  if (!Array.isArray(values) || values.length === 0) {
    return "@()";
  }
  return `@(${values.map((entry) => escapeForPowerShell(entry)).join(", ")})`;
}

function getHwpSecurityAllowList(sourcePath, outputPdfPath) {
  const paths = [];
  if (sourcePath) {
    paths.push(path.resolve(sourcePath));
  }
  if (outputPdfPath) {
    paths.push(path.resolve(outputPdfPath));
  }
  try {
    paths.push(ensureConvertedDir());
  } catch (_error) {
    // ignore converted dir lookup failures
  }
  return uniquePaths(paths);
}

function markHwpNativeSkip(reasonCode, reasonMessage = "") {
  const normalizedCode = String(reasonCode || "").toUpperCase();
  if (!normalizedCode) {
    return;
  }
  const ttlMs = normalizedCode === "ENGINE_TIMEOUT"
    ? HWP_NATIVE_TIMEOUT_SKIP_TTL_MS
    : normalizedCode === "ENGINE_MISSING"
      ? HWP_NATIVE_ENGINE_MISSING_SKIP_TTL_MS
      : HWP_NATIVE_SKIP_TTL_MS;
  hwpNativeSessionState = {
    skipUntil: Date.now() + ttlMs,
    reasonCode: normalizedCode,
    reasonMessage: String(reasonMessage || "").trim()
  };
}

function clearHwpNativeSkip() {
  hwpNativeSessionState = {
    skipUntil: 0,
    reasonCode: "",
    reasonMessage: ""
  };
}

function getHwpNativeSkipState() {
  if (Date.now() <= Number(hwpNativeSessionState.skipUntil || 0)) {
    return {
      skip: true,
      reasonCode: String(hwpNativeSessionState.reasonCode || ""),
      reasonMessage: String(hwpNativeSessionState.reasonMessage || "")
    };
  }
  return {
    skip: false,
    reasonCode: "",
    reasonMessage: ""
  };
}

function parseMarkerValue(stdout, markerName) {
  const match = String(stdout || "").match(new RegExp(`${markerName}=([^\\r\\n]+)`, "i"));
  return match ? String(match[1] || "").trim() : "";
}

function enrichHancomRunResult(runResult = {}, mode = "") {
  const stdout = String(runResult?.stdout || "");
  const securityRegistered = parseMarkerValue(stdout, "lookup_security_registered") === "1";
  const moduleName = parseMarkerValue(stdout, "lookup_module_name");
  const modulePath = parseMarkerValue(stdout, "lookup_module_path");
  const stepState = parseMarkerValue(stdout, "lookup_engine_state");
  const outputPath = parseMarkerValue(stdout, "lookup_output_path");
  return {
    ...runResult,
    mode,
    securityModuleRegistered: securityRegistered,
    securityModuleName: moduleName,
    securityModulePath: modulePath,
    engineState: stepState || (runResult?.ok ? "success" : "failed"),
    outputPath
  };
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

function findCandidateScoreRow(candidateScores, mode) {
  if (!Array.isArray(candidateScores)) {
    return null;
  }
  return candidateScores.find((entry) => String(entry?.mode || "") === String(mode || "")) || null;
}

function hasInvalidSelectorScore(candidateScores, selectedMode) {
  const selectedScore = findCandidateScoreRow(candidateScores, selectedMode);
  if (!selectedScore) {
    return true;
  }
  const inspected = Boolean(selectedScore.inspected);
  const finalScore = Number(selectedScore.finalScore ?? selectedScore.score ?? 0);
  const qualityScore = Number(selectedScore.qualityScore ?? finalScore);
  const mode = String(selectedMode || "").toLowerCase();
  if (!inspected) {
    return true;
  }
  if (!Number.isFinite(finalScore) || finalScore <= 0 || !Number.isFinite(qualityScore)) {
    return true;
  }
  if (mode === "fallback" || mode === "fallback-text") {
    return true;
  }
  if (mode === "hwpjs-html" && qualityScore < 60) {
    return true;
  }
  if (mode === "hwpjs-html") {
    const noisyTable = Number(selectedScore.markdownNoisyTableLineCount || 0);
    const tableCount = Number(selectedScore.markdownTableLineCount || 0);
    const symbolNoiseRatio = Number(selectedScore.markdownSymbolNoiseRatio || 0);
    const noisyRatio = tableCount > 0 ? noisyTable / tableCount : 0;
    if ((tableCount >= 6 && noisyRatio >= 0.35) || symbolNoiseRatio >= 0.3) {
      return true;
    }
  }
  return qualityScore < 24;
}

function readHwpSelectionMeta(sourcePath, options = {}) {
  try {
    const metaPath = createHwpSelectionMetaPath(sourcePath);
    if (!fs.existsSync(metaPath)) {
      return null;
    }
    const parsed = JSON.parse(fs.readFileSync(metaPath, "utf8"));
    const sourceSignature = String(options.sourceSignature || buildHwpSelectionSourceSignature(sourcePath));
    const forcedPageTarget = Math.max(0, Math.floor(Number(options.forcedPageTarget || 0)));
    if (String(parsed?.selectorVersion || "") !== HWP_SELECTION_META_VERSION) {
      return null;
    }
    if (!String(parsed?.selectedMode || "").trim()) {
      return null;
    }
    if (isBlockedHwpMode(parsed?.selectedMode || "")) {
      return null;
    }
    if (!parsed?.qualityTier) {
      return null;
    }
    if (!Array.isArray(parsed?.candidateScores) || parsed.candidateScores.length === 0) {
      return null;
    }
    if (!Array.isArray(parsed?.candidateFailures)) {
      return null;
    }
    if (!parsed?.engineState || typeof parsed.engineState !== "object") {
      return null;
    }
    if (parsed?.validationPassed !== true) {
      return null;
    }
    if (String(parsed?.sourceSignature || "") !== sourceSignature) {
      return null;
    }
    if (forcedPageTarget > 0 && Number(parsed?.pageCount || 0) !== forcedPageTarget) {
      return null;
    }
    const selectedPdfPath = String(parsed?.selectedPdfPath || "");
    if (!selectedPdfPath || !fs.existsSync(selectedPdfPath)) {
      return null;
    }
    return {
      selectedMode: String(parsed?.selectedMode || ""),
      qualityTier: String(parsed?.qualityTier || ""),
      selectedPdfPath,
      warningMessage: String(parsed?.warningMessage || ""),
      fallbackReason: String(parsed?.fallbackReason || ""),
      analysisMode: String(parsed?.analysisMode || "quick"),
      candidateScores: Array.isArray(parsed?.candidateScores) ? parsed.candidateScores : [],
      candidateFailures: Array.isArray(parsed?.candidateFailures) ? parsed.candidateFailures : [],
      pageCount: Number(parsed?.pageCount || 0),
      targetPageCount: Number(parsed?.targetPageCount || 0),
      artifactScore: Number(parsed?.artifactScore || 0),
      hwpPageModel: parsed?.hwpPageModel && typeof parsed.hwpPageModel === "object" ? parsed.hwpPageModel : null,
      engineVersion: String(parsed?.engineVersion || ""),
      engineState: parsed?.engineState && typeof parsed.engineState === "object" ? parsed.engineState : {},
      securityModuleRegistered: parsed?.securityModuleRegistered === true,
      validationPassed: parsed?.validationPassed === true,
      sourceSignature,
      updatedAt: String(parsed?.updatedAt || "")
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
      sourceSignature: String(payload.sourceSignature || buildHwpSelectionSourceSignature(sourcePath)),
      selectedMode: String(payload.selectedMode || ""),
      qualityTier: String(payload.qualityTier || resolveQualityTier(payload.selectedMode || "")),
      selectedPdfPath: String(payload.selectedPdfPath || ""),
      warningMessage: String(payload.warningMessage || ""),
      fallbackReason: String(payload.fallbackReason || ""),
      analysisMode: String(payload.analysisMode || "quick"),
      candidateScores: Array.isArray(payload.candidateScores) ? payload.candidateScores : [],
      candidateFailures: Array.isArray(payload.candidateFailures) ? payload.candidateFailures : [],
      pageCount: Number(payload.pageCount || 0),
      targetPageCount: Number(payload.targetPageCount || 0),
      artifactScore: Number(payload.artifactScore || 0),
      hwpPageModel: payload.hwpPageModel && typeof payload.hwpPageModel === "object" ? payload.hwpPageModel : null,
      engineVersion: String(payload.engineVersion || ""),
      engineState: payload.engineState && typeof payload.engineState === "object" ? payload.engineState : {},
      securityModuleRegistered: payload.securityModuleRegistered === true,
      validationPassed: payload.validationPassed !== false,
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
  try {
    const stat = fs.statSync(cachePath);
    if (!stat.isFile() || stat.size < (options.minSizeBytes || MIN_VALID_PDF_BYTES)) {
      return null;
    }
    const fd = fs.openSync(cachePath, "r");
    const header = Buffer.alloc(5);
    fs.readSync(fd, header, 0, 5, 0);
    fs.closeSync(fd);
    if (header.toString("utf8") !== "%PDF-") {
      return null;
    }
  } catch (_error) {
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

async function validateConvertedPdfPath(pdfPath, options = {}) {
  if (!pdfPath || !fs.existsSync(pdfPath)) {
    return {
      ok: false,
      errorCode: "INVALID_CONVERTED_PDF",
      reason: "file_not_found",
      pageCount: 0
    };
  }
  try {
    const stat = await fsp.stat(pdfPath);
    if (!stat.isFile() || stat.size < (options.minSizeBytes || MIN_VALID_PDF_BYTES)) {
      return {
        ok: false,
        errorCode: "INVALID_CONVERTED_PDF",
        reason: "file_too_small",
        pageCount: 0
      };
    }
    const fd = await fsp.open(pdfPath, "r");
    const header = Buffer.alloc(5);
    await fd.read(header, 0, 5, 0);
    await fd.close();
    if (header.toString("utf8") !== "%PDF-") {
      return {
        ok: false,
        errorCode: "INVALID_CONVERTED_PDF",
        reason: "pdf_header_missing",
        pageCount: 0
      };
    }
    const pageCount = await getPdfPageCount(pdfPath);
    if (!Number.isFinite(pageCount) || pageCount < 1) {
      return {
        ok: false,
        errorCode: "PAGE_COUNT_ZERO",
        reason: "page_count_zero",
        pageCount: 0
      };
    }
    return {
      ok: true,
      errorCode: "",
      reason: "",
      pageCount
    };
  } catch (_error) {
    return {
      ok: false,
      errorCode: "INVALID_CONVERTED_PDF",
      reason: "pdf_parse_failed",
      pageCount: 0
    };
  }
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
    const inspection = inspectSpreadsheetWorkbook(readSpreadsheetWorkbook(filePath));
    if (inspection.hasContent) return { ...inspection, reason: "cell-data" };
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
  if (["INVALID_HWP", "HWP_CONTENT_UNREADABLE", "SPREADSHEET_LIMIT", "EMPTY_DOCUMENT"].includes(code)) return code;
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
  if (/invalid pdf structure|pdf_header_missing|page_count_zero|invalid_converted_pdf|손상된 pdf/i.test(message)) {
    return "CONVERT_FAILED";
  }
  if (/변환|convert|hwp|office/i.test(message)) {
    return "CONVERT_FAILED";
  }
  return "OPEN_FAILED";
}

function parseDocumentRequestPayload(payload) {
  if (typeof payload === "string") {
    return {
      filePath: payload,
      options: {}
    };
  }
  if (payload && typeof payload === "object") {
    return {
      filePath: String(payload.filePath || ""),
      options: payload.options && typeof payload.options === "object" ? payload.options : {}
    };
  }
  return {
    filePath: "",
    options: {}
  };
}

function buildFeedbackKey(prefix, filePath, detail = "") {
  const seed = `${String(prefix || "feedback")}|${String(filePath || "")}|${String(detail || "")}`;
  return crypto.createHash("sha1").update(seed).digest("hex").slice(0, 16);
}

async function recordFeedbackProblem(entry) {
  try {
    await feedbackLogger.recordProblem(entry);
  } catch (_error) {
    // feedback log failures should not block document flow
  }
}

async function recordFeedbackSolution(entry) {
  if (entry?.userConfirmed !== true) return;
  try {
    await feedbackLogger.recordSolution(entry);
  } catch (_error) {
    // feedback log failures should not block document flow
  }
}

function fileUriToWindowsPath(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  let candidate = raw.split(/\r?\n/).map((entry) => String(entry || "").trim()).find(Boolean) || "";
  if (!candidate) {
    return "";
  }
  if (/^[a-zA-Z]:\\/.test(candidate) || /^\\\\/.test(candidate)) {
    return candidate;
  }
  if (!candidate.toLowerCase().startsWith("file://")) {
    return "";
  }
  candidate = candidate.replace(/^file:\/\/+/i, "");
  try {
    candidate = decodeURIComponent(candidate);
  } catch (_error) {
    // keep original candidate when decode fails
  }
  candidate = candidate.replace(/\//g, "\\");
  if (/^\\[a-zA-Z]:\\/.test(candidate)) {
    return candidate.slice(1);
  }
  return candidate;
}

function normalizeDropPathCandidate(value) {
  const raw = String(value || "").trim().replace(/^"+|"+$/g, "");
  if (!raw) {
    return "";
  }
  if (/^[a-zA-Z]:\\/.test(raw) || /^\\\\/.test(raw)) {
    return path.normalize(raw);
  }
  if (/^[a-zA-Z]:\//.test(raw)) {
    return path.normalize(raw.replace(/\//g, "\\"));
  }
  if (raw.toLowerCase().startsWith("file://")) {
    return path.normalize(fileUriToWindowsPath(raw));
  }
  return "";
}

function resolveDropPathPayload(payload = {}) {
  const input = payload && typeof payload === "object" ? payload : {};
  const candidates = [];
  const push = (value) => {
    const normalized = normalizeDropPathCandidate(value);
    if (!normalized) {
      return;
    }
    const key = normalized.toLowerCase();
    if (!candidates.some((entry) => entry.toLowerCase() === key)) {
      candidates.push(normalized);
    }
  };

  push(input.filePath);
  push(input.directPath);
  if (Array.isArray(input.candidates)) {
    input.candidates.forEach(push);
  }
  push(input.uriList);
  push(input.plainText);
  const downloadUrl = String(input.downloadUrl || "").trim();
  if (downloadUrl) {
    const parts = downloadUrl.split(":");
    if (parts.length >= 3) {
      push(parts.slice(2).join(":"));
    }
  }

  for (const candidate of candidates) {
    try {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
        return candidate;
      }
    } catch (_error) {
      // continue
    }
  }
  return candidates[0] || "";
}

function buildDiagnosticDetail(error) {
  const message = String(error?.message || "").trim();
  const code = String(error?.code || "").trim();
  const parts = [];
  if (code) {
    parts.push(`code=${code}`);
  }
  if (message) {
    parts.push(message);
  }
  return parts.join(" | ").slice(0, 1200);
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

function parseWorkerStatusLine(lineText) {
  const normalized = String(lineText || "").trim();
  if (!normalized || !normalized.startsWith("LOOKUP_STATUS")) {
    return null;
  }
  const rawJson = normalized.slice("LOOKUP_STATUS".length).trim();
  if (!rawJson) {
    return null;
  }
  try {
    const parsed = JSON.parse(rawJson);
    return {
      stage: String(parsed.stage || "").trim(),
      percent: Number.isFinite(Number(parsed.percent)) ? Number(parsed.percent) : null,
      message: String(parsed.message || "").trim(),
      messageKey: String(parsed.messageKey || "").trim(),
      mode: String(parsed.mode || "").trim()
    };
  } catch (_error) {
    return null;
  }
}

async function runPowerShellWorker(scriptContent, timeoutMs = 180000, options = {}) {
  const workerDir = path.join(app.getPath("temp"), "lookup-hwp-workers");
  await fsp.mkdir(workerDir, { recursive: true });
  const scriptPath = path.join(workerDir, `worker-${Date.now()}-${Math.random().toString(16).slice(2)}.ps1`);
  await fsp.writeFile(scriptPath, `\uFEFF${String(scriptContent || "")}`, "utf8");
  let stdoutBuffer = "";
  const onStatus = typeof options?.onStatus === "function" ? options.onStatus : null;
  const consumeStatusLines = (textChunk, flush = false) => {
    if (!onStatus) {
      return;
    }
    stdoutBuffer += String(textChunk || "");
    let newlineIndex = stdoutBuffer.search(/\r?\n/);
    while (newlineIndex >= 0) {
      const line = stdoutBuffer.slice(0, newlineIndex);
      stdoutBuffer = stdoutBuffer.slice(newlineIndex + (stdoutBuffer[newlineIndex] === "\r" ? 2 : 1));
      const statusPayload = parseWorkerStatusLine(line);
      if (statusPayload) {
        onStatus(statusPayload);
      }
      newlineIndex = stdoutBuffer.search(/\r?\n/);
    }
    if (flush && stdoutBuffer.trim().length > 0) {
      const statusPayload = parseWorkerStatusLine(stdoutBuffer);
      if (statusPayload) {
        onStatus(statusPayload);
      }
      stdoutBuffer = "";
    }
  };
  try {
    const result = await runProcess(
      "powershell",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
      timeoutMs,
      {
        onStdout: (chunkText) => consumeStatusLines(chunkText, false)
      }
    );
    consumeStatusLines("", true);
    return {
      ok: Boolean(result.ok),
      status: result.status,
      timedOut: Boolean(result.timedOut),
      stdout: String(result.stdout || "").trim(),
      stderr: String(result.stderr || "").trim(),
      error: String(result.error || "")
    };
  } finally {
    try {
      await fsp.unlink(scriptPath);
    } catch (_error) {
      // ignore worker script cleanup failures
    }
  }
}

function runProcess(command, args, timeoutMs = 180000, options = {}) {
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
      const text = chunk.toString("utf8");
      stdout += text;
      if (typeof options?.onStdout === "function") {
        try {
          options.onStdout(text);
        } catch (_error) {
          // ignore callback errors
        }
      }
    });
    child.stderr.on("data", (chunk) => {
      const text = chunk.toString("utf8");
      stderr += text;
      if (typeof options?.onStderr === "function") {
        try {
          options.onStderr(text);
        } catch (_error) {
          // ignore callback errors
        }
      }
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

function findAccdbJavaExecutable() {
  const candidates = uniquePaths([
    path.join(process.resourcesPath || "", "accdb", "java", "bin", "java.exe"),
    path.join(process.resourcesPath || "", "accdb", "java", "bin", "javaw.exe"),
    path.join(APP_ROOT, "build", "accdb", "java", "bin", "java.exe"),
    path.join(APP_ROOT, "build", "accdb", "java", "bin", "javaw.exe")
  ]);
  for (const candidate of candidates) {
    if (candidate && fs.existsSync(candidate)) {
      return candidate;
    }
  }
  const whereJava = spawnSync("where", ["java"], {
    encoding: "utf8",
    windowsHide: true
  });
  if (whereJava.status === 0) {
    const first = String(whereJava.stdout || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find(Boolean);
    if (first && fs.existsSync(first)) {
      return first;
    }
  }
  return "";
}

function listAccdbRuntimeJars() {
  const roots = uniquePaths([
    path.join(process.resourcesPath || "", "accdb", "lib"),
    path.join(APP_ROOT, "build", "accdb", "lib")
  ]);
  for (const root of roots) {
    try {
      if (!root || !fs.existsSync(root)) {
        continue;
      }
      const jars = fs.readdirSync(root)
        .filter((name) => String(name || "").toLowerCase().endsWith(".jar"))
        .map((name) => path.join(root, name));
      if (jars.length > 0) {
        return jars;
      }
    } catch (_error) {
      // ignore runtime lookup errors
    }
  }
  return [];
}

function resolveAccdbDumpClassDir() {
  const candidates = uniquePaths([
    path.join(process.resourcesPath || "", "accdb", "classes"),
    path.join(APP_ROOT, "build", "accdb", "classes")
  ]);
  return candidates.find((candidate) => candidate && fs.existsSync(candidate)) || "";
}

async function runAccdbDump(sourcePath, options = {}) {
  const javaPath = findAccdbJavaExecutable();
  const jarEntries = listAccdbRuntimeJars();
  const classDir = resolveAccdbDumpClassDir();
  if (!javaPath || jarEntries.length === 0 || !classDir) {
    return {
      ok: false,
      errorCode: "ENGINE_MISSING",
      message: "accdb_runtime_missing",
      tables: []
    };
  }

  const maxTables = Math.max(1, Math.floor(Number(options.maxTables) || ACCDB_PREVIEW_MAX_TABLES));
  const maxRows = Math.max(1, Math.floor(Number(options.maxRows) || ACCDB_PREVIEW_MAX_ROWS));
  const maxCols = Math.max(1, Math.floor(Number(options.maxCols) || ACCDB_PREVIEW_MAX_COLS));
  const outJsonPath = path.join(
    ensureConvertedDir(),
    `accdb-dump-${Date.now()}-${Math.random().toString(16).slice(2)}.json`
  );
  const classpath = [...jarEntries, classDir].join(path.delimiter);
  const args = [
    "-cp",
    classpath,
    "LookupAccdbDump",
    sourcePath,
    outJsonPath,
    String(maxTables),
    String(maxRows),
    String(maxCols)
  ];
  const result = await runProcess(javaPath, args, ACCDB_DUMP_TIMEOUT_MS);
  if (!result.ok || !fs.existsSync(outJsonPath)) {
    return {
      ok: false,
      errorCode: result.timedOut ? "ENGINE_TIMEOUT" : "CONVERT_FAILED",
      message: String(result.stderr || result.error || "accdb_dump_failed"),
      tables: []
    };
  }
  try {
    const parsed = JSON.parse(await fsp.readFile(outJsonPath, "utf8"));
    const tables = Array.isArray(parsed?.tables) ? parsed.tables : [];
    return {
      ok: true,
      tables
    };
  } catch (_error) {
    return {
      ok: false,
      errorCode: "CONVERT_FAILED",
      message: "accdb_dump_parse_failed",
      tables: []
    };
  } finally {
    try {
      await fsp.unlink(outJsonPath);
    } catch (_error) {
      // ignore temp cleanup errors
    }
  }
}

function findLibreOfficeExecutable() {
  const bundledCandidates = [
    path.join(process.resourcesPath || "", "libreoffice", "program", "soffice.com"),
    path.join(process.resourcesPath || "", "libreoffice", "program", "soffice.exe"),
    path.join(APP_ROOT, "build", "libreoffice", "program", "soffice.com"),
    path.join(APP_ROOT, "build", "libreoffice", "program", "soffice.exe")
  ];
  for (const candidate of bundledCandidates) {
    if (candidate && fs.existsSync(candidate)) {
      return { path: candidate, mode: "libreoffice-bundled" };
    }
  }

  const knownPaths = [
    path.join(process.env["ProgramFiles"] || "C:\\Program Files", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["ProgramFiles"] || "C:\\Program Files", "LibreOffice", "program", "soffice.com"),
    path.join(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)", "LibreOffice", "program", "soffice.exe"),
    path.join(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)", "LibreOffice", "program", "soffice.com")
  ];
  for (const candidate of knownPaths) {
    if (fs.existsSync(candidate)) {
      return { path: candidate, mode: "libreoffice-system" };
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
      return { path: first, mode: "libreoffice-system" };
    }
  }

  return { path: "", mode: "" };
}

function isTimeoutResult(result) {
  if (!result) {
    return false;
  }
  return Boolean(result.timedOut) || /timed out|timeout/i.test(String(result.error || result.stderr || ""));
}

async function detectHancomComAvailability(forceCheck = false) {
  const now = Date.now();
  if (!forceCheck && now - hancomComAvailabilityCache.checkedAt < 60 * 1000) {
    return false;
  }
  // lookup 정책: HWP 설치 프로그램 경유 경로를 사용하지 않는다.
  const available = false;
  hancomComAvailabilityCache = {
    checkedAt: now,
    available
  };
  return available;
}

function isNativeHwpMode(mode) {
  return mode === "hwp-saveas" || mode === "hwp-export";
}

function isBlockedHwpMode(mode) {
  const normalized = String(mode || "").toLowerCase();
  return (
    normalized === "hwp-saveas"
    || normalized === "hwp-export"
    || normalized === "hwp-print-pdf"
    || normalized === "hwp-converter-exe"
    || normalized === "hancom-com-saveas"
    || normalized === "hancom-com-export"
    || normalized === "hancom-com-print"
  );
}

function getHwpModePriority(mode) {
  switch (String(mode || "")) {
    case "hwp-saveas":
      return 5;
    case "hwp-export":
      return 4;
    case "hwp-print-pdf":
      return 1;
    case "hwp-converter-exe":
      return 1;
    case "hwpjs-html":
      return 2;
    case "hwp-rhwp":
      return 3;
    case "hwpx-reader":
      return 3;
    case "fallback":
      return 1;
    default:
      return 0;
  }
}

function detectHwpRunFailureCode(runResult) {
  const merged = `${String(runResult?.stderr || "")} ${String(runResult?.error || "")}`.toLowerCase();
  if (!runResult) {
    return "";
  }
  if (isTimeoutResult(runResult)) {
    return "ENGINE_TIMEOUT";
  }
  if (runResult.missingEngine || /security_module_missing|class not registered|cannot find|not found/.test(merged)) {
    return "ENGINE_MISSING";
  }
  if (/artifact|booklet/.test(merged)) {
    return "RENDER_ARTIFACT_DETECTED";
  }
  return "CONVERT_FAILED";
}

function updateHwpNativeSessionStateFromRunResult(runResult) {
  if (!runResult || typeof runResult !== "object") {
    return;
  }
  if (runResult.ok === true) {
    clearHwpNativeSkip();
    return;
  }
  const code = detectHwpRunFailureCode(runResult);
  if (code === "ENGINE_MISSING") {
    const reason = "보안 모듈 또는 한컴 엔진을 확인하지 못해 이번 세션에서 원본 변환을 잠시 건너뜁니다.";
    markHwpNativeSkip(code, reason);
    return;
  }
  if (code === "ENGINE_TIMEOUT") {
    const reason = "한컴 엔진 응답이 지연되어 잠시 후 원본 변환을 다시 시도합니다.";
    markHwpNativeSkip(code, reason);
  }
}

function mapWorkerStageToConvertStage(stageText) {
  const normalized = String(stageText || "").trim().toLowerCase();
  if (!normalized) {
    return "convert";
  }
  if (normalized === "engine") {
    return "engine";
  }
  if (normalized === "open") {
    return "open";
  }
  if (normalized === "save" || normalized === "export") {
    return "convert";
  }
  if (normalized === "validate") {
    return "validate";
  }
  if (normalized === "display") {
    return "display";
  }
  return "convert";
}

function resolveWorkerMessageKey(workerStatus = {}) {
  const directKey = String(workerStatus.messageKey || "").trim();
  if (directKey) {
    return directKey;
  }
  const stage = String(workerStatus.stage || "").trim().toLowerCase();
  switch (stage) {
    case "engine":
      return "convert.step.engine";
    case "open":
      return "convert.step.open";
    case "save":
      return "convert.step.save";
    case "validate":
      return "convert.step.validate";
    case "display":
      return "convert.step.display";
    default:
      return "";
  }
}

function createHwpWorkerStatusForwarder(sourceExt, operationId = "") {
  return (workerStatus) => {
    if (!workerStatus || typeof workerStatus !== "object") {
      return;
    }
    const stage = mapWorkerStageToConvertStage(workerStatus.stage);
    const percent = Number.isFinite(workerStatus.percent) ? clampNumber(workerStatus.percent, 0, 100) : null;
    const messageKey = resolveWorkerMessageKey(workerStatus);
    sendConvertStatus(stage, percent, "", {
      sourceExt,
      selectedMode: String(workerStatus.mode || ""),
      messageKey,
      operationId
    });
  };
}

function buildConversionWarning(sourceExt, engineMessage, hint) {
  const extLabel = String(sourceExt || "").replace(".", "").toUpperCase();
  const details = String(engineMessage || "").trim();
  const reason = hint || "변환 엔진을 사용하지 못해 호환 보기로 전환했습니다.";
  if (details) {
    logConvertTrace("convert-warning-details", {
      sourceExt: extLabel,
      details: details.slice(0, 300)
    });
  }
  return `${extLabel} 문서: ${reason}`;
}

function normalizePublicConvertMode(mode, converted) {
  if (!converted) {
    return "native-pdf";
  }
  switch (String(mode || "").toLowerCase()) {
    case "hwp-saveas":
      return "hancom-com-saveas";
    case "hwp-export":
      return "hancom-com-export";
    case "hwp-print-pdf":
      return "hancom-com-print";
    case "cache":
      return "cache";
    case "hwpjs-html":
      return "hwpjs-compat";
    case "hwpx-reader":
      return "hwpx-reader";
    case "libreoffice-bundled":
      return "libreoffice-bundled";
    case "libreoffice-system":
    case "libreoffice":
      return "libreoffice-system";
    case "accdb-ucanaccess":
      return "accdb-ucanaccess";
    case "notion-web":
      return "notion-web";
    case "web-html":
      return "web-html";
    case "xlsx-html":
    case "fallback":
    case "fallback-text":
      return "fallback-text";
    case "native":
      return "native-pdf";
    default:
      return String(mode || "fallback-text");
  }
}

function resolveQualityTier(mode = "") {
  const normalized = String(mode || "").toLowerCase();
  if (normalized === "native" || normalized === "native-pdf") {
    return "native-original";
  }
  if (normalized === "hwp-saveas" || normalized === "hwp-export" || normalized === "hancom-com-saveas" || normalized === "hancom-com-export") {
    return "native-original";
  }
  if (
    normalized === "hwpjs-html"
    || normalized === "hwp-rhwp"
    || normalized === "hwpx-reader"
    || normalized === "hwpjs-compat"
    || normalized === "hwp-print-pdf"
    || normalized === "hancom-com-print"
    || normalized === "libreoffice-bundled"
    || normalized === "libreoffice-system"
    || normalized === "office-com"
    || normalized === "xlsx-html"
    || normalized === "accdb-ucanaccess"
    || normalized === "notion-web"
    || normalized === "web-html"
  ) {
    return "compatibility-layout";
  }
  if (normalized === "fallback" || normalized === "fallback-text") {
    return "compatibility-text";
  }
  if (normalized === "cache") {
    return "native-original";
  }
  return "compatibility-text";
}

function normalizeConvertDiagnostics(diagnostics, selectedMode = "fallback", fallbackReason = "") {
  const input = diagnostics && typeof diagnostics === "object" ? diagnostics : {};
  const normalizedMode = String(input.selectedMode || selectedMode || "fallback");
  return {
    selectorVersion: input.selectorVersion || HWP_QUALITY_SELECTOR_VERSION,
    selectedMode: normalizedMode,
    qualityTier: String(input.qualityTier || resolveQualityTier(normalizedMode)),
    validationPassed: input.validationPassed === true,
    pageCount: Number(input.pageCount || 0),
    candidateScores: Array.isArray(input.candidateScores) ? input.candidateScores : [],
    candidateFailures: Array.isArray(input.candidateFailures) ? input.candidateFailures : [],
    artifactScore: Number(input.artifactScore || 0),
    engineState: input.engineState && typeof input.engineState === "object" ? input.engineState : {},
    securityModuleRegistered: input.securityModuleRegistered === true,
    fallbackReason: String(input.fallbackReason || fallbackReason || ""),
    sourceType: String(input.sourceType || ""),
    webPage: input.webPage && typeof input.webPage === "object" ? input.webPage : null
  };
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
  return normalizeSpecialSymbols(String(text || ""))
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .trim();
}

function sanitizeHwpCompatibilityMarkdown(markdownRaw) {
  let normalized = normalizeSpecialSymbols(String(markdownRaw || ""));
  if (!normalized.trim()) {
    return {
      markdown: "",
      strikeRatio: 0,
      strikeSegments: 0
    };
  }

  const strikeRegex = /~~([^~]+)~~/g;
  let strikeChars = 0;
  let strikeSegments = 0;
  let strikeMatch = null;
  while ((strikeMatch = strikeRegex.exec(normalized)) !== null) {
    strikeSegments += 1;
    strikeChars += String(strikeMatch[1] || "").replace(/\s+/g, "").length;
  }
  const plainChars = normalized
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/`[^`]*`/g, " ")
    .replace(/[#>*_\-\[\]\(\)!]/g, " ")
    .replace(/\s+/g, "")
    .length;
  const strikeRatio = plainChars > 0 ? strikeChars / plainChars : 0;

  // HWP fallback에서 과도한 취소선 표기가 나오는 문서를 방어한다.
  if (strikeSegments >= 8 || strikeRatio >= 0.18) {
    normalized = normalized.replace(/~~([^~]+)~~/g, "$1");
  }

  normalized = normalized
    .replace(/\uFFFD+/g, " ")
    .replace(/^\s{0,3}#{1,6}\s*/gm, "")
    .replace(/^\s*>\s?/gm, "")
    .replace(/^\s*[-*]{3,}\s*$/gm, "")
    .replace(/^\s*\|[-:\s|]+\|\s*$/gm, "")
    .replace(/`{1,3}/g, "")
    // 볼드(**text**) 표시는 유지하고, 과도한 기호 노이즈만 제거한다.
    .replace(/[*_~]{4,}/g, "")
    .replace(/[|]{3,}/g, " | ")
    .replace(/^\s*[|｜]\s*$/gm, "")
    .replace(/^\s*(?:[-*•]\s*)?[|｜](?:\s*[|｜])+\s*$/gm, "")
    .replace(/^\s*[#]{1,6}\s*$/gm, "")
    .replace(/[^\S\r\n]{3,}/g, "  ")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ");

  return {
    markdown: normalizeText(normalized),
    strikeRatio,
    strikeSegments
  };
}

function renderCompatInline(text) {
  const source = String(text || "");
  if (!source) {
    return "";
  }
  let html = "";
  let offset = 0;
  const tokenRegex = /(\*\*[^*\n][\s\S]*?\*\*|__[^_\n][\s\S]*?__)/g;
  let match = null;
  while ((match = tokenRegex.exec(source)) !== null) {
    const start = match.index;
    const token = match[0] || "";
    if (start > offset) {
      html += escapeHtml(source.slice(offset, start));
    }
    const inner = token.slice(2, -2);
    html += `<strong>${escapeHtml(inner)}</strong>`;
    offset = start + token.length;
  }
  if (offset < source.length) {
    html += escapeHtml(source.slice(offset));
  }
  return html;
}

function renderCompatInlineWithBreaks(text) {
  return renderCompatInline(text).replace(/\n/g, "<br />");
}

function escapeHtmlPreserveBreaks(text) {
  return renderCompatInlineWithBreaks(text);
}

const HWP_COMPAT_ORDERED_MARKER_PATTERN = "(?:\\d{1,3}[.)]|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|[가-힣][.)])";
const HWP_COMPAT_SYMBOL_MARKER_PATTERN = "(?:[○●◦▪▫■□◆◇▶▷▸▹※·])";
const HWP_COMPAT_BULLET_MARKER_PATTERN = "(?:[○●◦▪▫■□◆◇▶▷▸▹※·\\-–—])";
const HWP_COMPAT_LIST_MARKER_PATTERN = `${HWP_COMPAT_ORDERED_MARKER_PATTERN}|${HWP_COMPAT_BULLET_MARKER_PATTERN}`;
const HWP_COMPAT_INLINE_LIST_REGEX = new RegExp(`(^|\\s+)(${HWP_COMPAT_ORDERED_MARKER_PATTERN}|${HWP_COMPAT_SYMBOL_MARKER_PATTERN})(?=\\s*\\S)`, "g");
const HWP_COMPAT_LOOSE_INLINE_LIST_REGEX = new RegExp(`(${HWP_COMPAT_LIST_MARKER_PATTERN})(?=\\s*\\S)`, "g");

function normalizePipeSeparators(text) {
  return String(text || "").replace(/[｜¦]/g, "|");
}

function trimDecorativeLeadingBars(text) {
  return String(text || "").replace(/^([|｜¦]\s*)+/, "").trimStart();
}

function isOrderedListMarker(marker) {
  const token = String(marker || "").trim();
  return /^\d{1,3}[.)]$/.test(token) || /^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]$/.test(token) || /^[가-힣][.)]$/.test(token);
}

function parseCompatListSegments(rawLine) {
  const normalized = normalizePipeSeparators(String(rawLine || ""))
    .replace(/\t/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) {
    return null;
  }
  const line = trimDecorativeLeadingBars(normalized) || normalized;
  const singleBullet = line.match(/^([-–—])\s+(.+)$/);
  if (singleBullet) {
    return {
      prefix: "",
      items: [{ marker: singleBullet[1], text: singleBullet[2].trim(), ordered: false }],
      listType: "ul"
    };
  }
  HWP_COMPAT_INLINE_LIST_REGEX.lastIndex = 0;
  const markers = [];
  let match = null;
  while ((match = HWP_COMPAT_INLINE_LIST_REGEX.exec(line)) !== null) {
    const leading = match[1] || "";
    const marker = match[2] || "";
    const markerIndex = match.index + leading.length;
    let contentStart = HWP_COMPAT_INLINE_LIST_REGEX.lastIndex;
    while (contentStart < line.length && /\s/.test(line[contentStart])) {
      contentStart += 1;
    }
    markers.push({ marker, markerIndex, contentStart });
  }
  if (markers.length < 2) {
    HWP_COMPAT_LOOSE_INLINE_LIST_REGEX.lastIndex = 0;
    const looseMarkers = [];
    let loose = null;
    while ((loose = HWP_COMPAT_LOOSE_INLINE_LIST_REGEX.exec(line)) !== null) {
      const marker = loose[1] || "";
      const markerIndex = loose.index;
      const prevChar = markerIndex > 0 ? line[markerIndex - 1] : "";
      const isSafeBoundary = markerIndex === 0 || /[\s|｜¦()[\]{}:;,.!?]/.test(prevChar);
      const appearsInsideWord = !isSafeBoundary && /[\p{L}\p{N}]/u.test(prevChar);
      if (appearsInsideWord) {
        continue;
      }
      let contentStart = HWP_COMPAT_LOOSE_INLINE_LIST_REGEX.lastIndex;
      while (contentStart < line.length && /\s/.test(line[contentStart])) {
        contentStart += 1;
      }
      looseMarkers.push({ marker, markerIndex, contentStart });
    }
    if (looseMarkers.length >= 2 || looseMarkers.some((entry) => entry.markerIndex === 0)) {
      markers.splice(0, markers.length, ...looseMarkers);
    }
  }
  if (!markers.length) {
    return null;
  }
  const startsAtLineStart = markers[0].markerIndex === 0;
  if (!startsAtLineStart && markers.length < 2) {
    return null;
  }
  const prefix = startsAtLineStart ? "" : line.slice(0, markers[0].markerIndex).trim();
  const items = [];
  for (let i = 0; i < markers.length; i += 1) {
    const current = markers[i];
    const nextStart = i + 1 < markers.length ? markers[i + 1].markerIndex : line.length;
    const text = trimDecorativeLeadingBars(line.slice(current.contentStart, nextStart)).trim();
    if (!text) {
      continue;
    }
    items.push({
      marker: current.marker,
      text,
      ordered: isOrderedListMarker(current.marker)
    });
  }
  if (!items.length) {
    return null;
  }
  const orderedCount = items.filter((item) => item.ordered).length;
  return {
    prefix,
    items,
    listType: orderedCount >= Math.ceil(items.length / 2) ? "ol" : "ul"
  };
}

function appendListContinuation(listItems, line) {
  if (!listItems.length) {
    return;
  }
  const text = trimDecorativeLeadingBars(String(line || "")).trim();
  if (!text) {
    return;
  }
  const last = listItems[listItems.length - 1];
  last.text = `${String(last.text || "")} ${text}`.replace(/\s{2,}/g, " ").trim();
}

function parsePipeCells(rawLine) {
  const normalized = normalizePipeSeparators(rawLine).trim();
  if (!normalized.includes("|")) {
    return [];
  }
  const cells = normalized.split("|").map((cell) => cell.trim());
  if (cells.length && !cells[0]) {
    cells.shift();
  }
  if (cells.length && !cells[cells.length - 1]) {
    cells.pop();
  }
  return cells;
}

function isPipeDividerLine(rawLine) {
  return /^\s*\|?\s*[:\- ]+\s*(\|\s*[:\- ]+\s*)+\|?\s*$/.test(normalizePipeSeparators(rawLine));
}

function isLikelyPipeTableCandidateLine(rawLine) {
  const normalized = normalizePipeSeparators(String(rawLine || "").trim());
  if (!normalized) {
    return false;
  }
  const barCount = (normalized.match(/\|/g) || []).length;
  if (barCount < 2) {
    return false;
  }
  const cells = parsePipeCells(normalized);
  return cells.length >= 2;
}

function isLikelyListContinuation(rawLine, trimmedLine) {
  if (!trimmedLine) {
    return false;
  }
  if (/^#{1,6}\s+/.test(trimmedLine)) {
    return false;
  }
  if (isLikelyPipeTableCandidateLine(trimmedLine) && !parseCompatListSegments(rawLine)) {
    return false;
  }
  if (/^\s+/.test(String(rawLine || ""))) {
    return true;
  }
  return true;
}

function isBarHeadlineLine(rawLine) {
  const normalized = normalizePipeSeparators(rawLine).trim();
  const barCount = (normalized.match(/\|/g) || []).length;
  if (barCount < 3) {
    return false;
  }
  const cells = parsePipeCells(normalized);
  if (cells.length < 2) {
    return false;
  }
  const nonEmptyCount = cells.filter((cell) => cell.length > 0).length;
  return nonEmptyCount === 1;
}

function buildBarHeadlineBlock(rawLine) {
  const cells = parsePipeCells(rawLine);
  const textIndex = cells.findIndex((cell) => cell.length > 0);
  const text = textIndex >= 0 ? cells[textIndex] : normalizePipeSeparators(rawLine).replace(/\|/g, " ").trim();
  const leftBars = Math.max(1, textIndex);
  const rightBars = Math.max(1, cells.length - textIndex - 1);
  const leftMark = Array.from({ length: leftBars }, () => "|").join(" ");
  const rightMark = Array.from({ length: rightBars }, () => "|").join(" ");
  return `<h3 class="bar-headline"><span class="bar-headline-bars">${escapeHtml(leftMark)}</span><span class="bar-headline-text">${renderCompatInline(text)}</span><span class="bar-headline-bars">${escapeHtml(rightMark)}</span></h3>`;
}

function renderPipeTraceBlocks(pipeLines) {
  const normalizedLines = pipeLines
    .map((line) => normalizePipeSeparators(line).replace(/\s+/g, " ").trim())
    .filter(Boolean);
  if (!normalizedLines.length) {
    return [];
  }

  const blocks = [];
  let tableLines = normalizedLines.slice();
  while (tableLines.length && isBarHeadlineLine(tableLines[0])) {
    blocks.push(buildBarHeadlineBlock(tableLines[0]));
    tableLines = tableLines.slice(1);
  }

  if (!tableLines.length) {
    return blocks;
  }

  const contentLines = tableLines.filter((line) => !isPipeDividerLine(line));
  const parsedRows = contentLines.map((line) => parsePipeCells(line));
  const tableCandidate =
    parsedRows.length >= 2 &&
    parsedRows.every((cells) => Array.isArray(cells) && cells.length >= 2 && cells.some((cell) => cell.length > 0));

  if (!tableCandidate) {
    if (blocks.length) {
      for (const line of tableLines) {
        blocks.push(`<p>${renderCompatInlineWithBreaks(line)}</p>`);
      }
      return blocks;
    }
    return [];
  }

  const maxCols = Math.max(2, ...parsedRows.map((cells) => cells.length));
  const normalizedRows = parsedRows.map((cells) => {
    const row = cells.slice(0, maxCols);
    while (row.length < maxCols) {
      row.push("");
    }
    return row;
  });
  const headerRow = normalizedRows[0];
  const bodyRows = normalizedRows.slice(1);
  const thead = `<thead><tr>${headerRow.map((cell) => `<th>${renderCompatInline(cell)}</th>`).join("")}</tr></thead>`;
  const tbody = bodyRows.length
    ? `<tbody>${bodyRows
      .map((cells) => `<tr>${cells.map((cell) => `<td>${renderCompatInline(cell)}</td>`).join("")}</tr>`)
      .join("")}</tbody>`
    : "";
  blocks.push(`<table class="lookup-table-trace">${thead}${tbody}</table>`);
  return blocks;
}

function buildCompatListBlock(listType, listItems) {
  const tag = listType === "ol" ? "ol" : "ul";
  const itemsHtml = listItems
    .map(
      (item) =>
        `<li class="lookup-list-item"><span class="lookup-list-marker">${escapeHtml(item.marker)}</span><span class="lookup-list-body">${renderCompatInline(item.text)}</span></li>`
    )
    .join("");
  return `<${tag} class="lookup-list ${tag === "ol" ? "lookup-list-ordered" : "lookup-list-unordered"}">${itemsHtml}</${tag}>`;
}

function buildHwpCompatibilityBlocksFromMarkdown(markdownText) {
  const lines = String(markdownText || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+$/g, ""));
  const blocks = [];
  let index = 0;

  const flushParagraph = (bucket) => {
    if (!bucket.length) {
      return;
    }
    const paragraph = bucket.join("\n").trim();
    if (paragraph) {
      blocks.push(`<p>${escapeHtmlPreserveBreaks(paragraph)}</p>`);
    }
    bucket.length = 0;
  };

  const paragraphBuffer = [];
  while (index < lines.length) {
    const rawLine = lines[index];
    const line = rawLine.trim();
    if (!line) {
      flushParagraph(paragraphBuffer);
      index += 1;
      continue;
    }

    const headingMatch = line.match(/^(#{1,6})\s*(.+)$/);
    if (headingMatch) {
      flushParagraph(paragraphBuffer);
      const level = Math.min(4, Math.max(2, headingMatch[1].length + 1));
      const text = headingMatch[2].trim();
      blocks.push(`<h${level}>${renderCompatInline(text)}</h${level}>`);
      index += 1;
      continue;
    }

    const listStart = parseCompatListSegments(rawLine);
    if (listStart) {
      flushParagraph(paragraphBuffer);
      if (listStart.prefix) {
        blocks.push(`<p>${renderCompatInline(listStart.prefix)}</p>`);
      }
      const listItems = listStart.items.map((item) => ({ ...item }));
      const listType = listStart.listType;
      index += 1;

      while (index < lines.length) {
        const nextRaw = lines[index];
        const nextLine = nextRaw.trim();
        if (!nextLine) {
          break;
        }
        if (/^#{1,6}\s+/.test(nextLine)) {
          break;
        }
        const nextList = parseCompatListSegments(nextRaw);
        if (nextList) {
          if (nextList.listType !== listType) {
            break;
          }
          if (nextList.prefix) {
            appendListContinuation(listItems, nextList.prefix);
          }
          for (const item of nextList.items) {
            listItems.push({ ...item });
          }
          index += 1;
          continue;
        }
        if (isLikelyPipeTableCandidateLine(nextLine)) {
          break;
        }
        if (isLikelyListContinuation(nextRaw, nextLine)) {
          appendListContinuation(listItems, nextLine);
          index += 1;
          continue;
        }
        break;
      }
      blocks.push(buildCompatListBlock(listType, listItems));
      continue;
    }

    if (/[|｜¦]/.test(line)) {
      let end = index;
      const pipeLines = [];
      while (end < lines.length) {
        const candidate = lines[end].trim();
        if (!candidate || !/[|｜¦]/.test(candidate)) {
          break;
        }
        pipeLines.push(candidate);
        end += 1;
      }
      flushParagraph(paragraphBuffer);
      const traceBlocks = renderPipeTraceBlocks(pipeLines);
      if (traceBlocks.length) {
        blocks.push(...traceBlocks);
      } else {
        const rawText = pipeLines.map((entry) => trimDecorativeLeadingBars(entry) || entry).join("\n");
        blocks.push(`<p>${renderCompatInlineWithBreaks(rawText)}</p>`);
      }
      index = end;
      continue;
    }

    paragraphBuffer.push(trimDecorativeLeadingBars(line) || line);
    index += 1;
  }

  flushParagraph(paragraphBuffer);
  return blocks;
}

function resolveBundledKoreanFontPath() {
  const candidates = [
    path.join(APP_ROOT, "node_modules", "@ohah", "hwpjs", "fonts", "NotoSansKR-Regular.ttf"),
    path.join(process.resourcesPath || "", "app.asar.unpacked", "node_modules", "@ohah", "hwpjs", "fonts", "NotoSansKR-Regular.ttf"),
    path.join(process.resourcesPath || "", "node_modules", "@ohah", "hwpjs", "fonts", "NotoSansKR-Regular.ttf")
  ];
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    try {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
        return candidate;
      }
    } catch (_error) {
      // ignore invalid path
    }
  }
  return "";
}

function buildBundledFontFaceCss() {
  const fontPath = resolveBundledKoreanFontPath();
  if (!fontPath) {
    return "";
  }
  const fontUrl = pathToFileURL(fontPath).toString();
  return `@font-face {
      font-family: "LookupNotoSansKR";
      src: url("${fontUrl}") format("truetype");
      font-style: normal;
      font-weight: 400;
      font-display: swap;
    }`;
}

function buildHwpCompatibilityDocumentHtmlFromBlocks(blocks, title, options = {}) {
  const safeTitle = escapeHtml(title || "lookup HWP");
  const safeSourceExt = escapeHtml(String(options.sourceExt || "").replace(".", "").toUpperCase());
  const warningMessage = String(options.warningMessage || "").trim();
  const fontFaceCss = buildBundledFontFaceCss();
  const headerBlock =
    options.includeHeader !== false && (safeSourceExt || warningMessage)
      ? `<header class="compat-head">
    <p class="compat-title">${safeTitle}</p>
    ${safeSourceExt ? `<div class="compat-meta">원본 형식: ${safeSourceExt} / lookup 호환 보기</div>` : ""}
    ${warningMessage ? `<div class="compat-warning">${escapeHtml(warningMessage)}</div>` : ""}
  </header>`
      : "";
  const bodyHtml = Array.isArray(blocks) ? blocks.join("\n") : "";
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
  <style>
    ${fontFaceCss}
    @page { size: A4; margin: 18mm; }
    html, body { margin: 0; padding: 0; font-family: "LookupNotoSansKR", "Malgun Gothic", "Segoe UI", sans-serif; color: #101827; font-size: 14px; }
    .sheet { padding: 0; }
    .compat-head { margin: 0 0 14px; border-bottom: 1px solid #d6deea; padding-bottom: 9px; }
    .compat-title { margin: 0 0 4px; font-size: 18px; font-weight: 700; }
    .compat-meta { font-size: 12px; color: #475569; }
    .compat-warning { margin: 8px 0 0; padding: 7px 9px; border: 1px solid #f2d6a8; background: #fff7ea; color: #9a5f08; font-size: 12px; border-radius: 6px; }
    h2, h3, h4 { margin: 12px 0 7px; line-height: 1.40; }
    p { margin: 0 0 10px; line-height: 1.62; white-space: normal; word-break: break-word; }
    .lookup-list { margin: 0 0 10px; padding: 0; list-style: none; }
    .lookup-list-item { margin: 0 0 5px; line-height: 1.58; padding-left: 2.05em; text-indent: -2.05em; word-break: break-word; }
    .lookup-list-marker { display: inline-block; min-width: 1.6em; text-indent: 0; font-weight: 600; color: #0f2749; }
    .lookup-list-body { display: inline; }
    .bar-headline { margin: 8px 0 10px; font-size: 15px; font-weight: 700; line-height: 1.4; color: #0f172a; letter-spacing: 0.01em; }
    .bar-headline-bars { color: #334155; margin: 0 4px 0 0; font-weight: 600; }
    .bar-headline .bar-headline-bars:last-child { margin: 0 0 0 6px; }
    table.lookup-table-trace { width: 100%; border-collapse: collapse; margin: 10px 0 12px; font-size: 13px; table-layout: fixed; break-inside: avoid-page; page-break-inside: avoid; }
    .lookup-table-trace tr { break-inside: avoid; page-break-inside: avoid; }
    .lookup-table-trace:not(.lookup-table-structured) th, .lookup-table-trace:not(.lookup-table-structured) td { border: 1px solid #ccd6e6; padding: 5px 6px; vertical-align: middle; word-break: break-word; line-height: 1.52; }
    .lookup-table-trace:not(.lookup-table-structured) th { background: #f3f6fb; font-weight: 700; }
    .lookup-table-trace .lookup-cell-paragraph { margin: 0 0 6px; }
    .lookup-table-trace .lookup-cell-paragraph:last-child { margin-bottom: 0; }
    table.lookup-table-structured { width: auto; max-width: 100%; border-collapse: collapse; table-layout: fixed; }
    .lookup-table-structured th, .lookup-table-structured td { border: none; padding: 4px 5px; vertical-align: top; word-break: keep-all; overflow-wrap: anywhere; line-height: 1.42; box-sizing: border-box; }
    .lookup-table-structured th { font-weight: 700; }
    .lookup-table-structured td strong,
    .lookup-table-structured th strong { font-weight: 700; }
    .lookup-hwpx-page-break { break-before:page; page-break-before:always; }
    .lookup-hwpx-column-break { break-before:column; }
    .lookup-hwpx-section + .lookup-hwpx-section { break-before:page; page-break-before:always; }
    .lookup-hwpx-picture { vertical-align:middle; }
    ${options.extraCss || ""}
  </style>
</head>
<body>
  <article class="sheet">
    ${headerBlock}
    ${bodyHtml || `<p>${escapeHtml("문서 내용을 추출하지 못했습니다.")}</p>`}
  </article>
</body>
</html>`;
}

function buildHwpCompatibilityHtmlFromMarkdown(markdownText, title, options = {}) {
  const blocks = buildHwpCompatibilityBlocksFromMarkdown(markdownText);
  return buildHwpCompatibilityDocumentHtmlFromBlocks(blocks, title, options);
}

function createDocumentCharStyleResolver(document) {
  return createStructuredCharStyleResolver(document);
}

function createDocumentParaStyleResolver(document) {
  return createStructuredParaStyleResolver(document);
}

function renderStructuredTextRunToHtml(run, resolveCharStyle) {
  const text = normalizeSpecialSymbols(String(run?.text || ""));
  if (!text) {
    return "";
  }
  const escaped = escapeHtml(text);
  const style = typeof resolveCharStyle === "function" ? resolveCharStyle(run?.charPrIDRef) : {};
  const styleAttr = inlineStyleToCss(style || {});
  if (!styleAttr) {
    return escaped;
  }
  return `<span style="${styleAttr}">${escaped}</span>`;
}

function renderStructuredCellParagraphsToHtml(paragraphs, resolveCharStyle, resolveParaStyle, renderPicture = null) {
  const parts = [];
  const alignVotes = { left: 0, right: 0, center: 0, justify: 0 };
  for (const paragraph of Array.isArray(paragraphs) ? paragraphs : []) {
    const runs = Array.isArray(paragraph?.runs) ? paragraph.runs : [];
    const paraStyleId = paragraph?.paraPrIDRef ?? paragraph?.para_shape_id ?? paragraph?.paraShapeId;
    const paraStyle = typeof resolveParaStyle === "function" ? resolveParaStyle(paraStyleId) : {};
    const paraAlign = String(paraStyle?.textAlign || "").trim();
    if (Object.prototype.hasOwnProperty.call(alignVotes, paraAlign)) {
      alignVotes[paraAlign] += 1;
    }
    let lineHtml = "";
    for (const run of runs) {
      if (run?.t === "text") {
        lineHtml += renderStructuredTextRunToHtml(run, resolveCharStyle);
      } else if (run?.t === "picture" && renderPicture) {
        lineHtml += renderPicture(run.picture);
      } else if (run?.t === "break") {
        lineHtml += "<br />";
      }
    }
    const normalized = normalizeText(String(lineHtml).replace(/<br \/>/g, "\n"));
    if (!normalized) {
      continue;
    }
    const paraStyleAttr = paragraphStyleToCss(paraStyle || {});
    const styleAttr = paraStyleAttr ? ` style="${paraStyleAttr}"` : "";
    parts.push(`<p class="lookup-cell-paragraph"${styleAttr}>${lineHtml || "&nbsp;"}</p>`);
  }
  if (!parts.length) {
    return { html: "&nbsp;", horizontalAlign: "" };
  }
  const bestAlign = Object.entries(alignVotes).sort((a, b) => b[1] - a[1])[0];
  return {
    html: parts.join(""),
    horizontalAlign: bestAlign && bestAlign[1] > 0 && bestAlign[0] !== "justify" ? bestAlign[0] : ""
  };
}

function readStructuredTableNumber(source, keys, fallback = 0) {
  const target = source && typeof source === "object" ? source : null;
  const candidates = [
    target,
    target?.attributes,
    target?.cellAttributes,
    target?.cell_attributes
  ].filter(Boolean);
  for (const candidate of candidates) {
    for (const key of Array.isArray(keys) ? keys : []) {
      if (!key) {
        continue;
      }
      const value = Number(candidate[key]);
      if (Number.isFinite(value)) {
        return value;
      }
    }
  }
  return fallback;
}

function readStructuredTableCellAddress(cell, axis, fallback = 0) {
  const address = cell?.addr || cell?.address || cell?.position || {};
  const fromAddress = axis === "row"
    ? Number(address.r ?? address.row ?? address.rowIndex ?? address.row_index)
    : Number(address.c ?? address.col ?? address.column ?? address.colIndex ?? address.col_index);
  if (Number.isFinite(fromAddress)) {
    return Math.max(0, Math.floor(fromAddress));
  }
  const keys = axis === "row"
    ? ["rowAddress", "row_address", "row", "rowIndex", "row_index", "r"]
    : ["colAddress", "col_address", "col", "column", "colIndex", "col_index", "columnIndex", "c"];
  return Math.max(0, Math.floor(readStructuredTableNumber(cell, keys, fallback)));
}

function readStructuredTableCellSpan(cell, axis) {
  const span = cell?.span || {};
  const fromSpan = axis === "row"
    ? Number(span.rows ?? span.row ?? span.rowSpan ?? span.row_span)
    : Number(span.cols ?? span.col ?? span.colSpan ?? span.col_span);
  if (Number.isFinite(fromSpan)) {
    return Math.max(1, Math.floor(fromSpan));
  }
  const keys = axis === "row"
    ? ["rowSpan", "row_span", "rowspan", "rowMerge", "row_merge"]
    : ["colSpan", "col_span", "colspan", "colMerge", "col_merge"];
  return Math.max(1, Math.floor(readStructuredTableNumber(cell, keys, 1)));
}

function readStructuredTableCellSize(cell, axis) {
  const size = cell?.size || {};
  const fromSize = axis === "width"
    ? Number(size.width ?? size.w)
    : Number(size.height ?? size.h);
  if (Number.isFinite(fromSize)) {
    return Math.max(0, fromSize);
  }
  const keys = axis === "width"
    ? ["width", "cellWidth", "cell_width", "w"]
    : ["height", "cellHeight", "cell_height", "h"];
  return Math.max(0, readStructuredTableNumber(cell, keys, 0));
}

function flattenStructuredNumericValues(value) {
  const out = [];
  const visit = (item) => {
    if (Array.isArray(item)) {
      item.forEach(visit);
      return;
    }
    if (item && typeof item === "object") {
      if ("width" in item || "w" in item || "height" in item || "h" in item || "value" in item) {
        visit(item.width ?? item.w ?? item.height ?? item.h ?? item.value);
      }
      return;
    }
    const numeric = Number(item);
    if (Number.isFinite(numeric) && numeric > 0) {
      out.push(numeric);
    }
  };
  visit(value);
  return out;
}

function distributeStructuredDimensions(count, constraints, declared = []) {
  const values = Array.from({ length: count }, (_value, index) => Math.max(0, Number(declared[index] || 0)));
  const spans = [];
  for (const constraint of Array.isArray(constraints) ? constraints : []) {
    const start = Math.max(0, Math.floor(Number(constraint.start || 0)));
    const span = Math.max(1, Math.floor(Number(constraint.span || 1)));
    const measure = Math.max(0, Number(constraint.measure || 0));
    if (start >= count || measure <= 0) {
      continue;
    }
    const end = Math.min(count, start + span);
    if (end <= start) {
      continue;
    }
    if (end - start <= 1) {
      values[start] = Math.max(values[start], measure);
    } else {
      spans.push({ start, end, measure });
    }
  }
  for (let pass = 0; pass < 4; pass += 1) {
    for (const item of spans) {
      const width = item.end - item.start;
      const current = values
        .slice(item.start, item.end)
        .reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0);
      if (current <= 0) {
        const even = item.measure / width;
        for (let index = item.start; index < item.end; index += 1) {
          values[index] = Math.max(values[index], even);
        }
        continue;
      }
      if (current < item.measure * 0.985) {
        const add = (item.measure - current) / width;
        for (let index = item.start; index < item.end; index += 1) {
          values[index] = Math.max(0, values[index] + add);
        }
      }
    }
  }
  return values.map((value) => Math.max(0, Number(value) || 0));
}

function readStructuredTableBorderFillId(cell) {
  return Math.floor(readStructuredTableNumber(
    cell,
    ["borderFillIDRef", "borderFillId", "border_fill_id", "border_fill", "borderFill"],
    -1
  ));
}

function normalizeHeaderKeywordToken(value) {
  return String(value || "")
    .normalize("NFC")
    .replace(/\s+/g, "")
    .replace(/[^\p{L}\p{N}]/gu, "")
    .toLowerCase();
}

function computeHeaderSignalScore(rowTexts = []) {
  const normalized = rowTexts.map((text) => normalizeHeaderKeywordToken(text)).filter(Boolean);
  if (!normalized.length) {
    return 0;
  }
  let score = 0;
  const hasQuestion = normalized.some((token) => token.includes("질문") || token.includes("문항") || token.includes("question"));
  const hasAnswer = normalized.some((token) => token.includes("답변") || token.includes("응답") || token.includes("answer"));
  const hasIndex = normalized.some((token) =>
    token.includes("연번")
    || token.includes("번호")
    || token.includes("순번")
    || token === "no"
    || token.includes("index")
  );
  if (hasQuestion && hasAnswer) {
    score += 3;
  }
  if (hasIndex) {
    score += 2;
  }
  if (normalized.length >= 3) {
    score += 1;
  }
  return score;
}

function resolveStructuredHeaderRowIndex(rowCount, colCount, getCellText) {
  const scanRows = Math.max(1, Math.min(3, Number(rowCount) || 1));
  let bestIndex = -1;
  let bestScore = -1;
  for (let rowIndex = 0; rowIndex < scanRows; rowIndex += 1) {
    const rowTexts = [];
    for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
      const value = String(getCellText(rowIndex, colIndex) || "").trim();
      if (value) {
        rowTexts.push(value);
      }
    }
    if (!rowTexts.length) {
      continue;
    }
    if (shouldTreatFirstRowAsHeader(rowTexts)) {
      return rowIndex;
    }
    const score = computeHeaderSignalScore(rowTexts);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }
  return bestScore >= 3 ? bestIndex : -1;
}

function isNumericLikeText(value) {
  const text = String(value || "").trim();
  if (!text) {
    return false;
  }
  return /^[\d.,%\-+~()\s]+$/.test(text);
}

function detectStructuredCenterPreferredColumns({
  rowCount = 0,
  colCount = 0,
  headerRowIndex = -1,
  getCellText = () => "",
  colWidths = []
} = {}) {
  const columns = new Set();
  const safeColCount = Math.max(0, Math.floor(Number(colCount) || 0));
  const totalWidth = (Array.isArray(colWidths) ? colWidths : [])
    .map((value) => Math.max(0, Number(value) || 0))
    .reduce((sum, value) => sum + value, 0);

  for (let colIndex = 0; colIndex < safeColCount; colIndex += 1) {
    const headerText = headerRowIndex >= 0 ? String(getCellText(headerRowIndex, colIndex) || "") : "";
    const token = normalizeHeaderKeywordToken(headerText);
    const looksIndexHeader =
      token.includes("연번")
      || token.includes("번호")
      || token.includes("순번")
      || token === "no"
      || token.includes("index");
    let numericCount = 0;
    let textCount = 0;
    const scanRows = Math.max(0, Math.min(8, Math.floor(Number(rowCount) || 0)));
    for (let rowIndex = Math.max(0, headerRowIndex + 1); rowIndex < scanRows; rowIndex += 1) {
      const value = String(getCellText(rowIndex, colIndex) || "").trim();
      if (!value) {
        continue;
      }
      textCount += 1;
      if (isNumericLikeText(value)) {
        numericCount += 1;
      }
    }
    const width = Number(Array.isArray(colWidths) ? colWidths[colIndex] : 0);
    const isNarrowColumn = totalWidth > 0 && width > 0 && width / totalWidth <= 0.12;
    if (looksIndexHeader || (isNarrowColumn && textCount > 0 && numericCount >= Math.max(1, textCount - 1))) {
      columns.add(colIndex);
    }
  }
  return columns;
}

function normalizeStructuredTableGrid(table) {
  const rows = Array.isArray(table?.rows) ? table.rows : [];
  if (!rows.length) {
    return null;
  }

  const ownerMap = new Map();
  const coveredMap = new Set();
  const colConstraints = [];
  const rowConstraints = [];
  let maxRow = 0;
  let maxCol = 0;

  rows.forEach((row, rowIndex) => {
    const cells = Array.isArray(row?.cells) ? row.cells : [];
    let colCursor = 0;
    for (const cell of cells) {
      const resolvedRow = readStructuredTableCellAddress(cell, "row", rowIndex);
      let resolvedCol = readStructuredTableNumber(
        cell,
        ["colAddress", "col_address", "col", "column", "colIndex", "col_index", "columnIndex", "c"],
        Number.NaN
      );
      if (!Number.isFinite(resolvedCol)) {
        while (coveredMap.has(`${resolvedRow}:${colCursor}`) || ownerMap.has(`${resolvedRow}:${colCursor}`)) {
          colCursor += 1;
        }
        resolvedCol = colCursor;
      }
      const col = Math.max(0, Math.floor(resolvedCol));
      const rowSpan = readStructuredTableCellSpan(cell, "row");
      const colSpan = readStructuredTableCellSpan(cell, "col");
      ownerMap.set(`${resolvedRow}:${col}`, cell);
      const cellWidth = readStructuredTableCellSize(cell, "width");
      const cellHeight = readStructuredTableCellSize(cell, "height");
      if (cellWidth > 0) {
        colConstraints.push({ start: col, span: colSpan, measure: cellWidth });
      }
      if (cellHeight > 0) {
        rowConstraints.push({ start: resolvedRow, span: rowSpan, measure: cellHeight });
      }

      for (let r = resolvedRow; r < resolvedRow + rowSpan; r += 1) {
        for (let c = col; c < col + colSpan; c += 1) {
          if (r === resolvedRow && c === col) {
            continue;
          }
          coveredMap.add(`${r}:${c}`);
        }
      }

      colCursor = Math.max(colCursor, col + colSpan);
      maxRow = Math.max(maxRow, resolvedRow + rowSpan);
      maxCol = Math.max(maxCol, col + colSpan);
    }
  });

  const declaredWidths = Array.isArray(table?.colWidths)
    ? table.colWidths.map((value) => Math.max(0, Number(value) || 0))
    : flattenStructuredNumericValues(
      table?.attributes?.row_cols
      ?? table?.attributes?.rowCols
      ?? table?.attributes?.col_widths
      ?? table?.attributes?.colWidths
      ?? table?.colWidths
    );
  if (declaredWidths.length > maxCol) {
    maxCol = declaredWidths.length;
  }

  const rowCount = Math.max(1, maxRow, rows.length);
  const colCount = Math.max(1, maxCol);
  const colWidths = distributeStructuredDimensions(colCount, colConstraints, declaredWidths);
  const rowHeights = distributeStructuredDimensions(rowCount, rowConstraints, []);

  return {
    rowCount,
    colCount,
    ownerMap,
    coveredMap,
    colWidths,
    rowHeights
  };
}

function buildStructuredTableBlock(table, resolveCharStyle, resolveParaStyle, borderFillResolver, renderPicture = null) {
  const grid = normalizeStructuredTableGrid(table);
  if (!grid) {
    return "";
  }
  const { rowCount, colCount, ownerMap, coveredMap, colWidths, rowHeights } = grid;
  const plainTextCache = new Map();
  const getCellPlainText = (rowIndex, colIndex) => {
    const key = `${rowIndex}:${colIndex}`;
    if (plainTextCache.has(key)) {
      return plainTextCache.get(key);
    }
    const cell = ownerMap.get(key);
    if (!cell) {
      plainTextCache.set(key, "");
      return "";
    }
    const content = renderStructuredCellParagraphsToHtml(cell?.paragraphs || [], resolveCharStyle, resolveParaStyle, renderPicture);
    const plainText = normalizeText(
      String(content.html || "")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/gi, " ")
    );
    plainTextCache.set(key, plainText);
    return plainText;
  };
  const headerRowIndex = resolveStructuredHeaderRowIndex(rowCount, colCount, getCellPlainText);
  const validWidths = colWidths.filter((value) => value > 0);
  const totalWidth = validWidths.reduce((sum, value) => sum + value, 0);
  const tableWidthPx = hwpUnitToPx(totalWidth);
  const tableStyle = tableWidthPx > 32
    ? ` style="width:${Math.min(760, Math.max(120, tableWidthPx)).toFixed(1)}px;max-width:100%;"`
    : "";
  const colgroup = validWidths.length && totalWidth > 0
    ? `<colgroup>${colWidths.map((value) => {
      if (value <= 0) {
        return "<col />";
      }
      return `<col style="width:${((value / totalWidth) * 100).toFixed(3)}%;" />`;
    }).join("")}</colgroup>`
    : "";
  const centerColumns = detectStructuredCenterPreferredColumns({
    rowCount,
    colCount,
    headerRowIndex,
    getCellText: getCellPlainText,
    colWidths
  });

  const rowHtml = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const cells = [];
    const rowHeightPx = hwpUnitToPx(Array.isArray(rowHeights) ? rowHeights[rowIndex] : 0);
    const rowStyle = rowHeightPx > 10
      ? ` style="height:${Math.min(1800, Math.max(10, rowHeightPx)).toFixed(1)}px;"`
      : "";
    for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
      const key = `${rowIndex}:${colIndex}`;
      if (coveredMap.has(key)) {
        continue;
      }
      const cell = ownerMap.get(key);
      if (!cell) {
        cells.push("<td>&nbsp;</td>");
        continue;
      }
      const colSpan = readStructuredTableCellSpan(cell, "col");
      const rowSpan = readStructuredTableCellSpan(cell, "row");
      const isHeader = rowIndex === headerRowIndex;
      const content = renderStructuredCellParagraphsToHtml(cell?.paragraphs || [], resolveCharStyle, resolveParaStyle, renderPicture);
      const centerPreferred = centerColumns.has(colIndex);
      const preferredAlign = isHeader ? "center" : (content.horizontalAlign || (centerPreferred ? "center" : ""));
      const baseCellAttributes = (
        (cell?.cellAttributes && typeof cell.cellAttributes === "object" && cell.cellAttributes)
        || (cell?.cell_attributes && typeof cell.cell_attributes === "object" && cell.cell_attributes)
        || (cell?.attributes && typeof cell.attributes === "object" && cell.attributes)
        || {}
      );
      const cellAttributes = {
        ...baseCellAttributes,
        border_fill_id: readStructuredTableBorderFillId(cell)
      };
      const cellStyle = buildTableCellStyle({
        cellAttributes,
        tableAttributes: (table?.attributes && typeof table.attributes === "object") ? table.attributes : (table || {}),
        borderFillResolver: typeof borderFillResolver === "function" ? borderFillResolver : () => null,
        horizontalAlign: preferredAlign,
        columnRole: centerPreferred ? "index" : "",
        isHeader,
        rowSpan,
        cellWidth: readStructuredTableCellSize(cell, "width"),
        cellHeight: readStructuredTableCellSize(cell, "height"),
        forceMiddle: preferredAlign === "center" || isHeader || rowSpan > 1 || centerPreferred
      });
      const styleEntries = Object.entries(cellStyle).filter(([, value]) => value !== undefined && value !== null && value !== "");
      const styleAttr = styleEntries.length
        ? ` style="${styleEntries.map(([name, value]) => `${name}:${value}`).join(";")}"`
        : "";
      const tag = isHeader ? "th" : "td";
      cells.push(`<${tag} colspan="${colSpan}" rowspan="${rowSpan}"${styleAttr}>${content.html}</${tag}>`);
    }
    rowHtml.push(cells.length ? `<tr${rowStyle}>${cells.join("")}</tr>` : `<tr${rowStyle}></tr>`);
  }

  if (!rowHtml.length) {
    return "";
  }
  return `<table class="lookup-table-trace lookup-table-structured"${tableStyle}>${colgroup}<tbody>${rowHtml.join("")}</tbody></table>`;
}

function buildStructuredCompatibilityHtmlFromDocument(document, title, options = {}) {
  const sections = Array.isArray(document?.sections) ? document.sections : [];
  if (!sections.length) {
    return "";
  }
  const resolveCharStyle = createDocumentCharStyleResolver(document);
  const resolveParaStyle = createDocumentParaStyleResolver(document);
  const borderFillResolver = createBorderFillResolver(document);
  const blocks = [];
  const pageStyles = [];
  let hasVisibleContent = false;
  const renderPicture = picture => buildHwpxPictureHtml(picture, Array.isArray(document?.binData) ? document.binData : []);

  for (const [sectionIndex, section] of sections.entries()) {
    const layout = buildHwpxSectionLayout(section.def, sectionIndex);
    pageStyles.push(layout.css);
    blocks.push(`<section class="lookup-hwpx-section" style="${layout.style}">`);
    const paragraphs = Array.isArray(section?.paragraphs) ? section.paragraphs : [];
    for (const paragraph of paragraphs) {
      if (paragraph?.pageBreak) blocks.push('<div class="lookup-hwpx-page-break"></div>');
      const runs = Array.isArray(paragraph?.runs) ? paragraph.runs : [];
      let paragraphHtml = "";
      for (const run of runs) {
        if (run?.t === "table" && run?.table) {
          if (String(paragraphHtml || "").trim()) {
            const paraStyle = resolveParaStyle(paragraph?.paraPrIDRef ?? paragraph?.para_shape_id ?? paragraph?.paraShapeId);
            const paraStyleAttr = paragraphStyleToCss(paraStyle || {});
            const styleAttr = paraStyleAttr ? ` style="${paraStyleAttr}"` : "";
            blocks.push(`<p${styleAttr}>${paragraphHtml}</p>`);
            paragraphHtml = "";
          }
          const tableBlock = buildStructuredTableBlock(run.table, resolveCharStyle, resolveParaStyle, borderFillResolver, renderPicture);
          if (tableBlock) {
            blocks.push(tableBlock);
            hasVisibleContent = true;
          }
          continue;
        }
        if (run?.t === "text") {
          paragraphHtml += renderStructuredTextRunToHtml(run, resolveCharStyle);
          continue;
        }
        if (run?.t === "picture") {
          paragraphHtml += renderPicture(run.picture);
          continue;
        }
        if (run?.t === "break") {
          paragraphHtml += run.breakType === "PAGE"
            ? '<span class="lookup-hwpx-page-break" style="display:block"></span>'
            : run.breakType === "COLUMN" ? '<span class="lookup-hwpx-column-break" style="display:block"></span>' : "<br />";
        }
      }
      if (String(paragraphHtml || "").trim()) {
        hasVisibleContent = true;
        const paraStyle = resolveParaStyle(paragraph?.paraPrIDRef ?? paragraph?.para_shape_id ?? paragraph?.paraShapeId);
        const paraStyleAttr = paragraphStyleToCss(paraStyle || {});
        const styleAttr = paraStyleAttr ? ` style="${paraStyleAttr}"` : "";
        blocks.push(`<p${styleAttr}>${paragraphHtml}</p>`);
      } else if (!runs.length || runs.every(run => run?.t === "text" && !String(run.text || "").trim())) {
        blocks.push('<p class="lookup-hwpx-empty-paragraph">&nbsp;</p>');
      }
    }
    blocks.push("</section>");
  }

  if (!hasVisibleContent) {
    return "";
  }
  return buildHwpCompatibilityDocumentHtmlFromBlocks(blocks, title, { ...options, includeHeader: false, extraCss: pageStyles.join("\n") });
}

function parseHwpJsonDocument(jsonValue) {
  let parsed = null;
  if (!jsonValue) {
    return null;
  }
  if (typeof jsonValue === "string") {
    try {
      parsed = JSON.parse(jsonValue);
    } catch (_error) {
      return null;
    }
  } else if (typeof jsonValue === "object") {
    parsed = jsonValue;
  }
  if (!parsed || typeof parsed !== "object") {
    return null;
  }
  if (Array.isArray(parsed.sections)) {
    return parsed;
  }
  if (parsed.document && Array.isArray(parsed.document.sections)) {
    return parsed.document;
  }
  if (parsed.body && Array.isArray(parsed.body.sections)) {
    return {
      head: parsed.head || {},
      sections: parsed.body.sections
    };
  }
  return null;
}

function buildStructuredCompatibilityHtmlFromHwpJson(jsonValue, title, options = {}) {
  const hwpBlocks = buildHwpBlocksFromHwpJson(jsonValue);
  if (Array.isArray(hwpBlocks) && hwpBlocks.length) {
    return buildHwpCompatibilityDocumentHtmlFromBlocks(hwpBlocks, title, options);
  }
  const document = parseHwpJsonDocument(jsonValue);
  if (!document) {
    return "";
  }
  return buildStructuredCompatibilityHtmlFromDocument(document, title, options);
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

async function getPdfPageCount(pdfPath) {
  if (!pdfPath || !fs.existsSync(pdfPath)) {
    return 0;
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
    const pageCount = Number(doc.numPages || 0);
    await doc.destroy();
    try {
      await loadingTask.destroy();
    } catch (_error) {
      // ignore cleanup failure
    }
    return Number.isFinite(pageCount) ? pageCount : 0;
  } catch (_error) {
    return 0;
  }
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
        pageCount: 0,
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
        pageCount: Number(doc.numPages || 0),
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
        pageCount: 0,
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

function sortHwpCandidatesByQuality(candidates = []) {
  return [...candidates].sort((a, b) => {
    if (b.qualityScore !== a.qualityScore) {
      return b.qualityScore - a.qualityScore;
    }
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    if (getHwpModePriority(b.mode) !== getHwpModePriority(a.mode)) {
      return getHwpModePriority(b.mode) - getHwpModePriority(a.mode);
    }
    return Number(b.pageCount || 0) - Number(a.pageCount || 0);
  });
}

function chooseBestHwpCandidate(candidates) {
  const valid = candidates.filter((candidate) => candidate.available && candidate.selectionEligible !== false);
  if (!valid.length) {
    return null;
  }
  const sortedAll = sortHwpCandidatesByQuality(valid);
  const bestOverall = sortedAll[0] || null;
  const bestNative = sortHwpCandidatesByQuality(valid.filter((candidate) => isNativeHwpMode(candidate.mode)))[0] || null;
  if (!bestOverall) {
    return null;
  }
  if (!bestNative) {
    return bestOverall;
  }

  const nativeScore = Number(bestNative.qualityScore || bestNative.score || 0);
  const overallScore = Number(bestOverall.qualityScore || bestOverall.score || 0);
  const nativeArtifact = Number(bestNative.artifactScore || 0);
  const overallArtifact = Number(bestOverall.artifactScore || 0);
  const nativeLooksWeak = nativeScore < 60 || nativeArtifact >= 45;
  const overallLooksStrong = overallScore >= 72 && overallArtifact <= 34;
  const qualityGapLarge = overallScore >= nativeScore + 12;

  if (bestOverall.mode !== bestNative.mode && ((nativeLooksWeak && overallLooksStrong) || qualityGapLarge)) {
    return bestOverall;
  }

  return bestNative;
}

function applyHwpPageCountGate(candidates, options = {}) {
  const available = candidates.filter((candidate) => candidate.available && Number(candidate.pageCount || 0) > 0);
  if (!available.length) {
    return;
  }
  const forcedPageTarget = Math.max(0, Math.floor(Number(options.forcedPageTarget || 0)));
  if (forcedPageTarget > 0) {
    for (const candidate of available) {
      const pageCount = Number(candidate.pageCount || 0);
      if (pageCount !== forcedPageTarget) {
        candidate.selectionEligible = false;
        candidate.excludedReason = "PAGE_COUNT_TARGET_MISMATCH";
        candidate.pageTargetDelta = Math.abs(pageCount - forcedPageTarget);
        candidate.qualityScore = clampNumber(Number(candidate.qualityScore || candidate.score || 0) - 58, 0, 100);
      } else {
        candidate.pageTargetDelta = 0;
      }
    }
    return;
  }
  const maxPageCount = Math.max(...available.map((candidate) => Number(candidate.pageCount || 0)));
  if (!Number.isFinite(maxPageCount) || maxPageCount < 2) {
    return;
  }
  const minAllowed = Math.max(2, Math.ceil(maxPageCount * 0.8));
  for (const candidate of available) {
    if (Number(candidate.pageCount || 0) < minAllowed) {
      candidate.selectionEligible = false;
      candidate.excludedReason = "PAGE_COUNT_LOW";
      candidate.qualityScore = clampNumber(Number(candidate.qualityScore || candidate.score || 0) - 45, 0, 100);
    }
  }
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
  const startedAt = Date.now();
  const result = await runPowerShellWorker(buildExcelReadOnlyScript(sourcePath, outputPdfPath), 45000);
  if (result.timedOut) {
    const processId = Number(String(result.stdout || "").match(/"excelProcessId":(\d+)/)?.[1] || 0);
    const cleanupScript = buildOwnedExcelCleanupScript(processId, startedAt);
    if (cleanupScript) await runPowerShellAsync(cleanupScript, 5000);
  }
  return result;
}

async function convertWithPowerPointCom(sourcePath, outputPdfPath) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$powerPoint = $null",
    "$presentation = $null",
    "try {",
    "  $powerPoint = New-Object -ComObject PowerPoint.Application",
    "  $presentation = $powerPoint.Presentations.Open(" + source + ", $false, $true, $false)",
    "  $presentation.SaveAs(" + output + ", 32)",
    "} finally {",
    "  if ($presentation -ne $null) { $presentation.Close() | Out-Null }",
    "  if ($powerPoint -ne $null) { $powerPoint.Quit() | Out-Null }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'powerpoint_export_failed' }"
  ].join("; ");

  return await runPowerShellAsync(script, 180000);
}

async function convertWithHancomComPrintPdf(sourcePath, outputPdfPath, ext = ".hwp", statusCallback = null) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const registrationPlan = getHancomSecurityRegistrationPlan();
  const moduleCandidates = buildPowerShellStringArray(registrationPlan.installedModuleCandidates);
  const allowPaths = buildPowerShellStringArray(getHwpSecurityAllowList(sourcePath, outputPdfPath));
  const lookupModuleName = escapeForPowerShell(registrationPlan.lookupModuleName || "lookupFilePathChecker");
  const trustedLookupModulePath = escapeForPowerShell(registrationPlan.lookupModulePath || "");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "$moduleRoot = 'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules'",
    "$allowRoot = 'HKCU:\\SOFTWARE\\lookup\\HwpSecurity\\AllowedPaths'",
    "$moduleName = " + lookupModuleName,
    "$lookupModulePath = " + trustedLookupModulePath,
    "$modulePath = ''",
    "$registered = $false",
    "$prevMsgMode = 0",
    `$moduleCandidates = ${moduleCandidates}`,
    `$allowPaths = ${allowPaths}`,
    "Write-Output 'LOOKUP_STATUS {\"stage\":\"engine\",\"percent\":10,\"messageKey\":\"convert.step.engine\",\"mode\":\"hancom-com-print\"}'",
    "try {",
    "  if (-not (Test-Path $moduleRoot)) { New-Item -Path $moduleRoot -Force | Out-Null }",
    "  if (-not (Test-Path $allowRoot)) { New-Item -Path $allowRoot -Force | Out-Null }",
    "  try { Get-ItemProperty -Path $allowRoot -ErrorAction SilentlyContinue | Out-Null } catch {}",
    "  $pathIndex = 0",
    "  foreach ($allowed in $allowPaths) {",
    "    if (-not $allowed) { continue }",
    "    $pathIndex = $pathIndex + 1",
    "    try { Set-ItemProperty -Path $allowRoot -Name ('Path' + $pathIndex) -Value ([string]$allowed) -Type String -Force } catch {}",
    "  }",
    "  if ($lookupModulePath -and (Test-Path $lookupModulePath)) {",
    "    $modulePath = [string]$lookupModulePath",
    "  }",
    "  if (-not $modulePath) {",
    "    foreach ($candidate in $moduleCandidates) {",
    "      if ($candidate -and (Test-Path $candidate)) {",
    "        $modulePath = [string]$candidate",
    "        break",
    "      }",
    "    }",
    "  }",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name $moduleName -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  if ($modulePath -and (-not ($modulePath -eq $lookupModulePath))) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name 'FilePathCheckerModule' -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  Write-Output 'lookup_engine_state=hwp_object_created'",
    "  try { $prevMsgMode = $hwp.SetMessageBoxMode(0x10011) } catch {}",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', $moduleName) | Out-Null; $registered = $true; Write-Output ('lookup_module_name=' + $moduleName); Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if (-not $registered) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null; $registered = $true; Write-Output 'lookup_module_name=FilePathCheckerModule'; Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if ($registered) { Write-Output 'lookup_security_registered=1' } else { Write-Output 'lookup_security_registered=0'; Write-Output 'lookup_engine_state=security_module_missing'; throw 'hwp_security_module_missing' }",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"open\",\"percent\":28,\"messageKey\":\"convert.step.open\",\"mode\":\"hancom-com-print\"}'",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { Write-Output 'lookup_engine_state=open_failed'; throw 'hwp_open_failed' }",
    "  Write-Output 'lookup_engine_state=open_ok'",
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
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"save\",\"percent\":62,\"messageKey\":\"convert.step.print-fallback\",\"mode\":\"hancom-com-print\"}'",
    "  $printed = $false",
    "  try { $printed = $hwp.HAction.Execute('Print', $hwp.HParameterSet.HPrint.HSet) } catch { $printed = $false }",
    "  if (-not $printed) { Write-Output 'lookup_engine_state=print_failed'; throw 'hwp_print_failed' }",
    "  Write-Output 'lookup_engine_state=print_ok'",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"validate\",\"percent\":86,\"messageKey\":\"convert.step.validate\",\"mode\":\"hancom-com-print\"}'",
    "  Write-Output ('lookup_output_path=' + " + output + ")",
    "} finally {",
    "  if ($hwp -ne $null) {",
    "    try { $hwp.SetMessageBoxMode(0xFFFFFF) | Out-Null } catch {}",
    "    try { $hwp.Quit() | Out-Null } catch {}",
    "  }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("\n");
  const result = await runPowerShellWorker(script, 120000, { onStatus: statusCallback });
  return enrichHancomRunResult(result, "hwp-print-pdf");
}

async function convertWithHancomComExport(sourcePath, outputPdfPath, ext = ".hwp", statusCallback = null) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const registrationPlan = getHancomSecurityRegistrationPlan();
  const moduleCandidates = buildPowerShellStringArray(registrationPlan.installedModuleCandidates);
  const allowPaths = buildPowerShellStringArray(getHwpSecurityAllowList(sourcePath, outputPdfPath));
  const lookupModuleName = escapeForPowerShell(registrationPlan.lookupModuleName || "lookupFilePathChecker");
  const trustedLookupModulePath = escapeForPowerShell(registrationPlan.lookupModulePath || "");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "$moduleRoot = 'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules'",
    "$allowRoot = 'HKCU:\\SOFTWARE\\lookup\\HwpSecurity\\AllowedPaths'",
    "$moduleName = " + lookupModuleName,
    "$lookupModulePath = " + trustedLookupModulePath,
    "$modulePath = ''",
    "$registered = $false",
    "$exported = $false",
    "$prevMsgMode = 0",
    `$moduleCandidates = ${moduleCandidates}`,
    `$allowPaths = ${allowPaths}`,
    "Write-Output 'LOOKUP_STATUS {\"stage\":\"engine\",\"percent\":10,\"messageKey\":\"convert.step.engine\",\"mode\":\"hancom-com-export\"}'",
    "try {",
    "  if (-not (Test-Path $moduleRoot)) { New-Item -Path $moduleRoot -Force | Out-Null }",
    "  if (-not (Test-Path $allowRoot)) { New-Item -Path $allowRoot -Force | Out-Null }",
    "  $pathIndex = 0",
    "  foreach ($allowed in $allowPaths) {",
    "    if (-not $allowed) { continue }",
    "    $pathIndex = $pathIndex + 1",
    "    try { Set-ItemProperty -Path $allowRoot -Name ('Path' + $pathIndex) -Value ([string]$allowed) -Type String -Force } catch {}",
    "  }",
    "  if ($lookupModulePath -and (Test-Path $lookupModulePath)) {",
    "    $modulePath = [string]$lookupModulePath",
    "  }",
    "  if (-not $modulePath) {",
    "    foreach ($candidate in $moduleCandidates) {",
    "      if ($candidate -and (Test-Path $candidate)) {",
    "        $modulePath = [string]$candidate",
    "        break",
    "      }",
    "    }",
    "  }",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name $moduleName -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  if ($modulePath -and (-not ($modulePath -eq $lookupModulePath))) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name 'FilePathCheckerModule' -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  Write-Output 'lookup_engine_state=hwp_object_created'",
    "  try { $prevMsgMode = $hwp.SetMessageBoxMode(0x10011) } catch {}",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', $moduleName) | Out-Null; $registered = $true; Write-Output ('lookup_module_name=' + $moduleName); Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if (-not $registered) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null; $registered = $true; Write-Output 'lookup_module_name=FilePathCheckerModule'; Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if ($registered) { Write-Output 'lookup_security_registered=1' } else { Write-Output 'lookup_security_registered=0'; Write-Output 'lookup_engine_state=security_module_missing'; throw 'hwp_security_module_missing' }",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"open\",\"percent\":28,\"messageKey\":\"convert.step.open\",\"mode\":\"hancom-com-export\"}'",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { Write-Output 'lookup_engine_state=open_failed'; throw 'hwp_open_failed' }",
    "  Write-Output 'lookup_engine_state=open_ok'",
    "  $doc = $null",
    "  try { $doc = $hwp.XHwpDocuments.Item(0) } catch { $doc = $null }",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"save\",\"percent\":62,\"messageKey\":\"convert.step.save\",\"mode\":\"hancom-com-export\"}'",
    "  if ($doc -ne $null) {",
      "    try {",
      "      $member = $doc | Get-Member -Name 'ExportAsFixedFormat' -ErrorAction SilentlyContinue",
    "      if ($member) {",
    "        $doc.ExportAsFixedFormat(" + output + ", 'PDF')",
    "        if (Test-Path " + output + ") { $exported = $true }",
    "      }",
    "    } catch { $exported = $false }",
    "  }",
    "  if (-not $exported) { Write-Output 'lookup_engine_state=export_unavailable'; throw 'hwp_export_method_unavailable' }",
    "  Write-Output 'lookup_engine_state=export_ok'",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"validate\",\"percent\":86,\"messageKey\":\"convert.step.validate\",\"mode\":\"hancom-com-export\"}'",
    "  Write-Output ('lookup_output_path=' + " + output + ")",
    "} finally {",
    "  if ($hwp -ne $null) {",
    "    try { $hwp.SetMessageBoxMode(0xFFFFFF) | Out-Null } catch {}",
    "    try { $hwp.Quit() | Out-Null } catch {}",
    "  }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("\n");
  const result = await runPowerShellWorker(script, 90000, { onStatus: statusCallback });
  return enrichHancomRunResult(result, "hwp-export");
}

async function convertWithHancomComSaveAs(sourcePath, outputPdfPath, ext = ".hwp", statusCallback = null) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const format = ext === ".hwpx" ? "HWPX" : "HWP";
  const registrationPlan = getHancomSecurityRegistrationPlan();
  const moduleCandidates = buildPowerShellStringArray(registrationPlan.installedModuleCandidates);
  const allowPaths = buildPowerShellStringArray(getHwpSecurityAllowList(sourcePath, outputPdfPath));
  const lookupModuleName = escapeForPowerShell(registrationPlan.lookupModuleName || "lookupFilePathChecker");
  const trustedLookupModulePath = escapeForPowerShell(registrationPlan.lookupModulePath || "");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "$moduleRoot = 'HKCU:\\SOFTWARE\\HNC\\HwpAutomation\\Modules'",
    "$allowRoot = 'HKCU:\\SOFTWARE\\lookup\\HwpSecurity\\AllowedPaths'",
    "$moduleName = " + lookupModuleName,
    "$lookupModulePath = " + trustedLookupModulePath,
    "$modulePath = ''",
    "$registered = $false",
    "$prevMsgMode = 0",
    `$moduleCandidates = ${moduleCandidates}`,
    `$allowPaths = ${allowPaths}`,
    "Write-Output 'LOOKUP_STATUS {\"stage\":\"engine\",\"percent\":10,\"messageKey\":\"convert.step.engine\",\"mode\":\"hancom-com-saveas\"}'",
    "try {",
    "  if (-not (Test-Path $moduleRoot)) { New-Item -Path $moduleRoot -Force | Out-Null }",
    "  if (-not (Test-Path $allowRoot)) { New-Item -Path $allowRoot -Force | Out-Null }",
    "  $pathIndex = 0",
    "  foreach ($allowed in $allowPaths) {",
    "    if (-not $allowed) { continue }",
    "    $pathIndex = $pathIndex + 1",
    "    try { Set-ItemProperty -Path $allowRoot -Name ('Path' + $pathIndex) -Value ([string]$allowed) -Type String -Force } catch {}",
    "  }",
    "  if ($lookupModulePath -and (Test-Path $lookupModulePath)) {",
    "    $modulePath = [string]$lookupModulePath",
    "  }",
    "  if (-not $modulePath) {",
    "    foreach ($candidate in $moduleCandidates) {",
    "      if ($candidate -and (Test-Path $candidate)) {",
    "        $modulePath = [string]$candidate",
    "        break",
    "      }",
    "    }",
    "  }",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name $moduleName -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  if ($modulePath -and (-not ($modulePath -eq $lookupModulePath))) {",
    "    try { Set-ItemProperty -Path $moduleRoot -Name 'FilePathCheckerModule' -Value $modulePath -Type String -Force } catch {}",
    "  }",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  Write-Output 'lookup_engine_state=hwp_object_created'",
    "  try { $prevMsgMode = $hwp.SetMessageBoxMode(0x10011) } catch {}",
    "  if ($modulePath -and ($modulePath -eq $lookupModulePath)) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', $moduleName) | Out-Null; $registered = $true; Write-Output ('lookup_module_name=' + $moduleName); Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if (-not $registered) {",
    "    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null; $registered = $true; Write-Output 'lookup_module_name=FilePathCheckerModule'; Write-Output ('lookup_module_path=' + $modulePath) } catch {}",
    "  }",
    "  if ($registered) { Write-Output 'lookup_security_registered=1' } else { Write-Output 'lookup_security_registered=0'; Write-Output 'lookup_engine_state=security_module_missing'; throw 'hwp_security_module_missing' }",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"open\",\"percent\":28,\"messageKey\":\"convert.step.open\",\"mode\":\"hancom-com-saveas\"}'",
    "  $opened = $hwp.Open(" + source + ", '" + format + "', 'forceopen:true;versionwarning:false')",
    "  if (-not $opened) { Write-Output 'lookup_engine_state=open_failed'; throw 'hwp_open_failed' }",
    "  Write-Output 'lookup_engine_state=open_ok'",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"save\",\"percent\":62,\"messageKey\":\"convert.step.save\",\"mode\":\"hancom-com-saveas\"}'",
    "  $saved = $hwp.SaveAs(" + output + ", 'PDF', 'embedfont:true;bookmark:true;compress:false')",
    "  if (-not $saved) { Write-Output 'lookup_engine_state=saveas_failed'; throw 'hwp_save_failed' }",
    "  Write-Output 'lookup_engine_state=saveas_ok'",
    "  Write-Output 'LOOKUP_STATUS {\"stage\":\"validate\",\"percent\":86,\"messageKey\":\"convert.step.validate\",\"mode\":\"hancom-com-saveas\"}'",
    "  Write-Output ('lookup_output_path=' + " + output + ")",
    "} finally {",
    "  if ($hwp -ne $null) {",
    "    try { $hwp.SetMessageBoxMode(0xFFFFFF) | Out-Null } catch {}",
    "    try { $hwp.Quit() | Out-Null } catch {}",
    "  }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("\n");
  const result = await runPowerShellWorker(script, 90000, { onStatus: statusCallback });
  return enrichHancomRunResult(result, "hwp-saveas");
}

function findHancomConverterExecutable() {
  const roots = [
    process.env["ProgramFiles"] || "C:\\Program Files",
    process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)"
  ];
  const relativeCandidates = [
    "HNC\\Office 2024\\HOffice120\\Bin\\HpdfHwpConverter.exe",
    "HNC\\Office 2022\\HOffice120\\Bin\\HpdfHwpConverter.exe",
    "HNC\\Office 2020\\HOffice110\\Bin\\HpdfHwpConverter.exe",
    "Hancom\\Office 2024\\HOffice120\\Bin\\HpdfHwpConverter.exe",
    "Hancom\\Office 2022\\HOffice120\\Bin\\HpdfHwpConverter.exe",
    "Hancom\\Office 2020\\HOffice110\\Bin\\HpdfHwpConverter.exe",
    "Hancom\\HOffice\\Bin\\HpdfHwpConverter.exe",
    "HNC\\HOffice\\Bin\\HpdfHwpConverter.exe",
    "Hancom\\HOffice\\Bin\\HwpConverter.exe",
    "HNC\\HOffice\\Bin\\HwpConverter.exe"
  ];
  for (const root of roots) {
    for (const relPath of relativeCandidates) {
      const candidate = path.join(root, relPath);
      if (fs.existsSync(candidate)) {
        return candidate;
      }
    }
  }
  return "";
}

async function convertWithHancomConverterExe(sourcePath, outputPdfPath) {
  const converterExe = findHancomConverterExecutable();
  if (!converterExe) {
    return {
      ok: false,
      missingEngine: true,
      stderr: "hancom_converter_missing"
    };
  }

  const attempts = [
    [sourcePath, outputPdfPath],
    ["-i", sourcePath, "-o", outputPdfPath],
    ["/i", sourcePath, "/o", outputPdfPath],
    ["--input", sourcePath, "--output", outputPdfPath],
    ["/input", sourcePath, "/output", outputPdfPath]
  ];

  for (const args of attempts) {
    const result = await runProcess(converterExe, args, 90000);
    if (result.ok && fs.existsSync(outputPdfPath)) {
      return {
        ok: true,
        stdout: result.stdout || "",
        stderr: result.stderr || "",
        error: ""
      };
    }
  }

  return {
    ok: false,
    stderr: "hancom_converter_failed",
    error: "hancom_converter_failed"
  };
}

async function convertWithLibreOffice(sourcePath, outputPdfPath) {
  return convertOfficeWithLibreOffice(sourcePath, outputPdfPath, findLibreOfficeExecutable());
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
    const markdownRaw =
      typeof markdownResult === "string"
        ? markdownResult
        : markdownResult && typeof markdownResult.markdown === "string"
          ? markdownResult.markdown
          : "";
    if (!markdownRaw) {
      return "";
    }
    return sanitizeHwpCompatibilityMarkdown(markdownRaw).markdown;
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

function analyzeHwpJsMarkdownArtifacts(hwpjs, rawBuffer, suppliedMarkdown = null) {
  try {
    const markdownResult = suppliedMarkdown === null ? hwpjs.toMarkdown(rawBuffer) : suppliedMarkdown;
    const markdownRaw =
      typeof markdownResult === "string"
        ? markdownResult
        : typeof markdownResult?.markdown === "string"
          ? markdownResult.markdown
          : "";
    if (!markdownRaw) {
      return {
        inspected: false,
        totalChars: 0,
        strikeChars: 0,
        strikeSegments: 0,
        strikeRatio: 0,
        totalLines: 0,
        tableLineCount: 0,
        noisyTableLineCount: 0,
        headingOnlyCount: 0,
        symbolNoiseRatio: 0
      };
    }
    let strikeChars = 0;
    let strikeSegments = 0;
    const strikeRegex = /~~([^~]+)~~/g;
    let match = null;
    while ((match = strikeRegex.exec(markdownRaw)) !== null) {
      strikeSegments += 1;
      strikeChars += String(match[1] || "").replace(/\s+/g, "").length;
    }
    const plainChars = markdownRaw
      .replace(/```[\s\S]*?```/g, " ")
      .replace(/`[^`]*`/g, " ")
      .replace(/[#>*_\-\[\]\(\)!]/g, " ")
      .replace(/\s+/g, "")
      .length;
    const lines = markdownRaw.split(/\r?\n/);
    const tableLineCount = lines.filter((line) => /^\s*\|.*\|\s*$/.test(String(line || "").trim())).length;
    const noisyTableLineCount = lines.filter((line) => {
      const trimmed = String(line || "").trim();
      if (!/^\s*\|.*\|\s*$/.test(trimmed)) {
        return false;
      }
      const cells = trimmed
        .split("|")
        .map((cell) => cell.trim())
        .filter(Boolean);
      if (!cells.length) {
        return true;
      }
      const shortCells = cells.filter((cell) => cell.length <= 1).length;
      const hasBrokenToken = /[�□]{2,}|[`~*_]{3,}/.test(trimmed);
      return hasBrokenToken || shortCells / cells.length >= 0.62;
    }).length;
    const headingOnlyCount = lines.filter((line) => /^\s*#{1,6}\s*$/.test(String(line || "").trim())).length;
    const rawChars = markdownRaw.replace(/\s+/g, "");
    const noiseCharCount = rawChars.replace(/[0-9A-Za-z가-힣ㄱ-ㅎㅏ-ㅣ.,:;!?()[\]{}'"“”‘’/\-+%=<>]/g, "").length;
    const symbolNoiseRatio = rawChars.length > 0 ? clampNumber(noiseCharCount / rawChars.length, 0, 1) : 0;
    const totalChars = Math.max(plainChars, 0);
    const strikeRatio = totalChars > 0 ? clampNumber(strikeChars / totalChars, 0, 1) : 0;
    return {
      inspected: true,
      totalChars,
      strikeChars,
      strikeSegments,
      strikeRatio,
      totalLines: lines.length,
      tableLineCount,
      noisyTableLineCount,
      headingOnlyCount,
      symbolNoiseRatio
    };
  } catch (_error) {
    return {
      inspected: false,
      totalChars: 0,
      strikeChars: 0,
      strikeSegments: 0,
      strikeRatio: 0,
      totalLines: 0,
      tableLineCount: 0,
      noisyTableLineCount: 0,
      headingOnlyCount: 0,
      symbolNoiseRatio: 0
    };
  }
}

function countHtmlTagOccurrences(html, tagName) {
  const source = String(html || "");
  if (!source) {
    return 0;
  }
  const matcher = new RegExp(`<${tagName}\\b`, "gi");
  return (source.match(matcher) || []).length;
}

function readHtmlSpanAttribute(cellTag, attrName) {
  const source = String(cellTag || "");
  const matcher = new RegExp(`${attrName}\\s*=\\s*[\"']?(\\d+)`, "i");
  const match = source.match(matcher);
  if (!match) {
    return 1;
  }
  const value = Number(match[1]);
  return Number.isFinite(value) ? Math.max(1, Math.floor(value)) : 1;
}

function countMergedCellsInHtml(htmlDocument) {
  const source = String(htmlDocument || "");
  if (!source) {
    return 0;
  }
  const cellTags = source.match(/<t[dh]\b[^>]*>/gi) || [];
  let mergedCount = 0;
  for (const tag of cellTags) {
    const rowSpan = readHtmlSpanAttribute(tag, "rowspan");
    const colSpan = readHtmlSpanAttribute(tag, "colspan");
    if (rowSpan > 1 || colSpan > 1) {
      mergedCount += 1;
    }
  }
  return mergedCount;
}

function analyzeCompatibilityHtmlRichness(htmlDocument) {
  const safeHtml = String(htmlDocument || "");
  const colorStyleCount = (safeHtml.match(/color\s*:\s*[^;"]+/gi) || []).length;
  const highlightStyleCount = (safeHtml.match(/background-color\s*:\s*[^;"]+/gi) || []).length;
  const fontSizeCount = (safeHtml.match(/font-size\s*:\s*[^;"]+/gi) || []).length;
  const textStrokeCount = (safeHtml.match(/-webkit-text-stroke\s*:\s*[^;"]+/gi) || []).length;
  const textShadowCount = (safeHtml.match(/text-shadow\s*:\s*[^;"]+/gi) || []).length;
  const cellDimensionCount = (safeHtml.match(/(?:width|height)\s*:\s*\d+(?:\.\d+)?px/gi) || []).length;
  const tableBorderStyleCount = (safeHtml.match(/border-(?:left|right|top|bottom)\s*:\s*[^;"]+/gi) || []).length;
  const centerAlignCount = (safeHtml.match(/(?:text-align\s*:\s*center|vertical-align\s*:\s*middle)/gi) || []).length;
  const tableHeaderCount = countHtmlTagOccurrences(safeHtml, "th");
  return {
    tableCount: countHtmlTagOccurrences(safeHtml, "table"),
    strongCount: countHtmlTagOccurrences(safeHtml, "strong") + countHtmlTagOccurrences(safeHtml, "b"),
    paragraphCount: countHtmlTagOccurrences(safeHtml, "p") + countHtmlTagOccurrences(safeHtml, "li") + countHtmlTagOccurrences(safeHtml, "tr"),
    mergedCellCount: countMergedCellsInHtml(safeHtml),
    colorStyleCount,
    highlightStyleCount,
    fontSizeCount,
    tableHeaderCount,
    textStrokeCount,
    textShadowCount,
    cellDimensionCount,
    tableBorderStyleCount,
    centerAlignCount
  };
}

function normalizeHtmlTextForComparison(value) {
  return String(value || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/\s+/g, "");
}

function shouldPreferMarkdownCompatibilityHtml(htmlDocument, sanitizedMarkdownText, markdownArtifact = {}) {
  const safeHtml = String(htmlDocument || "").trim();
  const safeMarkdown = String(sanitizedMarkdownText || "").trim();
  if (!safeMarkdown) {
    return false;
  }
  if (!safeHtml) {
    return true;
  }
  const markdownTextDensity = safeMarkdown.replace(/\s+/g, "").length;
  const htmlTextDensity = normalizeHtmlTextForComparison(safeHtml).length;
  const richness = analyzeCompatibilityHtmlRichness(safeHtml);
  const htmlTableCount = richness.tableCount;
  const htmlParagraphCount = richness.paragraphCount;
  const richStructureDetected = htmlTableCount > 0
    || richness.strongCount >= 8
    || Number(richness.mergedCellCount || 0) > 0
    || Number(richness.tableHeaderCount || 0) > 0
    || Number(richness.colorStyleCount || 0) >= 6
    || Number(richness.highlightStyleCount || 0) >= 2
    || Number(richness.cellDimensionCount || 0) >= 8
    || Number(richness.tableBorderStyleCount || 0) >= 8
    || Number(richness.centerAlignCount || 0) >= 4;
  const markdownTableLineCount = Number(markdownArtifact.tableLineCount || 0);
  const noisyTableLineCount = Number(markdownArtifact.noisyTableLineCount || 0);
  const markdownNoisyRatio = markdownTableLineCount > 0
    ? clampNumber(noisyTableLineCount / markdownTableLineCount, 0, 1)
    : 0;
  if (!richStructureDetected && markdownTextDensity >= 160 && htmlTextDensity < markdownTextDensity * 0.42) {
    return true;
  }
  if (markdownTableLineCount >= 6 && htmlTableCount === 0) {
    return true;
  }
  if (!richStructureDetected && htmlParagraphCount <= 3 && markdownTextDensity >= 260) {
    return true;
  }
  if (markdownNoisyRatio >= 0.42 && htmlTableCount === 0) {
    return true;
  }
  if (Number(markdownArtifact.strikeRatio || 0) >= 0.24 && htmlTableCount === 0) {
    return true;
  }
  return false;
}

async function convertHwpToPdfWithHwpJs(sourcePath, outputPdfPath, options = {}) {
  try {
    const hwpjs = ensureHwpJsModule();
    const raw = fs.readFileSync(sourcePath);
    const sourceSignature = String(options.sourceSignature || buildHwpSelectionSourceSignature(sourcePath));
    const forcedPageTarget = Math.max(
      0,
      Math.floor(
        Number(
          options.forcedPageTarget
          || resolveForcedPageTarget({ sourcePath, sourceExt: ".hwp", sourceSignature })
          || 0
        )
      )
    );
    const variants = readHwpDocumentVariants(hwpjs, raw);
    const markdownRaw = variants.markdown;
    const markdownArtifact = analyzeHwpJsMarkdownArtifacts(hwpjs, raw, markdownRaw);
    const sanitizedMarkdown = sanitizeHwpCompatibilityMarkdown(markdownRaw);
    const htmlRaw = variants.html;
    const jsonResult = variants.json;
    const pageModelResult = buildHwpPageModelCompatibilityHtmlFromJson(
      jsonResult,
      path.basename(sourcePath),
      {
        sourceExt: ".hwp",
        targetPageCount: forcedPageTarget,
        fontFaceCss: buildBundledFontFaceCss()
      }
    );
    const structuredHtmlFromJson = buildStructuredCompatibilityHtmlFromHwpJson(
      jsonResult,
      path.basename(sourcePath),
      { sourceExt: ".hwp" }
    );
    const normalizedHtml = normalizeHwpHtmlDocument(htmlRaw, path.basename(sourcePath));
    let html = normalizedHtml;
    let selectedHtmlSource = html ? "hwpjs-html" : "";
    let selectedRichness = analyzeCompatibilityHtmlRichness(html);
    let selectedPageModelDiagnostics = null;
    if (pageModelResult?.html) {
      const pageModelDiagnostics = pageModelResult.diagnostics || {};
      const pageModelAllowed =
        Number(forcedPageTarget || 0) <= 0
        || (
          Number(pageModelDiagnostics.pageCount || 0) === Number(forcedPageTarget || 0)
          && pageModelDiagnostics.anchorMatched !== false
        );
      if (pageModelAllowed) {
        html = pageModelResult.html;
        selectedHtmlSource = "hwpjs-page-model";
        selectedRichness = analyzeCompatibilityHtmlRichness(html);
        selectedPageModelDiagnostics = pageModelDiagnostics;
      }
    }
    if (structuredHtmlFromJson) {
      const structuredRichness = analyzeCompatibilityHtmlRichness(structuredHtmlFromJson);
      const currentRichness = analyzeCompatibilityHtmlRichness(html);
      const preferStructured =
        selectedHtmlSource !== "hwpjs-page-model"
        && (
          !html
          || structuredRichness.mergedCellCount > currentRichness.mergedCellCount
          || structuredRichness.tableCount > currentRichness.tableCount
          || structuredRichness.strongCount > currentRichness.strongCount + 4
          || structuredRichness.tableHeaderCount > currentRichness.tableHeaderCount
          || structuredRichness.colorStyleCount > currentRichness.colorStyleCount + 8
          || structuredRichness.highlightStyleCount > currentRichness.highlightStyleCount
          || structuredRichness.fontSizeCount > currentRichness.fontSizeCount + 12
          || structuredRichness.cellDimensionCount > currentRichness.cellDimensionCount + 8
          || structuredRichness.tableBorderStyleCount > currentRichness.tableBorderStyleCount + 8
          || structuredRichness.centerAlignCount > currentRichness.centerAlignCount + 4
          || structuredRichness.textStrokeCount > currentRichness.textStrokeCount
          || structuredRichness.textShadowCount > currentRichness.textShadowCount
        );
      if (preferStructured) {
        html = structuredHtmlFromJson;
        selectedHtmlSource = "hwpjs-json-structured";
        selectedRichness = structuredRichness;
        selectedPageModelDiagnostics = null;
      }
    }
    if (markdownArtifact.inspected && (markdownArtifact.strikeRatio >= 0.18 || markdownArtifact.strikeSegments >= 8)) {
      html = html
        .replace(/<\s*(del|s)\b[^>]*>/gi, "")
        .replace(/<\s*\/\s*(del|s)\s*>/gi, "")
        .replace(/line-through/gi, "none");
    }
    selectedRichness = analyzeCompatibilityHtmlRichness(html);
    const preferMarkdownHtml = shouldPreferMarkdownCompatibilityHtml(
      html,
      sanitizedMarkdown.markdown,
      markdownArtifact
    );
    const keepStructuredMergeShape = selectedRichness.tableCount > 0 && selectedRichness.mergedCellCount > 0;
    const keepPageModelLayout = selectedHtmlSource === "hwpjs-page-model";
    if ((!html || (preferMarkdownHtml && !keepStructuredMergeShape && !keepPageModelLayout)) && sanitizedMarkdown.markdown) {
      html = buildHwpCompatibilityHtmlFromMarkdown(sanitizedMarkdown.markdown, path.basename(sourcePath), {
        sourceExt: ".hwp"
      });
      selectedHtmlSource = "hwpjs-markdown";
      selectedRichness = analyzeCompatibilityHtmlRichness(html);
      selectedPageModelDiagnostics = null;
    }
    if (!html) {
      return { ok: false, reason: "hwpjs_html_empty", parserFailures: variants.failures };
    }
    const basePrintOptions = {
      printBackground: true,
      pageSize: "A4",
      preferCSSPageSize: true
    };
    const calibratedResult = await calibratePdfByPageTarget({
      renderPdf: (printToPdfOptions) => htmlToPdfBuffer(html, printToPdfOptions),
      basePrintOptions,
      targetPageCount: forcedPageTarget,
      maxAttempts: forcedPageTarget > 0 ? 8 : 1
    });
    const pdfBuffer =
      calibratedResult?.buffer && Buffer.isBuffer(calibratedResult.buffer) && calibratedResult.buffer.length > 0
        ? calibratedResult.buffer
        : await htmlToPdfBuffer(html, basePrintOptions);
    await fsp.writeFile(outputPdfPath, pdfBuffer);
    return {
      ok: true,
      markdownArtifact,
      parserFailures: variants.failures,
      selectedHtmlSource,
      pageModelUsed: selectedHtmlSource === "hwpjs-page-model",
      hwpPageModel: selectedPageModelDiagnostics,
      mergedCellCount: Number(selectedRichness.mergedCellCount || 0),
      tableCount: Number(selectedRichness.tableCount || 0),
      strongCount: Number(selectedRichness.strongCount || 0),
      colorStyleCount: Number(selectedRichness.colorStyleCount || 0),
      highlightStyleCount: Number(selectedRichness.highlightStyleCount || 0),
      fontSizeCount: Number(selectedRichness.fontSizeCount || 0),
      tableHeaderCount: Number(selectedRichness.tableHeaderCount || 0),
      textStrokeCount: Number(selectedRichness.textStrokeCount || 0),
      textShadowCount: Number(selectedRichness.textShadowCount || 0),
      cellDimensionCount: Number(selectedRichness.cellDimensionCount || 0),
      tableBorderStyleCount: Number(selectedRichness.tableBorderStyleCount || 0),
      centerAlignCount: Number(selectedRichness.centerAlignCount || 0),
      pageCount: Number(calibratedResult?.pageCount || 0),
      targetPageCount: forcedPageTarget,
      pageTargetMatched: forcedPageTarget > 0
        ? Number(calibratedResult?.pageCount || 0) === forcedPageTarget
        : true,
      calibrated: Boolean(calibratedResult?.calibrated),
      calibrationAttempts: Number(calibratedResult?.attempts || 1),
      calibrationScale: Number(calibratedResult?.finalScale || 1)
    };
  } catch (error) {
    return { ok: false, reason: error?.message || "hwpjs_convert_failed" };
  }
}

async function convertHwpToPdfWithRhwp(sourcePath, outputPdfPath) {
  try {
    const raw = await fsp.readFile(sourcePath);
    const result = await renderHwpWithRhwp(BrowserWindow, raw, path.basename(sourcePath), buildBundledFontFaceCss());
    await fsp.writeFile(outputPdfPath, result.pdfBuffer);
    return { ok: true, pageCount: result.pageCount, pageGeometry:result.pageGeometry, selectedHtmlSource: result.selectedHtmlSource };
  } catch (error) {
    return { ok: false, reason: error?.message || "rhwp_convert_failed" };
  }
}

async function convertHwpxToPdfWithReader(sourcePath, outputPdfPath, options = {}) {
  try {
    const readHwpxDocument = ensureHwpxReader();
    const raw = fs.readFileSync(sourcePath);
    const document = normalizeHwpxDocumentStyles(readHwpxDocument(new Uint8Array(raw)));
    const sourceSignature = String(options.sourceSignature || buildHwpSelectionSourceSignature(sourcePath));
    const forcedPageTarget = Math.max(
      0,
      Math.floor(
        Number(
          options.forcedPageTarget
          || resolveForcedPageTarget({ sourcePath, sourceExt: ".hwpx", sourceSignature })
          || 0
        )
      )
    );
    const basePrintOptions = {
      printBackground: true,
      pageSize: "A4",
      preferCSSPageSize: true
    };
    const structuredHtml = buildStructuredCompatibilityHtmlFromDocument(document, path.basename(sourcePath), {
      sourceExt: ".hwpx"
    });
    if (structuredHtml) {
      const structuredRichness = analyzeCompatibilityHtmlRichness(structuredHtml);
      const calibratedResult = await calibratePdfByPageTarget({
        renderPdf: (printToPdfOptions) => htmlToPdfBuffer(structuredHtml, printToPdfOptions),
        basePrintOptions,
        targetPageCount: forcedPageTarget,
        maxAttempts: forcedPageTarget > 0 ? 8 : 1
      });
      const pdfBuffer =
        calibratedResult?.buffer && Buffer.isBuffer(calibratedResult.buffer) && calibratedResult.buffer.length > 0
          ? calibratedResult.buffer
          : await htmlToPdfBuffer(structuredHtml, basePrintOptions);
      await fsp.writeFile(outputPdfPath, pdfBuffer);
      return {
        ok: true,
        pageCount: Number(calibratedResult?.pageCount || 0),
        targetPageCount: forcedPageTarget,
        pageTargetMatched: forcedPageTarget > 0
          ? Number(calibratedResult?.pageCount || 0) === forcedPageTarget
          : true,
        calibrated: Boolean(calibratedResult?.calibrated),
        calibrationAttempts: Number(calibratedResult?.attempts || 1),
        calibrationScale: Number(calibratedResult?.finalScale || 1),
        mergedCellCount: Number(structuredRichness.mergedCellCount || 0),
        tableCount: Number(structuredRichness.tableCount || 0),
        strongCount: Number(structuredRichness.strongCount || 0),
        colorStyleCount: Number(structuredRichness.colorStyleCount || 0),
        highlightStyleCount: Number(structuredRichness.highlightStyleCount || 0),
        fontSizeCount: Number(structuredRichness.fontSizeCount || 0),
        tableHeaderCount: Number(structuredRichness.tableHeaderCount || 0),
        textStrokeCount: Number(structuredRichness.textStrokeCount || 0),
        textShadowCount: Number(structuredRichness.textShadowCount || 0),
        cellDimensionCount: Number(structuredRichness.cellDimensionCount || 0),
        tableBorderStyleCount: Number(structuredRichness.tableBorderStyleCount || 0),
        centerAlignCount: Number(structuredRichness.centerAlignCount || 0)
      };
    }
    const extractedText = normalizeText(extractHwpFamilyText(sourcePath));
    if (!extractedText) {
      return { ok: false, reason: "hwpx_reader_empty" };
    }
    const html = buildFallbackHtml({
      sourcePath,
      sourceExt: ".hwpx",
      contentText: extractedText,
      warningMessage: ""
    });
    const calibratedResult = await calibratePdfByPageTarget({
      renderPdf: (printToPdfOptions) => htmlToPdfBuffer(html, printToPdfOptions),
      basePrintOptions,
      targetPageCount: forcedPageTarget,
      maxAttempts: forcedPageTarget > 0 ? 6 : 1
    });
    const pdfBuffer =
      calibratedResult?.buffer && Buffer.isBuffer(calibratedResult.buffer) && calibratedResult.buffer.length > 0
        ? calibratedResult.buffer
        : await htmlToPdfBuffer(html, basePrintOptions);
    await fsp.writeFile(outputPdfPath, pdfBuffer);
    return {
      ok: true,
      pageCount: Number(calibratedResult?.pageCount || 0),
      targetPageCount: forcedPageTarget,
      pageTargetMatched: forcedPageTarget > 0
        ? Number(calibratedResult?.pageCount || 0) === forcedPageTarget
        : true,
      calibrated: Boolean(calibratedResult?.calibrated),
      calibrationAttempts: Number(calibratedResult?.attempts || 1),
      calibrationScale: Number(calibratedResult?.finalScale || 1)
    };
  } catch (error) {
    return { ok: false, reason: error?.message || "hwpx_reader_failed" };
  }
}

function buildFallbackHtml({ sourcePath, sourceExt, contentText, warningMessage }) {
  const cleanedText = sanitizeFallbackTextContent(contentText || "");
  const lowerExt = String(sourceExt || "").toLowerCase();
  if (lowerExt === ".hwp" || lowerExt === ".hwpx") {
    return buildHwpCompatibilityHtmlFromMarkdown(cleanedText, path.basename(sourcePath), {
      sourceExt: lowerExt,
      warningMessage
    });
  }
  const safeTitle = escapeHtml(path.basename(sourcePath));
  const safeExt = escapeHtml(sourceExt.replace(".", "").toUpperCase());
  const safeBody = escapeHtml(cleanedText || "(문서 내용을 추출하지 못했습니다.)");

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
    pre { white-space: pre-wrap; word-break: break-word; line-height: 1.45; font-size: 12px; margin: 0; }
  </style>
</head>
<body>
  <div class="head">
    <p class="title">${safeTitle}</p>
    <div class="meta">원본 형식: ${safeExt} / lookup 임시 변환 문서</div>
  </div>
  <pre>${safeBody}</pre>
</body>
</html>`;
}

function sanitizeFallbackTextContent(rawText) {
  let text = normalizeSpecialSymbols(String(rawText || ""));
  if (!text.trim()) {
    return "";
  }
  text = text
    .replace(/\uFFFD+/g, " ")
    .replace(/�+/g, " ")
    .replace(/\?{3,}/g, " ")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ")
    .replace(/\\[rnt]/gi, " ")
    .replace(/`{1,3}/g, "")
    .replace(/~~([^~]+)~~/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*\n]+)\*/g, "$1")
    .replace(/^\s{0,3}#{1,6}\s*/gm, "")
    .replace(/^\s*>\s?/gm, "")
    .replace(/^\s*[-*•]\s+(?=[#*_~`|])/gm, "")
    .replace(/^\s*\|[-:\s|]+\|\s*$/gm, "")
    .replace(/[|]{3,}/g, " | ")
    .replace(/(?:\|\s*){5,}/g, "| ")
    .replace(/^\s*(?:[|｜]\s*){2,}$/gm, "")
    .replace(/^[\-\s|]{4,}$/gm, "")
    .replace(/[^\S\r\n]{3,}/g, "  ")
    .replace(/\n{4,}/g, "\n\n")
    .replace(/[ \t]+\n/g, "\n");
  return normalizeText(text);
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

function normalizeSpreadsheetMergeRange(rawRange) {
  const startRowRaw = Number(rawRange?.s?.r);
  const endRowRaw = Number(rawRange?.e?.r);
  const startColRaw = Number(rawRange?.s?.c);
  const endColRaw = Number(rawRange?.e?.c);
  if (!Number.isFinite(startRowRaw) || !Number.isFinite(endRowRaw) || !Number.isFinite(startColRaw) || !Number.isFinite(endColRaw)) {
    return null;
  }
  const startRow = Math.max(0, Math.floor(Math.min(startRowRaw, endRowRaw)));
  const endRow = Math.max(0, Math.floor(Math.max(startRowRaw, endRowRaw)));
  const startCol = Math.max(0, Math.floor(Math.min(startColRaw, endColRaw)));
  const endCol = Math.max(0, Math.floor(Math.max(startColRaw, endColRaw)));
  return {
    startRow,
    endRow,
    startCol,
    endCol
  };
}

function buildSpreadsheetMergeMaps(mergeRanges, bounds) {
  const ownerMap = new Map();
  const coveredMap = new Set();
  const ownerRows = new Set();
  const coveredRows = new Set();
  const startToEndRow = new Map();
  const safeBounds = bounds || {};
  const minRow = Math.max(0, Number(safeBounds.startRow || 0));
  const maxRow = Math.max(minRow, Number(safeBounds.endRow || minRow));
  const minCol = Math.max(0, Number(safeBounds.startCol || 0));
  const maxCol = Math.max(minCol, Number(safeBounds.endCol || minCol));

  for (const mergeRange of Array.isArray(mergeRanges) ? mergeRanges : []) {
    const normalized = normalizeSpreadsheetMergeRange(mergeRange);
    if (!normalized) {
      continue;
    }
    if (normalized.endRow < minRow || normalized.startRow > maxRow || normalized.endCol < minCol || normalized.startCol > maxCol) {
      continue;
    }
    const startRow = Math.max(minRow, normalized.startRow);
    const endRow = Math.min(maxRow, normalized.endRow);
    const startCol = Math.max(minCol, normalized.startCol);
    const endCol = Math.min(maxCol, normalized.endCol);
    if (startRow > endRow || startCol > endCol) {
      continue;
    }
    const rowSpan = endRow - startRow + 1;
    const colSpan = endCol - startCol + 1;
    if (rowSpan <= 1 && colSpan <= 1) {
      continue;
    }
    const ownerKey = `${startRow}:${startCol}`;
    const currentOwner = ownerMap.get(ownerKey);
    const nextOwner = {
      startRow,
      endRow,
      startCol,
      endCol,
      rowSpan,
      colSpan
    };
    if (!currentOwner || currentOwner.rowSpan * currentOwner.colSpan < rowSpan * colSpan) {
      ownerMap.set(ownerKey, nextOwner);
    }
    ownerRows.add(startRow);
    startToEndRow.set(startRow, Math.max(Number(startToEndRow.get(startRow) || startRow), endRow));
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        if (row === startRow && col === startCol) {
          continue;
        }
        coveredMap.add(`${row}:${col}`);
        coveredRows.add(row);
      }
    }
  }

  return {
    ownerMap,
    coveredMap,
    ownerRows,
    coveredRows,
    startToEndRow
  };
}

function buildMergeAwareSpreadsheetPages(rowEntries, chunkRows) {
  const pages = [];
  const rows = Array.isArray(rowEntries) ? rowEntries : [];
  if (!rows.length) {
    return pages;
  }
  const safeChunkRows = Math.max(1, Math.floor(Number(chunkRows) || 42));
  let cursor = 0;
  while (cursor < rows.length) {
    const tentativeEnd = Math.min(rows.length - 1, cursor + safeChunkRows - 1);
    const boundaryRow = Number(rows[tentativeEnd]?.rowIndex || 0);
    let crossingIndex = -1;
    for (let index = cursor; index <= tentativeEnd; index += 1) {
      const mergeEndRow = Number(rows[index]?.mergeEndRow ?? rows[index]?.rowIndex ?? 0);
      if (mergeEndRow > boundaryRow) {
        crossingIndex = index;
        break;
      }
    }

    let pageEnd = tentativeEnd;
    if (crossingIndex > cursor) {
      pageEnd = crossingIndex - 1;
    } else if (crossingIndex === cursor) {
      const requiredEndRow = Number(rows[cursor]?.mergeEndRow ?? boundaryRow);
      for (let index = tentativeEnd + 1; index < rows.length; index += 1) {
        const rowIndex = Number(rows[index]?.rowIndex ?? requiredEndRow + 1);
        if (rowIndex > requiredEndRow) {
          break;
        }
        pageEnd = index;
      }
    }

    if (pageEnd < cursor) {
      pageEnd = tentativeEnd;
    }
    pages.push(rows.slice(cursor, pageEnd + 1));
    cursor = pageEnd + 1;
  }
  return pages;
}

function normalizeSpreadsheetHexColor(color) {
  const raw = String(color?.rgb || "").trim();
  if (!raw) {
    return "";
  }
  const compact = raw.replace(/^FF/i, "");
  if (!/^[0-9a-fA-F]{6}$/.test(compact)) {
    return "";
  }
  return `#${compact.toLowerCase()}`;
}

function normalizeSpreadsheetHorizontalAlign(value) {
  switch (String(value || "").toLowerCase()) {
    case "center":
    case "centercontinuous":
      return "center";
    case "right":
      return "right";
    case "justify":
      return "justify";
    case "left":
      return "left";
    default:
      return "";
  }
}

function normalizeSpreadsheetVerticalAlign(value) {
  switch (String(value || "").toLowerCase()) {
    case "center":
      return "middle";
    case "bottom":
      return "bottom";
    case "top":
      return "top";
    default:
      return "";
  }
}

function readSpreadsheetColumnWidth(colDef = {}) {
  const direct = Number(colDef?.wpx || colDef?.width || 0);
  if (Number.isFinite(direct) && direct > 0) {
    return direct;
  }
  const wch = Number(colDef?.wch || 0);
  if (Number.isFinite(wch) && wch > 0) {
    return wch * 8;
  }
  return 0;
}

function buildSpreadsheetCellStyle(cell = {}, options = {}) {
  const style = {};
  const alignment = (cell?.s && typeof cell.s === "object" && cell.s.alignment) || {};
  const font = (cell?.s && typeof cell.s === "object" && cell.s.font) || {};
  const fill = (cell?.s && typeof cell.s === "object" && cell.s.fill) || {};
  const horizontal = normalizeSpreadsheetHorizontalAlign(alignment.horizontal);
  const vertical = normalizeSpreadsheetVerticalAlign(alignment.vertical);
  const isHeader = options.isHeader === true;
  const isIndexColumn = options.isIndexColumn === true;
  if (horizontal) {
    style["text-align"] = horizontal;
  } else if (isHeader || isIndexColumn) {
    style["text-align"] = "center";
  }
  if (vertical) {
    style["vertical-align"] = vertical;
  } else if (isHeader || isIndexColumn) {
    style["vertical-align"] = "middle";
  }
  if (font.bold === true) {
    style["font-weight"] = "700";
  }
  if (font.italic === true) {
    style["font-style"] = "italic";
  }
  if (Number.isFinite(Number(font.sz)) && Number(font.sz) > 0) {
    style["font-size"] = `${Math.max(8, Math.min(30, Number(font.sz)))}px`;
  }
  const fontColor = normalizeSpreadsheetHexColor(font.color);
  if (fontColor) {
    style.color = fontColor;
  }
  const fillColor = normalizeSpreadsheetHexColor(fill.fgColor || fill.bgColor || null);
  if (fillColor) {
    style["background-color"] = fillColor;
  } else if (isHeader) {
    style["background-color"] = "#f3f6fb";
  }
  return Object.entries(style)
    .filter(([, value]) => String(value || "").trim() !== "")
    .map(([key, value]) => `${key}:${value}`)
    .join(";");
}

function buildSpreadsheetFallbackHtml({ sourcePath, sourceExt, warningMessage }) {
  const prepared = buildSpreadsheetCompatibility({ sourcePath });
  const visual = detectSpreadsheetVisualAssets(sourcePath);
  const warnings = [prepared.warningMessage];
  if (visual.hasVisual) warnings.push("표 보기 · 그림/차트 제외");
  else if (warningMessage) warnings.push("표 보기");
  return { ...prepared, warningMessage: warnings.filter(Boolean).join(" · ") };
}

async function renderSpreadsheetHtmlPdf(html, printOptions = {}) {
  // Large spreadsheets exceed Chromium's data URL length limit. Use an owned
  // temporary file and remove it after the disposable rendering window closes.
  const htmlPath = path.join(ensureConvertedDir(), `xlsx-render-${crypto.randomUUID()}.html`);
  try {
    await fsp.writeFile(htmlPath, html, "utf8");
    const rendered = await renderHtmlToPdfBuffer({
      sourcePath: htmlPath,
      printOptions: { ...printOptions, sourceType: "spreadsheet-html" }
    });
    return rendered.pdfBuffer;
  } finally {
    await fsp.unlink(htmlPath).catch(() => {});
  }
}

function buildAccdbFallbackHtml({ sourcePath, warningMessage = "", tables = [] }) {
  const sections = [];
  const sheetMap = [];
  const chunkRows = 42;
  let nextPageNumber = 1;
  for (const table of tables) {
    const tableName = String(table?.name || "").trim();
    const columns = Array.isArray(table?.columns)
      ? table.columns.map((name) => String(name || "").trim()).filter(Boolean)
      : [];
    const rows = Array.isArray(table?.rows) ? table.rows : [];
    if (!tableName || columns.length === 0) {
      continue;
    }
    const sanitizedRows = rows
      .slice(0, ACCDB_PREVIEW_MAX_ROWS)
      .map((row) => (Array.isArray(row) ? row : []));
    if (sanitizedRows.length === 0) {
      continue;
    }

    const headerHtml = `<tr>${columns.map((name) => `<th>${escapeHtml(name)}</th>`).join("")}</tr>`;
    const rowHtmlList = sanitizedRows.map((row) => {
      const cells = columns.map((_column, index) => {
        const rawValue = index < row.length ? row[index] : "";
        const value = String(rawValue ?? "");
        return `<td>${escapeHtml(value)}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    });
    const pageCountForTable = Math.max(1, Math.ceil(rowHtmlList.length / chunkRows));
    const startPage = nextPageNumber;
    for (let chunkIndex = 0; chunkIndex < pageCountForTable; chunkIndex += 1) {
      const chunkStart = chunkIndex * chunkRows;
      const chunkEnd = chunkStart + chunkRows;
      const rowsChunk = rowHtmlList.slice(chunkStart, chunkEnd).join("");
      const chunkTitle =
        pageCountForTable > 1
          ? `${escapeHtml(tableName)} (${chunkIndex + 1}/${pageCountForTable})`
          : escapeHtml(tableName);
      sections.push(
        `<section class="sheet"><h2>${chunkTitle}</h2><table><thead>${headerHtml}</thead><tbody>${rowsChunk}</tbody></table></section>`
      );
      nextPageNumber += 1;
    }
    sheetMap.push({
      sheetName: tableName,
      startPage,
      endPage: nextPageNumber - 1
    });
  }

  if (sections.length === 0) {
    return {
      html: "",
      sheetMap: []
    };
  }

  const safeTitle = escapeHtml(path.basename(sourcePath));
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
    th, td { border: 1px solid #d6deea; padding: 4px 5px; vertical-align: top; word-break: break-word; white-space: pre-wrap; }
    th { background: #f8fafc; }
  </style>
</head>
<body>
  <div class="head">
    <p class="title">${safeTitle}</p>
    <div class="meta">원본 형식: ACCDB / lookup 데이터 보기</div>
  </div>
  ${warning}
  ${sections.join("\n")}
</body>
</html>`,
    sheetMap
  };
}

function detectNotionHtmlSource(html, sourcePath = "") {
  const sample = `${String(sourcePath || "")}\n${String(html || "").slice(0, 250000)}`.toLowerCase();
  return (
    sample.includes("notion")
    || sample.includes("data-block-id")
    || sample.includes("notion-page-content")
    || sample.includes("notion-scroller")
    || sample.includes("notion-callout")
    || sample.includes("notion-collection")
  );
}

function getWebPrintSourceType({ html = "", sourcePath = "", sourceType = "" } = {}) {
  const normalized = String(sourceType || "").trim().toLowerCase();
  if (normalized) {
    return normalized;
  }
  return detectNotionHtmlSource(html, sourcePath) ? "notion-web" : "web-html";
}

function stripHtmlPdfCustomOptions(printOptions = {}) {
  const {
    sourcePath: _sourcePath,
    sourceType: _sourceType,
    normalizeScrollableContent: _normalizeScrollableContent,
    allowRasterFallback: _allowRasterFallback,
    forceRasterFallback: _forceRasterFallback,
    preparationTimeoutMs: _preparationTimeoutMs,
    preserveScreenMedia: _preserveScreenMedia,
    ...safePrintOptions
  } = printOptions && typeof printOptions === "object" ? printOptions : {};
  return safePrintOptions;
}

async function waitForHtmlPdfReadiness(webContents) {
  return waitForWebPageReadiness(webContents);
}

async function inspectWebPrintLayout(webContents, sourceType = "web-html") {
  return inspectWebPage(webContents, { sourceType });
}

async function rasterizeFullWebPageToPdfBuffer(tempWindow, diagnostics = {}) {
  const webContents = tempWindow.webContents;
  const pageWidthPt = 595.28;
  const pageHeightPt = 841.89;
  const pdfDoc = await PDFDocument.create();
  let metrics = await inspectWebPrintLayout(webContents, diagnostics?.sourceType || "web-html");
  const viewportWidth = Math.max(760, Math.min(1600, Math.ceil(Number(metrics.totalWidth || metrics.viewportWidth || 1040))));
  const sliceHeight = Math.max(980, Math.min(1900, Math.floor(viewportWidth * (pageHeightPt / pageWidthPt))));
  try {
    tempWindow.setContentSize(viewportWidth, sliceHeight);
  } catch (_error) {
    // Keep current viewport if the runtime rejects a large hidden-window size.
  }
  await new Promise((resolve) => setTimeout(resolve, 120));
  metrics = await inspectWebPrintLayout(webContents, diagnostics?.sourceType || "web-html");
  const totalHeight = Math.max(1, Math.ceil(Number(metrics.totalHeight || metrics.viewportHeight || sliceHeight)));
  const pageCount = Math.max(1, Math.ceil(totalHeight / sliceHeight));
  if (pageCount > 200) throw new Error("웹페이지가 너무 길어 전체 화면 저장을 완료하지 못했습니다.");
  let meaningfulCaptures = 0;
  for (let index = 0; index < pageCount; index += 1) {
    const top = index * sliceHeight;
    const actualTop = await webContents.executeJavaScript(`(() => {
      window.scrollTo(0, ${Math.floor(top)});
      return Number(window.scrollY || document.scrollingElement?.scrollTop || 0);
    })()`);
    await new Promise((resolve) => setTimeout(resolve, 80));
    const remaining = Math.max(1, totalHeight - top);
    const captureHeight = Math.min(sliceHeight, remaining);
    let image = await webContents.capturePage({
      x: 0,
      y: Math.max(0, Math.round(top - actualTop)),
      width: viewportWidth,
      height: captureHeight
    }, { stayHidden: true, stayAwake: true });
    let captureCheck = inspectCapturedImage(image);
    if (!captureCheck.valid && captureCheck.reason !== "uniform-image") {
      await new Promise((resolve) => setTimeout(resolve, 160));
      image = await webContents.capturePage({ x: 0, y: Math.max(0, Math.round(top - actualTop)), width: viewportWidth, height: captureHeight }, { stayHidden: true, stayAwake: true });
      captureCheck = inspectCapturedImage(image);
    }
    if (!captureCheck.valid && captureCheck.reason !== "uniform-image") throw new Error("웹페이지 화면이 비어 있어 PDF 저장을 중단했습니다.");
    if (captureCheck.valid) meaningfulCaptures += 1;
    const embedded = await pdfDoc.embedPng(image.toPNG());
    const page = pdfDoc.addPage([pageWidthPt, pageHeightPt]);
    const drawHeight = Math.min(pageHeightPt, captureHeight * (pageWidthPt / viewportWidth));
    page.drawImage(embedded, {
      x: 0,
      y: pageHeightPt - drawHeight,
      width: pageWidthPt,
      height: drawHeight
    });
  }
  if (!meaningfulCaptures) throw new Error("웹페이지 화면이 비어 있어 PDF 저장을 중단했습니다.");
  return Buffer.from(await pdfDoc.save());
}

function shouldUseWebRasterFallback(diagnostics, printOptions = {}) {
  if (printOptions.forceRasterFallback === true) {
    return true;
  }
  if (printOptions.allowRasterFallback === false) {
    return false;
  }
  const after = diagnostics?.after || {};
  return Number(after.maxOverflowY || 0) > 240;
}

async function loadConversionPage(tempWindow, url) {
  let timer;
  try {
    await Promise.race([
      tempWindow.loadURL(url),
      new Promise((_resolve, reject) => {
        timer = setTimeout(() => {
          const error = new Error("사이트 응답 지연 · 다시 시도");
          error.code = "WEB_LOAD_TIMEOUT";
          reject(error);
        }, 20000);
      })
    ]);
  } finally { clearTimeout(timer); }
}

async function renderHtmlToPdfBuffer({ html = "", sourcePath = "", sourceUrl = "", printOptions = {} } = {}) {
  const safeOptions = printOptions && typeof printOptions === "object" ? printOptions : {};
  let normalizedSourceUrl = "";
  if (sourceUrl) {
    const parsed = new URL(sourceUrl);
    if (!["https:", "http:"].includes(parsed.protocol)) throw new Error("http 또는 https 주소를 입력해 주세요.");
    if (parsed.username || parsed.password) throw new Error("로그인 정보가 포함된 주소는 열 수 없습니다.");
    normalizedSourceUrl = parsed.toString();
  }
  const sourceType = getWebPrintSourceType({
    html,
    sourcePath,
    sourceType: safeOptions.sourceType
  });
  const shouldNormalize = safeOptions.normalizeScrollableContent === true || sourceType === "notion-web" || Boolean(normalizedSourceUrl);
  const tempWindow = new BrowserWindow({
    show: false,
    width: shouldNormalize ? 1280 : 1040,
    height: 1440,
    autoHideMenuBar: true,
    webPreferences: {
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
      backgroundThrottling: false,
      offscreen: shouldNormalize
    }
  });

  tempWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  let resetMedia = null;
  let inlineSourceDirectory = "";
  try {
    if (normalizedSourceUrl) {
      await loadConversionPage(tempWindow, normalizedSourceUrl);
      await waitForWebAccess(tempWindow.webContents);
    } else if (sourcePath) {
      await tempWindow.loadURL(pathToFileURL(sourcePath).toString());
    } else {
      // Long tables and image-rich documents exceed Chromium's data URL limit.
      inlineSourceDirectory = await fsp.mkdtemp(path.join(app.getPath("temp"), "lookup-html-"));
      const inlineSourcePath = path.join(inlineSourceDirectory, "document.html");
      await fsp.writeFile(inlineSourcePath, html, "utf8");
      await tempWindow.loadURL(pathToFileURL(inlineSourcePath).toString());
    }
    let webPageDiagnostics = null;
    if (shouldNormalize) {
      if (safeOptions.preserveScreenMedia !== false) resetMedia = await useScreenMediaForPdf(tempWindow.webContents);
      webPageDiagnostics = await prepareWebPageForPdf(tempWindow.webContents, { sourceType, timeoutMs: safeOptions.preparationTimeoutMs });
      webPageDiagnostics.screenMedia = Boolean(resetMedia);
      const assessment = assessWebPageContent(webPageDiagnostics);
      webPageDiagnostics.contentAssessment = assessment;
      webPageDiagnostics.warningMessage = assessment.incompleteImages > 0
        ? `이미지 ${assessment.incompleteImages}개 미확인 · 다시 시도`
        : assessment.reachedBottom ? "" : "일부 내용 미확인 · 다시 시도";
      if (!assessment.hasContent) throw new Error("웹페이지 내용이 비어 있습니다. 원본 주소와 로딩 상태를 확인해 주세요.");
      if (normalizedSourceUrl) await waitForWebAccess(tempWindow.webContents, { timeoutMs: 1000 });
    } else await waitForHtmlPdfReadiness(tempWindow.webContents);
    const finalOptions = {
      printBackground: true,
      pageSize: "A4",
      preferCSSPageSize: true,
      ...stripHtmlPdfCustomOptions(safeOptions)
    };
    let pdfBuffer = await tempWindow.webContents.printToPDF(finalOptions);
    if (webPageDiagnostics && shouldUseWebRasterFallback(webPageDiagnostics, safeOptions)) {
      try {
        const rasterBuffer = await rasterizeFullWebPageToPdfBuffer(tempWindow, webPageDiagnostics.after);
        if (Buffer.isBuffer(rasterBuffer) && rasterBuffer.length >= MIN_VALID_PDF_BYTES) {
          webPageDiagnostics.rasterFallbackUsed = true;
          pdfBuffer = rasterBuffer;
        }
      } catch (error) {
        webPageDiagnostics.rasterFallbackError = error?.message || String(error);
        webPageDiagnostics.warningMessage = webPageDiagnostics.warningMessage || "일부 내용 미확인 · 다시 시도";
        if (safeOptions.forceRasterFallback === true) throw error;
      }
    }
    if (shouldNormalize) {
      const captureContent = await inspectPdfCaptureContent(pdfBuffer);
      webPageDiagnostics.pdfContent = captureContent;
      if (captureContent.placeholderOnly || (captureContent.checked && captureContent.pageCount > 0 && captureContent.blankPages.length === captureContent.pageCount)) {
        const error = new Error("웹페이지 내용이 PDF에 담기지 않아 저장을 중단했습니다. 다시 시도해 주세요.");
        error.code = "WEB_PDF_EMPTY";
        throw error;
      }
    }
    return {
      pdfBuffer,
      diagnostics: webPageDiagnostics || {
        sourceType,
        rasterFallbackUsed: false
      }
    };
  } finally {
    if (resetMedia) await resetMedia().catch(() => {});
    if (!tempWindow.isDestroyed()) {
      tempWindow.destroy();
    }
    if (inlineSourceDirectory) {
      await fsp.unlink(path.join(inlineSourceDirectory, "document.html")).catch(() => {});
      await fsp.rmdir(inlineSourceDirectory).catch(() => {});
    }
  }
}

async function htmlToPdfBuffer(html, printOptions = {}) {
  const result = await renderHtmlToPdfBuffer({ html, printOptions });
  return result.pdfBuffer;
}

async function htmlFileToPdfBuffer(sourcePath, printOptions = {}) {
  return renderHtmlToPdfBuffer({ sourcePath, printOptions });
}

async function convertWebHtmlFileToPdf(sourcePath) {
  const sourceExt = getFileExtension(sourcePath) || ".html";
  let previewHtml = "";
  try {
    previewHtml = (await fsp.readFile(sourcePath, "utf8")).slice(0, 250000);
  } catch (_error) {
    previewHtml = "";
  }
  const sourceType = getWebPrintSourceType({ html: previewHtml, sourcePath });
  const engine = sourceType === "notion-web" ? "notion-web" : "web-html";
  const pdfPath = createConvertedPdfPath(sourcePath, engine);
  const cached = readCachedPdf(pdfPath);
  if (cached) {
    const cachedValidation = await validateConvertedPdfPath(cached);
    if (cachedValidation.ok) {
      return {
        ok: true,
        engine,
        warningMessage: "",
        pdfPath: cached,
        fromCache: true,
        errorCode: "",
        pageCount: cachedValidation.pageCount,
        diagnostics: {
          sourceType,
          webPage: {
            sourceType,
            fromCache: true,
            rasterFallbackUsed: false
          }
        }
      };
    }
    try {
      await fsp.unlink(cached);
    } catch (_error) {
      // ignore stale cache cleanup failure
    }
  }

  try {
    const result = await htmlFileToPdfBuffer(sourcePath, {
      sourceType,
      normalizeScrollableContent: true,
      allowRasterFallback: true
    });
    await fsp.writeFile(pdfPath, result.pdfBuffer);
    const validation = await validateConvertedPdfPath(pdfPath);
    if (!validation.ok) {
      return {
        ok: false,
        engine,
        warningMessage: buildConversionWarning(sourceExt, validation.reason, "웹 페이지 PDF 변환 결과가 손상되었습니다."),
        errorCode: validation.errorCode || "INVALID_CONVERTED_PDF",
        diagnostics: {
          sourceType,
          webPage: result.diagnostics || null
        }
      };
    }
    return {
      ok: true,
      engine,
      warningMessage: result.diagnostics?.warningMessage || "",
      pdfPath,
      fromCache: false,
      errorCode: "",
      pageCount: validation.pageCount,
      diagnostics: {
        sourceType,
        webPage: result.diagnostics || null
      }
    };
  } catch (error) {
    return {
      ok: false,
      engine,
      warningMessage: buildConversionWarning(sourceExt, error?.message || String(error), "웹 페이지를 PDF로 변환하지 못했습니다."),
      errorCode: "CONVERT_FAILED",
      diagnostics: {
        sourceType,
        webPage: {
          sourceType,
          error: error?.message || String(error)
        }
      }
    };
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
  if (SPREADSHEET_EXTENSIONS.has(ext)) {
    const fromSheet = extractSpreadsheetText(filePath);
    if (fromSheet) {
      return fromSheet;
    }
    return extractOfficeAstText(filePath);
  }
  if (ext === ".accdb") {
    const dump = await runAccdbDump(filePath, { maxRows: 40, maxTables: 12, maxCols: 12 });
    if (!dump.ok || !Array.isArray(dump.tables)) {
      return "";
    }
    const blocks = [];
    for (const table of dump.tables) {
      const tableName = String(table?.name || "").trim();
      const rows = Array.isArray(table?.rows) ? table.rows : [];
      if (!tableName || rows.length === 0) {
        continue;
      }
      const lines = rows
        .slice(0, 20)
        .map((row) => (Array.isArray(row) ? row.map((cell) => String(cell ?? "")).join(" | ") : ""))
        .filter(Boolean);
      if (lines.length > 0) {
        blocks.push(`[${tableName}]\n${lines.join("\n")}`);
      }
    }
    return normalizeText(blocks.join("\n\n"));
  }
  if (WEB_DOCUMENT_EXTENSIONS.has(ext)) {
    try {
      return normalizeText(String(await fsp.readFile(filePath, "utf8"))
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/gi, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">"));
    } catch (_error) {
      return "";
    }
  }
  if (ext === ".ppt" || ext === ".pptx") {
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
  const spreadsheetSheetNames = SPREADSHEET_EXTENSIONS.has(ext) ? readSpreadsheetSheetNames(sourcePath) : [];
  const spreadsheetUnknownSheetMap = spreadsheetSheetNames.map((sheetName) => ({
    sheetName,
    startPage: null,
    endPage: null
  }));
  if (SPREADSHEET_EXTENSIONS.has(ext)) {
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
    const cachedValidation = await validateConvertedPdfPath(cachedOffice);
    if (cachedValidation.ok) {
      return {
        ok: true,
        engine: "office-com",
        warningMessage: "",
        pdfPath: cachedOffice,
        fromCache: true,
        errorCode: "",
        sheetMap: spreadsheetUnknownSheetMap,
        pageCount: cachedValidation.pageCount
      };
    }
    try {
      await fsp.unlink(cachedOffice);
    } catch (_error) {
      // ignore stale cache cleanup failure
    }
  }

  let comResult = { ok: false, stderr: "unsupported_office_ext" };
  if (ext === ".doc" || ext === ".docx") {
    comResult = await convertWithWordCom(sourcePath, officePdfPath);
  } else if (SPREADSHEET_EXTENSIONS.has(ext)) {
    comResult = await convertWithExcelCom(sourcePath, officePdfPath);
  } else if (ext === ".ppt" || ext === ".pptx") {
    comResult = await convertWithPowerPointCom(sourcePath, officePdfPath);
  }

  if (comResult.ok && fs.existsSync(officePdfPath)) {
    const comValidation = await validateConvertedPdfPath(officePdfPath);
    if (comValidation.ok) {
      return {
        ok: true,
        engine: "office-com",
        warningMessage: "",
        pdfPath: officePdfPath,
        fromCache: false,
        errorCode: "",
        sheetMap: spreadsheetUnknownSheetMap,
        pageCount: comValidation.pageCount
      };
    }
    try {
      await fsp.unlink(officePdfPath);
    } catch (_error) {
      // ignore invalid conversion cleanup failure
    }
    comResult = {
      ...comResult,
      ok: false,
      stderr: `office_invalid_pdf:${comValidation.reason}`,
      error: comValidation.errorCode || "INVALID_CONVERTED_PDF"
    };
  }

  const libreModeHint = findLibreOfficeExecutable().mode || "libreoffice-system";
  const librePdfPath = createConvertedPdfPath(sourcePath, libreModeHint);
  const cachedLibre = readCachedPdf(librePdfPath);
  if (cachedLibre) {
    const cachedValidation = await validateConvertedPdfPath(cachedLibre);
    if (cachedValidation.ok) {
      return {
        ok: true,
        engine: libreModeHint,
        warningMessage: "Office 기본 변환에 실패해 LibreOffice 호환 엔진으로 열었습니다.",
        pdfPath: cachedLibre,
        fromCache: true,
        errorCode: "",
        sheetMap: spreadsheetUnknownSheetMap,
        pageCount: cachedValidation.pageCount
      };
    }
    try {
      await fsp.unlink(cachedLibre);
    } catch (_error) {
      // ignore stale cache cleanup failure
    }
  }

  const libreResult = SPREADSHEET_EXTENSIONS.has(ext)
    ? await convertSpreadsheetWithLibreOffice(sourcePath, librePdfPath, findLibreOfficeExecutable())
    : await convertWithLibreOffice(sourcePath, librePdfPath);
  const libreEngineMode = libreResult?.engineMode || libreModeHint;
  if (libreResult.ok && fs.existsSync(librePdfPath)) {
    const libreValidation = await validateConvertedPdfPath(librePdfPath);
    if (libreValidation.ok) {
      return {
        ok: true,
        engine: libreEngineMode,
        warningMessage: "Office 기본 변환에 실패해 LibreOffice 호환 엔진으로 열었습니다.",
        pdfPath: librePdfPath,
        fromCache: false,
        errorCode: "",
        sheetMap: spreadsheetUnknownSheetMap,
        pageCount: libreValidation.pageCount
      };
    }
    try {
      await fsp.unlink(librePdfPath);
    } catch (_error) {
      // ignore invalid conversion cleanup failure
    }
  }

  const timedOut = isTimeoutResult(comResult) || isTimeoutResult(libreResult);
  const reasonHint = timedOut
    ? "변환 시간이 초과되어 호환 보기로 전환했습니다."
    : libreResult.missingEngine
      ? "Word/Excel/PowerPoint 또는 LibreOffice 변환 엔진을 찾지 못해 호환 보기로 전환했습니다."
      : "고품질 변환에 실패해 호환 보기로 전환했습니다.";

  const details = [comResult?.stderr, comResult?.error, libreResult?.stderr, libreResult?.error]
    .map((value) => String(value || "").trim())
    .find(Boolean);

  return {
    ok: false,
    engine: "fallback-text",
    warningMessage: buildConversionWarning(ext, details, reasonHint),
    errorCode: timedOut ? "ENGINE_TIMEOUT" : libreResult.missingEngine ? "ENGINE_MISSING" : "CONVERT_FAILED",
    emptyDocument: false,
    sheetMap: spreadsheetUnknownSheetMap
  };
}

async function convertAccdbFileToPdf(sourcePath) {
  const accdbPdfPath = createConvertedPdfPath(sourcePath, "accdb-ucanaccess");
  const cached = readCachedPdf(accdbPdfPath);
  if (cached) {
    const validation = await validateConvertedPdfPath(cached);
    if (validation.ok) {
      let cachedSheetMap = [];
      const cachedDump = await runAccdbDump(sourcePath, {
        maxRows: ACCDB_PREVIEW_MAX_ROWS,
        maxTables: ACCDB_PREVIEW_MAX_TABLES,
        maxCols: ACCDB_PREVIEW_MAX_COLS
      });
      if (cachedDump.ok) {
        const cachedFallback = buildAccdbFallbackHtml({
          sourcePath,
          warningMessage: "",
          tables: Array.isArray(cachedDump.tables) ? cachedDump.tables : []
        });
        cachedSheetMap = Array.isArray(cachedFallback.sheetMap) ? cachedFallback.sheetMap : [];
      }
      return {
        ok: true,
        engine: "accdb-ucanaccess",
        warningMessage: "",
        pdfPath: cached,
        errorCode: "",
        fromCache: true,
        sheetMap: cachedSheetMap,
        pageCount: validation.pageCount
      };
    }
    try {
      await fsp.unlink(cached);
    } catch (_error) {
      // ignore stale cache cleanup failure
    }
  }

  const dump = await runAccdbDump(sourcePath, {
    maxRows: ACCDB_PREVIEW_MAX_ROWS,
    maxTables: ACCDB_PREVIEW_MAX_TABLES,
    maxCols: ACCDB_PREVIEW_MAX_COLS
  });
  if (!dump.ok) {
    const reasonHint = dump.errorCode === "ENGINE_TIMEOUT"
      ? "ACCDB 읽기 엔진 응답이 지연되어 호환 보기로 전환했습니다."
      : dump.errorCode === "ENGINE_MISSING"
        ? "ACCDB 읽기 엔진을 찾지 못해 호환 보기로 전환했습니다."
        : "ACCDB 데이터 추출에 실패해 호환 보기로 전환했습니다.";
    return {
      ok: false,
      engine: "fallback-text",
      warningMessage: buildConversionWarning(".accdb", dump.message, reasonHint),
      errorCode: dump.errorCode || "CONVERT_FAILED",
      sheetMap: []
    };
  }

  const tableList = Array.isArray(dump.tables) ? dump.tables : [];
  const fallback = buildAccdbFallbackHtml({
    sourcePath,
    warningMessage: "",
    tables: tableList
  });
  if (!fallback?.html) {
    return {
      ok: false,
      engine: "fallback-text",
      warningMessage: "ACCDB 문서에서 표시할 테이블 데이터를 찾지 못했습니다.",
      errorCode: "EMPTY_DOCUMENT",
      sheetMap: []
    };
  }

  const pdfBuffer = await htmlToPdfBuffer(fallback.html);
  await fsp.writeFile(accdbPdfPath, pdfBuffer);
  const validation = await validateConvertedPdfPath(accdbPdfPath);
  if (!validation.ok) {
    throw new Error("INVALID_CONVERTED_PDF");
  }
  return {
    ok: true,
    engine: "accdb-ucanaccess",
    warningMessage: "",
    pdfPath: accdbPdfPath,
    errorCode: "",
    fromCache: false,
    sheetMap: Array.isArray(fallback.sheetMap) ? fallback.sheetMap : [],
    pageCount: validation.pageCount
  };
}

async function collectHwpCandidate({ mode, pdfPath, fromCache, runResult, analysisMode = "full", forcedPageTarget = 0 }) {
  const normalizedMode = String(mode || "").toLowerCase();
  if (isBlockedHwpMode(normalizedMode)) {
    return {
      mode,
      path: pdfPath,
      available: false,
      selectionEligible: false,
      fromCache: Boolean(fromCache),
      runResult: runResult || null,
      engineState: "blocked-by-policy",
      securityModuleRegistered: false,
      failureCode: "ENGINE_DISABLED",
      pageCount: 0,
      targetPageCount: Math.max(0, Math.floor(Number(forcedPageTarget || 0))),
      pageTargetDelta: 0,
      artifactScore: 100,
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
      score: 0,
      qualityScore: 0
    };
  }
  const exists = Boolean(pdfPath && fs.existsSync(pdfPath));
  if (!exists) {
    return {
      mode,
      path: pdfPath,
      available: false,
      selectionEligible: false,
      fromCache: Boolean(fromCache),
      runResult: runResult || null,
      engineState: String(runResult?.engineState || "not-run"),
      securityModuleRegistered: runResult?.securityModuleRegistered === true,
      failureCode: detectHwpRunFailureCode(runResult),
      pageCount: 0,
      targetPageCount: Math.max(0, Math.floor(Number(forcedPageTarget || 0))),
      pageTargetDelta: 0,
      artifactScore: 100,
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
      score: 0,
      qualityScore: 0
    };
  }
  const validation = await validateConvertedPdfPath(pdfPath);
  if (!validation.ok) {
    try {
      await fsp.unlink(pdfPath);
    } catch (_error) {
      // ignore invalid candidate cleanup failure
    }
    return {
      mode,
      path: pdfPath,
      available: false,
      selectionEligible: false,
      fromCache: Boolean(fromCache),
      runResult: runResult || null,
      engineState: String(runResult?.engineState || "invalid-pdf"),
      securityModuleRegistered: runResult?.securityModuleRegistered === true,
      failureCode: validation.errorCode || "INVALID_CONVERTED_PDF",
      pageCount: 0,
      targetPageCount: Math.max(0, Math.floor(Number(forcedPageTarget || 0))),
      pageTargetDelta: 0,
      artifactScore: 100,
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
      score: 0,
      qualityScore: 0
    };
  }
  const normalizedAnalysisMode = String(analysisMode || "full").toLowerCase() === "quick" ? "quick" : "full";
  const analysis = await analyzeHwpPdfQuality(pdfPath, {
    analysisMode: normalizedAnalysisMode,
    maxPages: normalizedAnalysisMode === "quick" ? 1 : 2
  });
  const runResultPageCount = Number(runResult?.pageCount || 0);
  const pageCount = Number(analysis?.metrics?.pageCount || runResultPageCount || validation.pageCount || 0);
  const bookletLike = Boolean(analysis?.metrics?.bookletLikePages > 0 || (await isLikelyBookletPdf(pdfPath)));
  const overlapLineCount = Number(analysis?.metrics?.overlapLineCount || 0);
  const rasterLineHitCount = Number(analysis?.metrics?.rasterLineHitCount || 0);
  const textLinePierceRatio = Number(analysis?.metrics?.textLinePierceRatio || 0);
  const markdownArtifact = runResult?.markdownArtifact || null;
  const mergedCellCount = Number(runResult?.mergedCellCount || 0);
  const colorStyleCount = Number(runResult?.colorStyleCount || 0);
  const highlightStyleCount = Number(runResult?.highlightStyleCount || 0);
  const fontSizeCount = Number(runResult?.fontSizeCount || 0);
  const tableHeaderCount = Number(runResult?.tableHeaderCount || 0);
  const textStrokeCount = Number(runResult?.textStrokeCount || 0);
  const textShadowCount = Number(runResult?.textShadowCount || 0);
  const cellDimensionCount = Number(runResult?.cellDimensionCount || 0);
  const tableBorderStyleCount = Number(runResult?.tableBorderStyleCount || 0);
  const centerAlignCount = Number(runResult?.centerAlignCount || 0);
  const targetPageCount = Math.max(
    0,
    Math.floor(Number(forcedPageTarget || runResult?.targetPageCount || 0))
  );
  const pageTargetDelta = targetPageCount > 0 ? Math.abs(pageCount - targetPageCount) : 0;
  let artifactScore = clampNumber(
    overlapLineCount * 16 + rasterLineHitCount * 12 + textLinePierceRatio * 75 + (bookletLike ? 28 : 0),
    0,
    100
  );
  if (mode === "hwpjs-html" && markdownArtifact?.inspected) {
    const strikeRatio = Number(markdownArtifact.strikeRatio || 0);
    const strikeSegments = Number(markdownArtifact.strikeSegments || 0);
    const tableLineCount = Number(markdownArtifact.tableLineCount || 0);
    const noisyTableLineCount = Number(markdownArtifact.noisyTableLineCount || 0);
    const noisyTableRatio = tableLineCount > 0 ? clampNumber(noisyTableLineCount / tableLineCount, 0, 1) : 0;
    const symbolNoiseRatio = Number(markdownArtifact.symbolNoiseRatio || 0);
    const headingOnlyCount = Number(markdownArtifact.headingOnlyCount || 0);
    artifactScore = clampNumber(
      artifactScore
      + strikeRatio * 95
      + Math.min(24, strikeSegments * 0.8)
      + noisyTableRatio * 52
      + symbolNoiseRatio * 58
      + Math.min(10, headingOnlyCount * 1.4),
      0,
      100
    );
  }
  if (mode === "hwpjs-html" && mergedCellCount > 0) {
    artifactScore = clampNumber(artifactScore - Math.min(16, mergedCellCount * 1.2), 0, 100);
  }
  let score = Number(analysis?.score || 0);
  if (bookletLike) {
    score = Math.max(0, score - 30);
  }
  if (mode === "hwp-saveas") {
    score += 2;
  }
  if (mode === "hwpjs-html") {
    score = Math.max(0, score - 22);
    if (runResult?.pageModelUsed === true) {
      score += 26;
      artifactScore = clampNumber(artifactScore - 18, 0, 100);
      const pageModel = runResult?.hwpPageModel || {};
      if (pageModel.anchorMatched === false) {
        score = Math.max(0, score - 55);
        artifactScore = clampNumber(artifactScore + 42, 0, 100);
      }
    }
    const styleStructureBonus = clampNumber(
      Math.log2(Math.max(1, colorStyleCount) + 1) * 2.4
      + Math.log2(Math.max(1, fontSizeCount) + 1) * 1.8
      + highlightStyleCount * 3.4
      + tableHeaderCount * 6.5
      + Math.log2(Math.max(1, cellDimensionCount) + 1) * 2.2
      + Math.log2(Math.max(1, tableBorderStyleCount) + 1) * 2.0
      + Math.log2(Math.max(1, centerAlignCount) + 1) * 2.1
      + textStrokeCount * 1.8
      + textShadowCount * 1.2,
      0,
      42
    );
    score += styleStructureBonus;
    if (markdownArtifact?.inspected) {
      const tableLineCount = Number(markdownArtifact.tableLineCount || 0);
      const noisyTableLineCount = Number(markdownArtifact.noisyTableLineCount || 0);
      const noisyTableRatio = tableLineCount > 0 ? clampNumber(noisyTableLineCount / tableLineCount, 0, 1) : 0;
      const symbolNoiseRatio = Number(markdownArtifact.symbolNoiseRatio || 0);
      if (tableLineCount >= 6 && noisyTableRatio >= 0.35) {
        score = Math.max(0, score - 16);
      }
      if (symbolNoiseRatio >= 0.28) {
        score = Math.max(0, score - 14);
      }
    }
  }
  if (targetPageCount > 0) {
    if (pageTargetDelta === 0) {
      score += 18;
      artifactScore = clampNumber(artifactScore - 18, 0, 100);
    } else {
      score = Math.max(0, score - Math.min(70, pageTargetDelta * 18));
      artifactScore = clampNumber(artifactScore + Math.min(52, pageTargetDelta * 14), 0, 100);
    }
  }
  const mergeStructureBonus = mode === "hwpjs-html"
    ? clampNumber(Math.log2(Math.max(1, mergedCellCount) + 1) * 2.6, 0, 10)
    : 0;
  const qualityScore = clampNumber(score - artifactScore * 0.58 + mergeStructureBonus, 0, 100);
  return {
    mode,
    path: pdfPath,
    available: true,
    selectionEligible: mode === "hwp-converter-exe" ? process.env.LOOKUP_ENABLE_HWP_CONVERTER_EXE === "1" : true,
    fromCache: Boolean(fromCache),
    runResult: runResult || null,
    engineState: String(runResult?.engineState || "success"),
    securityModuleRegistered: runResult?.securityModuleRegistered === true,
    failureCode: "",
    pageCount,
    artifactScore: clampNumber(artifactScore, 0, 100),
    analysis,
    bookletLike,
    score: clampNumber(score, 0, 100),
    qualityScore,
    mergedCellCount,
    colorStyleCount,
    highlightStyleCount,
    fontSizeCount,
    tableHeaderCount,
    textStrokeCount,
    textShadowCount,
    cellDimensionCount,
    tableBorderStyleCount,
    centerAlignCount,
    targetPageCount,
    pageTargetDelta,
    pageTargetMatched: targetPageCount > 0 ? pageTargetDelta === 0 : true
  };
}

function buildHwpSelectionWarning(selected, candidates) {
  if (!selected) {
    return "";
  }
  const hasForcedTargetMismatch = candidates.some((candidate) => candidate.excludedReason === "PAGE_COUNT_TARGET_MISMATCH");
  if (hasForcedTargetMismatch) {
    return "페이지 수 목표와 맞지 않는 결과를 제외하고 가장 유사한 결과를 선택했습니다.";
  }
  const hasPageCountReplacement = candidates.some((candidate) => candidate.excludedReason === "PAGE_COUNT_LOW");
  if (hasPageCountReplacement) {
    return "페이지 수가 비정상이라 더 안전한 변환 결과로 교체했습니다.";
  }
  if (selected.mode === "hwpjs-html" || !isNativeHwpMode(selected.mode)) {
    return "호환 보기입니다. 원본과 배치가 다를 수 있습니다.";
  }
  const switchedByArtifact = candidates.some(
    (candidate) => candidate.mode !== selected.mode
      && candidate.available
      && Number(candidate.artifactScore || 0) >= 45
      && isNativeHwpMode(candidate.mode)
  );
  if (switchedByArtifact || Number(selected.artifactScore || 0) >= 45) {
    return "변환 결과에 표시 오류가 감지되어 다른 방식으로 자동 전환했습니다.";
  }
  return "";
}

function buildHwpCandidateScoreRows(candidates) {
  return candidates.map((candidate) => ({
    ...(candidate?.runResult?.markdownArtifact && candidate.runResult.markdownArtifact.inspected
      ? {
        markdownStrikeRatio: Number(candidate.runResult.markdownArtifact.strikeRatio || 0),
        markdownStrikeSegments: Number(candidate.runResult.markdownArtifact.strikeSegments || 0),
        markdownTableLineCount: Number(candidate.runResult.markdownArtifact.tableLineCount || 0),
        markdownNoisyTableLineCount: Number(candidate.runResult.markdownArtifact.noisyTableLineCount || 0),
        markdownSymbolNoiseRatio: Number(candidate.runResult.markdownArtifact.symbolNoiseRatio || 0)
      }
      : {}),
    mode: candidate.mode,
    score: candidate.score,
    qualityScore: candidate.qualityScore,
    mergedCellCount: Number(candidate.mergedCellCount || candidate?.runResult?.mergedCellCount || 0),
    colorStyleCount: Number(candidate.colorStyleCount || candidate?.runResult?.colorStyleCount || 0),
    pageModelUsed: candidate?.runResult?.pageModelUsed === true,
    hwpPageModel: candidate?.runResult?.hwpPageModel || null,
    highlightStyleCount: Number(candidate.highlightStyleCount || candidate?.runResult?.highlightStyleCount || 0),
    fontSizeCount: Number(candidate.fontSizeCount || candidate?.runResult?.fontSizeCount || 0),
    tableHeaderCount: Number(candidate.tableHeaderCount || candidate?.runResult?.tableHeaderCount || 0),
    textStrokeCount: Number(candidate.textStrokeCount || candidate?.runResult?.textStrokeCount || 0),
    textShadowCount: Number(candidate.textShadowCount || candidate?.runResult?.textShadowCount || 0),
    cellDimensionCount: Number(candidate.cellDimensionCount || candidate?.runResult?.cellDimensionCount || 0),
    tableBorderStyleCount: Number(candidate.tableBorderStyleCount || candidate?.runResult?.tableBorderStyleCount || 0),
    centerAlignCount: Number(candidate.centerAlignCount || candidate?.runResult?.centerAlignCount || 0),
    engineState: candidate.engineState || "",
    securityModuleRegistered: candidate.securityModuleRegistered === true,
    fromCache: candidate.fromCache,
    selectionEligible: candidate.selectionEligible !== false,
    excludedReason: candidate.excludedReason || "",
    bookletLike: candidate.bookletLike,
    inspected: candidate.analysis.inspected,
    analysisMode: candidate.analysis.analysisMode || "full",
    pageCount: Number(candidate.pageCount || candidate.analysis?.metrics?.pageCount || 0),
    targetPageCount: Number(candidate.targetPageCount || candidate?.runResult?.targetPageCount || 0),
    pageTargetDelta: Number(candidate.pageTargetDelta || 0),
    pageTargetMatched: candidate.targetPageCount > 0
      ? Number(candidate.pageTargetDelta || 0) === 0
      : true,
    calibrated: candidate.runResult?.calibrated === true,
    calibrationAttempts: Number(candidate.runResult?.calibrationAttempts || 0),
    calibrationScale: Number(candidate.runResult?.calibrationScale || 0),
    artifactScore: Number(candidate.artifactScore || 0),
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

function buildHwpCandidateFailureRows(candidates) {
  return candidates
    .filter((candidate) => !candidate.available || candidate.selectionEligible === false)
    .map((candidate) => ({
      mode: candidate.mode,
      failureCode: candidate.failureCode || (candidate.selectionEligible === false ? (candidate.excludedReason || "PAGE_COUNT_LOW") : "CONVERT_FAILED"),
      engineState: candidate.engineState || "",
      securityModuleRegistered: candidate.securityModuleRegistered === true,
      excludedReason: candidate.excludedReason || "",
      pageCount: Number(candidate.pageCount || 0),
      targetPageCount: Number(candidate.targetPageCount || candidate?.runResult?.targetPageCount || 0),
      pageTargetDelta: Number(candidate.pageTargetDelta || 0),
      stderr: String(candidate.runResult?.stderr || "").slice(0, 200),
      error: String(candidate.runResult?.error || "").slice(0, 200)
    }));
}

function buildHwpEngineState(candidates) {
  const state = {};
  for (const candidate of candidates) {
    const mode = String(candidate?.mode || "").trim();
    if (!mode) {
      continue;
    }
    if (candidate?.available) {
      state[mode] = "success";
      continue;
    }
    if (candidate?.failureCode === "ENGINE_MISSING") {
      state[mode] = "engine_missing";
      continue;
    }
    if (candidate?.failureCode === "ENGINE_TIMEOUT") {
      state[mode] = "timeout";
      continue;
    }
    if (candidate?.failureCode === "PAGE_COUNT_ZERO" || candidate?.failureCode === "INVALID_CONVERTED_PDF") {
      state[mode] = "invalid_pdf";
      continue;
    }
    if (candidate?.failureCode === "PAGE_COUNT_TARGET_MISMATCH" || candidate?.excludedReason === "PAGE_COUNT_TARGET_MISMATCH") {
      state[mode] = "page_target_mismatch";
      continue;
    }
    state[mode] = "failed";
  }
  return state;
}

function detectSecurityModuleRegistered(candidates, selectedCandidate = null) {
  if (selectedCandidate?.runResult?.securityModuleRegistered === true) {
    return true;
  }
  return candidates.some((candidate) => candidate?.runResult?.securityModuleRegistered === true);
}

function isHwpQuickCandidateStable(candidate) {
  if (!candidate?.available) {
    return false;
  }
  if (!isNativeHwpMode(candidate.mode)) {
    return false;
  }
  const metrics = candidate.analysis?.metrics || {};
  if (Number(candidate.pageCount || 0) <= 0) {
    return false;
  }
  if (candidate.bookletLike) {
    return false;
  }
  if (Number(candidate.artifactScore || 0) >= 36) {
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
  return Number(candidate.qualityScore || candidate.score || 0) >= 80;
}

async function convertDocumentToPdf(sourcePath, options = {}) {
  await ensureFeedbackGate();
  if (typeof sourcePath !== "string" || !sourcePath.trim()) {
    throw new Error("파일 경로가 올바르지 않습니다.");
  }

  const resolved = path.resolve(sourcePath);
  const operationId = String(options?.operationId || convertProgressState.activeOperationId || "").trim();
  const sendStageStatus = (stage, percent, message, extra = {}) => {
    sendConvertStatus(stage, percent, message, {
      ...extra,
      operationId
    });
  };
  const ext = getFileExtension(resolved);
  if (!SUPPORTED_DOCUMENT_EXTENSIONS.has(ext)) {
    throw new Error("지원하지 않는 파일 형식입니다.");
  }

  const stat = await fsp.stat(resolved);
  if (!stat.isFile()) {
    throw new Error("파일이 존재하지 않습니다.");
  }

  if (ext === ".pdf") {
    const pdfBuffer = await fsp.readFile(resolved);
    const captureContent = await inspectPdfCaptureContent(pdfBuffer, { stopWhenNonPlaceholder: true });
    sendStageStatus("표시", 100, "PDF 문서를 바로 열고 있습니다.", {
      sourceExt: ext,
      convertMode: "native-pdf"
    });
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: false,
      convertMode: "native",
      warningMessage: captureContent.placeholderOnly ? "빈 캡처 PDF · 원본 페이지에서 다시 저장" : "",
      errorCode: "",
      convertDiagnostics: {
        selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
        selectedMode: "native",
        validationPassed: true,
        pageCount: captureContent.pageCount || 0,
        candidateScores: [],
        candidateFailures: [],
        artifactScore: 0,
        engineState: {},
        securityModuleRegistered: false,
        fallbackReason: ""
      },
      sheetMap: [],
      pdfPath: resolved,
      pdfBuffer
    };
  }

  let warningMessage = "";
  let errorCode = "";
  sendStageStatus("엔진 확인", 5, "변환 엔진을 확인하고 있습니다.", {
    sourceExt: ext,
    messageKey: "convert.step.engine"
  });

  if (WEB_DOCUMENT_EXTENSIONS.has(ext)) {
    sendStageStatus("분석", 18, "웹 페이지 전체 높이를 확인하고 있습니다.", {
      sourceExt: ext,
      selectedMode: "web-html",
      messageKey: "convert.step.web-analyze"
    });
    sendStageStatus("변환", 38, "숨은 내용을 끝까지 불러오고 있습니다.", {
      sourceExt: ext,
      selectedMode: "web-html",
      messageKey: "convert.step.web-load"
    });
    const webResult = await convertWebHtmlFileToPdf(resolved);
    if (webResult.ok && webResult.pdfPath && fs.existsSync(webResult.pdfPath)) {
      sendStageStatus("검사", 72, "PDF 페이지 나누기 결과를 검사하고 있습니다.", {
        sourceExt: ext,
        selectedMode: webResult.engine || "web-html",
        sourceType: webResult.diagnostics?.sourceType || ""
      });
      const webValidation = await validateConvertedPdfPath(webResult.pdfPath);
      if (!webValidation.ok) {
        warningMessage = buildConversionWarning(
          ext,
          webValidation.reason,
          "웹 페이지 PDF 변환 결과가 손상되어 텍스트 기반 보기로 전환했습니다."
        );
        errorCode = webValidation.errorCode || "INVALID_CONVERTED_PDF";
      } else {
        sendStageStatus("표시", 100, "문서를 화면에 표시합니다.", {
          sourceExt: ext,
          selectedMode: webResult.engine || "web-html",
          pageCount: webValidation.pageCount
        });
        return {
          sourcePath: resolved,
          sourceExt: ext,
          converted: true,
          convertMode: webResult.engine || "web-html",
          warningMessage: webResult.warningMessage || "",
          errorCode: "",
          convertDiagnostics: {
            selectedMode: webResult.engine || "web-html",
            sourceType: webResult.diagnostics?.sourceType || "",
            validationPassed: true,
            pageCount: webValidation.pageCount,
            candidateScores: [],
            fallbackReason: "",
            webPage: webResult.diagnostics?.webPage || null
          },
          sheetMap: [],
          pdfPath: webResult.pdfPath,
          pdfBuffer: await fsp.readFile(webResult.pdfPath)
        };
      }
    } else {
      warningMessage = webResult.warningMessage || "웹 페이지를 PDF로 변환하지 못해 텍스트 기반 보기로 전환했습니다.";
      errorCode = webResult.errorCode || "";
    }
  } else if (ACCESS_EXTENSIONS.has(ext)) {
    sendStageStatus("변환", 35, "데이터베이스 문서를 읽고 있습니다.", {
      sourceExt: ext
    });
    const accessResult = await convertAccdbFileToPdf(resolved);
    if (accessResult.ok && accessResult.pdfPath && fs.existsSync(accessResult.pdfPath)) {
      sendStageStatus("검사", 70, "변환 결과를 검사하고 있습니다.", {
        sourceExt: ext,
        selectedMode: accessResult.engine || "accdb-ucanaccess"
      });
      const accessValidation = await validateConvertedPdfPath(accessResult.pdfPath);
      if (!accessValidation.ok) {
        warningMessage = buildConversionWarning(
          ext,
          accessValidation.reason,
          "변환된 PDF가 손상되어 다른 방식으로 다시 시도했습니다."
        );
        errorCode = accessValidation.errorCode || "INVALID_CONVERTED_PDF";
      } else {
        sendStageStatus("표시", 100, "문서를 화면에 표시합니다.", {
          sourceExt: ext,
          selectedMode: accessResult.engine || "accdb-ucanaccess",
          pageCount: accessValidation.pageCount
        });
        return {
          sourcePath: resolved,
          sourceExt: ext,
          converted: true,
          convertMode: accessResult.engine || "accdb-ucanaccess",
          warningMessage: accessResult.warningMessage || "",
          errorCode: "",
          convertDiagnostics: {
            selectedMode: accessResult.engine || "accdb-ucanaccess",
            validationPassed: true,
            pageCount: accessValidation.pageCount,
            candidateScores: [],
            fallbackReason: ""
          },
          sheetMap: Array.isArray(accessResult.sheetMap) ? accessResult.sheetMap : [],
          pdfPath: accessResult.pdfPath,
          pdfBuffer: await fsp.readFile(accessResult.pdfPath)
        };
      }
    } else {
      warningMessage = accessResult.warningMessage || "ACCDB 변환 엔진을 찾지 못해 텍스트 기반 보기로 전환했습니다.";
      errorCode = accessResult.errorCode || "";
    }
  } else if (OFFICE_EXTENSIONS.has(ext)) {
    sendStageStatus("변환", 35, "문서를 PDF로 변환하고 있습니다.", {
      sourceExt: ext
    });
    const officeResult = await convertOfficeFileToPdf(resolved, ext);
    if (officeResult.ok && officeResult.pdfPath && fs.existsSync(officeResult.pdfPath)) {
      sendStageStatus("검사", 70, "변환 결과를 검사하고 있습니다.", {
        sourceExt: ext,
        selectedMode: officeResult.engine || "office-com"
      });
      const officeValidation = await validateConvertedPdfPath(officeResult.pdfPath);
      if (!officeValidation.ok) {
        warningMessage = buildConversionWarning(
          ext,
          officeValidation.reason,
          "변환된 PDF가 손상되어 다른 방식으로 다시 시도했습니다."
        );
        errorCode = officeValidation.errorCode || "INVALID_CONVERTED_PDF";
      } else {
        sendStageStatus("표시", 100, "문서를 화면에 표시합니다.", {
          sourceExt: ext,
          selectedMode: officeResult.engine || "office-com",
          pageCount: officeValidation.pageCount
        });
        return {
          sourcePath: resolved,
          sourceExt: ext,
          converted: true,
          convertMode: officeResult.engine || "office-com",
          warningMessage: officeResult.warningMessage || "",
          errorCode: "",
          convertDiagnostics: {
            selectedMode: officeResult.engine || "office-com",
            validationPassed: true,
            pageCount: officeValidation.pageCount,
            candidateScores: [],
            fallbackReason: ""
          },
          sheetMap: Array.isArray(officeResult.sheetMap) ? officeResult.sheetMap : [],
          pdfPath: officeResult.pdfPath,
          pdfBuffer: await fsp.readFile(officeResult.pdfPath)
        };
      }
    } else {
      warningMessage = officeResult.warningMessage || "고품질 변환 엔진을 찾지 못해 텍스트 기반 보기 모드로 변환했습니다.";
      errorCode = officeResult.errorCode || "";
    }
  } else if (HWP_EXTENSIONS.has(ext)) {
    const sourceReadiness = inspectHwpFamilyBytes(await fsp.readFile(resolved), ext);
    const documentExt = sourceReadiness.documentExt;
    logConvertTrace("hwp-open-start", { source: path.basename(resolved), ext });
    sendStageStatus("변환", 30, "한글 문서를 여는 중입니다.", {
      sourceExt: ext,
      messageKey: "convert.step.open"
    });
    const forceCompatibility = options && options.forceCompatibility === true;
    const readHwpCachedPdf = (pdfPath, cacheOptions = {}) => forceCompatibility ? null : readCachedPdf(pdfPath, cacheOptions);
    const sourceSignature = buildHwpSelectionSourceSignature(resolved);
    const forcedPageTarget = resolveForcedPageTarget({
      sourcePath: resolved,
      sourceExt: ext,
      sourceSignature
    });
    const allowNativeHwpEngines = false;
    const allowPrintStyleHwpFallback = false;
    const hwpSaveAsPdfPath = createConvertedPdfPath(resolved, "hwp-saveas");
    const hwpExportPdfPath = createConvertedPdfPath(resolved, "hwp-export");
    const hwpPrintPdfPath = createConvertedPdfPath(resolved, "hwp-print-pdf");
    const hwpConverterExePdfPath = createConvertedPdfPath(resolved, "hwp-converter-exe");
    const hwpJsPdfPath = createConvertedPdfPath(resolved, "hwpjs-html");
    const hwpxReaderPdfPath = createConvertedPdfPath(resolved, "hwpx-reader");
    const rhwpPdfPath = createConvertedPdfPath(resolved, "hwp-rhwp");
    const hwpLibreModeHint = findLibreOfficeExecutable().mode || "libreoffice-system";
    const hwpLibrePdfPath = createConvertedPdfPath(resolved, hwpLibreModeHint);
    const hwpStatusForwarder = createHwpWorkerStatusForwarder(ext, operationId);
    const hancomComAvailable = false;
    const hasNativeSecurityCandidate = false;
    let nativeAttemptBudget = allowNativeHwpEngines ? 2 : 0;
    const consumeNativeAttempt = () => {
      if (nativeAttemptBudget <= 0) {
        return false;
      }
      nativeAttemptBudget -= 1;
      return true;
    };
    if (forceCompatibility) {
      sendStageStatus("변환", 34, "원본 변환을 건너뛰고 호환 보기로 전환합니다.", {
        sourceExt: ext,
        selectedMode: "hwpjs-html"
      });
    } else {
      sendStageStatus("변환", 34, "비설치 변환 엔진으로 문서를 여는 중입니다.", {
        sourceExt: ext,
        selectedMode: "hwpjs-html"
      });
    }
    let hwpPrintResult = null;
    let hwpSaveAsResult = null;
    let hwpExportResult = null;
    let hwpConverterExeResult = null;
    let hwpJsResult = null;
    let hwpxReaderResult = null;
    let hwpLibreResult = null;
    let rhwpResult = null;
    const hwpCandidates = [];
    const collectCandidate = (payload) => collectHwpCandidate({
      ...(payload || {}),
      forcedPageTarget
    });

    const selectorMeta = forceCompatibility ? null : readHwpSelectionMeta(resolved, {
      sourceSignature,
      forcedPageTarget
    });
    if (selectorMeta?.selectedMode && selectorMeta?.selectedPdfPath && fs.existsSync(selectorMeta.selectedPdfPath)) {
      let selectorInvalidReason = "";
      if (selectorMeta.validationPassed !== true) {
        selectorInvalidReason = "validation_not_passed";
      } else if (!selectorMeta.engineState || typeof selectorMeta.engineState !== "object") {
        selectorInvalidReason = "engine_state_missing";
      } else if (hasInvalidSelectorScore(selectorMeta.candidateScores, selectorMeta.selectedMode)) {
        selectorInvalidReason = "score_not_trusted";
      } else if (Number(selectorMeta.pageCount || 0) <= 0) {
        selectorInvalidReason = "page_count_missing";
      } else if (Number(forcedPageTarget || 0) > 0 && Number(selectorMeta.pageCount || 0) !== Number(forcedPageTarget || 0)) {
        selectorInvalidReason = "forced_page_target_mismatch";
      }
      if (!selectorInvalidReason && (selectorMeta.qualityTier === "compatibility-layout" || selectorMeta.qualityTier === "compatibility-text")) {
        const updatedAtTs = selectorMeta.updatedAt ? Date.parse(selectorMeta.updatedAt) : NaN;
        if (!Number.isFinite(updatedAtTs) || Date.now() - updatedAtTs > COMPATIBILITY_SELECTION_META_TTL_MS) {
          selectorInvalidReason = "compatibility_cache_expired";
        }
      }
      if (!selectorInvalidReason && selectorMeta.selectedMode === "hwpjs-html") {
        const scoreRow = findCandidateScoreRow(selectorMeta.candidateScores, selectorMeta.selectedMode);
        const qualityScore = Number(scoreRow?.qualityScore ?? scoreRow?.finalScore ?? scoreRow?.score ?? 0);
        if (!Number.isFinite(qualityScore) || qualityScore < 60) {
          selectorInvalidReason = "compatibility_score_low";
        }
      }
      if (!selectorInvalidReason) {
        const cachedValidation = await validateConvertedPdfPath(selectorMeta.selectedPdfPath);
        if (cachedValidation.ok) {
          sendStageStatus("표시", 100, "검증된 캐시 문서를 바로 표시합니다.", {
            sourceExt: ext,
            selectedMode: selectorMeta.selectedMode,
            pageCount: cachedValidation.pageCount
          });
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
              validationPassed: true,
              analysisMode: selectorMeta.analysisMode || "quick",
              candidateScores: Array.isArray(selectorMeta.candidateScores) ? selectorMeta.candidateScores : [],
              candidateFailures: Array.isArray(selectorMeta.candidateFailures) ? selectorMeta.candidateFailures : [],
              pageCount: Number(cachedValidation.pageCount || selectorMeta.pageCount || 0),
              targetPageCount: Number(forcedPageTarget || 0),
              pageTargetMatched: Number(forcedPageTarget || 0) > 0
                ? Number(cachedValidation.pageCount || selectorMeta.pageCount || 0) === Number(forcedPageTarget || 0)
                : true,
              artifactScore: Number(selectorMeta.artifactScore || 0),
              hwpPageModel: selectorMeta.hwpPageModel || null,
              engineState: selectorMeta.engineState && typeof selectorMeta.engineState === "object" ? selectorMeta.engineState : {},
              securityModuleRegistered: selectorMeta.securityModuleRegistered === true,
              fallbackReason: selectorMeta.fallbackReason || selectorMeta.warningMessage || ""
            },
            sheetMap: [],
            pdfPath: selectorMeta.selectedPdfPath,
            pdfBuffer: await fsp.readFile(selectorMeta.selectedPdfPath)
          };
        }
        selectorInvalidReason = cachedValidation.reason || "cached_pdf_invalid";
      }
      logConvertTrace("hwp-meta-rejected", {
        selectedMode: selectorMeta.selectedMode,
        reason: selectorInvalidReason,
        source: path.basename(resolved)
      });
    }

    if (!forceCompatibility) {
      sendStageStatus("변환", 45, "비설치 엔진 변환을 시도합니다.", {
        sourceExt: ext,
        messageKey: "convert.step.save"
      });
    }
    const hwpSaveAsCached = allowNativeHwpEngines ? readHwpCachedPdf(hwpSaveAsPdfPath) : null;
    const createNativeUnavailableRunResult = () => {
      const skipStateNow = getHwpNativeSkipState();
      const securityUnavailable = !hasNativeSecurityCandidate;
      const missingEngine = !hancomComAvailable || skipStateNow.reasonCode === "ENGINE_MISSING" || securityUnavailable;
      const sessionSkipped = skipStateNow.skip;
      const stderr = forceCompatibility
        ? "hwp_native_forced_compatibility"
        : !hancomComAvailable
          ? "hwp_com_engine_missing"
          : securityUnavailable
            ? "hwp_security_module_untrusted"
          : sessionSkipped
            ? "hwp_native_skipped_session"
            : "hwp_native_not_attempted";
      const error = forceCompatibility
        ? "hwp_native_forced_compatibility"
        : (skipStateNow.reasonMessage || stderr);
      const engineState = forceCompatibility
        ? "forced_compatibility"
        : !hancomComAvailable
          ? "engine_missing"
          : securityUnavailable
            ? "security_untrusted"
          : sessionSkipped
            ? "session-skip"
            : "not-attempted";
      return {
        ok: false,
        missingEngine,
        stderr,
        error,
        engineState,
        securityModuleRegistered: false
      };
    };
    let saveAsCandidate = hwpSaveAsCached
      ? await collectCandidate({
        mode: "hwp-saveas",
        pdfPath: hwpSaveAsCached,
        fromCache: true,
        analysisMode: "quick"
      })
      : await (async () => {
        if (!consumeNativeAttempt()) {
          return collectCandidate({
            mode: "hwp-saveas",
            pdfPath: "",
            fromCache: false,
            runResult: createNativeUnavailableRunResult(),
            analysisMode: "quick"
          });
        }
        hwpSaveAsResult = await convertWithHancomComSaveAs(resolved, hwpSaveAsPdfPath, ext, hwpStatusForwarder);
        updateHwpNativeSessionStateFromRunResult(hwpSaveAsResult);
        return collectCandidate({
          mode: "hwp-saveas",
          pdfPath: hwpSaveAsPdfPath,
          fromCache: false,
          runResult: hwpSaveAsResult,
          analysisMode: "quick"
        });
      })();
    hwpCandidates.push(saveAsCandidate);

    const hwpExportCached = allowNativeHwpEngines ? readHwpCachedPdf(hwpExportPdfPath) : null;
    let exportCandidate = hwpExportCached
      ? await collectCandidate({
        mode: "hwp-export",
        pdfPath: hwpExportCached,
        fromCache: true,
        analysisMode: "quick"
      })
      : await (async () => {
        if (!consumeNativeAttempt()) {
          return collectCandidate({
            mode: "hwp-export",
            pdfPath: "",
            fromCache: false,
            runResult: createNativeUnavailableRunResult(),
            analysisMode: "quick"
          });
        }
        hwpExportResult = await convertWithHancomComExport(resolved, hwpExportPdfPath, ext, hwpStatusForwarder);
        updateHwpNativeSessionStateFromRunResult(hwpExportResult);
        return collectCandidate({
          mode: "hwp-export",
          pdfPath: hwpExportPdfPath,
          fromCache: false,
          runResult: hwpExportResult,
          analysisMode: "quick"
        });
      })();
    if (exportCandidate?.available && String(exportCandidate.analysis?.analysisMode || "").toLowerCase() !== "full") {
      exportCandidate = await collectCandidate({
        mode: "hwp-export",
        pdfPath: exportCandidate.path,
        fromCache: true,
        runResult: hwpExportResult,
        analysisMode: "full"
      });
    }
    hwpCandidates.push(exportCandidate);

    if (documentExt === ".hwp" && !forceCompatibility) {
      const hwpLibreCached = readHwpCachedPdf(hwpLibrePdfPath);
      const libreCandidate = hwpLibreCached
        ? await collectCandidate({
          mode: hwpLibreModeHint,
          pdfPath: hwpLibreCached,
          fromCache: true,
          analysisMode: "full"
        })
        : await (async () => {
          hwpLibreResult = await convertWithLibreOffice(resolved, hwpLibrePdfPath);
          return collectCandidate({
            mode: hwpLibreResult?.engineMode || hwpLibreModeHint,
            pdfPath: hwpLibrePdfPath,
            fromCache: false,
            runResult: hwpLibreResult,
            analysisMode: "full"
          });
        })();
      hwpCandidates.push(libreCandidate);
    }
    if (documentExt === ".hwpx") {
      const hwpxReaderCached = readHwpCachedPdf(hwpxReaderPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
      const hwpxReaderCandidate = hwpxReaderCached
        ? await collectCandidate({
          mode: "hwpx-reader",
          pdfPath: hwpxReaderCached,
          fromCache: true,
          analysisMode: "full"
        })
        : await (async () => {
          hwpxReaderResult = await convertHwpxToPdfWithReader(resolved, hwpxReaderPdfPath, {
            forcedPageTarget,
            sourceSignature
          });
          return collectCandidate({
            mode: "hwpx-reader",
            pdfPath: hwpxReaderPdfPath,
            fromCache: false,
            runResult: hwpxReaderResult,
            analysisMode: "full"
          });
        })();
      hwpCandidates.push(hwpxReaderCandidate);
    }

    if (ext !== ".hwp" && isHwpQuickCandidateStable(saveAsCandidate)) {
      sendStageStatus("검사", 75, "변환 결과를 확인하고 있습니다.", {
        sourceExt: ext,
        selectedMode: saveAsCandidate.mode,
        messageKey: "convert.step.validate"
      });
      logConvertTrace("hwp-quick-pass", {
        selectedMode: saveAsCandidate.mode,
        score: saveAsCandidate.score,
        qualityScore: saveAsCandidate.qualityScore,
        pageCount: saveAsCandidate.pageCount,
        source: path.basename(resolved)
      });
      if (isNativeHwpMode(saveAsCandidate.mode)) {
        clearHwpNativeSkip();
      }
      const candidateScores = buildHwpCandidateScoreRows(hwpCandidates);
      const candidateFailures = buildHwpCandidateFailureRows(hwpCandidates);
      writeHwpSelectionMeta(resolved, {
        selectedMode: saveAsCandidate.mode,
        selectedPdfPath: saveAsCandidate.path,
        warningMessage: "",
        fallbackReason: "",
        analysisMode: "quick",
        sourceSignature,
        candidateScores,
        candidateFailures,
        pageCount: Number(saveAsCandidate.pageCount || 0),
        targetPageCount: Number(forcedPageTarget || 0),
        artifactScore: Number(saveAsCandidate.artifactScore || 0),
        engineVersion: getConversionEngineToken(saveAsCandidate.mode),
        engineState: buildHwpEngineState(hwpCandidates),
        securityModuleRegistered: detectSecurityModuleRegistered(hwpCandidates, saveAsCandidate),
        validationPassed: true
      });
      sendStageStatus("표시", 100, "문서를 화면에 표시합니다.", {
        sourceExt: ext,
        selectedMode: saveAsCandidate.mode,
        pageCount: Number(saveAsCandidate.pageCount || 0)
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: saveAsCandidate.mode,
        warningMessage: "",
        errorCode: "",
        convertDiagnostics: {
          selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
          selectedMode: saveAsCandidate.mode,
          validationPassed: true,
          analysisMode: "quick",
          candidateScores,
          candidateFailures,
          pageCount: Number(saveAsCandidate.pageCount || 0),
          targetPageCount: Number(forcedPageTarget || 0),
          pageTargetMatched: Number(forcedPageTarget || 0) > 0
            ? Number(saveAsCandidate.pageCount || 0) === Number(forcedPageTarget || 0)
            : true,
          artifactScore: Number(saveAsCandidate.artifactScore || 0),
          engineState: buildHwpEngineState(hwpCandidates),
          securityModuleRegistered: detectSecurityModuleRegistered(hwpCandidates, saveAsCandidate),
          fallbackReason: ""
        },
        sheetMap: [],
        pdfPath: saveAsCandidate.path,
        pdfBuffer: await fsp.readFile(saveAsCandidate.path)
      };
    }

    if (saveAsCandidate?.available && String(saveAsCandidate.analysis?.analysisMode || "").toLowerCase() !== "full") {
      saveAsCandidate = await collectCandidate({
        mode: "hwp-saveas",
        pdfPath: saveAsCandidate.path,
        fromCache: true,
        runResult: hwpSaveAsResult,
        analysisMode: "full"
      });
      hwpCandidates[0] = saveAsCandidate;
    }

    if (allowNativeHwpEngines && process.env.LOOKUP_ENABLE_HWP_CONVERTER_EXE === "1") {
      const hwpConverterExeCached = readHwpCachedPdf(hwpConverterExePdfPath);
      const converterExeCandidate = hwpConverterExeCached
        ? await collectCandidate({
          mode: "hwp-converter-exe",
          pdfPath: hwpConverterExeCached,
          fromCache: true,
          analysisMode: "full"
        })
        : await (async () => {
          hwpConverterExeResult = await convertWithHancomConverterExe(resolved, hwpConverterExePdfPath);
          return collectCandidate({
            mode: "hwp-converter-exe",
            pdfPath: hwpConverterExePdfPath,
            fromCache: false,
            runResult: hwpConverterExeResult,
            analysisMode: "full"
          });
        })();
      hwpCandidates.push(converterExeCandidate);
    }

    applyHwpPageCountGate(hwpCandidates, { forcedPageTarget });
    let selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    const needCompatibilityCandidate = documentExt === ".hwp";
    if (needCompatibilityCandidate) {
      const cachedHwpJs = readHwpCachedPdf(hwpJsPdfPath, { maxAgeMs: HWPJS_SELECTOR_META_TTL_MS });
      let hwpJsCandidate = cachedHwpJs
        ? await collectCandidate({
          mode: "hwpjs-html",
          pdfPath: cachedHwpJs,
          fromCache: true,
          analysisMode: "full"
        })
        : null;
      if (hwpJsCandidate && forcedPageTarget > 0 && Number(hwpJsCandidate.pageCount || 0) !== forcedPageTarget) {
        try {
          await fsp.unlink(cachedHwpJs);
        } catch (_error) {
          // ignore stale cache cleanup failure
        }
        hwpJsCandidate = null;
      }
      if (!hwpJsCandidate) {
        hwpJsResult = await convertHwpToPdfWithHwpJs(resolved, hwpJsPdfPath, {
          forcedPageTarget,
          sourceSignature
        });
        if (hwpJsResult.ok && fs.existsSync(hwpJsPdfPath)) {
          hwpJsCandidate = await collectCandidate({
            mode: "hwpjs-html",
            pdfPath: hwpJsPdfPath,
            fromCache: false,
            runResult: hwpJsResult,
            analysisMode: "full"
          });
        }
      }
      if (hwpJsCandidate) {
        hwpCandidates.push(hwpJsCandidate);
      }
      applyHwpPageCountGate(hwpCandidates, { forcedPageTarget });
      selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    }

    if (!selectedCandidate || Number(selectedCandidate.qualityScore || selectedCandidate.score || 0) < 60) {
      sendStageStatus("변환", 78, "한글 페이지 복원 중", { sourceExt: ext, messageKey: "convert.step.compatibility" });
      const rhwpCached = readHwpCachedPdf(rhwpPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
      if (!rhwpCached) rhwpResult = await convertHwpToPdfWithRhwp(resolved, rhwpPdfPath);
      const rhwpCandidate = await collectCandidate({
        mode: "hwp-rhwp", pdfPath: rhwpCached || rhwpPdfPath, fromCache: Boolean(rhwpCached),
        runResult: rhwpResult, analysisMode: "full"
      });
      if (selectedCandidate?.available && rhwpCandidate.available) {
        const basePages = Number(selectedCandidate.pageCount || 0);
        const alternatePages = Number(rhwpCandidate.pageCount || 0);
        if (basePages > 0 && (alternatePages > Math.max(basePages * 2, basePages + 3) || alternatePages < basePages / 2)) {
          rhwpCandidate.available = false;
          rhwpCandidate.selectionEligible = false;
          rhwpCandidate.excludedReason = "ALTERNATIVE_PAGE_COUNT_DIVERGED";
          rhwpCandidate.failureCode = "RENDER_ARTIFACT_DETECTED";
        }
      }
      hwpCandidates.push(rhwpCandidate);
      applyHwpPageCountGate(hwpCandidates, { forcedPageTarget });
      selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    }

    if (!forceCompatibility && allowPrintStyleHwpFallback && (!selectedCandidate || Number(selectedCandidate.qualityScore || selectedCandidate.score || 0) < 58)) {
      sendStageStatus("변환", 70, "PDF 변환용 출력 보조 경로를 시도합니다.", {
        sourceExt: ext,
        selectedMode: "hwp-print-pdf"
      });
      const hwpPrintCached = readHwpCachedPdf(hwpPrintPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
      let printCandidate = hwpPrintCached
        ? await collectCandidate({
          mode: "hwp-print-pdf",
          pdfPath: hwpPrintCached,
          fromCache: true,
          analysisMode: "quick"
        })
        : await (async () => {
          hwpPrintResult = await convertWithHancomComPrintPdf(resolved, hwpPrintPdfPath, ext, hwpStatusForwarder);
          updateHwpNativeSessionStateFromRunResult(hwpPrintResult);
          return collectCandidate({
            mode: "hwp-print-pdf",
            pdfPath: hwpPrintPdfPath,
            fromCache: false,
            runResult: hwpPrintResult,
            analysisMode: "quick"
          });
        })();
      if (printCandidate?.available && String(printCandidate.analysis?.analysisMode || "").toLowerCase() !== "full") {
        printCandidate = await collectCandidate({
          mode: "hwp-print-pdf",
          pdfPath: printCandidate.path,
          fromCache: true,
          runResult: hwpPrintResult,
          analysisMode: "full"
        });
      }
      hwpCandidates.push(printCandidate);
      applyHwpPageCountGate(hwpCandidates, { forcedPageTarget });
      selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
    }

    if (selectedCandidate?.available && selectedCandidate.path && fs.existsSync(selectedCandidate.path)) {
      sendStageStatus("검사", 82, "변환 결과를 확인하고 있습니다.", {
        sourceExt: ext,
        selectedMode: selectedCandidate.mode,
        messageKey: "convert.step.validate"
      });
      const selectedValidation = await validateConvertedPdfPath(selectedCandidate.path);
      if (!selectedValidation.ok) {
        selectedCandidate.selectionEligible = false;
        selectedCandidate.available = false;
        selectedCandidate.failureCode = selectedValidation.errorCode || "INVALID_CONVERTED_PDF";
        selectedCandidate.excludedReason = "INVALID_CONVERTED_PDF";
        applyHwpPageCountGate(hwpCandidates, { forcedPageTarget });
        selectedCandidate = chooseBestHwpCandidate(hwpCandidates);
      }
    }

    if (selectedCandidate?.available && selectedCandidate.path && fs.existsSync(selectedCandidate.path)) {
      const selectedMode = selectedCandidate.mode;
      if (isNativeHwpMode(selectedMode)) {
        clearHwpNativeSkip();
      }
      const selectionWarning = buildHwpSelectionWarning(selectedCandidate, hwpCandidates);
      let selectionErrorCode = "";
      if (selectionWarning.includes("원본 변환 엔진")) {
        selectionErrorCode = "ENGINE_MISSING";
      } else if (selectionWarning.includes("표시 오류") || selectionWarning.includes("페이지 수")) {
        selectionErrorCode = "RENDER_ARTIFACT_DETECTED";
      }
      const candidateScores = buildHwpCandidateScoreRows(hwpCandidates);
      const candidateFailures = buildHwpCandidateFailureRows(hwpCandidates);
      const hwpEngineState = buildHwpEngineState(hwpCandidates);
      const securityModuleRegistered = detectSecurityModuleRegistered(hwpCandidates, selectedCandidate);
      logConvertTrace("hwp-select-final", {
        selectedMode,
        score: selectedCandidate.score,
        qualityScore: selectedCandidate.qualityScore,
        analysisMode: selectedCandidate.analysis?.analysisMode || "full",
        pageCount: selectedCandidate.pageCount,
        source: path.basename(resolved)
      });
      writeHwpSelectionMeta(resolved, {
        selectedMode,
        selectedPdfPath: selectedCandidate.path,
        warningMessage: selectionWarning,
        fallbackReason: selectionWarning || "",
        analysisMode: selectedCandidate.analysis?.analysisMode || "full",
        sourceSignature,
        candidateScores,
        candidateFailures,
        pageCount: Number(selectedCandidate.pageCount || 0),
        targetPageCount: Number(forcedPageTarget || 0),
        artifactScore: Number(selectedCandidate.artifactScore || 0),
        engineVersion: getConversionEngineToken(selectedMode),
        hwpPageModel: selectedCandidate?.runResult?.hwpPageModel || null,
        engineState: hwpEngineState,
        securityModuleRegistered,
        validationPassed: true
      });
      sendStageStatus("표시", 100, "문서를 화면에 표시합니다.", {
        sourceExt: ext,
        selectedMode,
        pageCount: Number(selectedCandidate.pageCount || 0)
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: selectedMode,
        warningMessage: selectionWarning,
        errorCode: selectionErrorCode,
        convertDiagnostics: {
          selectorVersion: HWP_QUALITY_SELECTOR_VERSION,
          selectedMode,
          validationPassed: true,
          analysisMode: selectedCandidate.analysis?.analysisMode || "full",
          candidateScores,
          candidateFailures,
          pageCount: Number(selectedCandidate.pageCount || 0),
          targetPageCount: Number(forcedPageTarget || 0),
          pageTargetMatched: Number(forcedPageTarget || 0) > 0
            ? Number(selectedCandidate.pageCount || 0) === Number(forcedPageTarget || 0)
            : true,
          artifactScore: Number(selectedCandidate.artifactScore || 0),
          hwpPageModel: selectedCandidate?.runResult?.hwpPageModel || null,
          engineState: hwpEngineState,
          securityModuleRegistered,
          fallbackReason: selectionWarning || ""
        },
        sheetMap: [],
        pdfPath: selectedCandidate.path,
        pdfBuffer: await fsp.readFile(selectedCandidate.path)
      };
    }

    const hwpErrorDetail = [
      hwpSaveAsResult?.stderr,
      hwpSaveAsResult?.error,
      hwpExportResult?.stderr,
      hwpExportResult?.error,
      hwpPrintResult?.stderr,
      hwpPrintResult?.error,
      hwpConverterExeResult?.stderr,
      hwpConverterExeResult?.error,
      hwpJsResult?.reason,
      rhwpResult?.reason,
      hwpxReaderResult?.reason,
      hwpLibreResult?.stderr,
      hwpLibreResult?.error
    ]
      .map((value) => String(value || "").trim())
      .find(Boolean);
    const mergedFailureText = `${String(hwpSaveAsResult?.stderr || "")} ${String(hwpSaveAsResult?.error || "")} ${String(hwpExportResult?.stderr || "")} ${String(hwpExportResult?.error || "")} ${String(hwpPrintResult?.stderr || "")} ${String(hwpPrintResult?.error || "")}`;
    const hwpTimedOut = isTimeoutResult(hwpPrintResult) || isTimeoutResult(hwpSaveAsResult) || isTimeoutResult(hwpExportResult) || isTimeoutResult(hwpConverterExeResult);
    const hwpSecurityMissing =
      /hwp_security_module_missing/i.test(mergedFailureText)
      || hwpCandidates.some((candidate) => String(candidate.failureCode || "") === "ENGINE_MISSING");
    const hwpRenderArtifact = hwpCandidates.some(
      (candidate) =>
        Number(candidate.artifactScore || 0) >= 45
        || candidate.excludedReason === "PAGE_COUNT_LOW"
        || candidate.excludedReason === "PAGE_COUNT_TARGET_MISMATCH"
    );
    const reasonHint = forceCompatibility
      ? "사용자 요청으로 호환 보기 모드로 열었습니다."
      : hwpSecurityMissing
        ? "원본 변환을 위한 보안 확인이 필요해 호환 보기로 열었습니다."
        : hwpTimedOut
          ? "한컴 변환 엔진 응답이 지연되어 호환 보기로 전환했습니다."
          : hwpRenderArtifact
            ? "변환 결과에 표시 오류가 감지되어 다른 방식으로 자동 전환했습니다."
            : "원본 레이아웃 변환에 실패해 호환 보기로 전환했습니다.";
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
      validationPassed: false,
      analysisMode: "full",
      candidateScores: buildHwpCandidateScoreRows(hwpCandidates),
      candidateFailures: buildHwpCandidateFailureRows(hwpCandidates),
      pageCount: 0,
      targetPageCount: Number(forcedPageTarget || 0),
      pageTargetMatched: false,
      artifactScore: 100,
      engineState: buildHwpEngineState(hwpCandidates),
      securityModuleRegistered: detectSecurityModuleRegistered(hwpCandidates),
      fallbackReason: reasonHint
    };
    logConvertTrace("hwp-fallback", {
      errorCode,
      fallbackReason: reasonHint,
      source: path.basename(resolved)
    });
    const fallbackPdfPath = createConvertedPdfPath(resolved, "fallback");
    const fallbackCache = readHwpCachedPdf(fallbackPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
    if (fallbackCache) {
      const fallbackValidation = await validateConvertedPdfPath(fallbackCache);
      const pageTargetMatched = Number(forcedPageTarget || 0) > 0
        ? Number(fallbackValidation.pageCount || 0) === Number(forcedPageTarget || 0)
        : true;
      if (fallbackValidation.ok && pageTargetMatched) {
        sendStageStatus("변환", 90, "호환 보기로 전환합니다.", {
          sourceExt: ext,
          selectedMode: "fallback-text",
          fallbackReason: reasonHint,
          messageKey: "convert.step.compatibility"
        });
        return {
          sourcePath: resolved,
          sourceExt: ext,
          converted: true,
          convertMode: "fallback",
          warningMessage,
          errorCode,
          convertDiagnostics: {
            ...hwpDiagnosticPayload,
            pageCount: Number(fallbackValidation.pageCount || 0),
            pageTargetMatched
          },
          sheetMap: [],
          pdfPath: fallbackCache,
          pdfBuffer: await fsp.readFile(fallbackCache)
        };
      }
      try {
        await fsp.unlink(fallbackCache);
      } catch (_error) {
        // ignore stale cache cleanup failure
      }
    }
    const extractedText = await extractDocumentTextForFallback(resolved, ext);
    if (!String(extractedText || "").trim()) {
      const error = new Error("hwp_no_readable_content");
      error.code = "HWP_CONTENT_UNREADABLE";
      error.publicMessage = "한글 내용 복원 실패";
      throw error;
    }
    sendStageStatus("변환", 90, "호환 보기로 전환합니다.", {
      sourceExt: ext,
      messageKey: "convert.step.compatibility"
    });
    const html = buildFallbackHtml({
      sourcePath: resolved,
      sourceExt: ext,
      contentText: extractedText || "",
      warningMessage
    });
    const basePrintOptions = {
      printBackground: true,
      pageSize: "A4",
      preferCSSPageSize: true
    };
    const fallbackCalibrated = await calibratePdfByPageTarget({
      renderPdf: (printToPdfOptions) => htmlToPdfBuffer(html, printToPdfOptions),
      basePrintOptions,
      targetPageCount: Number(forcedPageTarget || 0),
      maxAttempts: Number(forcedPageTarget || 0) > 0 ? 12 : 1
    });
    const fallbackPdf =
      fallbackCalibrated?.buffer && Buffer.isBuffer(fallbackCalibrated.buffer) && fallbackCalibrated.buffer.length > 0
        ? fallbackCalibrated.buffer
        : await htmlToPdfBuffer(html, basePrintOptions);
    await fsp.writeFile(fallbackPdfPath, fallbackPdf);
    const fallbackPageCount = Number(fallbackCalibrated?.pageCount || 0);
    const pageTargetMatched = Number(forcedPageTarget || 0) > 0
      ? fallbackPageCount === Number(forcedPageTarget || 0)
      : true;
    sendStageStatus("표시", 100, "호환 보기 문서를 표시합니다.", {
      sourceExt: ext,
      selectedMode: "fallback-text"
    });
    return {
      sourcePath: resolved,
      sourceExt: ext,
      converted: true,
      convertMode: "fallback",
      warningMessage,
      errorCode,
      convertDiagnostics: {
        ...hwpDiagnosticPayload,
        pageCount: fallbackPageCount,
        pageTargetMatched,
        calibrated: Boolean(fallbackCalibrated?.calibrated),
        calibrationAttempts: Number(fallbackCalibrated?.attempts || 1),
        calibrationScale: Number(fallbackCalibrated?.finalScale || 1)
      },
      sheetMap: [],
      pdfPath: fallbackPdfPath,
      pdfBuffer: fallbackPdf
    };
  }

  if (SPREADSHEET_EXTENSIONS.has(ext)) {
    const xlsxHtmlPdfPath = createConvertedPdfPath(resolved, "xlsx-html");
    const xlsxMapPath = `${xlsxHtmlPdfPath}.sheets.json`;
    const xlsxHtmlCache = readCachedPdf(xlsxHtmlPdfPath);
    if (xlsxHtmlCache) {
      const xlsxValidation = await validateConvertedPdfPath(xlsxHtmlCache);
      let cachedSheets = null;
      try { cachedSheets = JSON.parse(await fsp.readFile(xlsxMapPath, "utf8")); } catch (_) {}
      if (xlsxValidation.ok && cachedSheets?.pageCount === xlsxValidation.pageCount && Array.isArray(cachedSheets.sheetMap) && cachedSheets.sheetMap.length) {
        sendStageStatus("표시", 100, "엑셀 시트 보기 문서를 표시합니다.", {
          sourceExt: ext,
          selectedMode: "xlsx-html",
          pageCount: xlsxValidation.pageCount
        });
        return {
          sourcePath: resolved,
          sourceExt: ext,
          converted: true,
          convertMode: "xlsx-html",
          warningMessage: cachedSheets.warningMessage || "표 보기",
          errorCode: "",
          convertDiagnostics: {
            selectedMode: "xlsx-html",
            validationPassed: true,
            pageCount: xlsxValidation.pageCount,
            candidateScores: [],
            fallbackReason: warningMessage || "",
            spreadsheet: cachedSheets.diagnostics || null
          },
          sheetMap: cachedSheets.sheetMap,
          pdfPath: xlsxHtmlCache,
          pdfBuffer: await fsp.readFile(xlsxHtmlCache)
        };
      }
      try {
        await fsp.unlink(xlsxHtmlCache);
      } catch (_error) {
        // ignore invalid xlsx fallback cache cleanup failure
      }
    }
    if (errorCode === "EMPTY_DOCUMENT") {
      const error = new Error("표시할 셀 내용이 없습니다.");
      error.code = "EMPTY_DOCUMENT";
      throw error;
    }
    {
      try {
        sendStageStatus("변환", 88, "엑셀 시트 보기 문서를 생성하고 있습니다.", {
          sourceExt: ext
        });
        const xlsxFallback = buildSpreadsheetFallbackHtml({
          sourcePath: resolved,
          sourceExt: ext,
          warningMessage: warningMessage || "고품질 변환 엔진을 찾지 못해 표 보기 모드로 전환했습니다."
        });
        if (xlsxFallback?.html) {
          const renderedSpreadsheet = await renderSpreadsheetCompatibilityPdf(xlsxFallback, renderSpreadsheetHtmlPdf, PDFDocument);
          const xlsxPdf = renderedSpreadsheet.pdfBuffer;
          await fsp.writeFile(xlsxHtmlPdfPath, xlsxPdf);
          const xlsxValidation = await validateConvertedPdfPath(xlsxHtmlPdfPath);
          if (!xlsxValidation.ok) {
            throw new Error(`xlsx_html_invalid_pdf:${xlsxValidation.reason}`);
          }
          await fsp.writeFile(xlsxMapPath, JSON.stringify({
            pageCount: xlsxValidation.pageCount,
            sheetMap: renderedSpreadsheet.sheetMap,
            warningMessage: xlsxFallback.warningMessage || "표 보기",
            diagnostics: xlsxFallback.diagnostics || null
          }), "utf8");
          sendStageStatus("표시", 100, "엑셀 시트 보기 문서를 표시합니다.", {
            sourceExt: ext,
            selectedMode: "xlsx-html",
            pageCount: xlsxValidation.pageCount
          });
          return {
            sourcePath: resolved,
            sourceExt: ext,
            converted: true,
            convertMode: "xlsx-html",
            warningMessage: xlsxFallback.warningMessage || "표 보기",
            errorCode: "",
            convertDiagnostics: {
              selectedMode: "xlsx-html",
              validationPassed: true,
              pageCount: xlsxValidation.pageCount,
              candidateScores: [],
              fallbackReason: warningMessage || "표 보기",
              spreadsheet: xlsxFallback.diagnostics || null
            },
            sheetMap: renderedSpreadsheet.sheetMap,
            pdfPath: xlsxHtmlPdfPath,
            pdfBuffer: xlsxPdf
          };
        }
        const emptyError = new Error("표시할 셀 내용이 없습니다.");
        emptyError.code = "EMPTY_DOCUMENT";
        throw emptyError;
      } catch (error) {
        await fsp.unlink(xlsxHtmlPdfPath).catch(() => {});
        await fsp.unlink(xlsxMapPath).catch(() => {});
        // Never turn a truncated or failed table into a successful text preview.
        throw error;
      }
    }
  }

  const fallbackPdfPath = createConvertedPdfPath(resolved, "fallback");
  const fallbackCache = readCachedPdf(fallbackPdfPath, { maxAgeMs: FALLBACK_CACHE_TTL_MS });
  if (fallbackCache) {
    const fallbackValidation = await validateConvertedPdfPath(fallbackCache);
    if (fallbackValidation.ok) {
      sendStageStatus("표시", 100, "호환 보기 문서를 표시합니다.", {
        sourceExt: ext,
        selectedMode: "fallback-text",
        pageCount: fallbackValidation.pageCount
      });
      return {
        sourcePath: resolved,
        sourceExt: ext,
        converted: true,
        convertMode: "fallback",
        warningMessage,
        errorCode,
        convertDiagnostics: {
          selectedMode: "fallback-text",
          validationPassed: true,
          pageCount: fallbackValidation.pageCount,
          candidateScores: [],
          fallbackReason: warningMessage || ""
        },
        sheetMap: [],
        pdfPath: fallbackCache,
        pdfBuffer: await fsp.readFile(fallbackCache)
      };
    }
    try {
      await fsp.unlink(fallbackCache);
    } catch (_error) {
      // ignore stale fallback cache cleanup failure
    }
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
  const fallbackValidation = await validateConvertedPdfPath(fallbackPdfPath);
  if (!fallbackValidation.ok) {
    throw new Error("INVALID_CONVERTED_PDF");
  }
  sendStageStatus("표시", 100, "호환 보기 문서를 표시합니다.", {
    sourceExt: ext,
    selectedMode: "fallback-text",
    pageCount: fallbackValidation.pageCount
  });

  return {
    sourcePath: resolved,
    sourceExt: ext,
    converted: true,
    convertMode: "fallback",
    warningMessage,
    errorCode,
    convertDiagnostics: {
      selectedMode: "fallback-text",
      validationPassed: true,
      pageCount: fallbackValidation.pageCount,
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
        {
          label: mt("hwpPrintFallback"),
          type: "checkbox",
          checked: false,
          enabled: false,
          click: () => {}
        },
        { type: "separator" },
        {
          label: mt("updateCheck"),
          enabled: !updateBusy,
          click: () => sendToRenderer("menu-action", "check-update")
        },
        {
          label: mt("copySupportLink"),
          click: () => sendToRenderer("menu-action", "copy-support-link")
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
    minWidth: 880,
    minHeight: 560,
    title: "lookup",
    show: false,
    autoHideMenuBar: false,
    backgroundColor: "#f3f5f8",
    icon: iconPath,
    webPreferences: {
      preload: path.join(APP_ROOT, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webviewTag: false
    }
  });

  mainWindow.loadFile(path.join(APP_ROOT, "src", "renderer", "index.html"));

  mainWindow.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  mainWindow.webContents.on("will-navigate", (event, url) => {
    const normalized = String(url || "").trim().toLowerCase();
    if (!normalized.startsWith("file://")) {
      event.preventDefault();
    }
  });

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
  const pkgPath = path.join(APP_ROOT, "package.json");
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
    path.join(APP_ROOT, "update-config.json"),
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
      { name: "지원 문서", extensions: ["pdf", "hwp", "hwpx", "doc", "docx", "xls", "xlsx", "xlsm", "ppt", "pptx", "accdb", "html", "htm"] },
      { name: "PDF 문서", extensions: ["pdf"] },
      { name: "한글 문서", extensions: ["hwp", "hwpx"] },
      { name: "Word 문서", extensions: ["doc", "docx"] },
      { name: "Excel 문서", extensions: ["xls", "xlsx", "xlsm"] },
      { name: "PowerPoint 문서", extensions: ["ppt", "pptx"] },
      { name: "Access 데이터베이스", extensions: ["accdb"] },
      { name: "웹/Notion HTML", extensions: ["html", "htm"] }
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

ipcMain.handle("dialog:confirm-document-switch", async () => {
  const result = await dialog.showMessageBox(mainWindow, {
    type: "question", title: "저장하지 않은 변경", message: "문서 변경 사항 저장",
    buttons: ["저장", "저장 안 함", "취소"], defaultId: 0, cancelId: 2
  });
  return ["save", "discard", "cancel"][result.response] || "cancel";
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

ipcMain.handle("document:open", async (_event, requestPayload) => {
  let operationId = "";
  try {
    await ensureFeedbackGate();
    const request = parseDocumentRequestPayload(requestPayload);
    const filePath = request.filePath;
    operationId = createConvertOperationId("open");
    beginConvertOperation(operationId);
    sendConvertStatus("엔진 확인", 2, "문서를 열 준비를 하고 있습니다.", {
      sourceExt: getFileExtension(filePath || ""),
      messageKey: "convert.step.prepare-open",
      operationId
    });
    const converted = await convertDocumentToPdf(filePath, {
      ...(request.options || {}),
      operationId
    });
    const publicConvertMode = normalizePublicConvertMode(converted.convertMode, converted.converted);
    if (HWP_EXTENSIONS.has(String(converted.sourceExt || "").toLowerCase())) {
      await recordFeedbackSolution({
        key: buildFeedbackKey("hwp-compat-structure", converted.sourcePath, converted.convertMode),
        title: "HWP 구조 복원 경로 적용",
        changesMade: "hwpjs JSON 구조를 직접 읽어 목록/표/강조 복원 규칙을 우선 적용했습니다.",
        qualityGain: "목록 분리와 표 폭 복원 정확도를 높이고 외부 HWP 경유를 차단했습니다.",
        satisfactionPoint: "비설치 환경에서도 원본 구조감에 가까운 열람 품질을 목표로 고정했습니다."
      });
    }
    return {
      ok: true,
      sourcePath: converted.sourcePath,
      sourceExt: converted.sourceExt,
      converted: converted.converted,
      convertMode: publicConvertMode,
      warningMessage: converted.warningMessage,
      errorCode: converted.errorCode || "",
      convertDiagnostics: normalizeConvertDiagnostics(
        converted.convertDiagnostics,
        converted.convertMode || "fallback",
        converted.warningMessage || ""
      ),
      sheetMap: Array.isArray(converted.sheetMap) ? converted.sheetMap : [],
      pdfPath: converted.pdfPath,
      data: converted.pdfBuffer
    };
  } catch (error) {
    const errorCode = classifyDocumentOpenError(error);
    const publicMessage = String(error?.publicMessage || (errorCode === "EMPTY_DOCUMENT" ? "빈 문서" : ""));
    const request = parseDocumentRequestPayload(requestPayload);
    await recordFeedbackProblem({
      key: buildFeedbackKey("document-open-error", request.filePath, errorCode),
      title: "문서 열기 실패 재발 기록",
      requestSummary: `document:open (${String(request.filePath || "").slice(0, 220)})`,
      issueSummary: `${errorCode}: ${publicMessage || "문서를 열지 못했습니다."}`,
      rootCause: buildDiagnosticDetail(error),
      improvementPlan: "변환 엔진 상태와 복원 규칙을 다시 점검하고, 안전한 호환 보기로 즉시 전환합니다."
    });
    return {
      ok: false,
      errorCode,
      message: publicMessage,
      diagnosticDetail: buildDiagnosticDetail(error)
    };
  } finally {
    endConvertOperation(operationId);
  }
});

let webOpenBusy = false;
async function openWebDocument(rawUrl) {
  if (webOpenBusy) return { ok: false, errorCode: "WEB_BUSY", message: "웹페이지 처리 중" };
  let operationId = "";
  let pdfPath = "";
  try {
    const sourceUrl = normalizeWebSourceUrl(rawUrl);
    webOpenBusy = true;
    operationId = createConvertOperationId("web");
    beginConvertOperation(operationId);
    const status = (stage, percent, messageKey) => sendConvertStatus(stage, percent, "", {
      operationId, sourceExt: ".html", selectedMode: "web-html", messageKey
    });
    status("변환", 12, "convert.step.web-load");
    let result;
    let detailOnly = false;
    const printOptions = { sourceType: "web-html", normalizeScrollableContent: true, allowRasterFallback: true };
    try { result = await renderHtmlToPdfBuffer({ sourceUrl, printOptions }); }
    catch (error) {
      const detailUrl = ["WEB_ACCESS_CHALLENGE", "WEB_LOAD_TIMEOUT"].includes(error?.code) ? getPublicProductDetailUrl(sourceUrl) : "";
      if (!detailUrl) throw error;
      status("변환", 46, "convert.step.web-analyze");
      result = await renderHtmlToPdfBuffer({ sourceUrl: detailUrl, printOptions });
      detailOnly = true;
      result.diagnostics.captureUrl = detailUrl;
      result.diagnostics.warningMessage = ["상품 상세설명 · 상단 정보 미포함", result.diagnostics.warningMessage].filter(Boolean).join(" / ");
    }
    status("검증", 88, "convert.step.validate");
    const doc = await PDFDocument.load(result.pdfBuffer);
    if (!doc.getPageCount()) throw new Error("표시할 내용 없음");
    const captureContent = await inspectPdfCaptureContent(result.pdfBuffer);
    if (captureContent.placeholderOnly || (captureContent.checked && captureContent.blankPages.length === captureContent.pageCount)) {
      throw new Error("빈 페이지 감지 · 다시 시도");
    }
    const title = detailOnly ? "G마켓 상품 상세설명" : String(result.diagnostics?.title || new URL(sourceUrl).hostname);
    const name = getWebPdfName(title, sourceUrl);
    // Keep each result separate: a refresh must never reuse an incomplete prior capture.
    const captureDir = path.join(ensureConvertedDir(), `web-${operationId}`);
    await fsp.mkdir(captureDir, { recursive: true });
    pdfPath = path.join(captureDir, name);
    await fsp.writeFile(pdfPath, result.pdfBuffer);
    status("표시", 100, "convert.step.display");
    return {
      ok: true, sourcePath: pdfPath, sourceUrl, title, sourceExt: ".html", converted: true,
      convertMode: "web-html", warningMessage: result.diagnostics?.warningMessage || "",
      errorCode: "", sheetMap: [], pdfPath, data: result.pdfBuffer,
      convertDiagnostics: normalizeConvertDiagnostics({
        selectedMode: "web-html", validationPassed: true, pageCount: doc.getPageCount(),
        sourceType: "web-html", webPage: result.diagnostics
      }, "web-html")
    };
  } catch (error) {
    if (pdfPath) await fsp.unlink(pdfPath).catch(() => {});
    return {
      ok: false, errorCode: error?.code || "WEB_CAPTURE_FAILED", message: error?.message || "웹페이지 저장 실패",
      diagnosticDetail: String(error?.stack || error || "").slice(0, 1800)
    };
  } finally {
    if (operationId) { endConvertOperation(operationId); webOpenBusy = false; }
  }
}
ipcMain.handle("document:open-web", (_event, rawUrl) => openWebDocument(rawUrl));

ipcMain.handle("document:convert", async (_event, requestPayload) => {
  let operationId = "";
  try {
    await ensureFeedbackGate();
    const request = parseDocumentRequestPayload(requestPayload);
    const filePath = request.filePath;
    operationId = createConvertOperationId("convert");
    beginConvertOperation(operationId);
    sendConvertStatus("엔진 확인", 2, "문서를 변환할 준비를 하고 있습니다.", {
      sourceExt: getFileExtension(filePath || ""),
      messageKey: "convert.step.prepare-convert",
      operationId
    });
    const converted = await convertDocumentToPdf(filePath, {
      ...(request.options || {}),
      operationId
    });
    const publicConvertMode = normalizePublicConvertMode(converted.convertMode, converted.converted);
    if (HWP_EXTENSIONS.has(String(converted.sourceExt || "").toLowerCase())) {
      await recordFeedbackSolution({
        key: buildFeedbackKey("hwp-convert-quality", converted.sourcePath, converted.convertMode),
        title: "HWP 변환 품질 개선 적용",
        changesMade: "HWP JSON 기반 구조 복원 후보를 추가하고 목록/표 규칙을 강화했습니다.",
        qualityGain: "목록 줄바꿈/표 칸 정렬/강조 표현을 기존보다 안정적으로 재현합니다.",
        satisfactionPoint: "변환 진행 흐름과 결과 품질을 함께 관리할 수 있게 바꿨습니다."
      });
    }
    return {
      ok: true,
      sourcePath: converted.sourcePath,
      sourceExt: converted.sourceExt,
      converted: converted.converted,
      convertMode: publicConvertMode,
      warningMessage: converted.warningMessage,
      errorCode: converted.errorCode || "",
      convertDiagnostics: normalizeConvertDiagnostics(
        converted.convertDiagnostics,
        converted.convertMode || "fallback",
        converted.warningMessage || ""
      ),
      sheetMap: Array.isArray(converted.sheetMap) ? converted.sheetMap : [],
      pdfPath: converted.pdfPath
    };
  } catch (error) {
    const errorCode = classifyDocumentOpenError(error);
    const publicMessage = String(error?.publicMessage || (errorCode === "EMPTY_DOCUMENT" ? "빈 문서" : ""));
    const request = parseDocumentRequestPayload(requestPayload);
    await recordFeedbackProblem({
      key: buildFeedbackKey("document-convert-error", request.filePath, errorCode),
      title: "문서 변환 실패 재발 기록",
      requestSummary: `document:convert (${String(request.filePath || "").slice(0, 220)})`,
      issueSummary: `${errorCode}: ${publicMessage || "문서 변환에 실패했습니다."}`,
      rootCause: buildDiagnosticDetail(error),
      improvementPlan: "실패 유형을 누적하고, 동일 조건 재발 시 즉시 우회 경로를 우선 적용합니다."
    });
    return {
      ok: false,
      errorCode,
      message: publicMessage,
      diagnosticDetail: buildDiagnosticDetail(error)
    };
  } finally {
    endConvertOperation(operationId);
  }
});

ipcMain.handle("document:is-supported", async (_event, filePath) => {
  if (!filePath || typeof filePath !== "string") {
    return false;
  }
  return SUPPORTED_DOCUMENT_EXTENSIONS.has(getFileExtension(filePath));
});

ipcMain.handle("document:resolve-drop-path", async (_event, payload) => {
  const filePath = resolveDropPathPayload(payload);
  if (!filePath) {
    return {
      ok: false,
      filePath: "",
      reason: "drop_path_unreadable"
    };
  }
  return {
    ok: true,
    filePath
  };
});

ipcMain.handle("clipboard:copy-text", async (_event, text) => {
  clipboard.writeText(String(text || ""));
  return true;
});

ipcMain.handle("settings:get", async () => {
  return {
    language: appSettings.language,
    allowPrintStyleHwpFallback: false,
    hwpPrintFallbackExplicitlySet: false,
    hwpNativeRetryPolicy: appSettings.hwpNativeRetryPolicy || "quality-first-single-native-try"
  };
});

ipcMain.handle("settings:set-language", async (_event, language) => {
  appSettings.language = language === "en" ? "en" : "ko";
  appSettings.allowPrintStyleHwpFallback = false;
  appSettings.hwpPrintFallbackExplicitlySet = false;
  appSettings.printFallbackUserApprovedAtV132 = false;
  saveSettings();
  createMenu();
  return {
    language: appSettings.language,
    allowPrintStyleHwpFallback: false,
    hwpPrintFallbackExplicitlySet: false,
    hwpNativeRetryPolicy: appSettings.hwpNativeRetryPolicy || "quality-first-single-native-try"
  };
});

ipcMain.handle("settings:set-hwp-print-fallback", async (_event, enabled) => {
  void enabled;
  appSettings.allowPrintStyleHwpFallback = false;
  appSettings.hwpPrintFallbackExplicitlySet = false;
  appSettings.printFallbackUserApprovedAtV132 = false;
  saveSettings();
  createMenu();
  return {
    language: appSettings.language,
    allowPrintStyleHwpFallback: false,
    hwpPrintFallbackExplicitlySet: false,
    hwpNativeRetryPolicy: appSettings.hwpNativeRetryPolicy || "quality-first-single-native-try"
  };
});

ipcMain.handle("pdf:write-file", async (_event, payload) => {
  const filePath = payload?.filePath;
  if (typeof filePath !== "string" || !filePath.trim()) {
    throw new Error("저장 경로가 비어 있습니다.");
  }
  const resolvedTarget = path.resolve(filePath);
  if (getFileExtension(resolvedTarget) !== ".pdf") {
    throw new Error("PDF 파일로만 저장할 수 있습니다.");
  }
  const buffer = toBuffer(payload?.data);
  await fsp.writeFile(resolvedTarget, buffer);
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

ipcMain.handle("app:get-version", async () => APP_RELEASE_VERSION);

if (gotSingleInstanceLock) {
  app.whenReady().then(() => {
    void ensureFeedbackGate();
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










