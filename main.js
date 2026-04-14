const { app, BrowserWindow, Menu, dialog, ipcMain, clipboard, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { spawnSync } = require("node:child_process");
const { parseOffice } = require("officeparser");
const WordExtractor = require("word-extractor");
const XLSX = require("xlsx");
const hwpjs = require("@ohah/hwpjs");
const { read: readHwpxDocument } = require("hwpx-js");

let mainWindow = null;
let pendingDocumentToOpen = null;
let updateConfig = null;
let updateDownloaded = false;
let updateTargetVersion = "";
let updateInstallTriggered = false;
let appSettings = { language: "ko" };

const menuText = {
  ko: {
    file: "파일",
    open: "열기...",
    saveAs: "다른 이름 저장...",
    saveOverwrite: "원본 덮어쓰기 저장",
    print: "인쇄 미리보기...",
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
    print: "Print Preview...",
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

function createConvertedPdfPath(sourcePath) {
  const resolved = path.resolve(sourcePath);
  const stat = fs.statSync(resolved);
  const ext = getFileExtension(resolved).replace(".", "");
  const baseName = path.basename(resolved, path.extname(resolved)).replace(/[^\\w.\\-가-힣]/g, "_").slice(0, 56);
  const signature = `${resolved}|${stat.mtimeMs}|${stat.size}`;
  const digest = crypto.createHash("sha1").update(signature).digest("hex").slice(0, 12);
  return path.join(ensureConvertedDir(), `${baseName}-${ext}-${digest}.pdf`);
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

  return runPowerShell(script);
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

  return runPowerShell(script);
}

function convertWithHancomCom(sourcePath, outputPdfPath) {
  const source = escapeForPowerShell(sourcePath);
  const output = escapeForPowerShell(outputPdfPath);
  const script = [
    "$ErrorActionPreference = 'Stop'",
    "$hwp = $null",
    "try {",
    "  $hwp = New-Object -ComObject HWPFrame.HwpObject",
    "  try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') | Out-Null } catch {}",
    "  $hwp.XHwpWindows.Item(0).Visible = $false",
    "  $opened = $hwp.Open(" + source + ")",
    "  if (-not $opened) { throw 'hwp_open_failed' }",
    "  $saved = $hwp.SaveAs(" + output + ", 'PDF')",
    "  if (-not $saved) { throw 'hwp_save_failed' }",
    "} finally {",
    "  if ($hwp -ne $null) { try { $hwp.Quit() | Out-Null } catch {} }",
    "}",
    "if (-not (Test-Path " + output + ")) { throw 'hwp_export_failed' }"
  ].join("; ");
  return runPowerShell(script);
}

async function extractWordLegacyText(filePath) {
  const extractor = new WordExtractor();
  const document = await extractor.extract(filePath);
  return normalizeText(document?.getBody?.() || "");
}

async function extractOfficeAstText(filePath) {
  try {
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
    const raw = fs.readFileSync(sourcePath);
    const htmlRaw = hwpjs.toHtml(raw);
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
  if (ext === ".doc" || ext === ".docx") {
    return convertWithWordCom(sourcePath, outputPdfPath);
  }
  if (ext === ".xls" || ext === ".xlsx") {
    return convertWithExcelCom(sourcePath, outputPdfPath);
  }
  return { ok: false, stderr: "unsupported_office_ext" };
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
        convertMode: "office",
        warningMessage: "",
        pdfPath: targetPdfPath,
        pdfBuffer: await fsp.readFile(targetPdfPath)
      };
    }
    warningMessage = "고품질 변환 엔진을 찾지 못해 텍스트 기반 보기 모드로 변환했습니다.";
  } else if (HWP_EXTENSIONS.has(ext)) {
    const hwpResult = convertWithHancomCom(resolved, targetPdfPath);
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
    warningMessage = "원본 레이아웃 변환에 실패해 텍스트 기반 보기로 전환했습니다.";
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

  mainWindow.webContents.on("did-finish-load", () => {
    if (pendingDocumentToOpen) {
      sendToRenderer("system-open-file", pendingDocumentToOpen);
      pendingDocumentToOpen = null;
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
    sendUpdateStatus("disabled", {
      stage: "disabled",
      message: "업데이트 설정이 없습니다. update-config.json에 owner/repo를 입력해 주세요."
    });
    return;
  }

  updateDownloaded = false;
  updateTargetVersion = "";
  updateInstallTriggered = false;
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
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
    sendUpdateStatus("checking", { stage: "checking", message: "업데이트 확인중..." });
  });

  autoUpdater.on("update-available", (info) => {
    updateTargetVersion = String(info?.version || "");
    sendUpdateStatus("available", {
      stage: "downloading",
      targetVersion: updateTargetVersion,
      percent: 0,
      message: `새 버전 v${updateTargetVersion}을 찾았습니다. 다운로드를 시작합니다.`
    });
  });

  autoUpdater.on("update-not-available", () => {
    updateTargetVersion = "";
    sendUpdateStatus("not-available", { stage: "idle", percent: 100, message: "현재 최신 버전입니다." });
  });

  autoUpdater.on("download-progress", (progress) => {
    const percent = Math.max(0, Math.min(100, Number(progress?.percent || 0)));
    sendUpdateStatus("downloading", {
      stage: "downloading",
      targetVersion: updateTargetVersion,
      message: `업데이트 다운로드중... ${Math.round(percent)}%`,
      percent
    });
  });

  autoUpdater.on("update-downloaded", (info) => {
    updateDownloaded = true;
    updateInstallTriggered = false;
    updateTargetVersion = String(info?.version || updateTargetVersion || "");
    sendUpdateStatus("downloaded", {
      stage: "ready-to-install",
      targetVersion: updateTargetVersion,
      percent: 100,
      message: `업데이트 v${updateTargetVersion} 다운로드 완료. 설치를 준비합니다.`
    });

    if (app.isPackaged) {
      const triggerInstall = () => {
        if (updateInstallTriggered) {
          return;
        }
        try {
          sendUpdateStatus("installing", {
            stage: "installing",
            targetVersion: updateTargetVersion,
            percent: 100,
            message: "앱을 다시 시작해 업데이트를 설치합니다..."
          });
          autoUpdater.quitAndInstall(false, true);
          updateInstallTriggered = true;
        } catch (error) {
          updateInstallTriggered = false;
          sendUpdateStatus("error", {
            stage: "error",
            message: `업데이트 재시작 설치 실패: ${error?.message || "알 수 없는 오류"}`
          });
        }
      };
      setTimeout(triggerInstall, 1200);
      setTimeout(() => {
        if (!updateInstallTriggered) {
          triggerInstall();
        }
      }, 4200);
    }
  });

  autoUpdater.on("before-quit-for-update", () => {
    sendUpdateStatus("installing", {
      stage: "installing",
      targetVersion: updateTargetVersion,
      percent: 100,
      message: "업데이트 설치를 위해 앱을 종료합니다..."
    });
  });

  autoUpdater.on("error", (error) => {
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

  const converted = await convertDocumentToPdf(filePath);
  return converted.pdfBuffer;
});

ipcMain.handle("document:open", async (_event, filePath) => {
  const converted = await convertDocumentToPdf(filePath);
  return {
    sourcePath: converted.sourcePath,
    sourceExt: converted.sourceExt,
    converted: converted.converted,
    convertMode: converted.convertMode,
    warningMessage: converted.warningMessage,
    pdfPath: converted.pdfPath,
    data: converted.pdfBuffer
  };
});

ipcMain.handle("document:convert", async (_event, filePath) => {
  const converted = await convertDocumentToPdf(filePath);
  return {
    sourcePath: converted.sourcePath,
    sourceExt: converted.sourceExt,
    converted: converted.converted,
    convertMode: converted.convertMode,
    warningMessage: converted.warningMessage,
    pdfPath: converted.pdfPath
  };
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
  try {
    await autoUpdater.checkForUpdates();
    return { ok: true };
  } catch (error) {
    return { ok: false, message: error?.message || "업데이트 확인 실패" };
  }
});

ipcMain.handle("update:get-config", async () => {
  return {
    enabled: Boolean(updateConfig),
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








