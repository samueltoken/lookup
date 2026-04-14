
import * as pdfjsLib from "../../node_modules/pdfjs-dist/legacy/build/pdf.mjs";
import { PDFDocument, StandardFonts, degrees, rgb } from "../../node_modules/pdf-lib/dist/pdf-lib.esm.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs",
  import.meta.url
).toString();

let tesseractCreateWorker = null;

async function ensureTesseractCreateWorker() {
  if (typeof tesseractCreateWorker === "function") {
    return tesseractCreateWorker;
  }
  const module = await import("../../node_modules/tesseract.js/dist/tesseract.esm.min.js");
  if (typeof module?.createWorker !== "function") {
    throw new Error("OCR 모듈을 불러오지 못했습니다.");
  }
  tesseractCreateWorker = module.createWorker;
  return tesseractCreateWorker;
}

const storage = {
  darkMode: "lookup-dark-mode",
  language: "lookup-language",
  fullscreenMode: "lookup-fullscreen-view-mode",
  leftPanelWidth: "lookup-left-panel-width",
  rightPanelWidth: "lookup-right-panel-width",
  leftPanelVisible: "lookup-left-panel-visible",
  rightPanelVisible: "lookup-right-panel-visible"
};

function getStoredNumber(key, fallback) {
  const raw = Number.parseFloat(localStorage.getItem(key) || "");
  return Number.isFinite(raw) ? raw : fallback;
}

function getStoredBool(key, fallback) {
  const raw = localStorage.getItem(key);
  if (raw === "1") {
    return true;
  }
  if (raw === "0") {
    return false;
  }
  return fallback;
}

const i18n = {
  ko: {
    open: "열기",
    saveAs: "다른 이름 저장",
    saveOverwrite: "덮어쓰기",
    prev: "이전",
    next: "다음",
    zoomReset: "원래 크기",
    rotateLeft: "왼쪽 회전",
    rotateRight: "오른쪽 회전",
    deletePage: "페이지 삭제",
    modeView: "보기",
    modeHighlight: "형광펜",
    modePen: "펜",
    modeText: "텍스트 메모",
    undo: "되돌리기",
    redo: "다시하기",
    thumbToggleShow: "미리보기 표시",
    thumbToggleHide: "미리보기 숨기기",
    searchPanelToggleShow: "검색 패널 표시",
    searchPanelToggleHide: "검색 패널 숨기기",
    fullscreen: "전체화면",
    fullscreenExit: "전체화면 종료",
    darkMode: "다크모드",
    lightMode: "라이트모드",
    searchPlaceholder: "문서 검색",
    search: "검색",
    prevHit: "이전 결과",
    nextHit: "다음 결과",
    checkUpdate: "업데이트 확인",
    thumbPanelTitle: "미리보기 (드래그로 순서 변경)",
    dropHint: "문서를 끌어놓거나 열기 버튼을 눌러 주세요.",
    fullscreenModeContinuous: "전체화면: 연속 스크롤",
    fullscreenModeSingle: "전체화면: 현재 페이지",
    searchResults: "검색 결과",
    copiedContact: "개발자 이메일이 복사되었습니다.",
    languageChangedKo: "언어가 한국어로 변경되었습니다.",
    languageChangedEn: "언어가 English로 변경되었습니다.",
    updateStarted: "업데이트 확인을 시작했습니다.",
    versionCurrent: "현재 버전",
    versionTarget: "대상 버전",
    searchCount: "결과 {count}개",
    searchPanelCount: "{count}건",
    pageLabel: "페이지",
    updateReady: "업데이트 연동 준비됨",
    updateDisabled: "업데이트 비활성: 저장소 설정을 찾지 못했습니다.",
    updateChecking: "업데이트를 확인하고 있습니다...",
    printPreparing: "인쇄 미리보기를 준비하고 있습니다...",
    printOpened: "인쇄 미리보기를 열었습니다.",
    printFailed: "인쇄 미리보기를 열지 못했습니다.",
    ocrSearching: "이미지 문서에서 글자를 찾고 있습니다...",
    ocrFailed: "OCR 인식에 실패해 텍스트 검색만 사용합니다.",
    textMemoHint: "메모 입력",
    textMemoAdd: "추가",
    textMemoCancel: "취소",
    textMemoAdded: "텍스트 메모를 추가했습니다.",
    undoDone: "되돌리기 완료",
    redoDone: "다시하기 완료",
    updateStageIdle: "업데이트 대기",
    updateStageChecking: "업데이트 확인중",
    updateStageDownloading: "다운로드중",
    updateStageReady: "설치 준비",
    updateStageRestarting: "재시작 중",
    updateStageInstalled: "업데이트 완료",
    updateStageError: "업데이트 오류",
    updateBusyMessage: "업데이트가 이미 진행 중입니다.",
    updateNotesTitle: "업데이트 완료",
    updateNotesFallback: "변경 사항 요약을 불러오지 못했습니다.",
    updateNotesClose: "확인",
    openErrorNotFound: "파일을 찾을 수 없습니다. 경로를 확인해 주세요.",
    openErrorPermission: "파일 권한이 없어 열 수 없습니다.",
    openErrorUnsupported: "지원하지 않는 파일 형식입니다.",
    openErrorConvert: "문서 변환에 실패했습니다. 파일을 다시 확인해 주세요.",
    openErrorGeneric: "문서를 열지 못했습니다.",
    openErrorPdfFallback: "일반 열기에 실패해 PDF 직접 열기로 복구했습니다."
  },
  en: {
    open: "Open",
    saveAs: "Save As",
    saveOverwrite: "Overwrite",
    prev: "Prev",
    next: "Next",
    zoomReset: "Reset",
    rotateLeft: "Rotate Left",
    rotateRight: "Rotate Right",
    deletePage: "Delete Page",
    modeView: "View",
    modeHighlight: "Highlight",
    modePen: "Pen",
    modeText: "Text Note",
    undo: "Undo",
    redo: "Redo",
    thumbToggleShow: "Show Thumbnails",
    thumbToggleHide: "Hide Thumbnails",
    searchPanelToggleShow: "Show Search Panel",
    searchPanelToggleHide: "Hide Search Panel",
    fullscreen: "Fullscreen",
    fullscreenExit: "Exit Fullscreen",
    darkMode: "Dark Mode",
    lightMode: "Light Mode",
    searchPlaceholder: "Search document",
    search: "Search",
    prevHit: "Prev Hit",
    nextHit: "Next Hit",
    checkUpdate: "Check Update",
    thumbPanelTitle: "Thumbnails (drag to reorder)",
    dropHint: "Drop a document or click Open.",
    fullscreenModeContinuous: "Fullscreen: Continuous",
    fullscreenModeSingle: "Fullscreen: Single Page",
    searchResults: "Search Results",
    copiedContact: "Developer email copied to clipboard.",
    languageChangedKo: "Language switched to Korean.",
    languageChangedEn: "Language switched to English.",
    updateStarted: "Update check started.",
    versionCurrent: "Current Version",
    versionTarget: "Target Version",
    searchCount: "Results {count}",
    searchPanelCount: "{count} items",
    pageLabel: "page",
    updateReady: "Update connected",
    updateDisabled: "Update disabled: repository info not found.",
    updateChecking: "Checking for updates...",
    printPreparing: "Preparing print preview...",
    printOpened: "Print preview opened.",
    printFailed: "Unable to open print preview.",
    ocrSearching: "Scanning image pages with OCR...",
    ocrFailed: "OCR failed. Using text-layer search only.",
    textMemoHint: "Enter memo",
    textMemoAdd: "Add",
    textMemoCancel: "Cancel",
    textMemoAdded: "Text memo added.",
    undoDone: "Undo complete",
    redoDone: "Redo complete",
    updateStageIdle: "Idle",
    updateStageChecking: "Checking for updates",
    updateStageDownloading: "Downloading update",
    updateStageReady: "Preparing install",
    updateStageRestarting: "Restarting",
    updateStageInstalled: "Update complete",
    updateStageError: "Update error",
    updateBusyMessage: "An update is already in progress.",
    updateNotesTitle: "Update completed",
    updateNotesFallback: "Unable to load the release notes.",
    updateNotesClose: "Close",
    openErrorNotFound: "File not found. Please check the path.",
    openErrorPermission: "No permission to open this file.",
    openErrorUnsupported: "Unsupported file format.",
    openErrorConvert: "Failed to convert this document.",
    openErrorGeneric: "Unable to open this document.",
    openErrorPdfFallback: "Standard open failed. Recovered using direct PDF read."
  }
};

function t(key, vars = {}) {
  const dict = i18n[state.language] || i18n.ko;
  const template = dict[key] || i18n.ko[key] || key;
  return template.replace(/\{(\w+)\}/g, (_all, name) => String(vars[name] ?? ""));
}

const state = {
  pdfDoc: null,
  sourceBytes: null,
  filePath: "",
  sourceExt: ".pdf",
  sourceConverted: false,
  sourceConvertMode: "native",
  pageOrder: [],
  pageCache: new Map(),
  pageViews: new Map(),
  thumbnails: new Map(),
  pageRotations: new Map(),
  annotations: new Map(),
  textItemsCache: new Map(),
  searchPageCache: new Map(),
  ocrWordCache: new Map(),
  ocrPagePromises: new Map(),
  ocrWorker: null,
  ocrWorkerPromise: null,
  ocrWorkerReady: false,
  ocrErrorShown: false,
  scale: 1,
  currentPage: 1,
  editingMode: "view",
  drawing: null,
  searchQuery: "",
  searchMatches: [],
  perPageMatchIndexes: new Map(),
  activeSearchIndex: -1,
  renderVersion: 0,
  thumbRenderVersion: 0,
  isFullScreen: false,
  fullScreenViewMode: localStorage.getItem(storage.fullscreenMode) || "continuous",
  thumbPanelVisible: getStoredBool(storage.leftPanelVisible, true),
  searchPanelVisible: getStoredBool(storage.rightPanelVisible, false),
  fullscreenThumbVisible: false,
  fullScreenAutoFitDone: false,
  zoomMode: "fit",
  scrollRaf: 0,
  saveDirty: false,
  wheelZoomRaf: 0,
  wheelZoomApplying: false,
  wheelZoomDelta: 0,
  wheelZoomAnchor: null,
  singlePageWheelStepTime: 0,
  mainRenderQuality: 1.75,
  thumbRenderQuality: 2.6,
  leftPanelWidth: clamp(getStoredNumber(storage.leftPanelWidth, 250), 180, 560),
  rightPanelWidth: clamp(getStoredNumber(storage.rightPanelWidth, 280), 220, 620),
  activeResizer: null,
  fullRenderPassVersion: 0,
  fullRenderTimer: 0,
  layoutRecoveryTimer: 0,
  appVersion: "",
  updateTargetVersion: "",
  updateBusy: false,
  updateStage: "idle",
  updateBannerHideTimer: 0,
  layoutRecoveryToken: 0,
  language: localStorage.getItem(storage.language) === "en" ? "en" : "ko",
  applyingLanguage: false,
  pendingZoomJob: null,
  zoomJobRunning: false,
  viewerRenderRecoveryCount: 0,
  thumbRerenderTimer: 0,
  pageRenderTasks: new Map(),
  thumbRenderTasks: new Map(),
  historyPast: [],
  historyFuture: [],
  historyRestoring: false,
  historyLimit: 120,
  pendingTextMemo: null
};

const els = {
  workspace: document.getElementById("workspace"),
  openFileBtn: document.getElementById("openFileBtn"),
  saveAsBtn: document.getElementById("saveAsBtn"),
  saveOverwriteBtn: document.getElementById("saveOverwriteBtn"),
  printBtn: document.getElementById("printBtn"),
  prevPageBtn: document.getElementById("prevPageBtn"),
  nextPageBtn: document.getElementById("nextPageBtn"),
  pageInput: document.getElementById("pageInput"),
  pageCountLabel: document.getElementById("pageCountLabel"),
  zoomOutBtn: document.getElementById("zoomOutBtn"),
  zoomInBtn: document.getElementById("zoomInBtn"),
  zoomResetBtn: document.getElementById("zoomResetBtn"),
  zoomLabel: document.getElementById("zoomLabel"),
  rotateLeftBtn: document.getElementById("rotateLeftBtn"),
  rotateRightBtn: document.getElementById("rotateRightBtn"),
  deletePageBtn: document.getElementById("deletePageBtn"),
  editModeButtons: Array.from(document.querySelectorAll(".mode")),
  undoBtn: document.getElementById("undoBtn"),
  redoBtn: document.getElementById("redoBtn"),
  searchInput: document.getElementById("searchInput"),
  searchBtn: document.getElementById("searchBtn"),
  searchPrevBtn: document.getElementById("searchPrevBtn"),
  searchNextBtn: document.getElementById("searchNextBtn"),
  searchCountLabel: document.getElementById("searchCountLabel"),
  toggleThumbPanelBtn: document.getElementById("toggleThumbPanelBtn"),
  toggleSearchPanelBtn: document.getElementById("toggleSearchPanelBtn"),
  toggleFullscreenBtn: document.getElementById("toggleFullscreenBtn"),
  toggleDarkBtn: document.getElementById("toggleDarkBtn"),
  thumbPanel: document.getElementById("thumbPanel"),
  leftResizer: document.getElementById("leftResizer"),
  rightResizer: document.getElementById("rightResizer"),
  thumbnailList: document.getElementById("thumbnailList"),
  viewerPanel: document.getElementById("viewerPanel"),
  pagesContainer: document.getElementById("pagesContainer"),
  emptyHint: document.getElementById("emptyHint"),
  fullscreenMiniBar: document.getElementById("fullscreenMiniBar"),
  toggleFullscreenViewModeBtn: document.getElementById("toggleFullscreenViewModeBtn"),
  toggleThumbInFullscreenBtn: document.getElementById("toggleThumbInFullscreenBtn"),
  searchPanel: document.getElementById("searchPanel"),
  searchPanelCount: document.getElementById("searchPanelCount"),
  searchResultList: document.getElementById("searchResultList"),
  updateBanner: document.getElementById("updateBanner"),
  updateStageText: document.getElementById("updateStageText"),
  updateBannerPercentText: document.getElementById("updateBannerPercentText"),
  updateBannerBar: document.getElementById("updateBannerBar"),
  updateNotesModal: document.getElementById("updateNotesModal"),
  updateNotesTitle: document.getElementById("updateNotesTitle"),
  updateNotesVersion: document.getElementById("updateNotesVersion"),
  updateNotesList: document.getElementById("updateNotesList"),
  updateNotesCloseBtn: document.getElementById("updateNotesCloseBtn"),
  statusBar: document.getElementById("statusBar"),
  statusText: document.getElementById("statusText"),
  currentVersionLabel: document.getElementById("currentVersionLabel"),
  targetVersionLabel: document.getElementById("targetVersionLabel"),
  updateProgressWrap: document.getElementById("updateProgressWrap"),
  updateProgressBar: document.getElementById("updateProgressBar"),
  updateProgressText: document.getElementById("updateProgressText"),
  textMemoEditor: document.getElementById("textMemoEditor"),
  textMemoInput: document.getElementById("textMemoInput"),
  textMemoAddBtn: document.getElementById("textMemoAddBtn"),
  textMemoCancelBtn: document.getElementById("textMemoCancelBtn")
};

function toUint8Array(raw) {
  if (raw instanceof Uint8Array) {
    return raw;
  }
  if (raw instanceof ArrayBuffer) {
    return new Uint8Array(raw);
  }
  if (Array.isArray(raw)) {
    return Uint8Array.from(raw);
  }
  if (raw && raw.type === "Buffer" && Array.isArray(raw.data)) {
    return Uint8Array.from(raw.data);
  }
  throw new Error("파일 데이터를 읽지 못했습니다.");
}

function setStatus(message, isError = false) {
  els.statusText.textContent = message;
  els.statusText.style.color = isError ? "#d73333" : "";
}

function cloneAnnotationBucket(bucket) {
  return {
    highlights: (bucket?.highlights || []).map((item) => ({ ...item })),
    pens: (bucket?.pens || []).map((pen) => ({
      color: pen.color,
      width: pen.width,
      points: (pen.points || []).map((point) => ({ ...point }))
    })),
    texts: (bucket?.texts || []).map((note) => ({ ...note }))
  };
}

function createHistorySnapshot() {
  return {
    pageOrder: [...state.pageOrder],
    pageRotations: Array.from(state.pageRotations.entries()).map(([pageNum, rotation]) => [Number(pageNum), rotation]),
    annotations: Array.from(state.annotations.entries()).map(([pageNum, bucket]) => [Number(pageNum), cloneAnnotationBucket(bucket)]),
    currentPage: state.currentPage
  };
}

function updateHistoryButtons() {
  if (!els.undoBtn || !els.redoBtn) {
    return;
  }
  els.undoBtn.disabled = !state.pdfDoc || state.historyPast.length === 0;
  els.redoBtn.disabled = !state.pdfDoc || state.historyFuture.length === 0;
}

function clearHistory() {
  state.historyPast = [];
  state.historyFuture = [];
  updateHistoryButtons();
}

function pushHistorySnapshot() {
  if (!state.pdfDoc || state.historyRestoring) {
    return;
  }
  state.historyPast.push(createHistorySnapshot());
  if (state.historyPast.length > state.historyLimit) {
    state.historyPast.shift();
  }
  state.historyFuture = [];
  updateHistoryButtons();
}

async function restoreHistorySnapshot(snapshot) {
  if (!snapshot) {
    return;
  }
  state.historyRestoring = true;
  try {
    state.pageOrder = Array.isArray(snapshot.pageOrder) ? [...snapshot.pageOrder] : [];
    state.pageRotations = new Map((snapshot.pageRotations || []).map(([pageNum, rotation]) => [Number(pageNum), rotation]));
    state.annotations = new Map(
      (snapshot.annotations || []).map(([pageNum, bucket]) => [Number(pageNum), cloneAnnotationBucket(bucket)])
    );
    state.currentPage = Number(snapshot.currentPage || state.currentPage || 1);
    state.pageCache.clear();
    state.textItemsCache.clear();
    state.searchPageCache.clear();
    state.ocrWordCache.clear();
    ensureCurrentPageExists();
    await rebuildPageViews();
    await renderThumbnails();
    updatePageBadges();
    await goToPage(state.currentPage, false);
    if (state.searchQuery) {
      await performSearch(state.searchQuery, false);
    } else {
      clearSearchState();
    }
    state.saveDirty = true;
    updateToolbarState();
  } finally {
    state.historyRestoring = false;
    updateHistoryButtons();
  }
}

async function undoLastAction() {
  if (!state.pdfDoc || state.historyPast.length === 0) {
    return;
  }
  const current = createHistorySnapshot();
  const previous = state.historyPast.pop();
  state.historyFuture.push(current);
  await restoreHistorySnapshot(previous);
  setStatus(t("undoDone"));
}

async function redoLastAction() {
  if (!state.pdfDoc || state.historyFuture.length === 0) {
    return;
  }
  const current = createHistorySnapshot();
  const next = state.historyFuture.pop();
  state.historyPast.push(current);
  await restoreHistorySnapshot(next);
  setStatus(t("redoDone"));
}

function applyLanguageToStaticTexts() {
  if (state.applyingLanguage) {
    return;
  }
  state.applyingLanguage = true;
  try {
    document.documentElement.lang = state.language;
    document.querySelectorAll("[data-i18n]").forEach((el) => {
      const key = el.dataset.i18n;
      if (!key) {
        return;
      }
      el.textContent = t(key);
    });
    document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
      const key = el.dataset.i18nPlaceholder;
      if (!key) {
        return;
      }
      el.setAttribute("placeholder", t(key));
    });
    if (els.printBtn) {
      const label = state.language === "en" ? "Print" : "인쇄";
      els.printBtn.setAttribute("aria-label", label);
      els.printBtn.setAttribute("title", `${label} (Ctrl+P)`);
    }
    if (els.undoBtn) {
      els.undoBtn.setAttribute("aria-label", t("undo"));
      els.undoBtn.setAttribute("title", `${t("undo")} (Ctrl+Z)`);
    }
    if (els.redoBtn) {
      els.redoBtn.setAttribute("aria-label", t("redo"));
      els.redoBtn.setAttribute("title", `${t("redo")} (Ctrl+Y)`);
    }
    if (els.textMemoInput) {
      els.textMemoInput.setAttribute("placeholder", t("textMemoHint"));
    }
    if (els.textMemoAddBtn) {
      els.textMemoAddBtn.textContent = t("textMemoAdd");
    }
    if (els.textMemoCancelBtn) {
      els.textMemoCancelBtn.textContent = t("textMemoCancel");
    }
    if (els.updateNotesTitle) {
      els.updateNotesTitle.textContent = t("updateNotesTitle");
    }
    if (els.updateNotesCloseBtn) {
      els.updateNotesCloseBtn.textContent = t("updateNotesClose");
    }
    updateSearchCountText();
    updateVersionLabels();
    if (els.updateStageText) {
      els.updateStageText.textContent = updateStageLabel(state.updateStage);
    }
    applyPanelLayout();
    updateFullscreenButtons();
    updateToolbarState();
  } finally {
    state.applyingLanguage = false;
  }
}

async function setLanguage(language, persist = true) {
  state.language = language === "en" ? "en" : "ko";
  localStorage.setItem(storage.language, state.language);
  applyLanguageToStaticTexts();
  if (persist) {
    try {
      await window.lookupAPI.setLanguage(state.language);
    } catch (_error) {
      // non-fatal
    }
  }
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

const MAX_CANVAS_EDGE = 8192;
const MAX_CANVAS_PIXELS = 33554432;

function computeSafeRenderScale(width, height, requestedScale) {
  let safeScale = Math.max(0.2, requestedScale);
  const safeWidth = Math.max(1, width);
  const safeHeight = Math.max(1, height);
  let guard = 0;
  while (guard < 20) {
    const pixelWidth = safeWidth * safeScale;
    const pixelHeight = safeHeight * safeScale;
    const totalPixels = pixelWidth * pixelHeight;
    if (pixelWidth <= MAX_CANVAS_EDGE && pixelHeight <= MAX_CANVAS_EDGE && totalPixels <= MAX_CANVAS_PIXELS) {
      break;
    }
    safeScale *= 0.86;
    guard += 1;
  }
  return Math.max(0.2, safeScale);
}

function setZoomMode(mode) {
  state.zoomMode = mode === "manual" ? "manual" : "fit";
}

function normalizeSearchText(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function fileNameFromPath(filePath) {
  if (!filePath) {
    return "";
  }
  return filePath.replaceAll("\\", "/").split("/").pop() || filePath;
}

function getFileExt(filePath) {
  const normalized = String(filePath || "").trim().toLowerCase();
  const index = normalized.lastIndexOf(".");
  if (index < 0) {
    return "";
  }
  return normalized.slice(index);
}

function mapOpenErrorMessage(errorCode, fallbackMessage = "") {
  switch (String(errorCode || "").toUpperCase()) {
    case "NOT_FOUND":
      return t("openErrorNotFound");
    case "NO_PERMISSION":
      return t("openErrorPermission");
    case "UNSUPPORTED_FORMAT":
      return t("openErrorUnsupported");
    case "CONVERT_FAILED":
      return t("openErrorConvert");
    default:
      return fallbackMessage || t("openErrorGeneric");
  }
}

function buildWindowTitle(filePath) {
  const name = fileNameFromPath(filePath);
  return name ? `lookup - ${name}` : "lookup";
}

async function syncWindowTitle(filePath = "") {
  try {
    await window.lookupAPI.setWindowTitle(buildWindowTitle(filePath));
  } catch (_error) {
    // non-fatal
  }
}

function makeDefaultSavePath(currentPath) {
  if (!currentPath) {
    return "lookup-edited.pdf";
  }
  const normalized = currentPath.replaceAll("\\", "/");
  const dot = normalized.lastIndexOf(".");
  if (dot < 0) {
    return `${normalized}-edited.pdf`;
  }
  return `${normalized.slice(0, dot)}-edited.pdf`;
}

function getRotation(pageNum) {
  return state.pageRotations.get(pageNum) || 0;
}

function getEffectivePageRotation(page, pageNum) {
  const baseRotation = Number(page?.rotate || 0);
  const userRotation = getRotation(pageNum);
  return ((baseRotation + userRotation) % 360 + 360) % 360;
}

function getCurrentDisplayIndex() {
  return state.pageOrder.indexOf(state.currentPage);
}

function getNeighborPage(direction) {
  const currentIndex = getCurrentDisplayIndex();
  if (currentIndex < 0) {
    return null;
  }
  const targetIndex = currentIndex + (direction > 0 ? 1 : -1);
  if (targetIndex < 0 || targetIndex >= state.pageOrder.length) {
    return null;
  }
  return state.pageOrder[targetIndex];
}

async function stepPage(direction, smooth = true) {
  const nextPage = getNeighborPage(direction);
  if (!nextPage) {
    return false;
  }
  await goToPage(nextPage, smooth);
  return true;
}

function getAnnotationBucket(pageNum) {
  if (!state.annotations.has(pageNum)) {
    state.annotations.set(pageNum, { highlights: [], pens: [], texts: [] });
  }
  return state.annotations.get(pageNum);
}

function setDarkMode(enabled) {
  document.body.classList.toggle("dark", enabled);
  els.toggleDarkBtn.textContent = enabled ? t("lightMode") : t("darkMode");
  localStorage.setItem(storage.darkMode, enabled ? "1" : "0");
}

function applySavedDarkMode() {
  setDarkMode(localStorage.getItem(storage.darkMode) === "1");
}

function persistLayoutState() {
  localStorage.setItem(storage.leftPanelVisible, state.thumbPanelVisible ? "1" : "0");
  localStorage.setItem(storage.rightPanelVisible, state.searchPanelVisible ? "1" : "0");
  localStorage.setItem(storage.leftPanelWidth, String(Math.round(state.leftPanelWidth)));
  localStorage.setItem(storage.rightPanelWidth, String(Math.round(state.rightPanelWidth)));
  localStorage.setItem(storage.fullscreenMode, state.fullScreenViewMode);
}

function getEffectiveLeftPanelVisible() {
  return state.isFullScreen ? state.fullscreenThumbVisible : state.thumbPanelVisible;
}

function getEffectiveRightPanelVisible() {
  if (state.isFullScreen) {
    return false;
  }
  return state.searchPanelVisible;
}

function applyWorkspaceColumnSizes(leftVisible, rightVisible) {
  const leftWidth = leftVisible ? Math.round(state.leftPanelWidth) : 0;
  const rightWidth = rightVisible ? Math.round(state.rightPanelWidth) : 0;
  const leftResizerWidth = leftVisible ? 8 : 0;
  const rightResizerWidth = rightVisible ? 8 : 0;

  els.workspace.style.setProperty("--left-panel-width", `${leftWidth}px`);
  els.workspace.style.setProperty("--right-panel-width", `${rightWidth}px`);
  els.workspace.style.setProperty("--left-resizer-width", `${leftResizerWidth}px`);
  els.workspace.style.setProperty("--right-resizer-width", `${rightResizerWidth}px`);
}

function applyPanelLayout() {
  const leftVisible = getEffectiveLeftPanelVisible();
  const rightVisible = getEffectiveRightPanelVisible();

  applyWorkspaceColumnSizes(leftVisible, rightVisible);
  els.workspace.classList.toggle("left-collapsed", !leftVisible);
  els.workspace.classList.toggle("right-collapsed", !rightVisible);
  els.thumbPanel.classList.toggle("hidden", !leftVisible);
  els.leftResizer.classList.toggle("hidden", !leftVisible);
  els.searchPanel.classList.toggle("hidden", !rightVisible);
  els.rightResizer.classList.toggle("hidden", !rightVisible);

  els.toggleThumbPanelBtn.textContent = leftVisible ? t("thumbToggleHide") : t("thumbToggleShow");
  els.toggleSearchPanelBtn.textContent = rightVisible ? t("searchPanelToggleHide") : t("searchPanelToggleShow");
  els.toggleThumbInFullscreenBtn.textContent = leftVisible ? t("thumbToggleHide") : t("thumbToggleShow");
}

function normalizeUpdateStage(stage) {
  switch (String(stage || "").toLowerCase()) {
    case "checking":
      return "checking";
    case "available":
    case "downloading":
      return "downloading";
    case "ready-to-install":
    case "downloaded":
      return "ready";
    case "installing":
    case "restarting":
      return "restarting";
    case "installed":
      return "installed";
    case "error":
      return "error";
    case "disabled":
      return "disabled";
    case "idle":
    case "not-available":
    default:
      return "idle";
  }
}

function updateStageLabel(stage) {
  switch (normalizeUpdateStage(stage)) {
    case "checking":
      return t("updateStageChecking");
    case "downloading":
      return t("updateStageDownloading");
    case "ready":
      return t("updateStageReady");
    case "restarting":
      return t("updateStageRestarting");
    case "installed":
      return t("updateStageInstalled");
    case "error":
      return t("updateStageError");
    default:
      return t("updateStageIdle");
  }
}

function showUpdateBanner(show) {
  if (state.updateBannerHideTimer) {
    clearTimeout(state.updateBannerHideTimer);
    state.updateBannerHideTimer = 0;
  }
  if (!els.updateBanner) {
    return;
  }
  // v1.1.9부터 상단 큰 배너는 기본 동선에서 사용하지 않는다.
  els.updateBanner.classList.add("hidden");
}

function showUpdateProgressBar(show) {
  els.updateProgressWrap.classList.toggle("hidden", !show);
}

function setUpdateProgress(percent) {
  const safe = clamp(Math.round(percent), 0, 100);
  els.updateProgressBar.style.width = `${safe}%`;
  els.updateProgressText.textContent = `${safe}%`;
  if (els.updateBannerBar) {
    els.updateBannerBar.style.width = `${safe}%`;
  }
  if (els.updateBannerPercentText) {
    els.updateBannerPercentText.textContent = `${safe}%`;
  }
}

function applyUpdateVisualState(status, stage, percent) {
  const normalized = normalizeUpdateStage(stage || status);
  state.updateStage = normalized;
  if (els.updateStageText) {
    els.updateStageText.textContent = updateStageLabel(normalized);
  }
  if (typeof percent === "number") {
    setUpdateProgress(percent);
  } else if (normalized === "restarting" || normalized === "ready" || normalized === "installed") {
    setUpdateProgress(100);
  }

  const shouldShow = ["checking", "downloading", "ready", "restarting", "installed", "error"].includes(normalized);
  showUpdateProgressBar(shouldShow);
  showUpdateBanner(shouldShow);

  if (normalized === "installed") {
    state.updateBannerHideTimer = setTimeout(() => {
      state.updateBannerHideTimer = 0;
      showUpdateBanner(false);
      showUpdateProgressBar(false);
    }, 3600);
  } else if (normalized === "error") {
    state.updateBannerHideTimer = setTimeout(() => {
      state.updateBannerHideTimer = 0;
      showUpdateBanner(false);
      showUpdateProgressBar(false);
    }, 4200);
  }
}

function updateVersionLabels(currentVersion = state.appVersion, targetVersion = state.updateTargetVersion) {
  els.currentVersionLabel.textContent = `${t("versionCurrent")} v${currentVersion || "-"}`;
  if (targetVersion) {
    els.targetVersionLabel.classList.remove("hidden");
    els.targetVersionLabel.textContent = `${t("versionTarget")} v${targetVersion}`;
  } else {
    els.targetVersionLabel.classList.add("hidden");
    els.targetVersionLabel.textContent = `${t("versionTarget")} -`;
  }
}

function sanitizeReleaseNotesText(value) {
  let text = String(value || "").replace(/\uFEFF/g, "");
  text = text.replace(/```[\s\S]*?```/g, " ");
  text = text.replace(/`([^`]+)`/g, " $1 ");
  text = text.replace(/<\/?[^>]+>/g, " ");
  text = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1");
  text = text.replace(/[>*_~]/g, " ");
  text = text.replace(/\r/g, "");
  return text;
}

function toReleaseNoteLines(releaseNotes) {
  const source = Array.isArray(releaseNotes) ? releaseNotes.join("\n") : String(releaseNotes || "");
  const lines = sanitizeReleaseNotesText(source)
    .split(/\r?\n/)
    .map((line) => line.replace(/^[\s#\-*•\d.)]+/, "").trim())
    .filter(Boolean);
  const deduped = [];
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
    deduped.push(normalized);
    if (deduped.length >= 80) {
      break;
    }
  }
  return deduped;
}

function hideUpdateNotesModal() {
  if (!els.updateNotesModal) {
    return;
  }
  els.updateNotesModal.classList.add("hidden");
}

function showUpdateNotesModal(version, releaseNotes) {
  if (!els.updateNotesModal || !els.updateNotesList) {
    return;
  }
  const lines = toReleaseNoteLines(releaseNotes);
  els.updateNotesVersion.textContent = version ? `v${version}` : "";
  els.updateNotesList.innerHTML = "";
  if (!lines.length) {
    const li = document.createElement("li");
    li.textContent = t("updateNotesFallback");
    els.updateNotesList.appendChild(li);
  } else {
    lines.forEach((line) => {
      const li = document.createElement("li");
      li.textContent = line;
      els.updateNotesList.appendChild(li);
    });
  }
  els.updateNotesModal.classList.remove("hidden");
}

function updateFullscreenButtons() {
  els.toggleFullscreenBtn.textContent = state.isFullScreen ? t("fullscreenExit") : t("fullscreen");
  els.toggleFullscreenViewModeBtn.textContent =
    state.fullScreenViewMode === "single" ? t("fullscreenModeSingle") : t("fullscreenModeContinuous");
  document.body.classList.toggle("fullscreen-single", isSinglePageFullscreen());
}

function applyPageVisibility() {
  const singleMode = state.isFullScreen && state.fullScreenViewMode === "single";
  for (const [pageNum, view] of state.pageViews.entries()) {
    const hidden = singleMode && pageNum !== state.currentPage;
    view.wrap.classList.toggle("hidden-page", hidden);
  }
}

function ensureCurrentPageVisibleInSingleMode() {
  if (!isSinglePageFullscreen()) {
    return;
  }
  const currentView = state.pageViews.get(state.currentPage);
  if (!currentView) {
    return;
  }
  currentView.wrap.classList.remove("hidden-page");
}

function updateToolbarState() {
  const hasDoc = Boolean(state.pdfDoc);
  const total = state.pageOrder.length;
  const currentIndex = getCurrentDisplayIndex();

  els.prevPageBtn.disabled = !hasDoc || currentIndex <= 0;
  els.nextPageBtn.disabled = !hasDoc || currentIndex < 0 || currentIndex >= total - 1;
  els.pageInput.disabled = !hasDoc;
  els.zoomInBtn.disabled = !hasDoc;
  els.zoomOutBtn.disabled = !hasDoc;
  els.zoomResetBtn.disabled = !hasDoc;
  els.rotateLeftBtn.disabled = !hasDoc;
  els.rotateRightBtn.disabled = !hasDoc;
  els.deletePageBtn.disabled = !hasDoc || total <= 1;
  els.saveAsBtn.disabled = !hasDoc;
  els.saveOverwriteBtn.disabled = !hasDoc || state.sourceExt !== ".pdf";
  els.printBtn.disabled = !hasDoc;
  els.searchBtn.disabled = !hasDoc;
  els.searchInput.disabled = !hasDoc;
  els.searchPrevBtn.disabled = !hasDoc || state.searchMatches.length <= 1;
  els.searchNextBtn.disabled = !hasDoc || state.searchMatches.length <= 1;
  if (els.undoBtn) {
    els.undoBtn.disabled = !hasDoc || state.historyPast.length === 0;
  }
  if (els.redoBtn) {
    els.redoBtn.disabled = !hasDoc || state.historyFuture.length === 0;
  }

  els.pageCountLabel.textContent = `/ ${hasDoc ? total : 0}`;
  els.pageInput.value = `${hasDoc ? Math.max(1, currentIndex + 1) : 1}`;
  els.zoomLabel.textContent = `${Math.round(state.scale * 100)}%`;

  updateFullscreenButtons();
  updateHistoryButtons();
}

function updateActiveThumbnail() {
  for (const [pageNum, thumb] of state.thumbnails.entries()) {
    thumb.classList.toggle("active", pageNum === state.currentPage);
  }
}

function updateSearchCountText() {
  const count = state.searchMatches.length;
  els.searchCountLabel.textContent = t("searchCount", { count });
  els.searchPanelCount.textContent = t("searchPanelCount", { count });
}

function updateThumbnailSearchMark() {
  const hitPages = new Set(state.searchMatches.map((match) => match.pageNum));
  for (const [pageNum, thumb] of state.thumbnails.entries()) {
    thumb.classList.toggle("search-hit", hitPages.has(pageNum));
  }
}

function updateCurrentPage(pageNum, reasonText = "") {
  if (!state.pageOrder.includes(pageNum)) {
    return;
  }
  state.currentPage = pageNum;
  updateActiveThumbnail();
  updateToolbarState();
  applyPageVisibility();
  if (reasonText) {
    setStatus(reasonText);
  }
}

function clearSearchState() {
  state.searchQuery = "";
  state.searchMatches = [];
  state.perPageMatchIndexes = new Map();
  state.searchPageCache = new Map();
  state.activeSearchIndex = -1;
  els.searchResultList.innerHTML = "";
  updateSearchCountText();
  updateThumbnailSearchMark();
  for (const pageNum of state.pageOrder) {
    drawSearchHighlightsForPage(pageNum);
  }
}
async function getPdfPage(pageNum) {
  if (state.pageCache.has(pageNum)) {
    return state.pageCache.get(pageNum);
  }
  const page = await state.pdfDoc.getPage(pageNum);
  state.pageCache.set(pageNum, page);
  return page;
}

function buildPageElement(pageNum) {
  const wrap = document.createElement("div");
  wrap.className = "page-wrap";
  wrap.dataset.pageNum = String(pageNum);

  const canvas = document.createElement("canvas");
  canvas.className = "page-canvas";
  wrap.appendChild(canvas);

  const annotationCanvas = document.createElement("canvas");
  annotationCanvas.className = "annotation-canvas";
  wrap.appendChild(annotationCanvas);

  const searchOverlay = document.createElement("div");
  searchOverlay.className = "search-overlay";
  wrap.appendChild(searchOverlay);

  const badge = document.createElement("div");
  badge.className = "page-badge";
  badge.textContent = `${state.pageOrder.indexOf(pageNum) + 1}`;
  wrap.appendChild(badge);

  bindAnnotationEvents(pageNum, annotationCanvas);
  return { wrap, canvas, annotationCanvas, searchOverlay, badge, viewport: null, renderToken: 0 };
}

async function renderPage(pageNum) {
  const view = state.pageViews.get(pageNum);
  if (!view) {
    return;
  }
  const page = await getPdfPage(pageNum);
  const viewport = page.getViewport({ scale: state.scale, rotation: getEffectivePageRotation(page, pageNum) });
  view.viewport = viewport;

  view.wrap.style.width = `${viewport.width}px`;
  view.wrap.style.height = `${viewport.height}px`;

  const dpr = window.devicePixelRatio || 1;
  const requestedScale = dpr * state.mainRenderQuality;
  let renderScale = computeSafeRenderScale(viewport.width, viewport.height, requestedScale);
  const canvas = view.canvas;
  const syncCanvasSize = () => {
    canvas.width = Math.max(1, Math.floor(viewport.width * renderScale));
    canvas.height = Math.max(1, Math.floor(viewport.height * renderScale));
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;

    view.annotationCanvas.width = canvas.width;
    view.annotationCanvas.height = canvas.height;
    view.annotationCanvas.style.width = canvas.style.width;
    view.annotationCanvas.style.height = canvas.style.height;
  };
  syncCanvasSize();

  const renderToken = ++view.renderToken;
  const previousTask = state.pageRenderTasks.get(pageNum);
  if (previousTask && typeof previousTask.cancel === "function") {
    try {
      previousTask.cancel();
    } catch (_error) {
      // ignore cancellation errors
    }
  }

  const renderOnce = async (quality) => {
    const context = canvas.getContext("2d", { alpha: false });
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = quality;
    const renderTask = page.render({
      canvasContext: context,
      viewport,
      transform: renderScale === 1 ? null : [renderScale, 0, 0, renderScale, 0, 0]
    });
    state.pageRenderTasks.set(pageNum, renderTask);
    try {
      await renderTask.promise;
      return "ok";
    } catch (error) {
      if (error?.name === "RenderingCancelledException") {
        return "cancelled";
      }
      throw error;
    } finally {
      if (state.pageRenderTasks.get(pageNum) === renderTask) {
        state.pageRenderTasks.delete(pageNum);
      }
    }
  };

  let rendered = false;
  try {
    const result = await renderOnce("high");
    if (result === "cancelled") {
      return;
    }
    rendered = result === "ok";
  } catch (_error) {
    renderScale = computeSafeRenderScale(viewport.width, viewport.height, renderScale * 0.72);
    syncCanvasSize();
    try {
      const retry = await renderOnce("medium");
      if (retry === "cancelled") {
        return;
      }
      rendered = retry === "ok";
    } catch (_secondError) {
      rendered = false;
    }
  }
  if (!rendered) {
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
  }
  if (renderToken !== view.renderToken) {
    return;
  }

  drawAnnotationsForPage(pageNum);
  drawSearchHighlightsForPage(pageNum);
}

function getVisiblePageNumbers() {
  if (!state.pageOrder.length) {
    return [];
  }
  const viewportTop = els.viewerPanel.scrollTop - 40;
  const viewportBottom = viewportTop + els.viewerPanel.clientHeight + 80;
  const visible = [];
  for (const pageNum of state.pageOrder) {
    const view = state.pageViews.get(pageNum);
    if (!view || view.wrap.classList.contains("hidden-page")) {
      continue;
    }
    const top = view.wrap.offsetTop;
    const bottom = top + view.wrap.offsetHeight;
    if (bottom < viewportTop || top > viewportBottom) {
      continue;
    }
    visible.push(pageNum);
  }
  if (visible.length === 0 && state.currentPage) {
    return [state.currentPage];
  }
  return visible;
}

function buildPriorityRenderOrder() {
  const visible = getVisiblePageNumbers();
  const priority = new Set();
  for (const pageNum of visible) {
    priority.add(pageNum);
    const index = state.pageOrder.indexOf(pageNum);
    if (index > 0) {
      priority.add(state.pageOrder[index - 1]);
    }
    if (index >= 0 && index < state.pageOrder.length - 1) {
      priority.add(state.pageOrder[index + 1]);
    }
  }

  const orderedPriority = state.pageOrder.filter((pageNum) => priority.has(pageNum));
  const orderedRest = state.pageOrder.filter((pageNum) => !priority.has(pageNum));
  return { orderedPriority, orderedRest };
}

async function renderPagesList(pageNums, version) {
  for (const pageNum of pageNums) {
    if (version !== state.renderVersion) {
      return;
    }
    await renderPage(pageNum);
  }
}

function scheduleBackgroundRender(orderedRest, version) {
  if (state.fullRenderTimer) {
    clearTimeout(state.fullRenderTimer);
    state.fullRenderTimer = 0;
  }
  if (!orderedRest.length) {
    return;
  }
  const passVersion = ++state.fullRenderPassVersion;
  state.fullRenderTimer = setTimeout(async () => {
    if (passVersion !== state.fullRenderPassVersion || version !== state.renderVersion) {
      return;
    }
    await renderPagesList(orderedRest, version);
  }, 0);
}

async function renderAllPages(options = {}) {
  if (!state.pdfDoc) {
    return;
  }
  for (const renderTask of state.pageRenderTasks.values()) {
    if (renderTask && typeof renderTask.cancel === "function") {
      try {
        renderTask.cancel();
      } catch (_error) {
        // ignore cancellation errors
      }
    }
  }
  state.pageRenderTasks.clear();
  const version = ++state.renderVersion;
  const prioritizeVisible = options.prioritizeVisible !== false;
  if (!prioritizeVisible) {
    await renderPagesList(state.pageOrder, version);
    return;
  }

  const { orderedPriority, orderedRest } = buildPriorityRenderOrder();
  await renderPagesList(orderedPriority, version);
  scheduleBackgroundRender(orderedRest, version);
}

async function renderThumbnail(pageNum, thumbCanvas) {
  const page = await getPdfPage(pageNum);
  const rotation = getEffectivePageRotation(page, pageNum);
  const viewport = page.getViewport({ scale: 1, rotation });
  const panelWidth = Math.max(120, els.thumbnailList.clientWidth || state.leftPanelWidth - 20);
  const targetWidth = clamp(panelWidth - 8, 120, 860);
  const thumbScale = targetWidth / viewport.width;
  const scaledViewport = page.getViewport({ scale: thumbScale, rotation });
  const dpr = window.devicePixelRatio || 1;
  let renderScale = computeSafeRenderScale(scaledViewport.width, scaledViewport.height, dpr * state.thumbRenderQuality);
  const context = thumbCanvas.getContext("2d", { alpha: false });
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";

  thumbCanvas.width = Math.max(1, Math.floor(scaledViewport.width * renderScale));
  thumbCanvas.height = Math.max(1, Math.floor(scaledViewport.height * renderScale));
  thumbCanvas.style.width = `${scaledViewport.width}px`;
  thumbCanvas.style.height = `${scaledViewport.height}px`;

  const previousTask = state.thumbRenderTasks.get(pageNum);
  if (previousTask && typeof previousTask.cancel === "function") {
    try {
      previousTask.cancel();
    } catch (_error) {
      // ignore cancellation errors
    }
  }

  const renderThumbOnce = async () => {
    const task = page.render({
      canvasContext: context,
      viewport: scaledViewport,
      transform: renderScale === 1 ? null : [renderScale, 0, 0, renderScale, 0, 0]
    });
    state.thumbRenderTasks.set(pageNum, task);
    try {
      await task.promise;
      return "ok";
    } catch (error) {
      if (error?.name === "RenderingCancelledException") {
        return "cancelled";
      }
      throw error;
    } finally {
      if (state.thumbRenderTasks.get(pageNum) === task) {
        state.thumbRenderTasks.delete(pageNum);
      }
    }
  };

  try {
    const result = await renderThumbOnce();
    if (result === "cancelled") {
      return;
    }
  } catch (_error) {
    renderScale = computeSafeRenderScale(scaledViewport.width, scaledViewport.height, renderScale * 0.72);
    thumbCanvas.width = Math.max(1, Math.floor(scaledViewport.width * renderScale));
    thumbCanvas.height = Math.max(1, Math.floor(scaledViewport.height * renderScale));
    const retry = await renderThumbOnce();
    if (retry === "cancelled") {
      return;
    }
  }
}

function queueThumbnailRerender() {
  if (state.thumbRerenderTimer) {
    clearTimeout(state.thumbRerenderTimer);
  }
  state.thumbRerenderTimer = setTimeout(async () => {
    state.thumbRerenderTimer = 0;
    if (!state.pdfDoc || !state.thumbnails.size) {
      return;
    }
    const version = ++state.thumbRenderVersion;
    for (const pageNum of state.pageOrder) {
      if (version !== state.thumbRenderVersion) {
        return;
      }
      const thumb = state.thumbnails.get(pageNum);
      const canvas = thumb?.querySelector("canvas");
      if (!canvas) {
        continue;
      }
      await renderThumbnail(pageNum, canvas);
    }
  }, 120);
}

async function renderThumbnails() {
  const version = ++state.thumbRenderVersion;
  els.thumbnailList.innerHTML = "";
  state.thumbnails.clear();

  for (const pageNum of state.pageOrder) {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "thumb-item";
    item.dataset.pageNum = String(pageNum);
    item.draggable = true;

    const canvas = document.createElement("canvas");
    item.appendChild(canvas);

    const label = document.createElement("div");
    label.className = "thumb-label";
    label.textContent = `${state.pageOrder.indexOf(pageNum) + 1}`;
    item.appendChild(label);

    item.addEventListener("click", async () => {
      await goToPage(pageNum, true);
    });

    item.addEventListener("dragstart", (event) => {
      event.dataTransfer?.setData("text/plain", String(pageNum));
      event.dataTransfer.effectAllowed = "move";
    });
    item.addEventListener("dragover", (event) => {
      event.preventDefault();
      item.classList.add("drag-over");
    });
    item.addEventListener("dragleave", () => {
      item.classList.remove("drag-over");
    });
    item.addEventListener("drop", async (event) => {
      event.preventDefault();
      item.classList.remove("drag-over");
      const draggedPage = Number.parseInt(event.dataTransfer?.getData("text/plain") || "", 10);
      if (!Number.isFinite(draggedPage) || draggedPage === pageNum) {
        return;
      }
      await movePageOrder(draggedPage, pageNum);
    });

    els.thumbnailList.appendChild(item);
    state.thumbnails.set(pageNum, item);
    if (version !== state.thumbRenderVersion) {
      return;
    }
    await renderThumbnail(pageNum, canvas);
  }

  updateActiveThumbnail();
  updateThumbnailSearchMark();
}

async function rebuildPageViews() {
  state.pageViews.clear();
  els.pagesContainer.innerHTML = "";

  for (const pageNum of state.pageOrder) {
    const pageView = buildPageElement(pageNum);
    state.pageViews.set(pageNum, pageView);
    els.pagesContainer.appendChild(pageView.wrap);
  }
  applyAnnotationInteractivity();
  applyPageVisibility();
  await renderAllPages();
}

function updatePageBadges() {
  for (const pageNum of state.pageOrder) {
    const view = state.pageViews.get(pageNum);
    const thumb = state.thumbnails.get(pageNum);
    if (view) {
      view.badge.textContent = `${state.pageOrder.indexOf(pageNum) + 1}`;
    }
    if (thumb) {
      const label = thumb.querySelector(".thumb-label");
      if (label) {
        label.textContent = `${state.pageOrder.indexOf(pageNum) + 1}`;
      }
    }
  }
}

async function goToPage(pageNum, smooth = false) {
  if (!state.pageOrder.includes(pageNum)) {
    return;
  }
  updateCurrentPage(pageNum);

  const view = state.pageViews.get(pageNum);
  if (!view) {
    return;
  }
  if (isSinglePageFullscreen()) {
    view.wrap.scrollIntoView({
      behavior: smooth ? "smooth" : "auto",
      block: "center",
      inline: "center"
    });
    alignCurrentPageToViewerCenter();
    return;
  }
  view.wrap.scrollIntoView({
    behavior: smooth ? "smooth" : "auto",
    block: "center",
    inline: "center"
  });
}

function alignCurrentPageToViewerCenter() {
  const view = state.pageViews.get(state.currentPage);
  if (!view) {
    return;
  }
  const top = Math.max(0, view.wrap.offsetTop + view.wrap.offsetHeight / 2 - els.viewerPanel.clientHeight / 2);
  const left = Math.max(0, view.wrap.offsetLeft + view.wrap.offsetWidth / 2 - els.viewerPanel.clientWidth / 2);
  els.viewerPanel.scrollTop = top;
  els.viewerPanel.scrollLeft = left;
}

function updateCurrentPageByScroll() {
  if (!state.pdfDoc) {
    return;
  }
  if (isSinglePageFullscreen()) {
    return;
  }

  const panelTop = els.viewerPanel.scrollTop + els.viewerPanel.clientHeight * 0.35;
  let bestPage = state.currentPage;
  let bestDistance = Number.POSITIVE_INFINITY;

  for (const pageNum of state.pageOrder) {
    const view = state.pageViews.get(pageNum);
    if (!view || view.wrap.classList.contains("hidden-page")) {
      continue;
    }
    const center = view.wrap.offsetTop + view.wrap.offsetHeight / 2;
    const distance = Math.abs(center - panelTop);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestPage = pageNum;
    }
  }

  if (bestPage !== state.currentPage) {
    updateCurrentPage(bestPage);
  }
}

function queueScrollSync() {
  if (state.scrollRaf) {
    return;
  }
  state.scrollRaf = requestAnimationFrame(() => {
    state.scrollRaf = 0;
    updateCurrentPageByScroll();
  });
}

function ensureCurrentPageExists() {
  if (!state.pageOrder.length) {
    return;
  }
  if (!state.pageOrder.includes(state.currentPage)) {
    state.currentPage = state.pageOrder[0];
  }
}

function isSinglePageFullscreen() {
  return state.isFullScreen && state.fullScreenViewMode === "single";
}

function focusViewerPanel() {
  if (!els.viewerPanel) {
    return;
  }
  try {
    els.viewerPanel.focus({ preventScroll: true });
  } catch (_error) {
    els.viewerPanel.focus();
  }
}

function viewerSizeIsValid() {
  return els.viewerPanel.clientWidth > 80 && els.viewerPanel.clientHeight > 80;
}

function countPagesIntersectingViewer() {
  const panelRect = els.viewerPanel.getBoundingClientRect();
  let count = 0;
  for (const pageNum of state.pageOrder) {
    const view = state.pageViews.get(pageNum);
    if (!view || view.wrap.classList.contains("hidden-page")) {
      continue;
    }
    const rect = view.wrap.getBoundingClientRect();
    if (rect.width < 4 || rect.height < 4) {
      continue;
    }
    const intersects =
      rect.bottom > panelRect.top + 2 &&
      rect.top < panelRect.bottom - 2 &&
      rect.right > panelRect.left + 2 &&
      rect.left < panelRect.right - 2;
    if (intersects) {
      count += 1;
      if (count > 0) {
        return count;
      }
    }
  }
  return count;
}

async function ensureViewerPageVisible() {
  if (!state.pdfDoc || !state.pageOrder.length) {
    return;
  }
  if (countPagesIntersectingViewer() > 0) {
    return;
  }

  ensureCurrentPageExists();
  const currentView = state.pageViews.get(state.currentPage);
  if (currentView) {
    currentView.wrap.classList.remove("hidden-page");
  }
  if (state.currentPage) {
    await renderPage(state.currentPage);
    await goToPage(state.currentPage, false);
    if (isSinglePageFullscreen()) {
      alignCurrentPageToViewerCenter();
    }
  }
  if (countPagesIntersectingViewer() > 0) {
    return;
  }

  // 최종 안전 복구: 문서 뷰를 다시 만들고 현재 페이지를 다시 보장한다.
  await rebuildPageViews();
  ensureCurrentPageExists();
  const rebuiltView = state.pageViews.get(state.currentPage);
  if (rebuiltView) {
    rebuiltView.wrap.classList.remove("hidden-page");
  }
  if (state.currentPage) {
    await renderPage(state.currentPage);
    await goToPage(state.currentPage, false);
    if (isSinglePageFullscreen()) {
      alignCurrentPageToViewerCenter();
    }
  }
}

async function fitCurrentPageToViewport() {
  if (!state.pdfDoc || !state.currentPage) {
    return;
  }
  const page = await getPdfPage(state.currentPage);
  const baseViewport = page.getViewport({ scale: 1, rotation: getEffectivePageRotation(page, state.currentPage) });
  const maxWidth = Math.max(100, els.viewerPanel.clientWidth - 42);
  const maxHeight = Math.max(100, els.viewerPanel.clientHeight - 44);
  const fitWidthScale = maxWidth / Math.max(1, baseViewport.width);
  const fitHeightScale = maxHeight / Math.max(1, baseViewport.height);
  const nextScale = state.fullScreenViewMode === "single" ? Math.min(fitWidthScale, fitHeightScale) : fitWidthScale;
  await zoomTo(clamp(nextScale, 0.25, 6), null, { prioritizeVisible: true, zoomMode: "fit" });
  state.fullScreenAutoFitDone = true;
}

function shouldApplyFullscreenFit(options = {}) {
  if (!state.isFullScreen) {
    return false;
  }
  if (options.forceFit) {
    return true;
  }
  if (state.zoomMode === "manual") {
    return false;
  }
  return !state.fullScreenAutoFitDone;
}

function queueLayoutRecoveryRender(options = {}) {
  if (!state.pdfDoc) {
    return;
  }
  const token = ++state.layoutRecoveryToken;
  const attempt = Number(options.attempt || 0);
  const runRecovery = async () => {
    if (token !== state.layoutRecoveryToken) {
      return;
    }
    ensureCurrentPageExists();
    if (!viewerSizeIsValid()) {
      if (attempt < 5) {
        setTimeout(() => {
          queueLayoutRecoveryRender({ ...options, attempt: attempt + 1 });
        }, 120);
      }
      return;
    }
    if (token !== state.layoutRecoveryToken) {
      return;
    }
    state.viewerRenderRecoveryCount += 1;
    applyPageVisibility();
    ensureCurrentPageVisibleInSingleMode();
    if (state.currentPage && state.pageViews.has(state.currentPage)) {
      await renderPage(state.currentPage);
    }
    await goToPage(state.currentPage, false);
    if (shouldApplyFullscreenFit(options)) {
      await fitCurrentPageToViewport();
      await goToPage(state.currentPage, false);
    } else if (isSinglePageFullscreen()) {
      alignCurrentPageToViewerCenter();
    }
    await renderAllPages({ prioritizeVisible: true });
    if (isSinglePageFullscreen()) {
      alignCurrentPageToViewerCenter();
    }
    await ensureViewerPageVisible();
    if (state.isFullScreen) {
      focusViewerPanel();
    }
  };
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      runRecovery().catch(() => {});
    });
  });
  if (state.layoutRecoveryTimer) {
    clearTimeout(state.layoutRecoveryTimer);
  }
  state.layoutRecoveryTimer = setTimeout(() => {
    state.layoutRecoveryTimer = 0;
    requestAnimationFrame(() => {
      runRecovery().catch(() => {});
    });
  }, 120);
}

function itemRectInViewport(item, viewport) {
  if (item?.isOcr) {
    return {
      left: (item.leftRatio || 0) * viewport.width,
      top: (item.topRatio || 0) * viewport.height,
      width: Math.max(2, (item.widthRatio || 0) * viewport.width),
      height: Math.max(2, (item.heightRatio || 0) * viewport.height)
    };
  }
  const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
  const baseX = Math.max(0.0001, Math.hypot(item.transform?.[0] || 0, item.transform?.[1] || 0));
  const baseY = Math.max(0.0001, Math.hypot(item.transform?.[2] || 0, item.transform?.[3] || 0));
  const scaledX = Math.max(0.0001, Math.hypot(tx[0], tx[1]));
  const scaledY = Math.max(0.0001, Math.hypot(tx[2], tx[3]));
  const widthScale = scaledX / baseX;
  const heightScale = scaledY / baseY;
  const measuredWidth = Number(item.width || 0) * widthScale;
  const measuredHeight = Number(item.height || 0) * heightScale;
  const width = Number.isFinite(measuredWidth) && measuredWidth > 0
    ? measuredWidth
    : Math.max(8, item.str.length * Math.max(8, measuredHeight || 12) * 0.5);
  const fontHeight = Math.max(8, Number.isFinite(measuredHeight) && measuredHeight > 0 ? measuredHeight : scaledY);
  const left = tx[4];
  const top = tx[5] - fontHeight;
  return {
    left,
    top,
    width: Math.max(4, width),
    height: fontHeight
  };
}

function itemSegmentRectInViewport(item, viewport, segmentStart, segmentEnd) {
  const full = itemRectInViewport(item, viewport);
  const safeLength = Math.max(1, Number(item.searchableLength || item.searchable?.length || item.str?.length || 1));
  const startRatio = clamp(segmentStart / safeLength, 0, 1);
  const endRatio = clamp(segmentEnd / safeLength, 0, 1);
  const left = full.left + full.width * Math.min(startRatio, endRatio);
  const right = full.left + full.width * Math.max(startRatio, endRatio);
  return {
    left,
    top: full.top,
    width: Math.max(2, right - left),
    height: full.height
  };
}

function drawSearchHighlightsForPage(pageNum) {
  const view = state.pageViews.get(pageNum);
  if (!view || !view.viewport) {
    return;
  }

  view.searchOverlay.innerHTML = "";
  const matchIndexes = state.perPageMatchIndexes.get(pageNum) || [];
  const items = state.searchPageCache.get(pageNum) || state.textItemsCache.get(pageNum) || state.ocrWordCache.get(pageNum) || [];

  for (const matchIndex of matchIndexes) {
    const match = state.searchMatches[matchIndex];
    for (const segment of match.segments || []) {
      const itemIndex = segment.itemIndex;
      const item = items[itemIndex];
      if (!item) {
        continue;
      }
      const rect = itemSegmentRectInViewport(item, view.viewport, segment.startOffset, segment.endOffset);
      const box = document.createElement("div");
      box.className = "search-hit-box";
      if (matchIndex === state.activeSearchIndex) {
        box.classList.add("active");
      }
      box.style.left = `${rect.left}px`;
      box.style.top = `${rect.top}px`;
      box.style.width = `${rect.width}px`;
      box.style.height = `${rect.height}px`;
      view.searchOverlay.appendChild(box);
    }
  }
}
function drawAnnotationsForPage(pageNum, transientDrawing = null) {
  const view = state.pageViews.get(pageNum);
  if (!view || !view.viewport) {
    return;
  }
  const canvas = view.annotationCanvas;
  const context = canvas.getContext("2d");
  const cssWidth = view.viewport.width;
  const cssHeight = view.viewport.height;
  const pixelScaleX = canvas.width / Math.max(1, cssWidth);
  const pixelScaleY = canvas.height / Math.max(1, cssHeight);
  context.setTransform(pixelScaleX, 0, 0, pixelScaleY, 0, 0);
  context.clearRect(0, 0, cssWidth, cssHeight);

  const bucket = getAnnotationBucket(pageNum);

  for (const mark of bucket.highlights) {
    const p1 = view.viewport.convertToViewportPoint(mark.x1, mark.y1);
    const p2 = view.viewport.convertToViewportPoint(mark.x2, mark.y2);
    const left = Math.min(p1[0], p2[0]);
    const top = Math.min(p1[1], p2[1]);
    const width = Math.max(4, Math.abs(p1[0] - p2[0]));
    const height = Math.max(4, Math.abs(p1[1] - p2[1]));
    context.fillStyle = "rgba(255, 226, 46, 0.38)";
    context.fillRect(left, top, width, height);
  }

  for (const pen of bucket.pens) {
    if (!pen.points.length) {
      continue;
    }
    context.strokeStyle = pen.color;
    context.lineWidth = pen.width;
    context.lineJoin = "round";
    context.lineCap = "round";
    context.beginPath();
    pen.points.forEach((point, index) => {
      const [x, y] = view.viewport.convertToViewportPoint(point.x, point.y);
      if (index === 0) {
        context.moveTo(x, y);
      } else {
        context.lineTo(x, y);
      }
    });
    context.stroke();
  }

  for (const note of bucket.texts) {
    const [x, y] = view.viewport.convertToViewportPoint(note.x, note.y);
    context.font = "13px Segoe UI";
    const padX = 6;
    const padY = 4;
    const textWidth = context.measureText(note.text).width;
    const boxWidth = textWidth + padX * 2;
    const boxHeight = 22;
    context.fillStyle = "rgba(255, 241, 133, 0.9)";
    context.fillRect(x, y - boxHeight, boxWidth, boxHeight);
    context.strokeStyle = "rgba(180, 145, 20, 0.95)";
    context.strokeRect(x, y - boxHeight, boxWidth, boxHeight);
    context.fillStyle = "#1a1a1a";
    context.fillText(note.text, x + padX, y - padY);
  }

  if (transientDrawing && transientDrawing.pageNum === pageNum) {
    if (transientDrawing.type === "highlight" && transientDrawing.points.length >= 2) {
      const first = transientDrawing.points[0];
      const last = transientDrawing.points[transientDrawing.points.length - 1];
      const p1 = view.viewport.convertToViewportPoint(first.x, first.y);
      const p2 = view.viewport.convertToViewportPoint(last.x, last.y);
      const left = Math.min(p1[0], p2[0]);
      const top = Math.min(p1[1], p2[1]);
      const width = Math.max(4, Math.abs(p1[0] - p2[0]));
      const height = Math.max(4, Math.abs(p1[1] - p2[1]));
      context.fillStyle = "rgba(255, 226, 46, 0.38)";
      context.fillRect(left, top, width, height);
    } else if (transientDrawing.type === "pen" && transientDrawing.points.length >= 2) {
      context.strokeStyle = "#ff5353";
      context.lineWidth = 2.2;
      context.lineJoin = "round";
      context.lineCap = "round";
      context.beginPath();
      transientDrawing.points.forEach((point, index) => {
        const [x, y] = view.viewport.convertToViewportPoint(point.x, point.y);
        if (index === 0) {
          context.moveTo(x, y);
        } else {
          context.lineTo(x, y);
        }
      });
      context.stroke();
    }
  }
}

function cssToPdfPoint(pageNum, cssX, cssY) {
  const view = state.pageViews.get(pageNum);
  if (!view?.viewport) {
    return null;
  }
  const [x, y] = view.viewport.convertToPdfPoint(cssX, cssY);
  return { x, y };
}

function setEditingMode(mode) {
  state.editingMode = mode;
  for (const button of els.editModeButtons) {
    button.classList.toggle("active", button.dataset.mode === mode);
  }
  applyAnnotationInteractivity();
  setStatus(
    mode === "view"
      ? "보기 모드"
      : mode === "highlight"
        ? "형광펜 모드"
        : mode === "pen"
          ? "펜 모드"
          : "텍스트 메모 모드"
  );
}

function applyAnnotationInteractivity() {
  const isEditMode = state.editingMode !== "view";
  for (const view of state.pageViews.values()) {
    view.annotationCanvas.style.pointerEvents = isEditMode ? "auto" : "none";
    view.annotationCanvas.style.cursor =
      state.editingMode === "text"
        ? "cell"
        : state.editingMode === "highlight" || state.editingMode === "pen"
          ? "crosshair"
          : "default";
  }
}

function closeTextMemoEditor() {
  if (!els.textMemoEditor || !els.textMemoInput) {
    return;
  }
  els.textMemoEditor.classList.add("hidden");
  els.textMemoInput.value = "";
  state.pendingTextMemo = null;
}

function openTextMemoEditor(pageNum, clientX, clientY, startPoint) {
  if (!els.textMemoEditor || !els.textMemoInput || !els.textMemoAddBtn) {
    return;
  }
  const panelRect = els.viewerPanel.getBoundingClientRect();
  const maxLeft = Math.max(20, els.viewerPanel.scrollWidth - 320);
  const maxTop = Math.max(20, els.viewerPanel.scrollHeight - 140);
  const left = clamp(clientX - panelRect.left + els.viewerPanel.scrollLeft, 12, maxLeft);
  const top = clamp(clientY - panelRect.top + els.viewerPanel.scrollTop, 12, maxTop);
  els.textMemoEditor.style.left = `${left}px`;
  els.textMemoEditor.style.top = `${top}px`;
  els.textMemoEditor.classList.remove("hidden");
  state.pendingTextMemo = { pageNum, x: startPoint.x, y: startPoint.y };
  els.textMemoInput.value = "";
  els.textMemoInput.focus();
  els.textMemoInput.select();
}

function commitTextMemoFromEditor() {
  const pending = state.pendingTextMemo;
  if (!pending || !els.textMemoInput) {
    closeTextMemoEditor();
    return;
  }
  const text = String(els.textMemoInput.value || "").trim();
  if (!text) {
    closeTextMemoEditor();
    return;
  }
  pushHistorySnapshot();
  const bucket = getAnnotationBucket(pending.pageNum);
  bucket.texts.push({ x: pending.x, y: pending.y, text });
  state.saveDirty = true;
  drawAnnotationsForPage(pending.pageNum);
  closeTextMemoEditor();
  updateToolbarState();
  setStatus(t("textMemoAdded"));
}

function bindAnnotationEvents(pageNum, annotationCanvas) {
  annotationCanvas.addEventListener("pointerdown", (event) => {
    if (state.editingMode === "view" || !state.pdfDoc) {
      return;
    }
    event.preventDefault();
    updateCurrentPage(pageNum);

    const rect = annotationCanvas.getBoundingClientRect();
    const cssX = event.clientX - rect.left;
    const cssY = event.clientY - rect.top;
    const startPoint = cssToPdfPoint(pageNum, cssX, cssY);
    if (!startPoint) {
      return;
    }

    if (state.editingMode === "text") {
      openTextMemoEditor(pageNum, event.clientX, event.clientY, startPoint);
      return;
    }

    state.drawing = {
      pageNum,
      pointerId: event.pointerId,
      type: state.editingMode,
      points: [startPoint]
    };
    annotationCanvas.setPointerCapture(event.pointerId);
    drawAnnotationsForPage(pageNum, state.drawing);
  });

  annotationCanvas.addEventListener("pointermove", (event) => {
    if (!state.drawing || state.drawing.pageNum !== pageNum || state.drawing.pointerId !== event.pointerId) {
      return;
    }
    event.preventDefault();
    const rect = annotationCanvas.getBoundingClientRect();
    const cssX = event.clientX - rect.left;
    const cssY = event.clientY - rect.top;
    const point = cssToPdfPoint(pageNum, cssX, cssY);
    if (!point) {
      return;
    }
    state.drawing.points.push(point);
    drawAnnotationsForPage(pageNum, state.drawing);
  });

  annotationCanvas.addEventListener("pointerup", (event) => {
    if (!state.drawing || state.drawing.pageNum !== pageNum || state.drawing.pointerId !== event.pointerId) {
      return;
    }
    event.preventDefault();
    const drawing = state.drawing;
    state.drawing = null;

    const bucket = getAnnotationBucket(pageNum);
    if (drawing.type === "highlight" && drawing.points.length >= 2) {
      pushHistorySnapshot();
      const first = drawing.points[0];
      const last = drawing.points[drawing.points.length - 1];
      bucket.highlights.push({
        x1: first.x,
        y1: first.y,
        x2: last.x,
        y2: last.y
      });
      state.saveDirty = true;
    } else if (drawing.type === "pen" && drawing.points.length >= 2) {
      pushHistorySnapshot();
      bucket.pens.push({
        points: drawing.points.map((point) => ({ x: point.x, y: point.y })),
        color: "#ff5252",
        width: 2.2
      });
      state.saveDirty = true;
    }
    drawAnnotationsForPage(pageNum);
  });
}

async function ensureTextItems(pageNum) {
  if (state.textItemsCache.has(pageNum)) {
    return state.textItemsCache.get(pageNum);
  }
  const page = await getPdfPage(pageNum);
  const textContent = await page.getTextContent();
  const items = [];
  for (const item of textContent.items) {
    const displayText = String(item.str || "").replace(/\s+/g, " ").trim();
    if (!displayText) {
      continue;
    }

    const lowered = displayText.toLowerCase();
    const searchable = lowered;
    if (!searchable) {
      continue;
    }

    items.push({
      index: items.length,
      str: displayText,
      isOcr: false,
      lower: lowered,
      searchable,
      searchableLength: searchable.length,
      width: Number(item.width || 0),
      height: Number(item.height || 0),
      transform: Array.isArray(item.transform) ? [...item.transform] : [1, 0, 0, 1, 0, 0]
    });
  }

  state.textItemsCache.set(pageNum, items);
  return items;
}

async function ensureOcrWorker() {
  if (state.ocrWorkerReady && state.ocrWorker) {
    return state.ocrWorker;
  }
  if (state.ocrWorkerPromise) {
    return state.ocrWorkerPromise;
  }
  state.ocrWorkerPromise = (async () => {
    const createWorker = await ensureTesseractCreateWorker();
    try {
      const worker = await createWorker("kor+eng");
      state.ocrWorker = worker;
      state.ocrWorkerReady = true;
      return worker;
    } catch (_error) {
      const worker = await createWorker("eng");
      state.ocrWorker = worker;
      state.ocrWorkerReady = true;
      return worker;
    }
  })();
  try {
    return await state.ocrWorkerPromise;
  } finally {
    state.ocrWorkerPromise = null;
  }
}

async function ensureOcrItems(pageNum) {
  if (state.ocrWordCache.has(pageNum)) {
    return state.ocrWordCache.get(pageNum);
  }
  if (state.ocrPagePromises.has(pageNum)) {
    return state.ocrPagePromises.get(pageNum);
  }
  const promise = (async () => {
    const page = await getPdfPage(pageNum);
    const effectiveRotation = getEffectivePageRotation(page, pageNum);
    const baseViewport = page.getViewport({ scale: 1, rotation: effectiveRotation });
    const longerEdge = Math.max(baseViewport.width, baseViewport.height);
    const rasterScale = clamp(2100 / Math.max(1, longerEdge), 1, 3);
    const rasterViewport = page.getViewport({ scale: rasterScale, rotation: effectiveRotation });
    const rasterCanvas = document.createElement("canvas");
    rasterCanvas.width = Math.max(1, Math.round(rasterViewport.width));
    rasterCanvas.height = Math.max(1, Math.round(rasterViewport.height));
    const rasterContext = rasterCanvas.getContext("2d", { alpha: false });
    rasterContext.imageSmoothingEnabled = true;
    rasterContext.imageSmoothingQuality = "high";
    await page.render({ canvasContext: rasterContext, viewport: rasterViewport }).promise;

    const worker = await ensureOcrWorker();
    const ocrResult = await worker.recognize(rasterCanvas);
    const words = Array.isArray(ocrResult?.data?.words) ? ocrResult.data.words : [];
    const items = [];
    for (const word of words) {
      const raw = String(word?.text || "").replace(/\s+/g, " ").trim();
      if (!raw) {
        continue;
      }
      const lower = raw.toLowerCase();
      const x0 = Number(word?.bbox?.x0 ?? word?.bbox?.xMin ?? 0);
      const y0 = Number(word?.bbox?.y0 ?? word?.bbox?.yMin ?? 0);
      const x1 = Number(word?.bbox?.x1 ?? word?.bbox?.xMax ?? 0);
      const y1 = Number(word?.bbox?.y1 ?? word?.bbox?.yMax ?? 0);
      const left = clamp(Math.min(x0, x1) / Math.max(1, rasterViewport.width), 0, 1);
      const top = clamp(Math.min(y0, y1) / Math.max(1, rasterViewport.height), 0, 1);
      const width = clamp(Math.abs(x1 - x0) / Math.max(1, rasterViewport.width), 0.001, 1);
      const height = clamp(Math.abs(y1 - y0) / Math.max(1, rasterViewport.height), 0.001, 1);
      items.push({
        index: items.length,
        isOcr: true,
        str: raw,
        lower,
        searchable: lower,
        searchableLength: lower.length,
        leftRatio: left,
        topRatio: top,
        widthRatio: width,
        heightRatio: height
      });
    }
    state.ocrWordCache.set(pageNum, items);
    return items;
  })()
    .catch((error) => {
      if (!state.ocrErrorShown) {
        state.ocrErrorShown = true;
        setStatus(`${t("ocrFailed")} (${error?.message || "OCR"})`, true);
      }
      state.ocrWordCache.set(pageNum, []);
      return [];
    })
    .finally(() => {
      state.ocrPagePromises.delete(pageNum);
    });
  state.ocrPagePromises.set(pageNum, promise);
  return promise;
}

function shouldUseOcrFallback(textItems) {
  if (!Array.isArray(textItems) || textItems.length === 0) {
    return true;
  }
  const charCount = textItems.reduce((sum, item) => sum + Number(item.searchableLength || 0), 0);
  return charCount <= 1;
}

function buildSearchPreview(items, hitItemIndexes) {
  if (!items.length || !hitItemIndexes.length) {
    return "";
  }
  const startIndex = Math.max(0, hitItemIndexes[0] - 2);
  const endIndex = Math.min(items.length - 1, hitItemIndexes[hitItemIndexes.length - 1] + 2);
  const raw = items
    .slice(startIndex, endIndex + 1)
    .map((item) => String(item.str || "").replace(/\s+/g, " ").trim())
    .join(" ");
  return raw.length <= 120 ? raw : `${raw.slice(0, 117)}...`;
}

function renderSearchResultList() {
  els.searchResultList.innerHTML = "";
  state.searchMatches.forEach((match, index) => {
    const li = document.createElement("li");
    li.className = "search-result-item";
    if (index === state.activeSearchIndex) {
      li.classList.add("active");
    }
    const displayIndex = state.pageOrder.indexOf(match.pageNum) + 1;
    li.textContent =
      state.language === "en" ? `Page ${displayIndex}: ${match.text}` : `${displayIndex}페이지: ${match.text}`;
    li.addEventListener("click", async () => {
      await activateSearchResult(index, true);
    });
    els.searchResultList.appendChild(li);
  });
}

async function activateSearchResult(index, shouldScroll) {
  if (index < 0 || index >= state.searchMatches.length) {
    return;
  }
  state.activeSearchIndex = index;
  const match = state.searchMatches[index];
  await goToPage(match.pageNum, shouldScroll);
  if (shouldScroll) {
    await scrollActiveMatchToCenter(match);
  }
  renderSearchResultList();
  for (const pageNum of state.pageOrder) {
    drawSearchHighlightsForPage(pageNum);
  }
  const pageNumber = state.pageOrder.indexOf(match.pageNum) + 1;
  setStatus(
    state.language === "en"
      ? `Result ${index + 1}/${state.searchMatches.length} - page ${pageNumber}`
      : `검색 결과 ${index + 1}/${state.searchMatches.length} - ${pageNumber}페이지`
  );
}

async function scrollActiveMatchToCenter(match) {
  if (!match) {
    return;
  }
  const view = state.pageViews.get(match.pageNum);
  if (!view || !view.viewport || !match.segments?.length) {
    return;
  }
  const firstSegment = match.segments[0];
  const items = state.searchPageCache.get(match.pageNum) || state.textItemsCache.get(match.pageNum) || state.ocrWordCache.get(match.pageNum) || [];
  const item = items[firstSegment.itemIndex];
  if (!item) {
    return;
  }
  const rect = itemSegmentRectInViewport(item, view.viewport, firstSegment.startOffset, firstSegment.endOffset);
  const targetCenter = view.wrap.offsetTop + rect.top + rect.height / 2;
  const nextScrollTop = Math.max(0, targetCenter - els.viewerPanel.clientHeight * 0.45);
  els.viewerPanel.scrollTo({ top: nextScrollTop, behavior: "smooth" });
}

async function performSearch(rawQuery, jumpFirst = true) {
  const query = normalizeSearchText(rawQuery);
  const queryNeedle = query.toLowerCase();
  state.searchQuery = query;
  state.searchMatches = [];
  state.perPageMatchIndexes = new Map();
  state.searchPageCache = new Map();
  state.activeSearchIndex = -1;

  if (!queryNeedle) {
    renderSearchResultList();
    updateSearchCountText();
    updateThumbnailSearchMark();
    for (const pageNum of state.pageOrder) {
      drawSearchHighlightsForPage(pageNum);
    }
    setStatus(state.language === "en" ? "Please enter a search keyword." : "검색어를 입력해 주세요.");
    return;
  }

  if (!state.isFullScreen) {
    state.searchPanelVisible = true;
    applyPanelLayout();
    persistLayoutState();
  }

  setStatus(state.language === "en" ? "Searching..." : "검색 중...");
  let ocrUsed = false;
  for (const pageNum of state.pageOrder) {
    let items = await ensureTextItems(pageNum);
    if (shouldUseOcrFallback(items)) {
      ocrUsed = true;
      if (!state.ocrWordCache.has(pageNum)) {
        setStatus(t("ocrSearching"));
      }
      items = await ensureOcrItems(pageNum);
    }
    state.searchPageCache.set(pageNum, items);
    if (!items.length) {
      continue;
    }
    for (const item of items) {
      const text = item.searchable || "";
      if (!text) {
        continue;
      }
      let from = 0;
      while (from < text.length) {
        const found = text.indexOf(queryNeedle, from);
        if (found < 0) {
          break;
        }
        const foundEnd = found + queryNeedle.length;
        const hitItemIndexes = [item.index];
        const segments = [
          {
            itemIndex: item.index,
            startOffset: found,
            endOffset: foundEnd
          }
        ];
        const preview = buildSearchPreview(items, hitItemIndexes);
        const matchIndex = state.searchMatches.length;
        state.searchMatches.push({
          pageNum,
          itemIndexes: hitItemIndexes,
          segments,
          text: preview || "(검색 결과)"
        });
        if (!state.perPageMatchIndexes.has(pageNum)) {
          state.perPageMatchIndexes.set(pageNum, []);
        }
        state.perPageMatchIndexes.get(pageNum).push(matchIndex);
        from = found + Math.max(1, queryNeedle.length);
      }
    }
  }

  updateSearchCountText();
  renderSearchResultList();
  updateThumbnailSearchMark();
  for (const pageNum of state.pageOrder) {
    drawSearchHighlightsForPage(pageNum);
  }

  if (state.searchMatches.length === 0) {
    setStatus(
      state.language === "en" ? `No results for "${query}".` : `"${query}" 검색 결과가 없습니다.`,
      true
    );
    return;
  }

  if (ocrUsed) {
    setStatus(state.language === "en" ? `Found ${state.searchMatches.length} result(s).` : `${state.searchMatches.length}개 결과를 찾았습니다.`);
  }

  if (jumpFirst) {
    await activateSearchResult(0, true);
  } else {
    state.activeSearchIndex = clamp(state.activeSearchIndex, 0, state.searchMatches.length - 1);
  }
}

async function searchNext(direction) {
  if (state.searchMatches.length === 0) {
    await performSearch(els.searchInput.value, true);
    return;
  }
  let nextIndex = state.activeSearchIndex + direction;
  if (nextIndex < 0) {
    nextIndex = state.searchMatches.length - 1;
  }
  if (nextIndex >= state.searchMatches.length) {
    nextIndex = 0;
  }
  await activateSearchResult(nextIndex, true);
}
async function movePageOrder(draggedPageNum, targetPageNum) {
  const fromIndex = state.pageOrder.indexOf(draggedPageNum);
  const toIndex = state.pageOrder.indexOf(targetPageNum);
  if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) {
    return;
  }

  pushHistorySnapshot();
  state.pageOrder.splice(fromIndex, 1);
  state.pageOrder.splice(toIndex, 0, draggedPageNum);
  state.saveDirty = true;

  await rebuildPageViews();
  await renderThumbnails();
  updatePageBadges();
  await goToPage(draggedPageNum, false);
  if (state.searchQuery) {
    await performSearch(state.searchQuery, false);
  }
  setStatus("페이지 순서를 변경했습니다.");
}

async function deleteCurrentPage() {
  if (!state.pdfDoc || state.pageOrder.length <= 1) {
    return;
  }
  const removeIndex = getCurrentDisplayIndex();
  if (removeIndex < 0) {
    return;
  }
  pushHistorySnapshot();
  const removedPageNum = state.pageOrder[removeIndex];
  state.pageOrder.splice(removeIndex, 1);
  state.pageRotations.delete(removedPageNum);
  state.annotations.delete(removedPageNum);
  state.textItemsCache.delete(removedPageNum);
  state.searchPageCache.delete(removedPageNum);
  state.pageViews.delete(removedPageNum);
  state.saveDirty = true;

  const fallbackPage = state.pageOrder[Math.max(0, Math.min(removeIndex, state.pageOrder.length - 1))];
  state.currentPage = fallbackPage;

  await rebuildPageViews();
  await renderThumbnails();
  updatePageBadges();
  await goToPage(fallbackPage, false);
  if (state.searchQuery) {
    await performSearch(state.searchQuery, false);
  }
  setStatus("현재 페이지를 삭제했습니다.");
}

async function rotateCurrentPage(delta) {
  if (!state.pdfDoc) {
    return;
  }
  const current = getRotation(state.currentPage);
  const next = (current + delta + 360) % 360;
  pushHistorySnapshot();
  state.pageRotations.set(state.currentPage, next);
  state.saveDirty = true;
  await renderPage(state.currentPage);

  const thumb = state.thumbnails.get(state.currentPage);
  if (thumb) {
    const canvas = thumb.querySelector("canvas");
    if (canvas) {
      await renderThumbnail(state.currentPage, canvas);
    }
  }
  setStatus("현재 페이지를 회전했습니다.");
}

function buildZoomAnchorFromClient(clientX, clientY) {
  const rect = els.viewerPanel.getBoundingClientRect();
  return {
    x: clientX - rect.left + els.viewerPanel.scrollLeft,
    y: clientY - rect.top + els.viewerPanel.scrollTop,
    dx: clientX - rect.left,
    dy: clientY - rect.top
  };
}

function normalizeWheelDelta(event) {
  if (event.deltaMode === WheelEvent.DOM_DELTA_LINE) {
    return event.deltaY * 16;
  }
  if (event.deltaMode === WheelEvent.DOM_DELTA_PAGE) {
    return event.deltaY * els.viewerPanel.clientHeight;
  }
  return event.deltaY;
}

function requestWheelZoomApply() {
  if (state.wheelZoomRaf) {
    return;
  }
  state.wheelZoomRaf = requestAnimationFrame(() => {
    state.wheelZoomRaf = 0;
    if (!state.pdfDoc || Math.abs(state.wheelZoomDelta) < 0.01) {
      state.wheelZoomDelta = 0;
      return;
    }
    const delta = state.wheelZoomDelta;
    const anchor = state.wheelZoomAnchor;
    state.wheelZoomDelta = 0;
    const factor = clamp(Math.exp(-delta * 0.0023), 0.78, 1.42);
    queueZoomJob(state.scale * factor, anchor);
    if (Math.abs(state.wheelZoomDelta) >= 0.01) {
      requestWheelZoomApply();
    }
  });
}

function queueZoomJob(scale, anchor) {
  state.pendingZoomJob = {
    scale,
    anchor,
    options: { prioritizeVisible: true }
  };
  if (!state.zoomJobRunning) {
    runZoomJobs().catch(() => {});
  }
}

async function runZoomJobs() {
  if (state.zoomJobRunning) {
    return;
  }
  state.zoomJobRunning = true;
  try {
    while (state.pendingZoomJob) {
      const job = state.pendingZoomJob;
      state.pendingZoomJob = null;
      await zoomTo(job.scale, job.anchor, job.options);
    }
  } finally {
    state.zoomJobRunning = false;
  }
}

function handleCtrlWheelZoom(event) {
  if (!event.ctrlKey || !state.pdfDoc) {
    return;
  }
  event.preventDefault();
  state.wheelZoomAnchor = buildZoomAnchorFromClient(event.clientX, event.clientY);
  state.wheelZoomDelta += normalizeWheelDelta(event);
  requestWheelZoomApply();
}

function isViewerEventTarget(event) {
  const target = event.target;
  return Boolean(target instanceof Node && els.viewerPanel.contains(target));
}

function handleViewerWheel(event) {
  if (!state.pdfDoc || !isViewerEventTarget(event)) {
    return;
  }
  if (event.ctrlKey) {
    handleCtrlWheelZoom(event);
    return;
  }
  if (!isSinglePageFullscreen()) {
    return;
  }
  const delta = normalizeWheelDelta(event);
  if (Math.abs(delta) < 1.2) {
    return;
  }
  const now = Date.now();
  if (now - state.singlePageWheelStepTime < 90) {
    event.preventDefault();
    return;
  }
  state.singlePageWheelStepTime = now;
  event.preventDefault();
  stepPage(delta > 0 ? 1 : -1, false).catch(() => {});
}

async function zoomTo(newScale, anchorInput = null, options = {}) {
  if (!state.pdfDoc) {
    return;
  }
  const nextScale = clamp(newScale, 0.25, 6);
  if (Math.abs(nextScale - state.scale) < 0.001) {
    return;
  }

  let anchor = null;
  if (anchorInput && typeof anchorInput.clientX === "number" && typeof anchorInput.clientY === "number") {
    anchor = buildZoomAnchorFromClient(anchorInput.clientX, anchorInput.clientY);
  } else if (
    anchorInput &&
    typeof anchorInput.x === "number" &&
    typeof anchorInput.y === "number" &&
    typeof anchorInput.dx === "number" &&
    typeof anchorInput.dy === "number"
  ) {
    anchor = anchorInput;
  }

  const oldScale = state.scale;
  state.scale = nextScale;
  setZoomMode(options.zoomMode || "manual");
  await renderAllPages({ prioritizeVisible: options.prioritizeVisible !== false });
  updateToolbarState();

  if (anchor) {
    const ratio = nextScale / oldScale;
    els.viewerPanel.scrollLeft = anchor.x * ratio - anchor.dx;
    els.viewerPanel.scrollTop = anchor.y * ratio - anchor.dy;
  } else if (state.currentPage) {
    await goToPage(state.currentPage, false);
  }
}

function toggleFullscreenViewMode() {
  state.fullScreenViewMode = state.fullScreenViewMode === "single" ? "continuous" : "single";
  localStorage.setItem(storage.fullscreenMode, state.fullScreenViewMode);
  updateFullscreenButtons();
  applyPageVisibility();
  persistLayoutState();
  if (state.isFullScreen) {
    setZoomMode("fit");
    state.fullScreenAutoFitDone = false;
    queueLayoutRecoveryRender({ forceFit: true });
    return;
  }
  queueLayoutRecoveryRender();
}

async function buildEditedPdfBytes() {
  if (!state.sourceBytes) {
    throw new Error("저장할 PDF가 없습니다.");
  }

  const srcDoc = await PDFDocument.load(state.sourceBytes);
  const outputDoc = await PDFDocument.create();
  const font = await outputDoc.embedFont(StandardFonts.Helvetica);

  for (const pageNum of state.pageOrder) {
    const [copiedPage] = await outputDoc.copyPages(srcDoc, [pageNum - 1]);
    const extraRotation = getRotation(pageNum);
    if (extraRotation) {
      const base = copiedPage.getRotation()?.angle || 0;
      copiedPage.setRotation(degrees((base + extraRotation) % 360));
    }
    outputDoc.addPage(copiedPage);
  }

  for (let displayIndex = 0; displayIndex < state.pageOrder.length; displayIndex += 1) {
    const originalPageNum = state.pageOrder[displayIndex];
    const annotations = state.annotations.get(originalPageNum);
    if (!annotations) {
      continue;
    }
    const page = outputDoc.getPage(displayIndex);

    for (const mark of annotations.highlights) {
      const left = Math.min(mark.x1, mark.x2);
      const bottom = Math.min(mark.y1, mark.y2);
      const width = Math.abs(mark.x1 - mark.x2);
      const height = Math.abs(mark.y1 - mark.y2);
      if (width < 1 || height < 1) {
        continue;
      }
      page.drawRectangle({
        x: left,
        y: bottom,
        width,
        height,
        color: rgb(1, 0.9, 0.2),
        opacity: 0.35
      });
    }

    for (const pen of annotations.pens) {
      for (let i = 1; i < pen.points.length; i += 1) {
        const start = pen.points[i - 1];
        const end = pen.points[i];
        page.drawLine({
          start: { x: start.x, y: start.y },
          end: { x: end.x, y: end.y },
          color: rgb(1, 0.28, 0.28),
          thickness: 1.8,
          opacity: 0.92
        });
      }
    }

    for (const note of annotations.texts) {
      const size = 12;
      const textWidth = font.widthOfTextAtSize(note.text, size);
      page.drawRectangle({
        x: note.x,
        y: note.y - size - 6,
        width: textWidth + 12,
        height: size + 8,
        color: rgb(1, 0.95, 0.6),
        opacity: 0.95
      });
      page.drawText(note.text, {
        x: note.x + 6,
        y: note.y - size - 1,
        size,
        font,
        color: rgb(0.1, 0.1, 0.1)
      });
    }
  }

  return outputDoc.save();
}

async function savePdfAs() {
  if (!state.pdfDoc) {
    return;
  }
  const savePath = await window.lookupAPI.savePdfDialog({
    defaultPath: makeDefaultSavePath(state.filePath)
  });
  if (!savePath) {
    return;
  }
  setStatus("PDF 저장 중...");
  const bytes = await buildEditedPdfBytes();
  await window.lookupAPI.writePdfFile(savePath, bytes);
  await loadPdfFromBytes(bytes, savePath);
  setStatus(`저장 완료: ${fileNameFromPath(savePath)}`);
}

async function savePdfOverwrite() {
  if (!state.pdfDoc) {
    return;
  }
  if (state.sourceExt !== ".pdf") {
    setStatus(
      state.language === "en"
        ? "Overwrite is available only when original file is PDF. Use Save As."
        : "덮어쓰기는 원본이 PDF일 때만 가능합니다. 다른 이름 저장을 사용해 주세요.",
      true
    );
    return;
  }
  if (!state.filePath) {
    await savePdfAs();
    return;
  }
  const confirmed = await window.lookupAPI.confirmOverwrite({
    message: "현재 PDF 파일을 덮어쓰시겠습니까?",
    detail: fileNameFromPath(state.filePath)
  });
  if (!confirmed) {
    return;
  }
  setStatus("원본 파일 덮어쓰기 저장 중...");
  const bytes = await buildEditedPdfBytes();
  await window.lookupAPI.writePdfFile(state.filePath, bytes);
  await loadPdfFromBytes(bytes, state.filePath);
  setStatus("원본 파일 덮어쓰기 저장 완료");
}

async function openPrintPreview() {
  if (!state.pdfDoc) {
    return false;
  }
  setStatus(t("printPreparing"));
  const bytes = await buildEditedPdfBytes();
  const currentName = fileNameFromPath(state.filePath) || "document.pdf";
  const safeBase = currentName.replace(/\.[^./\\]+$/, "");
  const fileName = `${safeBase}-print-preview.pdf`;
  const result = await window.lookupAPI.printPreview(bytes, fileName);
  if (!result?.ok) {
    setStatus(result?.message || t("printFailed"), true);
    return false;
  }
  setStatus(t("printOpened"));
  return true;
}

async function checkForUpdatesFromUI() {
  if (state.updateBusy) {
    setStatus(t("updateBusyMessage"), true);
    return false;
  }
  state.updateBusy = true;
  applyUpdateVisualState("checking", "checking", 0);
  setStatus(t("updateChecking"));
  const result = await window.lookupAPI.checkForUpdates();
  if (!result.ok) {
    state.updateBusy = false;
    applyUpdateVisualState("error", "error", state.updateStage === "downloading" ? undefined : 0);
    setStatus(
      state.language === "en" ? `Update check failed: ${result.message}` : `업데이트 확인 실패: ${result.message}`,
      true
    );
    return false;
  }
  return true;
}

async function loadPdfFromBytes(rawBytes, filePath, meta = {}) {
  const bytes = toUint8Array(rawBytes);
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  const pdfDoc = await loadingTask.promise;

  state.pdfDoc = pdfDoc;
  state.sourceBytes = bytes;
  state.filePath = filePath || "";
  state.sourceExt = meta.sourceExt || ".pdf";
  state.sourceConverted = Boolean(meta.converted);
  state.sourceConvertMode = meta.convertMode || (state.sourceExt === ".pdf" ? "native" : "fallback");
  state.pageOrder = Array.from({ length: pdfDoc.numPages }, (_v, i) => i + 1);
  state.pageCache.clear();
  state.pageRotations.clear();
  state.annotations.clear();
  state.textItemsCache.clear();
  state.searchPageCache.clear();
  state.ocrWordCache.clear();
  state.ocrPagePromises.clear();
  for (const renderTask of state.pageRenderTasks.values()) {
    if (renderTask && typeof renderTask.cancel === "function") {
      try {
        renderTask.cancel();
      } catch (_error) {
        // ignore cancellation errors
      }
    }
  }
  for (const renderTask of state.thumbRenderTasks.values()) {
    if (renderTask && typeof renderTask.cancel === "function") {
      try {
        renderTask.cancel();
      } catch (_error) {
        // ignore cancellation errors
      }
    }
  }
  state.pageRenderTasks.clear();
  state.thumbRenderTasks.clear();
  state.ocrErrorShown = false;
  clearSearchState();
  state.scale = 1;
  setZoomMode("fit");
  state.fullScreenAutoFitDone = false;
  state.currentPage = 1;
  state.saveDirty = false;
  clearHistory();
  closeTextMemoEditor();
  await syncWindowTitle(filePath || "");

  els.emptyHint.classList.add("hidden");
  await rebuildPageViews();
  await renderThumbnails();
  updatePageBadges();
  await goToPage(state.currentPage, false);
  const convertTail =
    meta && meta.converted
      ? state.language === "en"
        ? ` / Converted from ${String(meta.sourceExt || "").replace(".", "").toUpperCase()}`
        : ` / ${String(meta.sourceExt || "").replace(".", "").toUpperCase()} 변환 열람`
      : "";
  setStatus(`열림: ${fileNameFromPath(filePath)} (${state.pageOrder.length}페이지)${convertTail}`);
  if (meta.warningMessage) {
    setStatus(meta.warningMessage);
  }
}

async function loadDocumentFromPath(filePath) {
  if (!filePath || typeof filePath !== "string") {
    return false;
  }

  const resolvedPath = filePath.trim();
  const ext = getFileExt(resolvedPath);
  setStatus(state.language === "en" ? "Opening document..." : "문서를 열고 있습니다...");

  try {
    const payload = await window.lookupAPI.openDocument(resolvedPath);
    if (payload?.ok === true && payload.data) {
      await loadPdfFromBytes(payload.data, payload.sourcePath || resolvedPath, payload);
      return true;
    }
    if (payload && payload.data && payload.ok !== false) {
      await loadPdfFromBytes(payload.data, payload.sourcePath || resolvedPath, payload);
      return true;
    }

    if (ext === ".pdf") {
      const directBytes = await window.lookupAPI.readPdfFile(resolvedPath);
      await loadPdfFromBytes(directBytes, resolvedPath, {
        sourcePath: resolvedPath,
        sourceExt: ".pdf",
        converted: false,
        convertMode: "native"
      });
      setStatus(t("openErrorPdfFallback"));
      return true;
    }

    throw new Error(mapOpenErrorMessage(payload?.errorCode, payload?.message));
  } catch (error) {
    const fallbackMessage = mapOpenErrorMessage("", error?.message || "");
    setStatus(
      state.language === "en"
        ? `Failed to open file: ${fallbackMessage}`
        : `파일 열기 실패: ${fallbackMessage}`,
      true
    );
    if (!state.pdfDoc) {
      els.emptyHint.classList.remove("hidden");
    }
    return false;
  }
}

async function openFileDialog() {
  const selectedPath = await window.lookupAPI.openDocumentDialog();
  if (!selectedPath) {
    return;
  }
  await loadDocumentFromPath(selectedPath);
}

function fileUriToPath(text) {
  const raw = String(text || "").trim();
  if (!raw) {
    return "";
  }
  const line = raw.split(/\r?\n/).find((entry) => entry && !entry.startsWith("#")) || "";
  if (!line) {
    return "";
  }
  if (/^[a-zA-Z]:\\/.test(line) || /^[a-zA-Z]:\//.test(line)) {
    return line.replaceAll("/", "\\");
  }
  if (!line.toLowerCase().startsWith("file://")) {
    return "";
  }
  let decoded = decodeURIComponent(line.replace(/^file:\/\/+/i, ""));
  decoded = decoded.replace(/\//g, "\\");
  if (/^[a-zA-Z]:\\/.test(decoded)) {
    return decoded;
  }
  if (/^\\{2}/.test(decoded)) {
    return decoded;
  }
  if (/^\\[a-zA-Z]:\\/.test(decoded)) {
    return decoded.slice(1);
  }
  return decoded;
}

function extractDropFilePath(event) {
  const list = Array.from(event.dataTransfer?.files || []);
  const direct = list.find((file) => typeof file?.path === "string" && file.path.trim().length > 0);
  if (direct?.path) {
    return direct.path;
  }
  const uriList = event.dataTransfer?.getData("text/uri-list");
  const fromUri = fileUriToPath(uriList);
  if (fromUri) {
    return fromUri;
  }
  const text = event.dataTransfer?.getData("text/plain");
  return fileUriToPath(text);
}

function toggleLeftPanelVisibility() {
  if (state.isFullScreen) {
    state.fullscreenThumbVisible = !state.fullscreenThumbVisible;
  } else {
    state.thumbPanelVisible = !state.thumbPanelVisible;
  }
  applyPanelLayout();
  persistLayoutState();
  if (getEffectiveLeftPanelVisible()) {
    queueThumbnailRerender();
  }
  const forceFit = state.isFullScreen && state.zoomMode !== "manual";
  if (forceFit) {
    setZoomMode("fit");
    state.fullScreenAutoFitDone = false;
  }
  queueLayoutRecoveryRender({
    forceFit
  });
}

function toggleRightPanelVisibility() {
  if (state.isFullScreen) {
    return;
  }
  state.searchPanelVisible = !state.searchPanelVisible;
  applyPanelLayout();
  persistLayoutState();
  queueLayoutRecoveryRender();
}

function handlePanelResizeStart(side, startEvent) {
  if (!(startEvent instanceof PointerEvent)) {
    return;
  }
  startEvent.preventDefault();
  state.activeResizer = side;
  const startX = startEvent.clientX;
  const startLeftWidth = state.leftPanelWidth;
  const startRightWidth = state.rightPanelWidth;
  const workspaceRect = els.workspace.getBoundingClientRect();

  const activeResizerEl = side === "left" ? els.leftResizer : els.rightResizer;
  activeResizerEl.classList.add("active");
  activeResizerEl.setPointerCapture(startEvent.pointerId);

  const onMove = (event) => {
    const delta = event.clientX - startX;
    if (side === "left") {
      const maxWidth = Math.max(220, workspaceRect.width - 380);
      state.leftPanelWidth = clamp(startLeftWidth + delta, 180, maxWidth);
      if (!state.isFullScreen) {
        state.thumbPanelVisible = true;
      } else {
        state.fullscreenThumbVisible = true;
      }
    } else {
      const maxWidth = Math.max(260, workspaceRect.width - 380);
      state.rightPanelWidth = clamp(startRightWidth - delta, 220, maxWidth);
      if (!state.isFullScreen) {
        state.searchPanelVisible = true;
      }
    }
    applyPanelLayout();
    if (side === "left") {
      queueThumbnailRerender();
    }
  };

  const onFinish = (event) => {
    state.activeResizer = null;
    activeResizerEl.classList.remove("active");
    activeResizerEl.releasePointerCapture(startEvent.pointerId);
    window.removeEventListener("pointermove", onMove);
    window.removeEventListener("pointerup", onFinish);
    window.removeEventListener("pointercancel", onFinish);
    persistLayoutState();
    if (side === "left") {
      queueThumbnailRerender();
    }
    queueLayoutRecoveryRender({
      forceFit: state.isFullScreen && side === "left" && state.zoomMode !== "manual"
    });
    if (event) {
      event.preventDefault();
    }
  };

  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onFinish);
  window.addEventListener("pointercancel", onFinish);
}

function bindPanelResizeHandles() {
  els.leftResizer.addEventListener("pointerdown", (event) => handlePanelResizeStart("left", event));
  els.rightResizer.addEventListener("pointerdown", (event) => handlePanelResizeStart("right", event));
}

let resizeDebounceTimer = 0;
function handleWindowResize() {
  if (resizeDebounceTimer) {
    clearTimeout(resizeDebounceTimer);
  }
  resizeDebounceTimer = setTimeout(() => {
    resizeDebounceTimer = 0;
    if (getEffectiveLeftPanelVisible()) {
      queueThumbnailRerender();
    }
    queueLayoutRecoveryRender({ preserveZoom: state.isFullScreen && state.zoomMode === "manual" });
  }, 80);
}

function bindToolbarActions() {
  els.openFileBtn.addEventListener("click", () => openFileDialog());
  els.saveAsBtn.addEventListener("click", () => savePdfAs().catch((error) => setStatus(error.message, true)));
  els.saveOverwriteBtn.addEventListener("click", () =>
    savePdfOverwrite().catch((error) => setStatus(error.message, true))
  );
  els.printBtn.addEventListener("click", () => openPrintPreview().catch((error) => setStatus(error.message, true)));

  els.prevPageBtn.addEventListener("click", async () => {
    const index = getCurrentDisplayIndex();
    if (index > 0) {
      await goToPage(state.pageOrder[index - 1], true);
    }
  });
  els.nextPageBtn.addEventListener("click", async () => {
    const index = getCurrentDisplayIndex();
    if (index >= 0 && index < state.pageOrder.length - 1) {
      await goToPage(state.pageOrder[index + 1], true);
    }
  });
  els.pageInput.addEventListener("change", async () => {
    if (!state.pdfDoc) {
      return;
    }
    const wantedIndex = Number.parseInt(els.pageInput.value, 10);
    if (Number.isNaN(wantedIndex)) {
      updateToolbarState();
      return;
    }
    const clampedIndex = clamp(wantedIndex - 1, 0, state.pageOrder.length - 1);
    await goToPage(state.pageOrder[clampedIndex], true);
  });

  els.zoomInBtn.addEventListener("click", () => zoomTo(state.scale * 1.12, null, { prioritizeVisible: true }).catch(() => {}));
  els.zoomOutBtn.addEventListener("click", () => zoomTo(state.scale / 1.12, null, { prioritizeVisible: true }).catch(() => {}));
  els.zoomResetBtn.addEventListener("click", () => zoomTo(1).catch(() => {}));

  els.rotateLeftBtn.addEventListener("click", () => rotateCurrentPage(-90).catch(() => {}));
  els.rotateRightBtn.addEventListener("click", () => rotateCurrentPage(90).catch(() => {}));
  els.deletePageBtn.addEventListener("click", () => deleteCurrentPage().catch(() => {}));

  els.editModeButtons.forEach((button) => {
    button.addEventListener("click", () => setEditingMode(button.dataset.mode));
  });
  if (els.undoBtn) {
    els.undoBtn.addEventListener("click", () => undoLastAction().catch(() => {}));
  }
  if (els.redoBtn) {
    els.redoBtn.addEventListener("click", () => redoLastAction().catch(() => {}));
  }

  els.searchBtn.addEventListener("click", () => performSearch(els.searchInput.value, true).catch(() => {}));
  els.searchPrevBtn.addEventListener("click", () => searchNext(-1).catch(() => {}));
  els.searchNextBtn.addEventListener("click", () => searchNext(1).catch(() => {}));
  els.searchInput.addEventListener("keydown", async (event) => {
    if (event.key === "Enter") {
      await performSearch(els.searchInput.value, true);
    }
  });

  els.toggleThumbPanelBtn.addEventListener("click", () => {
    toggleLeftPanelVisibility();
  });
  els.toggleSearchPanelBtn.addEventListener("click", () => {
    toggleRightPanelVisibility();
  });

  els.toggleDarkBtn.addEventListener("click", () => {
    setDarkMode(!document.body.classList.contains("dark"));
  });

  els.toggleFullscreenBtn.addEventListener("click", () => {
    window.lookupAPI.toggleFullScreen();
  });
  els.toggleFullscreenViewModeBtn.addEventListener("click", () => {
    toggleFullscreenViewMode();
  });
  els.toggleThumbInFullscreenBtn.addEventListener("click", () => {
    toggleLeftPanelVisibility();
  });
  if (els.textMemoAddBtn) {
    els.textMemoAddBtn.addEventListener("click", () => commitTextMemoFromEditor());
  }
  if (els.textMemoCancelBtn) {
    els.textMemoCancelBtn.addEventListener("click", () => closeTextMemoEditor());
  }
  if (els.textMemoInput) {
    els.textMemoInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        commitTextMemoFromEditor();
      } else if (event.key === "Escape") {
        event.preventDefault();
        closeTextMemoEditor();
      }
    });
  }
  if (els.updateNotesCloseBtn) {
    els.updateNotesCloseBtn.addEventListener("click", () => hideUpdateNotesModal());
  }
  if (els.updateNotesModal) {
    els.updateNotesModal.addEventListener("click", (event) => {
      if (event.target === els.updateNotesModal) {
        hideUpdateNotesModal();
      }
    });
  }
}

function bindWindowActions() {
  els.viewerPanel.addEventListener("scroll", queueScrollSync);
  els.viewerPanel.addEventListener("click", () => {
    focusViewerPanel();
  });
  els.viewerPanel.addEventListener("wheel", handleViewerWheel, { passive: false });
  window.addEventListener("resize", handleWindowResize);

  els.viewerPanel.addEventListener("dragover", (event) => {
    event.preventDefault();
    els.viewerPanel.classList.add("drag-over");
  });
  els.viewerPanel.addEventListener("dragleave", () => {
    els.viewerPanel.classList.remove("drag-over");
  });
  els.viewerPanel.addEventListener("drop", async (event) => {
    event.preventDefault();
    els.viewerPanel.classList.remove("drag-over");
    const droppedPath = extractDropFilePath(event);
    if (!droppedPath) {
      setStatus(state.language === "en" ? "Failed to read dropped file path." : "드롭한 파일 경로를 읽지 못했습니다.", true);
      return;
    }
    await loadDocumentFromPath(droppedPath);
  });

  window.addEventListener("keydown", async (event) => {
    const active = document.activeElement;
    const isTypingTarget =
      active &&
      (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.isContentEditable === true);
    if (event.key === "Escape" && els.updateNotesModal && !els.updateNotesModal.classList.contains("hidden")) {
      event.preventDefault();
      hideUpdateNotesModal();
      return;
    }
    if (event.key === "Escape" && state.pendingTextMemo) {
      event.preventDefault();
      closeTextMemoEditor();
      return;
    }
    if (event.ctrlKey && !event.shiftKey && event.key.toLowerCase() === "z" && !isTypingTarget) {
      event.preventDefault();
      await undoLastAction();
      return;
    }
    if (event.ctrlKey && !event.shiftKey && event.key.toLowerCase() === "y" && !isTypingTarget) {
      event.preventDefault();
      await redoLastAction();
      return;
    }
    if (event.ctrlKey && event.key.toLowerCase() === "f") {
      event.preventDefault();
      els.searchInput.focus();
      els.searchInput.select();
      return;
    }
    if (event.ctrlKey && event.key.toLowerCase() === "s") {
      event.preventDefault();
      await savePdfOverwrite();
      return;
    }
    if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === "s") {
      event.preventDefault();
      await savePdfAs();
      return;
    }
    if (event.ctrlKey && event.key.toLowerCase() === "p") {
      event.preventDefault();
      await openPrintPreview();
      return;
    }
    if (event.key === "F11") {
      event.preventDefault();
      await window.lookupAPI.toggleFullScreen();
      return;
    }
    if (event.key === "Escape" && state.isFullScreen) {
      event.preventDefault();
      await window.lookupAPI.setFullScreen(false);
      return;
    }
    if (
      isSinglePageFullscreen() &&
      ["ArrowUp", "ArrowLeft", "PageUp"].includes(event.key) &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      !isTypingTarget
    ) {
      event.preventDefault();
      await stepPage(-1, false);
      return;
    }
    if (
      isSinglePageFullscreen() &&
      ["ArrowDown", "ArrowRight", "PageDown"].includes(event.key) &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      !isTypingTarget
    ) {
      event.preventDefault();
      await stepPage(1, false);
      return;
    }
    if (event.key === "PageUp") {
      event.preventDefault();
      await stepPage(-1, true);
      return;
    }
    if (event.key === "PageDown") {
      event.preventDefault();
      await stepPage(1, true);
      return;
    }
    if (event.key === "Delete" && state.pdfDoc && !event.ctrlKey && !event.metaKey) {
      if (!isTypingTarget) {
        await deleteCurrentPage();
      }
    }
  });

  window.addEventListener("beforeunload", () => {
    if (state.ocrWorker) {
      Promise.resolve(state.ocrWorker.terminate()).catch(() => {});
      state.ocrWorker = null;
      state.ocrWorkerReady = false;
    }
  });
}

function bindMainProcessEvents() {
  window.lookupAPI.onSystemOpenFile((filePath) => {
    loadDocumentFromPath(filePath).catch((error) => {
      setStatus(error.message, true);
    });
  });

  window.lookupAPI.onMenuAction((action) => {
    const map = {
      "open-file": () => openFileDialog(),
      "save-as": () => savePdfAs(),
      "save-overwrite": () => savePdfOverwrite(),
      print: () => openPrintPreview(),
      "prev-page": async () => {
        await stepPage(-1, true);
      },
      "next-page": async () => {
        await stepPage(1, true);
      },
      "zoom-in": () => zoomTo(state.scale * 1.12, null, { prioritizeVisible: true }),
      "zoom-out": () => zoomTo(state.scale / 1.12, null, { prioritizeVisible: true }),
      "zoom-reset": () => zoomTo(1),
      "toggle-dark": () => setDarkMode(!document.body.classList.contains("dark")),
      "toggle-fullscreen-view-mode": () => toggleFullscreenViewMode(),
      "toggle-thumb-panel": () => toggleLeftPanelVisibility(),
      "toggle-search-panel": () => toggleRightPanelVisibility(),
      "check-update": async () => {
        const ok = await checkForUpdatesFromUI();
        if (ok) {
          setStatus(t("updateStarted"));
        }
      },
      "copy-developer-email": async () => {
        await window.lookupAPI.copyText("lamsaiku65@gmail.com");
        setStatus(t("copiedContact"));
      },
      "set-language-ko": async () => {
        await setLanguage("ko", true);
        setStatus(t("languageChangedKo"));
      },
      "set-language-en": async () => {
        await setLanguage("en", true);
        setStatus(t("languageChangedEn"));
      }
    };
    const fn = map[action];
    if (fn) {
      Promise.resolve(fn()).catch((error) => {
        setStatus(error?.message || "명령 실행 중 오류", true);
      });
    }
  });

  window.lookupAPI.onFullScreenChanged((isFullScreen) => {
    const wasFullScreen = state.isFullScreen;
    state.isFullScreen = isFullScreen;
    document.body.classList.toggle("fullscreen", isFullScreen);

    if (isFullScreen) {
      state.fullscreenThumbVisible = false;
      state.fullScreenAutoFitDone = false;
      setZoomMode("fit");
      els.fullscreenMiniBar.classList.remove("hidden");
    } else {
      els.fullscreenMiniBar.classList.add("hidden");
    }

    applyPanelLayout();
    ensureCurrentPageExists();
    applyPageVisibility();
    ensureCurrentPageVisibleInSingleMode();
    updateToolbarState();
    persistLayoutState();
    if (isFullScreen) {
      queueLayoutRecoveryRender({ forceFit: true });
    } else {
      queueLayoutRecoveryRender({ preserveZoom: state.zoomMode === "manual" });
    }
    if (isFullScreen || wasFullScreen) {
      focusViewerPanel();
    }
  });

  window.lookupAPI.onUpdateStatus((payload) => {
    if (!payload?.status) {
      return;
    }
    if (payload.currentVersion) {
      state.appVersion = payload.currentVersion;
    }
    if (payload.targetVersion) {
      state.updateTargetVersion = payload.targetVersion;
    }

    const normalizedStage = normalizeUpdateStage(payload.stage || payload.status);
    const busyStages = new Set(["checking", "downloading", "ready", "restarting"]);
    state.updateBusy = busyStages.has(normalizedStage);

    if (normalizedStage === "idle") {
      state.updateTargetVersion = "";
    }

    let percent = typeof payload.percent === "number" ? payload.percent : undefined;
    if (percent === undefined && ["ready", "restarting", "installed"].includes(normalizedStage)) {
      percent = 100;
    } else if (percent === undefined && normalizedStage === "checking") {
      percent = 0;
    }
    applyUpdateVisualState(payload.status, payload.stage, percent);
    updateVersionLabels();
    if (payload.message) {
      setStatus(payload.message, payload.status === "error");
    }
    if (normalizedStage === "installed") {
      showUpdateNotesModal(payload.targetVersion || state.appVersion, payload.releaseNotes || []);
    }
  });
}

async function initializeUpdateStatus() {
  state.appVersion = await window.lookupAPI.getAppVersion();
  updateVersionLabels();
  const config = await window.lookupAPI.getUpdateConfig();
  if (config?.currentVersion) {
    state.appVersion = config.currentVersion;
  }
  if (config?.targetVersion) {
    state.updateTargetVersion = config.targetVersion;
  }
  state.updateBusy = Boolean(config?.busy);
  updateVersionLabels();
  if (!config.enabled) {
    state.updateBusy = false;
    applyUpdateVisualState("disabled", "disabled", 0);
    setStatus(t("updateDisabled"));
  } else {
    if (state.updateBusy) {
      applyUpdateVisualState("checking", "checking", 0);
    } else {
      applyUpdateVisualState("idle", "idle", 0);
      showUpdateProgressBar(false);
      showUpdateBanner(false);
    }
    setStatus(`${t("updateReady")}: ${config.owner}/${config.repo}`);
  }
}

async function init() {
  try {
    const settings = await window.lookupAPI.getSettings();
    if (settings?.language === "ko" || settings?.language === "en") {
      state.language = settings.language;
    }
  } catch (_error) {
    // fallback to local storage language
  }
  localStorage.setItem(storage.language, state.language);
  applyLanguageToStaticTexts();
  applySavedDarkMode();
  bindToolbarActions();
  bindWindowActions();
  bindPanelResizeHandles();
  bindMainProcessEvents();
  setEditingMode("view");
  applyPanelLayout();
  updateToolbarState();

  const isFullScreen = await window.lookupAPI.isFullScreen();
  state.isFullScreen = isFullScreen;
  document.body.classList.toggle("fullscreen", isFullScreen);
  if (isFullScreen) {
    state.fullscreenThumbVisible = false;
    state.fullScreenAutoFitDone = false;
    setZoomMode("fit");
    els.fullscreenMiniBar.classList.remove("hidden");
  } else {
    els.fullscreenMiniBar.classList.add("hidden");
  }
  applyPanelLayout();
  updateFullscreenButtons();
  persistLayoutState();
  focusViewerPanel();
  showUpdateProgressBar(false);
  showUpdateBanner(false);
  hideUpdateNotesModal();
  setUpdateProgress(0);
  if (els.updateStageText) {
    els.updateStageText.textContent = updateStageLabel("idle");
  }
  await syncWindowTitle("");
  initializeUpdateStatus().catch((error) => {
    setStatus(error?.message || "업데이트 초기화 오류", true);
  });
}

init().catch((error) => {
  setStatus(error?.message || "초기화 오류", true);
});
