
import * as pdfjsLib from "../../node_modules/pdfjs-dist/legacy/build/pdf.mjs";
import { PDFDocument, StandardFonts, degrees, rgb } from "../../node_modules/pdf-lib/dist/pdf-lib.esm.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs",
  import.meta.url
).toString();

const state = {
  pdfDoc: null,
  sourceBytes: null,
  filePath: "",
  pageOrder: [],
  pageCache: new Map(),
  pageViews: new Map(),
  thumbnails: new Map(),
  pageRotations: new Map(),
  annotations: new Map(),
  textItemsCache: new Map(),
  searchPageCache: new Map(),
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
  fullScreenViewMode: localStorage.getItem("lookup-fullscreen-view-mode") || "continuous",
  thumbPanelVisible: true,
  fullscreenThumbVisible: false,
  scrollRaf: 0,
  saveDirty: false,
  wheelZoomBusy: false,
  mainRenderQuality: 1.35,
  thumbRenderQuality: 2
};

const els = {
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
  searchInput: document.getElementById("searchInput"),
  searchBtn: document.getElementById("searchBtn"),
  searchPrevBtn: document.getElementById("searchPrevBtn"),
  searchNextBtn: document.getElementById("searchNextBtn"),
  searchCountLabel: document.getElementById("searchCountLabel"),
  toggleFullscreenBtn: document.getElementById("toggleFullscreenBtn"),
  toggleDarkBtn: document.getElementById("toggleDarkBtn"),
  checkUpdateBtn: document.getElementById("checkUpdateBtn"),
  installUpdateBtn: document.getElementById("installUpdateBtn"),
  thumbPanel: document.getElementById("thumbPanel"),
  thumbnailList: document.getElementById("thumbnailList"),
  viewerPanel: document.getElementById("viewerPanel"),
  pagesContainer: document.getElementById("pagesContainer"),
  emptyHint: document.getElementById("emptyHint"),
  fullscreenMiniBar: document.getElementById("fullscreenMiniBar"),
  toggleFullscreenViewModeBtn: document.getElementById("toggleFullscreenViewModeBtn"),
  toggleThumbInFullscreenBtn: document.getElementById("toggleThumbInFullscreenBtn"),
  searchPanelCount: document.getElementById("searchPanelCount"),
  searchResultList: document.getElementById("searchResultList"),
  statusBar: document.getElementById("statusBar")
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
  els.statusBar.textContent = message;
  els.statusBar.style.color = isError ? "#d73333" : "";
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
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

function getCurrentDisplayIndex() {
  return state.pageOrder.indexOf(state.currentPage);
}

function getAnnotationBucket(pageNum) {
  if (!state.annotations.has(pageNum)) {
    state.annotations.set(pageNum, { highlights: [], pens: [], texts: [] });
  }
  return state.annotations.get(pageNum);
}

function setDarkMode(enabled) {
  document.body.classList.toggle("dark", enabled);
  els.toggleDarkBtn.textContent = enabled ? "라이트모드" : "다크모드";
  localStorage.setItem("lookup-dark-mode", enabled ? "1" : "0");
}

function applySavedDarkMode() {
  setDarkMode(localStorage.getItem("lookup-dark-mode") === "1");
}

function applyThumbPanelVisibility() {
  const visible = state.isFullScreen ? state.fullscreenThumbVisible : state.thumbPanelVisible;
  els.thumbPanel.classList.toggle("hidden", !visible);
  els.toggleThumbInFullscreenBtn.textContent = visible ? "미리보기 숨김" : "미리보기 표시";
}

function updateFullscreenButtons() {
  els.toggleFullscreenBtn.textContent = state.isFullScreen ? "전체화면 종료" : "전체화면";
  els.toggleFullscreenViewModeBtn.textContent =
    state.fullScreenViewMode === "single" ? "전체화면: 현재 페이지" : "전체화면: 연속 스크롤";
}

function applyPageVisibility() {
  const singleMode = state.isFullScreen && state.fullScreenViewMode === "single";
  for (const [pageNum, view] of state.pageViews.entries()) {
    const hidden = singleMode && pageNum !== state.currentPage;
    view.wrap.classList.toggle("hidden-page", hidden);
  }
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
  els.saveOverwriteBtn.disabled = !hasDoc;
  els.printBtn.disabled = !hasDoc;
  els.searchBtn.disabled = !hasDoc;
  els.searchInput.disabled = !hasDoc;
  els.searchPrevBtn.disabled = !hasDoc || state.searchMatches.length <= 1;
  els.searchNextBtn.disabled = !hasDoc || state.searchMatches.length <= 1;

  els.pageCountLabel.textContent = `/ ${hasDoc ? total : 0}`;
  els.pageInput.value = `${hasDoc ? Math.max(1, currentIndex + 1) : 1}`;
  els.zoomLabel.textContent = `${Math.round(state.scale * 100)}%`;

  updateFullscreenButtons();
}

function updateActiveThumbnail() {
  for (const [pageNum, thumb] of state.thumbnails.entries()) {
    thumb.classList.toggle("active", pageNum === state.currentPage);
  }
}

function updateSearchCountText() {
  const count = state.searchMatches.length;
  els.searchCountLabel.textContent = `결과 ${count}개`;
  els.searchPanelCount.textContent = `${count}건`;
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
  const viewport = page.getViewport({ scale: state.scale, rotation: getRotation(pageNum) });
  view.viewport = viewport;

  view.wrap.style.width = `${viewport.width}px`;
  view.wrap.style.height = `${viewport.height}px`;

  const dpr = window.devicePixelRatio || 1;
  const renderScale = dpr * state.mainRenderQuality;
  const canvas = view.canvas;
  canvas.width = Math.floor(viewport.width * renderScale);
  canvas.height = Math.floor(viewport.height * renderScale);
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;

  view.annotationCanvas.width = canvas.width;
  view.annotationCanvas.height = canvas.height;
  view.annotationCanvas.style.width = canvas.style.width;
  view.annotationCanvas.style.height = canvas.style.height;

  const renderToken = ++view.renderToken;
  const context = canvas.getContext("2d", { alpha: false });
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  await page.render({
    canvasContext: context,
    viewport,
    transform: renderScale === 1 ? null : [renderScale, 0, 0, renderScale, 0, 0]
  }).promise;
  if (renderToken !== view.renderToken || renderToken !== state.renderVersion) {
    return;
  }

  drawAnnotationsForPage(pageNum);
  drawSearchHighlightsForPage(pageNum);
}

async function renderAllPages() {
  if (!state.pdfDoc) {
    return;
  }
  const version = ++state.renderVersion;
  for (const pageNum of state.pageOrder) {
    if (version !== state.renderVersion) {
      return;
    }
    await renderPage(pageNum);
  }
}

async function renderThumbnail(pageNum, thumbCanvas) {
  const page = await getPdfPage(pageNum);
  const viewport = page.getViewport({ scale: 1, rotation: getRotation(pageNum) });
  const targetWidth = 170;
  const thumbScale = targetWidth / viewport.width;
  const scaledViewport = page.getViewport({ scale: thumbScale, rotation: getRotation(pageNum) });
  const dpr = window.devicePixelRatio || 1;
  const renderScale = dpr * state.thumbRenderQuality;
  const context = thumbCanvas.getContext("2d", { alpha: false });
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";

  thumbCanvas.width = Math.floor(scaledViewport.width * renderScale);
  thumbCanvas.height = Math.floor(scaledViewport.height * renderScale);
  thumbCanvas.style.width = `${scaledViewport.width}px`;
  thumbCanvas.style.height = `${scaledViewport.height}px`;

  await page.render({
    canvasContext: context,
    viewport: scaledViewport,
    transform: renderScale === 1 ? null : [renderScale, 0, 0, renderScale, 0, 0]
  }).promise;
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

  if (state.isFullScreen && state.fullScreenViewMode === "single") {
    els.viewerPanel.scrollTop = 0;
    return;
  }

  const view = state.pageViews.get(pageNum);
  if (!view) {
    return;
  }
  view.wrap.scrollIntoView({
    behavior: smooth ? "smooth" : "auto",
    block: "center",
    inline: "nearest"
  });
}

function updateCurrentPageByScroll() {
  if (!state.pdfDoc) {
    return;
  }
  if (state.isFullScreen && state.fullScreenViewMode === "single") {
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

function itemRectInViewport(item, viewport) {
  const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
  const fontHeight = Math.max(8, Math.hypot(tx[2], tx[3]));
  let width = item.width * tx[0];
  if (!Number.isFinite(width) || Math.abs(width) < 4) {
    width = Math.max(8, item.str.length * fontHeight * 0.45);
  }
  const left = width < 0 ? tx[4] + width : tx[4];
  const top = tx[5] - fontHeight;
  return {
    left,
    top,
    width: Math.max(4, Math.abs(width)),
    height: fontHeight
  };
}

function drawSearchHighlightsForPage(pageNum) {
  const view = state.pageViews.get(pageNum);
  if (!view || !view.viewport) {
    return;
  }

  view.searchOverlay.innerHTML = "";
  const matchIndexes = state.perPageMatchIndexes.get(pageNum) || [];
  const items = state.textItemsCache.get(pageNum) || [];

  for (const matchIndex of matchIndexes) {
    const match = state.searchMatches[matchIndex];
    const uniqueItemIndexes = new Set(match.itemIndexes || []);
    for (const itemIndex of uniqueItemIndexes) {
      const item = items[itemIndex];
      if (!item) {
        continue;
      }
      const rect = itemRectInViewport(item, view.viewport);
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
      const text = window.prompt("메모 내용을 입력하세요.", "");
      if (!text || !text.trim()) {
        return;
      }
      const bucket = getAnnotationBucket(pageNum);
      bucket.texts.push({ x: startPoint.x, y: startPoint.y, text: text.trim() });
      state.saveDirty = true;
      drawAnnotationsForPage(pageNum);
      setStatus("텍스트 메모를 추가했습니다.");
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
  const collapsedParts = [];
  let cursor = 0;

  for (const item of textContent.items) {
    const displayText = String(item.str || "").replace(/\s+/g, " ").trim();
    if (!displayText) {
      continue;
    }

    const lowered = displayText.toLowerCase();
    const searchable = lowered.replace(/\s+/g, "");
    if (!searchable) {
      continue;
    }

    const searchStart = cursor;
    const searchEnd = searchStart + searchable.length;
    cursor = searchEnd;

    items.push({
      index: items.length,
      str: displayText,
      lower: lowered,
      searchable,
      searchStart,
      searchEnd,
      width: item.width || 0,
      transform: item.transform
    });
    collapsedParts.push(searchable);
  }

  state.searchPageCache.set(pageNum, collapsedParts.join(""));
  state.textItemsCache.set(pageNum, items);
  return items;
}

function buildSearchPreview(items, hitItemIndexes) {
  if (!items.length || !hitItemIndexes.length) {
    return "";
  }
  const startIndex = Math.max(0, hitItemIndexes[0] - 2);
  const endIndex = Math.min(items.length - 1, hitItemIndexes[hitItemIndexes.length - 1] + 2);
  const raw = items
    .slice(startIndex, endIndex + 1)
    .map((item) => item.str)
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
    li.textContent = `${displayIndex}페이지: ${match.text}`;
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
  renderSearchResultList();
  for (const pageNum of state.pageOrder) {
    drawSearchHighlightsForPage(pageNum);
  }
  setStatus(
    `검색 결과 ${index + 1}/${state.searchMatches.length} - ${state.pageOrder.indexOf(match.pageNum) + 1}페이지`
  );
}

async function performSearch(rawQuery, jumpFirst = true) {
  const query = normalizeSearchText(rawQuery);
  const queryNeedle = query.replace(/\s+/g, "");
  state.searchQuery = query;
  state.searchMatches = [];
  state.perPageMatchIndexes = new Map();
  state.activeSearchIndex = -1;

  if (!queryNeedle) {
    renderSearchResultList();
    updateSearchCountText();
    updateThumbnailSearchMark();
    for (const pageNum of state.pageOrder) {
      drawSearchHighlightsForPage(pageNum);
    }
    setStatus("검색어를 입력해 주세요.");
    return;
  }

  setStatus("검색 중...");
  for (const pageNum of state.pageOrder) {
    const items = await ensureTextItems(pageNum);
    const pageSearchText = state.searchPageCache.get(pageNum) || "";
    if (!pageSearchText) {
      continue;
    }
    let from = 0;
    while (from < pageSearchText.length) {
      const found = pageSearchText.indexOf(queryNeedle, from);
      if (found < 0) {
        break;
      }
      const foundEnd = found + queryNeedle.length;
      const hitItemIndexes = [];
      for (const item of items) {
        if (item.searchEnd <= found) {
          continue;
        }
        if (item.searchStart >= foundEnd) {
          break;
        }
        hitItemIndexes.push(item.index);
      }
      if (!hitItemIndexes.length) {
        from = found + 1;
        continue;
      }
      const preview = buildSearchPreview(items, hitItemIndexes);
      const matchIndex = state.searchMatches.length;
      state.searchMatches.push({
        pageNum,
        itemIndexes: hitItemIndexes,
        text: preview || "(검색 결과)"
      });
      if (!state.perPageMatchIndexes.has(pageNum)) {
        state.perPageMatchIndexes.set(pageNum, []);
      }
      state.perPageMatchIndexes.get(pageNum).push(matchIndex);
      from = found + Math.max(1, queryNeedle.length);
    }
  }

  updateSearchCountText();
  renderSearchResultList();
  updateThumbnailSearchMark();
  for (const pageNum of state.pageOrder) {
    drawSearchHighlightsForPage(pageNum);
  }

  if (state.searchMatches.length === 0) {
    setStatus(`"${query}" 검색 결과가 없습니다.`, true);
    return;
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

async function zoomTo(newScale, anchorEvent = null) {
  if (!state.pdfDoc) {
    return;
  }
  const nextScale = clamp(newScale, 0.25, 4);
  if (Math.abs(nextScale - state.scale) < 0.001) {
    return;
  }

  let anchor = null;
  if (anchorEvent) {
    const rect = els.viewerPanel.getBoundingClientRect();
    anchor = {
      x: anchorEvent.clientX - rect.left + els.viewerPanel.scrollLeft,
      y: anchorEvent.clientY - rect.top + els.viewerPanel.scrollTop,
      dx: anchorEvent.clientX - rect.left,
      dy: anchorEvent.clientY - rect.top
    };
  }

  const oldScale = state.scale;
  state.scale = nextScale;
  await renderAllPages();
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
  localStorage.setItem("lookup-fullscreen-view-mode", state.fullScreenViewMode);
  updateFullscreenButtons();
  applyPageVisibility();
  goToPage(state.currentPage, false).catch(() => {});
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

async function loadPdfFromBytes(rawBytes, filePath) {
  const bytes = toUint8Array(rawBytes);
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  const pdfDoc = await loadingTask.promise;

  state.pdfDoc = pdfDoc;
  state.sourceBytes = bytes;
  state.filePath = filePath || "";
  state.pageOrder = Array.from({ length: pdfDoc.numPages }, (_v, i) => i + 1);
  state.pageCache.clear();
  state.pageRotations.clear();
  state.annotations.clear();
  state.textItemsCache.clear();
  state.searchPageCache.clear();
  clearSearchState();
  state.scale = 1;
  state.currentPage = 1;
  state.saveDirty = false;

  els.emptyHint.classList.add("hidden");
  await rebuildPageViews();
  await renderThumbnails();
  updatePageBadges();
  await goToPage(state.currentPage, false);
  setStatus(`열림: ${fileNameFromPath(filePath)} (${state.pageOrder.length}페이지)`);
}

async function loadPdfFromPath(filePath) {
  if (!filePath || typeof filePath !== "string") {
    return;
  }
  if (!filePath.toLowerCase().endsWith(".pdf")) {
    setStatus("PDF 파일만 열 수 있습니다.", true);
    return;
  }

  try {
    setStatus("PDF를 열고 있습니다...");
    const raw = await window.lookupAPI.readPdfFile(filePath);
    await loadPdfFromBytes(raw, filePath);
  } catch (error) {
    setStatus(`파일 열기 실패: ${error?.message || "알 수 없는 오류"}`, true);
  }
}

async function openFileDialog() {
  const selectedPath = await window.lookupAPI.openPdfDialog();
  if (!selectedPath) {
    return;
  }
  await loadPdfFromPath(selectedPath);
}
function bindToolbarActions() {
  els.openFileBtn.addEventListener("click", () => openFileDialog());
  els.saveAsBtn.addEventListener("click", () => savePdfAs().catch((error) => setStatus(error.message, true)));
  els.saveOverwriteBtn.addEventListener("click", () =>
    savePdfOverwrite().catch((error) => setStatus(error.message, true))
  );
  els.printBtn.addEventListener("click", async () => {
    const success = await window.lookupAPI.printDocument();
    if (!success) {
      setStatus("인쇄를 시작하지 못했습니다.", true);
    }
  });

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

  els.zoomInBtn.addEventListener("click", () => zoomTo(state.scale + 0.1).catch(() => {}));
  els.zoomOutBtn.addEventListener("click", () => zoomTo(state.scale - 0.1).catch(() => {}));
  els.zoomResetBtn.addEventListener("click", () => zoomTo(1).catch(() => {}));

  els.rotateLeftBtn.addEventListener("click", () => rotateCurrentPage(-90).catch(() => {}));
  els.rotateRightBtn.addEventListener("click", () => rotateCurrentPage(90).catch(() => {}));
  els.deletePageBtn.addEventListener("click", () => deleteCurrentPage().catch(() => {}));

  els.editModeButtons.forEach((button) => {
    button.addEventListener("click", () => setEditingMode(button.dataset.mode));
  });

  els.searchBtn.addEventListener("click", () => performSearch(els.searchInput.value, true).catch(() => {}));
  els.searchPrevBtn.addEventListener("click", () => searchNext(-1).catch(() => {}));
  els.searchNextBtn.addEventListener("click", () => searchNext(1).catch(() => {}));
  els.searchInput.addEventListener("keydown", async (event) => {
    if (event.key === "Enter") {
      await performSearch(els.searchInput.value, true);
    }
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
    state.fullscreenThumbVisible = !state.fullscreenThumbVisible;
    applyThumbPanelVisibility();
  });

  els.checkUpdateBtn.addEventListener("click", async () => {
    setStatus("업데이트를 확인하고 있습니다...");
    const result = await window.lookupAPI.checkForUpdates();
    if (!result.ok) {
      setStatus(`업데이트 확인 실패: ${result.message}`, true);
    }
  });

  els.installUpdateBtn.addEventListener("click", async () => {
    const result = await window.lookupAPI.installUpdateNow();
    if (!result.ok) {
      setStatus(result.message, true);
    }
  });
}

function bindWindowActions() {
  els.viewerPanel.addEventListener("scroll", queueScrollSync);

  window.addEventListener(
    "wheel",
    async (event) => {
      if (!event.ctrlKey || !state.pdfDoc) {
        return;
      }
      if (!(event.target instanceof Node) || !els.viewerPanel.contains(event.target)) {
        return;
      }
      event.preventDefault();
      if (state.wheelZoomBusy) {
        return;
      }
      state.wheelZoomBusy = true;
      try {
        const zoomDelta = event.deltaY < 0 ? 0.1 : -0.1;
        await zoomTo(state.scale + zoomDelta, event);
      } finally {
        state.wheelZoomBusy = false;
      }
    },
    { passive: false }
  );

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
    const file = event.dataTransfer?.files?.[0];
    if (!file?.path) {
      return;
    }
    await loadPdfFromPath(file.path);
  });

  window.addEventListener("keydown", async (event) => {
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
      await window.lookupAPI.printDocument();
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
    if (event.key === "PageUp") {
      event.preventDefault();
      const index = getCurrentDisplayIndex();
      if (index > 0) {
        await goToPage(state.pageOrder[index - 1], true);
      }
      return;
    }
    if (event.key === "PageDown") {
      event.preventDefault();
      const index = getCurrentDisplayIndex();
      if (index >= 0 && index < state.pageOrder.length - 1) {
        await goToPage(state.pageOrder[index + 1], true);
      }
      return;
    }
    if (event.key === "Delete" && state.pdfDoc && !event.ctrlKey && !event.metaKey) {
      const active = document.activeElement;
      const isTyping =
        active &&
        (active.tagName === "INPUT" || active.tagName === "TEXTAREA" || active.isContentEditable === true);
      if (!isTyping) {
        await deleteCurrentPage();
      }
    }
  });
}

function bindMainProcessEvents() {
  window.lookupAPI.onSystemOpenFile((filePath) => {
    loadPdfFromPath(filePath).catch((error) => {
      setStatus(error.message, true);
    });
  });

  window.lookupAPI.onMenuAction((action) => {
    const map = {
      "open-file": () => openFileDialog(),
      "save-as": () => savePdfAs(),
      "save-overwrite": () => savePdfOverwrite(),
      print: () => window.lookupAPI.printDocument(),
      "prev-page": async () => {
        const index = getCurrentDisplayIndex();
        if (index > 0) {
          await goToPage(state.pageOrder[index - 1], true);
        }
      },
      "next-page": async () => {
        const index = getCurrentDisplayIndex();
        if (index >= 0 && index < state.pageOrder.length - 1) {
          await goToPage(state.pageOrder[index + 1], true);
        }
      },
      "zoom-in": () => zoomTo(state.scale + 0.1),
      "zoom-out": () => zoomTo(state.scale - 0.1),
      "zoom-reset": () => zoomTo(1),
      "toggle-dark": () => setDarkMode(!document.body.classList.contains("dark")),
      "toggle-fullscreen-view-mode": () => toggleFullscreenViewMode(),
      "toggle-thumb-panel": () => {
        if (state.isFullScreen) {
          state.fullscreenThumbVisible = !state.fullscreenThumbVisible;
        } else {
          state.thumbPanelVisible = !state.thumbPanelVisible;
        }
        applyThumbPanelVisibility();
      },
      "check-update": () => window.lookupAPI.checkForUpdates(),
      "install-update": () => window.lookupAPI.installUpdateNow()
    };
    const fn = map[action];
    if (fn) {
      Promise.resolve(fn()).catch((error) => {
        setStatus(error?.message || "명령 실행 중 오류", true);
      });
    }
  });

  window.lookupAPI.onFullScreenChanged((isFullScreen) => {
    state.isFullScreen = isFullScreen;
    document.body.classList.toggle("fullscreen", isFullScreen);
    document.body.classList.toggle("hide-search-panel", isFullScreen);

    if (isFullScreen) {
      state.fullscreenThumbVisible = false;
      els.fullscreenMiniBar.classList.remove("hidden");
    } else {
      els.fullscreenMiniBar.classList.add("hidden");
    }

    applyThumbPanelVisibility();
    applyPageVisibility();
    updateToolbarState();
    goToPage(state.currentPage, false).catch(() => {});
  });

  window.lookupAPI.onUpdateStatus((payload) => {
    if (!payload?.status) {
      return;
    }
    if (payload.status === "downloaded") {
      els.installUpdateBtn.classList.remove("hidden");
    } else if (payload.status === "checking" || payload.status === "downloading" || payload.status === "available") {
      els.installUpdateBtn.classList.add("hidden");
    }
    if (payload.message) {
      setStatus(payload.message, payload.status === "error");
    }
  });
}

async function initializeUpdateStatus() {
  const config = await window.lookupAPI.getUpdateConfig();
  if (!config.enabled) {
    setStatus("업데이트 비활성: update-config.json에 GitHub owner/repo를 입력하세요.");
  } else {
    setStatus(`업데이트 연동 준비됨: ${config.owner}/${config.repo}`);
  }
}

async function init() {
  applySavedDarkMode();
  bindToolbarActions();
  bindWindowActions();
  bindMainProcessEvents();
  setEditingMode("view");
  applyThumbPanelVisibility();
  updateToolbarState();

  const isFullScreen = await window.lookupAPI.isFullScreen();
  state.isFullScreen = isFullScreen;
  document.body.classList.toggle("fullscreen", isFullScreen);
  document.body.classList.toggle("hide-search-panel", isFullScreen);
  if (isFullScreen) {
    state.fullscreenThumbVisible = false;
    els.fullscreenMiniBar.classList.remove("hidden");
  } else {
    els.fullscreenMiniBar.classList.add("hidden");
  }
  applyThumbPanelVisibility();
  updateFullscreenButtons();
  await initializeUpdateStatus();
}

init().catch((error) => {
  setStatus(error?.message || "초기화 오류", true);
});
