"use strict";

// All page changes happen in the disposable conversion window. Source files are never rewritten.
const WEB_PRINT_ENGINE_VERSION = "fullpage-v3";
const DEFAULT_TIMEOUT_MS = 24000;

function createPageScript(operation, options = {}) {
  return `(${pageRuntime.toString()})(${JSON.stringify(operation)}, ${JSON.stringify(options)})`;
}

// This function is serialized into a sandboxed renderer; it must have no Node dependencies.
async function pageRuntime(operation, options) {
  const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const now = () => Date.now();
  const deadline = now() + Math.max(500, Math.min(30000, Number(options.timeoutMs) || 24000));
  const elements = () => Array.from(document.querySelectorAll("*")).slice(0, 50000);
  const visible = (el) => {
    const style = getComputedStyle(el);
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity) !== 0
      && el.getBoundingClientRect().width > 0 && el.getBoundingClientRect().height > 0;
  };
  const describe = (el) => (String(el.tagName || "").toLowerCase()
    + (el.id ? `#${String(el.id).slice(0, 60)}` : "")).slice(0, 100);
  const settlePaint = () => Promise.race([
    new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))), wait(150)
  ]);
  const metrics = () => {
    const root = document.scrollingElement || document.documentElement;
    let scrollContainerCount = 0;
    let maxOverflowY = 0;
    let maxOverflowX = 0;
    const scrollContainers = [];
    for (const el of elements()) {
      if (!visible(el)) continue;
      const style = getComputedStyle(el);
      const clippedY = Math.max(0, el.scrollHeight - el.clientHeight);
      const clippedX = Math.max(0, el.scrollWidth - el.clientWidth);
      const isRoot = el === root || el === document.body || el === document.documentElement;
      // Small crop boxes and horizontal carousels are intentional, not lost document content.
      const relevantY = !isRoot && el.clientWidth >= 240 && el.clientHeight >= 120
        && /auto|scroll|hidden|clip/.test(style.overflowY) && clippedY > 120;
      const relevantX = !isRoot && el.clientHeight >= 240 && /auto|scroll/.test(style.overflowX) && clippedX > 120;
      if (!relevantY && !relevantX) continue;
      scrollContainerCount += 1;
      maxOverflowY = Math.max(maxOverflowY, relevantY ? clippedY : 0);
      maxOverflowX = Math.max(maxOverflowX, relevantX ? clippedX : 0);
      if (scrollContainers.length < 24) scrollContainers.push({
        name: describe(el), clientHeight: el.clientHeight, scrollHeight: el.scrollHeight,
        clientWidth: el.clientWidth, scrollWidth: el.scrollWidth, clippedY, clippedX
      });
    }
    const images = Array.from(document.images).filter(visible);
    const loadedImages = images.filter((img) => img.complete && img.naturalWidth > 1 && img.naturalHeight > 1);
    // A missing image can have zero intrinsic height, so it must not disappear
    // from the resource report merely because visible() cannot measure it yet.
    const imageResources = Array.from(document.images).filter((img) => img.getClientRects().length > 0
      && getComputedStyle(img).visibility !== "hidden");
    const pendingImages = imageResources.filter((img) => !img.complete);
    const failedImages = imageResources.filter((img) => img.complete && (!img.naturalWidth || !img.naturalHeight));
    const imageDetails = (list) => list.slice(0, 16).map((img) => ({
      name: describe(img), source: String(img.currentSrc || img.src || "").split(/[?#]/)[0].slice(0, 180)
    }));
    return {
      sourceType: options.sourceType || "web-html",
      totalHeight: Math.max(root?.scrollHeight || 0, document.documentElement?.scrollHeight || 0, document.body?.scrollHeight || 0),
      totalWidth: Math.max(root?.scrollWidth || 0, document.documentElement?.scrollWidth || 0, document.body?.scrollWidth || 0),
      viewportHeight: window.innerHeight, viewportWidth: window.innerWidth,
      textLength: String(document.body?.innerText || "").trim().length,
      imageCount: images.length, loadedImageCount: loadedImages.length,
      totalImageCount: document.images.length,
      pendingImageCount: pendingImages.length, failedImageCount: failedImages.length,
      pendingImages: imageDetails(pendingImages), failedImages: imageDetails(failedImages),
      canvasCount: Array.from(document.querySelectorAll("canvas,svg")).filter(visible).length,
      backgroundImageCount: elements().filter((el) => visible(el) && getComputedStyle(el).backgroundImage !== "none").length,
      scrollContainerCount, maxOverflowY, maxOverflowX, scrollContainers,
      iframeCount: document.querySelectorAll("iframe,frame").length
    };
  };
  const ready = async (budgetMs = 6000) => {
    const end = Math.min(deadline, now() + budgetMs);
    if (document.fonts?.ready) await Promise.race([document.fonts.ready.catch(() => {}), wait(Math.max(0, end - now()))]);
    while (now() < end) {
      const pending = Array.from(document.images).filter((img) => !img.complete && img.getAttribute("loading") !== "lazy");
      if (!pending.length) break;
      await wait(100);
    }
    // decode() can hang on broken images, so it shares the same deadline.
    const decodes = Array.from(document.images).filter((img) => img.complete && img.naturalWidth > 1 && img.decode)
      .map((img) => img.decode().catch(() => {}));
    await Promise.race([Promise.all(decodes), wait(Math.max(0, end - now()))]);
    await settlePaint();
  };
  if (operation === "inspect") return metrics();
  if (operation === "ready") { await ready(); return metrics(); }
  if (operation === "access") {
    const title = String(document.title || "").trim();
    const challengeTitle = /^(?:잠시만\s*기다리십시오|잠시만\s*기다려\s*주세요|just a moment|attention required|checking your browser|verify you are human|access denied|접속이 차단)/i.test(title);
    const visibleChallenge = Array.from(document.querySelectorAll("#challenge-running,#challenge-stage,#cf-challenge-running,form[action*='__cf_chl'],.cf-challenge"))
      .some(visible);
    return { title, challenged: challengeTitle || visibleChallenge, readyState: document.readyState };
  }
  if (operation === "expand") {
    let expanded = 0;
    if (/(^|\.)gmarket\.co\.kr$/i.test(location.hostname)) {
      for (const button of document.querySelectorAll("button.button__detail-more.js-toggle-button")) {
        if (visible(button) && /상세정보\s*더보기/.test(button.textContent || "")) { button.click(); expanded += 1; }
      }
    }
    if (expanded) { await wait(350); await settlePaint(); }
    return { expanded };
  }

  const before = metrics();
  let promotedImages = 0;
  for (const img of Array.from(document.images)) {
    img.loading = "eager";
    img.decoding = "sync";
    const current = String(img.getAttribute("src") || "");
    const placeholder = !current || /(?:placeholder|transparent|spacer|blank)(?:[._/-]|$)/i.test(current)
      || (/^data:image\//i.test(current) && img.naturalWidth <= 16 && img.hasAttribute("data-src"))
      || (img.complete && img.naturalWidth <= 16 && img.naturalHeight <= 16);
    if (!placeholder) continue;
    const lazySrc = ["data-src", "data-original", "data-lazy-src", "data-url", "data-echo"]
      .map((name) => img.getAttribute(name)).find((value) => {
        if (!value) return false;
        try { return /^(?:https?:|file:|data:)$/.test(new URL(value, document.baseURI).protocol); } catch (_) { return false; }
      });
    const lazySrcset = img.getAttribute("data-srcset") || img.getAttribute("data-lazy-srcset");
    if (lazySrcset) img.setAttribute("srcset", lazySrcset);
    if (lazySrc) { img.src = lazySrc; promotedImages += 1; }
  }
  for (const el of elements()) {
    if (getComputedStyle(el).contentVisibility === "auto") el.style.setProperty("content-visibility", "visible", "important");
  }
  await ready(2500);

  // Re-read height at every step: lazy sections often extend the document while scrolling.
  const root = document.scrollingElement || document.documentElement;
  const scrollBoxes = [root, ...elements().filter((el) => {
    if (el === root || el === document.body || el === document.documentElement || !visible(el)) return false;
    return el.clientHeight >= 150 && el.clientWidth >= 240
      && /auto|scroll/.test(getComputedStyle(el).overflowY) && el.scrollHeight > el.clientHeight + 120;
  })].filter(Boolean).slice(0, 12);
  let scrollSteps = 0;
  let reachedBottom = true;
  const finishingReserveMs = Math.max(1000, Math.min(3500, Math.floor((deadline - now()) * 0.3)));
  for (const el of scrollBoxes) {
    const original = el.scrollTop;
    const isRoot = el === root;
    let y = 0;
    let lastBottom = -1;
    let stableBottom = 0;
    const step = Math.max(250, Math.floor((el.clientHeight || window.innerHeight || 900) * 0.72));
    while (now() < deadline - finishingReserveMs && scrollSteps < 650) {
      const bottom = Math.max(0, el.scrollHeight - (el.clientHeight || window.innerHeight));
      y = Math.min(y, bottom);
      if (isRoot) window.scrollTo(0, y); else el.scrollTop = y;
      scrollSteps += 1;
      await wait(65);
      if (y >= bottom) {
        stableBottom = bottom === lastBottom ? stableBottom + 1 : 0;
        lastBottom = bottom;
        if (stableBottom >= 2) break;
        await wait(140);
      } else y = Math.min(bottom, y + step);
    }
    if (y < Math.max(0, el.scrollHeight - el.clientHeight) - 2 || stableBottom < 2) reachedBottom = false;
    if (isRoot) window.scrollTo(0, original); else el.scrollTop = original;
  }
  await ready(Math.min(2500, finishingReserveMs));

  const style = document.createElement("style");
  style.id = "lookup-web-fullpage-print-style";
  style.textContent = `
    @page { size: A4; margin: 10mm; }
    html, body { height: auto !important; max-height: none !important; overflow: visible !important; }
    html { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; scroll-behavior: auto !important; }
    body { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
    img { max-width: 100%; }
    [data-lookup-image-slice] { display: block !important; position: relative !important; overflow: hidden !important; break-inside: avoid !important; page-break-inside: avoid !important; }
    [data-lookup-image-slice] > canvas { display: block !important; width: 100% !important; height: auto !important; max-height: none !important; margin: 0 !important; }
  `;
  document.head.appendChild(style);
  let normalized = 0;
  let hiddenFloating = 0;
  let staticFloating = 0;
  for (const el of elements()) {
    if (!visible(el) || /^(INPUT|TEXTAREA|SELECT|CANVAS|SVG|VIDEO|IMG|IFRAME)$/.test(el.tagName)) continue;
    const computed = getComputedStyle(el);
    if (options.sourceType === "notion-web" && /notion-(?:topbar|sidebar|help|overlay)|pseudoSelection|presence/i.test(String(el.className))) {
      el.style.setProperty("display", "none", "important"); hiddenFloating += 1; continue;
    }
    if (computed.position === "fixed" || computed.position === "sticky") {
      el.style.setProperty("position", "static", "important");
      el.style.setProperty("inset", "auto", "important"); staticFloating += 1;
    }
    const clippedY = el.scrollHeight - el.clientHeight;
    const scrollBox = el.clientWidth >= 240 && el.clientHeight >= 120 && clippedY > 120
      && /auto|scroll|hidden|clip/.test(computed.overflowY);
    if (scrollBox) {
      el.style.setProperty("overflow-y", "visible", "important");
      el.style.setProperty("overflow-x", "visible", "important");
      el.style.setProperty("height", "auto", "important");
      el.style.setProperty("max-height", "none", "important");
      el.style.setProperty("contain", "none", "important");
      normalized += 1;
    }
    if (Number(el.getBoundingClientRect().height) > 950 && /avoid/.test(computed.breakInside || computed.pageBreakInside)) {
      el.style.setProperty("break-inside", "auto", "important");
      el.style.setProperty("page-break-inside", "auto", "important");
    }
  }
  // Chromium treats one enormous product-detail image as an indivisible item. Split it
  // into printable strips with the original decoded image, preserving text-sized detail.
  // Each printed item contains only its own pixels, without relying on transformed
  // image overflow or asking Chromium to fragment one enormous replaced element.
  let splitImages = 0;
  let imageSlices = 0;
  let stripPixelCount = 0;
  for (const img of Array.from(document.images)) {
    if (!visible(img) || !img.complete || img.naturalWidth <= 0 || img.closest("[data-lookup-image-slice]")) continue;
    const width = Math.min(718, img.getBoundingClientRect().width);
    const scaledHeight = img.naturalHeight * width / img.naturalWidth;
    if (width < 180 || scaledHeight < 1200 || getComputedStyle(img).objectFit === "cover") continue;
    const naturalStep = Math.max(1, Math.floor(880 * img.naturalWidth / width));
    const stripCount = Math.ceil(img.naturalHeight / naturalStep);
    if (stripCount > 200 || imageSlices + stripCount > 300) throw new Error("웹페이지 이미지가 너무 길어 전체 PDF 저장을 완료하지 못했습니다.");
    const canvasWidth = Math.min(1400, img.naturalWidth);
    const requiredPixels = Math.ceil(canvasWidth * img.naturalHeight * canvasWidth / img.naturalWidth);
    if (stripPixelCount + requiredPixels > 40000000) throw new Error("웹페이지 이미지가 너무 커서 전체 PDF 저장을 완료하지 못했습니다.");
    const group = document.createElement("div");
    group.setAttribute("data-lookup-long-image", "");
    group.style.cssText = `width:${width}px;max-width:100%;margin:0 auto;`;
    for (let top = 0; top < img.naturalHeight; top += naturalStep) {
      const slice = document.createElement("div");
      const height = Math.min(naturalStep, img.naturalHeight - top);
      slice.setAttribute("data-lookup-image-slice", "");
      slice.style.aspectRatio = `${img.naturalWidth} / ${height}`;
      const canvas = document.createElement("canvas");
      canvas.width = canvasWidth;
      canvas.height = Math.max(1, Math.round(height * canvas.width / img.naturalWidth));
      canvas.setAttribute("role", "img");
      if (top === 0 && img.alt) canvas.setAttribute("aria-label", img.alt);
      const context = canvas.getContext("2d");
      if (!context) throw new Error("상품 이미지를 나누지 못했습니다.");
      context.drawImage(img, 0, top, img.naturalWidth, height, 0, 0, canvas.width, canvas.height);
      slice.appendChild(canvas); group.appendChild(slice); imageSlices += 1;
    }
    stripPixelCount += requiredPixels;
    img.replaceWith(group); splitImages += 1;
  }
  let trimmedTrailingBlanks = 0;
  for (const container of document.querySelectorAll("body,main,article,.ee-contents,.notion-page-content")) {
    for (let count = 0; count < 12; count += 1) {
      const tail = Array.from(container.children).reverse().find((el) => !/^(SCRIPT|STYLE|LINK|NOSCRIPT)$/.test(el.tagName) && getComputedStyle(el).display !== "none");
      if (!tail || tail.tagName !== "P" || String(tail.textContent || "").trim()
        || tail.querySelector("img,svg,canvas,iframe,video,audio,object,input,button,hr")) break;
      const computed = getComputedStyle(tail);
      if (tail.getBoundingClientRect().height > 80 || computed.backgroundImage !== "none"
        || computed.breakBefore !== "auto" || computed.breakAfter !== "auto"
        || ["::before", "::after"].some((pseudo) => !["none", "normal", '""'].includes(getComputedStyle(tail, pseudo).content))) break;
      tail.remove(); trimmedTrailingBlanks += 1;
    }
  }
  window.scrollTo(0, 0);
  await ready(2500);
  return {
    sourceType: options.sourceType || "web-html", before, after: metrics(),
    lazyLoad: { touched: scrollSteps, scrollSteps, promotedImages, reachedBottom },
    normalization: { normalized, hiddenFloating, staticFloating, splitImages, imageSlices, trimmedTrailingBlanks },
    rasterFallbackUsed: false
  };
}

async function inspectWebPage(webContents, options = {}) {
  return webContents.executeJavaScript(createPageScript("inspect", options));
}

async function waitForWebPageReadiness(webContents, options = {}) {
  return webContents.executeJavaScript(createPageScript("ready", options));
}

async function waitForWebAccess(webContents, { timeoutMs = 20000 } = {}) {
  const end = Date.now() + Math.max(500, Math.min(30000, timeoutMs));
  let access;
  do {
    access = await webContents.executeJavaScript(createPageScript("access"));
    if (!access.challenged && access.readyState === "complete") return access;
    await new Promise((resolve) => setTimeout(resolve, 300));
  } while (Date.now() < end);
  const error = new Error(access?.challenged ? "사이트 접속 확인이 끝나지 않았습니다. 잠시 후 다시 시도해 주세요." : "웹페이지 로딩이 끝나지 않았습니다. 다시 시도해 주세요.");
  error.code = access?.challenged ? "WEB_ACCESS_CHALLENGE" : "WEB_LOAD_TIMEOUT";
  throw error;
}

async function prepareWebPageForPdf(webContents, options = {}) {
  const startedAt = Date.now();
  const timeoutMs = Math.max(5000, Math.min(45000, Number(options.timeoutMs) || DEFAULT_TIMEOUT_MS));
  const expansion = await webContents.executeJavaScript(createPageScript("expand", options));
  const frames = Array.from(webContents.mainFrame?.framesInSubtree || []).filter((frame) => frame !== webContents.mainFrame).slice(0, 12).reverse();
  const frameResults = [];
  // Prepare child documents before expanding their embedding boxes. Browser security is
  // kept enabled; Electron evaluates in each frame without granting page code Node access.
  let unpreparedFrameCount = 0;
  for (const frame of frames) {
    if (Date.now() - startedAt > timeoutMs - 7500) { unpreparedFrameCount += 1; continue; }
    try {
      const preliminary = await frame.executeJavaScript(createPageScript("inspect", { sourceType: "web-frame" }));
      if (preliminary.viewportWidth < 240 || preliminary.viewportHeight < 120 || (!preliminary.textLength && !preliminary.totalImageCount && !preliminary.canvasCount && !preliminary.backgroundImageCount)) continue;
      const result = await frame.executeJavaScript(createPageScript("prepare", {
        sourceType: "web-frame", timeoutMs: Math.min(14000, timeoutMs - (Date.now() - startedAt) - 7500)
      }));
      const parent = frame.parent;
      const index = parent?.frames.indexOf(frame);
      if (parent && index >= 0) await parent.executeJavaScript(`(() => {
        const frame = document.querySelectorAll("iframe,frame")[${index}];
        if (!frame || frame.getBoundingClientRect().width < 240) return;
        frame.style.setProperty("height", ${JSON.stringify(`${Math.min(180000, Math.ceil(result.after.totalHeight))}px`)}, "important");
        frame.style.setProperty("max-height", "none", "important");
        frame.setAttribute("scrolling", "no");
      })()`);
      frameResults.push({ before: result.before, after: result.after, normalization: result.normalization, lazyLoad: result.lazyLoad });
    } catch (error) { unpreparedFrameCount += 1; frameResults.push({ error: error?.message || String(error) }); }
  }
  const result = await webContents.executeJavaScript(createPageScript("prepare", {
    ...options, timeoutMs: Math.max(5000, timeoutMs - (Date.now() - startedAt))
  }));
  result.frames = frameResults;
  result.unpreparedFrameCount = unpreparedFrameCount;
  result.expansion = expansion;
  result.title = webContents.getTitle();
  result.elapsedMs = Date.now() - startedAt;
  result.engineVersion = WEB_PRINT_ENGINE_VERSION;
  return result;
}

async function useScreenMediaForPdf(webContents) {
  const debuggerApi = webContents.debugger;
  if (!debuggerApi) throw new Error("화면 출력 연결을 준비하지 못했습니다.");
  const attachedHere = !debuggerApi.isAttached();
  if (attachedHere) debuggerApi.attach("1.3");
  try { await debuggerApi.sendCommand("Emulation.setEmulatedMedia", { media: "screen" }); }
  catch (error) { if (attachedHere && debuggerApi.isAttached()) debuggerApi.detach(); throw error; }
  return async () => {
    try { if (debuggerApi.isAttached()) await debuggerApi.sendCommand("Emulation.setEmulatedMedia", { media: "" }); }
    finally { if (attachedHere && debuggerApi.isAttached()) debuggerApi.detach(); }
  };
}

function assessWebPageContent(diagnostics) {
  const after = diagnostics?.after || diagnostics || {};
  const hasContent = Number(after.textLength || 0) > 0 || Number(after.loadedImageCount || 0) > 0
    || Number(after.canvasCount || 0) > 0 || Number(after.backgroundImageCount || 0) > 0
    || (diagnostics?.frames || []).some((frame) => Number(frame.after?.textLength || 0) > 0 || Number(frame.after?.loadedImageCount || 0) > 0);
  return {
    hasContent,
    incompleteImages: Number(after.pendingImageCount || 0) + Number(after.failedImageCount || 0)
      + (diagnostics?.frames || []).reduce((total, frame) => total + Number(frame.after?.pendingImageCount || 0) + Number(frame.after?.failedImageCount || 0), 0),
    reachedBottom: diagnostics?.lazyLoad?.reachedBottom !== false
      && (diagnostics?.frames || []).every((frame) => !frame.error && frame.lazyLoad?.reachedBottom !== false)
      && Number(diagnostics?.unpreparedFrameCount || 0) === 0
  };
}

function inspectCapturedImage(nativeImage) {
  if (!nativeImage || nativeImage.isEmpty()) return { valid: false, reason: "empty-image" };
  const size = nativeImage.getSize();
  if (size.width < 8 || size.height < 8) return { valid: false, reason: "invalid-image-size", ...size };
  const pixels = nativeImage.toBitmap();
  if (!pixels || pixels.length < size.width * size.height * 4) return { valid: false, reason: "missing-image-pixels", ...size };
  // Sample every few pixels, including alpha. Empty hidden-window captures are often
  // entirely transparent or a single white colour and must never count as a document.
  const stride = Math.max(4, Math.floor(pixels.length / 80000 / 4) * 4);
  let opaque = 0;
  let min = 255;
  let max = 0;
  for (let index = 0; index < pixels.length; index += stride) {
    if (pixels[index + 3] < 16) continue;
    opaque += 1;
    const tone = (pixels[index] + pixels[index + 1] + pixels[index + 2]) / 3;
    min = Math.min(min, tone); max = Math.max(max, tone);
  }
  return { valid: opaque > 0 && max - min >= 2, reason: !opaque ? "transparent-image" : max - min < 2 ? "uniform-image" : "", ...size, contrast: max - min };
}

module.exports = {
  WEB_PRINT_ENGINE_VERSION, createPageScript, inspectWebPage, waitForWebPageReadiness,
  waitForWebAccess, prepareWebPageForPdf, useScreenMediaForPdf, assessWebPageContent, inspectCapturedImage
};
