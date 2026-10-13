import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const helperSource = await readFile(new URL("../src/renderer/view/focus-scroll.js", import.meta.url), "utf8");
const {
  createPageWheelNavigator,
  getPageScrollBounds,
  isFocusPageClippedByViewport
} = await import(`data:text/javascript;base64,${Buffer.from(helperSource).toString("base64")}`);
const appSource = await readFile(new URL("../src/renderer/app-core.js", import.meta.url), "utf8");

const bounds = { minTop: 100, maxTop: 1100, minLeft: 20, maxLeft: 420 };
function consume(navigator, overrides = {}) {
  return navigator.consume({
    pageKey: 1, bounds, scrollTop: 500, scrollLeft: 120,
    deltaX: 0, deltaY: 120, deltaMode: 0, shiftKey: false,
    viewportHeight: 600, viewportWidth: 800, now: 1000, ...overrides
  });
}

test("page bounds include reading margins without exceeding the scrollable panel", () => {
  assert.deepEqual(getPageScrollBounds(
    { scrollHeight: 1600, clientHeight: 600, scrollWidth: 1500, clientWidth: 800 },
    { offsetTop: 300, offsetHeight: 1200, offsetLeft: 100, offsetWidth: 1400 }
  ), { minTop: 276, maxTop: 924, minLeft: 76, maxLeft: 700 });
});

test("centered small pages have zero bounds when the panel cannot scroll", () => {
  assert.deepEqual(getPageScrollBounds(
    { scrollHeight: 600, clientHeight: 600, scrollWidth: 800, clientWidth: 800 },
    { offsetTop: 150, offsetHeight: 300, offsetLeft: 200, offsetWidth: 400 }
  ), { minTop: 0, maxTop: 0, minLeft: 0, maxLeft: 0 });
});

test("bounds at the document end stay reachable", () => {
  assert.deepEqual(getPageScrollBounds(
    { scrollHeight: 1000, clientHeight: 600, scrollWidth: 1000, clientWidth: 800 },
    { offsetTop: 950, offsetHeight: 100, offsetLeft: 950, offsetWidth: 100 }
  ), { minTop: 400, maxTop: 400, minLeft: 200, maxLeft: 200 });
});

test("interior scrolling moves within the selected page", () => {
  const result = consume(createPageWheelNavigator(), { deltaX: 25 });
  assert.deepEqual(result, { scrollTop: 620, scrollLeft: 145, pageDirection: 0 });
});

test("a large downward input reaches the bottom before a later input turns the page", () => {
  const navigator = createPageWheelNavigator();
  assert.deepEqual(consume(navigator, { scrollTop: 1000, deltaY: 5000 }),
    { scrollTop: 1100, scrollLeft: 120, pageDirection: 0 });
  assert.equal(consume(navigator, { scrollTop: 1100, now: 1010 }).pageDirection, 1);
});

test("a large upward input reaches the top before a later input turns back", () => {
  const navigator = createPageWheelNavigator();
  assert.equal(consume(navigator, { scrollTop: 200, deltaY: -5000 }).pageDirection, 0);
  assert.equal(consume(navigator, { scrollTop: 100, deltaY: -120, now: 1010 }).pageDirection, -1);
});

test("subpixel browser scroll positions count as the same edge", () => {
  assert.equal(consume(createPageWheelNavigator(), { scrollTop: 1099.5 }).pageDirection, 1);
  assert.equal(consume(createPageWheelNavigator(), { scrollTop: 100.5, deltaY: -120 }).pageDirection, -1);
});

test("small trackpad edge inputs accumulate without losing motion", () => {
  const navigator = createPageWheelNavigator();
  for (let index = 0; index < 5; index += 1) {
    assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 6, now: 1000 + index * 10 }).pageDirection, 0);
  }
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 6, now: 1050 }).pageDirection, 1);
});

test("changing direction clears accumulated edge movement", () => {
  const navigator = createPageWheelNavigator();
  const flat = { ...bounds, minTop: 0, maxTop: 0 };
  assert.equal(consume(navigator, { bounds: flat, scrollTop: 0, deltaY: 30 }).pageDirection, 0);
  assert.equal(consume(navigator, { bounds: flat, scrollTop: 0, deltaY: -10, now: 1010 }).pageDirection, 0);
  assert.equal(consume(navigator, { bounds: flat, scrollTop: 0, deltaY: -26, now: 1020 }).pageDirection, -1);
});

test("a pause begins a fresh trackpad edge gesture", () => {
  const navigator = createPageWheelNavigator();
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 30 }).pageDirection, 0);
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 6, now: 1181 }).pageDirection, 0);
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 30, now: 1190 }).pageDirection, 1);
});

test("changing pages does not carry a partly accumulated edge gesture", () => {
  const navigator = createPageWheelNavigator();
  consume(navigator, { scrollTop: 1100, deltaY: 30 });
  assert.equal(consume(navigator, { pageKey: 2, scrollTop: 1100, deltaY: 6, now: 1010 }).pageDirection, 0);
});

test("another page turn is blocked for 280ms without blocking internal scrolling", () => {
  const navigator = createPageWheelNavigator();
  assert.equal(consume(navigator, { scrollTop: 1100 }).pageDirection, 1);
  const internal = consume(navigator, { pageKey: 2, scrollTop: 100, now: 1020 });
  assert.equal(internal.scrollTop, 220);
  assert.equal(internal.pageDirection, 0);
  assert.equal(consume(navigator, { pageKey: 2, scrollTop: 1100, now: 1279 }).pageDirection, 0);
  assert.equal(consume(navigator, { pageKey: 2, scrollTop: 1100, now: 1280 }).pageDirection, 1);
});

test("leaving an edge clears previous edge pressure", () => {
  const navigator = createPageWheelNavigator();
  consume(navigator, { scrollTop: 1100, deltaY: 30 });
  consume(navigator, { scrollTop: 1100, deltaY: -20, now: 1010 });
  consume(navigator, { scrollTop: 1080, deltaY: 20, now: 1020 });
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 6, now: 1030 }).pageDirection, 0);
});

test("Shift wheel pans horizontally without turning pages", () => {
  const result = consume(createPageWheelNavigator(), { scrollTop: 1100, shiftKey: true });
  assert.deepEqual(result, { scrollTop: 1100, scrollLeft: 240, pageDirection: 0 });
});

test("horizontal and equal-axis trackpad gestures never turn pages", () => {
  assert.equal(consume(createPageWheelNavigator(), { scrollTop: 1100, deltaX: 150 }).pageDirection, 0);
  assert.equal(consume(createPageWheelNavigator(), { scrollTop: 1100, deltaX: 120 }).pageDirection, 0);
});

test("line and page wheel units use the appropriate axis dimensions", () => {
  assert.deepEqual(consume(createPageWheelNavigator(), { deltaX: 2, deltaY: 3, deltaMode: 1 }),
    { scrollTop: 548, scrollLeft: 152, pageDirection: 0 });
  assert.deepEqual(consume(createPageWheelNavigator(), { deltaX: 0.25, deltaY: 0.5, deltaMode: 2 }),
    { scrollTop: 800, scrollLeft: 320, pageDirection: 0 });
});

test("horizontal-only clipping and pages with no vertical scroll still allow vertical page turns", () => {
  const panel = { scrollHeight: 600, clientHeight: 600, scrollWidth: 1500, clientWidth: 800 };
  const page = { offsetTop: 24, offsetHeight: 300, offsetLeft: 24, offsetWidth: 1452 };
  assert.equal(isFocusPageClippedByViewport(panel, page), true);
  const result = consume(createPageWheelNavigator(), { bounds: getPageScrollBounds(panel, page), scrollTop: 0 });
  assert.equal(result.pageDirection, 1);
  assert.equal(result.scrollTop, 0);
});

test("reset removes an earlier turn lock and partial gesture", () => {
  const navigator = createPageWheelNavigator();
  consume(navigator, { scrollTop: 1100 });
  navigator.reset();
  assert.equal(consume(navigator, { scrollTop: 1100, now: 1001 }).pageDirection, 1);
  navigator.reset();
  consume(navigator, { scrollTop: 1100, deltaY: 30 });
  navigator.reset();
  assert.equal(consume(navigator, { scrollTop: 1100, deltaY: 6, now: 1010 }).pageDirection, 0);
});

function sourceBetween(start, end) {
  const from = appSource.indexOf(start);
  const to = appSource.indexOf(end, from + start.length);
  assert.ok(from >= 0 && to > from, `Renderer function boundary: ${start}`);
  return appSource.slice(from, to);
}

// Execute the production wheel handler, page navigation and scroll selection in a
// small DOM geometry fixture. Actual rendering is checked separately in the app.
const rendererFunctions = [
  sourceBetween("function getCurrentDisplayIndex()", "async function stepPage("),
  sourceBetween("function normalizeGoToPageOptions(", "function alignCurrentPageToViewerCenter()"),
  sourceBetween("function updateCurrentPageByScroll()", "function queueScrollSync()"),
  sourceBetween("function rememberWheelScrollPosition()", "async function zoomTo(")
].join("\n");

function createRendererFixture({ mode = "focus", zoomMode = "manual", pageOrder = [1, 2, 3], pages, loadPdfPage } = {}) {
  const geometry = pages || pageOrder.map(() => ({ width: 1200, height: 1400 }));
  const state = { pdfDoc: {}, currentPage: pageOrder[0], pageOrder, zoomMode, scale: 1,
    wheelScrollPosition: null, wheelPageTurn: null, pageViews: new Map() };
  const panel = { clientWidth: 800, clientHeight: 600, scrollHeight: 1, scrollWidth: 1 };
  let top = 0;
  let left = 0;
  Object.defineProperties(panel, {
    scrollTop: { get: () => Math.min(top, Math.max(0, panel.scrollHeight - panel.clientHeight)), set: value => { top = Math.min(Math.max(0, value), Math.max(0, panel.scrollHeight - panel.clientHeight)); } },
    scrollLeft: { get: () => Math.min(left, Math.max(0, panel.scrollWidth - panel.clientWidth)), set: value => { left = Math.min(Math.max(0, value), Math.max(0, panel.scrollWidth - panel.clientWidth)); } }
  });
  panel.scrollTo = ({ top: nextTop, left: nextLeft }) => { panel.scrollTop = nextTop; panel.scrollLeft = nextLeft; };
  let nextTop = 24;
  pageOrder.forEach((page, index) => {
    const size = geometry[index];
    const classes = new Set();
    state.pageViews.set(page, { wrap: {
      offsetTop: mode === "focus" ? 24 : (size.top ?? nextTop), offsetLeft: size.left ?? 24,
      style: {},
      get offsetWidth() { return parseFloat(this.style.width) || size.width; },
      get offsetHeight() { return parseFloat(this.style.height) || size.height; },
      classList: { contains: name => classes.has(name), remove: name => classes.delete(name) }
    } });
    nextTop += size.height + 24;
  });
  const updateGeometry = () => {
    const visible = mode === "focus" ? [state.pageViews.get(state.currentPage)] : [...state.pageViews.values()];
    panel.scrollHeight = Math.max(panel.clientHeight, ...visible.map(view => view.wrap.offsetTop + view.wrap.offsetHeight + 24));
    panel.scrollWidth = Math.max(panel.clientWidth, ...visible.map(view => view.wrap.offsetLeft + view.wrap.offsetWidth + 24));
  };
  updateGeometry();
  let clock = 1000;
  let zoomEvents = 0;
  const requestedPages = [];
  const viewportRequests = [];
  const defaultPdfPage = page => ({
    getViewport(options) {
      viewportRequests.push({ page, ...options });
      const size = geometry[pageOrder.indexOf(page)];
      const width = (size.pdfWidth ?? size.width) * options.scale;
      const height = (size.pdfHeight ?? size.height) * options.scale;
      return options.rotation % 180 ? { width: height, height: width } : { width, height };
    }
  });
  const context = vm.createContext({
    state, els: { viewerPanel: panel }, pageWheelNavigator: createPageWheelNavigator(),
    getPageScrollBounds, isFocusPageClippedByViewport,
    clamp: (value, min, max) => Math.min(max, Math.max(min, value)),
    performance: { now: () => clock }, getEffectiveViewMode: () => mode,
    isViewerEventTarget: event => event.inViewer !== false,
    handleCtrlWheelZoom: event => { event.preventDefault(); zoomEvents += 1; },
    getPdfPage: page => { requestedPages.push(page); return loadPdfPage ? loadPdfPage(page, defaultPdfPage) : Promise.resolve(defaultPdfPage(page)); },
    getEffectivePageRotation: (_pdfPage, page) => geometry[pageOrder.indexOf(page)].rotation || 0,
    applyPageVisibility: updateGeometry, ensureCurrentPageVisibleInSingleMode() {},
    getPreferredNeighborDepth: () => 0, scheduleVisibleRegionRender() {},
    updateCurrentPage: page => { state.currentPage = page; state.wheelScrollPosition = null; updateGeometry(); }
  });
  vm.runInContext(rendererFunctions, context);
  return {
    state, panel, context, requestedPages, viewportRequests, defaultPdfPage,
    get zoomEvents() { return zoomEvents; },
    currentBounds: () => getPageScrollBounds(panel, state.pageViews.get(state.currentPage).wrap),
    syncScroll: () => context.updateCurrentPageByScroll(),
    async wheel(values = {}, time = clock + 10) {
      clock = time;
      const event = { deltaX: 0, deltaY: 120, deltaMode: 0, shiftKey: false, ctrlKey: false,
        prevented: false, preventDefault() { this.prevented = true; }, ...values };
      context.handleViewerWheel(event);
      for (let turn = 0; turn < 6; turn += 1) await Promise.resolve();
      return event;
    }
  };
}

for (const mode of ["focus", "single", "spread"]) {
  test(`renderer ${mode}: read the bottom, then enter the next page at its top`, async () => {
    const fixture = createRendererFixture({ mode });
    const before = fixture.currentBounds();
    fixture.panel.scrollTop = before.maxTop - 20;
    assert.equal((await fixture.wheel()).prevented, true);
    assert.equal(fixture.state.currentPage, 1);
    assert.equal(fixture.panel.scrollTop, before.maxTop);
    await fixture.wheel();
    assert.equal(fixture.state.currentPage, 2);
    assert.equal(fixture.panel.scrollTop, fixture.currentBounds().minTop);
    fixture.syncScroll();
    assert.equal(fixture.state.currentPage, 2);
  });
}

test("renderer backward transition enters the previous page at its bottom", async () => {
  const fixture = createRendererFixture();
  fixture.context.updateCurrentPage(2);
  fixture.panel.scrollTop = 0;
  await fixture.wheel({ deltaY: -120 });
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.panel.scrollTop, fixture.currentBounds().maxTop);
});

test("renderer follows reordered pages rather than numeric page IDs", async () => {
  const fixture = createRendererFixture({ pageOrder: [3, 1, 2] });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  assert.equal(fixture.state.currentPage, 1);
});

test("renderer document boundaries retain the first and last page", async () => {
  const fixture = createRendererFixture();
  await fixture.wheel({ deltaY: -120 });
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.panel.scrollTop, 0);
  fixture.context.updateCurrentPage(3);
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel({}, 1400);
  assert.equal(fixture.state.currentPage, 3);
  assert.equal(fixture.panel.scrollTop, fixture.currentBounds().maxTop);
});

test("renderer preserves horizontal reading position across differently sized pages", async () => {
  const fixture = createRendererFixture({ pages: [{ width: 1200, height: 1400 }, { width: 1600, height: 800 }, { width: 1000, height: 1000 }] });
  const before = fixture.currentBounds();
  fixture.panel.scrollTop = before.maxTop;
  fixture.panel.scrollLeft = before.minLeft + (before.maxLeft - before.minLeft) * 0.25;
  await fixture.wheel();
  const after = fixture.currentBounds();
  assert.equal(fixture.panel.scrollTop, after.minTop);
  assert.equal(fixture.panel.scrollLeft, after.minLeft + (after.maxLeft - after.minLeft) * 0.25);
});

test("renderer manual and fit focus pages with no vertical range can turn", async () => {
  for (const zoomMode of ["manual", "fit"]) {
    const fixture = createRendererFixture({ zoomMode, pages: [{ width: 400, height: 300 }, { width: 400, height: 300 }, { width: 400, height: 300 }] });
    await fixture.wheel();
    assert.equal(fixture.state.currentPage, 2);
    assert.equal(fixture.panel.scrollTop, 0);
  }
});

test("renderer Ctrl zoom and horizontal wheel leave the selected page alone", async () => {
  const fixture = createRendererFixture();
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel({ ctrlKey: true });
  assert.equal(fixture.zoomEvents, 1);
  assert.equal(fixture.state.currentPage, 1);
  await fixture.wheel({ shiftKey: true });
  assert.equal(fixture.panel.scrollLeft, 120);
  assert.equal(fixture.state.currentPage, 1);
  await fixture.wheel({ deltaX: 180 });
  assert.equal(fixture.state.currentPage, 1);
});

test("renderer preserves normal continuous scrolling and ignores outside input", async () => {
  const fixture = createRendererFixture({ mode: "single", zoomMode: "fit" });
  assert.equal((await fixture.wheel()).prevented, false);
  fixture.state.zoomMode = "manual";
  assert.equal((await fixture.wheel({ inViewer: false })).prevented, false);
  assert.equal(fixture.state.currentPage, 1);
});

test("manual scrollbar movement releases wheel page selection", async () => {
  const fixture = createRendererFixture({ mode: "single" });
  await fixture.wheel();
  assert.ok(fixture.state.wheelScrollPosition);
  fixture.panel.scrollTop = fixture.state.pageViews.get(2).wrap.offsetTop;
  fixture.syncScroll();
  assert.equal(fixture.state.currentPage, 2);
  assert.equal(fixture.state.wheelScrollPosition, null);
});

test("spread navigation reaches the neighboring page in the same row", async () => {
  const fixture = createRendererFixture({ mode: "spread", pages: [
    { width: 1200, height: 1400, top: 24, left: 24 },
    { width: 1000, height: 1200, top: 24, left: 1246 },
    { width: 1200, height: 1400, top: 1448, left: 24 }
  ] });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  assert.equal(fixture.state.currentPage, 2);
  assert.equal(fixture.panel.scrollTop, fixture.currentBounds().minTop);
  assert.ok(fixture.panel.scrollLeft >= fixture.currentBounds().minLeft);
  fixture.syncScroll();
  assert.equal(fixture.state.currentPage, 2);
  await fixture.wheel({ deltaY: -120 }, 1400);
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.panel.scrollTop, fixture.currentBounds().maxTop);
});

test("a more visible neighbor does not steal selection during a small wheel edge gesture", async () => {
  const fixture = createRendererFixture({ mode: "single", pages: [
    { width: 1200, height: 100 },
    { width: 1200, height: 1400 },
    { width: 1200, height: 1400 }
  ] });
  await fixture.wheel({ deltaY: 20 });
  fixture.syncScroll();
  assert.equal(fixture.state.currentPage, 1);
  await fixture.wheel({ deltaY: 20 });
  assert.equal(fixture.state.currentPage, 2);
  fixture.syncScroll();
  assert.equal(fixture.state.currentPage, 2);
});

test("entering an unvisited previous page uses its real bottom instead of its placeholder", async () => {
  const fixture = createRendererFixture({ pages: [
    { width: 300, height: 150, pdfWidth: 1000, pdfHeight: 1800 },
    { width: 1200, height: 1400 },
    { width: 1200, height: 1400 }
  ] });
  fixture.context.updateCurrentPage(2);
  await fixture.wheel({ deltaY: -120 });
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.state.pageViews.get(1).wrap.offsetHeight, 1800);
  assert.equal(fixture.panel.scrollTop, 1248);
  assert.equal(fixture.state.wheelPageTurn, null);
});

test("entering an unvisited next page preserves horizontal position using real dimensions", async () => {
  const fixture = createRendererFixture({ pages: [
    { width: 1200, height: 1400 },
    { width: 300, height: 150, pdfWidth: 1600, pdfHeight: 2200 },
    { width: 1200, height: 1400 }
  ] });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  fixture.panel.scrollLeft = fixture.currentBounds().maxLeft * 0.75;
  await fixture.wheel();
  assert.equal(fixture.state.currentPage, 2);
  assert.equal(fixture.state.pageViews.get(2).wrap.offsetWidth, 1600);
  assert.equal(fixture.panel.scrollTop, 0);
  assert.equal(fixture.panel.scrollLeft, 636);
});

test("unvisited page geometry respects the active zoom and rotation", async () => {
  const fixture = createRendererFixture({ pages: [
    { width: 1200, height: 1400 },
    { width: 300, height: 150, pdfWidth: 600, pdfHeight: 800, rotation: 90 },
    { width: 1200, height: 1400 }
  ] });
  fixture.state.scale = 2;
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  assert.deepEqual(fixture.viewportRequests, [{ page: 2, scale: 2, rotation: 90 }]);
  assert.equal(fixture.state.pageViews.get(2).wrap.offsetWidth, 1600);
  assert.equal(fixture.state.pageViews.get(2).wrap.offsetHeight, 1200);
});

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((finish, fail) => { resolve = finish; reject = fail; });
  return { promise, resolve, reject };
}

async function flushAsyncNavigation() {
  for (let turn = 0; turn < 8; turn += 1) await Promise.resolve();
}

test("pending page sizing blocks duplicate wheel turns", async () => {
  const pending = deferred();
  const fixture = createRendererFixture({ loadPdfPage: () => pending.promise });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  assert.ok(fixture.state.wheelPageTurn);
  assert.equal(fixture.state.currentPage, 1);
  assert.equal((await fixture.wheel({}, 1600)).prevented, true);
  assert.deepEqual(fixture.requestedPages, [2]);
  pending.resolve(fixture.defaultPdfPage(2));
  await flushAsyncNavigation();
  assert.equal(fixture.state.currentPage, 2);
  assert.equal(fixture.state.wheelPageTurn, null);
});

for (const change of ["document", "page", "scale"]) {
  test(`pending wheel turn is cancelled after the ${change} changes`, async () => {
    const pending = deferred();
    const fixture = createRendererFixture({ loadPdfPage: () => pending.promise });
    fixture.panel.scrollTop = fixture.currentBounds().maxTop;
    await fixture.wheel();
    if (change === "document") fixture.state.pdfDoc = {};
    if (change === "page") fixture.context.updateCurrentPage(3);
    if (change === "scale") fixture.state.scale = 2;
    pending.resolve(fixture.defaultPdfPage(2));
    await flushAsyncNavigation();
    assert.equal(fixture.state.currentPage, change === "page" ? 3 : 1);
    assert.deepEqual(fixture.viewportRequests, []);
    assert.equal(fixture.state.wheelPageTurn, null);
  });
}

test("Ctrl zoom cancels a pending wheel turn without applying its late result", async () => {
  const pending = deferred();
  const fixture = createRendererFixture({ loadPdfPage: () => pending.promise });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  await fixture.wheel({ ctrlKey: true });
  assert.equal(fixture.zoomEvents, 1);
  assert.equal(fixture.state.wheelPageTurn, null);
  pending.resolve(fixture.defaultPdfPage(2));
  await flushAsyncNavigation();
  assert.equal(fixture.state.currentPage, 1);
  assert.deepEqual(fixture.viewportRequests, []);
});

test("failed page sizing releases navigation for a later retry", async () => {
  let attempts = 0;
  const fixture = createRendererFixture({ loadPdfPage: (page, makePage) => {
    attempts += 1;
    return attempts === 1 ? Promise.reject(new Error("Page unavailable")) : Promise.resolve(makePage(page));
  } });
  fixture.panel.scrollTop = fixture.currentBounds().maxTop;
  await fixture.wheel();
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.state.wheelPageTurn, null);
  await fixture.wheel({}, 1500);
  assert.equal(fixture.state.currentPage, 2);
  assert.equal(fixture.state.wheelPageTurn, null);
});

test("a superseded request cannot clear or navigate over a newer request", async () => {
  const earlier = deferred();
  const later = deferred();
  const fixture = createRendererFixture({ loadPdfPage: page => page === 2 ? earlier.promise : later.promise });
  const oldTurn = fixture.context.turnPageFromWheel(2, 1, 0.5);
  fixture.state.wheelPageTurn = null;
  const newTurn = fixture.context.turnPageFromWheel(3, 1, 0.5);
  const activeRequest = fixture.state.wheelPageTurn;
  earlier.resolve(fixture.defaultPdfPage(2));
  await oldTurn;
  assert.equal(fixture.state.currentPage, 1);
  assert.equal(fixture.state.wheelPageTurn, activeRequest);
  later.resolve(fixture.defaultPdfPage(3));
  await newTurn;
  assert.equal(fixture.state.currentPage, 3);
  assert.equal(fixture.state.wheelPageTurn, null);
});
