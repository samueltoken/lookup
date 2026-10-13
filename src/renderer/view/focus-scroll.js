function clamp(value, min, max) {
  return Math.min(max, Math.max(min, Number(value) || 0));
}

export function isFocusPageClippedByViewport(viewerPanel, pageWrap) {
  if (!viewerPanel || !pageWrap) {
    return false;
  }
  const clippedY = pageWrap.offsetHeight > viewerPanel.clientHeight + 8;
  const clippedX = pageWrap.offsetWidth > viewerPanel.clientWidth + 8;
  return clippedY || clippedX;
}

export function estimateFocusFitScale(viewerPanel, pageWrap, currentScale) {
  if (!viewerPanel || !pageWrap) {
    return 0;
  }
  const safeScale = Math.max(0.01, Number(currentScale) || 1);
  const baseWidth = Math.max(1, pageWrap.offsetWidth / safeScale);
  const baseHeight = Math.max(1, pageWrap.offsetHeight / safeScale);
  const maxWidth = Math.max(100, viewerPanel.clientWidth - 42);
  const maxHeight = Math.max(100, viewerPanel.clientHeight - 44);
  const fitWidthScale = maxWidth / baseWidth;
  const fitHeightScale = maxHeight / baseHeight;
  const fitScale = Math.min(fitWidthScale, fitHeightScale);
  if (!Number.isFinite(fitScale) || fitScale <= 0) {
    return 0;
  }
  return clamp(fitScale, 0.1, 10);
}

export function isFocusContinuousMode({
  effectiveViewMode,
  zoomMode,
  scale,
  viewerPanel,
  currentPageWrap
}) {
  if (effectiveViewMode !== "focus") {
    return false;
  }
  if (isFocusPageClippedByViewport(viewerPanel, currentPageWrap)) {
    return true;
  }
  if (zoomMode === "manual") {
    const fitScale = estimateFocusFitScale(viewerPanel, currentPageWrap, scale);
    if (fitScale > 0 && Number(scale || 0) > fitScale + 0.01) {
      return true;
    }
  }
  return false;
}

export function preferredNeighborDepth(isContinuousMode) {
  return isContinuousMode ? 4 : 1;
}

export function getPageScrollBounds(panel, page) {
  const panelMaxTop = Math.max(0, panel.scrollHeight - panel.clientHeight);
  const panelMaxLeft = Math.max(0, panel.scrollWidth - panel.clientWidth);
  const minTop = clamp(page.offsetTop - 24, 0, panelMaxTop);
  const minLeft = clamp(page.offsetLeft - 24, 0, panelMaxLeft);
  return {
    minTop,
    maxTop: clamp(page.offsetTop + page.offsetHeight + 24 - panel.clientHeight, minTop, panelMaxTop),
    minLeft,
    maxLeft: clamp(page.offsetLeft + page.offsetWidth + 24 - panel.clientWidth, minLeft, panelMaxLeft)
  };
}

export function createPageWheelNavigator() {
  let edgePage = null;
  let edgeDirection = 0;
  let edgeDistance = 0;
  let lastInputTime = -Infinity;
  let lastTurnTime = -Infinity;

  function reset() {
    edgePage = null;
    edgeDirection = 0;
    edgeDistance = 0;
    lastInputTime = -Infinity;
    lastTurnTime = -Infinity;
  }

  function consume({ pageKey, bounds, scrollTop, scrollLeft, deltaX = 0, deltaY = 0,
    deltaMode = 0, shiftKey = false, viewportHeight, viewportWidth, now }) {
    const unitY = deltaMode === 1 ? 16 : deltaMode === 2 ? viewportHeight : 1;
    const unitX = deltaMode === 1 ? 16 : deltaMode === 2 ? viewportWidth : 1;
    const dy = shiftKey ? 0 : deltaY * unitY;
    const dx = (shiftKey ? deltaX || deltaY : deltaX) * unitX;
    const result = {
      scrollTop: clamp(scrollTop + dy, bounds.minTop, bounds.maxTop),
      scrollLeft: clamp(scrollLeft + dx, bounds.minLeft, bounds.maxLeft),
      pageDirection: 0
    };
    const direction = Math.sign(dy);
    // Reaching the edge consumes this input. Only further vertical input may turn a page.
    const atEdge = direction > 0 ? scrollTop >= bounds.maxTop - 1.5 : scrollTop <= bounds.minTop + 1.5;
    if (!direction || Math.abs(dy) <= Math.abs(dx) || !atEdge) {
      edgePage = null;
      edgeDistance = 0;
      lastInputTime = now;
      return result;
    }
    if (edgePage !== pageKey || edgeDirection !== direction || now - lastInputTime > 180) {
      edgeDistance = 0;
    }
    edgePage = pageKey;
    edgeDirection = direction;
    lastInputTime = now;
    if (now - lastTurnTime < 280) {
      edgeDistance = 0;
      return result;
    }
    edgeDistance += Math.abs(dy);
    if (edgeDistance >= 36) {
      result.pageDirection = direction;
      lastTurnTime = now;
      edgeDistance = 0;
    }
    return result;
  }

  return { consume, reset };
}
