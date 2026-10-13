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
