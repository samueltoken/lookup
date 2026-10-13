const { hwpUnitToPx, toCssColor } = require("./style-profile");

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, Number(value) || 0));
}

function hasColor(color) {
  return Boolean(toCssColor(color));
}

function borderWidthToPx(widthCode) {
  const numeric = Number(widthCode);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  if (numeric <= 0) {
    return 0;
  }
  return clamp(numeric * 0.14, 0.45, 3.2);
}

function normalizeBorderLineStyle(lineType) {
  const type = Number(lineType);
  if (!Number.isFinite(type) || type <= 0) {
    return "none";
  }
  if (type === 2 || type === 5 || type === 6) {
    return "dashed";
  }
  if (type === 3 || type === 4 || type === 7) {
    return "dotted";
  }
  if (type === 8 || type === 9 || type === 10) {
    return "double";
  }
  return "solid";
}

function resolveBorderFillList(document = {}) {
  if (Array.isArray(document?.doc_info?.border_fill)) {
    return document.doc_info.border_fill;
  }
  if (Array.isArray(document?.doc_info?.border_fills)) {
    return document.doc_info.border_fills;
  }
  if (Array.isArray(document?.head?.borderFills)) {
    return document.head.borderFills;
  }
  return [];
}

function createBorderFillResolver(document = {}) {
  const list = resolveBorderFillList(document);
  const byExplicitId = new Map();
  list.forEach((item, index) => {
    const explicitId = Number(item?.id);
    if (Number.isFinite(explicitId)) {
      byExplicitId.set(explicitId, item);
    }
  });
  return (borderFillId) => {
    const raw = Number(borderFillId);
    if (!Number.isFinite(raw)) {
      return null;
    }
    if (byExplicitId.has(raw)) {
      return byExplicitId.get(raw) || null;
    }
    const floor = Math.floor(raw);
    const candidates = [];
    if (floor >= 0) {
      if (floor > 0) {
        candidates.push(floor - 1);
      }
      candidates.push(floor);
      if (floor + 1 < list.length) {
        candidates.push(floor + 1);
      }
    }
    const unique = Array.from(new Set(candidates));
    for (const index of unique) {
      if (index >= 0 && index < list.length && list[index]) {
        return list[index];
      }
    }
    return null;
  };
}

function buildBorderCss(borderFill) {
  const style = {};
  if (!borderFill || typeof borderFill !== "object") {
    return style;
  }

  const borders = Array.isArray(borderFill.borders) ? borderFill.borders : [];
  const sideNames = ["left", "right", "top", "bottom"];
  for (let index = 0; index < 4; index += 1) {
    const border = borders[index] || {};
    const lineType = Number(border.line_type || border.lineType || 0);
    const lineStyle = normalizeBorderLineStyle(lineType);
    let widthPx = borderWidthToPx(border.width);
    if (lineStyle === "none") {
      style[`border-${sideNames[index]}`] = "none";
      continue;
    }
    if (widthPx <= 0) {
      widthPx = 0.45;
    }
    const color = hasColor(border.color) ? toCssColor(border.color) : "rgb(30, 41, 59)";
    style[`border-${sideNames[index]}`] = `${widthPx.toFixed(2)}px ${lineStyle} ${color}`;
  }

  const fill = borderFill.fill || {};
  const fillType = String(fill.type || "").toLowerCase();
  const fillColor = fill.background_color || fill.backgroundColor || fill.color || null;
  if ((fillType === "solid" || fillType === "color") && hasColor(fillColor)) {
    style["background-color"] = toCssColor(fillColor);
  }
  return style;
}

function readVerticalAlign(cellAttributes = {}) {
  const raw = String(
    cellAttributes.vertical_align
    ?? cellAttributes.verticalAlign
    ?? cellAttributes.vAlign
    ?? cellAttributes.valign
    ?? cellAttributes.textVerticalAlign
    ?? ""
  ).trim().toLowerCase();
  if (raw.includes("center") || raw.includes("middle") || raw === "1") {
    return "middle";
  }
  if (raw.includes("bottom") || raw === "2") {
    return "bottom";
  }
  if (raw.includes("top") || raw === "0") {
    return "top";
  }
  return "";
}

function buildTableCellStyle({
  cellAttributes = {},
  tableAttributes = {},
  borderFillResolver = () => null,
  horizontalAlign = "",
  columnRole = "",
  isHeader = false,
  rowSpan = 1,
  cellWidth = 0,
  cellHeight = 0,
  forceMiddle = false
} = {}) {
  const style = {
    "box-sizing": "border-box"
  };
  const borderId = Number(
    cellAttributes.border_fill_id
    ?? cellAttributes.borderFillId
    ?? tableAttributes.border_fill_id
    ?? tableAttributes.borderFillId
    ?? -1
  );
  const borderFill = Number.isFinite(borderId) ? borderFillResolver(borderId) : null;
  Object.assign(style, buildBorderCss(borderFill));

  const widthPx = hwpUnitToPx(cellWidth || cellAttributes.width || 0);
  const heightPx = hwpUnitToPx(cellHeight || cellAttributes.height || 0);
  if (widthPx > 12) {
    style.width = `${clamp(widthPx, 24, 1800).toFixed(1)}px`;
  }
  if (heightPx > 12) {
    const safeHeight = clamp(heightPx, 12, 1800).toFixed(1);
    style.height = `${safeHeight}px`;
    style["min-height"] = `${safeHeight}px`;
  }

  const leftPaddingPx = hwpUnitToPx(cellAttributes.left_margin || cellAttributes.leftMargin || 0);
  const rightPaddingPx = hwpUnitToPx(cellAttributes.right_margin || cellAttributes.rightMargin || 0);
  const topPaddingPx = hwpUnitToPx(cellAttributes.top_margin || cellAttributes.topMargin || 0);
  const bottomPaddingPx = hwpUnitToPx(cellAttributes.bottom_margin || cellAttributes.bottomMargin || 0);
  if (leftPaddingPx > 0 || rightPaddingPx > 0 || topPaddingPx > 0 || bottomPaddingPx > 0) {
    const top = clamp(topPaddingPx, 0, 32);
    const right = clamp(rightPaddingPx, 0, 32);
    const bottom = clamp(bottomPaddingPx, 0, 32);
    const left = clamp(leftPaddingPx, 0, 32);
    style.padding = `${top.toFixed(1)}px ${right.toFixed(1)}px ${bottom.toFixed(1)}px ${left.toFixed(1)}px`;
  }

  if (horizontalAlign) {
    style["text-align"] = horizontalAlign;
  } else if (columnRole === "index") {
    style["text-align"] = "center";
  } else if (isHeader) {
    style["text-align"] = "center";
  }

  const explicitVerticalAlign = readVerticalAlign(cellAttributes);
  const shouldMiddle = forceMiddle || isHeader || Number(rowSpan || 1) > 1 || columnRole === "index";
  style["vertical-align"] = explicitVerticalAlign || (shouldMiddle ? "middle" : "top");
  if (isHeader) {
    style["font-weight"] = "700";
  }
  return style;
}

function normalizeHeaderToken(value) {
  return String(value || "")
    .normalize("NFC")
    .replace(/\s+/g, "")
    .replace(/[^\p{L}\p{N}]/gu, "")
    .toLowerCase();
}

function containsAny(text, fragments) {
  return fragments.some((fragment) => fragment && text.includes(fragment));
}

function shouldTreatFirstRowAsHeader(rowTexts = []) {
  if (!Array.isArray(rowTexts) || !rowTexts.length) {
    return false;
  }

  const normalized = rowTexts
    .map((value) => normalizeHeaderToken(value))
    .filter(Boolean);
  if (!normalized.length) {
    return false;
  }

  const hasQuestion = normalized.some((token) =>
    containsAny(token, ["질문", "문항", "질의", "question"])
  );
  const hasAnswer = normalized.some((token) =>
    containsAny(token, ["답변", "응답", "답", "answer"])
  );
  const hasIndex = normalized.some((token) =>
    containsAny(token, ["연번", "번호", "순번", "index", "no"])
  );

  if (hasQuestion && hasAnswer) {
    return true;
  }
  return hasIndex && normalized.length >= 3;
}

module.exports = {
  createBorderFillResolver,
  buildTableCellStyle,
  shouldTreatFirstRowAsHeader
};

