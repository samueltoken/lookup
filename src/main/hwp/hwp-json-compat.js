const { normalizeSpecialSymbols, normalizeListMarker } = require("./symbol-normalizer");
const {
  hwpUnitToPx,
  createHwpCharStyleResolver,
  createHwpParaStyleResolver,
  inlineStyleToCss,
  paragraphStyleToCss
} = require("./style-profile");
const { createBorderFillResolver, buildTableCellStyle, shouldTreatFirstRowAsHeader } = require("./table-style");

const ORDERED_MARKER_PATTERN = "(?:\\d{1,3}[.)]|[\\u2460-\\u2473]|[\\uAC00-\\uD7A3][.)])";
const SYMBOL_MARKER_PATTERN = "(?:[\\u25CB\\u25CF\\u25E6\\u25AA\\u25AB\\u25A0\\u25A1\\u25C6\\u25C7\\u25B6\\u25B7\\u25B8\\u25B9\\u203B\\u00B7\\u2022\\-\\u2013\\u2014])";
const BULLET_MARKER_PATTERN = `(?:${ORDERED_MARKER_PATTERN}|${SYMBOL_MARKER_PATTERN})`;
const INLINE_LIST_REGEX = new RegExp(`(^|\\s+)(${BULLET_MARKER_PATTERN})(?=\\s*\\S)`, "g");
const LOOSE_INLINE_LIST_REGEX = new RegExp(`(${BULLET_MARKER_PATTERN})(?=\\s*\\S)`, "g");

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizePipeSeparators(text) {
  return String(text || "").replace(/[|｜¦]+/g, "|");
}

function trimDecorativeLeadingBars(text) {
  return String(text || "").replace(/^[|｜¦\s]+/, "").trimStart();
}

function normalizeText(text) {
  const normalized = normalizeSpecialSymbols(String(text || ""))
    .replace(/\r\n/g, "\n")
    .replace(/\u0000/g, "")
    .replace(/\u000b/g, "")
    .replace(/\f/g, "")
    .replace(/\u001f/g, "")
    .replace(/[ \t]+$/gm, "");
  return normalized.trim();
}

function normalizeParagraphText(text) {
  return normalizeSpecialSymbols(String(text || ""))
    .replace(/\r/g, "")
    .replace(/\u0000/g, "")
    .replace(/\u000b/g, "")
    .replace(/\f/g, "")
    .replace(/\u001f/g, "");
}

function parseJsonValue(jsonValue) {
  if (!jsonValue) {
    return null;
  }
  if (typeof jsonValue === "string") {
    try {
      return JSON.parse(jsonValue);
    } catch (_error) {
      return null;
    }
  }
  if (typeof jsonValue === "object") {
    return jsonValue;
  }
  return null;
}

function isOrderedListMarker(marker) {
  const token = String(marker || "").trim();
  return /^\d{1,3}[.)]$/.test(token) || /^[\u2460-\u2473]$/.test(token) || /^[\uAC00-\uD7A3][.)]$/.test(token);
}

function parseListSegmentsFromLine(rawLine) {
  const normalized = normalizePipeSeparators(normalizeSpecialSymbols(String(rawLine || "")))
    .replace(/\t/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) {
    return null;
  }

  const line = trimDecorativeLeadingBars(normalized) || normalized;
  const singleBullet = line.match(new RegExp(`^(${BULLET_MARKER_PATTERN})\\s*(.+)$`));
  if (singleBullet) {
    const marker = normalizeListMarker(singleBullet[1]);
    const itemText = trimDecorativeLeadingBars(singleBullet[2] || "").trim();
    if (!itemText) {
      return null;
    }
    return {
      prefix: "",
      items: [{ marker, text: itemText, ordered: isOrderedListMarker(marker) }],
      listType: isOrderedListMarker(marker) ? "ol" : "ul"
    };
  }

  INLINE_LIST_REGEX.lastIndex = 0;
  const markers = [];
  let match = null;
  while ((match = INLINE_LIST_REGEX.exec(line)) !== null) {
    const leading = match[1] || "";
    const marker = normalizeListMarker(match[2] || "");
    const markerIndex = match.index + leading.length;
    let contentStart = INLINE_LIST_REGEX.lastIndex;
    while (contentStart < line.length && /\s/.test(line[contentStart])) {
      contentStart += 1;
    }
    markers.push({ marker, markerIndex, contentStart });
  }

  if (markers.length < 2) {
    LOOSE_INLINE_LIST_REGEX.lastIndex = 0;
    const looseMarkers = [];
    let loose = null;
    while ((loose = LOOSE_INLINE_LIST_REGEX.exec(line)) !== null) {
      const marker = normalizeListMarker(loose[1] || "");
      const markerIndex = loose.index;
      const prevChar = markerIndex > 0 ? line[markerIndex - 1] : "";
      const isSafeBoundary = markerIndex === 0 || /[\s|｜¦)[\]{}:;,.!?]/.test(prevChar);
      const appearsInsideWord = !isSafeBoundary && /[\p{L}\p{N}]/u.test(prevChar);
      if (appearsInsideWord) {
        continue;
      }
      let contentStart = LOOSE_INLINE_LIST_REGEX.lastIndex;
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
  for (let index = 0; index < markers.length; index += 1) {
    const current = markers[index];
    const nextStart = index + 1 < markers.length ? markers[index + 1].markerIndex : line.length;
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

function appendListContinuation(items, line) {
  if (!items.length) {
    return;
  }
  const text = trimDecorativeLeadingBars(normalizeSpecialSymbols(String(line || ""))).trim();
  if (!text) {
    return;
  }
  const last = items[items.length - 1];
  last.text = `${String(last.text || "")} ${text}`.replace(/\s{2,}/g, " ").trim();
}

function buildCompatListBlock(listType, listItems) {
  const tag = listType === "ol" ? "ol" : "ul";
  const itemsHtml = listItems
    .map((item) => {
      const marker = escapeHtml(normalizeListMarker(item.marker));
      return `<li class="lookup-list-item"><span class="lookup-list-marker">${marker}</span><span class="lookup-list-body">${escapeHtml(item.text)}</span></li>`;
    })
    .join("");
  return `<${tag} class="lookup-list ${tag === "ol" ? "lookup-list-ordered" : "lookup-list-unordered"}">${itemsHtml}</${tag}>`;
}

function extractParagraphTextRecord(paragraph) {
  const records = Array.isArray(paragraph?.records) ? paragraph.records : [];
  return records.find((record) => record && record.type === "para_text") || null;
}

function extractParagraphShapeRecord(paragraph) {
  const records = Array.isArray(paragraph?.records) ? paragraph.records : [];
  return records.find((record) => record && record.type === "para_char_shape") || null;
}

function buildShapeRanges(shapeRecord) {
  const shapeEntries = Array.isArray(shapeRecord?.shapes)
    ? shapeRecord.shapes
    : Array.isArray(shapeRecord?.char_shapes)
      ? shapeRecord.char_shapes
      : [];
  const ranges = shapeEntries
    .map((shape) => ({
      position: Math.max(0, Number(shape?.position || 0)),
      shapeId: Number(shape?.shape_id)
    }))
    .filter((shape) => Number.isFinite(shape.shapeId))
    .sort((a, b) => a.position - b.position);
  if (!ranges.length) {
    ranges.push({ position: 0, shapeId: -1 });
  }
  return ranges;
}

function getShapeIdAt(ranges, charIndex) {
  let activeShapeId = ranges[0]?.shapeId;
  for (const range of ranges) {
    if (range.position <= charIndex) {
      activeShapeId = range.shapeId;
      continue;
    }
    break;
  }
  return activeShapeId;
}

function isSameInlineStyle(a = {}, b = {}) {
  return (
    Boolean(a.bold) === Boolean(b.bold)
    && Boolean(a.italic) === Boolean(b.italic)
    && Boolean(a.underline) === Boolean(b.underline)
    && Boolean(a.strike) === Boolean(b.strike)
    && Number(a.fontSizePx || 0).toFixed(2) === Number(b.fontSizePx || 0).toFixed(2)
    && String(a.color || "") === String(b.color || "")
    && String(a.backgroundColor || "") === String(b.backgroundColor || "")
    && String(a.fontFamily || "") === String(b.fontFamily || "")
    && Number(a.letterSpacingEm || 0).toFixed(3) === Number(b.letterSpacingEm || 0).toFixed(3)
    && Number(a.textStrokeWidthPx || 0).toFixed(2) === Number(b.textStrokeWidthPx || 0).toFixed(2)
    && String(a.textStrokeColor || "") === String(b.textStrokeColor || "")
    && String(a.textShadow || "") === String(b.textShadow || "")
  );
}

function extractParagraphStyledSegments(paragraph, resolveCharStyle) {
  const textRecord = extractParagraphTextRecord(paragraph);
  const rawText = normalizeParagraphText(textRecord?.text || "");
  if (!rawText) {
    return [];
  }

  const shapeRecord = extractParagraphShapeRecord(paragraph);
  const ranges = buildShapeRanges(shapeRecord);
  const segments = [];
  let buffer = "";
  let bufferStyle = resolveCharStyle(-1);

  const flushBuffer = () => {
    if (!buffer) {
      return;
    }
    segments.push({
      text: buffer,
      style: bufferStyle
    });
    buffer = "";
  };

  for (let index = 0; index < rawText.length; index += 1) {
    const char = rawText[index];
    if (!char || /[\u0000\r]/.test(char)) {
      continue;
    }
    const shapeId = getShapeIdAt(ranges, index);
    const style = resolveCharStyle(shapeId);
    if (buffer && !isSameInlineStyle(bufferStyle, style)) {
      flushBuffer();
    }
    bufferStyle = style;
    buffer += char;
  }
  flushBuffer();
  return segments;
}

function segmentsToPlainText(segments) {
  return segments.map((segment) => normalizeSpecialSymbols(segment.text || "")).join("");
}

function renderStyledSegments(segments) {
  if (!Array.isArray(segments) || !segments.length) {
    return "";
  }
  return segments
    .map((segment) => {
      const escaped = escapeHtml(normalizeSpecialSymbols(segment.text || "")).replace(/\n/g, "<br />");
      if (!escaped) {
        return "";
      }
      const styleAttr = inlineStyleToCss(segment.style || {});
      if (!styleAttr) {
        return escaped;
      }
      return `<span style="${styleAttr}">${escaped}</span>`;
    })
    .join("");
}

function isLikelyHeadingLine(text) {
  const line = String(text || "").trim();
  if (!line) {
    return false;
  }
  if (/^#{1,6}\s+/.test(line)) {
    return true;
  }
  if (/^[?졻뀫?™뀭?ㅲ뀯?╈뀱?ⓥ뀳][\s.)]/.test(line)) {
    return true;
  }
  if (/^[IVXLCM]{1,6}[\s.)]/i.test(line)) {
    return true;
  }
  return false;
}

function isLikelyListContinuation(lineText, paraStyle) {
  const text = String(lineText || "").trim();
  if (!text) {
    return false;
  }
  if (parseListSegmentsFromLine(text)) {
    return false;
  }
  if (isLikelyHeadingLine(text)) {
    return false;
  }
  if (String(paraStyle?.textAlign || "") === "justify") {
    return true;
  }
  return true;
}

function findTableRecord(paragraph) {
  const records = Array.isArray(paragraph?.records) ? paragraph.records : [];
  for (const record of records) {
    if (!record || record.type !== "ctrl_header") {
      continue;
    }
    if (String(record.ctrl_id || "").trim() !== "tbl") {
      continue;
    }
    const children = Array.isArray(record.children) ? record.children : [];
    const tableChild = children.find((child) => child && child.type === "table" && child.table);
    if (tableChild && tableChild.table) {
      return tableChild.table;
    }
  }
  return null;
}

function readNumericValue(source, keys, fallback = 0) {
  const target = source && typeof source === "object" ? source : {};
  for (const key of Array.isArray(keys) ? keys : []) {
    if (!key) {
      continue;
    }
    const value = Number(target[key]);
    if (Number.isFinite(value)) {
      return value;
    }
  }
  return fallback;
}

function readTableCellAddress(attr, axis, fallback = 0) {
  const keys = axis === "row"
    ? ["row_address", "rowAddress", "row", "row_index", "rowIndex", "r"]
    : ["col_address", "colAddress", "col", "column", "col_index", "colIndex", "columnIndex", "c"];
  return Math.max(0, Math.floor(readNumericValue(attr, keys, fallback)));
}

function readTableCellSpan(attr, axis) {
  const keys = axis === "row"
    ? ["row_span", "rowSpan", "rowspan", "row_merge", "rowMerge"]
    : ["col_span", "colSpan", "colspan", "col_merge", "colMerge"];
  return Math.max(1, Math.floor(readNumericValue(attr, keys, 1)));
}

function readTableCellMeasure(attr, axis) {
  const keys = axis === "width"
    ? ["width", "cell_width", "cellWidth", "w"]
    : ["height", "cell_height", "cellHeight", "h"];
  return Math.max(0, readNumericValue(attr, keys, 0));
}

function flattenNumericValues(value) {
  const out = [];
  const visit = (item) => {
    if (Array.isArray(item)) {
      item.forEach(visit);
      return;
    }
    if (item && typeof item === "object") {
      if ("width" in item || "w" in item || "value" in item) {
        visit(item.width ?? item.w ?? item.value);
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

function resolveTableSize(table) {
  const attributes = table?.attributes || {};
  const explicitRows = Math.max(
    0,
    Math.floor(readNumericValue(attributes, ["row_count", "rowCount", "rows", "row_cnt"], 0))
  );
  const explicitCols = Math.max(
    0,
    Math.floor(readNumericValue(attributes, ["col_count", "colCount", "cols", "column_count", "columnCount", "col_cnt"], 0))
  );
  const cells = Array.isArray(table?.cells) ? table.cells : [];
  let maxRow = 0;
  let maxCol = 0;
  for (const cell of cells) {
    const attr = cell?.cell_attributes || {};
    const row = readTableCellAddress(attr, "row", 0);
    const col = readTableCellAddress(attr, "col", 0);
    const rowSpan = readTableCellSpan(attr, "row");
    const colSpan = readTableCellSpan(attr, "col");
    maxRow = Math.max(maxRow, row + rowSpan);
    maxCol = Math.max(maxCol, col + colSpan);
  }
  return {
    rows: Math.max(1, explicitRows || maxRow),
    cols: Math.max(1, explicitCols || maxCol)
  };
}

function buildTableDimensionValues(table, count, axis) {
  const values = Array.from({ length: count }, () => 0);
  const cells = Array.isArray(table?.cells) ? table.cells : [];
  const spans = [];

  if (axis === "col") {
    const declared = flattenNumericValues(
      table?.attributes?.row_cols
      ?? table?.attributes?.rowCols
      ?? table?.attributes?.col_widths
      ?? table?.attributes?.colWidths
      ?? table?.colWidths
    );
    declared.slice(0, count).forEach((value, index) => {
      values[index] = Math.max(values[index], value);
    });
  }

  for (const cell of cells) {
    const attr = cell?.cell_attributes || {};
    const start = readTableCellAddress(attr, axis === "row" ? "row" : "col", 0);
    const span = readTableCellSpan(attr, axis === "row" ? "row" : "col");
    const measure = readTableCellMeasure(attr, axis === "row" ? "height" : "width");
    if (start >= count || measure <= 0) {
      continue;
    }
    const end = Math.min(count, start + span);
    if (end <= start) {
      continue;
    }
    if (span <= 1) {
      values[start] = Math.max(values[start], measure);
    } else {
      spans.push({ start, end, measure });
    }
  }

  for (let pass = 0; pass < 4; pass += 1) {
    for (const item of spans) {
      const width = item.end - item.start;
      if (width <= 0) {
        continue;
      }
      const currentTotal = values
        .slice(item.start, item.end)
        .reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0);
      if (currentTotal <= 0) {
        const even = item.measure / width;
        for (let index = item.start; index < item.end; index += 1) {
          values[index] = Math.max(values[index], even);
        }
        continue;
      }
      if (currentTotal < item.measure * 0.985) {
        const add = (item.measure - currentTotal) / width;
        for (let index = item.start; index < item.end; index += 1) {
          values[index] = Math.max(0, values[index] + add);
        }
      }
    }
  }

  return values.map((value) => Math.max(0, Number(value) || 0));
}

function buildTableColWidths(table, colCount) {
  return buildTableDimensionValues(table, colCount, "col");
}

function buildTableRowHeights(table, rowCount) {
  return buildTableDimensionValues(table, rowCount, "row");
}

function extractCellPlainText(cell, resolveCharStyle) {
  const lines = [];
  for (const paragraph of Array.isArray(cell?.paragraphs) ? cell.paragraphs : []) {
    const segments = extractParagraphStyledSegments(paragraph, resolveCharStyle);
    const plain = normalizeText(trimDecorativeLeadingBars(segmentsToPlainText(segments)));
    if (plain) {
      lines.push(plain);
    }
  }
  return lines.join(" ").trim();
}

function buildCellContent(cell, resolveCharStyle, resolveParaStyle) {
  const paragraphs = Array.isArray(cell?.paragraphs) ? cell.paragraphs : [];
  const parts = [];
  let listState = null;
  const alignVotes = { left: 0, right: 0, center: 0, justify: 0 };

  const flushList = () => {
    if (!listState || !Array.isArray(listState.items) || !listState.items.length) {
      listState = null;
      return;
    }
    parts.push(buildCompatListBlock(listState.listType, listState.items));
    listState = null;
  };

  for (const paragraph of paragraphs) {
    const paraShapeId = Number(paragraph?.para_header?.para_shape_id);
    const paraStyle = resolveParaStyle(paraShapeId);
    const align = String(paraStyle?.textAlign || "").trim();
    if (align && Object.prototype.hasOwnProperty.call(alignVotes, align)) {
      alignVotes[align] += 1;
    }
    const segments = extractParagraphStyledSegments(paragraph, resolveCharStyle);
    const plain = trimDecorativeLeadingBars(segmentsToPlainText(segments)).trim();
    if (!plain) {
      flushList();
      continue;
    }
    const parsedList = parseListSegmentsFromLine(plain);
    if (parsedList) {
      if (!listState || listState.listType !== parsedList.listType) {
        flushList();
        listState = {
          listType: parsedList.listType,
          items: []
        };
      }
      if (parsedList.prefix) {
        if (listState.items.length) {
          appendListContinuation(listState.items, parsedList.prefix);
        } else {
          parts.push(escapeHtml(parsedList.prefix));
        }
      }
      for (const item of parsedList.items) {
        listState.items.push({
          marker: normalizeListMarker(item.marker),
          text: normalizeSpecialSymbols(item.text)
        });
      }
      continue;
    }
    if (listState && isLikelyListContinuation(plain, paraStyle)) {
      appendListContinuation(listState.items, plain);
      continue;
    }
    flushList();
    const paragraphStyle = paragraphStyleToCss(paraStyle);
    const styleAttr = paragraphStyle ? ` style="${paragraphStyle}"` : "";
    parts.push(`<p class="lookup-cell-paragraph"${styleAttr}>${renderStyledSegments(segments)}</p>`);
  }
  flushList();

  const voteEntries = Object.entries(alignVotes).sort((a, b) => b[1] - a[1]);
  const dominantAlign = voteEntries[0] && voteEntries[0][1] > 0 ? voteEntries[0][0] : "";
  return {
    html: parts.length ? parts.join("") : "&nbsp;",
    horizontalAlign: dominantAlign === "justify" ? "" : dominantAlign
  };
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
  const hasQuestion = normalized.some((token) => token.includes("질문") || token.includes("문항") || token.includes("질의") || token.includes("question"));
  const hasAnswer = normalized.some((token) => token.includes("답변") || token.includes("응답") || token.includes("답") || token.includes("answer"));
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
function resolveHeaderRowIndex(rowCount, colCount, getCellText) {
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

function detectCenterPreferredColumns({
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
    const scanRows = Math.max(0, Math.min(10, Math.floor(Number(rowCount) || 0)));
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
    const isNarrowColumn = totalWidth > 0 && width > 0 && width / totalWidth <= 0.14;
    if (looksIndexHeader || (isNarrowColumn && textCount > 0 && numericCount >= Math.max(1, textCount - 1))) {
      columns.add(colIndex);
    }
  }
  return columns;
}
function buildTableBlock(table, context = {}) {
  const resolveCharStyle = typeof context.resolveCharStyle === "function" ? context.resolveCharStyle : () => ({});
  const resolveParaStyle = typeof context.resolveParaStyle === "function" ? context.resolveParaStyle : () => ({});
  const borderFillResolver = typeof context.borderFillResolver === "function" ? context.borderFillResolver : () => null;
  const size = resolveTableSize(table);
  const rowCount = size.rows;
  const colCount = size.cols;
  const rowStart = Math.max(0, Math.floor(Number(context.rowStart || 0)));
  const rowEnd = Math.max(rowStart, Math.min(rowCount, Math.floor(Number(context.rowEnd || rowCount))));
  const repeatHeader = context.repeatHeader === true && rowStart > 0;
  const rowHeightScale = Number.isFinite(Number(context.rowHeightScale)) && Number(context.rowHeightScale) > 0
    ? Number(context.rowHeightScale)
    : 1;
  const extraClass = String(context.extraClass || "").replace(/[^a-z0-9_-]+/gi, " ").trim();
  const cells = Array.isArray(table?.cells) ? table.cells : [];
  if (!cells.length) {
    return "";
  }

  const ownerMap = new Map();
  const coveredMap = new Set();
  for (const cell of cells) {
    const attr = cell?.cell_attributes || {};
    const row = readTableCellAddress(attr, "row", 0);
    const col = readTableCellAddress(attr, "col", 0);
    const rowSpan = readTableCellSpan(attr, "row");
    const colSpan = readTableCellSpan(attr, "col");
    ownerMap.set(`${row}:${col}`, cell);
    for (let r = row; r < Math.min(row + rowSpan, rowCount); r += 1) {
      for (let c = col; c < Math.min(col + colSpan, colCount); c += 1) {
        if (r === row && c === col) {
          continue;
        }
        coveredMap.add(`${r}:${c}`);
      }
    }
  }

  const colWidths = buildTableColWidths(table, colCount);
  const rowHeights = buildTableRowHeights(table, rowCount);
  const plainTextCache = new Map();
  const getCellPlainText = (rowIndex, colIndex) => {
    const key = `${rowIndex}:${colIndex}`;
    if (plainTextCache.has(key)) {
      return plainTextCache.get(key);
    }
    const cell = ownerMap.get(key);
    const text = cell ? extractCellPlainText(cell, resolveCharStyle) : "";
    plainTextCache.set(key, text);
    return text;
  };
  const headerRowIndex = resolveHeaderRowIndex(rowCount, colCount, getCellPlainText);
  const centerColumns = detectCenterPreferredColumns({
    rowCount,
    colCount,
    headerRowIndex,
    getCellText: getCellPlainText,
    colWidths
  });

  const validWidths = colWidths.filter((value) => value > 0);
  const totalWidth = validWidths.reduce((sum, value) => sum + value, 0);
  const tableWidthPx = hwpUnitToPx(totalWidth);
  const tableStyle = tableWidthPx > 32
    ? ` style="width:${Math.min(760, Math.max(120, tableWidthPx)).toFixed(1)}px;max-width:100%;"`
    : "";
  const colgroup = validWidths.length && totalWidth > 0
    ? `<colgroup>${colWidths
      .map((value) => {
        if (value <= 0) {
          return "<col />";
        }
        const ratio = (value / totalWidth) * 100;
        return `<col style="width:${ratio.toFixed(3)}%;" />`;
      })
      .join("")}</colgroup>`
    : "";

  const rowHtml = [];
  const tableAttributes = table?.attributes || {};
  const renderRows = [];
  if (repeatHeader) {
    renderRows.push(0);
  }
  for (let rowIndex = rowStart; rowIndex < rowEnd; rowIndex += 1) {
    if (!renderRows.includes(rowIndex)) {
      renderRows.push(rowIndex);
    }
  }

  for (const rowIndex of renderRows) {
    const cellHtml = [];
    const rowHeightPx = hwpUnitToPx(rowHeights[rowIndex] || 0) * rowHeightScale;
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
        cellHtml.push("<td>&nbsp;</td>");
        continue;
      }
      const attr = cell?.cell_attributes || {};
      const colSpan = readTableCellSpan(attr, "col");
      const rowSpan = Math.max(1, Math.min(readTableCellSpan(attr, "row"), rowEnd - rowIndex));
      const cellContent = buildCellContent(cell, resolveCharStyle, resolveParaStyle);
      const isHeader = rowIndex === headerRowIndex;
      const centerPreferred = centerColumns.has(colIndex);
      const preferredAlign = isHeader ? "center" : (cellContent.horizontalAlign || (centerPreferred ? "center" : ""));
      const inferredCenter = preferredAlign === "center" || (isHeader && rowSpan === 1) || centerPreferred;
      const cellStyle = buildTableCellStyle({
        cellAttributes: attr,
        tableAttributes,
        borderFillResolver,
        horizontalAlign: preferredAlign,
        columnRole: centerPreferred ? "index" : "",
        isHeader,
        rowSpan,
        cellWidth: readTableCellMeasure(attr, "width"),
        cellHeight: readTableCellMeasure(attr, "height"),
        forceMiddle: inferredCenter
      });
      const styleEntries = Object.entries(cellStyle);
      const styleAttr = styleEntries.length
        ? ` style="${styleEntries.map(([name, value]) => `${name}:${value}`).join(";")}"`
        : "";
      const tag = isHeader ? "th" : "td";
      cellHtml.push(`<${tag} colspan="${colSpan}" rowspan="${rowSpan}"${styleAttr}>${cellContent.html}</${tag}>`);
    }
    rowHtml.push(`<tr${rowStyle}>${cellHtml.join("")}</tr>`);
  }

  const className = `lookup-table-trace lookup-table-structured${extraClass ? ` ${extraClass}` : ""}`;
  return `<table class="${className}"${tableStyle}>${colgroup}<tbody>${rowHtml.join("")}</tbody></table>`;
}

function buildHwpBlocksFromHwpJson(jsonValue) {
  const parsed = parseJsonValue(jsonValue);
  if (!parsed || !parsed.body_text || !Array.isArray(parsed.body_text.sections)) {
    return [];
  }

  const sections = parsed.body_text.sections;
  const resolveCharStyle = createHwpCharStyleResolver(parsed);
  const resolveParaStyle = createHwpParaStyleResolver(parsed);
  const borderFillResolver = createBorderFillResolver(parsed);
  const blocks = [];
  let listState = null;

  const flushList = () => {
    if (!listState || !Array.isArray(listState.items) || !listState.items.length) {
      listState = null;
      return;
    }
    blocks.push(buildCompatListBlock(listState.listType, listState.items));
    listState = null;
  };

  for (const section of sections) {
    const paragraphs = Array.isArray(section?.paragraphs) ? section.paragraphs : [];
    for (const paragraph of paragraphs) {
      const table = findTableRecord(paragraph);
      if (table) {
        flushList();
        const tableBlock = buildTableBlock(table, {
          resolveCharStyle,
          resolveParaStyle,
          borderFillResolver
        });
        if (tableBlock) {
          blocks.push(tableBlock);
        }
        continue;
      }

      const paraShapeId = Number(paragraph?.para_header?.para_shape_id);
      const paraStyle = resolveParaStyle(paraShapeId);
      const segments = extractParagraphStyledSegments(paragraph, resolveCharStyle);
      const plainText = trimDecorativeLeadingBars(segmentsToPlainText(segments)).trim();
      if (!plainText) {
        flushList();
        continue;
      }

      const parsedList = parseListSegmentsFromLine(plainText);
      if (parsedList) {
        if (!listState || listState.listType !== parsedList.listType) {
          flushList();
          listState = {
            listType: parsedList.listType,
            items: []
          };
        }
        if (parsedList.prefix) {
          if (listState.items.length) {
            appendListContinuation(listState.items, parsedList.prefix);
          } else {
            blocks.push(`<p>${escapeHtml(parsedList.prefix)}</p>`);
          }
        }
        for (const item of parsedList.items) {
          listState.items.push({
            marker: normalizeListMarker(item.marker),
            text: normalizeSpecialSymbols(item.text)
          });
        }
        continue;
      }

      if (listState && isLikelyListContinuation(plainText, paraStyle)) {
        appendListContinuation(listState.items, plainText);
        continue;
      }

      flushList();
      const paraStyleText = paragraphStyleToCss(paraStyle);
      const paraAttr = paraStyleText ? ` style="${paraStyleText}"` : "";
      blocks.push(`<p${paraAttr}>${renderStyledSegments(segments)}</p>`);
    }
  }

  flushList();
  return blocks;
}

module.exports = {
  buildHwpBlocksFromHwpJson,
  parseJsonValue,
  escapeHtml,
  trimDecorativeLeadingBars,
  parseListSegmentsFromLine,
  buildCompatListBlock,
  findTableRecord,
  readTableCellAddress,
  readTableCellSpan,
  readTableCellMeasure,
  resolveTableSize,
  buildTableRowHeights,
  buildTableColWidths,
  extractCellPlainText,
  extractParagraphStyledSegments,
  segmentsToPlainText,
  renderStyledSegments,
  buildTableBlock
};


