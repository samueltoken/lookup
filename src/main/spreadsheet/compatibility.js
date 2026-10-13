"use strict";

const path = require("node:path");
const XLSX = require("xlsx");

const SPREADSHEET_RENDER_VERSION = "complete-sheets-v2";
const MAX_RENDERED_CELLS = 250000;
const MAX_RENDERED_SHEETS = 100;
const PAGE_TABLE_WIDTH = 1000;

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}

function readSpreadsheetWorkbook(sourcePath) {
  // Dense worksheets contain row arrays, not A1-addressed cell objects. Keep one
  // representation for inspection, formatting and the compatibility renderer.
  const workbook = XLSX.readFile(sourcePath, {
    dense: false, cellDates: false, cellStyles: true, cellNF: true,
    cellText: true, cellFormula: true, bookFiles: true
  });
  enrichWorkbookCellStyles(workbook);
  delete workbook.files;
  delete workbook.keys;
  return workbook;
}

function xmlAttributes(tag) {
  const attributes = {};
  for (const match of tag.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attributes[match[1]] = match[2] ?? match[3];
  return attributes;
}

function xmlPart(workbook, name) {
  const content = workbook.files?.[name]?.content;
  return content ? Buffer.from(content).toString("utf8") : "";
}

function resolveStyleColor(color, workbook) {
  if (!color) return color;
  let rgb = color.rgb;
  if (!rgb && Number.isInteger(Number(color.theme))) rgb = workbook.Themes?.themeElements?.clrScheme?.[Number(color.theme)]?.rgb;
  if (!rgb || !/^(?:[\da-f]{6}|[\da-f]{8})$/i.test(rgb)) return color;
  rgb = rgb.slice(-6);
  const tint = Number(color.tint || 0);
  if (tint) rgb = rgb.match(/../g).map((component) => {
    const value = parseInt(component, 16);
    const adjusted = tint > 0 ? value + (255 - value) * tint : value * (1 + tint);
    return Math.max(0, Math.min(255, Math.round(adjusted))).toString(16).padStart(2, "0");
  }).join("");
  return { ...color, rgb };
}

function enrichWorkbookCellStyles(workbook) {
  const styles = workbook.Styles;
  if (!styles?.CellXf || !workbook.files) return;
  const relations = new Map();
  for (const match of xmlPart(workbook, "xl/_rels/workbook.xml.rels").matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const attributes = xmlAttributes(match[0]);
    if (attributes.Target && attributes.Id && /\/worksheet$/.test(attributes.Type || "")) {
      const target = attributes.Target.replace(/\\/g, "/");
      relations.set(attributes.Id, target.startsWith("/") ? target.slice(1) : path.posix.normalize(`xl/${target}`));
    }
  }
  for (const [index, name] of (workbook.SheetNames || []).entries()) {
    const definition = workbook.Workbook?.Sheets?.[index];
    const file = relations.get(definition?.id) || `xl/worksheets/sheet${index + 1}.xml`;
    const sheet = workbook.Sheets[name];
    if (!sheet) continue;
    for (const match of xmlPart(workbook, file).matchAll(/<(?:\w+:)?c\b[^>]*>/g)) {
      const attributes = xmlAttributes(match[0]);
      const cell = sheet[attributes.r];
      if (!cell || attributes.s === undefined) continue;
      const xf = styles.CellXf[Number(attributes.s)];
      if (!xf) continue;
      const font = styles.Fonts?.[xf.fontId] || {};
      const fill = styles.Fills?.[xf.fillId] || {};
      cell.s = {
        ...(typeof cell.s === "object" ? cell.s : {}),
        font: { ...font, color: resolveStyleColor(font.color, workbook) },
        fill: { ...fill, fgColor: resolveStyleColor(fill.fgColor, workbook) },
        border: styles.Borders?.[xf.borderId] || {},
        alignment: xf.alignment || {}
      };
    }
  }
}

function cellHasContent(cell) {
  return Boolean(cell && (cell.f || (cell.v !== undefined && cell.v !== null && String(cell.v).trim()) || String(cell.w ?? "").trim()));
}

function* cells(sheet) {
  if (!sheet) return;
  if (Array.isArray(sheet)) {
    for (const rowKey of Object.keys(sheet)) {
      if (!/^\d+$/.test(rowKey) || !Array.isArray(sheet[rowKey])) continue;
      for (const colKey of Object.keys(sheet[rowKey])) {
        if (/^\d+$/.test(colKey)) yield { r: Number(rowKey), c: Number(colKey), cell: sheet[rowKey][colKey] };
      }
    }
  } else {
    for (const address of Object.keys(sheet)) {
      if (!/^[A-Z]+[1-9]\d*$/.test(address)) continue;
      yield { ...XLSX.utils.decode_cell(address), cell: sheet[address] };
    }
  }
}

function getCell(sheet, row, col) {
  return Array.isArray(sheet) ? sheet[row]?.[col] : sheet[XLSX.utils.encode_cell({ r: row, c: col })];
}

function inspectSpreadsheetWorkbook(workbook) {
  let populatedCells = 0;
  let formulaCells = 0;
  for (const name of workbook.SheetNames || []) {
    for (const { cell } of cells(workbook.Sheets[name])) {
      if (!cellHasContent(cell)) continue;
      populatedCells += 1;
      if (cell.f) formulaCells += 1;
    }
  }
  return { hasContent: populatedCells > 0, populatedCells, formulaCells, sheetCount: (workbook.SheetNames || []).length };
}

function formatCell(cell, diagnostics) {
  if (!cell) return "";
  // Preserve cached formula values, including zero and false. Uncalculated
  // formulas remain visible rather than becoming an apparently empty cell.
  if (cell.f && (cell.v === undefined || cell.v === null) && !String(cell.w ?? "")) {
    diagnostics.uncalculatedFormulaCount += 1;
    return `=${cell.f}`;
  }
  if (cell.w !== undefined) return String(cell.w);
  try { return String(XLSX.utils.format_cell(cell) ?? cell.v ?? ""); }
  catch (_) { return String(cell.v ?? ""); }
}

function getColumnWidth(sheet, col) {
  const definition = sheet["!cols"]?.[col] || {};
  const width = Number(definition.wpx) || (Number(definition.wch) * 7 + 5) || (Number(definition.width) * 7 + 5) || 84;
  return Math.min(280, Math.max(56, width));
}

function splitColumnBands(sheet, columns) {
  const bands = [];
  let band = [];
  let width = 0;
  for (const column of columns) {
    const nextWidth = getColumnWidth(sheet, column);
    if (band.length && width + nextWidth > PAGE_TABLE_WIDTH) { bands.push(band); band = []; width = 0; }
    band.push(column);
    width += nextWidth;
  }
  if (band.length) bands.push(band);
  return bands;
}

function getVisibleBounds(sheet) {
  let bounds = null;
  for (const { r, c, cell } of cells(sheet)) {
    if (!cellHasContent(cell) || sheet["!rows"]?.[r]?.hidden || sheet["!cols"]?.[c]?.hidden) continue;
    if (!bounds) bounds = { startRow: r, endRow: r, startCol: c, endCol: c };
    else { bounds.startRow = Math.min(bounds.startRow, r); bounds.endRow = Math.max(bounds.endRow, r); bounds.startCol = Math.min(bounds.startCol, c); bounds.endCol = Math.max(bounds.endCol, c); }
  }
  if (!bounds) return null;
  for (const merge of sheet["!merges"] || []) {
    if (!merge?.s || !merge?.e || !cellHasContent(getCell(sheet, merge.s.r, merge.s.c))) continue;
    if (merge.e.r < bounds.startRow || merge.s.r > bounds.endRow || merge.e.c < bounds.startCol || merge.s.c > bounds.endCol) continue;
    bounds.startRow = Math.min(bounds.startRow, merge.s.r); bounds.endRow = Math.max(bounds.endRow, merge.e.r);
    bounds.startCol = Math.min(bounds.startCol, merge.s.c); bounds.endCol = Math.max(bounds.endCol, merge.e.c);
  }
  return bounds;
}

function limitError(message) {
  const error = new Error(message);
  error.code = "SPREADSHEET_LIMIT";
  error.publicMessage = "엑셀 크기 제한 · 파일 분할 필요";
  return error;
}

function colorHex(color) {
  const raw = String(color?.rgb || "");
  const rgb = raw.length === 8 ? raw.slice(2) : raw;
  return /^[a-f\d]{6}$/i.test(rgb) ? `#${rgb}` : "";
}

function cellStyle(cell) {
  const style = cell?.s || {};
  const rules = [];
  const alignment = style.alignment || {};
  if (["left", "center", "right", "justify"].includes(alignment.horizontal)) rules.push(`text-align:${alignment.horizontal}`);
  else if (cell?.t === "n") rules.push("text-align:right");
  const vertical = { center: "middle", top: "top", bottom: "bottom" }[alignment.vertical];
  if (vertical) rules.push(`vertical-align:${vertical}`);
  if (style.font?.bold) rules.push("font-weight:700");
  if (style.font?.italic) rules.push("font-style:italic");
  if (style.font?.underline) rules.push("text-decoration:underline");
  if (Number(style.font?.sz) > 0) rules.push(`font-size:${Math.min(96, Number(style.font.sz))}pt`);
  const fontName = String(style.font?.name || "").replace(/[^\p{L}\p{N} _-]/gu, "");
  if (fontName) rules.push(`font-family:'${fontName}','Malgun Gothic',sans-serif`);
  const textColor = colorHex(style.font?.color);
  const background = colorHex(style.fill?.fgColor);
  if (textColor) rules.push(`color:${textColor}`);
  if (background) rules.push(`background:${background}`);
  for (const edge of ["top", "right", "bottom", "left"]) {
    const border = style.border?.[edge];
    if (!border?.style) continue;
    const line = /dash|dot/i.test(border.style) ? (/dot/i.test(border.style) ? "dotted" : "dashed") : "solid";
    const width = /medium|thick|double/i.test(border.style) ? 2 : 1;
    rules.push(`border-${edge}:${width}px ${line} ${colorHex(border.color) || "#64748b"}`);
  }
  return rules.join(";");
}

function buildBandTable(sheet, rows, columns, diagnostics) {
  const owners = new Map();
  const covered = new Set();
  for (const merge of sheet["!merges"] || []) {
    if (!merge?.s || !merge?.e) continue;
    const mergeRows = rows.filter((row) => row >= merge.s.r && row <= merge.e.r);
    const mergeCols = columns.filter((col) => col >= merge.s.c && col <= merge.e.c);
    if (!mergeRows.length || !mergeCols.length) continue;
    const key = `${mergeRows[0]}:${mergeCols[0]}`;
    owners.set(key, { rows: mergeRows.length, cols: mergeCols.length, cell: getCell(sheet, merge.s.r, merge.s.c) });
    for (const row of mergeRows) for (const col of mergeCols) if (`${row}:${col}` !== key) covered.add(`${row}:${col}`);
  }
  let html = "<table><colgroup><col class=\"row-number\" />";
  for (const col of columns) html += `<col style="width:${getColumnWidth(sheet, col)}px" />`;
  html += "</colgroup><thead><tr><th></th>" + columns.map((col) => `<th>${XLSX.utils.encode_col(col)}</th>`).join("") + "</tr></thead><tbody>";
  for (const row of rows) {
    const rowDefinition = sheet["!rows"]?.[row] || {};
    const rowHeight = Number(rowDefinition.hpx) || Number(rowDefinition.hpt) * 96 / 72;
    html += `<tr${rowHeight > 0 ? ` style="height:${rowHeight.toFixed(2)}px"` : ""}><th class="row-number">${row + 1}</th>`;
    for (const col of columns) {
      if (covered.has(`${row}:${col}`)) continue;
      const owner = owners.get(`${row}:${col}`);
      const cell = owner ? owner.cell : getCell(sheet, row, col);
      const spans = owner ? ` rowspan="${owner.rows}" colspan="${owner.cols}"` : "";
      html += `<td${spans} style="${cellStyle(cell)}">${escapeHtml(formatCell(cell, diagnostics))}</td>`;
    }
    html += "</tr>";
  }
  return html + "</tbody></table>";
}

function wrapHtml(title, sections) {
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>
    @page { size: A4 landscape; margin: 12mm; }
    html,body { margin:0;padding:0;font-family:'Malgun Gothic','Segoe UI',sans-serif;color:#17202b;font-size:11px; }
    h1 { font-size:15px;margin:0 0 5px; } h2 { font-size:12px;margin:0 0 9px;font-weight:500;color:#475569; }
    section { break-after:page;page-break-after:always; } section:last-child { break-after:auto;page-break-after:auto; }
    table { border-collapse:collapse;table-layout:fixed;width:auto;max-width:100%; }
    thead { display:table-header-group; } tr { break-inside:avoid;page-break-inside:avoid; }
    td,th { border:1px solid #d4dbe4;padding:5px 6px;vertical-align:top;white-space:pre-wrap;overflow-wrap:anywhere; }
    th { background:#f5f7fa;color:#64748b;font-size:10px;font-weight:500;text-align:center; }
    .row-number { width:38px;min-width:38px;white-space:nowrap; }
  </style></head><body>${sections.map((section) => `<section><h1>${escapeHtml(title)}</h1><h2>${escapeHtml(section.title)}</h2>${section.table}</section>`).join("")}</body></html>`;
}

function buildSpreadsheetCompatibility({ sourcePath = "", workbook, maxCells = MAX_RENDERED_CELLS } = {}) {
  const data = workbook || readSpreadsheetWorkbook(sourcePath);
  const diagnostics = { engineVersion: SPREADSHEET_RENDER_VERSION, renderedCellCount: 0, hiddenSheetCount: 0, emptySheetCount: 0, uncalculatedFormulaCount: 0, horizontalBands: 0 };
  const sheets = [];
  const allSections = [];
  const sheetNames = data.SheetNames || [];
  if (sheetNames.length > MAX_RENDERED_SHEETS) throw limitError("시트가 너무 많습니다. Excel에서 필요한 시트를 따로 저장해 주세요.");
  for (const [index, name] of sheetNames.entries()) {
    if (Number(data.Workbook?.Sheets?.[index]?.Hidden || 0) > 0) { diagnostics.hiddenSheetCount += 1; continue; }
    const sheet = data.Sheets[name];
    const bounds = getVisibleBounds(sheet);
    if (!bounds) { diagnostics.emptySheetCount += 1; continue; }
    const rowCount = bounds.endRow - bounds.startRow + 1;
    const colCount = bounds.endCol - bounds.startCol + 1;
    if (rowCount * colCount > maxCells || diagnostics.renderedCellCount + rowCount * colCount > maxCells) {
      throw limitError("표가 너무 큽니다. Excel에서 필요한 범위를 따로 저장해 주세요.");
    }
    const rows = Array.from({ length: rowCount }, (_, offset) => bounds.startRow + offset).filter((row) => !sheet["!rows"]?.[row]?.hidden);
    const columns = Array.from({ length: colCount }, (_, offset) => bounds.startCol + offset).filter((col) => !sheet["!cols"]?.[col]?.hidden);
    const bands = splitColumnBands(sheet, columns);
    const sections = bands.map((band) => ({
      title: bands.length > 1 ? `${name} · ${XLSX.utils.encode_col(band[0])}-${XLSX.utils.encode_col(band[band.length - 1])}` : name,
      table: buildBandTable(sheet, rows, band, diagnostics)
    }));
    diagnostics.renderedCellCount += rows.length * columns.length;
    diagnostics.horizontalBands += bands.length;
    allSections.push(...sections);
    sheets.push({ sheetName: name, html: wrapHtml(path.basename(sourcePath) || "Excel", sections), rowCount: rows.length, columnCount: columns.length, bandCount: bands.length });
  }
  const warnings = [];
  if (diagnostics.uncalculatedFormulaCount) warnings.push(`계산값 없는 수식 ${diagnostics.uncalculatedFormulaCount}개`);
  return {
    html: allSections.length ? wrapHtml(path.basename(sourcePath) || "Excel", allSections) : "",
    sheets,
    sheetMap: sheets.map(({ sheetName }) => ({ sheetName, startPage: null, endPage: null })),
    warningMessage: warnings.join(" · "),
    diagnostics
  };
}

async function renderSpreadsheetCompatibilityPdf(prepared, renderHtml, PDFDocument) {
  const output = await PDFDocument.create();
  const sheetMap = [];
  for (const sheet of prepared.sheets || []) {
    const bytes = await renderHtml(sheet.html, { landscape: true, printBackground: true, pageSize: "A4", preferCSSPageSize: true });
    const source = await PDFDocument.load(bytes);
    if (!source.getPageCount()) throw new Error("엑셀 시트의 PDF 페이지가 비어 있습니다.");
    const startPage = output.getPageCount() + 1;
    const copied = await output.copyPages(source, source.getPageIndices());
    for (const page of copied) output.addPage(page);
    sheetMap.push({ sheetName: sheet.sheetName, startPage, endPage: output.getPageCount() });
  }
  if (!output.getPageCount()) {
    const error = new Error("표시할 셀 내용이 없습니다.");
    error.code = "EMPTY_DOCUMENT";
    throw error;
  }
  return { pdfBuffer: Buffer.from(await output.save()), sheetMap, pageCount: output.getPageCount() };
}

module.exports = {
  SPREADSHEET_RENDER_VERSION, MAX_RENDERED_CELLS, readSpreadsheetWorkbook,
  inspectSpreadsheetWorkbook, buildSpreadsheetCompatibility, renderSpreadsheetCompatibilityPdf
};
