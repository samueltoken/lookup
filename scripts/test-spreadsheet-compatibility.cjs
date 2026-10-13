const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const XLSX = require("xlsx");
const { PDFDocument } = require("pdf-lib");
const {
  readSpreadsheetWorkbook, inspectSpreadsheetWorkbook,
  buildSpreadsheetCompatibility, renderSpreadsheetCompatibilityPdf
} = require("../src/main/spreadsheet/compatibility");
const { buildExcelReadOnlyScript, buildOwnedExcelCleanupScript, convertSpreadsheetWithLibreOffice } = require("../src/main/spreadsheet/native-engine");

function workbook(sheets, metadata) {
  return { SheetNames: Object.keys(sheets), Sheets: sheets, Workbook: { Sheets: metadata || Object.keys(sheets).map(() => ({ Hidden: 0 })) } };
}

test("dense and sparse worksheets both count real cells including zero and false", () => {
  const rows = [[{ t: "n", v: 0 }, { t: "b", v: false }, { t: "n", f: "SUM(A1:A2)" }]];
  const sparse = { A1: rows[0][0], B1: rows[0][1], C1: rows[0][2], "!ref": "A1:C1" };
  for (const sheet of [rows, sparse]) {
    const result = inspectSpreadsheetWorkbook(workbook({ Main: sheet }));
    assert.equal(result.hasContent, true);
    assert.equal(result.populatedCells, 3);
    assert.equal(result.formulaCells, 1);
  }
});

test("a formatted but empty worksheet remains empty", () => {
  const result = inspectSpreadsheetWorkbook(workbook({ Empty: { A1: { s: { font: { bold: true } } }, "!ref": "A1:XFD1048576" } }));
  assert.equal(result.hasContent, false);
});

test("all rows and columns survive beyond the old 300 by 32 limit", () => {
  const sheet = XLSX.utils.aoa_to_sheet(Array.from({ length: 450 }, (_, r) => Array.from({ length: 40 }, (_, c) => `R${r + 1}C${c + 1}`)));
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: sheet }) });
  assert.ok(result.html.includes("R450C40"));
  assert.ok(result.html.includes("R301C33"));
  assert.equal(result.sheets[0].rowCount, 450);
  assert.equal(result.sheets[0].columnCount, 40);
  assert.ok(result.sheets[0].bandCount > 1);
});

test("cached numeric formats, false and zero are preserved", () => {
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: {
    A1: { t: "n", v: 12345.5, w: "12,345.50", z: "#,##0.00" },
    B1: { t: "n", v: 0.125, w: "12.5%", z: "0.0%" },
    C1: { t: "n", f: "A2-A2", v: 0, w: "0" },
    D1: { t: "b", v: false }, "!ref": "A1:D1"
  } }) });
  for (const text of ["12,345.50", "12.5%", ">0</td>", "FALSE"]) assert.ok(result.html.includes(text), text);
  assert.equal(result.diagnostics.uncalculatedFormulaCount, 0);
});

test("formulas with no cached result stay visible with a short warning", () => {
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: { A1: { t: "n", f: "SUM(B1:B2)" }, "!ref": "A1" } }) });
  assert.ok(result.html.includes("=SUM(B1:B2)"));
  assert.equal(result.diagnostics.uncalculatedFormulaCount, 1);
  assert.ok(result.warningMessage.includes("1개"));
});

test("hidden sheets, rows and columns retain their source visibility", () => {
  const visible = XLSX.utils.aoa_to_sheet([["visible", "hidden-column"], ["hidden-row", "hidden-both"], ["end", "hidden-end"]]);
  visible["!rows"] = [{}, { hidden: true }];
  visible["!cols"] = [{}, { hidden: true }];
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Visible: visible, Hidden: XLSX.utils.aoa_to_sheet([["hidden-sheet"]]), Empty: {} }, [{}, { Hidden: 1 }, {}]) });
  assert.ok(result.html.includes("visible"));
  for (const text of ["hidden-column", "hidden-row", "hidden-sheet"]) assert.equal(result.html.includes(text), false);
  assert.deepEqual(result.sheets.map((sheet) => sheet.sheetName), ["Visible"]);
  assert.equal(result.diagnostics.hiddenSheetCount, 1);
  assert.equal(result.diagnostics.emptySheetCount, 1);
});

test("merged text remains present on every horizontal band it spans", () => {
  const sheet = { A1: { t: "s", v: "MERGED-TITLE" }, AN3: { t: "s", v: "END" }, "!ref": "A1:AN3", "!merges": [{ s: { r: 0, c: 0 }, e: { r: 1, c: 39 } }] };
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: sheet }) });
  assert.equal((result.html.match(/MERGED-TITLE/g) || []).length, result.sheets[0].bandCount);
  assert.ok(result.html.includes('rowspan="2"'));
  assert.ok(result.html.includes("END"));
});

test("font, alignment, fill, border, width and height are rendered", () => {
  const sheet = { A1: { t: "s", v: "styled", s: {
    font: { name: "Malgun Gothic", sz: 15, bold: true, italic: true, color: { rgb: "FF112233" } },
    alignment: { horizontal: "center", vertical: "center" }, fill: { fgColor: { rgb: "FFEEDD" } },
    border: { bottom: { style: "medium", color: { rgb: "112233" } } }
  } }, "!cols": [{ wpx: 140 }], "!rows": [{ hpx: 40 }], "!ref": "A1" };
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: sheet }) });
  for (const value of ["font-size:15pt", "text-align:center", "vertical-align:middle", "background:#FFEEDD", "color:#112233", "width:140px", "height:40.00px", "border-bottom:2px"]) assert.ok(result.html.includes(value), value);
});

test("a stale enormous used range does not create a million blank rows", () => {
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: { A1: { t: "s", v: "only-cell" }, "!ref": "A1:XFD1048576" } }) });
  assert.equal(result.sheets[0].rowCount, 1);
  assert.equal(result.sheets[0].columnCount, 1);
});

test("real oversized content fails explicitly instead of truncating", () => {
  assert.throws(() => buildSpreadsheetCompatibility({ workbook: workbook({ Main: { A1: { t: "s", v: "START" }, XFD1048576: { t: "s", v: "END" }, "!ref": "A1:XFD1048576" } }) }), (error) => error.code === "SPREADSHEET_LIMIT");
});

test("cell strings are escaped rather than executed as HTML", () => {
  const result = buildSpreadsheetCompatibility({ workbook: workbook({ Main: { A1: { t: "s", v: '<img src=x onerror="alert(1)">' }, "!ref": "A1" } }) });
  assert.ok(result.html.includes("&lt;img"));
  assert.equal(result.html.includes("<img"), false);
});

test("sheet navigation uses actual PDF page counts, not row guesses", async () => {
  const prepared = { sheets: [{ sheetName: "First", html: "first" }, { sheetName: "Second", html: "second" }] };
  const rendered = await renderSpreadsheetCompatibilityPdf(prepared, async (html) => {
    const document = await PDFDocument.create();
    for (let i = 0; i < (html === "first" ? 2 : 3); i += 1) document.addPage();
    return document.save();
  }, PDFDocument);
  assert.deepEqual(rendered.sheetMap, [{ sheetName: "First", startPage: 1, endPage: 2 }, { sheetName: "Second", startPage: 3, endPage: 5 }]);
  assert.equal((await PDFDocument.load(rendered.pdfBuffer)).getPageCount(), 5);
});

test("native Excel script uses a read-only workbook and disables macros and links", () => {
  const script = buildExcelReadOnlyScript("C:\\한글 공백\\O'Brien.xlsx", "C:\\출력\\view.pdf");
  assert.ok(script.includes("O''Brien.xlsx"));
  assert.ok(script.includes("AutomationSecurity = 3"));
  assert.ok(script.includes("AskToUpdateLinks = $false"));
  assert.ok(script.includes("0, $true"));
  assert.ok(script.includes("Close($false)"));
  assert.equal(/\.Save\(/.test(script), false);
});

test("Excel timeout cleanup is restricted to our newly created process ID", () => {
  assert.equal(buildOwnedExcelCleanupScript(0, Date.now()), "");
  const script = buildOwnedExcelCleanupScript(12345, Date.now());
  assert.ok(script.includes("Get-Process -Id 12345"));
  assert.ok(script.includes("StartTime.ToUniversalTime() -ge"));
  assert.ok(script.includes("Stop-Process -Id 12345"));
  assert.equal(script.includes("Stop-Process -Name"), false);
});

test("a missing LibreOffice engine returns promptly", async () => {
  const result = await convertSpreadsheetWithLibreOffice("source.xlsx", "output.pdf", {});
  assert.equal(result.ok, false);
  assert.equal(result.missingEngine, true);
});

const fixture = path.join(__dirname, "../tmp/spreadsheet-fixtures/한글 공백 Excel 호환성.xlsx");
test("real XLSX fixture preserves source bytes, last cell, font styles and formats", { skip: !fs.existsSync(fixture) }, () => {
  const before = crypto.createHash("sha256").update(fs.readFileSync(fixture)).digest("hex");
  const data = readSpreadsheetWorkbook(fixture);
  const result = buildSpreadsheetCompatibility({ sourcePath: fixture, workbook: data });
  assert.equal(inspectSpreadsheetWorkbook(data).hasContent, true);
  assert.ok(result.html.includes("FINAL-CELL-AN450"));
  assert.ok(result.html.includes("SECOND-SHEET-END"));
  assert.ok(result.html.includes("12,345.50"));
  assert.ok(result.html.includes("font-size:15pt"));
  assert.ok(result.html.includes("background:#DBEAFE"));
  assert.equal(result.sheets.length, 2);
  assert.equal(crypto.createHash("sha256").update(fs.readFileSync(fixture)).digest("hex"), before);
});
