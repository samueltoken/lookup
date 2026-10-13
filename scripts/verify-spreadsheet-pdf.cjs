// Disposable integration profile: simulate unavailable external engines and
// exercise the real spreadsheet fallback, Chromium printing and PDF merge.
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const assert = require("node:assert/strict");
const { createRequire } = require("node:module");
const root = path.resolve(__dirname, "..");
const runtime = createRequire(path.join(process.env.USERPROFILE, ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/package.json"));
const { _electron: electron } = runtime("playwright");
const temp = path.join(root, "tmp", "spreadsheet-verification");
fs.mkdirSync(temp, { recursive: true });
const fixture = path.join(root, "tmp", "spreadsheet-fixtures", "한글 공백 Excel 호환성.xlsx");
const hashBefore = crypto.createHash("sha256").update(fs.readFileSync(fixture)).digest("hex");
const bootstrap = path.join(temp, "bootstrap.cjs");
fs.writeFileSync(bootstrap, `
const { app } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
app.setPath('userData', ${JSON.stringify(path.join(temp, "profile"))});
app.whenReady().then(() => {
  const filename = ${JSON.stringify(path.join(root, "src/main/app-main.js"))};
  let source = fs.readFileSync(filename, 'utf8');
  source = source.slice(0, source.lastIndexOf('\\nif (gotSingleInstanceLock) {'));
  source = source.replace('app.requestSingleInstanceLock()', 'true');
  source = source.replace('mainWindow?.show();', '');
  source = source.replace('await convertSpreadsheetWithLibreOffice(sourcePath, librePdfPath, findLibreOfficeExecutable())', '({ok:false,missingEngine:true,stderr:"test_engine_unavailable"})');
  source += '\\nconvertWithExcelCom = async () => ({ok:false,stderr:"test_engine_unavailable"});\\nmodule.exports = { convertDocumentToPdf, createWindow };';
  const module = new Module(filename);
  module.filename = filename;
  module.paths = Module._nodeModulePaths(path.dirname(filename));
  module._compile(source, filename);
  global.lookupSpreadsheetTest = module.exports;
  global.lookupTestFs = fs;
  global.lookupSpreadsheetTest.createWindow();
});
`, "utf8");

(async () => {
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  const app = await electron.launch({ executablePath: path.join(root, "node_modules/electron/dist/electron.exe"), args: [bootstrap], env, timeout: 30000 });
  try {
    await app.firstWindow();
    await app.evaluate(async () => {
      while (!global.lookupSpreadsheetTest) await new Promise((resolve) => setTimeout(resolve, 30));
    });
    const results = await app.evaluate(async (_electron, { fixture, output }) => {
      const first = await global.lookupSpreadsheetTest.convertDocumentToPdf(fixture);
      const second = await global.lookupSpreadsheetTest.convertDocumentToPdf(fixture);
      global.lookupTestFs.writeFileSync(output, first.pdfBuffer);
      return {
        first: { mode: first.convertMode, sheetMap: first.sheetMap, diagnostics: first.convertDiagnostics, warning: first.warningMessage, bytes: first.pdfBuffer.length },
        cache: { mode: second.convertMode, sheetMap: second.sheetMap, bytes: second.pdfBuffer.length }
      };
    }, { fixture, output: path.join(temp, "compatibility-fixture.pdf") });
    assert.equal(results.first.mode, "xlsx-html");
    assert.equal(results.first.sheetMap.length, 2);
    assert.deepEqual(results.first.sheetMap, results.cache.sheetMap);
    assert.ok(results.first.bytes > 2000);
    assert.equal(crypto.createHash("sha256").update(fs.readFileSync(fixture)).digest("hex"), hashBefore);
    fs.writeFileSync(path.join(temp, "result.json"), JSON.stringify(results, null, 2));
    console.log(JSON.stringify(results, null, 2));
  } finally { await app.close(); }
})().catch((error) => { console.error(error); process.exitCode = 1; });
