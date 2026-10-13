// Run with: .\node_modules\electron\dist\electron.exe scripts\verify-web-preparation.cjs <output-directory>
// This exercises Chromium layout, lazy loading and actual printed PDF text.
const { app, BrowserWindow } = require("electron");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const { PDFDocument } = require("pdf-lib");
const {
  prepareWebPageForPdf, useScreenMediaForPdf, inspectCapturedImage, assessWebPageContent, waitForWebAccess
} = require("../src/main/web/page-preparation");

const output = path.resolve(process.argv[2] || "tmp/web-preparation-check");
app.setPath("userData", path.join(output, "profile"));
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="720" height="3200"><rect width="720" height="3200" fill="#eee"/><path d="M0 0L720 3200M720 0L0 3200" stroke="#222" stroke-width="25"/><text x="70" y="3140" font-size="50">LAST_IMAGE_STRIP</text></svg>`;
const imageUrl = `data:image/svg+xml,${encodeURIComponent(svg)}`;
const frame = `<html><body><h2>FRAME_START_MARKER</h2><div style="height:1500px;background:linear-gradient(white,#aac)">FRAME_CONTENT</div><p>FRAME_END_MARKER</p></body></html>`;
const fixture = `<!doctype html><html><head><style>
  body{margin:0;font-family:Arial;width:720px}#inner{width:700px;height:200px;overflow:auto}
  #screen-only{padding:20px}@media print{#screen-only{display:none}}
  .lazy-space{height:1900px} .row{height:300px} iframe{width:700px;height:200px;border:0}
  #tiny-crop{width:80px;height:30px;overflow:hidden}
</style></head><body>
<h1 id="screen-only">SCREEN_ONLY_MARKER</h1>
<div id="tiny-crop"><div style="height:80px">DECORATIVE_CROP</div></div>
<div id="inner"><div class="row">SCROLL_START_MARKER</div><div class="row">SCROLL_MIDDLE_MARKER</div><div class="row">SCROLL_END_MARKER</div></div>
<iframe srcdoc="${frame.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"></iframe>
<div class="lazy-space">BEFORE_LAZY_CONTENT</div><img id="lazy" style="width:700px" loading="lazy" src="data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=" data-src="${imageUrl.replace(/"/g, "&quot;")}"/>
<p>DOCUMENT_END_MARKER</p><p><br></p><p><br></p></body></html>`;

app.whenReady().then(async () => {
  const win = new BrowserWindow({ show: false, width: 1040, height: 1000, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, offscreen: true } });
  let resetMedia;
  try {
    await fs.mkdir(output, { recursive: true });
    await win.loadURL(`data:text/html,${encodeURIComponent(fixture)}`);
    resetMedia = await useScreenMediaForPdf(win.webContents);
    const diagnostics = await prepareWebPageForPdf(win.webContents, { timeoutMs: 18000 });
    assert.equal(diagnostics.lazyLoad.reachedBottom, true, "walk reaches the dynamically measured bottom");
    assert.equal(diagnostics.normalization.splitImages, 1, "long image is divided into readable strips");
    assert.ok(diagnostics.normalization.imageSlices >= 3);
    assert.equal(diagnostics.normalization.trimmedTrailingBlanks, 2, "empty trailing paragraphs cannot create an extra white sheet");
    assert.equal(diagnostics.after.pendingImageCount, 0);
    assert.equal(diagnostics.after.failedImageCount, 0);
    assert.ok(diagnostics.frames.some((frame) => frame.after?.textLength > 20), "embedded document is prepared");
    const layout = await win.webContents.executeJavaScript(`({innerHeight:document.getElementById('inner').clientHeight, cropHeight:document.getElementById('tiny-crop').clientHeight, iframeHeight:document.querySelector('iframe').clientHeight,screenVisible:getComputedStyle(document.getElementById('screen-only')).display!=='none',media:matchMedia('screen').matches})`);
    assert.ok(layout.innerHeight >= 900, "internal scroller no longer clips its ending");
    assert.equal(layout.cropHeight, 30, "intentional small crop keeps its shape");
    assert.ok(layout.iframeHeight >= 1500, "iframe contains its full document height");
    assert.equal(layout.screenVisible, true);
    assert.equal(layout.media, true);
    assert.equal(assessWebPageContent(diagnostics).hasContent, true);
    const pdf = await win.webContents.printToPDF({ pageSize: "A4", printBackground: true, preferCSSPageSize: true });
    const parsed = await PDFDocument.load(pdf);
    assert.ok(parsed.getPageCount() >= 5, "actual PDF spans internal, embedded and long-image content");
    await fs.writeFile(path.join(output, "fixture.pdf"), pdf);
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const pdfHandle = await pdfjs.getDocument({ data: new Uint8Array(pdf), disableFontFace: true }).promise;
    let text = "";
    for (let page = 1; page <= pdfHandle.numPages; page += 1) text += (await (await pdfHandle.getPage(page)).getTextContent()).items.map((item) => item.str).join(" ") + "\n";
    for (const marker of ["SCREEN_ONLY_MARKER", "SCROLL_END_MARKER", "FRAME_END_MARKER", "DOCUMENT_END_MARKER"]) assert.ok(text.includes(marker), `printed PDF retains ${marker}`);
    await pdfHandle.destroy();
    const capture = await win.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true });
    assert.equal(inspectCapturedImage(capture).valid, true);
    await fs.writeFile(path.join(output, "diagnostics.json"), JSON.stringify({ diagnostics, layout, pageCount: parsed.getPageCount(), text }, null, 2));
    await win.loadURL("data:text/html,<html><head><title>Just a moment...</title></head><body>Access check</body></html>");
    await assert.rejects(waitForWebAccess(win.webContents, { timeoutMs: 500 }), (error) => error.code === "WEB_ACCESS_CHALLENGE", "access-check page cannot be saved as requested content");
    await win.loadURL("data:text/html,<html><body></body></html>");
    const empty = await prepareWebPageForPdf(win.webContents, { timeoutMs: 5000 });
    assert.equal(assessWebPageContent(empty).hasContent, false, "empty page does not count as source content");
    await win.loadURL("data:text/html,<html><body><p>CONTENT_WITH_MISSING_IMAGE</p><img src='data:image/png;base64,broken'></body></html>");
    const missingImage = await prepareWebPageForPdf(win.webContents, { timeoutMs: 5000 });
    assert.equal(assessWebPageContent(missingImage).incompleteImages, 1, "zero-height failed image remains in the missing-resource report");
    console.log(JSON.stringify({ ok: true, pageCount: parsed.getPageCount(), splitImages: diagnostics.normalization.splitImages, imageSlices: diagnostics.normalization.imageSlices, markers: 4, output }));
  } catch (error) { console.error(error.stack || error); process.exitCode = 1; }
  finally {
    if (resetMedia) await resetMedia().catch(() => {});
    if (!win.isDestroyed()) win.destroy();
    app.exit(process.exitCode || 0);
  }
});
