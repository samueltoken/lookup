const fs = require("node:fs/promises");
const path = require("node:path");
const os = require("node:os");
const { escapeHtml } = require("./hwp-json-compat");
const { normalizeHwpxPackageForRenderer } = require("./document-readiness");

// Character measurements use a real browser canvas in one short-lived renderer.
async function renderHwpWithRhwp(BrowserWindow, bytes, title = "", fontFaceCss = "") {
  const entryPath = require.resolve("@rhwp/core");
  const [runtimeSource, wasmBytes] = await Promise.all([
    fs.readFile(entryPath, "utf8"),
    fs.readFile(path.join(path.dirname(entryPath), "rhwp_bg.wasm"))
  ]);
  const documentBase64 = normalizeHwpxPackageForRenderer(bytes).toString("base64");
  const wasmBase64 = wasmBytes.toString("base64");
  const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
    ${fontFaceCss}
    @page { size:A4; margin:0; }
    html,body { margin:0; padding:0; background:#fff; }
    .lookup-rhwp-page { break-after:page; page-break-after:always; }
    .lookup-rhwp-page:last-child { break-after:auto; page-break-after:auto; }
    * { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  </style></head><body><script type="module">
    ${runtimeSource.replace(/<\/script/gi, "<\\/script")}
    let lookupDoc;
    try {
      const decodeBytes = value => Uint8Array.from(atob(value), char => char.charCodeAt(0));
      const context = document.createElement('canvas').getContext('2d');
      globalThis.measureTextWidth = (font, text) => { context.font = font; return context.measureText(text).width; };
      await document.fonts.ready;
      initSync({module:decodeBytes('${wasmBase64}')});
      lookupDoc = new HwpDocument(decodeBytes('${documentBase64}'));
      const count = lookupDoc.pageCount();
      if (!count) throw new Error('rhwp_empty_document');
      const pageGeometry=[];
      for (let index=0; index<count; index++) {
        const page = document.createElement('section');
        page.className = 'lookup-rhwp-page';
        page.innerHTML = lookupDoc.renderPageHtml(index);
        if (!page.innerHTML.trim()) throw new Error('rhwp_page_empty');
        document.body.appendChild(page);
        if (index===0) pageGeometry.push({info:lookupDoc.getPageInfo(index),html:page.innerHTML.slice(0,180),width:page.firstElementChild?.getBoundingClientRect().width,height:page.firstElementChild?.getBoundingClientRect().height});
      }
      await document.fonts.ready;
      if (!document.body.innerText.trim() && !document.querySelector('.lookup-rhwp-page img, .lookup-rhwp-page image')) throw new Error('rhwp_no_visible_content');
      window.lookupHwpRenderState = {ok:true,pageCount:count,pageGeometry};
    } catch(error) {
      window.lookupHwpRenderState = {ok:false,reason:String(error.message || error)};
    } finally { try { lookupDoc?.free(); } catch(_) {} }
  </script></body></html>`;
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "lookup-hwp-rhwp-"));
  const rendererPath = path.join(directory, "renderer.html");
  let window;
  let timeout;
  try {
    window = new BrowserWindow({
      show: false, width: 1100, height: 1500, autoHideMenuBar: true,
      webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, offscreen: true }
    });
    window.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    await fs.writeFile(rendererPath, html, "utf8");
    await Promise.race([
      window.loadFile(rendererPath),
      new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error("rhwp_renderer_timeout")), 30000); })
    ]);
    const state = await window.webContents.executeJavaScript("window.lookupHwpRenderState || {ok:false,reason:'rhwp_renderer_not_ready'}");
    if (!state.ok) throw new Error(state.reason || "rhwp_renderer_failed");
    await window.webContents.executeJavaScript("Promise.all([document.fonts.ready,...Array.from(document.images).map(image=>image.complete ? Promise.resolve() : image.decode().catch(()=>{}))])");
    const pdfBuffer = await window.webContents.printToPDF({ printBackground: true, pageSize: "A4", preferCSSPageSize: true });
    return { ok: true, pdfBuffer, pageCount: state.pageCount, pageGeometry:state.pageGeometry, selectedHtmlSource: "rhwp-page-html" };
  } finally {
    clearTimeout(timeout);
    if (window && !window.isDestroyed()) window.destroy();
    await fs.unlink(rendererPath).catch(() => {});
    await fs.rmdir(directory).catch(() => {});
  }
}

module.exports = { renderHwpWithRhwp };
