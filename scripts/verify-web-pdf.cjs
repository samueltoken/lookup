// Run the real Electron renderer and PDF path in an isolated profile.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const root = path.resolve(__dirname, '..');
const runtimeRequire = createRequire(path.join(process.env.USERPROFILE, '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/package.json'));
const { _electron: electron } = runtimeRequire('playwright');
const tmp = path.join(root, 'tmp', 'web-verification');
fs.mkdirSync(tmp, { recursive: true });
const bootstrap = path.join(tmp, 'bootstrap.cjs');
fs.writeFileSync(bootstrap, `
const { app } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
global.lookupFs = fs;
app.setPath('userData', ${JSON.stringify(path.join(tmp, 'profile'))});
app.whenReady().then(() => {
  const filename = ${JSON.stringify(path.join(root, 'src/main/app-main.js'))};
  let source = fs.readFileSync(filename, 'utf8');
  source = source.slice(0, source.lastIndexOf('\\nif (gotSingleInstanceLock) {'));
  source = source.replace('app.requestSingleInstanceLock()', 'true').replace('mainWindow?.show();', '');
  source = source.replace('webviewTag: false', 'webviewTag: false, backgroundThrottling: false, offscreen: true');
  source += '\\nmodule.exports = { renderHtmlToPdfBuffer, openWebDocument, convertDocumentToPdf, createWindow, ipcMain, PDFDocument };';
  const module = new Module(filename);
  module.filename = filename;
  module.paths = Module._nodeModulePaths(path.dirname(filename));
  module._compile(source, filename);
  global.lookupTest = module.exports;
  global.lookupTest.createWindow();
});
`, 'utf8');

async function run() {
  let localServer;
  let localUrl;
  if (process.argv.includes('--web-ui-local')) {
    localServer = require('node:http').createServer((_request, response) => {
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      response.end(`<html><head><title>웹페이지 UI 검증</title></head><body><h1>LOCAL-WEB-TITLE</h1>${Array.from({length: 100}, (_, i) => `<p style="height:45px">ROW-${i + 1}</p>`).join('')}<p>LOCAL-WEB-END</p></body></html>`);
    });
    await new Promise(resolve => localServer.listen(0, '127.0.0.1', resolve));
    localUrl = `http://127.0.0.1:${localServer.address().port}`;
  }
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  const app = await electron.launch({ executablePath: path.join(root, 'node_modules/electron/dist/electron.exe'), args: [bootstrap], env, timeout: 30000 });
  try {
    const page = await app.firstWindow();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.waitForLoadState('domcontentloaded');
    await page.waitForFunction(() => !!window.lookupAPI && !document.querySelector('#currentVersionLabel')?.textContent.includes('-'), null, { timeout: 15000 });
    await app.evaluate(async ({ BrowserWindow }, outputPath) => {
      const image = await BrowserWindow.getAllWindows()[0].webContents.capturePage(undefined, { stayHidden: true, stayAwake: true });
      global.lookupFs.writeFileSync(outputPath, image.toPNG());
    }, path.join(tmp, 'viewer-empty.png'));
    if (process.argv.includes('--web-ui') || localUrl) {
      await page.locator('#openWebPageBtn').click();
      await page.locator('#webPageUrlInput').fill(localUrl || process.env.LOOKUP_WEB_REGRESSION_URL || 'https://example.org/');
      await page.locator('#webPageOpenBtn').click();
      await page.waitForFunction(() => {
        const count = Number(document.querySelector('#pageCountLabel').textContent.replace(/[^\d]/g, ''));
        return (count >= 3 && document.querySelector('#convertOverlay').classList.contains('hidden'))
          || document.querySelector('#convertOverlayTitle').textContent.includes('실패');
      }, null, { timeout: 90000 });
      const result = await page.evaluate(() => ({
        status: document.querySelector('#statusText').textContent,
        warning: document.querySelector('#documentWarningBanner').textContent,
        pages: document.querySelector('#pageCountLabel').textContent
      }));
      assert.ok(Number(result.pages.replace(/[^\d]/g, '')) >= (localUrl ? 3 : 20), JSON.stringify(result));
      await app.evaluate(async ({ BrowserWindow }, outputPath) => {
        const image = await BrowserWindow.getAllWindows()[0].webContents.capturePage(undefined, { stayHidden: true, stayAwake: true });
        global.lookupFs.writeFileSync(outputPath, image.toPNG());
      }, path.join(tmp, localUrl ? 'viewer-web-local.png' : 'viewer-gmarket.png'));
      console.log(JSON.stringify({ webUiChecks: 'passed', result }));
    }
    if (process.argv.includes('--viewer-check')) {
      const fixturePath = path.join(root, 'tmp', 'web-preparation-check', 'fixture.pdf');
      assert.ok(fs.existsSync(fixturePath));
      await app.evaluate(({ BrowserWindow }, fixturePath) => {
        BrowserWindow.getAllWindows()[0].webContents.send('system-open-file', fixturePath);
      }, fixturePath);
      await page.waitForFunction(() => document.querySelector('#pageCountLabel').textContent.includes('10'), null, { timeout: 20000 });
      await page.locator('#tabViewBtn').click();
      await page.locator('#viewFocusBtn').click();
      await page.locator('#tabHomeBtn').click();
      for (let index = 0; index < 7; index++) await page.locator('#zoomInBtn').click();
      await page.waitForTimeout(600);
      const panel = page.locator('#viewerPanel');
      const box = await panel.boundingBox();
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.evaluate(() => { document.querySelector('#viewerPanel').scrollTop = 30; });
      await page.mouse.wheel(0, 450);
      await page.waitForTimeout(300);
      const after = await page.evaluate(() => ({
        top: document.querySelector('#viewerPanel').scrollTop,
        current: document.querySelector('#pageInput').value,
        visiblePages: document.querySelectorAll('.page-wrap:not(.hidden-page)').length,
        zoom: document.querySelector('#zoomLabel').textContent
      }));
      assert.ok(after.top > 30, 'wheel scrolls inside the magnified page');
      assert.equal(after.current, '1', 'wheel keeps the current page selected');
      assert.equal(after.visiblePages, 1, 'focus view shows only the current page');
      await page.evaluate(() => { const panel = document.querySelector('#viewerPanel'); panel.scrollTop = panel.scrollHeight; });
      await page.mouse.wheel(0, 900);
      await page.waitForTimeout(200);
      assert.equal(await page.locator('#pageInput').inputValue(), '1', 'wheel at page bottom does not skip ahead');
      await page.locator('#nextPageBtn').click();
      await page.waitForTimeout(200);
      assert.equal(await page.locator('#pageInput').inputValue(), '2', 'explicit next button changes page');
      await page.locator('#tabViewBtn').click();
      await page.locator('#viewSingleBtn').click();
      await page.locator('#tabHomeBtn').click();
      await page.locator('#pageInput').fill('1');
      await page.locator('#pageInput').press('Enter');
      await page.waitForTimeout(200);
      await page.evaluate(() => {
        const panel = document.querySelector('#viewerPanel');
        const wrap = document.querySelector('.page-wrap');
        panel.scrollTop = wrap.offsetTop + wrap.offsetHeight - panel.clientHeight;
      });
      await page.mouse.wheel(0, 3000);
      await page.waitForTimeout(200);
      assert.equal(await page.locator('#pageInput').inputValue(), '1', 'continuous layout also keeps zoomed wheel on selected page');
      const clippedLeft = await page.evaluate(() => {
        const panel = document.querySelector('#viewerPanel');
        panel.scrollLeft = 0;
        const wrap = document.querySelector('.page-wrap:not(.hidden-page)');
        return wrap.getBoundingClientRect().left - panel.getBoundingClientRect().left;
      });
      assert.ok(clippedLeft >= 0, 'magnified page left edge remains reachable');
      await app.evaluate(async ({ BrowserWindow }, outputPath) => {
        const image = await BrowserWindow.getAllWindows()[0].webContents.capturePage(undefined, { stayHidden: true, stayAwake: true });
        global.lookupFs.writeFileSync(outputPath, image.toPNG());
      }, path.join(tmp, 'viewer-zoom.png'));
      await page.locator('#openWebPageBtn').click();
      await page.waitForFunction(() => !document.querySelector('#webPageModal').classList.contains('hidden'));
      await page.keyboard.press('Escape');
      assert.ok(await page.locator('#webPageModal').evaluate(el => el.classList.contains('hidden')));
      await page.keyboard.press('Control+f');
      assert.equal(await page.locator('#searchInput').evaluate(el => el === document.activeElement), true);
      await page.locator('#tabToolsBtn').click();
      await page.locator('#editTextBtn').click();
      const memoBox = await panel.boundingBox();
      await page.mouse.click(memoBox.x + memoBox.width / 2, memoBox.y + memoBox.height / 2);
      await page.waitForFunction(() => !document.querySelector('#textMemoEditor').classList.contains('hidden'));
      await page.locator('#textMemoInput').fill('한글 메모 저장 검증');
      await page.locator('#textMemoAddBtn').click();
      const saveOutput = path.join(tmp, 'saved-viewer.pdf');
      await app.evaluate((_electron, saveOutput) => {
        global.lookupTest.ipcMain.removeHandler('dialog:save-pdf');
        global.lookupTest.ipcMain.handle('dialog:save-pdf', () => saveOutput);
      }, saveOutput);
      await page.locator('#saveAsBtn').click();
      try {
        await page.waitForFunction(() => document.querySelector('#statusText').textContent.includes('저장 완료'), null, { timeout: 15000 });
      } catch (error) {
        console.log(JSON.stringify({ saveFailure: await page.locator('#statusText').textContent(), overlay: await page.locator('#convertOverlayMessage').textContent(), errors }));
        throw error;
      }
      const savedPages = await app.evaluate(async (_electron, saveOutput) => {
        const doc = await global.lookupTest.PDFDocument.load(global.lookupFs.readFileSync(saveOutput));
        return doc.getPageCount();
      }, saveOutput);
      assert.equal(savedPages, 10, 'Save As produces a valid PDF retaining all pages');
      console.log(JSON.stringify({ viewerChecks: 'passed', scroll: after, leftEdge: clippedLeft }));
    }
    if (process.argv.includes('--inspect-live')) {
      const sourceUrl = process.env.LOOKUP_WEB_REGRESSION_URL;
      if (!sourceUrl || !/^https?:\/\//i.test(sourceUrl)) throw new Error('Set LOOKUP_WEB_REGRESSION_URL to an HTTP(S) URL for live inspection.');
      const live = await app.evaluate(async ({ BrowserWindow }, sourceUrl) => {
        const window = new BrowserWindow({ show: false, width: 1280, height: 1000, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, offscreen: true } });
        global.lookupLiveWindow = window;
        await window.loadURL(sourceUrl);
        await new Promise(resolve => setTimeout(resolve, 2500));
        return {
          title: window.webContents.getTitle(),
          frames: window.webContents.mainFrame.framesInSubtree.map(frame => ({ url: frame.url, name: frame.name })),
          dom: await window.webContents.executeJavaScript(`({height: document.body.scrollHeight, images:document.images.length, iframe: [...document.querySelectorAll('iframe')].map(el=>({id:el.id,src:el.src,height:el.clientHeight})),buttons:[...document.querySelectorAll('button,a')].filter(el=>el.textContent.includes('상세정보 더보기')).map(el=>({tag:el.tagName,id:el.id,cls:el.className,href:el.getAttribute('href')}))})`)
        };
      }, sourceUrl);
      console.log(JSON.stringify({ live }, null, 2));
      await app.evaluate(async (_electron, outputPath) => {
        const image = await global.lookupLiveWindow.webContents.capturePage();
        global.lookupFs.writeFileSync(outputPath, image.toPNG());
        global.lookupLiveWindow.destroy();
      }, path.join(tmp, 'web-screen.png'));
      return;
    }
    if (process.argv.includes('--live-pdf')) {
      const sourceUrl = process.env.LOOKUP_WEB_REGRESSION_URL;
      if (!sourceUrl || !/^https?:\/\//i.test(sourceUrl)) throw new Error('Set LOOKUP_WEB_REGRESSION_URL to an HTTP(S) URL for live capture.');
      const output = path.join(root, 'output', 'pdf', 'web-capture.pdf');
      fs.mkdirSync(path.dirname(output), { recursive: true });
      const result = await app.evaluate(async (_electron, { output, sourceUrl }) => {
        const result = await global.lookupTest.renderHtmlToPdfBuffer({
          sourceUrl,
          printOptions: { sourceType: 'web-html', normalizeScrollableContent: true }
        });
        global.lookupFs.writeFileSync(output, result.pdfBuffer);
        return { bytes: result.pdfBuffer.length, diagnostics: result.diagnostics };
      }, { output, sourceUrl });
      fs.writeFileSync(path.join(tmp, 'web-diagnostics.json'), JSON.stringify(result, null, 2));
      console.log(JSON.stringify(result, null, 2));
    }
    console.log(JSON.stringify({ rendererErrors: errors, status: await page.locator('#statusText').textContent() }));
    assert.deepEqual(errors, []);
  } finally {
    await app.close();
    if (localServer) await new Promise(resolve => localServer.close(resolve));
  }
}
run().catch(error => { console.error(error); process.exitCode = 1; });
