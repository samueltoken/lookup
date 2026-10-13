const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const root = path.resolve(__dirname, '..');
const runtime = createRequire(path.join(process.env.USERPROFILE, '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/package.json'));
const { _electron: electron } = runtime('playwright');
const output = path.join(root, 'tmp', 'packaged-verification');
fs.mkdirSync(output, { recursive: true });

async function run() {
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  const profile = path.join(output, 'profile');
  const fixture = path.join(root, 'tmp', 'web-preparation-check', 'fixture.pdf');
  const app = await electron.launch({
    executablePath: path.join(root, 'release', 'v1.3.6', 'win-unpacked', 'lookup.exe'),
    args: [`--user-data-dir=${profile}`, fixture], env, timeout: 30000
  });
  try {
    const page = await app.firstWindow();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const actualProfile = await app.evaluate(({ app }) => app.getPath('userData'));
    assert.equal(path.resolve(actualProfile), path.resolve(profile), 'packaged test uses an isolated profile');
    await page.waitForFunction(() => document.querySelector('#pageCountLabel')?.textContent.includes('10'), null, { timeout: 20000 });
    assert.equal(await page.evaluate(() => window.lookupAPI.getAppVersion()), '1.3.6');
    await page.locator('#tabViewBtn').click();
    await page.locator('#viewFocusBtn').click();
    await page.locator('#tabHomeBtn').click();
    for (let i = 0; i < 7; i++) await page.locator('#zoomInBtn').click();
    await page.waitForTimeout(400);
    const box = await page.locator('#viewerPanel').boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.wheel(0, 3000);
    await page.waitForTimeout(200);
    assert.equal(await page.locator('#pageInput').inputValue(), '1');
    await page.locator('#tabViewBtn').click();
    await page.locator('#toggleDarkBtn').click();
    assert.equal(await page.locator('body').evaluate(el => el.classList.contains('dark')), true);
    await page.locator('#toggleDarkBtn').click();
    await page.screenshot({ path: path.join(output, 'packaged-viewer.png') });
    assert.deepEqual(errors, []);
    const summary = { packagedApp: 'passed', version: '1.3.6', pages: 10, magnifiedWheelRetainsPage: true, darkMode: true, profile: actualProfile, rendererErrors: errors };
    fs.writeFileSync(path.join(output, 'result.json'), JSON.stringify(summary, null, 2));
    console.log(JSON.stringify(summary));
  } finally { await app.close(); }
}
run().catch(error => { console.error(error); process.exitCode = 1; });
