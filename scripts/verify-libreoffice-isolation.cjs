const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const { PDFDocument } = require('pdf-lib');
const { convertOfficeWithLibreOffice } = require('../src/main/spreadsheet/native-engine');

async function run() {
  const root = path.resolve(__dirname, '..');
  const candidates = [
    path.join(root, 'build', 'libreoffice', 'program', 'soffice.exe'),
    path.join(process.env.ProgramFiles || 'C:/Program Files', 'LibreOffice', 'program', 'soffice.exe')
  ];
  let executable;
  for (const candidate of candidates) {
    try { await fs.access(candidate); executable = candidate; break; } catch {}
  }
  if (!executable) {
    console.log(JSON.stringify({ officeIsolation: 'unavailable', reason: 'LibreOffice runtime is not installed' }));
    return;
  }
  const output = path.join(root, 'tmp', 'office-verification');
  await fs.mkdir(output, { recursive: true });
  const source = path.join(output, '한글 공백 원본.txt');
  await fs.writeFile(source, '원본 보존 검사\nLibreOffice isolated conversion\n마지막 내용\n', 'utf8');
  const before = crypto.createHash('sha256').update(await fs.readFile(source)).digest('hex');
  const result = await convertOfficeWithLibreOffice(source, path.join(output, 'result.pdf'), {
    path: executable, mode: 'libreoffice-test'
  });
  assert.equal(result.ok, true, JSON.stringify(result));
  const doc = await PDFDocument.load(await fs.readFile(path.join(output, 'result.pdf')));
  assert.ok(doc.getPageCount() > 0);
  assert.equal(crypto.createHash('sha256').update(await fs.readFile(source)).digest('hex'), before);
  assert.equal((await fs.readdir(output)).some(name => name.startsWith('.lookup-office-')), false);
  console.log(JSON.stringify({ officeIsolation: 'passed', pageCount: doc.getPageCount(), sourceUnchanged: true, temporaryProfileRemoved: true }));
}
run().catch(error => { console.error(error); process.exitCode = 1; });
