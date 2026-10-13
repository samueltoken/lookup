"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { pathToFileURL } = require("node:url");
const { spawn } = require("node:child_process");

function quotePowerShell(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function buildExcelReadOnlyScript(sourcePath, outputPdfPath) {
  return `
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
  Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class LookupExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId); }'
  $excel = New-Object -ComObject Excel.Application
  [uint32]$lookupExcelProcessId = 0
  [LookupExcelOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$lookupExcelProcessId) | Out-Null
  Write-Output ('LOOKUP_STATUS {"excelProcessId":' + $lookupExcelProcessId + '}')
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  $workbook = $excel.Workbooks.Open(${quotePowerShell(sourcePath)}, 0, $true)
  $workbook.ExportAsFixedFormat(0, ${quotePowerShell(outputPdfPath)})
} finally {
  if ($workbook -ne $null) {
    try { $workbook.Close($false) | Out-Null } catch {}
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null
  }
  if ($excel -ne $null) {
    try { $excel.Quit() | Out-Null } catch {}
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
if (-not (Test-Path -LiteralPath ${quotePowerShell(outputPdfPath)})) { throw 'excel_export_failed' }
`;
}

function buildOwnedExcelCleanupScript(processId, startedAt) {
  const pid = Math.floor(Number(processId));
  if (!Number.isInteger(pid) || pid <= 0) return "";
  const earliest = new Date(Number(startedAt) - 2000).toISOString();
  return `$lookupExcelProcess = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; if ($lookupExcelProcess -and $lookupExcelProcess.ProcessName -eq 'EXCEL' -and $lookupExcelProcess.StartTime.ToUniversalTime() -ge [DateTime]::Parse(${quotePowerShell(earliest)}).ToUniversalTime()) { Stop-Process -Id ${pid} -Force }`;
}

function runBoundedProcess(executable, args, timeoutMs) {
  return new Promise((resolve) => {
    const child = spawn(executable, args, { windowsHide: true });
    let stdout = "";
    let stderr = "";
    let finished = false;
    let timedOut = false;
    let forcedFinish;
    const finish = (value) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer); clearTimeout(forcedFinish);
      resolve({ stdout: stdout.trim(), stderr: stderr.trim(), timedOut, ...value });
    };
    child.stdout.on("data", (chunk) => { stdout = (stdout + chunk.toString()).slice(-65536); });
    child.stderr.on("data", (chunk) => { stderr = (stderr + chunk.toString()).slice(-65536); });
    child.on("error", (error) => finish({ ok: false, error: error.message }));
    child.on("close", (status) => finish({ ok: status === 0 && !timedOut, status }));
    const timer = setTimeout(() => {
      timedOut = true;
      if (process.platform === "win32" && child.pid && child.exitCode === null) {
        // This is the PID returned by our spawn, never a name-wide Office kill.
        const killer = spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true });
        killer.on("error", () => {});
      } else child.kill("SIGKILL");
      forcedFinish = setTimeout(() => finish({ ok: false, error: "engine_timeout" }), 5000);
    }, timeoutMs);
  });
}

async function convertOfficeWithLibreOffice(sourcePath, outputPdfPath, detected, { timeoutMs = 45000, pdfFilter = "" } = {}) {
  if (!detected?.path) return { ok: false, missingEngine: true, stderr: "libreoffice_missing", engineMode: "" };
  const outputDir = path.resolve(path.dirname(outputPdfPath));
  const workingDir = path.join(outputDir, `.lookup-office-${crypto.randomUUID()}`);
  const privateOutput = path.join(workingDir, "output");
  try {
    await fs.mkdir(privateOutput, { recursive: true });
    const result = await runBoundedProcess(detected.path, [
      `-env:UserInstallation=${pathToFileURL(path.join(workingDir, "profile")).href}`,
      "--headless", "--invisible", "--nologo", "--norestore", "--convert-to", pdfFilter ? `pdf:${pdfFilter}` : "pdf",
      "--outdir", privateOutput, sourcePath
    ], timeoutMs);
    if (!result.ok) return { ...result, engineMode: detected.mode || "libreoffice-system" };
    const produced = path.join(privateOutput, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
    try { await fs.copyFile(produced, outputPdfPath); }
    catch (error) { return { ok: false, stderr: "libreoffice_pdf_not_created", error: error.message, engineMode: detected.mode || "libreoffice-system" }; }
    return { ...result, ok: true, engineMode: detected.mode || "libreoffice-system" };
  } finally {
    if (path.dirname(path.resolve(workingDir)) === outputDir && path.basename(workingDir).startsWith(".lookup-office-")) {
      await fs.rm(workingDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

async function convertSpreadsheetWithLibreOffice(sourcePath, outputPdfPath, detected, options = {}) {
  return convertOfficeWithLibreOffice(sourcePath, outputPdfPath, detected, { ...options, pdfFilter: "calc_pdf_Export" });
}

module.exports = { buildExcelReadOnlyScript, buildOwnedExcelCleanupScript, convertSpreadsheetWithLibreOffice, convertOfficeWithLibreOffice };
