import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import { rcedit } from "rcedit";

const MANIFEST_RELATIVE_PATH = "resources/lookup-install.json";

function isInside(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function assertOrdinaryEntry(root, entry) {
  if (!isInside(root, entry)) {
    throw new Error(`Install manifest path escapes packaged application: ${entry}`);
  }
  const stat = fs.lstatSync(entry);
  if (stat.isSymbolicLink() || (!stat.isDirectory() && !stat.isFile())) {
    throw new Error(`Install manifest refuses a link or special file: ${entry}`);
  }
  if (!isInside(root, fs.realpathSync.native(entry))) {
    throw new Error(`Install manifest resolved path escapes packaged application: ${entry}`);
  }
  return stat;
}

// Windows has reparse points other than links. Check their native attributes
// before descending; Directory.EnumerateFileSystemEntries does not recurse.
function assertNoWindowsReparsePoints(root) {
  if (process.platform !== "win32") return;
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $manifestRoot = $env:LOOKUP_INSTALL_MANIFEST_ROOT
  $pending = [System.Collections.Generic.Stack[string]]::new()
  $pending.Push($manifestRoot)
  while ($pending.Count -gt 0) {
    $entry = $pending.Pop()
    $attributes = [System.IO.File]::GetAttributes($entry)
    if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw "Install manifest refuses a Windows reparse point: $entry"
    }
    if (($attributes -band [System.IO.FileAttributes]::Directory) -ne 0) {
      foreach ($child in [System.IO.Directory]::EnumerateFileSystemEntries($entry)) {
        $pending.Push($child)
      }
    }
  }
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
`;
  const result = spawnSync("powershell.exe", [
    "-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand",
    Buffer.from(script, "utf16le").toString("base64"),
  ], {
    env: { ...process.env, LOOKUP_INSTALL_MANIFEST_ROOT: root },
    encoding: "utf8", windowsHide: true, timeout: 120_000, maxBuffer: 1024 * 1024,
  });
  if (result.error || result.status !== 0) {
    throw new Error(`Install manifest safety check failed: ${result.error?.message || result.stderr?.trim() || `PowerShell exit ${result.status}`}`);
  }
}

function hashFile(file) {
  const hash = crypto.createHash("sha256");
  const buffer = Buffer.allocUnsafe(1024 * 1024);
  const descriptor = fs.openSync(file, "r");
  try {
    let bytesRead;
    while ((bytesRead = fs.readSync(descriptor, buffer, 0, buffer.length, null)) > 0) {
      hash.update(buffer.subarray(0, bytesRead));
    }
  } finally {
    fs.closeSync(descriptor);
  }
  return hash.digest("hex");
}

export function writeInstallOwnershipManifest(appOutDir, appInfo) {
  if (!appOutDir || !path.isAbsolute(appOutDir)) {
    throw new Error("Install manifest requires an absolute packaged application directory.");
  }
  const root = path.resolve(appOutDir);
  const resolvedRoot = fs.realpathSync.native(root);
  if (path.relative(root, resolvedRoot) !== "") {
    throw new Error(`Install manifest refuses an application directory reached through a link: ${root}`);
  }
  if (!assertOrdinaryEntry(root, root).isDirectory()) {
    throw new Error(`Install manifest application directory is not a directory: ${root}`);
  }
  assertNoWindowsReparsePoints(root);

  const appId = String(appInfo?.id || appInfo?.appId || "");
  const packageVersion = String(appInfo?.version || "");
  const productFilename = String(appInfo?.productFilename || "lookup");
  if (appId !== "com.lookup.pdfviewer" || !packageVersion || productFilename !== "lookup") {
    throw new Error("Install manifest requires lookup application identity and a package version.");
  }

  const executable = `${productFilename}.exe`;
  if (!assertOrdinaryEntry(root, path.join(root, executable)).isFile()) {
    throw new Error(`Install manifest application executable is missing: ${executable}`);
  }
  const files = [];
  const visit = (directory) => {
    for (const name of fs.readdirSync(directory).sort()) {
      const entry = path.join(directory, name);
      const stat = assertOrdinaryEntry(root, entry);
      if (stat.isDirectory()) {
        visit(entry);
      } else {
        const relative = path.relative(root, entry).split(path.sep).join("/");
        if (relative !== MANIFEST_RELATIVE_PATH) {
          files.push({ path: relative, sha256: hashFile(entry) });
        }
      }
    }
  };
  visit(root);
  files.sort((a, b) => a.path < b.path ? -1 : a.path > b.path ? 1 : 0);
  const manifest = {
    schemaVersion: 1, appId, packageVersion,
    productName: String(appInfo?.productName || "lookup"), executable, files,
  };
  const resources = path.join(root, "resources");
  if (!fs.existsSync(resources)) fs.mkdirSync(resources);
  if (!assertOrdinaryEntry(root, resources).isDirectory()) {
    throw new Error("Install manifest resources directory is not a directory.");
  }
  const manifestPath = path.join(root, ...MANIFEST_RELATIVE_PATH.split("/"));
  if (fs.existsSync(manifestPath) && !assertOrdinaryEntry(root, manifestPath).isFile()) {
    throw new Error("Install manifest destination is not an ordinary file.");
  }
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
  return manifest;
}

export default async function afterPack(context) {
  if (context?.electronPlatformName !== "win32") return;
  const projectDir = String(context?.packager?.projectDir || process.cwd());
  const appInfo = context?.packager?.appInfo;
  const productFilename = String(appInfo?.productFilename || "lookup");
  const appOutDir = String(context?.appOutDir || "");

  try {
    const exePath = path.join(appOutDir, `${productFilename}.exe`);
    const iconPath = path.join(projectDir, "build", "icon.ico");
    if (fs.existsSync(exePath) && fs.existsSync(iconPath)) {
      await rcedit(exePath, { icon: iconPath });
      console.log(`afterPack icon patched: ${exePath}`);
    }
  } catch (error) {
    console.warn(`afterPack icon patch skipped: ${error?.message || error}`);
  }

  const manifest = writeInstallOwnershipManifest(appOutDir, appInfo);
  console.log(`afterPack install manifest written: ${manifest.files.length} files (${manifest.packageVersion})`);
}
