import fs from "node:fs";
import fsp from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import https from "node:https";

const projectRoot = process.cwd();
const accdbRoot = path.join(projectRoot, "build", "accdb");
const libDir = path.join(accdbRoot, "lib");
const classDir = path.join(accdbRoot, "classes");
const javaDir = path.join(accdbRoot, "java");
const sourceFile = path.join(accdbRoot, "src", "LookupAccdbDump.java");

const JAR_URLS = [
  "https://repo1.maven.org/maven2/net/sf/ucanaccess/ucanaccess/5.0.1/ucanaccess-5.0.1.jar",
  "https://repo1.maven.org/maven2/org/hsqldb/hsqldb/2.5.0/hsqldb-2.5.0.jar",
  "https://repo1.maven.org/maven2/com/healthmarketscience/jackcess/jackcess/3.0.1/jackcess-3.0.1.jar",
  "https://repo1.maven.org/maven2/org/apache/commons/commons-lang3/3.8.1/commons-lang3-3.8.1.jar",
  "https://repo1.maven.org/maven2/commons-logging/commons-logging/1.2/commons-logging-1.2.jar"
];

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { "User-Agent": "lookup-accdb-setup" } }, (res) => {
      if (!res.statusCode || res.statusCode < 200 || res.statusCode >= 300) {
        reject(new Error(`JSON fetch failed: ${url} (${res.statusCode || "unknown"})`));
        return;
      }
      let raw = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => {
        raw += chunk;
      });
      res.on("end", () => {
        try {
          resolve(JSON.parse(raw));
        } catch (error) {
          reject(error);
        }
      });
    }).on("error", reject);
  });
}

function downloadFile(url, destPath) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(destPath);
    const request = https.get(url, { headers: { "User-Agent": "lookup-accdb-setup" } }, (res) => {
      if (res.statusCode && res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        file.close();
        fs.unlink(destPath, () => {
          downloadFile(res.headers.location, destPath).then(resolve).catch(reject);
        });
        return;
      }
      if (!res.statusCode || res.statusCode < 200 || res.statusCode >= 300) {
        file.close();
        fs.unlink(destPath, () => {
          reject(new Error(`Download failed: ${url} (${res.statusCode || "unknown"})`));
        });
        return;
      }
      const totalBytes = Number.parseInt(String(res.headers["content-length"] || "0"), 10);
      let downloadedBytes = 0;
      let lastLogAt = Date.now();
      res.on("data", (chunk) => {
        downloadedBytes += chunk.length;
        const now = Date.now();
        if (now - lastLogAt >= 2500) {
          if (Number.isFinite(totalBytes) && totalBytes > 0) {
            const percent = Math.min(100, (downloadedBytes / totalBytes) * 100);
            console.log(`[accdb] downloading ${path.basename(destPath)} ${percent.toFixed(1)}%`);
          } else {
            console.log(`[accdb] downloading ${path.basename(destPath)} ${Math.round(downloadedBytes / (1024 * 1024))}MB`);
          }
          lastLogAt = now;
        }
      });
      res.pipe(file);
      file.on("finish", () => {
        file.close(() => resolve(destPath));
      });
    });
    request.on("error", (error) => {
      file.close();
      fs.unlink(destPath, () => reject(error));
    });
  });
}

function runOrThrow(command, args, cwd = projectRoot) {
  const result = spawnSync(command, args, {
    cwd,
    stdio: "inherit",
    windowsHide: true
  });
  if (result.status !== 0) {
    throw new Error(`Command failed: ${command} ${args.join(" ")}`);
  }
}

function extractZip(zipPath, targetDir) {
  ensureDir(targetDir);
  runOrThrow("powershell", [
    "-NoLogo",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `Expand-Archive -LiteralPath '${zipPath.replace(/'/g, "''")}' -DestinationPath '${targetDir.replace(/'/g, "''")}' -Force`
  ]);
}

async function findSingleSubdir(parentDir) {
  const entries = await fsp.readdir(parentDir, { withFileTypes: true });
  const dirs = entries.filter((entry) => entry.isDirectory());
  if (dirs.length === 0) {
    throw new Error(`No extracted directory found in ${parentDir}`);
  }
  return path.join(parentDir, dirs[0].name);
}

async function setupJars() {
  ensureDir(libDir);
  for (const jarUrl of JAR_URLS) {
    const fileName = path.basename(jarUrl);
    const target = path.join(libDir, fileName);
    if (!fs.existsSync(target)) {
      console.log(`[accdb] download ${fileName}`);
      await downloadFile(jarUrl, target);
    }
  }
}

async function setupJavaRuntime(tempRoot) {
  ensureDir(javaDir);
  const jreApi = "https://api.adoptium.net/v3/assets/latest/21/hotspot?architecture=x64&image_type=jre&os=windows&vendor=eclipse";
  const ecjUrl = "https://repo1.maven.org/maven2/org/eclipse/jdt/ecj/3.42.0/ecj-3.42.0.jar";

  const jreMeta = await fetchJson(jreApi);
  const jreLink = jreMeta?.[0]?.binary?.package?.link;
  if (!jreLink) {
    throw new Error("Failed to resolve Temurin runtime download links.");
  }

  const jreZip = path.join(tempRoot, "jre.zip");
  const ecjJar = path.join(tempRoot, "ecj.jar");
  console.log("[accdb] download Temurin JRE");
  await downloadFile(jreLink, jreZip);
  console.log("[accdb] download ECJ");
  await downloadFile(ecjUrl, ecjJar);

  const jreExtract = path.join(tempRoot, "jre-extract");
  extractZip(jreZip, jreExtract);

  const jreHome = await findSingleSubdir(jreExtract);

  fs.rmSync(javaDir, { recursive: true, force: true });
  fs.cpSync(jreHome, javaDir, { recursive: true });

  return { ecjJar };
}

async function compileHelper(ecjJar) {
  ensureDir(classDir);
  const jarPaths = fs.readdirSync(libDir)
    .filter((name) => name.toLowerCase().endsWith(".jar"))
    .map((name) => path.join(libDir, name));
  if (!fs.existsSync(sourceFile)) {
    throw new Error(`Source file not found: ${sourceFile}`);
  }
  const javaPath = path.join(javaDir, "bin", "java.exe");
  if (!fs.existsSync(javaPath)) {
    throw new Error("java.exe not found in bundled runtime.");
  }
  const classpath = jarPaths.join(path.delimiter);
  runOrThrow(javaPath, ["-jar", ecjJar, "-1.8", "-encoding", "UTF-8", "-cp", classpath, "-d", classDir, sourceFile]);
}

async function main() {
  ensureDir(accdbRoot);
  ensureDir(path.join(accdbRoot, "src"));
  const tempRoot = await fsp.mkdtemp(path.join(os.tmpdir(), "lookup-accdb-"));
  try {
    await setupJars();
    const { ecjJar } = await setupJavaRuntime(tempRoot);
    await compileHelper(ecjJar);
    console.log("[accdb] runtime setup complete");
  } finally {
    fs.rmSync(tempRoot, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error?.message || error);
  process.exit(1);
});
