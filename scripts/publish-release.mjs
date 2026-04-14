import { spawnSync } from "node:child_process";
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import https from "node:https";

const OWNER = "samueltoken";
const REPO = "lookup";
const TARGET_TAG = "v1.2.2";
const TARGET_NAME = "lookup v1.2.2";

const REPAIR_TAGS = ["v1.1.7", "v1.2.0", "v1.2.1"];
const RELEASE_NOTES_DIR = path.resolve("release-notes");
const RELEASE_DIR = path.resolve("release");
const TARGET_ASSETS = [
  "lookup-Setup-1.2.2.exe",
  "latest.yml",
  "lookup-Setup-1.2.2.exe.blockmap"
];

function getGitCredential() {
  const result = spawnSync("git", ["credential", "fill"], {
    input: "protocol=https\nhost=github.com\n\n",
    encoding: "utf8"
  });
  if (result.status !== 0 || !result.stdout) {
    throw new Error("GitHub 인증 정보를 찾지 못했습니다.");
  }
  const map = new Map();
  for (const line of result.stdout.split(/\r?\n/)) {
    const idx = line.indexOf("=");
    if (idx <= 0) {
      continue;
    }
    map.set(line.slice(0, idx), line.slice(idx + 1));
  }
  const username = map.get("username") || "";
  const password = map.get("password") || "";
  if (!username || !password) {
    throw new Error("GitHub 인증 정보(username/password)가 비어 있습니다.");
  }
  return `Basic ${Buffer.from(`${username}:${password}`, "utf8").toString("base64")}`;
}

function requestJson(authHeader, method, apiPath, body = null, host = "api.github.com") {
  return new Promise((resolve, reject) => {
    const bodyText = body ? JSON.stringify(body) : "";
    const req = https.request(
      {
        protocol: "https:",
        hostname: host,
        path: apiPath,
        method,
        headers: {
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "User-Agent": "lookup-release-publisher",
          Authorization: authHeader,
          "Content-Type": "application/json; charset=utf-8",
          "Content-Length": Buffer.byteLength(bodyText, "utf8")
        }
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          const raw = Buffer.concat(chunks).toString("utf8");
          const statusCode = Number(res.statusCode || 0);
          if (statusCode < 200 || statusCode >= 300) {
            reject(new Error(`GitHub API ${statusCode}: ${raw.slice(0, 260)}`));
            return;
          }
          if (!raw) {
            resolve({});
            return;
          }
          try {
            resolve(JSON.parse(raw));
          } catch (_error) {
            resolve({});
          }
        });
      }
    );
    req.on("error", reject);
    if (bodyText) {
      req.write(bodyText, "utf8");
    }
    req.end();
  });
}

function uploadBinary(authHeader, uploadPath, fileName, fileBuffer) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        protocol: "https:",
        hostname: "uploads.github.com",
        path: `${uploadPath}?name=${encodeURIComponent(fileName)}`,
        method: "POST",
        headers: {
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "User-Agent": "lookup-release-publisher",
          Authorization: authHeader,
          "Content-Type": "application/octet-stream",
          "Content-Length": fileBuffer.length
        }
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          const raw = Buffer.concat(chunks).toString("utf8");
          const statusCode = Number(res.statusCode || 0);
          if (statusCode < 200 || statusCode >= 300) {
            reject(new Error(`Asset upload ${statusCode}: ${raw.slice(0, 260)}`));
            return;
          }
          try {
            resolve(JSON.parse(raw));
          } catch (_error) {
            resolve({});
          }
        });
      }
    );
    req.on("error", reject);
    req.write(fileBuffer);
    req.end();
  });
}

function hasBrokenText(body) {
  return /\?{2,}/.test(String(body || ""));
}

async function loadReleaseNotes(tag) {
  const notePath = path.join(RELEASE_NOTES_DIR, `${tag}.md`);
  return await fsp.readFile(notePath, "utf8");
}

async function ensureRepairedReleases(authHeader) {
  const releases = await requestJson(authHeader, "GET", `/repos/${OWNER}/${REPO}/releases?per_page=100`);
  for (const tag of REPAIR_TAGS) {
    const release = releases.find((entry) => entry.tag_name === tag);
    if (!release?.id) {
      continue;
    }
    if (!hasBrokenText(release.body)) {
      continue;
    }
    const body = await loadReleaseNotes(tag);
    await requestJson(authHeader, "PATCH", `/repos/${OWNER}/${REPO}/releases/${release.id}`, {
      name: `lookup ${tag}`,
      body,
      draft: false,
      prerelease: false
    });
    console.log(`repaired release body: ${tag}`);
  }
}

async function createOrUpdateTargetRelease(authHeader) {
  const releases = await requestJson(authHeader, "GET", `/repos/${OWNER}/${REPO}/releases?per_page=100`);
  const existing = releases.find((entry) => entry.tag_name === TARGET_TAG);
  const body = await loadReleaseNotes(TARGET_TAG);

  if (existing?.id) {
    const updated = await requestJson(authHeader, "PATCH", `/repos/${OWNER}/${REPO}/releases/${existing.id}`, {
      name: TARGET_NAME,
      body,
      draft: false,
      prerelease: false
    });
    return updated;
  }

  return await requestJson(authHeader, "POST", `/repos/${OWNER}/${REPO}/releases`, {
    tag_name: TARGET_TAG,
    name: TARGET_NAME,
    body,
    draft: false,
    prerelease: false
  });
}

async function uploadAssets(authHeader, release) {
  if (!release?.id || !release?.upload_url) {
    throw new Error("릴리즈 업로드 URL을 찾지 못했습니다.");
  }
  const uploadPath = String(release.upload_url).replace(/^https:\/\/uploads\.github\.com/, "").replace(/\{.*$/, "");
  const currentAssets = Array.isArray(release.assets) ? release.assets : [];

  for (const fileName of TARGET_ASSETS) {
    const filePath = path.join(RELEASE_DIR, fileName);
    if (!fs.existsSync(filePath)) {
      throw new Error(`릴리즈 자산 파일이 없습니다: ${fileName}`);
    }

    const existingAsset = currentAssets.find((asset) => asset.name === fileName);
    if (existingAsset?.id) {
      await requestJson(authHeader, "DELETE", `/repos/${OWNER}/${REPO}/releases/assets/${existingAsset.id}`);
    }

    const buffer = await fsp.readFile(filePath);
    await uploadBinary(authHeader, uploadPath, fileName, buffer);
    console.log(`uploaded asset: ${fileName}`);
  }
}

async function main() {
  const authHeader = getGitCredential();
  await ensureRepairedReleases(authHeader);
  const release = await createOrUpdateTargetRelease(authHeader);
  await uploadAssets(authHeader, release);
  console.log(`release published: ${TARGET_TAG}`);
}

main().catch((error) => {
  console.error(error?.message || error);
  process.exitCode = 1;
});
