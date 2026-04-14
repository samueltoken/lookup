import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import https from "node:https";

const OWNER = "samueltoken";
const REPO = "lookup";

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
  return { username, password };
}

function requestGitHub(authHeader, method, apiPath, body = null) {
  return new Promise((resolve, reject) => {
    const bodyText = body ? JSON.stringify(body) : "";
    const request = https.request(
      {
        protocol: "https:",
        hostname: "api.github.com",
        path: apiPath,
        method,
        headers: {
          Accept: "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "User-Agent": "lookup-release-notes-repair",
          Authorization: authHeader,
          "Content-Type": "application/json; charset=utf-8",
          "Content-Length": Buffer.byteLength(bodyText, "utf8")
        }
      },
      (response) => {
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => {
          const raw = Buffer.concat(chunks).toString("utf8");
          const statusCode = Number(response.statusCode || 0);
          if (statusCode < 200 || statusCode >= 300) {
            reject(new Error(`GitHub API ${statusCode}: ${raw.slice(0, 240)}`));
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

    request.on("error", reject);
    if (bodyText) {
      request.write(bodyText, "utf8");
    }
    request.end();
  });
}

async function main() {
  const credential = getGitCredential();
  const authHeader = `Basic ${Buffer.from(`${credential.username}:${credential.password}`, "utf8").toString("base64")}`;
  const releases = await requestGitHub(authHeader, "GET", `/repos/${OWNER}/${REPO}/releases?per_page=100`);

  const targets = [
    { tag: "v1.1.7", name: "lookup v1.1.7", noteFile: "release-notes/v1.1.7.md" },
    { tag: "v1.2.0", name: "lookup v1.2.0", noteFile: "release-notes/v1.2.0.md" },
    { tag: "v1.2.1", name: "lookup v1.2.1", noteFile: "release-notes/v1.2.1.md" }
  ];

  for (const target of targets) {
    const release = releases.find((entry) => entry.tag_name === target.tag);
    if (!release?.id) {
      continue;
    }
    const notePath = path.resolve(target.noteFile);
    const body = await fs.readFile(notePath, "utf8");
    await requestGitHub(authHeader, "PATCH", `/repos/${OWNER}/${REPO}/releases/${release.id}`, {
      name: target.name,
      body,
      draft: false,
      prerelease: false
    });
    console.log(`release body updated: ${target.tag}`);
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
