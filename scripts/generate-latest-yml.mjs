import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const releaseDir = path.join(rootDir, "release");

function toYaml(data) {
  return [
    `version: ${data.version}`,
    "files:",
    `  - url: ${data.fileName}`,
    `    sha512: ${data.sha512}`,
    `    size: ${data.size}`,
    `path: ${data.fileName}`,
    `sha512: ${data.sha512}`,
    `releaseDate: '${data.releaseDate}'`,
    ""
  ].join("\n");
}

async function main() {
  const pkgPath = path.join(rootDir, "package.json");
  const pkg = JSON.parse(await fs.readFile(pkgPath, "utf8"));
  const version = String(pkg.version || "").trim();
  if (!version) {
    throw new Error("package.json에서 version을 찾지 못했습니다.");
  }

  const fileName = `lookup-Setup-${version}.exe`;
  const exePath = path.join(releaseDir, fileName);
  const stat = await fs.stat(exePath);
  if (!stat.isFile()) {
    throw new Error(`설치 파일이 없습니다: ${exePath}`);
  }

  const fileBuffer = await fs.readFile(exePath);
  const sha512 = crypto.createHash("sha512").update(fileBuffer).digest("base64");
  const yml = toYaml({
    version,
    fileName,
    sha512,
    size: stat.size,
    releaseDate: new Date(stat.mtime).toISOString()
  });

  const outPath = path.join(releaseDir, "latest.yml");
  await fs.writeFile(outPath, yml, "utf8");
  console.log(`latest.yml 생성 완료: ${outPath}`);
}

main().catch((error) => {
  console.error("latest.yml 생성 실패:", error);
  process.exit(1);
});
