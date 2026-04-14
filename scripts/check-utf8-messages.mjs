import fs from "node:fs";
import path from "node:path";

const targetPath = path.join(process.cwd(), "build", "messages.yml");
const bytes = fs.readFileSync(targetPath);

let decoded = "";
try {
  decoded = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
} catch (error) {
  console.error("messages.yml is not valid UTF-8:", error?.message || error);
  process.exit(1);
}

if (!decoded.includes("업데이트 중입니다. 잠시만 기다려주세요.")) {
  console.error("messages.yml does not contain the required Korean update text.");
  process.exit(1);
}

console.log("messages.yml UTF-8 check passed.");
