import fs from "node:fs";
import path from "node:path";
import { rcedit } from "rcedit";

export default async function afterPack(context) {
  try {
    if (context?.electronPlatformName !== "win32") {
      return;
    }
    const projectDir = String(context?.packager?.projectDir || process.cwd());
    const productFilename = String(context?.packager?.appInfo?.productFilename || "lookup");
    const appOutDir = String(context?.appOutDir || "");
    if (!appOutDir) {
      return;
    }

    const exePath = path.join(appOutDir, `${productFilename}.exe`);
    const iconPath = path.join(projectDir, "build", "icon.ico");
    if (!fs.existsSync(exePath) || !fs.existsSync(iconPath)) {
      return;
    }

    await rcedit(exePath, { icon: iconPath });
    console.log(`afterPack icon patched: ${exePath}`);
  } catch (error) {
    console.warn(`afterPack icon patch skipped: ${error?.message || error}`);
  }
}
