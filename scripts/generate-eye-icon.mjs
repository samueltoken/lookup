import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PNG } from "pngjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const buildDir = path.join(rootDir, "build");

function clamp01(value) {
  return Math.max(0, Math.min(1, value));
}

function smoothstep(a, b, x) {
  const t = clamp01((x - a) / (b - a));
  return t * t * (3 - 2 * t);
}

function blendPixel(png, x, y, color, alpha) {
  const index = (png.width * y + x) << 2;
  const srcA = clamp01(alpha);
  if (srcA <= 0) {
    return;
  }

  const dstA = png.data[index + 3] / 255;
  const outA = srcA + dstA * (1 - srcA);
  if (outA <= 0) {
    return;
  }

  const srcR = color[0] / 255;
  const srcG = color[1] / 255;
  const srcB = color[2] / 255;
  const dstR = png.data[index] / 255;
  const dstG = png.data[index + 1] / 255;
  const dstB = png.data[index + 2] / 255;

  const outR = (srcR * srcA + dstR * dstA * (1 - srcA)) / outA;
  const outG = (srcG * srcA + dstG * dstA * (1 - srcA)) / outA;
  const outB = (srcB * srcA + dstB * dstA * (1 - srcA)) / outA;

  png.data[index] = Math.round(outR * 255);
  png.data[index + 1] = Math.round(outG * 255);
  png.data[index + 2] = Math.round(outB * 255);
  png.data[index + 3] = Math.round(outA * 255);
}

function drawDuck(size) {
  const png = new PNG({ width: size, height: size });
  const bodyCx = size * 0.5;
  const bodyCy = size * 0.58;
  const bodyR = size * 0.29;
  const headCx = size * 0.65;
  const headCy = size * 0.37;
  const headR = size * 0.19;
  const wingCx = size * 0.48;
  const wingCy = size * 0.56;
  const wingRx = size * 0.16;
  const wingRy = size * 0.115;
  const beakCx = size * 0.81;
  const beakCy = size * 0.42;
  const beakRx = size * 0.08;
  const beakRy = size * 0.045;
  const eyeCx = size * 0.68;
  const eyeCy = size * 0.35;
  const eyeR = size * 0.024;

  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const px = x + 0.5;
      const py = y + 0.5;

      const bodyDx = px - bodyCx;
      const bodyDy = py - bodyCy;
      const bodyDist = Math.sqrt(bodyDx * bodyDx + bodyDy * bodyDy);
      const bodyMask = 1 - smoothstep(bodyR, bodyR * 1.08, bodyDist);
      if (bodyMask > 0) {
        blendPixel(png, x, y, [254, 217, 60], bodyMask);
      }
      const bodyOutline = smoothstep(bodyR * 0.9, bodyR, bodyDist) - smoothstep(bodyR, bodyR * 1.1, bodyDist);
      if (bodyOutline > 0) {
        blendPixel(png, x, y, [214, 163, 24], bodyOutline * 0.9);
      }

      const headDx = px - headCx;
      const headDy = py - headCy;
      const headDist = Math.sqrt(headDx * headDx + headDy * headDy);
      const headMask = 1 - smoothstep(headR, headR * 1.08, headDist);
      if (headMask > 0) {
        blendPixel(png, x, y, [255, 224, 77], headMask);
      }
      const headOutline = smoothstep(headR * 0.9, headR, headDist) - smoothstep(headR, headR * 1.1, headDist);
      if (headOutline > 0) {
        blendPixel(png, x, y, [214, 163, 24], headOutline * 0.9);
      }

      const wingDx = (px - wingCx) / wingRx;
      const wingDy = (py - wingCy) / wingRy;
      const wingNorm = wingDx * wingDx + wingDy * wingDy;
      const wingMask = 1 - smoothstep(1, 1.12, wingNorm);
      if (wingMask > 0) {
        blendPixel(png, x, y, [245, 190, 42], wingMask * 0.55);
      }

      const beakNorm =
        ((px - beakCx) * (px - beakCx)) / (beakRx * beakRx) +
        ((py - beakCy) * (py - beakCy)) / (beakRy * beakRy);
      const beakMask = 1 - smoothstep(1, 1.1, beakNorm);
      if (beakMask > 0) {
        blendPixel(png, x, y, [245, 130, 36], beakMask);
      }

      const eyeDx = px - eyeCx;
      const eyeDy = py - eyeCy;
      const eyeDist = Math.sqrt(eyeDx * eyeDx + eyeDy * eyeDy);
      const eyeMask = 1 - smoothstep(eyeR, eyeR * 1.2, eyeDist);
      if (eyeMask > 0) {
        blendPixel(png, x, y, [20, 26, 39], eyeMask);
      }

      const shineDx = px - (headCx - size * 0.06);
      const shineDy = py - (headCy - size * 0.06);
      const shineDist = Math.sqrt(shineDx * shineDx + shineDy * shineDy);
      const shineMask = 1 - smoothstep(size * 0.03, size * 0.05, shineDist);
      if (shineMask > 0) {
        blendPixel(png, x, y, [255, 255, 255], shineMask * 0.65);
      }
    }
  }

  return PNG.sync.write(png);
}

function buildIco(pngBuffers, sizes) {
  const count = pngBuffers.length;
  const header = Buffer.alloc(6 + count * 16);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(count, 4);

  let offset = header.length;
  for (let i = 0; i < count; i += 1) {
    const size = sizes[i];
    const buffer = pngBuffers[i];
    const base = 6 + i * 16;
    header[base] = size === 256 ? 0 : size;
    header[base + 1] = size === 256 ? 0 : size;
    header[base + 2] = 0;
    header[base + 3] = 0;
    header.writeUInt16LE(1, base + 4);
    header.writeUInt16LE(32, base + 6);
    header.writeUInt32LE(buffer.length, base + 8);
    header.writeUInt32LE(offset, base + 12);
    offset += buffer.length;
  }

  return Buffer.concat([header, ...pngBuffers]);
}

async function main() {
  await fs.mkdir(buildDir, { recursive: true });
  const sizes = [16, 24, 32, 48, 64, 128, 256];
  const pngBuffers = sizes.map((size) => drawDuck(size));
  const icoBuffer = buildIco(pngBuffers, sizes);

  await fs.writeFile(path.join(buildDir, "icon.ico"), icoBuffer);
  await fs.writeFile(path.join(buildDir, "icon.png"), pngBuffers[pngBuffers.length - 1]);
  console.log("오리 아이콘 생성 완료: build/icon.ico, build/icon.png");
}

main().catch((error) => {
  console.error("오리 아이콘 생성 실패:", error);
  process.exit(1);
});
