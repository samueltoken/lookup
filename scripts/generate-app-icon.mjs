import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PNG } from 'pngjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const sourcePath = path.join(root, 'src', 'renderer', 'assets', 'logo.png');

// Compile the supplied artwork into Windows icon sizes without redrawing it.
function resize(source, size) {
  const output = new PNG({ width: size, height: size });
  const sx = source.width / size;
  const sy = source.height / size;
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const x0 = x * sx, x1 = (x + 1) * sx;
      const y0 = y * sy, y1 = (y + 1) * sy;
      let area = 0, alpha = 0, red = 0, green = 0, blue = 0;
      for (let py = Math.floor(y0); py < Math.ceil(y1); py++) {
        for (let px = Math.floor(x0); px < Math.ceil(x1); px++) {
          const weight = (Math.min(x1, px + 1) - Math.max(x0, px)) * (Math.min(y1, py + 1) - Math.max(y0, py));
          const index = (py * source.width + px) * 4;
          const coverage = source.data[index + 3] / 255 * weight;
          area += weight; alpha += coverage;
          red += source.data[index] * coverage;
          green += source.data[index + 1] * coverage;
          blue += source.data[index + 2] * coverage;
        }
      }
      const index = (y * size + x) * 4;
      output.data[index] = alpha ? Math.round(red / alpha) : 0;
      output.data[index + 1] = alpha ? Math.round(green / alpha) : 0;
      output.data[index + 2] = alpha ? Math.round(blue / alpha) : 0;
      output.data[index + 3] = Math.round(alpha / area * 255);
    }
  }
  return PNG.sync.write(output);
}

function buildIco(images, sizes) {
  const header = Buffer.alloc(6 + images.length * 16);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(images.length, 4);
  let offset = header.length;
  for (let i = 0; i < images.length; i++) {
    const base = 6 + i * 16;
    header[base] = sizes[i] === 256 ? 0 : sizes[i];
    header[base + 1] = header[base];
    header.writeUInt16LE(1, base + 4);
    header.writeUInt16LE(32, base + 6);
    header.writeUInt32LE(images[i].length, base + 8);
    header.writeUInt32LE(offset, base + 12);
    offset += images[i].length;
  }
  return Buffer.concat([header, ...images]);
}

const artwork = await fs.readFile(sourcePath);
const source = PNG.sync.read(artwork);
const sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256];
const images = sizes.map(size => resize(source, size));
await fs.mkdir(path.join(root, 'build'), { recursive: true });
await fs.writeFile(path.join(root, 'build', 'icon.ico'), buildIco(images, sizes));
await fs.writeFile(path.join(root, 'build', 'icon.png'), artwork);
console.log('Supplied logo compiled into Windows icons.');
