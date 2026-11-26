import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ASSETS_DIR = path.resolve(__dirname, '../src/client/assets');

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function makeChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcBuf = Buffer.alloc(4);
  const combined = Buffer.concat([typeBuf, data]);
  crcBuf.writeUInt32BE(crc32(combined), 0);
  return Buffer.concat([len, combined, crcBuf]);
}

/**
 * Create a crisp Excel-green rounded badge PNG with white inner translation bars.
 */
function createBadgePng(size) {
  const raw = Buffer.alloc(size * (1 + size * 4));
  const radius = Math.max(2, Math.floor(size * 0.18));

  for (let y = 0; y < size; y += 1) {
    const rowStart = y * (1 + size * 4);
    raw[rowStart] = 0; // filter type 0

    for (let x = 0; x < size; x += 1) {
      const idx = rowStart + 1 + x * 4;
      // Rounded corner mask
      const dx = Math.max(radius - x, 0, x - (size - 1 - radius));
      const dy = Math.max(radius - y, 0, y - (size - 1 - radius));
      const inRoundedRect = dx * dx + dy * dy <= radius * radius;

      if (!inRoundedRect) {
        raw[idx] = 0;
        raw[idx + 1] = 0;
        raw[idx + 2] = 0;
        raw[idx + 3] = 0;
        continue;
      }

      // Draw a simple white "T / arrow" motif inside the green badge
      const nx = x / size;
      const ny = y / size;
      const isTopBar = ny >= 0.28 && ny <= 0.40 && nx >= 0.24 && nx <= 0.76;
      const isStem = nx >= 0.44 && nx <= 0.56 && ny >= 0.28 && ny <= 0.75;

      if (isTopBar || isStem) {
        raw[idx] = 255;
        raw[idx + 1] = 255;
        raw[idx + 2] = 255;
        raw[idx + 3] = 255;
      } else {
        // Excel green #107C41
        raw[idx] = 16;
        raw[idx + 1] = 124;
        raw[idx + 2] = 65;
        raw[idx + 3] = 255;
      }
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const compressed = zlib.deflateSync(raw);

  return Buffer.concat([
    signature,
    makeChunk('IHDR', ihdr),
    makeChunk('IDAT', compressed),
    makeChunk('IEND', Buffer.alloc(0))
  ]);
}

fs.mkdirSync(ASSETS_DIR, { recursive: true });
for (const size of [16, 32, 80]) {
  const outPath = path.join(ASSETS_DIR, `icon-${size}.png`);
  fs.writeFileSync(outPath, createBadgePng(size));
  console.log(`Generated ${outPath}`);
}
