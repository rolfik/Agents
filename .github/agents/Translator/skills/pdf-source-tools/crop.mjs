// Usage: node crop.mjs <crops.json> [--overwrite] [--sheet contact.png]
// crops.json (paths relative to the JSON file):
// { "pages_dir": "pages", "out_dir": "figures",
//   "crops": [ { "name": "fig-1", "page": 55, "x": 120, "y": 1030, "w": 950, "h": 435 } ] }
// x/y/w/h are pixels of page-N.png rendered by render.mjs (note its --scale).
// --sheet writes one labelled contact sheet of all crops for visual verification.
import fs from 'node:fs';
import path from 'node:path';
import { createCanvas, loadImage } from '@napi-rs/canvas';
import { option, flag } from './lib.mjs';

const args = process.argv.slice(2);
const sheet = option(args, '--sheet');
const overwrite = flag(args, '--overwrite');
const [specFile] = args;
if (!specFile) throw new Error('Usage: node crop.mjs <crops.json> [--overwrite] [--sheet contact.png]');
const base = path.dirname(path.resolve(specFile));
const spec = JSON.parse(fs.readFileSync(specFile, 'utf8'));
const pagesDir = path.resolve(base, spec.pages_dir);
const outDir = path.resolve(base, spec.out_dir);
fs.mkdirSync(outDir, { recursive: true });
const made = [];
for (const c of spec.crops) {
  const file = path.join(outDir, `${c.name}.png`);
  if (fs.existsSync(file) && !overwrite) throw new Error(`Preserve existing image: ${file}`);
  const image = await loadImage(path.join(pagesDir, `page-${c.page}.png`));
  if (c.x < 0 || c.y < 0 || c.x + c.w > image.width || c.y + c.h > image.height) throw new Error(`Crop outside page image: ${c.name}`);
  const canvas = createCanvas(c.w, c.h);
  canvas.getContext('2d').drawImage(image, c.x, c.y, c.w, c.h, 0, 0, c.w, c.h);
  fs.writeFileSync(file, canvas.toBuffer('image/png'));
  made.push({ ...c, file });
  console.log(`${c.name}: page ${c.page}, ${c.w}x${c.h} -> ${file}`);
}
if (sheet) {
  const width = 1000;
  const items = [];
  for (const m of made) {
    const img = await loadImage(m.file);
    const k = Math.min(1, (width - 20) / img.width);
    items.push({ m, img, w: img.width * k, h: img.height * k });
  }
  const height = items.reduce((s, i) => s + i.h + 40, 10);
  const canvas = createCanvas(width, Math.ceil(height));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, canvas.height);
  let y = 10;
  for (const i of items) {
    ctx.fillStyle = '#b00000'; ctx.font = '18px Arial';
    ctx.fillText(`${i.m.name} — PDF page ${i.m.page}`, 10, y + 20);
    ctx.drawImage(i.img, 10, y + 28, i.w, i.h);
    ctx.strokeStyle = '#999999'; ctx.strokeRect(10, y + 28, i.w, i.h);
    y += i.h + 40;
  }
  fs.writeFileSync(sheet, canvas.toBuffer('image/png'));
  console.log(`sheet: ${sheet}`);
}
