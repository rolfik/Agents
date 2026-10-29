// Usage: node render.mjs <input.pdf> <out-dir> <pages: 5 | 1-20 | 3,7,9-12> [--scale 1.6] [--overwrite]
// Writes page-N.png (N = 1-based PDF page order). Existing images are preserved unless --overwrite.
import fs from 'node:fs';
import path from 'node:path';
import { createCanvas } from '@napi-rs/canvas';
import { openPdf, parsePages, option, flag } from './lib.mjs';

const args = process.argv.slice(2);
const scale = Number(option(args, '--scale', '1.6'));
const overwrite = flag(args, '--overwrite');
const [input, outDir, spec] = args;
if (!input || !outDir || !spec) throw new Error('Usage: node render.mjs <input.pdf> <out-dir> <pages> [--scale 1.6] [--overwrite]');
fs.mkdirSync(outDir, { recursive: true });
const { pdf, close } = await openPdf(input);
try {
  for (const n of parsePages(spec, pdf.numPages)) {
    const file = path.join(outDir, `page-${n}.png`);
    if (fs.existsSync(file) && !overwrite) { console.log(`${n}: exists, kept`); continue; }
    const page = await pdf.getPage(n);
    const viewport = page.getViewport({ scale });
    const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    fs.writeFileSync(file, canvas.toBuffer('image/png'));
    console.log(`${n}: ${canvas.width}x${canvas.height}`);
    page.cleanup();
  }
} finally {
  await close();
}
