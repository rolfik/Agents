// Usage: node proof-ocr.mjs <ocr-dir> <image-dir> <queries.json> <out-prefix> [--per-sheet 12] [--context 1]
// Visual proofing for scanned books without a reliable PDF text layer: locates suspect words through the word
// boxes of an OCR run (ocr-windows.ps1 output: ocr-<page>.json with lines[].words[] {t,x,y,w,h}) and cuts the line
// (plus --context lines) from the rendered page image (<image-dir>/<image> named in the OCR json).
// queries.json: [ { "id": "b0123-1", "pages": [44, 45], "word": "expansioh", "occurrence": 1 } ]
//   word is matched punctuation-insensitively as a prefix of the OCR word (so "com-" finds "com-"); multiword
//   queries match consecutive words. occurrence counts across the listed pages.
// Writes <out-prefix>-01.png … and <out-prefix>.json. Decide every reading from the image, never from the label.
import fs from 'node:fs';
import path from 'node:path';
import { createCanvas, loadImage } from '@napi-rs/canvas';

const args = process.argv.slice(2);
const opt = (n, d) => { const i = args.indexOf(n); if (i < 0) return d; const v = args[i + 1]; args.splice(i, 2); return v; };
const perSheet = Number(opt('--per-sheet', '12'));
const context = Number(opt('--context', '1'));
const [ocrDir, imgDir, queryFile, prefix] = args;
if (!prefix) throw new Error('Usage: node proof-ocr.mjs <ocr-dir> <image-dir> <queries.json> <out-prefix> [--per-sheet 12] [--context 1]');
fs.mkdirSync(path.dirname(path.resolve(prefix)), { recursive: true });
const queries = JSON.parse(fs.readFileSync(queryFile, 'utf8').replace(/^\uFEFF/, ''));
const norm = s => s.normalize('NFKC').replace(/[^\p{L}\p{N}]+/gu, '').toLowerCase();
const ocrCache = new Map(), imgCache = new Map();
const ocr = n => { if (!ocrCache.has(n)) { const f = path.join(ocrDir, `ocr-${n}.json`); ocrCache.set(n, fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, '')) : null); } return ocrCache.get(n); };
async function img(n) {
  if (!imgCache.has(n)) { if (imgCache.size > 4) imgCache.delete(imgCache.keys().next().value); imgCache.set(n, await loadImage(path.join(imgDir, ocr(n).image))); }
  return imgCache.get(n);
}
const strips = [], report = [];
for (const q of queries) {
  const want = String(q.word).split(/\s+/).map(norm).filter(Boolean);
  let left = q.occurrence || 1, hit = null;
  for (const pg of [].concat(q.pages ?? q.page)) {
    const o = ocr(pg); if (!o) continue;
    for (let li = 0; li < o.lines.length && !hit; li++) {
      const ws = o.lines[li].words;
      for (let wi = 0; wi + want.length <= ws.length; wi++) {
        const ok = want.every((w, k) => { const t = norm(ws[wi + k].t); return k === want.length - 1 ? t.startsWith(w) : t === w; });
        if (ok && --left === 0) { hit = { pg, li, words: ws.slice(wi, wi + want.length) }; break; }
      }
    }
    if (hit) break;
  }
  if (!hit) { report.push({ id: q.id, word: q.word, found: false }); continue; }
  const o = ocr(hit.pg), lines = o.lines.slice(Math.max(0, hit.li - context), hit.li + context + 1);
  const y0 = Math.max(0, Math.min(...lines.map(l => l.y)) - 8), y1 = Math.min(o.height, Math.max(...lines.map(l => l.y + l.h)) + 8);
  strips.push({ q, hit, y0, y1, w: o.width });
  report.push({ id: q.id, word: q.word, found: true, page: hit.pg });
}
const label = 34;
for (let s = 0; s * perSheet < strips.length; s++) {
  const part = strips.slice(s * perSheet, (s + 1) * perSheet);
  const W = Math.max(...part.map(p => p.w));
  const H = part.reduce((a, p) => a + p.y1 - p.y0 + label, 0);
  const c = createCanvas(W, H), g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
  let y = 0;
  for (const p of part) {
    g.fillStyle = '#1a3d8f'; g.font = 'bold 24px sans-serif';
    g.fillText(`${p.q.id}  p.${p.hit.pg}  OCR: ${p.q.word}`, 8, y + 26);
    y += label;
    const im = await img(p.hit.pg);
    g.drawImage(im, 0, p.y0, p.w, p.y1 - p.y0, 0, y, p.w, p.y1 - p.y0);
    g.strokeStyle = '#e00'; g.lineWidth = 3;
    const xs = p.hit.words.map(w => w.x), ys = p.hit.words.map(w => w.y);
    const x0 = Math.min(...xs) - 4, x1 = Math.max(...p.hit.words.map(w => w.x + w.w)) + 4;
    const wy0 = Math.min(...ys) - 4, wy1 = Math.max(...p.hit.words.map(w => w.y + w.h)) + 4;
    g.strokeRect(x0, y + wy0 - p.y0, x1 - x0, wy1 - wy0);
    y += p.y1 - p.y0;
    for (const r of report) if (r.id === p.q.id) { r.sheet = s + 1; }
  }
  fs.writeFileSync(`${prefix}-${String(s + 1).padStart(2, '0')}.png`, c.toBuffer('image/png'));
}
fs.writeFileSync(`${prefix}.json`, JSON.stringify(report, null, 1));
console.log(`queries ${queries.length}, found ${report.filter(r => r.found).length}, sheets ${Math.ceil(strips.length / perSheet)}`);
for (const r of report) if (!r.found) console.log('NOT FOUND', r.id, r.word);
