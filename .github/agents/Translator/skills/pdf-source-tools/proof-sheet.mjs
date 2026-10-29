// Usage: node proof-sheet.mjs <input.pdf> <queries.json> <out-prefix> [--scale 2.2] [--per-sheet 14] [--context 1]
// Visual proofing of suspicious source words (OCR errors, spell-check flags, numbers, symbols).
// queries.json: [ { "id": "b0123-1", "page": 44, "word": "wvast", "occurrence": 1, "note": "spell" } ]
//   page = 1-based PDF page; word = OCR text to locate (punctuation-insensitive); occurrence optional (1-based).
//   Instead of page, "pages": [44, 45] searches the pages in order (a paragraph crossing a page break);
//   occurrence then counts across those pages. The page where the word was found is reported back.
// For each query the full text line (plus --context lines above/below) is cut from the rendered page,
// the located word is boxed in red and labelled with its OCR reading. Look at every sheet and decide the
// correct reading from the image, never from the OCR label.
// Writes <out-prefix>-01.png … and <out-prefix>.json (found/not found, sheet and slot of each query).
import fs from 'node:fs';
import { createCanvas } from '@napi-rs/canvas';
import { openPdf, option, pageLines } from './lib.mjs';

const args = process.argv.slice(2);
const scale = Number(option(args, '--scale', '2.2'));
const perSheet = Number(option(args, '--per-sheet', '14'));
const context = Number(option(args, '--context', '1'));
const [input, queryFile, prefix] = args;
if (!input || !queryFile || !prefix) throw new Error('Usage: node proof-sheet.mjs <input.pdf> <queries.json> <out-prefix> [--scale 2.2] [--per-sheet 14] [--context 1]');
const queries = JSON.parse(fs.readFileSync(queryFile, 'utf8'));
const norm = s => s.normalize('NFKC').replace(/[^\p{L}\p{N}]+/gu, '').toLowerCase();
const { pdf, close } = await openPdf(input);
const cache = new Map();
async function pageData(n) {
  if (cache.has(n)) return cache.get(n);
  if (cache.size > 6) cache.delete(cache.keys().next().value);
  const page = await pdf.getPage(n);
  const viewport = page.getViewport({ scale });
  const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  const data = { canvas, viewport, lines: await pageLines(page) };
  page.cleanup();
  cache.set(n, data);
  return data;
}
const results = [];
const tiles = [];
try {
  for (const q of queries) {
    const target = norm(q.word);
    const hits = [];
    for (const pn of (Array.isArray(q.pages) && q.pages.length ? q.pages : [q.page])) {
      const pd = await pageData(pn);
      pd.lines.forEach((l, li) => l.words.forEach(w => { const s = norm(w.s); if (s && (s === target || (target.length > 2 && s.includes(target)))) hits.push({ li, w, pn }); }));
    }
    const hit = hits[(q.occurrence ?? 1) - 1];
    if (!hit) { results.push({ ...q, found: false, candidates: hits.length }); continue; }
    q.page = hit.pn;
    const d = await pageData(hit.pn);
    const lo = Math.max(0, hit.li - context), hi = Math.min(d.lines.length - 1, hit.li + context);
    const band = d.lines.slice(lo, hi + 1);
    const top = Math.min(...band.map(l => d.viewport.convertToViewportPoint(0, l.y + l.h)[1])) - 6;
    const bottom = Math.max(...band.map(l => d.viewport.convertToViewportPoint(0, l.y - l.h * 0.35)[1])) + 6;
    const left = Math.max(0, Math.min(...band.map(l => d.viewport.convertToViewportPoint(l.x, 0)[0])) - 12);
    const right = Math.min(d.canvas.width, Math.max(...band.map(l => d.viewport.convertToViewportPoint(l.right, 0)[0])) + 12);
    const [wx, wy] = d.viewport.convertToViewportPoint(hit.w.x, hit.w.y + hit.w.h);
    const [wx2, wy2] = d.viewport.convertToViewportPoint(hit.w.x + hit.w.w, hit.w.y - hit.w.h * 0.3);
    const crop = { x: Math.floor(left), y: Math.max(0, Math.floor(top)), w: Math.ceil(right - left), h: Math.ceil(bottom - top) };
    const tile = createCanvas(crop.w, crop.h);
    const tctx = tile.getContext('2d');
    tctx.drawImage(d.canvas, crop.x, crop.y, crop.w, crop.h, 0, 0, crop.w, crop.h);
    tctx.strokeStyle = '#e00000'; tctx.lineWidth = 2;
    tctx.strokeRect(wx - crop.x - 2, wy - crop.y - 2, wx2 - wx + 4, wy2 - wy + 4);
    tiles.push({ q, tile, line: d.lines[hit.li].text });
    results.push({ ...q, found: true, line: d.lines[hit.li].text, candidates: hits.length });
  }
} finally {
  await close();
}
const width = 1500;
for (let s = 0; s * perSheet < tiles.length; s++) {
  const group = tiles.slice(s * perSheet, (s + 1) * perSheet).map(t => ({ ...t, k: Math.min(1, (width - 20) / t.tile.width) }));
  const height = group.reduce((sum, t) => sum + t.tile.height * t.k + 34, 10);
  const canvas = createCanvas(width, Math.ceil(height));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, canvas.height);
  let y = 10;
  group.forEach((t, i) => {
    ctx.fillStyle = '#0040a0'; ctx.font = '17px Arial';
    ctx.fillText(`[${i + 1}] ${t.q.id ?? ''}  PDF ${t.q.page}  OCR: "${t.q.word}"${t.q.note ? '  — ' + t.q.note : ''}`, 10, y + 18);
    ctx.drawImage(t.tile, 10, y + 24, t.tile.width * t.k, t.tile.height * t.k);
    y += t.tile.height * t.k + 34;
  });
  const file = `${prefix}-${String(s + 1).padStart(2, '0')}.png`;
  fs.writeFileSync(file, canvas.toBuffer('image/png'));
  group.forEach((t, i) => { const r = results.find(x => x.id === t.q.id && x.page === t.q.page && x.word === t.q.word); if (r) { r.sheet = file; r.slot = i + 1; } });
  console.log(`${file}: ${group.length} items`);
}
fs.writeFileSync(`${prefix}.json`, JSON.stringify(results, null, 1), 'utf8');
const missing = results.filter(r => !r.found);
console.log(`queries ${queries.length}, located ${results.length - missing.length}, not located ${missing.length}`);
for (const m of missing) console.log(`  not located: ${m.id ?? ''} PDF ${m.page} "${m.word}"`);
