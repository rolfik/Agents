import fs from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

export async function openPdf(file) {
  const task = getDocument({ data: new Uint8Array(fs.readFileSync(file)), useSystemFonts: true });
  const pdf = await task.promise;
  return { pdf, close: () => task.destroy() };
}

// Accepts "5", "1-20", "3,7,9-12". Numbers are 1-based PDF page order, not printed folios.
export function parsePages(spec, max) {
  const pages = [];
  for (const part of String(spec ?? `1-${max}`).split(',').map(s => s.trim()).filter(Boolean)) {
    const [a, b] = part.split('-').map(Number);
    const last = Number.isFinite(b) ? b : a;
    if (!Number.isInteger(a) || !Number.isInteger(last) || a < 1 || last > max || last < a) throw new Error(`Invalid page range: ${part} (document has ${max} pages)`);
    for (let n = a; n <= last; n++) pages.push(n);
  }
  return pages;
}

export function option(args, name, fallback) {
  const i = args.indexOf(name);
  if (i < 0) return fallback;
  const value = args[i + 1];
  args.splice(i, 2);
  return value;
}

export function flag(args, name) {
  const i = args.indexOf(name);
  if (i < 0) return false;
  args.splice(i, 1);
  return true;
}

// Groups positioned text items into visual lines (top to bottom) with word boxes in PDF units.
export async function pageLines(page) {
  const content = await page.getTextContent();
  // Scanned pages are often slightly rotated: correct the baseline by the median text slope (b/a of the item matrix).
  const slopes = content.items.filter(it => it.str && it.str.trim() && Math.abs(it.transform[0]) > 0.1).map(it => it.transform[1] / it.transform[0]).sort((a, b) => a - b);
  const skew = slopes.length ? slopes[Math.floor(slopes.length / 2)] : 0;
  const lines = [];
  for (const it of content.items) {
    if (!it.str || !it.str.trim()) continue;
    const x = it.transform[4], y = it.transform[5] - (Math.abs(skew) < 0.2 ? skew * it.transform[4] : 0);
    const h = Math.abs(it.transform[3]) || it.height || 10;
    let line = lines.find(l => Math.abs(l.y - y) < Math.max(3, h * 0.35));
    if (!line) { line = { y, h, words: [] }; lines.push(line); }
    line.h = Math.max(line.h, h);
    line.words.push({ s: it.str, x, y, ry: it.transform[5], w: it.width, h });
  }
  lines.sort((a, b) => b.y - a.y);
  for (const l of lines) {
    l.words.sort((a, b) => a.x - b.x);
    l.x = l.words[0].x;
    l.right = Math.max(...l.words.map(w => w.x + w.w));
    l.text = l.words.map(w => w.s).join(' ').replace(/\s+/g, ' ').trim();
  }
  return lines;
}

export function dominantMargin(lines) {
  const counts = {};
  for (const l of lines) {
    const x = Math.round(l.x);
    for (let d = -3; d <= 3; d++) counts[x + d] = (counts[x + d] || 0) + 1;
  }
  return Number(Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0] ?? 0);
}
