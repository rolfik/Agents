// Usage: node extract-text.mjs <input.pdf> <out-dir> [--pages 1-20] [--words]
// Writes p001.txt… (one per PDF page), pages.json (lines with x/y and optionally word boxes)
// and summary.json (page count, SHA-256, metadata, pages without a usable text layer).
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { openPdf, parsePages, option, flag, pageLines } from './lib.mjs';

const args = process.argv.slice(2);
const pagesSpec = option(args, '--pages');
const withWords = flag(args, '--words');
const [input, outDir] = args;
if (!input || !outDir) throw new Error('Usage: node extract-text.mjs <input.pdf> <out-dir> [--pages 1-20] [--words]');
fs.mkdirSync(outDir, { recursive: true });
const sha256 = crypto.createHash('sha256').update(fs.readFileSync(input)).digest('hex').toUpperCase();
const { pdf, close } = await openPdf(input);
try {
  const width = String(pdf.numPages).length < 3 ? 3 : String(pdf.numPages).length;
  const pages = [];
  for (const n of parsePages(pagesSpec, pdf.numPages)) {
    const page = await pdf.getPage(n);
    const lines = await pageLines(page);
    const text = lines.map(l => l.text).join('\n');
    fs.writeFileSync(path.join(outDir, `p${String(n).padStart(width, '0')}.txt`), text + '\n', 'utf8');
    pages.push({
      page: n,
      size: page.view.slice(2),
      characters: text.length,
      lines: lines.map(l => ({ text: l.text, x: +l.x.toFixed(1), y: +l.y.toFixed(1), h: +l.h.toFixed(1), ...(withWords ? { words: l.words.map(w => ({ s: w.s, x: +w.x.toFixed(1), y: +w.y.toFixed(1), w: +w.w.toFixed(1), h: +w.h.toFixed(1) })) } : {}) }))
    });
    page.cleanup();
  }
  fs.writeFileSync(path.join(outDir, 'pages.json'), JSON.stringify(pages), 'utf8');
  const meta = await pdf.getMetadata().catch(() => null);
  const summary = {
    source: path.resolve(input),
    source_sha256: sha256,
    pdf_pages: pdf.numPages,
    extracted_pages: pages.length,
    text_characters: pages.reduce((s, p) => s + p.characters, 0),
    pages_without_text_layer: pages.filter(p => p.characters < 30).map(p => p.page),
    metadata: meta ? { info: meta.info, xmp: meta.metadata?.getAll?.() ?? null } : null
  };
  fs.writeFileSync(path.join(outDir, 'summary.json'), JSON.stringify(summary, null, 2), 'utf8');
  console.log(JSON.stringify(summary, null, 2));
} finally {
  await close();
}
