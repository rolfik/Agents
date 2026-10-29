// Usage: node lines.mjs <input.pdf> --pages 40-45 [--para] [--skip "<regex>"] [--json out.json]
// Lines mode prints every visual line with its indent relative to the page's dominant left margin
// ("¶+24" marks a probable paragraph start). Paragraph mode (--para) joins lines into paragraphs,
// drops lines matching --skip (running heads) and bare page numbers, and de-hyphenates "word-" line ends.
// Paragraphs continue across pages when the next page starts without indentation (flag "cont").
// The output is a draft for the original-language source; every result still needs visual proofing.
import fs from 'node:fs';
import { openPdf, parsePages, option, flag, pageLines, dominantMargin } from './lib.mjs';

const args = process.argv.slice(2);
const pagesSpec = option(args, '--pages');
const skip = option(args, '--skip');
const jsonOut = option(args, '--json');
const indent = Number(option(args, '--indent', '6'));
const para = flag(args, '--para');
const [input] = args;
if (!input) throw new Error('Usage: node lines.mjs <input.pdf> --pages 1-5 [--para] [--skip regex] [--json out.json]');
const skipRe = skip ? new RegExp(skip) : null;
const { pdf, close } = await openPdf(input);
const paragraphs = [];
try {
  for (const n of parsePages(pagesSpec, pdf.numPages)) {
    const page = await pdf.getPage(n);
    const lines = await pageLines(page);
    const margin = dominantMargin(lines);
    const rights = lines.map(l => l.right).sort((a, b) => a - b);
    const rightEdge = rights[Math.floor(rights.length * 0.8)] ?? 0;
    if (!para) console.log(`=== PDF ${n} (margin ${margin}) ===`);
    let first = true;
    for (const l of lines) {
      const dx = Math.round(l.x - margin);
      if (!para) { console.log(`${dx > indent ? '¶+' + dx : '   '} ${l.text}`); continue; }
      if (/^\d{1,4}$/.test(l.text) || (skipRe && skipRe.test(l.text))) continue;
      const prev = paragraphs[paragraphs.length - 1];
      // A line ending well before the right edge (heading, last line) closes its paragraph.
      const closed = l.right < rightEdge - (rightEdge - margin) * 0.15;
      if (!prev || dx > indent || prev.closed) {
        paragraphs.push({ pages: [n], indent: dx, text: l.text, cont: false, closed });
      } else {
        if (first) prev.cont = true;
        if (!prev.pages.includes(n)) prev.pages.push(n);
        prev.text = /[A-Za-z\u00C0-\u024F]-$/.test(prev.text) ? prev.text.slice(0, -1) + l.text : prev.text + ' ' + l.text;
        prev.closed = closed;
      }
      first = false;
    }
    page.cleanup();
  }
} finally {
  await close();
}
if (para) {
  for (const p of paragraphs) console.log(`[${p.pages.join(',')}${p.cont ? ' cont' : ''}${p.indent > indent ? ' +' + p.indent : ''}] ${p.text}`);
  if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify(paragraphs.map(({ closed, ...p }) => p), null, 1), 'utf8');
}
