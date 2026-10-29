// Structural alignment check between the checked original and its translation.
// Usage: node align-check.mjs <original> <translation> [--report out.json] [--translator-prefix "Text…"]... [--ratio-min 0.5] [--ratio-max 2]
//   <original>    original-blocks.json from export-original.ps1, or the original's project JSON (book-docx-tools format)
//   <translation> the translation's project JSON, or a single batch JSON with "blocks"
// Exit code 1 = structural errors (missing/extra/reordered blocks, type or table mismatch, empty or stale translation).
import fs from 'node:fs';
import path from 'node:path';

const args = process.argv.slice(2);
const opt = { report: null, prefixes: [], ratioMin: 0.5, ratioMax: 2 };
const files = [];
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === '--report') opt.report = args[++i];
  else if (a === '--translator-prefix') opt.prefixes.push(args[++i]);
  else if (a === '--ratio-min') opt.ratioMin = Number(args[++i]);
  else if (a === '--ratio-max') opt.ratioMax = Number(args[++i]);
  else files.push(a);
}
if (files.length !== 2) {
  console.error('Usage: node align-check.mjs <original> <translation> [--report out.json] [--translator-prefix text]...');
  process.exit(2);
}
const DEFAULT_PREFIXES = ['Originál: PDF strany', 'Poznámka překladatele', 'Česká legenda', 'České popisky', 'Český popisek', 'Překlad původního seznamu'];
const prefixes = opt.prefixes.length ? opt.prefixes : DEFAULT_PREFIXES;

const readJson = (f) => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
const pad = (n) => String(n).padStart(4, '0');

function load(file) {
  const full = path.resolve(file);
  const j = readJson(full);
  if (Array.isArray(j.batches)) {
    const dir = path.dirname(full);
    const out = [];
    for (const b of j.batches) {
      const bf = path.isAbsolute(b) ? b : path.join(dir, b);
      const data = readJson(bf);
      for (const blk of data.blocks) out.push({ ...blk, _batch: path.basename(bf), _bookmark: `${j.bookmark_prefix}_b${pad(out.length + 1)}` });
    }
    return { kind: 'project', blocks: out };
  }
  if (Array.isArray(j.blocks)) return { kind: j.source_docx ? 'export' : 'batch', blocks: j.blocks.map((b, i) => ({ ...b, _batch: path.basename(full), _bookmark: b.src ?? `#${i + 1}` })) };
  throw new Error(`${file}: neither a project (batches) nor a block list (blocks).`);
}

const orig = load(files[0]);
const tran = load(files[1]);
const origBlocks = orig.blocks.map((b) => ({ ...b, src: b.src ?? b._bookmark }));
const isTranslatorBlock = (b) => b.origin === 'translator' || (typeof b.text === 'string' && prefixes.some((p) => b.text.startsWith(p)));
const tranBlocks = tran.blocks.filter((b) => !isTranslatorBlock(b));
const skipped = tran.blocks.length - tranBlocks.length;

const errors = [];
const warnings = [];
const err = (msg, t, o) => errors.push({ msg, translation: t?._bookmark, batch: t?._batch, original: o?.src });
const warn = (msg, t, o) => warnings.push({ msg, translation: t?._bookmark, batch: t?._batch, original: o?.src });

const pairs = [];
const bySrc = tranBlocks.some((b) => b.src);
if (bySrc) {
  const index = new Map(origBlocks.map((b, i) => [b.src, i]));
  const seen = new Set();
  let last = -1;
  for (const t of tranBlocks) {
    if (!t.src) { err('Translation block without "src" (mark translator additions with origin:"translator")', t); continue; }
    const i = index.get(t.src);
    if (i === undefined) { err(`Unknown src ${t.src}`, t); continue; }
    if (seen.has(t.src)) err(`Duplicate translation of ${t.src}`, t);
    if (i < last) err(`Order differs: ${t.src} comes after ${origBlocks[last].src}`, t);
    seen.add(t.src); last = Math.max(last, i);
    pairs.push([origBlocks[i], t]);
  }
  for (const o of origBlocks) if (!seen.has(o.src)) err(`Original block ${o.src} has no translation`, null, o);
} else {
  const n = Math.min(origBlocks.length, tranBlocks.length);
  for (let i = 0; i < n; i++) pairs.push([origBlocks[i], tranBlocks[i]]);
  if (origBlocks.length !== tranBlocks.length) err(`Block count differs: original ${origBlocks.length}, translation ${tranBlocks.length} (without ${skipped} translator blocks)`);
}

const textOf = (b) => (b.type === 'table' ? (b.rows ?? []).flat().join(' ') : String(b.text ?? ''));
const ratios = [];
let firstTypeMismatch = null;
for (const [o, t] of pairs) {
  if (o.type !== t.type) {
    err(`Type differs: original ${o.type}, translation ${t.type}`, t, o);
    if (!firstTypeMismatch) firstTypeMismatch = { original: o.src, translation: t._bookmark, original_text: textOf(o).slice(0, 80), translation_text: textOf(t).slice(0, 80) };
    continue;
  }
  if (o.type === 'table') {
    const or = o.rows ?? [], tr = t.rows ?? [];
    if (or.length !== tr.length || (or[0]?.length ?? 0) !== (tr[0]?.length ?? 0)) err(`Table size differs: ${or.length}x${or[0]?.length} vs ${tr.length}x${tr[0]?.length}`, t, o);
    else if (tr.flat().every((c) => !String(c).trim()) && or.flat().some((c) => String(c).trim())) err('Table not translated (all cells empty)', t, o);
    if (t.orig_rows && JSON.stringify(t.orig_rows) !== JSON.stringify(or)) err('Stale: the original table changed after export', t, o);
  } else {
    const ot = String(o.text ?? ''), tt = String(t.text ?? '');
    if (!tt.trim() && ot.trim()) err('Empty translation', t, o);
    if (typeof t.orig === 'string' && t.orig !== ot) err('Stale: the original text changed after export; retranslate this block', t, o);
    if (ot.length >= 20 && tt === ot && !/^[\d\s.,;:()+\-=×·/]+$/.test(ot)) warn('Translation identical to the original (untranslated?)', t, o);
    if (['p', 'note', 'quote'].includes(o.type) && ot.length >= 60 && tt.length) ratios.push({ r: tt.length / ot.length, o, t });
  }
  const op = o.source_pages, tp = t.source_pages;
  if (Array.isArray(op) && Array.isArray(tp) && op.length && tp.length) {
    const oMin = Math.min(...op), oMax = Math.max(...op), tMin = Math.min(...tp), tMax = Math.max(...tp);
    if (tMax < oMin || tMin > oMax) err(`source_pages differ: original ${op.join(',')}, translation ${tp.join(',')}`, t, o);
  }
  if (o.type === 'image' && o.file && t.file && path.basename(o.file) !== path.basename(t.file)) warn(`Image file differs: ${o.file} vs ${t.file}`, t, o);
}
let median = null;
if (ratios.length) {
  const sorted = ratios.map((x) => x.r).sort((a, b) => a - b);
  median = sorted[Math.floor(sorted.length / 2)];
  for (const { r, o, t } of ratios) {
    if (r < median * opt.ratioMin || r > median * opt.ratioMax) warn(`Length ratio ${r.toFixed(2)} (median ${median.toFixed(2)}): possible omission, merge or addition`, t, o);
  }
}

const result = {
  original: path.resolve(files[0]), translation: path.resolve(files[1]), matched_by: bySrc ? 'src' : 'order',
  original_blocks: origBlocks.length, translation_blocks: tran.blocks.length, translator_blocks_skipped: skipped,
  pairs: pairs.length, median_length_ratio: median, first_type_mismatch: firstTypeMismatch,
  errors, warnings,
};
if (opt.report) fs.writeFileSync(opt.report, JSON.stringify(result, null, 2), 'utf8');
console.log(`Alignment (${result.matched_by}): ${origBlocks.length} original blocks, ${tran.blocks.length} translation blocks (${skipped} translator additions skipped), ${pairs.length} pairs, median length ratio ${median?.toFixed(2) ?? '-'}; ${errors.length} errors, ${warnings.length} warnings.`);
for (const e of errors.slice(0, 30)) console.log(`  ERROR: ${e.msg} [${e.original ?? ''} -> ${e.translation ?? ''} ${e.batch ?? ''}]`);
if (errors.length > 30) console.log(`  … ${errors.length - 30} more errors (see --report)`);
for (const w of warnings.slice(0, 15)) console.log(`  WARNING: ${w.msg} [${w.original ?? ''} -> ${w.translation ?? ''} ${w.batch ?? ''}]`);
if (warnings.length > 15) console.log(`  … ${warnings.length - 15} more warnings (see --report)`);
process.exit(errors.length ? 1 : 0);
