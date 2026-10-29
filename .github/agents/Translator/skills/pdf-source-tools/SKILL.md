---
name: pdf-source-tools
description: Lokální nástroje Translator agenta pro práci se zdrojovým PDF – extrakce textové vrstvy, segmentace řádků do odstavců, vykreslení stránek, výřezy obrázků a vizuální ověřovací archy podezřelých slov. Použij ve fázi 1 při tvorbě zkontrolovaného originálu ve Wordu.
---

# pdf-source-tools

Node.js skripty (ESM) nad `pdfjs-dist` a `@napi-rs/canvas`. Závislosti jsou nainstalované v `node_modules` této složky; pokud chybí, spusť zde `npm install` (registr může vyžadovat přihlášení – pak zkopíruj ověřené `node_modules` stejné verze). Spouštěj z libovolné složky, cesty k PDF zadávej relativně k aktuálnímu adresáři nebo absolutně.

Čísla stran jsou vždy **pořadí PDF stran od 1**, nikoli tištěné číslování. Rozsahy: `5`, `1-20`, `3,7,9-12`.

Zdrojové PDF žádný skript nemění. Výstupní obrázky se nepřepisují bez `--overwrite`.

## Skripty

| Skript | Použití | Výstup |
|---|---|---|
| `extract-text.mjs` | `node extract-text.mjs <pdf> <out-dir> [--pages 1-20] [--words]` | `pNNN.txt` po stranách, `pages.json` (řádky s polohou, s `--words` i slova), `summary.json` (SHA256, počet stran, strany bez textové vrstvy → nutné OCR, metadata) |
| `lines.mjs` | `node lines.mjs <pdf> --pages 40-45 [--para] [--skip "^(RUNNING HEAD)$"] [--indent 6] [--json out.json]` | řádky s odsazením `¶+dx`; s `--para` odstavce: vynechá samotná čísla stran a záhlaví podle `--skip`, spojí rozdělená slova, `cont` = odstavec pokračuje z předchozí strany |
| `render.mjs` | `node render.mjs <pdf> <out-dir> <pages> [--scale 1.6] [--overwrite]` | `pNNN.png` s bílým pozadím |
| `crop.mjs` | `node crop.mjs <crops.json> [--overwrite] [--sheet contact.png]` | výřezy obrázků z vykreslených stran; spec `{ "pages_dir", "out_dir", "crops": [ {"name","page","x","y","w","h"} ] }`, cesty relativně ke spec souboru, souřadnice v pixelech vykreslené strany |
| `proof-sheet.mjs` | `node proof-sheet.mjs <pdf> <queries.json> <out-prefix> [--scale 2.2] [--per-sheet 14] [--context 1]` | archy `<prefix>-01.png…` s výřezem řádku, červeným rámečkem kolem slova a OCR textem; `<prefix>.json` se souřadnicemi a nenalezenými dotazy |

`queries.json` pro proof-sheet: `[ {"id":"b0123-1","page":44,"word":"wvast","occurrence":1,"note":"spell"} ]`. Párování ignoruje interpunkci a velikost písmen a bere i slovo, které hledaný řetězec obsahuje; při více výskytech na straně zadej `occurrence`.

## Pravidla

1. Textová vrstva PDF je jen návrh. Rozhoduje obraz stránky. Ověřeno na Davsonovi: OCR „wvast“ = „vast“, „ecale“ = „calculation“, uvozovky `‘‘` = `“`.
2. Každé slovo označené kontrolou pravopisu (viz `book-docx-tools\check-spelling.ps1`) a každé podezřelé místo (cizí znaky, `|`, rozbité uvozovky, číslice uvnitř slov, spojovník na konci řádku) ověř na proof-sheetu a archy si skutečně prohlédni (`view`).
3. Rovnice, tabulky, popisky obrázků, nadpisy a začátky a konce kapitol přepisuj podle vykreslené strany, ne podle extrakce.
4. Strany bez textové vrstvy (`summary.json`) vyžadují OCR nebo ruční přepis podle obrazu. Nedostupné OCR přiznej; nevymýšlej text.
5. Zkontroluj duplicitní a chybějící strany (porovnej tištěná čísla stran s pořadím PDF) a zapiš je do poznámky originálu.
6. Obrázky po výřezu vizuálně zkontroluj (ostrost, celé popisky, žádné ořezy).
