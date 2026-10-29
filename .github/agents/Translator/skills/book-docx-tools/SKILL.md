---
name: book-docx-tools
description: Sestavení, bezpečné doplňování, audit, kontrola pravopisu a export knižních Word dokumentů (zkontrolovaný originál i překlad) z pracovních JSON dávek přes vlastní instanci Word COM. Použij ve fázi 1 pro originál a ve fázi 2 pro překlad.
---

# book-docx-tools

PowerShell 7 skripty nad modulem `BookDocx.psm1`, které ovládají Microsoft Word přes COM. Každý skript:

- spouští **vlastní skrytou instanci Wordu**, ověřuje `Documents.Count == 0` a zavírá jen to, co sám otevřel;
- nikdy nepřepisuje existující výstup naslepo: pracuje na kopii `<config>-working.docx`, po uložení dokument znovu otevře a ověří a teprve pak výsledek přesune na místo výstupu;
- zapisuje JSON zprávu (`<config>-validation.json`, případně `-Report`).

Spouštěj přes `& <cesta>\skript.ps1 ...`. Cesty v konfiguraci jsou relativní ke složce konfigurace. Při chybě nemaž naslepo zálohy ani pracovní kopie; nejdřív zjisti příčinu.

## Konfigurace projektu (JSON)

Vzor je v `project.example.json`. Originál a překlad mají každý **vlastní** konfiguraci, výstup, `bookmark_prefix` a dávky.

| Pole | Povinné | Význam |
|---|---|---|
| `output` | ano | výsledný .docx |
| `language_id` | ano | jazyk textu dokumentu (1029 čeština, 2057 en-GB, 1033 en-US, 1031 němčina…) |
| `original_language_id` | u překladu | jazyk originálu pro původní název, autora a údaje o vydání |
| `bookmark_prefix` | ano | písmena/číslice, max. 16 znaků; bloky dostanou záložky `<prefix>_b0001…` |
| `style_prefix` | ano | předpona vlastních stylů (`<prefix>Text`, `Note`, `Caption`…) |
| `batches` | ano | seznam dávek v pořadí knihy |
| `source_file`, `source_sha256` | doporučené | zdroj (PDF u originálu, zkontrolovaný originál .docx u překladu) a jeho hash; dávka s jiným `source_sha256` je odmítnuta |
| `assets_dir` | ne | kořen pro `file` obrázků (výchozí složka konfigurace) |
| `font`, `font_size`, `image_width_cm` | ne | výchozí Cambria, 11 b., 12 cm |
| `toc_title`, `toc_levels` | ne | „Obsah“, `1-2`; u anglického originálu „Contents“ |
| `page_break_before` | ne | typy bloků začínající na nové stránce, výchozí `["h1","h2"]` |
| `title_page` | ano | viz níže |

### Titulní strany (`title_page`)

Povinné pořadí podle pokynů agenta:

Úprava odpovídá vzoru schválenému uživatelem; vše je vystředěné:

1. **Strana 1:** `title` (styl Název) a nepovinně `subtitle` (styl Podtitul) v jazyce dokumentu, vodorovná čára; u překladu pak `original_title` (Název) a `original_subtitle` (Podtitul) v jazyce originálu; prázdný řádek; `authors` (Podtitul, jeden autor na položku, jen jméno) a nepovinně `author_details` (běžný text, např. tituly a členství „M.I.MECH.E., M.INST.PET.“).
2. **Strana 2 (nová stránka):** `edition` jsou údaje o vydání převzaté z originálu (rok, nakladatel, adresa, tisk…), po řádcích; prázdný řetězec `""` vloží prázdný řádek pro oddělení skupin (první položka nesmí být prázdná). Nic nevymýšlej. U překladu pod nimi po prázdném řádku `ai_note`, přesně ve znění **„Automatický překlad pomocí <název AI modelu> z originálu.“**, kde je skutečný název použitého modelu nebo modelů (např. „GPT-6 Astra a Claude Opus 5.5“). Zástupný text v `<…>` build odmítne. Volitelně pak po prázdném řádku `notes` drobným písmem (řetězec nebo `{id, text, lang}`; `id` vytvoří záložku `<prefix>_<id>`, např. `scope` pro oznámení o rozsahu).
3. **Strana 3:** klikací obsah, pole `TOC \o "1-2" \h \z`.

Přítomnost `original_title` určuje, že jde o překlad; pak je `ai_note` povinné. U originálu se `original_title` ani `ai_note` nevyplňují.

## Dávky a bloky

```json
{ "source_pages": [12, 15], "source_sha256": "…", "status": "checked",
  "blocks": [ { "type": "p", "text": "…", "source_pages": [12], "src": "orig_b0042" } ] }
```

| `type` | Pole | Sazba |
|---|---|---|
| `h1`, `h2`, `h3` | `text` | Nadpis 1–3 (vstupují do obsahu podle `toc_levels`) |
| `p` | `text` | odstavec do bloku; `\n` vytvoří zalomení řádku (verše, adresy), odstavec se pak zarovná vlevo |
| `note` | `text` | menší poznámka (autorova poznámka pod čarou, překladatelská poznámka) |
| `label` | `text` | tučný štítek držený s dalším odstavcem |
| `quote` | `text` | kurzívní citát na střed |
| `signature` | `text` | podpis vpravo |
| `image` | `file`, `text` (popisek), nepovinně `width_cm` | obrázek se zachovaným poměrem stran a popiskem pod ním |
| `table` | `rows` (obdélníkové pole), nepovinně `col_widths_cm` | editovatelná tabulka, záhlaví se opakuje, řádky se nedělí |

`status` dávky: `todo` (nevyplněná kostra z `export-original.ps1` – sestavení i doplnění ji odmítne), `translated`, `checked`. Pole `orig`/`orig_rows` v kostře překladu nesou text originálu pro `align-check` a do dokumentu se nevkládají.

Společná nepovinná pole: `source_pages` (pořadí PDF stran originálu), `lang` (odlišný jazyk bloku), `new_page` (`true`/`false` přebíjí `page_break_before`), `src` (záložka odpovídajícího bloku originálu, pro `align-check`), `origin: "translator"` (blok přidaný překladatelem, nemá protějšek v originálu).

## Skripty

| Skript | Použití |
|---|---|
| `build-docx.ps1 -Project p.json [-Preview náhled.pdf]` | Nový dokument: titulní strany, obsah, všechny bloky, zápatí s polem PAGE, ustálení obsahu, ověření, kontrola balíčku. Existující výstup nepřepíše. |
| `append-batch.ps1 -Project p.json [-ScopeText text] [-Preview náhled.pdf]` | Připojí bloky, které v dokumentu ještě nejsou (podle počtu záložek `<prefix>_bNNNN`). Vytvoří zálohu `<výstup>.backup-<čas>.docx`, zachová text před obsahem a všechny dosavadní bloky včetně uživatelových úprav, komentářů a revizí a obnoví sledování změn. Po kontrole hashe nahradí výstup. `-ScopeText` přepíše poznámku se záložkou `<prefix>_scope`. |
| `audit-docx.ps1 -Project p.json [-Docx] [-Report] [-SkipTitlePages] [-Baseline starší.docx]` | Kontrola na dočasné kopii: každý blok proti JSON, tabulky po buňkách, obrázky a popisky, titulní strany, odkazy obsahu (záložka a skutečná stránka), jazyk, XML/ZIP integrita a média. Rozdíly textu hlásí jako uživatelské revize, neopravuje je. Návratový kód 1 = problém. `-SkipTitlePages` slouží pro starší dokumenty bez nové titulní úpravy. |
| `check-spelling.ps1 -Project p.json -Out podezreni.json [-Queries dotazy.json] [-Allow povolena.txt] [-Docx]` | Wordem nalezené pravopisné chyby a regexová podezření (OCR: `\|`, ligatury, číslice ve slově, rozdělená slova, dvojité mezery, neobvyklé znaky). `-Queries` vytvoří vstup pro `pdf-source-tools/proof-sheet.mjs`. `-Allow` je seznam povolených slov (jedno na řádek). |
| `export-pdf.ps1 -Docx d.docx -Pdf d.pdf [-RenderPages 1-5] [-RenderDir png]` | Export PDF se záložkami z nadpisů; volitelně vykreslí stránky PNG přes `pdf-source-tools/render.mjs` k vizuální kontrole. |

## Postup

1. `build-docx.ps1` z první ověřené dávky (nebo ze všech).
2. Pro každou další dávku `append-batch.ps1`; nikdy nepřestavuj dokument, který uživatel upravil.
3. `audit-docx.ps1`, který musí skončit s 0 problémy.
4. U originálu `check-spelling.ps1 -Queries`, pak `proof-sheet.mjs` nad zdrojovým PDF. Archy si prohlédni, oprav dávky (u nového dokumentu) nebo přímo Word (u upraveného) a audit zopakuj.
5. `export-pdf.ps1 -RenderPages …` a prohlédni titulní strany, obsah, začátky kapitol, obrázky, tabulky, vzorce a poslední stránky.

## Známá omezení

- Vyžaduje nainstalovaný Microsoft Word (Windows) a PowerShell 7.
- Zamčený nebo otevřený výstup se nepřepisuje; skript skončí chybou.
- Jazyk symbolů a vzorců může Word hlásit odlišně; audit to vede jako varování, nikoli chybu.
- Po pádu může zůstat proces `WINWORD.EXE /Automation -Embedding`. Ukonči jen ten konkrétní PID (`Stop-Process -Id`) a jen po ověření příkazové řádky.
