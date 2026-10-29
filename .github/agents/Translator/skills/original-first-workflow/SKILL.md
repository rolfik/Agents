---
name: original-first-workflow
description: Povinný dvoufázový postup Translator agenta – nejdřív z dodaného souboru vytvořit a vizuálně ověřit Word originál v původním jazyce, pak překládat výhradně z tohoto Wordu (bez opakovaného skenování PDF) a strukturálně ověřit, že překlad odpovídá originálu blok po bloku. Obsahuje export koster překladu a kontrolu zarovnání.
---

# original-first-workflow

Výsledkem každé překladatelské zakázky jsou **dva Word soubory**:

| Soubor | Obsah | Jazyk kontroly pravopisu |
|---|---|---|
| `<Původní název> - <rok> - originál.docx` | zkontrolovaná předloha v původním jazyce, ověřená proti obrazu stránek | jazyk originálu (např. en-GB 2057) |
| `<Český název> - <rok> - překlad.docx` | překlad sestavený z originálu blok po bloku | cs-CZ 1029 |

Názvy lze přizpůsobit zadání. Oba soubory musí být ve stejné složce jako vstup, pokud uživatel neurčí jinak.

Nástroje:

- `..\pdf-source-tools`: extrakce, segmentace, vykreslení, výřezy obrázků a proof-sheety (fáze 1);
- `..\book-docx-tools`: sestavení, doplňování, audit, kontrola pravopisu a export do PDF (obě fáze);
- tato složka: `export-original.ps1` (předání z fáze 1 do fáze 2) a `align-check.mjs` (kontrola shody).

## Fáze 1 – zkontrolovaný originál

1. `pdf-source-tools\extract-text.mjs` zjistí počet stran, SHA256 a strany bez textové vrstvy. Bez textové vrstvy je nutné OCR nebo ruční přepis podle obrazu.
2. `render.mjs` vykreslí titulní a bibliografické strany, obsah, začátek a konec knihy, strany s obrázky, tabulkami a vzorci. Prohlédni je (`view`) a sestav mapu částí.
3. `lines.mjs --para --skip "<živá záhlaví>"` rozdělí text po dávkách do odstavců. Zapiš je do dávek JSON ve formátu `book-docx-tools` (typy `h1`, `h2`, `h3`, `p`, `note`, `label`, `quote`, `signature`, `image`, `table`) a ke každému bloku uveď `source_pages`.
4. Každou dávku porovnej s vykreslenými stranami:
   - nadpisy, hranice odstavců, verše, poznámky, tabulky (buňka po buňce), vzorce, popisky a pořadí sloupců;
   - opravuj chyby OCR, ne autora;
   - nečitelné místo označ, nevymýšlej ho.
5. Obrázky vyřízni pomocí `crop.mjs` a vizuálně zkontroluj.
6. Sestav originál (`book-docx-tools\build-docx.ps1`) s titulní stranou v původním jazyce:
   - `title`, `subtitle`, `authors`, `edition` přesně podle bibliografické strany;
   - poznámka v `notes` (s `lang` 1029), že jde o přepis originálu ověřený proti obrazu stránek, včetně známých vad předlohy (např. duplicitní nebo chybějící strany).
7. `check-spelling.ps1` → každé podezřelé slovo ověř na `proof-sheet.mjs` a archy si prohlédni. Opravy zapisuj do dávek, nebo (u hotového dokumentu) přímo do Wordu.
8. `audit-docx.ps1` (0 problémů) a `export-pdf.ps1` + `render.mjs`:
   - vizuálně porovnej vzorek stran exportu s odpovídajícími stranami originálu;
   - vzorek vždy zahrnuje titulní strany, obsah, začátky kapitol, tabulky, obrázky, vzorce a poslední strany.
9. Zapiš do průběžného stavu, které strany byly vizuálně ověřeny. Neověřené části nevydávej za ověřené.

**Uživatel může originál ve Wordu dále opravovat.** Od této chvíle je autoritativní aktuální Word originálu, ne pracovní JSON dávky fáze 1. Záložky bloků (`<prefix>_bNNNN`) nesmí zmizet. Nový odstavec bez záložky ohlásí export jako varování; takový text je nutné přidat do bloku nebo doplnit do dávek a originál znovu sestavit.

## Předání – export-original.ps1

```powershell
& export-original.ps1 -Project <originál-projekt.json> -OutDir <pracovní-složka> -Prefix cz [-BatchSize 40] [-Force]
```

- Otevře dočasnou kopii **aktuálního** originálu jen pro čtení ve vlastní instanci Wordu a přečte každý blok podle jeho záložky, tedy včetně uživatelových oprav.
- Zapíše `original-blocks.json`: `source_docx`, `source_sha256`, `assets_dir` a bloky se `src` (záložka originálu), `type`, `text`/`rows`, `file`, `source_pages` a `new_page`.
- Zapíše kostry dávek překladu `cz-01.json…`:
  - každý blok má `src`, `orig` (nebo `orig_rows`) a prázdné `text` (nebo prázdné `rows` stejného rozměru);
  - `status: "todo"`, `source_sha256` = hash originálu;
  - dávky se dělí na hranicích `h1`/`h2`.
- Existující kostry nikdy nepřepisuje. `original-blocks.json` přepíše jen s `-Force`.
- Varuje před textem mimo záložky bloků. Pokud originál chybí nebo se během exportu změní, skončí chybou.

## Fáze 2 – překlad z originálu

1. Projekt překladu (formát `book-docx-tools`) obsahuje:
   - `source_file` = Word originálu a `source_sha256` z exportu;
   - `assets_dir` = `assets_dir` z exportu (obrázky originálu se použijí znovu);
   - `language_id` 1029 a `original_language_id`;
   - `batches` = vyplněné kostry;
   - titulní stranu podle pravidel agenta (český název, čára, původní název, autoři; na nové straně údaje vydání a `ai_note`).
2. Překládej **jen z `orig`/`orig_rows`** v kostrách. PDF už znovu neskenuj. Výjimkou je pochybnost o samotném originálu: pak oprav originál (fáze 1) a export zopakuj.
3. Vyplň `text`/`rows`, `orig` ponech beze změny a `status` nastav na `translated` (po revizi `checked`).
   - Sestavení odmítne dávku se stavem `todo`.
   - Blok, který v originálu není (např. „Poznámka překladatele“, česká legenda k obrázku), vlož s `"origin": "translator"` a bez `src`.
4. Před každým sestavením a na konci spusť kontrolu zarovnání:

```powershell
node align-check.mjs <pracovní-složka>\original-blocks.json <překlad-projekt.json> --report align.json
```

5. Sestavení a doplňování (`build-docx.ps1`, `append-batch.ps1`) ověří, že se Word originálu od exportu nezměnil (`source_sha256`).
   - Pokud se změnil: export zopakuj s `-Force` do stejné složky (kostry se nepřepíšou).
   - Pak spusť `align-check.mjs`: bloky hlášené jako *Stale* přelož znovu a v kostrách aktualizuj `orig`.
   - Nakonec v projektu a dávkách nastav nový `source_sha256`.
6. Kontrola překladu: `audit-docx.ps1`, `check-spelling.ps1`, `export-pdf.ps1` + vizuální kontrola vzorku stran.

## align-check.mjs

```text
node align-check.mjs <original-blocks.json | originál-projekt.json> <překlad-projekt.json | dávka.json>
     [--report out.json] [--translator-prefix "Text"]... [--ratio-min 0.5] [--ratio-max 2]
```

Páruje bloky podle `src`. Když překlad `src` nemá (starší překlady), páruje podle pořadí. Bloky překladatele přeskočí:

- bloky s `origin:"translator"`;
- bloky s textem začínajícím výchozími prefixy „Originál: PDF strany“, „Poznámka překladatele“, „Česká legenda“, „České popisky“, „Český popisek“ a „Překlad původního seznamu“;
- vlastní prefixy přes `--translator-prefix` (nahradí výchozí).

**Chyby (exit 1):**

- chybějící, nadbývající, zdvojený nebo přeházený blok;
- rozdílný počet bloků (párování podle pořadí);
- rozdílný typ bloku (zpráva uvede první neshodu);
- rozdílný rozměr tabulky;
- prázdný překlad;
- nepřekrývající se `source_pages`;
- zastaralý překlad (`orig` ≠ aktuální text originálu).

**Varování:** text shodný s originálem (nepřeložený?), poměr délek mimo `ratio-min`–`ratio-max` násobek mediánu (možné vynechání nebo sloučení) a jiný soubor obrázku.

Shoda struktury nedokazuje věrnost překladu. Obsah odstavců porovnávej s originálem i nadále a varování o délce projdi ručně.

## Ověřeno

Na testovacím projektu (8 bloků: nadpisy, odstavec, obrázek s popiskem, tabulka, verš, poznámka) bylo ověřeno:

- export → vyplnění koster → sestavení → audit: 0 problémů;
- `align-check` hlásí 0 chyb a správně přeskočí blok překladatele;
- odhalí záměnu pořadí, chybějící blok i zastaralý `orig`;
- oprava provedená uživatelem přímo ve Wordu originálu se propíše do nového exportu a dotčený blok překladu je označen jako *Stale*;
- sestavení odmítne kostru se stavem `todo`.
