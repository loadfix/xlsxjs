# xlsxjs — project notes for Claude

Browser-side XLSX→HTML renderer. TypeScript, built with rollup, tested with
Karma+jasmine and a jsdom harness. Sibling to `../docxjs`; the layout and
workflows mirror it so the two projects are interchangeable to work in.

## Scope (current slice)

Open `.xlsx`, parse **all** sheets (binding via `xl/_rels/workbook.xml.rels`,
with positional fallback), render each as an HTML `<section class="xlsx">`
containing column-letter headers, row-number gutter, cell values, merged
cells (horizontal + vertical), column widths + hidden columns from
`<cols>`, row heights + hidden rows from `<row ht>/<row hidden>`, and cell
formatting from `xl/styles.xml` — fonts (bold/italic/underline/size/color),
solid pattern fills, borders (per side), alignment, and number formats
(currency, percent, dates via `yyyy-mm-dd`/`d-mmm-yy`/etc). Named-style
inheritance is honoured: a `cellXfs` entry with `xfId` picks up its base
formatting from `cellStyleXfs` whenever its own `applyX` flags are off.

Colour handling resolves three sources: direct `rgb`, theme index + tint
(via `xl/theme/theme1.xml`, using the Excel HSL tint formula), and the
legacy indexed palette (`<color indexed=…>`, 64-entry table).

**Conditional formatting**: `<dxfs>` from styles.xml pair with
`<conditionalFormatting>` blocks per sheet. Supported dxf-driven rule
types are `cellIs` (all operators), `containsText` / `notContainsText` /
`beginsWith` / `endsWith`, `duplicateValues` / `uniqueValues`, `top10`
(including bottom + percent), and a narrow `expression` form that
matches `=<cellRef> <op> <literal>`. Graphical rules `colorScale` and
`dataBar` render directly: colour scales interpolate between 2- or
3-stop cfvos (with `min` / `max` / `num` / `percent` / `percentile`
thresholds), data bars emit a horizontal linear gradient spanning the
cell. `iconSet` renders as an inline SVG prepended to the cell via a
small palette (3TrafficLights1, 3Arrows, 3Symbols, 3Symbols2); unknown
sets fall through to no icon. Rules are resolved in Excel's priority
order; dxf matches layer on top of the base xf and the cell gains
`.xlsx-cf`, `.xlsx-cf-colorscale`, `.xlsx-cf-databar`, or
`.xlsx-cf-iconset` as appropriate.

**Frozen panes + autoFilter + tables**: per-sheet metadata is surfaced
on `Sheet.frozenPanes`, `Sheet.autoFilter`, and `Sheet.tables`. Frozen
cells receive `.xlsx-frozen-row` / `.xlsx-frozen-col` /
`.xlsx-frozen-both` with CSS `position: sticky`. AutoFilter header
cells gain `.xlsx-autofilter` (which paints a small ▾ indicator).
Tables are parsed from `xl/tables/tableN.xml` via the sheet's rels; a
small `<div class="xlsx-table-caption">` per table is emitted after
the sheet's `<table>` so consumers can see the range/name in the DOM
(table-level styling is intentionally left to the cell-level xf chain).

**Images**: drawings referenced via the sheet's rels are resolved;
the embedded media binary is inlined as a `data:` URL on an `<img>`
wrapped in a `<figure class="xlsx-image">` after the table. Anchor
coordinates + offsets are surfaced on the figure as data-attributes so
callers who want real in-flow positioning can overlay using those.

**R1C1 notation**: `showFormulas: true` replaces cell text with the
formula (prefixed with `=`). Set `formulaNotation: 'r1c1'` to render
in Excel's R1C1 form (anchor-relative deltas bracketed; absolutes not
bracketed). `a1ToR1c1` / `r1c1ToA1` are also exposed as pure helpers.

Shared strings and the common `c/@t` types are handled (`s`, `inlineStr`,
`n`, `b`, `str`, `e`). Rich-text runs
(`<si><r><rPr>…</rPr><t>…</t></r></si>`) render as one `<span>` per run
with per-run font properties; plain strings follow the simple
`textContent` path. Formulas surface their `<f>` text on the Cell model;
the cached `<v>` (when present) is used as the display value. xlsxjs
does not evaluate formulas — a cell with no cached value renders blank.

Deliberately deferred: chart rendering, drawings beyond raster images
(shapes / connectors / SmartArt), diagonal / double borders, in-flow
image positioning (images render after the table, not overlaid on the
cell grid), full expression-rule interpretation beyond the narrow form.

### Excel width → pixel conversion

`src/html-renderer.ts` uses a simplified `7·width + 5` formula (Calibri
11pt, MDW=7). This matches Excel's rounded pixel output within ±1px
across common widths. The exact ECMA-376 formula is documented in the
same file; swap it in when/if styles.xml (and therefore the actual sheet
font) is parsed.

## Build & test

- `npm run build` — dev UMD bundle (`dist/xlsx-preview.js`).
- `npm run build-prod` — also emits `.mjs` and minified variants.
- `npm run test:render` — jsdom harness (`scripts/test-render.mjs`). Builds
  the UMD first, then loads the `basic` fixture through the library.
- `npm run e2e` — Karma suite against real Chrome.
- `node scripts/make-fixture.mjs` — regenerate the `basic` fixture XLSX.
  Run this if the fixture shape needs to change; do not hand-edit the
  binary.

### Keep README.md and TODO.md current

Whenever a feature is added, removed, or a public option changes, update both of these files *in the same PR* as the code change — stale docs have bitten us before.

- **`README.md`** — the API block reflects the real public surface. If you add/remove a function or option, add/remove the matching entry. If you add or remove an export, reflect it in the API section. Any prose sections (Status, Contributing, project-specific sections) should also match reality.
- **`TODO.md`** — if the change resolves a tracked issue, move that entry into a "Resolved in fork" / "Done" section with a one-line description and the PR/commit reference. Update any counts table at the top and bump the "last updated" date.

Minimum check before every PR that touches source: `grep -n "<feature name>" README.md TODO.md` to catch stale references.

## Architecture

- `src/workbook.ts` — zip open + part extraction. Hands XML strings to the parser.
- `src/workbook-parser.ts` — XLSX XML → `{ sheets: Sheet[], styles }` model.
- `src/styles.ts` — `xl/styles.xml` → typed styles (`Styles`, `CellXf`, fonts/fills/borders/numFmts). Colour elements become `ColorRef` (rgb or theme+tint) via `parseColorElement`; rgb values pass through `sanitizeHexColor` before they can reach the DOM.
- `src/theme.ts` — `xl/theme/theme1.xml` → 12-entry colour table in SpreadsheetML index order (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accent1..6, 10 hlink, 11 folHlink). `resolveColor(ref, theme)` hands back the final `#rrggbb`, applying the Excel HSL tint formula for theme+tint and the legacy 64-entry palette (`indexedColor(index)`) for `<color indexed=…>` refs.
- `src/conditional-format.ts` — parses `<conditionalFormatting>` blocks into `ConditionalFormatting[]` and evaluates rules. `evaluateRule(rule, cell, range, ctx)` covers cellIs, containsText / begins- / endsWith, duplicate/unique, top10, and a narrow expression form. `ctx.cellsInRange` handles rules that need to see the whole range (duplicates, top10).
- `src/number-format.ts` — Excel format-code → display text. Handles `General`, thousands + decimals, percent, currency, date/time, `@`, and `positive;negative;zero` sections. Colour modifiers (`[Red]`, `[Blue]`) and conditional sections are parsed but dropped.
- `src/html-renderer.ts` — model → `<section class="xlsx">` with `<table>`. Resolves each cell's XF, applies fonts/fills/borders via inline styles (never class-name interpolation), and formats numeric values through `number-format.ts`.
- `src/xlsx-preview.ts` — public API: `parseAsync`, `renderWorkbook`, `renderAsync`, `defaultOptions`, plus test-visible helpers `formatNumber`, `parseStyles`, `sanitizeHexColor`.
- `src/utils.ts` — column-letter math and cell-ref parsing.

## Known future work

- Charts (`xl/charts/*.xml`, both classic `c:chartSpace` and chartEx).
- Drawings beyond raster images (shapes, connectors, SmartArt).
- In-flow image positioning (currently images render after the table).
- Gradient fills and diagonal / double borders.
- Full expression-rule interpretation (currently only `=<cellRef> <op> <literal>`).
- Full multi-cellStyle inheritance (e.g. cellStyle referencing another cellStyle).
- Drawing layout for one-cell + absolute anchors (only twoCellAnchor sizing is fully honoured).

## Fixture generation

- `tests/render-test/basic/` + `tests/render-test/merged/` come from
  `scripts/make-fixture.mjs` (Node / JSZip).
- `tests/render-test/python-xlsx/` comes from the sibling `python-xlsx`
  library at `~/code/python-xlsx`. Regenerate with
  `~/code/python-xlsx/.venv/bin/python scripts/make-python-xlsx-fixture.py`.
  This is the fixture that exercises real-world styles + number formats,
  so treat it as the canary for regressions in that code path.
- `tests/render-test/conditional-format/` comes from python-xlsx as well:
  `~/code/python-xlsx/.venv/bin/python scripts/make-cf-fixture.py`.
  Adjust the python script (not the binary) to evolve the cf fixture.
- `tests/render-test/cf-graphical/` exercises the graphical cf rules
  (colorScale + dataBar). Regenerate with
  `~/code/python-xlsx/.venv/bin/python scripts/make-cf-graphical-fixture.py`.
- `tests/render-test/cf-icons/` — iconSet cf. Regenerate with
  `~/code/python-xlsx/.venv/bin/python scripts/make-icons-fixture.py`.
- `tests/render-test/tables/` — frozen panes + autoFilter + named table.
  `~/code/python-xlsx/.venv/bin/python scripts/make-tables-fixture.py`.
- `tests/render-test/image/` — one embedded PNG anchored at B3.
  `~/code/python-xlsx/.venv/bin/python scripts/make-image-fixture.py`
  (PIL required).

## Security constraints (inherited pattern)

All XLSX content is attacker-controlled. Follow docxjs's rules:

- Never interpolate XLSX strings into a CSS class, selector, or `innerHTML`.
- Keyed maps on XLSX-derived strings use `Map` or validate against a strict
  regex.
- Attribute values set via `setAttribute` or `dataset.*` are safe.
- Cell text goes through `textContent`, never `innerHTML`.
