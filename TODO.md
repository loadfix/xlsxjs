# xlsxjs — TODO

Everything that the real-world Excel-365 smoke test (see
`scripts/smoke-test.mjs`) surfaced against `/mnt/data/Temp/365/*.xlsx`,
plus pre-existing known gaps. Organised by effort tier, not by visibility.

## Resolved in fork

- **Page layout metadata** (`feat/page-layout`) — manual row/column page
  breaks (`<rowBreaks>`/`<colBreaks>` with `man="1"`) land on
  `Sheet.pageBreaks = { rows, cols }` (0-based indices, automatic breaks
  dropped) and surface as `data-page-break-rows` / `data-page-break-cols`
  on the rendered `<section class="xlsx">`. `_xlnm.Print_Area` defined
  names resolve into cell ranges on `Sheet.printArea: PrintAreaRange[] |
  null` (via `localSheetId`). `<headerFooter>`'s `oddHeader`/`oddFooter`
  split into `{ left, center, right }` zones on `Sheet.headerFooter`; the
  `&D`/`&T`/`&A` substitution codes resolve eagerly (date/time via
  `new Date()`, sheet name from the model), `&P`/`&N` stay literal
  (`"(page)"` / `"(total)"` — xlsxjs doesn't paginate), and `&F` is left
  verbatim. The renderer emits `<div class="xlsx-header">` +
  `<div class="xlsx-footer">` grids after the table, one
  `<div data-zone="left|center|right">` per zone. `FrozenPanes` gained a
  required `kind: 'frozen' | 'split'` field; `state="split"` panes now
  round-trip through the parser and render with the same sticky classes
  as frozen panes. Harness scenarios 57–60; fixture at
  `tests/render-test/page-layout/`.

- **Graphical-CF data-bar ext + dxf strike/numFmt** (`feat/cf-graphical-ext`) —
  `DataBar` now carries the post-2010 ext attributes (`border`,
  `borderColor`, `negativeFillColor`, `negativeBorderColor`,
  `axisPosition`, `axisColor`, `gradient`, `direction`). The parser
  walks the cfRule's `<extLst>/<ext>/<x14:id>` GUID, then splices in
  the matching `<x14:cfRule>` from the sheet-level `<extLst>`. The
  renderer honours the extra fields: negative values pick up the
  negative fill colour, `border` draws a 1px edge in the declared
  border colour, and `axisPosition="middle"` tags the td with
  `data-cf-databar-axis` + an inset box-shadow. `applyDxf` now also
  applies `dxf.font.strike` / `dxf.font.underline` alongside its font
  and re-runs `formatNumber` against the cell when `dxf.numFmtCode`
  is set (non-rich-text, numeric cells only). Covered by render-harness
  scenarios 50–52; fixture at `tests/render-test/cf-ext-databar/`.

- **Alignment flags on `xf/alignment/`**
  (`feat/alignment-flags`) — `wrapText`, `shrinkToFit`, `indent`,
  `textRotation` (including `255` = stacked vertical), widened
  `horizontal` (`justify` / `distributed` / `centerContinuous` /
  `fill`) and `vertical` (`justify` / `distributed`) enums, and
  `readingOrder` (0 context / 1 LTR / 2 RTL) now parse into
  `Alignment` and apply to the td via `applyAlignment`. CSS
  compromises: `shrinkToFit` is an `.xlsx-shrink-to-fit` class
  (browsers have no auto-fit rule); `fill` aligns to start (no pure
  CSS "repeat to fill"); `distributed`/`justify` vertical
  approximate as `middle`; rotated cells use `transform:rotate(…)`
  inside the td, which doesn't grow the cell box.

## Tiny wins (quick fixes)

- [ ] **Detect encrypted files and throw a clear error.** Today
  `encrypted.xlsx` fails with `Can't find end of central directory : is
  this a zip file ?`. Check the first 4 bytes for the OLE CFB magic
  (`D0 CF 11 E0`) before handing the buffer to JSZip and throw an
  `XlsxEncryptedError` with a useful message. Touches `src/workbook.ts`.
- [ ] **Filter `gray125` from the smoke tool's `non-solid fills` warning.**
  Excel writes `<patternFill patternType="gray125"/>` as `fills[1]` on
  every workbook by default; we correctly skip it at render time, but the
  smoke tool flags it on every file and drowns out real signal. Drop it
  from the count. Touches `scripts/smoke-test.mjs`.
- [ ] **Filter render-irrelevant parts from the dropped-parts tally.**
  `calcChain.xml` is a formula dependency graph (not render-relevant);
  `printerSettings/*.bin` is binary print config. Don't flag these as
  "dropped" — they're dropped on purpose. Touches `scripts/smoke-test.mjs`.

## Real features (one slice each)

- [ ] **Classic comments** — `xl/comments*.xml` + `xl/drawings/vmlDrawing*.vml`.
  Parse the `<commentList>/<comment ref="…" authorId="…">` structure plus
  the author table; render as a hoverable annotation (a small marker in
  the corner of the cell with a tooltip containing author + text). VML
  drawing parsing is needed only to know *which* comments should render
  (Excel uses it to hide-by-default); we can ignore VML shape layout.

- [ ] **Threaded comments** — `xl/threadedComments/threadedComment*.xml` +
  `xl/persons/person.xml`. Excel 365's modern comment model: person
  registry + per-comment threads. Same renderer sink as classic
  comments but the data model is cleaner (no VML). The threadedComments
  file carries `{id, ref, personId, text, parentId}`; person.xml maps
  `personId → displayName`.

- [ ] **Detect charts and surface them on the sheet model.** `xl/charts/*.xml`
  (classic `c:chartSpace`) and `xl/charts/*.xml` (chartEx `cx:chartSpace`,
  Excel-2013+ types: treemap, sunburst, waterfall, funnel, histogram,
  pareto, box-whisker, map). Don't try to render (big project); parse
  the anchor + chart type name onto `Sheet.charts[]` so consumers can
  render their own placeholder, and so our render output can emit a
  `<div class="xlsx-chart-placeholder">` where the chart would go.

- [ ] **Detect pivot tables.** `xl/pivotTables/pivotTable*.xml` references
  a cache in `xl/pivotCache/pivotCacheDefinition*.xml` +
  `pivotCacheRecords*.xml`. The sheet *already* renders the materialised
  values correctly (they live in the sheet XML). Surface the range +
  name on `Sheet.pivotTables[]` so consumers know the region is a pivot.

- [ ] **Sheet-level `<extLst>` extensions.** The sheet XML's `extLst`
  is where Excel hides post-2010 features: sparklines (`x14`), data
  validation formulas (`x14`), protected ranges (`x14`), dynamic-array
  spill metadata (`xda`). Today we silently drop every `<ext>` whose
  URI we don't know. At minimum we should enumerate the URIs seen and
  expose them on `Sheet.extensions[]` so the smoke tool can roll them up
  and we can decide which to implement.

## Big projects (library-scale work)

- [ ] **Chart rendering** — two separate schemas. Every stock / line /
  bar / pie chart goes through `c:chartSpace`. The modern types
  (treemap, sunburst, waterfall, funnel, pareto, box-whisker, map)
  live in `cx:chartSpace`. Rendering these properly means implementing
  a DrawingML subset + axis/series layout. Probably its own library
  (the docxjs sibling, "chartjs-xlsx"?). Out of scope for xlsxjs.

- [ ] **Pivot table interactivity** — filtering, grouping, drill-down.
  The materialised values already render; this is interactive UI on
  top. Out of scope for xlsxjs; leave to consumers.

- [ ] **Encryption: actually decrypt.** `msoffcrypto-tool` is the
  reference implementation in Python; a browser-side port would need
  SHA-512 + AES-CBC + RC4-40 support. Real project; put behind an
  optional dependency if we ever do this.

## Known pre-existing gaps (pre-smoke-test)

- [ ] Multi-cellStyle inheritance (cellStyle referencing another cellStyle).
- [ ] Gradient fills.
- [ ] Diagonal / double borders beyond the style name.
- [ ] Full expression-rule interpretation (today only `=<cellRef> <op> <literal>`).
- [ ] iconSet rendering for sets beyond 3TrafficLights1 / 3Arrows / 3Symbols /
  3Symbols2 (we expose the set name; consumers can render their own).
- [ ] In-flow image positioning (images render after the table, not overlaid
  on the cell grid).
- [ ] Drawing one-cell + absolute anchors (only twoCellAnchor sizing is
  fully honoured).

## Smoke-test improvements (to keep the triage signal useful)

- [ ] Add per-file golden HTML output capture + diff mode, `docxjs`-style.
  When the first pass looks right for a given fixture, snapshot it and
  later runs detect regressions. The current smoke tool only catches
  crashes, not silent output drift.
- [ ] Playwright / browser-harness pass. jsdom and Chrome disagree on
  `position: sticky` inside tables, `linear-gradient` on `<td>`, and
  `display: none` on `<col>` — three features xlsxjs renders but only
  verifies in jsdom.
- [ ] Security review. SECURITY_REVIEW.md for xlsxjs in the docxjs
  tradition: XML billion-laughs on the parse path, URL scheme allowlist
  if/when we render hyperlinks, `<img src>` with arbitrary MIME types
  from inline media, CSS identifier validation on any class we
  interpolate from XLSX-derived strings.

## Read-only feature gaps vs Microsoft Excel

*Added 2026-05-04. Items that xlsxjs doesn't render or parse,
ordered by how noticeable their absence is when viewing a real workbook. Anything
already tracked elsewhere in this file is excluded.*

### Cell content
- **Hyperlinks** (`worksheet/hyperlinks/hyperlink` + sheet rels of type `/hyperlink`) — cell-anchored links (web, bookmark, email); today the formatted text renders but never becomes clickable.
- **Data validation dropdowns** (`worksheet/dataValidations/dataValidation[@type='list']`) — list-source values pinned to a cell; Excel paints a ▾ affordance even in read-only view.
- **Cell metadata / dynamic-array spill** (`xl/metadata.xml` + `c/@cm`/`@vm`) — marks cells that carry rich data types or are part of a spilled array; without it linked data types render as plain strings.

### Cell formatting
- **Strikethrough** (`font/strike`) — struck-out text; common in change tracking and finished-task lists.
- **Subscript / superscript** (`font/vertAlign` = `subscript`|`superscript`) — scientific and chemistry sheets rely on this for legibility.
- **Underline variants** (`font/u/@val` = `double`|`singleAccounting`|`doubleAccounting`) — accounting totals use the double accounting underline by convention; today every `<u>` renders as a single underline.
- **Font family / face** (`font/name/@val`) — deliberately dropped over CSS-injection concerns; the sheet loses its chosen face (e.g. monospace code fonts, Cambria headings).
- **Non-solid pattern fills** (`fill/patternFill/@patternType` = `darkGray`/`lightGray`/`darkHorizontal`/`lightVertical`/`darkGrid`/`darkTrellis`/…) — only `solid` and `gray125` are recognised; every other pattern drops the fill entirely.
- **Diagonal borders** (`border/diagonal` + `@diagonalUp`/`@diagonalDown`) — the crossed-out-cell convention; parser reads only the four orthogonal sides.

### Number formats
- **Fraction formats** (`# ?/?`, `# ??/??` — numFmt IDs 12/13 and custom) — displayed via the generic numeric path today, so "0.25" shows as `0.25` instead of `1/4`.
- **Conditional section formats** (`[>100]#,##0;[Red]-#,##0`) — the bracketed conditional predicate is stripped and the first section is always used regardless of value.
- **Colour modifiers in format codes** (`[Red]`, `[Blue]`, `[Color 14]`) — parsed as brackets and dropped; the intended per-section colour never reaches the td.
- **Scientific notation** (numFmt IDs 11/48 and custom `0.00E+00`) — the E+NN exponent syntax is not expanded; numbers render via the generic numeric path.

## Resolved in fork

- **Elapsed-time markers** (`[h]`/`[hh]`/`[m]`/`[mm]`/`[s]`/`[ss]`) — `formatDateTime`
  now detects elapsed tokens before the bracket-strip pass and substitutes the
  accumulated hours/minutes/seconds (serial × 24/1440/86400). Non-bracketed
  tokens after the elapsed marker keep their modulo wall-clock meaning, so
  `[h]:mm` on 1.5 renders "36:00". Harness scenario 34.
- **Accounting padding** (`_<char>`) — `renderLiteralsAroundNumber` now emits
  a single ASCII space per `_<char>` sequence instead of dropping both
  characters silently, so `_( _) _-` forms align. Harness scenario 35.
- **Fill character** (`*<char>`) — the sequence is now silently stripped
  (previously the `*` surfaced as a literal in the output). Documented as a
  compromise in the file header of `number-format.ts`: a web renderer can't
  measure remaining column width without a layout pass. Harness scenario 35.
- **Locale currency symbols** (`[$€-2]`, `[$¥-411]`, `[$-409]`) — a new
  `extractLocaleCurrency` pass rewrites `[$<symbol>-<localeHex>]` to a quoted
  literal carrying just the symbol before the generic bracket strip. Empty
  symbols (`[$-409]`) become nothing. Harness scenario 36.
- **1904 date system** — `Workbook.date1904` now lives on the Workbook model,
  populated from `<workbookPr date1904="1"/>` via a new `parseWorkbookMeta`
  reader. `formatNumber` takes an options bag with `date1904` that flips the
  epoch from 1899-12-30 to 1904-01-01 (no leap-bug fudge). `html-renderer.ts`
  threads the flag through per sheet. Harness scenario 37.
- **Sheet visibility + view state** (`feat/sheet-view-state`) — `Sheet.state`
  carries the `<sheet @state>` value (`visible` / `hidden` / `veryHidden`);
  non-visible sheets are skipped by the renderer. A new `Sheet.view`
  (`SheetView`) parses `<sheetView>` attributes `rightToLeft`,
  `showGridLines`, `showRowColHeaders`, `zoomScale`, plus the
  `<sheetPr><tabColor/>` colour. The renderer flips `section.dir="rtl"`,
  adds `.xlsx-no-gridlines` / `.xlsx-no-headers` classes (CSS strips cell
  borders / hides the gutter + header row), sets `section.style.zoom` for
  non-100% zoom, and surfaces the resolved tab colour on
  `data-tab-color`. Harness scenarios 38–40.

### Display & layout
- (no open items — see "Resolved in fork" for recent page-layout work.)

### Conditional formatting
- **`containsBlanks` / `notContainsBlanks`** (`cfRule/@type`) — highlights blank or non-blank cells in a range; returns no-match today.
- **`containsErrors` / `notContainsErrors`** (`cfRule/@type`) — highlights `#DIV/0!`, `#N/A`, etc.; returns no-match today.
- **`aboveAverage` / `belowAverage`** (`cfRule/@type` + `@aboveAverage`/`@equalAverage`/`@stdDev`) — statistical banding; common on scorecard sheets.
- **`timePeriod`** (`cfRule/@type` + `@timePeriod`=`today`|`yesterday`|`thisWeek`|`lastMonth`|…) — schedule-colouring rules used heavily in project plans.
- **Custom icon-set rule lists** (`iconSet/@custom='1'` + `cfIcon` children) — per-threshold icon overrides; xlsxjs uses the default palette for the set name only.

### Objects & drawings
- **Shapes and connectors** (`xdr:sp`, `xdr:cxnSp` in `xl/drawings/drawingN.xml`) — callout arrows, rectangles, text boxes with cell-anchored position; currently only `<xdr:pic>` and `<xdr:graphicFrame>` (chart) are walked.
- **WordArt and SmartArt** (`xdr:sp` with `a:txBody`/`dgm:relIds`) — decorative text and hierarchy/process diagrams embedded as drawings.
- **Text boxes** (`xdr:sp/xdr:txBody`) — free-floating annotations, used heavily for dashboards.
- **Form controls** (`xl/ctrlProps/*.xml` + VML / `xdr:sp` buttons, checkboxes, list boxes) — checked/unchecked state and labels aren't rendered even as static values.
- **Slicers and timelines** (`xl/slicers/slicer*.xml`, `xl/timelines/timeline*.xml`) — the pivot/table filter chips that show current selection state.
- **OLE objects and embedded files** (`xl/embeddings/*.bin` + `xdr:sp`) — embedded PDFs, Word docs, equations render as missing content today.

### Structural
- **Row and column outlines** (`row/@outlineLevel`, `col/@outlineLevel`, `sheetFormatPr/@outlineLevelRow`, `@outlineLevelCol`) — the grouping gutter with +/- handles beside row/column headers; outline level is dropped.
- **Summary row / column position** (`sheetPr/outlinePr/@summaryBelow`, `@summaryRight`) — controls which side the group-summary row sits on for outline rendering.
- **Defined names** (`workbook/definedNames/definedName`) — named ranges surface as cell refs in formulas and inside dropdown sources; not exposed on the Workbook model.
- **Sheet protection state** (`worksheet/sheetProtection`) — read-only status in the sheet tab and locked-cell indicators; not surfaced for consumers.

### Internationalisation & accessibility
- **Phonetic ruby (furigana)** (`si/rPh`, `worksheet/phoneticPr`) — ruby text above Japanese characters; the `<rPh>` runs are dropped when we flatten an `<si>`.
- **Alt text on images** (`xdr:pic/xdr:nvPicPr/xdr:cNvPr/@descr`) — parsed today but also emitted as the `<img alt>` only when `descr` is present; the newer `a:extLst` "decorative" marker is not honoured.
- **Alt text on tables / charts** (`table/@altText`, `table/@altTextSummary`) — accessibility label on defined tables; not exposed on `TableDef`.
- **Major/minor font scheme resolution** (`xl/theme/theme1.xml` `a:fontScheme/a:majorFont`/`a:minorFont`) — styles referencing `@scheme='major'|'minor'` on a font fall back to the default family because we don't read the theme's font pair.
