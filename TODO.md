# xlsxjs — TODO

Everything that the real-world Excel-365 smoke test (see
`scripts/smoke-test.mjs`) surfaced against `/mnt/data/Temp/365/*.xlsx`,
plus pre-existing known gaps. Organised by effort tier, not by visibility.

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
