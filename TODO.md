# xlsxjs — TODO

What's still open. The "Resolved in fork" block at the bottom tracks
features that have shipped on `master` and live in the harness / fixtures.
Last reconciled 2026-05-04 after Wave 11 (SmartArt orgchart + cycle,
chart radar / doughnut / data labels, OLE CFB reader).

## Open — medium items (one slice each)

_Empty — all medium slices shipped in Wave 7._

## Open — big projects (library-scale)

- [ ] **Chart rendering (further)** — Waves 9-11 ship column / bar / line /
  pie / scatter / area / radar / doughnut and stacked + percentStacked
  variants, plus data-label annotations. Still open: 3D, stock, bubble,
  surface, dual axes, trendlines, scatter smoothing, stacked line,
  concentric doughnut rings, per-chart `<c:numFmt>` for labels, and every
  chartEx `cx:chartSpace` variant (treemap / sunburst / waterfall / funnel /
  pareto / box-whisker / histogram / map).
- [ ] **SmartArt layouts (further)** — Waves 10-11 ship hierarchy,
  orgchart (branch-rail connectors), and cycle (circular arc arrows)
  via `smartArtLayout: 'svg' | 'both'`. Still open: matrix / radial /
  bracketed / process / Venn / pyramid layouts from `layout1.xml`,
  children-of-children in cycle diagrams, per-layout colour themes from
  `colors1.xml`, per-node font scaling.
- [ ] **Pivot table interactivity** — filtering, grouping, drill-down UI.
  Materialised values already render correctly; this is interaction on top.
- [ ] **Slicer / timeline re-materialisation** — Wave 10 ships opt-in
  chips + dual-handle slider (`interactiveSlicers: true`) that fire
  `xlsx:slicer-change` / `xlsx:timeline-change` CustomEvents. Still open:
  actually re-running the pivot data through the new filter server-side or
  in-browser when the event fires.
- [ ] **OLE payload rendering** — Wave 11 ships `parseCfb` + opt-in
  `parseOleCfb` that extracts named streams from CFB `.bin` payloads
  (kind='ole'). `Sheet.embeddings[i].cfb` surfaces root CLSID + streams
  with a 256-stream / 16 MiB cap. Still open: actually rendering the
  stream contents — the WordDocument BIFF table, Excel's Workbook stream,
  Equation Editor's MTEF byte format, etc. Those are each their own
  library-scale project.
- [ ] **Form-control interactivity (further)** — Wave 9 ships opt-in
  live widgets (`interactiveFormControls`). Still open: sheet-prefixed
  linkedCell refs (`Sheet2!$A$1`), macro-assigned button click dispatch, and
  the legacy VML fallback for producers that write controls via
  `xl/drawings/vmlDrawing*.vml` without a ctrlProp part.
- [ ] **Encryption: actually decrypt** — password-protected `.xlsx` are OLE
  CFB containers. Today we detect them and throw. Decryption needs SHA-512
  + AES-CBC (+ RC4-40 for legacy) and an OLE CFB reader. Real project;
  should live behind an optional dependency if we ever do it.

## Resolved in fork

Most recent first (Wave 11 landed 2026-05-04). Earlier groupings blurred
together in the interest of a readable tail.

### Wave 11 (extensions II, 2026-05-04)
- ✅ **SmartArt orgchart + cycle layouts** — `renderSmartArtSvg` now
  dispatches on `model.layout` (or explicit `opts.layout`). Orgchart
  paints 3-segment branch-rail connectors (vertical drop + horizontal
  rail + per-child vertical); cycle arranges nodes evenly around a circle
  with curved `<path>` arrows using the `A` arc command, plus a shared
  `<marker id="xlsx-smartart-arrow">` arrowhead. Fixture:
  `tests/render-test/smartart-layouts/`.
- ✅ **Chart radar + doughnut + data labels** — `parseChart` adds
  `'radar'` + `'doughnut'` kinds and `ChartSeries.dataLabels`. Radar emits
  concentric polygon gridlines + radial axis lines + series polygons at
  30% fill opacity; doughnut paints each slice as a `<path>` with two
  arc commands (outer + inner reverse). Data labels reach the DOM as
  `<text class="xlsx-chart-data-label">` positioned per `<c:dLblPos>`
  (`outEnd` / `ctr` / `inBase` / `t` / `b`); chart-level `<c:dLbls>`
  inherits to per-series. Fixture: `tests/render-test/radar-labels/`.
- ✅ **OLE CFB reader** — new `src/ole-cfb.ts` with `parseCfb(bytes)`
  decoding the 512-byte header, DIFAT-chained FAT, directory stream, and
  the mini-FAT for small streams. Surfaces root CLSID + named streams
  with preservation of `\x01` / `\x05` control prefixes. Guardrails:
  ≤256 streams, ≤16 MiB/stream, reject buffers > 32 MiB. Detection and
  parsing are opt-in via `Options.parseOleCfb` (default false, Wave 8
  byte-stable). Fixture: `tests/render-test/ole-cfb/`.

### Wave 10 (extensions, 2026-05-04)
- ✅ **Chart scatter / area / stacked variants** — `ChartModel.kind` adds
  `'scatter'` + `'area'`; `ChartModel.grouping` (`'standard' | 'stacked' |
  'percentStacked' | null`) captures the `<c:grouping>` element on
  bar / column / line / area. `ChartSeries.xValues` surfaces scatter's
  `<c:xVal>`. Renderer paints stacked/percentStacked bar+column, scatter
  circles (markers-only), and area paths (closed-fill at 70% opacity,
  stacking when grouping demands). Fixture: `tests/render-test/charts-ext/`.
  Still deferred: 3D / stock / surface / radar / chartEx variants,
  scatter smoothing, stacked line.
- ✅ **Interactive slicers + timeline** — `Options.interactiveSlicers`
  (default off) swaps the detect-only slicer `<ul>` for a row of
  `<button class="xlsx-slicer-chip" aria-pressed>` toggle chips (one per
  cache item, not just selected). `SheetSlicer.allItems` now surfaces the
  full item list; `SheetTimeline.bounds` exposes the cache's
  `<state><bounds/>` date range so the interactive renderer can emit a
  `<div class="xlsx-timeline-slider">` with two `<input type="range">`
  handles. Chip clicks / handle drags fire `xlsx:slicer-change` and
  `xlsx:timeline-change` CustomEvents on the aside; xlsxjs does not
  re-materialise the pivot.
- ✅ **SmartArt hierarchy SVG** — `renderSmartArtSvg(model)` paints the
  tree as a top-down hierarchy: rounded-rect nodes sized ~100×40 with
  a depth-indexed palette, centred labels, straight `<line>` connectors.
  `Options.smartArtLayout` adds `'tree'` (default, byte-stable),
  `'svg'` (SVG only), and `'both'` (SVG then `<ul>`). Orgchart and
  cycle layouts currently fall back to the hierarchy renderer; native
  variants to come.

### Wave 9 (first-class rendering + interactivity, 2026-05-04)
- ✅ **Classic chart rendering** — `parseChart` (`src/chart-parser.ts`) +
  `renderChart` (`src/chart-renderer.ts`). Column (clustered), horizontal
  bar, line (with point markers), pie (incl. doughnut). ChartModel
  surfaces on `SheetChart.model` with typed categories / series /
  values / colour / title / legend position. Fixture: `tests/render-test/charts/`.
  Gated CSS keeps existing snapshots byte-stable. Falls back to the
  dashed placeholder for scatter / area / chartEx / unknown chart types
  (and whenever `Options.renderCharts: false`).
- ✅ **SmartArt tree detection** — `parseSmartArt`
  (`src/smartart-parser.ts`) builds a hierarchy from `xl/diagrams/data1.xml`
  + `layout1.xml`. `Sheet.smartArt: SheetSmartArt[]` surfaces the parsed
  tree plus the layout uniqueId. Renderer emits
  `<aside class="xlsx-smartart">` with a nested `<ul>` where each `<li>`
  carries `data-level` matching its depth. 32-hop cycle guard for
  malformed diagrams. Actual diagram-shape geometry is still deferred.
- ✅ **Interactive form controls** — `Options.interactiveFormControls`
  (default off) swaps the detect-only `<aside>` body for real inputs:
  `<input type="checkbox">` / `type="radio"` / `type="range"` /
  `type="number"` / `<select>` / `<button>`. Change + input events route
  through the exported `applyFormControlUpdate(container, linkedCell, value)`
  helper which writes the new value into the linked cell's `<td>`.
  Default-off guarantees Wave 8 snapshots remain byte-stable. Sheet-
  prefixed `linkedCell` refs log a console.warn + skip (future wave).

### Wave 8 (detect-only slices, 2026-05-04)
- ✅ **Form-control detection** — `Sheet.formControls: SheetFormControl[]`
  surfaces kind (button / checkbox / radio / combo / list / scrollbar /
  spinner / groupBox / label / dialog), label, `linkedCell`, `inputRange`,
  `checked`, `min` / `max` / `inc` / `page` / `val`, `dropLines`, `altText`.
  The renderer emits `<aside class="xlsx-form-control">` inside the
  image-layer with data attributes + ☐/☑/○/● glyphs for boolean controls
  and a `<legend>` for group boxes. No live widget — consumers hydrate
  against the DOM.
- ✅ **Slicer + timeline detection** — `Sheet.slicers: SheetSlicer[]` and
  `Sheet.timelines: SheetTimeline[]`. Slicer fields: name, caption,
  cache, sourceName, columnCount, style, showCaption, rowHeight,
  selectedItems. Timeline fields: name, caption, cache, sourceName,
  level (years/quarters/months/days), selectedRange, showHeader /
  showSelectionLabel / showTimeLevel / showHorizontalScrollbar, style.
  Renderer emits `<aside class="xlsx-slicer">` + `<aside class="xlsx-timeline">`
  after the table with caption headers and selected-item `<li>`s.
- ✅ **OLE / package embedding detection** — `Sheet.embeddings: SheetEmbedding[]`
  surfaces kind ('ole' | 'package'), contentType, fileName, progId, anchor
  coords, size, altText. Binaries stay raw by default; `inlineEmbeddings: true`
  in Options projects the payload through `bytesToDataUrl` (32 MiB cap) with
  a MIME allowlist (pdf / xlsx / docx / pptx / txt / images), emitting an
  `<a download>` child. OLE `.bin` CFB payloads are never inlined (outside
  the allowlist) even with the flag on — raw CFB bytes never reach a
  `data:` URL.

### Wave 7 (medium items, 2026-05-04)
- ✅ **Golden HTML diff mode** — `scripts/golden-diff.mjs` captures each
  fixture's rendered `<section class="xlsx">` as `tests/render-test/*/result.html`;
  `npm run test:golden` diffs against it, `npm run test:golden:capture` rewrites.
  37 snapshots in the harness; normalised for date stamps in headers/footers.
- ✅ **In-flow image positioning** — images now render as `position: absolute`
  `<figure>`s inside a zero-height `.xlsx-image-layer` that sits directly
  above the `<table>`. twoCell / oneCell anchors compute CSS `left` / `top`
  by summing declared column widths + row heights (with Excel's 8.43-char /
  15-pt defaults for undeclared dimensions + a ~30px gutter estimate for
  the row-number column); absolute anchors keep their EMU pixel offsets.
- ✅ **Shape preset geometry rendering** — `src/shape-presets.ts` maps
  26 prstGeom presets (rect, roundRect, ellipse, line, triangle, rtTriangle,
  diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, star5,
  rightArrow / leftArrow / upArrow / downArrow / leftRightArrow,
  flowChartProcess / Decision / Terminator / Connector, callout1,
  wedgeRectCallout, wedgeEllipseCallout, cloudCallout) to inline SVG.
  Unknown presets fall through to the plain bordered `<aside>`.
- ✅ **Full expression-rule interpretation** — `conditional-format.ts`
  `evalExpression` now handles AND / OR / NOT combinators, ISNUMBER /
  ISBLANK / ISERROR predicates, SEARCH (case-insensitive), MOD(ROW()) /
  MOD(COLUMN()) banding, ISEVEN / ISODD, LEFT / RIGHT equality, plus
  the pre-existing narrow cellRef-op-literal form. Unknown formulas
  silently return false.
- ✅ **Security review** — `SECURITY_REVIEW.md` covers 12 attack surfaces;
  4 findings fixed in the same pass: control-char bypass in
  `isSafeHyperlinkHref`, unbounded range expansion in hyperlinks +
  data-validations (capped at 1,048,576 cells), unbounded media inlining
  (32 MiB cap + MIME allowlist in `sanitizeMediaMime`).

### Wave 6 (small items, 2026-05-04)
- ✅ **Diagonal borders** — `BorderStyle.diagonal` + `diagonalUp` /
  `diagonalDown`; renderer paints stacked `linear-gradient` overlays
  (`to bottom right` / `to top right`) on the cell's `background-image`.
- ✅ **Gradient fills** — `FillStyle` is a discriminated union
  (`pattern` / `gradient` / `none`); `<gradientFill>` parses into
  `{ kind: 'gradient', type, degree, stops: [{position, color}] }` and
  renders as `linear-gradient(<degree>deg, …)`. Path gradients parse
  but render as a flat fallback to the first stop's colour.
- ✅ **Non-solid pattern fills** — `darkGray` / `mediumGray` / `lightGray` /
  `dark*` / `light*` directional stripes + `*Grid` / `*Trellis` render
  as `repeating-linear-gradient` approximations on top of `bgColor`.
  `gray125` is still intentionally ignored (Excel's default).
- ✅ **Sheet protection state** — `Sheet.protection` (SheetProtection)
  exposes enabled + passwordHashed + per-operation toggles; the
  renderer sets `data-sheet-protected="true"` on the section.
- ✅ **Alt text on tables** — `TableDef.altText` +
  `TableDef.altTextSummary`; the caption's `aria-label` picks up
  altText (summary falls back to `title`).
- ✅ **Cell metadata / dynamic-array spill** — `xl/metadata.xml` parses
  into `Workbook.metadata`; `Cell.cellMetadataIndex`, `valueMetadataIndex`,
  `isSpillAnchor` carry the resolved wiring; renderer tags spill-anchor
  cells with `.xlsx-spill-anchor`.
- ✅ **Multi-step cellStyle inheritance** — `resolveEffectiveXf` walks
  up to 8 hops with a cycle guard when a named style references another
  named style. One-hop callers unaffected.
- ✅ **Custom icon-set rule lists** — `iconSet/@custom='1'` + `cfIcon`
  children parsed onto `IconSet.customIcons`; renderer swaps in the
  per-position override (each position can pull from a different set).

### Wave 5 (drawings / theme / i18n)
- ✅ **Shapes / connectors / text boxes** — `Sheet.shapes`, `<aside
  class="xlsx-shape">` with `data-kind` / `data-preset`.
- ✅ **Drawing anchors** — proper `oneCellAnchor` / `absoluteAnchor` +
  decorative flag (alt="" + aria-hidden), `emuToPx` helper.
- ✅ **Theme font scheme + phonetics** — `Theme.majorFont` /
  `Theme.minorFont`, `FontStyle.scheme`, `<rPh>` → HTML `<ruby><rt>`.

### Wave 4 (number formats r2 + page layout)
- ✅ **Number formats round 2** — colour modifiers (`[Red]` etc.),
  conditional section predicates (`[>100]…`), fractions (`# ?/?`,
  `# ??/??`), scientific + engineering notation.
- ✅ **Page layout** — row/column manual page breaks, print area
  (from `_xlnm.Print_Area`), header/footer zones with `&D`/`&P`/`&A`
  substitution, `FrozenPanes.kind: 'frozen' | 'split'`.

### Wave 3 (conditional formatting expansion)
- ✅ **CF new rule types** — `containsBlanks` / `notContainsBlanks`,
  `containsErrors` / `notContainsErrors`, `aboveAverage` / `belowAverage`
  (with `stdDev` + `equalAverage`), `timePeriod`.
- ✅ **CF graphical ext attrs** — dataBar `negativeFillColor`,
  `axisPosition`, `border`, `borderColor`; dxf `strike` + `numFmtCode`
  application.

### Wave 2 (view / content / structure)
- ✅ **Sheet view state** — hidden / veryHidden suppression, tab colour,
  RTL direction, gridline / header toggles, zoom.
- ✅ **Hyperlinks + data validation** — URL-scheme allowlist via
  `isSafeHyperlinkHref`; `<dataValidation type="list">` surfaces ▾ with
  pinned options.
- ✅ **Outlines + defined names** — row/column `outlineLevel`,
  `SheetOutline`, `Workbook.definedNames` (incl. `_xlnm.*`).

### Wave 1 (formatting correctness)
- ✅ **Alignment flags** — `wrapText`, `shrinkToFit`, `indent`,
  `textRotation` (incl. 255 stacked), `readingOrder`, widened
  horizontal / vertical enums.
- ✅ **Font extras** — strike, sub/super, underline variants
  (single / double / accounting), sanitised font family via
  `sanitizeFontFamily`.
- ✅ **Number formats round 1** — elapsed time (`[h]` / `[mm]` / `[ss]`),
  accounting padding (`_<char>`), fill character (`*<char>` stripped),
  locale currency (`[$€-2]` etc.), 1904 date system.

### Earlier (session 1 / baseline)
- ✅ **Tiny wins** — encrypted-file detection, `gray125` filter,
  `calcChain`/`printerSettings` filtered from the dropped-parts tally.
- ✅ **Classic comments** (`xl/comments*.xml`) — `Sheet.comments`,
  inline `●` marker.
- ✅ **Threaded comments** (`xl/threadedComments/` + `xl/persons/`) —
  `Sheet.threadedComments`, `Workbook.persons`, inline `💬` marker.
- ✅ **Charts + pivots detection** — `Sheet.charts` / `Sheet.pivots` /
  `Sheet.extensions` (URI roll-up); charts render as a placeholder.
- ✅ **Indexed colours** — ECMA-376 64-entry palette via
  `indexedColor(i)`; `ColorRef` gained an `indexed` variant.
- ✅ **Multi-sheet + workbook rels** — `<sheet r:id>` resolution via
  `xl/_rels/workbook.xml.rels`.
- ✅ **Named styles inheritance** (single hop) — `cellXf.xfId` →
  `cellStyleXfs[j]` via `resolveEffectiveXf`. Extended to a cycle-safe
  chain walk in Wave 6.
- ✅ **Row heights + hidden rows/columns** — `RowDimension`,
  `ColumnWidth.hidden`, `tr.style.height` + `display: none`.
- ✅ **Graphical CF rules** — colourScale (2/3-stop interpolation),
  dataBar (gradient), iconSet (4 built-in sets; custom overrides
  added in Wave 6).
- ✅ **Cached formula values** — `<f>` surfaces on `Cell.formula`;
  cached `<v>` is the displayed value.
- ✅ **Baseline render surface** — merged cells, column widths, frozen
  panes, autoFilter, tables, images, styles.xml fonts / fills / borders,
  rich text, dxf-driven CF, R1C1 notation, iconSet SVGs.

### Tooling
- ✅ **Playwright browser harness** — 35 tests, real Chrome on :3002,
  one render spec per fixture plus a library-surface smoke.
