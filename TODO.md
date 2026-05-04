# xlsxjs — TODO

What's still open. The "Resolved in fork" block at the bottom tracks
features that have shipped on `master` and live in the harness / fixtures.
Last reconciled 2026-05-04 after Wave 6 (fills+borders, sheet metadata,
cell metadata, cellStyle chain + custom icons).

## Open — medium items (one slice each)

- [ ] **Golden HTML diff mode for the smoke tool** — docxjs-style. First run
  against a fixture captures a `result.html` snapshot; subsequent runs diff
  the rendered container against it and flag drift. Catches regressions that
  the current "does it crash?" smoke can't see. Touches `scripts/smoke-test.mjs`.
- [ ] **In-flow image positioning** — images render in a `<figure>` after
  the table today. A true overlay would `position: absolute` them over the
  correct cell range, computed from `twoCellAnchor` + the actual table
  layout. Needs a layout measurement pass post-render.
- [ ] **Shape preset geometry rendering** — `Sheet.shapes` models the
  preset name (`rect` / `line` / `flowChartProcess` / …) but the renderer
  emits a plain `<aside>`. Emit an SVG per preset so a text box actually
  looks like a text box instead of a generic bordered card.
- [ ] **Full expression-rule interpretation** — today's narrow form is
  `=<cellRef> <op> <literal>`. Common Excel expression rules also use
  `AND(…)`, `OR(…)`, `SEARCH("…")`, `ISNUMBER(…)`, `MOD(ROW(), 2)=0`
  (banding). A tiny formula interpreter covering those five would unlock
  most real expression-cf cases.
- [ ] **Security review** — write `SECURITY_REVIEW.md` in the docxjs style.
  Audit: XML billion-laughs guard on the parse path, URL-scheme allowlist
  on hyperlinks (we have it, but review), `<img src>` MIME allowlist on
  inlined media (we embed raw media — verify), CSS identifier validation
  on any class interpolated from XLSX-derived strings, `textContent` vs
  `innerHTML` sinks.

## Open — big projects (library-scale)

- [ ] **Chart rendering** — classic `c:chartSpace` (bar/line/pie) +
  chartEx `cx:chartSpace` (treemap / sunburst / waterfall / funnel /
  pareto / box-whisker / histogram / map). Implementing properly means a
  DrawingML subset + axis/series layout. Probably its own library.
- [ ] **SmartArt** (`xl/diagrams/*.xml` — `data1.xml` + `layout1.xml` +
  `quickStyle1.xml`) — hierarchy / process diagrams. Outer shape container
  surfaces today but the graph layout is untouched.
- [ ] **Form controls** (`xl/ctrlProps/*.xml` + VML fallback) — buttons,
  checkboxes, dropdowns. The outer `<xdr:sp>` surfaces as a SheetShape;
  checked/unchecked state and macro labels aren't decoded.
- [ ] **Slicers and timelines** — widgets rendered over pivot tables. We
  detect them (`Sheet.pivots`) but don't paint the filter chip UI.
- [ ] **OLE objects / embedded files** (`xl/embeddings/*.bin`) — embedded
  PDFs, Word docs, equations. Render as missing content today.
- [ ] **Pivot table interactivity** — filtering, grouping, drill-down UI.
  Materialised values already render correctly; this is interaction on top.
- [ ] **Encryption: actually decrypt** — password-protected `.xlsx` are OLE
  CFB containers. Today we detect them and throw. Decryption needs SHA-512
  + AES-CBC (+ RC4-40 for legacy) and an OLE CFB reader. Real project;
  should live behind an optional dependency if we ever do it.

## Resolved in fork

Most recent first (Wave 6 landed 2026-05-04). Earlier groupings blurred
together in the interest of a readable tail.

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
