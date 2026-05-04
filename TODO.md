# xlsxjs — TODO

What's still open. The "Resolved in fork" block at the bottom tracks
features that have shipped on `master` and live in the harness / fixtures.
Cleaned up 2026-05-04 to reflect reality after Waves 1–5.

## Open — small items (one file each)

- [ ] **Sheet protection state** (`<sheetProtection>`) — surface the locked /
  password-protected flag on `Sheet.protection` for consumers who want a
  read-only indicator. No rendering side effect (cells are already read-only).
- [ ] **Alt text on tables** (`table/@altText`, `table/@altTextSummary`) —
  accessibility label on defined tables; extend `TableDef`.
- [ ] **Cell metadata / dynamic-array spill** (`xl/metadata.xml` + `c/@cm`,
  `c/@vm`) — linked data types and spill-range metadata. Without this, rich
  data types render as plain strings. Parse the metadata index and surface
  it on the Cell model so consumers can style spill regions.
- [ ] **Custom icon-set rule lists** (`iconSet/@custom='1'` + `cfIcon`
  children) — per-threshold icon overrides. We honour the default palette
  for the named set but drop per-rule overrides.
- [ ] **Multi-step cellStyle inheritance** — today `resolveEffectiveXf`
  follows `cellXfs[i].xfId → cellStyleXfs[j]` one hop. Excel permits named
  styles to chain (rare but legal); walk the chain until a default or a
  cycle.

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

- ✅ **Tiny wins** (encrypted file detection, `gray125` filter,
  `calcChain`/`printerSettings` filtered from dropped-parts tally).
- ✅ **Classic comments** (`xl/comments*.xml`) — `Sheet.comments`, inline
  `●` marker.
- ✅ **Threaded comments** (`xl/threadedComments/` + `xl/persons/`) —
  `Sheet.threadedComments`, `Workbook.persons`, inline `💬` marker.
- ✅ **Charts + pivots detection** — `Sheet.charts` / `Sheet.pivots` /
  `Sheet.extensions` (URI roll-up); charts render as a placeholder.
- ✅ **Alignment flags** — `wrapText`, `shrinkToFit`, `indent`,
  `textRotation` (incl. 255 stacked), `readingOrder`, widened
  horizontal / vertical enums.
- ✅ **Font extras** — strike, sub/super, underline variants
  (single / double / accounting), sanitised font family via
  `sanitizeFontFamily`.
- ✅ **Number formats round 1** — elapsed time (`[h]` / `[mm]` / `[ss]`),
  accounting padding (`_<char>`), fill character (`*<char>` stripped),
  locale currency (`[$€-2]` etc.), 1904 date system.
- ✅ **Number formats round 2** — colour modifiers (`[Red]` etc.),
  conditional section predicates (`[>100]…`), fractions (`# ?/?`,
  `# ??/??`), scientific + engineering notation.
- ✅ **Sheet view state** — hidden / veryHidden suppression, tab colour,
  RTL direction, gridline / header toggles, zoom.
- ✅ **Hyperlinks + data validation** — URL-scheme allowlist via
  `isSafeHyperlinkHref`; `<dataValidation type="list">` surfaces ▾ with
  pinned options.
- ✅ **Outlines + defined names** — row/column `outlineLevel`,
  `SheetOutline`, `Workbook.definedNames` (incl. `_xlnm.*`).
- ✅ **CF new rule types** — `containsBlanks` / `notContainsBlanks`,
  `containsErrors` / `notContainsErrors`, `aboveAverage` / `belowAverage`
  (with `stdDev` + `equalAverage`), `timePeriod`.
- ✅ **CF graphical ext attrs** — dataBar `negativeFillColor`,
  `axisPosition`, `border`, `borderColor`; dxf `strike` + `numFmtCode`
  application.
- ✅ **Page layout** — row/column manual page breaks, print area
  (from `_xlnm.Print_Area`), header/footer zones with `&D`/`&P`/`&A`
  substitution, `FrozenPanes.kind: 'frozen' | 'split'`.
- ✅ **Shapes / connectors / text boxes** — `Sheet.shapes`, `<aside
  class="xlsx-shape">` with `data-kind` / `data-preset`.
- ✅ **Drawing anchors** — proper `oneCellAnchor` / `absoluteAnchor` +
  decorative flag (alt="" + aria-hidden), `emuToPx` helper.
- ✅ **Diagonal borders** — `BorderStyle.diagonal` + `diagonalUp` /
  `diagonalDown`; renderer paints stacked `linear-gradient` overlays
  (`to bottom right` / `to top right`) on the cell's `background-image`.
- ✅ **Gradient fills** — `FillStyle` is now a discriminated union
  (`pattern` / `gradient` / `none`); `<gradientFill>` parses into
  `{ kind: 'gradient', type, degree, stops: [{position, color}] }` and
  renders as a CSS `linear-gradient(<degree>deg, …)`. Path gradients
  parse but render as a flat fallback to the first stop's colour.
- ✅ **Non-solid pattern fills** — `darkGray` / `mediumGray` / `lightGray` /
  `dark*` / `light*` directional stripes + `*Grid` / `*Trellis` render
  as `repeating-linear-gradient` approximations on top of `bgColor`.
  `gray125` is still intentionally ignored (Excel's default).
- ✅ **Theme font scheme + phonetics** — `Theme.majorFont` /
  `Theme.minorFont`, `FontStyle.scheme`, `<rPh>` → HTML `<ruby><rt>`.
- ✅ **Indexed colours** — ECMA-376 64-entry palette via
  `indexedColor(i)`; `ColorRef` gained an `indexed` variant.
- ✅ **Multi-sheet + workbook rels** — `<sheet r:id>` resolution via
  `xl/_rels/workbook.xml.rels`.
- ✅ **Named styles inheritance** (one hop) — `cellXf.xfId` →
  `cellStyleXfs[j]` via `resolveEffectiveXf`.
- ✅ **Row heights + hidden rows/columns** — `RowDimension`,
  `ColumnWidth.hidden`, rendered via `tr.style.height` + `display: none`.
- ✅ **Graphical CF rules** — colourScale (2/3-stop interpolation),
  dataBar (gradient), iconSet (4 built-in sets).
- ✅ **Cached formula values** — `<f>` surfaces on `Cell.formula`;
  cached `<v>` is the displayed value.
- ✅ **Merged cells + column widths + frozen panes + autoFilter + tables
  + images + styles.xml fonts/fills/borders + rich text + dxf-driven cf
  + R1C1 notation + iconSet SVGs + classic + threaded comments**
  — baseline work from session 1.
- ✅ **Playwright browser harness** — 31 tests, real Chrome on :3002.
