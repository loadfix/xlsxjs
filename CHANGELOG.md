# Changelog

All notable changes to `xlsx-preview` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.2] - 2026-05-05

### Added

- **Formula evaluator POC** (`evaluateFormulas`, `evaluateFormulasForce`) —
  opt-in calc engine covering SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF,
  AND, OR, NOT plus arithmetic, comparison, string concat, and cell
  references. Cells whose cached `<v>` is empty get rewritten to the
  computed value; `evaluateFormulasForce` overwrites Excel's own cache.
- **Accessibility** — each sheet `<table>` now carries `role="table"`,
  `aria-labelledby`, `aria-rowcount`, and `aria-colcount`. Column-letter
  headers get `scope="col"`, row-number gutters get `scope="row"`, and
  the top-left corner cell is marked `aria-hidden="true"`.
- **Responsive mode** (`responsive: true`) — adds a `data-responsive`
  hook on each sheet `<section>` so the stylesheet turns on horizontal
  overflow plus sticky first column / first row for mobile viewports.
  Default off; existing consumers keep current behaviour.
- **Shared `oox-*` CSS classes** — rows / cells / sheet sections now also
  carry cross-format classes (`oox-table`, `oox-table-row`,
  `oox-table-cell`, `oox-page`) alongside the existing `xlsx-*` classes,
  so downstream stylesheets can style DOCX pages, PPTX slides, and XLSX
  sheets in one place.

### Fixed

- **Cached formula fallback** — when `<c><f>…</f></c>` carries no `<v>`,
  render the formula text (or `#ERROR!`) rather than a blank cell.

## [0.0.1] - 2026-05-04

### Added

- Initial public release. Browser-side XLSX → HTML renderer:
  multi-sheet parse, `xl/styles.xml` fonts / fills / borders /
  alignment / number formats, merged cells, frozen & split panes,
  autoFilter markers, tables, conditional formatting (colour scales,
  data bars, icon sets), embedded images, classic + threaded
  comments, per-sheet display state (visibility, RTL, gridlines,
  zoom, tab colour), page-layout metadata (breaks, print area,
  header/footer zones), and `xl/metadata.xml` dynamic-array spill
  anchors.
- Classic chart rendering (column / bar / line / pie / scatter /
  area / radar / doughnut, including stacked + percentStacked
  variants and data-label annotations) as inline SVG.
- SmartArt detection with opt-in SVG layouts (`smartArtLayout: 'svg'
  | 'both'`) covering hierarchy, orgchart, and cycle.
- Detect-only models for form controls, slicers, timelines, and
  OLE / package embeddings, with opt-in interactive widgets
  (`interactiveFormControls`, `interactiveSlicers`) and opt-in CFB
  stream extraction (`parseOleCfb`).
- Shape preset geometry rendering for 26 `prstGeom` presets.
- Full expression-rule conditional-formatting interpretation
  (AND / OR / NOT, ISNUMBER / ISBLANK / ISERROR, SEARCH,
  MOD(ROW() / COLUMN()), ISEVEN / ISODD, LEFT / RIGHT equality).
- Golden HTML diff harness (`npm run test:golden`) and Playwright
  browser smoke tests.

[0.0.2]: https://github.com/loadfix/xlsxjs/releases/tag/v0.0.2
[0.0.1]: https://github.com/loadfix/xlsxjs/releases/tag/v0.0.1
