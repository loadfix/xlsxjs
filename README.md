# xlsx-preview

Browser-side XLSX → HTML renderer, written from scratch in TypeScript and following the architecture of the sibling [docxjs](https://github.com/loadfix/docxjs) library (TypeScript + rollup, jsdom for the depth harness, Playwright for the browser smoke, JSZip for package I/O). Not affiliated with SheetJS or other libraries that share the xlsx name.

Opens an `.xlsx`, walks every sheet, and emits an HTML table per sheet with column-letter headers, row-number gutters, cell formatting from `xl/styles.xml` (fonts, fills, borders, alignment — including `wrapText`, `shrinkToFit`, `indent`, `textRotation` (incl. stacked), `readingOrder`, and the widened horizontal/vertical enums — number formats), merged cells, frozen or split panes, autoFilter markers, parsed tables, conditional formatting (including colour scales, data bars, and icon sets), embedded images, classic + threaded comments, per-sheet display state (visibility, RTL direction, grid-line / header toggles, zoom, and tab colour), page-layout metadata (manual row/column page breaks, `_xlnm.Print_Area` resolution, and three-zone `<headerFooter>` text), and `xl/metadata.xml` — dynamic-array spill anchors surface on `Cell.isSpillAnchor` and get a dashed `.xlsx-spill-anchor` outline.

## Installation

```
npm install xlsx-preview
```

## Usage

```html
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="dist/xlsx-preview.js"></script>
<script>
  fileInput.addEventListener('change', async () => {
    await xlsx.renderAsync(fileInput.files[0], document.getElementById('out'));
  });
</script>
```

## API

The public surface is:

- `parseAsync(data, options)` — open an XLSX blob/ArrayBuffer/Uint8Array and
  return the parsed workbook model (`{ sheets, styles, theme, date1904, … }`).
- `renderWorkbook(wb, options)` — render a previously-parsed workbook into
  `<section class="xlsx">` elements (one per sheet).
- `renderAsync(data, bodyContainer, styleContainer?, options)` — the common
  path: parse + render in one call.
- `defaultOptions` — the options object used when none is passed.
- Test-visible helpers: `formatNumber`, `parseStyles`, `resolveEffectiveXf`,
  `sanitizeHexColor`, `sanitizeFontFamily`, `sanitizeMediaMime`,
  `isSafeHyperlinkHref`, `bytesToDataUrl`, `a1ToR1c1`, `r1c1ToA1`,
  `emuToPx`, `evaluateRule`, `parseChart`, `renderChart`, `parseSmartArt`,
  `renderSmartArtSvg`, `applyFormControlUpdate`.

Options of note:

- `showFormulas: boolean` — replace cell text with the `<f>` formula.
- `formulaNotation: 'a1' | 'r1c1'` — which formula dialect to render.
- `inlineEmbeddings: boolean` — opt in to base64-inlining embedded files
  (PDFs, inner xlsx/docx/pptx, images) as `data:` URLs with an
  `<a download>` affordance. 32 MiB cap + MIME allowlist; OLE CFB `.bin`
  streams stay un-inlined.
- `renderCharts: boolean` — render classic `c:chartSpace` charts as inline
  SVG (default `true`). When `false`, charts fall back to the dashed
  placeholder — useful if the consumer wants to paint their own chart UI.
- `interactiveFormControls: boolean` — opt in to live form widgets.
  Checkbox / radio / scrollbar / spinner / combo / button render as real
  HTML inputs whose change events update the `linkedCell` cell in the
  rendered DOM. Default `false` (the detect-only aside from Wave 8 ships
  unchanged).
- `interactiveSlicers: boolean` — opt in to clickable slicer chips +
  dual-handle timeline slider. Changes fire `xlsx:slicer-change` /
  `xlsx:timeline-change` CustomEvents on the aside. Default `false`.
- `smartArtLayout: 'tree' | 'svg' | 'both'` — controls how SmartArt
  diagrams render. `'tree'` (default, byte-stable) emits only the indented
  `<ul>`; `'svg'` emits only the inline SVG hierarchy; `'both'` emits both
  (SVG first for visual, `<ul>` after for accessibility / text).

See `src/xlsx-preview.ts` for the full options list.

## Status

Early. The public API (`parseAsync` / `renderWorkbook` / `renderAsync`) is
stable enough for downstream use, but the model emitted by `parseAsync` is
still evolving as new XLSX features come online (see `TODO.md`). Pin a
specific version if you rely on the shape of the workbook model.

Drawings beyond raster images (`xdr:sp` text boxes / WordArt and
`xdr:cxnSp` connectors) surface on `Sheet.shapes` with their preset
geometry, text body, and anchor coordinates; the renderer emits an
`<aside class="xlsx-shape">` per shape with an inline SVG for recognised
presets (rect, ellipse, line, triangle, arrows, callouts, flowchart shapes).

Form controls (`Sheet.formControls`), slicers (`Sheet.slicers`),
timelines (`Sheet.timelines`), and OLE / package embeddings
(`Sheet.embeddings`) all surface as detect-only models with metadata —
kind, anchor, linkedCell, selectedItems, progId, etc. — and render as
informational `<aside>` elements. Form-control widgets become live
on opt-in (`interactiveFormControls: true`); slicer/timeline UI and
pivot interactivity are still consumer territory.

Classic chartSpace charts (column / bar / line / pie / scatter / area,
including stacked + percentStacked variants) render as inline SVG via
`ChartModel` + `renderChart`. SmartArt hierarchy diagrams surface on
`Sheet.smartArt` with the parsed tree and render as a nested `<ul>`
aside by default; opt in to `smartArtLayout: 'svg'` / `'both'` for an
inline SVG hierarchy via `renderSmartArtSvg`.

Images render as absolutely-positioned `<figure>`s inside a zero-height
`.xlsx-image-layer` above the `<table>`, so anchor coordinates place the
figure over the correct cell range.

Expression-rule conditional formatting supports `AND` / `OR` / `NOT`,
`ISNUMBER` / `ISBLANK` / `ISERROR`, `SEARCH`, `MOD(ROW()/COLUMN(), n)`,
`ISEVEN` / `ISODD`, and `LEFT` / `RIGHT` equality alongside the narrow
`=<cellRef> <op> <literal>` form.

Deliberately deferred: chart rendering, SmartArt, form-control VML
fallbacks, and double borders.

## Contributing

```bash
npm install
npm run build
npm run test:render   # jsdom depth harness (105 scenarios)
npm run test:golden   # golden HTML diff against result.html snapshots
npm test              # Playwright browser smoke (real Chrome, port :3002)
npm run dev           # static demo server at :8767
```

PRs welcome. Keep `README.md` and `TODO.md` in sync with code changes in
the same PR (see `CLAUDE.md`).

Apache-2.0.

## Related projects

Part of a family of document-rendering libraries:

- [docxjs](https://github.com/loadfix/docxjs) — browser-side DOCX → HTML renderer (TypeScript)
- [pptxjs](https://github.com/loadfix/pptxjs) — browser-side PPTX → HTML renderer (TypeScript)
- [python-docx](https://github.com/loadfix/python-docx) — Python DOCX parser/generator
- [python-pptx](https://github.com/loadfix/python-pptx) — Python PPTX parser/generator
- [python-xlsx](https://github.com/loadfix/python-xlsx) — Python XLSX parser/generator
