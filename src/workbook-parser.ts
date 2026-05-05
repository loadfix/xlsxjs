// XLSX XML → internal model. Scope for the initial slice: sheet names, shared
// strings, and per-cell values (string / number / inline string / boolean).
// No styles, no number formats, no formulas — those land in follow-ups.
//
// The model shape is intentionally small and stable so the renderer can walk
// it without caring about the on-disk XML structure.

import { parseCellRef } from './utils';
import type { Options } from './xlsx-preview';
import { parseStyles, parseColorElement, type Styles, type FontStyle, type UnderlineStyle } from './styles';
import { parseTheme, type Theme, type ColorRef } from './theme';
import { parseConditionalFormatting, parseSqref, type ConditionalFormatting } from './conditional-format';
import { bytesToDataUrl, sanitizeMediaMime } from './workbook';
import { parseChart as parseChartModel, type ChartModel } from './chart-parser';
import { parseSmartArt, type SmartArtModel } from './smartart-parser';
import { parseCfb, type CfbModel } from './ole-cfb';
export type { SmartArtModel, SmartArtNode } from './smartart-parser';
export type { CfbModel, CfbStream } from './ole-cfb';
export { parseSmartArt } from './smartart-parser';
export { parseCfb } from './ole-cfb';

// URL schemes we'll emit as an `<a href="…">` in the rendered sheet. Anything
// outside this set (most importantly `javascript:` / `data:` / `vbscript:` /
// `file:` / `blob:`) is dropped to inert plain text. Fragment-only (`#foo`)
// and relative paths (no scheme) are accepted: they can't carry script.
// Mirrors docxjs's isSafeHyperlinkHref — keep them in lockstep.
const SAFE_HREF_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:']);

// DoS guardrail: a single hyperlink / data-validation range is expanded to
// one entry per covered cell so the renderer can look up by (row, col). An
// attacker-controlled `ref="A1:XFD1048576"` would otherwise balloon to 17
// billion entries and exhaust memory before the sheet ever renders. Ranges
// wider than this cap are clipped to the cap and a note is dropped on the
// console.warn channel. The cap is generous enough to cover legitimate
// whole-column selections (`A:A` = 1,048,576 cells) but stops single-range
// bombs dead.
const MAX_RANGE_EXPANSION_CELLS = 1_048_576;

/**
 * Returns `true` iff `raw` can be safely emitted as the `href` of an `<a>` in
 * a read-only xlsx viewer. Accepts absolute URLs with known-safe schemes,
 * fragment-only URLs (`#anchor`), and relative paths. `javascript:`,
 * `data:`, `vbscript:`, `file:`, `blob:`, and every other scheme outside the
 * allowlist are rejected.
 */
export function isSafeHyperlinkHref(raw: string | null | undefined): boolean {
    if (raw == null) return true; // empty href is inert
    if (typeof raw !== 'string') return false;
    // Control characters (tab, newline, null byte, etc.) in a URL are
    // normalised away by both the WHATWG URL parser AND by the browser
    // when it resolves an anchor's href — meaning `"java\tscript:alert(1)"`
    // passes a naive URL(…).protocol check (the URL parser sees it as a
    // relative path and assigns http:) but the browser executes it as
    // `javascript:` on click. Reject any string containing a control
    // character outright; a legitimate hyperlink never contains one.
    if (/[\x00-\x1f\x7f]/.test(raw)) return false;
    const trimmed = raw.trim();
    if (trimmed === '') return true;
    if (trimmed.startsWith('#')) return true;
    try {
        // Resolve against a synthetic base so relative URLs parse but never
        // inherit the host document's base URI (which could be file:// in
        // embedding apps and pass the scheme allowlist accidentally).
        const parsed = new URL(trimmed, 'http://xlsxjs.invalid/');
        return SAFE_HREF_SCHEMES.has(parsed.protocol);
    } catch {
        // Non-URL strings that aren't fragments are almost always relative
        // paths like `foo/bar.html`. Treat as safe — no scheme, no sink.
        return !/^[a-z][a-z0-9+.-]*:/i.test(trimmed);
    }
}

export interface RichTextRun {
    text: string;
    // Subset of FontStyle — runs carry their own rPr. Properties not present
    // on the run inherit from the cell's xf font.
    bold: boolean;
    italic: boolean;
    underline: UnderlineStyle;
    strike: boolean;
    vertAlign: 'subscript' | 'superscript' | null;
    size: number | null;
    color: ColorRef;
    name: string | null;
}

// A phonetic-ruby annotation attached to a <si>. East-Asian text (most often
// Japanese kanji) can carry one or more <rPh> children that map a phonetic
// reading onto a substring of the main text via [startIdx..endIdx) offsets.
// The renderer emits these as HTML5 <ruby>…<rt>…</rt></ruby> markup.
export interface PhoneticRun {
    // The kanji (or other base) substring spanning [startIdx..endIdx) of the
    // shared-string / inline-string body. Precomputed for consumer
    // convenience — equivalent to `text.slice(startIdx, endIdx)`.
    base: string;
    // The phonetic reading — hiragana / katakana / romaji — that Excel wants
    // rendered above the base characters.
    phonetic: string;
    // 0-based inclusive start index of the base span in the main text.
    startIdx: number;
    // 0-based exclusive end index of the base span in the main text.
    endIdx: number;
}

// A shared-string entry. Plain strings carry text only; rich entries also
// carry the runs so the renderer can emit <span> per run with the correct
// styling. `text` is always the concatenated plain-text projection.
export interface SharedString {
    text: string;
    runs: RichTextRun[] | null;
    // Phonetic-ruby annotations (<rPh>) attached to the string, or null when
    // the <si> declared none. Each entry maps a phonetic reading onto a
    // substring of `text`. The renderer wraps covered spans in HTML5 <ruby>.
    phonetics: PhoneticRun[] | null;
}

export interface Cell {
    col: number;
    row: number;
    value: string;
    kind: 'string' | 'number' | 'boolean' | 'inlineStr' | 'error' | 'empty';
    // Index into Styles.cellXfs. -1 when the cell has no `s=` attribute.
    styleIndex: number;
    // Formula text when the cell carries an `<f>` child. Null otherwise.
    // xlsxjs does not evaluate formulas — when a cached `<v>` is present it
    // becomes the displayed value; when absent the cell renders empty.
    formula: string | null;
    // For string / inlineStr cells whose source had <r>/<rPr> runs, the
    // runs are preserved here so the renderer can emit formatted spans.
    // Null on non-string cells or plain strings.
    runs: RichTextRun[] | null;
    // Phonetic-ruby annotations inherited from the shared-string / inline-
    // string source. Null when the source <si>/<is> declared no <rPh>.
    phonetics: PhoneticRun[] | null;
    // Index into Workbook.metadata.cellMetadata, resolved from c/@cm. null
    // when the cell has no cm attribute. 0-based in our model even though
    // the XML side is 1-based.
    cellMetadataIndex: number | null;
    // Index into Workbook.metadata.valueMetadata, resolved from c/@vm.
    // Same null/0-based convention as cellMetadataIndex.
    valueMetadataIndex: number | null;
    // Shortcut: true when the resolved cellMetadata block is a dynamic-array
    // (XLDAPR) type and the futureMetadata fDynamic flag is set. Lets the
    // renderer tag spill-anchor cells without walking the indirection chain.
    isSpillAnchor: boolean;
}

// A single block in the cellMetadata / valueMetadata list.
export interface CellMetadataBlock {
    // Index into the metadataTypes array; 0-based in our model even though
    // the XML side is 1-based.
    typeIndex: number;
    // For dynamic-array types (XLDAPR), true when this cell is a spill anchor.
    // False for any other metadata type, or when the resolved futureMetadata
    // block had fDynamic="0".
    dynamicArray: boolean;
}

// Workbook-wide cell metadata lifted from xl/metadata.xml. Excel 365 uses
// this part to describe dynamic-array spill anchors (XLDAPR) and linked /
// rich data types; xlsxjs exposes the resolved blocks and lets the renderer
// tag spill-anchor cells via Cell.isSpillAnchor.
export interface WorkbookMetadata {
    // cellMetadata[i] is the block at 1-based index i+1 on the XML side.
    cellMetadata: CellMetadataBlock[];
    // Same for valueMetadata.
    valueMetadata: CellMetadataBlock[];
}

export interface MergedRange {
    // Inclusive bounds, 0-based.
    col: number;
    row: number;
    colSpan: number;
    rowSpan: number;
}

export interface ColumnWidth {
    // Inclusive bounds, 0-based, matching <col min/max> but zero-indexed.
    min: number;
    max: number;
    // Width in Excel character units, or null when the column has no custom
    // width but is present for another reason (e.g. hidden="1"). Renderer
    // converts to pixels when non-null.
    width: number | null;
    hidden: boolean;
    // <col outlineLevel="N"/> — 0 when absent. Groups adjacent columns in
    // the Excel UI; xlsxjs doesn't render a collapse affordance but exposes
    // the level on the model + as a data attribute in the DOM.
    outlineLevel: number;
}

export interface RowDimension {
    // 0-based row index. Present only when the row carries a custom height,
    // a hidden flag, or an outlineLevel — sparse by design.
    row: number;
    // Height in points. null when only the hidden flag was set.
    height: number | null;
    hidden: boolean;
    // <row outlineLevel="N"/> — 0 when absent. Same semantics as
    // ColumnWidth.outlineLevel, applied to rows.
    outlineLevel: number;
}

// Sheet-wide outline configuration (sheetFormatPr + sheetPr/outlinePr). The
// max levels mirror the deepest row/col outline observed in the sheet. The
// summary flags govern which side of a group the summary row/column sits on
// (default: below for rows, right for columns).
export interface SheetOutline {
    maxRowLevel: number;    // from sheetFormatPr/@outlineLevelRow (default 0)
    maxColLevel: number;    // from sheetFormatPr/@outlineLevelCol (default 0)
    summaryBelow: boolean;  // sheetPr/outlinePr/@summaryBelow (default true)
    summaryRight: boolean;  // sheetPr/outlinePr/@summaryRight (default true)
}

// A workbook-scoped defined name. localSheetId is null for workbook-scoped
// names (the default) or a 0-based sheet index when the name is local to
// one sheet. Excel stores print areas, print titles, and autoFilter state
// as `_xlnm.*` defined names — we expose them here without filtering so
// callers can pick out the bits they care about.
export interface DefinedName {
    name: string;
    localSheetId: number | null;
    // The raw reference / formula (e.g. "Sheet1!$A$1:$B$2"). xlsxjs does not
    // evaluate or resolve this — consumers can parse it themselves.
    formula: string;
    // @hidden="1" marks print-area / print-titles / autoFilter internals
    // that Excel hides from the user-facing name manager UI.
    hidden: boolean;
}

// Frozen / split pane metadata from <sheetViews><sheetView><pane ...>. xlsxjs
// renders frozen panes as sticky <th>/<td> borders; a plain `state="split"`
// (scroll split without freeze) carries the same visual treatment, but
// `kind` distinguishes the two so consumers can tell them apart.
export interface FrozenPanes {
    // 'frozen' for state="frozen" or "frozenSplit"; 'split' for state="split"
    // (scroll split without freezing). Renderer treats both identically.
    kind: 'frozen' | 'split';
    // 0-based column index at which horizontal freeze begins, or null when
    // no horizontal freeze. A value of 3 means cols 0..2 are frozen.
    xSplit: number | null;
    // Same for rows.
    ySplit: number | null;
}

// Page-break metadata from <rowBreaks>/<colBreaks>. Only manual breaks
// (man="1") are captured — automatic breaks are Excel's pagination hints,
// not author intent. Indices are 0-based.
export interface PageBreaks {
    rows: number[];
    cols: number[];
}

// A resolved print area range from a `_xlnm.Print_Area` defined name. The
// defined name's `formula` is parsed into cell-range coordinates here so
// consumers don't need to re-parse the reference string. Multiple entries
// arise when a single Print_Area formula carries a comma-separated list
// of ranges (rare, but legal per ECMA-376).
export interface PrintAreaRange {
    col: number;
    row: number;
    endCol: number;
    endRow: number;
}

// Header/footer zones split out of the &L / &C / &R codes in a single
// header or footer string. Substitution codes (&P, &N, &D, &T, &F, &A)
// are resolved eagerly during parsing using the sheet / workbook context;
// &P and &N remain literal since xlsxjs does not paginate.
export interface HeaderFooterZones {
    left: string;
    center: string;
    right: string;
}

// Sheet <headerFooter> metadata. Only oddHeader/oddFooter are surfaced —
// even / first-page variants are schema-legal but rarely used in practice
// and a viewer that doesn't paginate can't select between them anyway.
export interface HeaderFooter {
    oddHeader: HeaderFooterZones | null;
    oddFooter: HeaderFooterZones | null;
}

// <autoFilter ref="A1:D10"/> — the range that has filter dropdowns. We
// don't render the dropdown UI; the flag surfaces the filtered region so
// consumers can render their own affordance.
export interface AutoFilter {
    col: number; row: number;
    endCol: number; endRow: number;
}

// A detected chart anchored inside a sheet. xlsxjs does NOT render the chart
// content — the model carries the position + a rough chart-type name so
// consumers (and the renderer) can emit a placeholder.
export interface SheetChart {
    // 'classic' is c:chartSpace (charts 1.0), 'chartex' is cx:chartSpace
    // (the 2014 extensions: treemap, sunburst, waterfall, funnel, pareto,
    // box-whisker, histogram, map).
    kind: 'classic' | 'chartex';
    // Best-effort chart-type tag: e.g. 'barChart', 'lineChart', 'pieChart'
    // for classic; 'treemap', 'sunburst', 'waterfall', … for chartEx.
    // null when the chart xml couldn't be found or had no recognisable type.
    chartType: string | null;
    col: number; row: number;
    endCol: number | null;
    endRow: number | null;
    // Parsed classic chart (c:chartSpace) model, when the chart part could
    // be resolved and contains a supported plot type. chartEx charts and
    // malformed / missing XML leave this null. The renderer projects this
    // model to inline SVG via src/chart-renderer.ts; null falls through to
    // the dashed placeholder.
    model: ChartModel | null;
}

// A pivot table anchored inside a sheet. Detection-only — the sheet's cell
// values already carry the materialised pivot output.
export interface SheetPivot {
    name: string;
    col: number; row: number;
    endCol: number; endRow: number;
}

// An Excel 2010+ slicer surfaced from xl/slicers/slicerN.xml + its cache
// under xl/slicerCaches/slicerCacheN.xml. Detection-only — xlsxjs does NOT
// render the clickable slicer UI; the model carries enough metadata for a
// consumer to synthesise one (caption, selected items, source column
// name, dimensioning hints). Selected items are the plain-text values
// Excel kept with `s="1"` under <items> — presenting them alone is a
// faithful read-only summary of the slicer's current state.
export interface SheetSlicer {
    name: string;           // <slicer name="…"/>
    caption: string | null; // <slicer caption="…"/> — header label
    cache: string;          // <slicer cache="…"/> — slicerCacheDefinition name
    // Pivot-source column the slicer filters by. Lifted from the slicer
    // cache's <tabular pivotCacheId="…"> → <items>, with the field label
    // coming from <slicerCacheDefinition sourceName="…"/>. Null when the
    // cache is either missing or not the tabular variant.
    sourceName: string | null;
    columnCount: number | null;    // <slicer columnCount="…"/>  — tile columns
    style: string | null;          // <slicer style="…"/>        — Excel built-in
    showCaption: boolean;          // <slicer showCaption="…"/>  — defaults true
    rowHeight: number | null;      // <slicer rowHeight="…"/>    — EMU
    // Plain-text values with @s="1" in the slicer cache's <items>. Empty
    // array when the cache declared no selection (Excel treats that as
    // "all selected", but surfacing an empty selection list is the honest
    // read-only representation).
    selectedItems: string[];
    // Every plain-text item the cache enumerated (both selected and
    // unselected, in @x index order). Empty when the cache either
    // didn't list items at all or listed only unparseable entries. Used by
    // the interactive renderer to emit chips for every option, not just
    // the currently-selected ones.
    allItems: string[];
}

// An Excel 2013+ timeline (date slicer) surfaced from xl/timelines/
// timelineN.xml + its cache under xl/timelineCaches/timelineCacheN.xml.
// Detection-only — xlsxjs does NOT render the draggable timeline track;
// the model carries the caption, the current level (years/quarters/
// months/days), and the currently-selected range endpoints.
export interface SheetTimeline {
    name: string;           // <timeline name="…"/>
    caption: string | null; // <timeline caption="…"/> — header label
    cache: string;          // <timeline cache="…"/>   — timelineCacheDefinition name
    // Pivot-source field (usually a date column). Lifted from the timeline
    // cache's @sourceName. Null when the cache is missing or didn't carry
    // the attribute.
    sourceName: string | null;
    // Current timeline granularity. Excel writes this as the @level attr on
    // <timeline> (or <timelineViewState>), using one of the four lexical
    // forms listed here. Null when absent or unrecognised.
    level: 'years' | 'quarters' | 'months' | 'days' | null;
    // Selected date range. Both endpoints are left as raw strings (Excel's
    // `startDate` / `endDate` attrs are ISO-like "yyyy-mm-ddThh:mm:ss"
    // fragments); null when neither is declared.
    selectedRange: { start: string | null; end: string | null } | null;
    showHeader: boolean;            // default true
    showSelectionLabel: boolean;    // default true
    showTimeLevel: boolean;         // default true
    showHorizontalScrollbar: boolean; // default true
    style: string | null;           // <timeline style="…"/>
    // Absolute slider bounds surfaced from the timeline cache's
    // `<state><bounds startDate=… endDate=…/>` block. Interactive consumers
    // use these to pin the slider min/max; null when the cache didn't
    // enumerate bounds. Both endpoints are raw ISO-like strings, same shape
    // as `selectedRange`.
    bounds: { min: string; max: string } | null;
}

// Count of sheet-level <extLst><ext uri="…"> entries. Surfaced so the smoke
// tool can roll up unknown extension uses (sparklines, dynamic-array spill,
// protected ranges, etc.) without re-scanning the raw XML.
export interface SheetExtensionUri {
    uri: string;
    count: number;
}

// A drawing shape or connector anchored inside a sheet. Covers `<xdr:sp>`
// (rectangles, callouts, text boxes, WordArt, form-control surfaces) and
// `<xdr:cxnSp>` (connectors: straight, elbow, curved lines between two
// anchors). xlsxjs does NOT render the shape geometry — the model carries
// the anchor, the preset name, and any text body so consumers can emit
// their own placeholder; the built-in renderer emits an `<aside>` per
// shape with the text, marked with the preset for consumer styling.
//
// Form controls (xl/ctrlProps/*.xml) embed as an `<xdr:sp>` wrapped in
// `<mc:AlternateContent>`. We don't decode the VML fallback here — the
// shape surfaces with `kind='shape'`, the preset+alt+name from the
// `<xdr:cNvPr>`, and the text body if any. `kind='unknown'` is reserved
// for future drawing subtypes we haven't special-cased.
export interface SheetShape {
    kind: 'shape' | 'connector' | 'unknown';
    name: string | null;     // <xdr:cNvPr>/@name
    alt: string | null;      // <xdr:cNvPr>/@descr  ||  @title  (descr wins)
    preset: string | null;   // <a:prstGeom>/@prst — e.g. 'rect', 'ellipse',
                             // 'line', 'flowChartProcess', 'cloud'.
    text: string | null;     // flattened <xdr:txBody> text; paragraphs (<a:p>)
                             // joined with '\n', <a:t> within a paragraph
                             // concatenated without separators. null when
                             // the shape carries no text body.
    col: number; row: number;
    endCol: number | null;
    endRow: number | null;
}

// A SmartArt diagram anchored inside a sheet. The outer `<xdr:graphicFrame>`
// shape surfaces on `Sheet.shapes` independently; this entry is the parsed
// data model (`xl/diagrams/data1.xml`) + the layout uniqueId from
// `xl/diagrams/layout1.xml`. Detection-only: xlsxjs renders an indented
// tree aside — no layout graph interpretation. See `src/smartart-parser.ts`
// for the model shape and `renderSmartArt` in html-renderer for the DOM.
export interface SheetSmartArt {
    // cNvPr @name on the hosting <xdr:graphicFrame>, or null when absent.
    // Never used in CSS / selectors — only reaches the DOM via setAttribute.
    name: string | null;
    // 0-based anchor coordinates straight off the <xdr:twoCellAnchor>'s
    // <xdr:from> / <xdr:to>. oneCellAnchor-wrapped SmartArt is rare; when
    // encountered, endCol/endRow fall back to null.
    col: number;
    row: number;
    endCol: number | null;
    endRow: number | null;
    // Parsed data model. null when the diagram's data1.xml is missing or
    // unparseable (the outer graphicFrame still surfaces on shapes[]).
    model: SmartArtModel | null;
}

// A form-control anchored inside a sheet. Detect-only: we surface the
// metadata from `xl/ctrlProps/ctrlProp{N}.xml` + the hosting `<xdr:sp>` so
// consumers can discover buttons / checkboxes / scrollbars / dropdowns
// without having to re-parse the drawing layer. xlsxjs does NOT paint an
// interactive widget — the renderer emits an `<aside class="xlsx-form-control">`
// that captures kind + state for consumer-side styling / hydration.
//
// The `kind` values mirror the ctrlProp `@objectType` set documented in the
// XLSX form-control schema (a macrosheet extension in the 2009/9/main
// namespace). Unknown / malformed object types decay to `'unknown'` rather
// than throwing so a hostile file cannot break the parser by inventing a
// value.
export interface SheetFormControl {
    kind:
        | 'button'
        | 'checkbox'
        | 'radio'
        | 'combo'
        | 'list'
        | 'scrollbar'
        | 'spinner'
        | 'groupBox'
        | 'label'
        | 'dialog'
        | 'unknown';
    // Anchor coordinates (0-based) carried straight through from the
    // surrounding `<xdr:twoCellAnchor>`. The form-control renderer uses
    // these to position its aside inside the image overlay layer using the
    // same arithmetic as SheetImage.
    col: number;
    row: number;
    colOff: number;
    rowOff: number;
    endCol: number | null;
    endRow: number | null;
    endColOff: number | null;
    endRowOff: number | null;
    // Label text. For a button-style control it's the txBody flatten of the
    // hosting `<xdr:sp>`; for everything else it's the cNvPr @name when the
    // shape carries no visible text body. null when neither is declared.
    label: string | null;
    // Linked cell from `fmlaLink` (checkbox / radio / scrollbar / spinner /
    // combo). Raw A1 ref (e.g. `$Sheet2!$B$3`) — we do not dereference.
    linkedCell: string | null;
    // Input source from `fmlaRange` (combo/list). Raw A1 range.
    inputRange: string | null;
    // Checkbox / radio state. `null` when the attribute is absent.
    checked: boolean | null;
    // Numeric parameters (scrollbar / spinner). All `null` unless the
    // ctrlProp declares them.
    min: number | null;
    max: number | null;
    inc: number | null;
    page: number | null;
    val: number | null;
    // Combo / list dropdown height, in rows. `null` when absent.
    dropLines: number | null;
    // Accessibility: prefer `cNvPr/@descr`, fall back to `cNvPr/a:extLst`
    // alt-text ext (same as SheetImage.alt / SheetShape.alt).
    altText: string | null;
}

// A rendered image anchored inside a sheet. Coordinates are 0-based.
// xlsxjs distinguishes three DrawingML anchor modes:
//   - 'twoCell'   : <xdr:from> + <xdr:to>. endCol/endRow + widthEmu/heightEmu
//                   are populated (size may be derived from cells).
//   - 'oneCell'   : <xdr:from> + <xdr:ext cx cy/>. endCol/endRow stay null;
//                   widthEmu/heightEmu come from the anchor's <ext>.
//   - 'absolute'  : <xdr:pos x y/> + <xdr:ext cx cy/>. col/row/endCol/endRow
//                   are null (or 0 for col/row compatibility); absoluteX/Y
//                   carry the pixel-absolute position in EMU.
export interface SheetImage {
    col: number; row: number;           // top-left anchor cell (0 for absolute)
    endCol: number | null;              // bottom-right anchor cell (twoCellAnchor)
    endRow: number | null;
    // Offsets from the top-left anchor cell, in EMUs (English Metric Units;
    // 914400 per inch, 9525 per pixel). Renderer converts to px via emuToPx.
    colOff: number; rowOff: number;
    // Data URL + intrinsic dimensions (or null if the drawing didn't declare them).
    dataUrl: string;
    widthEmu: number | null;
    heightEmu: number | null;
    alt: string | null;
    // Which DrawingML anchor element wrapped this image.
    anchorMode: 'twoCell' | 'oneCell' | 'absolute';
    // Pixel-absolute position in EMU (absoluteAnchor only; null otherwise).
    absoluteX: number | null;
    absoluteY: number | null;
    // True when the image carries a "decorative=1" marker in its cNvPr
    // extLst (the modern accessibility hint, distinct from empty alt).
    // Renderer surfaces this as alt="" + aria-hidden="true" so screen
    // readers skip the image entirely.
    decorative: boolean;
}

// A detected OLE / embedded-file attachment anchored inside a sheet. Detect-
// only: the payload is never inlined by default — the renderer emits a
// placeholder `<aside class="xlsx-embedding">`. Two flavours are surfaced:
//   - `kind='ole'`   — a raw OLE CFB container (the `.bin` payload under
//                      xl/embeddings/), referenced via an "/oleObject" rel.
//                      `progId` is the Excel hint ("Word.Document.12",
//                      "Equation.DSMT4", …) when present.
//   - `kind='package'` — a real file (xlsx, docx, pptx, pdf, …) packaged
//                      inside the xlsx via a "/package" rel. `progId` is
//                      null — the contentType carries the MIME.
// `dataUrl` is null unless the caller opted in via `Options.inlineEmbeddings`;
// when enabled, payloads up to MAX_MEDIA_BYTES pass through sanitizeMediaMime
// and decay to null on rejection.
export interface SheetEmbedding {
    kind: 'ole' | 'package';
    contentType: string;
    fileName: string | null;
    progId: string | null;
    col: number | null;
    row: number | null;
    endCol: number | null;
    endRow: number | null;
    size: number;
    dataUrl: string | null;
    altText: string | null;
    // CFB stream list extracted from the raw `.bin` payload when the caller
    // opted in via Options.parseOleCfb AND kind === 'ole'. Null otherwise
    // (including when kind === 'package', since package embeddings are not
    // CFB containers). The list is capped + guardrailed — see parseCfb() in
    // src/ole-cfb.ts.
    cfb: CfbModel | null;
}

// A defined table (<table> inside xl/tables/tableN.xml, referenced from
// the sheet's rels). Separate from autofilter — tables carry names,
// header/totals info, and per-column filters. xlsxjs exposes them on the
// sheet so callers can render their own table-aware UI.
export interface TableDef {
    name: string;
    displayName: string;
    col: number; row: number;
    endCol: number; endRow: number;
    headerRowCount: number;
    totalsRowCount: number;
    columns: { name: string }[];
    // Table accessibility metadata from the <table> root attributes.
    // `altText` is a short title suitable for `aria-label`; `altTextSummary`
    // is a longer description meant for `title` fallback. Both are null when
    // the producer did not set them.
    altText: string | null;
    altTextSummary: string | null;
}

// Sheet-protection state surfaced from <sheetProtection …/>. A locked
// operation is represented as attribute value `"1"`; explicitly unlocked is
// `"0"`. Per ECMA-376 §18.3.1.85, absent attributes default to `"1"` so we
// parse each flag as `attr !== "0"`. xlsxjs does not enforce the protection
// — it only surfaces the state so consumers can read it / style with it.
export interface SheetProtection {
    enabled: boolean;             // true when <sheetProtection sheet="1"/>
    passwordHashed: boolean;      // true when @password is present
    // Granular toggles are rarely inspected by readers; we surface the common
    // ones individually. Each is `true` when the corresponding operation is
    // locked (including via absence, per ECMA default).
    selectLockedCells: boolean;
    selectUnlockedCells: boolean;
    formatCells: boolean;
    formatColumns: boolean;
    formatRows: boolean;
    insertColumns: boolean;
    insertRows: boolean;
    deleteColumns: boolean;
    deleteRows: boolean;
    sort: boolean;
    autoFilter: boolean;
    pivotTables: boolean;
    objects: boolean;
    scenarios: boolean;
}

// A classic comment attached to a single anchor cell. Parsed from
// xl/comments{N}.xml and bound to a sheet via its rels. xlsxjs ignores
// the VML bubble-layout (xl/drawings/vmlDrawing{N}.vml) — only the
// cell anchor, author, and body matter for inline-marker rendering.
export interface SheetComment {
    col: number; row: number;    // 0-based anchor, derived from ref="B3"
    author: string | null;
    text: string;                // concatenated plain-text projection
    runs: RichTextRun[] | null;  // preserved when the <text> body had runs
}

// A single entry in a threaded-comment thread. Parent comments carry
// `parentId = null`; replies reference the starter by GUID. `author` has
// already been resolved through the workbook-wide person.xml registry,
// so the renderer never needs to cross-reference IDs. `date` is the raw
// ISO string off the `dT` attribute — we don't parse / reformat it here.
export interface ThreadedCommentEntry {
    id: string;
    col: number; row: number;
    author: string | null;    // resolved from personId via person.xml
    date: string | null;      // raw ISO string from dT attribute
    text: string;
    parentId: string | null;  // null for thread starter
}

// Per-sheet view state surfaced from <sheetViews><sheetView ...>/> plus the
// tab colour under <sheetPr><tabColor/>. These are display-only hints —
// xlsxjs renders them as section-level attributes / classes so consumers can
// react (skip hidden sheets, flip direction, paint a tab strip, etc.).
export interface SheetView {
    // True when <sheetView rightToLeft="1"/>; the renderer flips the section's
    // `dir` attribute to "rtl".
    rightToLeft: boolean;
    // False only when <sheetView showGridLines="0"/> — default is true.
    showGridLines: boolean;
    // False only when <sheetView showRowColHeaders="0"/> — default is true.
    showRowColHeaders: boolean;
    // Zoom percentage (100 = 100%). Null when no zoomScale was declared.
    zoomScale: number | null;
    // <sheetPr><tabColor .../>. Parsed via parseColorElement so the same rgb
    // / theme+tint / indexed paths as cell fills resolve it.
    tabColor: ColorRef | null;
}

// A per-cell hyperlink, resolved from `<hyperlinks><hyperlink>` + the sheet's
// rels. Range-valued sources (`ref="A1:C3"`) are expanded to one entry per
// covered cell so the renderer can find a link by (row,col). Only the
// top-left anchor is generally clickable in Excel but we carry the metadata
// on every covered cell for look-up purposes.
export interface Hyperlink {
    col: number;
    row: number;
    target: string | null;    // external URL from rels (e.g. "https://…")
    location: string | null;  // intra-workbook anchor ("Sheet2!A1")
    tooltip: string | null;
    display: string | null;
}

// A type="list" data-validation entry. Numbers/date/custom validations are
// intentionally ignored — only list validations surface as a ▾ affordance.
// When the formula1 source is a literal quoted list ("Red,Green,Blue") the
// options are pinned on the model; range-reference sources (Sheet2!$A$1:$A$5)
// leave options=null and only the ▾ indicator renders.
export interface DataValidationList {
    col: number; row: number;
    endCol: number; endRow: number;
    options: string[] | null;
}

export interface Sheet {
    name: string;
    // From the <sheet state="…"/> attribute on xl/workbook.xml's <sheets>.
    // 'visible' is the default; 'hidden' and 'veryHidden' suppress rendering
    // (the former is user-toggleable in Excel, the latter VBA-only).
    state: 'visible' | 'hidden' | 'veryHidden';
    rows: Cell[][]; // sparse: rows[rowIndex] may be undefined
    maxCol: number;
    maxRow: number;
    merges: MergedRange[];
    columns: ColumnWidth[];
    rowDimensions: RowDimension[];
    conditionalFormatting: ConditionalFormatting[];
    frozenPanes: FrozenPanes | null;
    autoFilter: AutoFilter | null;
    tables: TableDef[];
    images: SheetImage[];
    charts: SheetChart[];
    shapes: SheetShape[];
    // SmartArt diagrams anchored on the sheet. Detection-only: the model is
    // a hierarchy tree parsed from xl/diagrams/data1.xml; the outer
    // <xdr:graphicFrame> shape still surfaces on shapes[] under the same
    // anchor. Callers that want to render a fancy layout can branch on
    // model.layout (the layoutDef @uniqueId, e.g.
    // ".../layout/hierarchy1") while the DOM-side tree stays stable.
    smartArt: SheetSmartArt[];
    // Form controls (buttons, checkboxes, radios, scrollbars, dropdowns…).
    // Detect-only: the renderer emits an informational `<aside>` rather
    // than an interactive widget. Kept separate from `shapes[]` because
    // form controls carry their own metadata (linkedCell, range, min/max).
    formControls: SheetFormControl[];
    embeddings: SheetEmbedding[];
    pivots: SheetPivot[];
    slicers: SheetSlicer[];
    timelines: SheetTimeline[];
    extensions: SheetExtensionUri[];
    comments: SheetComment[];
    threadedComments: ThreadedCommentEntry[];
    // Per-sheet display state. Never null — a sheet with no <sheetView> gets
    // the documented defaults (showGridLines=true, showRowColHeaders=true,
    // rightToLeft=false, zoomScale=null, tabColor=null).
    view: SheetView;
    outline: SheetOutline;
    hyperlinks: Hyperlink[];
    dataValidationLists: DataValidationList[];
    // Manual row/column page breaks (0-based indices). Empty arrays when the
    // sheet declares no <rowBreaks>/<colBreaks> or every break is automatic.
    pageBreaks: PageBreaks;
    // Print ranges resolved from `_xlnm.Print_Area` defined names scoped to
    // this sheet (via `localSheetId`). Null when the workbook declares no
    // print area for the sheet.
    printArea: PrintAreaRange[] | null;
    // <headerFooter> metadata. Null when the sheet has no headerFooter entry
    // (or neither oddHeader nor oddFooter carries any substituted content).
    headerFooter: HeaderFooter | null;
    // <sheetProtection> state. Null when the sheet has no sheetProtection
    // element; an object (with `enabled=true`) when the sheet is protected.
    protection: SheetProtection | null;
}

export interface Workbook {
    sheets: Sheet[];
    styles: Styles | null;
    theme: Theme | null;
    // Workbook-wide author registry (personId GUID → displayName). Empty
    // Map when the package carries no xl/persons/person.xml.
    persons: Map<string, string>;
    // True when xl/workbook.xml declares <workbookPr date1904="1"/>. Flips
    // date-serial interpretation from the 1900 system (PC convention, with
    // the historical 1900-02-29 leap bug) to the 1904 system (Mac Office
    // pre-2011 convention, no leap bug). Threaded into formatNumber so date
    // cells render correctly on files authored under either convention.
    date1904: boolean;
    // Workbook-scoped defined names. Includes `_xlnm.*` internals (print
    // areas, print titles, autoFilter state) — no filtering here so
    // consumers see everything Excel persists.
    definedNames: DefinedName[];
    // Resolved xl/metadata.xml blocks. Null when the package has no
    // metadata part.
    metadata: WorkbookMetadata | null;
}

const NS = {
    main:  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    rel:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rels:  'http://schemas.openxmlformats.org/package/2006/relationships',
    xdr:   'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    a:     'http://schemas.openxmlformats.org/drawingml/2006/main',
    c:     'http://schemas.openxmlformats.org/drawingml/2006/chart',
    cx:    'http://schemas.microsoft.com/office/drawing/2014/chartex',
    tc:    'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
    xda:   'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray',
    // Excel 2010 x14 — slicers + form-control (ctrlProp) live in the same NS.
    // Form-control `<formControlPr>` is also valid in 2006/main without a
    // prefix; we match on localName to tolerate both.
    x14:   'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    // Excel 2013 x15 — timelines. (x14 slicers live in x14; timelines in x15.)
    x15:   'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
    // DrawingML diagram (SmartArt). The <dgm:relIds> child of a graphicFrame's
    // <a:graphicData> binds r:dm / r:lo / r:qs / r:cs rIds to data1.xml /
    // layout1.xml / quickStyle1.xml / colors1.xml respectively.
    dgm:   'http://schemas.openxmlformats.org/drawingml/2006/diagram',
};

// URI that marks a graphicFrame's <a:graphicData> as carrying a SmartArt
// diagram (rather than a chart or ink annotation). Identity constant — not
// used in a CSS selector or innerHTML.
const SMARTART_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

export class WorkbookParser {
    constructor(private _options: Options) {}

    parse(
        parts: Record<string, string>,
        media: Record<string, string> = {},
        embeddingBytes: Record<string, Uint8Array> = {},
    ): Workbook {
        const workbookXml = parts['xl/workbook.xml'];
        if (!workbookXml) throw new Error('xlsx-preview: workbook.xml missing');

        const sharedStrings = parts['xl/sharedStrings.xml']
            ? parseSharedStrings(parts['xl/sharedStrings.xml'])
            : [];

        const styles = parts['xl/styles.xml'] ? parseStyles(parts['xl/styles.xml']) : null;
        // Excel writes theme under xl/theme/theme1.xml. Producers (LibreOffice,
        // python-xlsx) follow the same convention. We look for any theme*.xml
        // under xl/theme/ and pick the first.
        const themePath = Object.keys(parts).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
        const theme = themePath ? parseTheme(parts[themePath]) : null;

        // Workbook rels → rId → target path. The <sheet r:id="rIdN"/>
        // attribute binds each sheet to the actual xml file. If rels are
        // missing, fall back to positional resolution for older / malformed
        // producers — still matches Excel, LibreOffice, and python-xlsx
        // output.
        const rels = parts['xl/_rels/workbook.xml.rels']
            ? parseRelationships(parts['xl/_rels/workbook.xml.rels'])
            : new Map<string, { target: string; type: string }>();

        // Workbook-wide person registry (threaded comments). The rel type
        // ends in "/person" and points at a package-wide person.xml. We fall
        // back to the conventional xl/persons/person.xml path if no rel is
        // declared but the file is present.
        const persons = resolvePersons(rels, parts);

        const sheetMeta = parseSheetList(workbookXml);
        const { date1904, definedNames } = parseWorkbookMeta(workbookXml);
        const metadata = parts['xl/metadata.xml']
            ? parseWorkbookMetadata(parts['xl/metadata.xml'])
            : null;
        // Content-type resolver — reads [Content_Types].xml once per workbook
        // so embedding-resolver lookups stay O(1). Falls back to a no-op map
        // when the part is absent (extension-only resolution still works).
        const contentTypes = parts['[Content_Types].xml']
            ? parseContentTypes(parts['[Content_Types].xml'])
            : { defaults: new Map<string, string>(), overrides: new Map<string, string>() };
        const sheets: Sheet[] = [];
        for (let i = 0; i < sheetMeta.length; i++) {
            const { name, rId, state } = sheetMeta[i];
            let xmlPath: string | null = null;
            const rel = rId ? rels.get(rId) : undefined;
            if (rel) {
                xmlPath = resolveWorkbookRelTarget(rel.target);
            }
            if (!xmlPath || !parts[xmlPath]) {
                // Positional fallback.
                xmlPath = `xl/worksheets/sheet${i + 1}.xml`;
            }
            const xml = parts[xmlPath];
            if (!xml) continue;
            const tables = resolveTablesForSheet(xmlPath, parts);
            const { images, charts, shapes, formControls, smartArt } = resolveDrawingsForSheet(xmlPath, parts, media, sharedStrings);
            const pivots = resolvePivotsForSheet(xmlPath, parts);
            const { slicers, timelines } = resolveSlicersAndTimelinesForSheet(xmlPath, parts);
            const comments = resolveCommentsForSheet(xmlPath, parts);
            const threadedComments = resolveThreadedCommentsForSheet(xmlPath, parts, persons);
            const hyperlinkTargets = resolveHyperlinkTargets(xmlPath, parts);
            const embeddings = resolveEmbeddingsForSheet(
                xmlPath, xml, parts, embeddingBytes, contentTypes,
                this._options.inlineEmbeddings === true,
                this._options.parseOleCfb === true,
            );
            sheets.push(parseSheet(name, state, xml, sharedStrings, tables, images, charts, shapes, formControls, embeddings, pivots, slicers, timelines, comments, threadedComments, hyperlinkTargets, i, definedNames, metadata, smartArt));
        }

        return { sheets, styles, theme, persons, date1904, definedNames, metadata };
    }
}

// Parsed projection of [Content_Types].xml. `defaults` maps lower-cased file
// extension → contentType (from `<Default Extension="…" ContentType="…"/>`).
// `overrides` maps package-absolute part name (with leading "/") → contentType
// (from `<Override PartName="/x/y.xlsx" ContentType="…"/>`). Resolution order:
// overrides first, then the file extension's default, then a literal fallback.
interface ContentTypeMap {
    defaults: Map<string, string>;  // extension (lower-cased) → MIME
    overrides: Map<string, string>; // part name (leading "/") → MIME
}

function parseContentTypes(xml: string): ContentTypeMap {
    const doc = parseXml(xml);
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    const ns = 'http://schemas.openxmlformats.org/package/2006/content-types';
    const defEls = doc.getElementsByTagNameNS(ns, 'Default');
    for (let i = 0; i < defEls.length; i++) {
        const el = defEls[i];
        const ext = el.getAttribute('Extension');
        const ct = el.getAttribute('ContentType');
        if (ext && ct) defaults.set(ext.toLowerCase(), ct);
    }
    const ovEls = doc.getElementsByTagNameNS(ns, 'Override');
    for (let i = 0; i < ovEls.length; i++) {
        const el = ovEls[i];
        const part = el.getAttribute('PartName');
        const ct = el.getAttribute('ContentType');
        if (part && ct) overrides.set(part, ct);
    }
    return { defaults, overrides };
}

// Resolve the contentType for a package-embedded part. Check the overrides
// map (needs a leading "/" package-absolute path), then the file extension's
// Default entry, then an extension-derived best-guess, then
// application/octet-stream as last-resort. Returns the raw contentType —
// callers apply sanitizeMediaMime before using the value as the `data:` URL's
// MIME.
function resolveContentType(path: string, map: ContentTypeMap): string {
    const pkgPath = path.startsWith('/') ? path : `/${path}`;
    const ov = map.overrides.get(pkgPath);
    if (ov) return ov;
    const dot = path.lastIndexOf('.');
    if (dot >= 0) {
        const ext = path.slice(dot + 1).toLowerCase();
        const def = map.defaults.get(ext);
        if (def) return def;
        // Best-effort extension-derived MIME. Keeps OLE ".bin" distinct so
        // callers can tell the two kinds apart. Anything else falls through
        // to application/octet-stream so sanitizeMediaMime keeps its teeth.
        switch (ext) {
            case 'bin':  return 'application/vnd.openxmlformats-officedocument.oleObject';
            case 'pdf':  return 'application/pdf';
            case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
            case 'docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            case 'pptx': return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
            case 'xls':  return 'application/vnd.ms-excel';
            case 'doc':  return 'application/msword';
            case 'ppt':  return 'application/vnd.ms-powerpoint';
            case 'txt':  return 'text/plain';
        }
    }
    return 'application/octet-stream';
}

// Resolve every OLE / package embedding referenced from a single sheet. Walks
// the sheet's rels for "/oleObject" (raw CFB bin) and "/package" (packaged
// real file) types, matches each to its anchor + progId via the sheet's
// `<oleObjects>` block, and falls through to an anchor-less "after-table"
// entry when no `<oleObject>` element points at the rel.
//
// ProgId is a hint Excel writes onto <oleObject progId="…"/>; we do NOT key
// any logic on it (XLSX strings are attacker-controlled) — the renderer
// surfaces it verbatim as a data-attribute through setAttribute only.
//
// `inline=true` flips dataUrl population on: payloads up to
// MAX_EMBEDDING_BYTES (via bytesToDataUrl) project through sanitizeMediaMime
// and land on the model; rejections stay null and the renderer emits the
// placeholder without a download link.
function resolveEmbeddingsForSheet(
    sheetPath: string,
    sheetXml: string,
    parts: Record<string, string>,
    embeddingBytes: Record<string, Uint8Array>,
    contentTypes: ContentTypeMap,
    inline: boolean,
    parseOleCfb: boolean,
): SheetEmbedding[] {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return [];
    const rels = parseRelationships(relsXml);
    const sheetDir = sheetPath.replace(/\/[^/]+$/, '');

    // Walk rels for embedding types. Keep the rel id so we can map back to
    // <oleObject r:id="…"/> entries in the sheet xml.
    interface EmbeddingRel {
        rId: string;
        kind: 'ole' | 'package';
        target: string;   // resolved package-relative path (no leading slash)
        fileName: string | null;
    }
    const embRels: EmbeddingRel[] = [];
    for (const [rId, rel] of rels) {
        // The rel type URL ends in "/oleObject" for OLE CFB bin payloads and
        // "/package" for a packaged real file. Everything else (drawings,
        // hyperlinks, tables, …) is already handled by other resolvers.
        const isOle = rel.type.endsWith('/oleObject');
        const isPackage = rel.type.endsWith('/package');
        if (!isOle && !isPackage) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${sheetDir}/${rel.target}`);
        const fileName = target.slice(target.lastIndexOf('/') + 1) || null;
        embRels.push({
            rId,
            kind: isOle ? 'ole' : 'package',
            target,
            fileName,
        });
    }
    if (embRels.length === 0) return [];

    // Map from rId → <oleObject> anchor metadata. An <oleObject progId="…"
    // shapeId="…" r:id="…"> child usually carries a sibling <objectPr>
    // wrapping an <anchor> with <from>/<to> cell markers; the ECMA shape is
    // variable — we pull col/row/endCol/endRow best-effort from whichever
    // `<from>` / `<to>` with child `<col>` / `<row>` text nodes is present.
    const anchorByRid = parseOleObjectAnchors(sheetXml);

    const out: SheetEmbedding[] = [];
    for (const rel of embRels) {
        const bytes = embeddingBytes[rel.target];
        if (!bytes) {
            // Rel points at a part that isn't in the embeddings directory (or
            // was dropped for some reason). Skip — there's nothing to surface.
            continue;
        }
        const contentType = resolveContentType(rel.target, contentTypes);
        const anchor = anchorByRid.get(rel.rId) ?? null;
        let dataUrl: string | null = null;
        if (inline) {
            // MIME must pass the allowlist; bytesToDataUrl returns null for
            // oversized payloads. Either rejection keeps dataUrl=null, which
            // tells the renderer to emit the placeholder without a download.
            const safeMime = sanitizeMediaMime(contentType);
            if (safeMime !== 'application/octet-stream' || contentType === 'application/octet-stream') {
                // Only project when the original MIME is itself in the
                // allowlist (so a hostile/unknown MIME doesn't silently decay
                // to a working octet-stream download). Raw OLE CFB ".bin"
                // payloads sit on 'application/vnd.openxmlformats-officedocument.oleObject'
                // which is outside the allowlist and therefore stays null —
                // intentional: we don't want to hand users a .bin OLE container.
                dataUrl = bytesToDataUrl(bytes, contentType);
            }
        }
        // Opt-in CFB extraction. Only ole payloads are CFB containers;
        // packaged real files are their own format and shouldn't be parsed
        // as CFB. parseCfb returns null on any error, so consumers can't
        // distinguish "flag off" from "invalid CFB" without comparing the
        // flag — intentional, keeps the API shape simple.
        const cfb = (parseOleCfb && rel.kind === 'ole') ? parseCfb(bytes) : null;
        out.push({
            kind: rel.kind,
            contentType,
            fileName: rel.fileName,
            progId: anchor?.progId ?? null,
            col: anchor?.col ?? null,
            row: anchor?.row ?? null,
            endCol: anchor?.endCol ?? null,
            endRow: anchor?.endRow ?? null,
            size: bytes.byteLength,
            dataUrl,
            altText: anchor?.altText ?? null,
            cfb,
        });
    }
    return out;
}

// Parse any <oleObject progId="…" r:id="…"/> entries from a sheet xml and
// return a map keyed by relationship id. Anchor coordinates come from the
// sibling <objectPr>/<anchor>/<from|to> descendants when present. Shape-id
// fallbacks (for legacy form controls) are ignored — xlsxjs doesn't surface
// them anywhere else.
function parseOleObjectAnchors(xml: string): Map<string, {
    progId: string | null;
    col: number | null; row: number | null;
    endCol: number | null; endRow: number | null;
    altText: string | null;
}> {
    const out = new Map<string, {
        progId: string | null;
        col: number | null; row: number | null;
        endCol: number | null; endRow: number | null;
        altText: string | null;
    }>();
    const doc = parseXml(xml);
    // Modern Excel wraps oleObjects in an <mc:AlternateContent> so the
    // element may sit at multiple depths. getElementsByTagNameNS walks the
    // full subtree so both the main and fallback copies surface; we keep the
    // first one per rId (the <Choice> before <Fallback> is the preferred
    // copy per MCE rules).
    const els = doc.getElementsByTagNameNS(NS.main, 'oleObject');
    for (let i = 0; i < els.length; i++) {
        const el = els[i];
        const rId = el.getAttributeNS(NS.rel, 'id');
        if (!rId || out.has(rId)) continue;
        const progId = el.getAttribute('progId') || null;
        // <objectPr> carries an <anchor> with <from>/<to> xdr-like markers.
        // Older Excel writes <oleObject><twoCellAnchor>…</twoCellAnchor></oleObject>
        // directly; support both shapes by scanning descendants.
        let col: number | null = null;
        let row: number | null = null;
        let endCol: number | null = null;
        let endRow: number | null = null;
        // Look in the xdr namespace (the DrawingML anchor namespace) first —
        // Excel's modern output uses xdr for these markers.
        const findFrom = el.getElementsByTagNameNS(NS.xdr, 'from').item(0)
            ?? el.getElementsByTagNameNS(NS.main, 'from').item(0);
        const findTo = el.getElementsByTagNameNS(NS.xdr, 'to').item(0)
            ?? el.getElementsByTagNameNS(NS.main, 'to').item(0);
        if (findFrom) {
            col = anchorCellValue(findFrom, 'col');
            row = anchorCellValue(findFrom, 'row');
            // fallback to main-namespace children when xdr lookup missed
            if (col === null) {
                const c = findFrom.getElementsByTagNameNS(NS.main, 'col').item(0);
                col = c ? Number(c.textContent) : null;
                if (col !== null && !Number.isFinite(col)) col = null;
            }
            if (row === null) {
                const r = findFrom.getElementsByTagNameNS(NS.main, 'row').item(0);
                row = r ? Number(r.textContent) : null;
                if (row !== null && !Number.isFinite(row)) row = null;
            }
        }
        if (findTo) {
            endCol = anchorCellValue(findTo, 'col');
            endRow = anchorCellValue(findTo, 'row');
            if (endCol === null) {
                const c = findTo.getElementsByTagNameNS(NS.main, 'col').item(0);
                endCol = c ? Number(c.textContent) : null;
                if (endCol !== null && !Number.isFinite(endCol)) endCol = null;
            }
            if (endRow === null) {
                const r = findTo.getElementsByTagNameNS(NS.main, 'row').item(0);
                endRow = r ? Number(r.textContent) : null;
                if (endRow !== null && !Number.isFinite(endRow)) endRow = null;
            }
        }
        out.set(rId, { progId, col, row, endCol, endRow, altText: null });
    }
    return out;
}

// Parse xl/metadata.xml into a resolved WorkbookMetadata. We care about two
// things:
//   - <metadataTypes> — name → index map so we know which cellMetadata
//     blocks are dynamic-array (XLDAPR) vs. other types (XLRICHVALUE etc.).
//   - <futureMetadata name="XLDAPR"> — each <bk>/<extLst>/<ext>/<xda:…> child
//     indicates whether that future-metadata index is a dynamic-array flag
//     (fDynamic="1" = spill anchor; "0" = suppressed).
//   - <cellMetadata> / <valueMetadata> — each <bk> wraps one or more <rc>
//     entries; we read the first <rc t="typeIndex" v="valueIndex"/> pair.
//     The type's name + the resolved futureMetadata block combine to tell
//     the renderer whether this block flips Cell.isSpillAnchor.
export function parseWorkbookMetadata(xml: string): WorkbookMetadata {
    const doc = parseXml(xml);

    // 1-based XML index → type name. Absent / out-of-range indexes fall
    // through to null in the resolver.
    const typeNames: string[] = [];
    const typeEls = doc.getElementsByTagNameNS(NS.main, 'metadataType');
    for (let i = 0; i < typeEls.length; i++) {
        typeNames.push(typeEls[i].getAttribute('name') ?? '');
    }

    // Map metadata type name → list of bool flags (one per <bk> entry).
    // For XLDAPR, the flag is fDynamic; other types stay empty (we don't
    // interpret them). The <bk> ordering is 0-based from xlsxjs's perspective
    // and matches the `v` attribute on <rc t="…" v="…"/>.
    const futureFlagsByType = new Map<string, boolean[]>();
    const fmEls = doc.getElementsByTagNameNS(NS.main, 'futureMetadata');
    for (let i = 0; i < fmEls.length; i++) {
        const fm = fmEls[i];
        const name = fm.getAttribute('name') ?? '';
        const bkEls = fm.getElementsByTagNameNS(NS.main, 'bk');
        const flags: boolean[] = [];
        for (let j = 0; j < bkEls.length; j++) {
            const bk = bkEls[j];
            let dynamic = false;
            // Dynamic-array flag lives under bk/extLst/ext/xda:dynamicArrayProperties.
            const daProps = bk.getElementsByTagNameNS(NS.xda, 'dynamicArrayProperties').item(0);
            if (daProps) {
                const attr = daProps.getAttribute('fDynamic');
                dynamic = attr === '1' || attr === 'true';
            }
            flags.push(dynamic);
        }
        futureFlagsByType.set(name, flags);
    }

    // Resolve a single <bk> block under <cellMetadata> or <valueMetadata>.
    // Each <bk> has one or more <rc t="typeIdx1based" v="valueIdx0based"/>;
    // we take the first, which is Excel's convention for the XLDAPR case.
    const resolveBlock = (bk: Element): CellMetadataBlock => {
        const rc = bk.getElementsByTagNameNS(NS.main, 'rc').item(0);
        const t = rc ? Number(rc.getAttribute('t')) : NaN;
        const v = rc ? Number(rc.getAttribute('v')) : NaN;
        const typeIndex = Number.isFinite(t) && t >= 1 ? t - 1 : -1;
        const typeName = typeIndex >= 0 && typeIndex < typeNames.length
            ? typeNames[typeIndex]
            : '';
        let dynamicArray = false;
        if (typeName === 'XLDAPR') {
            const flags = futureFlagsByType.get('XLDAPR');
            if (flags && Number.isFinite(v) && v >= 0 && v < flags.length) {
                dynamicArray = flags[v];
            }
        }
        return { typeIndex, dynamicArray };
    };

    const readList = (listTag: string): CellMetadataBlock[] => {
        const out: CellMetadataBlock[] = [];
        const wrap = doc.getElementsByTagNameNS(NS.main, listTag).item(0);
        if (!wrap) return out;
        const bks = wrap.getElementsByTagNameNS(NS.main, 'bk');
        for (let i = 0; i < bks.length; i++) out.push(resolveBlock(bks[i]));
        return out;
    };

    return {
        cellMetadata: readList('cellMetadata'),
        valueMetadata: readList('valueMetadata'),
    };
}

// Parse the small set of workbook-wide flags we care about from xl/workbook.xml.
// Today that covers `<workbookPr date1904="1"/>` plus the `<definedNames>`
// registry; more flags can join this reader as they come online.
function parseWorkbookMeta(workbookXml: string): { date1904: boolean; definedNames: DefinedName[] } {
    const doc = parseXml(workbookXml);
    const pr = doc.getElementsByTagNameNS(NS.main, 'workbookPr').item(0);
    let date1904 = false;
    if (pr) {
        const attr = pr.getAttribute('date1904');
        date1904 = attr === '1' || attr === 'true';
    }
    const definedNames: DefinedName[] = [];
    const dnEls = doc.getElementsByTagNameNS(NS.main, 'definedName');
    for (let i = 0; i < dnEls.length; i++) {
        const el = dnEls[i];
        const name = el.getAttribute('name');
        if (!name) continue;
        const localAttr = el.getAttribute('localSheetId');
        const localSheetId = localAttr != null && Number.isFinite(Number(localAttr))
            ? Number(localAttr)
            : null;
        const hidden = el.getAttribute('hidden') === '1';
        const formula = el.textContent ?? '';
        definedNames.push({ name, localSheetId, formula, hidden });
    }
    return { date1904, definedNames };
}

// xl/persons/person.xml holds the GUID → displayName registry shared by
// every threaded comment in the package. Excel 365 writes exactly one
// such file; the rel type is a Microsoft-specific URL ending in
// "/person" (not the office-document namespace). If the rel is missing
// but the file is there, fall back to the conventional path.
function resolvePersons(
    workbookRels: Map<string, { target: string; type: string }>,
    parts: Record<string, string>,
): Map<string, string> {
    let xmlPath: string | null = null;
    for (const [, rel] of workbookRels) {
        if (rel.type.endsWith('/person')) {
            xmlPath = resolveWorkbookRelTarget(rel.target);
            break;
        }
    }
    if (!xmlPath || !parts[xmlPath]) {
        // Fallback: convention is xl/persons/person.xml, but we accept any
        // xl/persons/*.xml file if only one is present.
        const candidates = Object.keys(parts).filter((p) => /^xl\/persons\/.*\.xml$/i.test(p));
        xmlPath = candidates[0] ?? null;
    }
    if (!xmlPath || !parts[xmlPath]) return new Map();
    return parsePersons(parts[xmlPath]);
}

function parsePersons(xml: string): Map<string, string> {
    const out = new Map<string, string>();
    const doc = parseXml(xml);
    const nodes = doc.getElementsByTagNameNS(NS.tc, 'person');
    for (let i = 0; i < nodes.length; i++) {
        const el = nodes[i];
        const id = el.getAttribute('id');
        const displayName = el.getAttribute('displayName');
        if (!id) continue;
        out.set(id, displayName ?? '');
    }
    return out;
}

// Resolve every threadedComments*.xml referenced from the sheet's rels
// (type ending in "/threadedComment") and parse it into ThreadedCommentEntry[].
// Each personId is resolved to a display name via the workbook-wide
// registry; unknown IDs fall back to `null` (renderer displays "Unknown").
function resolveThreadedCommentsForSheet(
    sheetPath: string,
    parts: Record<string, string>,
    persons: Map<string, string>,
): ThreadedCommentEntry[] {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return [];
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const out: ThreadedCommentEntry[] = [];
    for (const [, rel] of rels) {
        if (!rel.type.endsWith('/threadedComment')) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const xml = parts[target];
        if (!xml) continue;
        out.push(...parseThreadedComments(xml, persons));
    }
    return out;
}

// Parse a single threadedComment{N}.xml into a flat list of entries.
// Thread structure (parent vs. reply) is preserved via parentId; the
// renderer groups them back into threads at display time.
export function parseThreadedComments(xml: string, persons: Map<string, string>): ThreadedCommentEntry[] {
    const doc = parseXml(xml);
    const nodes = doc.getElementsByTagNameNS(NS.tc, 'threadedComment');
    const out: ThreadedCommentEntry[] = [];
    for (let i = 0; i < nodes.length; i++) {
        const el = nodes[i];
        const ref = el.getAttribute('ref');
        const id = el.getAttribute('id');
        if (!ref || !id) continue;
        const cell = parseCellRef(ref);
        if (!cell) continue;
        const personId = el.getAttribute('personId');
        const author = personId && persons.has(personId) ? persons.get(personId)! : null;
        const textEl = el.getElementsByTagNameNS(NS.tc, 'text').item(0);
        const text = textEl?.textContent ?? '';
        const date = el.getAttribute('dT');
        const parentId = el.getAttribute('parentId');
        out.push({
            id,
            col: cell.col,
            row: cell.row,
            author: author && author.length > 0 ? author : null,
            date: date ?? null,
            text,
            parentId: parentId ?? null,
        });
    }
    return out;
}

function parseRelationships(xml: string): Map<string, { target: string; type: string }> {
    const out = new Map<string, { target: string; type: string }>();
    const doc = parseXml(xml);
    // Relationships elements sit in the package-relationships namespace. We
    // walk childNodes rather than querying by tag name to be robust to
    // producers that omit the xmlns prefix.
    const rootEls = doc.getElementsByTagNameNS(NS.rels, 'Relationship');
    for (let i = 0; i < rootEls.length; i++) {
        const el = rootEls[i];
        const id = el.getAttribute('Id');
        const target = el.getAttribute('Target');
        const type = el.getAttribute('Type') ?? '';
        if (!id || !target) continue;
        out.set(id, { target, type });
    }
    return out;
}

// Target paths in xl/_rels/workbook.xml.rels are relative to the workbook
// part (xl/workbook.xml), so they resolve against xl/. A target that begins
// with "/" is package-absolute and used verbatim (minus the leading slash).
function resolveWorkbookRelTarget(target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    // Strip any leading "./".
    const t = target.replace(/^\.\//, '');
    return `xl/${t}`;
}

// For a sheet part at e.g. "xl/worksheets/sheet1.xml", resolve + parse any
// tables referenced via its rels part. Relative targets resolve against
// the sheet's containing directory.
function resolveTablesForSheet(sheetPath: string, parts: Record<string, string>): TableDef[] {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return [];
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const out: TableDef[] = [];
    for (const [, rel] of rels) {
        if (!rel.type.endsWith('/table')) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : `${dir}/${rel.target}`.replace(/\/\.\//g, '/');
        // Normalise "../tables/…" forms.
        const normalised = normaliseRelPath(target);
        const xml = parts[normalised];
        if (!xml) continue;
        const parsed = parseTable(xml);
        if (parsed) out.push(parsed);
    }
    return out;
}

function normaliseRelPath(path: string): string {
    const stack: string[] = [];
    for (const seg of path.split('/')) {
        if (seg === '' || seg === '.') continue;
        if (seg === '..') stack.pop();
        else stack.push(seg);
    }
    return stack.join('/');
}

function resolveDrawingsForSheet(
    sheetPath: string,
    parts: Record<string, string>,
    media: Record<string, string>,
    sharedStrings: SharedString[] = [],
): { images: SheetImage[]; charts: SheetChart[]; shapes: SheetShape[]; formControls: SheetFormControl[]; smartArt: SheetSmartArt[] } {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return { images: [], charts: [], shapes: [], formControls: [], smartArt: [] };
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const images: SheetImage[] = [];
    const charts: SheetChart[] = [];
    const shapes: SheetShape[] = [];
    const formControls: SheetFormControl[] = [];
    const smartArt: SheetSmartArt[] = [];
    for (const [, rel] of rels) {
        if (!rel.type.endsWith('/drawing')) continue;
        const drawingPath = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const drawingXml = parts[drawingPath];
        if (!drawingXml) continue;
        // The drawing may have its own rels part binding rIds to media
        // files and chart parts. Shape-only drawings (text boxes /
        // connectors) can legitimately ship without a rels file, so a
        // missing rels part no longer aborts the walk.
        const drawingRelsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const drawingRelsXml = parts[drawingRelsPath];
        const drawingRels = drawingRelsXml
            ? parseRelationships(drawingRelsXml)
            : new Map<string, { target: string; type: string }>();
        const drawingDir = drawingPath.replace(/\/[^/]+$/, '');
        const parsed = parseDrawing(drawingXml, drawingRels, drawingDir, parts, media, sharedStrings);
        images.push(...parsed.images);
        charts.push(...parsed.charts);
        shapes.push(...parsed.shapes);
        formControls.push(...parsed.formControls);
        smartArt.push(...parsed.smartArt);
    }
    return { images, charts, shapes, formControls, smartArt };
}

function resolvePivotsForSheet(
    sheetPath: string,
    parts: Record<string, string>,
): SheetPivot[] {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return [];
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const out: SheetPivot[] = [];
    for (const [, rel] of rels) {
        // The official rel type ends in "/pivotTable".
        if (!rel.type.endsWith('/pivotTable')) continue;
        const pivotPath = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const xml = parts[pivotPath];
        if (!xml) continue;
        const parsed = parsePivotTable(xml);
        if (parsed) out.push(parsed);
    }
    return out;
}

function parsePivotTable(xml: string): SheetPivot | null {
    const doc = parseXml(xml);
    const def = doc.getElementsByTagNameNS(NS.main, 'pivotTableDefinition').item(0);
    if (!def) return null;
    const name = def.getAttribute('name') ?? '';
    const loc = def.getElementsByTagNameNS(NS.main, 'location').item(0);
    if (!loc) return null;
    const ref = loc.getAttribute('ref');
    if (!ref) return null;
    const parts = ref.split(':');
    const a = parseCellRef(parts[0]);
    const b = parts[1] ? parseCellRef(parts[1]) : a;
    if (!a || !b) return null;
    return {
        name,
        col: Math.min(a.col, b.col),
        row: Math.min(a.row, b.row),
        endCol: Math.max(a.col, b.col),
        endRow: Math.max(a.row, b.row),
    };
}

// Slicer / timeline detection. The sheet's rels can carry three flavours of
// link into the slicer / timeline chain:
//   1. A `.../slicer` rel → xl/slicers/slicerN.xml  (one <slicers> list
//      per file, each <slicer> references a cache by @name).
//   2. A `.../timeline` rel → xl/timelines/timelineN.xml (one <timelines>
//      list per file, each <timeline> references a timelineCache by @name).
//   3. Nothing — the references live behind an `<extLst>` block on the
//      worksheet (URIs {A8765BA9-…} for slicers, {7E03D99C-…} for
//      timelines). Both URIs wrap an `<x14:slicerList>` or
//      `<x15:timelineRefs>` with `<x14:slicer r:id="…"/>` / similar inside
//      the same sheet's rels — we still resolve through the sheet rels.
//
// Rather than walk the extLst explicitly, we scan the sheet's rels for
// anything whose type ends in /slicer or /timeline; that catches both the
// direct-rel producers (Excel 2013+, some LibreOffice versions) and the
// extLst-wrapped producers (Excel 2010/2016 "classic" slicers) without
// needing two code paths.
//
// Caches live at xl/slicerCaches/ and xl/timelineCaches/. Each cache's
// name attribute is what the per-sheet widget references via @cache. We
// build two workbook-wide name → XML Maps so both widgets can look up
// their cache in O(1) without re-resolving package-wide rels.
function resolveSlicersAndTimelinesForSheet(
    sheetPath: string,
    parts: Record<string, string>,
): { slicers: SheetSlicer[]; timelines: SheetTimeline[] } {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    const slicers: SheetSlicer[] = [];
    const timelines: SheetTimeline[] = [];
    if (!relsXml) return { slicers, timelines };
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');

    // Index every slicerCache / timelineCache part in the package by the
    // cache's declared @name. Using a Map (not a plain object) keeps
    // attacker-controlled cache names off Object.prototype.
    const slicerCacheByName = indexSlicerCaches(parts);
    const timelineCacheByName = indexTimelineCaches(parts);

    for (const [, rel] of rels) {
        const type = rel.type;
        const isSlicer = type.endsWith('/slicer');
        const isTimeline = type.endsWith('/timeline');
        if (!isSlicer && !isTimeline) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const xml = parts[target];
        if (!xml) continue;
        if (isSlicer) {
            slicers.push(...parseSlicerFile(xml, slicerCacheByName));
        } else {
            timelines.push(...parseTimelineFile(xml, timelineCacheByName));
        }
    }
    return { slicers, timelines };
}

// Walk the package for xl/slicerCaches/*.xml and key by the <slicerCacheDefinition @name>
// attribute. The x14:slicerCacheDefinition element is the root of a slicer
// cache file; producers may omit the x14 prefix and write an unprefixed
// <slicerCacheDefinition>, so we match by localName.
function indexSlicerCaches(parts: Record<string, string>): Map<string, string> {
    const out = new Map<string, string>();
    for (const path of Object.keys(parts)) {
        if (!/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path)) continue;
        const xml = parts[path];
        if (!xml) continue;
        const name = peekRootAttr(xml, 'slicerCacheDefinition', 'name');
        if (name) out.set(name, xml);
    }
    return out;
}

function indexTimelineCaches(parts: Record<string, string>): Map<string, string> {
    const out = new Map<string, string>();
    for (const path of Object.keys(parts)) {
        if (!/^xl\/timelineCaches\/[^/]+\.xml$/i.test(path)) continue;
        const xml = parts[path];
        if (!xml) continue;
        // timelineCacheDefinition lives in the x15 namespace; we still
        // localName-match so unprefixed producers resolve too.
        const name = peekRootAttr(xml, 'timelineCacheDefinition', 'name');
        if (name) out.set(name, xml);
    }
    return out;
}

// Pluck the first element whose localName matches `rootLocal` from `xml`
// and return its `attr` attribute. Used to build cache-name indexes without
// parsing the full document twice.
function peekRootAttr(xml: string, rootLocal: string, attr: string): string | null {
    const doc = parseXml(xml);
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName === rootLocal) {
            return el.getAttribute(attr);
        }
    }
    return null;
}

// Parse one xl/slicers/slicerN.xml file. The root is `<slicers>` carrying
// one or more `<slicer>` children. Each slicer references a cache by
// @cache; we look that up in the workbook-wide index and pull the
// sourceName + selectedItems off the cache's tabular block.
function parseSlicerFile(xml: string, slicerCacheByName: Map<string, string>): SheetSlicer[] {
    const doc = parseXml(xml);
    const out: SheetSlicer[] = [];
    // Match by localName (x14:slicer or unprefixed <slicer>). We walk the
    // whole doc rather than using getElementsByTagNameNS so we don't miss
    // entries written under a differently-prefixed namespace.
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName !== 'slicer') continue;
        const name = el.getAttribute('name');
        const cache = el.getAttribute('cache');
        if (!name || !cache) continue;
        const caption = el.getAttribute('caption');
        const columnCountAttr = el.getAttribute('columnCount');
        const columnCount = numOrNull(columnCountAttr);
        const style = el.getAttribute('style') || null;
        // @showCaption defaults to true; only `"0"`/`"false"` flips it off.
        const showCaptionAttr = el.getAttribute('showCaption');
        const showCaption = !(showCaptionAttr === '0' || showCaptionAttr === 'false');
        const rowHeightAttr = el.getAttribute('rowHeight');
        const rowHeight = numOrNull(rowHeightAttr);

        const cacheXml = slicerCacheByName.get(cache);
        let sourceName: string | null = null;
        let selectedItems: string[] = [];
        let allItems: string[] = [];
        if (cacheXml) {
            const details = readSlicerCache(cacheXml);
            sourceName = details.sourceName;
            selectedItems = details.selectedItems;
            allItems = details.allItems;
        }
        out.push({
            name,
            caption: caption || null,
            cache,
            sourceName,
            columnCount,
            style,
            showCaption,
            rowHeight,
            selectedItems,
            allItems,
        });
    }
    return out;
}

// Parse one xl/timelines/timelineN.xml file. The root is `<timelines>` with
// one or more `<timeline>` children. Excel writes the currently-displayed
// level + selected range either as attributes directly on `<timeline>` or
// inside a `<timelineViewState>` child — we read the level off either
// location, preferring the inner state when both are present.
function parseTimelineFile(xml: string, timelineCacheByName: Map<string, string>): SheetTimeline[] {
    const doc = parseXml(xml);
    const out: SheetTimeline[] = [];
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName !== 'timeline') continue;
        const name = el.getAttribute('name');
        const cache = el.getAttribute('cache');
        if (!name || !cache) continue;
        const caption = el.getAttribute('caption');
        const style = el.getAttribute('style') || null;
        const showHeader = !truthyZero(el.getAttribute('showHeader'));
        const showSelectionLabel = !truthyZero(el.getAttribute('showSelectionLabel'));
        const showTimeLevel = !truthyZero(el.getAttribute('showTimeLevel'));
        const showHorizontalScrollbar = !truthyZero(el.getAttribute('showHorizontalScrollbar'));

        // View state can be a nested `<timelineViewState>` with level +
        // selection range, OR the same attributes can appear directly on
        // `<timeline>`. We prefer the nested block when present.
        let level: SheetTimeline['level'] = null;
        let selectedRange: { start: string | null; end: string | null } | null = null;
        const nested = firstDescendantByLocalName(el, 'timelineViewState');
        const src = nested ?? el;
        const levelAttr = src.getAttribute('level');
        level = normaliseTimelineLevel(levelAttr);
        // Modern producers write selection as `<selection startDate=… endDate=…/>`
        // under timelineViewState. Older producers write `startDate`/`endDate`
        // attributes directly on the parent. Try both.
        const selectionEl = firstDescendantByLocalName(src, 'selection') ?? src;
        const startAttr = selectionEl.getAttribute('startDate');
        const endAttr = selectionEl.getAttribute('endDate');
        if (startAttr !== null || endAttr !== null) {
            selectedRange = { start: startAttr, end: endAttr };
        }

        const cacheXml = timelineCacheByName.get(cache);
        let sourceName: string | null = null;
        let bounds: { min: string; max: string } | null = null;
        if (cacheXml) {
            const details = readTimelineCache(cacheXml);
            sourceName = details.sourceName;
            bounds = details.bounds;
        }
        out.push({
            name,
            caption: caption || null,
            cache,
            sourceName,
            level,
            selectedRange,
            showHeader,
            showSelectionLabel,
            showTimeLevel,
            showHorizontalScrollbar,
            style,
            bounds,
        });
    }
    return out;
}

// Inspect a slicer cache's tabular block. Returns the source field name (the
// pivot column the slicer drives) plus every item currently flagged
// selected. A producer may spell the root as <slicerCacheDefinition> with
// @sourceName, OR put the source name in an inner <pivotTable name="…"/>
// reference — we only promise the outer @sourceName form here and return
// null when that's missing.
function readSlicerCache(xml: string): { sourceName: string | null; selectedItems: string[]; allItems: string[] } {
    const doc = parseXml(xml);
    let sourceName: string | null = null;
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName === 'slicerCacheDefinition') {
            sourceName = el.getAttribute('sourceName');
            break;
        }
    }
    const selectedItems: string[] = [];
    const allItems: string[] = [];
    // Each <tabular> block carries an <items> list with <i x="…" s="1"/>
    // entries. The `@s="1"` flag marks a selected item; the displayed
    // text lives either as @n ("name") on the item or via a parent map.
    // We accept either @n, the inner textContent, or a numeric index into
    // a sibling `<value>` list — but in practice the Excel-written form
    // under the tabular/items chain uses @n for string items. A numeric
    // index without a resolvable value is dropped.
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName !== 'tabular') continue;
        const items = firstDescendantByLocalName(el, 'items');
        if (!items) continue;
        for (let j = 0; j < items.childNodes.length; j++) {
            const node = items.childNodes[j];
            if (node.nodeType !== 1) continue;
            const item = node as Element;
            if (item.localName !== 'i') continue;
            const name = item.getAttribute('n') ?? item.textContent;
            if (!name || name.length === 0) continue;
            allItems.push(name);
            const selected = item.getAttribute('s');
            // Excel writes `s="1"` for selected items (explicitly filtered
            // in) and either omits the attribute or writes `s="0"` for
            // unselected. We treat only the truthy form as a positive
            // selection signal.
            if (selected === '1' || selected === 'true') selectedItems.push(name);
        }
    }
    return { sourceName, selectedItems, allItems };
}

function readTimelineCache(xml: string): { sourceName: string | null; bounds: { min: string; max: string } | null } {
    const doc = parseXml(xml);
    let sourceName: string | null = null;
    let bounds: { min: string; max: string } | null = null;
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName === 'timelineCacheDefinition') {
            sourceName = el.getAttribute('sourceName');
        }
        // Bounds can live as `<bounds startDate=… endDate=…/>` under
        // `<state>`, or as `<cache startDate=… endDate=…/>` under the
        // `<pivotTables>` block in older producers. We prefer the named
        // `<bounds>` element and only emit when both endpoints resolve.
        if (el.localName === 'bounds') {
            const minAttr = el.getAttribute('startDate');
            const maxAttr = el.getAttribute('endDate');
            if (minAttr && maxAttr) {
                bounds = { min: minAttr, max: maxAttr };
            }
        }
    }
    return { sourceName, bounds };
}

function firstDescendantByLocalName(root: Element, localName: string): Element | null {
    const all = root.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        if (all[i].localName === localName) return all[i];
    }
    return null;
}

function numOrNull(attr: string | null): number | null {
    if (attr === null || attr === '') return null;
    const n = Number(attr);
    return Number.isFinite(n) ? n : null;
}

function truthyZero(v: string | null): boolean {
    return v === '0' || v === 'false';
}

function normaliseTimelineLevel(raw: string | null): SheetTimeline['level'] {
    if (!raw) return null;
    // Excel 2013+ writes the level as a numeric code ("0".."3") on the
    // timeline element; older/niche producers spell it out
    // ("months" / "Months"). Accept both forms and map to the long name.
    const lower = raw.toLowerCase();
    if (lower === 'years' || lower === 'quarters' || lower === 'months' || lower === 'days') {
        return lower;
    }
    // Numeric form: 0 = years, 1 = quarters, 2 = months, 3 = days. Any
    // other numeric value is an unknown level, so we fall through to null.
    const numericMap: Record<string, SheetTimeline['level']> = {
        '0': 'years',
        '1': 'quarters',
        '2': 'months',
        '3': 'days',
    };
    return numericMap[raw] ?? null;
}

// Classic comments: xl/comments{N}.xml. Bound to a sheet through the
// sheet's rels part (type ".../relationships/comments"). xlsxjs does not
// parse the sibling VML drawing — the bubble layout is irrelevant to an
// inline marker — only the anchor + author + text.
function resolveCommentsForSheet(sheetPath: string, parts: Record<string, string>): SheetComment[] {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return [];
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const out: SheetComment[] = [];
    for (const [, rel] of rels) {
        if (!rel.type.endsWith('/comments')) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const xml = parts[target];
        if (!xml) continue;
        out.push(...parseComments(xml));
    }
    return out;
}

// Resolve the sheet's rels to produce an rId → target URL map for hyperlink
// entries. Only hyperlink-typed rels surface (the sheet's rels part also
// carries drawings/tables/comments). `target` is the URL verbatim — we do
// NOT sanitize here; the renderer applies the URL allowlist when it wraps
// the cell in an anchor. External hyperlinks use the TargetMode="External"
// attribute in the rels but the URL itself lives in `Target`; we carry the
// whole string through so consumers that want full fidelity can inspect it.
function resolveHyperlinkTargets(
    sheetPath: string,
    parts: Record<string, string>,
): Map<string, string> {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return new Map();
    const rels = parseRelationships(relsXml);
    const out = new Map<string, string>();
    for (const [id, rel] of rels) {
        if (!rel.type.endsWith('/hyperlink')) continue;
        out.set(id, rel.target);
    }
    return out;
}

function parseComments(xml: string): SheetComment[] {
    const doc = parseXml(xml);
    // Authors: positional array indexed by authorId.
    const authors: string[] = [];
    const authorEls = doc.getElementsByTagNameNS(NS.main, 'author');
    for (let i = 0; i < authorEls.length; i++) {
        authors.push(authorEls[i].textContent ?? '');
    }
    const out: SheetComment[] = [];
    const commentEls = doc.getElementsByTagNameNS(NS.main, 'comment');
    for (let i = 0; i < commentEls.length; i++) {
        const c = commentEls[i];
        const ref = c.getAttribute('ref');
        if (!ref) continue;
        const parsed = parseCellRef(ref);
        if (!parsed) continue;
        const authorIdAttr = c.getAttribute('authorId');
        const authorIdx = authorIdAttr != null ? Number(authorIdAttr) : NaN;
        const author = Number.isFinite(authorIdx) && authorIdx >= 0 && authorIdx < authors.length
            ? authors[authorIdx]
            : null;
        // The <text> child has the same shape as a shared-string <si>:
        // plain <t>, or <r>/<rPr>/<t> runs. parseSi handles both.
        const textEl = c.getElementsByTagNameNS(NS.main, 'text').item(0);
        const body = textEl ? parseSi(textEl) : { text: '', runs: null, phonetics: null };
        out.push({
            col: parsed.col,
            row: parsed.row,
            author: author && author.length > 0 ? author : null,
            text: body.text,
            runs: body.runs,
        });
    }
    return out;
}

function parseDrawing(
    xml: string,
    rels: Map<string, { target: string; type: string }>,
    drawingDir: string,
    parts: Record<string, string>,
    media: Record<string, string>,
    sharedStrings: SharedString[] = [],
): { images: SheetImage[]; charts: SheetChart[]; shapes: SheetShape[]; formControls: SheetFormControl[]; smartArt: SheetSmartArt[] } {
    const doc = parseXml(xml);
    const images: SheetImage[] = [];
    const charts: SheetChart[] = [];
    const shapes: SheetShape[] = [];
    const formControls: SheetFormControl[] = [];
    const smartArt: SheetSmartArt[] = [];
    // DrawingML anchors come in three flavours; each carries <pic> /
    // <graphicFrame> / <sp> / <cxnSp> children we care about. We tag the
    // mode here so the model + renderer can apply the right sizing path.
    const labelled: { el: Element; mode: 'twoCell' | 'oneCell' | 'absolute' }[] = [
        ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'twoCellAnchor'))
            .map((el) => ({ el, mode: 'twoCell' as const })),
        ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'oneCellAnchor'))
            .map((el) => ({ el, mode: 'oneCell' as const })),
        ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'absoluteAnchor'))
            .map((el) => ({ el, mode: 'absolute' as const })),
    ];
    for (const { el: anchor, mode } of labelled) {
        const from = anchor.getElementsByTagNameNS(NS.xdr, 'from').item(0);
        const to = mode === 'twoCell'
            ? anchor.getElementsByTagNameNS(NS.xdr, 'to').item(0)
            : null;
        const col = from ? anchorCellValue(from, 'col') : null;
        const row = from ? anchorCellValue(from, 'row') : null;
        const endCol = to ? anchorCellValue(to, 'col') : null;
        const endRow = to ? anchorCellValue(to, 'row') : null;
        const colOff = from ? anchorCellValue(from, 'colOff') : null;
        const rowOff = from ? anchorCellValue(from, 'rowOff') : null;

        // <xdr:ext cx cy/> is a direct child of one-cell and absolute anchors.
        // It sizes the drawing in EMU. Two-cell anchors don't carry it (size
        // is derived from the anchor cells); we still tolerate one if it
        // appears in a tolerant producer's output, but only look shallow so
        // the a:ext inside pic/spPr/a:xfrm isn't accidentally matched (it's
        // in the `a` namespace, so getElementsByTagNameNS(xdr, 'ext') won't
        // catch it — but we're conservative for future-proofing).
        const ext = findDirectChildNS(anchor, NS.xdr, 'ext');
        const extCx = ext ? Number(ext.getAttribute('cx')) : NaN;
        const extCy = ext ? Number(ext.getAttribute('cy')) : NaN;
        const widthEmu = Number.isFinite(extCx) ? extCx : null;
        const heightEmu = Number.isFinite(extCy) ? extCy : null;

        // <xdr:pos x y/> is a direct child of absoluteAnchor only.
        let absoluteX: number | null = null;
        let absoluteY: number | null = null;
        if (mode === 'absolute') {
            const pos = findDirectChildNS(anchor, NS.xdr, 'pos');
            if (pos) {
                const x = Number(pos.getAttribute('x'));
                const y = Number(pos.getAttribute('y'));
                if (Number.isFinite(x)) absoluteX = x;
                if (Number.isFinite(y)) absoluteY = y;
            }
        }

        const pic = anchor.getElementsByTagNameNS(NS.xdr, 'pic').item(0);
        if (pic) {
            const blip = pic.getElementsByTagNameNS(NS.a, 'blip').item(0);
            const embed = blip?.getAttributeNS(NS.rel, 'embed');
            const rel = embed ? rels.get(embed) : undefined;
            if (rel) {
                const mediaPath = rel.target.startsWith('/')
                    ? rel.target.slice(1)
                    : normaliseRelPath(`${drawingDir}/${rel.target}`);
                const dataUrl = media[mediaPath];
                if (dataUrl) {
                    // Alt text lives under pic/nvPicPr/cNvPr.
                    const cNvPr = pic.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                    const alt = cNvPr?.getAttribute('descr') ?? cNvPr?.getAttribute('title') ?? null;
                    // Decorative flag. The accessibility marker lives inside
                    // <xdr:cNvPr><a:extLst><a:ext uri="…"><… decorative="1"/></a:ext></a:extLst>.
                    // Different producers use different URIs (Office 365 uses
                    // {C183D7F6-B498-43B3-948B-1728B52AA6E4} with an
                    // <adec:decorative val="1"/> child). Rather than match a
                    // fixed URI, scan every descendant of the cNvPr's extLst
                    // for any element carrying a `decorative="1"` (or
                    // `val="1"` on an element named "decorative") attribute.
                    const decorative = cNvPr ? detectDecorative(cNvPr) : false;
                    images.push({
                        col: col ?? 0,
                        row: row ?? 0,
                        endCol,
                        endRow,
                        colOff: colOff ?? 0,
                        rowOff: rowOff ?? 0,
                        dataUrl,
                        widthEmu,
                        heightEmu,
                        alt,
                        anchorMode: mode,
                        absoluteX,
                        absoluteY,
                        decorative,
                    });
                }
            }
        }

        // Chart detection. A chart-bearing graphicFrame wraps either a
        // <c:chart r:id="…"/> (classic) or a <cx:chart r:id="…"/> (chartEx).
        // Modern Excel wraps the whole graphicFrame in mc:AlternateContent so
        // old consumers fall back to a <sp>. We use getElementsByTagNameNS on
        // the anchor so either shape is discovered.
        const chartEl = anchor.getElementsByTagNameNS(NS.c, 'chart').item(0)
            ?? anchor.getElementsByTagNameNS(NS.cx, 'chart').item(0);
        if (chartEl) {
            const kind: 'classic' | 'chartex' = chartEl.namespaceURI === NS.cx ? 'chartex' : 'classic';
            const rId = chartEl.getAttributeNS(NS.rel, 'id');
            let chartType: string | null = null;
            let chartXml: string | undefined;
            if (rId) {
                const rel = rels.get(rId);
                if (rel) {
                    const chartPath = rel.target.startsWith('/')
                        ? rel.target.slice(1)
                        : normaliseRelPath(`${drawingDir}/${rel.target}`);
                    chartXml = parts[chartPath];
                    if (chartXml) chartType = peekChartType(chartXml, kind);
                }
            }
            // Parse both classic and chartEx chart parts into a ChartModel.
            // Classic chartSpace produces a kind of 'bar' / 'column' / …;
            // chartEx produces kind 'chartex' with a chartExSubtype tag. The
            // renderer projects classic models to inline SVG, and routes
            // chartEx models through the dashed placeholder (per-subtype
            // renderers are a future wave). Defensive: any parse error
            // leaves model=null rather than aborting the sheet.
            let model: ChartModel | null = null;
            if (chartXml) {
                try {
                    const parsed = parseChartModel(chartXml, sharedStrings);
                    if (parsed.kind !== 'unknown') model = parsed;
                } catch {
                    model = null;
                }
            }
            charts.push({
                kind,
                chartType,
                col: col ?? 0,
                row: row ?? 0,
                endCol,
                endRow,
                model,
            });
        }

        // SmartArt detection. A SmartArt diagram surfaces as a
        // <xdr:graphicFrame> whose <a:graphicData uri="…/diagram"> wraps a
        // <dgm:relIds r:dm="…" r:lo="…" r:qs="…" r:cs="…"/>. The four rIds
        // resolve against the drawing's rels into xl/diagrams/data1.xml
        // (data model), layout1.xml, quickStyle1.xml, and colors1.xml.
        // We care about data1.xml (hierarchy / text) + layout1.xml
        // (layoutDef @uniqueId); the other two describe styling we don't
        // render.
        //
        // The outer graphicFrame already contributes to `shapes[]` via
        // <xdr:graphicFrame> shape detection elsewhere; SmartArt surfaces
        // the parsed data model on a dedicated sheet-level array so a
        // consumer can walk the tree without having to re-parse the
        // drawing xml.
        const graphicFrame = anchor.getElementsByTagNameNS(NS.xdr, 'graphicFrame').item(0);
        if (graphicFrame) {
            const graphicData = graphicFrame.getElementsByTagNameNS(NS.a, 'graphicData').item(0);
            const uri = graphicData?.getAttribute('uri') ?? '';
            if (uri === SMARTART_GRAPHIC_URI) {
                const relIdsEl = graphicData!.getElementsByTagNameNS(NS.dgm, 'relIds').item(0);
                if (relIdsEl) {
                    const dmId = relIdsEl.getAttributeNS(NS.rel, 'dm');
                    const loId = relIdsEl.getAttributeNS(NS.rel, 'lo');
                    const dmRel = dmId ? rels.get(dmId) : undefined;
                    const loRel = loId ? rels.get(loId) : undefined;
                    let dataXml: string | null = null;
                    let layoutXml: string | null = null;
                    if (dmRel) {
                        const dataPath = dmRel.target.startsWith('/')
                            ? dmRel.target.slice(1)
                            : normaliseRelPath(`${drawingDir}/${dmRel.target}`);
                        dataXml = parts[dataPath] ?? null;
                    }
                    if (loRel) {
                        const layoutPath = loRel.target.startsWith('/')
                            ? loRel.target.slice(1)
                            : normaliseRelPath(`${drawingDir}/${loRel.target}`);
                        layoutXml = parts[layoutPath] ?? null;
                    }
                    // cNvPr sits under graphicFrame/nvGraphicFramePr/cNvPr.
                    // Match via the subtree so wrapping layers don't break us.
                    const cNvPr = graphicFrame.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                    const name = cNvPr?.getAttribute('name') || null;
                    const model = dataXml ? parseSmartArt(dataXml, layoutXml) : null;
                    smartArt.push({
                        name,
                        col: col ?? 0,
                        row: row ?? 0,
                        endCol,
                        endRow,
                        model,
                    });
                }
            }
        }

        // Shapes (xdr:sp) + connectors (xdr:cxnSp). Form controls wrap their
        // shape in <mc:AlternateContent>; getElementsByTagNameNS still finds
        // the inner <xdr:sp> so the form control surfaces with the same
        // preset/name/alt tuple as a plain shape.
        //
        // We deliberately scope the search to direct-ish descendants of the
        // anchor rather than the whole document, and skip shapes whose
        // ancestry chain includes a <xdr:grpSp> so a group's container
        // shape doesn't duplicate with its children. (Top-level groups are
        // still surfaced via their child <xdr:sp> entries.)
        const spEls = [
            ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'sp')),
            ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'cxnSp')),
        ];
        // Form-control detection. A form control lives inside an
        // <mc:AlternateContent> pairing <mc:Choice> (the modern
        // drawingML shape) with <mc:Fallback> (legacy VML). The anchor
        // always carries an <xdr:clientData/> sibling, but so do plain
        // shapes — that marker alone doesn't identify a form control.
        // The positive signal we rely on is a rel of type `/ctrlProp`
        // pointing at ctrlProps/ctrlProp{N}.xml in the drawing's rels.
        // When that's present, we emit a SheetFormControl and suppress
        // the generic SheetShape entry. Everything else falls through
        // to the existing shape parsing.
        const ctrlPropRelHit = anchorHasCtrlPropRel(anchor, rels, drawingDir, parts);
        let formControlConsumed = false;
        if (ctrlPropRelHit && spEls.length > 0) {
            const fc = parseFormControl(
                spEls[0],
                rels,
                drawingDir,
                parts,
                col, row, colOff, rowOff, endCol, endRow,
                anchor,
            );
            if (fc) {
                formControls.push(fc);
                formControlConsumed = true;
            }
        }
        if (!formControlConsumed) {
            for (const el of spEls) {
                const kind: 'shape' | 'connector' = el.localName === 'cxnSp' ? 'connector' : 'shape';
                const shape = parseShape(el, kind, col, row, endCol, endRow);
                if (shape) shapes.push(shape);
            }
        }
    }
    return { images, charts, shapes, formControls, smartArt };
}

// Scan an anchor's descendants for any `r:id` attribute resolving to a
// ctrlProp xml part. Returns true when we can bind at least one rId →
// rels entry → `xl/ctrlProps/ctrlProp*.xml`. Plain shapes may declare
// <xdr:clientData/> without a ctrlProp rel; this check prevents those
// from getting mis-categorised as form controls.
function anchorHasCtrlPropRel(
    anchor: Element,
    rels: Map<string, { target: string; type: string }>,
    drawingDir: string,
    parts: Record<string, string>,
): boolean {
    const descendants = anchor.getElementsByTagName('*');
    for (let i = 0; i < descendants.length; i++) {
        const rId = descendants[i].getAttributeNS(NS.rel, 'id');
        if (!rId) continue;
        const rel = rels.get(rId);
        if (!rel) continue;
        // Accept either an explicit rel type ending in `/ctrlProp` or a
        // target pointing at xl/ctrlProps/* (handles producers that emit
        // a generic type URI).
        if (rel.type.endsWith('/ctrlProp')) {
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${drawingDir}/${rel.target}`);
            if (parts[target]) return true;
        }
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${drawingDir}/${rel.target}`);
        if (/^xl\/ctrlProps\/.*\.xml$/i.test(target) && parts[target]) return true;
    }
    return false;
}

// Form-control kinds — the strict allowlist. Any ctrlProp `@objectType`
// not in this map decays to 'unknown' (keyed maps on attacker-controlled
// XLSX strings always validate against a strict allowlist). The map keys
// are the canonical ObjectType values Excel writes; callers are also
// tolerant of the lowercase variant by comparing after toLowerCase().
const FORM_CONTROL_OBJECT_TYPES: Record<string, SheetFormControl['kind']> = {
    'button':    'button',
    'checkbox':  'checkbox',
    'radio':     'radio',
    'drop':      'combo',      // Excel's "Drop-Down" (form control) = combo
    'combobox':  'combo',
    'list':      'list',
    'listbox':   'list',
    'scroll':    'scrollbar',
    'scrollbar': 'scrollbar',
    'spin':      'spinner',
    'spinner':   'spinner',
    'groupbox':  'groupBox',
    'label':     'label',
    'dialog':    'dialog',
};

function normaliseFormControlKind(raw: string | null | undefined): SheetFormControl['kind'] {
    if (!raw) return 'unknown';
    const key = raw.toLowerCase();
    return FORM_CONTROL_OBJECT_TYPES[key] ?? 'unknown';
}

// Resolve one form-control aside from its hosting <xdr:sp> + the matching
// ctrlProp xml. Returns null if the surface doesn't look like a form
// control (no rel, no recognisable ctrlProp root, etc.) so the caller
// can fall back to the generic SheetShape path.
function parseFormControl(
    sp: Element,
    rels: Map<string, { target: string; type: string }>,
    drawingDir: string,
    parts: Record<string, string>,
    col: number | null,
    row: number | null,
    colOff: number | null,
    rowOff: number | null,
    endCol: number | null,
    endRow: number | null,
    anchor: Element,
): SheetFormControl | null {
    // Resolve the ctrlProp rel. Real Excel puts the rId on the clientData's
    // @r:id attribute; older producers sometimes put it on the anchor. We
    // accept either so hand-built fixtures don't need to hunt for the exact
    // ECMA placement.
    //
    // clientData in production files often carries the rel via a
    // `fmlaLink=` attribute but NOT an r:id — the actual ctrlProp rel lives
    // on the outer <mc:AlternateContent>'s <mc:Choice><xdr:sp> or on the
    // anchor itself. We walk every descendant looking for a relationship
    // r:id attribute, then check whether the target path matches a
    // ctrlProps/ctrlProp*.xml file.
    let ctrlPropXml: string | null = null;
    const descendants = anchor.getElementsByTagName('*');
    for (let i = 0; i < descendants.length; i++) {
        const rId = descendants[i].getAttributeNS(NS.rel, 'id');
        if (!rId) continue;
        const rel = rels.get(rId);
        if (!rel) continue;
        const target = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${drawingDir}/${rel.target}`);
        if (!/^xl\/ctrlProps\/.*\.xml$/i.test(target)) continue;
        ctrlPropXml = parts[target] ?? null;
        if (ctrlPropXml) break;
    }

    // cNvPr / name / alt — same lookup as parseShape.
    const cNvPr = sp.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
    const shapeName = cNvPr?.getAttribute('name') || null;
    const altText = cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('title') || null;

    // Shape text body (label for buttons).
    let labelFromText: string | null = null;
    const txBody = sp.getElementsByTagNameNS(NS.xdr, 'txBody').item(0);
    if (txBody) {
        const paragraphs: string[] = [];
        const pEls = txBody.getElementsByTagNameNS(NS.a, 'p');
        for (let i = 0; i < pEls.length; i++) {
            const p = pEls[i];
            const tEls = p.getElementsByTagNameNS(NS.a, 't');
            let line = '';
            for (let j = 0; j < tEls.length; j++) line += tEls[j].textContent ?? '';
            paragraphs.push(line);
        }
        const joined = paragraphs.join('\n');
        if (joined.length > 0) labelFromText = joined;
    }

    // Parse the ctrlProp xml. Accept any root element whose localName is
    // `formControlPr` (the usual name); the attributes we care about all
    // live on that element. Missing ctrlProp still yields a SheetFormControl
    // so the detection path is useful — the shape name (from cNvPr) is
    // frequently "Check Box 1" etc. which we can coarsely classify.
    let kind: SheetFormControl['kind'] = 'unknown';
    let linkedCell: string | null = null;
    let inputRange: string | null = null;
    let checked: boolean | null = null;
    let min: number | null = null;
    let max: number | null = null;
    let inc: number | null = null;
    let page: number | null = null;
    let val: number | null = null;
    let dropLines: number | null = null;

    if (ctrlPropXml) {
        const ctrlDoc = parseXml(ctrlPropXml);
        // Walk root first; fall back to any descendant named formControlPr
        // if the root is an <mc:AlternateContent> wrapper.
        let root: Element | null = ctrlDoc.documentElement;
        if (root && root.localName !== 'formControlPr') {
            const found = ctrlDoc.getElementsByTagName('formControlPr').item(0);
            if (found) root = found;
        }
        if (root) {
            kind = normaliseFormControlKind(root.getAttribute('objectType'));
            linkedCell = root.getAttribute('fmlaLink') || null;
            inputRange = root.getAttribute('fmlaRange') || null;
            const checkedAttr = root.getAttribute('checked');
            if (checkedAttr === 'Checked' || checkedAttr === '1' || checkedAttr === 'true') {
                checked = true;
            } else if (checkedAttr === 'Unchecked' || checkedAttr === '0' || checkedAttr === 'false') {
                checked = false;
            } else if (kind === 'checkbox' || kind === 'radio') {
                // Absence == unchecked for the boolean controls.
                checked = false;
            }
            const numAttr = (name: string): number | null => {
                const raw = root!.getAttribute(name);
                if (raw === null) return null;
                const n = Number(raw);
                return Number.isFinite(n) ? n : null;
            };
            min = numAttr('min');
            max = numAttr('max');
            inc = numAttr('inc');
            page = numAttr('page');
            val = numAttr('val');
            dropLines = numAttr('dropLines');
        }
    } else if (shapeName) {
        // No ctrlProp rel — best-effort classification from the shape name.
        // Helps when a minimal producer emits only the hosting <xdr:sp> with
        // a conventional name (e.g. "Check Box 1").
        const lower = shapeName.toLowerCase();
        if (lower.startsWith('check box')) kind = 'checkbox';
        else if (lower.startsWith('option button')) kind = 'radio';
        else if (lower.startsWith('button')) kind = 'button';
        else if (lower.startsWith('scroll bar')) kind = 'scrollbar';
        else if (lower.startsWith('spinner') || lower.startsWith('spin button')) kind = 'spinner';
        else if (lower.startsWith('drop down')) kind = 'combo';
        else if (lower.startsWith('list box')) kind = 'list';
        else if (lower.startsWith('group box')) kind = 'groupBox';
        else if (lower.startsWith('label')) kind = 'label';
    }

    // Label: button/label prefer the shape's text body; other kinds fall
    // back to the shape name (which Excel uses for accessibility name).
    let label: string | null = null;
    if (kind === 'button' || kind === 'label' || kind === 'groupBox') {
        label = labelFromText ?? shapeName;
    } else {
        label = labelFromText ?? shapeName;
    }

    return {
        kind,
        col: col ?? 0,
        row: row ?? 0,
        colOff: colOff ?? 0,
        rowOff: rowOff ?? 0,
        endCol,
        endRow,
        endColOff: resolveEndOffset(anchor, 'colOff'),
        endRowOff: resolveEndOffset(anchor, 'rowOff'),
        label,
        linkedCell,
        inputRange,
        checked,
        min, max, inc, page, val,
        dropLines,
        altText,
    };
}

// Fish the <xdr:to>'s colOff / rowOff out of an anchor, if present.
// twoCellAnchor declares this; one- / absoluteAnchor do not, in which case
// we return null so the renderer can apply default sizing.
function resolveEndOffset(anchor: Element, attr: 'colOff' | 'rowOff'): number | null {
    const to = findDirectChildNS(anchor, NS.xdr, 'to');
    if (!to) return null;
    const el = findDirectChildNS(to, NS.xdr, attr);
    if (!el) return null;
    const n = Number(el.textContent);
    return Number.isFinite(n) ? n : null;
}

// Build a SheetShape from a single <xdr:sp> or <xdr:cxnSp>. The anchor
// coordinates are passed through from the surrounding <xdr:twoCellAnchor>
// / <xdr:oneCellAnchor>; the shape only contributes preset, text, name,
// and alt text.
function parseShape(
    el: Element,
    kind: 'shape' | 'connector',
    col: number | null,
    row: number | null,
    endCol: number | null,
    endRow: number | null,
): SheetShape | null {
    // cNvPr is normally at sp/nvSpPr/cNvPr (or cxnSp/nvCxnSpPr/cNvPr); we
    // search the subtree so either shape comes out the same.
    const cNvPr = el.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
    const name = cNvPr?.getAttribute('name') || null;
    const alt = cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('title') || null;

    // Preset geometry sits under spPr/prstGeom. <a:custGeom> (custom
    // freeform) has no preset — we leave preset=null there.
    const spPr = el.getElementsByTagNameNS(NS.xdr, 'spPr').item(0);
    let preset: string | null = null;
    if (spPr) {
        const prst = spPr.getElementsByTagNameNS(NS.a, 'prstGeom').item(0);
        preset = prst?.getAttribute('prst') || null;
    }

    // Flatten <xdr:txBody>. Each <a:p> is a paragraph; paragraphs join
    // with '\n'. Inside a paragraph, every <a:t> descendant joins without
    // a separator so runs within a paragraph read as one string.
    let text: string | null = null;
    const txBody = el.getElementsByTagNameNS(NS.xdr, 'txBody').item(0);
    if (txBody) {
        const paragraphs: string[] = [];
        const pEls = txBody.getElementsByTagNameNS(NS.a, 'p');
        for (let i = 0; i < pEls.length; i++) {
            const p = pEls[i];
            const tEls = p.getElementsByTagNameNS(NS.a, 't');
            let line = '';
            for (let j = 0; j < tEls.length; j++) line += tEls[j].textContent ?? '';
            paragraphs.push(line);
        }
        const joined = paragraphs.join('\n');
        if (joined.length > 0) text = joined;
    }

    return {
        kind,
        name,
        alt,
        preset,
        text,
        col: col ?? 0,
        row: row ?? 0,
        endCol,
        endRow,
    };
}

// Peek at a chart part and pull a chart-type name. For classic charts this is
// the first child element tag under c:chartSpace/c:chart/c:plotArea (e.g.
// "barChart", "lineChart"). For chartEx it's the series' layoutId attribute
// (e.g. "treemap", "sunburst", "waterfall"); when that's missing we fall
// through to the first recognisable child of cx:chart.
function peekChartType(xml: string, kind: 'classic' | 'chartex'): string | null {
    const doc = parseXml(xml);
    if (kind === 'classic') {
        const plotArea = doc.getElementsByTagNameNS(NS.c, 'plotArea').item(0);
        if (!plotArea) return null;
        for (let i = 0; i < plotArea.childNodes.length; i++) {
            const node = plotArea.childNodes[i];
            if (node.nodeType !== 1) continue;
            const el = node as Element;
            if (el.namespaceURI !== NS.c) continue;
            // Axis children aren't chart-type entries — skip them.
            if (/Ax$/.test(el.localName) || el.localName === 'numFmt') continue;
            if (/Chart$/.test(el.localName) || el.localName === 'chartEx') return el.localName;
        }
        return null;
    }
    // chartEx
    const series = doc.getElementsByTagNameNS(NS.cx, 'series').item(0);
    const layoutId = series?.getAttribute('layoutId');
    if (layoutId) return layoutId;
    const chart = doc.getElementsByTagNameNS(NS.cx, 'chart').item(0);
    if (!chart) return null;
    for (let i = 0; i < chart.childNodes.length; i++) {
        const node = chart.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI !== NS.cx) continue;
        if (el.localName !== 'plotArea' && el.localName !== 'title') return el.localName;
    }
    return null;
}

function anchorCellValue(parent: Element | null, tag: string): number | null {
    if (!parent) return null;
    const el = parent.getElementsByTagNameNS(NS.xdr, tag).item(0);
    if (!el) return null;
    const n = Number(el.textContent);
    return Number.isFinite(n) ? n : null;
}

// Shallow lookup: only walk the direct element children of `parent`, so a
// same-named descendant in a different subtree isn't accidentally returned.
// Useful for the anchor-level <xdr:ext>/<xdr:pos> where a tree-wide
// getElementsByTagNameNS would also pick up a nested a-namespace <ext>
// that leaked into the xdr namespace (we've seen producers mis-declare
// namespaces).
function findDirectChildNS(parent: Element, ns: string, localName: string): Element | null {
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI === ns && el.localName === localName) return el;
    }
    return null;
}

// Accessibility: DrawingML's "decorative" flag lives inside a cNvPr extLst.
// Office uses {C183D7F6-B498-43B3-948B-1728B52AA6E4} with an
// <adec:decorative val="1"/> child; other producers attach the attribute
// under different URIs. Rather than pin to a specific extension URI we
// accept any descendant that either:
//   - has a localName of "decorative" with val/value="1" (or "true"), OR
//   - carries a `decorative="1"`/`decorative="true"` attribute directly.
function detectDecorative(cNvPr: Element): boolean {
    const extLst = cNvPr.getElementsByTagNameNS(NS.a, 'extLst').item(0);
    if (!extLst) return false;
    const truthy = (v: string | null) => v === '1' || v === 'true';
    const all = extLst.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName === 'decorative' && (truthy(el.getAttribute('val')) || truthy(el.getAttribute('value')))) {
            return true;
        }
        if (truthy(el.getAttribute('decorative'))) return true;
    }
    return false;
}

function parseTable(xml: string): TableDef | null {
    const doc = parseXml(xml);
    const t = doc.getElementsByTagNameNS(NS.main, 'table').item(0);
    if (!t) return null;
    const ref = t.getAttribute('ref');
    if (!ref) return null;
    const parts = ref.split(':');
    const a = parseCellRef(parts[0]);
    const b = parts[1] ? parseCellRef(parts[1]) : a;
    if (!a || !b) return null;
    const columns: { name: string }[] = [];
    const colEls = t.getElementsByTagNameNS(NS.main, 'tableColumn');
    for (let i = 0; i < colEls.length; i++) {
        columns.push({ name: colEls[i].getAttribute('name') ?? `Column${i + 1}` });
    }
    return {
        name: t.getAttribute('name') ?? '',
        displayName: t.getAttribute('displayName') ?? '',
        col: Math.min(a.col, b.col),
        row: Math.min(a.row, b.row),
        endCol: Math.max(a.col, b.col),
        endRow: Math.max(a.row, b.row),
        headerRowCount: Number(t.getAttribute('headerRowCount') ?? '1') || 0,
        totalsRowCount: Number(t.getAttribute('totalsRowCount') ?? '0') || 0,
        columns,
        altText: t.getAttribute('altText') || null,
        altTextSummary: t.getAttribute('altTextSummary') || null,
    };
}

function parseXml(xml: string): Document {
    return new DOMParser().parseFromString(xml, 'application/xml');
}

function parseSheetList(workbookXml: string): { name: string; rId: string | null; state: 'visible' | 'hidden' | 'veryHidden' }[] {
    const doc = parseXml(workbookXml);
    const nodes = doc.getElementsByTagNameNS(NS.main, 'sheet');
    const out: { name: string; rId: string | null; state: 'visible' | 'hidden' | 'veryHidden' }[] = [];
    for (let i = 0; i < nodes.length; i++) {
        const rId = nodes[i].getAttributeNS(NS.rel, 'id');
        // <sheet state="…"> is one of visible / hidden / veryHidden. Absent
        // attribute defaults to visible; any unrecognised value falls back to
        // visible so we don't silently drop sheets on malformed producers.
        const stateAttr = nodes[i].getAttribute('state');
        const state: 'visible' | 'hidden' | 'veryHidden' =
            stateAttr === 'hidden' || stateAttr === 'veryHidden' ? stateAttr : 'visible';
        out.push({
            name: nodes[i].getAttribute('name') ?? `Sheet${i + 1}`,
            rId: rId || null,
            state,
        });
    }
    return out;
}

function parseSharedStrings(xml: string): SharedString[] {
    const doc = parseXml(xml);
    const sis = doc.getElementsByTagNameNS(NS.main, 'si');
    const out: SharedString[] = [];
    for (let i = 0; i < sis.length; i++) {
        out.push(parseSi(sis[i]));
    }
    return out;
}

// A <si> is one of:
//   - <t>plain text</t>
//   - <r><rPr>…</rPr><t>run text</t></r>… (optionally many runs, with or
//     without a terminating bare <t>)
//   - any combination of the above + <rPh sb="..." eb="..."><t>…</t></rPh>
//     phonetic-ruby annotations
// We walk direct children in order so a mix produces the correct runs; rPh
// children are collected separately and resolved against the final text
// after the main body is assembled.
function parseSi(si: Element): SharedString {
    const runs: RichTextRun[] = [];
    let sawRun = false;
    const phonetics: PhoneticRun[] = [];
    for (let i = 0; i < si.childNodes.length; i++) {
        const node = si.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI !== NS.main) continue;
        if (el.localName === 't') {
            runs.push(emptyRun(el.textContent ?? ''));
            continue;
        }
        if (el.localName === 'r') {
            sawRun = true;
            runs.push(parseRun(el));
            continue;
        }
        if (el.localName === 'rPh') {
            const parsed = parseRPh(el);
            if (parsed) phonetics.push(parsed);
        }
    }
    const text = runs.map((r) => r.text).join('');
    // Resolve each phonetic's `base` substring once the body is assembled,
    // so consumers don't need to re-slice. Clamp indices defensively so a
    // producer that emits out-of-range sb/eb can't crash the renderer.
    for (const p of phonetics) {
        const start = Math.max(0, Math.min(text.length, p.startIdx));
        const end = Math.max(start, Math.min(text.length, p.endIdx));
        p.startIdx = start;
        p.endIdx = end;
        p.base = text.slice(start, end);
    }
    return { text, runs: sawRun ? runs : null, phonetics: phonetics.length > 0 ? phonetics : null };
}

// Parse a <rPh sb="startIdx" eb="endIdx"><t>phonetic</t></rPh> into a
// PhoneticRun. `sb` and `eb` are required per the schema; missing/invalid
// numeric attributes drop the entry rather than produce a degenerate run.
// The `base` field is resolved later by parseSi once the body text is known.
function parseRPh(el: Element): PhoneticRun | null {
    const sbAttr = el.getAttribute('sb');
    const ebAttr = el.getAttribute('eb');
    if (sbAttr === null || ebAttr === null) return null;
    const sb = Number(sbAttr);
    const eb = Number(ebAttr);
    if (!Number.isFinite(sb) || !Number.isFinite(eb)) return null;
    if (sb < 0 || eb < sb) return null;
    const tEl = el.getElementsByTagNameNS(NS.main, 't').item(0);
    const phonetic = tEl?.textContent ?? '';
    return {
        base: '', // filled in by parseSi once the body text is assembled
        phonetic,
        startIdx: Math.floor(sb),
        endIdx: Math.floor(eb),
    };
}

function emptyRun(text: string): RichTextRun {
    return {
        text,
        bold: false,
        italic: false,
        underline: null,
        strike: false,
        vertAlign: null,
        size: null,
        color: null,
        name: null,
    };
}

function parseRun(el: Element): RichTextRun {
    const rPr = el.getElementsByTagNameNS(NS.main, 'rPr').item(0);
    // Multiple <t>s inside a single <r> are rare but schema-legal; concat.
    const ts = el.getElementsByTagNameNS(NS.main, 't');
    let text = '';
    for (let i = 0; i < ts.length; i++) text += ts[i].textContent ?? '';

    if (!rPr) return emptyRun(text);

    const has = (tag: string) => rPr.getElementsByTagNameNS(NS.main, tag).length > 0;
    const firstAttr = (tag: string, attr: string): string | null => {
        const node = rPr.getElementsByTagNameNS(NS.main, tag).item(0);
        return node ? node.getAttribute(attr) : null;
    };
    const color = rPr.getElementsByTagNameNS(NS.main, 'color').item(0);
    const sizeAttr = firstAttr('sz', 'val');
    const uEl = rPr.getElementsByTagNameNS(NS.main, 'u').item(0);
    let underline: UnderlineStyle = null;
    if (uEl) {
        const val = uEl.getAttribute('val');
        if (val === null || val === '') underline = 'single';
        else if (val === 'single' || val === 'double' || val === 'singleAccounting' || val === 'doubleAccounting') underline = val;
        else if (val === 'none') underline = null;
    }
    const vaEl = rPr.getElementsByTagNameNS(NS.main, 'vertAlign').item(0);
    let vertAlign: 'subscript' | 'superscript' | null = null;
    if (vaEl) {
        const val = vaEl.getAttribute('val');
        if (val === 'subscript' || val === 'superscript') vertAlign = val;
    }
    return {
        text,
        bold: has('b'),
        italic: has('i'),
        underline,
        strike: has('strike'),
        vertAlign,
        size: sizeAttr ? Number(sizeAttr) : null,
        color: parseColorElement(color),
        name: firstAttr('rFont', 'val') ?? firstAttr('name', 'val'),
    };
}

function parseSheet(
    name: string,
    state: 'visible' | 'hidden' | 'veryHidden',
    xml: string,
    sharedStrings: SharedString[],
    tables: TableDef[] = [],
    images: SheetImage[] = [],
    charts: SheetChart[] = [],
    shapes: SheetShape[] = [],
    formControls: SheetFormControl[] = [],
    embeddings: SheetEmbedding[] = [],
    pivots: SheetPivot[] = [],
    slicers: SheetSlicer[] = [],
    timelines: SheetTimeline[] = [],
    comments: SheetComment[] = [],
    threadedComments: ThreadedCommentEntry[] = [],
    hyperlinkTargets: Map<string, string> = new Map(),
    sheetIndex: number = 0,
    definedNames: DefinedName[] = [],
    metadata: WorkbookMetadata | null = null,
    smartArt: SheetSmartArt[] = [],
): Sheet {
    const doc = parseXml(xml);
    const rowEls = doc.getElementsByTagNameNS(NS.main, 'row');
    const rows: Cell[][] = [];
    const rowDimensions: RowDimension[] = [];
    let maxCol = -1;
    let maxRow = -1;

    for (let i = 0; i < rowEls.length; i++) {
        const rowEl = rowEls[i];
        const rowAttr = rowEl.getAttribute('r');
        const rowIndex = rowAttr ? Number(rowAttr) - 1 : i;
        if (!Number.isFinite(rowIndex) || rowIndex < 0) continue;

        // customHeight="1" + ht="X" means the row carries a custom height.
        // Excel also emits ht without customHeight when auto-sizing against
        // wrapped content — accept either signal.
        const htAttr = rowEl.getAttribute('ht');
        const customHeight = rowEl.getAttribute('customHeight') === '1';
        const hidden = rowEl.getAttribute('hidden') === '1';
        let height: number | null = null;
        if (htAttr && (customHeight || !Number.isNaN(Number(htAttr)))) {
            const h = Number(htAttr);
            if (Number.isFinite(h) && h >= 0) height = h;
        }
        const outlineAttr = rowEl.getAttribute('outlineLevel');
        const outlineLevel = outlineAttr != null && Number.isFinite(Number(outlineAttr))
            ? Math.max(0, Number(outlineAttr))
            : 0;
        if (height !== null || hidden || outlineLevel > 0) {
            rowDimensions.push({ row: rowIndex, height, hidden, outlineLevel });
            if (rowIndex > maxRow) maxRow = rowIndex;
        }

        const cellEls = rowEl.getElementsByTagNameNS(NS.main, 'c');
        const cells: Cell[] = [];
        for (let j = 0; j < cellEls.length; j++) {
            const cell = parseCell(cellEls[j], rowIndex, sharedStrings, metadata);
            if (!cell) continue;
            cells.push(cell);
            if (cell.col > maxCol) maxCol = cell.col;
        }
        if (cells.length === 0) continue;
        rows[rowIndex] = cells;
        if (rowIndex > maxRow) maxRow = rowIndex;
    }

    const merges = parseMerges(doc);
    // Merged ranges can extend the visible grid past any populated cell — a
    // merge that covers empty anchor cells is still visually present.
    for (const m of merges) {
        const rightCol = m.col + m.colSpan - 1;
        const bottomRow = m.row + m.rowSpan - 1;
        if (rightCol > maxCol) maxCol = rightCol;
        if (bottomRow > maxRow) maxRow = bottomRow;
    }

    const columns = parseCols(doc);
    const conditionalFormatting = parseConditionalFormatting(doc);
    const frozenPanes = parseFrozenPanes(doc);
    const autoFilter = parseAutoFilter(doc);
    const extensions = collectExtensionUris(doc);
    const view = parseSheetView(doc);
    const outline = parseSheetOutline(doc);
    const hyperlinks = parseHyperlinks(doc, hyperlinkTargets);
    const dataValidationLists = parseDataValidationLists(doc);
    const pageBreaks = parsePageBreaks(doc);
    const printArea = resolvePrintArea(sheetIndex, definedNames);
    const headerFooter = parseHeaderFooter(doc, name);
    const protection = parseSheetProtection(doc);

    // Extend maxCol/maxRow to cover hyperlink and data-validation ranges even
    // when the underlying cells are empty — the ▾ indicator / anchor still
    // needs a <td> to land on.
    for (const h of hyperlinks) {
        if (h.col > maxCol) maxCol = h.col;
        if (h.row > maxRow) maxRow = h.row;
    }
    for (const v of dataValidationLists) {
        if (v.endCol > maxCol) maxCol = v.endCol;
        if (v.endRow > maxRow) maxRow = v.endRow;
    }

    return {
        name, state, rows, maxCol, maxRow, merges, columns, rowDimensions,
        conditionalFormatting, frozenPanes, autoFilter, tables, images,
        charts, shapes, smartArt, formControls, embeddings, pivots, slicers, timelines, extensions, comments, threadedComments, view, outline,
        hyperlinks, dataValidationLists, pageBreaks, printArea, headerFooter,
        protection,
    };
}

// Parse <sheetProtection …/>. Returns null when the sheet has no such
// element. Each granular flag is `true` when the corresponding operation is
// locked — an absent attribute defaults to locked per ECMA-376 §18.3.1.85,
// so we parse as `attr !== "0"`. `enabled` mirrors the `@sheet` attribute;
// `passwordHashed` is true when any hash (`@password`, `@algorithmName +
// @hashValue`, or the extList SHA-512 variant) is present.
function parseSheetProtection(doc: Document): SheetProtection | null {
    const el = doc.getElementsByTagNameNS(NS.main, 'sheetProtection').item(0);
    if (!el) return null;
    const locked = (attr: string): boolean => el.getAttribute(attr) !== '0';
    const enabled = el.getAttribute('sheet') === '1';
    const passwordHashed =
        el.getAttribute('password') !== null ||
        el.getAttribute('hashValue') !== null;
    return {
        enabled,
        passwordHashed,
        selectLockedCells: locked('selectLockedCells'),
        selectUnlockedCells: locked('selectUnlockedCells'),
        formatCells: locked('formatCells'),
        formatColumns: locked('formatColumns'),
        formatRows: locked('formatRows'),
        insertColumns: locked('insertColumns'),
        insertRows: locked('insertRows'),
        deleteColumns: locked('deleteColumns'),
        deleteRows: locked('deleteRows'),
        sort: locked('sort'),
        autoFilter: locked('autoFilter'),
        pivotTables: locked('pivotTables'),
        objects: locked('objects'),
        scenarios: locked('scenarios'),
    };
}

// Read per-sheet display state. The <sheetView> element lives under
// <sheetViews> (first entry wins when Excel writes multiple, e.g. split
// panes); the <tabColor> lives under <sheetPr> at the sheet root. All
// defaults follow ECMA-376 §18.3.1.87 — grid lines and headers default
// to true, everything else defaults to "no value".
function parseSheetView(doc: Document): SheetView {
    const defaults: SheetView = {
        rightToLeft: false,
        showGridLines: true,
        showRowColHeaders: true,
        zoomScale: null,
        tabColor: null,
    };
    const sv = doc.getElementsByTagNameNS(NS.main, 'sheetView').item(0);
    if (sv) {
        defaults.rightToLeft = sv.getAttribute('rightToLeft') === '1';
        const sgl = sv.getAttribute('showGridLines');
        // Absent attribute → default true; explicit "0" → false.
        if (sgl === '0') defaults.showGridLines = false;
        const sh = sv.getAttribute('showRowColHeaders');
        if (sh === '0') defaults.showRowColHeaders = false;
        const zs = sv.getAttribute('zoomScale');
        if (zs !== null) {
            const n = Number(zs);
            if (Number.isFinite(n) && n > 0) defaults.zoomScale = n;
        }
    }
    // <sheetPr><tabColor .../> — scoped under the sheet-root's <sheetPr>
    // child. parseColorElement handles rgb / theme+tint / indexed uniformly.
    const sheetPr = doc.getElementsByTagNameNS(NS.main, 'sheetPr').item(0);
    if (sheetPr) {
        const tc = sheetPr.getElementsByTagNameNS(NS.main, 'tabColor').item(0);
        if (tc) defaults.tabColor = parseColorElement(tc);
    }
    return defaults;
}

// Pull sheet-wide outline config. `<sheetFormatPr outlineLevelRow outlineLevelCol/>`
// carries the max observed levels; `<sheetPr><outlinePr summaryBelow summaryRight/></sheetPr>`
// carries the summary placement flags. Both attributes default to "1" when
// absent, so the sheet's default renderer behaviour matches Excel.
function parseSheetOutline(doc: Document): SheetOutline {
    let maxRowLevel = 0;
    let maxColLevel = 0;
    const fmtPr = doc.getElementsByTagNameNS(NS.main, 'sheetFormatPr').item(0);
    if (fmtPr) {
        const r = Number(fmtPr.getAttribute('outlineLevelRow'));
        if (Number.isFinite(r) && r > 0) maxRowLevel = r;
        const c = Number(fmtPr.getAttribute('outlineLevelCol'));
        if (Number.isFinite(c) && c > 0) maxColLevel = c;
    }
    // summaryBelow / summaryRight default to true (Excel's convention). The
    // attribute uses the XSD boolean lexical form — accept "0"/"false" as
    // false, everything else including absent + "1"/"true" is true.
    let summaryBelow = true;
    let summaryRight = true;
    const sheetPr = doc.getElementsByTagNameNS(NS.main, 'sheetPr').item(0);
    if (sheetPr) {
        const outlinePr = sheetPr.getElementsByTagNameNS(NS.main, 'outlinePr').item(0);
        if (outlinePr) {
            const sb = outlinePr.getAttribute('summaryBelow');
            if (sb === '0' || sb === 'false') summaryBelow = false;
            const sr = outlinePr.getAttribute('summaryRight');
            if (sr === '0' || sr === 'false') summaryRight = false;
        }
    }
    return { maxRowLevel, maxColLevel, summaryBelow, summaryRight };
}

// Walk <hyperlinks><hyperlink …/> elements. Each carries a `ref` (single cell
// or range), optional `r:id` (→ rels target), optional `location` (intra-
// workbook anchor like "Sheet2!A1"), optional `tooltip` and `display`. We
// expand ranges to one Hyperlink per covered cell so the renderer can look up
// by (row,col); the top-left anchor is the canonical entry but every covered
// cell carries the same target/location/tooltip/display tuple.
function parseHyperlinks(doc: Document, targets: Map<string, string>): Hyperlink[] {
    const out: Hyperlink[] = [];
    const list = doc.getElementsByTagNameNS(NS.main, 'hyperlinks').item(0);
    if (!list) return out;
    const hls = list.getElementsByTagNameNS(NS.main, 'hyperlink');
    for (let i = 0; i < hls.length; i++) {
        const el = hls[i];
        const ref = el.getAttribute('ref');
        if (!ref) continue;
        const rId = el.getAttributeNS(NS.rel, 'id');
        const target = rId ? (targets.get(rId) ?? null) : null;
        const location = el.getAttribute('location');
        const tooltip = el.getAttribute('tooltip');
        const display = el.getAttribute('display');
        // ref can be "A1", "A1:C3", or (rare) a space-separated list — reuse
        // the existing sqref parser so the expansion lives in one place.
        const ranges = parseSqref(ref);
        for (const range of ranges) {
            const width = range.endCol - range.col + 1;
            const height = range.endRow - range.row + 1;
            const total = width * height;
            // DoS guard: refuse to expand pathological ranges (a crafted
            // ref="A1:XFD1048576" would otherwise push 17 billion entries).
            // Legitimate ranges fit well under this cap; oversized ones still
            // surface the top-left anchor so the URL isn't silently dropped.
            if (total > MAX_RANGE_EXPANSION_CELLS) {
                out.push({
                    col: range.col,
                    row: range.row,
                    target,
                    location: location || null,
                    tooltip: tooltip || null,
                    display: display || null,
                });
                continue;
            }
            for (let r = range.row; r <= range.endRow; r++) {
                for (let c = range.col; c <= range.endCol; c++) {
                    out.push({
                        col: c,
                        row: r,
                        target,
                        location: location || null,
                        tooltip: tooltip || null,
                        display: display || null,
                    });
                }
            }
        }
    }
    return out;
}

// Walk <dataValidations><dataValidation type="list" sqref="…"> blocks. Only
// the `type="list"` case is surfaced (other validation types — whole,
// decimal, date, time, textLength, custom — get no visual affordance). The
// sqref is expanded to one entry per contiguous range. `formula1` holds
// either a literal quoted list ("Red,Green,Blue") — in which case we split
// it out into `options` — or a range reference like `Sheet2!$A$1:$A$5`
// which leaves `options=null` (only the ▾ indicator renders, no pinned
// value set).
function parseDataValidationLists(doc: Document): DataValidationList[] {
    const out: DataValidationList[] = [];
    const list = doc.getElementsByTagNameNS(NS.main, 'dataValidations').item(0);
    if (!list) return out;
    const dvs = list.getElementsByTagNameNS(NS.main, 'dataValidation');
    for (let i = 0; i < dvs.length; i++) {
        const el = dvs[i];
        if (el.getAttribute('type') !== 'list') continue;
        const sqref = el.getAttribute('sqref');
        if (!sqref) continue;
        const f1 = el.getElementsByTagNameNS(NS.main, 'formula1').item(0);
        const formulaText = (f1?.textContent ?? '').trim();
        let options: string[] | null = null;
        const quoted = /^"(.*)"$/s.exec(formulaText);
        if (quoted) {
            // Excel uses the list separator comma; there's no documented
            // escape mechanism inside the quoted form, so a plain split is
            // correct. Empty list ("" → "" → [""]) keeps one empty string,
            // which matches Excel's behaviour of offering a single blank
            // option; consumers can filter if they want.
            options = quoted[1].split(',');
        }
        const ranges = parseSqref(sqref);
        for (const range of ranges) {
            // DoS guard: clamp pathological validation ranges before the
            // renderer expands them into a per-cell Map. Excel's own limits
            // are 1,048,576 rows × 16,384 cols; validation ranges beyond that
            // are producer bugs or deliberate bombs.
            const width = range.endCol - range.col + 1;
            const height = range.endRow - range.row + 1;
            if (width * height > MAX_RANGE_EXPANSION_CELLS) {
                const clampedCols = Math.min(width, MAX_RANGE_EXPANSION_CELLS);
                out.push({
                    col: range.col,
                    row: range.row,
                    endCol: range.col + clampedCols - 1,
                    endRow: range.row, // single row covers the anchor
                    options,
                });
                continue;
            }
            out.push({
                col: range.col,
                row: range.row,
                endCol: range.endCol,
                endRow: range.endRow,
                options,
            });
        }
    }
    return out;
}

// Walk every <ext uri="…"> in the sheet xml (sheet-level <extLst>, and any
// nested <extLst> that Excel writes inside conditional-format /
// dataValidations / etc.) and roll up by uri. xlsxjs does not render any of
// these — this is purely a triage signal.
function collectExtensionUris(doc: Document): SheetExtensionUri[] {
    // <ext> elements are emitted in multiple XSD namespaces (main, x14, etc.).
    // Rather than enumerate them we walk localName 'ext' across the tree.
    const counts = new Map<string, number>();
    const all = doc.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (el.localName !== 'ext') continue;
        const uri = el.getAttribute('uri');
        if (!uri) continue;
        counts.set(uri, (counts.get(uri) ?? 0) + 1);
    }
    const out: SheetExtensionUri[] = [];
    for (const [uri, count] of counts) out.push({ uri, count });
    return out;
}

function parseFrozenPanes(doc: Document): FrozenPanes | null {
    const pane = doc.getElementsByTagNameNS(NS.main, 'pane').item(0);
    if (!pane) return null;
    const state = pane.getAttribute('state');
    // `frozen` / `frozenSplit` are freezes; plain `split` is a scroll split
    // without freezing. Renderer treats both identically (sticky cells via
    // the existing .xlsx-frozen-* classes); `kind` preserves the distinction
    // on the model so consumers can tell one from the other.
    let kind: 'frozen' | 'split';
    if (state === 'frozen' || state === 'frozenSplit') kind = 'frozen';
    else if (state === 'split') kind = 'split';
    else return null;
    const xSplit = Number(pane.getAttribute('xSplit'));
    const ySplit = Number(pane.getAttribute('ySplit'));
    const x = Number.isFinite(xSplit) && xSplit > 0 ? xSplit : null;
    const y = Number.isFinite(ySplit) && ySplit > 0 ? ySplit : null;
    if (x === null && y === null) return null;
    return { kind, xSplit: x, ySplit: y };
}

// Parse <rowBreaks>/<colBreaks>. Each holds <brk id="N" man="1"> entries;
// `id` is the 1-based row or column after which the page break lies, so we
// subtract one to store 0-based indices. Automatic breaks (man absent or
// "0") are Excel's auto-pagination hints and skipped here — we only want
// manual breaks placed by the author.
function parsePageBreaks(doc: Document): PageBreaks {
    const rows: number[] = [];
    const cols: number[] = [];
    const rowWrap = doc.getElementsByTagNameNS(NS.main, 'rowBreaks').item(0);
    if (rowWrap) {
        const brks = rowWrap.getElementsByTagNameNS(NS.main, 'brk');
        for (let i = 0; i < brks.length; i++) {
            const el = brks[i];
            if (el.getAttribute('man') !== '1') continue;
            const id = Number(el.getAttribute('id'));
            if (!Number.isFinite(id) || id <= 0) continue;
            rows.push(id - 1);
        }
    }
    const colWrap = doc.getElementsByTagNameNS(NS.main, 'colBreaks').item(0);
    if (colWrap) {
        const brks = colWrap.getElementsByTagNameNS(NS.main, 'brk');
        for (let i = 0; i < brks.length; i++) {
            const el = brks[i];
            if (el.getAttribute('man') !== '1') continue;
            const id = Number(el.getAttribute('id'));
            if (!Number.isFinite(id) || id <= 0) continue;
            cols.push(id - 1);
        }
    }
    return { rows, cols };
}

// Parse <headerFooter><oddHeader>…</oddHeader><oddFooter>…</oddFooter>.
// The header/footer string is split into left/center/right zones by the
// `&L` / `&C` / `&R` section codes; substitution codes (&D, &T, &F, &A,
// &P, &N) are resolved eagerly using the sheet context so the returned
// zones are ready to render. `sheetName` seeds &A; there is no workbook
// file name accessible to a browser-side viewer, so &F stays literal.
function parseHeaderFooter(doc: Document, sheetName: string): HeaderFooter | null {
    const hf = doc.getElementsByTagNameNS(NS.main, 'headerFooter').item(0);
    if (!hf) return null;
    const oddHeader = hf.getElementsByTagNameNS(NS.main, 'oddHeader').item(0);
    const oddFooter = hf.getElementsByTagNameNS(NS.main, 'oddFooter').item(0);
    const parseOne = (el: Element | null): HeaderFooterZones | null => {
        if (!el) return null;
        const raw = el.textContent ?? '';
        if (raw === '') return null;
        const zones = splitHeaderFooterZones(raw);
        return {
            left: substituteHeaderFooterCodes(zones.left, sheetName),
            center: substituteHeaderFooterCodes(zones.center, sheetName),
            right: substituteHeaderFooterCodes(zones.right, sheetName),
        };
    };
    const header = parseOne(oddHeader);
    const footer = parseOne(oddFooter);
    if (!header && !footer) return null;
    return { oddHeader: header, oddFooter: footer };
}

// Split a header/footer string into its left/center/right zones. Excel
// delimits zones with `&L`, `&C`, `&R` (uppercase). A string with no zone
// markers defaults to the center zone. Zone markers can appear in any
// order and more than once; the last occurrence wins per Excel's
// behaviour. Double-ampersand (`&&`) is a literal `&` — we leave it
// in place so the substitution pass can see the escape.
function splitHeaderFooterZones(raw: string): HeaderFooterZones {
    const out: HeaderFooterZones = { left: '', center: '', right: '' };
    let zone: 'left' | 'center' | 'right' = 'center';
    let i = 0;
    while (i < raw.length) {
        if (raw[i] === '&' && i + 1 < raw.length) {
            const next = raw[i + 1];
            if (next === 'L') { zone = 'left'; i += 2; continue; }
            if (next === 'C') { zone = 'center'; i += 2; continue; }
            if (next === 'R') { zone = 'right'; i += 2; continue; }
            // Keep every other &X sequence verbatim so the substitution
            // pass can resolve it; `&&` falls into this branch too.
            out[zone] += raw[i] + raw[i + 1];
            i += 2;
            continue;
        }
        out[zone] += raw[i];
        i += 1;
    }
    return out;
}

// Resolve the substitution codes inside one header/footer zone. Codes
// supported:
//   &D — current date        →  ISO yyyy-mm-dd from `new Date()`
//   &T — current time        →  HH:MM:SS from `new Date()`
//   &F — workbook file name  →  literal "&F" (not known to the viewer)
//   &A — sheet name          →  the passed-in sheetName
//   &P — current page        →  literal "(page)"   (no pagination)
//   &N — total pages         →  literal "(total)"  (no pagination)
//   && — escaped ampersand   →  single "&"
// Every other `&X` sequence is left verbatim so consumers who want to
// interpret them themselves can.
function substituteHeaderFooterCodes(text: string, sheetName: string): string {
    let out = '';
    let i = 0;
    const now = new Date();
    const pad = (n: number) => (n < 10 ? `0${n}` : String(n));
    const dateStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const timeStr = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    while (i < text.length) {
        if (text[i] === '&' && i + 1 < text.length) {
            const next = text[i + 1];
            if (next === '&') { out += '&'; i += 2; continue; }
            if (next === 'D') { out += dateStr; i += 2; continue; }
            if (next === 'T') { out += timeStr; i += 2; continue; }
            if (next === 'F') { out += '&F'; i += 2; continue; }
            if (next === 'A') { out += sheetName; i += 2; continue; }
            if (next === 'P') { out += '(page)'; i += 2; continue; }
            if (next === 'N') { out += '(total)'; i += 2; continue; }
            // Unknown code — preserve verbatim.
            out += text[i] + text[i + 1];
            i += 2;
            continue;
        }
        out += text[i];
        i += 1;
    }
    return out;
}

// Resolve `_xlnm.Print_Area` defined names scoped to a particular sheet.
// Excel persists the print area as a defined name whose formula is one or
// more cell references (e.g. "Sheet1!$A$1:$C$5"). We parse the range part
// of each reference into cell coordinates; the sheet-name prefix is not
// validated — the defined name's localSheetId already binds it to this
// sheet. Unparseable refs are skipped rather than failing the whole list
// so a malformed producer doesn't lose the valid entries.
function resolvePrintArea(sheetIndex: number, definedNames: DefinedName[]): PrintAreaRange[] | null {
    const hits = definedNames.filter(
        (n) => n.name === '_xlnm.Print_Area' && n.localSheetId === sheetIndex,
    );
    if (hits.length === 0) return null;
    const out: PrintAreaRange[] = [];
    for (const n of hits) {
        // A print area can be a comma-separated list of refs (rare but legal).
        // Each ref is "[SheetName!]$A$1:$B$2" or "[SheetName!]$A$1".
        const parts = n.formula.split(',');
        for (const part of parts) {
            const range = parsePrintAreaRef(part.trim());
            if (range) out.push(range);
        }
    }
    return out.length > 0 ? out : null;
}

// Parse a single print-area reference like "Sheet1!$A$1:$C$5" or the
// degenerate single-cell form "Sheet1!$A$1". The sheet-name prefix (before
// a `!`, optionally single-quoted) is stripped; the absolute-ref dollar
// signs are ignored. Returns null when either endpoint fails to parse.
function parsePrintAreaRef(ref: string): PrintAreaRange | null {
    let body = ref;
    const bangIdx = body.lastIndexOf('!');
    if (bangIdx >= 0) body = body.slice(bangIdx + 1);
    const clean = body.replace(/\$/g, '');
    const [tl, br] = clean.split(':');
    const a = parseCellRef(tl);
    if (!a) return null;
    const b = br ? parseCellRef(br) : a;
    if (!b) return null;
    return {
        col: Math.min(a.col, b.col),
        row: Math.min(a.row, b.row),
        endCol: Math.max(a.col, b.col),
        endRow: Math.max(a.row, b.row),
    };
}

function parseAutoFilter(doc: Document): AutoFilter | null {
    const af = doc.getElementsByTagNameNS(NS.main, 'autoFilter').item(0);
    if (!af) return null;
    const ref = af.getAttribute('ref');
    if (!ref) return null;
    const parts = ref.split(':');
    const a = parseCellRef(parts[0]);
    const b = parts[1] ? parseCellRef(parts[1]) : a;
    if (!a || !b) return null;
    return {
        col: Math.min(a.col, b.col),
        row: Math.min(a.row, b.row),
        endCol: Math.max(a.col, b.col),
        endRow: Math.max(a.row, b.row),
    };
}

function parseMerges(doc: Document): MergedRange[] {
    const out: MergedRange[] = [];
    const els = doc.getElementsByTagNameNS(NS.main, 'mergeCell');
    for (let i = 0; i < els.length; i++) {
        const ref = els[i].getAttribute('ref');
        if (!ref) continue;
        const [tl, br] = ref.split(':');
        if (!tl || !br) continue;
        const a = parseCellRef(tl);
        const b = parseCellRef(br);
        if (!a || !b) continue;
        if (b.col < a.col || b.row < a.row) continue;
        out.push({
            col: a.col,
            row: a.row,
            colSpan: b.col - a.col + 1,
            rowSpan: b.row - a.row + 1,
        });
    }
    return out;
}

function parseCols(doc: Document): ColumnWidth[] {
    const out: ColumnWidth[] = [];
    const els = doc.getElementsByTagNameNS(NS.main, 'col');
    for (let i = 0; i < els.length; i++) {
        const el = els[i];
        const min = Number(el.getAttribute('min'));
        const max = Number(el.getAttribute('max'));
        const widthAttr = el.getAttribute('width');
        const customWidth = el.getAttribute('customWidth') === '1';
        const hidden = el.getAttribute('hidden') === '1';
        const outlineAttr = el.getAttribute('outlineLevel');
        const outlineLevel = outlineAttr != null && Number.isFinite(Number(outlineAttr))
            ? Math.max(0, Number(outlineAttr))
            : 0;
        // Excel writes a <col> for every column range even without a custom
        // width (carrying style info, hidden flag, or outline level). Keep
        // the entry whenever any of those signals is meaningful.
        const hasWidth = widthAttr !== null && (customWidth || !Number.isNaN(Number(widthAttr)));
        if (!hasWidth && !hidden && outlineLevel === 0) continue;
        if (!Number.isFinite(min) || !Number.isFinite(max)) continue;
        if (min < 1 || max < min) continue;
        const width = hasWidth && Number.isFinite(Number(widthAttr)) ? Number(widthAttr) : null;
        out.push({ min: min - 1, max: max - 1, width, hidden, outlineLevel });
    }
    return out;
}

function parseCell(
    c: Element,
    fallbackRow: number,
    sharedStrings: SharedString[],
    metadata: WorkbookMetadata | null = null,
): Cell | null {
    const ref = c.getAttribute('r');
    let col = 0, row = fallbackRow;
    if (ref) {
        const parsed = parseCellRef(ref);
        if (!parsed) return null;
        col = parsed.col;
        row = parsed.row;
    }

    const type = c.getAttribute('t') ?? 'n';
    const vEl = c.getElementsByTagNameNS(NS.main, 'v').item(0);
    const raw = vEl?.textContent ?? '';
    const fEl = c.getElementsByTagNameNS(NS.main, 'f').item(0);
    const formula = fEl ? (fEl.textContent ?? '') : null;
    const styleAttr = c.getAttribute('s');
    const styleIndex = styleAttr != null && Number.isFinite(Number(styleAttr)) ? Number(styleAttr) : -1;

    // Metadata indices (c/@cm, c/@vm) — 1-based on the XML side; store 0-based.
    const resolveMetaIdx = (attr: string, list: CellMetadataBlock[] | undefined): number | null => {
        const raw = c.getAttribute(attr);
        if (raw == null || raw === '') return null;
        const n = Number(raw);
        if (!Number.isFinite(n) || n <= 0) return null;
        const zero = Math.floor(n) - 1;
        if (!list) return zero;
        // Out-of-range indexes still populate so consumers can see the
        // producer's intent; the isSpillAnchor flag just falls through to
        // false.
        return zero;
    };
    const cellMetadataIndex = resolveMetaIdx('cm', metadata?.cellMetadata);
    const valueMetadataIndex = resolveMetaIdx('vm', metadata?.valueMetadata);
    let isSpillAnchor = false;
    if (cellMetadataIndex !== null && metadata) {
        const block = metadata.cellMetadata[cellMetadataIndex];
        if (block && block.dynamicArray) isSpillAnchor = true;
    }

    const base = {
        col, row, styleIndex, formula,
        cellMetadataIndex, valueMetadataIndex, isSpillAnchor,
    };

    switch (type) {
        case 's': {
            const idx = Number(raw);
            const entry = Number.isFinite(idx) && idx >= 0 && idx < sharedStrings.length
                ? sharedStrings[idx]
                : { text: '', runs: null, phonetics: null };
            return { ...base, value: entry.text, runs: entry.runs, phonetics: entry.phonetics, kind: 'string' };
        }
        case 'inlineStr': {
            const is = c.getElementsByTagNameNS(NS.main, 'is').item(0);
            const entry = is ? parseSi(is) : { text: '', runs: null, phonetics: null };
            return { ...base, value: entry.text, runs: entry.runs, phonetics: entry.phonetics, kind: 'inlineStr' };
        }
        case 'b':
            return { ...base, value: raw === '1' ? 'TRUE' : 'FALSE', runs: null, phonetics: null, kind: 'boolean' };
        case 'e':
            return { ...base, value: raw, runs: null, phonetics: null, kind: 'error' };
        case 'str':
            return { ...base, value: raw, runs: null, phonetics: null, kind: 'string' };
        case 'n':
        default:
            if (raw === '') return { ...base, value: '', runs: null, phonetics: null, kind: 'empty' };
            return { ...base, value: raw, runs: null, phonetics: null, kind: 'number' };
    }
}
