import { Workbook } from './workbook';
export { XlsxEncryptedError, sanitizeMediaMime, bytesToDataUrl, MAX_EMBEDDING_BYTES } from './workbook';
import { WorkbookParser } from './workbook-parser';
import { HtmlRenderer } from './html-renderer';
export { applyFormControlUpdate } from './html-renderer';
import { h } from './html';
import { evaluateSheetFormulas } from './formula-eval';
export { evaluateFormula, parseFormula, evalAst, evaluateSheetFormulas, makeSheetResolver } from './formula-eval';
export type { FormulaValue, CellResolver, SheetLike as FormulaSheet, CellLike as FormulaCell } from './formula-eval';

export type { Workbook as ParsedWorkbook, Sheet, SheetView, Cell, RichTextRun, SharedString, PhoneticRun, FrozenPanes, AutoFilter, TableDef, SheetChart, SheetPivot, SheetSlicer, SheetTimeline, SheetExtensionUri, SheetImage, SheetEmbedding, SheetComment, ThreadedCommentEntry, SheetOutline, DefinedName, ColumnWidth, RowDimension, Hyperlink, DataValidationList, PageBreaks, PrintAreaRange, HeaderFooter, HeaderFooterZones, SheetSmartArt, SmartArtModel, SmartArtNode, CfbModel, CfbStream } from './workbook-parser';
export { parseThreadedComments, isSafeHyperlinkHref, parseSmartArt, parseCfb } from './workbook-parser';
export type { Styles, CellXf, FontStyle, FillStyle, BorderStyle, Dxf, UnderlineStyle } from './styles';
export type { Theme, ColorRef } from './theme';
export type { ConditionalFormatting, CfRule, CfRuleType, CfOperator, CellRange, Cfvo, ColorScale, DataBar, IconSet } from './conditional-format';
export { parseStyles, lookupNumberFormat, sanitizeHexColor, sanitizeFontFamily, parseColorElement, resolveEffectiveXf } from './styles';
export { parseTheme, applyTint, resolveColor, indexedColor } from './theme';
export { parseConditionalFormatting, evaluateRule, resolveCfvo, interpolateColorScale } from './conditional-format';
export { formatNumber } from './number-format';
export { a1ToR1c1, r1c1ToA1 } from './formula-notation';
export { emuToPx } from './utils';
export type { ChartModel, ChartSeries } from './chart-parser';
export { parseChart } from './chart-parser';
export { renderChart } from './chart-renderer';
export { renderSmartArtSvg } from './smartart-renderer';
export type { RenderSmartArtOptions } from './smartart-renderer';

export interface Options {
    className: string;
    inWrapper: boolean;
    debug: boolean;
    // When true, formula cells render their formula text (with leading "=")
    // instead of the cached value. Useful for audit views.
    showFormulas: boolean;
    // Formula notation used for display. Has no effect unless showFormulas
    // is true. 'a1' is Excel's default; 'r1c1' is an author-relative form
    // popular in accounting contexts.
    formulaNotation: 'a1' | 'r1c1';
    // Opt-in inlining of embedded OLE / package payloads as `data:` URLs on
    // the rendered `<aside class="xlsx-embedding">`. Detection + metadata
    // surfacing happen unconditionally; only the payload projection is gated
    // on this flag, since a naive `dataUrl` lands the full embedding in both
    // model and DOM memory. Oversized payloads (>32 MiB) or MIMEs outside
    // the sanitizer's allowlist keep `dataUrl` null even when this is true.
    inlineEmbeddings: boolean;
    // Opt-in OLE Compound File Binary parsing for raw `.bin` payloads
    // attached via `/oleObject` rels. When true, each SheetEmbedding with
    // kind === 'ole' gets its `cfb` field populated with the root CLSID and
    // a capped list of streams (name + bytes). Used by downstream consumers
    // to route the payload to a Word / Excel / equation renderer. The flag
    // is orthogonal to `inlineEmbeddings` — CFB parsing runs on the raw
    // bytes xlsxjs already has in memory; no `data:` URL is projected.
    // Default off to keep Wave-8 golden output byte-stable.
    parseOleCfb: boolean;
    // Controls whether classic chartSpace charts render as inline SVG
    // (default) or fall back to the dashed placeholder. Producers that
    // want to style their own chart widget — and consumers who only need
    // the chart's anchor coordinates — can set this to false to keep the
    // placeholder-only output. Detection still happens unconditionally;
    // Sheet.charts[] is populated either way.
    renderCharts: boolean;
    // Opt-in live form-control widgets. When true, the renderer swaps the
    // detect-only <aside class="xlsx-form-control"> body for a real form
    // input (checkbox / radio / range / number / select / button) tied to
    // the control's `linkedCell`. The aside's data-attributes still surface
    // so consumers can still inspect the control's metadata. Default off to
    // keep Wave-8 golden output byte-stable.
    interactiveFormControls: boolean;
    // Opt-in interactive slicer + timeline widgets. When true, the detect-
    // only `<ul>` body of `<aside class="xlsx-slicer">` is replaced by a
    // row of `<button class="xlsx-slicer-chip">` toggle chips, and the
    // detect-only range label of `<aside class="xlsx-timeline">` picks up
    // a two-handle `<div class="xlsx-timeline-slider">` when the cache
    // supplied concrete bounds. Clicking a chip or dragging a handle fires
    // an `xlsx:slicer-change` / `xlsx:timeline-change` CustomEvent on the
    // aside — xlsxjs does NOT re-materialise the underlying pivot table;
    // consumers listen for the event to drive their own filter UI. Default
    // off to keep Wave-8 golden output byte-stable.
    interactiveSlicers: boolean;
    // Controls how `<aside class="xlsx-smartart">` paints its diagram body:
    //   · 'tree' — the Wave 9 indented `<ul>` only (default, byte-stable).
    //   · 'svg'  — an inline SVG rendered via `renderSmartArtSvg`, and no
    //              `<ul>`. Falls back to the `<ul>` when the model has no
    //              root nodes.
    //   · 'both' — SVG first, then the `<ul>` (useful for accessible fallback
    //              or for consumers who want to style the tree + SVG side by
    //              side via CSS).
    // The SVG always renders the hierarchy layout regardless of
    // `model.layout` — cycle / orgchart variants will grow native layouts in
    // a future wave; for now the hierarchy layout is a reasonable fallback.
    smartArtLayout: 'tree' | 'svg' | 'both';
    // Opt-in formula evaluation. When true, every cell with a `<f>` element
    // gets its formula parsed and evaluated post-load; cells whose cached
    // `<v>` was empty (or when `evaluateFormulasForce` is set, every formula
    // cell) are rewritten to the computed value. Only a minimal POC set of
    // functions is covered — SUM / AVERAGE / MIN / MAX / COUNT / COUNTA /
    // IF / AND / OR / NOT, plus arithmetic and cell ranges. Formulas outside
    // that surface evaluate to the string "#ERROR!" and land in the cell
    // as a kind="error" value. Default off; existing golden snapshots stay
    // byte-stable.
    evaluateFormulas: boolean;
    // When true, the evaluator overrides any existing cached `<v>` with its
    // own computed result. Only has effect alongside `evaluateFormulas`.
    // Default false — Excel's own cache is authoritative for files that
    // actually went through a saved-from-Excel pipeline.
    evaluateFormulasForce: boolean;
    h: typeof h;
}

export const defaultOptions: Options = {
    className: 'xlsx',
    inWrapper: true,
    debug: false,
    showFormulas: false,
    formulaNotation: 'a1',
    inlineEmbeddings: false,
    parseOleCfb: false,
    renderCharts: true,
    interactiveFormControls: false,
    interactiveSlicers: false,
    smartArtLayout: 'tree',
    evaluateFormulas: false,
    evaluateFormulasForce: false,
    h,
};

function mergeOptions(userOptions?: Partial<Options>): Options {
    return { ...defaultOptions, ...userOptions };
}

export async function parseAsync(data: Blob | ArrayBuffer | Uint8Array, userOptions?: Partial<Options>): Promise<Workbook> {
    const ops = mergeOptions(userOptions);
    const wb = await Workbook.load(data, new WorkbookParser(ops));
    if (ops.evaluateFormulas && wb.parsed) {
        for (const sheet of wb.parsed.sheets) {
            evaluateSheetFormulas(sheet, { force: ops.evaluateFormulasForce });
        }
    }
    return wb;
}

export async function renderWorkbook(workbook: Workbook, userOptions?: Partial<Options>): Promise<Node[]> {
    const ops = mergeOptions(userOptions);
    const renderer = new HtmlRenderer();
    if (!workbook.parsed) throw new Error('xlsx-preview: workbook is not parsed');
    return await renderer.render(workbook.parsed, ops);
}

export async function renderAsync(
    data: Blob | ArrayBuffer | Uint8Array,
    bodyContainer: HTMLElement,
    styleContainer?: HTMLElement,
    userOptions?: Partial<Options>,
): Promise<Workbook> {
    const wb = await parseAsync(data, userOptions);
    const nodes = await renderWorkbook(wb, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = '';
    bodyContainer.innerHTML = '';

    for (const n of nodes) {
        const c = n.nodeName === 'STYLE' ? styleContainer : bodyContainer;
        c.appendChild(n);
    }
    return wb;
}
