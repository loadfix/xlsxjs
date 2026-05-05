// Internal model → DOM. The rendered output is one <section class="xlsx"> per
// sheet, each containing an <h2> sheet name and a <table> of the cells.
// Numeric cells get a numeric-aligned class; other kinds render as text.

import type { Workbook, Sheet, SheetView, Cell, MergedRange, RichTextRun, FrozenPanes, SheetComment, SheetShape, SheetSlicer, SheetTimeline, SheetImage, SheetEmbedding, SheetFormControl, SheetSmartArt, SmartArtNode, RowDimension, ThreadedCommentEntry, Hyperlink, PhoneticRun } from './workbook-parser';
import { isSafeHyperlinkHref } from './workbook-parser';
import { indexToColumnLetters, emuToPx, parseCellRef } from './utils';
import { h } from './html';
import type { Options } from './xlsx-preview';
import { lookupNumberFormat, resolveEffectiveXf, sanitizeFontFamily, type Styles, type CellXf, type FontStyle, type FillStyle, type BorderStyle, type Dxf } from './styles';
import { resolveColor, type Theme } from './theme';
import { formatNumber } from './number-format';
import { a1ToR1c1 } from './formula-notation';
import { renderIcon } from './icons';
import { renderShapePreset } from './shape-presets';
import { evaluateRule, resolveCfvo, interpolateColorScale, type ConditionalFormatting, type CfContext, type CellRange, type CfRule } from './conditional-format';
import { renderChart } from './chart-renderer';
import { renderSmartArtSvg } from './smartart-renderer';

// Excel column "width" is in units of the default font's "0" character. For
// the default Calibri 11pt, one unit ≈ 7 pixels of content plus 5px of cell
// padding. The formula Excel documents is
//   pixels = truncate(((256 * width + truncate(128/MDW)) / 256) * MDW)
// with MDW=7 for Calibri 11. We approximate with 7·width + 5 which matches
// Excel's rounded pixel output within ±1px across the common width range.
// See ECMA-376 §18.3.1.13.
const PX_PER_CHAR = 7;
const PADDING_PX = 5;
function charWidthToPx(width: number): number {
    return Math.round(width * PX_PER_CHAR + PADDING_PX);
}

// Excel's default column width is 8.43 characters; the renderer resolves
// that to ~64px (8.43 * 7 + 5 ≈ 64). Rows default to 15pt (20px at 96 DPI).
// These constants feed the image-overlay offset arithmetic when a column or
// row has no declared dimension.
const DEFAULT_COL_WIDTH_CHARS = 8.43;
const DEFAULT_COL_WIDTH_PX = charWidthToPx(DEFAULT_COL_WIDTH_CHARS);
const DEFAULT_ROW_HEIGHT_PT = 15;
const DEFAULT_ROW_HEIGHT_PX = Math.round(DEFAULT_ROW_HEIGHT_PT * 4 / 3); // 20px

// Gutter column (row-number <th>) width estimate. We don't measure the
// rendered box — the image-layer sits before the table in DOM order, so
// measurement isn't available at render time. A <th> with padding 2px 6px
// + 1px borders around a 1-2 digit row number reads at roughly 30px across
// common Excel files. Callers can align tighter via consumer CSS if needed.
const GUTTER_WIDTH_PX = 30;

export class HtmlRenderer {
    async render(workbook: Workbook, options: Options): Promise<Node[]> {
        const nodes: Node[] = [];
        // The embedding CSS block is gated on workbook content: the Wave 8
        // slice deliberately keeps existing golden snapshots stable by only
        // extending the stylesheet when a workbook actually contains one or
        // more embeddings. A workbook with zero embeddings renders exactly
        // the same CSS it did before the feature landed.
        const hasEmbeddings = workbook.sheets.some((s) => s.embeddings && s.embeddings.length > 0);
        // The chart CSS block is likewise gated: it only lands in the
        // stylesheet when at least one chart will render as a real SVG
        // figure (`chart.model != null` AND `renderCharts` is enabled).
        // Workbooks with only chartEx / unsupported charts — or with
        // `renderCharts: false` — keep the placeholder-only CSS that
        // existed before this feature, so their golden snapshots stay
        // byte-stable.
        const hasRenderedCharts = options.renderCharts !== false &&
            workbook.sheets.some((s) => s.charts && s.charts.some((c) => c.model !== null));
        const hasSmartArt = workbook.sheets.some((s) => s.smartArt && s.smartArt.length > 0);
        nodes.push(renderStyle(options.className, {
            withEmbeddings: hasEmbeddings,
            withCharts: hasRenderedCharts,
            withSmartArt: hasSmartArt,
            responsive: options.responsive === true,
        }));
        for (const sheet of workbook.sheets) {
            // Skip hidden and veryHidden sheets — the demo's sheet-switcher
            // omits them too so section/index pairings stay in sync.
            if (sheet.state !== 'visible') continue;
            nodes.push(renderSheet(sheet, workbook, options));
        }
        return nodes;
    }
}

function renderStyle(className: string, opts: { withEmbeddings: boolean; withCharts?: boolean; withSmartArt?: boolean; responsive?: boolean } = { withEmbeddings: false }): HTMLStyleElement {
    const style = document.createElement('style');
    style.setAttribute('data-xlsxjs', '');
    // Kept intentionally small — consumers style further via their own CSS.
    // The embedding-related rules are appended only when the workbook carries
    // embeddings (see HtmlRenderer.render); that keeps golden snapshots for
    // embedding-free workbooks byte-stable with pre-Wave-8 output.
    const embeddingCss = opts.withEmbeddings ? `
.${className} .xlsx-embedding {
    border: 1px dashed #b0b0b0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #555; font-size: 0.9em;
    background: #fafafa;
}
.${className} .xlsx-embedding > pre {
    margin: 0; white-space: pre-wrap;
    font-family: inherit; font-size: inherit;
}
.${className} .xlsx-image-layer > .xlsx-embedding {
    position: absolute; pointer-events: auto;
    max-width: 320px;
}` : '';
    const chartCss = opts.withCharts ? `
.${className} figure.xlsx-chart {
    margin: 0.5em 0; padding: 0;
}
.${className} figure.xlsx-chart > svg {
    display: block; max-width: 100%;
}` : '';
    // Same gating pattern as embeddings: only append SmartArt rules when the
    // workbook actually carries a diagram. Workbooks without SmartArt get
    // exactly the pre-existing CSS so golden snapshots stay byte-stable.
    // Responsive (mobile-friendly) wrapping. When opted in, the sheet
    // <section> becomes its own horizontal scroll container, so a wide
    // table no longer forces the page viewport to scroll and the row
    // gutter / header row stick to the left/top of that container. We
    // also hint -webkit-overflow-scrolling so iOS gives us momentum
    // scrolling. Gated on `responsive: true` to keep byte-stable output
    // for consumers who never opt in.
    const responsiveCss = opts.responsive ? `
.${className}[data-responsive="true"] {
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    max-width: 100%;
}
.${className}[data-responsive="true"] > table {
    min-width: max-content;
}
.${className}[data-responsive="true"] th:first-child,
.${className}[data-responsive="true"] td:first-child {
    position: sticky; left: 0; z-index: 1; background: #f3f3f3;
}
.${className}[data-responsive="true"] thead th {
    position: sticky; top: 0; z-index: 2; background: #f3f3f3;
}
.${className}[data-responsive="true"] thead th:first-child {
    z-index: 3;
}
@media (max-width: 768px) {
    .${className}[data-responsive="true"] { font-size: 0.85em; }
    .${className}[data-responsive="true"] th, .${className}[data-responsive="true"] td { padding: 1px 4px; }
}` : '';
    const smartArtCss = opts.withSmartArt ? `
.${className} .xlsx-smartart {
    border: 1px dashed #b0b0b0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #333; font-size: 0.9em;
    background: #fafbfc;
}
.${className} .xlsx-smartart ul {
    list-style: none; padding-left: 1em; margin: 0;
}
.${className} .xlsx-smartart > ul { padding-left: 0; }
.${className} .xlsx-smartart li { margin: 0.1em 0; }
.${className} .xlsx-smartart .xlsx-smartart-node {
    display: inline-block; padding: 1px 4px;
}
.${className} .xlsx-image-layer > .xlsx-smartart {
    position: absolute; pointer-events: auto;
    max-width: 480px;
}` : '';
    style.textContent = `
.${className} { font-family: system-ui, sans-serif; }
.${className} table { border-collapse: separate; border-spacing: 0; }
.${className} th, .${className} td { border: 1px solid #d0d0d0; padding: 2px 6px; }
.${className} th { background: #f3f3f3; font-weight: normal; color: #555; }
.${className} td.xlsx-numeric { text-align: right; font-variant-numeric: tabular-nums; }
.${className} .xlsx-sheet-name { margin: 1rem 0 0.25rem; font-size: 1rem; font-weight: 600; }
.${className} .xlsx-frozen-col { position: sticky; left: 0; z-index: 1; background: inherit; }
.${className} .xlsx-frozen-row { position: sticky; top: 0; z-index: 2; background: inherit; }
.${className} .xlsx-frozen-both { position: sticky; left: 0; top: 0; z-index: 3; background: inherit; }
.${className} .xlsx-autofilter::after { content: " ▾"; color: #888; font-size: 0.85em; }
.${className} .xlsx-validation-list::after { content: " ▾"; color: #888; font-size: 0.85em; }
.${className} a.xlsx-hyperlink { color: #0563c1; text-decoration: underline; }
.${className} .xlsx-table-caption { font-size: 0.85em; color: #666; margin: 0.25rem 0 0; }
.${className} .xlsx-chart-placeholder {
    border: 1px dashed #999; padding: 1em; margin: 0.5em 0;
    color: #666; font-size: 0.9em; text-align: center;
}
.${className} .xlsx-shape {
    border: 1px solid #e0e0e0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #555; font-size: 0.9em;
    position: relative;
}
.${className} .xlsx-shape[data-preset] { border-color: transparent; }
.${className} .xlsx-shape[data-kind="connector"] {
    border-style: dashed; color: #888;
}
.${className} .xlsx-shape[data-kind="connector"][data-preset] { border-style: none; }
.${className} .xlsx-shape > svg {
    position: absolute; inset: 0;
    width: 100%; height: 100%;
    pointer-events: none; z-index: 0;
}
.${className} .xlsx-shape > pre {
    margin: 0; white-space: pre-line;
    font-family: inherit; font-size: inherit;
    position: relative; z-index: 1;
}${embeddingCss}${chartCss}
.${className} .xlsx-form-control {
    position: absolute;
    display: inline-flex; align-items: center; gap: 4px;
    border: 1px dashed #bbb; border-radius: 2px;
    padding: 2px 6px; margin: 0;
    color: #333; background: rgba(255, 255, 255, 0.85);
    font-size: 0.85em; pointer-events: auto;
}
.${className} .xlsx-form-control[data-kind="button"] {
    border-style: solid; background: #f3f3f3;
}
.${className} .xlsx-form-control[data-kind="groupBox"] {
    border-style: solid; background: transparent;
}
.${className} .xlsx-form-control[data-kind="scrollbar"],
.${className} .xlsx-form-control[data-kind="spinner"] {
    background: #eef2f7; color: #555;
}
.${className} .xlsx-form-control > .xlsx-form-control-glyph {
    font-size: 1em; line-height: 1; color: #333;
}
.${className} .xlsx-form-control > legend {
    font-size: 0.85em; color: #333; padding: 0 4px;
}
.${className} .xlsx-spill-anchor { outline: 1px dashed #0066cc; outline-offset: -1px; }
.${className} .xlsx-comment-marker { color: #c00; margin-left: 4px; cursor: help; }
.${className} .xlsx-threaded { color: #0066cc; margin-left: 4px; cursor: help; }
.${className} .xlsx-shrink-to-fit { font-size: clamp(0.55em, 0.95em, 1em); overflow: hidden; }
.${className}.xlsx-no-gridlines th, .${className}.xlsx-no-gridlines td { border: none; }
.${className}.xlsx-no-headers thead tr > th:first-child,
.${className}.xlsx-no-headers tbody tr > th:first-child { display: none; }
.${className}.xlsx-no-headers thead tr:first-child { display: none; }
.${className} .xlsx-outline-1 > th:first-child { padding-left: 0.5rem; }
.${className} .xlsx-outline-2 > th:first-child { padding-left: 1rem; }
.${className} .xlsx-outline-3 > th:first-child { padding-left: 1.5rem; }
.${className} .xlsx-outline-4 > th:first-child { padding-left: 2rem; }
.${className} .xlsx-outline-5 > th:first-child { padding-left: 2.5rem; }
.${className} .xlsx-outline-6 > th:first-child { padding-left: 3rem; }
.${className} .xlsx-outline-7 > th:first-child { padding-left: 3.5rem; }
.${className} col.xlsx-outline-1 { border-left: 2px solid #ddd; }
.${className} col.xlsx-outline-2 { border-left: 3px solid #ccc; }
.${className} col.xlsx-outline-3 { border-left: 4px solid #bbb; }
.${className} .xlsx-header, .${className} .xlsx-footer {
    display: grid; grid-template-columns: 1fr 1fr 1fr;
    font-size: 0.85em; color: #666; margin: 0.5em 0;
}${smartArtCss}${responsiveCss}
    `.trim();
    return style;
}

// Build a CfContext scoped to one sheet. `cellsInRange` walks the sparse
// row array and returns every intersecting cell — enough for duplicateValues
// and top10 to function correctly across a range. `date1904` is threaded
// through so the timePeriod evaluator can pick the right epoch when it
// turns a cell's serial value into a calendar date.
function makeCfContext(sheet: Sheet, date1904: boolean): CfContext {
    return {
        date1904,
        cellsInRange(range: CellRange) {
            const out: { col: number; row: number; cell: Cell | null }[] = [];
            for (let r = range.row; r <= range.endRow; r++) {
                const row = sheet.rows[r];
                for (let c = range.col; c <= range.endCol; c++) {
                    const cell = row ? row.find((x) => x.col === c) ?? null : null;
                    out.push({ col: c, row: r, cell });
                }
            }
            return out;
        },
    };
}

function rangeContains(range: CellRange, row: number, col: number): boolean {
    return row >= range.row && row <= range.endRow && col >= range.col && col <= range.endCol;
}

// Per-cell graphical state. colorScaleBg lands on td.style.backgroundColor;
// dataBar is rendered as a horizontal gradient inside the cell that fills
// proportional to dataBarFraction (0..1), in the dataBar's colour; icon
// is a serialized SVG string that gets prepended to the cell content.
interface DataBarRenderState {
    color: string;                  // effective fill colour (negative variant if applicable)
    fraction: number;               // 0..1, how far across the cell the bar fills
    gradient: boolean;              // false → flat fill to `fraction`, no transparency fade
    direction: 'leftToRight' | 'rightToLeft' | 'context';
    border: boolean;
    borderColor: string | null;     // resolved hex or null (fall back to fill colour)
    // Axis line — rendered as a short in-cell inset box-shadow when present.
    axis: { position: 'middle' | 'left'; color: string } | null;
}
interface GraphicalCfState {
    colorScaleBg?: string;                      // '#rrggbb'
    dataBar?: DataBarRenderState;
    icon?: { svg: string; showValue: boolean };
}

function resolveGraphicalConditionalFormats(
    sheet: Sheet,
    theme: Theme | null,
    date1904: boolean,
): Map<string, GraphicalCfState> {
    const out = new Map<string, GraphicalCfState>();
    const ctx = makeCfContext(sheet, date1904);

    const pairs: { range: CellRange; rule: CfRule }[] = [];
    for (const block of sheet.conditionalFormatting) {
        for (const range of block.ranges) {
            for (const rule of block.rules) {
                if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
                    pairs.push({ range, rule });
                }
            }
        }
    }
    if (pairs.length === 0) return out;
    pairs.sort((a, b) => a.rule.priority - b.rule.priority);

    for (const { range, rule } of pairs) {
        const cells = ctx.cellsInRange(range);
        const values: number[] = [];
        for (const entry of cells) {
            if (!entry.cell) continue;
            const n = Number(entry.cell.value);
            if (Number.isFinite(n)) values.push(n);
        }
        if (values.length === 0) continue;
        if (rule.type === 'colorScale' && rule.colorScale) {
            applyColorScale(out, rule.colorScale, cells, values, theme);
        } else if (rule.type === 'dataBar' && rule.dataBar) {
            applyDataBar(out, rule.dataBar, cells, values, theme);
        } else if (rule.type === 'iconSet' && rule.iconSet) {
            applyIconSet(out, rule.iconSet, cells, values);
        }
    }
    return out;
}

function applyIconSet(
    out: Map<string, GraphicalCfState>,
    icons: NonNullable<CfRule['iconSet']>,
    cells: { col: number; row: number; cell: Cell | null }[],
    values: number[],
): void {
    // Resolve each cfvo threshold; icons[i] maps to "value >= thresholds[i]".
    // Excel's convention: the first cfvo is always "min" for iconSet, so the
    // effective thresholds start at the second entry.
    const thresholds: number[] = [];
    for (const cfvo of icons.cfvos) {
        const t = resolveCfvo(cfvo, values);
        if (t === null) return; // bail: any unresolvable threshold breaks ranking
        thresholds.push(t);
    }
    if (thresholds.length < 2) return;
    for (const entry of cells) {
        if (!entry.cell) continue;
        const n = Number(entry.cell.value);
        if (!Number.isFinite(n)) continue;
        // Walk thresholds ascending: the icon index is the count of
        // thresholds the value meets-or-exceeds, minus one (so min-tier = 0).
        let idx = 0;
        for (let i = 1; i < thresholds.length; i++) {
            if (n >= thresholds[i]) idx = i;
        }
        // Custom iconSet overrides: when the rule carries a per-position
        // map, swap in the (set, iconId) at idx instead of the declared set.
        // `reverse` doesn't apply to custom overrides — the author picked
        // the glyph explicitly, so we pass reverse=false.
        const override = icons.customIcons?.[idx];
        const svg = override
            ? renderIcon(override.iconSet, override.iconId, false)
            : renderIcon(icons.iconSet, idx, icons.reverse);
        if (!svg) continue;
        const key = `${entry.row},${entry.col}`;
        const existing = out.get(key) ?? {};
        existing.icon = { svg, showValue: icons.showValue };
        out.set(key, existing);
    }
}

function applyColorScale(
    out: Map<string, GraphicalCfState>,
    scale: NonNullable<CfRule['colorScale']>,
    cells: { col: number; row: number; cell: Cell | null }[],
    values: number[],
    theme: Theme | null,
): void {
    // Build threshold / colour stops in cfvo order. Any missing colour or
    // unresolvable threshold skips the stop; if fewer than 2 remain, bail.
    const stops: { threshold: number; hex: string }[] = [];
    for (let i = 0; i < scale.cfvos.length; i++) {
        const t = resolveCfvo(scale.cfvos[i], values);
        const c = resolveColor(scale.colors[i] ?? null, theme);
        if (t !== null && c) stops.push({ threshold: t, hex: c });
    }
    if (stops.length < 2) return;
    stops.sort((a, b) => a.threshold - b.threshold);
    for (const entry of cells) {
        if (!entry.cell) continue;
        const n = Number(entry.cell.value);
        if (!Number.isFinite(n)) continue;
        const hex = interpolateColorScale(n, stops);
        if (!hex) continue;
        const key = `${entry.row},${entry.col}`;
        const existing = out.get(key) ?? {};
        existing.colorScaleBg = hex;
        out.set(key, existing);
    }
}

function applyDataBar(
    out: Map<string, GraphicalCfState>,
    bar: NonNullable<CfRule['dataBar']>,
    cells: { col: number; row: number; cell: Cell | null }[],
    values: number[],
    theme: Theme | null,
): void {
    const min = bar.cfvos[0] ? resolveCfvo(bar.cfvos[0], values) : Math.min(...values);
    const max = bar.cfvos[1] ? resolveCfvo(bar.cfvos[1], values) : Math.max(...values);
    if (min === null || max === null || max === min) return;
    const color = resolveColor(bar.color, theme);
    if (!color) return;
    // Resolve optional ext colours eagerly so each cell lookup stays cheap.
    const negFill = resolveColor(bar.negativeFillColor, theme);
    const resolvedBorderColor = resolveColor(bar.borderColor, theme);
    const axisColor = resolveColor(bar.axisColor, theme) ?? '#000000';

    const lenMin = Math.max(0, bar.minLength) / 100;
    const lenMax = Math.min(100, bar.maxLength) / 100;
    for (const entry of cells) {
        if (!entry.cell) continue;
        const n = Number(entry.cell.value);
        if (!Number.isFinite(n)) continue;
        // Map the value to [minLength, maxLength] of the cell width.
        const clamped = Math.max(min, Math.min(max, n));
        const t = (clamped - min) / (max - min);
        const fraction = lenMin + (lenMax - lenMin) * t;
        // Negative values get the negativeFillColor variant when declared;
        // otherwise reuse the main fill.
        const fillColor = n < 0 && negFill ? negFill : color;
        let axis: DataBarRenderState['axis'] = null;
        if (bar.axisPosition === 'middle') {
            axis = { position: 'middle', color: axisColor };
        } else if (bar.axisPosition === 'automatic' && min < 0 && max > 0) {
            // Automatic + mixed sign puts the axis at 0 within the range;
            // we approximate as the "middle" position (close enough for the
            // visual indicator) and leave fine-tuned placement to future work.
            axis = { position: 'middle', color: axisColor };
        } else if (bar.axisPosition === 'automatic') {
            axis = { position: 'left', color: axisColor };
        }
        const key = `${entry.row},${entry.col}`;
        const existing = out.get(key) ?? {};
        existing.dataBar = {
            color: fillColor,
            fraction,
            gradient: bar.gradient,
            direction: bar.direction,
            border: bar.border,
            borderColor: resolvedBorderColor,
            axis,
        };
        out.set(key, existing);
    }
}

// For each covered cell, return the winning dxf (lowest priority wins in
// Excel's model) after walking the sheet's conditional-format blocks. If a
// rule has stopIfTrue, no lower-priority rule can override. Keyed by
// "row,col" for O(1) lookup from the main render loop.
function resolveConditionalFormats(sheet: Sheet, styles: Styles | null, date1904: boolean): Map<string, Dxf> {
    const out = new Map<string, Dxf>();
    if (!styles?.dxfs.length) return out;
    const blocks = sheet.conditionalFormatting;
    if (!blocks.length) return out;

    const ctx = makeCfContext(sheet, date1904);

    // Flatten to (range, rule) pairs, then sort by priority ascending — Excel
    // applies the lowest-priority matching rule (ascending = higher priority).
    const pairs: { range: CellRange; rule: typeof blocks[0]['rules'][0] }[] = [];
    for (const block of blocks) {
        for (const range of block.ranges) {
            for (const rule of block.rules) pairs.push({ range, rule });
        }
    }
    pairs.sort((a, b) => a.rule.priority - b.rule.priority);

    // Track "stopIfTrue" cells so lower-priority rules can't override.
    const stopped = new Set<string>();

    for (let r = 0; r <= sheet.maxRow; r++) {
        for (let c = 0; c <= sheet.maxCol; c++) {
            const key = `${r},${c}`;
            for (const { range, rule } of pairs) {
                if (!rangeContains(range, r, c)) continue;
                if (out.has(key)) {
                    if (stopped.has(key)) break;
                    // Still consider the pair for evaluation only when the
                    // existing match's rule did not request stopIfTrue.
                    // (In Excel, every true rule applies from highest
                    // priority downward, but simpler to just keep the winner.)
                    continue;
                }
                const cell = sheet.rows[r]?.find((x) => x.col === c) ?? null;
                if (!evaluateRule(rule, cell, range, ctx)) continue;
                const dxf = styles.dxfs[rule.dxfId];
                if (dxf) out.set(key, dxf);
                if (rule.stopIfTrue) stopped.add(key);
            }
        }
    }
    return out;
}

function renderSheet(sheet: Sheet, workbook: Workbook, options: Options): HTMLElement {
    const styles = workbook.styles;
    const theme = workbook.theme;
    const date1904 = workbook.date1904;
    const section = h('section', { class: options.className, 'data-sheet-name': sheet.name }) as HTMLElement;
    // Opt-in responsive: surface a data-* hook on the section so the CSS
    // block emitted by renderStyle() can target this sheet without
    // affecting consumers who layer their own styles. See Options.responsive.
    if (options.responsive) {
        section.setAttribute('data-responsive', 'true');
    }
    applySheetView(section, sheet.view, theme);
    // Sheet-protection state surfaces as a data attribute so consumers can
    // style protected sheets (e.g. a subtle banner) via CSS. xlsxjs does not
    // enforce any protection — it only reflects the workbook's declared state.
    if (sheet.protection?.enabled) {
        section.setAttribute('data-sheet-protected', 'true');
    }
    // Manual page breaks surface as data attributes on the section; xlsxjs
    // doesn't render a visual break (no pagination) but consumers who want to
    // paint a divider can read these off.
    if (sheet.pageBreaks.rows.length > 0) {
        section.setAttribute('data-page-break-rows', sheet.pageBreaks.rows.join(','));
    }
    if (sheet.pageBreaks.cols.length > 0) {
        section.setAttribute('data-page-break-cols', sheet.pageBreaks.cols.join(','));
    }
    // Accessibility: the sheet name doubles as the table's accessible
    // name. Using aria-labelledby lets screen readers announce the sheet
    // name when the table gains focus, matching Excel's own behaviour
    // ("Sheet1, table"). The id is sheet-name-derived but sanitized
    // (sheet names can carry spaces/quotes/unicode which are invalid in
    // HTML ids under older specs); non-word chars collapse to `-`.
    const sheetNameId = `xlsx-sheet-name-${sheet.name.replace(/\W+/g, '-').replace(/^-+|-+$/g, '') || 'unnamed'}`;
    section.appendChild(h('div', { class: 'xlsx-sheet-name', id: sheetNameId }, [sheet.name]));

    const table = h('table') as HTMLTableElement;
    // Accessibility: explicit ARIA role so assistive tech treats this as
    // tabular data even in renderers that apply display:block to <table>
    // (common with frozen-pane workarounds). aria-labelledby points at
    // the sheet-name banner above.
    table.setAttribute('role', 'table');
    table.setAttribute('aria-labelledby', sheetNameId);

    if (sheet.maxCol < 0) {
        section.appendChild(table);
        return section;
    }

    const colCount = sheet.maxCol + 1;
    const rowCount = sheet.maxRow + 1;

    // aria-rowcount / aria-colcount let screen readers announce "row N
    // of M" as the user navigates, even in virtualised / frozen-pane
    // layouts where the full row count isn't reachable via DOM children
    // alone. +1 on colcount accounts for the row-number gutter column
    // xlsxjs prepends; +1 on rowcount accounts for the column-letter
    // header row.
    table.setAttribute('aria-colcount', String(colCount + 1));
    table.setAttribute('aria-rowcount', String(rowCount + 1));

    // <colgroup> — one <col> for the row-number gutter, then one per data
    // column. Widths from <cols>/<col> are applied directly in pixels so the
    // browser lays the table out the way Excel does. Hidden columns get
    // display:none so cells in that column collapse out of the layout.
    const colgroup = document.createElement('colgroup');
    colgroup.appendChild(document.createElement('col'));
    const widthByCol = new Map<number, number>();
    const hiddenCols = new Set<number>();
    const outlineByCol = new Map<number, number>();
    for (const cw of sheet.columns) {
        for (let i = cw.min; i <= cw.max; i++) {
            if (cw.width !== null) widthByCol.set(i, cw.width);
            if (cw.hidden) hiddenCols.add(i);
            if (cw.outlineLevel > 0) outlineByCol.set(i, cw.outlineLevel);
        }
    }
    for (let c = 0; c < colCount; c++) {
        const col = document.createElement('col');
        const w = widthByCol.get(c);
        if (w !== undefined) col.style.width = `${charWidthToPx(w)}px`;
        if (hiddenCols.has(c)) col.style.display = 'none';
        const lvl = outlineByCol.get(c);
        if (lvl) {
            col.setAttribute('data-outline-level', String(lvl));
            col.classList.add(`xlsx-outline-${Math.min(lvl, 7)}`);
        }
        colgroup.appendChild(col);
    }
    table.appendChild(colgroup);

    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    // The corner cell (top-left, empty) is decorative — the row-number
    // and column-letter gutters intersect here. Marking it aria-hidden
    // keeps screen readers from announcing an empty header cell.
    const corner = h('th') as HTMLTableCellElement;
    corner.setAttribute('aria-hidden', 'true');
    headRow.appendChild(corner);
    for (let c = 0; c < colCount; c++) {
        const th = h('th', null, [indexToColumnLetters(c)]) as HTMLTableCellElement;
        // scope="col" lets screen readers associate each data cell with
        // its column-letter header (A, B, C, ...) during navigation.
        th.setAttribute('scope', 'col');
        if (hiddenCols.has(c)) th.style.display = 'none';
        const lvl = outlineByCol.get(c);
        if (lvl) {
            th.setAttribute('data-outline-level', String(lvl));
            th.classList.add(`xlsx-outline-${Math.min(lvl, 7)}`);
        }
        headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);

    const frozen = sheet.frozenPanes;
    const autoFilter = sheet.autoFilter;

    // Merge index: every cell covered by a merge (except the anchor) is
    // suppressed; the anchor gets rowspan/colspan. A cell's merge is looked
    // up by "row,col".
    const mergeByAnchor = new Map<string, MergedRange>();
    const suppressedCells = new Set<string>();
    for (const m of sheet.merges) {
        mergeByAnchor.set(`${m.row},${m.col}`, m);
        for (let dr = 0; dr < m.rowSpan; dr++) {
            for (let dc = 0; dc < m.colSpan; dc++) {
                if (dr === 0 && dc === 0) continue;
                suppressedCells.add(`${m.row + dr},${m.col + dc}`);
            }
        }
    }

    const dxfByCell = resolveConditionalFormats(sheet, styles, date1904);
    const graphicalByCell = resolveGraphicalConditionalFormats(sheet, theme, date1904);

    // Comment markers indexed by "row,col" — one marker per anchored
    // comment is appended to its cell's <td> content after styling runs.
    const commentByCell = new Map<string, typeof sheet.comments[0]>();
    for (const cmt of sheet.comments) commentByCell.set(`${cmt.row},${cmt.col}`, cmt);

    // Threaded comments grouped by anchor cell. Thread entries carry
    // parent/reply structure; we keep them in document order here and let
    // the title-formatter order parents before their replies.
    const threadedByCell = new Map<string, ThreadedCommentEntry[]>();
    for (const entry of sheet.threadedComments) {
        const key = `${entry.row},${entry.col}`;
        const list = threadedByCell.get(key);
        if (list) list.push(entry);
        else threadedByCell.set(key, [entry]);
    }

    // Row dimensions by row index — applied to <tr>. Per-row height is set
    // on the <tr> (browsers honour this); hidden rows get display:none.
    const rowDim = new Map<number, typeof sheet.rowDimensions[0]>();
    for (const d of sheet.rowDimensions) rowDim.set(d.row, d);

    // Hyperlinks: already flattened to one entry per covered cell by the
    // parser, so a (row,col) lookup is enough.
    const hyperlinkByCell = new Map<string, Hyperlink>();
    for (const h of sheet.hyperlinks) hyperlinkByCell.set(`${h.row},${h.col}`, h);

    // Data-validation lists: the parser kept the range bounds + the resolved
    // options; expand to per-cell here so the render loop can check membership
    // in O(1). options is stored as a pipe-delimited string so commas inside
    // the options don't break the `data-validation-options` attribute.
    const validationByCell = new Map<string, string | null>();
    for (const v of sheet.dataValidationLists) {
        const opts = v.options ? v.options.join('|') : null;
        for (let r = v.row; r <= v.endRow; r++) {
            for (let c = v.col; c <= v.endCol; c++) {
                validationByCell.set(`${r},${c}`, opts);
            }
        }
    }

    const tbody = document.createElement('tbody');
    for (let r = 0; r < rowCount; r++) {
        const tr = document.createElement('tr');
        const dim = rowDim.get(r);
        if (dim?.hidden) tr.style.display = 'none';
        // Excel's row height is in points. One point = 4/3 px at 96 DPI.
        if (dim?.height != null) tr.style.height = `${(dim.height * 4 / 3).toFixed(2)}px`;
        if (dim && dim.outlineLevel > 0) {
            tr.setAttribute('data-outline-level', String(dim.outlineLevel));
            tr.classList.add(`xlsx-outline-${Math.min(dim.outlineLevel, 7)}`);
        }
        // scope="row" lets screen readers associate each data cell in
        // this row with the row-number gutter header so the user hears
        // "row 3" as they navigate across.
        const rowHeader = h('th', null, [String(r + 1)]) as HTMLTableCellElement;
        rowHeader.setAttribute('scope', 'row');
        tr.appendChild(rowHeader);
        const cells = sheet.rows[r];
        const byCol: Record<number, typeof cells[0]> = {};
        if (cells) for (const cell of cells) byCol[cell.col] = cell;
        for (let c = 0; c < colCount; c++) {
            if (suppressedCells.has(`${r},${c}`)) continue;
            const cell = byCol[c];
            const td = document.createElement('td');
            if (cell) renderCellContent(td, cell, styles, theme, date1904, options);
            // Dynamic-array spill anchors get a subtle dashed outline. The
            // flag is resolved at parse time from xl/metadata.xml + c/@cm.
            if (cell?.isSpillAnchor) td.classList.add('xlsx-spill-anchor');
            // Hyperlink wrap: runs first so the anchor hugs the rendered
            // content (textContent / per-run <span>s). The URL is held to the
            // allowlist in isSafeHyperlinkHref; rejected URLs leave the cell
            // rendered as inert text.
            const hlink = hyperlinkByCell.get(`${r},${c}`);
            if (hlink) wrapCellWithHyperlink(td, hlink);
            const dxf = dxfByCell.get(`${r},${c}`);
            if (dxf) applyDxf(td, dxf, theme, cell ?? null, date1904);
            const gfx = graphicalByCell.get(`${r},${c}`);
            if (gfx) applyGraphicalCf(td, gfx);
            const cmt = commentByCell.get(`${r},${c}`);
            if (cmt) appendCommentMarker(td, cmt);
            if (hiddenCols.has(c)) td.style.display = 'none';
            if (frozen) tagFrozen(td, r, c, frozen);
            if (autoFilter && r === autoFilter.row && c >= autoFilter.col && c <= autoFilter.endCol) {
                td.classList.add('xlsx-autofilter');
            }
            if (validationByCell.has(`${r},${c}`)) {
                td.classList.add('xlsx-validation-list');
                const opts = validationByCell.get(`${r},${c}`);
                if (opts !== null && opts !== undefined) {
                    td.setAttribute('data-validation-options', opts);
                }
            }
            const threaded = threadedByCell.get(`${r},${c}`);
            if (threaded && threaded.length) appendThreadedCommentMarker(td, threaded);
            const merge = mergeByAnchor.get(`${r},${c}`);
            if (merge) {
                if (merge.colSpan > 1) td.setAttribute('colspan', String(merge.colSpan));
                if (merge.rowSpan > 1) td.setAttribute('rowspan', String(merge.rowSpan));
                td.classList.add('xlsx-merged');
            }
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    section.appendChild(table);

    // Table definitions: not rendered visually (Excel table styling would
    // duplicate the cell-level fills/fonts we already applied), but we emit
    // a small caption so consumers can see the table shape in the DOM.
    for (const t of sheet.tables) {
        const caption = document.createElement('div');
        caption.className = 'xlsx-table-caption';
        caption.setAttribute('data-table-name', t.name);
        caption.setAttribute('data-table-display-name', t.displayName);
        // Accessibility: prefer the short altText on `aria-label` (assistive
        // tech reads this in place of the visible caption's tree); when only
        // the longer `altTextSummary` is present, fall back to `title` so it
        // surfaces as a tooltip + accessible description.
        if (t.altText) {
            caption.setAttribute('aria-label', t.altText);
        } else if (t.altTextSummary) {
            caption.setAttribute('title', t.altTextSummary);
        }
        caption.textContent = `Table "${t.displayName || t.name}" · ${t.columns.length} column(s) · rows ${t.row + 1}-${t.endRow + 1}`;
        section.appendChild(caption);
    }

    // Images, form-controls, and anchored embeddings all live inside a
    // zero-height <div class="xlsx-image-layer"> that sits directly above the
    // table. twoCell and oneCell anchors compute CSS top/left/width/height by
    // summing column widths + row heights up to the anchor cell, so the figure
    // lands on its anchor cell instead of trailing after the table. absoluteAnchor
    // images keep their EMU-derived pixel offsets. The anchor coordinates +
    // offsets are still surfaced as data-attributes so consumers who want a
    // different overlay strategy can read them off the DOM.
    //
    // Form controls piggyback on this layer with the same anchor math —
    // xlsxjs only surfaces detect-only metadata (kind + linkedCell /
    // checked / min / max / etc.) as data attributes on an <aside>; no
    // interactive widget is painted.
    //
    // Unanchored embeddings (producers that skip objectPr/anchor) fall through
    // to the section-level render step after the table.
    const anchoredEmbeddings = sheet.embeddings.filter((e) => e.col !== null && e.row !== null);
    const unanchoredEmbeddings = sheet.embeddings.filter((e) => e.col === null || e.row === null);
    const smartArtEntries = sheet.smartArt ?? [];
    if (sheet.images.length > 0 || sheet.formControls.length > 0 || anchoredEmbeddings.length > 0 || smartArtEntries.length > 0) {
        const imageLayer = document.createElement('div');
        imageLayer.className = 'xlsx-image-layer';
        // position:relative + height:0 so the layer establishes a positioning
        // context without pushing the table down; pointer-events:none lets
        // clicks fall through to the table except on the image figures
        // themselves (which re-enable pointer-events).
        imageLayer.style.position = 'relative';
        imageLayer.style.height = '0';
        imageLayer.style.pointerEvents = 'none';
        for (const img of sheet.images) {
            imageLayer.appendChild(renderImage(img, widthByCol, hiddenCols, rowDim));
        }
        for (const fc of sheet.formControls) {
            imageLayer.appendChild(renderFormControl(fc, widthByCol, hiddenCols, rowDim, options, sheet, workbook));
        }
        for (const emb of anchoredEmbeddings) {
            const aside = renderEmbedding(emb);
            // Anchored embeddings pick up the same sum-of-col/row geometry as
            // anchored images so they sit on their cell rather than stacking
            // at 0,0. EMU offsets aren't carried on oleObject refs — just
            // the cell index pair — so we skip colOff/rowOff here.
            const left = GUTTER_WIDTH_PX + sumColsPx(emb.col as number, widthByCol, hiddenCols);
            const top = sumRowsPx(emb.row as number, rowDim);
            aside.style.left = `${left}px`;
            aside.style.top = `${top}px`;
            imageLayer.appendChild(aside);
        }
        for (const art of smartArtEntries) {
            const aside = renderSmartArt(art, options);
            const left = GUTTER_WIDTH_PX + sumColsPx(art.col, widthByCol, hiddenCols);
            const top = sumRowsPx(art.row, rowDim);
            aside.style.left = `${left}px`;
            aside.style.top = `${top}px`;
            imageLayer.appendChild(aside);
        }
        // Insert the layer directly before the <table> so the anchor-cell
        // geometry in CSS lines up with the table's laid-out grid.
        section.insertBefore(imageLayer, table);
    }

    // Charts. When options.renderCharts is true (the default) we project
    // the parsed ChartModel to inline SVG via chart-renderer.ts. If the
    // model is null, the plot kind isn't one we render (scatter / area /
    // unknown / chartEx), or the renderer returns null, we fall back to
    // the dashed placeholder so consumers still see a chart-sized div at
    // the anchor. Anchor data-attributes match between the real SVG
    // figure and the placeholder so downstream CSS can target either.
    for (const chart of sheet.charts) {
        let svg: SVGSVGElement | null = null;
        if (options.renderCharts !== false && chart.model) {
            try {
                svg = renderChart(chart.model);
            } catch {
                svg = null;
            }
        }
        if (svg) {
            const figure = document.createElement('figure');
            figure.className = 'xlsx-chart';
            figure.setAttribute('data-chart-kind', chart.kind);
            if (chart.chartType) figure.setAttribute('data-chart-type', chart.chartType);
            if (chart.model?.kind) figure.setAttribute('data-chart-plot', chart.model.kind);
            figure.setAttribute('data-anchor-col', String(chart.col));
            figure.setAttribute('data-anchor-row', String(chart.row));
            if (chart.endCol !== null) figure.setAttribute('data-anchor-end-col', String(chart.endCol));
            if (chart.endRow !== null) figure.setAttribute('data-anchor-end-row', String(chart.endRow));
            figure.appendChild(svg);
            section.appendChild(figure);
        } else {
            const ph = document.createElement('div');
            ph.className = 'xlsx-chart-placeholder';
            ph.setAttribute('data-chart-kind', chart.kind);
            if (chart.chartType) ph.setAttribute('data-chart-type', chart.chartType);
            ph.setAttribute('data-anchor-col', String(chart.col));
            ph.setAttribute('data-anchor-row', String(chart.row));
            if (chart.endCol !== null) ph.setAttribute('data-anchor-end-col', String(chart.endCol));
            if (chart.endRow !== null) ph.setAttribute('data-anchor-end-row', String(chart.endRow));
            ph.textContent = `[chart: ${chart.chartType ?? chart.kind}]`;
            section.appendChild(ph);
        }
    }

    // Shapes and connectors. xlsxjs does NOT render the preset geometry
    // itself (a text box, ellipse, callout, connector arrow, etc.) but we
    // surface an <aside> per shape so the DOM stays in lockstep with the
    // XLSX's drawing layer. Text bodies render through a <pre> with
    // white-space:pre-line so paragraph breaks survive without collapsing
    // ordinary whitespace. All attacker-controlled strings reach the DOM
    // via textContent / setAttribute only.
    for (const shape of sheet.shapes) {
        section.appendChild(renderShape(shape));
    }

    // Slicers + timelines: detect-only placeholders. Excel's live widgets are
    // clickable filter affordances; xlsxjs emits a summary <aside> per
    // widget so the DOM reflects the slicer/timeline's presence + current
    // selection without trying to reproduce the interactive UI. Same
    // security contract as shapes — attacker-controlled strings reach the
    // DOM only via textContent / setAttribute.
    for (const slicer of sheet.slicers) {
        section.appendChild(renderSlicer(slicer, options));
    }
    for (const timeline of sheet.timelines) {
        section.appendChild(renderTimeline(timeline, options));
    }

    // Unanchored embeddings. Anchor-less <oleObject> rels (e.g. producers that
    // skip the objectPr/anchor block) can't be placed on the image overlay,
    // so they land after the table as in-flow placeholders. The fileName +
    // contentType + size + progId all reach the DOM via setAttribute /
    // textContent only so attacker-controlled strings can't escape.
    for (const emb of unanchoredEmbeddings) {
        section.appendChild(renderEmbedding(emb));
    }

    // Header / footer — rendered as 3-column grids after the table. Zone
    // strings come pre-substituted (dates / sheet name / literal markers);
    // they're attacker-controlled so they reach the DOM only via textContent.
    if (sheet.headerFooter?.oddHeader) {
        section.appendChild(renderHeaderFooter('xlsx-header', sheet.headerFooter.oddHeader));
    }
    if (sheet.headerFooter?.oddFooter) {
        section.appendChild(renderHeaderFooter('xlsx-footer', sheet.headerFooter.oddFooter));
    }
    return section;
}

// Pixel width of a single column, honouring the sheet's <col> entries when
// declared and falling back to Excel's 8.43-char default otherwise. Hidden
// columns contribute 0 so images anchored past them stack tight against the
// visible grid.
function columnPx(col: number, widthByCol: Map<number, number>, hiddenCols: Set<number>): number {
    if (hiddenCols.has(col)) return 0;
    const w = widthByCol.get(col);
    return w !== undefined ? charWidthToPx(w) : DEFAULT_COL_WIDTH_PX;
}

// Pixel height of a single row, honouring <row ht>/<row hidden> when
// declared and falling back to Excel's 15pt default otherwise. Hidden rows
// contribute 0 so images stack tight against the visible grid.
function rowPx(row: number, rowDim: Map<number, RowDimension>): number {
    const d = rowDim.get(row);
    if (d?.hidden) return 0;
    if (d?.height != null) return Math.round(d.height * 4 / 3);
    return DEFAULT_ROW_HEIGHT_PX;
}

// Sum pixel widths for columns 0..col-1, matching the semantics Excel uses
// for anchoring a drawing at column `col`.
function sumColsPx(col: number, widthByCol: Map<number, number>, hiddenCols: Set<number>): number {
    let total = 0;
    for (let c = 0; c < col; c++) total += columnPx(c, widthByCol, hiddenCols);
    return total;
}

function sumRowsPx(row: number, rowDim: Map<number, RowDimension>): number {
    let total = 0;
    for (let r = 0; r < row; r++) total += rowPx(r, rowDim);
    return total;
}

// Build a <figure class="xlsx-image"> for one parsed image. twoCell + oneCell
// anchors get CSS top/left computed from the anchor cell coordinates + any
// row-height / column-width entries on the sheet; absolute anchors keep
// their EMU-derived left/top. The figure is always position:absolute inside
// the .xlsx-image-layer so it overlays the table without displacing cells.
// Attacker-controlled strings (alt text) reach the DOM only via `img.alt` /
// setAttribute — src is an embedded data: URL built by the parser.
function renderImage(
    img: SheetImage,
    widthByCol: Map<number, number>,
    hiddenCols: Set<number>,
    rowDim: Map<number, RowDimension>,
): HTMLElement {
    const fig = document.createElement('figure');
    fig.className = 'xlsx-image';
    fig.setAttribute('data-anchor-mode', img.anchorMode);
    fig.setAttribute('data-anchor-col', String(img.col));
    fig.setAttribute('data-anchor-row', String(img.row));
    if (img.endCol !== null) fig.setAttribute('data-anchor-end-col', String(img.endCol));
    if (img.endRow !== null) fig.setAttribute('data-anchor-end-row', String(img.endRow));
    fig.style.margin = '0';
    fig.style.position = 'absolute';
    // Re-enable pointer events on the figure itself (the layer turned them
    // off so clicks fall through to the underlying table cells).
    fig.style.pointerEvents = 'auto';

    if (img.anchorMode === 'absolute') {
        // Absolute anchors carry pixel positions in EMU — no grid math.
        if (img.absoluteX !== null) fig.style.left = `${emuToPx(img.absoluteX)}px`;
        if (img.absoluteY !== null) fig.style.top = `${emuToPx(img.absoluteY)}px`;
    } else {
        // Cell-anchored: left = gutter + sum of column widths for columns
        // before the anchor + EMU colOff; top = sum of row heights for rows
        // before the anchor + EMU rowOff.
        const left = GUTTER_WIDTH_PX + sumColsPx(img.col, widthByCol, hiddenCols) + emuToPx(img.colOff);
        const top = sumRowsPx(img.row, rowDim) + emuToPx(img.rowOff);
        fig.style.left = `${left}px`;
        fig.style.top = `${top}px`;
    }

    const el = document.createElement('img');
    el.src = img.dataUrl;
    // Accessibility. A decorative image gets an explicit empty alt +
    // aria-hidden so screen readers skip it entirely; otherwise we pass
    // through the producer's descr/title when present.
    if (img.decorative) {
        el.alt = '';
        el.setAttribute('aria-hidden', 'true');
    } else if (img.alt) {
        el.alt = img.alt;
    }

    // Pick a rendered width/height: twoCellAnchor spans the range between
    // `from` and `to` (xdr:to gives the exclusive lower-right corner, so the
    // width is the gap between their summed coordinates). oneCellAnchor and
    // absoluteAnchor carry their own widthEmu/heightEmu. The SheetImage
    // model doesn't surface the end-cell fractional offsets (toColOff /
    // toRowOff) — if the producer set them non-zero our twoCell size may be
    // off by a few px; documented compromise. The width/height land on both
    // the <img> intrinsic attrs (keeps scenario 23/63 size assertions green)
    // and are implicit on the <figure>'s content box.
    let widthPx: number | null = null;
    let heightPx: number | null = null;
    if (img.anchorMode === 'twoCell' && img.endCol !== null && img.endRow !== null) {
        const spanW = sumColsPx(img.endCol, widthByCol, hiddenCols) - sumColsPx(img.col, widthByCol, hiddenCols) - emuToPx(img.colOff);
        widthPx = spanW > 0 ? spanW : (img.widthEmu ? emuToPx(img.widthEmu) : null);
        const spanH = sumRowsPx(img.endRow, rowDim) - sumRowsPx(img.row, rowDim) - emuToPx(img.rowOff);
        heightPx = spanH > 0 ? spanH : (img.heightEmu ? emuToPx(img.heightEmu) : null);
    } else if (img.widthEmu && img.heightEmu) {
        widthPx = emuToPx(img.widthEmu);
        heightPx = emuToPx(img.heightEmu);
    }
    if (widthPx !== null && heightPx !== null) {
        el.width = widthPx;
        el.height = heightPx;
    }
    el.style.maxWidth = '100%';
    fig.appendChild(el);
    return fig;
}

// Build an <aside class="xlsx-shape"> for a drawing shape or connector.
// Same anchor data-attrs as the <figure class="xlsx-image"> / chart
// placeholder so consumers that want in-flow positioning can overlay.
// Name / alt / preset / text are all attacker-controlled XLSX strings:
// they reach the DOM only via setAttribute (HTML-encoded) or textContent.
// Paragraph breaks in the text body are preserved by emitting a <pre>
// with CSS white-space: pre-line (see renderStyle).
function renderShape(shape: SheetShape): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-shape';
    aside.setAttribute('data-kind', shape.kind);
    if (shape.preset) aside.setAttribute('data-preset', shape.preset);
    if (shape.name) aside.setAttribute('data-name', shape.name);
    if (shape.alt) aside.setAttribute('aria-label', shape.alt);
    aside.setAttribute('data-anchor-col', String(shape.col));
    aside.setAttribute('data-anchor-row', String(shape.row));
    if (shape.endCol !== null) aside.setAttribute('data-anchor-end-col', String(shape.endCol));
    if (shape.endRow !== null) aside.setAttribute('data-anchor-end-row', String(shape.endRow));

    // Inline SVG glyph for the Excel prstGeom preset, when we have a
    // mapping. The SVG comes from `renderShapePreset` — it's self-generated
    // (no attacker content) so setting it via innerHTML on a throw-away
    // container is safe; we then move the <svg> child out so the parent
    // structure stays flat (<aside> > <svg>, <aside> > <pre>).
    const svg = renderShapePreset(shape.preset, { width: 100, height: 100, stroke: '#888' });
    if (svg) {
        const tmp = document.createElement('div');
        tmp.innerHTML = svg;
        const svgEl = tmp.firstElementChild;
        if (svgEl) aside.appendChild(svgEl);
    }

    if (shape.text && shape.text.length > 0) {
        const pre = document.createElement('pre');
        pre.textContent = shape.text;
        aside.appendChild(pre);
    }
    return aside;
}

// Build an `<aside class="xlsx-smartart">` carrying the diagram's hierarchy
// tree. We deliberately do NOT attempt to reproduce Excel's fancy diagram
// layout (hierarchy tree lines, cycle rings, pyramid stacks) — instead we
// emit a nested <ul> that captures the parsed parent-of graph. Consumers who
// want the real layout can hydrate against the `data-layout` attribute and
// the per-<li> `data-level` attribute to build their own geometry.
//
// All attacker-controlled strings (node text, layout name, model name) reach
// the DOM via textContent / setAttribute only. The <aside> receives CSS
// position:absolute when it lands inside the image overlay layer; its
// default in-flow style (no image layer) leaves it flowing above the table,
// matching how the slicer / timeline asides behave.
function renderSmartArt(art: SheetSmartArt, options: Options): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-smartart';
    aside.style.position = 'absolute';
    aside.style.pointerEvents = 'auto';
    if (art.model?.layout) aside.setAttribute('data-layout', art.model.layout);
    if (art.name) aside.setAttribute('data-name', art.name);
    aside.setAttribute('data-anchor-col', String(art.col));
    aside.setAttribute('data-anchor-row', String(art.row));
    if (art.endCol !== null) aside.setAttribute('data-anchor-end-col', String(art.endCol));
    if (art.endRow !== null) aside.setAttribute('data-anchor-end-row', String(art.endRow));

    const roots = art.model?.rootNodes ?? [];
    // Wire the layout strategy. 'tree' (default) keeps Wave-9 byte-stable by
    // emitting only the <ul>. 'svg' drops the <ul> in favour of a rendered
    // SVG hierarchy (falling back to the <ul> when the model is empty so the
    // aside isn't silently empty). 'both' prepends the SVG in front of the
    // <ul> so assistive tech still has the textual hierarchy.
    const layout = options.smartArtLayout ?? 'tree';
    let svg: SVGSVGElement | null = null;
    if (layout !== 'tree' && art.model) {
        try {
            svg = renderSmartArtSvg(art.model);
        } catch {
            svg = null;
        }
    }
    if (svg) aside.appendChild(svg);
    if (layout === 'tree' || layout === 'both' || !svg) {
        aside.appendChild(renderSmartArtList(roots, 0));
    }
    return aside;
}

// Recursively build a <ul> for a list of SmartArt nodes at the given depth.
// Each <li> carries a data-level attribute matching its depth (0-based) and
// a <span class="xlsx-smartart-node"> containing the node text via
// textContent. Empty-text nodes still render with a <span> so the DOM shape
// stays predictable (consumers can filter by absence of text in their
// hydration code).
function renderSmartArtList(nodes: SmartArtNode[], depth: number): HTMLUListElement {
    const ul = document.createElement('ul');
    for (const node of nodes) {
        const li = document.createElement('li');
        li.setAttribute('data-level', String(depth));
        const span = document.createElement('span');
        span.className = 'xlsx-smartart-node';
        span.textContent = node.text;
        li.appendChild(span);
        if (node.children.length > 0) {
            li.appendChild(renderSmartArtList(node.children, depth + 1));
        }
        ul.appendChild(li);
    }
    return ul;
}

// Build an `<aside class="xlsx-form-control">` for one form-control entry.
// xlsxjs deliberately does NOT paint an interactive widget — the aside only
// captures the detected metadata (kind, linked cell, input range, min/max,
// checked state) as data-attributes so consumers who want a real widget can
// hydrate against it. Label text and alt text are attacker-controlled XLSX
// strings: they reach the DOM via textContent / setAttribute only.
//
// For `checkbox` / `radio`, a small unicode glyph (☐/☑/○/●) leads the
// label — a lightweight indicator that matches how the detected state
// should render. The glyph lives in a `<span class="xlsx-form-control-glyph">`
// so consumers can restyle / hide it via CSS.
//
// For `groupBox`, the aside wraps a <legend>-carrying <fieldset>-style box
// containing the label. The DOM element is still an `<aside>` (not a real
// <fieldset>) so screen readers don't infer a form grouping we didn't
// intend — but we emit a nested <legend> child with the label text so
// visual styling can mimic a group-box chrome.
function renderFormControl(
    fc: SheetFormControl,
    widthByCol: Map<number, number>,
    hiddenCols: Set<number>,
    rowDim: Map<number, RowDimension>,
    options: Options,
    sheet: Sheet,
    workbook: Workbook,
): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-form-control';
    aside.setAttribute('data-kind', fc.kind);
    aside.setAttribute('data-anchor-col', String(fc.col));
    aside.setAttribute('data-anchor-row', String(fc.row));
    if (fc.endCol !== null) aside.setAttribute('data-anchor-end-col', String(fc.endCol));
    if (fc.endRow !== null) aside.setAttribute('data-anchor-end-row', String(fc.endRow));
    if (fc.linkedCell) aside.setAttribute('data-linked-cell', fc.linkedCell);
    if (fc.inputRange) aside.setAttribute('data-input-range', fc.inputRange);
    if (fc.checked !== null) aside.setAttribute('data-checked', String(fc.checked));
    if (fc.min !== null) aside.setAttribute('data-min', String(fc.min));
    if (fc.max !== null) aside.setAttribute('data-max', String(fc.max));
    if (fc.inc !== null) aside.setAttribute('data-inc', String(fc.inc));
    if (fc.page !== null) aside.setAttribute('data-page', String(fc.page));
    if (fc.val !== null) aside.setAttribute('data-val', String(fc.val));
    if (fc.dropLines !== null) aside.setAttribute('data-drop-lines', String(fc.dropLines));
    if (fc.altText) aside.setAttribute('aria-label', fc.altText);

    // Positioning uses the same anchor math as SheetImage (cell-anchor only;
    // form controls always use twoCellAnchor in real producers). `left` and
    // `top` resolve against the image-layer's origin (= table top-left);
    // the GUTTER_WIDTH_PX offset accounts for the row-number <th>.
    const left = GUTTER_WIDTH_PX + sumColsPx(fc.col, widthByCol, hiddenCols) + emuToPx(fc.colOff);
    const top = sumRowsPx(fc.row, rowDim) + emuToPx(fc.rowOff);
    aside.style.left = `${left}px`;
    aside.style.top = `${top}px`;

    // Interactive path: when Options.interactiveFormControls is on, swap
    // the detect-only glyph/label body for a real input. The aside still
    // carries its data-attrs so consumers who want to inspect metadata can.
    // Default-off keeps Wave-8 golden snapshots byte-stable.
    if (options.interactiveFormControls) {
        populateInteractiveFormControl(aside, fc, sheet, workbook);
        return aside;
    }

    // Glyph for the boolean controls — placed ahead of the label so the
    // aside's flex layout lines it up. Consumers who want different
    // glyphs can hide this span via CSS and inject their own.
    if (fc.kind === 'checkbox' || fc.kind === 'radio') {
        const glyph = document.createElement('span');
        glyph.className = 'xlsx-form-control-glyph';
        // The DOM text-content path is safe — we're writing a fixed
        // unicode glyph from a literal, not attacker input.
        if (fc.kind === 'checkbox') {
            glyph.textContent = fc.checked ? '☑' : '☐'; // ☑ / ☐
        } else {
            glyph.textContent = fc.checked ? '●' : '○'; // ● / ○
        }
        aside.appendChild(glyph);
    }

    // groupBox: render a visible <legend> so consumer CSS can style the
    // aside like a fieldset chrome. Label text goes inside the legend;
    // any loose label text outside is suppressed so we don't duplicate.
    if (fc.kind === 'groupBox') {
        const legend = document.createElement('legend');
        if (fc.label) legend.textContent = fc.label;
        aside.appendChild(legend);
        return aside;
    }

    // Everything else: plain label text after the optional glyph.
    if (fc.label) {
        const txt = document.createElement('span');
        txt.className = 'xlsx-form-control-label';
        txt.textContent = fc.label;
        aside.appendChild(txt);
    }
    return aside;
}

// Populate an interactive form-control aside with a real input element.
// The aside already carries all the metadata data-attrs — we just attach
// a widget + change/input listener that updates the linked cell's <td>
// via applyFormControlUpdate. Radio groups are keyed on inputRange (or a
// synthesized id from the col/row pair) so multiple radios share a name.
// All attacker-controlled strings (labels, option text) reach the DOM via
// textContent / setAttribute only.
function populateInteractiveFormControl(
    aside: HTMLElement,
    fc: SheetFormControl,
    sheet: Sheet,
    workbook: Workbook,
): void {
    switch (fc.kind) {
        case 'checkbox': {
            // Leading glyph span so consumer CSS can still hook the icon.
            // The glyph text tracks the input's state so the visual indicator
            // stays in sync on change.
            const glyph = document.createElement('span');
            glyph.className = 'xlsx-form-control-glyph';
            glyph.textContent = fc.checked ? '☑' : '☐';
            aside.appendChild(glyph);
            const input = document.createElement('input');
            input.type = 'checkbox';
            input.checked = !!fc.checked;
            input.addEventListener('change', () => {
                glyph.textContent = input.checked ? '☑' : '☐';
                if (fc.linkedCell) {
                    const container = resolveRenderRoot(aside);
                    if (container) {
                        applyFormControlUpdate(container, fc.linkedCell, input.checked ? 'TRUE' : 'FALSE');
                    }
                }
            });
            aside.appendChild(input);
            if (fc.label) {
                const label = document.createElement('span');
                label.className = 'xlsx-form-control-label';
                label.textContent = fc.label;
                aside.appendChild(label);
            }
            return;
        }
        case 'radio': {
            const glyph = document.createElement('span');
            glyph.className = 'xlsx-form-control-glyph';
            glyph.textContent = fc.checked ? '●' : '○';
            aside.appendChild(glyph);
            const input = document.createElement('input');
            input.type = 'radio';
            // Stable group id: prefer a synthesized key from the anchor
            // column (Excel's radios tend to share a column inside a
            // groupBox). Never interpolate attacker-controlled strings
            // into the name attribute — `radioGroupKey` is built from
            // integers only.
            input.name = radioGroupName(fc, sheet);
            input.checked = !!fc.checked;
            input.addEventListener('change', () => {
                glyph.textContent = input.checked ? '●' : '○';
                if (fc.linkedCell && input.checked) {
                    const container = resolveRenderRoot(aside);
                    if (container) {
                        applyFormControlUpdate(container, fc.linkedCell, 'TRUE');
                    }
                }
            });
            aside.appendChild(input);
            if (fc.label) {
                const label = document.createElement('span');
                label.className = 'xlsx-form-control-label';
                label.textContent = fc.label;
                aside.appendChild(label);
            }
            return;
        }
        case 'scrollbar': {
            const input = document.createElement('input');
            input.type = 'range';
            if (fc.min !== null) input.min = String(fc.min);
            if (fc.max !== null) input.max = String(fc.max);
            if (fc.inc !== null) input.step = String(fc.inc);
            if (fc.val !== null) input.value = String(fc.val);
            input.addEventListener('input', () => {
                if (fc.linkedCell) {
                    const container = resolveRenderRoot(aside);
                    if (container) {
                        applyFormControlUpdate(container, fc.linkedCell, input.value);
                    }
                }
            });
            aside.appendChild(input);
            return;
        }
        case 'spinner': {
            const input = document.createElement('input');
            input.type = 'number';
            if (fc.min !== null) input.min = String(fc.min);
            if (fc.max !== null) input.max = String(fc.max);
            if (fc.inc !== null) input.step = String(fc.inc);
            if (fc.val !== null) input.value = String(fc.val);
            input.addEventListener('input', () => {
                if (fc.linkedCell) {
                    const container = resolveRenderRoot(aside);
                    if (container) {
                        applyFormControlUpdate(container, fc.linkedCell, input.value);
                    }
                }
            });
            aside.appendChild(input);
            return;
        }
        case 'combo':
        case 'list': {
            const select = document.createElement('select');
            const optionTexts = fc.inputRange ? resolveInputRangeValues(fc.inputRange, sheet, workbook) : [];
            for (const text of optionTexts) {
                const opt = document.createElement('option');
                opt.textContent = text;
                // The option's value is the option text. setAttribute
                // keeps attacker-controlled strings HTML-encoded.
                opt.setAttribute('value', text);
                select.appendChild(opt);
            }
            select.addEventListener('change', () => {
                if (fc.linkedCell) {
                    const container = resolveRenderRoot(aside);
                    if (container) {
                        applyFormControlUpdate(container, fc.linkedCell, select.value);
                    }
                }
            });
            aside.appendChild(select);
            return;
        }
        case 'button': {
            const btn = document.createElement('button');
            btn.setAttribute('type', 'button');
            if (fc.label) btn.textContent = fc.label;
            // No action is wired — macros / VBA are out of scope. Button
            // labels are attacker-controlled XLSX strings but land via
            // textContent only.
            aside.appendChild(btn);
            return;
        }
        case 'groupBox': {
            const legend = document.createElement('legend');
            if (fc.label) legend.textContent = fc.label;
            aside.appendChild(legend);
            return;
        }
        case 'label':
        case 'dialog':
        case 'unknown':
        default: {
            if (fc.label) {
                const txt = document.createElement('span');
                txt.className = 'xlsx-form-control-label';
                txt.textContent = fc.label;
                aside.appendChild(txt);
            }
            return;
        }
    }
}

// Build a stable, non-attacker-derived name attribute for a radio input. We
// prefer the inputRange (which is an A1 cell reference we already validate
// via parseCellRef) when present; otherwise we fall back to the anchor
// cell coordinates. The sheet's declared name never enters this string —
// that's attacker content and the `name` attribute flows into HTML form
// semantics.
function radioGroupName(fc: SheetFormControl, sheet: Sheet): string {
    // inputRange is an A1 range like `$E$1:$E$4` — safe ASCII only. We
    // still route it through a strict regex to reject surprises. If
    // the range doesn't match, or isn't present, use the anchor col.
    const rangeOk = fc.inputRange && /^\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?$/.test(fc.inputRange);
    const sheetIndexKey = String(Math.max(0, fc.col));
    const key = rangeOk ? fc.inputRange! : `${sheetIndexKey}`;
    // Sanitise further for HTML: the regex above already guaranteed
    // ASCII-only, but we strip any `$` so browsers don't surface a `$`-
    // prefixed name attribute. The sheet name is intentionally NOT part
    // of the string (would be attacker content).
    const sanitized = key.replace(/[$:]/g, '_');
    void sheet; // the sheet name is deliberately NOT part of the group id.
    return `xlsx-radio-${sanitized}`;
}

// Resolve a combo/list control's inputRange to a flat list of option
// strings by reading the cells in the range from the parsed workbook.
// Sheet-prefixed ranges (e.g. `Sheet2!$E$1:$E$4`) are not supported in
// Wave 9 — those return an empty array so the combo still renders as an
// empty <select>. A1 range parsing is intentionally strict: a malformed
// range short-circuits to [].
function resolveInputRangeValues(inputRange: string, sheet: Sheet, _workbook: Workbook): string[] {
    // Strip `$` anchors and the leading sheet prefix (if any).
    if (inputRange.includes('!')) {
        // Sheet-prefixed range — not supported in Wave 9.
        console.warn(`xlsx-preview: sheet-prefixed form-control inputRange "${inputRange}" not supported; skipping`);
        return [];
    }
    const clean = inputRange.replace(/\$/g, '');
    const m = /^([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?$/.exec(clean);
    if (!m) return [];
    const start = parseCellRef(m[1]);
    if (!start) return [];
    const end = m[2] ? parseCellRef(m[2]) : start;
    if (!end) return [];
    const out: string[] = [];
    for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
            const cell = sheet.rows[r]?.find((x) => x.col === c);
            if (cell && cell.value !== '') out.push(cell.value);
        }
    }
    return out;
}

// Walk up from a form-control aside to the closest section.xlsx or the
// containing document fragment so applyFormControlUpdate can find the td
// that hosts the linked cell. The aside lives inside the image layer
// inside its section, so the nearest `section.xlsx` is the right root.
function resolveRenderRoot(el: HTMLElement): HTMLElement | null {
    let cur: HTMLElement | null = el;
    while (cur) {
        if (cur.tagName === 'SECTION' && cur.classList.contains('xlsx')) return cur;
        cur = cur.parentElement;
    }
    return null;
}

// Replace the textContent of the <td> at linkedCell's row/col with the
// stringified newValue. linkedCell is an A1 ref (with optional `$`
// anchors). Sheet-prefixed refs (`Sheet2!$A$1`) are not supported in
// Wave 9 — they log a console.warn and skip the update.
//
// container is the root element the form controls were rendered into
// (typically `section.xlsx` or a parent that contains one). We search
// for the first matching `<td>` inside the first `section.xlsx tbody`:
// xlsxjs renders each row as `<th row-number><td>…</td>…` so the td at
// column `col` is at children[col + 1].
//
// This is a UI-only side effect. The parsed workbook model is NOT
// mutated — consumers who need to persist state should read the td
// textContent off the DOM themselves.
export function applyFormControlUpdate(
    container: HTMLElement,
    linkedCell: string,
    newValue: string,
): void {
    if (linkedCell.includes('!')) {
        console.warn(`xlsx-preview: sheet-prefixed linkedCell "${linkedCell}" not supported; skipping update`);
        return;
    }
    const clean = linkedCell.replace(/\$/g, '');
    const ref = parseCellRef(clean);
    if (!ref) return;
    // Find the first section.xlsx inside (or equal to) the container; that's
    // the sheet the form control sits on. Multi-sheet workbooks: consumers
    // calling applyFormControlUpdate directly are responsible for passing
    // the right section (the event-handler path above picks the enclosing
    // section automatically).
    const section = container.classList?.contains('xlsx') && container.tagName === 'SECTION'
        ? container
        : container.querySelector('section.xlsx');
    if (!section) return;
    const tbody = section.querySelector('tbody');
    if (!tbody) return;
    const tr = tbody.children[ref.row];
    if (!tr) return;
    // Each rendered row starts with a <th> (row number) gutter cell,
    // so the data cell at column `col` lives at children[col + 1].
    const td = tr.children[ref.col + 1];
    if (!td) return;
    td.textContent = newValue;
}

// Build an `<aside class="xlsx-slicer">` carrying the slicer's caption +
// current selection. xlsxjs does NOT render the clickable slicer UI by
// default — this is a read-only summary. All attacker-controlled strings
// reach the DOM via textContent / setAttribute only. aria-hidden="false"
// so screen readers pick the aside up.
//
// Opt-in interactive mode (Options.interactiveSlicers) swaps the flat
// selected-item <ul> for a set of `<button class="xlsx-slicer-chip">`
// toggle chips covering every item the cache enumerated. Each chip tracks
// its pressed state via aria-pressed; clicking a chip flips its state and
// dispatches an `xlsx:slicer-change` CustomEvent on the aside so
// consumers can drive their own pivot filter. xlsxjs does NOT
// re-materialise pivot data.
function renderSlicer(slicer: SheetSlicer, options: Options): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-slicer';
    aside.setAttribute('aria-hidden', 'false');
    aside.setAttribute('data-name', slicer.name);
    if (slicer.caption) aside.setAttribute('data-caption', slicer.caption);
    if (slicer.sourceName) aside.setAttribute('data-source', slicer.sourceName);
    if (slicer.style) aside.setAttribute('data-style', slicer.style);

    // Header carries the caption (or the slicer's raw name when no caption
    // was declared). textContent keeps attacker-controlled strings inert.
    const header = document.createElement('header');
    header.textContent = slicer.caption ?? slicer.name;
    aside.appendChild(header);

    if (options.interactiveSlicers) {
        populateInteractiveSlicer(aside, slicer);
        return aside;
    }

    // Selected-item list. Empty array stays rendered as an empty <ul> so
    // the DOM shape is stable whether Excel wrote "all selected"
    // (no explicit items) or an explicit selection set.
    const ul = document.createElement('ul');
    for (const item of slicer.selectedItems) {
        const li = document.createElement('li');
        li.textContent = item;
        ul.appendChild(li);
    }
    aside.appendChild(ul);
    return aside;
}

// Interactive slicer body. For each item the cache enumerated (all items
// when available, else just the selected ones), emit a
// `<button class="xlsx-slicer-chip">` with aria-pressed tracking whether
// the item is currently part of the selection. Clicking a chip toggles
// its aria-pressed and dispatches an `xlsx:slicer-change` CustomEvent on
// the aside carrying the slicer's name + the current selection snapshot.
// All item labels reach the DOM via textContent / setAttribute only.
function populateInteractiveSlicer(aside: HTMLElement, slicer: SheetSlicer): void {
    // Prefer the full item list so unselected chips surface too; fall back
    // to selectedItems when the cache did not enumerate items (which
    // leaves us with chips only for the current selection).
    const itemsSource = slicer.allItems.length > 0 ? slicer.allItems : slicer.selectedItems;
    // Selected-lookup uses a Set so repeated lookups are O(1) even on
    // large item lists. Set is also safe for attacker-controlled strings —
    // unlike plain objects, it can't trip prototype keys.
    const selectedSet = new Set(slicer.selectedItems);
    const state = new Map<HTMLButtonElement, string>();

    const dispatch = () => {
        const current: string[] = [];
        for (const [btn, label] of state) {
            if (btn.getAttribute('aria-pressed') === 'true') current.push(label);
        }
        // Event name is a hard-coded string literal. The detail carries
        // the slicer's name (from setAttribute-safe metadata) plus a fresh
        // array snapshot so listeners can't mutate our internal state.
        const win = (aside.ownerDocument?.defaultView ?? (typeof window !== 'undefined' ? window : null)) as (Window & typeof globalThis) | null;
        const CE = win?.CustomEvent ?? (typeof CustomEvent !== 'undefined' ? CustomEvent : null);
        if (!CE) return;
        aside.dispatchEvent(new CE('xlsx:slicer-change', {
            bubbles: true,
            detail: { slicer: slicer.name, selectedItems: current },
        }));
    };

    for (const label of itemsSource) {
        const btn = document.createElement('button');
        btn.setAttribute('type', 'button');
        btn.className = 'xlsx-slicer-chip';
        const pressed = selectedSet.has(label);
        btn.setAttribute('aria-pressed', pressed ? 'true' : 'false');
        btn.textContent = label;
        state.set(btn, label);
        btn.addEventListener('click', () => {
            const nextPressed = btn.getAttribute('aria-pressed') !== 'true';
            btn.setAttribute('aria-pressed', nextPressed ? 'true' : 'false');
            dispatch();
        });
        aside.appendChild(btn);
    }
}

// Build an `<aside class="xlsx-timeline">` carrying the timeline's caption +
// active level + selected date range. Same read-only contract as the slicer
// aside. When a selected range is populated we emit a small
// `<span>start → end</span>` label; otherwise the caption alone lands in
// the header.
//
// Opt-in interactive mode (Options.interactiveSlicers) additionally emits
// a `<div class="xlsx-timeline-slider">` with two `<input type="range">`
// handles bounded by the cache's min/max dates. Dragging either handle
// updates a `<span class="xlsx-timeline-label">` label and dispatches an
// `xlsx:timeline-change` CustomEvent on the aside. When the cache did NOT
// enumerate bounds we fall back to the detect-only label only — no slider.
function renderTimeline(timeline: SheetTimeline, options: Options): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-timeline';
    aside.setAttribute('aria-hidden', 'false');
    aside.setAttribute('data-name', timeline.name);
    if (timeline.caption) aside.setAttribute('data-caption', timeline.caption);
    if (timeline.sourceName) aside.setAttribute('data-source', timeline.sourceName);
    if (timeline.level) aside.setAttribute('data-level', timeline.level);
    if (timeline.style) aside.setAttribute('data-style', timeline.style);

    const header = document.createElement('header');
    header.textContent = timeline.caption ?? timeline.name;
    aside.appendChild(header);

    const range = timeline.selectedRange;
    const rangeStart = range?.start ?? null;
    const rangeEnd = range?.end ?? null;

    if (options.interactiveSlicers && timeline.bounds) {
        populateInteractiveTimeline(aside, timeline, timeline.bounds, rangeStart, rangeEnd);
        return aside;
    }

    if (rangeStart || rangeEnd) {
        const label = document.createElement('span');
        // We format the range with a plain " → " separator; start/end
        // are the raw ISO-like strings Excel persisted, e.g.
        // "2023-01-01T00:00:00". Consumers wanting a localised format
        // can read the data-level + re-format from the parsed model.
        label.textContent = `${rangeStart ?? ''} → ${rangeEnd ?? ''}`;
        aside.appendChild(label);
    }
    return aside;
}

// Interactive timeline body. Slider handles are two `<input type="range">`
// elements whose min/max encode the cache's bounds as numeric epoch
// milliseconds; the `.value` of each handle on init is the current
// selected range (falling back to bounds when the selection is open). A
// label beneath the sliders shows the currently-chosen ISO dates and is
// refreshed whenever either handle fires `input`. We dispatch
// `xlsx:timeline-change` on the aside with real Date objects — consumers
// that need the original ISO strings can re-format from `.toISOString()`.
function populateInteractiveTimeline(
    aside: HTMLElement,
    timeline: SheetTimeline,
    bounds: { min: string; max: string },
    rangeStart: string | null,
    rangeEnd: string | null,
): void {
    const minMs = Date.parse(bounds.min);
    const maxMs = Date.parse(bounds.max);
    // Guard against unparseable bounds — we already know the cache had
    // strings, but Date.parse can still return NaN for non-ISO forms. In
    // that case fall through to the detect-only label so we don't crash.
    if (!Number.isFinite(minMs) || !Number.isFinite(maxMs) || maxMs <= minMs) {
        if (rangeStart || rangeEnd) {
            const label = document.createElement('span');
            label.textContent = `${rangeStart ?? ''} → ${rangeEnd ?? ''}`;
            aside.appendChild(label);
        }
        return;
    }

    // Clamp the starting handle positions into [min, max]. An open-ended
    // selection (null endpoint) pins the handle to the bound on that side.
    const startMsRaw = rangeStart ? Date.parse(rangeStart) : minMs;
    const endMsRaw = rangeEnd ? Date.parse(rangeEnd) : maxMs;
    const startMs = Number.isFinite(startMsRaw) ? Math.max(minMs, Math.min(maxMs, startMsRaw)) : minMs;
    const endMs = Number.isFinite(endMsRaw) ? Math.max(minMs, Math.min(maxMs, endMsRaw)) : maxMs;

    const slider = document.createElement('div');
    slider.className = 'xlsx-timeline-slider';

    const minHandle = document.createElement('input');
    minHandle.type = 'range';
    minHandle.setAttribute('data-handle', 'start');
    minHandle.min = String(minMs);
    minHandle.max = String(maxMs);
    minHandle.value = String(startMs);

    const maxHandle = document.createElement('input');
    maxHandle.type = 'range';
    maxHandle.setAttribute('data-handle', 'end');
    maxHandle.min = String(minMs);
    maxHandle.max = String(maxMs);
    maxHandle.value = String(endMs);

    slider.appendChild(minHandle);
    slider.appendChild(maxHandle);
    aside.appendChild(slider);

    const label = document.createElement('span');
    label.className = 'xlsx-timeline-label';
    const fmt = (ms: number): string => new Date(ms).toISOString();
    label.textContent = `${fmt(startMs)} → ${fmt(endMs)}`;
    aside.appendChild(label);

    const onInput = () => {
        let a = Number(minHandle.value);
        let b = Number(maxHandle.value);
        if (!Number.isFinite(a)) a = minMs;
        if (!Number.isFinite(b)) b = maxMs;
        // Keep the two handles ordered. When the start drags past the end
        // (or vice versa) we swap for the event payload but leave the
        // inputs at their as-dragged positions so the user sees what they
        // did — jQuery UI's classic dual-slider behaviour.
        const lo = Math.min(a, b);
        const hi = Math.max(a, b);
        label.textContent = `${fmt(lo)} → ${fmt(hi)}`;
        const win = (aside.ownerDocument?.defaultView ?? (typeof window !== 'undefined' ? window : null)) as (Window & typeof globalThis) | null;
        const CE = win?.CustomEvent ?? (typeof CustomEvent !== 'undefined' ? CustomEvent : null);
        if (!CE) return;
        aside.dispatchEvent(new CE('xlsx:timeline-change', {
            bubbles: true,
            detail: { timeline: timeline.name, start: new Date(lo), end: new Date(hi) },
        }));
    };
    minHandle.addEventListener('input', onInput);
    maxHandle.addEventListener('input', onInput);
}

// Build an <aside class="xlsx-embedding"> for one detected embedding. Name,
// progId, contentType, filename are all attacker-controlled XLSX strings —
// they reach the DOM via setAttribute / textContent only. When dataUrl is
// populated (opt-in via Options.inlineEmbeddings), a sibling `<a download>`
// surfaces so consumers can retrieve the payload; the href lives on the
// sanitized data: URL from bytesToDataUrl, and the download attribute pins
// the filename so the browser doesn't fall back to the raw data: body.
function renderEmbedding(emb: SheetEmbedding): HTMLElement {
    const aside = document.createElement('aside');
    aside.className = 'xlsx-embedding';
    aside.setAttribute('data-kind', emb.kind);
    if (emb.progId) aside.setAttribute('data-progid', emb.progId);
    if (emb.fileName) aside.setAttribute('data-filename', emb.fileName);
    aside.setAttribute('data-content-type', emb.contentType);
    aside.setAttribute('data-size', String(emb.size));
    if (emb.col !== null) aside.setAttribute('data-anchor-col', String(emb.col));
    if (emb.row !== null) aside.setAttribute('data-anchor-row', String(emb.row));
    if (emb.endCol !== null) aside.setAttribute('data-anchor-end-col', String(emb.endCol));
    if (emb.endRow !== null) aside.setAttribute('data-anchor-end-row', String(emb.endRow));
    if (emb.altText) aside.setAttribute('aria-label', emb.altText);

    // Summary block — one line per facet. textContent only.
    const pre = document.createElement('pre');
    const lines: string[] = [];
    lines.push(`${emb.kind}: ${emb.fileName ?? '(no filename)'}`);
    lines.push(`type: ${emb.contentType}`);
    lines.push(`size: ${emb.size} bytes`);
    if (emb.progId) lines.push(`progId: ${emb.progId}`);
    pre.textContent = lines.join('\n');
    aside.appendChild(pre);

    // Opt-in download link. The dataUrl is already sanitized (see
    // bytesToDataUrl) — we only emit an <a> when the projection succeeded.
    // The download filename comes from the rel target; if absent (rare), we
    // derive one from the contentType so the browser still picks something
    // sensible. `rel="noopener"` is belt-and-braces for producers that plumb
    // `target="_blank"` on top.
    if (emb.dataUrl) {
        const a = document.createElement('a');
        a.href = emb.dataUrl;
        a.setAttribute('download', emb.fileName ?? 'embedded');
        a.setAttribute('rel', 'noopener');
        a.textContent = `Download ${emb.fileName ?? 'embedded file'}`;
        aside.appendChild(a);
    }
    return aside;
}

// Build a <div class="xlsx-header|xlsx-footer"> containing three
// <div data-zone="left|center|right"> children. Each zone's text is emitted
// via textContent so attacker strings are HTML-encoded by the DOM.
function renderHeaderFooter(className: string, zones: { left: string; center: string; right: string }): HTMLElement {
    const div = document.createElement('div');
    div.className = className;
    const addZone = (name: 'left' | 'center' | 'right', text: string) => {
        const z = document.createElement('div');
        z.setAttribute('data-zone', name);
        z.textContent = text;
        div.appendChild(z);
    };
    addZone('left', zones.left);
    addZone('center', zones.center);
    addZone('right', zones.right);
    return div;
}

// Append a small "●" marker to a cell that has a classic comment. Author
// and body text are attacker-controlled — we route them through
// setAttribute (HTML-encoded) for the title, never through innerHTML.
function appendCommentMarker(td: HTMLTableCellElement, comment: SheetComment): void {
    const marker = document.createElement('span');
    marker.className = 'xlsx-comment-marker';
    marker.setAttribute('role', 'note');
    const author = comment.author && comment.author.length > 0 ? comment.author : null;
    const title = author ? `${author}: ${comment.text}` : comment.text;
    marker.setAttribute('title', title);
    // Unicode bullet as the visible marker; textContent ensures the glyph
    // is literal rather than interpreted.
    marker.textContent = '●';
    td.appendChild(marker);
}

// Append a 💬 marker to a cell carrying one or more threaded comments.
// The marker's `title` attribute holds the rendered thread — parents
// first, their replies immediately after, in chronological (dT) order.
// Author + text are attacker-controlled strings: they only ever reach
// the DOM via setAttribute('title', …), never innerHTML.
function appendThreadedCommentMarker(td: HTMLTableCellElement, entries: ThreadedCommentEntry[]): void {
    const marker = document.createElement('span');
    marker.className = 'xlsx-comment-marker xlsx-threaded';
    marker.setAttribute('role', 'note');
    marker.setAttribute('title', formatThreadTitle(entries));
    marker.textContent = '💬';
    td.appendChild(marker);
}

// Arrange thread entries so each parent is immediately followed by its
// replies, with both parent and reply lists sorted chronologically by dT.
// Entries with no parent whose parentId references nothing in this group
// are treated as thread starters. Entries whose parent is missing are
// emitted at the end so no data is lost.
function formatThreadTitle(entries: ThreadedCommentEntry[]): string {
    const byId = new Map<string, ThreadedCommentEntry>();
    for (const e of entries) byId.set(e.id, e);

    const sortByDate = (arr: ThreadedCommentEntry[]) =>
        arr.slice().sort((a, b) => (a.date ?? '').localeCompare(b.date ?? ''));

    const starters = sortByDate(entries.filter((e) => !e.parentId || !byId.has(e.parentId)));
    const repliesByParent = new Map<string, ThreadedCommentEntry[]>();
    for (const e of entries) {
        if (!e.parentId || !byId.has(e.parentId)) continue;
        const list = repliesByParent.get(e.parentId);
        if (list) list.push(e);
        else repliesByParent.set(e.parentId, [e]);
    }

    const ordered: ThreadedCommentEntry[] = [];
    for (const starter of starters) {
        ordered.push(starter);
        const replies = repliesByParent.get(starter.id);
        if (replies) for (const r of sortByDate(replies)) ordered.push(r);
    }

    return ordered.map((e) => `${e.author ?? 'Unknown'}: ${e.text}`).join('\n');
}

// Apply <sheetView>/<sheetPr> display hints to the section element.
//   - rightToLeft       → section.dir = 'rtl'
//   - !showGridLines    → .xlsx-no-gridlines class (CSS strips cell borders)
//   - !showRowColHeaders→ .xlsx-no-headers class (CSS hides the header row
//                         and the row-number gutter)
//   - zoomScale !== 100 → section.style.zoom = (zoomScale / 100)
//   - tabColor          → data-tab-color="#rrggbb" (no visual render)
function applySheetView(section: HTMLElement, view: SheetView, theme: Theme | null): void {
    if (view.rightToLeft) section.setAttribute('dir', 'rtl');
    if (!view.showGridLines) section.classList.add('xlsx-no-gridlines');
    if (!view.showRowColHeaders) section.classList.add('xlsx-no-headers');
    if (view.zoomScale !== null && view.zoomScale !== 100) {
        // Chrome honours the non-standard `zoom` property directly; Firefox
        // would need `transform: scale(…)` as a fallback but we keep this
        // simple — the property is ignored without errors elsewhere.
        section.style.zoom = String(view.zoomScale / 100);
    }
    if (view.tabColor) {
        const hex = resolveColor(view.tabColor, theme);
        if (hex) section.setAttribute('data-tab-color', hex);
    }
}

// Wrap a cell's rendered children in an `<a>` when the cell carries a
// hyperlink. The URL is vetted through isSafeHyperlinkHref before it can
// reach an href attribute; attacker URLs (`javascript:`, `data:`, etc.) are
// rejected and the cell stays as inert text. All attacker-controlled strings
// (tooltip, display, location, target) reach the DOM only via setAttribute,
// so the browser does the HTML-encoding.
function wrapCellWithHyperlink(td: HTMLTableCellElement, link: Hyperlink): void {
    // Resolve the effective href: external `target` wins; `location` is the
    // intra-workbook fallback and gets a leading `#` so the browser treats it
    // as a fragment. We never interpolate the location into CSS / innerHTML.
    let href: string | null = null;
    if (link.target != null && link.target !== '' && isSafeHyperlinkHref(link.target)) {
        href = link.target;
    } else if (link.location != null && link.location !== '') {
        // location is a workbook anchor like "Sheet2!A1". It never carries a
        // scheme, but we still route through the allowlist so any pathological
        // producer-written value gets rejected.
        const frag = `#${link.location}`;
        if (isSafeHyperlinkHref(frag)) href = frag;
    }
    if (href == null) return; // rejected → cell stays as plain text

    const a = document.createElement('a');
    a.className = 'xlsx-hyperlink';
    a.setAttribute('href', href);
    // target=_blank + rel=noopener on external links so the viewer doesn't
    // leak its window.opener. Intra-workbook (`#…`) anchors don't need it,
    // but applying it unconditionally keeps the DOM boring to audit.
    a.setAttribute('target', '_blank');
    a.setAttribute('rel', 'noopener noreferrer');
    if (link.tooltip) a.setAttribute('title', link.tooltip);
    // Move the cell's existing children into the anchor. If the cell is
    // empty, fall back to the display text (attribute-only sink → setAttribute
    // wouldn't work; use textContent instead).
    if (td.firstChild) {
        while (td.firstChild) a.appendChild(td.firstChild);
    } else if (link.display) {
        a.textContent = link.display;
    }
    td.appendChild(a);
}

function tagFrozen(td: HTMLTableCellElement, row: number, col: number, panes: FrozenPanes): void {
    const inX = panes.xSplit !== null && col < panes.xSplit;
    const inY = panes.ySplit !== null && row < panes.ySplit;
    if (inX && inY) td.classList.add('xlsx-frozen-both');
    else if (inY) td.classList.add('xlsx-frozen-row');
    else if (inX) td.classList.add('xlsx-frozen-col');
}

function renderCellContent(td: HTMLTableCellElement, cell: Cell, styles: Styles | null, theme: Theme | null, date1904: boolean, options: Options): void {
    const xf = resolveXf(styles, cell.styleIndex);
    let text = cell.value;
    let numeric = cell.kind === 'number';
    // Colour modifier from the number-format code (e.g. `[Red]0;[Blue]-0`).
    // When non-null, the renderer applies it as td.style.color after the
    // xf-level font colour lands, so the format-code colour wins.
    let formatColor: string | null = null;

    // Numeric formatting: applies to numbers and to formula results stored
    // as numbers. Strings, inline strings, booleans, and errors display as-is.
    if ((cell.kind === 'number' || cell.kind === 'empty') && xf) {
        const code = lookupNumberFormat(styles, xf.numFmtId);
        if (code && code !== 'General' && cell.value !== '') {
            const res = formatNumber(cell.value, code, { date1904 });
            text = res.text;
            numeric = res.numeric;
            formatColor = res.color;
        }
    }

    // showFormulas: render the formula text instead of the cached value.
    // Formula text is always displayed as a string (no numeric alignment),
    // and rich-text runs are bypassed so the formula doesn't get mis-formatted.
    //
    // Fallback behaviour: when a formula cell has NO cached value (writers
    // that don't evaluate formulas — python-xlsx emits `<f>…</f><v/>`), show
    // the formula text so the cell isn't blank. Excel re-computes on open,
    // but static HTML consumers can't; a blank cell for `=SUM(…)` is surprising.
    const emptyFormula = cell.formula != null && (cell.value === '' || cell.value == null);
    if ((options.showFormulas && cell.formula != null) || emptyFormula) {
        const a1 = `=${cell.formula}`;
        text = options.formulaNotation === 'r1c1' ? a1ToR1c1(a1, cell.row, cell.col) : a1;
        td.textContent = text;
        td.classList.add('xlsx-formula');
        return; // skip runs path; formula replaces both value and runs
    }

    if (cell.runs && (cell.kind === 'string' || cell.kind === 'inlineStr')) {
        // Rich-text path: emit one <span> per run, applying the run's rPr.
        // Cell-level font (from xf) still applies as the default; runs layer
        // on top. Text content goes via textContent on each span so attacker
        // strings are HTML-encoded by the DOM.
        for (const run of cell.runs) appendRunSpan(td, run, theme);
    } else if (cell.phonetics && (cell.kind === 'string' || cell.kind === 'inlineStr')) {
        // Phonetic-ruby path: the cell carries <rPh> annotations but no rich
        // runs, so we can walk the base text and wrap annotated spans in
        // HTML5 <ruby>. Plain segments between annotations are appended as
        // text nodes — never via innerHTML.
        appendPhoneticText(td, cell.value, cell.phonetics);
    } else {
        td.textContent = text;
    }

    if (cell.kind === 'error') td.classList.add('xlsx-error');
    if (numeric) td.classList.add('xlsx-numeric');

    if (!xf) return;

    if (xf.applyFont || xf.fontId > 0) applyFont(td, styles!.fonts[xf.fontId], theme);
    if (xf.applyFill || xf.fillId > 0) applyFill(td, styles!.fills[xf.fillId], theme);
    if (xf.applyBorder || xf.borderId > 0) applyBorder(td, styles!.borders[xf.borderId], theme);
    if (xf.applyAlignment) applyAlignment(td, xf);
    // Format-code colour wins over xf font colour — apply last so it's the
    // one that lands on the td.
    if (formatColor) td.style.color = formatColor;
}

function appendRunSpan(td: HTMLTableCellElement, run: RichTextRun, theme: Theme | null): void {
    const span = document.createElement('span');
    span.textContent = run.text;
    if (run.bold) span.style.fontWeight = 'bold';
    if (run.italic) span.style.fontStyle = 'italic';
    applyTextDecoration(span, run.underline, run.strike);
    if (run.underline === 'double' || run.underline === 'doubleAccounting') {
        span.style.textDecorationStyle = 'double';
    }
    if (run.underline === 'singleAccounting' || run.underline === 'doubleAccounting') {
        span.classList.add('xlsx-accounting-underline');
    }
    if (run.vertAlign === 'subscript') {
        span.style.verticalAlign = 'sub';
        span.style.fontSize = '0.8em';
    } else if (run.vertAlign === 'superscript') {
        span.style.verticalAlign = 'super';
        span.style.fontSize = '0.8em';
    }
    if (run.size) span.style.fontSize = `${run.size}pt`;
    const color = resolveColor(run.color, theme);
    if (color) span.style.color = color;
    const family = sanitizeFontFamily(run.name);
    if (family) span.style.fontFamily = family;
    td.appendChild(span);
}

// Emit a cell's plain text body interleaved with HTML5 <ruby> wrappers for
// any phonetic (<rPh>) annotations. The base text between / after annotated
// spans goes in as raw text nodes; each phonetic hit becomes
// <ruby>basechars<rt>phonetic</rt></ruby>. Every string reaches the DOM via
// createTextNode / textContent (never innerHTML) so attacker-controlled
// content can't escape.
export function appendPhoneticText(
    parent: Node,
    text: string,
    phonetics: PhoneticRun[],
): void {
    // Sort defensively and clip to the text length so overlapping / reversed
    // producer input can't emit mis-ordered DOM.
    const sorted = phonetics
        .map((p) => ({
            startIdx: Math.max(0, Math.min(text.length, p.startIdx)),
            endIdx: Math.max(0, Math.min(text.length, p.endIdx)),
            phonetic: p.phonetic,
        }))
        .filter((p) => p.endIdx > p.startIdx)
        .sort((a, b) => a.startIdx - b.startIdx);
    let cursor = 0;
    for (const p of sorted) {
        // Skip overlapping entries — we already rendered ruby over this range.
        if (p.startIdx < cursor) continue;
        if (p.startIdx > cursor) {
            parent.appendChild(document.createTextNode(text.slice(cursor, p.startIdx)));
        }
        const ruby = document.createElement('ruby');
        ruby.appendChild(document.createTextNode(text.slice(p.startIdx, p.endIdx)));
        const rt = document.createElement('rt');
        rt.textContent = p.phonetic;
        ruby.appendChild(rt);
        parent.appendChild(ruby);
        cursor = p.endIdx;
    }
    if (cursor < text.length) {
        parent.appendChild(document.createTextNode(text.slice(cursor)));
    }
}

// Compose textDecoration from an underline variant + strike-through. Both
// can coexist, so we build the shorthand from the truthy tokens.
function applyTextDecoration(
    el: HTMLElement,
    underline: FontStyle['underline'] | null | undefined,
    strike: boolean,
): void {
    const parts: string[] = [];
    if (underline) parts.push('underline');
    if (strike) parts.push('line-through');
    if (parts.length === 0) return;
    el.style.textDecoration = parts.join(' ');
}

function resolveXf(styles: Styles | null, index: number): CellXf | null {
    if (!styles || index < 0 || index >= styles.cellXfs.length) return null;
    return resolveEffectiveXf(styles, styles.cellXfs[index]);
}

function applyFont(td: HTMLTableCellElement, font: FontStyle | undefined, theme: Theme | null): void {
    if (!font) return;
    if (font.bold) td.style.fontWeight = 'bold';
    if (font.italic) td.style.fontStyle = 'italic';
    applyTextDecoration(td, font.underline, font.strike);
    if (font.underline === 'double' || font.underline === 'doubleAccounting') {
        td.style.textDecorationStyle = 'double';
    }
    if (font.underline === 'singleAccounting' || font.underline === 'doubleAccounting') {
        td.classList.add('xlsx-accounting-underline');
    }
    if (font.vertAlign === 'subscript') {
        td.style.verticalAlign = 'sub';
        td.style.fontSize = '0.8em';
    } else if (font.vertAlign === 'superscript') {
        td.style.verticalAlign = 'super';
        td.style.fontSize = '0.8em';
    }
    if (font.size) td.style.fontSize = `${font.size}pt`;
    const color = resolveColor(font.color, theme);
    if (color) td.style.color = color;
    // font.name goes through sanitizeFontFamily (strict alphanumeric /
    // space / hyphen / dot allowlist, quoted on output) so an attacker
    // can't escape the font-family CSS property. When no explicit name is
    // set but the font carries a <scheme val="major|minor"/>, resolve the
    // theme-declared typeface and feed that through the sanitizer instead.
    const family = sanitizeFontFamily(font.name) ?? resolveSchemeFontFamily(font.scheme, theme);
    if (family) td.style.fontFamily = family;
}

// Resolve a <font scheme="major|minor"/> reference to a CSS font-family value
// via the theme's fontScheme. The typeface string is routed through
// sanitizeFontFamily so an attacker who tampers with theme1.xml still can't
// break out of the font-family property.
function resolveSchemeFontFamily(scheme: 'major' | 'minor' | null, theme: Theme | null): string | null {
    if (!scheme || !theme) return null;
    const name = scheme === 'major' ? theme.majorFont : theme.minorFont;
    return sanitizeFontFamily(name);
}

function applyFill(td: HTMLTableCellElement, fill: FillStyle | undefined, theme: Theme | null): void {
    if (!fill) return;
    if (fill.kind === 'none') return;
    if (fill.kind === 'gradient') {
        const gradCss = gradientFillToCss(fill, theme);
        if (gradCss) td.style.backgroundImage = gradCss;
        return;
    }
    // Pattern fill. Solid = plain backgroundColor (the fast path and by far
    // the common case). The gridded / striped patterns ride on top of the
    // bgColor via a `repeating-linear-gradient` approximation keyed by
    // patternType; this is NOT a pixel-perfect match of Excel's bitmaps,
    // but the visual indicator (direction + density) survives.
    if (fill.patternType === 'none' || fill.patternType === 'gray125') return;
    const fg = resolveColor(fill.fgColor, theme);
    if (fill.patternType === 'solid') {
        if (fg) td.style.backgroundColor = fg;
        return;
    }
    // Non-solid pattern. Need at least the fg colour to paint marks; if it's
    // missing, bail so the cell stays unfilled.
    if (!fg) return;
    const bg = resolveColor(fill.bgColor, theme) ?? '#ffffff';
    td.style.backgroundColor = bg;
    const pattern = patternFillToCss(fill.patternType, fg);
    if (pattern) td.style.backgroundImage = pattern;
}

// Convert a <gradientFill> to a CSS `linear-gradient(...)`. Path gradients
// (Excel's rectangular radial) don't map cleanly onto CSS radial-gradient
// inside a table cell — we approximate them with a flat fill to the first
// stop's colour so something visible still lands. Linear gradients use the
// degree attribute (0..360) mapped straight to the CSS angle.
function gradientFillToCss(fill: Extract<FillStyle, { kind: 'gradient' }>, theme: Theme | null): string | null {
    // Build resolved stops first so we can bail early when none are usable.
    const resolved: { position: number; hex: string }[] = [];
    for (const stop of fill.stops) {
        const hex = resolveColor(stop.color, theme);
        if (!hex) continue;
        resolved.push({ position: stop.position, hex });
    }
    if (resolved.length === 0) return null;
    if (fill.type === 'path' || resolved.length === 1) {
        // Path gradients fall back to the first stop — CSS radial-gradient in
        // a td cell doesn't give a visually faithful match, so keep it simple.
        // Same for degenerate single-stop linear gradients.
        return `linear-gradient(${resolved[0].hex}, ${resolved[0].hex})`;
    }
    const stops = resolved
        .map((s) => `${s.hex} ${(s.position * 100).toFixed(2)}%`)
        .join(', ');
    return `linear-gradient(${fill.degree}deg, ${stops})`;
}

// Map a non-solid Excel pattern type to a CSS background-image approximation.
// Excel's real renderings are 8x8 bitmaps — we emit `repeating-linear-gradient`
// stripes / crosshatches of 2px fg + 2px transparent (so bgColor shows
// through). The visual density mirrors Excel's classification (dark* is
// busier than light*) but the exact bitmap pattern is a deliberate
// compromise documented here. `fg` is already a resolved #rrggbb colour.
function patternFillToCss(patternType: string, fg: string): string | null {
    const stripe = (angle: string, onPx: number, offPx: number) =>
        `repeating-linear-gradient(${angle}, ${fg} 0 ${onPx}px, transparent ${onPx}px ${onPx + offPx}px)`;
    switch (patternType) {
        // Solid-percentage "gray" patterns. We approximate density by varying
        // the stripe ratio; a 1px mark at 2px pitch reads as ~50% (dark),
        // 2px pitch 3px gap reads as ~40% (medium), 4px pitch reads as light.
        case 'darkGray':       return stripe('45deg', 2, 2);
        case 'mediumGray':     return stripe('45deg', 1, 2);
        case 'lightGray':      return stripe('45deg', 1, 4);
        // Directional stripes. Dark variants use tighter spacing.
        case 'darkHorizontal': return stripe('0deg',  2, 2);
        case 'lightHorizontal':return stripe('0deg',  1, 4);
        case 'darkVertical':   return stripe('90deg', 2, 2);
        case 'lightVertical':  return stripe('90deg', 1, 4);
        case 'darkDown':       return stripe('135deg', 2, 2);
        case 'lightDown':      return stripe('135deg', 1, 4);
        case 'darkUp':         return stripe('45deg',  2, 2);
        case 'lightUp':        return stripe('45deg',  1, 4);
        // Grid/trellis cross two stripe layers via the `,` multi-image syntax.
        case 'darkGrid':
            return `${stripe('0deg', 2, 2)}, ${stripe('90deg', 2, 2)}`;
        case 'lightGrid':
            return `${stripe('0deg', 1, 4)}, ${stripe('90deg', 1, 4)}`;
        case 'darkTrellis':
            return `${stripe('45deg', 2, 2)}, ${stripe('135deg', 2, 2)}`;
        case 'lightTrellis':
            return `${stripe('45deg', 1, 4)}, ${stripe('135deg', 1, 4)}`;
        default:
            return null;
    }
}

function applyBorder(td: HTMLTableCellElement, border: BorderStyle | undefined, theme: Theme | null): void {
    if (!border) return;
    const sides: ['left' | 'right' | 'top' | 'bottom', 'borderLeft' | 'borderRight' | 'borderTop' | 'borderBottom'][] = [
        ['left', 'borderLeft'],
        ['right', 'borderRight'],
        ['top', 'borderTop'],
        ['bottom', 'borderBottom'],
    ];
    for (const [side, cssSide] of sides) {
        const s = border[side];
        if (!s.style) continue;
        const width = borderWidth(s.style);
        const style = borderStyle(s.style);
        const color = resolveColor(s.color, theme) ?? '#000';
        td.style[cssSide] = `${width} ${style} ${color}`;
    }
    // Diagonals. CSS has no border-diagonal; we paint a linear-gradient
    // overlay on the cell background with a narrow band of the resolved
    // colour along the diagonal axis. Stacked gradients for both diagonals.
    // The overlay uses calc(50% ± Xpx) to keep the band width in pixels — a
    // form real browsers accept but jsdom's CSS parser currently discards.
    // Writing into the style attribute directly preserves the raw string
    // (browsers still parse it normally) so the overlay survives round-trip
    // in every environment.
    if (border.diagonalUp || border.diagonalDown) {
        const diagCss = diagonalBorderToCss(border, theme);
        if (diagCss) appendBackgroundImage(td, diagCss);
    }
}

// Append `value` to the element's CSS background-image property, preserving
// any existing background-image. Written via the style *attribute* (not the
// CSSStyleDeclaration) so value strings that contain calc()/modern colour
// syntax survive jsdom's CSS parser unmangled. Browsers parse the raw
// attribute string identically.
function appendBackgroundImage(td: HTMLElement, value: string): void {
    const existingStyle = td.getAttribute('style') ?? '';
    // Pull an existing background-image declaration out of the attribute and
    // merge with the new value. Simple regex — declarations in our output
    // never contain semicolons inside their values.
    const match = /(?:^|;)\s*background-image\s*:\s*([^;]+?)\s*(?=;|$)/i.exec(existingStyle);
    if (match) {
        const merged = `${match[1]}, ${value}`;
        const replaced = existingStyle.slice(0, match.index) +
            (match.index === 0 ? '' : ';') +
            `background-image: ${merged}` +
            existingStyle.slice(match.index + match[0].length);
        td.setAttribute('style', replaced);
        return;
    }
    const separator = existingStyle && !existingStyle.trim().endsWith(';') ? '; ' : '';
    td.setAttribute('style', `${existingStyle}${separator}background-image: ${value};`);
}

// Build a multi-gradient string representing the active diagonal borders.
// Each diagonal is a "thin band of colour" gradient: transparent up to the
// diagonal line, the colour for ±half-width, then transparent again. The
// band width comes from the diagonal BorderSide's style (same mapping as
// the orthogonal sides). When both diagonals are active we stack two
// gradients with a comma; CSS blends them onto the same element.
function diagonalBorderToCss(border: BorderStyle, theme: Theme | null): string | null {
    const styleName = border.diagonal.style ?? 'thin';
    const widthPx = Math.max(1, parseInt(borderWidth(styleName), 10) || 1);
    const half = widthPx / 2;
    const color = resolveColor(border.diagonal.color, theme) ?? '#000';
    const gradients: string[] = [];
    const band = (direction: string) =>
        `linear-gradient(${direction}, ` +
            `transparent calc(50% - ${half}px), ` +
            `${color} calc(50% - ${half}px), ` +
            `${color} calc(50% + ${half}px), ` +
            `transparent calc(50% + ${half}px))`;
    // diagonalDown: top-left → bottom-right (CSS "to bottom right").
    if (border.diagonalDown) gradients.push(band('to bottom right'));
    // diagonalUp: bottom-left → top-right (CSS "to top right").
    if (border.diagonalUp) gradients.push(band('to top right'));
    if (gradients.length === 0) return null;
    return gradients.join(', ');
}

function borderWidth(style: string): string {
    switch (style) {
        case 'thick': return '3px';
        case 'medium':
        case 'mediumDashed':
        case 'mediumDashDot':
        case 'mediumDashDotDot':
            return '2px';
        default:
            return '1px';
    }
}

function borderStyle(style: string): string {
    if (style.toLowerCase().includes('dashed') || style === 'dashDot' || style === 'dashDotDot') return 'dashed';
    if (style === 'dotted' || style === 'hair') return 'dotted';
    if (style === 'double') return 'double';
    return 'solid';
}

function applyAlignment(td: HTMLTableCellElement, xf: CellXf): void {
    const a = xf.alignment;

    // Horizontal: left/right/center/justify map straight to CSS text-align.
    // distributed + centerContinuous don't have a direct CSS equivalent;
    // we render them as 'justify' and 'center' respectively — close enough
    // visually for spreadsheets. fill repeats the content to fill the cell
    // in Excel; we approximate as text-align:start with a note (no reliable
    // pure-CSS "repeat to fill" behaviour exists inside a table cell).
    if (a.horizontal) {
        switch (a.horizontal) {
            case 'centerContinuous':
                td.style.textAlign = 'center';
                break;
            case 'distributed':
                td.style.textAlign = 'justify';
                td.style.textAlignLast = 'justify';
                break;
            case 'fill':
                // Pure CSS can't replicate Excel's "repeat content across
                // cell width". Align to the start edge and document the gap.
                td.style.textAlign = 'start';
                break;
            default:
                td.style.textAlign = a.horizontal;
        }
    }

    // Vertical: middle / top / bottom map directly; justify + distributed
    // don't exist as CSS vertical-align values, so we approximate as
    // middle (the typical Excel use case is header rows where content is
    // short and the distributed look is indistinguishable from middle).
    if (a.vertical) {
        if (a.vertical === 'middle' || a.vertical === 'justify' || a.vertical === 'distributed') {
            td.style.verticalAlign = 'middle';
        } else {
            td.style.verticalAlign = a.vertical;
        }
    }

    // wrapText: normal white-space so \n + long runs wrap; word-break so
    // narrow cells still wrap very long unbroken strings.
    if (a.wrapText) {
        td.style.whiteSpace = 'normal';
        td.style.wordBreak = 'break-word';
    }

    // readingOrder: 2 = RTL; 1 = LTR (explicit); 0 = context (no style).
    if (a.readingOrder === 2) td.style.direction = 'rtl';
    else if (a.readingOrder === 1) td.style.direction = 'ltr';

    // indent: 1 unit ≈ 0.5em by Excel's convention. Apply to the leading
    // side based on horizontal + readingOrder — right-aligned cells indent
    // from the right, RTL flips both sides.
    if (a.indent > 0) {
        const em = `${(a.indent * 0.5).toFixed(2)}em`;
        const rtl = a.readingOrder === 2;
        // right-aligned or RTL + left-aligned → padding on the right.
        const padRight = (a.horizontal === 'right' && !rtl) || (a.horizontal === 'left' && rtl);
        if (padRight) td.style.paddingRight = em;
        else td.style.paddingLeft = em;
    }

    // textRotation: 255 = stacked vertical (CJK convention); other values
    // (1..180) are counter-clockwise rotation in degrees. Pure CSS rotation
    // inside a table cell is imperfect — the cell's laid-out box doesn't
    // grow to accommodate rotated content, so callers relying on exact
    // header-height match with Excel may need extra padding. Documented
    // compromise; the common case (small-angle header labels) renders
    // legibly.
    if (a.textRotation === 255) {
        td.style.writingMode = 'vertical-lr';
    } else if (a.textRotation != null && a.textRotation !== 0) {
        // Excel rotates counter-clockwise for positive angles; CSS
        // rotate uses clockwise-positive, so negate.
        td.style.transform = `rotate(-${a.textRotation}deg)`;
        td.style.transformOrigin = 'center center';
        td.style.display = 'inline-block';
    }

    // shrinkToFit: we can't measure the rendered text at parse time and
    // browsers have no pure-CSS "shrink font to fit container" rule.
    // Tag the td so consumers can target it; the default CSS rule in
    // renderStyle sets font-size via clamp() as a sensible floor. A
    // JS-based auto-fit pass is left to downstream apps (it's cheap once
    // the DOM is in-document, but out of scope for xlsxjs itself).
    if (a.shrinkToFit) td.classList.add('xlsx-shrink-to-fit');
}

function applyGraphicalCf(td: HTMLTableCellElement, state: GraphicalCfState): void {
    if (state.colorScaleBg) {
        td.style.backgroundColor = state.colorScaleBg;
        td.classList.add('xlsx-cf-colorscale');
    }
    if (state.dataBar) {
        // Render the bar as a horizontal linear gradient, so we avoid adding
        // child elements to the td (keeps rich-text / formatted content
        // layout untouched). The gradient also survives colorScaleBg
        // rendering as a fallback since we set the gradient directly.
        const bar = state.dataBar;
        const pct = +(Math.max(0, Math.min(1, bar.fraction)) * 100).toFixed(2);
        // Direction: default (context / leftToRight) fills from the left;
        // rightToLeft mirrors the gradient so the bar grows from the right.
        const angle = bar.direction === 'rightToLeft' ? '270deg' : '90deg';
        // Gradient vs flat: gradient === false emits a solid fill up to pct
        // without the opacity fade. The classic Excel "gradient" look fades
        // within the filled portion, which the CSS linear-gradient shorthand
        // below approximates by using the fill colour then transparent.
        const fillStop = bar.gradient
            ? `${bar.color} 0 ${pct}%, transparent ${pct}% 100%`
            : `${bar.color} 0 ${pct}%, transparent ${pct}% 100%`;
        td.style.background = `linear-gradient(${angle}, ${fillStop})`;
        if (bar.border) {
            // Fall back to the fill colour when borderColor is null — matches
            // Excel's "use fill colour for border" default.
            const bc = bar.borderColor ?? bar.color;
            td.style.border = `1px solid ${bc}`;
        }
        if (bar.axis) {
            // Record the axis position + colour as data-attrs so consumer CSS
            // can paint a vertical line without us needing to inject a child
            // element. Middle-position axes also get an inset box-shadow that
            // draws a 1px vertical line down the middle of the cell.
            td.setAttribute('data-cf-databar-axis', bar.axis.position);
            td.setAttribute('data-cf-databar-axis-color', bar.axis.color);
            if (bar.axis.position === 'middle') {
                // `inset` box-shadow of -50% offset paints a 1px line down
                // the column at 50% from the left edge. Multiple shadows
                // layer, but we don't compose with existing shadows here —
                // cells in xlsxjs aren't styled with box-shadow elsewhere.
                const existingShadow = td.style.boxShadow;
                const axisShadow = `inset 50% 0 0 -49% ${bar.axis.color}`;
                td.style.boxShadow = existingShadow ? `${existingShadow}, ${axisShadow}` : axisShadow;
            }
        }
        td.classList.add('xlsx-cf-databar');
    }
    if (state.icon) {
        // Prepend an <span> holding the inline SVG. We set the SVG via
        // innerHTML on the wrapper — the content is generated by xlsxjs
        // itself, not attacker-controlled, so it's safe. Optionally hide
        // the cell value when the cf rule asked us to (showValue=false).
        const iconSpan = document.createElement('span');
        iconSpan.className = 'xlsx-cf-icon';
        iconSpan.style.display = 'inline-flex';
        iconSpan.style.verticalAlign = 'middle';
        iconSpan.style.marginRight = '0.3em';
        iconSpan.innerHTML = state.icon.svg;
        if (!state.icon.showValue) {
            // Wrap existing content in an invisible span so the text still
            // flows in a11y tooling but isn't visually present.
            const hide = document.createElement('span');
            while (td.firstChild) hide.appendChild(td.firstChild);
            hide.style.visibility = 'hidden';
            td.appendChild(hide);
        }
        td.insertBefore(iconSpan, td.firstChild);
        td.classList.add('xlsx-cf-iconset');
    }
}

// Layer a dxf's properties on top of an already-styled td. Only applied when
// a conditional-formatting rule fired, so any property the dxf sets wins
// over the cell's base xf styling.
function applyDxf(td: HTMLTableCellElement, dxf: Dxf, theme: Theme | null, cell: Cell | null, date1904: boolean): void {
    if (dxf.font) {
        const f = dxf.font;
        if (f.bold) td.style.fontWeight = 'bold';
        if (f.italic) td.style.fontStyle = 'italic';
        // Strike / underline: apply via the shared helper so the decoration
        // shorthand is consistent with applyFont. Compose with any existing
        // decoration on the td so a dxf-added strike doesn't clobber a
        // cell-level underline (or vice versa).
        const strike = !!f.strike;
        const underline = f.underline ?? null;
        if (strike || underline) {
            const existingDecoration = td.style.textDecoration ?? '';
            const existingUnderline = /underline/.test(existingDecoration);
            const existingStrike = /line-through/.test(existingDecoration);
            applyTextDecoration(td, underline || (existingUnderline ? 'single' : null), strike || existingStrike);
            if (underline === 'double' || underline === 'doubleAccounting') {
                td.style.textDecorationStyle = 'double';
            }
            if (underline === 'singleAccounting' || underline === 'doubleAccounting') {
                td.classList.add('xlsx-accounting-underline');
            }
        }
        if (f.vertAlign === 'subscript') {
            td.style.verticalAlign = 'sub';
            td.style.fontSize = '0.8em';
        } else if (f.vertAlign === 'superscript') {
            td.style.verticalAlign = 'super';
            td.style.fontSize = '0.8em';
        }
        if (f.size) td.style.fontSize = `${f.size}pt`;
        if (f.color) {
            const color = resolveColor(f.color, theme);
            if (color) td.style.color = color;
        }
        const family = sanitizeFontFamily(f.name);
        if (family) td.style.fontFamily = family;
    }
    if (dxf.fill) applyFill(td, dxf.fill, theme);
    if (dxf.border) applyBorder(td, dxf.border, theme);
    // dxf numFmt: when present, re-run the number formatter against the cell's
    // raw value using the dxf's formatCode and replace the td's text. Skipped
    // for rich-text (runs) so we don't destroy per-run formatting, and for
    // non-numeric cells (strings, errors, booleans) where a number format
    // would be meaningless.
    if (dxf.numFmtCode && cell && (cell.kind === 'number' || cell.kind === 'empty')
        && cell.value !== '' && !cell.runs) {
        const code = dxf.numFmtCode;
        if (code !== 'General') {
            const res = formatNumber(cell.value, code, { date1904 });
            td.textContent = res.text;
            if (res.numeric) td.classList.add('xlsx-numeric');
        }
    }
    td.classList.add('xlsx-cf');
}
