// Internal model → DOM. The rendered output is one <section class="xlsx"> per
// sheet, each containing an <h2> sheet name and a <table> of the cells.
// Numeric cells get a numeric-aligned class; other kinds render as text.

import type { Workbook, Sheet, SheetView, Cell, MergedRange, RichTextRun, FrozenPanes, SheetComment, SheetShape, ThreadedCommentEntry, Hyperlink, PhoneticRun } from './workbook-parser';
import { isSafeHyperlinkHref } from './workbook-parser';
import { indexToColumnLetters, emuToPx } from './utils';
import { h } from './html';
import type { Options } from './xlsx-preview';
import { lookupNumberFormat, resolveEffectiveXf, sanitizeFontFamily, type Styles, type CellXf, type FontStyle, type FillStyle, type BorderStyle, type Dxf } from './styles';
import { resolveColor, type Theme } from './theme';
import { formatNumber } from './number-format';
import { a1ToR1c1 } from './formula-notation';
import { renderIcon } from './icons';
import { evaluateRule, resolveCfvo, interpolateColorScale, type ConditionalFormatting, type CfContext, type CellRange, type CfRule } from './conditional-format';

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

export class HtmlRenderer {
    async render(workbook: Workbook, options: Options): Promise<Node[]> {
        const nodes: Node[] = [];
        nodes.push(renderStyle(options.className));
        for (const sheet of workbook.sheets) {
            // Skip hidden and veryHidden sheets — the demo's sheet-switcher
            // omits them too so section/index pairings stay in sync.
            if (sheet.state !== 'visible') continue;
            nodes.push(renderSheet(sheet, workbook.styles, workbook.theme, workbook.date1904, options));
        }
        return nodes;
    }
}

function renderStyle(className: string): HTMLStyleElement {
    const style = document.createElement('style');
    style.setAttribute('data-xlsxjs', '');
    // Kept intentionally small — consumers style further via their own CSS.
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
}
.${className} .xlsx-shape[data-kind="connector"] {
    border-style: dashed; color: #888;
}
.${className} .xlsx-shape pre {
    margin: 0; white-space: pre-line;
    font-family: inherit; font-size: inherit;
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
}
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
        const svg = renderIcon(icons.iconSet, idx, icons.reverse);
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

function renderSheet(sheet: Sheet, styles: Styles | null, theme: Theme | null, date1904: boolean, options: Options): HTMLElement {
    const section = h('section', { class: options.className, 'data-sheet-name': sheet.name }) as HTMLElement;
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
    section.appendChild(h('div', { class: 'xlsx-sheet-name' }, [sheet.name]));

    const table = h('table') as HTMLTableElement;

    if (sheet.maxCol < 0) {
        section.appendChild(table);
        return section;
    }

    const colCount = sheet.maxCol + 1;

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
    headRow.appendChild(h('th'));
    for (let c = 0; c < colCount; c++) {
        const th = h('th', null, [indexToColumnLetters(c)]);
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
    const rowCount = sheet.maxRow + 1;
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
        tr.appendChild(h('th', null, [String(r + 1)]));
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

    // Images. Each gets its own <figure> with an <img> whose size is derived
    // from the image's extent (EMUs). We don't attempt to place it inside the
    // table — positioning images against a browser table layout is fragile —
    // but we do annotate the anchor on the DOM so consumers who want to
    // overlay can read off the coordinates. absoluteAnchor images get a
    // position:absolute + left/top in CSS pixels so a caller who gives the
    // section a position:relative ancestor gets free pixel-absolute layout.
    for (const img of sheet.images) {
        const fig = document.createElement('figure');
        fig.className = 'xlsx-image';
        fig.setAttribute('data-anchor-mode', img.anchorMode);
        fig.setAttribute('data-anchor-col', String(img.col));
        fig.setAttribute('data-anchor-row', String(img.row));
        if (img.endCol !== null) fig.setAttribute('data-anchor-end-col', String(img.endCol));
        if (img.endRow !== null) fig.setAttribute('data-anchor-end-row', String(img.endRow));
        fig.style.margin = '0.5rem 0';
        if (img.anchorMode === 'absolute') {
            fig.style.position = 'absolute';
            if (img.absoluteX !== null) fig.style.left = `${emuToPx(img.absoluteX)}px`;
            if (img.absoluteY !== null) fig.style.top = `${emuToPx(img.absoluteY)}px`;
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
        if (img.widthEmu && img.heightEmu) {
            el.width = emuToPx(img.widthEmu);
            el.height = emuToPx(img.heightEmu);
        }
        el.style.maxWidth = '100%';
        fig.appendChild(el);
        section.appendChild(fig);
    }

    // Chart placeholders. We don't render chart content (the chart XML
    // describes a plot, not a renderable bitmap), but we surface a dashed
    // placeholder so consumers can see a chart lives at this anchor — and so
    // diff tools notice drift against an Excel re-save.
    for (const chart of sheet.charts) {
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
    if (shape.text && shape.text.length > 0) {
        const pre = document.createElement('pre');
        pre.textContent = shape.text;
        aside.appendChild(pre);
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
    if (options.showFormulas && cell.formula != null) {
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
    const color = resolveColor(fill?.fgColor ?? null, theme);
    if (!color) return;
    td.style.backgroundColor = color;
}

function applyBorder(td: HTMLTableCellElement, border: BorderStyle | undefined, theme: Theme | null): void {
    if (!border) return;
    const sides: [keyof BorderStyle, 'borderLeft' | 'borderRight' | 'borderTop' | 'borderBottom'][] = [
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
