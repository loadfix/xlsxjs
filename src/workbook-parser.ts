// XLSX XML → internal model. Scope for the initial slice: sheet names, shared
// strings, and per-cell values (string / number / inline string / boolean).
// No styles, no number formats, no formulas — those land in follow-ups.
//
// The model shape is intentionally small and stable so the renderer can walk
// it without caring about the on-disk XML structure.

import { parseCellRef } from './utils';
import type { Options } from './xlsx-preview';
import { parseStyles, parseColorElement, type Styles, type FontStyle } from './styles';
import { parseTheme, type Theme, type ColorRef } from './theme';
import { parseConditionalFormatting, type ConditionalFormatting } from './conditional-format';

export interface RichTextRun {
    text: string;
    // Subset of FontStyle — runs carry their own rPr. Properties not present
    // on the run inherit from the cell's xf font.
    bold: boolean;
    italic: boolean;
    underline: boolean;
    size: number | null;
    color: ColorRef;
    name: string | null;
}

// A shared-string entry. Plain strings carry text only; rich entries also
// carry the runs so the renderer can emit <span> per run with the correct
// styling. `text` is always the concatenated plain-text projection.
export interface SharedString {
    text: string;
    runs: RichTextRun[] | null;
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
}

export interface RowDimension {
    // 0-based row index. Present only when the row carries a custom height
    // and/or hidden flag — sparse by design.
    row: number;
    // Height in points. null when only the hidden flag was set.
    height: number | null;
    hidden: boolean;
}

// Frozen / split pane metadata from <sheetViews><sheetView><pane ...>. xlsxjs
// renders frozen panes as sticky <th>/<td> borders; splits without freeze
// are not represented visually.
export interface FrozenPanes {
    // 0-based column index at which horizontal freeze begins, or null when
    // no horizontal freeze. A value of 3 means cols 0..2 are frozen.
    xSplit: number | null;
    // Same for rows.
    ySplit: number | null;
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
}

// A pivot table anchored inside a sheet. Detection-only — the sheet's cell
// values already carry the materialised pivot output.
export interface SheetPivot {
    name: string;
    col: number; row: number;
    endCol: number; endRow: number;
}

// Count of sheet-level <extLst><ext uri="…"> entries. Surfaced so the smoke
// tool can roll up unknown extension uses (sparklines, dynamic-array spill,
// protected ranges, etc.) without re-scanning the raw XML.
export interface SheetExtensionUri {
    uri: string;
    count: number;
}

// A rendered image anchored inside a sheet. Coordinates are 0-based.
// xlsxjs ships only the "twoCellAnchor" and "oneCellAnchor" positioning
// — absolute pixel anchors are rare and reported with `col=0, row=0`.
export interface SheetImage {
    col: number; row: number;           // top-left anchor cell
    endCol: number | null;              // bottom-right anchor cell (twoCellAnchor)
    endRow: number | null;
    // Offsets from the top-left anchor cell, in EMUs (English Metric Units;
    // 914400 per inch, 9525 per pixel). Renderer converts to px.
    colOff: number; rowOff: number;
    // Data URL + intrinsic dimensions (or null if the drawing didn't declare them).
    dataUrl: string;
    widthEmu: number | null;
    heightEmu: number | null;
    alt: string | null;
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
}

export interface Sheet {
    name: string;
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
    pivots: SheetPivot[];
    extensions: SheetExtensionUri[];
}

export interface Workbook {
    sheets: Sheet[];
    styles: Styles | null;
    theme: Theme | null;
}

const NS = {
    main:  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    rel:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rels:  'http://schemas.openxmlformats.org/package/2006/relationships',
    xdr:   'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    a:     'http://schemas.openxmlformats.org/drawingml/2006/main',
    c:     'http://schemas.openxmlformats.org/drawingml/2006/chart',
    cx:    'http://schemas.microsoft.com/office/drawing/2014/chartex',
};

export class WorkbookParser {
    constructor(private _options: Options) {}

    parse(parts: Record<string, string>, media: Record<string, string> = {}): Workbook {
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

        const sheetMeta = parseSheetList(workbookXml);
        const sheets: Sheet[] = [];
        for (let i = 0; i < sheetMeta.length; i++) {
            const { name, rId } = sheetMeta[i];
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
            const { images, charts } = resolveDrawingsForSheet(xmlPath, parts, media);
            const pivots = resolvePivotsForSheet(xmlPath, parts);
            sheets.push(parseSheet(name, xml, sharedStrings, tables, images, charts, pivots));
        }

        return { sheets, styles, theme };
    }
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
): { images: SheetImage[]; charts: SheetChart[] } {
    const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relsXml = parts[relsPath];
    if (!relsXml) return { images: [], charts: [] };
    const rels = parseRelationships(relsXml);
    const dir = sheetPath.replace(/\/[^/]+$/, '');
    const images: SheetImage[] = [];
    const charts: SheetChart[] = [];
    for (const [, rel] of rels) {
        if (!rel.type.endsWith('/drawing')) continue;
        const drawingPath = rel.target.startsWith('/')
            ? rel.target.slice(1)
            : normaliseRelPath(`${dir}/${rel.target}`);
        const drawingXml = parts[drawingPath];
        if (!drawingXml) continue;
        // The drawing has its own rels part binding rIds to media files
        // and to chart parts.
        const drawingRelsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const drawingRelsXml = parts[drawingRelsPath];
        if (!drawingRelsXml) continue;
        const drawingRels = parseRelationships(drawingRelsXml);
        const drawingDir = drawingPath.replace(/\/[^/]+$/, '');
        const parsed = parseDrawing(drawingXml, drawingRels, drawingDir, parts, media);
        images.push(...parsed.images);
        charts.push(...parsed.charts);
    }
    return { images, charts };
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

function parseDrawing(
    xml: string,
    rels: Map<string, { target: string; type: string }>,
    drawingDir: string,
    parts: Record<string, string>,
    media: Record<string, string>,
): { images: SheetImage[]; charts: SheetChart[] } {
    const doc = parseXml(xml);
    const images: SheetImage[] = [];
    const charts: SheetChart[] = [];
    // Both twoCellAnchor and oneCellAnchor carry child <pic> / <graphicFrame>
    // elements. We walk each anchor in order, checking both.
    const anchors = [
        ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'twoCellAnchor')),
        ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'oneCellAnchor')),
    ];
    for (const anchor of anchors) {
        const from = anchor.getElementsByTagNameNS(NS.xdr, 'from').item(0);
        const to = anchor.getElementsByTagNameNS(NS.xdr, 'to').item(0);
        const col = anchorCellValue(from, 'col');
        const row = anchorCellValue(from, 'row');
        const endCol = to ? anchorCellValue(to, 'col') : null;
        const endRow = to ? anchorCellValue(to, 'row') : null;
        const colOff = anchorCellValue(from, 'colOff');
        const rowOff = anchorCellValue(from, 'rowOff');

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
                    // Extent (explicit size for oneCellAnchor / absolute). Not
                    // always present on twoCellAnchor; renderer falls back to
                    // anchor-based sizing.
                    const ext = anchor.getElementsByTagNameNS(NS.xdr, 'ext').item(0);
                    const widthEmu = ext ? Number(ext.getAttribute('cx')) : null;
                    const heightEmu = ext ? Number(ext.getAttribute('cy')) : null;
                    // Alt text lives under pic/nvPicPr/cNvPr.
                    const cNvPr = pic.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                    const alt = cNvPr?.getAttribute('descr') ?? cNvPr?.getAttribute('title') ?? null;
                    images.push({
                        col: col ?? 0,
                        row: row ?? 0,
                        endCol,
                        endRow,
                        colOff: colOff ?? 0,
                        rowOff: rowOff ?? 0,
                        dataUrl,
                        widthEmu: Number.isFinite(widthEmu as number) ? (widthEmu as number) : null,
                        heightEmu: Number.isFinite(heightEmu as number) ? (heightEmu as number) : null,
                        alt,
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
            if (rId) {
                const rel = rels.get(rId);
                if (rel) {
                    const chartPath = rel.target.startsWith('/')
                        ? rel.target.slice(1)
                        : normaliseRelPath(`${drawingDir}/${rel.target}`);
                    const chartXml = parts[chartPath];
                    if (chartXml) chartType = peekChartType(chartXml, kind);
                }
            }
            charts.push({
                kind,
                chartType,
                col: col ?? 0,
                row: row ?? 0,
                endCol,
                endRow,
            });
        }
    }
    return { images, charts };
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
    };
}

function parseXml(xml: string): Document {
    return new DOMParser().parseFromString(xml, 'application/xml');
}

function parseSheetList(workbookXml: string): { name: string; rId: string | null }[] {
    const doc = parseXml(workbookXml);
    const nodes = doc.getElementsByTagNameNS(NS.main, 'sheet');
    const out: { name: string; rId: string | null }[] = [];
    for (let i = 0; i < nodes.length; i++) {
        const rId = nodes[i].getAttributeNS(NS.rel, 'id');
        out.push({
            name: nodes[i].getAttribute('name') ?? `Sheet${i + 1}`,
            rId: rId || null,
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
// We walk direct children in order so a mix produces the correct runs.
function parseSi(si: Element): SharedString {
    const runs: RichTextRun[] = [];
    let sawRun = false;
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
        }
    }
    const text = runs.map((r) => r.text).join('');
    return { text, runs: sawRun ? runs : null };
}

function emptyRun(text: string): RichTextRun {
    return {
        text,
        bold: false,
        italic: false,
        underline: false,
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
    return {
        text,
        bold: has('b'),
        italic: has('i'),
        underline: has('u'),
        size: sizeAttr ? Number(sizeAttr) : null,
        color: parseColorElement(color),
        name: firstAttr('rFont', 'val') ?? firstAttr('name', 'val'),
    };
}

function parseSheet(
    name: string,
    xml: string,
    sharedStrings: SharedString[],
    tables: TableDef[] = [],
    images: SheetImage[] = [],
    charts: SheetChart[] = [],
    pivots: SheetPivot[] = [],
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
        if (height !== null || hidden) {
            rowDimensions.push({ row: rowIndex, height, hidden });
            if (rowIndex > maxRow) maxRow = rowIndex;
        }

        const cellEls = rowEl.getElementsByTagNameNS(NS.main, 'c');
        const cells: Cell[] = [];
        for (let j = 0; j < cellEls.length; j++) {
            const cell = parseCell(cellEls[j], rowIndex, sharedStrings);
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

    return {
        name, rows, maxCol, maxRow, merges, columns, rowDimensions,
        conditionalFormatting, frozenPanes, autoFilter, tables, images,
        charts, pivots, extensions,
    };
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
    // Only render freeze / frozenSplit states; a plain "split" is a split
    // without freezing and we don't support that visually.
    if (state !== 'frozen' && state !== 'frozenSplit') return null;
    const xSplit = Number(pane.getAttribute('xSplit'));
    const ySplit = Number(pane.getAttribute('ySplit'));
    const x = Number.isFinite(xSplit) && xSplit > 0 ? xSplit : null;
    const y = Number.isFinite(ySplit) && ySplit > 0 ? ySplit : null;
    if (x === null && y === null) return null;
    return { xSplit: x, ySplit: y };
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
        // Excel writes a <col> for every column range even without a custom
        // width (carrying style info or the hidden flag). Keep the entry
        // whenever either width or hidden is meaningful.
        const hasWidth = widthAttr !== null && (customWidth || !Number.isNaN(Number(widthAttr)));
        if (!hasWidth && !hidden) continue;
        if (!Number.isFinite(min) || !Number.isFinite(max)) continue;
        if (min < 1 || max < min) continue;
        const width = hasWidth && Number.isFinite(Number(widthAttr)) ? Number(widthAttr) : null;
        out.push({ min: min - 1, max: max - 1, width, hidden });
    }
    return out;
}

function parseCell(c: Element, fallbackRow: number, sharedStrings: SharedString[]): Cell | null {
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
    const base = { col, row, styleIndex, formula };

    switch (type) {
        case 's': {
            const idx = Number(raw);
            const entry = Number.isFinite(idx) && idx >= 0 && idx < sharedStrings.length
                ? sharedStrings[idx]
                : { text: '', runs: null };
            return { ...base, value: entry.text, runs: entry.runs, kind: 'string' };
        }
        case 'inlineStr': {
            const is = c.getElementsByTagNameNS(NS.main, 'is').item(0);
            const entry = is ? parseSi(is) : { text: '', runs: null };
            return { ...base, value: entry.text, runs: entry.runs, kind: 'inlineStr' };
        }
        case 'b':
            return { ...base, value: raw === '1' ? 'TRUE' : 'FALSE', runs: null, kind: 'boolean' };
        case 'e':
            return { ...base, value: raw, runs: null, kind: 'error' };
        case 'str':
            return { ...base, value: raw, runs: null, kind: 'string' };
        case 'n':
        default:
            if (raw === '') return { ...base, value: '', runs: null, kind: 'empty' };
            return { ...base, value: raw, runs: null, kind: 'number' };
    }
}
