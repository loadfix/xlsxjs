// Chart-model parser. Converts a single `xl/charts/chart*.xml`
// (classic `c:chartSpace`) into a typed ChartModel that the renderer can
// project to SVG. Defensive on unknown chart types: returns a blank
// `kind: 'unknown'` model so callers can fall through to the dashed
// placeholder without throwing.
//
// Scope is deliberately narrow — just the common chart kinds needed for a
// first rendering pass: bar (horizontal), column (vertical bar), line, pie,
// scatter, and area. Everything else (3D, stock, bubble, surface, radar,
// chartEx) decays to `kind: 'unknown'`.
//
// Attacker surface: title / category / series-name strings all reach the
// renderer as plain string model fields; the renderer is responsible for
// writing them to the DOM via `textContent` (never innerHTML). We preserve
// raw text here — no stripping — so the renderer sees exactly what the
// .xlsx declared.

import type { SharedString } from './workbook-parser';

export interface ChartSeries {
    name: string | null;
    values: (number | null)[];
    // Series colour from `<c:spPr>/<a:solidFill>/<a:srgbClr val="…"/>`.
    // Leave null when the chart XML doesn't spell one out — the renderer
    // assigns a palette colour by series index.
    color: string | null;
}

export interface ChartModel {
    kind: 'bar' | 'column' | 'line' | 'pie' | 'scatter' | 'area' | 'unknown';
    title: string | null;
    categories: string[];
    series: ChartSeries[];
    legend: 'top' | 'right' | 'bottom' | 'left' | 'none';
}

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function parseXml(xml: string): Document {
    return new DOMParser().parseFromString(xml, 'application/xml');
}

function firstChild(parent: Element | null, ns: string, localName: string): Element | null {
    if (!parent) return null;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI === ns && el.localName === localName) return el;
    }
    return null;
}

function directChildren(parent: Element, ns: string, localName: string): Element[] {
    const out: Element[] = [];
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI === ns && el.localName === localName) out.push(el);
    }
    return out;
}

// Read `<c:v>` text children of `<c:pt>` entries under a container,
// honouring the `idx` attribute so sparse-indexed points land at the right
// slot. Used for both string caches (category labels) and numeric caches
// (category / value series). Returns strings; caller coerces to number.
function readPts(parent: Element): string[] {
    const pts = directChildren(parent, NS_C, 'pt');
    if (!pts.length) return [];
    let maxIdx = -1;
    const raw: { idx: number; v: string }[] = [];
    for (const pt of pts) {
        const idxAttr = pt.getAttribute('idx');
        const idx = idxAttr !== null ? Number(idxAttr) : raw.length;
        const vEl = firstChild(pt, NS_C, 'v');
        const v = vEl ? (vEl.textContent ?? '') : '';
        raw.push({ idx: Number.isFinite(idx) ? idx : raw.length, v });
        if (idx > maxIdx) maxIdx = idx;
    }
    const out: string[] = [];
    for (let i = 0; i <= maxIdx; i++) out.push('');
    for (const r of raw) out[r.idx] = r.v;
    return out;
}

// Resolve category labels. Looks first at `<c:cat>/<c:strRef>/<c:strCache>`
// (the common case — Excel caches category text alongside the formula
// reference), then falls back to a numRef's numCache when the axis is
// numeric. Returns an empty array when no cache is present; that's the
// contract the renderer relies on to lay out "index" x-axis ticks.
function parseCategories(ser: Element): string[] {
    const cat = firstChild(ser, NS_C, 'cat');
    if (!cat) return [];
    const strRef = firstChild(cat, NS_C, 'strRef');
    if (strRef) {
        const cache = firstChild(strRef, NS_C, 'strCache');
        if (cache) return readPts(cache);
    }
    const numRef = firstChild(cat, NS_C, 'numRef');
    if (numRef) {
        const cache = firstChild(numRef, NS_C, 'numCache');
        if (cache) return readPts(cache);
    }
    // Literal string / numeric arrays (strLit / numLit).
    const strLit = firstChild(cat, NS_C, 'strLit');
    if (strLit) return readPts(strLit);
    const numLit = firstChild(cat, NS_C, 'numLit');
    if (numLit) return readPts(numLit);
    return [];
}

// Resolve series values. `<c:val>/<c:numRef>/<c:numCache>/<c:pt>/<c:v>` is
// the happy path for every chart type we render. Missing / malformed
// numeric text decays to null so the renderer can skip-or-gap cleanly.
function parseValues(ser: Element): (number | null)[] {
    const val = firstChild(ser, NS_C, 'val');
    if (!val) return [];
    const numRef = firstChild(val, NS_C, 'numRef');
    const cache = numRef ? firstChild(numRef, NS_C, 'numCache') : null;
    const source = cache ?? firstChild(val, NS_C, 'numLit');
    if (!source) return [];
    const raw = readPts(source);
    return raw.map((s) => {
        if (s === '') return null;
        const n = Number(s);
        return Number.isFinite(n) ? n : null;
    });
}

// Series name resolution. In declining priority:
//   · `<c:tx>/<c:strRef>/<c:strCache>/<c:pt>/<c:v>` (cached label)
//   · `<c:tx>/<c:v>` (literal — older producers)
//   · `<c:tx>/<c:rich>/<a:p>/<a:r>/<a:t>` (rich-text label, runs joined)
function parseSeriesName(ser: Element): string | null {
    const tx = firstChild(ser, NS_C, 'tx');
    if (!tx) return null;
    const strRef = firstChild(tx, NS_C, 'strRef');
    if (strRef) {
        const cache = firstChild(strRef, NS_C, 'strCache');
        if (cache) {
            const pts = readPts(cache);
            if (pts.length) return pts[0];
        }
    }
    const vEl = firstChild(tx, NS_C, 'v');
    if (vEl) return vEl.textContent ?? null;
    const rich = firstChild(tx, NS_C, 'rich');
    if (rich) return collectRichText(rich);
    return null;
}

// Walk an `<a:p>` body and join every `<a:r>/<a:t>` + direct `<a:t>` text
// into a single string, mirroring how Excel projects a title to a label.
function collectRichText(root: Element): string {
    const parts: string[] = [];
    const visit = (el: Element) => {
        for (let i = 0; i < el.childNodes.length; i++) {
            const node = el.childNodes[i];
            if (node.nodeType !== 1) continue;
            const child = node as Element;
            if (child.namespaceURI === NS_A && child.localName === 't') {
                parts.push(child.textContent ?? '');
                continue;
            }
            visit(child);
        }
    };
    visit(root);
    return parts.join('');
}

// Series colour: first `<c:spPr>/<a:solidFill>/<a:srgbClr val="…"/>`. We
// deliberately only honour the srgbClr variant here — theme colours and
// schemeClr would need the workbook theme to resolve, which the parser
// doesn't have in scope. Unresolved colours fall through to the renderer's
// palette.
function parseSeriesColor(ser: Element): string | null {
    const spPr = firstChild(ser, NS_C, 'spPr');
    if (!spPr) return null;
    const solidFill = firstChild(spPr, NS_A, 'solidFill');
    if (!solidFill) return null;
    const srgb = firstChild(solidFill, NS_A, 'srgbClr');
    if (!srgb) return null;
    const val = srgb.getAttribute('val');
    if (!val || !/^[0-9a-fA-F]{6}$/.test(val)) return null;
    return `#${val.toLowerCase()}`;
}

// Parse a `<c:ser>` block into one ChartSeries. Used by every chart kind.
function parseSer(ser: Element, fallbackCategories: string[] | null): ChartSeries {
    return {
        name: parseSeriesName(ser),
        values: parseValues(ser),
        color: parseSeriesColor(ser),
    };
    // `fallbackCategories` is unused here — categories come from the first
    // series and are hoisted by `parseChart`. Parameter kept for call-site
    // clarity in future extensions (e.g. per-series category refinement).
    void fallbackCategories;
}

// Parse the `<c:title>` block. Prefers `<c:title>/<c:tx>/<c:rich>` over
// `<c:title>/<c:tx>/<c:strRef>/<c:strCache>` — Excel typically writes rich
// for titles, and the cache path is rare.
function parseTitle(chart: Element): string | null {
    const title = firstChild(chart, NS_C, 'title');
    if (!title) return null;
    const tx = firstChild(title, NS_C, 'tx');
    if (!tx) return null;
    const rich = firstChild(tx, NS_C, 'rich');
    if (rich) {
        const text = collectRichText(rich).trim();
        return text.length ? text : null;
    }
    const strRef = firstChild(tx, NS_C, 'strRef');
    if (strRef) {
        const cache = firstChild(strRef, NS_C, 'strCache');
        if (cache) {
            const pts = readPts(cache);
            if (pts.length && pts[0]) return pts[0];
        }
    }
    return null;
}

// Legend position. `<c:legend>/<c:legendPos val="…"/>` — Excel writes `r`,
// `b`, `t`, `l`, `tr`, `none`. We map `tr` to `right` (close enough for a
// first-pass renderer). A missing `<c:legend>` resolves to `'none'`.
function parseLegend(chart: Element): ChartModel['legend'] {
    const legend = firstChild(chart, NS_C, 'legend');
    if (!legend) return 'none';
    const pos = firstChild(legend, NS_C, 'legendPos');
    const val = pos?.getAttribute('val');
    switch (val) {
        case 't': return 'top';
        case 'b': return 'bottom';
        case 'l': return 'left';
        case 'r':
        case 'tr':
        case null:
        case undefined:
            return 'right';
        default:
            return 'right';
    }
}

// Dispatch: look at the first child of `<c:plotArea>` that names a chart
// type and return the matching ChartModel.kind (plus category + series data
// pulled out of the series children).
function dispatchChart(plotArea: Element): { kind: ChartModel['kind']; sers: Element[] } {
    for (let i = 0; i < plotArea.childNodes.length; i++) {
        const node = plotArea.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI !== NS_C) continue;
        switch (el.localName) {
            case 'barChart': {
                const barDir = firstChild(el, NS_C, 'barDir');
                const dir = barDir?.getAttribute('val');
                // Bar = horizontal; col (default) = vertical columns.
                const kind: ChartModel['kind'] = dir === 'bar' ? 'bar' : 'column';
                return { kind, sers: directChildren(el, NS_C, 'ser') };
            }
            case 'lineChart':
                return { kind: 'line', sers: directChildren(el, NS_C, 'ser') };
            case 'pieChart':
            case 'doughnutChart':
                return { kind: 'pie', sers: directChildren(el, NS_C, 'ser') };
            case 'scatterChart':
                return { kind: 'scatter', sers: directChildren(el, NS_C, 'ser') };
            case 'areaChart':
                return { kind: 'area', sers: directChildren(el, NS_C, 'ser') };
            default:
                continue;
        }
    }
    return { kind: 'unknown', sers: [] };
}

export function parseChart(xml: string, sharedStrings: SharedString[]): ChartModel {
    // `sharedStrings` is currently unused — series text always lives inside
    // the chart's own strCache/numCache. Kept on the signature so future
    // chartSpace extensions (e.g. formula resolution back to the sheet)
    // have a place to land without a public API break.
    void sharedStrings;

    const blank: ChartModel = {
        kind: 'unknown',
        title: null,
        categories: [],
        series: [],
        legend: 'none',
    };

    let doc: Document;
    try {
        doc = parseXml(xml);
    } catch {
        return blank;
    }

    const chartSpace = doc.getElementsByTagNameNS(NS_C, 'chartSpace').item(0);
    if (!chartSpace) return blank;
    const chart = firstChild(chartSpace, NS_C, 'chart');
    if (!chart) return blank;
    const plotArea = firstChild(chart, NS_C, 'plotArea');
    if (!plotArea) return blank;

    const { kind, sers } = dispatchChart(plotArea);
    if (kind === 'unknown') {
        // Still surface title / legend even when the plot type isn't one we
        // render — consumers can use the title for their own placeholder.
        return {
            ...blank,
            title: parseTitle(chart),
            legend: parseLegend(chart),
        };
    }

    // Categories come from the first series that declares a `<c:cat>`.
    // Series order preserves document order (Excel's drawing order too).
    let categories: string[] = [];
    for (const ser of sers) {
        const cat = parseCategories(ser);
        if (cat.length) { categories = cat; break; }
    }

    const series: ChartSeries[] = sers.map((ser) => parseSer(ser, categories));

    return {
        kind,
        title: parseTitle(chart),
        categories,
        series,
        legend: parseLegend(chart),
    };
}
