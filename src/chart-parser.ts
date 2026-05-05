// Chart-model parser. Converts a single `xl/charts/chart*.xml`
// (classic `c:chartSpace`) into a typed ChartModel that the renderer can
// project to SVG. Defensive on unknown chart types: returns a blank
// `kind: 'unknown'` model so callers can fall through to the dashed
// placeholder without throwing.
//
// Scope is deliberately narrow — just the common chart kinds: bar
// (horizontal), column (vertical bar), line, pie, doughnut, scatter, area,
// and radar. Everything else (3D, stock, bubble, surface, chartEx) decays
// to `kind: 'unknown'`.
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
    // Scatter series carry a parallel `<c:xVal>` cache alongside `<c:val>` —
    // both coordinate streams are needed so the renderer can plot each
    // marker at its (x, y). Non-scatter series leave this null; renderers
    // fall back to category-index ordering.
    xValues: (number | null)[] | null;
    // Series colour from `<c:spPr>/<a:solidFill>/<a:srgbClr val="…"/>`.
    // Leave null when the chart XML doesn't spell one out — the renderer
    // assigns a palette colour by series index.
    color: string | null;
    // Data-label configuration for this series. `show` is true when a
    // `<c:showVal val="1"/>` (or a parent chart-level `<c:dLbls>`) flagged
    // the series. `position` picks up `<c:dLblPos val="…"/>` (ctr, outEnd,
    // t, b, l, r, inBase, inEnd, bestFit). Chart-level dLbls inherit to
    // every series; per-series dLbls overrides.
    dataLabels: { show: boolean; position: string | null };
}

export interface ChartModel {
    kind: 'bar' | 'column' | 'line' | 'pie' | 'scatter' | 'area' | 'radar' | 'doughnut' | 'chartex' | 'unknown';
    title: string | null;
    categories: string[];
    series: ChartSeries[];
    legend: 'top' | 'right' | 'bottom' | 'left' | 'none';
    // Classic bar / column / line chart grouping, lifted straight from
    // `<c:grouping val="…"/>`. `null` for chart types that don't carry one
    // (pie, scatter) or when the element was missing. The renderer reads
    // this to decide between clustered vs stacked vs 100%-stacked layouts;
    // `kind` stays 'bar' / 'column' / 'line' — grouping is orthogonal.
    grouping: 'standard' | 'stacked' | 'percentStacked' | null;
    // chartEx (cx:chartSpace, Office 2014) subtype. Lifted straight from
    // `<cx:series layoutId="…">` — e.g. 'treemap', 'sunburst', 'waterfall',
    // 'funnel', 'paretoLine', 'boxWhisker', 'regionMap', 'histogram',
    // 'clusteredColumn'. null when the chart part isn't chartEx or the
    // layoutId couldn't be resolved. xlsxjs does NOT render chartEx plots
    // to SVG in this wave — the field exists so consumers + the placeholder
    // path can surface the subtype as a `data-chartex-type` attribute.
    chartExSubtype: string | null;
}

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
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
    return readNumericRefOrLit(ser, 'val');
}

// Scatter (and bubble, in the future) series declare an `<c:xVal>` sibling
// to `<c:val>` that holds the horizontal coordinate stream. Same structure
// as `<c:val>` — a numRef/numCache happy path with numLit fallback. We
// return null rather than an empty array when no xVal is declared, so the
// renderer can distinguish "scatter with missing data" from "non-scatter
// series" cleanly.
function parseXValues(ser: Element): (number | null)[] | null {
    const xVal = firstChild(ser, NS_C, 'xVal');
    if (!xVal) return null;
    return readNumericCacheOrLit(xVal);
}

// Shared reader for `<c:val>` / `<c:xVal>` blocks: each wraps a
// `<c:numRef>/<c:numCache>` or a bare `<c:numLit>`. Missing / non-numeric
// cells decay to null.
function readNumericRefOrLit(ser: Element, localName: string): (number | null)[] {
    const el = firstChild(ser, NS_C, localName);
    if (!el) return [];
    return readNumericCacheOrLit(el);
}

function readNumericCacheOrLit(container: Element): (number | null)[] {
    const numRef = firstChild(container, NS_C, 'numRef');
    const cache = numRef ? firstChild(numRef, NS_C, 'numCache') : null;
    const source = cache ?? firstChild(container, NS_C, 'numLit');
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
// `kind` influences which coordinate streams to read — scatter pulls
// `<c:yVal>` into `values` and `<c:xVal>` into `xValues`; everything else
// reads the single `<c:val>` stream and leaves `xValues` null.
function parseSer(
    ser: Element,
    fallbackCategories: string[] | null,
    kind: ChartModel['kind'],
    chartLevelLabels: { show: boolean; position: string | null },
): ChartSeries {
    let values: (number | null)[];
    let xValues: (number | null)[] | null;
    if (kind === 'scatter') {
        // Scatter series spell their y-coordinate as `<c:yVal>`. Happy
        // path; if the producer wrote `<c:val>` instead (some legacy /
        // LibreOffice outputs do) fall through to the common reader.
        const yVal = readNumericRefOrLit(ser, 'yVal');
        values = yVal.length ? yVal : parseValues(ser);
        xValues = parseXValues(ser);
    } else {
        values = parseValues(ser);
        xValues = null;
    }
    return {
        name: parseSeriesName(ser),
        values,
        xValues,
        color: parseSeriesColor(ser),
        dataLabels: parseSeriesDataLabels(ser, chartLevelLabels),
    };
    // `fallbackCategories` is unused here — categories come from the first
    // series and are hoisted by `parseChart`. Parameter kept for call-site
    // clarity in future extensions (e.g. per-series category refinement).
    void fallbackCategories;
}

// Read a `<c:dLbls>` block — either chart-level (on the plot element) or
// per-series (inside `<c:ser>`). `<c:showVal val="1"/>` anywhere under the
// block flags the labels as visible; `<c:dLblPos val="…"/>` pulls a
// position keyword. Missing / malformed values return the defaults
// { show: false, position: null }.
function parseDataLabelsBlock(block: Element | null): { show: boolean; position: string | null } {
    if (!block) return { show: false, position: null };
    let show = false;
    let position: string | null = null;
    // Walk direct children first, then dLbl-scoped nested children — the
    // chart-level `<c:dLbls>` may wrap its flags inside `<c:dLbl>` siblings
    // but the common shape xlsxjs reads is the flat form used by Excel.
    for (let i = 0; i < block.childNodes.length; i++) {
        const node = block.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI !== NS_C) continue;
        if (el.localName === 'showVal') {
            if (el.getAttribute('val') === '1') show = true;
        } else if (el.localName === 'dLblPos') {
            const v = el.getAttribute('val');
            if (v) position = v;
        } else if (el.localName === 'dLbl') {
            // A per-point override — read the same two flags so the chart
            // can still opt in via a sibling dLbl.
            for (let j = 0; j < el.childNodes.length; j++) {
                const c = el.childNodes[j];
                if (c.nodeType !== 1) continue;
                const ce = c as Element;
                if (ce.namespaceURI !== NS_C) continue;
                if (ce.localName === 'showVal' && ce.getAttribute('val') === '1') show = true;
                if (ce.localName === 'dLblPos') {
                    const v = ce.getAttribute('val');
                    if (v && position === null) position = v;
                }
            }
        }
    }
    return { show, position };
}

// Per-series data labels. Looks at `<c:ser>/<c:dLbls>`. If the series
// doesn't declare its own block, the chart-level defaults flow through.
// When the series does declare a block, its flags win for keys it sets;
// unset keys still inherit.
function parseSeriesDataLabels(
    ser: Element,
    chartLevel: { show: boolean; position: string | null },
): { show: boolean; position: string | null } {
    const block = firstChild(ser, NS_C, 'dLbls');
    if (!block) return { show: chartLevel.show, position: chartLevel.position };
    const perSer = parseDataLabelsBlock(block);
    return {
        show: perSer.show || chartLevel.show,
        position: perSer.position ?? chartLevel.position,
    };
}

// Grouping (`standard` | `stacked` | `percentStacked`) applies to bar,
// column, line, and area plots. Absent / unrecognised values map to
// 'standard' so the renderer can treat it as the non-stacked default.
function parseGrouping(chartEl: Element): ChartModel['grouping'] {
    const g = firstChild(chartEl, NS_C, 'grouping');
    const val = g?.getAttribute('val');
    if (val === 'stacked') return 'stacked';
    if (val === 'percentStacked') return 'percentStacked';
    if (val === 'standard' || val === 'clustered') return 'standard';
    return null;
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
// pulled out of the series children). The plot element itself is returned
// so the caller can lift `<c:grouping>` off it without re-walking.
function dispatchChart(plotArea: Element): { kind: ChartModel['kind']; sers: Element[]; plotEl: Element | null } {
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
                return { kind, sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            }
            case 'lineChart':
                return { kind: 'line', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            case 'pieChart':
                return { kind: 'pie', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            case 'doughnutChart':
                return { kind: 'doughnut', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            case 'scatterChart':
                return { kind: 'scatter', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            case 'areaChart':
                return { kind: 'area', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            case 'radarChart':
                return { kind: 'radar', sers: directChildren(el, NS_C, 'ser'), plotEl: el };
            default:
                continue;
        }
    }
    return { kind: 'unknown', sers: [], plotEl: null };
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
        grouping: null,
        chartExSubtype: null,
    };

    let doc: Document;
    try {
        doc = parseXml(xml);
    } catch {
        return blank;
    }

    // chartEx (cx:chartSpace) detection. The 2014 extension namespace
    // carries treemap / sunburst / waterfall / funnel / paretoLine /
    // boxWhisker / regionMap / histogram / clusteredColumn layouts.
    // xlsxjs surfaces the subtype as `chartExSubtype` and lifts categories
    // + first-series values from the first `<cx:data>` block so consumers
    // get a round-trippable model, but rendering falls through to the
    // dashed placeholder — per-subtype SVG renderers are a future wave.
    const cxChartSpace = doc.getElementsByTagNameNS(NS_CX, 'chartSpace').item(0);
    if (cxChartSpace) return parseChartEx(cxChartSpace, blank);

    const chartSpace = doc.getElementsByTagNameNS(NS_C, 'chartSpace').item(0);
    if (!chartSpace) return blank;
    const chart = firstChild(chartSpace, NS_C, 'chart');
    if (!chart) return blank;
    const plotArea = firstChild(chart, NS_C, 'plotArea');
    if (!plotArea) return blank;

    const { kind, sers, plotEl } = dispatchChart(plotArea);
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
    // Scatter charts don't have categories — their x-axis is numeric.
    let categories: string[] = [];
    if (kind !== 'scatter') {
        for (const ser of sers) {
            const cat = parseCategories(ser);
            if (cat.length) { categories = cat; break; }
        }
    }

    // Chart-level `<c:dLbls>` sits directly under the plot element (e.g.
    // inside `<c:barChart>`). It applies to every series unless the series
    // declares its own `<c:dLbls>`. When it's absent we pass { show: false,
    // position: null } through so individual series still pick up their own
    // flags or decay to the defaults.
    const chartLevelLabels = plotEl
        ? parseDataLabelsBlock(firstChild(plotEl, NS_C, 'dLbls'))
        : { show: false, position: null };

    const series: ChartSeries[] = sers.map((ser) => parseSer(ser, categories, kind, chartLevelLabels));

    // Grouping applies to bar / column / line / area. Pie / doughnut /
    // scatter / radar leave it null — stacking is meaningless on them.
    const grouping = (kind === 'bar' || kind === 'column' || kind === 'line' || kind === 'area') && plotEl
        ? parseGrouping(plotEl)
        : null;

    return {
        kind,
        title: parseTitle(chart),
        categories,
        series,
        legend: parseLegend(chart),
        grouping,
        chartExSubtype: null,
    };
}

// chartEx (cx:chartSpace) parser. Surfaces the subtype (`treemap`,
// `sunburst`, `waterfall`, `funnel`, `paretoLine`, `boxWhisker`,
// `regionMap`, `histogram`, `clusteredColumn`) plus a best-effort
// round-trip of the first data block's categories + values so consumers
// can display a title / summary. The model's `kind` is always 'chartex'
// — we do NOT fold chartEx subtypes into the classic kind enum because
// the renderer deliberately falls through to the dashed placeholder
// (per-subtype SVG renderers are a future wave).
//
// The chartEx data layout is fundamentally different from classic:
//   <cx:chartData>/<cx:data id="0">/
//       <cx:strDim type="cat"><cx:pt>…</cx:pt></cx:strDim>
//       <cx:numDim type="val"><cx:pt>…</cx:pt></cx:numDim>
//   <cx:plotArea>/<cx:plotAreaRegion>/<cx:series layoutId="…">
//       <cx:tx><cx:txData><cx:v>…</cx:v></cx:txData></cx:tx>
//       <cx:dataId val="0"/>
// We read the first <cx:data> block regardless of its `id` — xlsxjs
// doesn't resolve the series/data `dataId` pointer in this wave.
function parseChartEx(chartSpace: Element, blank: ChartModel): ChartModel {
    const chart = firstChild(chartSpace, NS_CX, 'chart');
    const title = chart ? parseChartExTitle(chart) : null;

    // layoutId lives on the first <cx:series> inside
    // <cx:plotArea>/<cx:plotAreaRegion>. Fallback: any <cx:series> anywhere.
    let subtype: string | null = null;
    let seriesName: string | null = null;
    if (chart) {
        const plotArea = firstChild(chart, NS_CX, 'plotArea');
        const plotRegion = plotArea ? firstChild(plotArea, NS_CX, 'plotAreaRegion') : null;
        const firstSer = plotRegion ? firstChild(plotRegion, NS_CX, 'series') : null;
        const ser = firstSer ?? chartSpace.getElementsByTagNameNS(NS_CX, 'series').item(0);
        if (ser) {
            const layoutId = ser.getAttribute('layoutId');
            if (layoutId) subtype = layoutId;
            seriesName = parseChartExSeriesName(ser);
        }
    }

    // Categories + values from the first <cx:data> block. In the ECMA-376
    // schema <cx:chartData> sits as a child of <cx:chartSpace> (sibling
    // of <cx:chart>), not inside <cx:chart>. Defensive: also walk inside
    // <cx:chart> so producers that flatten the hierarchy still round-trip.
    const categories: string[] = [];
    const values: (number | null)[] = [];
    const chartData = firstChild(chartSpace, NS_CX, 'chartData')
        ?? (chart ? firstChild(chart, NS_CX, 'chartData') : null);
    const firstData = chartData ? firstChild(chartData, NS_CX, 'data') : null;
    if (firstData) {
        for (let i = 0; i < firstData.childNodes.length; i++) {
            const node = firstData.childNodes[i];
            if (node.nodeType !== 1) continue;
            const el = node as Element;
            if (el.namespaceURI !== NS_CX) continue;
            if (el.localName === 'strDim') {
                for (const s of readChartExPts(el)) categories.push(s);
            } else if (el.localName === 'numDim') {
                for (const s of readChartExPts(el)) {
                    if (s === '') { values.push(null); continue; }
                    const n = Number(s);
                    values.push(Number.isFinite(n) ? n : null);
                }
            }
        }
    }

    const series: ChartSeries[] = [];
    // Surface a single synthetic series when we actually lifted values —
    // consumers (and scenario 116) expect `series[0].values` to round-trip.
    if (values.length || seriesName !== null) {
        series.push({
            name: seriesName,
            values,
            xValues: null,
            color: null,
            dataLabels: { show: false, position: null },
        });
    }

    return {
        ...blank,
        kind: 'chartex',
        title,
        categories,
        series,
        legend: 'none',
        chartExSubtype: subtype,
    };
}

// Read the <cx:pt> text children of a chartEx dimension block (strDim /
// numDim). Honours the `idx` attribute; unindexed points fall into doc
// order. Mirrors readPts() for classic but on the cx: namespace.
function readChartExPts(parent: Element): string[] {
    const pts: Element[] = [];
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.namespaceURI === NS_CX && el.localName === 'pt') pts.push(el);
    }
    if (!pts.length) return [];
    let maxIdx = -1;
    const raw: { idx: number; v: string }[] = [];
    for (const pt of pts) {
        const idxAttr = pt.getAttribute('idx');
        const idx = idxAttr !== null ? Number(idxAttr) : raw.length;
        const text = pt.textContent ?? '';
        raw.push({ idx: Number.isFinite(idx) ? idx : raw.length, v: text });
        if (idx > maxIdx) maxIdx = idx;
    }
    const out: string[] = [];
    for (let i = 0; i <= maxIdx; i++) out.push('');
    for (const r of raw) out[r.idx] = r.v;
    return out;
}

// chartEx title: <cx:chart>/<cx:title>/<cx:tx>/<cx:rich>/<a:p>/<a:r>/<a:t>
// or <cx:tx>/<cx:txData>/<cx:v>. Mirrors the classic title resolver but on
// the cx: namespace.
function parseChartExTitle(chart: Element): string | null {
    const title = firstChild(chart, NS_CX, 'title');
    if (!title) return null;
    const tx = firstChild(title, NS_CX, 'tx');
    if (!tx) return null;
    const rich = firstChild(tx, NS_CX, 'rich');
    if (rich) {
        const text = collectRichText(rich).trim();
        if (text.length) return text;
    }
    const txData = firstChild(tx, NS_CX, 'txData');
    if (txData) {
        const v = firstChild(txData, NS_CX, 'v');
        if (v) {
            const text = (v.textContent ?? '').trim();
            if (text.length) return text;
        }
    }
    return null;
}

// Series name: <cx:series>/<cx:tx>/<cx:txData>/<cx:v>.
function parseChartExSeriesName(ser: Element): string | null {
    const tx = firstChild(ser, NS_CX, 'tx');
    if (!tx) return null;
    const txData = firstChild(tx, NS_CX, 'txData');
    if (!txData) return null;
    const v = firstChild(txData, NS_CX, 'v');
    if (!v) return null;
    const text = v.textContent ?? '';
    return text.length ? text : null;
}
