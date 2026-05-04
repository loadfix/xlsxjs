// ChartModel → inline SVG. First-pass renderer: bar / column / line / pie.
// Unknown / unsupported kinds return null so callers can fall through to
// the dashed placeholder.
//
// Security contract:
//   · All XLSX-derived strings (title, series names, category labels) reach
//     the DOM via `textContent`. SVG `<text>` supports textContent; we never
//     use innerHTML here.
//   · Numeric inputs are coerced with Number.isFinite and clamped into the
//     plot area — a NaN / Infinity value can never escape to an attribute.
//   · Attributes are set with setAttribute; the only string we interpolate
//     into a colour attribute is validated against a 6-hex-digit regex
//     (colours coming from the model are already validated by the parser;
//     the palette is a hard-coded constant).
//
// Layout model:
//   · width × height (defaults 480×300) — callers can pass larger for e.g.
//     embedded / full-screen renderings. viewBox mirrors the pixel box.
//   · Margins: top=40, right=40, bottom=60, left=80. The title sits in the
//     top margin; the x-axis category labels use the bottom margin; the
//     y-axis numeric labels use the left margin.
//   · Legend sits just inside the side/bottom/top margin that matches
//     `model.legend`. `none` suppresses the legend entirely.

import type { ChartModel, ChartSeries } from './chart-parser';

export type { ChartModel, ChartSeries } from './chart-parser';

const SVG_NS = 'http://www.w3.org/2000/svg';

// Excel 2013 default colour cycle. Lowercase hex, matches Office's accent
// palette within a frame of one or two. Series beyond index 7 wrap.
const PALETTE = [
    '#5b9bd5',
    '#ed7d31',
    '#a5a5a5',
    '#ffc000',
    '#4472c4',
    '#70ad47',
    '#264478',
    '#9e480e',
] as const;

const GRID_COLOR = '#ddd';
const AXIS_COLOR = '#888';
const TEXT_COLOR = '#333';

interface Layout {
    width: number;
    height: number;
    plotX: number;
    plotY: number;
    plotW: number;
    plotH: number;
}

function makeSvg(width: number, height: number): SVGSVGElement {
    const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    svg.setAttribute('xmlns', SVG_NS);
    svg.setAttribute('width', String(width));
    svg.setAttribute('height', String(height));
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
    svg.setAttribute('class', 'xlsx-chart-svg');
    return svg;
}

function el(tag: string, attrs: Record<string, string | number>): SVGElement {
    const node = document.createElementNS(SVG_NS, tag);
    for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
    return node;
}

function textEl(
    x: number, y: number, text: string, attrs: Record<string, string | number> = {},
): SVGElement {
    const t = document.createElementNS(SVG_NS, 'text');
    t.setAttribute('x', String(x));
    t.setAttribute('y', String(y));
    t.setAttribute('fill', TEXT_COLOR);
    t.setAttribute('font-family', 'system-ui, sans-serif');
    t.setAttribute('font-size', '12');
    for (const [k, v] of Object.entries(attrs)) t.setAttribute(k, String(v));
    // Attacker-controlled strings land here as textContent; SVG text honours
    // it exactly like HTML, so no inner-markup can escape.
    t.textContent = text;
    return t;
}

function colorFor(series: ChartSeries, idx: number): string {
    if (series.color && /^#[0-9a-f]{6}$/i.test(series.color)) return series.color;
    return PALETTE[idx % PALETTE.length];
}

// Find a "nice" axis max. Simple algorithm: compute the raw max, pick a
// scale that is a power of 10 below it, and round up to the nearest
// 1/2/5/10 multiple of that scale. Min is assumed zero for bar/column/line
// (negative values supported via `axisMinFor`).
function niceMax(max: number): number {
    if (!Number.isFinite(max) || max <= 0) return 1;
    const mag = Math.pow(10, Math.floor(Math.log10(max)));
    const norm = max / mag;
    let nice: number;
    if (norm <= 1) nice = 1;
    else if (norm <= 2) nice = 2;
    else if (norm <= 5) nice = 5;
    else nice = 10;
    return nice * mag;
}

function niceMin(min: number): number {
    if (min >= 0) return 0;
    // Mirror niceMax for negatives so 0 stays centred on the axis.
    const abs = Math.abs(min);
    return -niceMax(abs);
}

// Collect every numeric value across every series. Null entries are
// skipped; the range min/max is used to size the numeric axis.
function seriesRange(series: ChartSeries[]): { min: number; max: number } {
    let min = 0;
    let max = 0;
    let seen = false;
    for (const s of series) {
        for (const v of s.values) {
            if (v === null || !Number.isFinite(v)) continue;
            if (!seen) { min = v; max = v; seen = true; continue; }
            if (v < min) min = v;
            if (v > max) max = v;
        }
    }
    if (!seen) return { min: 0, max: 1 };
    // Never let the axis collapse to a single value.
    if (min === max) {
        if (max > 0) min = 0;
        else if (max < 0) max = 0;
        else { max = 1; }
    }
    return { min, max };
}

function baseLayout(width: number, height: number, hasLegend: boolean, legendPos: ChartModel['legend']): Layout {
    const top = 40;
    const right = 40 + (hasLegend && legendPos === 'right' ? 100 : 0);
    const bottom = 60 + (hasLegend && legendPos === 'bottom' ? 20 : 0);
    const left = 80 + (hasLegend && legendPos === 'left' ? 80 : 0);
    const adjustedTop = top + (hasLegend && legendPos === 'top' ? 20 : 0);
    return {
        width,
        height,
        plotX: left,
        plotY: adjustedTop,
        plotW: Math.max(10, width - left - right),
        plotH: Math.max(10, height - adjustedTop - bottom),
    };
}

// Title text at the top-centre.
function appendTitle(svg: SVGSVGElement, title: string, layout: Layout): void {
    const t = textEl(layout.width / 2, 20, title, {
        'text-anchor': 'middle',
        'font-size': '14',
        'font-weight': '600',
    });
    t.setAttribute('class', 'xlsx-chart-title');
    svg.appendChild(t);
}

// Legend row. Each entry is a <g> containing a swatch <rect> + a <text> with
// the series name. Positioned based on model.legend.
function appendLegend(svg: SVGSVGElement, model: ChartModel, layout: Layout): void {
    if (model.legend === 'none') return;
    const entries = model.series
        .map((s, i) => ({ name: s.name, color: colorFor(s, i) }))
        .filter((e) => e.name !== null);
    if (!entries.length) return;
    const group = document.createElementNS(SVG_NS, 'g');
    group.setAttribute('class', 'xlsx-chart-legend');

    let x: number;
    let y: number;
    let stepX = 0;
    let stepY = 0;
    const rowHeight = 16;
    const swatchSize = 10;
    const gap = 6;
    const entryGap = 14;

    if (model.legend === 'right') {
        x = layout.plotX + layout.plotW + 16;
        y = layout.plotY + 4;
        stepY = rowHeight;
    } else if (model.legend === 'left') {
        x = 10;
        y = layout.plotY + 4;
        stepY = rowHeight;
    } else if (model.legend === 'top') {
        x = layout.plotX;
        y = layout.plotY - 16;
        stepX = 80;
    } else {
        // bottom
        x = layout.plotX;
        y = layout.height - 16;
        stepX = 80;
    }

    entries.forEach((entry, i) => {
        const item = document.createElementNS(SVG_NS, 'g');
        item.setAttribute('class', 'xlsx-chart-legend-entry');
        const lx = x + stepX * i;
        const ly = y + stepY * i;
        const rect = el('rect', {
            x: lx, y: ly, width: swatchSize, height: swatchSize,
            fill: entry.color,
        });
        item.appendChild(rect);
        const label = textEl(lx + swatchSize + gap, ly + swatchSize - 1,
            entry.name ?? '', { 'font-size': '11' });
        item.appendChild(label);
        group.appendChild(item);
        void entryGap;
    });
    svg.appendChild(group);
}

// Draw a horizontal gridline + tick label for each nice value between
// axisMin and axisMax. Used by bar/column/line; pie skips axes entirely.
function appendYAxis(svg: SVGSVGElement, layout: Layout, axisMin: number, axisMax: number): void {
    const g = document.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'xlsx-chart-yaxis');
    const ticks = 5;
    for (let i = 0; i <= ticks; i++) {
        const frac = i / ticks;
        const value = axisMin + (axisMax - axisMin) * frac;
        const y = layout.plotY + layout.plotH - frac * layout.plotH;
        g.appendChild(el('line', {
            x1: layout.plotX, x2: layout.plotX + layout.plotW,
            y1: y, y2: y,
            stroke: GRID_COLOR, 'stroke-width': 1,
        }));
        g.appendChild(textEl(layout.plotX - 6, y + 4, formatTick(value), {
            'text-anchor': 'end', 'font-size': '10',
        }));
    }
    // Axis line.
    g.appendChild(el('line', {
        x1: layout.plotX, x2: layout.plotX,
        y1: layout.plotY, y2: layout.plotY + layout.plotH,
        stroke: AXIS_COLOR, 'stroke-width': 1,
    }));
    svg.appendChild(g);
}

function appendXAxisCategory(svg: SVGSVGElement, layout: Layout, categories: string[]): void {
    if (!categories.length) return;
    const g = document.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'xlsx-chart-xaxis');
    const slot = layout.plotW / categories.length;
    categories.forEach((label, i) => {
        const cx = layout.plotX + slot * (i + 0.5);
        g.appendChild(textEl(cx, layout.plotY + layout.plotH + 14, label ?? '', {
            'text-anchor': 'middle', 'font-size': '10',
        }));
    });
    // Baseline.
    g.appendChild(el('line', {
        x1: layout.plotX, x2: layout.plotX + layout.plotW,
        y1: layout.plotY + layout.plotH, y2: layout.plotY + layout.plotH,
        stroke: AXIS_COLOR, 'stroke-width': 1,
    }));
    svg.appendChild(g);
}

function formatTick(v: number): string {
    if (!Number.isFinite(v)) return '';
    // Keep ticks compact: one decimal max, trailing zeros trimmed.
    const abs = Math.abs(v);
    if (abs >= 1000) return String(Math.round(v));
    if (abs >= 10 || abs === 0) return String(Math.round(v));
    const s = v.toFixed(1);
    return s.endsWith('.0') ? s.slice(0, -2) : s;
}

// Column chart (vertical bars). Clustered by default; `grouping='stacked'`
// stacks vertically per category; `'percentStacked'` normalises each
// category to 100%.
function renderColumn(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
    if (n === 0) return;
    const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const percent = model.grouping === 'percentStacked';

    const { min, max } = stacked
        ? stackedRange(model.series, n, percent)
        : seriesRange(model.series);
    const axisMax = max > 0 ? niceMax(max) : 0;
    const axisMin = min < 0 ? niceMin(min) : 0;
    appendYAxis(svg, layout, axisMin, axisMax);

    const categoryLabels = model.categories.length ? model.categories
        : Array.from({ length: n }, (_, i) => String(i + 1));
    appendXAxisCategory(svg, layout, categoryLabels);

    const slot = layout.plotW / n;
    const groupPad = slot * 0.2;
    const seriesCount = Math.max(1, model.series.length);
    const barW = stacked ? slot - groupPad * 2 : (slot - groupPad * 2) / seriesCount;
    const range = axisMax - axisMin || 1;
    const zeroY = layout.plotY + layout.plotH * (axisMax / range);

    // Pre-compute per-category totals (for percentStacked normalisation) and
    // running stack tops for stacked layouts. Both are keyed by category
    // index; we advance during the series loop below.
    const totals: number[] = percent ? categoryTotals(model.series, n) : [];
    const stackPos: number[] = new Array(n).fill(0);
    const stackNeg: number[] = new Array(n).fill(0);

    model.series.forEach((s, si) => {
        const group = document.createElementNS(SVG_NS, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', String(si));
        if (s.name) group.setAttribute('data-series-name', sanitiseForAttr(s.name));
        const color = colorFor(s, si);
        for (let ci = 0; ci < n; ci++) {
            const vRaw = s.values[ci];
            const vRawFinite = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
            const x = stacked
                ? layout.plotX + slot * ci + groupPad
                : layout.plotX + slot * ci + groupPad + barW * si;
            let y: number;
            let h: number;
            if (vRawFinite === null) {
                y = zeroY;
                h = 0;
            } else if (stacked) {
                // Stacked column: bars run from the current stack top
                // upwards. Percent normalises each slice to its share of
                // the category total. Negative values stack downwards
                // separately so the axis doesn't collapse.
                const categoryTotal = totals[ci] ?? 0;
                const vShare = percent
                    ? (categoryTotal > 0 ? (vRawFinite / categoryTotal) * 100 : 0)
                    : vRawFinite;
                if (vShare >= 0) {
                    const base = stackPos[ci];
                    const top = base + vShare;
                    const yTop = layout.plotY + layout.plotH - ((top - axisMin) / range) * layout.plotH;
                    const yBase = layout.plotY + layout.plotH - ((base - axisMin) / range) * layout.plotH;
                    y = yTop;
                    h = yBase - yTop;
                    stackPos[ci] = top;
                } else {
                    const base = stackNeg[ci];
                    const bottom = base + vShare;
                    const yTop = layout.plotY + layout.plotH - ((base - axisMin) / range) * layout.plotH;
                    const yBot = layout.plotY + layout.plotH - ((bottom - axisMin) / range) * layout.plotH;
                    y = yTop;
                    h = yBot - yTop;
                    stackNeg[ci] = bottom;
                }
            } else if (vRawFinite >= 0) {
                const top = layout.plotY + layout.plotH - (vRawFinite / range) * layout.plotH - (axisMin < 0 ? -axisMin / range * layout.plotH : 0);
                y = top;
                h = zeroY - top;
            } else {
                const bottom = layout.plotY + layout.plotH - ((vRawFinite - axisMin) / range) * layout.plotH;
                y = zeroY;
                h = bottom - zeroY;
            }
            group.appendChild(el('rect', {
                x, y, width: Math.max(0, barW - 1), height: Math.max(0, h),
                fill: color,
            }));
        }
        svg.appendChild(group);
    });
}

// Sum each category's positive / negative slices independently and return
// the extremes. Used by stacked bar / column / area to size the axis.
function stackedRange(series: ChartSeries[], n: number, percent: boolean): { min: number; max: number } {
    if (percent) return { min: 0, max: 100 };
    let max = 0;
    let min = 0;
    for (let ci = 0; ci < n; ci++) {
        let pos = 0, neg = 0;
        for (const s of series) {
            const v = s.values[ci];
            if (v === null || v === undefined || !Number.isFinite(v)) continue;
            if (v >= 0) pos += v;
            else neg += v;
        }
        if (pos > max) max = pos;
        if (neg < min) min = neg;
    }
    if (max === 0 && min === 0) max = 1;
    return { min, max };
}

function categoryTotals(series: ChartSeries[], n: number): number[] {
    const out: number[] = new Array(n).fill(0);
    for (let ci = 0; ci < n; ci++) {
        for (const s of series) {
            const v = s.values[ci];
            if (v === null || v === undefined || !Number.isFinite(v)) continue;
            if (v > 0) out[ci] += v; // percent normalisation against positive-side total
        }
    }
    return out;
}

// Horizontal bar chart — flip axes. Categories along the y-axis, values
// along the x-axis.
function renderBar(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
    if (n === 0) return;
    const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const percent = model.grouping === 'percentStacked';
    const { min, max } = stacked
        ? stackedRange(model.series, n, percent)
        : seriesRange(model.series);
    const axisMax = max > 0 ? niceMax(max) : 0;
    const axisMin = min < 0 ? niceMin(min) : 0;

    // X-axis becomes numeric; Y-axis becomes categorical.
    const gx = document.createElementNS(SVG_NS, 'g');
    gx.setAttribute('class', 'xlsx-chart-xaxis');
    const ticks = 5;
    for (let i = 0; i <= ticks; i++) {
        const frac = i / ticks;
        const value = axisMin + (axisMax - axisMin) * frac;
        const x = layout.plotX + frac * layout.plotW;
        gx.appendChild(el('line', {
            x1: x, x2: x, y1: layout.plotY, y2: layout.plotY + layout.plotH,
            stroke: GRID_COLOR, 'stroke-width': 1,
        }));
        gx.appendChild(textEl(x, layout.plotY + layout.plotH + 14, formatTick(value), {
            'text-anchor': 'middle', 'font-size': '10',
        }));
    }
    gx.appendChild(el('line', {
        x1: layout.plotX, x2: layout.plotX + layout.plotW,
        y1: layout.plotY + layout.plotH, y2: layout.plotY + layout.plotH,
        stroke: AXIS_COLOR, 'stroke-width': 1,
    }));
    svg.appendChild(gx);

    const categoryLabels = model.categories.length ? model.categories
        : Array.from({ length: n }, (_, i) => String(i + 1));
    const gy = document.createElementNS(SVG_NS, 'g');
    gy.setAttribute('class', 'xlsx-chart-yaxis');
    const slot = layout.plotH / n;
    categoryLabels.forEach((label, i) => {
        const cy = layout.plotY + slot * (i + 0.5);
        gy.appendChild(textEl(layout.plotX - 6, cy + 4, label ?? '', {
            'text-anchor': 'end', 'font-size': '10',
        }));
    });
    gy.appendChild(el('line', {
        x1: layout.plotX, x2: layout.plotX,
        y1: layout.plotY, y2: layout.plotY + layout.plotH,
        stroke: AXIS_COLOR, 'stroke-width': 1,
    }));
    svg.appendChild(gy);

    const groupPad = slot * 0.2;
    const seriesCount = Math.max(1, model.series.length);
    const barH = stacked ? slot - groupPad * 2 : (slot - groupPad * 2) / seriesCount;
    const range = axisMax - axisMin || 1;
    const zeroX = layout.plotX + ((0 - axisMin) / range) * layout.plotW;

    const totals: number[] = percent ? categoryTotals(model.series, n) : [];
    const stackPos: number[] = new Array(n).fill(0);
    const stackNeg: number[] = new Array(n).fill(0);

    model.series.forEach((s, si) => {
        const group = document.createElementNS(SVG_NS, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', String(si));
        if (s.name) group.setAttribute('data-series-name', sanitiseForAttr(s.name));
        const color = colorFor(s, si);
        for (let ci = 0; ci < n; ci++) {
            const vRaw = s.values[ci];
            const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
            const y = stacked
                ? layout.plotY + slot * ci + groupPad
                : layout.plotY + slot * ci + groupPad + barH * si;
            let x: number;
            let w: number;
            if (v === null) {
                x = zeroX; w = 0;
            } else if (stacked) {
                const categoryTotal = totals[ci] ?? 0;
                const vShare = percent
                    ? (categoryTotal > 0 ? (v / categoryTotal) * 100 : 0)
                    : v;
                if (vShare >= 0) {
                    const base = stackPos[ci];
                    const top = base + vShare;
                    const xStart = layout.plotX + ((base - axisMin) / range) * layout.plotW;
                    const xEnd = layout.plotX + ((top - axisMin) / range) * layout.plotW;
                    x = xStart;
                    w = xEnd - xStart;
                    stackPos[ci] = top;
                } else {
                    const base = stackNeg[ci];
                    const bottom = base + vShare;
                    const xStart = layout.plotX + ((bottom - axisMin) / range) * layout.plotW;
                    const xEnd = layout.plotX + ((base - axisMin) / range) * layout.plotW;
                    x = xStart;
                    w = xEnd - xStart;
                    stackNeg[ci] = bottom;
                }
            } else if (v >= 0) {
                x = zeroX;
                w = (v / range) * layout.plotW;
            } else {
                const end = layout.plotX + ((v - axisMin) / range) * layout.plotW;
                x = end;
                w = zeroX - end;
            }
            group.appendChild(el('rect', {
                x, y, width: Math.max(0, w), height: Math.max(0, barH - 1),
                fill: color,
            }));
        }
        svg.appendChild(group);
    });
}

// Line chart. One polyline per series with circle markers per point.
function renderLine(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
    if (n === 0) return;
    const { min, max } = seriesRange(model.series);
    const axisMax = max > 0 ? niceMax(max) : 0;
    const axisMin = min < 0 ? niceMin(min) : 0;
    appendYAxis(svg, layout, axisMin, axisMax);
    const categoryLabels = model.categories.length ? model.categories
        : Array.from({ length: n }, (_, i) => String(i + 1));
    appendXAxisCategory(svg, layout, categoryLabels);

    const slot = n > 1 ? layout.plotW / (n - 1) : layout.plotW;

    const pointX = (ci: number) => n > 1
        ? layout.plotX + slot * ci
        : layout.plotX + layout.plotW / 2;
    const pointY = (v: number) => layout.plotY + layout.plotH - ((v - axisMin) / (axisMax - axisMin || 1)) * layout.plotH;

    model.series.forEach((s, si) => {
        const group = document.createElementNS(SVG_NS, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', String(si));
        if (s.name) group.setAttribute('data-series-name', sanitiseForAttr(s.name));
        const color = colorFor(s, si);
        const pts: string[] = [];
        for (let ci = 0; ci < n; ci++) {
            const vRaw = s.values[ci];
            const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
            if (v === null) continue;
            const px = pointX(ci);
            const py = pointY(v);
            pts.push(`${px.toFixed(2)},${py.toFixed(2)}`);
        }
        if (pts.length) {
            group.appendChild(el('polyline', {
                points: pts.join(' '),
                fill: 'none',
                stroke: color,
                'stroke-width': 2,
            }));
            // Markers.
            for (let ci = 0; ci < n; ci++) {
                const vRaw = s.values[ci];
                const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
                if (v === null) continue;
                group.appendChild(el('circle', {
                    cx: pointX(ci), cy: pointY(v), r: 3,
                    fill: color,
                }));
            }
        }
        svg.appendChild(group);
    });
}

// Scatter chart. Both axes are numeric — series carry parallel
// (xValues, values) streams. Every series lands as its own group of
// `<circle>` markers; no polyline is drawn (smoothing paths are a future
// extension). The x-axis mirrors the y-axis in rendering: nice-round ticks
// derived from the combined min/max of every declared xValue.
function renderScatter(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const xRange = scatterXRange(model.series);
    const { min, max } = seriesRange(model.series);
    const axisMaxY = max > 0 ? niceMax(max) : 0;
    const axisMinY = min < 0 ? niceMin(min) : 0;
    const axisMinX = xRange.min < 0 ? niceMin(xRange.min) : (xRange.min === xRange.max ? 0 : Math.min(0, xRange.min));
    const axisMaxX = xRange.max > 0 ? niceMax(xRange.max) : xRange.max;
    const rangeY = axisMaxY - axisMinY || 1;
    const rangeX = axisMaxX - axisMinX || 1;
    appendYAxis(svg, layout, axisMinY, axisMaxY);
    appendXAxisNumeric(svg, layout, axisMinX, axisMaxX);

    const projX = (v: number) => layout.plotX + ((v - axisMinX) / rangeX) * layout.plotW;
    const projY = (v: number) => layout.plotY + layout.plotH - ((v - axisMinY) / rangeY) * layout.plotH;

    model.series.forEach((s, si) => {
        const group = document.createElementNS(SVG_NS, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', String(si));
        if (s.name) group.setAttribute('data-series-name', sanitiseForAttr(s.name));
        const color = colorFor(s, si);
        const xs = s.xValues ?? [];
        const ys = s.values;
        const count = Math.max(xs.length, ys.length);
        for (let i = 0; i < count; i++) {
            const xv = xs[i];
            const yv = ys[i];
            if (xv === null || xv === undefined || !Number.isFinite(xv)) continue;
            if (yv === null || yv === undefined || !Number.isFinite(yv)) continue;
            group.appendChild(el('circle', {
                cx: projX(xv).toFixed(2),
                cy: projY(yv).toFixed(2),
                r: 4,
                fill: color,
                stroke: '#fff',
                'stroke-width': 1,
            }));
        }
        svg.appendChild(group);
    });
}

// Collect every numeric x across every series. Mirrors `seriesRange` for
// the y-axis but walks `xValues`. Empty streams fall back to {0, 1}.
function scatterXRange(series: ChartSeries[]): { min: number; max: number } {
    let min = 0, max = 0, seen = false;
    for (const s of series) {
        const xs = s.xValues;
        if (!xs) continue;
        for (const v of xs) {
            if (v === null || !Number.isFinite(v)) continue;
            if (!seen) { min = v; max = v; seen = true; continue; }
            if (v < min) min = v;
            if (v > max) max = v;
        }
    }
    if (!seen) return { min: 0, max: 1 };
    if (min === max) {
        if (max > 0) min = 0;
        else if (max < 0) max = 0;
        else { max = 1; }
    }
    return { min, max };
}

// Numeric x-axis (used by scatter). Mirrors appendXAxisCategory's look but
// draws equally-spaced nice-round tick labels.
function appendXAxisNumeric(svg: SVGSVGElement, layout: Layout, axisMin: number, axisMax: number): void {
    const g = document.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'xlsx-chart-xaxis');
    const ticks = 5;
    for (let i = 0; i <= ticks; i++) {
        const frac = i / ticks;
        const value = axisMin + (axisMax - axisMin) * frac;
        const x = layout.plotX + frac * layout.plotW;
        g.appendChild(el('line', {
            x1: x, x2: x, y1: layout.plotY + layout.plotH, y2: layout.plotY + layout.plotH + 4,
            stroke: AXIS_COLOR, 'stroke-width': 1,
        }));
        g.appendChild(textEl(x, layout.plotY + layout.plotH + 14, formatTick(value), {
            'text-anchor': 'middle', 'font-size': '10',
        }));
    }
    g.appendChild(el('line', {
        x1: layout.plotX, x2: layout.plotX + layout.plotW,
        y1: layout.plotY + layout.plotH, y2: layout.plotY + layout.plotH,
        stroke: AXIS_COLOR, 'stroke-width': 1,
    }));
    svg.appendChild(g);
}

// Area chart. One `<path>` per series with a closed shape: top follows the
// series values, bottom returns along the baseline (or the previous stack
// top in stacked / percentStacked modes). Fill at 70% opacity, outline at
// full opacity. Stacked mode accumulates running totals per category so
// series paint above each other without overlap; percentStacked normalises
// each category to 100%.
function renderArea(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
    if (n === 0) return;
    const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const percent = model.grouping === 'percentStacked';
    const { min, max } = stacked
        ? stackedRange(model.series, n, percent)
        : seriesRange(model.series);
    const axisMax = max > 0 ? niceMax(max) : 0;
    const axisMin = min < 0 ? niceMin(min) : 0;
    const range = axisMax - axisMin || 1;
    appendYAxis(svg, layout, axisMin, axisMax);
    const categoryLabels = model.categories.length ? model.categories
        : Array.from({ length: n }, (_, i) => String(i + 1));
    appendXAxisCategory(svg, layout, categoryLabels);

    const slot = n > 1 ? layout.plotW / (n - 1) : layout.plotW;
    const pointX = (ci: number) => n > 1
        ? layout.plotX + slot * ci
        : layout.plotX + layout.plotW / 2;
    const yFor = (v: number) => layout.plotY + layout.plotH - ((v - axisMin) / range) * layout.plotH;
    const baselineY = yFor(0);

    const totals = percent ? categoryTotals(model.series, n) : [];
    // Running stack tops per category. Starts at 0 (axis baseline); we
    // mutate as each series paints. Non-stacked mode ignores these.
    const prevTop: number[] = new Array(n).fill(0);

    model.series.forEach((s, si) => {
        const group = document.createElementNS(SVG_NS, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', String(si));
        if (s.name) group.setAttribute('data-series-name', sanitiseForAttr(s.name));
        const color = colorFor(s, si);

        // Top edge: walk categories and emit Move / Line commands.
        // Bottom edge: walk back along the baseline (or previous stack) and
        // close with Z so the shape paints filled.
        const topPts: { x: number; y: number }[] = [];
        const bottomPts: { x: number; y: number }[] = [];
        for (let ci = 0; ci < n; ci++) {
            const vRaw = s.values[ci];
            const vFinite = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : 0;
            const x = pointX(ci);
            if (stacked) {
                const share = percent
                    ? (totals[ci] > 0 ? (vFinite / totals[ci]) * 100 : 0)
                    : vFinite;
                const baseValue = prevTop[ci];
                const topValue = baseValue + share;
                topPts.push({ x, y: yFor(topValue) });
                bottomPts.push({ x, y: yFor(baseValue) });
                prevTop[ci] = topValue;
            } else {
                topPts.push({ x, y: yFor(vFinite) });
                bottomPts.push({ x, y: baselineY });
            }
        }

        // Build the d-string: M first-top → L ... → L last-bottom →
        // L ... (reverse) → Z. We spell out each point explicitly so the
        // fill shape is closed and the outline trace matches Excel's look.
        const parts: string[] = [];
        topPts.forEach((p, i) => {
            parts.push(`${i === 0 ? 'M' : 'L'} ${p.x.toFixed(2)} ${p.y.toFixed(2)}`);
        });
        for (let i = bottomPts.length - 1; i >= 0; i--) {
            const p = bottomPts[i];
            parts.push(`L ${p.x.toFixed(2)} ${p.y.toFixed(2)}`);
        }
        parts.push('Z');
        const d = parts.join(' ');

        group.appendChild(el('path', {
            d,
            fill: color,
            'fill-opacity': 0.7,
            stroke: color,
            'stroke-width': 1.5,
            'stroke-opacity': 1,
        }));
        svg.appendChild(group);
    });
}

// Pie chart. First series only. One <path> per slice, built with the
// standard M (centre) → L (first vertex) → A (arc) → Z (close) command
// sequence. Legend sits alongside using the category labels as entry names.
function renderPie(model: ChartModel, layout: Layout, svg: SVGSVGElement): void {
    const series = model.series[0];
    if (!series) return;
    const values = series.values
        .map((v) => (v !== null && Number.isFinite(v) && v > 0 ? v : 0));
    const total = values.reduce((a, b) => a + b, 0);
    if (total <= 0) return;

    // Pick the largest square that fits inside the plot area.
    const size = Math.min(layout.plotW, layout.plotH);
    const cx = layout.plotX + layout.plotW / 2;
    const cy = layout.plotY + layout.plotH / 2;
    const r = size / 2 - 4;

    const group = document.createElementNS(SVG_NS, 'g');
    group.setAttribute('class', 'xlsx-chart-series');
    group.setAttribute('data-series-index', '0');
    if (series.name) group.setAttribute('data-series-name', sanitiseForAttr(series.name));

    let acc = 0;
    values.forEach((v, i) => {
        if (v <= 0) {
            // Even zero slices still emit a path so the slice count matches
            // the category count — the harness asserts this. Degenerate
            // paths are kept invisible by matching start + end points.
            const path = el('path', {
                d: `M ${cx} ${cy} Z`,
                fill: PALETTE[i % PALETTE.length],
            });
            group.appendChild(path);
            return;
        }
        const startAngle = (acc / total) * Math.PI * 2 - Math.PI / 2;
        acc += v;
        const endAngle = (acc / total) * Math.PI * 2 - Math.PI / 2;
        const x1 = cx + r * Math.cos(startAngle);
        const y1 = cy + r * Math.sin(startAngle);
        const x2 = cx + r * Math.cos(endAngle);
        const y2 = cy + r * Math.sin(endAngle);
        const large = endAngle - startAngle > Math.PI ? 1 : 0;
        const d = `M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${large} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
        const color = PALETTE[i % PALETTE.length];
        const path = el('path', {
            d,
            fill: color,
            stroke: '#fff',
            'stroke-width': 1,
        });
        group.appendChild(path);
    });
    svg.appendChild(group);

    // Category legend — for pie charts the categories are the "series names"
    // from the viewer's perspective, so we fabricate a legend of
    // {swatch, category}.
    if (model.legend !== 'none' && model.categories.length) {
        const legend = document.createElementNS(SVG_NS, 'g');
        legend.setAttribute('class', 'xlsx-chart-legend');
        const x = layout.plotX + layout.plotW + 16;
        const y = layout.plotY + 4;
        const rowH = 16;
        const swatch = 10;
        model.categories.forEach((label, i) => {
            const item = document.createElementNS(SVG_NS, 'g');
            item.setAttribute('class', 'xlsx-chart-legend-entry');
            const ly = y + rowH * i;
            item.appendChild(el('rect', {
                x, y: ly, width: swatch, height: swatch,
                fill: PALETTE[i % PALETTE.length],
            }));
            item.appendChild(textEl(x + swatch + 6, ly + swatch - 1, label ?? '', {
                'font-size': '11',
            }));
            legend.appendChild(item);
        });
        svg.appendChild(legend);
    }
}

// Strip control characters and cap length for data-* attribute safety.
// The string will be URL-encoded by setAttribute anyway — this is a
// belt-and-braces guard so attacker-controlled series names can't smuggle
// newlines / nulls through the attribute edge.
function sanitiseForAttr(s: string): string {
    // Strip ASCII control characters (0x00-0x1F and 0x7F). The cap is a
    // conservative defensive bound; series names should be short in
    // practice. Writing this as an explicit loop keeps the source free of
    // literal control bytes and keeps the regex engine out of a hot path.
    let out = '';
    for (let i = 0; i < s.length; i++) {
        const c = s.charCodeAt(i);
        if (c < 0x20 || c === 0x7F) continue;
        out += s[i];
    }
    return out.slice(0, 256);
}

export function renderChart(model: ChartModel, width: number = 480, height: number = 300): SVGSVGElement | null {
    if (model.kind === 'unknown') return null;
    const svg = makeSvg(width, height);
    const hasLegend = model.legend !== 'none' && model.series.some((s) => s.name);
    const legendPos = model.legend;
    const layout = baseLayout(width, height, hasLegend, legendPos);

    if (model.title) appendTitle(svg, model.title, layout);

    switch (model.kind) {
        case 'column':
            renderColumn(model, layout, svg);
            appendLegend(svg, model, layout);
            return svg;
        case 'bar':
            renderBar(model, layout, svg);
            appendLegend(svg, model, layout);
            return svg;
        case 'line':
            renderLine(model, layout, svg);
            appendLegend(svg, model, layout);
            return svg;
        case 'pie':
            renderPie(model, layout, svg);
            return svg;
        case 'scatter':
            renderScatter(model, layout, svg);
            appendLegend(svg, model, layout);
            return svg;
        case 'area':
            renderArea(model, layout, svg);
            appendLegend(svg, model, layout);
            return svg;
        default:
            return null;
    }
}
