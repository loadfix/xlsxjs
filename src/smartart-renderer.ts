// SmartArtModel → inline SVG. Wave 10 shipped a single top-down hierarchy
// layout; Wave 11 adds two more layout strategies so the common diagram
// variants render as something closer to what Excel paints:
//
//   · 'hierarchy' — every depth level lays out horizontally across the
//     viewport, parent nodes connect to children via straight diagonal
//     lines. This is the pre-existing default.
//   · 'orgchart'  — same top-down positioning, but connectors draw as a
//     branch rail: parent-drop vertical, horizontal rail spanning children,
//     then a short vertical per child down to the child's top-centre. For
//     a single-child parent the rail collapses to a straight vertical line.
//   · 'cycle'     — nodes arrange around a circle, connected clockwise by
//     arc paths with a shared arrow marker so the flow closes back to the
//     first node.
//
// Dispatch is `opts.layout` first (explicit override), then a regex match
// against `model.layout` uniqueId (…/orgChart/… → orgchart,
// …/cycle/… → cycle), falling through to the hierarchy default.
//
// Security contract:
//   · Node text reaches the DOM via `textContent` on `<text>` elements. We
//     never set innerHTML and never interpolate XLSX-derived strings into
//     attributes.
//   · Colours come from a hard-coded palette; numeric coordinates are
//     computed from the tree shape (depth count + siblings-at-depth), so
//     no attacker-controlled data can escape into an SVG attribute.
//   · The rendered `<svg>` is a detached element — callers attach it into
//     the DOM wherever they want (see `renderSmartArt` for the aside-first
//     wiring). Returning `null` for an empty tree lets callers skip the
//     append entirely.

import type { SmartArtModel, SmartArtNode } from './smartart-parser';

const SVG_NS = 'http://www.w3.org/2000/svg';

const DEFAULT_WIDTH = 480;
const DEFAULT_HEIGHT = 320;
const NODE_WIDTH = 100;
const NODE_HEIGHT = 40;
const NODE_RX = 4;
const MARGIN_X = 12;
const MARGIN_Y = 16;

// Cycle-specific node dimensions: smaller than hierarchy so all nodes fit
// comfortably around a circle in the default viewport.
const CYCLE_NODE_WIDTH = 80;
const CYCLE_NODE_HEIGHT = 32;
const CYCLE_NODE_RX = 4;

// Depth-indexed palette. Mirrors the chart-renderer Office-accent palette but
// capitalised to match the spec in CLAUDE's task block. The text colour is
// intentionally white so nodes stay legible against the coloured fill.
const PALETTE = [
    '#5B9BD5',
    '#ED7D31',
    '#A5A5A5',
    '#FFC000',
    '#4472C4',
    '#70AD47',
] as const;

const CONNECTOR_COLOR = '#888';
const TEXT_COLOR = '#ffffff';

export type SmartArtLayoutStrategy = 'hierarchy' | 'orgchart' | 'cycle' | 'auto';

export interface RenderSmartArtOptions {
    width?: number;
    height?: number;
    // Explicit layout override. Defaults to 'auto' — dispatch off
    // `model.layout` uniqueId. Pass a concrete strategy to force a given
    // layout regardless of what the diagram announced.
    layout?: SmartArtLayoutStrategy;
}

/**
 * Render a SmartArt diagram as an inline SVG. Returns `null` when the
 * parsed model has no root nodes — callers can branch on that to fall back
 * to the `<ul>` tree emitted by Wave 9 renderer.
 *
 * Dispatch:
 *   · `opts.layout` overrides everything, except `'auto'` which defers to
 *     the model's announced layout.
 *   · `model.layout` uniqueId matching `/orgChart/i` → orgchart strategy.
 *   · `model.layout` uniqueId matching `/cycle/i`   → cycle strategy.
 *   · everything else → hierarchy (default) strategy.
 */
export function renderSmartArtSvg(
    model: SmartArtModel,
    opts: RenderSmartArtOptions = {},
): SVGSVGElement | null {
    const roots = model.rootNodes;
    if (!roots || roots.length === 0) return null;

    const width = opts.width ?? DEFAULT_WIDTH;
    const height = opts.height ?? DEFAULT_HEIGHT;

    const strategy = resolveStrategy(model.layout, opts.layout ?? 'auto');

    if (strategy === 'cycle') {
        return renderCycle(model, width, height);
    }
    // Both hierarchy and orgchart share the level-based placement; the only
    // difference is how connectors are drawn.
    return renderHierarchyOrOrgchart(model, width, height, strategy === 'orgchart');
}

// Resolve the layout strategy. Explicit `opts.layout` wins unless it's
// 'auto'; otherwise we dispatch on `model.layout` uniqueId via a case-
// insensitive regex. Unknown layouts fall through to hierarchy.
function resolveStrategy(
    layoutName: string | null,
    override: SmartArtLayoutStrategy,
): 'hierarchy' | 'orgchart' | 'cycle' {
    if (override === 'hierarchy' || override === 'orgchart' || override === 'cycle') {
        return override;
    }
    if (layoutName) {
        if (/orgChart/i.test(layoutName)) return 'orgchart';
        if (/cycle/i.test(layoutName)) return 'cycle';
    }
    return 'hierarchy';
}

// Hierarchy + orgchart share node placement: each depth level lays out
// horizontally, evenly spaced. The `orgchart` flag only switches connector
// drawing between the straight-diagonal line (hierarchy) and the branch-rail
// form (orgchart).
function renderHierarchyOrOrgchart(
    model: SmartArtModel,
    width: number,
    height: number,
    orgchart: boolean,
): SVGSVGElement {
    interface Placed {
        node: SmartArtNode;
        parentIndex: number; // -1 for roots
        x: number;
        y: number;
        depth: number;
    }
    const levels: Placed[][] = [];
    function walk(node: SmartArtNode, depth: number, parentIndex: number): void {
        if (!levels[depth]) levels[depth] = [];
        const placed: Placed = { node, parentIndex, x: 0, y: 0, depth };
        levels[depth].push(placed);
        const myIndex = levels[depth].length - 1;
        for (const child of node.children) {
            walk(child, depth + 1, myIndex);
        }
    }
    for (const root of model.rootNodes) walk(root, 0, -1);

    const depthCount = levels.length;
    const usableH = Math.max(NODE_HEIGHT, height - MARGIN_Y * 2);
    for (let d = 0; d < depthCount; d++) {
        const row = levels[d];
        const yCentre = depthCount === 1
            ? height / 2
            : MARGIN_Y + NODE_HEIGHT / 2 + (usableH - NODE_HEIGHT) * (d / (depthCount - 1));
        const usableW = Math.max(NODE_WIDTH, width - MARGIN_X * 2);
        const count = row.length;
        for (let i = 0; i < count; i++) {
            const xCentre = count === 1
                ? width / 2
                : MARGIN_X + NODE_WIDTH / 2 + (usableW - NODE_WIDTH) * (i / (count - 1));
            row[i].x = xCentre;
            row[i].y = yCentre;
        }
    }

    const svg = createSvgRoot(width, height);

    if (orgchart) {
        // Orgchart branch-rail connectors. Group children by their parent
        // index (on the level above), then for each non-empty group:
        //   · drop a short vertical from parent-bottom-centre to the rail-y
        //   · draw a horizontal rail spanning the left-most to right-most
        //     child x-centre (skip when only one child — rail is zero-width
        //     and the straight vertical covers the connection)
        //   · drop a short vertical per child from the rail to the child's
        //     top-centre
        for (let d = 1; d < depthCount; d++) {
            const row = levels[d];
            const parents = levels[d - 1];
            // Group children by parent index.
            const byParent = new Map<number, Placed[]>();
            for (const placed of row) {
                const list = byParent.get(placed.parentIndex) ?? [];
                list.push(placed);
                byParent.set(placed.parentIndex, list);
            }
            for (const [parentIndex, group] of byParent) {
                const parent = parents[parentIndex];
                if (!parent) continue;
                const parentBottom = parent.y + NODE_HEIGHT / 2;
                const childTop = group[0].y - NODE_HEIGHT / 2;
                // Halfway rail y between parent-bottom and child-top.
                const railY = (parentBottom + childTop) / 2;

                if (group.length === 1) {
                    // Single child: just a straight vertical line from
                    // parent-bottom-centre to child-top-centre.
                    const line = document.createElementNS(SVG_NS, 'line');
                    line.setAttribute('x1', String(parent.x));
                    line.setAttribute('y1', String(parentBottom));
                    line.setAttribute('x2', String(group[0].x));
                    line.setAttribute('y2', String(childTop));
                    line.setAttribute('stroke', CONNECTOR_COLOR);
                    line.setAttribute('stroke-width', '1');
                    svg.appendChild(line);
                    continue;
                }

                // Parent drop.
                const drop = document.createElementNS(SVG_NS, 'line');
                drop.setAttribute('x1', String(parent.x));
                drop.setAttribute('y1', String(parentBottom));
                drop.setAttribute('x2', String(parent.x));
                drop.setAttribute('y2', String(railY));
                drop.setAttribute('stroke', CONNECTOR_COLOR);
                drop.setAttribute('stroke-width', '1');
                svg.appendChild(drop);

                // Horizontal rail across children.
                let minX = group[0].x;
                let maxX = group[0].x;
                for (const c of group) {
                    if (c.x < minX) minX = c.x;
                    if (c.x > maxX) maxX = c.x;
                }
                const rail = document.createElementNS(SVG_NS, 'line');
                rail.setAttribute('x1', String(minX));
                rail.setAttribute('y1', String(railY));
                rail.setAttribute('x2', String(maxX));
                rail.setAttribute('y2', String(railY));
                rail.setAttribute('stroke', CONNECTOR_COLOR);
                rail.setAttribute('stroke-width', '1');
                svg.appendChild(rail);

                // Per-child drop from rail to child top.
                for (const c of group) {
                    const leg = document.createElementNS(SVG_NS, 'line');
                    leg.setAttribute('x1', String(c.x));
                    leg.setAttribute('y1', String(railY));
                    leg.setAttribute('x2', String(c.x));
                    leg.setAttribute('y2', String(c.y - NODE_HEIGHT / 2));
                    leg.setAttribute('stroke', CONNECTOR_COLOR);
                    leg.setAttribute('stroke-width', '1');
                    svg.appendChild(leg);
                }
            }
        }
    } else {
        // Hierarchy: straight diagonal lines from parent-bottom-centre to
        // child-top-centre. This is the pre-existing behaviour and remains
        // byte-stable for the hierarchy fixture's golden snapshot.
        for (let d = 1; d < depthCount; d++) {
            const row = levels[d];
            const parents = levels[d - 1];
            for (const placed of row) {
                const parent = parents[placed.parentIndex];
                if (!parent) continue;
                const line = document.createElementNS(SVG_NS, 'line');
                line.setAttribute('x1', String(parent.x));
                line.setAttribute('y1', String(parent.y + NODE_HEIGHT / 2));
                line.setAttribute('x2', String(placed.x));
                line.setAttribute('y2', String(placed.y - NODE_HEIGHT / 2));
                line.setAttribute('stroke', CONNECTOR_COLOR);
                line.setAttribute('stroke-width', '1');
                svg.appendChild(line);
            }
        }
    }

    // Draw nodes on top.
    for (let d = 0; d < depthCount; d++) {
        const row = levels[d];
        const fill = PALETTE[d % PALETTE.length];
        for (const placed of row) {
            appendNode(svg, placed.x, placed.y, NODE_WIDTH, NODE_HEIGHT, NODE_RX, fill, placed.node.text);
        }
    }

    return svg;
}

// Cycle layout: place a set of nodes evenly around a circle, connect them
// clockwise with arc paths, and wrap the last one back to the first.
//
// If `rootNodes.length > 1`, use the roots directly. Otherwise (the usual
// "one doc root → 4 steps" shape) use `rootNodes[0].children`. If neither
// yields at least two nodes, fall through to a minimal one-node rendering
// so consumers don't get an unexpectedly empty SVG.
function renderCycle(model: SmartArtModel, width: number, height: number): SVGSVGElement {
    const roots = model.rootNodes;
    let nodes: SmartArtNode[] = [];
    if (roots.length > 1) {
        nodes = roots;
    } else if (roots.length === 1) {
        nodes = roots[0].children.length > 0 ? roots[0].children : roots;
    }

    const svg = createSvgRoot(width, height);

    // Shared arrow marker so every connector path reuses the same tip.
    // `orient="auto"` rotates the marker along the path tangent at its
    // end point, which is what makes the arrow point "forward" around the
    // loop. `markerUnits="strokeWidth"` keeps the arrowhead proportional
    // to the stroke; we set stroke-width=1 on paths so the raw refX/refY
    // values read as pixels relative to a 10-unit marker box.
    const defs = document.createElementNS(SVG_NS, 'defs');
    const marker = document.createElementNS(SVG_NS, 'marker');
    marker.setAttribute('id', 'xlsx-smartart-arrow');
    marker.setAttribute('viewBox', '0 0 10 10');
    marker.setAttribute('refX', '8');
    marker.setAttribute('refY', '5');
    marker.setAttribute('markerWidth', '6');
    marker.setAttribute('markerHeight', '6');
    marker.setAttribute('orient', 'auto');
    const arrow = document.createElementNS(SVG_NS, 'polygon');
    arrow.setAttribute('points', '0,0 10,5 0,10');
    arrow.setAttribute('fill', CONNECTOR_COLOR);
    marker.appendChild(arrow);
    defs.appendChild(marker);
    svg.appendChild(defs);

    if (nodes.length === 0) return svg;

    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.max(40, Math.min(width, height) / 3 - 30);

    // Compute per-node centres. Start at the 12-o'clock position
    // (angle = -π/2) and go clockwise so the order reads naturally.
    interface CyclePlaced {
        node: SmartArtNode;
        x: number;
        y: number;
    }
    const placed: CyclePlaced[] = [];
    const n = nodes.length;
    for (let i = 0; i < n; i++) {
        const angle = -Math.PI / 2 + (2 * Math.PI * i) / Math.max(1, n);
        placed.push({
            node: nodes[i],
            x: cx + radius * Math.cos(angle),
            y: cy + radius * Math.sin(angle),
        });
    }

    // Connector arcs. Each arc starts at the outer edge of the current
    // node (on the line from centre → node, past the node) and ends at the
    // outer edge of the next node (on the line from next node → centre,
    // short of the node). We shorten both endpoints by a small "halo" so
    // the path doesn't cut through the rectangles.
    if (n >= 2) {
        const halo = 18; // keep the arc clear of the node rectangle
        for (let i = 0; i < n; i++) {
            const a = placed[i];
            const b = placed[(i + 1) % n];
            // Vector centre→a, normalised.
            const dax = a.x - cx;
            const day = a.y - cy;
            const la = Math.hypot(dax, day) || 1;
            const ux1 = dax / la;
            const uy1 = day / la;
            // Vector centre→b, normalised.
            const dbx = b.x - cx;
            const dby = b.y - cy;
            const lb = Math.hypot(dbx, dby) || 1;
            const ux2 = dbx / lb;
            const uy2 = dby / lb;

            // Start a halo outside the current node (bigger radius so the
            // arc curves outside the circle of node centres — visually
            // pleasing and keeps the arc clear of every node rectangle).
            const arcRadius = radius + halo;
            const sx = cx + ux1 * arcRadius;
            const sy = cy + uy1 * arcRadius;
            const ex = cx + ux2 * arcRadius;
            const ey = cy + uy2 * arcRadius;

            // Arc command: sweep along the outer circle (larger than the
            // node-centre circle) clockwise. `0 0 1` = small arc, clockwise.
            const d = `M ${sx.toFixed(2)} ${sy.toFixed(2)} A ${arcRadius.toFixed(2)} ${arcRadius.toFixed(2)} 0 0 1 ${ex.toFixed(2)} ${ey.toFixed(2)}`;
            const path = document.createElementNS(SVG_NS, 'path');
            path.setAttribute('d', d);
            path.setAttribute('fill', 'none');
            path.setAttribute('stroke', CONNECTOR_COLOR);
            path.setAttribute('stroke-width', '1');
            path.setAttribute('marker-end', 'url(#xlsx-smartart-arrow)');
            svg.appendChild(path);
        }
    }

    // Draw nodes on top. All cycle nodes share the first-level palette entry
    // (the cycle has no depth dimension — every step is semantically peer).
    const fill = PALETTE[0];
    for (const p of placed) {
        appendNode(svg, p.x, p.y, CYCLE_NODE_WIDTH, CYCLE_NODE_HEIGHT, CYCLE_NODE_RX, fill, p.node.text);
    }

    return svg;
}

// Shared SVG root. Width / height / viewBox + the xlsx-smartart-svg class.
function createSvgRoot(width: number, height: number): SVGSVGElement {
    const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    svg.setAttribute('xmlns', SVG_NS);
    svg.setAttribute('width', String(width));
    svg.setAttribute('height', String(height));
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
    svg.setAttribute('class', 'xlsx-smartart-svg');
    return svg;
}

// Append one rounded-rect node + centred label. Centre coordinates in,
// attacker-controlled text goes through textContent only.
function appendNode(
    svg: SVGSVGElement,
    cx: number,
    cy: number,
    w: number,
    h: number,
    rx: number,
    fill: string,
    text: string,
): void {
    const g = document.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'xlsx-smartart-node');

    const rect = document.createElementNS(SVG_NS, 'rect');
    rect.setAttribute('x', String(cx - w / 2));
    rect.setAttribute('y', String(cy - h / 2));
    rect.setAttribute('width', String(w));
    rect.setAttribute('height', String(h));
    rect.setAttribute('rx', String(rx));
    rect.setAttribute('ry', String(rx));
    rect.setAttribute('fill', fill);
    rect.setAttribute('stroke', fill);
    g.appendChild(rect);

    const t = document.createElementNS(SVG_NS, 'text');
    t.setAttribute('x', String(cx));
    t.setAttribute('y', String(cy));
    t.setAttribute('text-anchor', 'middle');
    t.setAttribute('dominant-baseline', 'middle');
    t.setAttribute('fill', TEXT_COLOR);
    t.setAttribute('font-family', 'system-ui, sans-serif');
    t.setAttribute('font-size', '12');
    // Node text is attacker-controlled — textContent only.
    t.textContent = text;
    g.appendChild(t);

    svg.appendChild(g);
}
