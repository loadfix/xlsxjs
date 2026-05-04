// SmartArtModel → inline SVG. First-pass hierarchy renderer: each level of
// the parsed tree lays out horizontally across the viewport, nodes render as
// rounded rectangles with centred text, and parents connect to children via
// straight lines from the parent's bottom-centre to the child's top-centre.
//
// This slice intentionally covers ONE layout strategy (top-down hierarchy)
// even when `model.layout` announces a cycle / orgchart / pyramid variant —
// those remain on the roadmap. Falling back to the hierarchy layout for any
// diagram keeps the opt-in SVG useful while we grow the repertoire.
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

export interface RenderSmartArtOptions {
    width?: number;
    height?: number;
}

/**
 * Render a SmartArt hierarchy as an inline SVG. Returns `null` when the
 * parsed model has no root nodes — callers can branch on that to fall back
 * to the `<ul>` tree emitted by Wave 9 renderer.
 *
 * The layout is always top-down hierarchy regardless of `model.layout`. Each
 * depth level lays its nodes out evenly spaced across the available width;
 * parents connect to their children via straight lines from the parent's
 * bottom-centre to the child's top-centre.
 */
export function renderSmartArtSvg(
    model: SmartArtModel,
    opts: RenderSmartArtOptions = {},
): SVGSVGElement | null {
    const roots = model.rootNodes;
    if (!roots || roots.length === 0) return null;

    const width = opts.width ?? DEFAULT_WIDTH;
    const height = opts.height ?? DEFAULT_HEIGHT;

    // Flatten the tree into levels. Each level is a list of { node, parentIndex }
    // tuples where parentIndex points at the level-above entry. We lay each
    // level out horizontally, evenly spaced across the viewport.
    interface Placed {
        node: SmartArtNode;
        parentIndex: number; // -1 for roots
        x: number;           // centre x of the node box
        y: number;           // centre y of the node box
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
    for (const root of roots) walk(root, 0, -1);

    const depthCount = levels.length;
    // Vertical positioning: evenly distribute level centres across the height.
    // Top and bottom margins reserved so rounded rects don't clip.
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

    const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
    svg.setAttribute('xmlns', SVG_NS);
    svg.setAttribute('width', String(width));
    svg.setAttribute('height', String(height));
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
    svg.setAttribute('class', 'xlsx-smartart-svg');

    // Draw connectors first so they sit beneath the node rectangles.
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

    // Draw nodes on top.
    for (let d = 0; d < depthCount; d++) {
        const row = levels[d];
        const fill = PALETTE[d % PALETTE.length];
        for (const placed of row) {
            const g = document.createElementNS(SVG_NS, 'g');
            g.setAttribute('class', 'xlsx-smartart-node');

            const rect = document.createElementNS(SVG_NS, 'rect');
            rect.setAttribute('x', String(placed.x - NODE_WIDTH / 2));
            rect.setAttribute('y', String(placed.y - NODE_HEIGHT / 2));
            rect.setAttribute('width', String(NODE_WIDTH));
            rect.setAttribute('height', String(NODE_HEIGHT));
            rect.setAttribute('rx', String(NODE_RX));
            rect.setAttribute('ry', String(NODE_RX));
            rect.setAttribute('fill', fill);
            rect.setAttribute('stroke', fill);
            g.appendChild(rect);

            const text = document.createElementNS(SVG_NS, 'text');
            text.setAttribute('x', String(placed.x));
            text.setAttribute('y', String(placed.y));
            text.setAttribute('text-anchor', 'middle');
            text.setAttribute('dominant-baseline', 'middle');
            text.setAttribute('fill', TEXT_COLOR);
            text.setAttribute('font-family', 'system-ui, sans-serif');
            text.setAttribute('font-size', '12');
            // Node text is attacker-controlled — textContent only.
            text.textContent = placed.node.text;
            g.appendChild(text);

            svg.appendChild(g);
        }
    }

    return svg;
}
