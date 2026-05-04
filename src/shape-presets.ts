// Shape-preset SVG palette. xlsxjs surfaces `SheetShape.preset` as the
// `<a:prstGeom prst=…>` name from the drawing XML. This module maps the
// most common Excel prstGeom values to an inline SVG string the renderer
// can overlay on the `<aside class="xlsx-shape">`.
//
// SVGs use `viewBox="0 0 100 100"` with `preserveAspectRatio="none"` so CSS
// can resize them to fill the aside (which is sized by anchor / CSS, not
// by the SVG itself). All strokes default to `opts.stroke` (pass `#888`
// for the neutral palette). Fills are `none` by default; shapes whose
// Excel rendering implies a filled body (arrows, stars, callouts) take
// a light fill (`#e0e0e0`) so the glyph reads as a solid Excel shape
// rather than an outline.
//
// Unknown presets return null; the renderer falls back to the plain
// bordered `<aside>`.

export interface ShapePresetOptions {
    width: number;
    height: number;
    stroke: string;
}

const DEFAULT_FILL = '#e0e0e0';

// Build the outer <svg>. We emit preserveAspectRatio="none" so the path
// stretches with the aside (Excel's prstGeom scales the same way).
function wrap(inner: string): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">${inner}</svg>`;
}

// Geometry helpers: a regular n-gon centred at (50,50) with circumradius 48,
// point-up (first vertex at the top).
function regularPolygon(sides: number): string {
    const cx = 50, cy = 50, r = 48;
    const pts: string[] = [];
    for (let i = 0; i < sides; i++) {
        const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
        const x = cx + r * Math.cos(angle);
        const y = cy + r * Math.sin(angle);
        pts.push(`${x.toFixed(2)},${y.toFixed(2)}`);
    }
    return pts.join(' ');
}

// 5-point star centred at (50,50), outer radius 48, inner radius 20.
function starPoints(): string {
    const cx = 50, cy = 50, rOuter = 48, rInner = 20;
    const pts: string[] = [];
    for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? rOuter : rInner;
        const angle = -Math.PI / 2 + (i * Math.PI) / 5;
        const x = cx + r * Math.cos(angle);
        const y = cy + r * Math.sin(angle);
        pts.push(`${x.toFixed(2)},${y.toFixed(2)}`);
    }
    return pts.join(' ');
}

// Map a preset name to the SVG body (no outer <svg>). Returns null for
// unknown presets; callers should fall back to the plain <aside>.
function presetBody(preset: string, stroke: string): string | null {
    const s = stroke;
    switch (preset) {
        case 'rect':
        case 'flowChartProcess':
            return `<rect x="1" y="1" width="98" height="98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'roundRect':
            return `<rect x="1" y="1" width="98" height="98" rx="8" ry="8" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'ellipse':
        case 'flowChartConnector':
            return `<ellipse cx="50" cy="50" rx="49" ry="49" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'line':
            return `<line x1="0" y1="0" x2="100" y2="100" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'triangle':
            return `<polygon points="50,2 98,98 2,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'rtTriangle':
            // Right triangle, right angle at bottom-left.
            return `<polygon points="2,2 2,98 98,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'diamond':
        case 'flowChartDecision':
            return `<polygon points="50,2 98,50 50,98 2,50" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'parallelogram':
            // 20% shear to the right along the top edge.
            return `<polygon points="22,2 98,2 78,98 2,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'trapezoid':
            // Isosceles trapezoid, wide base on the bottom.
            return `<polygon points="22,2 78,2 98,98 2,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'pentagon':
            return `<polygon points="${regularPolygon(5)}" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'hexagon':
            return `<polygon points="${regularPolygon(6)}" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'octagon':
            return `<polygon points="${regularPolygon(8)}" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'star5':
            return `<polygon points="${starPoints()}" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'rightArrow':
            // Horizontal arrow pointing right. Shaft occupies vertical
            // band 30..70, head is the triangle 60..98 across the full
            // 10..90 vertical range. Path keeps it as one closed glyph.
            return `<path d="M2 30 L60 30 L60 10 L98 50 L60 90 L60 70 L2 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'leftArrow':
            return `<path d="M98 30 L40 30 L40 10 L2 50 L40 90 L40 70 L98 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'upArrow':
            return `<path d="M30 98 L30 40 L10 40 L50 2 L90 40 L70 40 L70 98 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'downArrow':
            return `<path d="M30 2 L30 60 L10 60 L50 98 L90 60 L70 60 L70 2 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'leftRightArrow':
            // Double-ended horizontal arrow. Shaft 30..70, heads at each end.
            return `<path d="M2 50 L20 30 L20 40 L80 40 L80 30 L98 50 L80 70 L80 60 L20 60 L20 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'flowChartTerminator':
            // Pill: rect with height-sized corner radii.
            return `<rect x="1" y="1" width="98" height="98" rx="49" ry="49" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'callout1':
        case 'wedgeRectCallout':
            // Rectangle body plus a small triangular tail pointing to the
            // bottom-left corner of the shape's anchor box.
            return `<rect x="1" y="1" width="98" height="78" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<polygon points="15,79 35,79 8,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'wedgeEllipseCallout':
            return `<ellipse cx="50" cy="40" rx="49" ry="39" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<polygon points="20,75 38,75 8,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        case 'cloudCallout':
            // Approximate a cloud with 4 overlapping circles plus two small
            // trailing bubbles pointing to the callout source (bottom-left).
            return `<ellipse cx="30" cy="45" rx="22" ry="20" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<ellipse cx="60" cy="35" rx="26" ry="22" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<ellipse cx="78" cy="55" rx="18" ry="18" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<ellipse cx="50" cy="60" rx="28" ry="18" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<circle cx="18" cy="82" r="6" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                `<circle cx="8" cy="94" r="4" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;

        default:
            return null;
    }
}

// Return an inline SVG string representing the Excel prstGeom preset, or
// null when the preset is unknown. The caller supplies `opts.stroke`; any
// SVG fill is either `none` or `DEFAULT_FILL` (#e0e0e0) — see header.
//
// `opts.width` / `opts.height` are accepted for forward compatibility but
// aren't used today: the viewBox is unit-scaled (0..100) and the element
// is sized by CSS.
export function renderShapePreset(
    preset: string | null,
    opts: ShapePresetOptions,
): string | null {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const _opts = opts;
    if (preset === null || preset === '') return null;
    const body = presetBody(preset, opts.stroke);
    if (body === null) return null;
    return wrap(body);
}
