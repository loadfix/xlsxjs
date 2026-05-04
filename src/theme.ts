// Parse xl/theme/theme1.xml into the 12-element SpreadsheetML theme colour
// table and provide a tint-aware colour resolver.
//
// ECMA-376 §18.8.19 `ST_ColorType`: a colour element can carry rgb, indexed,
// theme, or auto, optionally modified by `tint` in [-1, 1]. We handle rgb and
// theme here (indexed / auto are TODO). The tint formula operates on the
// HSL luminance component:
//   tint < 0 → L' = L * (1 + tint)
//   tint > 0 → L' = L + (1 - L) * tint
//
// The theme clrScheme XML children are in the order <dk1> <lt1> <dk2> <lt2>
// <accent1>…<accent6> <hlink> <folHlink>, but SpreadsheetML's theme index
// swaps 0↔1 and 2↔3 so that 0=lt1 (background 1), 1=dk1 (text 1), etc. We
// do the swap at parse time so downstream code can just index by the
// SpreadsheetML convention.

const NS_DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export type Theme = {
    // 12 entries, indexed per SpreadsheetML theme= convention.
    // 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accent1..6, 10 hlink, 11 folHlink
    colors: (string | null)[];
};

export type ColorRef =
    | { kind: 'rgb'; value: string }                        // '#rrggbb'
    | { kind: 'theme'; index: number; tint: number }
    | { kind: 'indexed'; index: number; tint: number }
    | null;

// Excel's legacy 64-entry "indexed colors" palette from ECMA-376 §18.8.27.
// Indices 64/65 are system fg/bg; we emit nulls there and let the renderer
// fall back to default text / background.
const INDEXED_PALETTE: (string | null)[] = [
    '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
    '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
    '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
    '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
    '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
    '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
    '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
    '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
    null, null,
];

export function indexedColor(index: number): string | null {
    if (!Number.isFinite(index) || index < 0 || index >= INDEXED_PALETTE.length) return null;
    return INDEXED_PALETTE[index];
}

const HEX_6 = /^([A-Fa-f0-9]{6})$/;
const HEX_ARGB = /^([A-Fa-f0-9]{2})([A-Fa-f0-9]{6})$/;

export function rgbToHex(value: string | null | undefined): string | null {
    if (!value) return null;
    const trimmed = value.trim();
    const argb = HEX_ARGB.exec(trimmed);
    if (argb) return `#${argb[2].toLowerCase()}`;
    const rgb = HEX_6.exec(trimmed);
    if (rgb) return `#${rgb[1].toLowerCase()}`;
    return null;
}

function parseXml(xml: string): Document {
    return new DOMParser().parseFromString(xml, 'application/xml');
}

export function parseTheme(xml: string): Theme {
    // Empty theme defaults — black / white for the first four, null for the
    // rest. Any reference to an unresolvable theme index falls through to
    // null so the renderer leaves the cell unstyled.
    const colors: (string | null)[] = new Array(12).fill(null);
    colors[0] = '#ffffff'; colors[1] = '#000000';
    colors[2] = '#e7e6e6'; colors[3] = '#44546a';

    const doc = parseXml(xml);
    const scheme = doc.getElementsByTagNameNS(NS_DRAW, 'clrScheme').item(0);
    if (!scheme) return { colors };

    // Walk children by local name instead of positional index so we're
    // robust to schemes that omit optional members.
    const get = (localName: string): string | null => {
        const el = scheme.getElementsByTagNameNS(NS_DRAW, localName).item(0);
        if (!el) return null;
        return resolveClrChild(el);
    };

    // ECMA-376 names → SpreadsheetML index (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4+=accents).
    const mapping: [string, number][] = [
        ['lt1', 0], ['dk1', 1], ['lt2', 2], ['dk2', 3],
        ['accent1', 4], ['accent2', 5], ['accent3', 6], ['accent4', 7],
        ['accent5', 8], ['accent6', 9], ['hlink', 10], ['folHlink', 11],
    ];
    for (const [name, idx] of mapping) {
        const c = get(name);
        if (c) colors[idx] = c;
    }
    return { colors };
}

function resolveClrChild(wrapper: Element): string | null {
    // <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    // <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    const srgb = wrapper.getElementsByTagNameNS(NS_DRAW, 'srgbClr').item(0);
    if (srgb) return rgbToHex(srgb.getAttribute('val'));
    const sysClr = wrapper.getElementsByTagNameNS(NS_DRAW, 'sysClr').item(0);
    if (sysClr) {
        const last = sysClr.getAttribute('lastClr');
        if (last) return rgbToHex(last);
        const val = sysClr.getAttribute('val');
        if (val === 'windowText') return '#000000';
        if (val === 'window') return '#ffffff';
    }
    return null;
}

// ── tint + HSL helpers ─────────────────────────────────────────────────────

function hexToRgb(hex: string): [number, number, number] | null {
    const m = /^#([0-9a-f]{6})$/i.exec(hex);
    if (!m) return null;
    return [
        parseInt(m[1].slice(0, 2), 16),
        parseInt(m[1].slice(2, 4), 16),
        parseInt(m[1].slice(4, 6), 16),
    ];
}

function rgbToHexStr(r: number, g: number, b: number): string {
    const hex = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
    return `#${hex(r)}${hex(g)}${hex(b)}`;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r: h = ((g - b) / d) + (g < b ? 6 : 0); break;
            case g: h = ((b - r) / d) + 2; break;
            case b: h = ((r - g) / d) + 4; break;
        }
        h /= 6;
    }
    return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const hue2rgb = (p: number, q: number, t: number) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    const r = hue2rgb(p, q, h + 1 / 3);
    const g = hue2rgb(p, q, h);
    const b = hue2rgb(p, q, h - 1 / 3);
    return [r * 255, g * 255, b * 255];
}

export function applyTint(hex: string, tint: number): string {
    if (!tint || !Number.isFinite(tint)) return hex;
    const rgb = hexToRgb(hex);
    if (!rgb) return hex;
    const [h, s, l] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
    const clamped = Math.max(-1, Math.min(1, tint));
    const newL = clamped < 0 ? l * (1 + clamped) : l + (1 - l) * clamped;
    const [r, g, b] = hslToRgb(h, s, newL);
    return rgbToHexStr(r, g, b);
}

export function resolveColor(ref: ColorRef, theme: Theme | null): string | null {
    if (!ref) return null;
    if (ref.kind === 'rgb') return ref.value;
    if (ref.kind === 'theme') {
        const base = theme?.colors[ref.index] ?? null;
        if (!base) return null;
        return ref.tint ? applyTint(base, ref.tint) : base;
    }
    if (ref.kind === 'indexed') {
        const base = indexedColor(ref.index);
        if (!base) return null;
        return ref.tint ? applyTint(base, ref.tint) : base;
    }
    return null;
}
