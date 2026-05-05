// Parse xl/styles.xml into a small, renderer-friendly model. Everything here
// is pure XML → typed data; no DOM, no formatting decisions. Colour elements
// are parsed into ColorRef (rgb or theme+tint) so the renderer can resolve
// them against xl/theme/theme1.xml when it has it.
//
// Scope:
//   - <numFmts>          custom number format strings (+ built-in lookup)
//   - <fonts>            bold, italic, underline, size, colour, name
//   - <fills>            patternFill (solid + stripe patterns) + gradientFill
//   - <borders>          per-side style + colour, including diagonal
//   - <cellXfs>          the indexed array a cell's s= attribute points at
//
// Not yet: named styles beyond single-step cellStyleXfs inheritance,
// tableStyles, path-gradient fills (parsed but rendered as a flat fallback).

import type { ColorRef } from './theme';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | null;

export interface FontStyle {
    bold: boolean;
    italic: boolean;
    // Underline variants per ECMA-376 §18.8.36: 'single' (the default when
    // <u/> is present with no val), 'double', 'singleAccounting',
    // 'doubleAccounting'. null = no underline.
    underline: UnderlineStyle;
    strike: boolean;
    // Sub/superscript (ECMA-376 §18.4.13 vertAlign). null = baseline.
    vertAlign: 'subscript' | 'superscript' | null;
    size: number | null;
    color: ColorRef;
    name: string | null;
    // <scheme val="major|minor"/> — a reference to the theme's fontScheme.
    // Resolved by the renderer against `Theme.majorFont` / `Theme.minorFont`
    // when no explicit `name` is set. null when the font has no <scheme/>.
    scheme: 'major' | 'minor' | null;
}

// Discriminated union: pattern fills (solid + the stripe/grid patterns),
// gradient fills (linear + path), and the explicit "none" variant which is
// what `<patternFill patternType="none"/>` and `gray125` decompose to.
//
// The renderer switches on `kind` and applies the appropriate CSS. Pattern
// preserves both fgColor and bgColor so non-solid stripe patterns can use
// fgColor for the marks and bgColor behind them.
export type PatternType =
    | 'none' | 'solid' | 'gray125'
    | 'darkGray' | 'mediumGray' | 'lightGray'
    | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp'
    | 'darkGrid' | 'darkTrellis'
    | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp'
    | 'lightGrid' | 'lightTrellis';

export interface PatternFill {
    kind: 'pattern';
    patternType: PatternType;
    // For solid fills fgColor is the canonical fill colour; for non-solid
    // stripe/grid patterns, fgColor paints the marks and bgColor paints the
    // background behind them.
    fgColor: ColorRef;
    bgColor: ColorRef;
}

export interface GradientStop {
    position: number;   // 0.0 .. 1.0
    color: ColorRef;
}

export interface GradientFill {
    kind: 'gradient';
    // 'linear' = classic angled gradient; 'path' = radial from a rectangular
    // region (left/right/top/bottom anchor attrs). We parse both but render
    // path gradients as a flat fallback to the first stop's colour.
    type: 'linear' | 'path';
    degree: number;     // 0..360 rotation (linear only); 0 for path
    stops: GradientStop[];
}

export interface NoFill {
    kind: 'none';
}

export type FillStyle = PatternFill | GradientFill | NoFill;

export interface BorderSide {
    style: string | null; // 'thin' | 'medium' | 'thick' | 'dashed' | ...
    color: ColorRef;
}

export interface BorderStyle {
    left: BorderSide;
    right: BorderSide;
    top: BorderSide;
    bottom: BorderSide;
    // Diagonal line — only painted when diagonalUp or diagonalDown is true.
    // CSS has no border-diagonal so the renderer paints this as an inline
    // linear-gradient overlay (see html-renderer.ts applyBorder).
    diagonal: BorderSide;
    diagonalUp: boolean;    // <border diagonalUp="1"/> (bottom-left → top-right)
    diagonalDown: boolean;  // <border diagonalDown="1"/> (top-left → bottom-right)
}

export interface Alignment {
    horizontal: 'left' | 'right' | 'center' | 'justify' | 'distributed' | 'centerContinuous' | 'fill' | null;
    // ST_VerticalAlignment values from ECMA-376 §18.18.88 — kept verbatim
    // (note: "center", not the CSS spelling "middle"; the renderer maps
    // `center` → CSS vertical-align: middle).
    vertical: 'top' | 'center' | 'bottom' | 'justify' | 'distributed' | null;
    // True when <alignment @wrapText="1"/> is present. Renderer maps this to
    // white-space: normal + word-break so multi-line content stays in cell.
    wrapText: boolean;
    // True when <alignment @shrinkToFit="1"/> is present. CSS cannot fully
    // replicate Excel's shrink behaviour (which measures the rendered text
    // against cell width), so the renderer tags the td with a class and
    // leaves fine-tuning to the consumer.
    shrinkToFit: boolean;
    // Indent level in Excel character-width units. Conventionally 1 unit ≈
    // 0.5em; applied as paddingLeft/paddingRight depending on horizontal
    // alignment + reading order.
    indent: number;
    // 0..180 = rotation in degrees (counter-clockwise); 255 = stacked
    // vertical (East-Asian convention). null = not specified.
    textRotation: number | null;
    // 0 = context (default), 1 = LTR, 2 = RTL. 0 leaves direction unset on
    // the td; 2 emits direction: rtl.
    readingOrder: 0 | 1 | 2;
}

// The "untouched" default. Factored out so cellXf and cellStyleXf share the
// same baseline and the renderer can cheaply check "is any alignment set?"
// by comparing against this shape.
function defaultAlignment(): Alignment {
    return {
        horizontal: null,
        vertical: null,
        wrapText: false,
        shrinkToFit: false,
        indent: 0,
        textRotation: null,
        readingOrder: 0,
    };
}

export interface CellXf {
    numFmtId: number;
    fontId: number;
    fillId: number;
    borderId: number;
    // xfId points into cellStyleXfs. -1 means "no base style" (e.g. the
    // default "Normal" xf at index 0, or a cellXf emitted with no xfId).
    xfId: number;
    applyNumberFormat: boolean;
    applyFont: boolean;
    applyFill: boolean;
    applyBorder: boolean;
    applyAlignment: boolean;
    alignment: Alignment;
}

// A "differential format". Used by conditional formatting (and table styles,
// but we don't render those). Every field is optional — a dxf is applied
// on top of the existing cell formatting, overriding only what it specifies.
export interface Dxf {
    font?: Partial<FontStyle>;
    fill?: FillStyle;
    border?: BorderStyle;
    numFmtCode?: string;
    // Unlike a cellXf, dxfs carry the formatCode directly on the dxf (no id
    // lookup). Nullable because most dxfs don't change number formatting.
}

export interface Styles {
    numFmts: Map<number, string>; // id → format code
    fonts: FontStyle[];
    fills: FillStyle[];
    borders: BorderStyle[];
    cellXfs: CellXf[];
    // Base XFs referenced by named styles. `cellXfs[i].xfId` is an index
    // into this array. We merge these under the cellXf at resolution time.
    cellStyleXfs: CellXf[];
    dxfs: Dxf[];
}

// The canonical ECMA-376 built-in number formats (IDs 0..49). Ids 41..44 are
// reserved per the spec; everything else comes straight from §18.8.30.
const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
};

export function lookupNumberFormat(styles: Styles | null, numFmtId: number): string | null {
    if (styles?.numFmts.has(numFmtId)) return styles.numFmts.get(numFmtId)!;
    if (numFmtId in BUILTIN_NUMBER_FORMATS) return BUILTIN_NUMBER_FORMATS[numFmtId];
    return null;
}

// Allow only 6- or 8-hex colour codes (Excel serialises colours as ARGB).
// Anything else — rgb() strings we fail to parse, bare identifiers — returns
// null so the renderer falls back to the default.
const HEX_COLOR = /^([A-Fa-f0-9]{6})([A-Fa-f0-9]{2})?$/;
const ARGB_COLOR = /^([A-Fa-f0-9]{2})([A-Fa-f0-9]{6})$/;

// Sanitize a <font><name val="…"/></font> value for safe echo into
// td.style.fontFamily. Attacker-controlled strings cannot reach the DOM
// uninspected — a name like 'Arial; display: none' would break out of the
// font-family property if naively interpolated. Mirrors docxjs's
// sanitizeFontFamily pattern.
//
// Rules:
//   - null / empty → null (no font-family applied).
//   - Any of ; { } < > or a newline → null (injection attempt rejected).
//   - Remaining chars must be alphanumeric / spaces / hyphens / dots;
//     anything else → null.
//   - Accepted names are wrapped in double quotes so font names like
//     "Times New Roman" work and no further escaping is needed.
const SAFE_FONT_FAMILY = /^[A-Za-z0-9 .\-]+$/;

export function sanitizeFontFamily(name: string | null | undefined): string | null {
    if (!name) return null;
    const trimmed = name.trim();
    if (!trimmed) return null;
    // Reject obvious CSS-injection markers outright.
    if (/[;{}<>\n\r]/.test(trimmed)) return null;
    if (!SAFE_FONT_FAMILY.test(trimmed)) return null;
    return `"${trimmed}"`;
}

export function sanitizeHexColor(value: string | null | undefined): string | null {
    if (!value) return null;
    const trimmed = value.trim();
    const argb = ARGB_COLOR.exec(trimmed);
    if (argb) return `#${argb[2].toLowerCase()}`;
    const rgb = HEX_COLOR.exec(trimmed);
    if (rgb) return `#${rgb[1].toLowerCase()}`;
    return null;
}

// Parse a <color rgb=.../theme=.../indexed=.../tint=.../> element into a
// ColorRef. Used by both the styles parser and the rich-text runs in shared
// strings. Attribute precedence follows the schema: rgb > theme > indexed.
export function parseColorElement(el: Element | null): ColorRef {
    if (!el) return null;
    const tintAttr = el.getAttribute('tint');
    const tint = tintAttr ? Number(tintAttr) : 0;
    const safeTint = Number.isFinite(tint) ? tint : 0;

    const rgb = el.getAttribute('rgb');
    if (rgb) {
        const sanitized = sanitizeHexColor(rgb);
        if (sanitized) return { kind: 'rgb', value: sanitized };
        return null;
    }
    const themeAttr = el.getAttribute('theme');
    if (themeAttr != null) {
        const index = Number(themeAttr);
        if (!Number.isFinite(index) || index < 0) return null;
        return { kind: 'theme', index, tint: safeTint };
    }
    const indexedAttr = el.getAttribute('indexed');
    if (indexedAttr != null) {
        const index = Number(indexedAttr);
        if (!Number.isFinite(index) || index < 0) return null;
        return { kind: 'indexed', index, tint: safeTint };
    }
    return null;
}

function parseXml(xml: string): Document {
    return new DOMParser().parseFromString(xml, 'application/xml');
}

export function parseStyles(xml: string): Styles {
    const doc = parseXml(xml);
    return {
        numFmts: parseNumFmts(doc),
        fonts: parseFonts(doc),
        fills: parseFills(doc),
        borders: parseBorders(doc),
        cellXfs: parseCellXfs(doc, 'cellXfs'),
        cellStyleXfs: parseCellXfs(doc, 'cellStyleXfs'),
        dxfs: parseDxfs(doc),
    };
}

function parseNumFmts(doc: Document): Map<number, string> {
    const out = new Map<number, string>();
    const els = doc.getElementsByTagNameNS(NS_MAIN, 'numFmt');
    for (let i = 0; i < els.length; i++) {
        const id = Number(els[i].getAttribute('numFmtId'));
        const code = els[i].getAttribute('formatCode') ?? '';
        if (Number.isFinite(id) && id >= 0) out.set(id, code);
    }
    return out;
}

function parseFonts(doc: Document): FontStyle[] {
    const fontsRoot = doc.getElementsByTagNameNS(NS_MAIN, 'fonts').item(0);
    const out: FontStyle[] = [];
    if (!fontsRoot) return out;
    const fonts = fontsRoot.getElementsByTagNameNS(NS_MAIN, 'font');
    for (let i = 0; i < fonts.length; i++) out.push(parseFont(fonts[i]));
    return out;
}

function parseFont(el: Element): FontStyle {
    const has = (tag: string) => el.getElementsByTagNameNS(NS_MAIN, tag).length > 0;
    const firstAttr = (tag: string, attr: string): string | null => {
        const node = el.getElementsByTagNameNS(NS_MAIN, tag).item(0);
        return node ? node.getAttribute(attr) : null;
    };
    const colorEl = el.getElementsByTagNameNS(NS_MAIN, 'color').item(0);
    const sizeAttr = firstAttr('sz', 'val');
    return {
        bold: has('b'),
        italic: has('i'),
        underline: parseUnderline(el),
        strike: has('strike'),
        vertAlign: parseVertAlign(el),
        size: sizeAttr ? Number(sizeAttr) : null,
        color: parseColorElement(colorEl),
        name: firstAttr('name', 'val') ?? firstAttr('rFont', 'val'),
        scheme: parseScheme(el),
    };
}

// ECMA-376 §18.8.35: <scheme val="major|minor"/> inside a <font/> marks the
// font as the theme's major or minor face. Anything unrecognised (including
// legacy `val="none"`) falls back to null so the renderer ignores the hint.
function parseScheme(parent: Element): 'major' | 'minor' | null {
    const s = parent.getElementsByTagNameNS(NS_MAIN, 'scheme').item(0);
    if (!s) return null;
    const val = s.getAttribute('val');
    if (val === 'major' || val === 'minor') return val;
    return null;
}

// ECMA-376 §18.8.36: <u/> with no val attribute means 'single'; explicit
// values can be 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'.
// Unknown values fall back to null so we don't invent styles.
function parseUnderline(parent: Element): UnderlineStyle {
    const u = parent.getElementsByTagNameNS(NS_MAIN, 'u').item(0);
    if (!u) return null;
    const val = u.getAttribute('val');
    if (val === null || val === '') return 'single';
    if (val === 'single' || val === 'double' || val === 'singleAccounting' || val === 'doubleAccounting') {
        return val;
    }
    if (val === 'none') return null;
    return null;
}

function parseVertAlign(parent: Element): 'subscript' | 'superscript' | null {
    const va = parent.getElementsByTagNameNS(NS_MAIN, 'vertAlign').item(0);
    if (!va) return null;
    const val = va.getAttribute('val');
    if (val === 'subscript' || val === 'superscript') return val;
    return null;
}

function parseFills(doc: Document): FillStyle[] {
    const fillsRoot = doc.getElementsByTagNameNS(NS_MAIN, 'fills').item(0);
    const out: FillStyle[] = [];
    if (!fillsRoot) return out;
    const fills = fillsRoot.getElementsByTagNameNS(NS_MAIN, 'fill');
    for (let i = 0; i < fills.length; i++) out.push(parseFill(fills[i]));
    return out;
}

// Valid patternType values per ECMA-376 §18.18.55. Anything unrecognised
// (or missing) decays to 'none' so the renderer leaves the cell unfilled.
const PATTERN_TYPES = new Set<string>([
    'none', 'solid', 'gray125',
    'darkGray', 'mediumGray', 'lightGray',
    'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp',
    'darkGrid', 'darkTrellis',
    'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
    'lightGrid', 'lightTrellis',
]);

function parseFill(el: Element, opts?: { allowBgFallback?: boolean }): FillStyle {
    const gf = el.getElementsByTagNameNS(NS_MAIN, 'gradientFill').item(0);
    if (gf) return parseGradientFill(gf);
    const pf = el.getElementsByTagNameNS(NS_MAIN, 'patternFill').item(0);
    if (!pf) return { kind: 'none' };
    const rawType = pf.getAttribute('patternType');
    let patternType: PatternType = rawType && PATTERN_TYPES.has(rawType)
        ? rawType as PatternType
        : 'none';
    const fgEl = pf.getElementsByTagNameNS(NS_MAIN, 'fgColor').item(0);
    const bgEl = pf.getElementsByTagNameNS(NS_MAIN, 'bgColor').item(0);
    let fgColor = parseColorElement(fgEl);
    const bgColor = parseColorElement(bgEl);
    // Legacy fallback: dxf fills and some older Excel writers omit the
    // patternType and put the colour on bgColor instead of fgColor. When
    // caller opts in and fgColor is missing, promote bgColor into fgColor
    // and treat the fill as a solid (so the renderer lands a background
    // colour rather than skipping the 'none' fill).
    if (!fgColor && opts?.allowBgFallback && bgColor) {
        fgColor = bgColor;
        if (patternType === 'none') patternType = 'solid';
    }
    return { kind: 'pattern', patternType, fgColor, bgColor };
}

function parseGradientFill(el: Element): GradientFill {
    const typeAttr = el.getAttribute('type');
    const type: GradientFill['type'] = typeAttr === 'path' ? 'path' : 'linear';
    const degAttr = el.getAttribute('degree');
    const degN = degAttr != null ? Number(degAttr) : 0;
    const degree = Number.isFinite(degN) ? ((degN % 360) + 360) % 360 : 0;
    const stops: GradientStop[] = [];
    const stopEls = el.getElementsByTagNameNS(NS_MAIN, 'stop');
    for (let i = 0; i < stopEls.length; i++) {
        const s = stopEls[i];
        const posAttr = s.getAttribute('position');
        const posN = posAttr != null ? Number(posAttr) : NaN;
        if (!Number.isFinite(posN)) continue;
        const colorEl = s.getElementsByTagNameNS(NS_MAIN, 'color').item(0);
        const color = parseColorElement(colorEl);
        stops.push({ position: Math.max(0, Math.min(1, posN)), color });
    }
    return { kind: 'gradient', type, degree, stops };
}

function parseBorders(doc: Document): BorderStyle[] {
    const bordersRoot = doc.getElementsByTagNameNS(NS_MAIN, 'borders').item(0);
    const out: BorderStyle[] = [];
    if (!bordersRoot) return out;
    const borders = bordersRoot.getElementsByTagNameNS(NS_MAIN, 'border');
    for (let i = 0; i < borders.length; i++) out.push(parseBorder(borders[i]));
    return out;
}

function parseBorder(el: Element): BorderStyle {
    const side = (tag: string): BorderSide => {
        const node = el.getElementsByTagNameNS(NS_MAIN, tag).item(0);
        if (!node) return { style: null, color: null };
        const style = node.getAttribute('style');
        const color = node.getElementsByTagNameNS(NS_MAIN, 'color').item(0);
        return { style, color: parseColorElement(color) };
    };
    return {
        left: side('left'),
        right: side('right'),
        top: side('top'),
        bottom: side('bottom'),
        diagonal: side('diagonal'),
        diagonalUp: el.getAttribute('diagonalUp') === '1',
        diagonalDown: el.getAttribute('diagonalDown') === '1',
    };
}

function parseCellXfs(doc: Document, tag: 'cellXfs' | 'cellStyleXfs'): CellXf[] {
    const root = doc.getElementsByTagNameNS(NS_MAIN, tag).item(0);
    const out: CellXf[] = [];
    if (!root) return out;
    const xfs = root.getElementsByTagNameNS(NS_MAIN, 'xf');
    for (let i = 0; i < xfs.length; i++) out.push(parseCellXf(xfs[i]));
    return out;
}

function parseDxfs(doc: Document): Dxf[] {
    const out: Dxf[] = [];
    const root = doc.getElementsByTagNameNS(NS_MAIN, 'dxfs').item(0);
    if (!root) return out;
    const dxfs = root.getElementsByTagNameNS(NS_MAIN, 'dxf');
    for (let i = 0; i < dxfs.length; i++) out.push(parseDxf(dxfs[i]));
    return out;
}

function parseDxf(el: Element): Dxf {
    const fontEl = el.getElementsByTagNameNS(NS_MAIN, 'font').item(0);
    const fillEl = el.getElementsByTagNameNS(NS_MAIN, 'fill').item(0);
    const borderEl = el.getElementsByTagNameNS(NS_MAIN, 'border').item(0);
    const numFmtEl = el.getElementsByTagNameNS(NS_MAIN, 'numFmt').item(0);

    const out: Dxf = {};
    if (fontEl) {
        // dxf fonts are partial — parseFont fills in defaults we don't want
        // to apply blindly, so pick out just the flags that are present.
        const has = (tag: string) => fontEl.getElementsByTagNameNS(NS_MAIN, tag).length > 0;
        const colorEl = fontEl.getElementsByTagNameNS(NS_MAIN, 'color').item(0);
        const sz = fontEl.getElementsByTagNameNS(NS_MAIN, 'sz').item(0);
        const font: Partial<FontStyle> = {};
        if (has('b')) font.bold = true;
        if (has('i')) font.italic = true;
        if (has('u')) font.underline = parseUnderline(fontEl);
        if (has('strike')) font.strike = true;
        const va = parseVertAlign(fontEl);
        if (va) font.vertAlign = va;
        if (sz?.getAttribute('val')) font.size = Number(sz.getAttribute('val')) || undefined as any;
        const c = parseColorElement(colorEl);
        if (c) font.color = c;
        if (Object.keys(font).length > 0) out.font = font;
    }
    if (fillEl) out.fill = parseFill(fillEl, { allowBgFallback: true });
    if (borderEl) out.border = parseBorder(borderEl);
    if (numFmtEl) {
        const code = numFmtEl.getAttribute('formatCode');
        if (code) out.numFmtCode = code;
    }
    return out;
}

function parseCellXf(el: Element): CellXf {
    const bool = (attr: string) => el.getAttribute(attr) === '1';
    const num = (attr: string) => Number(el.getAttribute(attr) ?? '0') | 0;
    const alignEl = el.getElementsByTagNameNS(NS_MAIN, 'alignment').item(0);
    const xfIdAttr = el.getAttribute('xfId');
    const xfId = xfIdAttr != null && Number.isFinite(Number(xfIdAttr)) ? Number(xfIdAttr) : -1;
    return {
        numFmtId: num('numFmtId'),
        fontId: num('fontId'),
        fillId: num('fillId'),
        borderId: num('borderId'),
        xfId,
        applyNumberFormat: bool('applyNumberFormat'),
        applyFont: bool('applyFont'),
        applyFill: bool('applyFill'),
        applyBorder: bool('applyBorder'),
        applyAlignment: bool('applyAlignment'),
        alignment: parseAlignment(alignEl),
    };
}

// Parse an <alignment/> child of an <xf>. Values are validated against a
// strict enum per ECMA-376 §18.8.1; anything unrecognised falls back to
// null (horizontal/vertical) or the neutral default (booleans, numbers).
function parseAlignment(el: Element | null): Alignment {
    if (!el) return defaultAlignment();
    const h = el.getAttribute('horizontal');
    const v = el.getAttribute('vertical');
    const horizontal: Alignment['horizontal'] =
        h === 'left' || h === 'right' || h === 'center' || h === 'justify' ||
        h === 'distributed' || h === 'centerContinuous' || h === 'fill' ? h : null;
    const vertical: Alignment['vertical'] =
        v === 'top' || v === 'center' || v === 'bottom' ||
        v === 'justify' || v === 'distributed' ? v : null;
    const indentAttr = el.getAttribute('indent');
    const indent = indentAttr != null && Number.isFinite(Number(indentAttr)) ? Math.max(0, Math.floor(Number(indentAttr))) : 0;
    const rotAttr = el.getAttribute('textRotation');
    let textRotation: number | null = null;
    if (rotAttr != null) {
        const n = Number(rotAttr);
        // Per spec: 0..180 inclusive is a rotation angle; 255 is the
        // stacked-vertical marker. Anything else is ignored.
        if (Number.isFinite(n) && ((n >= 0 && n <= 180) || n === 255)) {
            textRotation = Math.round(n);
        }
    }
    const roAttr = el.getAttribute('readingOrder');
    let readingOrder: 0 | 1 | 2 = 0;
    if (roAttr === '1') readingOrder = 1;
    else if (roAttr === '2') readingOrder = 2;
    return {
        horizontal,
        vertical,
        wrapText: el.getAttribute('wrapText') === '1',
        shrinkToFit: el.getAttribute('shrinkToFit') === '1',
        indent,
        textRotation,
        readingOrder,
    };
}

// Merge a cellXf with its referenced cellStyleXf base. Per ECMA-376, the
// cellXf's apply* flags control which categories override the base; the
// base provides the defaults for any category whose cellXf apply-flag is
// false. Resolution:
//   - numFmtId / fontId / fillId / borderId from cellXf when applyX=true,
//     otherwise from the base.
//   - alignment: cellXf when applyAlignment=true, otherwise base.
//
// A named style (cellStyleXf) may itself carry an `xfId` pointing at yet
// another cellStyleXf — rare in practice, but legal per the schema. We
// walk the chain iteratively so a pathological file can't blow the stack,
// capping at MAX_XF_HOPS hops and bailing on any revisited index.
const MAX_XF_HOPS = 8;

export function resolveEffectiveXf(styles: Styles, cellXf: CellXf): CellXf {
    let effective = cellXf;
    const seen = new Set<number>();
    for (let hop = 0; hop < MAX_XF_HOPS; hop++) {
        const nextId = effective.xfId;
        if (nextId < 0 || nextId >= styles.cellStyleXfs.length) break;
        // Cycle guard: if we've already merged with this cellStyleXf index,
        // bail — the chain loops back on itself and there's nothing new to
        // pick up.
        if (seen.has(nextId)) break;
        seen.add(nextId);
        const base = styles.cellStyleXfs[nextId];
        const merged: CellXf = {
            numFmtId: effective.applyNumberFormat ? effective.numFmtId : (base.numFmtId || effective.numFmtId),
            fontId:   effective.applyFont         ? effective.fontId   : (base.fontId   || effective.fontId),
            fillId:   effective.applyFill         ? effective.fillId   : (base.fillId   || effective.fillId),
            borderId: effective.applyBorder       ? effective.borderId : (base.borderId || effective.borderId),
            // Advance the chain: the base's own xfId tells us whether another
            // hop is needed on the next iteration. We preserve it here rather
            // than pinning to the original cellXf.xfId so the loop can
            // terminate when the chain reaches a terminal (-1) entry.
            xfId: base.xfId,
            // Effective xf always carries the "apply" flags as true when the
            // resulting id is non-default — the renderer uses those to decide
            // whether to call applyFont/applyFill/applyBorder at all.
            applyNumberFormat: effective.applyNumberFormat || base.applyNumberFormat,
            applyFont:         effective.applyFont         || base.applyFont,
            applyFill:         effective.applyFill         || base.applyFill,
            applyBorder:       effective.applyBorder       || base.applyBorder,
            applyAlignment:    effective.applyAlignment    || base.applyAlignment,
            alignment: effective.applyAlignment ? effective.alignment : base.alignment,
        };
        effective = merged;
    }
    // The returned xf keeps the *original* xfId so downstream code can
    // still see which named style this cell was rooted at, matching the
    // pre-chain contract. (The intermediate traversal used `base.xfId`
    // internally to walk the chain.)
    if (effective !== cellXf) {
        effective = { ...effective, xfId: cellXf.xfId };
    }
    return effective;
}
