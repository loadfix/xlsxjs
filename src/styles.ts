// Parse xl/styles.xml into a small, renderer-friendly model. Everything here
// is pure XML → typed data; no DOM, no formatting decisions. Colour elements
// are parsed into ColorRef (rgb or theme+tint) so the renderer can resolve
// them against xl/theme/theme1.xml when it has it.
//
// Scope:
//   - <numFmts>          custom number format strings (+ built-in lookup)
//   - <fonts>            bold, italic, underline, size, colour, name
//   - <fills>            patternFill solid fills (fgColor)
//   - <borders>          per-side style + colour
//   - <cellXfs>          the indexed array a cell's s= attribute points at
//
// Not yet: named styles, cellStyleXfs inheritance, tableStyles, dxfs,
// gradient fills, double / diagonal borders beyond the style name,
// indexed colours (the legacy 64-entry palette).

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
}

export interface FillStyle {
    // Only solid patternFill is rendered for now. Any other pattern type
    // returns null from resolveFill so cells stay unfilled.
    fgColor: ColorRef;
}

export interface BorderSide {
    style: string | null; // 'thin' | 'medium' | 'thick' | 'dashed' | ...
    color: ColorRef;
}

export interface BorderStyle {
    left: BorderSide;
    right: BorderSide;
    top: BorderSide;
    bottom: BorderSide;
}

export interface Alignment {
    horizontal: 'left' | 'right' | 'center' | 'justify' | 'distributed' | 'centerContinuous' | 'fill' | null;
    vertical: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed' | null;
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
    };
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

function parseFill(el: Element, opts?: { allowBgFallback?: boolean }): FillStyle {
    const pf = el.getElementsByTagNameNS(NS_MAIN, 'patternFill').item(0);
    if (!pf) return { fgColor: null };
    const patternType = pf.getAttribute('patternType');
    // For solid fills, fgColor is the canonical fill colour. dxf fills, the
    // legacy openpyxl convention, and some older Excel writers put the
    // colour on bgColor instead — allow the caller to opt into that fallback.
    const strictSolid = patternType === 'solid';
    if (!strictSolid && !opts?.allowBgFallback) return { fgColor: null };
    const fg = pf.getElementsByTagNameNS(NS_MAIN, 'fgColor').item(0);
    const fgRef = parseColorElement(fg);
    if (fgRef) return { fgColor: fgRef };
    if (opts?.allowBgFallback) {
        const bg = pf.getElementsByTagNameNS(NS_MAIN, 'bgColor').item(0);
        return { fgColor: parseColorElement(bg) };
    }
    return { fgColor: null };
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
    return { left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') };
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
        v === 'top' || v === 'middle' || v === 'bottom' ||
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
export function resolveEffectiveXf(styles: Styles, cellXf: CellXf): CellXf {
    if (cellXf.xfId < 0 || cellXf.xfId >= styles.cellStyleXfs.length) return cellXf;
    const base = styles.cellStyleXfs[cellXf.xfId];
    return {
        numFmtId: cellXf.applyNumberFormat ? cellXf.numFmtId : (base.numFmtId || cellXf.numFmtId),
        fontId:   cellXf.applyFont         ? cellXf.fontId   : (base.fontId   || cellXf.fontId),
        fillId:   cellXf.applyFill         ? cellXf.fillId   : (base.fillId   || cellXf.fillId),
        borderId: cellXf.applyBorder       ? cellXf.borderId : (base.borderId || cellXf.borderId),
        xfId: cellXf.xfId,
        // Effective xf always carries the "apply" flags as true when the
        // resulting id is non-default — the renderer uses those to decide
        // whether to call applyFont/applyFill/applyBorder at all.
        applyNumberFormat: cellXf.applyNumberFormat || base.applyNumberFormat,
        applyFont:         cellXf.applyFont         || base.applyFont,
        applyFill:         cellXf.applyFill         || base.applyFill,
        applyBorder:       cellXf.applyBorder       || base.applyBorder,
        applyAlignment:    cellXf.applyAlignment    || base.applyAlignment,
        alignment: cellXf.applyAlignment ? cellXf.alignment : base.alignment,
    };
}
