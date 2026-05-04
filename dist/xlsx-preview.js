/*
 * @license
 * xlsx-preview
 * Released under Apache License 2.0
 */
(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports) :
    typeof define === 'function' && define.amd ? define(['exports'], factory) :
    (global = typeof globalThis !== 'undefined' ? globalThis : global || self, factory(global.xlsx = {}));
})(this, (function (exports) { 'use strict';

    class XlsxEncryptedError extends Error {
        constructor() {
            super('xlsx-preview: this file is encrypted (OLE CFB container). xlsxjs does not decrypt — remove the password in Excel and re-save.');
            this.name = 'XlsxEncryptedError';
        }
    }
    const OLE_CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];
    async function bytesOf(data) {
        if (data instanceof Uint8Array)
            return data;
        if (data instanceof ArrayBuffer)
            return new Uint8Array(data);
        if (typeof data?.arrayBuffer === 'function') {
            return new Uint8Array(await data.arrayBuffer());
        }
        return null;
    }
    function isOleCfb(bytes) {
        if (!bytes || bytes.length < 4)
            return false;
        return OLE_CFB_MAGIC.every((b, i) => bytes[i] === b);
    }
    class Workbook {
        constructor() {
            this.parts = {};
            this.media = {};
            this.parsed = null;
        }
        static async load(data, parser) {
            const wb = new Workbook();
            const bytes = await bytesOf(data);
            if (isOleCfb(bytes))
                throw new XlsxEncryptedError();
            const zip = await JSZip.loadAsync(data);
            const readIfPresent = async (path) => {
                const f = zip.file(path);
                return f ? await f.async('string') : null;
            };
            const workbookXml = await readIfPresent('xl/workbook.xml');
            if (!workbookXml)
                throw new Error('xlsx-preview: missing xl/workbook.xml');
            wb.parts['xl/workbook.xml'] = workbookXml;
            const workbookRels = await readIfPresent('xl/_rels/workbook.xml.rels');
            if (workbookRels)
                wb.parts['xl/_rels/workbook.xml.rels'] = workbookRels;
            const sharedStrings = await readIfPresent('xl/sharedStrings.xml');
            if (sharedStrings)
                wb.parts['xl/sharedStrings.xml'] = sharedStrings;
            const styles = await readIfPresent('xl/styles.xml');
            if (styles)
                wb.parts['xl/styles.xml'] = styles;
            const metadata = await readIfPresent('xl/metadata.xml');
            if (metadata)
                wb.parts['xl/metadata.xml'] = metadata;
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/theme\/theme\d+\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/worksheets\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/worksheets\/_rels\/.*\.xml\.rels$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/tables\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/drawings\/.*\.xml$/i.test(p) ||
                    /^xl\/drawings\/_rels\/.*\.xml\.rels$/i.test(p) ||
                    /^xl\/drawings\/.*\.vml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/comments\d*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/charts\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/pivotTables\/.*\.xml$/i.test(p) ||
                    /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(p) ||
                    /^xl\/slicers\/.*\.xml$/i.test(p) ||
                    /^xl\/slicerCaches\/.*\.xml$/i.test(p) ||
                    /^xl\/timelines\/.*\.xml$/i.test(p) ||
                    /^xl\/timelineCaches\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/threadedComments\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/persons\/.*\.xml$/i.test(p)) {
                    const xml = await readIfPresent(p);
                    if (xml)
                        wb.parts[p] = xml;
                }
            }
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/media\/[^/]+$/i.test(p)) {
                    const bin = zip.file(p);
                    if (!bin)
                        continue;
                    const base64 = await bin.async('base64');
                    const mime = guessMime(p);
                    wb.media[p] = `data:${mime};base64,${base64}`;
                }
            }
            wb.parsed = parser.parse(wb.parts, wb.media);
            return wb;
        }
    }
    function guessMime(path) {
        const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
        switch (ext) {
            case 'png': return 'image/png';
            case 'jpg':
            case 'jpeg': return 'image/jpeg';
            case 'gif': return 'image/gif';
            case 'svg': return 'image/svg+xml';
            case 'webp': return 'image/webp';
            case 'bmp': return 'image/bmp';
            case 'tif':
            case 'tiff': return 'image/tiff';
            default: return 'application/octet-stream';
        }
    }

    function columnLettersToIndex(letters) {
        let n = 0;
        for (let i = 0; i < letters.length; i++) {
            const c = letters.charCodeAt(i);
            if (c < 0x41 || c > 0x5a)
                return -1;
            n = n * 26 + (c - 0x40);
        }
        return n - 1;
    }
    function indexToColumnLetters(index) {
        let n = index + 1;
        let s = '';
        while (n > 0) {
            const r = (n - 1) % 26;
            s = String.fromCharCode(0x41 + r) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    }
    const CELL_REF = /^([A-Z]+)([1-9][0-9]*)$/;
    function parseCellRef(ref) {
        const m = CELL_REF.exec(ref);
        if (!m)
            return null;
        const col = columnLettersToIndex(m[1]);
        if (col < 0)
            return null;
        return { col, row: Number(m[2]) - 1 };
    }
    function emuToPx(emu) {
        if (!Number.isFinite(emu))
            return 0;
        return Math.round(emu / 9525);
    }

    const NS_MAIN$1 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    function defaultAlignment() {
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
    const BUILTIN_NUMBER_FORMATS = {
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
    function lookupNumberFormat(styles, numFmtId) {
        if (styles?.numFmts.has(numFmtId))
            return styles.numFmts.get(numFmtId);
        if (numFmtId in BUILTIN_NUMBER_FORMATS)
            return BUILTIN_NUMBER_FORMATS[numFmtId];
        return null;
    }
    const HEX_COLOR = /^([A-Fa-f0-9]{6})([A-Fa-f0-9]{2})?$/;
    const ARGB_COLOR = /^([A-Fa-f0-9]{2})([A-Fa-f0-9]{6})$/;
    const SAFE_FONT_FAMILY = /^[A-Za-z0-9 .\-]+$/;
    function sanitizeFontFamily(name) {
        if (!name)
            return null;
        const trimmed = name.trim();
        if (!trimmed)
            return null;
        if (/[;{}<>\n\r]/.test(trimmed))
            return null;
        if (!SAFE_FONT_FAMILY.test(trimmed))
            return null;
        return `"${trimmed}"`;
    }
    function sanitizeHexColor(value) {
        if (!value)
            return null;
        const trimmed = value.trim();
        const argb = ARGB_COLOR.exec(trimmed);
        if (argb)
            return `#${argb[2].toLowerCase()}`;
        const rgb = HEX_COLOR.exec(trimmed);
        if (rgb)
            return `#${rgb[1].toLowerCase()}`;
        return null;
    }
    function parseColorElement(el) {
        if (!el)
            return null;
        const tintAttr = el.getAttribute('tint');
        const tint = tintAttr ? Number(tintAttr) : 0;
        const safeTint = Number.isFinite(tint) ? tint : 0;
        const rgb = el.getAttribute('rgb');
        if (rgb) {
            const sanitized = sanitizeHexColor(rgb);
            if (sanitized)
                return { kind: 'rgb', value: sanitized };
            return null;
        }
        const themeAttr = el.getAttribute('theme');
        if (themeAttr != null) {
            const index = Number(themeAttr);
            if (!Number.isFinite(index) || index < 0)
                return null;
            return { kind: 'theme', index, tint: safeTint };
        }
        const indexedAttr = el.getAttribute('indexed');
        if (indexedAttr != null) {
            const index = Number(indexedAttr);
            if (!Number.isFinite(index) || index < 0)
                return null;
            return { kind: 'indexed', index, tint: safeTint };
        }
        return null;
    }
    function parseXml$2(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function parseStyles(xml) {
        const doc = parseXml$2(xml);
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
    function parseNumFmts(doc) {
        const out = new Map();
        const els = doc.getElementsByTagNameNS(NS_MAIN$1, 'numFmt');
        for (let i = 0; i < els.length; i++) {
            const id = Number(els[i].getAttribute('numFmtId'));
            const code = els[i].getAttribute('formatCode') ?? '';
            if (Number.isFinite(id) && id >= 0)
                out.set(id, code);
        }
        return out;
    }
    function parseFonts(doc) {
        const fontsRoot = doc.getElementsByTagNameNS(NS_MAIN$1, 'fonts').item(0);
        const out = [];
        if (!fontsRoot)
            return out;
        const fonts = fontsRoot.getElementsByTagNameNS(NS_MAIN$1, 'font');
        for (let i = 0; i < fonts.length; i++)
            out.push(parseFont(fonts[i]));
        return out;
    }
    function parseFont(el) {
        const has = (tag) => el.getElementsByTagNameNS(NS_MAIN$1, tag).length > 0;
        const firstAttr = (tag, attr) => {
            const node = el.getElementsByTagNameNS(NS_MAIN$1, tag).item(0);
            return node ? node.getAttribute(attr) : null;
        };
        const colorEl = el.getElementsByTagNameNS(NS_MAIN$1, 'color').item(0);
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
    function parseScheme(parent) {
        const s = parent.getElementsByTagNameNS(NS_MAIN$1, 'scheme').item(0);
        if (!s)
            return null;
        const val = s.getAttribute('val');
        if (val === 'major' || val === 'minor')
            return val;
        return null;
    }
    function parseUnderline(parent) {
        const u = parent.getElementsByTagNameNS(NS_MAIN$1, 'u').item(0);
        if (!u)
            return null;
        const val = u.getAttribute('val');
        if (val === null || val === '')
            return 'single';
        if (val === 'single' || val === 'double' || val === 'singleAccounting' || val === 'doubleAccounting') {
            return val;
        }
        if (val === 'none')
            return null;
        return null;
    }
    function parseVertAlign(parent) {
        const va = parent.getElementsByTagNameNS(NS_MAIN$1, 'vertAlign').item(0);
        if (!va)
            return null;
        const val = va.getAttribute('val');
        if (val === 'subscript' || val === 'superscript')
            return val;
        return null;
    }
    function parseFills(doc) {
        const fillsRoot = doc.getElementsByTagNameNS(NS_MAIN$1, 'fills').item(0);
        const out = [];
        if (!fillsRoot)
            return out;
        const fills = fillsRoot.getElementsByTagNameNS(NS_MAIN$1, 'fill');
        for (let i = 0; i < fills.length; i++)
            out.push(parseFill(fills[i]));
        return out;
    }
    function parseFill(el, opts) {
        const pf = el.getElementsByTagNameNS(NS_MAIN$1, 'patternFill').item(0);
        if (!pf)
            return { fgColor: null };
        const patternType = pf.getAttribute('patternType');
        const strictSolid = patternType === 'solid';
        if (!strictSolid && !opts?.allowBgFallback)
            return { fgColor: null };
        const fg = pf.getElementsByTagNameNS(NS_MAIN$1, 'fgColor').item(0);
        const fgRef = parseColorElement(fg);
        if (fgRef)
            return { fgColor: fgRef };
        if (opts?.allowBgFallback) {
            const bg = pf.getElementsByTagNameNS(NS_MAIN$1, 'bgColor').item(0);
            return { fgColor: parseColorElement(bg) };
        }
        return { fgColor: null };
    }
    function parseBorders(doc) {
        const bordersRoot = doc.getElementsByTagNameNS(NS_MAIN$1, 'borders').item(0);
        const out = [];
        if (!bordersRoot)
            return out;
        const borders = bordersRoot.getElementsByTagNameNS(NS_MAIN$1, 'border');
        for (let i = 0; i < borders.length; i++)
            out.push(parseBorder(borders[i]));
        return out;
    }
    function parseBorder(el) {
        const side = (tag) => {
            const node = el.getElementsByTagNameNS(NS_MAIN$1, tag).item(0);
            if (!node)
                return { style: null, color: null };
            const style = node.getAttribute('style');
            const color = node.getElementsByTagNameNS(NS_MAIN$1, 'color').item(0);
            return { style, color: parseColorElement(color) };
        };
        return { left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') };
    }
    function parseCellXfs(doc, tag) {
        const root = doc.getElementsByTagNameNS(NS_MAIN$1, tag).item(0);
        const out = [];
        if (!root)
            return out;
        const xfs = root.getElementsByTagNameNS(NS_MAIN$1, 'xf');
        for (let i = 0; i < xfs.length; i++)
            out.push(parseCellXf(xfs[i]));
        return out;
    }
    function parseDxfs(doc) {
        const out = [];
        const root = doc.getElementsByTagNameNS(NS_MAIN$1, 'dxfs').item(0);
        if (!root)
            return out;
        const dxfs = root.getElementsByTagNameNS(NS_MAIN$1, 'dxf');
        for (let i = 0; i < dxfs.length; i++)
            out.push(parseDxf(dxfs[i]));
        return out;
    }
    function parseDxf(el) {
        const fontEl = el.getElementsByTagNameNS(NS_MAIN$1, 'font').item(0);
        const fillEl = el.getElementsByTagNameNS(NS_MAIN$1, 'fill').item(0);
        const borderEl = el.getElementsByTagNameNS(NS_MAIN$1, 'border').item(0);
        const numFmtEl = el.getElementsByTagNameNS(NS_MAIN$1, 'numFmt').item(0);
        const out = {};
        if (fontEl) {
            const has = (tag) => fontEl.getElementsByTagNameNS(NS_MAIN$1, tag).length > 0;
            const colorEl = fontEl.getElementsByTagNameNS(NS_MAIN$1, 'color').item(0);
            const sz = fontEl.getElementsByTagNameNS(NS_MAIN$1, 'sz').item(0);
            const font = {};
            if (has('b'))
                font.bold = true;
            if (has('i'))
                font.italic = true;
            if (has('u'))
                font.underline = parseUnderline(fontEl);
            if (has('strike'))
                font.strike = true;
            const va = parseVertAlign(fontEl);
            if (va)
                font.vertAlign = va;
            if (sz?.getAttribute('val'))
                font.size = Number(sz.getAttribute('val')) || undefined;
            const c = parseColorElement(colorEl);
            if (c)
                font.color = c;
            if (Object.keys(font).length > 0)
                out.font = font;
        }
        if (fillEl)
            out.fill = parseFill(fillEl, { allowBgFallback: true });
        if (borderEl)
            out.border = parseBorder(borderEl);
        if (numFmtEl) {
            const code = numFmtEl.getAttribute('formatCode');
            if (code)
                out.numFmtCode = code;
        }
        return out;
    }
    function parseCellXf(el) {
        const bool = (attr) => el.getAttribute(attr) === '1';
        const num = (attr) => Number(el.getAttribute(attr) ?? '0') | 0;
        const alignEl = el.getElementsByTagNameNS(NS_MAIN$1, 'alignment').item(0);
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
    function parseAlignment(el) {
        if (!el)
            return defaultAlignment();
        const h = el.getAttribute('horizontal');
        const v = el.getAttribute('vertical');
        const horizontal = h === 'left' || h === 'right' || h === 'center' || h === 'justify' ||
            h === 'distributed' || h === 'centerContinuous' || h === 'fill' ? h : null;
        const vertical = v === 'top' || v === 'middle' || v === 'bottom' ||
            v === 'justify' || v === 'distributed' ? v : null;
        const indentAttr = el.getAttribute('indent');
        const indent = indentAttr != null && Number.isFinite(Number(indentAttr)) ? Math.max(0, Math.floor(Number(indentAttr))) : 0;
        const rotAttr = el.getAttribute('textRotation');
        let textRotation = null;
        if (rotAttr != null) {
            const n = Number(rotAttr);
            if (Number.isFinite(n) && ((n >= 0 && n <= 180) || n === 255)) {
                textRotation = Math.round(n);
            }
        }
        const roAttr = el.getAttribute('readingOrder');
        let readingOrder = 0;
        if (roAttr === '1')
            readingOrder = 1;
        else if (roAttr === '2')
            readingOrder = 2;
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
    const MAX_XF_HOPS = 8;
    function resolveEffectiveXf(styles, cellXf) {
        let effective = cellXf;
        const seen = new Set();
        for (let hop = 0; hop < MAX_XF_HOPS; hop++) {
            const nextId = effective.xfId;
            if (nextId < 0 || nextId >= styles.cellStyleXfs.length)
                break;
            if (seen.has(nextId))
                break;
            seen.add(nextId);
            const base = styles.cellStyleXfs[nextId];
            const merged = {
                numFmtId: effective.applyNumberFormat ? effective.numFmtId : (base.numFmtId || effective.numFmtId),
                fontId: effective.applyFont ? effective.fontId : (base.fontId || effective.fontId),
                fillId: effective.applyFill ? effective.fillId : (base.fillId || effective.fillId),
                borderId: effective.applyBorder ? effective.borderId : (base.borderId || effective.borderId),
                xfId: base.xfId,
                applyNumberFormat: effective.applyNumberFormat || base.applyNumberFormat,
                applyFont: effective.applyFont || base.applyFont,
                applyFill: effective.applyFill || base.applyFill,
                applyBorder: effective.applyBorder || base.applyBorder,
                applyAlignment: effective.applyAlignment || base.applyAlignment,
                alignment: effective.applyAlignment ? effective.alignment : base.alignment,
            };
            effective = merged;
        }
        if (effective !== cellXf) {
            effective = { ...effective, xfId: cellXf.xfId };
        }
        return effective;
    }

    const NS_DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const INDEXED_PALETTE = [
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
    function indexedColor(index) {
        if (!Number.isFinite(index) || index < 0 || index >= INDEXED_PALETTE.length)
            return null;
        return INDEXED_PALETTE[index];
    }
    const HEX_6 = /^([A-Fa-f0-9]{6})$/;
    const HEX_ARGB = /^([A-Fa-f0-9]{2})([A-Fa-f0-9]{6})$/;
    function rgbToHex(value) {
        if (!value)
            return null;
        const trimmed = value.trim();
        const argb = HEX_ARGB.exec(trimmed);
        if (argb)
            return `#${argb[2].toLowerCase()}`;
        const rgb = HEX_6.exec(trimmed);
        if (rgb)
            return `#${rgb[1].toLowerCase()}`;
        return null;
    }
    function parseXml$1(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function parseTheme(xml) {
        const colors = new Array(12).fill(null);
        colors[0] = '#ffffff';
        colors[1] = '#000000';
        colors[2] = '#e7e6e6';
        colors[3] = '#44546a';
        const doc = parseXml$1(xml);
        const { majorFont, minorFont } = parseFontScheme(doc);
        const scheme = doc.getElementsByTagNameNS(NS_DRAW, 'clrScheme').item(0);
        if (!scheme)
            return { colors, majorFont, minorFont };
        const get = (localName) => {
            const el = scheme.getElementsByTagNameNS(NS_DRAW, localName).item(0);
            if (!el)
                return null;
            return resolveClrChild(el);
        };
        const mapping = [
            ['lt1', 0], ['dk1', 1], ['lt2', 2], ['dk2', 3],
            ['accent1', 4], ['accent2', 5], ['accent3', 6], ['accent4', 7],
            ['accent5', 8], ['accent6', 9], ['hlink', 10], ['folHlink', 11],
        ];
        for (const [name, idx] of mapping) {
            const c = get(name);
            if (c)
                colors[idx] = c;
        }
        return { colors, majorFont, minorFont };
    }
    function parseFontScheme(doc) {
        const fs = doc.getElementsByTagNameNS(NS_DRAW, 'fontScheme').item(0);
        if (!fs)
            return { majorFont: null, minorFont: null };
        const latinFrom = (parentLocalName) => {
            const parent = fs.getElementsByTagNameNS(NS_DRAW, parentLocalName).item(0);
            if (!parent)
                return null;
            const latin = parent.getElementsByTagNameNS(NS_DRAW, 'latin').item(0);
            if (!latin)
                return null;
            const typeface = latin.getAttribute('typeface');
            if (!typeface)
                return null;
            return typeface;
        };
        return { majorFont: latinFrom('majorFont'), minorFont: latinFrom('minorFont') };
    }
    function resolveClrChild(wrapper) {
        const srgb = wrapper.getElementsByTagNameNS(NS_DRAW, 'srgbClr').item(0);
        if (srgb)
            return rgbToHex(srgb.getAttribute('val'));
        const sysClr = wrapper.getElementsByTagNameNS(NS_DRAW, 'sysClr').item(0);
        if (sysClr) {
            const last = sysClr.getAttribute('lastClr');
            if (last)
                return rgbToHex(last);
            const val = sysClr.getAttribute('val');
            if (val === 'windowText')
                return '#000000';
            if (val === 'window')
                return '#ffffff';
        }
        return null;
    }
    function hexToRgb(hex) {
        const m = /^#([0-9a-f]{6})$/i.exec(hex);
        if (!m)
            return null;
        return [
            parseInt(m[1].slice(0, 2), 16),
            parseInt(m[1].slice(2, 4), 16),
            parseInt(m[1].slice(4, 6), 16),
        ];
    }
    function rgbToHexStr(r, g, b) {
        const hex = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
        return `#${hex(r)}${hex(g)}${hex(b)}`;
    }
    function rgbToHsl(r, g, b) {
        r /= 255;
        g /= 255;
        b /= 255;
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h = 0, s = 0;
        const l = (max + min) / 2;
        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r:
                    h = ((g - b) / d) + (g < b ? 6 : 0);
                    break;
                case g:
                    h = ((b - r) / d) + 2;
                    break;
                case b:
                    h = ((r - g) / d) + 4;
                    break;
            }
            h /= 6;
        }
        return [h, s, l];
    }
    function hslToRgb(h, s, l) {
        if (s === 0)
            return [l * 255, l * 255, l * 255];
        const hue2rgb = (p, q, t) => {
            if (t < 0)
                t += 1;
            if (t > 1)
                t -= 1;
            if (t < 1 / 6)
                return p + (q - p) * 6 * t;
            if (t < 1 / 2)
                return q;
            if (t < 2 / 3)
                return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        };
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        const r = hue2rgb(p, q, h + 1 / 3);
        const g = hue2rgb(p, q, h);
        const b = hue2rgb(p, q, h - 1 / 3);
        return [r * 255, g * 255, b * 255];
    }
    function applyTint(hex, tint) {
        if (!tint || !Number.isFinite(tint))
            return hex;
        const rgb = hexToRgb(hex);
        if (!rgb)
            return hex;
        const [h, s, l] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
        const clamped = Math.max(-1, Math.min(1, tint));
        const newL = clamped < 0 ? l * (1 + clamped) : l + (1 - l) * clamped;
        const [r, g, b] = hslToRgb(h, s, newL);
        return rgbToHexStr(r, g, b);
    }
    function resolveColor(ref, theme) {
        if (!ref)
            return null;
        if (ref.kind === 'rgb')
            return ref.value;
        if (ref.kind === 'theme') {
            const base = theme?.colors[ref.index] ?? null;
            if (!base)
                return null;
            return ref.tint ? applyTint(base, ref.tint) : base;
        }
        if (ref.kind === 'indexed') {
            const base = indexedColor(ref.index);
            if (!base)
                return null;
            return ref.tint ? applyTint(base, ref.tint) : base;
        }
        return null;
    }

    const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    function parseConditionalFormatting(doc) {
        const out = [];
        const blocks = doc.getElementsByTagNameNS(NS_MAIN, 'conditionalFormatting');
        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];
            const sqref = block.getAttribute('sqref') ?? '';
            const ranges = parseSqref(sqref);
            const rules = [];
            const ruleEls = block.getElementsByTagNameNS(NS_MAIN, 'cfRule');
            for (let j = 0; j < ruleEls.length; j++) {
                const rule = parseCfRule(ruleEls[j]);
                if (rule)
                    rules.push(rule);
            }
            out.push({ ranges, rules });
        }
        mergeDataBarExtAttrs(doc, out);
        return out;
    }
    function mergeDataBarExtAttrs(doc, blocks) {
        const byExtId = new Map();
        for (const block of blocks) {
            for (const rule of block.rules) {
                if (rule.type === 'dataBar' && rule.dataBar?.extId) {
                    byExtId.set(rule.dataBar.extId, rule.dataBar);
                }
            }
        }
        if (byExtId.size === 0)
            return;
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'cfRule')
                continue;
            if (el.getAttribute('type') !== 'dataBar')
                continue;
            const id = el.getAttribute('id');
            if (!id)
                continue;
            const target = byExtId.get(id);
            if (!target)
                continue;
            applyDataBarExt(el, target);
        }
    }
    function applyDataBarExt(x14rule, bar) {
        const children = x14rule.getElementsByTagName('*');
        let dataBarEl = null;
        for (let i = 0; i < children.length; i++) {
            if (children[i].localName === 'dataBar') {
                dataBarEl = children[i];
                break;
            }
        }
        if (!dataBarEl)
            return;
        const minAttr = dataBarEl.getAttribute('minLength');
        if (minAttr != null) {
            const n = Number(minAttr);
            if (Number.isFinite(n))
                bar.minLength = n;
        }
        const maxAttr = dataBarEl.getAttribute('maxLength');
        if (maxAttr != null) {
            const n = Number(maxAttr);
            if (Number.isFinite(n))
                bar.maxLength = n;
        }
        if (dataBarEl.getAttribute('border') === '1')
            bar.border = true;
        if (dataBarEl.getAttribute('gradient') === '0')
            bar.gradient = false;
        const axis = dataBarEl.getAttribute('axisPosition');
        if (axis === 'middle' || axis === 'none' || axis === 'automatic')
            bar.axisPosition = axis;
        const dir = dataBarEl.getAttribute('direction');
        if (dir === 'leftToRight' || dir === 'rightToLeft' || dir === 'context')
            bar.direction = dir;
        for (let i = 0; i < children.length; i++) {
            const c = children[i];
            switch (c.localName) {
                case 'borderColor':
                    bar.borderColor = parseCfColor(c);
                    break;
                case 'negativeFillColor':
                    bar.negativeFillColor = parseCfColor(c);
                    break;
                case 'negativeBorderColor':
                    bar.negativeBorderColor = parseCfColor(c);
                    break;
                case 'axisColor':
                    bar.axisColor = parseCfColor(c);
                    break;
            }
        }
    }
    function parseSqref(sqref) {
        const parts = sqref.trim().split(/\s+/).filter((p) => p.length);
        const out = [];
        for (const p of parts) {
            const r = parseRangeToken(p);
            if (r)
                out.push(r);
        }
        return out;
    }
    function parseRangeToken(tok) {
        const parts = tok.split(':');
        if (parts.length === 1) {
            const a = parseCellRef(parts[0].replace(/\$/g, ''));
            if (!a)
                return null;
            return { col: a.col, row: a.row, endCol: a.col, endRow: a.row };
        }
        const a = parseCellRef(parts[0].replace(/\$/g, ''));
        const b = parseCellRef(parts[1].replace(/\$/g, ''));
        if (!a || !b)
            return null;
        return {
            col: Math.min(a.col, b.col),
            row: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row),
        };
    }
    function parseCfRule(el) {
        const rawType = el.getAttribute('type') ?? '';
        const type = asRuleType(rawType);
        const priority = Number(el.getAttribute('priority')) || 0;
        const dxfIdAttr = el.getAttribute('dxfId');
        const dxfId = dxfIdAttr != null && Number.isFinite(Number(dxfIdAttr)) ? Number(dxfIdAttr) : -1;
        const formulas = [];
        const fEls = el.getElementsByTagNameNS(NS_MAIN, 'formula');
        for (let i = 0; i < fEls.length; i++)
            formulas.push(fEls[i].textContent ?? '');
        const operator = asOperator(el.getAttribute('operator'));
        const text = el.getAttribute('text') ?? undefined;
        const rank = el.getAttribute('rank') ? Number(el.getAttribute('rank')) : undefined;
        const percent = el.getAttribute('percent') === '1';
        const bottom = el.getAttribute('bottom') === '1';
        const stopIfTrue = el.getAttribute('stopIfTrue') === '1';
        const aboveAverageAttr = el.getAttribute('aboveAverage');
        const aboveAverage = aboveAverageAttr !== '0';
        const equalAverage = el.getAttribute('equalAverage') === '1';
        const stdDevAttr = el.getAttribute('stdDev');
        const stdDevN = stdDevAttr != null ? Number(stdDevAttr) : NaN;
        const stdDev = Number.isFinite(stdDevN) ? stdDevN : null;
        const timePeriod = el.getAttribute('timePeriod');
        const rule = {
            type, priority, dxfId, operator, formulas, text, rank, percent, bottom, stopIfTrue,
            aboveAverage, equalAverage, stdDev,
            timePeriod: timePeriod ?? null,
        };
        if (type === 'colorScale')
            rule.colorScale = parseColorScale(el);
        else if (type === 'dataBar')
            rule.dataBar = parseDataBar(el);
        else if (type === 'iconSet')
            rule.iconSet = parseIconSet(el);
        return rule;
    }
    function parseCfvos(parent) {
        const out = [];
        const els = parent.getElementsByTagNameNS(NS_MAIN, 'cfvo');
        for (let i = 0; i < els.length; i++) {
            const t = els[i].getAttribute('type') ?? 'num';
            const v = els[i].getAttribute('val');
            const type = (t === 'min' || t === 'max' || t === 'percent' || t === 'percentile' || t === 'num' || t === 'formula')
                ? t : 'num';
            out.push({ type, value: v });
        }
        return out;
    }
    function parseCfColor(el) {
        if (!el)
            return null;
        const tintAttr = el.getAttribute('tint');
        const tint = tintAttr ? Number(tintAttr) : 0;
        const safeTint = Number.isFinite(tint) ? tint : 0;
        const rgb = el.getAttribute('rgb');
        if (rgb) {
            const sanitized = sanitizeHexColor(rgb);
            return sanitized ? { kind: 'rgb', value: sanitized } : null;
        }
        const themeAttr = el.getAttribute('theme');
        if (themeAttr != null) {
            const index = Number(themeAttr);
            if (!Number.isFinite(index) || index < 0)
                return null;
            return { kind: 'theme', index, tint: safeTint };
        }
        const idxAttr = el.getAttribute('indexed');
        if (idxAttr != null) {
            const index = Number(idxAttr);
            if (!Number.isFinite(index) || index < 0)
                return null;
            return { kind: 'indexed', index, tint: safeTint };
        }
        return null;
    }
    function parseColorScale(ruleEl) {
        const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'colorScale').item(0);
        if (!root)
            return { cfvos: [], colors: [] };
        const cfvos = parseCfvos(root);
        const colors = [];
        const colorEls = root.getElementsByTagNameNS(NS_MAIN, 'color');
        for (let i = 0; i < colorEls.length; i++)
            colors.push(parseCfColor(colorEls[i]));
        return { cfvos, colors };
    }
    function dataBarDefaults() {
        return {
            cfvos: [], color: null, minLength: 10, maxLength: 90, showValue: true,
            border: false, borderColor: null,
            negativeFillColor: null, negativeBorderColor: null,
            axisPosition: 'automatic', axisColor: null,
            gradient: true, direction: 'context',
            extId: null,
        };
    }
    function parseDataBar(ruleEl) {
        const out = dataBarDefaults();
        const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'dataBar').item(0);
        if (!root)
            return { ...out, extId: parseCfRuleExtId(ruleEl) };
        out.cfvos = parseCfvos(root);
        out.color = parseCfColor(root.getElementsByTagNameNS(NS_MAIN, 'color').item(0));
        out.minLength = Number(root.getAttribute('minLength') ?? '10') || 10;
        out.maxLength = Number(root.getAttribute('maxLength') ?? '90') || 90;
        out.showValue = root.getAttribute('showValue') !== '0';
        out.extId = parseCfRuleExtId(ruleEl);
        return out;
    }
    function parseCfRuleExtId(ruleEl) {
        const extLst = ruleEl.getElementsByTagNameNS(NS_MAIN, 'extLst').item(0);
        if (!extLst)
            return null;
        const all = extLst.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'id')
                continue;
            const text = (el.textContent ?? '').trim();
            if (text)
                return text;
        }
        return null;
    }
    function parseIconSet(ruleEl) {
        const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'iconSet').item(0);
        if (!root)
            return {
                iconSet: '3TrafficLights1', cfvos: [], showValue: true, reverse: false,
                customIcons: null,
            };
        const cfvos = parseCfvos(root);
        const isCustom = root.getAttribute('custom') === '1';
        let customIcons = null;
        if (isCustom) {
            customIcons = new Array(cfvos.length).fill(null);
            const iconEls = root.getElementsByTagNameNS(NS_MAIN, 'cfIcon');
            for (let i = 0; i < iconEls.length && i < cfvos.length; i++) {
                const set = iconEls[i].getAttribute('iconSet');
                const idAttr = iconEls[i].getAttribute('iconId');
                const id = idAttr != null ? Number(idAttr) : NaN;
                if (!set || !Number.isFinite(id))
                    continue;
                customIcons[i] = { iconSet: set, iconId: id };
            }
        }
        return {
            iconSet: root.getAttribute('iconSet') ?? '3TrafficLights1',
            cfvos,
            showValue: root.getAttribute('showValue') !== '0',
            reverse: root.getAttribute('reverse') === '1',
            customIcons,
        };
    }
    function asRuleType(s) {
        switch (s) {
            case 'cellIs':
            case 'containsText':
            case 'notContainsText':
            case 'beginsWith':
            case 'endsWith':
            case 'duplicateValues':
            case 'uniqueValues':
            case 'top10':
            case 'expression':
            case 'containsBlanks':
            case 'notContainsBlanks':
            case 'containsErrors':
            case 'notContainsErrors':
            case 'aboveAverage':
            case 'timePeriod':
            case 'colorScale':
            case 'dataBar':
            case 'iconSet':
                return s;
            default:
                return 'unsupported';
        }
    }
    function asOperator(s) {
        if (!s)
            return undefined;
        switch (s) {
            case 'equal':
            case 'notEqual':
            case 'greaterThan':
            case 'greaterThanOrEqual':
            case 'lessThan':
            case 'lessThanOrEqual':
            case 'between':
            case 'notBetween':
                return s;
            default:
                return undefined;
        }
    }
    function evaluateRule(rule, cell, range, ctx) {
        if (rule.type === 'containsBlanks')
            return evalContainsBlanks(cell);
        if (rule.type === 'notContainsBlanks')
            return !evalContainsBlanks(cell);
        if (rule.type === 'containsErrors')
            return evalContainsErrors(cell);
        if (rule.type === 'notContainsErrors')
            return !evalContainsErrors(cell);
        if (!cell)
            return false;
        switch (rule.type) {
            case 'cellIs': return evalCellIs(rule, cell);
            case 'containsText': return containsText(cell, rule.text);
            case 'notContainsText': return !containsText(cell, rule.text);
            case 'beginsWith': return startsOrEnds(cell, rule.text, 'begin');
            case 'endsWith': return startsOrEnds(cell, rule.text, 'end');
            case 'duplicateValues': return evalDuplicate(cell, range, ctx, true);
            case 'uniqueValues': return evalDuplicate(cell, range, ctx, false);
            case 'top10': return evalTop10(rule, cell, range, ctx);
            case 'expression': return evalExpression(rule, cell);
            case 'aboveAverage': return evalAboveAverage(rule, cell, range, ctx);
            case 'timePeriod': return evalTimePeriod(rule, cell, ctx);
            case 'unsupported': return false;
        }
        return false;
    }
    function evalContainsBlanks(cell) {
        if (!cell)
            return true;
        if (cell.kind === 'empty')
            return true;
        if (cell.value === '' || cell.value == null)
            return true;
        return false;
    }
    function evalContainsErrors(cell) {
        if (!cell)
            return false;
        return cell.kind === 'error';
    }
    function numericValue(cell) {
        if (cell.kind !== 'number' && cell.kind !== 'boolean') {
            const n = Number(cell.value);
            return Number.isFinite(n) ? n : null;
        }
        if (cell.kind === 'boolean')
            return cell.value === 'TRUE' ? 1 : 0;
        const n = Number(cell.value);
        return Number.isFinite(n) ? n : null;
    }
    function parseFormulaLiteral(text) {
        if (text === undefined)
            return null;
        const t = text.trim();
        if (t.startsWith('"') && t.endsWith('"'))
            return t.slice(1, -1);
        const n = Number(t);
        if (Number.isFinite(n))
            return n;
        return null;
    }
    function evalCellIs(rule, cell) {
        const a = parseFormulaLiteral(rule.formulas[0]);
        const b = parseFormulaLiteral(rule.formulas[1]);
        if (a === null)
            return false;
        if (typeof a === 'number') {
            const v = numericValue(cell);
            if (v === null)
                return false;
            switch (rule.operator) {
                case 'equal': return v === a;
                case 'notEqual': return v !== a;
                case 'greaterThan': return v > a;
                case 'greaterThanOrEqual': return v >= a;
                case 'lessThan': return v < a;
                case 'lessThanOrEqual': return v <= a;
                case 'between': return typeof b === 'number' && v >= a && v <= b;
                case 'notBetween': return typeof b === 'number' && (v < a || v > b);
                default: return false;
            }
        }
        const sv = (cell.value ?? '').toLocaleLowerCase();
        const sa = a.toLocaleLowerCase();
        switch (rule.operator) {
            case 'equal': return sv === sa;
            case 'notEqual': return sv !== sa;
            default: return false;
        }
    }
    function containsText(cell, text, _not) {
        if (!text)
            return false;
        return (cell.value ?? '').toLocaleLowerCase().includes(text.toLocaleLowerCase());
    }
    function startsOrEnds(cell, text, side) {
        if (!text)
            return false;
        const v = (cell.value ?? '').toLocaleLowerCase();
        const t = text.toLocaleLowerCase();
        return side === 'begin' ? v.startsWith(t) : v.endsWith(t);
    }
    function evalDuplicate(cell, range, ctx, findDuplicate) {
        const cells = ctx.cellsInRange(range);
        let seen = 0;
        for (const entry of cells) {
            const other = entry.cell;
            if (!other)
                continue;
            if (other.value === cell.value)
                seen++;
            if (seen > 1)
                break;
        }
        return findDuplicate ? seen > 1 : seen === 1;
    }
    function evalTop10(rule, cell, range, ctx) {
        const cells = ctx.cellsInRange(range);
        const values = [];
        for (const entry of cells) {
            if (!entry.cell)
                continue;
            const n = numericValue(entry.cell);
            if (n === null)
                continue;
            values.push({ cell: entry.cell, value: n });
        }
        if (values.length === 0)
            return false;
        values.sort((a, b) => rule.bottom ? a.value - b.value : b.value - a.value);
        const rank = rule.rank ?? 10;
        const n = rule.percent ? Math.max(1, Math.ceil(values.length * rank / 100)) : Math.min(rank, values.length);
        const threshold = values[n - 1]?.value;
        if (threshold === undefined)
            return false;
        const v = numericValue(cell);
        if (v === null)
            return false;
        return rule.bottom ? v <= threshold : v >= threshold;
    }
    function evalAboveAverage(rule, cell, range, ctx) {
        const cells = ctx.cellsInRange(range);
        const values = [];
        for (const entry of cells) {
            if (!entry.cell)
                continue;
            const n = numericValue(entry.cell);
            if (n === null)
                continue;
            values.push(n);
        }
        if (values.length === 0)
            return false;
        const mean = values.reduce((a, b) => a + b, 0) / values.length;
        let threshold = mean;
        if (rule.stdDev !== null && rule.stdDev !== undefined) {
            let variance = 0;
            for (const v of values) {
                const d = v - mean;
                variance += d * d;
            }
            variance /= values.length;
            const stddev = Math.sqrt(variance);
            threshold = mean + rule.stdDev * stddev;
        }
        const v = numericValue(cell);
        if (v === null)
            return false;
        const above = rule.aboveAverage;
        const incl = rule.equalAverage;
        if (above)
            return incl ? v >= threshold : v > threshold;
        return incl ? v <= threshold : v < threshold;
    }
    const MS_PER_DAY$1 = 86400000;
    const EPOCH_MS_1900$1 = Date.UTC(1899, 11, 30);
    const EPOCH_MS_1904$1 = Date.UTC(1904, 0, 1);
    function serialToDate$1(serial, date1904) {
        let days = Math.floor(serial);
        if (!date1904 && days >= 60)
            days -= 1;
        const epoch = date1904 ? EPOCH_MS_1904$1 : EPOCH_MS_1900$1;
        return new Date(epoch + days * MS_PER_DAY$1);
    }
    function startOfDayUtc(d) {
        return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    function startOfWeekUtc(d) {
        const dow = d.getUTCDay();
        return startOfDayUtc(d) - dow * MS_PER_DAY$1;
    }
    function evalTimePeriod(rule, cell, ctx) {
        if (!rule.timePeriod)
            return false;
        const n = numericValue(cell);
        if (n === null)
            return false;
        const date1904 = ctx.date1904 === true;
        const cellDate = serialToDate$1(n, date1904);
        const cellDay = startOfDayUtc(cellDate);
        const now = new Date();
        const today = startOfDayUtc(now);
        const yesterday = today - MS_PER_DAY$1;
        const tomorrow = today + MS_PER_DAY$1;
        const thisWeekStart = startOfWeekUtc(now);
        const lastWeekStart = thisWeekStart - 7 * MS_PER_DAY$1;
        const nextWeekStart = thisWeekStart + 7 * MS_PER_DAY$1;
        const last7Start = today - 6 * MS_PER_DAY$1;
        const thisMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1);
        const thisMonthEnd = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1);
        const lastMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1);
        const nextMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1);
        const nextMonthEnd = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 2, 1);
        switch (rule.timePeriod) {
            case 'today': return cellDay === today;
            case 'yesterday': return cellDay === yesterday;
            case 'tomorrow': return cellDay === tomorrow;
            case 'last7Days': return cellDay >= last7Start && cellDay <= today;
            case 'thisWeek': return cellDay >= thisWeekStart && cellDay < nextWeekStart;
            case 'lastWeek': return cellDay >= lastWeekStart && cellDay < thisWeekStart;
            case 'nextWeek': return cellDay >= nextWeekStart && cellDay < nextWeekStart + 7 * MS_PER_DAY$1;
            case 'thisMonth': return cellDay >= thisMonthStart && cellDay < thisMonthEnd;
            case 'lastMonth': return cellDay >= lastMonthStart && cellDay < thisMonthStart;
            case 'nextMonth': return cellDay >= nextMonthStart && cellDay < nextMonthEnd;
            default: return false;
        }
    }
    function resolveCfvo(cfvo, values) {
        if (values.length === 0)
            return null;
        const min = Math.min(...values);
        const max = Math.max(...values);
        switch (cfvo.type) {
            case 'min': return min;
            case 'max': return max;
            case 'num': {
                const n = Number(cfvo.value);
                return Number.isFinite(n) ? n : null;
            }
            case 'percent': {
                const p = Number(cfvo.value);
                if (!Number.isFinite(p))
                    return null;
                return min + (max - min) * (p / 100);
            }
            case 'percentile': {
                const p = Number(cfvo.value);
                if (!Number.isFinite(p))
                    return null;
                const sorted = values.slice().sort((a, b) => a - b);
                const rank = (p / 100) * (sorted.length - 1);
                const lo = Math.floor(rank);
                const hi = Math.ceil(rank);
                return sorted[lo] + (sorted[hi] - sorted[lo]) * (rank - lo);
            }
            default:
                return null;
        }
    }
    function interpolateColorScale(value, stops) {
        if (stops.length === 0)
            return null;
        if (value <= stops[0].threshold)
            return stops[0].hex;
        if (value >= stops[stops.length - 1].threshold)
            return stops[stops.length - 1].hex;
        for (let i = 0; i < stops.length - 1; i++) {
            const a = stops[i], b = stops[i + 1];
            if (value >= a.threshold && value <= b.threshold) {
                const t = b.threshold === a.threshold ? 0 : (value - a.threshold) / (b.threshold - a.threshold);
                return lerpHex(a.hex, b.hex, t);
            }
        }
        return null;
    }
    function lerpHex(a, b, t) {
        const pa = parseHex(a), pb = parseHex(b);
        if (!pa || !pb)
            return a;
        const mix = (x, y) => Math.round(x + (y - x) * t);
        return `#${toHex(mix(pa[0], pb[0]))}${toHex(mix(pa[1], pb[1]))}${toHex(mix(pa[2], pb[2]))}`;
    }
    function parseHex(h) {
        const m = /^#([0-9a-f]{6})$/i.exec(h);
        if (!m)
            return null;
        return [parseInt(m[1].slice(0, 2), 16), parseInt(m[1].slice(2, 4), 16), parseInt(m[1].slice(4, 6), 16)];
    }
    function toHex(n) {
        return Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0');
    }
    function evalExpression(rule, cell, _range, _ctx) {
        const raw = rule.formulas[0] ?? '';
        const src = raw.trim().replace(/^=/, '');
        const m = /^\$?([A-Z]+)\$?([1-9][0-9]*)\s*(<=|>=|<>|=|<|>)\s*(.+)$/.exec(src);
        if (!m) {
            return false;
        }
        const [, , , op, rhs] = m;
        const v = numericValue(cell);
        const rhsLit = parseFormulaLiteral(rhs);
        if (v === null || typeof rhsLit !== 'number')
            return false;
        switch (op) {
            case '<': return v < rhsLit;
            case '<=': return v <= rhsLit;
            case '>': return v > rhsLit;
            case '>=': return v >= rhsLit;
            case '=': return v === rhsLit;
            case '<>': return v !== rhsLit;
            default: return false;
        }
    }

    const SAFE_HREF_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:']);
    function isSafeHyperlinkHref(raw) {
        if (raw == null)
            return true;
        if (typeof raw !== 'string')
            return false;
        const trimmed = raw.trim();
        if (trimmed === '')
            return true;
        if (trimmed.startsWith('#'))
            return true;
        try {
            const parsed = new URL(trimmed, 'http://xlsxjs.invalid/');
            return SAFE_HREF_SCHEMES.has(parsed.protocol);
        }
        catch {
            return !/^[a-z][a-z0-9+.-]*:/i.test(trimmed);
        }
    }
    const NS = {
        main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
        xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
        c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
        tc: 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
        xda: 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray',
    };
    class WorkbookParser {
        constructor(_options) {
            this._options = _options;
        }
        parse(parts, media = {}) {
            const workbookXml = parts['xl/workbook.xml'];
            if (!workbookXml)
                throw new Error('xlsx-preview: workbook.xml missing');
            const sharedStrings = parts['xl/sharedStrings.xml']
                ? parseSharedStrings(parts['xl/sharedStrings.xml'])
                : [];
            const styles = parts['xl/styles.xml'] ? parseStyles(parts['xl/styles.xml']) : null;
            const themePath = Object.keys(parts).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
            const theme = themePath ? parseTheme(parts[themePath]) : null;
            const rels = parts['xl/_rels/workbook.xml.rels']
                ? parseRelationships(parts['xl/_rels/workbook.xml.rels'])
                : new Map();
            const persons = resolvePersons(rels, parts);
            const sheetMeta = parseSheetList(workbookXml);
            const { date1904, definedNames } = parseWorkbookMeta(workbookXml);
            const metadata = parts['xl/metadata.xml']
                ? parseWorkbookMetadata(parts['xl/metadata.xml'])
                : null;
            const sheets = [];
            for (let i = 0; i < sheetMeta.length; i++) {
                const { name, rId, state } = sheetMeta[i];
                let xmlPath = null;
                const rel = rId ? rels.get(rId) : undefined;
                if (rel) {
                    xmlPath = resolveWorkbookRelTarget(rel.target);
                }
                if (!xmlPath || !parts[xmlPath]) {
                    xmlPath = `xl/worksheets/sheet${i + 1}.xml`;
                }
                const xml = parts[xmlPath];
                if (!xml)
                    continue;
                const tables = resolveTablesForSheet(xmlPath, parts);
                const { images, charts, shapes } = resolveDrawingsForSheet(xmlPath, parts, media);
                const pivots = resolvePivotsForSheet(xmlPath, parts);
                const comments = resolveCommentsForSheet(xmlPath, parts);
                const threadedComments = resolveThreadedCommentsForSheet(xmlPath, parts, persons);
                const hyperlinkTargets = resolveHyperlinkTargets(xmlPath, parts);
                sheets.push(parseSheet(name, state, xml, sharedStrings, tables, images, charts, shapes, pivots, comments, threadedComments, hyperlinkTargets, i, definedNames, metadata));
            }
            return { sheets, styles, theme, persons, date1904, definedNames, metadata };
        }
    }
    function parseWorkbookMetadata(xml) {
        const doc = parseXml(xml);
        const typeNames = [];
        const typeEls = doc.getElementsByTagNameNS(NS.main, 'metadataType');
        for (let i = 0; i < typeEls.length; i++) {
            typeNames.push(typeEls[i].getAttribute('name') ?? '');
        }
        const futureFlagsByType = new Map();
        const fmEls = doc.getElementsByTagNameNS(NS.main, 'futureMetadata');
        for (let i = 0; i < fmEls.length; i++) {
            const fm = fmEls[i];
            const name = fm.getAttribute('name') ?? '';
            const bkEls = fm.getElementsByTagNameNS(NS.main, 'bk');
            const flags = [];
            for (let j = 0; j < bkEls.length; j++) {
                const bk = bkEls[j];
                let dynamic = false;
                const daProps = bk.getElementsByTagNameNS(NS.xda, 'dynamicArrayProperties').item(0);
                if (daProps) {
                    const attr = daProps.getAttribute('fDynamic');
                    dynamic = attr === '1' || attr === 'true';
                }
                flags.push(dynamic);
            }
            futureFlagsByType.set(name, flags);
        }
        const resolveBlock = (bk) => {
            const rc = bk.getElementsByTagNameNS(NS.main, 'rc').item(0);
            const t = rc ? Number(rc.getAttribute('t')) : NaN;
            const v = rc ? Number(rc.getAttribute('v')) : NaN;
            const typeIndex = Number.isFinite(t) && t >= 1 ? t - 1 : -1;
            const typeName = typeIndex >= 0 && typeIndex < typeNames.length
                ? typeNames[typeIndex]
                : '';
            let dynamicArray = false;
            if (typeName === 'XLDAPR') {
                const flags = futureFlagsByType.get('XLDAPR');
                if (flags && Number.isFinite(v) && v >= 0 && v < flags.length) {
                    dynamicArray = flags[v];
                }
            }
            return { typeIndex, dynamicArray };
        };
        const readList = (listTag) => {
            const out = [];
            const wrap = doc.getElementsByTagNameNS(NS.main, listTag).item(0);
            if (!wrap)
                return out;
            const bks = wrap.getElementsByTagNameNS(NS.main, 'bk');
            for (let i = 0; i < bks.length; i++)
                out.push(resolveBlock(bks[i]));
            return out;
        };
        return {
            cellMetadata: readList('cellMetadata'),
            valueMetadata: readList('valueMetadata'),
        };
    }
    function parseWorkbookMeta(workbookXml) {
        const doc = parseXml(workbookXml);
        const pr = doc.getElementsByTagNameNS(NS.main, 'workbookPr').item(0);
        let date1904 = false;
        if (pr) {
            const attr = pr.getAttribute('date1904');
            date1904 = attr === '1' || attr === 'true';
        }
        const definedNames = [];
        const dnEls = doc.getElementsByTagNameNS(NS.main, 'definedName');
        for (let i = 0; i < dnEls.length; i++) {
            const el = dnEls[i];
            const name = el.getAttribute('name');
            if (!name)
                continue;
            const localAttr = el.getAttribute('localSheetId');
            const localSheetId = localAttr != null && Number.isFinite(Number(localAttr))
                ? Number(localAttr)
                : null;
            const hidden = el.getAttribute('hidden') === '1';
            const formula = el.textContent ?? '';
            definedNames.push({ name, localSheetId, formula, hidden });
        }
        return { date1904, definedNames };
    }
    function resolvePersons(workbookRels, parts) {
        let xmlPath = null;
        for (const [, rel] of workbookRels) {
            if (rel.type.endsWith('/person')) {
                xmlPath = resolveWorkbookRelTarget(rel.target);
                break;
            }
        }
        if (!xmlPath || !parts[xmlPath]) {
            const candidates = Object.keys(parts).filter((p) => /^xl\/persons\/.*\.xml$/i.test(p));
            xmlPath = candidates[0] ?? null;
        }
        if (!xmlPath || !parts[xmlPath])
            return new Map();
        return parsePersons(parts[xmlPath]);
    }
    function parsePersons(xml) {
        const out = new Map();
        const doc = parseXml(xml);
        const nodes = doc.getElementsByTagNameNS(NS.tc, 'person');
        for (let i = 0; i < nodes.length; i++) {
            const el = nodes[i];
            const id = el.getAttribute('id');
            const displayName = el.getAttribute('displayName');
            if (!id)
                continue;
            out.set(id, displayName ?? '');
        }
        return out;
    }
    function resolveThreadedCommentsForSheet(sheetPath, parts, persons) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return [];
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const out = [];
        for (const [, rel] of rels) {
            if (!rel.type.endsWith('/threadedComment'))
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${dir}/${rel.target}`);
            const xml = parts[target];
            if (!xml)
                continue;
            out.push(...parseThreadedComments(xml, persons));
        }
        return out;
    }
    function parseThreadedComments(xml, persons) {
        const doc = parseXml(xml);
        const nodes = doc.getElementsByTagNameNS(NS.tc, 'threadedComment');
        const out = [];
        for (let i = 0; i < nodes.length; i++) {
            const el = nodes[i];
            const ref = el.getAttribute('ref');
            const id = el.getAttribute('id');
            if (!ref || !id)
                continue;
            const cell = parseCellRef(ref);
            if (!cell)
                continue;
            const personId = el.getAttribute('personId');
            const author = personId && persons.has(personId) ? persons.get(personId) : null;
            const textEl = el.getElementsByTagNameNS(NS.tc, 'text').item(0);
            const text = textEl?.textContent ?? '';
            const date = el.getAttribute('dT');
            const parentId = el.getAttribute('parentId');
            out.push({
                id,
                col: cell.col,
                row: cell.row,
                author: author && author.length > 0 ? author : null,
                date: date ?? null,
                text,
                parentId: parentId ?? null,
            });
        }
        return out;
    }
    function parseRelationships(xml) {
        const out = new Map();
        const doc = parseXml(xml);
        const rootEls = doc.getElementsByTagNameNS(NS.rels, 'Relationship');
        for (let i = 0; i < rootEls.length; i++) {
            const el = rootEls[i];
            const id = el.getAttribute('Id');
            const target = el.getAttribute('Target');
            const type = el.getAttribute('Type') ?? '';
            if (!id || !target)
                continue;
            out.set(id, { target, type });
        }
        return out;
    }
    function resolveWorkbookRelTarget(target) {
        if (target.startsWith('/'))
            return target.slice(1);
        const t = target.replace(/^\.\//, '');
        return `xl/${t}`;
    }
    function resolveTablesForSheet(sheetPath, parts) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return [];
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const out = [];
        for (const [, rel] of rels) {
            if (!rel.type.endsWith('/table'))
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : `${dir}/${rel.target}`.replace(/\/\.\//g, '/');
            const normalised = normaliseRelPath(target);
            const xml = parts[normalised];
            if (!xml)
                continue;
            const parsed = parseTable(xml);
            if (parsed)
                out.push(parsed);
        }
        return out;
    }
    function normaliseRelPath(path) {
        const stack = [];
        for (const seg of path.split('/')) {
            if (seg === '' || seg === '.')
                continue;
            if (seg === '..')
                stack.pop();
            else
                stack.push(seg);
        }
        return stack.join('/');
    }
    function resolveDrawingsForSheet(sheetPath, parts, media) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return { images: [], charts: [], shapes: [] };
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const images = [];
        const charts = [];
        const shapes = [];
        for (const [, rel] of rels) {
            if (!rel.type.endsWith('/drawing'))
                continue;
            const drawingPath = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${dir}/${rel.target}`);
            const drawingXml = parts[drawingPath];
            if (!drawingXml)
                continue;
            const drawingRelsPath = drawingPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
            const drawingRelsXml = parts[drawingRelsPath];
            const drawingRels = drawingRelsXml
                ? parseRelationships(drawingRelsXml)
                : new Map();
            const drawingDir = drawingPath.replace(/\/[^/]+$/, '');
            const parsed = parseDrawing(drawingXml, drawingRels, drawingDir, parts, media);
            images.push(...parsed.images);
            charts.push(...parsed.charts);
            shapes.push(...parsed.shapes);
        }
        return { images, charts, shapes };
    }
    function resolvePivotsForSheet(sheetPath, parts) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return [];
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const out = [];
        for (const [, rel] of rels) {
            if (!rel.type.endsWith('/pivotTable'))
                continue;
            const pivotPath = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${dir}/${rel.target}`);
            const xml = parts[pivotPath];
            if (!xml)
                continue;
            const parsed = parsePivotTable(xml);
            if (parsed)
                out.push(parsed);
        }
        return out;
    }
    function parsePivotTable(xml) {
        const doc = parseXml(xml);
        const def = doc.getElementsByTagNameNS(NS.main, 'pivotTableDefinition').item(0);
        if (!def)
            return null;
        const name = def.getAttribute('name') ?? '';
        const loc = def.getElementsByTagNameNS(NS.main, 'location').item(0);
        if (!loc)
            return null;
        const ref = loc.getAttribute('ref');
        if (!ref)
            return null;
        const parts = ref.split(':');
        const a = parseCellRef(parts[0]);
        const b = parts[1] ? parseCellRef(parts[1]) : a;
        if (!a || !b)
            return null;
        return {
            name,
            col: Math.min(a.col, b.col),
            row: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row),
        };
    }
    function resolveCommentsForSheet(sheetPath, parts) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return [];
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const out = [];
        for (const [, rel] of rels) {
            if (!rel.type.endsWith('/comments'))
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${dir}/${rel.target}`);
            const xml = parts[target];
            if (!xml)
                continue;
            out.push(...parseComments(xml));
        }
        return out;
    }
    function resolveHyperlinkTargets(sheetPath, parts) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return new Map();
        const rels = parseRelationships(relsXml);
        const out = new Map();
        for (const [id, rel] of rels) {
            if (!rel.type.endsWith('/hyperlink'))
                continue;
            out.set(id, rel.target);
        }
        return out;
    }
    function parseComments(xml) {
        const doc = parseXml(xml);
        const authors = [];
        const authorEls = doc.getElementsByTagNameNS(NS.main, 'author');
        for (let i = 0; i < authorEls.length; i++) {
            authors.push(authorEls[i].textContent ?? '');
        }
        const out = [];
        const commentEls = doc.getElementsByTagNameNS(NS.main, 'comment');
        for (let i = 0; i < commentEls.length; i++) {
            const c = commentEls[i];
            const ref = c.getAttribute('ref');
            if (!ref)
                continue;
            const parsed = parseCellRef(ref);
            if (!parsed)
                continue;
            const authorIdAttr = c.getAttribute('authorId');
            const authorIdx = authorIdAttr != null ? Number(authorIdAttr) : NaN;
            const author = Number.isFinite(authorIdx) && authorIdx >= 0 && authorIdx < authors.length
                ? authors[authorIdx]
                : null;
            const textEl = c.getElementsByTagNameNS(NS.main, 'text').item(0);
            const body = textEl ? parseSi(textEl) : { text: '', runs: null};
            out.push({
                col: parsed.col,
                row: parsed.row,
                author: author && author.length > 0 ? author : null,
                text: body.text,
                runs: body.runs,
            });
        }
        return out;
    }
    function parseDrawing(xml, rels, drawingDir, parts, media) {
        const doc = parseXml(xml);
        const images = [];
        const charts = [];
        const shapes = [];
        const labelled = [
            ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'twoCellAnchor'))
                .map((el) => ({ el, mode: 'twoCell' })),
            ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'oneCellAnchor'))
                .map((el) => ({ el, mode: 'oneCell' })),
            ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'absoluteAnchor'))
                .map((el) => ({ el, mode: 'absolute' })),
        ];
        for (const { el: anchor, mode } of labelled) {
            const from = anchor.getElementsByTagNameNS(NS.xdr, 'from').item(0);
            const to = mode === 'twoCell'
                ? anchor.getElementsByTagNameNS(NS.xdr, 'to').item(0)
                : null;
            const col = from ? anchorCellValue(from, 'col') : null;
            const row = from ? anchorCellValue(from, 'row') : null;
            const endCol = to ? anchorCellValue(to, 'col') : null;
            const endRow = to ? anchorCellValue(to, 'row') : null;
            const colOff = from ? anchorCellValue(from, 'colOff') : null;
            const rowOff = from ? anchorCellValue(from, 'rowOff') : null;
            const ext = findDirectChildNS(anchor, NS.xdr, 'ext');
            const extCx = ext ? Number(ext.getAttribute('cx')) : NaN;
            const extCy = ext ? Number(ext.getAttribute('cy')) : NaN;
            const widthEmu = Number.isFinite(extCx) ? extCx : null;
            const heightEmu = Number.isFinite(extCy) ? extCy : null;
            let absoluteX = null;
            let absoluteY = null;
            if (mode === 'absolute') {
                const pos = findDirectChildNS(anchor, NS.xdr, 'pos');
                if (pos) {
                    const x = Number(pos.getAttribute('x'));
                    const y = Number(pos.getAttribute('y'));
                    if (Number.isFinite(x))
                        absoluteX = x;
                    if (Number.isFinite(y))
                        absoluteY = y;
                }
            }
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
                        const cNvPr = pic.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                        const alt = cNvPr?.getAttribute('descr') ?? cNvPr?.getAttribute('title') ?? null;
                        const decorative = cNvPr ? detectDecorative(cNvPr) : false;
                        images.push({
                            col: col ?? 0,
                            row: row ?? 0,
                            endCol,
                            endRow,
                            colOff: colOff ?? 0,
                            rowOff: rowOff ?? 0,
                            dataUrl,
                            widthEmu,
                            heightEmu,
                            alt,
                            anchorMode: mode,
                            absoluteX,
                            absoluteY,
                            decorative,
                        });
                    }
                }
            }
            const chartEl = anchor.getElementsByTagNameNS(NS.c, 'chart').item(0)
                ?? anchor.getElementsByTagNameNS(NS.cx, 'chart').item(0);
            if (chartEl) {
                const kind = chartEl.namespaceURI === NS.cx ? 'chartex' : 'classic';
                const rId = chartEl.getAttributeNS(NS.rel, 'id');
                let chartType = null;
                if (rId) {
                    const rel = rels.get(rId);
                    if (rel) {
                        const chartPath = rel.target.startsWith('/')
                            ? rel.target.slice(1)
                            : normaliseRelPath(`${drawingDir}/${rel.target}`);
                        const chartXml = parts[chartPath];
                        if (chartXml)
                            chartType = peekChartType(chartXml, kind);
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
            const spEls = [
                ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'sp')),
                ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'cxnSp')),
            ];
            for (const el of spEls) {
                const kind = el.localName === 'cxnSp' ? 'connector' : 'shape';
                const shape = parseShape(el, kind, col, row, endCol, endRow);
                if (shape)
                    shapes.push(shape);
            }
        }
        return { images, charts, shapes };
    }
    function parseShape(el, kind, col, row, endCol, endRow) {
        const cNvPr = el.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
        const name = cNvPr?.getAttribute('name') || null;
        const alt = cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('title') || null;
        const spPr = el.getElementsByTagNameNS(NS.xdr, 'spPr').item(0);
        let preset = null;
        if (spPr) {
            const prst = spPr.getElementsByTagNameNS(NS.a, 'prstGeom').item(0);
            preset = prst?.getAttribute('prst') || null;
        }
        let text = null;
        const txBody = el.getElementsByTagNameNS(NS.xdr, 'txBody').item(0);
        if (txBody) {
            const paragraphs = [];
            const pEls = txBody.getElementsByTagNameNS(NS.a, 'p');
            for (let i = 0; i < pEls.length; i++) {
                const p = pEls[i];
                const tEls = p.getElementsByTagNameNS(NS.a, 't');
                let line = '';
                for (let j = 0; j < tEls.length; j++)
                    line += tEls[j].textContent ?? '';
                paragraphs.push(line);
            }
            const joined = paragraphs.join('\n');
            if (joined.length > 0)
                text = joined;
        }
        return {
            kind,
            name,
            alt,
            preset,
            text,
            col: col ?? 0,
            row: row ?? 0,
            endCol,
            endRow,
        };
    }
    function peekChartType(xml, kind) {
        const doc = parseXml(xml);
        if (kind === 'classic') {
            const plotArea = doc.getElementsByTagNameNS(NS.c, 'plotArea').item(0);
            if (!plotArea)
                return null;
            for (let i = 0; i < plotArea.childNodes.length; i++) {
                const node = plotArea.childNodes[i];
                if (node.nodeType !== 1)
                    continue;
                const el = node;
                if (el.namespaceURI !== NS.c)
                    continue;
                if (/Ax$/.test(el.localName) || el.localName === 'numFmt')
                    continue;
                if (/Chart$/.test(el.localName) || el.localName === 'chartEx')
                    return el.localName;
            }
            return null;
        }
        const series = doc.getElementsByTagNameNS(NS.cx, 'series').item(0);
        const layoutId = series?.getAttribute('layoutId');
        if (layoutId)
            return layoutId;
        const chart = doc.getElementsByTagNameNS(NS.cx, 'chart').item(0);
        if (!chart)
            return null;
        for (let i = 0; i < chart.childNodes.length; i++) {
            const node = chart.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI !== NS.cx)
                continue;
            if (el.localName !== 'plotArea' && el.localName !== 'title')
                return el.localName;
        }
        return null;
    }
    function anchorCellValue(parent, tag) {
        if (!parent)
            return null;
        const el = parent.getElementsByTagNameNS(NS.xdr, tag).item(0);
        if (!el)
            return null;
        const n = Number(el.textContent);
        return Number.isFinite(n) ? n : null;
    }
    function findDirectChildNS(parent, ns, localName) {
        for (let i = 0; i < parent.childNodes.length; i++) {
            const node = parent.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI === ns && el.localName === localName)
                return el;
        }
        return null;
    }
    function detectDecorative(cNvPr) {
        const extLst = cNvPr.getElementsByTagNameNS(NS.a, 'extLst').item(0);
        if (!extLst)
            return false;
        const truthy = (v) => v === '1' || v === 'true';
        const all = extLst.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === 'decorative' && (truthy(el.getAttribute('val')) || truthy(el.getAttribute('value')))) {
                return true;
            }
            if (truthy(el.getAttribute('decorative')))
                return true;
        }
        return false;
    }
    function parseTable(xml) {
        const doc = parseXml(xml);
        const t = doc.getElementsByTagNameNS(NS.main, 'table').item(0);
        if (!t)
            return null;
        const ref = t.getAttribute('ref');
        if (!ref)
            return null;
        const parts = ref.split(':');
        const a = parseCellRef(parts[0]);
        const b = parts[1] ? parseCellRef(parts[1]) : a;
        if (!a || !b)
            return null;
        const columns = [];
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
            altText: t.getAttribute('altText') || null,
            altTextSummary: t.getAttribute('altTextSummary') || null,
        };
    }
    function parseXml(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function parseSheetList(workbookXml) {
        const doc = parseXml(workbookXml);
        const nodes = doc.getElementsByTagNameNS(NS.main, 'sheet');
        const out = [];
        for (let i = 0; i < nodes.length; i++) {
            const rId = nodes[i].getAttributeNS(NS.rel, 'id');
            const stateAttr = nodes[i].getAttribute('state');
            const state = stateAttr === 'hidden' || stateAttr === 'veryHidden' ? stateAttr : 'visible';
            out.push({
                name: nodes[i].getAttribute('name') ?? `Sheet${i + 1}`,
                rId: rId || null,
                state,
            });
        }
        return out;
    }
    function parseSharedStrings(xml) {
        const doc = parseXml(xml);
        const sis = doc.getElementsByTagNameNS(NS.main, 'si');
        const out = [];
        for (let i = 0; i < sis.length; i++) {
            out.push(parseSi(sis[i]));
        }
        return out;
    }
    function parseSi(si) {
        const runs = [];
        let sawRun = false;
        const phonetics = [];
        for (let i = 0; i < si.childNodes.length; i++) {
            const node = si.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI !== NS.main)
                continue;
            if (el.localName === 't') {
                runs.push(emptyRun(el.textContent ?? ''));
                continue;
            }
            if (el.localName === 'r') {
                sawRun = true;
                runs.push(parseRun(el));
                continue;
            }
            if (el.localName === 'rPh') {
                const parsed = parseRPh(el);
                if (parsed)
                    phonetics.push(parsed);
            }
        }
        const text = runs.map((r) => r.text).join('');
        for (const p of phonetics) {
            const start = Math.max(0, Math.min(text.length, p.startIdx));
            const end = Math.max(start, Math.min(text.length, p.endIdx));
            p.startIdx = start;
            p.endIdx = end;
            p.base = text.slice(start, end);
        }
        return { text, runs: sawRun ? runs : null, phonetics: phonetics.length > 0 ? phonetics : null };
    }
    function parseRPh(el) {
        const sbAttr = el.getAttribute('sb');
        const ebAttr = el.getAttribute('eb');
        if (sbAttr === null || ebAttr === null)
            return null;
        const sb = Number(sbAttr);
        const eb = Number(ebAttr);
        if (!Number.isFinite(sb) || !Number.isFinite(eb))
            return null;
        if (sb < 0 || eb < sb)
            return null;
        const tEl = el.getElementsByTagNameNS(NS.main, 't').item(0);
        const phonetic = tEl?.textContent ?? '';
        return {
            base: '',
            phonetic,
            startIdx: Math.floor(sb),
            endIdx: Math.floor(eb),
        };
    }
    function emptyRun(text) {
        return {
            text,
            bold: false,
            italic: false,
            underline: null,
            strike: false,
            vertAlign: null,
            size: null,
            color: null,
            name: null,
        };
    }
    function parseRun(el) {
        const rPr = el.getElementsByTagNameNS(NS.main, 'rPr').item(0);
        const ts = el.getElementsByTagNameNS(NS.main, 't');
        let text = '';
        for (let i = 0; i < ts.length; i++)
            text += ts[i].textContent ?? '';
        if (!rPr)
            return emptyRun(text);
        const has = (tag) => rPr.getElementsByTagNameNS(NS.main, tag).length > 0;
        const firstAttr = (tag, attr) => {
            const node = rPr.getElementsByTagNameNS(NS.main, tag).item(0);
            return node ? node.getAttribute(attr) : null;
        };
        const color = rPr.getElementsByTagNameNS(NS.main, 'color').item(0);
        const sizeAttr = firstAttr('sz', 'val');
        const uEl = rPr.getElementsByTagNameNS(NS.main, 'u').item(0);
        let underline = null;
        if (uEl) {
            const val = uEl.getAttribute('val');
            if (val === null || val === '')
                underline = 'single';
            else if (val === 'single' || val === 'double' || val === 'singleAccounting' || val === 'doubleAccounting')
                underline = val;
            else if (val === 'none')
                underline = null;
        }
        const vaEl = rPr.getElementsByTagNameNS(NS.main, 'vertAlign').item(0);
        let vertAlign = null;
        if (vaEl) {
            const val = vaEl.getAttribute('val');
            if (val === 'subscript' || val === 'superscript')
                vertAlign = val;
        }
        return {
            text,
            bold: has('b'),
            italic: has('i'),
            underline,
            strike: has('strike'),
            vertAlign,
            size: sizeAttr ? Number(sizeAttr) : null,
            color: parseColorElement(color),
            name: firstAttr('rFont', 'val') ?? firstAttr('name', 'val'),
        };
    }
    function parseSheet(name, state, xml, sharedStrings, tables = [], images = [], charts = [], shapes = [], pivots = [], comments = [], threadedComments = [], hyperlinkTargets = new Map(), sheetIndex = 0, definedNames = [], metadata = null) {
        const doc = parseXml(xml);
        const rowEls = doc.getElementsByTagNameNS(NS.main, 'row');
        const rows = [];
        const rowDimensions = [];
        let maxCol = -1;
        let maxRow = -1;
        for (let i = 0; i < rowEls.length; i++) {
            const rowEl = rowEls[i];
            const rowAttr = rowEl.getAttribute('r');
            const rowIndex = rowAttr ? Number(rowAttr) - 1 : i;
            if (!Number.isFinite(rowIndex) || rowIndex < 0)
                continue;
            const htAttr = rowEl.getAttribute('ht');
            const customHeight = rowEl.getAttribute('customHeight') === '1';
            const hidden = rowEl.getAttribute('hidden') === '1';
            let height = null;
            if (htAttr && (customHeight || !Number.isNaN(Number(htAttr)))) {
                const h = Number(htAttr);
                if (Number.isFinite(h) && h >= 0)
                    height = h;
            }
            const outlineAttr = rowEl.getAttribute('outlineLevel');
            const outlineLevel = outlineAttr != null && Number.isFinite(Number(outlineAttr))
                ? Math.max(0, Number(outlineAttr))
                : 0;
            if (height !== null || hidden || outlineLevel > 0) {
                rowDimensions.push({ row: rowIndex, height, hidden, outlineLevel });
                if (rowIndex > maxRow)
                    maxRow = rowIndex;
            }
            const cellEls = rowEl.getElementsByTagNameNS(NS.main, 'c');
            const cells = [];
            for (let j = 0; j < cellEls.length; j++) {
                const cell = parseCell(cellEls[j], rowIndex, sharedStrings, metadata);
                if (!cell)
                    continue;
                cells.push(cell);
                if (cell.col > maxCol)
                    maxCol = cell.col;
            }
            if (cells.length === 0)
                continue;
            rows[rowIndex] = cells;
            if (rowIndex > maxRow)
                maxRow = rowIndex;
        }
        const merges = parseMerges(doc);
        for (const m of merges) {
            const rightCol = m.col + m.colSpan - 1;
            const bottomRow = m.row + m.rowSpan - 1;
            if (rightCol > maxCol)
                maxCol = rightCol;
            if (bottomRow > maxRow)
                maxRow = bottomRow;
        }
        const columns = parseCols(doc);
        const conditionalFormatting = parseConditionalFormatting(doc);
        const frozenPanes = parseFrozenPanes(doc);
        const autoFilter = parseAutoFilter(doc);
        const extensions = collectExtensionUris(doc);
        const view = parseSheetView(doc);
        const outline = parseSheetOutline(doc);
        const hyperlinks = parseHyperlinks(doc, hyperlinkTargets);
        const dataValidationLists = parseDataValidationLists(doc);
        const pageBreaks = parsePageBreaks(doc);
        const printArea = resolvePrintArea(sheetIndex, definedNames);
        const headerFooter = parseHeaderFooter(doc, name);
        const protection = parseSheetProtection(doc);
        for (const h of hyperlinks) {
            if (h.col > maxCol)
                maxCol = h.col;
            if (h.row > maxRow)
                maxRow = h.row;
        }
        for (const v of dataValidationLists) {
            if (v.endCol > maxCol)
                maxCol = v.endCol;
            if (v.endRow > maxRow)
                maxRow = v.endRow;
        }
        return {
            name, state, rows, maxCol, maxRow, merges, columns, rowDimensions,
            conditionalFormatting, frozenPanes, autoFilter, tables, images,
            charts, shapes, pivots, extensions, comments, threadedComments, view, outline,
            hyperlinks, dataValidationLists, pageBreaks, printArea, headerFooter,
            protection,
        };
    }
    function parseSheetProtection(doc) {
        const el = doc.getElementsByTagNameNS(NS.main, 'sheetProtection').item(0);
        if (!el)
            return null;
        const locked = (attr) => el.getAttribute(attr) !== '0';
        const enabled = el.getAttribute('sheet') === '1';
        const passwordHashed = el.getAttribute('password') !== null ||
            el.getAttribute('hashValue') !== null;
        return {
            enabled,
            passwordHashed,
            selectLockedCells: locked('selectLockedCells'),
            selectUnlockedCells: locked('selectUnlockedCells'),
            formatCells: locked('formatCells'),
            formatColumns: locked('formatColumns'),
            formatRows: locked('formatRows'),
            insertColumns: locked('insertColumns'),
            insertRows: locked('insertRows'),
            deleteColumns: locked('deleteColumns'),
            deleteRows: locked('deleteRows'),
            sort: locked('sort'),
            autoFilter: locked('autoFilter'),
            pivotTables: locked('pivotTables'),
            objects: locked('objects'),
            scenarios: locked('scenarios'),
        };
    }
    function parseSheetView(doc) {
        const defaults = {
            rightToLeft: false,
            showGridLines: true,
            showRowColHeaders: true,
            zoomScale: null,
            tabColor: null,
        };
        const sv = doc.getElementsByTagNameNS(NS.main, 'sheetView').item(0);
        if (sv) {
            defaults.rightToLeft = sv.getAttribute('rightToLeft') === '1';
            const sgl = sv.getAttribute('showGridLines');
            if (sgl === '0')
                defaults.showGridLines = false;
            const sh = sv.getAttribute('showRowColHeaders');
            if (sh === '0')
                defaults.showRowColHeaders = false;
            const zs = sv.getAttribute('zoomScale');
            if (zs !== null) {
                const n = Number(zs);
                if (Number.isFinite(n) && n > 0)
                    defaults.zoomScale = n;
            }
        }
        const sheetPr = doc.getElementsByTagNameNS(NS.main, 'sheetPr').item(0);
        if (sheetPr) {
            const tc = sheetPr.getElementsByTagNameNS(NS.main, 'tabColor').item(0);
            if (tc)
                defaults.tabColor = parseColorElement(tc);
        }
        return defaults;
    }
    function parseSheetOutline(doc) {
        let maxRowLevel = 0;
        let maxColLevel = 0;
        const fmtPr = doc.getElementsByTagNameNS(NS.main, 'sheetFormatPr').item(0);
        if (fmtPr) {
            const r = Number(fmtPr.getAttribute('outlineLevelRow'));
            if (Number.isFinite(r) && r > 0)
                maxRowLevel = r;
            const c = Number(fmtPr.getAttribute('outlineLevelCol'));
            if (Number.isFinite(c) && c > 0)
                maxColLevel = c;
        }
        let summaryBelow = true;
        let summaryRight = true;
        const sheetPr = doc.getElementsByTagNameNS(NS.main, 'sheetPr').item(0);
        if (sheetPr) {
            const outlinePr = sheetPr.getElementsByTagNameNS(NS.main, 'outlinePr').item(0);
            if (outlinePr) {
                const sb = outlinePr.getAttribute('summaryBelow');
                if (sb === '0' || sb === 'false')
                    summaryBelow = false;
                const sr = outlinePr.getAttribute('summaryRight');
                if (sr === '0' || sr === 'false')
                    summaryRight = false;
            }
        }
        return { maxRowLevel, maxColLevel, summaryBelow, summaryRight };
    }
    function parseHyperlinks(doc, targets) {
        const out = [];
        const list = doc.getElementsByTagNameNS(NS.main, 'hyperlinks').item(0);
        if (!list)
            return out;
        const hls = list.getElementsByTagNameNS(NS.main, 'hyperlink');
        for (let i = 0; i < hls.length; i++) {
            const el = hls[i];
            const ref = el.getAttribute('ref');
            if (!ref)
                continue;
            const rId = el.getAttributeNS(NS.rel, 'id');
            const target = rId ? (targets.get(rId) ?? null) : null;
            const location = el.getAttribute('location');
            const tooltip = el.getAttribute('tooltip');
            const display = el.getAttribute('display');
            const ranges = parseSqref(ref);
            for (const range of ranges) {
                for (let r = range.row; r <= range.endRow; r++) {
                    for (let c = range.col; c <= range.endCol; c++) {
                        out.push({
                            col: c,
                            row: r,
                            target,
                            location: location || null,
                            tooltip: tooltip || null,
                            display: display || null,
                        });
                    }
                }
            }
        }
        return out;
    }
    function parseDataValidationLists(doc) {
        const out = [];
        const list = doc.getElementsByTagNameNS(NS.main, 'dataValidations').item(0);
        if (!list)
            return out;
        const dvs = list.getElementsByTagNameNS(NS.main, 'dataValidation');
        for (let i = 0; i < dvs.length; i++) {
            const el = dvs[i];
            if (el.getAttribute('type') !== 'list')
                continue;
            const sqref = el.getAttribute('sqref');
            if (!sqref)
                continue;
            const f1 = el.getElementsByTagNameNS(NS.main, 'formula1').item(0);
            const formulaText = (f1?.textContent ?? '').trim();
            let options = null;
            const quoted = /^"(.*)"$/s.exec(formulaText);
            if (quoted) {
                options = quoted[1].split(',');
            }
            const ranges = parseSqref(sqref);
            for (const range of ranges) {
                out.push({
                    col: range.col,
                    row: range.row,
                    endCol: range.endCol,
                    endRow: range.endRow,
                    options,
                });
            }
        }
        return out;
    }
    function collectExtensionUris(doc) {
        const counts = new Map();
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'ext')
                continue;
            const uri = el.getAttribute('uri');
            if (!uri)
                continue;
            counts.set(uri, (counts.get(uri) ?? 0) + 1);
        }
        const out = [];
        for (const [uri, count] of counts)
            out.push({ uri, count });
        return out;
    }
    function parseFrozenPanes(doc) {
        const pane = doc.getElementsByTagNameNS(NS.main, 'pane').item(0);
        if (!pane)
            return null;
        const state = pane.getAttribute('state');
        let kind;
        if (state === 'frozen' || state === 'frozenSplit')
            kind = 'frozen';
        else if (state === 'split')
            kind = 'split';
        else
            return null;
        const xSplit = Number(pane.getAttribute('xSplit'));
        const ySplit = Number(pane.getAttribute('ySplit'));
        const x = Number.isFinite(xSplit) && xSplit > 0 ? xSplit : null;
        const y = Number.isFinite(ySplit) && ySplit > 0 ? ySplit : null;
        if (x === null && y === null)
            return null;
        return { kind, xSplit: x, ySplit: y };
    }
    function parsePageBreaks(doc) {
        const rows = [];
        const cols = [];
        const rowWrap = doc.getElementsByTagNameNS(NS.main, 'rowBreaks').item(0);
        if (rowWrap) {
            const brks = rowWrap.getElementsByTagNameNS(NS.main, 'brk');
            for (let i = 0; i < brks.length; i++) {
                const el = brks[i];
                if (el.getAttribute('man') !== '1')
                    continue;
                const id = Number(el.getAttribute('id'));
                if (!Number.isFinite(id) || id <= 0)
                    continue;
                rows.push(id - 1);
            }
        }
        const colWrap = doc.getElementsByTagNameNS(NS.main, 'colBreaks').item(0);
        if (colWrap) {
            const brks = colWrap.getElementsByTagNameNS(NS.main, 'brk');
            for (let i = 0; i < brks.length; i++) {
                const el = brks[i];
                if (el.getAttribute('man') !== '1')
                    continue;
                const id = Number(el.getAttribute('id'));
                if (!Number.isFinite(id) || id <= 0)
                    continue;
                cols.push(id - 1);
            }
        }
        return { rows, cols };
    }
    function parseHeaderFooter(doc, sheetName) {
        const hf = doc.getElementsByTagNameNS(NS.main, 'headerFooter').item(0);
        if (!hf)
            return null;
        const oddHeader = hf.getElementsByTagNameNS(NS.main, 'oddHeader').item(0);
        const oddFooter = hf.getElementsByTagNameNS(NS.main, 'oddFooter').item(0);
        const parseOne = (el) => {
            if (!el)
                return null;
            const raw = el.textContent ?? '';
            if (raw === '')
                return null;
            const zones = splitHeaderFooterZones(raw);
            return {
                left: substituteHeaderFooterCodes(zones.left, sheetName),
                center: substituteHeaderFooterCodes(zones.center, sheetName),
                right: substituteHeaderFooterCodes(zones.right, sheetName),
            };
        };
        const header = parseOne(oddHeader);
        const footer = parseOne(oddFooter);
        if (!header && !footer)
            return null;
        return { oddHeader: header, oddFooter: footer };
    }
    function splitHeaderFooterZones(raw) {
        const out = { left: '', center: '', right: '' };
        let zone = 'center';
        let i = 0;
        while (i < raw.length) {
            if (raw[i] === '&' && i + 1 < raw.length) {
                const next = raw[i + 1];
                if (next === 'L') {
                    zone = 'left';
                    i += 2;
                    continue;
                }
                if (next === 'C') {
                    zone = 'center';
                    i += 2;
                    continue;
                }
                if (next === 'R') {
                    zone = 'right';
                    i += 2;
                    continue;
                }
                out[zone] += raw[i] + raw[i + 1];
                i += 2;
                continue;
            }
            out[zone] += raw[i];
            i += 1;
        }
        return out;
    }
    function substituteHeaderFooterCodes(text, sheetName) {
        let out = '';
        let i = 0;
        const now = new Date();
        const pad = (n) => (n < 10 ? `0${n}` : String(n));
        const dateStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
        const timeStr = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
        while (i < text.length) {
            if (text[i] === '&' && i + 1 < text.length) {
                const next = text[i + 1];
                if (next === '&') {
                    out += '&';
                    i += 2;
                    continue;
                }
                if (next === 'D') {
                    out += dateStr;
                    i += 2;
                    continue;
                }
                if (next === 'T') {
                    out += timeStr;
                    i += 2;
                    continue;
                }
                if (next === 'F') {
                    out += '&F';
                    i += 2;
                    continue;
                }
                if (next === 'A') {
                    out += sheetName;
                    i += 2;
                    continue;
                }
                if (next === 'P') {
                    out += '(page)';
                    i += 2;
                    continue;
                }
                if (next === 'N') {
                    out += '(total)';
                    i += 2;
                    continue;
                }
                out += text[i] + text[i + 1];
                i += 2;
                continue;
            }
            out += text[i];
            i += 1;
        }
        return out;
    }
    function resolvePrintArea(sheetIndex, definedNames) {
        const hits = definedNames.filter((n) => n.name === '_xlnm.Print_Area' && n.localSheetId === sheetIndex);
        if (hits.length === 0)
            return null;
        const out = [];
        for (const n of hits) {
            const parts = n.formula.split(',');
            for (const part of parts) {
                const range = parsePrintAreaRef(part.trim());
                if (range)
                    out.push(range);
            }
        }
        return out.length > 0 ? out : null;
    }
    function parsePrintAreaRef(ref) {
        let body = ref;
        const bangIdx = body.lastIndexOf('!');
        if (bangIdx >= 0)
            body = body.slice(bangIdx + 1);
        const clean = body.replace(/\$/g, '');
        const [tl, br] = clean.split(':');
        const a = parseCellRef(tl);
        if (!a)
            return null;
        const b = br ? parseCellRef(br) : a;
        if (!b)
            return null;
        return {
            col: Math.min(a.col, b.col),
            row: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row),
        };
    }
    function parseAutoFilter(doc) {
        const af = doc.getElementsByTagNameNS(NS.main, 'autoFilter').item(0);
        if (!af)
            return null;
        const ref = af.getAttribute('ref');
        if (!ref)
            return null;
        const parts = ref.split(':');
        const a = parseCellRef(parts[0]);
        const b = parts[1] ? parseCellRef(parts[1]) : a;
        if (!a || !b)
            return null;
        return {
            col: Math.min(a.col, b.col),
            row: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row),
        };
    }
    function parseMerges(doc) {
        const out = [];
        const els = doc.getElementsByTagNameNS(NS.main, 'mergeCell');
        for (let i = 0; i < els.length; i++) {
            const ref = els[i].getAttribute('ref');
            if (!ref)
                continue;
            const [tl, br] = ref.split(':');
            if (!tl || !br)
                continue;
            const a = parseCellRef(tl);
            const b = parseCellRef(br);
            if (!a || !b)
                continue;
            if (b.col < a.col || b.row < a.row)
                continue;
            out.push({
                col: a.col,
                row: a.row,
                colSpan: b.col - a.col + 1,
                rowSpan: b.row - a.row + 1,
            });
        }
        return out;
    }
    function parseCols(doc) {
        const out = [];
        const els = doc.getElementsByTagNameNS(NS.main, 'col');
        for (let i = 0; i < els.length; i++) {
            const el = els[i];
            const min = Number(el.getAttribute('min'));
            const max = Number(el.getAttribute('max'));
            const widthAttr = el.getAttribute('width');
            const customWidth = el.getAttribute('customWidth') === '1';
            const hidden = el.getAttribute('hidden') === '1';
            const outlineAttr = el.getAttribute('outlineLevel');
            const outlineLevel = outlineAttr != null && Number.isFinite(Number(outlineAttr))
                ? Math.max(0, Number(outlineAttr))
                : 0;
            const hasWidth = widthAttr !== null && (customWidth || !Number.isNaN(Number(widthAttr)));
            if (!hasWidth && !hidden && outlineLevel === 0)
                continue;
            if (!Number.isFinite(min) || !Number.isFinite(max))
                continue;
            if (min < 1 || max < min)
                continue;
            const width = hasWidth && Number.isFinite(Number(widthAttr)) ? Number(widthAttr) : null;
            out.push({ min: min - 1, max: max - 1, width, hidden, outlineLevel });
        }
        return out;
    }
    function parseCell(c, fallbackRow, sharedStrings, metadata = null) {
        const ref = c.getAttribute('r');
        let col = 0, row = fallbackRow;
        if (ref) {
            const parsed = parseCellRef(ref);
            if (!parsed)
                return null;
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
        const resolveMetaIdx = (attr, list) => {
            const raw = c.getAttribute(attr);
            if (raw == null || raw === '')
                return null;
            const n = Number(raw);
            if (!Number.isFinite(n) || n <= 0)
                return null;
            const zero = Math.floor(n) - 1;
            if (!list)
                return zero;
            return zero;
        };
        const cellMetadataIndex = resolveMetaIdx('cm', metadata?.cellMetadata);
        const valueMetadataIndex = resolveMetaIdx('vm', metadata?.valueMetadata);
        let isSpillAnchor = false;
        if (cellMetadataIndex !== null && metadata) {
            const block = metadata.cellMetadata[cellMetadataIndex];
            if (block && block.dynamicArray)
                isSpillAnchor = true;
        }
        const base = {
            col, row, styleIndex, formula,
            cellMetadataIndex, valueMetadataIndex, isSpillAnchor,
        };
        switch (type) {
            case 's': {
                const idx = Number(raw);
                const entry = Number.isFinite(idx) && idx >= 0 && idx < sharedStrings.length
                    ? sharedStrings[idx]
                    : { text: '', runs: null, phonetics: null };
                return { ...base, value: entry.text, runs: entry.runs, phonetics: entry.phonetics, kind: 'string' };
            }
            case 'inlineStr': {
                const is = c.getElementsByTagNameNS(NS.main, 'is').item(0);
                const entry = is ? parseSi(is) : { text: '', runs: null, phonetics: null };
                return { ...base, value: entry.text, runs: entry.runs, phonetics: entry.phonetics, kind: 'inlineStr' };
            }
            case 'b':
                return { ...base, value: raw === '1' ? 'TRUE' : 'FALSE', runs: null, phonetics: null, kind: 'boolean' };
            case 'e':
                return { ...base, value: raw, runs: null, phonetics: null, kind: 'error' };
            case 'str':
                return { ...base, value: raw, runs: null, phonetics: null, kind: 'string' };
            case 'n':
            default:
                if (raw === '')
                    return { ...base, value: '', runs: null, phonetics: null, kind: 'empty' };
                return { ...base, value: raw, runs: null, phonetics: null, kind: 'number' };
        }
    }

    function h(tag, attrs, children) {
        const el = document.createElement(tag);
        if (attrs) {
            for (const [k, v] of Object.entries(attrs)) {
                if (v === null || v === undefined)
                    continue;
                el.setAttribute(k, String(v));
            }
        }
        if (children) {
            for (const c of children) {
                if (c === null || c === undefined || c === false)
                    continue;
                el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
            }
        }
        return el;
    }

    const NUMERIC_SECTION_INDEX = {
        positive: 0,
        negative: 1,
        zero: 2,
    };
    const NAMED_COLORS = {
        black: '#000000',
        blue: '#0000ff',
        cyan: '#00ffff',
        green: '#00ff00',
        magenta: '#ff00ff',
        red: '#ff0000',
        white: '#ffffff',
        yellow: '#ffff00',
    };
    function formatNumber(value, formatCode, options) {
        if (formatCode === '' || formatCode.toLowerCase() === 'general') {
            return formatGeneral(value);
        }
        if (formatCode.trim() === '@') {
            return { text: value, numeric: false, color: null };
        }
        const num = Number(value);
        if (!Number.isFinite(num)) {
            return { text: value, numeric: false, color: null };
        }
        const sections = splitSections(formatCode);
        const section = selectSection(sections, num);
        const { code: sectionWithSymbol } = extractLocaleCurrency(section.code);
        const color = extractColorModifier(sectionWithSymbol);
        const cleaned = stripSquareBracketModifiers(sectionWithSymbol);
        const magnitude = Math.abs(num);
        const body = applyFormat(cleaned, magnitude, options);
        const text = (num < 0 && section.prefixMinus) ? '-' + body : body;
        return { text, numeric: true, color };
    }
    function selectSection(sections, num) {
        const predicates = sections.map(parseSectionPredicate);
        const anyConditional = predicates.some((p) => p.predicate !== null);
        if (anyConditional) {
            let fallback = null;
            for (let i = 0; i < sections.length; i++) {
                const p = predicates[i];
                if (p.predicate === null) {
                    if (fallback === null && !isTextSection(sections[i]))
                        fallback = sections[i];
                    continue;
                }
                if (evaluatePredicate(p.predicate, num)) {
                    return { code: p.body, prefixMinus: false };
                }
            }
            if (fallback !== null)
                return { code: fallback, prefixMinus: false };
            return { code: sections[0], prefixMinus: false };
        }
        let sectionIndex;
        if (num > 0)
            sectionIndex = NUMERIC_SECTION_INDEX.positive;
        else if (num < 0)
            sectionIndex = sections[NUMERIC_SECTION_INDEX.negative] ? NUMERIC_SECTION_INDEX.negative : NUMERIC_SECTION_INDEX.positive;
        else
            sectionIndex = sections[NUMERIC_SECTION_INDEX.zero] ? NUMERIC_SECTION_INDEX.zero : NUMERIC_SECTION_INDEX.positive;
        const code = sections[sectionIndex] ?? sections[0] ?? '';
        const prefixMinus = num < 0 && sectionIndex === NUMERIC_SECTION_INDEX.positive;
        return { code, prefixMinus };
    }
    function isTextSection(code) {
        const stripped = stripQuotedLiterals(code);
        return /@/.test(stripped) && !/[0#?]/.test(stripped);
    }
    function parseSectionPredicate(code) {
        const m = /^\[([<>=]+)(-?\d+(?:\.\d+)?)\]/.exec(code);
        if (!m)
            return { predicate: null, body: code };
        const op = m[1];
        if (!['>', '<', '=', '>=', '<=', '<>'].includes(op)) {
            return { predicate: null, body: code };
        }
        const value = Number(m[2]);
        if (!Number.isFinite(value))
            return { predicate: null, body: code };
        return { predicate: { op, value }, body: code.slice(m[0].length) };
    }
    function evaluatePredicate(p, num) {
        switch (p.op) {
            case '>': return num > p.value;
            case '<': return num < p.value;
            case '=': return num === p.value;
            case '>=': return num >= p.value;
            case '<=': return num <= p.value;
            case '<>': return num !== p.value;
        }
        return false;
    }
    function extractColorModifier(code) {
        let match;
        const re = /\[([^\]]*)\]/g;
        while ((match = re.exec(code)) !== null) {
            const inner = match[1].trim();
            const named = NAMED_COLORS[inner.toLowerCase()];
            if (named)
                return named;
            const colorM = /^color\s*(\d+)$/i.exec(inner);
            if (colorM) {
                const idx = Number(colorM[1]);
                const hex = indexedColor(idx - 1);
                if (hex)
                    return hex;
            }
        }
        return null;
    }
    function extractLocaleCurrency(code) {
        return {
            code: code.replace(/\[\$([^\]]*)\]/g, (_, inner) => {
                const dashIdx = inner.indexOf('-');
                const symbol = dashIdx >= 0 ? inner.slice(0, dashIdx) : inner;
                if (!symbol)
                    return '';
                return `"${symbol.replace(/"/g, '')}"`;
            }),
        };
    }
    function splitSections(code) {
        const out = [];
        let current = '';
        let inQuote = false;
        for (let i = 0; i < code.length; i++) {
            const ch = code[i];
            if (ch === '"') {
                inQuote = !inQuote;
                current += ch;
                continue;
            }
            if (ch === '\\' && i + 1 < code.length) {
                current += ch + code[i + 1];
                i++;
                continue;
            }
            if (ch === ';' && !inQuote) {
                out.push(current);
                current = '';
                continue;
            }
            current += ch;
        }
        out.push(current);
        return out;
    }
    function stripSquareBracketModifiers(code) {
        return code.replace(/\[([^\]]*)\]/g, (match, inner) => {
            if (/^(hh?|mm?|ss?)$/i.test(inner))
                return match;
            return '';
        });
    }
    function applyFormat(code, value, options) {
        if (code.trim().toLowerCase() === 'general') {
            return formatGeneralNumber(value);
        }
        if (/\[(hh?|mm?|ss?)\]/.test(code)) {
            return formatDateTime(code, value, options);
        }
        if (isFractionCode(code)) {
            return formatFraction(code, value);
        }
        if (/[eE][+-]/.test(stripQuotedLiterals(code))) {
            return formatScientific(code, value);
        }
        if (/[yMdhsmAP]/.test(stripQuotedLiterals(code))) {
            if (/[yMdAPhs]/.test(stripQuotedLiterals(code))) {
                return formatDateTime(code, value, options);
            }
        }
        return formatNumeric(code, value);
    }
    function stripQuotedLiterals(code) {
        return code
            .replace(/\\./g, '')
            .replace(/"([^"]*)"/g, '')
            .replace(/\[[^\]]*\]/g, '');
    }
    function formatNumeric(code, value) {
        const percentCount = (stripQuotedLiterals(code).match(/%/g) ?? []).length;
        let working = value;
        for (let i = 0; i < percentCount; i++)
            working *= 100;
        const dotIdx = firstUnquotedIndexOf(code, '.');
        let decimals = 0;
        if (dotIdx >= 0) {
            for (let i = dotIdx + 1; i < code.length; i++) {
                const c = code[i];
                if (c === '0' || c === '#' || c === '?')
                    decimals++;
                else if (c === '%' || c === ',' || c === ' ')
                    continue;
                else
                    break;
            }
        }
        const thousands = /#,##0|0,000/.test(stripQuotedLiterals(code));
        const fixed = Math.abs(working).toFixed(decimals);
        const [intPart, fracPart] = fixed.split('.');
        const withSep = thousands ? insertThousands(intPart) : intPart;
        const body = fracPart != null ? `${withSep}.${fracPart}` : withSep;
        const signed = working < 0 ? '-' + body : body;
        return renderLiteralsAroundNumber(code, signed);
    }
    function firstUnquotedIndexOf(code, ch) {
        let inQuote = false;
        for (let i = 0; i < code.length; i++) {
            const c = code[i];
            if (c === '"') {
                inQuote = !inQuote;
                continue;
            }
            if (c === '\\') {
                i++;
                continue;
            }
            if (c === ch && !inQuote)
                return i;
        }
        return -1;
    }
    function insertThousands(intPart) {
        return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }
    function renderLiteralsAroundNumber(code, numberText) {
        const hasPlaceholder = /[0#?]/.test(stripQuotedLiterals(code));
        let out = '';
        let inserted = false;
        let i = 0;
        while (i < code.length) {
            const c = code[i];
            if (c === '"') {
                i++;
                while (i < code.length && code[i] !== '"') {
                    out += code[i];
                    i++;
                }
                i++;
                continue;
            }
            if (c === '\\') {
                if (i + 1 < code.length)
                    out += code[i + 1];
                i += 2;
                continue;
            }
            if (c === '_') {
                if (i + 1 < code.length)
                    out += ' ';
                i += 2;
                continue;
            }
            if (c === '*') {
                i += 2;
                continue;
            }
            if (c === '0' || c === '#' || c === '?' || c === '.' || c === ',') {
                if (!inserted) {
                    out += numberText;
                    inserted = true;
                }
                while (i < code.length && '0#?.,'.includes(code[i]))
                    i++;
                continue;
            }
            if (c === '%') {
                out += '%';
                i++;
                continue;
            }
            out += c;
            i++;
        }
        if (!inserted && hasPlaceholder)
            out = numberText + out;
        return out;
    }
    function formatGeneral(value) {
        const n = Number(value);
        if (!Number.isFinite(n))
            return { text: value, numeric: false, color: null };
        return { text: formatGeneralNumber(n), numeric: true, color: null };
    }
    function formatGeneralNumber(n) {
        return n.toString();
    }
    function isFractionCode(code) {
        return /[0#]\s+\?+\/(\?+|\d+)/.test(code);
    }
    function formatFraction(code, value) {
        const m = /(\?+)\/(\?+|\d+)/.exec(code);
        if (!m)
            return String(value);
        const denomRaw = m[2];
        const fixedDenom = /^\d+$/.test(denomRaw) ? Number(denomRaw) : null;
        const denomPlaces = denomRaw.length;
        const maxDenom = fixedDenom ?? Math.pow(10, denomPlaces);
        const sign = value < 0 ? '-' : '';
        const abs = Math.abs(value);
        const whole = Math.floor(abs);
        const frac = abs - whole;
        let num, den;
        if (fixedDenom !== null && fixedDenom > 0) {
            den = fixedDenom;
            num = Math.round(frac * fixedDenom);
        }
        else {
            ({ num, den } = bestFraction(frac, maxDenom));
        }
        let outWhole = whole;
        let outNum = num;
        if (outNum === den && den !== 0) {
            outWhole += 1;
            outNum = 0;
        }
        if (outNum === 0) {
            return sign + String(outWhole);
        }
        return `${sign}${outWhole} ${outNum}/${den}`;
    }
    function bestFraction(x, maxDenom) {
        if (x === 0)
            return { num: 0, den: 1 };
        let h1 = 1, k1 = 0;
        let h = Math.floor(x), k = 1;
        let rem = x - h;
        while (rem > 1e-12) {
            const inv = 1 / rem;
            const a = Math.floor(inv);
            const newH = a * h + h1;
            const newK = a * k + k1;
            if (newK > maxDenom)
                break;
            h1 = h;
            k1 = k;
            h = newH;
            k = newK;
            rem = inv - a;
        }
        return { num: h, den: k };
    }
    function formatScientific(code, value) {
        const split = findExponentSplit(code);
        if (!split)
            return formatNumeric(code, value);
        const { mantissaCode, expSign, exponentCode, before, after } = split;
        const intPlaceholders = countIntegerPlaceholders(mantissaCode);
        const decimalPlaces = countDecimalPlaceholders(mantissaCode);
        const engineering = intPlaceholders >= 3;
        let mantissa;
        let exponent;
        if (value === 0) {
            mantissa = 0;
            exponent = 0;
        }
        else {
            exponent = Math.floor(Math.log10(Math.abs(value)));
            if (engineering) {
                exponent = Math.floor(exponent / 3) * 3;
            }
            mantissa = value / Math.pow(10, exponent);
            const rounded = Number(mantissa.toFixed(decimalPlaces));
            const boundary = engineering ? Math.pow(10, intPlaceholders) : 10;
            if (Math.abs(rounded) >= boundary) {
                mantissa = rounded / 10;
                exponent += 1;
                if (engineering) {
                    const misalign = exponent % 3;
                    if (misalign !== 0) {
                        mantissa *= Math.pow(10, misalign);
                        exponent -= misalign;
                    }
                }
            }
        }
        const mantissaText = formatNumeric(mantissaCode, mantissa);
        const expAbs = Math.abs(exponent);
        const expDigits = exponentCode.replace(/[^0#?]/g, '').length;
        const expBody = String(expAbs).padStart(expDigits, '0');
        const sign = exponent < 0 ? '-' : (expSign === '+' ? '+' : '');
        const expText = sign + expBody;
        return `${before}${mantissaText}E${expText}${after}`;
    }
    function findExponentSplit(code) {
        let inQuote = false;
        let inBracket = false;
        let eIdx = -1;
        let sign = '+';
        for (let i = 0; i < code.length - 1; i++) {
            const c = code[i];
            if (c === '"') {
                inQuote = !inQuote;
                continue;
            }
            if (!inQuote && c === '[') {
                inBracket = true;
                continue;
            }
            if (!inQuote && c === ']') {
                inBracket = false;
                continue;
            }
            if (inQuote || inBracket)
                continue;
            if (c === '\\') {
                i++;
                continue;
            }
            if ((c === 'E' || c === 'e') && (code[i + 1] === '+' || code[i + 1] === '-')) {
                eIdx = i;
                sign = code[i + 1];
                break;
            }
        }
        if (eIdx < 0)
            return null;
        let mStart = eIdx;
        while (mStart > 0) {
            const c = code[mStart - 1];
            if ('0#?.,'.includes(c)) {
                mStart--;
                continue;
            }
            break;
        }
        let eEnd = eIdx + 2;
        while (eEnd < code.length) {
            const c = code[eEnd];
            if ('0#?'.includes(c)) {
                eEnd++;
                continue;
            }
            break;
        }
        return {
            before: code.slice(0, mStart),
            mantissaCode: code.slice(mStart, eIdx),
            expSign: sign,
            exponentCode: code.slice(eIdx + 2, eEnd),
            after: code.slice(eEnd),
        };
    }
    function countIntegerPlaceholders(code) {
        const dotIdx = code.indexOf('.');
        const slice = dotIdx < 0 ? code : code.slice(0, dotIdx);
        return (slice.match(/[0#?]/g) ?? []).length;
    }
    function countDecimalPlaceholders(code) {
        const dotIdx = code.indexOf('.');
        if (dotIdx < 0)
            return 0;
        let n = 0;
        for (let i = dotIdx + 1; i < code.length; i++) {
            const c = code[i];
            if (c === '0' || c === '#' || c === '?')
                n++;
            else
                break;
        }
        return n;
    }
    const EPOCH_MS_1900 = Date.UTC(1899, 11, 30);
    const EPOCH_MS_1904 = Date.UTC(1904, 0, 1);
    const MS_PER_DAY = 86400000;
    function serialToDate(serial, date1904) {
        const whole = Math.floor(serial);
        const frac = serial - whole;
        const epoch = date1904 ? EPOCH_MS_1904 : EPOCH_MS_1900;
        const ms = epoch + whole * MS_PER_DAY + Math.round(frac * MS_PER_DAY);
        return new Date(ms);
    }
    function formatDateTime(code, value, options) {
        const date1904 = options?.date1904 === true;
        const date = serialToDate(value, date1904);
        const Y = date.getUTCFullYear();
        const M = date.getUTCMonth() + 1;
        const D = date.getUTCDate();
        const h24 = date.getUTCHours();
        const m = date.getUTCMinutes();
        const s = date.getUTCSeconds();
        const ampm = /AM\/PM|am\/pm/.test(stripQuotedLiterals(code));
        const h12 = ((h24 + 11) % 12) + 1;
        const hour = ampm ? h12 : h24;
        const totalHours = Math.floor(value * 24);
        const totalMinutes = Math.floor(value * 1440);
        const totalSeconds = Math.floor(value * 86400);
        let out = '';
        let i = 0;
        let lastSawHour = false;
        while (i < code.length) {
            const c = code[i];
            if (c === '"') {
                i++;
                while (i < code.length && code[i] !== '"') {
                    out += code[i];
                    i++;
                }
                i++;
                continue;
            }
            if (c === '\\' && i + 1 < code.length) {
                out += code[i + 1];
                i += 2;
                continue;
            }
            if (c === '[') {
                const end = code.indexOf(']', i);
                if (end < 0) {
                    i++;
                    continue;
                }
                const inner = code.slice(i + 1, end).toLowerCase();
                i = end + 1;
                if (inner === 'h' || inner === 'hh') {
                    out += inner.length >= 2 ? pad2(totalHours) : String(totalHours);
                    lastSawHour = true;
                }
                else if (inner === 'm' || inner === 'mm') {
                    out += inner.length >= 2 ? pad2(totalMinutes) : String(totalMinutes);
                    lastSawHour = false;
                }
                else if (inner === 's' || inner === 'ss') {
                    out += inner.length >= 2 ? pad2(totalSeconds) : String(totalSeconds);
                    lastSawHour = false;
                }
                continue;
            }
            const run = readRun(code, i, c.toLowerCase());
            if (run > 0) {
                const token = code.slice(i, i + run).toLowerCase();
                switch (token[0]) {
                    case 'y':
                        out += token.length >= 4 ? String(Y).padStart(4, '0') : String(Y % 100).padStart(2, '0');
                        lastSawHour = false;
                        break;
                    case 'd':
                        out += token.length >= 4 ? weekdayName(date, true)
                            : token.length === 3 ? weekdayName(date, false)
                                : token.length === 2 ? pad2(D)
                                    : String(D);
                        lastSawHour = false;
                        break;
                    case 'h':
                        out += token.length >= 2 ? pad2(hour) : String(hour);
                        lastSawHour = true;
                        break;
                    case 's':
                        out += token.length >= 2 ? pad2(s) : String(s);
                        lastSawHour = false;
                        break;
                    case 'm':
                        if (lastSawHour) {
                            out += token.length >= 2 ? pad2(m) : String(m);
                        }
                        else {
                            out += token.length >= 4 ? monthName(M, true)
                                : token.length === 3 ? monthName(M, false)
                                    : token.length === 2 ? pad2(M)
                                        : String(M);
                        }
                        break;
                    case 'a':
                    case 'p':
                        if (token === 'am/pm') {
                            out += h24 < 12 ? 'AM' : 'PM';
                        }
                        else {
                            out += code.slice(i, i + run);
                        }
                        break;
                }
                i += run;
                continue;
            }
            out += c;
            i++;
        }
        return out;
    }
    function readRun(code, start, token) {
        const lower = code.slice(start).toLowerCase();
        if (lower.startsWith('am/pm'))
            return 5;
        if (!'ymdhs'.includes(token))
            return 0;
        let n = 0;
        while (start + n < code.length && code[start + n].toLowerCase() === token)
            n++;
        return n;
    }
    function pad2(n) {
        return n < 10 ? '0' + n : String(n);
    }
    const MONTHS_FULL = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December',
    ];
    const MONTHS_ABBR = MONTHS_FULL.map((m) => m.slice(0, 3));
    function monthName(m, full) {
        const idx = m - 1;
        if (idx < 0 || idx > 11)
            return String(m);
        return (full ? MONTHS_FULL : MONTHS_ABBR)[idx];
    }
    const DAYS_FULL = [
        'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
    ];
    const DAYS_ABBR = DAYS_FULL.map((d) => d.slice(0, 3));
    function weekdayName(date, full) {
        return (full ? DAYS_FULL : DAYS_ABBR)[date.getUTCDay()];
    }

    const A1_REF = /(\$?)([A-Z]+)(\$?)([1-9][0-9]*)/g;
    function a1ToR1c1(formula, anchorRow, anchorCol) {
        return formula.replace(A1_REF, (_, colAbs, col, rowAbs, rowStr) => {
            const colIdx = columnLettersToIndex(col);
            const rowIdx = Number(rowStr) - 1;
            if (colIdx < 0 || !Number.isFinite(rowIdx))
                return _;
            const rowPart = rowAbs ? `R${rowIdx + 1}` : `R[${rowIdx - anchorRow}]`;
            const colPart = colAbs ? `C${colIdx + 1}` : `C[${colIdx - anchorCol}]`;
            return rowPart + colPart;
        });
    }
    const R1C1_REF = /R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?/g;
    function r1c1ToA1(formula, anchorRow, anchorCol) {
        return formula.replace(R1C1_REF, (match, rowPart, colPart) => {
            if (rowPart === undefined && colPart === undefined)
                return match;
            const { index: rowIdx, abs: rowAbs } = decodeAxis(rowPart, anchorRow);
            const { index: colIdx, abs: colAbs } = decodeAxis(colPart, anchorCol);
            if (rowIdx === null || colIdx === null)
                return match;
            const col = indexToColumnLetters(colIdx);
            const prefixCol = colAbs ? '$' : '';
            const prefixRow = rowAbs ? '$' : '';
            return `${prefixCol}${col}${prefixRow}${rowIdx + 1}`;
        });
    }
    function decodeAxis(part, anchor) {
        if (part === undefined)
            return { index: anchor, abs: false };
        if (part.startsWith('[') && part.endsWith(']')) {
            const delta = Number(part.slice(1, -1));
            if (!Number.isFinite(delta))
                return { index: null, abs: false };
            return { index: anchor + delta, abs: false };
        }
        const n = Number(part);
        if (!Number.isFinite(n) || n < 1)
            return { index: null, abs: false };
        return { index: n - 1, abs: true };
    }

    const TL_COLORS = ['#d13438', '#ffb900', '#107c10'];
    const ARROW_COLORS = ['#d13438', '#ffb900', '#107c10'];
    function circle(color) {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><circle cx="6" cy="6" r="5" fill="${color}"/></svg>`;
    }
    function arrow(color, angleDeg) {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><g transform="rotate(${angleDeg} 6 6)"><path d="M6 1 L10 7 L7 7 L7 11 L5 11 L5 7 L2 7 Z" fill="${color}"/></g></svg>`;
    }
    function cross() {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M3 3 L9 9 M9 3 L3 9" stroke="#d13438" stroke-width="2" stroke-linecap="round"/></svg>`;
    }
    function bang() {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M6 1 L11 11 L1 11 Z" fill="#ffb900" stroke="#8a6300" stroke-width="0.5"/><path d="M6 5 L6 8" stroke="#202020" stroke-width="1" stroke-linecap="round"/><circle cx="6" cy="10" r="0.6" fill="#202020"/></svg>`;
    }
    function tick() {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M2 7 L5 10 L10 3" stroke="#107c10" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    }
    const SETS = {
        '3TrafficLights1': [circle(TL_COLORS[0]), circle(TL_COLORS[1]), circle(TL_COLORS[2])],
        '3Arrows': [arrow(ARROW_COLORS[0], 180), arrow(ARROW_COLORS[1], 90), arrow(ARROW_COLORS[2], 0)],
        '3Symbols': [cross(), bang(), tick()],
        '3Symbols2': [cross(), bang(), tick()],
    };
    function renderIcon(set, index, reverse) {
        const icons = SETS[set];
        if (!icons)
            return null;
        const effective = reverse ? icons.length - 1 - index : index;
        if (effective < 0 || effective >= icons.length)
            return null;
        return icons[effective];
    }

    const PX_PER_CHAR = 7;
    const PADDING_PX = 5;
    function charWidthToPx(width) {
        return Math.round(width * PX_PER_CHAR + PADDING_PX);
    }
    class HtmlRenderer {
        async render(workbook, options) {
            const nodes = [];
            nodes.push(renderStyle(options.className));
            for (const sheet of workbook.sheets) {
                if (sheet.state !== 'visible')
                    continue;
                nodes.push(renderSheet(sheet, workbook.styles, workbook.theme, workbook.date1904, options));
            }
            return nodes;
        }
    }
    function renderStyle(className) {
        const style = document.createElement('style');
        style.setAttribute('data-xlsxjs', '');
        style.textContent = `
.${className} { font-family: system-ui, sans-serif; }
.${className} table { border-collapse: separate; border-spacing: 0; }
.${className} th, .${className} td { border: 1px solid #d0d0d0; padding: 2px 6px; }
.${className} th { background: #f3f3f3; font-weight: normal; color: #555; }
.${className} td.xlsx-numeric { text-align: right; font-variant-numeric: tabular-nums; }
.${className} .xlsx-sheet-name { margin: 1rem 0 0.25rem; font-size: 1rem; font-weight: 600; }
.${className} .xlsx-frozen-col { position: sticky; left: 0; z-index: 1; background: inherit; }
.${className} .xlsx-frozen-row { position: sticky; top: 0; z-index: 2; background: inherit; }
.${className} .xlsx-frozen-both { position: sticky; left: 0; top: 0; z-index: 3; background: inherit; }
.${className} .xlsx-autofilter::after { content: " ▾"; color: #888; font-size: 0.85em; }
.${className} .xlsx-validation-list::after { content: " ▾"; color: #888; font-size: 0.85em; }
.${className} a.xlsx-hyperlink { color: #0563c1; text-decoration: underline; }
.${className} .xlsx-table-caption { font-size: 0.85em; color: #666; margin: 0.25rem 0 0; }
.${className} .xlsx-chart-placeholder {
    border: 1px dashed #999; padding: 1em; margin: 0.5em 0;
    color: #666; font-size: 0.9em; text-align: center;
}
.${className} .xlsx-shape {
    border: 1px solid #e0e0e0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #555; font-size: 0.9em;
}
.${className} .xlsx-shape[data-kind="connector"] {
    border-style: dashed; color: #888;
}
.${className} .xlsx-shape pre {
    margin: 0; white-space: pre-line;
    font-family: inherit; font-size: inherit;
}
.${className} .xlsx-spill-anchor { outline: 1px dashed #0066cc; outline-offset: -1px; }
.${className} .xlsx-comment-marker { color: #c00; margin-left: 4px; cursor: help; }
.${className} .xlsx-threaded { color: #0066cc; margin-left: 4px; cursor: help; }
.${className} .xlsx-shrink-to-fit { font-size: clamp(0.55em, 0.95em, 1em); overflow: hidden; }
.${className}.xlsx-no-gridlines th, .${className}.xlsx-no-gridlines td { border: none; }
.${className}.xlsx-no-headers thead tr > th:first-child,
.${className}.xlsx-no-headers tbody tr > th:first-child { display: none; }
.${className}.xlsx-no-headers thead tr:first-child { display: none; }
.${className} .xlsx-outline-1 > th:first-child { padding-left: 0.5rem; }
.${className} .xlsx-outline-2 > th:first-child { padding-left: 1rem; }
.${className} .xlsx-outline-3 > th:first-child { padding-left: 1.5rem; }
.${className} .xlsx-outline-4 > th:first-child { padding-left: 2rem; }
.${className} .xlsx-outline-5 > th:first-child { padding-left: 2.5rem; }
.${className} .xlsx-outline-6 > th:first-child { padding-left: 3rem; }
.${className} .xlsx-outline-7 > th:first-child { padding-left: 3.5rem; }
.${className} col.xlsx-outline-1 { border-left: 2px solid #ddd; }
.${className} col.xlsx-outline-2 { border-left: 3px solid #ccc; }
.${className} col.xlsx-outline-3 { border-left: 4px solid #bbb; }
.${className} .xlsx-header, .${className} .xlsx-footer {
    display: grid; grid-template-columns: 1fr 1fr 1fr;
    font-size: 0.85em; color: #666; margin: 0.5em 0;
}
    `.trim();
        return style;
    }
    function makeCfContext(sheet, date1904) {
        return {
            date1904,
            cellsInRange(range) {
                const out = [];
                for (let r = range.row; r <= range.endRow; r++) {
                    const row = sheet.rows[r];
                    for (let c = range.col; c <= range.endCol; c++) {
                        const cell = row ? row.find((x) => x.col === c) ?? null : null;
                        out.push({ col: c, row: r, cell });
                    }
                }
                return out;
            },
        };
    }
    function rangeContains(range, row, col) {
        return row >= range.row && row <= range.endRow && col >= range.col && col <= range.endCol;
    }
    function resolveGraphicalConditionalFormats(sheet, theme, date1904) {
        const out = new Map();
        const ctx = makeCfContext(sheet, date1904);
        const pairs = [];
        for (const block of sheet.conditionalFormatting) {
            for (const range of block.ranges) {
                for (const rule of block.rules) {
                    if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
                        pairs.push({ range, rule });
                    }
                }
            }
        }
        if (pairs.length === 0)
            return out;
        pairs.sort((a, b) => a.rule.priority - b.rule.priority);
        for (const { range, rule } of pairs) {
            const cells = ctx.cellsInRange(range);
            const values = [];
            for (const entry of cells) {
                if (!entry.cell)
                    continue;
                const n = Number(entry.cell.value);
                if (Number.isFinite(n))
                    values.push(n);
            }
            if (values.length === 0)
                continue;
            if (rule.type === 'colorScale' && rule.colorScale) {
                applyColorScale(out, rule.colorScale, cells, values, theme);
            }
            else if (rule.type === 'dataBar' && rule.dataBar) {
                applyDataBar(out, rule.dataBar, cells, values, theme);
            }
            else if (rule.type === 'iconSet' && rule.iconSet) {
                applyIconSet(out, rule.iconSet, cells, values);
            }
        }
        return out;
    }
    function applyIconSet(out, icons, cells, values) {
        const thresholds = [];
        for (const cfvo of icons.cfvos) {
            const t = resolveCfvo(cfvo, values);
            if (t === null)
                return;
            thresholds.push(t);
        }
        if (thresholds.length < 2)
            return;
        for (const entry of cells) {
            if (!entry.cell)
                continue;
            const n = Number(entry.cell.value);
            if (!Number.isFinite(n))
                continue;
            let idx = 0;
            for (let i = 1; i < thresholds.length; i++) {
                if (n >= thresholds[i])
                    idx = i;
            }
            const override = icons.customIcons?.[idx];
            const svg = override
                ? renderIcon(override.iconSet, override.iconId, false)
                : renderIcon(icons.iconSet, idx, icons.reverse);
            if (!svg)
                continue;
            const key = `${entry.row},${entry.col}`;
            const existing = out.get(key) ?? {};
            existing.icon = { svg, showValue: icons.showValue };
            out.set(key, existing);
        }
    }
    function applyColorScale(out, scale, cells, values, theme) {
        const stops = [];
        for (let i = 0; i < scale.cfvos.length; i++) {
            const t = resolveCfvo(scale.cfvos[i], values);
            const c = resolveColor(scale.colors[i] ?? null, theme);
            if (t !== null && c)
                stops.push({ threshold: t, hex: c });
        }
        if (stops.length < 2)
            return;
        stops.sort((a, b) => a.threshold - b.threshold);
        for (const entry of cells) {
            if (!entry.cell)
                continue;
            const n = Number(entry.cell.value);
            if (!Number.isFinite(n))
                continue;
            const hex = interpolateColorScale(n, stops);
            if (!hex)
                continue;
            const key = `${entry.row},${entry.col}`;
            const existing = out.get(key) ?? {};
            existing.colorScaleBg = hex;
            out.set(key, existing);
        }
    }
    function applyDataBar(out, bar, cells, values, theme) {
        const min = bar.cfvos[0] ? resolveCfvo(bar.cfvos[0], values) : Math.min(...values);
        const max = bar.cfvos[1] ? resolveCfvo(bar.cfvos[1], values) : Math.max(...values);
        if (min === null || max === null || max === min)
            return;
        const color = resolveColor(bar.color, theme);
        if (!color)
            return;
        const negFill = resolveColor(bar.negativeFillColor, theme);
        const resolvedBorderColor = resolveColor(bar.borderColor, theme);
        const axisColor = resolveColor(bar.axisColor, theme) ?? '#000000';
        const lenMin = Math.max(0, bar.minLength) / 100;
        const lenMax = Math.min(100, bar.maxLength) / 100;
        for (const entry of cells) {
            if (!entry.cell)
                continue;
            const n = Number(entry.cell.value);
            if (!Number.isFinite(n))
                continue;
            const clamped = Math.max(min, Math.min(max, n));
            const t = (clamped - min) / (max - min);
            const fraction = lenMin + (lenMax - lenMin) * t;
            const fillColor = n < 0 && negFill ? negFill : color;
            let axis = null;
            if (bar.axisPosition === 'middle') {
                axis = { position: 'middle', color: axisColor };
            }
            else if (bar.axisPosition === 'automatic' && min < 0 && max > 0) {
                axis = { position: 'middle', color: axisColor };
            }
            else if (bar.axisPosition === 'automatic') {
                axis = { position: 'left', color: axisColor };
            }
            const key = `${entry.row},${entry.col}`;
            const existing = out.get(key) ?? {};
            existing.dataBar = {
                color: fillColor,
                fraction,
                gradient: bar.gradient,
                direction: bar.direction,
                border: bar.border,
                borderColor: resolvedBorderColor,
                axis,
            };
            out.set(key, existing);
        }
    }
    function resolveConditionalFormats(sheet, styles, date1904) {
        const out = new Map();
        if (!styles?.dxfs.length)
            return out;
        const blocks = sheet.conditionalFormatting;
        if (!blocks.length)
            return out;
        const ctx = makeCfContext(sheet, date1904);
        const pairs = [];
        for (const block of blocks) {
            for (const range of block.ranges) {
                for (const rule of block.rules)
                    pairs.push({ range, rule });
            }
        }
        pairs.sort((a, b) => a.rule.priority - b.rule.priority);
        const stopped = new Set();
        for (let r = 0; r <= sheet.maxRow; r++) {
            for (let c = 0; c <= sheet.maxCol; c++) {
                const key = `${r},${c}`;
                for (const { range, rule } of pairs) {
                    if (!rangeContains(range, r, c))
                        continue;
                    if (out.has(key)) {
                        if (stopped.has(key))
                            break;
                        continue;
                    }
                    const cell = sheet.rows[r]?.find((x) => x.col === c) ?? null;
                    if (!evaluateRule(rule, cell, range, ctx))
                        continue;
                    const dxf = styles.dxfs[rule.dxfId];
                    if (dxf)
                        out.set(key, dxf);
                    if (rule.stopIfTrue)
                        stopped.add(key);
                }
            }
        }
        return out;
    }
    function renderSheet(sheet, styles, theme, date1904, options) {
        const section = h('section', { class: options.className, 'data-sheet-name': sheet.name });
        applySheetView(section, sheet.view, theme);
        if (sheet.protection?.enabled) {
            section.setAttribute('data-sheet-protected', 'true');
        }
        if (sheet.pageBreaks.rows.length > 0) {
            section.setAttribute('data-page-break-rows', sheet.pageBreaks.rows.join(','));
        }
        if (sheet.pageBreaks.cols.length > 0) {
            section.setAttribute('data-page-break-cols', sheet.pageBreaks.cols.join(','));
        }
        section.appendChild(h('div', { class: 'xlsx-sheet-name' }, [sheet.name]));
        const table = h('table');
        if (sheet.maxCol < 0) {
            section.appendChild(table);
            return section;
        }
        const colCount = sheet.maxCol + 1;
        const colgroup = document.createElement('colgroup');
        colgroup.appendChild(document.createElement('col'));
        const widthByCol = new Map();
        const hiddenCols = new Set();
        const outlineByCol = new Map();
        for (const cw of sheet.columns) {
            for (let i = cw.min; i <= cw.max; i++) {
                if (cw.width !== null)
                    widthByCol.set(i, cw.width);
                if (cw.hidden)
                    hiddenCols.add(i);
                if (cw.outlineLevel > 0)
                    outlineByCol.set(i, cw.outlineLevel);
            }
        }
        for (let c = 0; c < colCount; c++) {
            const col = document.createElement('col');
            const w = widthByCol.get(c);
            if (w !== undefined)
                col.style.width = `${charWidthToPx(w)}px`;
            if (hiddenCols.has(c))
                col.style.display = 'none';
            const lvl = outlineByCol.get(c);
            if (lvl) {
                col.setAttribute('data-outline-level', String(lvl));
                col.classList.add(`xlsx-outline-${Math.min(lvl, 7)}`);
            }
            colgroup.appendChild(col);
        }
        table.appendChild(colgroup);
        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        headRow.appendChild(h('th'));
        for (let c = 0; c < colCount; c++) {
            const th = h('th', null, [indexToColumnLetters(c)]);
            if (hiddenCols.has(c))
                th.style.display = 'none';
            const lvl = outlineByCol.get(c);
            if (lvl) {
                th.setAttribute('data-outline-level', String(lvl));
                th.classList.add(`xlsx-outline-${Math.min(lvl, 7)}`);
            }
            headRow.appendChild(th);
        }
        thead.appendChild(headRow);
        table.appendChild(thead);
        const frozen = sheet.frozenPanes;
        const autoFilter = sheet.autoFilter;
        const mergeByAnchor = new Map();
        const suppressedCells = new Set();
        for (const m of sheet.merges) {
            mergeByAnchor.set(`${m.row},${m.col}`, m);
            for (let dr = 0; dr < m.rowSpan; dr++) {
                for (let dc = 0; dc < m.colSpan; dc++) {
                    if (dr === 0 && dc === 0)
                        continue;
                    suppressedCells.add(`${m.row + dr},${m.col + dc}`);
                }
            }
        }
        const dxfByCell = resolveConditionalFormats(sheet, styles, date1904);
        const graphicalByCell = resolveGraphicalConditionalFormats(sheet, theme, date1904);
        const commentByCell = new Map();
        for (const cmt of sheet.comments)
            commentByCell.set(`${cmt.row},${cmt.col}`, cmt);
        const threadedByCell = new Map();
        for (const entry of sheet.threadedComments) {
            const key = `${entry.row},${entry.col}`;
            const list = threadedByCell.get(key);
            if (list)
                list.push(entry);
            else
                threadedByCell.set(key, [entry]);
        }
        const rowDim = new Map();
        for (const d of sheet.rowDimensions)
            rowDim.set(d.row, d);
        const hyperlinkByCell = new Map();
        for (const h of sheet.hyperlinks)
            hyperlinkByCell.set(`${h.row},${h.col}`, h);
        const validationByCell = new Map();
        for (const v of sheet.dataValidationLists) {
            const opts = v.options ? v.options.join('|') : null;
            for (let r = v.row; r <= v.endRow; r++) {
                for (let c = v.col; c <= v.endCol; c++) {
                    validationByCell.set(`${r},${c}`, opts);
                }
            }
        }
        const tbody = document.createElement('tbody');
        const rowCount = sheet.maxRow + 1;
        for (let r = 0; r < rowCount; r++) {
            const tr = document.createElement('tr');
            const dim = rowDim.get(r);
            if (dim?.hidden)
                tr.style.display = 'none';
            if (dim?.height != null)
                tr.style.height = `${(dim.height * 4 / 3).toFixed(2)}px`;
            if (dim && dim.outlineLevel > 0) {
                tr.setAttribute('data-outline-level', String(dim.outlineLevel));
                tr.classList.add(`xlsx-outline-${Math.min(dim.outlineLevel, 7)}`);
            }
            tr.appendChild(h('th', null, [String(r + 1)]));
            const cells = sheet.rows[r];
            const byCol = {};
            if (cells)
                for (const cell of cells)
                    byCol[cell.col] = cell;
            for (let c = 0; c < colCount; c++) {
                if (suppressedCells.has(`${r},${c}`))
                    continue;
                const cell = byCol[c];
                const td = document.createElement('td');
                if (cell)
                    renderCellContent(td, cell, styles, theme, date1904, options);
                if (cell?.isSpillAnchor)
                    td.classList.add('xlsx-spill-anchor');
                const hlink = hyperlinkByCell.get(`${r},${c}`);
                if (hlink)
                    wrapCellWithHyperlink(td, hlink);
                const dxf = dxfByCell.get(`${r},${c}`);
                if (dxf)
                    applyDxf(td, dxf, theme, cell ?? null, date1904);
                const gfx = graphicalByCell.get(`${r},${c}`);
                if (gfx)
                    applyGraphicalCf(td, gfx);
                const cmt = commentByCell.get(`${r},${c}`);
                if (cmt)
                    appendCommentMarker(td, cmt);
                if (hiddenCols.has(c))
                    td.style.display = 'none';
                if (frozen)
                    tagFrozen(td, r, c, frozen);
                if (autoFilter && r === autoFilter.row && c >= autoFilter.col && c <= autoFilter.endCol) {
                    td.classList.add('xlsx-autofilter');
                }
                if (validationByCell.has(`${r},${c}`)) {
                    td.classList.add('xlsx-validation-list');
                    const opts = validationByCell.get(`${r},${c}`);
                    if (opts !== null && opts !== undefined) {
                        td.setAttribute('data-validation-options', opts);
                    }
                }
                const threaded = threadedByCell.get(`${r},${c}`);
                if (threaded && threaded.length)
                    appendThreadedCommentMarker(td, threaded);
                const merge = mergeByAnchor.get(`${r},${c}`);
                if (merge) {
                    if (merge.colSpan > 1)
                        td.setAttribute('colspan', String(merge.colSpan));
                    if (merge.rowSpan > 1)
                        td.setAttribute('rowspan', String(merge.rowSpan));
                    td.classList.add('xlsx-merged');
                }
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        section.appendChild(table);
        for (const t of sheet.tables) {
            const caption = document.createElement('div');
            caption.className = 'xlsx-table-caption';
            caption.setAttribute('data-table-name', t.name);
            caption.setAttribute('data-table-display-name', t.displayName);
            if (t.altText) {
                caption.setAttribute('aria-label', t.altText);
            }
            else if (t.altTextSummary) {
                caption.setAttribute('title', t.altTextSummary);
            }
            caption.textContent = `Table "${t.displayName || t.name}" · ${t.columns.length} column(s) · rows ${t.row + 1}-${t.endRow + 1}`;
            section.appendChild(caption);
        }
        for (const img of sheet.images) {
            const fig = document.createElement('figure');
            fig.className = 'xlsx-image';
            fig.setAttribute('data-anchor-mode', img.anchorMode);
            fig.setAttribute('data-anchor-col', String(img.col));
            fig.setAttribute('data-anchor-row', String(img.row));
            if (img.endCol !== null)
                fig.setAttribute('data-anchor-end-col', String(img.endCol));
            if (img.endRow !== null)
                fig.setAttribute('data-anchor-end-row', String(img.endRow));
            fig.style.margin = '0.5rem 0';
            if (img.anchorMode === 'absolute') {
                fig.style.position = 'absolute';
                if (img.absoluteX !== null)
                    fig.style.left = `${emuToPx(img.absoluteX)}px`;
                if (img.absoluteY !== null)
                    fig.style.top = `${emuToPx(img.absoluteY)}px`;
            }
            const el = document.createElement('img');
            el.src = img.dataUrl;
            if (img.decorative) {
                el.alt = '';
                el.setAttribute('aria-hidden', 'true');
            }
            else if (img.alt) {
                el.alt = img.alt;
            }
            if (img.widthEmu && img.heightEmu) {
                el.width = emuToPx(img.widthEmu);
                el.height = emuToPx(img.heightEmu);
            }
            el.style.maxWidth = '100%';
            fig.appendChild(el);
            section.appendChild(fig);
        }
        for (const chart of sheet.charts) {
            const ph = document.createElement('div');
            ph.className = 'xlsx-chart-placeholder';
            ph.setAttribute('data-chart-kind', chart.kind);
            if (chart.chartType)
                ph.setAttribute('data-chart-type', chart.chartType);
            ph.setAttribute('data-anchor-col', String(chart.col));
            ph.setAttribute('data-anchor-row', String(chart.row));
            if (chart.endCol !== null)
                ph.setAttribute('data-anchor-end-col', String(chart.endCol));
            if (chart.endRow !== null)
                ph.setAttribute('data-anchor-end-row', String(chart.endRow));
            ph.textContent = `[chart: ${chart.chartType ?? chart.kind}]`;
            section.appendChild(ph);
        }
        for (const shape of sheet.shapes) {
            section.appendChild(renderShape(shape));
        }
        if (sheet.headerFooter?.oddHeader) {
            section.appendChild(renderHeaderFooter('xlsx-header', sheet.headerFooter.oddHeader));
        }
        if (sheet.headerFooter?.oddFooter) {
            section.appendChild(renderHeaderFooter('xlsx-footer', sheet.headerFooter.oddFooter));
        }
        return section;
    }
    function renderShape(shape) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-shape';
        aside.setAttribute('data-kind', shape.kind);
        if (shape.preset)
            aside.setAttribute('data-preset', shape.preset);
        if (shape.name)
            aside.setAttribute('data-name', shape.name);
        if (shape.alt)
            aside.setAttribute('aria-label', shape.alt);
        aside.setAttribute('data-anchor-col', String(shape.col));
        aside.setAttribute('data-anchor-row', String(shape.row));
        if (shape.endCol !== null)
            aside.setAttribute('data-anchor-end-col', String(shape.endCol));
        if (shape.endRow !== null)
            aside.setAttribute('data-anchor-end-row', String(shape.endRow));
        if (shape.text && shape.text.length > 0) {
            const pre = document.createElement('pre');
            pre.textContent = shape.text;
            aside.appendChild(pre);
        }
        return aside;
    }
    function renderHeaderFooter(className, zones) {
        const div = document.createElement('div');
        div.className = className;
        const addZone = (name, text) => {
            const z = document.createElement('div');
            z.setAttribute('data-zone', name);
            z.textContent = text;
            div.appendChild(z);
        };
        addZone('left', zones.left);
        addZone('center', zones.center);
        addZone('right', zones.right);
        return div;
    }
    function appendCommentMarker(td, comment) {
        const marker = document.createElement('span');
        marker.className = 'xlsx-comment-marker';
        marker.setAttribute('role', 'note');
        const author = comment.author && comment.author.length > 0 ? comment.author : null;
        const title = author ? `${author}: ${comment.text}` : comment.text;
        marker.setAttribute('title', title);
        marker.textContent = '●';
        td.appendChild(marker);
    }
    function appendThreadedCommentMarker(td, entries) {
        const marker = document.createElement('span');
        marker.className = 'xlsx-comment-marker xlsx-threaded';
        marker.setAttribute('role', 'note');
        marker.setAttribute('title', formatThreadTitle(entries));
        marker.textContent = '💬';
        td.appendChild(marker);
    }
    function formatThreadTitle(entries) {
        const byId = new Map();
        for (const e of entries)
            byId.set(e.id, e);
        const sortByDate = (arr) => arr.slice().sort((a, b) => (a.date ?? '').localeCompare(b.date ?? ''));
        const starters = sortByDate(entries.filter((e) => !e.parentId || !byId.has(e.parentId)));
        const repliesByParent = new Map();
        for (const e of entries) {
            if (!e.parentId || !byId.has(e.parentId))
                continue;
            const list = repliesByParent.get(e.parentId);
            if (list)
                list.push(e);
            else
                repliesByParent.set(e.parentId, [e]);
        }
        const ordered = [];
        for (const starter of starters) {
            ordered.push(starter);
            const replies = repliesByParent.get(starter.id);
            if (replies)
                for (const r of sortByDate(replies))
                    ordered.push(r);
        }
        return ordered.map((e) => `${e.author ?? 'Unknown'}: ${e.text}`).join('\n');
    }
    function applySheetView(section, view, theme) {
        if (view.rightToLeft)
            section.setAttribute('dir', 'rtl');
        if (!view.showGridLines)
            section.classList.add('xlsx-no-gridlines');
        if (!view.showRowColHeaders)
            section.classList.add('xlsx-no-headers');
        if (view.zoomScale !== null && view.zoomScale !== 100) {
            section.style.zoom = String(view.zoomScale / 100);
        }
        if (view.tabColor) {
            const hex = resolveColor(view.tabColor, theme);
            if (hex)
                section.setAttribute('data-tab-color', hex);
        }
    }
    function wrapCellWithHyperlink(td, link) {
        let href = null;
        if (link.target != null && link.target !== '' && isSafeHyperlinkHref(link.target)) {
            href = link.target;
        }
        else if (link.location != null && link.location !== '') {
            const frag = `#${link.location}`;
            if (isSafeHyperlinkHref(frag))
                href = frag;
        }
        if (href == null)
            return;
        const a = document.createElement('a');
        a.className = 'xlsx-hyperlink';
        a.setAttribute('href', href);
        a.setAttribute('target', '_blank');
        a.setAttribute('rel', 'noopener noreferrer');
        if (link.tooltip)
            a.setAttribute('title', link.tooltip);
        if (td.firstChild) {
            while (td.firstChild)
                a.appendChild(td.firstChild);
        }
        else if (link.display) {
            a.textContent = link.display;
        }
        td.appendChild(a);
    }
    function tagFrozen(td, row, col, panes) {
        const inX = panes.xSplit !== null && col < panes.xSplit;
        const inY = panes.ySplit !== null && row < panes.ySplit;
        if (inX && inY)
            td.classList.add('xlsx-frozen-both');
        else if (inY)
            td.classList.add('xlsx-frozen-row');
        else if (inX)
            td.classList.add('xlsx-frozen-col');
    }
    function renderCellContent(td, cell, styles, theme, date1904, options) {
        const xf = resolveXf(styles, cell.styleIndex);
        let text = cell.value;
        let numeric = cell.kind === 'number';
        let formatColor = null;
        if ((cell.kind === 'number' || cell.kind === 'empty') && xf) {
            const code = lookupNumberFormat(styles, xf.numFmtId);
            if (code && code !== 'General' && cell.value !== '') {
                const res = formatNumber(cell.value, code, { date1904 });
                text = res.text;
                numeric = res.numeric;
                formatColor = res.color;
            }
        }
        if (options.showFormulas && cell.formula != null) {
            const a1 = `=${cell.formula}`;
            text = options.formulaNotation === 'r1c1' ? a1ToR1c1(a1, cell.row, cell.col) : a1;
            td.textContent = text;
            td.classList.add('xlsx-formula');
            return;
        }
        if (cell.runs && (cell.kind === 'string' || cell.kind === 'inlineStr')) {
            for (const run of cell.runs)
                appendRunSpan(td, run, theme);
        }
        else if (cell.phonetics && (cell.kind === 'string' || cell.kind === 'inlineStr')) {
            appendPhoneticText(td, cell.value, cell.phonetics);
        }
        else {
            td.textContent = text;
        }
        if (cell.kind === 'error')
            td.classList.add('xlsx-error');
        if (numeric)
            td.classList.add('xlsx-numeric');
        if (!xf)
            return;
        if (xf.applyFont || xf.fontId > 0)
            applyFont(td, styles.fonts[xf.fontId], theme);
        if (xf.applyFill || xf.fillId > 0)
            applyFill(td, styles.fills[xf.fillId], theme);
        if (xf.applyBorder || xf.borderId > 0)
            applyBorder(td, styles.borders[xf.borderId], theme);
        if (xf.applyAlignment)
            applyAlignment(td, xf);
        if (formatColor)
            td.style.color = formatColor;
    }
    function appendRunSpan(td, run, theme) {
        const span = document.createElement('span');
        span.textContent = run.text;
        if (run.bold)
            span.style.fontWeight = 'bold';
        if (run.italic)
            span.style.fontStyle = 'italic';
        applyTextDecoration(span, run.underline, run.strike);
        if (run.underline === 'double' || run.underline === 'doubleAccounting') {
            span.style.textDecorationStyle = 'double';
        }
        if (run.underline === 'singleAccounting' || run.underline === 'doubleAccounting') {
            span.classList.add('xlsx-accounting-underline');
        }
        if (run.vertAlign === 'subscript') {
            span.style.verticalAlign = 'sub';
            span.style.fontSize = '0.8em';
        }
        else if (run.vertAlign === 'superscript') {
            span.style.verticalAlign = 'super';
            span.style.fontSize = '0.8em';
        }
        if (run.size)
            span.style.fontSize = `${run.size}pt`;
        const color = resolveColor(run.color, theme);
        if (color)
            span.style.color = color;
        const family = sanitizeFontFamily(run.name);
        if (family)
            span.style.fontFamily = family;
        td.appendChild(span);
    }
    function appendPhoneticText(parent, text, phonetics) {
        const sorted = phonetics
            .map((p) => ({
            startIdx: Math.max(0, Math.min(text.length, p.startIdx)),
            endIdx: Math.max(0, Math.min(text.length, p.endIdx)),
            phonetic: p.phonetic,
        }))
            .filter((p) => p.endIdx > p.startIdx)
            .sort((a, b) => a.startIdx - b.startIdx);
        let cursor = 0;
        for (const p of sorted) {
            if (p.startIdx < cursor)
                continue;
            if (p.startIdx > cursor) {
                parent.appendChild(document.createTextNode(text.slice(cursor, p.startIdx)));
            }
            const ruby = document.createElement('ruby');
            ruby.appendChild(document.createTextNode(text.slice(p.startIdx, p.endIdx)));
            const rt = document.createElement('rt');
            rt.textContent = p.phonetic;
            ruby.appendChild(rt);
            parent.appendChild(ruby);
            cursor = p.endIdx;
        }
        if (cursor < text.length) {
            parent.appendChild(document.createTextNode(text.slice(cursor)));
        }
    }
    function applyTextDecoration(el, underline, strike) {
        const parts = [];
        if (underline)
            parts.push('underline');
        if (strike)
            parts.push('line-through');
        if (parts.length === 0)
            return;
        el.style.textDecoration = parts.join(' ');
    }
    function resolveXf(styles, index) {
        if (!styles || index < 0 || index >= styles.cellXfs.length)
            return null;
        return resolveEffectiveXf(styles, styles.cellXfs[index]);
    }
    function applyFont(td, font, theme) {
        if (!font)
            return;
        if (font.bold)
            td.style.fontWeight = 'bold';
        if (font.italic)
            td.style.fontStyle = 'italic';
        applyTextDecoration(td, font.underline, font.strike);
        if (font.underline === 'double' || font.underline === 'doubleAccounting') {
            td.style.textDecorationStyle = 'double';
        }
        if (font.underline === 'singleAccounting' || font.underline === 'doubleAccounting') {
            td.classList.add('xlsx-accounting-underline');
        }
        if (font.vertAlign === 'subscript') {
            td.style.verticalAlign = 'sub';
            td.style.fontSize = '0.8em';
        }
        else if (font.vertAlign === 'superscript') {
            td.style.verticalAlign = 'super';
            td.style.fontSize = '0.8em';
        }
        if (font.size)
            td.style.fontSize = `${font.size}pt`;
        const color = resolveColor(font.color, theme);
        if (color)
            td.style.color = color;
        const family = sanitizeFontFamily(font.name) ?? resolveSchemeFontFamily(font.scheme, theme);
        if (family)
            td.style.fontFamily = family;
    }
    function resolveSchemeFontFamily(scheme, theme) {
        if (!scheme || !theme)
            return null;
        const name = scheme === 'major' ? theme.majorFont : theme.minorFont;
        return sanitizeFontFamily(name);
    }
    function applyFill(td, fill, theme) {
        const color = resolveColor(fill?.fgColor ?? null, theme);
        if (!color)
            return;
        td.style.backgroundColor = color;
    }
    function applyBorder(td, border, theme) {
        if (!border)
            return;
        const sides = [
            ['left', 'borderLeft'],
            ['right', 'borderRight'],
            ['top', 'borderTop'],
            ['bottom', 'borderBottom'],
        ];
        for (const [side, cssSide] of sides) {
            const s = border[side];
            if (!s.style)
                continue;
            const width = borderWidth(s.style);
            const style = borderStyle(s.style);
            const color = resolveColor(s.color, theme) ?? '#000';
            td.style[cssSide] = `${width} ${style} ${color}`;
        }
    }
    function borderWidth(style) {
        switch (style) {
            case 'thick': return '3px';
            case 'medium':
            case 'mediumDashed':
            case 'mediumDashDot':
            case 'mediumDashDotDot':
                return '2px';
            default:
                return '1px';
        }
    }
    function borderStyle(style) {
        if (style.toLowerCase().includes('dashed') || style === 'dashDot' || style === 'dashDotDot')
            return 'dashed';
        if (style === 'dotted' || style === 'hair')
            return 'dotted';
        if (style === 'double')
            return 'double';
        return 'solid';
    }
    function applyAlignment(td, xf) {
        const a = xf.alignment;
        if (a.horizontal) {
            switch (a.horizontal) {
                case 'centerContinuous':
                    td.style.textAlign = 'center';
                    break;
                case 'distributed':
                    td.style.textAlign = 'justify';
                    td.style.textAlignLast = 'justify';
                    break;
                case 'fill':
                    td.style.textAlign = 'start';
                    break;
                default:
                    td.style.textAlign = a.horizontal;
            }
        }
        if (a.vertical) {
            if (a.vertical === 'middle' || a.vertical === 'justify' || a.vertical === 'distributed') {
                td.style.verticalAlign = 'middle';
            }
            else {
                td.style.verticalAlign = a.vertical;
            }
        }
        if (a.wrapText) {
            td.style.whiteSpace = 'normal';
            td.style.wordBreak = 'break-word';
        }
        if (a.readingOrder === 2)
            td.style.direction = 'rtl';
        else if (a.readingOrder === 1)
            td.style.direction = 'ltr';
        if (a.indent > 0) {
            const em = `${(a.indent * 0.5).toFixed(2)}em`;
            const rtl = a.readingOrder === 2;
            const padRight = (a.horizontal === 'right' && !rtl) || (a.horizontal === 'left' && rtl);
            if (padRight)
                td.style.paddingRight = em;
            else
                td.style.paddingLeft = em;
        }
        if (a.textRotation === 255) {
            td.style.writingMode = 'vertical-lr';
        }
        else if (a.textRotation != null && a.textRotation !== 0) {
            td.style.transform = `rotate(-${a.textRotation}deg)`;
            td.style.transformOrigin = 'center center';
            td.style.display = 'inline-block';
        }
        if (a.shrinkToFit)
            td.classList.add('xlsx-shrink-to-fit');
    }
    function applyGraphicalCf(td, state) {
        if (state.colorScaleBg) {
            td.style.backgroundColor = state.colorScaleBg;
            td.classList.add('xlsx-cf-colorscale');
        }
        if (state.dataBar) {
            const bar = state.dataBar;
            const pct = +(Math.max(0, Math.min(1, bar.fraction)) * 100).toFixed(2);
            const angle = bar.direction === 'rightToLeft' ? '270deg' : '90deg';
            const fillStop = bar.gradient
                ? `${bar.color} 0 ${pct}%, transparent ${pct}% 100%`
                : `${bar.color} 0 ${pct}%, transparent ${pct}% 100%`;
            td.style.background = `linear-gradient(${angle}, ${fillStop})`;
            if (bar.border) {
                const bc = bar.borderColor ?? bar.color;
                td.style.border = `1px solid ${bc}`;
            }
            if (bar.axis) {
                td.setAttribute('data-cf-databar-axis', bar.axis.position);
                td.setAttribute('data-cf-databar-axis-color', bar.axis.color);
                if (bar.axis.position === 'middle') {
                    const existingShadow = td.style.boxShadow;
                    const axisShadow = `inset 50% 0 0 -49% ${bar.axis.color}`;
                    td.style.boxShadow = existingShadow ? `${existingShadow}, ${axisShadow}` : axisShadow;
                }
            }
            td.classList.add('xlsx-cf-databar');
        }
        if (state.icon) {
            const iconSpan = document.createElement('span');
            iconSpan.className = 'xlsx-cf-icon';
            iconSpan.style.display = 'inline-flex';
            iconSpan.style.verticalAlign = 'middle';
            iconSpan.style.marginRight = '0.3em';
            iconSpan.innerHTML = state.icon.svg;
            if (!state.icon.showValue) {
                const hide = document.createElement('span');
                while (td.firstChild)
                    hide.appendChild(td.firstChild);
                hide.style.visibility = 'hidden';
                td.appendChild(hide);
            }
            td.insertBefore(iconSpan, td.firstChild);
            td.classList.add('xlsx-cf-iconset');
        }
    }
    function applyDxf(td, dxf, theme, cell, date1904) {
        if (dxf.font) {
            const f = dxf.font;
            if (f.bold)
                td.style.fontWeight = 'bold';
            if (f.italic)
                td.style.fontStyle = 'italic';
            const strike = !!f.strike;
            const underline = f.underline ?? null;
            if (strike || underline) {
                const existingDecoration = td.style.textDecoration ?? '';
                const existingUnderline = /underline/.test(existingDecoration);
                const existingStrike = /line-through/.test(existingDecoration);
                applyTextDecoration(td, underline || (existingUnderline ? 'single' : null), strike || existingStrike);
                if (underline === 'double' || underline === 'doubleAccounting') {
                    td.style.textDecorationStyle = 'double';
                }
                if (underline === 'singleAccounting' || underline === 'doubleAccounting') {
                    td.classList.add('xlsx-accounting-underline');
                }
            }
            if (f.vertAlign === 'subscript') {
                td.style.verticalAlign = 'sub';
                td.style.fontSize = '0.8em';
            }
            else if (f.vertAlign === 'superscript') {
                td.style.verticalAlign = 'super';
                td.style.fontSize = '0.8em';
            }
            if (f.size)
                td.style.fontSize = `${f.size}pt`;
            if (f.color) {
                const color = resolveColor(f.color, theme);
                if (color)
                    td.style.color = color;
            }
            const family = sanitizeFontFamily(f.name);
            if (family)
                td.style.fontFamily = family;
        }
        if (dxf.fill)
            applyFill(td, dxf.fill, theme);
        if (dxf.border)
            applyBorder(td, dxf.border, theme);
        if (dxf.numFmtCode && cell && (cell.kind === 'number' || cell.kind === 'empty')
            && cell.value !== '' && !cell.runs) {
            const code = dxf.numFmtCode;
            if (code !== 'General') {
                const res = formatNumber(cell.value, code, { date1904 });
                td.textContent = res.text;
                if (res.numeric)
                    td.classList.add('xlsx-numeric');
            }
        }
        td.classList.add('xlsx-cf');
    }

    const defaultOptions = {
        className: 'xlsx',
        inWrapper: true,
        debug: false,
        showFormulas: false,
        formulaNotation: 'a1',
        h,
    };
    function mergeOptions(userOptions) {
        return { ...defaultOptions, ...userOptions };
    }
    async function parseAsync(data, userOptions) {
        const ops = mergeOptions(userOptions);
        return Workbook.load(data, new WorkbookParser(ops));
    }
    async function renderWorkbook(workbook, userOptions) {
        const ops = mergeOptions(userOptions);
        const renderer = new HtmlRenderer();
        if (!workbook.parsed)
            throw new Error('xlsx-preview: workbook is not parsed');
        return await renderer.render(workbook.parsed, ops);
    }
    async function renderAsync(data, bodyContainer, styleContainer, userOptions) {
        const wb = await parseAsync(data, userOptions);
        const nodes = await renderWorkbook(wb, userOptions);
        styleContainer ?? (styleContainer = bodyContainer);
        styleContainer.innerHTML = '';
        bodyContainer.innerHTML = '';
        for (const n of nodes) {
            const c = n.nodeName === 'STYLE' ? styleContainer : bodyContainer;
            c.appendChild(n);
        }
        return wb;
    }

    exports.XlsxEncryptedError = XlsxEncryptedError;
    exports.a1ToR1c1 = a1ToR1c1;
    exports.applyTint = applyTint;
    exports.defaultOptions = defaultOptions;
    exports.emuToPx = emuToPx;
    exports.evaluateRule = evaluateRule;
    exports.formatNumber = formatNumber;
    exports.indexedColor = indexedColor;
    exports.interpolateColorScale = interpolateColorScale;
    exports.isSafeHyperlinkHref = isSafeHyperlinkHref;
    exports.lookupNumberFormat = lookupNumberFormat;
    exports.parseAsync = parseAsync;
    exports.parseColorElement = parseColorElement;
    exports.parseConditionalFormatting = parseConditionalFormatting;
    exports.parseStyles = parseStyles;
    exports.parseTheme = parseTheme;
    exports.parseThreadedComments = parseThreadedComments;
    exports.r1c1ToA1 = r1c1ToA1;
    exports.renderAsync = renderAsync;
    exports.renderWorkbook = renderWorkbook;
    exports.resolveCfvo = resolveCfvo;
    exports.resolveColor = resolveColor;
    exports.resolveEffectiveXf = resolveEffectiveXf;
    exports.sanitizeFontFamily = sanitizeFontFamily;
    exports.sanitizeHexColor = sanitizeHexColor;

}));
//# sourceMappingURL=xlsx-preview.js.map
