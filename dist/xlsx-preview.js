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
        };
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
    function resolveEffectiveXf(styles, cellXf) {
        if (cellXf.xfId < 0 || cellXf.xfId >= styles.cellStyleXfs.length)
            return cellXf;
        const base = styles.cellStyleXfs[cellXf.xfId];
        return {
            numFmtId: cellXf.applyNumberFormat ? cellXf.numFmtId : (base.numFmtId || cellXf.numFmtId),
            fontId: cellXf.applyFont ? cellXf.fontId : (base.fontId || cellXf.fontId),
            fillId: cellXf.applyFill ? cellXf.fillId : (base.fillId || cellXf.fillId),
            borderId: cellXf.applyBorder ? cellXf.borderId : (base.borderId || cellXf.borderId),
            xfId: cellXf.xfId,
            applyNumberFormat: cellXf.applyNumberFormat || base.applyNumberFormat,
            applyFont: cellXf.applyFont || base.applyFont,
            applyFill: cellXf.applyFill || base.applyFill,
            applyBorder: cellXf.applyBorder || base.applyBorder,
            applyAlignment: cellXf.applyAlignment || base.applyAlignment,
            alignment: cellXf.applyAlignment ? cellXf.alignment : base.alignment,
        };
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
        const scheme = doc.getElementsByTagNameNS(NS_DRAW, 'clrScheme').item(0);
        if (!scheme)
            return { colors };
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
        return { colors };
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
        return out;
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
        const rule = { type, priority, dxfId, operator, formulas, text, rank, percent, bottom, stopIfTrue };
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
    function parseDataBar(ruleEl) {
        const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'dataBar').item(0);
        if (!root)
            return { cfvos: [], color: null, minLength: 10, maxLength: 90, showValue: true };
        return {
            cfvos: parseCfvos(root),
            color: parseCfColor(root.getElementsByTagNameNS(NS_MAIN, 'color').item(0)),
            minLength: Number(root.getAttribute('minLength') ?? '10') || 10,
            maxLength: Number(root.getAttribute('maxLength') ?? '90') || 90,
            showValue: root.getAttribute('showValue') !== '0',
        };
    }
    function parseIconSet(ruleEl) {
        const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'iconSet').item(0);
        if (!root)
            return { iconSet: '3TrafficLights1', cfvos: [], showValue: true, reverse: false };
        return {
            iconSet: root.getAttribute('iconSet') ?? '3TrafficLights1',
            cfvos: parseCfvos(root),
            showValue: root.getAttribute('showValue') !== '0',
            reverse: root.getAttribute('reverse') === '1',
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
            case 'unsupported': return false;
        }
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

    const NS = {
        main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
        xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
        c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
        tc: 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
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
            const sheets = [];
            for (let i = 0; i < sheetMeta.length; i++) {
                const { name, rId } = sheetMeta[i];
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
                const { images, charts } = resolveDrawingsForSheet(xmlPath, parts, media);
                const pivots = resolvePivotsForSheet(xmlPath, parts);
                const comments = resolveCommentsForSheet(xmlPath, parts);
                const threadedComments = resolveThreadedCommentsForSheet(xmlPath, parts, persons);
                sheets.push(parseSheet(name, xml, sharedStrings, tables, images, charts, pivots, comments, threadedComments));
            }
            return { sheets, styles, theme, persons };
        }
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
            return { images: [], charts: [] };
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const images = [];
        const charts = [];
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
            if (!drawingRelsXml)
                continue;
            const drawingRels = parseRelationships(drawingRelsXml);
            const drawingDir = drawingPath.replace(/\/[^/]+$/, '');
            const parsed = parseDrawing(drawingXml, drawingRels, drawingDir, parts, media);
            images.push(...parsed.images);
            charts.push(...parsed.charts);
        }
        return { images, charts };
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
            const body = textEl ? parseSi(textEl) : { text: '', runs: null };
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
        const anchors = [
            ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'twoCellAnchor')),
            ...Array.from(doc.getElementsByTagNameNS(NS.xdr, 'oneCellAnchor')),
        ];
        for (const anchor of anchors) {
            const from = anchor.getElementsByTagNameNS(NS.xdr, 'from').item(0);
            const to = anchor.getElementsByTagNameNS(NS.xdr, 'to').item(0);
            const col = anchorCellValue(from, 'col');
            const row = anchorCellValue(from, 'row');
            const endCol = to ? anchorCellValue(to, 'col') : null;
            const endRow = to ? anchorCellValue(to, 'row') : null;
            const colOff = anchorCellValue(from, 'colOff');
            const rowOff = anchorCellValue(from, 'rowOff');
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
                        const ext = anchor.getElementsByTagNameNS(NS.xdr, 'ext').item(0);
                        const widthEmu = ext ? Number(ext.getAttribute('cx')) : null;
                        const heightEmu = ext ? Number(ext.getAttribute('cy')) : null;
                        const cNvPr = pic.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                        const alt = cNvPr?.getAttribute('descr') ?? cNvPr?.getAttribute('title') ?? null;
                        images.push({
                            col: col ?? 0,
                            row: row ?? 0,
                            endCol,
                            endRow,
                            colOff: colOff ?? 0,
                            rowOff: rowOff ?? 0,
                            dataUrl,
                            widthEmu: Number.isFinite(widthEmu) ? widthEmu : null,
                            heightEmu: Number.isFinite(heightEmu) ? heightEmu : null,
                            alt,
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
        }
        return { images, charts };
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
            out.push({
                name: nodes[i].getAttribute('name') ?? `Sheet${i + 1}`,
                rId: rId || null,
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
            }
        }
        const text = runs.map((r) => r.text).join('');
        return { text, runs: sawRun ? runs : null };
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
    function parseSheet(name, xml, sharedStrings, tables = [], images = [], charts = [], pivots = [], comments = [], threadedComments = []) {
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
            if (height !== null || hidden) {
                rowDimensions.push({ row: rowIndex, height, hidden });
                if (rowIndex > maxRow)
                    maxRow = rowIndex;
            }
            const cellEls = rowEl.getElementsByTagNameNS(NS.main, 'c');
            const cells = [];
            for (let j = 0; j < cellEls.length; j++) {
                const cell = parseCell(cellEls[j], rowIndex, sharedStrings);
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
        return {
            name, rows, maxCol, maxRow, merges, columns, rowDimensions,
            conditionalFormatting, frozenPanes, autoFilter, tables, images,
            charts, pivots, extensions, comments, threadedComments,
        };
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
        if (state !== 'frozen' && state !== 'frozenSplit')
            return null;
        const xSplit = Number(pane.getAttribute('xSplit'));
        const ySplit = Number(pane.getAttribute('ySplit'));
        const x = Number.isFinite(xSplit) && xSplit > 0 ? xSplit : null;
        const y = Number.isFinite(ySplit) && ySplit > 0 ? ySplit : null;
        if (x === null && y === null)
            return null;
        return { xSplit: x, ySplit: y };
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
            const hasWidth = widthAttr !== null && (customWidth || !Number.isNaN(Number(widthAttr)));
            if (!hasWidth && !hidden)
                continue;
            if (!Number.isFinite(min) || !Number.isFinite(max))
                continue;
            if (min < 1 || max < min)
                continue;
            const width = hasWidth && Number.isFinite(Number(widthAttr)) ? Number(widthAttr) : null;
            out.push({ min: min - 1, max: max - 1, width, hidden });
        }
        return out;
    }
    function parseCell(c, fallbackRow, sharedStrings) {
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
        const base = { col, row, styleIndex, formula };
        switch (type) {
            case 's': {
                const idx = Number(raw);
                const entry = Number.isFinite(idx) && idx >= 0 && idx < sharedStrings.length
                    ? sharedStrings[idx]
                    : { text: '', runs: null };
                return { ...base, value: entry.text, runs: entry.runs, kind: 'string' };
            }
            case 'inlineStr': {
                const is = c.getElementsByTagNameNS(NS.main, 'is').item(0);
                const entry = is ? parseSi(is) : { text: '', runs: null };
                return { ...base, value: entry.text, runs: entry.runs, kind: 'inlineStr' };
            }
            case 'b':
                return { ...base, value: raw === '1' ? 'TRUE' : 'FALSE', runs: null, kind: 'boolean' };
            case 'e':
                return { ...base, value: raw, runs: null, kind: 'error' };
            case 'str':
                return { ...base, value: raw, runs: null, kind: 'string' };
            case 'n':
            default:
                if (raw === '')
                    return { ...base, value: '', runs: null, kind: 'empty' };
                return { ...base, value: raw, runs: null, kind: 'number' };
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
    function formatNumber(value, formatCode) {
        if (formatCode === '' || formatCode.toLowerCase() === 'general') {
            return formatGeneral(value);
        }
        if (formatCode.trim() === '@') {
            return { text: value, numeric: false };
        }
        const num = Number(value);
        if (!Number.isFinite(num)) {
            return { text: value, numeric: false };
        }
        const sections = splitSections(formatCode);
        let sectionIndex;
        if (num > 0)
            sectionIndex = NUMERIC_SECTION_INDEX.positive;
        else if (num < 0)
            sectionIndex = sections[NUMERIC_SECTION_INDEX.negative] ? NUMERIC_SECTION_INDEX.negative : NUMERIC_SECTION_INDEX.positive;
        else
            sectionIndex = sections[NUMERIC_SECTION_INDEX.zero] ? NUMERIC_SECTION_INDEX.zero : NUMERIC_SECTION_INDEX.positive;
        const section = sections[sectionIndex] ?? formatCode;
        const cleaned = stripSquareBracketModifiers(section);
        let magnitude = Math.abs(num);
        if (num < 0 && sectionIndex === NUMERIC_SECTION_INDEX.positive) {
            return { text: '-' + applyFormat(cleaned, magnitude), numeric: true };
        }
        return { text: applyFormat(cleaned, magnitude), numeric: true };
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
        return code.replace(/\[[^\]]*\]/g, '');
    }
    function applyFormat(code, value) {
        if (code.trim().toLowerCase() === 'general') {
            return formatGeneralNumber(value);
        }
        if (/[yMdhsmAP]/.test(stripQuotedLiterals(code))) {
            if (/[yMdAPhs]/.test(stripQuotedLiterals(code))) {
                return formatDateTime(code, value);
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
        if (!inserted)
            out = numberText + out;
        return out;
    }
    function formatGeneral(value) {
        const n = Number(value);
        if (!Number.isFinite(n))
            return { text: value, numeric: false };
        return { text: formatGeneralNumber(n), numeric: true };
    }
    function formatGeneralNumber(n) {
        return n.toString();
    }
    const EPOCH_MS = Date.UTC(1899, 11, 30);
    const MS_PER_DAY = 86400000;
    function serialToDate(serial) {
        const whole = Math.floor(serial);
        const frac = serial - whole;
        const ms = EPOCH_MS + whole * MS_PER_DAY + Math.round(frac * MS_PER_DAY);
        return new Date(ms);
    }
    function formatDateTime(code, value) {
        const date = serialToDate(value);
        const Y = date.getUTCFullYear();
        const M = date.getUTCMonth() + 1;
        const D = date.getUTCDate();
        const h24 = date.getUTCHours();
        const m = date.getUTCMinutes();
        const s = date.getUTCSeconds();
        const ampm = /AM\/PM|am\/pm/.test(stripQuotedLiterals(code));
        const h12 = ((h24 + 11) % 12) + 1;
        const hour = ampm ? h12 : h24;
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
                const inner = code.slice(i + 1, end);
                i = end + 1;
                if (inner.toLowerCase() === 'h') {
                    out += String(h24);
                    lastSawHour = true;
                }
                else if (inner.toLowerCase() === 'mm') {
                    out += pad2(m);
                    lastSawHour = false;
                }
                else if (inner.toLowerCase() === 'ss') {
                    out += pad2(s);
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
                nodes.push(renderSheet(sheet, workbook.styles, workbook.theme, options));
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
.${className} .xlsx-table-caption { font-size: 0.85em; color: #666; margin: 0.25rem 0 0; }
.${className} .xlsx-chart-placeholder {
    border: 1px dashed #999; padding: 1em; margin: 0.5em 0;
    color: #666; font-size: 0.9em; text-align: center;
}
.${className} .xlsx-comment-marker { color: #c00; margin-left: 4px; cursor: help; }
.${className} .xlsx-threaded { color: #0066cc; margin-left: 4px; cursor: help; }
.${className} .xlsx-shrink-to-fit { font-size: clamp(0.55em, 0.95em, 1em); overflow: hidden; }
    `.trim();
        return style;
    }
    function makeCfContext(sheet) {
        return {
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
    function resolveGraphicalConditionalFormats(sheet, theme) {
        const out = new Map();
        const ctx = makeCfContext(sheet);
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
            const svg = renderIcon(icons.iconSet, idx, icons.reverse);
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
            const key = `${entry.row},${entry.col}`;
            const existing = out.get(key) ?? {};
            existing.dataBar = { color, fraction };
            out.set(key, existing);
        }
    }
    function resolveConditionalFormats(sheet, styles) {
        const out = new Map();
        if (!styles?.dxfs.length)
            return out;
        const blocks = sheet.conditionalFormatting;
        if (!blocks.length)
            return out;
        const ctx = makeCfContext(sheet);
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
    function renderSheet(sheet, styles, theme, options) {
        const section = h('section', { class: options.className, 'data-sheet-name': sheet.name });
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
        for (const cw of sheet.columns) {
            for (let i = cw.min; i <= cw.max; i++) {
                if (cw.width !== null)
                    widthByCol.set(i, cw.width);
                if (cw.hidden)
                    hiddenCols.add(i);
            }
        }
        for (let c = 0; c < colCount; c++) {
            const col = document.createElement('col');
            const w = widthByCol.get(c);
            if (w !== undefined)
                col.style.width = `${charWidthToPx(w)}px`;
            if (hiddenCols.has(c))
                col.style.display = 'none';
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
        const dxfByCell = resolveConditionalFormats(sheet, styles);
        const graphicalByCell = resolveGraphicalConditionalFormats(sheet, theme);
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
        const tbody = document.createElement('tbody');
        const rowCount = sheet.maxRow + 1;
        for (let r = 0; r < rowCount; r++) {
            const tr = document.createElement('tr');
            const dim = rowDim.get(r);
            if (dim?.hidden)
                tr.style.display = 'none';
            if (dim?.height != null)
                tr.style.height = `${(dim.height * 4 / 3).toFixed(2)}px`;
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
                    renderCellContent(td, cell, styles, theme, options);
                const dxf = dxfByCell.get(`${r},${c}`);
                if (dxf)
                    applyDxf(td, dxf, theme);
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
            caption.textContent = `Table "${t.displayName || t.name}" · ${t.columns.length} column(s) · rows ${t.row + 1}-${t.endRow + 1}`;
            section.appendChild(caption);
        }
        for (const img of sheet.images) {
            const fig = document.createElement('figure');
            fig.className = 'xlsx-image';
            fig.setAttribute('data-anchor-col', String(img.col));
            fig.setAttribute('data-anchor-row', String(img.row));
            if (img.endCol !== null)
                fig.setAttribute('data-anchor-end-col', String(img.endCol));
            if (img.endRow !== null)
                fig.setAttribute('data-anchor-end-row', String(img.endRow));
            fig.style.margin = '0.5rem 0';
            const el = document.createElement('img');
            el.src = img.dataUrl;
            if (img.alt)
                el.alt = img.alt;
            if (img.widthEmu && img.heightEmu) {
                el.width = Math.round(img.widthEmu / 9525);
                el.height = Math.round(img.heightEmu / 9525);
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
        return section;
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
    function renderCellContent(td, cell, styles, theme, options) {
        const xf = resolveXf(styles, cell.styleIndex);
        let text = cell.value;
        let numeric = cell.kind === 'number';
        if ((cell.kind === 'number' || cell.kind === 'empty') && xf) {
            const code = lookupNumberFormat(styles, xf.numFmtId);
            if (code && code !== 'General' && cell.value !== '') {
                const res = formatNumber(cell.value, code);
                text = res.text;
                numeric = res.numeric;
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
        const family = sanitizeFontFamily(font.name);
        if (family)
            td.style.fontFamily = family;
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
            const pct = +(Math.max(0, Math.min(1, state.dataBar.fraction)) * 100).toFixed(2);
            td.style.background = `linear-gradient(90deg, ${state.dataBar.color} 0 ${pct}%, transparent ${pct}% 100%)`;
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
    function applyDxf(td, dxf, theme) {
        if (dxf.font) {
            const f = dxf.font;
            if (f.bold)
                td.style.fontWeight = 'bold';
            if (f.italic)
                td.style.fontStyle = 'italic';
            if (f.underline || f.strike) {
                applyTextDecoration(td, f.underline ?? null, !!f.strike);
                if (f.underline === 'double' || f.underline === 'doubleAccounting') {
                    td.style.textDecorationStyle = 'double';
                }
                if (f.underline === 'singleAccounting' || f.underline === 'doubleAccounting') {
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
    exports.evaluateRule = evaluateRule;
    exports.formatNumber = formatNumber;
    exports.indexedColor = indexedColor;
    exports.interpolateColorScale = interpolateColorScale;
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
    exports.sanitizeFontFamily = sanitizeFontFamily;
    exports.sanitizeHexColor = sanitizeHexColor;

}));
//# sourceMappingURL=xlsx-preview.js.map
