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
    const MAX_MEDIA_BYTES = 32 * 1024 * 1024;
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
            this.embeddings = {};
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
                if (/^xl\/ctrlProps\/.*\.xml$/i.test(p)) {
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
                if (/^xl\/diagrams\/.*\.xml$/i.test(p)) {
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
                    const buf = await bin.async('uint8array');
                    if (buf.byteLength > MAX_MEDIA_BYTES)
                        continue;
                    const base64 = bytesToBase64(buf);
                    const mime = sanitizeMediaMime(guessMime(p));
                    wb.media[p] = `data:${mime};base64,${base64}`;
                }
            }
            const contentTypesXml = await readIfPresent('[Content_Types].xml');
            if (contentTypesXml)
                wb.parts['[Content_Types].xml'] = contentTypesXml;
            for (const p of Object.keys(zip.files)) {
                if (/^xl\/embeddings\/[^/]+$/i.test(p)) {
                    const bin = zip.file(p);
                    if (!bin)
                        continue;
                    const buf = await bin.async('uint8array');
                    wb.embeddings[p] = buf;
                }
            }
            wb.parsed = parser.parse(wb.parts, wb.media, wb.embeddings);
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
    const MEDIA_MIME_ALLOWLIST = new Set([
        'image/png',
        'image/jpeg',
        'image/gif',
        'image/svg+xml',
        'image/webp',
        'image/bmp',
        'image/tiff',
        'application/octet-stream',
        'text/plain',
        'application/pdf',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'application/vnd.ms-excel',
        'application/msword',
        'application/vnd.ms-powerpoint',
    ]);
    function sanitizeMediaMime(mime) {
        return MEDIA_MIME_ALLOWLIST.has(mime) ? mime : 'application/octet-stream';
    }
    const MAX_EMBEDDING_BYTES = MAX_MEDIA_BYTES;
    function bytesToDataUrl(bytes, mime) {
        if (bytes.byteLength > MAX_MEDIA_BYTES)
            return null;
        const safe = sanitizeMediaMime(mime);
        return `data:${safe};base64,${bytesToBase64(bytes)}`;
    }
    function bytesToBase64(bytes) {
        const CHUNK = 0x8000;
        let binary = '';
        for (let i = 0; i < bytes.length; i += CHUNK) {
            const slice = bytes.subarray(i, Math.min(i + CHUNK, bytes.length));
            binary += String.fromCharCode.apply(null, slice);
        }
        return btoa(binary);
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
    function parseXml$3(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function parseStyles(xml) {
        const doc = parseXml$3(xml);
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
    const PATTERN_TYPES = new Set([
        'none', 'solid', 'gray125',
        'darkGray', 'mediumGray', 'lightGray',
        'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp',
        'darkGrid', 'darkTrellis',
        'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
        'lightGrid', 'lightTrellis',
    ]);
    function parseFill(el, opts) {
        const gf = el.getElementsByTagNameNS(NS_MAIN$1, 'gradientFill').item(0);
        if (gf)
            return parseGradientFill(gf);
        const pf = el.getElementsByTagNameNS(NS_MAIN$1, 'patternFill').item(0);
        if (!pf)
            return { kind: 'none' };
        const rawType = pf.getAttribute('patternType');
        let patternType = rawType && PATTERN_TYPES.has(rawType)
            ? rawType
            : 'none';
        const fgEl = pf.getElementsByTagNameNS(NS_MAIN$1, 'fgColor').item(0);
        const bgEl = pf.getElementsByTagNameNS(NS_MAIN$1, 'bgColor').item(0);
        let fgColor = parseColorElement(fgEl);
        const bgColor = parseColorElement(bgEl);
        if (!fgColor && opts?.allowBgFallback && bgColor) {
            fgColor = bgColor;
            if (patternType === 'none')
                patternType = 'solid';
        }
        return { kind: 'pattern', patternType, fgColor, bgColor };
    }
    function parseGradientFill(el) {
        const typeAttr = el.getAttribute('type');
        const type = typeAttr === 'path' ? 'path' : 'linear';
        const degAttr = el.getAttribute('degree');
        const degN = degAttr != null ? Number(degAttr) : 0;
        const degree = Number.isFinite(degN) ? ((degN % 360) + 360) % 360 : 0;
        const stops = [];
        const stopEls = el.getElementsByTagNameNS(NS_MAIN$1, 'stop');
        for (let i = 0; i < stopEls.length; i++) {
            const s = stopEls[i];
            const posAttr = s.getAttribute('position');
            const posN = posAttr != null ? Number(posAttr) : NaN;
            if (!Number.isFinite(posN))
                continue;
            const colorEl = s.getElementsByTagNameNS(NS_MAIN$1, 'color').item(0);
            const color = parseColorElement(colorEl);
            stops.push({ position: Math.max(0, Math.min(1, posN)), color });
        }
        return { kind: 'gradient', type, degree, stops };
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
    function parseXml$2(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function parseTheme(xml) {
        const colors = new Array(12).fill(null);
        colors[0] = '#ffffff';
        colors[1] = '#000000';
        colors[2] = '#e7e6e6';
        colors[3] = '#44546a';
        const doc = parseXml$2(xml);
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
    function evalExpression(rule, cell, range, ctx) {
        const raw = rule.formulas[0] ?? '';
        const src = raw.trim().replace(/^=/, '');
        return evalExpr(src, cell);
    }
    function evalExpr(src, cell, range, ctx) {
        const s = src.trim();
        if (!s)
            return false;
        const boolCall = matchCall(s, ['AND', 'OR', 'NOT']);
        if (boolCall) {
            const args = splitArgs(boolCall.inner);
            if (boolCall.name === 'AND') {
                if (args.length === 0)
                    return false;
                for (const a of args)
                    if (!evalExpr(a, cell))
                        return false;
                return true;
            }
            if (boolCall.name === 'OR') {
                if (args.length === 0)
                    return false;
                for (const a of args)
                    if (evalExpr(a, cell))
                        return true;
                return false;
            }
            if (boolCall.name === 'NOT') {
                if (args.length !== 1)
                    return false;
                return !evalExpr(args[0], cell);
            }
        }
        const modMatch = /^MOD\s*\(\s*(ROW|COLUMN)\s*\(\s*\)\s*,\s*(-?\d+)\s*\)\s*(=|<>)\s*(-?\d+)\s*$/i.exec(s);
        if (modMatch) {
            const axis = modMatch[1].toUpperCase();
            const divisor = Number(modMatch[2]);
            const op = modMatch[3];
            const target = Number(modMatch[4]);
            if (divisor === 0)
                return false;
            const raw = (axis === 'ROW' ? cell.row : cell.col) + 1;
            const m = ((raw % divisor) + divisor) % divisor;
            return op === '=' ? m === target : m !== target;
        }
        const parityMatch = /^(ISEVEN|ISODD)\s*\(\s*(ROW|COLUMN)\s*\(\s*\)\s*\)\s*$/i.exec(s);
        if (parityMatch) {
            const fn = parityMatch[1].toUpperCase();
            const axis = parityMatch[2].toUpperCase();
            const raw = (axis === 'ROW' ? cell.row : cell.col) + 1;
            return fn === 'ISEVEN' ? raw % 2 === 0 : raw % 2 !== 0;
        }
        const isnumSearch = /^ISNUMBER\s*\(\s*(SEARCH\s*\(.*\))\s*\)\s*$/i.exec(s);
        if (isnumSearch)
            return evalSearch(isnumSearch[1], cell);
        const typePred = matchCall(s, ['ISNUMBER', 'ISBLANK', 'ISERROR', 'ISEVEN', 'ISODD']);
        if (typePred) {
            const args = splitArgs(typePred.inner);
            if (args.length !== 1)
                return false;
            const arg = args[0].trim();
            if (typePred.name === 'ISEVEN' || typePred.name === 'ISODD') {
                if (!isCellRef(arg))
                    return false;
                const n = numericValue(cell);
                if (n === null || !Number.isFinite(n))
                    return false;
                const even = Math.trunc(n) % 2 === 0;
                return typePred.name === 'ISEVEN' ? even : !even;
            }
            if (!isCellRef(arg))
                return false;
            if (typePred.name === 'ISNUMBER')
                return cell.kind === 'number' || cell.kind === 'boolean';
            if (typePred.name === 'ISBLANK')
                return evalContainsBlanks(cell);
            if (typePred.name === 'ISERROR')
                return cell.kind === 'error';
        }
        const searchCmp = /^(SEARCH\s*\(.*\))\s*(>=|>|<>|=)\s*(-?\d+)\s*$/i.exec(s);
        if (searchCmp) {
            const hit = evalSearch(searchCmp[1], cell);
            const target = Number(searchCmp[3]);
            switch (searchCmp[2]) {
                case '>':
                case '>=':
                    return hit && target >= 0;
                case '=':
                    return hit && target >= 1;
                case '<>':
                    return !hit;
                default: return false;
            }
        }
        if (/^SEARCH\s*\(.*\)\s*$/i.test(s))
            return evalSearch(s, cell);
        const leftRight = /^(LEFT|RIGHT)\s*\(\s*([^,]+?)\s*(?:,\s*(\d+)\s*)?\)\s*(=|<>)\s*(".*")\s*$/i.exec(s);
        if (leftRight) {
            const fn = leftRight[1].toUpperCase();
            const argRef = leftRight[2].trim();
            const n = leftRight[3] ? Number(leftRight[3]) : 1;
            const op = leftRight[4];
            const lit = parseQuoted(leftRight[5]);
            if (!isCellRef(argRef) || lit === null || !Number.isFinite(n) || n < 0)
                return false;
            const v = cell.value ?? '';
            const slice = fn === 'LEFT' ? v.slice(0, n) : (n === 0 ? '' : v.slice(-n));
            const hit = slice === lit;
            return op === '=' ? hit : !hit;
        }
        const m = /^\$?([A-Z]+)\$?([1-9][0-9]*)\s*(<=|>=|<>|=|<|>)\s*(.+)$/.exec(s);
        if (m) {
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
        return false;
    }
    function matchCall(src, names) {
        const s = src.trim();
        for (const n of names) {
            if (s.length < n.length + 2)
                continue;
            if (s.slice(0, n.length).toUpperCase() !== n)
                continue;
            let i = n.length;
            while (i < s.length && (s[i] === ' ' || s[i] === '\t'))
                i++;
            if (s[i] !== '(')
                continue;
            let depth = 0;
            let inStr = false;
            let end = -1;
            for (let j = i; j < s.length; j++) {
                const ch = s[j];
                if (inStr) {
                    if (ch === '"') {
                        if (s[j + 1] === '"') {
                            j++;
                            continue;
                        }
                        inStr = false;
                    }
                    continue;
                }
                if (ch === '"') {
                    inStr = true;
                    continue;
                }
                if (ch === '(')
                    depth++;
                else if (ch === ')') {
                    depth--;
                    if (depth === 0) {
                        end = j;
                        break;
                    }
                }
            }
            if (end !== s.length - 1)
                continue;
            return { name: n, inner: s.slice(i + 1, end) };
        }
        return null;
    }
    function splitArgs(s) {
        const out = [];
        let depth = 0;
        let inStr = false;
        let start = 0;
        for (let i = 0; i < s.length; i++) {
            const ch = s[i];
            if (inStr) {
                if (ch === '"') {
                    if (s[i + 1] === '"') {
                        i++;
                        continue;
                    }
                    inStr = false;
                }
                continue;
            }
            if (ch === '"') {
                inStr = true;
                continue;
            }
            if (ch === '(')
                depth++;
            else if (ch === ')')
                depth--;
            else if (ch === ',' && depth === 0) {
                out.push(s.slice(start, i).trim());
                start = i + 1;
            }
        }
        const tail = s.slice(start).trim();
        if (tail || out.length > 0)
            out.push(tail);
        return out;
    }
    function parseQuoted(s) {
        const t = s.trim();
        if (t.length < 2 || t[0] !== '"' || t[t.length - 1] !== '"')
            return null;
        let out = '';
        for (let i = 1; i < t.length - 1; i++) {
            if (t[i] === '"') {
                if (t[i + 1] === '"') {
                    out += '"';
                    i++;
                    continue;
                }
                return null;
            }
            out += t[i];
        }
        return out;
    }
    function isCellRef(s) {
        return /^\$?[A-Z]+\$?[1-9][0-9]*$/i.test(s.trim());
    }
    function evalSearch(callSrc, cell) {
        const call = matchCall(callSrc, ['SEARCH']);
        if (!call)
            return false;
        const args = splitArgs(call.inner);
        if (args.length < 2)
            return false;
        const needle = parseQuoted(args[0]);
        const targetArg = args[1].trim();
        if (needle === null)
            return false;
        if (!isCellRef(targetArg))
            return false;
        const hay = (cell.value ?? '').toLocaleLowerCase();
        return hay.includes(needle.toLocaleLowerCase());
    }

    const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    function parseXml$1(xml) {
        return new DOMParser().parseFromString(xml, 'application/xml');
    }
    function firstChild(parent, ns, localName) {
        if (!parent)
            return null;
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
    function directChildren(parent, ns, localName) {
        const out = [];
        for (let i = 0; i < parent.childNodes.length; i++) {
            const node = parent.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI === ns && el.localName === localName)
                out.push(el);
        }
        return out;
    }
    function readPts(parent) {
        const pts = directChildren(parent, NS_C, 'pt');
        if (!pts.length)
            return [];
        let maxIdx = -1;
        const raw = [];
        for (const pt of pts) {
            const idxAttr = pt.getAttribute('idx');
            const idx = idxAttr !== null ? Number(idxAttr) : raw.length;
            const vEl = firstChild(pt, NS_C, 'v');
            const v = vEl ? (vEl.textContent ?? '') : '';
            raw.push({ idx: Number.isFinite(idx) ? idx : raw.length, v });
            if (idx > maxIdx)
                maxIdx = idx;
        }
        const out = [];
        for (let i = 0; i <= maxIdx; i++)
            out.push('');
        for (const r of raw)
            out[r.idx] = r.v;
        return out;
    }
    function parseCategories(ser) {
        const cat = firstChild(ser, NS_C, 'cat');
        if (!cat)
            return [];
        const strRef = firstChild(cat, NS_C, 'strRef');
        if (strRef) {
            const cache = firstChild(strRef, NS_C, 'strCache');
            if (cache)
                return readPts(cache);
        }
        const numRef = firstChild(cat, NS_C, 'numRef');
        if (numRef) {
            const cache = firstChild(numRef, NS_C, 'numCache');
            if (cache)
                return readPts(cache);
        }
        const strLit = firstChild(cat, NS_C, 'strLit');
        if (strLit)
            return readPts(strLit);
        const numLit = firstChild(cat, NS_C, 'numLit');
        if (numLit)
            return readPts(numLit);
        return [];
    }
    function parseValues(ser) {
        return readNumericRefOrLit(ser, 'val');
    }
    function parseXValues(ser) {
        const xVal = firstChild(ser, NS_C, 'xVal');
        if (!xVal)
            return null;
        return readNumericCacheOrLit(xVal);
    }
    function readNumericRefOrLit(ser, localName) {
        const el = firstChild(ser, NS_C, localName);
        if (!el)
            return [];
        return readNumericCacheOrLit(el);
    }
    function readNumericCacheOrLit(container) {
        const numRef = firstChild(container, NS_C, 'numRef');
        const cache = numRef ? firstChild(numRef, NS_C, 'numCache') : null;
        const source = cache ?? firstChild(container, NS_C, 'numLit');
        if (!source)
            return [];
        const raw = readPts(source);
        return raw.map((s) => {
            if (s === '')
                return null;
            const n = Number(s);
            return Number.isFinite(n) ? n : null;
        });
    }
    function parseSeriesName(ser) {
        const tx = firstChild(ser, NS_C, 'tx');
        if (!tx)
            return null;
        const strRef = firstChild(tx, NS_C, 'strRef');
        if (strRef) {
            const cache = firstChild(strRef, NS_C, 'strCache');
            if (cache) {
                const pts = readPts(cache);
                if (pts.length)
                    return pts[0];
            }
        }
        const vEl = firstChild(tx, NS_C, 'v');
        if (vEl)
            return vEl.textContent ?? null;
        const rich = firstChild(tx, NS_C, 'rich');
        if (rich)
            return collectRichText(rich);
        return null;
    }
    function collectRichText(root) {
        const parts = [];
        const visit = (el) => {
            for (let i = 0; i < el.childNodes.length; i++) {
                const node = el.childNodes[i];
                if (node.nodeType !== 1)
                    continue;
                const child = node;
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
    function parseSeriesColor(ser) {
        const spPr = firstChild(ser, NS_C, 'spPr');
        if (!spPr)
            return null;
        const solidFill = firstChild(spPr, NS_A, 'solidFill');
        if (!solidFill)
            return null;
        const srgb = firstChild(solidFill, NS_A, 'srgbClr');
        if (!srgb)
            return null;
        const val = srgb.getAttribute('val');
        if (!val || !/^[0-9a-fA-F]{6}$/.test(val))
            return null;
        return `#${val.toLowerCase()}`;
    }
    function parseSer(ser, fallbackCategories, kind, chartLevelLabels) {
        let values;
        let xValues;
        if (kind === 'scatter') {
            const yVal = readNumericRefOrLit(ser, 'yVal');
            values = yVal.length ? yVal : parseValues(ser);
            xValues = parseXValues(ser);
        }
        else {
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
    }
    function parseDataLabelsBlock(block) {
        if (!block)
            return { show: false, position: null };
        let show = false;
        let position = null;
        for (let i = 0; i < block.childNodes.length; i++) {
            const node = block.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI !== NS_C)
                continue;
            if (el.localName === 'showVal') {
                if (el.getAttribute('val') === '1')
                    show = true;
            }
            else if (el.localName === 'dLblPos') {
                const v = el.getAttribute('val');
                if (v)
                    position = v;
            }
            else if (el.localName === 'dLbl') {
                for (let j = 0; j < el.childNodes.length; j++) {
                    const c = el.childNodes[j];
                    if (c.nodeType !== 1)
                        continue;
                    const ce = c;
                    if (ce.namespaceURI !== NS_C)
                        continue;
                    if (ce.localName === 'showVal' && ce.getAttribute('val') === '1')
                        show = true;
                    if (ce.localName === 'dLblPos') {
                        const v = ce.getAttribute('val');
                        if (v && position === null)
                            position = v;
                    }
                }
            }
        }
        return { show, position };
    }
    function parseSeriesDataLabels(ser, chartLevel) {
        const block = firstChild(ser, NS_C, 'dLbls');
        if (!block)
            return { show: chartLevel.show, position: chartLevel.position };
        const perSer = parseDataLabelsBlock(block);
        return {
            show: perSer.show || chartLevel.show,
            position: perSer.position ?? chartLevel.position,
        };
    }
    function parseGrouping(chartEl) {
        const g = firstChild(chartEl, NS_C, 'grouping');
        const val = g?.getAttribute('val');
        if (val === 'stacked')
            return 'stacked';
        if (val === 'percentStacked')
            return 'percentStacked';
        if (val === 'standard' || val === 'clustered')
            return 'standard';
        return null;
    }
    function parseTitle(chart) {
        const title = firstChild(chart, NS_C, 'title');
        if (!title)
            return null;
        const tx = firstChild(title, NS_C, 'tx');
        if (!tx)
            return null;
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
                if (pts.length && pts[0])
                    return pts[0];
            }
        }
        return null;
    }
    function parseLegend(chart) {
        const legend = firstChild(chart, NS_C, 'legend');
        if (!legend)
            return 'none';
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
    function dispatchChart(plotArea) {
        for (let i = 0; i < plotArea.childNodes.length; i++) {
            const node = plotArea.childNodes[i];
            if (node.nodeType !== 1)
                continue;
            const el = node;
            if (el.namespaceURI !== NS_C)
                continue;
            switch (el.localName) {
                case 'barChart': {
                    const barDir = firstChild(el, NS_C, 'barDir');
                    const dir = barDir?.getAttribute('val');
                    const kind = dir === 'bar' ? 'bar' : 'column';
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
    function parseChart(xml, sharedStrings) {
        const blank = {
            kind: 'unknown',
            title: null,
            categories: [],
            series: [],
            legend: 'none',
            grouping: null,
        };
        let doc;
        try {
            doc = parseXml$1(xml);
        }
        catch {
            return blank;
        }
        const chartSpace = doc.getElementsByTagNameNS(NS_C, 'chartSpace').item(0);
        if (!chartSpace)
            return blank;
        const chart = firstChild(chartSpace, NS_C, 'chart');
        if (!chart)
            return blank;
        const plotArea = firstChild(chart, NS_C, 'plotArea');
        if (!plotArea)
            return blank;
        const { kind, sers, plotEl } = dispatchChart(plotArea);
        if (kind === 'unknown') {
            return {
                ...blank,
                title: parseTitle(chart),
                legend: parseLegend(chart),
            };
        }
        let categories = [];
        if (kind !== 'scatter') {
            for (const ser of sers) {
                const cat = parseCategories(ser);
                if (cat.length) {
                    categories = cat;
                    break;
                }
            }
        }
        const chartLevelLabels = plotEl
            ? parseDataLabelsBlock(firstChild(plotEl, NS_C, 'dLbls'))
            : { show: false, position: null };
        const series = sers.map((ser) => parseSer(ser, categories, kind, chartLevelLabels));
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
        };
    }

    const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const MAX_TREE_DEPTH = 32;
    function parseSmartArt(dataXml, layoutXml) {
        const empty = { id: '', rootNodes: [], layout: parseLayoutName(layoutXml ?? null) };
        if (!dataXml)
            return empty;
        let doc;
        try {
            doc = new DOMParser().parseFromString(dataXml, 'application/xml');
        }
        catch {
            return empty;
        }
        const dataModel = doc.getElementsByTagNameNS(DGM_NS, 'dataModel').item(0);
        if (!dataModel)
            return empty;
        const ptLst = dataModel.getElementsByTagNameNS(DGM_NS, 'ptLst').item(0);
        if (!ptLst)
            return empty;
        const points = new Map();
        const ptEls = ptLst.getElementsByTagNameNS(DGM_NS, 'pt');
        for (let i = 0; i < ptEls.length; i++) {
            const pt = ptEls[i];
            const id = pt.getAttribute('modelId');
            if (!id)
                continue;
            const type = pt.getAttribute('type') ?? '';
            if (type === 'pres' || type === 'parTrans' || type === 'sibTrans')
                continue;
            const text = flattenPtText(pt);
            points.set(id, { id, type, text });
        }
        const parOfs = [];
        const cxnLst = dataModel.getElementsByTagNameNS(DGM_NS, 'cxnLst').item(0);
        if (cxnLst) {
            const cxnEls = cxnLst.getElementsByTagNameNS(DGM_NS, 'cxn');
            for (let i = 0; i < cxnEls.length; i++) {
                const cxn = cxnEls[i];
                const type = cxn.getAttribute('type') ?? 'parOf';
                if (type !== 'parOf')
                    continue;
                const src = cxn.getAttribute('srcId');
                const dest = cxn.getAttribute('destId');
                if (!src || !dest)
                    continue;
                if (!points.has(src) || !points.has(dest))
                    continue;
                const ordAttr = cxn.getAttribute('srcOrd');
                const ord = ordAttr !== null && Number.isFinite(Number(ordAttr))
                    ? Number(ordAttr)
                    : 0;
                parOfs.push({ src, dest, ord });
            }
        }
        const childrenOf = new Map();
        const referencedAsDest = new Set();
        const grouped = new Map();
        for (const edge of parOfs) {
            const list = grouped.get(edge.src) ?? [];
            list.push(edge);
            grouped.set(edge.src, list);
            referencedAsDest.add(edge.dest);
        }
        for (const [src, list] of grouped) {
            list.sort((a, b) => a.ord - b.ord);
            childrenOf.set(src, list.map((e) => e.dest));
        }
        const docPt = Array.from(points.values()).find((p) => p.type === 'doc');
        let rootIds;
        if (docPt && childrenOf.has(docPt.id)) {
            rootIds = childrenOf.get(docPt.id) ?? [];
        }
        else {
            rootIds = Array.from(points.values())
                .filter((p) => p.type !== 'doc' && !referencedAsDest.has(p.id))
                .map((p) => p.id);
        }
        const rootNodes = [];
        const onStack = new Set();
        for (const rootId of rootIds) {
            const root = buildNode(rootId, points, childrenOf, onStack, 0);
            if (root)
                rootNodes.push(root);
        }
        return { id: '', rootNodes, layout: parseLayoutName(layoutXml ?? null) };
    }
    function buildNode(id, points, childrenOf, onStack, depth) {
        const pt = points.get(id);
        if (!pt)
            return null;
        if (depth >= MAX_TREE_DEPTH) {
            return { id: pt.id, text: pt.text, children: [], level: depth };
        }
        if (onStack.has(id)) {
            return { id: pt.id, text: pt.text, children: [], level: depth };
        }
        onStack.add(id);
        const children = [];
        const childIds = childrenOf.get(id) ?? [];
        for (const childId of childIds) {
            const child = buildNode(childId, points, childrenOf, onStack, depth + 1);
            if (child)
                children.push(child);
        }
        onStack.delete(id);
        return { id: pt.id, text: pt.text, children, level: depth };
    }
    function flattenPtText(pt) {
        const t = pt.getElementsByTagNameNS(DGM_NS, 't').item(0);
        if (!t)
            return '';
        const pEls = t.getElementsByTagNameNS(A_NS, 'p');
        if (pEls.length === 0) {
            const tEls = t.getElementsByTagNameNS(A_NS, 't');
            let out = '';
            for (let i = 0; i < tEls.length; i++)
                out += tEls[i].textContent ?? '';
            return out;
        }
        const paragraphs = [];
        for (let i = 0; i < pEls.length; i++) {
            const p = pEls[i];
            const tEls = p.getElementsByTagNameNS(A_NS, 't');
            let line = '';
            for (let j = 0; j < tEls.length; j++)
                line += tEls[j].textContent ?? '';
            paragraphs.push(line);
        }
        return paragraphs.join('\n');
    }
    function parseLayoutName(layoutXml) {
        if (!layoutXml)
            return null;
        let doc;
        try {
            doc = new DOMParser().parseFromString(layoutXml, 'application/xml');
        }
        catch {
            return null;
        }
        const def = doc.getElementsByTagNameNS(DGM_NS, 'layoutDef').item(0);
        if (!def)
            return null;
        return def.getAttribute('uniqueId') || null;
    }

    const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    const FREESECT = 0xffffffff;
    const ENDOFCHAIN = 0xfffffffe;
    const FATSECT = 0xfffffffd;
    const DIFSECT = 0xfffffffc;
    const DIR_TYPE_UNUSED = 0x00;
    const DIR_TYPE_STREAM = 0x02;
    const DIR_TYPE_ROOT = 0x05;
    const MAX_STREAMS = 256;
    const MAX_STREAM_BYTES = 16 * 1024 * 1024;
    const MAX_BUFFER_BYTES = 32 * 1024 * 1024;
    const MAX_CHAIN_SECTORS = 1 << 20;
    function parseCfb(bytes) {
        if (!bytes || bytes.byteLength < 512)
            return null;
        if (bytes.byteLength > MAX_BUFFER_BYTES)
            return null;
        for (let i = 0; i < CFB_MAGIC.length; i++) {
            if (bytes[i] !== CFB_MAGIC[i])
                return null;
        }
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const sectorShift = view.getUint16(0x1e, true);
        if (sectorShift !== 9 && sectorShift !== 12)
            return null;
        const sectorSize = 1 << sectorShift;
        if (sectorSize < 128)
            return null;
        const miniSectorShift = view.getUint16(0x20, true);
        if (miniSectorShift < 1 || miniSectorShift > 16)
            return null;
        const miniSectorSize = 1 << miniSectorShift;
        const numFatSectors = view.getUint32(0x2c, true);
        const firstDirSector = view.getUint32(0x30, true);
        const miniStreamCutoff = view.getUint32(0x38, true);
        const firstMiniFatSector = view.getUint32(0x3c, true);
        const numMiniFatSectors = view.getUint32(0x40, true);
        const firstDifatSector = view.getUint32(0x44, true);
        const numDifatSectors = view.getUint32(0x48, true);
        if (numFatSectors > MAX_CHAIN_SECTORS)
            return null;
        if (numMiniFatSectors > MAX_CHAIN_SECTORS)
            return null;
        if (numDifatSectors > MAX_CHAIN_SECTORS)
            return null;
        const sectorDataStart = sectorSize;
        const availableBytes = bytes.byteLength - sectorDataStart;
        if (availableBytes < 0)
            return null;
        const totalSectors = Math.floor(availableBytes / sectorSize);
        if (totalSectors <= 0)
            return null;
        const sectorBytes = (sid) => {
            if (sid >= totalSectors)
                return null;
            const off = sectorDataStart + sid * sectorSize;
            if (off + sectorSize > bytes.byteLength)
                return null;
            return bytes.subarray(off, off + sectorSize);
        };
        const difat = [];
        for (let i = 0; i < 109; i++) {
            const sid = view.getUint32(0x4c + i * 4, true);
            if (sid === FREESECT)
                break;
            difat.push(sid);
        }
        {
            let next = firstDifatSector;
            let guard = 0;
            const seen = new Set();
            while (next !== ENDOFCHAIN && next !== FREESECT && guard < numDifatSectors + 4) {
                if (seen.has(next))
                    break;
                seen.add(next);
                const sec = sectorBytes(next);
                if (!sec)
                    break;
                const secView = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
                const entriesPerDifat = (sectorSize / 4) - 1;
                for (let i = 0; i < entriesPerDifat; i++) {
                    const sid = secView.getUint32(i * 4, true);
                    if (sid === FREESECT)
                        break;
                    difat.push(sid);
                }
                next = secView.getUint32(sectorSize - 4, true);
                guard++;
                if (difat.length > MAX_CHAIN_SECTORS)
                    return null;
            }
        }
        const fatEntriesPerSector = sectorSize / 4;
        const fatLength = difat.length * fatEntriesPerSector;
        if (fatLength > MAX_CHAIN_SECTORS)
            return null;
        const fat = new Uint32Array(fatLength);
        for (let i = 0; i < difat.length; i++) {
            const sec = sectorBytes(difat[i]);
            if (!sec)
                return null;
            const secView = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
            for (let j = 0; j < fatEntriesPerSector; j++) {
                fat[i * fatEntriesPerSector + j] = secView.getUint32(j * 4, true);
            }
        }
        function readFatChain(firstSid, size) {
            if (firstSid === ENDOFCHAIN || firstSid === FREESECT)
                return new Uint8Array(0);
            const cap = Math.min(size, MAX_STREAM_BYTES);
            const out = new Uint8Array(cap);
            let written = 0;
            let sid = firstSid;
            const seen = new Set();
            let hops = 0;
            while (sid !== ENDOFCHAIN && sid !== FREESECT && sid !== FATSECT && sid !== DIFSECT) {
                if (hops++ > MAX_CHAIN_SECTORS)
                    return null;
                if (seen.has(sid))
                    return null;
                seen.add(sid);
                const sec = sectorBytes(sid);
                if (!sec)
                    return null;
                const copyLen = Math.min(sec.length, cap - written);
                if (copyLen <= 0)
                    break;
                out.set(sec.subarray(0, copyLen), written);
                written += copyLen;
                if (written >= cap)
                    break;
                if (sid >= fat.length)
                    return null;
                sid = fat[sid];
            }
            return out.subarray(0, Math.min(written, size));
        }
        const dirBytes = readFatChain(firstDirSector, bytes.byteLength);
        if (!dirBytes)
            return null;
        const entryCount = Math.min(Math.floor(dirBytes.length / 128), MAX_STREAMS);
        if (entryCount === 0)
            return null;
        const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
        const entries = [];
        for (let i = 0; i < entryCount; i++) {
            const base = i * 128;
            if (base + 128 > dirBytes.length)
                break;
            const type = dirBytes[base + 0x42];
            if (type === DIR_TYPE_UNUSED) {
                entries.push({ name: '', type, startSector: 0, size: 0, clsid: null });
                continue;
            }
            const nameLenBytes = dirView.getUint16(base + 0x40, true);
            const clampedNameLen = Math.max(0, Math.min(nameLenBytes, 64));
            const name = decodeUtf16LeName(dirBytes.subarray(base, base + clampedNameLen));
            const startSector = dirView.getUint32(base + 0x74, true);
            const sizeLo = dirView.getUint32(base + 0x78, true);
            const sizeHi = dirView.getUint32(base + 0x7c, true);
            const size = sizeHi !== 0 ? MAX_STREAM_BYTES + 1 : sizeLo;
            const clsid = readClsid(dirBytes, base + 0x50);
            entries.push({ name, type, startSector, size, clsid });
        }
        const root = entries[0];
        if (!root || root.type !== DIR_TYPE_ROOT)
            return null;
        let miniStream = null;
        let miniFat = null;
        const ensureMini = () => {
            if (miniStream && miniFat)
                return true;
            miniStream = readFatChain(root.startSector, Math.min(root.size, MAX_STREAM_BYTES));
            if (!miniStream)
                return false;
            if (firstMiniFatSector === ENDOFCHAIN || firstMiniFatSector === FREESECT) {
                miniFat = new Uint32Array(0);
                return true;
            }
            const mfatBytes = readFatChain(firstMiniFatSector, numMiniFatSectors * sectorSize);
            if (!mfatBytes)
                return false;
            const entries = Math.floor(mfatBytes.length / 4);
            miniFat = new Uint32Array(entries);
            const mfatView = new DataView(mfatBytes.buffer, mfatBytes.byteOffset, mfatBytes.byteLength);
            for (let i = 0; i < entries; i++)
                miniFat[i] = mfatView.getUint32(i * 4, true);
            return true;
        };
        function readMiniChain(firstSid, size) {
            if (!ensureMini() || !miniStream || !miniFat)
                return null;
            if (firstSid === ENDOFCHAIN || firstSid === FREESECT)
                return new Uint8Array(0);
            const cap = Math.min(size, MAX_STREAM_BYTES);
            const out = new Uint8Array(cap);
            let written = 0;
            let sid = firstSid;
            const seen = new Set();
            let hops = 0;
            while (sid !== ENDOFCHAIN && sid !== FREESECT && sid !== FATSECT && sid !== DIFSECT) {
                if (hops++ > MAX_CHAIN_SECTORS)
                    return null;
                if (seen.has(sid))
                    return null;
                seen.add(sid);
                const off = sid * miniSectorSize;
                if (off + miniSectorSize > miniStream.length)
                    return null;
                const copyLen = Math.min(miniSectorSize, cap - written);
                if (copyLen <= 0)
                    break;
                out.set(miniStream.subarray(off, off + copyLen), written);
                written += copyLen;
                if (written >= cap)
                    break;
                if (sid >= miniFat.length)
                    return null;
                sid = miniFat[sid];
            }
            return out.subarray(0, Math.min(written, size));
        }
        const streams = [];
        for (let i = 1; i < entries.length && streams.length < MAX_STREAMS; i++) {
            const e = entries[i];
            if (e.type !== DIR_TYPE_STREAM)
                continue;
            if (e.size > MAX_STREAM_BYTES)
                continue;
            if (!e.name)
                continue;
            const useMini = e.size > 0 && e.size < miniStreamCutoff;
            const streamBytes = useMini
                ? readMiniChain(e.startSector, e.size)
                : readFatChain(e.startSector, e.size);
            if (!streamBytes)
                continue;
            streams.push({ name: e.name, bytes: streamBytes });
        }
        return {
            clsid: root.clsid,
            streams,
        };
    }
    function decodeUtf16LeName(buf) {
        let out = '';
        for (let i = 0; i + 1 < buf.length; i += 2) {
            const cu = buf[i] | (buf[i + 1] << 8);
            if (cu === 0)
                break;
            out += String.fromCharCode(cu);
        }
        if (out.length === 0)
            return '';
        const head = out[0];
        const rest = out.slice(1).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
        if (head === '\x01' || head === '\x05') {
            return head + rest;
        }
        if (head >= '\x00' && head < ' ' && head !== '\t' && head !== '\n' && head !== '\r') {
            return rest;
        }
        return head + rest;
    }
    function readClsid(buf, offset) {
        if (offset + 16 > buf.length)
            return null;
        const b = buf.subarray(offset, offset + 16);
        let allZero = true;
        for (let i = 0; i < 16; i++) {
            if (b[i] !== 0) {
                allZero = false;
                break;
            }
        }
        if (allZero)
            return null;
        const hex = (n) => n.toString(16).padStart(2, '0');
        const g1 = hex(b[3]) + hex(b[2]) + hex(b[1]) + hex(b[0]);
        const g2 = hex(b[5]) + hex(b[4]);
        const g3 = hex(b[7]) + hex(b[6]);
        const g4 = hex(b[8]) + hex(b[9]);
        const g5 = hex(b[10]) + hex(b[11]) + hex(b[12]) + hex(b[13]) + hex(b[14]) + hex(b[15]);
        return `{${g1}-${g2}-${g3}-${g4}-${g5}}`.toUpperCase();
    }

    const SAFE_HREF_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:']);
    const MAX_RANGE_EXPANSION_CELLS = 1048576;
    function isSafeHyperlinkHref(raw) {
        if (raw == null)
            return true;
        if (typeof raw !== 'string')
            return false;
        if (/[\x00-\x1f\x7f]/.test(raw))
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
        dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    };
    const SMARTART_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
    class WorkbookParser {
        constructor(_options) {
            this._options = _options;
        }
        parse(parts, media = {}, embeddingBytes = {}) {
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
            const contentTypes = parts['[Content_Types].xml']
                ? parseContentTypes(parts['[Content_Types].xml'])
                : { defaults: new Map(), overrides: new Map() };
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
                const { images, charts, shapes, formControls, smartArt } = resolveDrawingsForSheet(xmlPath, parts, media, sharedStrings);
                const pivots = resolvePivotsForSheet(xmlPath, parts);
                const { slicers, timelines } = resolveSlicersAndTimelinesForSheet(xmlPath, parts);
                const comments = resolveCommentsForSheet(xmlPath, parts);
                const threadedComments = resolveThreadedCommentsForSheet(xmlPath, parts, persons);
                const hyperlinkTargets = resolveHyperlinkTargets(xmlPath, parts);
                const embeddings = resolveEmbeddingsForSheet(xmlPath, xml, parts, embeddingBytes, contentTypes, this._options.inlineEmbeddings === true, this._options.parseOleCfb === true);
                sheets.push(parseSheet(name, state, xml, sharedStrings, tables, images, charts, shapes, formControls, embeddings, pivots, slicers, timelines, comments, threadedComments, hyperlinkTargets, i, definedNames, metadata, smartArt));
            }
            return { sheets, styles, theme, persons, date1904, definedNames, metadata };
        }
    }
    function parseContentTypes(xml) {
        const doc = parseXml(xml);
        const defaults = new Map();
        const overrides = new Map();
        const ns = 'http://schemas.openxmlformats.org/package/2006/content-types';
        const defEls = doc.getElementsByTagNameNS(ns, 'Default');
        for (let i = 0; i < defEls.length; i++) {
            const el = defEls[i];
            const ext = el.getAttribute('Extension');
            const ct = el.getAttribute('ContentType');
            if (ext && ct)
                defaults.set(ext.toLowerCase(), ct);
        }
        const ovEls = doc.getElementsByTagNameNS(ns, 'Override');
        for (let i = 0; i < ovEls.length; i++) {
            const el = ovEls[i];
            const part = el.getAttribute('PartName');
            const ct = el.getAttribute('ContentType');
            if (part && ct)
                overrides.set(part, ct);
        }
        return { defaults, overrides };
    }
    function resolveContentType(path, map) {
        const pkgPath = path.startsWith('/') ? path : `/${path}`;
        const ov = map.overrides.get(pkgPath);
        if (ov)
            return ov;
        const dot = path.lastIndexOf('.');
        if (dot >= 0) {
            const ext = path.slice(dot + 1).toLowerCase();
            const def = map.defaults.get(ext);
            if (def)
                return def;
            switch (ext) {
                case 'bin': return 'application/vnd.openxmlformats-officedocument.oleObject';
                case 'pdf': return 'application/pdf';
                case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                case 'docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                case 'pptx': return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
                case 'xls': return 'application/vnd.ms-excel';
                case 'doc': return 'application/msword';
                case 'ppt': return 'application/vnd.ms-powerpoint';
                case 'txt': return 'text/plain';
            }
        }
        return 'application/octet-stream';
    }
    function resolveEmbeddingsForSheet(sheetPath, sheetXml, parts, embeddingBytes, contentTypes, inline, parseOleCfb) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return [];
        const rels = parseRelationships(relsXml);
        const sheetDir = sheetPath.replace(/\/[^/]+$/, '');
        const embRels = [];
        for (const [rId, rel] of rels) {
            const isOle = rel.type.endsWith('/oleObject');
            const isPackage = rel.type.endsWith('/package');
            if (!isOle && !isPackage)
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${sheetDir}/${rel.target}`);
            const fileName = target.slice(target.lastIndexOf('/') + 1) || null;
            embRels.push({
                rId,
                kind: isOle ? 'ole' : 'package',
                target,
                fileName,
            });
        }
        if (embRels.length === 0)
            return [];
        const anchorByRid = parseOleObjectAnchors(sheetXml);
        const out = [];
        for (const rel of embRels) {
            const bytes = embeddingBytes[rel.target];
            if (!bytes) {
                continue;
            }
            const contentType = resolveContentType(rel.target, contentTypes);
            const anchor = anchorByRid.get(rel.rId) ?? null;
            let dataUrl = null;
            if (inline) {
                const safeMime = sanitizeMediaMime(contentType);
                if (safeMime !== 'application/octet-stream' || contentType === 'application/octet-stream') {
                    dataUrl = bytesToDataUrl(bytes, contentType);
                }
            }
            const cfb = (parseOleCfb && rel.kind === 'ole') ? parseCfb(bytes) : null;
            out.push({
                kind: rel.kind,
                contentType,
                fileName: rel.fileName,
                progId: anchor?.progId ?? null,
                col: anchor?.col ?? null,
                row: anchor?.row ?? null,
                endCol: anchor?.endCol ?? null,
                endRow: anchor?.endRow ?? null,
                size: bytes.byteLength,
                dataUrl,
                altText: anchor?.altText ?? null,
                cfb,
            });
        }
        return out;
    }
    function parseOleObjectAnchors(xml) {
        const out = new Map();
        const doc = parseXml(xml);
        const els = doc.getElementsByTagNameNS(NS.main, 'oleObject');
        for (let i = 0; i < els.length; i++) {
            const el = els[i];
            const rId = el.getAttributeNS(NS.rel, 'id');
            if (!rId || out.has(rId))
                continue;
            const progId = el.getAttribute('progId') || null;
            let col = null;
            let row = null;
            let endCol = null;
            let endRow = null;
            const findFrom = el.getElementsByTagNameNS(NS.xdr, 'from').item(0)
                ?? el.getElementsByTagNameNS(NS.main, 'from').item(0);
            const findTo = el.getElementsByTagNameNS(NS.xdr, 'to').item(0)
                ?? el.getElementsByTagNameNS(NS.main, 'to').item(0);
            if (findFrom) {
                col = anchorCellValue(findFrom, 'col');
                row = anchorCellValue(findFrom, 'row');
                if (col === null) {
                    const c = findFrom.getElementsByTagNameNS(NS.main, 'col').item(0);
                    col = c ? Number(c.textContent) : null;
                    if (col !== null && !Number.isFinite(col))
                        col = null;
                }
                if (row === null) {
                    const r = findFrom.getElementsByTagNameNS(NS.main, 'row').item(0);
                    row = r ? Number(r.textContent) : null;
                    if (row !== null && !Number.isFinite(row))
                        row = null;
                }
            }
            if (findTo) {
                endCol = anchorCellValue(findTo, 'col');
                endRow = anchorCellValue(findTo, 'row');
                if (endCol === null) {
                    const c = findTo.getElementsByTagNameNS(NS.main, 'col').item(0);
                    endCol = c ? Number(c.textContent) : null;
                    if (endCol !== null && !Number.isFinite(endCol))
                        endCol = null;
                }
                if (endRow === null) {
                    const r = findTo.getElementsByTagNameNS(NS.main, 'row').item(0);
                    endRow = r ? Number(r.textContent) : null;
                    if (endRow !== null && !Number.isFinite(endRow))
                        endRow = null;
                }
            }
            out.set(rId, { progId, col, row, endCol, endRow, altText: null });
        }
        return out;
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
    function resolveDrawingsForSheet(sheetPath, parts, media, sharedStrings = []) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        if (!relsXml)
            return { images: [], charts: [], shapes: [], formControls: [], smartArt: [] };
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const images = [];
        const charts = [];
        const shapes = [];
        const formControls = [];
        const smartArt = [];
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
            const parsed = parseDrawing(drawingXml, drawingRels, drawingDir, parts, media, sharedStrings);
            images.push(...parsed.images);
            charts.push(...parsed.charts);
            shapes.push(...parsed.shapes);
            formControls.push(...parsed.formControls);
            smartArt.push(...parsed.smartArt);
        }
        return { images, charts, shapes, formControls, smartArt };
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
    function resolveSlicersAndTimelinesForSheet(sheetPath, parts) {
        const relsPath = sheetPath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
        const relsXml = parts[relsPath];
        const slicers = [];
        const timelines = [];
        if (!relsXml)
            return { slicers, timelines };
        const rels = parseRelationships(relsXml);
        const dir = sheetPath.replace(/\/[^/]+$/, '');
        const slicerCacheByName = indexSlicerCaches(parts);
        const timelineCacheByName = indexTimelineCaches(parts);
        for (const [, rel] of rels) {
            const type = rel.type;
            const isSlicer = type.endsWith('/slicer');
            const isTimeline = type.endsWith('/timeline');
            if (!isSlicer && !isTimeline)
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${dir}/${rel.target}`);
            const xml = parts[target];
            if (!xml)
                continue;
            if (isSlicer) {
                slicers.push(...parseSlicerFile(xml, slicerCacheByName));
            }
            else {
                timelines.push(...parseTimelineFile(xml, timelineCacheByName));
            }
        }
        return { slicers, timelines };
    }
    function indexSlicerCaches(parts) {
        const out = new Map();
        for (const path of Object.keys(parts)) {
            if (!/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path))
                continue;
            const xml = parts[path];
            if (!xml)
                continue;
            const name = peekRootAttr(xml, 'slicerCacheDefinition', 'name');
            if (name)
                out.set(name, xml);
        }
        return out;
    }
    function indexTimelineCaches(parts) {
        const out = new Map();
        for (const path of Object.keys(parts)) {
            if (!/^xl\/timelineCaches\/[^/]+\.xml$/i.test(path))
                continue;
            const xml = parts[path];
            if (!xml)
                continue;
            const name = peekRootAttr(xml, 'timelineCacheDefinition', 'name');
            if (name)
                out.set(name, xml);
        }
        return out;
    }
    function peekRootAttr(xml, rootLocal, attr) {
        const doc = parseXml(xml);
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === rootLocal) {
                return el.getAttribute(attr);
            }
        }
        return null;
    }
    function parseSlicerFile(xml, slicerCacheByName) {
        const doc = parseXml(xml);
        const out = [];
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'slicer')
                continue;
            const name = el.getAttribute('name');
            const cache = el.getAttribute('cache');
            if (!name || !cache)
                continue;
            const caption = el.getAttribute('caption');
            const columnCountAttr = el.getAttribute('columnCount');
            const columnCount = numOrNull(columnCountAttr);
            const style = el.getAttribute('style') || null;
            const showCaptionAttr = el.getAttribute('showCaption');
            const showCaption = !(showCaptionAttr === '0' || showCaptionAttr === 'false');
            const rowHeightAttr = el.getAttribute('rowHeight');
            const rowHeight = numOrNull(rowHeightAttr);
            const cacheXml = slicerCacheByName.get(cache);
            let sourceName = null;
            let selectedItems = [];
            let allItems = [];
            if (cacheXml) {
                const details = readSlicerCache(cacheXml);
                sourceName = details.sourceName;
                selectedItems = details.selectedItems;
                allItems = details.allItems;
            }
            out.push({
                name,
                caption: caption || null,
                cache,
                sourceName,
                columnCount,
                style,
                showCaption,
                rowHeight,
                selectedItems,
                allItems,
            });
        }
        return out;
    }
    function parseTimelineFile(xml, timelineCacheByName) {
        const doc = parseXml(xml);
        const out = [];
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'timeline')
                continue;
            const name = el.getAttribute('name');
            const cache = el.getAttribute('cache');
            if (!name || !cache)
                continue;
            const caption = el.getAttribute('caption');
            const style = el.getAttribute('style') || null;
            const showHeader = !truthyZero(el.getAttribute('showHeader'));
            const showSelectionLabel = !truthyZero(el.getAttribute('showSelectionLabel'));
            const showTimeLevel = !truthyZero(el.getAttribute('showTimeLevel'));
            const showHorizontalScrollbar = !truthyZero(el.getAttribute('showHorizontalScrollbar'));
            let level = null;
            let selectedRange = null;
            const nested = firstDescendantByLocalName(el, 'timelineViewState');
            const src = nested ?? el;
            const levelAttr = src.getAttribute('level');
            level = normaliseTimelineLevel(levelAttr);
            const selectionEl = firstDescendantByLocalName(src, 'selection') ?? src;
            const startAttr = selectionEl.getAttribute('startDate');
            const endAttr = selectionEl.getAttribute('endDate');
            if (startAttr !== null || endAttr !== null) {
                selectedRange = { start: startAttr, end: endAttr };
            }
            const cacheXml = timelineCacheByName.get(cache);
            let sourceName = null;
            let bounds = null;
            if (cacheXml) {
                const details = readTimelineCache(cacheXml);
                sourceName = details.sourceName;
                bounds = details.bounds;
            }
            out.push({
                name,
                caption: caption || null,
                cache,
                sourceName,
                level,
                selectedRange,
                showHeader,
                showSelectionLabel,
                showTimeLevel,
                showHorizontalScrollbar,
                style,
                bounds,
            });
        }
        return out;
    }
    function readSlicerCache(xml) {
        const doc = parseXml(xml);
        let sourceName = null;
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === 'slicerCacheDefinition') {
                sourceName = el.getAttribute('sourceName');
                break;
            }
        }
        const selectedItems = [];
        const allItems = [];
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName !== 'tabular')
                continue;
            const items = firstDescendantByLocalName(el, 'items');
            if (!items)
                continue;
            for (let j = 0; j < items.childNodes.length; j++) {
                const node = items.childNodes[j];
                if (node.nodeType !== 1)
                    continue;
                const item = node;
                if (item.localName !== 'i')
                    continue;
                const name = item.getAttribute('n') ?? item.textContent;
                if (!name || name.length === 0)
                    continue;
                allItems.push(name);
                const selected = item.getAttribute('s');
                if (selected === '1' || selected === 'true')
                    selectedItems.push(name);
            }
        }
        return { sourceName, selectedItems, allItems };
    }
    function readTimelineCache(xml) {
        const doc = parseXml(xml);
        let sourceName = null;
        let bounds = null;
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === 'timelineCacheDefinition') {
                sourceName = el.getAttribute('sourceName');
            }
            if (el.localName === 'bounds') {
                const minAttr = el.getAttribute('startDate');
                const maxAttr = el.getAttribute('endDate');
                if (minAttr && maxAttr) {
                    bounds = { min: minAttr, max: maxAttr };
                }
            }
        }
        return { sourceName, bounds };
    }
    function firstDescendantByLocalName(root, localName) {
        const all = root.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            if (all[i].localName === localName)
                return all[i];
        }
        return null;
    }
    function numOrNull(attr) {
        if (attr === null || attr === '')
            return null;
        const n = Number(attr);
        return Number.isFinite(n) ? n : null;
    }
    function truthyZero(v) {
        return v === '0' || v === 'false';
    }
    function normaliseTimelineLevel(raw) {
        if (!raw)
            return null;
        const lower = raw.toLowerCase();
        if (lower === 'years' || lower === 'quarters' || lower === 'months' || lower === 'days') {
            return lower;
        }
        const numericMap = {
            '0': 'years',
            '1': 'quarters',
            '2': 'months',
            '3': 'days',
        };
        return numericMap[raw] ?? null;
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
    function parseDrawing(xml, rels, drawingDir, parts, media, sharedStrings = []) {
        const doc = parseXml(xml);
        const images = [];
        const charts = [];
        const shapes = [];
        const formControls = [];
        const smartArt = [];
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
                let chartXml;
                if (rId) {
                    const rel = rels.get(rId);
                    if (rel) {
                        const chartPath = rel.target.startsWith('/')
                            ? rel.target.slice(1)
                            : normaliseRelPath(`${drawingDir}/${rel.target}`);
                        chartXml = parts[chartPath];
                        if (chartXml)
                            chartType = peekChartType(chartXml, kind);
                    }
                }
                let model = null;
                if (kind === 'classic' && chartXml) {
                    try {
                        const parsed = parseChart(chartXml, sharedStrings);
                        if (parsed.kind !== 'unknown')
                            model = parsed;
                    }
                    catch {
                        model = null;
                    }
                }
                charts.push({
                    kind,
                    chartType,
                    col: col ?? 0,
                    row: row ?? 0,
                    endCol,
                    endRow,
                    model,
                });
            }
            const graphicFrame = anchor.getElementsByTagNameNS(NS.xdr, 'graphicFrame').item(0);
            if (graphicFrame) {
                const graphicData = graphicFrame.getElementsByTagNameNS(NS.a, 'graphicData').item(0);
                const uri = graphicData?.getAttribute('uri') ?? '';
                if (uri === SMARTART_GRAPHIC_URI) {
                    const relIdsEl = graphicData.getElementsByTagNameNS(NS.dgm, 'relIds').item(0);
                    if (relIdsEl) {
                        const dmId = relIdsEl.getAttributeNS(NS.rel, 'dm');
                        const loId = relIdsEl.getAttributeNS(NS.rel, 'lo');
                        const dmRel = dmId ? rels.get(dmId) : undefined;
                        const loRel = loId ? rels.get(loId) : undefined;
                        let dataXml = null;
                        let layoutXml = null;
                        if (dmRel) {
                            const dataPath = dmRel.target.startsWith('/')
                                ? dmRel.target.slice(1)
                                : normaliseRelPath(`${drawingDir}/${dmRel.target}`);
                            dataXml = parts[dataPath] ?? null;
                        }
                        if (loRel) {
                            const layoutPath = loRel.target.startsWith('/')
                                ? loRel.target.slice(1)
                                : normaliseRelPath(`${drawingDir}/${loRel.target}`);
                            layoutXml = parts[layoutPath] ?? null;
                        }
                        const cNvPr = graphicFrame.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
                        const name = cNvPr?.getAttribute('name') || null;
                        const model = dataXml ? parseSmartArt(dataXml, layoutXml) : null;
                        smartArt.push({
                            name,
                            col: col ?? 0,
                            row: row ?? 0,
                            endCol,
                            endRow,
                            model,
                        });
                    }
                }
            }
            const spEls = [
                ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'sp')),
                ...Array.from(anchor.getElementsByTagNameNS(NS.xdr, 'cxnSp')),
            ];
            const ctrlPropRelHit = anchorHasCtrlPropRel(anchor, rels, drawingDir, parts);
            let formControlConsumed = false;
            if (ctrlPropRelHit && spEls.length > 0) {
                const fc = parseFormControl(spEls[0], rels, drawingDir, parts, col, row, colOff, rowOff, endCol, endRow, anchor);
                if (fc) {
                    formControls.push(fc);
                    formControlConsumed = true;
                }
            }
            if (!formControlConsumed) {
                for (const el of spEls) {
                    const kind = el.localName === 'cxnSp' ? 'connector' : 'shape';
                    const shape = parseShape(el, kind, col, row, endCol, endRow);
                    if (shape)
                        shapes.push(shape);
                }
            }
        }
        return { images, charts, shapes, formControls, smartArt };
    }
    function anchorHasCtrlPropRel(anchor, rels, drawingDir, parts) {
        const descendants = anchor.getElementsByTagName('*');
        for (let i = 0; i < descendants.length; i++) {
            const rId = descendants[i].getAttributeNS(NS.rel, 'id');
            if (!rId)
                continue;
            const rel = rels.get(rId);
            if (!rel)
                continue;
            if (rel.type.endsWith('/ctrlProp')) {
                const target = rel.target.startsWith('/')
                    ? rel.target.slice(1)
                    : normaliseRelPath(`${drawingDir}/${rel.target}`);
                if (parts[target])
                    return true;
            }
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${drawingDir}/${rel.target}`);
            if (/^xl\/ctrlProps\/.*\.xml$/i.test(target) && parts[target])
                return true;
        }
        return false;
    }
    const FORM_CONTROL_OBJECT_TYPES = {
        'button': 'button',
        'checkbox': 'checkbox',
        'radio': 'radio',
        'drop': 'combo',
        'combobox': 'combo',
        'list': 'list',
        'listbox': 'list',
        'scroll': 'scrollbar',
        'scrollbar': 'scrollbar',
        'spin': 'spinner',
        'spinner': 'spinner',
        'groupbox': 'groupBox',
        'label': 'label',
        'dialog': 'dialog',
    };
    function normaliseFormControlKind(raw) {
        if (!raw)
            return 'unknown';
        const key = raw.toLowerCase();
        return FORM_CONTROL_OBJECT_TYPES[key] ?? 'unknown';
    }
    function parseFormControl(sp, rels, drawingDir, parts, col, row, colOff, rowOff, endCol, endRow, anchor) {
        let ctrlPropXml = null;
        const descendants = anchor.getElementsByTagName('*');
        for (let i = 0; i < descendants.length; i++) {
            const rId = descendants[i].getAttributeNS(NS.rel, 'id');
            if (!rId)
                continue;
            const rel = rels.get(rId);
            if (!rel)
                continue;
            const target = rel.target.startsWith('/')
                ? rel.target.slice(1)
                : normaliseRelPath(`${drawingDir}/${rel.target}`);
            if (!/^xl\/ctrlProps\/.*\.xml$/i.test(target))
                continue;
            ctrlPropXml = parts[target] ?? null;
            if (ctrlPropXml)
                break;
        }
        const cNvPr = sp.getElementsByTagNameNS(NS.xdr, 'cNvPr').item(0);
        const shapeName = cNvPr?.getAttribute('name') || null;
        const altText = cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('title') || null;
        let labelFromText = null;
        const txBody = sp.getElementsByTagNameNS(NS.xdr, 'txBody').item(0);
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
                labelFromText = joined;
        }
        let kind = 'unknown';
        let linkedCell = null;
        let inputRange = null;
        let checked = null;
        let min = null;
        let max = null;
        let inc = null;
        let page = null;
        let val = null;
        let dropLines = null;
        if (ctrlPropXml) {
            const ctrlDoc = parseXml(ctrlPropXml);
            let root = ctrlDoc.documentElement;
            if (root && root.localName !== 'formControlPr') {
                const found = ctrlDoc.getElementsByTagName('formControlPr').item(0);
                if (found)
                    root = found;
            }
            if (root) {
                kind = normaliseFormControlKind(root.getAttribute('objectType'));
                linkedCell = root.getAttribute('fmlaLink') || null;
                inputRange = root.getAttribute('fmlaRange') || null;
                const checkedAttr = root.getAttribute('checked');
                if (checkedAttr === 'Checked' || checkedAttr === '1' || checkedAttr === 'true') {
                    checked = true;
                }
                else if (checkedAttr === 'Unchecked' || checkedAttr === '0' || checkedAttr === 'false') {
                    checked = false;
                }
                else if (kind === 'checkbox' || kind === 'radio') {
                    checked = false;
                }
                const numAttr = (name) => {
                    const raw = root.getAttribute(name);
                    if (raw === null)
                        return null;
                    const n = Number(raw);
                    return Number.isFinite(n) ? n : null;
                };
                min = numAttr('min');
                max = numAttr('max');
                inc = numAttr('inc');
                page = numAttr('page');
                val = numAttr('val');
                dropLines = numAttr('dropLines');
            }
        }
        else if (shapeName) {
            const lower = shapeName.toLowerCase();
            if (lower.startsWith('check box'))
                kind = 'checkbox';
            else if (lower.startsWith('option button'))
                kind = 'radio';
            else if (lower.startsWith('button'))
                kind = 'button';
            else if (lower.startsWith('scroll bar'))
                kind = 'scrollbar';
            else if (lower.startsWith('spinner') || lower.startsWith('spin button'))
                kind = 'spinner';
            else if (lower.startsWith('drop down'))
                kind = 'combo';
            else if (lower.startsWith('list box'))
                kind = 'list';
            else if (lower.startsWith('group box'))
                kind = 'groupBox';
            else if (lower.startsWith('label'))
                kind = 'label';
        }
        let label = null;
        if (kind === 'button' || kind === 'label' || kind === 'groupBox') {
            label = labelFromText ?? shapeName;
        }
        else {
            label = labelFromText ?? shapeName;
        }
        return {
            kind,
            col: col ?? 0,
            row: row ?? 0,
            colOff: colOff ?? 0,
            rowOff: rowOff ?? 0,
            endCol,
            endRow,
            endColOff: resolveEndOffset(anchor, 'colOff'),
            endRowOff: resolveEndOffset(anchor, 'rowOff'),
            label,
            linkedCell,
            inputRange,
            checked,
            min, max, inc, page, val,
            dropLines,
            altText,
        };
    }
    function resolveEndOffset(anchor, attr) {
        const to = findDirectChildNS(anchor, NS.xdr, 'to');
        if (!to)
            return null;
        const el = findDirectChildNS(to, NS.xdr, attr);
        if (!el)
            return null;
        const n = Number(el.textContent);
        return Number.isFinite(n) ? n : null;
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
    function parseSheet(name, state, xml, sharedStrings, tables = [], images = [], charts = [], shapes = [], formControls = [], embeddings = [], pivots = [], slicers = [], timelines = [], comments = [], threadedComments = [], hyperlinkTargets = new Map(), sheetIndex = 0, definedNames = [], metadata = null, smartArt = []) {
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
            charts, shapes, smartArt, formControls, embeddings, pivots, slicers, timelines, extensions, comments, threadedComments, view, outline,
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
                const width = range.endCol - range.col + 1;
                const height = range.endRow - range.row + 1;
                const total = width * height;
                if (total > MAX_RANGE_EXPANSION_CELLS) {
                    out.push({
                        col: range.col,
                        row: range.row,
                        target,
                        location: location || null,
                        tooltip: tooltip || null,
                        display: display || null,
                    });
                    continue;
                }
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
                const width = range.endCol - range.col + 1;
                const height = range.endRow - range.row + 1;
                if (width * height > MAX_RANGE_EXPANSION_CELLS) {
                    const clampedCols = Math.min(width, MAX_RANGE_EXPANSION_CELLS);
                    out.push({
                        col: range.col,
                        row: range.row,
                        endCol: range.col + clampedCols - 1,
                        endRow: range.row,
                        options,
                    });
                    continue;
                }
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

    const DEFAULT_FILL = '#e0e0e0';
    function wrap(inner) {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">${inner}</svg>`;
    }
    function regularPolygon(sides) {
        const cx = 50, cy = 50, r = 48;
        const pts = [];
        for (let i = 0; i < sides; i++) {
            const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
            const x = cx + r * Math.cos(angle);
            const y = cy + r * Math.sin(angle);
            pts.push(`${x.toFixed(2)},${y.toFixed(2)}`);
        }
        return pts.join(' ');
    }
    function starPoints() {
        const cx = 50, cy = 50, rOuter = 48, rInner = 20;
        const pts = [];
        for (let i = 0; i < 10; i++) {
            const r = i % 2 === 0 ? rOuter : rInner;
            const angle = -Math.PI / 2 + (i * Math.PI) / 5;
            const x = cx + r * Math.cos(angle);
            const y = cy + r * Math.sin(angle);
            pts.push(`${x.toFixed(2)},${y.toFixed(2)}`);
        }
        return pts.join(' ');
    }
    function presetBody(preset, stroke) {
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
                return `<polygon points="2,2 2,98 98,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'diamond':
            case 'flowChartDecision':
                return `<polygon points="50,2 98,50 50,98 2,50" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'parallelogram':
                return `<polygon points="22,2 98,2 78,98 2,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'trapezoid':
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
                return `<path d="M2 30 L60 30 L60 10 L98 50 L60 90 L60 70 L2 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'leftArrow':
                return `<path d="M98 30 L40 30 L40 10 L2 50 L40 90 L40 70 L98 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'upArrow':
                return `<path d="M30 98 L30 40 L10 40 L50 2 L90 40 L70 40 L70 98 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'downArrow':
                return `<path d="M30 2 L30 60 L10 60 L50 98 L90 60 L70 60 L70 2 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'leftRightArrow':
                return `<path d="M2 50 L20 30 L20 40 L80 40 L80 30 L98 50 L80 70 L80 60 L20 60 L20 70 Z" fill="${DEFAULT_FILL}" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'flowChartTerminator':
                return `<rect x="1" y="1" width="98" height="98" rx="49" ry="49" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'callout1':
            case 'wedgeRectCallout':
                return `<rect x="1" y="1" width="98" height="78" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                    `<polygon points="15,79 35,79 8,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'wedgeEllipseCallout':
                return `<ellipse cx="50" cy="40" rx="49" ry="39" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>` +
                    `<polygon points="20,75 38,75 8,98" fill="none" stroke="${s}" stroke-width="1" vector-effect="non-scaling-stroke"/>`;
            case 'cloudCallout':
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
    function renderShapePreset(preset, opts) {
        if (preset === null || preset === '')
            return null;
        const body = presetBody(preset, opts.stroke);
        if (body === null)
            return null;
        return wrap(body);
    }

    const SVG_NS$1 = 'http://www.w3.org/2000/svg';
    const PALETTE$1 = [
        '#5b9bd5',
        '#ed7d31',
        '#a5a5a5',
        '#ffc000',
        '#4472c4',
        '#70ad47',
        '#264478',
        '#9e480e',
    ];
    const GRID_COLOR = '#ddd';
    const AXIS_COLOR = '#888';
    const TEXT_COLOR$1 = '#333';
    function makeSvg(width, height) {
        const svg = document.createElementNS(SVG_NS$1, 'svg');
        svg.setAttribute('xmlns', SVG_NS$1);
        svg.setAttribute('width', String(width));
        svg.setAttribute('height', String(height));
        svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
        svg.setAttribute('class', 'xlsx-chart-svg');
        return svg;
    }
    function el(tag, attrs) {
        const node = document.createElementNS(SVG_NS$1, tag);
        for (const [k, v] of Object.entries(attrs))
            node.setAttribute(k, String(v));
        return node;
    }
    function textEl(x, y, text, attrs = {}) {
        const t = document.createElementNS(SVG_NS$1, 'text');
        t.setAttribute('x', String(x));
        t.setAttribute('y', String(y));
        t.setAttribute('fill', TEXT_COLOR$1);
        t.setAttribute('font-family', 'system-ui, sans-serif');
        t.setAttribute('font-size', '12');
        for (const [k, v] of Object.entries(attrs))
            t.setAttribute(k, String(v));
        t.textContent = text;
        return t;
    }
    function colorFor(series, idx) {
        if (series.color && /^#[0-9a-f]{6}$/i.test(series.color))
            return series.color;
        return PALETTE$1[idx % PALETTE$1.length];
    }
    function niceMax(max) {
        if (!Number.isFinite(max) || max <= 0)
            return 1;
        const mag = Math.pow(10, Math.floor(Math.log10(max)));
        const norm = max / mag;
        let nice;
        if (norm <= 1)
            nice = 1;
        else if (norm <= 2)
            nice = 2;
        else if (norm <= 5)
            nice = 5;
        else
            nice = 10;
        return nice * mag;
    }
    function niceMin(min) {
        if (min >= 0)
            return 0;
        const abs = Math.abs(min);
        return -niceMax(abs);
    }
    function seriesRange(series) {
        let min = 0;
        let max = 0;
        let seen = false;
        for (const s of series) {
            for (const v of s.values) {
                if (v === null || !Number.isFinite(v))
                    continue;
                if (!seen) {
                    min = v;
                    max = v;
                    seen = true;
                    continue;
                }
                if (v < min)
                    min = v;
                if (v > max)
                    max = v;
            }
        }
        if (!seen)
            return { min: 0, max: 1 };
        if (min === max) {
            if (max > 0)
                min = 0;
            else if (max < 0)
                max = 0;
            else {
                max = 1;
            }
        }
        return { min, max };
    }
    function baseLayout(width, height, hasLegend, legendPos) {
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
    function appendTitle(svg, title, layout) {
        const t = textEl(layout.width / 2, 20, title, {
            'text-anchor': 'middle',
            'font-size': '14',
            'font-weight': '600',
        });
        t.setAttribute('class', 'xlsx-chart-title');
        svg.appendChild(t);
    }
    function appendLegend(svg, model, layout) {
        if (model.legend === 'none')
            return;
        const entries = model.series
            .map((s, i) => ({ name: s.name, color: colorFor(s, i) }))
            .filter((e) => e.name !== null);
        if (!entries.length)
            return;
        const group = document.createElementNS(SVG_NS$1, 'g');
        group.setAttribute('class', 'xlsx-chart-legend');
        let x;
        let y;
        let stepX = 0;
        let stepY = 0;
        const rowHeight = 16;
        const swatchSize = 10;
        const gap = 6;
        if (model.legend === 'right') {
            x = layout.plotX + layout.plotW + 16;
            y = layout.plotY + 4;
            stepY = rowHeight;
        }
        else if (model.legend === 'left') {
            x = 10;
            y = layout.plotY + 4;
            stepY = rowHeight;
        }
        else if (model.legend === 'top') {
            x = layout.plotX;
            y = layout.plotY - 16;
            stepX = 80;
        }
        else {
            x = layout.plotX;
            y = layout.height - 16;
            stepX = 80;
        }
        entries.forEach((entry, i) => {
            const item = document.createElementNS(SVG_NS$1, 'g');
            item.setAttribute('class', 'xlsx-chart-legend-entry');
            const lx = x + stepX * i;
            const ly = y + stepY * i;
            const rect = el('rect', {
                x: lx, y: ly, width: swatchSize, height: swatchSize,
                fill: entry.color,
            });
            item.appendChild(rect);
            const label = textEl(lx + swatchSize + gap, ly + swatchSize - 1, entry.name ?? '', { 'font-size': '11' });
            item.appendChild(label);
            group.appendChild(item);
        });
        svg.appendChild(group);
    }
    function appendYAxis(svg, layout, axisMin, axisMax) {
        const g = document.createElementNS(SVG_NS$1, 'g');
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
        g.appendChild(el('line', {
            x1: layout.plotX, x2: layout.plotX,
            y1: layout.plotY, y2: layout.plotY + layout.plotH,
            stroke: AXIS_COLOR, 'stroke-width': 1,
        }));
        svg.appendChild(g);
    }
    function appendXAxisCategory(svg, layout, categories) {
        if (!categories.length)
            return;
        const g = document.createElementNS(SVG_NS$1, 'g');
        g.setAttribute('class', 'xlsx-chart-xaxis');
        const slot = layout.plotW / categories.length;
        categories.forEach((label, i) => {
            const cx = layout.plotX + slot * (i + 0.5);
            g.appendChild(textEl(cx, layout.plotY + layout.plotH + 14, label ?? '', {
                'text-anchor': 'middle', 'font-size': '10',
            }));
        });
        g.appendChild(el('line', {
            x1: layout.plotX, x2: layout.plotX + layout.plotW,
            y1: layout.plotY + layout.plotH, y2: layout.plotY + layout.plotH,
            stroke: AXIS_COLOR, 'stroke-width': 1,
        }));
        svg.appendChild(g);
    }
    function formatTick(v) {
        if (!Number.isFinite(v))
            return '';
        const abs = Math.abs(v);
        if (abs >= 1000)
            return String(Math.round(v));
        if (abs >= 10 || abs === 0)
            return String(Math.round(v));
        const s = v.toFixed(1);
        return s.endsWith('.0') ? s.slice(0, -2) : s;
    }
    function renderColumn(model, layout, svg) {
        const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
        if (n === 0)
            return;
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
        const totals = percent ? categoryTotals(model.series, n) : [];
        const stackPos = new Array(n).fill(0);
        const stackNeg = new Array(n).fill(0);
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            for (let ci = 0; ci < n; ci++) {
                const vRaw = s.values[ci];
                const vRawFinite = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
                const x = stacked
                    ? layout.plotX + slot * ci + groupPad
                    : layout.plotX + slot * ci + groupPad + barW * si;
                let y;
                let h;
                if (vRawFinite === null) {
                    y = zeroY;
                    h = 0;
                }
                else if (stacked) {
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
                    }
                    else {
                        const base = stackNeg[ci];
                        const bottom = base + vShare;
                        const yTop = layout.plotY + layout.plotH - ((base - axisMin) / range) * layout.plotH;
                        const yBot = layout.plotY + layout.plotH - ((bottom - axisMin) / range) * layout.plotH;
                        y = yTop;
                        h = yBot - yTop;
                        stackNeg[ci] = bottom;
                    }
                }
                else if (vRawFinite >= 0) {
                    const top = layout.plotY + layout.plotH - (vRawFinite / range) * layout.plotH - (axisMin < 0 ? -axisMin / range * layout.plotH : 0);
                    y = top;
                    h = zeroY - top;
                }
                else {
                    const bottom = layout.plotY + layout.plotH - ((vRawFinite - axisMin) / range) * layout.plotH;
                    y = zeroY;
                    h = bottom - zeroY;
                }
                group.appendChild(el('rect', {
                    x, y, width: Math.max(0, barW - 1), height: Math.max(0, h),
                    fill: color,
                }));
                if (s.dataLabels?.show && vRawFinite !== null) {
                    const pos = s.dataLabels.position ?? 'outEnd';
                    const cx = x + Math.max(0, barW - 1) / 2;
                    let ly;
                    if (pos === 'ctr')
                        ly = y + h / 2 + 3;
                    else if (pos === 'inBase')
                        ly = y + h - 3;
                    else
                        ly = y - 3;
                    appendDataLabel(group, cx, ly, formatDataLabel(vRawFinite), 'middle');
                }
            }
            svg.appendChild(group);
        });
    }
    function stackedRange(series, n, percent) {
        if (percent)
            return { min: 0, max: 100 };
        let max = 0;
        let min = 0;
        for (let ci = 0; ci < n; ci++) {
            let pos = 0, neg = 0;
            for (const s of series) {
                const v = s.values[ci];
                if (v === null || v === undefined || !Number.isFinite(v))
                    continue;
                if (v >= 0)
                    pos += v;
                else
                    neg += v;
            }
            if (pos > max)
                max = pos;
            if (neg < min)
                min = neg;
        }
        if (max === 0 && min === 0)
            max = 1;
        return { min, max };
    }
    function categoryTotals(series, n) {
        const out = new Array(n).fill(0);
        for (let ci = 0; ci < n; ci++) {
            for (const s of series) {
                const v = s.values[ci];
                if (v === null || v === undefined || !Number.isFinite(v))
                    continue;
                if (v > 0)
                    out[ci] += v;
            }
        }
        return out;
    }
    function renderBar(model, layout, svg) {
        const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
        if (n === 0)
            return;
        const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
        const percent = model.grouping === 'percentStacked';
        const { min, max } = stacked
            ? stackedRange(model.series, n, percent)
            : seriesRange(model.series);
        const axisMax = max > 0 ? niceMax(max) : 0;
        const axisMin = min < 0 ? niceMin(min) : 0;
        const gx = document.createElementNS(SVG_NS$1, 'g');
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
        const gy = document.createElementNS(SVG_NS$1, 'g');
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
        const totals = percent ? categoryTotals(model.series, n) : [];
        const stackPos = new Array(n).fill(0);
        const stackNeg = new Array(n).fill(0);
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            for (let ci = 0; ci < n; ci++) {
                const vRaw = s.values[ci];
                const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
                const y = stacked
                    ? layout.plotY + slot * ci + groupPad
                    : layout.plotY + slot * ci + groupPad + barH * si;
                let x;
                let w;
                if (v === null) {
                    x = zeroX;
                    w = 0;
                }
                else if (stacked) {
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
                    }
                    else {
                        const base = stackNeg[ci];
                        const bottom = base + vShare;
                        const xStart = layout.plotX + ((bottom - axisMin) / range) * layout.plotW;
                        const xEnd = layout.plotX + ((base - axisMin) / range) * layout.plotW;
                        x = xStart;
                        w = xEnd - xStart;
                        stackNeg[ci] = bottom;
                    }
                }
                else if (v >= 0) {
                    x = zeroX;
                    w = (v / range) * layout.plotW;
                }
                else {
                    const end = layout.plotX + ((v - axisMin) / range) * layout.plotW;
                    x = end;
                    w = zeroX - end;
                }
                group.appendChild(el('rect', {
                    x, y, width: Math.max(0, w), height: Math.max(0, barH - 1),
                    fill: color,
                }));
                if (s.dataLabels?.show && v !== null) {
                    const pos = s.dataLabels.position ?? 'outEnd';
                    const cy = y + Math.max(0, barH - 1) / 2 + 3;
                    let lx;
                    let anchor = 'start';
                    if (pos === 'ctr') {
                        lx = x + w / 2;
                        anchor = 'middle';
                    }
                    else if (pos === 'inBase') {
                        lx = x + 3;
                        anchor = 'start';
                    }
                    else {
                        lx = x + w + 3;
                        anchor = 'start';
                    }
                    appendDataLabel(group, lx, cy, formatDataLabel(v), anchor);
                }
            }
            svg.appendChild(group);
        });
    }
    function renderLine(model, layout, svg) {
        const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
        if (n === 0)
            return;
        const { min, max } = seriesRange(model.series);
        const axisMax = max > 0 ? niceMax(max) : 0;
        const axisMin = min < 0 ? niceMin(min) : 0;
        appendYAxis(svg, layout, axisMin, axisMax);
        const categoryLabels = model.categories.length ? model.categories
            : Array.from({ length: n }, (_, i) => String(i + 1));
        appendXAxisCategory(svg, layout, categoryLabels);
        const slot = n > 1 ? layout.plotW / (n - 1) : layout.plotW;
        const pointX = (ci) => n > 1
            ? layout.plotX + slot * ci
            : layout.plotX + layout.plotW / 2;
        const pointY = (v) => layout.plotY + layout.plotH - ((v - axisMin) / (axisMax - axisMin || 1)) * layout.plotH;
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            const pts = [];
            for (let ci = 0; ci < n; ci++) {
                const vRaw = s.values[ci];
                const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
                if (v === null)
                    continue;
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
                for (let ci = 0; ci < n; ci++) {
                    const vRaw = s.values[ci];
                    const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : null;
                    if (v === null)
                        continue;
                    group.appendChild(el('circle', {
                        cx: pointX(ci), cy: pointY(v), r: 3,
                        fill: color,
                    }));
                    if (s.dataLabels?.show) {
                        const pos = s.dataLabels.position ?? 't';
                        let ly;
                        if (pos === 'b')
                            ly = pointY(v) + 14;
                        else if (pos === 'ctr')
                            ly = pointY(v) + 3;
                        else
                            ly = pointY(v) - 6;
                        appendDataLabel(group, pointX(ci), ly, formatDataLabel(v), 'middle');
                    }
                }
            }
            svg.appendChild(group);
        });
    }
    function renderScatter(model, layout, svg) {
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
        const projX = (v) => layout.plotX + ((v - axisMinX) / rangeX) * layout.plotW;
        const projY = (v) => layout.plotY + layout.plotH - ((v - axisMinY) / rangeY) * layout.plotH;
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            const xs = s.xValues ?? [];
            const ys = s.values;
            const count = Math.max(xs.length, ys.length);
            for (let i = 0; i < count; i++) {
                const xv = xs[i];
                const yv = ys[i];
                if (xv === null || xv === undefined || !Number.isFinite(xv))
                    continue;
                if (yv === null || yv === undefined || !Number.isFinite(yv))
                    continue;
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
    function scatterXRange(series) {
        let min = 0, max = 0, seen = false;
        for (const s of series) {
            const xs = s.xValues;
            if (!xs)
                continue;
            for (const v of xs) {
                if (v === null || !Number.isFinite(v))
                    continue;
                if (!seen) {
                    min = v;
                    max = v;
                    seen = true;
                    continue;
                }
                if (v < min)
                    min = v;
                if (v > max)
                    max = v;
            }
        }
        if (!seen)
            return { min: 0, max: 1 };
        if (min === max) {
            if (max > 0)
                min = 0;
            else if (max < 0)
                max = 0;
            else {
                max = 1;
            }
        }
        return { min, max };
    }
    function appendXAxisNumeric(svg, layout, axisMin, axisMax) {
        const g = document.createElementNS(SVG_NS$1, 'g');
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
    function renderArea(model, layout, svg) {
        const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
        if (n === 0)
            return;
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
        const pointX = (ci) => n > 1
            ? layout.plotX + slot * ci
            : layout.plotX + layout.plotW / 2;
        const yFor = (v) => layout.plotY + layout.plotH - ((v - axisMin) / range) * layout.plotH;
        const baselineY = yFor(0);
        const totals = percent ? categoryTotals(model.series, n) : [];
        const prevTop = new Array(n).fill(0);
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            const topPts = [];
            const bottomPts = [];
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
                }
                else {
                    topPts.push({ x, y: yFor(vFinite) });
                    bottomPts.push({ x, y: baselineY });
                }
            }
            const parts = [];
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
    function renderPie(model, layout, svg) {
        renderPieLike(model, layout, svg, 0);
    }
    function renderDoughnut(model, layout, svg) {
        renderPieLike(model, layout, svg, 0.5);
    }
    function renderPieLike(model, layout, svg, innerRatio) {
        const series = model.series[0];
        if (!series)
            return;
        const values = series.values
            .map((v) => (v !== null && Number.isFinite(v) && v > 0 ? v : 0));
        const total = values.reduce((a, b) => a + b, 0);
        if (total <= 0)
            return;
        const size = Math.min(layout.plotW, layout.plotH);
        const cx = layout.plotX + layout.plotW / 2;
        const cy = layout.plotY + layout.plotH / 2;
        const rOuter = size / 2 - 4;
        const rInner = innerRatio > 0 ? rOuter * innerRatio : 0;
        const group = document.createElementNS(SVG_NS$1, 'g');
        group.setAttribute('class', 'xlsx-chart-series');
        group.setAttribute('data-series-index', '0');
        if (series.name)
            group.setAttribute('data-series-name', sanitiseForAttr(series.name));
        let acc = 0;
        values.forEach((v, i) => {
            if (v <= 0) {
                const path = el('path', {
                    d: `M ${cx} ${cy} Z`,
                    fill: PALETTE$1[i % PALETTE$1.length],
                });
                group.appendChild(path);
                return;
            }
            const startAngle = (acc / total) * Math.PI * 2 - Math.PI / 2;
            acc += v;
            const endAngle = (acc / total) * Math.PI * 2 - Math.PI / 2;
            const x1 = cx + rOuter * Math.cos(startAngle);
            const y1 = cy + rOuter * Math.sin(startAngle);
            const x2 = cx + rOuter * Math.cos(endAngle);
            const y2 = cy + rOuter * Math.sin(endAngle);
            const large = endAngle - startAngle > Math.PI ? 1 : 0;
            let d;
            if (rInner > 0) {
                const ix2 = cx + rInner * Math.cos(endAngle);
                const iy2 = cy + rInner * Math.sin(endAngle);
                const ix1 = cx + rInner * Math.cos(startAngle);
                const iy1 = cy + rInner * Math.sin(startAngle);
                d = `M ${x1.toFixed(2)} ${y1.toFixed(2)} `
                    + `A ${rOuter} ${rOuter} 0 ${large} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} `
                    + `L ${ix2.toFixed(2)} ${iy2.toFixed(2)} `
                    + `A ${rInner} ${rInner} 0 ${large} 0 ${ix1.toFixed(2)} ${iy1.toFixed(2)} Z`;
            }
            else {
                d = `M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${rOuter} ${rOuter} 0 ${large} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
            }
            const color = PALETTE$1[i % PALETTE$1.length];
            const path = el('path', {
                d,
                fill: color,
                stroke: '#fff',
                'stroke-width': 1,
            });
            group.appendChild(path);
            if (series.dataLabels?.show) {
                const mid = (startAngle + endAngle) / 2;
                const pos = series.dataLabels.position ?? 'ctr';
                const rLabel = pos === 'outEnd'
                    ? rOuter * 1.1
                    : (rInner > 0 ? (rInner + rOuter) / 2 : rOuter * 0.65);
                const lx = cx + rLabel * Math.cos(mid);
                const ly = cy + rLabel * Math.sin(mid) + 3;
                appendDataLabel(group, lx, ly, formatDataLabel(v), 'middle');
            }
        });
        svg.appendChild(group);
        if (model.legend !== 'none' && model.categories.length) {
            const legend = document.createElementNS(SVG_NS$1, 'g');
            legend.setAttribute('class', 'xlsx-chart-legend');
            const x = layout.plotX + layout.plotW + 16;
            const y = layout.plotY + 4;
            const rowH = 16;
            const swatch = 10;
            model.categories.forEach((label, i) => {
                const item = document.createElementNS(SVG_NS$1, 'g');
                item.setAttribute('class', 'xlsx-chart-legend-entry');
                const ly = y + rowH * i;
                item.appendChild(el('rect', {
                    x, y: ly, width: swatch, height: swatch,
                    fill: PALETTE$1[i % PALETTE$1.length],
                }));
                item.appendChild(textEl(x + swatch + 6, ly + swatch - 1, label ?? '', {
                    'font-size': '11',
                }));
                legend.appendChild(item);
            });
            svg.appendChild(legend);
        }
    }
    function renderRadar(model, layout, svg) {
        const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length));
        if (n === 0)
            return;
        const { max } = seriesRange(model.series);
        const axisMax = max > 0 ? niceMax(max) : 1;
        const size = Math.min(layout.plotW, layout.plotH);
        const cx = layout.plotX + layout.plotW / 2;
        const cy = layout.plotY + layout.plotH / 2;
        const rOuter = size / 2 - 20;
        const angleFor = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
        const vertexAt = (i, rFrac) => ({
            x: cx + rOuter * rFrac * Math.cos(angleFor(i)),
            y: cy + rOuter * rFrac * Math.sin(angleFor(i)),
        });
        const gridGroup = document.createElementNS(SVG_NS$1, 'g');
        gridGroup.setAttribute('class', 'xlsx-chart-radar-grid');
        const ticks = [0.25, 0.5, 0.75, 1];
        for (const rFrac of ticks) {
            const pts = [];
            for (let i = 0; i < n; i++) {
                const p = vertexAt(i, rFrac);
                pts.push(`${p.x.toFixed(2)},${p.y.toFixed(2)}`);
            }
            gridGroup.appendChild(el('polygon', {
                points: pts.join(' '),
                fill: 'none',
                stroke: GRID_COLOR,
                'stroke-width': 1,
            }));
        }
        svg.appendChild(gridGroup);
        const axesGroup = document.createElementNS(SVG_NS$1, 'g');
        axesGroup.setAttribute('class', 'xlsx-chart-radar-axes');
        const categoryLabels = model.categories.length ? model.categories
            : Array.from({ length: n }, (_, i) => String(i + 1));
        for (let i = 0; i < n; i++) {
            const outer = vertexAt(i, 1);
            axesGroup.appendChild(el('line', {
                x1: cx, y1: cy,
                x2: outer.x, y2: outer.y,
                stroke: AXIS_COLOR, 'stroke-width': 1,
            }));
            const labelPt = vertexAt(i, 1.1);
            const anchor = Math.abs(labelPt.x - cx) < 1 ? 'middle'
                : (labelPt.x < cx ? 'end' : 'start');
            axesGroup.appendChild(textEl(labelPt.x, labelPt.y + 4, categoryLabels[i] ?? '', {
                'text-anchor': anchor, 'font-size': '10',
            }));
        }
        svg.appendChild(axesGroup);
        model.series.forEach((s, si) => {
            const group = document.createElementNS(SVG_NS$1, 'g');
            group.setAttribute('class', 'xlsx-chart-series');
            group.setAttribute('data-series-index', String(si));
            if (s.name)
                group.setAttribute('data-series-name', sanitiseForAttr(s.name));
            const color = colorFor(s, si);
            const pts = [];
            const vertices = [];
            for (let i = 0; i < n; i++) {
                const vRaw = s.values[i];
                const v = vRaw !== null && vRaw !== undefined && Number.isFinite(vRaw) ? vRaw : 0;
                const rFrac = axisMax > 0 ? Math.max(0, v / axisMax) : 0;
                const p = vertexAt(i, rFrac);
                pts.push(`${p.x.toFixed(2)},${p.y.toFixed(2)}`);
                vertices.push({ x: p.x, y: p.y, v });
            }
            if (pts.length) {
                group.appendChild(el('polygon', {
                    points: pts.join(' '),
                    fill: color,
                    'fill-opacity': 0.3,
                    stroke: color,
                    'stroke-width': 2,
                }));
                for (const vx of vertices) {
                    group.appendChild(el('circle', {
                        cx: vx.x.toFixed(2), cy: vx.y.toFixed(2), r: 3, fill: color,
                    }));
                }
                if (s.dataLabels?.show) {
                    for (let i = 0; i < n; i++) {
                        const vx = vertices[i];
                        const a = angleFor(i);
                        const lx = vx.x + Math.cos(a) * 8;
                        const ly = vx.y + Math.sin(a) * 8 + 3;
                        appendDataLabel(group, lx, ly, formatDataLabel(vx.v), 'middle');
                    }
                }
            }
            svg.appendChild(group);
        });
    }
    function formatDataLabel(v) {
        if (!Number.isFinite(v))
            return '';
        return formatNumber(String(v), 'General').text;
    }
    function appendDataLabel(parent, x, y, text, anchor) {
        const t = textEl(x, y, text, {
            'text-anchor': anchor,
            'font-size': '10',
        });
        t.setAttribute('class', 'xlsx-chart-data-label');
        parent.appendChild(t);
    }
    function sanitiseForAttr(s) {
        let out = '';
        for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 0x20 || c === 0x7F)
                continue;
            out += s[i];
        }
        return out.slice(0, 256);
    }
    function renderChart(model, width = 480, height = 300) {
        if (model.kind === 'unknown')
            return null;
        const svg = makeSvg(width, height);
        const hasLegend = model.legend !== 'none' && model.series.some((s) => s.name);
        const legendPos = model.legend;
        const layout = baseLayout(width, height, hasLegend, legendPos);
        if (model.title)
            appendTitle(svg, model.title, layout);
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
            case 'doughnut':
                renderDoughnut(model, layout, svg);
                return svg;
            case 'scatter':
                renderScatter(model, layout, svg);
                appendLegend(svg, model, layout);
                return svg;
            case 'area':
                renderArea(model, layout, svg);
                appendLegend(svg, model, layout);
                return svg;
            case 'radar':
                renderRadar(model, layout, svg);
                appendLegend(svg, model, layout);
                return svg;
            default:
                return null;
        }
    }

    const SVG_NS = 'http://www.w3.org/2000/svg';
    const DEFAULT_WIDTH = 480;
    const DEFAULT_HEIGHT = 320;
    const NODE_WIDTH = 100;
    const NODE_HEIGHT = 40;
    const NODE_RX = 4;
    const MARGIN_X = 12;
    const MARGIN_Y = 16;
    const CYCLE_NODE_WIDTH = 80;
    const CYCLE_NODE_HEIGHT = 32;
    const CYCLE_NODE_RX = 4;
    const PALETTE = [
        '#5B9BD5',
        '#ED7D31',
        '#A5A5A5',
        '#FFC000',
        '#4472C4',
        '#70AD47',
    ];
    const CONNECTOR_COLOR = '#888';
    const TEXT_COLOR = '#ffffff';
    function renderSmartArtSvg(model, opts = {}) {
        const roots = model.rootNodes;
        if (!roots || roots.length === 0)
            return null;
        const width = opts.width ?? DEFAULT_WIDTH;
        const height = opts.height ?? DEFAULT_HEIGHT;
        const strategy = resolveStrategy(model.layout, opts.layout ?? 'auto');
        if (strategy === 'cycle') {
            return renderCycle(model, width, height);
        }
        return renderHierarchyOrOrgchart(model, width, height, strategy === 'orgchart');
    }
    function resolveStrategy(layoutName, override) {
        if (override === 'hierarchy' || override === 'orgchart' || override === 'cycle') {
            return override;
        }
        if (layoutName) {
            if (/orgChart/i.test(layoutName))
                return 'orgchart';
            if (/cycle/i.test(layoutName))
                return 'cycle';
        }
        return 'hierarchy';
    }
    function renderHierarchyOrOrgchart(model, width, height, orgchart) {
        const levels = [];
        function walk(node, depth, parentIndex) {
            if (!levels[depth])
                levels[depth] = [];
            const placed = { node, parentIndex, x: 0, y: 0, depth };
            levels[depth].push(placed);
            const myIndex = levels[depth].length - 1;
            for (const child of node.children) {
                walk(child, depth + 1, myIndex);
            }
        }
        for (const root of model.rootNodes)
            walk(root, 0, -1);
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
            for (let d = 1; d < depthCount; d++) {
                const row = levels[d];
                const parents = levels[d - 1];
                const byParent = new Map();
                for (const placed of row) {
                    const list = byParent.get(placed.parentIndex) ?? [];
                    list.push(placed);
                    byParent.set(placed.parentIndex, list);
                }
                for (const [parentIndex, group] of byParent) {
                    const parent = parents[parentIndex];
                    if (!parent)
                        continue;
                    const parentBottom = parent.y + NODE_HEIGHT / 2;
                    const childTop = group[0].y - NODE_HEIGHT / 2;
                    const railY = (parentBottom + childTop) / 2;
                    if (group.length === 1) {
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
                    const drop = document.createElementNS(SVG_NS, 'line');
                    drop.setAttribute('x1', String(parent.x));
                    drop.setAttribute('y1', String(parentBottom));
                    drop.setAttribute('x2', String(parent.x));
                    drop.setAttribute('y2', String(railY));
                    drop.setAttribute('stroke', CONNECTOR_COLOR);
                    drop.setAttribute('stroke-width', '1');
                    svg.appendChild(drop);
                    let minX = group[0].x;
                    let maxX = group[0].x;
                    for (const c of group) {
                        if (c.x < minX)
                            minX = c.x;
                        if (c.x > maxX)
                            maxX = c.x;
                    }
                    const rail = document.createElementNS(SVG_NS, 'line');
                    rail.setAttribute('x1', String(minX));
                    rail.setAttribute('y1', String(railY));
                    rail.setAttribute('x2', String(maxX));
                    rail.setAttribute('y2', String(railY));
                    rail.setAttribute('stroke', CONNECTOR_COLOR);
                    rail.setAttribute('stroke-width', '1');
                    svg.appendChild(rail);
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
        }
        else {
            for (let d = 1; d < depthCount; d++) {
                const row = levels[d];
                const parents = levels[d - 1];
                for (const placed of row) {
                    const parent = parents[placed.parentIndex];
                    if (!parent)
                        continue;
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
        for (let d = 0; d < depthCount; d++) {
            const row = levels[d];
            const fill = PALETTE[d % PALETTE.length];
            for (const placed of row) {
                appendNode(svg, placed.x, placed.y, NODE_WIDTH, NODE_HEIGHT, NODE_RX, fill, placed.node.text);
            }
        }
        return svg;
    }
    function renderCycle(model, width, height) {
        const roots = model.rootNodes;
        let nodes = [];
        if (roots.length > 1) {
            nodes = roots;
        }
        else if (roots.length === 1) {
            nodes = roots[0].children.length > 0 ? roots[0].children : roots;
        }
        const svg = createSvgRoot(width, height);
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
        if (nodes.length === 0)
            return svg;
        const cx = width / 2;
        const cy = height / 2;
        const radius = Math.max(40, Math.min(width, height) / 3 - 30);
        const placed = [];
        const n = nodes.length;
        for (let i = 0; i < n; i++) {
            const angle = -Math.PI / 2 + (2 * Math.PI * i) / Math.max(1, n);
            placed.push({
                node: nodes[i],
                x: cx + radius * Math.cos(angle),
                y: cy + radius * Math.sin(angle),
            });
        }
        if (n >= 2) {
            const halo = 18;
            for (let i = 0; i < n; i++) {
                const a = placed[i];
                const b = placed[(i + 1) % n];
                const dax = a.x - cx;
                const day = a.y - cy;
                const la = Math.hypot(dax, day) || 1;
                const ux1 = dax / la;
                const uy1 = day / la;
                const dbx = b.x - cx;
                const dby = b.y - cy;
                const lb = Math.hypot(dbx, dby) || 1;
                const ux2 = dbx / lb;
                const uy2 = dby / lb;
                const arcRadius = radius + halo;
                const sx = cx + ux1 * arcRadius;
                const sy = cy + uy1 * arcRadius;
                const ex = cx + ux2 * arcRadius;
                const ey = cy + uy2 * arcRadius;
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
        const fill = PALETTE[0];
        for (const p of placed) {
            appendNode(svg, p.x, p.y, CYCLE_NODE_WIDTH, CYCLE_NODE_HEIGHT, CYCLE_NODE_RX, fill, p.node.text);
        }
        return svg;
    }
    function createSvgRoot(width, height) {
        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('xmlns', SVG_NS);
        svg.setAttribute('width', String(width));
        svg.setAttribute('height', String(height));
        svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
        svg.setAttribute('class', 'xlsx-smartart-svg');
        return svg;
    }
    function appendNode(svg, cx, cy, w, h, rx, fill, text) {
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
        t.textContent = text;
        g.appendChild(t);
        svg.appendChild(g);
    }

    const PX_PER_CHAR = 7;
    const PADDING_PX = 5;
    function charWidthToPx(width) {
        return Math.round(width * PX_PER_CHAR + PADDING_PX);
    }
    const DEFAULT_COL_WIDTH_CHARS = 8.43;
    const DEFAULT_COL_WIDTH_PX = charWidthToPx(DEFAULT_COL_WIDTH_CHARS);
    const DEFAULT_ROW_HEIGHT_PT = 15;
    const DEFAULT_ROW_HEIGHT_PX = Math.round(DEFAULT_ROW_HEIGHT_PT * 4 / 3);
    const GUTTER_WIDTH_PX = 30;
    class HtmlRenderer {
        async render(workbook, options) {
            const nodes = [];
            const hasEmbeddings = workbook.sheets.some((s) => s.embeddings && s.embeddings.length > 0);
            const hasRenderedCharts = options.renderCharts !== false &&
                workbook.sheets.some((s) => s.charts && s.charts.some((c) => c.model !== null));
            const hasSmartArt = workbook.sheets.some((s) => s.smartArt && s.smartArt.length > 0);
            nodes.push(renderStyle(options.className, {
                withEmbeddings: hasEmbeddings,
                withCharts: hasRenderedCharts,
                withSmartArt: hasSmartArt,
            }));
            for (const sheet of workbook.sheets) {
                if (sheet.state !== 'visible')
                    continue;
                nodes.push(renderSheet(sheet, workbook, options));
            }
            return nodes;
        }
    }
    function renderStyle(className, opts = { withEmbeddings: false }) {
        const style = document.createElement('style');
        style.setAttribute('data-xlsxjs', '');
        const embeddingCss = opts.withEmbeddings ? `
.${className} .xlsx-embedding {
    border: 1px dashed #b0b0b0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #555; font-size: 0.9em;
    background: #fafafa;
}
.${className} .xlsx-embedding > pre {
    margin: 0; white-space: pre-wrap;
    font-family: inherit; font-size: inherit;
}
.${className} .xlsx-image-layer > .xlsx-embedding {
    position: absolute; pointer-events: auto;
    max-width: 320px;
}` : '';
        const chartCss = opts.withCharts ? `
.${className} figure.xlsx-chart {
    margin: 0.5em 0; padding: 0;
}
.${className} figure.xlsx-chart > svg {
    display: block; max-width: 100%;
}` : '';
        const smartArtCss = opts.withSmartArt ? `
.${className} .xlsx-smartart {
    border: 1px dashed #b0b0b0; border-radius: 2px;
    padding: 0.5em; margin: 0.5em 0;
    color: #333; font-size: 0.9em;
    background: #fafbfc;
}
.${className} .xlsx-smartart ul {
    list-style: none; padding-left: 1em; margin: 0;
}
.${className} .xlsx-smartart > ul { padding-left: 0; }
.${className} .xlsx-smartart li { margin: 0.1em 0; }
.${className} .xlsx-smartart .xlsx-smartart-node {
    display: inline-block; padding: 1px 4px;
}
.${className} .xlsx-image-layer > .xlsx-smartart {
    position: absolute; pointer-events: auto;
    max-width: 480px;
}` : '';
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
    position: relative;
}
.${className} .xlsx-shape[data-preset] { border-color: transparent; }
.${className} .xlsx-shape[data-kind="connector"] {
    border-style: dashed; color: #888;
}
.${className} .xlsx-shape[data-kind="connector"][data-preset] { border-style: none; }
.${className} .xlsx-shape > svg {
    position: absolute; inset: 0;
    width: 100%; height: 100%;
    pointer-events: none; z-index: 0;
}
.${className} .xlsx-shape > pre {
    margin: 0; white-space: pre-line;
    font-family: inherit; font-size: inherit;
    position: relative; z-index: 1;
}${embeddingCss}${chartCss}
.${className} .xlsx-form-control {
    position: absolute;
    display: inline-flex; align-items: center; gap: 4px;
    border: 1px dashed #bbb; border-radius: 2px;
    padding: 2px 6px; margin: 0;
    color: #333; background: rgba(255, 255, 255, 0.85);
    font-size: 0.85em; pointer-events: auto;
}
.${className} .xlsx-form-control[data-kind="button"] {
    border-style: solid; background: #f3f3f3;
}
.${className} .xlsx-form-control[data-kind="groupBox"] {
    border-style: solid; background: transparent;
}
.${className} .xlsx-form-control[data-kind="scrollbar"],
.${className} .xlsx-form-control[data-kind="spinner"] {
    background: #eef2f7; color: #555;
}
.${className} .xlsx-form-control > .xlsx-form-control-glyph {
    font-size: 1em; line-height: 1; color: #333;
}
.${className} .xlsx-form-control > legend {
    font-size: 0.85em; color: #333; padding: 0 4px;
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
}${smartArtCss}
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
    function renderSheet(sheet, workbook, options) {
        const styles = workbook.styles;
        const theme = workbook.theme;
        const date1904 = workbook.date1904;
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
        const sheetNameId = `xlsx-sheet-name-${sheet.name.replace(/\W+/g, '-').replace(/^-+|-+$/g, '') || 'unnamed'}`;
        section.appendChild(h('div', { class: 'xlsx-sheet-name', id: sheetNameId }, [sheet.name]));
        const table = h('table');
        table.setAttribute('role', 'table');
        table.setAttribute('aria-labelledby', sheetNameId);
        if (sheet.maxCol < 0) {
            section.appendChild(table);
            return section;
        }
        const colCount = sheet.maxCol + 1;
        const rowCount = sheet.maxRow + 1;
        table.setAttribute('aria-colcount', String(colCount + 1));
        table.setAttribute('aria-rowcount', String(rowCount + 1));
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
        const corner = h('th');
        corner.setAttribute('aria-hidden', 'true');
        headRow.appendChild(corner);
        for (let c = 0; c < colCount; c++) {
            const th = h('th', null, [indexToColumnLetters(c)]);
            th.setAttribute('scope', 'col');
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
            const rowHeader = h('th', null, [String(r + 1)]);
            rowHeader.setAttribute('scope', 'row');
            tr.appendChild(rowHeader);
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
        const anchoredEmbeddings = sheet.embeddings.filter((e) => e.col !== null && e.row !== null);
        const unanchoredEmbeddings = sheet.embeddings.filter((e) => e.col === null || e.row === null);
        const smartArtEntries = sheet.smartArt ?? [];
        if (sheet.images.length > 0 || sheet.formControls.length > 0 || anchoredEmbeddings.length > 0 || smartArtEntries.length > 0) {
            const imageLayer = document.createElement('div');
            imageLayer.className = 'xlsx-image-layer';
            imageLayer.style.position = 'relative';
            imageLayer.style.height = '0';
            imageLayer.style.pointerEvents = 'none';
            for (const img of sheet.images) {
                imageLayer.appendChild(renderImage(img, widthByCol, hiddenCols, rowDim));
            }
            for (const fc of sheet.formControls) {
                imageLayer.appendChild(renderFormControl(fc, widthByCol, hiddenCols, rowDim, options, sheet));
            }
            for (const emb of anchoredEmbeddings) {
                const aside = renderEmbedding(emb);
                const left = GUTTER_WIDTH_PX + sumColsPx(emb.col, widthByCol, hiddenCols);
                const top = sumRowsPx(emb.row, rowDim);
                aside.style.left = `${left}px`;
                aside.style.top = `${top}px`;
                imageLayer.appendChild(aside);
            }
            for (const art of smartArtEntries) {
                const aside = renderSmartArt(art, options);
                const left = GUTTER_WIDTH_PX + sumColsPx(art.col, widthByCol, hiddenCols);
                const top = sumRowsPx(art.row, rowDim);
                aside.style.left = `${left}px`;
                aside.style.top = `${top}px`;
                imageLayer.appendChild(aside);
            }
            section.insertBefore(imageLayer, table);
        }
        for (const chart of sheet.charts) {
            let svg = null;
            if (options.renderCharts !== false && chart.model) {
                try {
                    svg = renderChart(chart.model);
                }
                catch {
                    svg = null;
                }
            }
            if (svg) {
                const figure = document.createElement('figure');
                figure.className = 'xlsx-chart';
                figure.setAttribute('data-chart-kind', chart.kind);
                if (chart.chartType)
                    figure.setAttribute('data-chart-type', chart.chartType);
                if (chart.model?.kind)
                    figure.setAttribute('data-chart-plot', chart.model.kind);
                figure.setAttribute('data-anchor-col', String(chart.col));
                figure.setAttribute('data-anchor-row', String(chart.row));
                if (chart.endCol !== null)
                    figure.setAttribute('data-anchor-end-col', String(chart.endCol));
                if (chart.endRow !== null)
                    figure.setAttribute('data-anchor-end-row', String(chart.endRow));
                figure.appendChild(svg);
                section.appendChild(figure);
            }
            else {
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
        }
        for (const shape of sheet.shapes) {
            section.appendChild(renderShape(shape));
        }
        for (const slicer of sheet.slicers) {
            section.appendChild(renderSlicer(slicer, options));
        }
        for (const timeline of sheet.timelines) {
            section.appendChild(renderTimeline(timeline, options));
        }
        for (const emb of unanchoredEmbeddings) {
            section.appendChild(renderEmbedding(emb));
        }
        if (sheet.headerFooter?.oddHeader) {
            section.appendChild(renderHeaderFooter('xlsx-header', sheet.headerFooter.oddHeader));
        }
        if (sheet.headerFooter?.oddFooter) {
            section.appendChild(renderHeaderFooter('xlsx-footer', sheet.headerFooter.oddFooter));
        }
        return section;
    }
    function columnPx(col, widthByCol, hiddenCols) {
        if (hiddenCols.has(col))
            return 0;
        const w = widthByCol.get(col);
        return w !== undefined ? charWidthToPx(w) : DEFAULT_COL_WIDTH_PX;
    }
    function rowPx(row, rowDim) {
        const d = rowDim.get(row);
        if (d?.hidden)
            return 0;
        if (d?.height != null)
            return Math.round(d.height * 4 / 3);
        return DEFAULT_ROW_HEIGHT_PX;
    }
    function sumColsPx(col, widthByCol, hiddenCols) {
        let total = 0;
        for (let c = 0; c < col; c++)
            total += columnPx(c, widthByCol, hiddenCols);
        return total;
    }
    function sumRowsPx(row, rowDim) {
        let total = 0;
        for (let r = 0; r < row; r++)
            total += rowPx(r, rowDim);
        return total;
    }
    function renderImage(img, widthByCol, hiddenCols, rowDim) {
        const fig = document.createElement('figure');
        fig.className = 'xlsx-image';
        fig.setAttribute('data-anchor-mode', img.anchorMode);
        fig.setAttribute('data-anchor-col', String(img.col));
        fig.setAttribute('data-anchor-row', String(img.row));
        if (img.endCol !== null)
            fig.setAttribute('data-anchor-end-col', String(img.endCol));
        if (img.endRow !== null)
            fig.setAttribute('data-anchor-end-row', String(img.endRow));
        fig.style.margin = '0';
        fig.style.position = 'absolute';
        fig.style.pointerEvents = 'auto';
        if (img.anchorMode === 'absolute') {
            if (img.absoluteX !== null)
                fig.style.left = `${emuToPx(img.absoluteX)}px`;
            if (img.absoluteY !== null)
                fig.style.top = `${emuToPx(img.absoluteY)}px`;
        }
        else {
            const left = GUTTER_WIDTH_PX + sumColsPx(img.col, widthByCol, hiddenCols) + emuToPx(img.colOff);
            const top = sumRowsPx(img.row, rowDim) + emuToPx(img.rowOff);
            fig.style.left = `${left}px`;
            fig.style.top = `${top}px`;
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
        let widthPx = null;
        let heightPx = null;
        if (img.anchorMode === 'twoCell' && img.endCol !== null && img.endRow !== null) {
            const spanW = sumColsPx(img.endCol, widthByCol, hiddenCols) - sumColsPx(img.col, widthByCol, hiddenCols) - emuToPx(img.colOff);
            widthPx = spanW > 0 ? spanW : (img.widthEmu ? emuToPx(img.widthEmu) : null);
            const spanH = sumRowsPx(img.endRow, rowDim) - sumRowsPx(img.row, rowDim) - emuToPx(img.rowOff);
            heightPx = spanH > 0 ? spanH : (img.heightEmu ? emuToPx(img.heightEmu) : null);
        }
        else if (img.widthEmu && img.heightEmu) {
            widthPx = emuToPx(img.widthEmu);
            heightPx = emuToPx(img.heightEmu);
        }
        if (widthPx !== null && heightPx !== null) {
            el.width = widthPx;
            el.height = heightPx;
        }
        el.style.maxWidth = '100%';
        fig.appendChild(el);
        return fig;
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
        const svg = renderShapePreset(shape.preset, { stroke: '#888' });
        if (svg) {
            const tmp = document.createElement('div');
            tmp.innerHTML = svg;
            const svgEl = tmp.firstElementChild;
            if (svgEl)
                aside.appendChild(svgEl);
        }
        if (shape.text && shape.text.length > 0) {
            const pre = document.createElement('pre');
            pre.textContent = shape.text;
            aside.appendChild(pre);
        }
        return aside;
    }
    function renderSmartArt(art, options) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-smartart';
        aside.style.position = 'absolute';
        aside.style.pointerEvents = 'auto';
        if (art.model?.layout)
            aside.setAttribute('data-layout', art.model.layout);
        if (art.name)
            aside.setAttribute('data-name', art.name);
        aside.setAttribute('data-anchor-col', String(art.col));
        aside.setAttribute('data-anchor-row', String(art.row));
        if (art.endCol !== null)
            aside.setAttribute('data-anchor-end-col', String(art.endCol));
        if (art.endRow !== null)
            aside.setAttribute('data-anchor-end-row', String(art.endRow));
        const roots = art.model?.rootNodes ?? [];
        const layout = options.smartArtLayout ?? 'tree';
        let svg = null;
        if (layout !== 'tree' && art.model) {
            try {
                svg = renderSmartArtSvg(art.model);
            }
            catch {
                svg = null;
            }
        }
        if (svg)
            aside.appendChild(svg);
        if (layout === 'tree' || layout === 'both' || !svg) {
            aside.appendChild(renderSmartArtList(roots, 0));
        }
        return aside;
    }
    function renderSmartArtList(nodes, depth) {
        const ul = document.createElement('ul');
        for (const node of nodes) {
            const li = document.createElement('li');
            li.setAttribute('data-level', String(depth));
            const span = document.createElement('span');
            span.className = 'xlsx-smartart-node';
            span.textContent = node.text;
            li.appendChild(span);
            if (node.children.length > 0) {
                li.appendChild(renderSmartArtList(node.children, depth + 1));
            }
            ul.appendChild(li);
        }
        return ul;
    }
    function renderFormControl(fc, widthByCol, hiddenCols, rowDim, options, sheet, workbook) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-form-control';
        aside.setAttribute('data-kind', fc.kind);
        aside.setAttribute('data-anchor-col', String(fc.col));
        aside.setAttribute('data-anchor-row', String(fc.row));
        if (fc.endCol !== null)
            aside.setAttribute('data-anchor-end-col', String(fc.endCol));
        if (fc.endRow !== null)
            aside.setAttribute('data-anchor-end-row', String(fc.endRow));
        if (fc.linkedCell)
            aside.setAttribute('data-linked-cell', fc.linkedCell);
        if (fc.inputRange)
            aside.setAttribute('data-input-range', fc.inputRange);
        if (fc.checked !== null)
            aside.setAttribute('data-checked', String(fc.checked));
        if (fc.min !== null)
            aside.setAttribute('data-min', String(fc.min));
        if (fc.max !== null)
            aside.setAttribute('data-max', String(fc.max));
        if (fc.inc !== null)
            aside.setAttribute('data-inc', String(fc.inc));
        if (fc.page !== null)
            aside.setAttribute('data-page', String(fc.page));
        if (fc.val !== null)
            aside.setAttribute('data-val', String(fc.val));
        if (fc.dropLines !== null)
            aside.setAttribute('data-drop-lines', String(fc.dropLines));
        if (fc.altText)
            aside.setAttribute('aria-label', fc.altText);
        const left = GUTTER_WIDTH_PX + sumColsPx(fc.col, widthByCol, hiddenCols) + emuToPx(fc.colOff);
        const top = sumRowsPx(fc.row, rowDim) + emuToPx(fc.rowOff);
        aside.style.left = `${left}px`;
        aside.style.top = `${top}px`;
        if (options.interactiveFormControls) {
            populateInteractiveFormControl(aside, fc, sheet);
            return aside;
        }
        if (fc.kind === 'checkbox' || fc.kind === 'radio') {
            const glyph = document.createElement('span');
            glyph.className = 'xlsx-form-control-glyph';
            if (fc.kind === 'checkbox') {
                glyph.textContent = fc.checked ? '☑' : '☐';
            }
            else {
                glyph.textContent = fc.checked ? '●' : '○';
            }
            aside.appendChild(glyph);
        }
        if (fc.kind === 'groupBox') {
            const legend = document.createElement('legend');
            if (fc.label)
                legend.textContent = fc.label;
            aside.appendChild(legend);
            return aside;
        }
        if (fc.label) {
            const txt = document.createElement('span');
            txt.className = 'xlsx-form-control-label';
            txt.textContent = fc.label;
            aside.appendChild(txt);
        }
        return aside;
    }
    function populateInteractiveFormControl(aside, fc, sheet, workbook) {
        switch (fc.kind) {
            case 'checkbox': {
                const glyph = document.createElement('span');
                glyph.className = 'xlsx-form-control-glyph';
                glyph.textContent = fc.checked ? '☑' : '☐';
                aside.appendChild(glyph);
                const input = document.createElement('input');
                input.type = 'checkbox';
                input.checked = !!fc.checked;
                input.addEventListener('change', () => {
                    glyph.textContent = input.checked ? '☑' : '☐';
                    if (fc.linkedCell) {
                        const container = resolveRenderRoot(aside);
                        if (container) {
                            applyFormControlUpdate(container, fc.linkedCell, input.checked ? 'TRUE' : 'FALSE');
                        }
                    }
                });
                aside.appendChild(input);
                if (fc.label) {
                    const label = document.createElement('span');
                    label.className = 'xlsx-form-control-label';
                    label.textContent = fc.label;
                    aside.appendChild(label);
                }
                return;
            }
            case 'radio': {
                const glyph = document.createElement('span');
                glyph.className = 'xlsx-form-control-glyph';
                glyph.textContent = fc.checked ? '●' : '○';
                aside.appendChild(glyph);
                const input = document.createElement('input');
                input.type = 'radio';
                input.name = radioGroupName(fc);
                input.checked = !!fc.checked;
                input.addEventListener('change', () => {
                    glyph.textContent = input.checked ? '●' : '○';
                    if (fc.linkedCell && input.checked) {
                        const container = resolveRenderRoot(aside);
                        if (container) {
                            applyFormControlUpdate(container, fc.linkedCell, 'TRUE');
                        }
                    }
                });
                aside.appendChild(input);
                if (fc.label) {
                    const label = document.createElement('span');
                    label.className = 'xlsx-form-control-label';
                    label.textContent = fc.label;
                    aside.appendChild(label);
                }
                return;
            }
            case 'scrollbar': {
                const input = document.createElement('input');
                input.type = 'range';
                if (fc.min !== null)
                    input.min = String(fc.min);
                if (fc.max !== null)
                    input.max = String(fc.max);
                if (fc.inc !== null)
                    input.step = String(fc.inc);
                if (fc.val !== null)
                    input.value = String(fc.val);
                input.addEventListener('input', () => {
                    if (fc.linkedCell) {
                        const container = resolveRenderRoot(aside);
                        if (container) {
                            applyFormControlUpdate(container, fc.linkedCell, input.value);
                        }
                    }
                });
                aside.appendChild(input);
                return;
            }
            case 'spinner': {
                const input = document.createElement('input');
                input.type = 'number';
                if (fc.min !== null)
                    input.min = String(fc.min);
                if (fc.max !== null)
                    input.max = String(fc.max);
                if (fc.inc !== null)
                    input.step = String(fc.inc);
                if (fc.val !== null)
                    input.value = String(fc.val);
                input.addEventListener('input', () => {
                    if (fc.linkedCell) {
                        const container = resolveRenderRoot(aside);
                        if (container) {
                            applyFormControlUpdate(container, fc.linkedCell, input.value);
                        }
                    }
                });
                aside.appendChild(input);
                return;
            }
            case 'combo':
            case 'list': {
                const select = document.createElement('select');
                const optionTexts = fc.inputRange ? resolveInputRangeValues(fc.inputRange, sheet) : [];
                for (const text of optionTexts) {
                    const opt = document.createElement('option');
                    opt.textContent = text;
                    opt.setAttribute('value', text);
                    select.appendChild(opt);
                }
                select.addEventListener('change', () => {
                    if (fc.linkedCell) {
                        const container = resolveRenderRoot(aside);
                        if (container) {
                            applyFormControlUpdate(container, fc.linkedCell, select.value);
                        }
                    }
                });
                aside.appendChild(select);
                return;
            }
            case 'button': {
                const btn = document.createElement('button');
                btn.setAttribute('type', 'button');
                if (fc.label)
                    btn.textContent = fc.label;
                aside.appendChild(btn);
                return;
            }
            case 'groupBox': {
                const legend = document.createElement('legend');
                if (fc.label)
                    legend.textContent = fc.label;
                aside.appendChild(legend);
                return;
            }
            case 'label':
            case 'dialog':
            case 'unknown':
            default: {
                if (fc.label) {
                    const txt = document.createElement('span');
                    txt.className = 'xlsx-form-control-label';
                    txt.textContent = fc.label;
                    aside.appendChild(txt);
                }
                return;
            }
        }
    }
    function radioGroupName(fc, sheet) {
        const rangeOk = fc.inputRange && /^\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?$/.test(fc.inputRange);
        const sheetIndexKey = String(Math.max(0, fc.col));
        const key = rangeOk ? fc.inputRange : `${sheetIndexKey}`;
        const sanitized = key.replace(/[$:]/g, '_');
        return `xlsx-radio-${sanitized}`;
    }
    function resolveInputRangeValues(inputRange, sheet, _workbook) {
        if (inputRange.includes('!')) {
            console.warn(`xlsx-preview: sheet-prefixed form-control inputRange "${inputRange}" not supported; skipping`);
            return [];
        }
        const clean = inputRange.replace(/\$/g, '');
        const m = /^([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?$/.exec(clean);
        if (!m)
            return [];
        const start = parseCellRef(m[1]);
        if (!start)
            return [];
        const end = m[2] ? parseCellRef(m[2]) : start;
        if (!end)
            return [];
        const out = [];
        for (let r = start.row; r <= end.row; r++) {
            for (let c = start.col; c <= end.col; c++) {
                const cell = sheet.rows[r]?.find((x) => x.col === c);
                if (cell && cell.value !== '')
                    out.push(cell.value);
            }
        }
        return out;
    }
    function resolveRenderRoot(el) {
        let cur = el;
        while (cur) {
            if (cur.tagName === 'SECTION' && cur.classList.contains('xlsx'))
                return cur;
            cur = cur.parentElement;
        }
        return null;
    }
    function applyFormControlUpdate(container, linkedCell, newValue) {
        if (linkedCell.includes('!')) {
            console.warn(`xlsx-preview: sheet-prefixed linkedCell "${linkedCell}" not supported; skipping update`);
            return;
        }
        const clean = linkedCell.replace(/\$/g, '');
        const ref = parseCellRef(clean);
        if (!ref)
            return;
        const section = container.classList?.contains('xlsx') && container.tagName === 'SECTION'
            ? container
            : container.querySelector('section.xlsx');
        if (!section)
            return;
        const tbody = section.querySelector('tbody');
        if (!tbody)
            return;
        const tr = tbody.children[ref.row];
        if (!tr)
            return;
        const td = tr.children[ref.col + 1];
        if (!td)
            return;
        td.textContent = newValue;
    }
    function renderSlicer(slicer, options) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-slicer';
        aside.setAttribute('aria-hidden', 'false');
        aside.setAttribute('data-name', slicer.name);
        if (slicer.caption)
            aside.setAttribute('data-caption', slicer.caption);
        if (slicer.sourceName)
            aside.setAttribute('data-source', slicer.sourceName);
        if (slicer.style)
            aside.setAttribute('data-style', slicer.style);
        const header = document.createElement('header');
        header.textContent = slicer.caption ?? slicer.name;
        aside.appendChild(header);
        if (options.interactiveSlicers) {
            populateInteractiveSlicer(aside, slicer);
            return aside;
        }
        const ul = document.createElement('ul');
        for (const item of slicer.selectedItems) {
            const li = document.createElement('li');
            li.textContent = item;
            ul.appendChild(li);
        }
        aside.appendChild(ul);
        return aside;
    }
    function populateInteractiveSlicer(aside, slicer) {
        const itemsSource = slicer.allItems.length > 0 ? slicer.allItems : slicer.selectedItems;
        const selectedSet = new Set(slicer.selectedItems);
        const state = new Map();
        const dispatch = () => {
            const current = [];
            for (const [btn, label] of state) {
                if (btn.getAttribute('aria-pressed') === 'true')
                    current.push(label);
            }
            const win = (aside.ownerDocument?.defaultView ?? (typeof window !== 'undefined' ? window : null));
            const CE = win?.CustomEvent ?? (typeof CustomEvent !== 'undefined' ? CustomEvent : null);
            if (!CE)
                return;
            aside.dispatchEvent(new CE('xlsx:slicer-change', {
                bubbles: true,
                detail: { slicer: slicer.name, selectedItems: current },
            }));
        };
        for (const label of itemsSource) {
            const btn = document.createElement('button');
            btn.setAttribute('type', 'button');
            btn.className = 'xlsx-slicer-chip';
            const pressed = selectedSet.has(label);
            btn.setAttribute('aria-pressed', pressed ? 'true' : 'false');
            btn.textContent = label;
            state.set(btn, label);
            btn.addEventListener('click', () => {
                const nextPressed = btn.getAttribute('aria-pressed') !== 'true';
                btn.setAttribute('aria-pressed', nextPressed ? 'true' : 'false');
                dispatch();
            });
            aside.appendChild(btn);
        }
    }
    function renderTimeline(timeline, options) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-timeline';
        aside.setAttribute('aria-hidden', 'false');
        aside.setAttribute('data-name', timeline.name);
        if (timeline.caption)
            aside.setAttribute('data-caption', timeline.caption);
        if (timeline.sourceName)
            aside.setAttribute('data-source', timeline.sourceName);
        if (timeline.level)
            aside.setAttribute('data-level', timeline.level);
        if (timeline.style)
            aside.setAttribute('data-style', timeline.style);
        const header = document.createElement('header');
        header.textContent = timeline.caption ?? timeline.name;
        aside.appendChild(header);
        const range = timeline.selectedRange;
        const rangeStart = range?.start ?? null;
        const rangeEnd = range?.end ?? null;
        if (options.interactiveSlicers && timeline.bounds) {
            populateInteractiveTimeline(aside, timeline, timeline.bounds, rangeStart, rangeEnd);
            return aside;
        }
        if (rangeStart || rangeEnd) {
            const label = document.createElement('span');
            label.textContent = `${rangeStart ?? ''} → ${rangeEnd ?? ''}`;
            aside.appendChild(label);
        }
        return aside;
    }
    function populateInteractiveTimeline(aside, timeline, bounds, rangeStart, rangeEnd) {
        const minMs = Date.parse(bounds.min);
        const maxMs = Date.parse(bounds.max);
        if (!Number.isFinite(minMs) || !Number.isFinite(maxMs) || maxMs <= minMs) {
            if (rangeStart || rangeEnd) {
                const label = document.createElement('span');
                label.textContent = `${rangeStart ?? ''} → ${rangeEnd ?? ''}`;
                aside.appendChild(label);
            }
            return;
        }
        const startMsRaw = rangeStart ? Date.parse(rangeStart) : minMs;
        const endMsRaw = rangeEnd ? Date.parse(rangeEnd) : maxMs;
        const startMs = Number.isFinite(startMsRaw) ? Math.max(minMs, Math.min(maxMs, startMsRaw)) : minMs;
        const endMs = Number.isFinite(endMsRaw) ? Math.max(minMs, Math.min(maxMs, endMsRaw)) : maxMs;
        const slider = document.createElement('div');
        slider.className = 'xlsx-timeline-slider';
        const minHandle = document.createElement('input');
        minHandle.type = 'range';
        minHandle.setAttribute('data-handle', 'start');
        minHandle.min = String(minMs);
        minHandle.max = String(maxMs);
        minHandle.value = String(startMs);
        const maxHandle = document.createElement('input');
        maxHandle.type = 'range';
        maxHandle.setAttribute('data-handle', 'end');
        maxHandle.min = String(minMs);
        maxHandle.max = String(maxMs);
        maxHandle.value = String(endMs);
        slider.appendChild(minHandle);
        slider.appendChild(maxHandle);
        aside.appendChild(slider);
        const label = document.createElement('span');
        label.className = 'xlsx-timeline-label';
        const fmt = (ms) => new Date(ms).toISOString();
        label.textContent = `${fmt(startMs)} → ${fmt(endMs)}`;
        aside.appendChild(label);
        const onInput = () => {
            let a = Number(minHandle.value);
            let b = Number(maxHandle.value);
            if (!Number.isFinite(a))
                a = minMs;
            if (!Number.isFinite(b))
                b = maxMs;
            const lo = Math.min(a, b);
            const hi = Math.max(a, b);
            label.textContent = `${fmt(lo)} → ${fmt(hi)}`;
            const win = (aside.ownerDocument?.defaultView ?? (typeof window !== 'undefined' ? window : null));
            const CE = win?.CustomEvent ?? (typeof CustomEvent !== 'undefined' ? CustomEvent : null);
            if (!CE)
                return;
            aside.dispatchEvent(new CE('xlsx:timeline-change', {
                bubbles: true,
                detail: { timeline: timeline.name, start: new Date(lo), end: new Date(hi) },
            }));
        };
        minHandle.addEventListener('input', onInput);
        maxHandle.addEventListener('input', onInput);
    }
    function renderEmbedding(emb) {
        const aside = document.createElement('aside');
        aside.className = 'xlsx-embedding';
        aside.setAttribute('data-kind', emb.kind);
        if (emb.progId)
            aside.setAttribute('data-progid', emb.progId);
        if (emb.fileName)
            aside.setAttribute('data-filename', emb.fileName);
        aside.setAttribute('data-content-type', emb.contentType);
        aside.setAttribute('data-size', String(emb.size));
        if (emb.col !== null)
            aside.setAttribute('data-anchor-col', String(emb.col));
        if (emb.row !== null)
            aside.setAttribute('data-anchor-row', String(emb.row));
        if (emb.endCol !== null)
            aside.setAttribute('data-anchor-end-col', String(emb.endCol));
        if (emb.endRow !== null)
            aside.setAttribute('data-anchor-end-row', String(emb.endRow));
        if (emb.altText)
            aside.setAttribute('aria-label', emb.altText);
        const pre = document.createElement('pre');
        const lines = [];
        lines.push(`${emb.kind}: ${emb.fileName ?? '(no filename)'}`);
        lines.push(`type: ${emb.contentType}`);
        lines.push(`size: ${emb.size} bytes`);
        if (emb.progId)
            lines.push(`progId: ${emb.progId}`);
        pre.textContent = lines.join('\n');
        aside.appendChild(pre);
        if (emb.dataUrl) {
            const a = document.createElement('a');
            a.href = emb.dataUrl;
            a.setAttribute('download', emb.fileName ?? 'embedded');
            a.setAttribute('rel', 'noopener');
            a.textContent = `Download ${emb.fileName ?? 'embedded file'}`;
            aside.appendChild(a);
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
        const emptyFormula = cell.formula != null && (cell.value === '' || cell.value == null);
        if ((options.showFormulas && cell.formula != null) || emptyFormula) {
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
        if (!fill)
            return;
        if (fill.kind === 'none')
            return;
        if (fill.kind === 'gradient') {
            const gradCss = gradientFillToCss(fill, theme);
            if (gradCss)
                td.style.backgroundImage = gradCss;
            return;
        }
        if (fill.patternType === 'none' || fill.patternType === 'gray125')
            return;
        const fg = resolveColor(fill.fgColor, theme);
        if (fill.patternType === 'solid') {
            if (fg)
                td.style.backgroundColor = fg;
            return;
        }
        if (!fg)
            return;
        const bg = resolveColor(fill.bgColor, theme) ?? '#ffffff';
        td.style.backgroundColor = bg;
        const pattern = patternFillToCss(fill.patternType, fg);
        if (pattern)
            td.style.backgroundImage = pattern;
    }
    function gradientFillToCss(fill, theme) {
        const resolved = [];
        for (const stop of fill.stops) {
            const hex = resolveColor(stop.color, theme);
            if (!hex)
                continue;
            resolved.push({ position: stop.position, hex });
        }
        if (resolved.length === 0)
            return null;
        if (fill.type === 'path' || resolved.length === 1) {
            return `linear-gradient(${resolved[0].hex}, ${resolved[0].hex})`;
        }
        const stops = resolved
            .map((s) => `${s.hex} ${(s.position * 100).toFixed(2)}%`)
            .join(', ');
        return `linear-gradient(${fill.degree}deg, ${stops})`;
    }
    function patternFillToCss(patternType, fg) {
        const stripe = (angle, onPx, offPx) => `repeating-linear-gradient(${angle}, ${fg} 0 ${onPx}px, transparent ${onPx}px ${onPx + offPx}px)`;
        switch (patternType) {
            case 'darkGray': return stripe('45deg', 2, 2);
            case 'mediumGray': return stripe('45deg', 1, 2);
            case 'lightGray': return stripe('45deg', 1, 4);
            case 'darkHorizontal': return stripe('0deg', 2, 2);
            case 'lightHorizontal': return stripe('0deg', 1, 4);
            case 'darkVertical': return stripe('90deg', 2, 2);
            case 'lightVertical': return stripe('90deg', 1, 4);
            case 'darkDown': return stripe('135deg', 2, 2);
            case 'lightDown': return stripe('135deg', 1, 4);
            case 'darkUp': return stripe('45deg', 2, 2);
            case 'lightUp': return stripe('45deg', 1, 4);
            case 'darkGrid':
                return `${stripe('0deg', 2, 2)}, ${stripe('90deg', 2, 2)}`;
            case 'lightGrid':
                return `${stripe('0deg', 1, 4)}, ${stripe('90deg', 1, 4)}`;
            case 'darkTrellis':
                return `${stripe('45deg', 2, 2)}, ${stripe('135deg', 2, 2)}`;
            case 'lightTrellis':
                return `${stripe('45deg', 1, 4)}, ${stripe('135deg', 1, 4)}`;
            default:
                return null;
        }
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
        if (border.diagonalUp || border.diagonalDown) {
            const diagCss = diagonalBorderToCss(border, theme);
            if (diagCss)
                appendBackgroundImage(td, diagCss);
        }
    }
    function appendBackgroundImage(td, value) {
        const existingStyle = td.getAttribute('style') ?? '';
        const match = /(?:^|;)\s*background-image\s*:\s*([^;]+?)\s*(?=;|$)/i.exec(existingStyle);
        if (match) {
            const merged = `${match[1]}, ${value}`;
            const replaced = existingStyle.slice(0, match.index) +
                (match.index === 0 ? '' : ';') +
                `background-image: ${merged}` +
                existingStyle.slice(match.index + match[0].length);
            td.setAttribute('style', replaced);
            return;
        }
        const separator = existingStyle && !existingStyle.trim().endsWith(';') ? '; ' : '';
        td.setAttribute('style', `${existingStyle}${separator}background-image: ${value};`);
    }
    function diagonalBorderToCss(border, theme) {
        const styleName = border.diagonal.style ?? 'thin';
        const widthPx = Math.max(1, parseInt(borderWidth(styleName), 10) || 1);
        const half = widthPx / 2;
        const color = resolveColor(border.diagonal.color, theme) ?? '#000';
        const gradients = [];
        const band = (direction) => `linear-gradient(${direction}, ` +
            `transparent calc(50% - ${half}px), ` +
            `${color} calc(50% - ${half}px), ` +
            `${color} calc(50% + ${half}px), ` +
            `transparent calc(50% + ${half}px))`;
        if (border.diagonalDown)
            gradients.push(band('to bottom right'));
        if (border.diagonalUp)
            gradients.push(band('to top right'));
        if (gradients.length === 0)
            return null;
        return gradients.join(', ');
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

    const ERROR = Symbol('formula-error');
    const CELL_REF_TOKEN = /^\$?([A-Z]+)\$?([1-9][0-9]*)/;
    function lex(src) {
        const out = [];
        let i = 0;
        while (i < src.length) {
            const c = src[i];
            if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
                i++;
                continue;
            }
            if (c === '"') {
                let j = i + 1;
                let s = '';
                while (j < src.length) {
                    if (src[j] === '"') {
                        if (src[j + 1] === '"') {
                            s += '"';
                            j += 2;
                            continue;
                        }
                        break;
                    }
                    s += src[j];
                    j++;
                }
                out.push({ kind: 'str', text: s });
                i = j + 1;
                continue;
            }
            if ((c >= '0' && c <= '9') || (c === '.' && src[i + 1] >= '0' && src[i + 1] <= '9')) {
                let j = i;
                while (j < src.length && ((src[j] >= '0' && src[j] <= '9') || src[j] === '.'))
                    j++;
                if (j < src.length && (src[j] === 'e' || src[j] === 'E')) {
                    j++;
                    if (src[j] === '+' || src[j] === '-')
                        j++;
                    while (j < src.length && src[j] >= '0' && src[j] <= '9')
                        j++;
                }
                const text = src.slice(i, j);
                out.push({ kind: 'num', text, numeric: Number(text) });
                i = j;
                continue;
            }
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '$') {
                const upper = src.slice(i, i + 24).toUpperCase();
                const m = CELL_REF_TOKEN.exec(upper);
                if (m) {
                    const refText = src.slice(i, i + m[0].length);
                    const col = columnLettersToIndex(m[1]);
                    const row = Number(m[2]) - 1;
                    if (col >= 0) {
                        out.push({ kind: 'ref', text: refText, ref: { col, row } });
                        i += m[0].length;
                        continue;
                    }
                }
                if (c === '$') {
                    i++;
                    continue;
                }
                let j = i;
                while (j < src.length && ((src[j] >= 'A' && src[j] <= 'Z') ||
                    (src[j] >= 'a' && src[j] <= 'z') ||
                    (src[j] >= '0' && src[j] <= '9') ||
                    src[j] === '_' || src[j] === '.'))
                    j++;
                const id = src.slice(i, j).toUpperCase();
                if (id === 'TRUE')
                    out.push({ kind: 'bool', text: 'TRUE' });
                else if (id === 'FALSE')
                    out.push({ kind: 'bool', text: 'FALSE' });
                else
                    out.push({ kind: 'ident', text: id });
                i = j;
                continue;
            }
            if (c === '#') {
                let j = i + 1;
                while (j < src.length && src[j] !== ' ' && src[j] !== ',' && src[j] !== ')' && src[j] !== '(')
                    j++;
                out.push({ kind: 'err', text: src.slice(i, j) });
                i = j;
                continue;
            }
            if (c === '(') {
                out.push({ kind: 'lparen', text: c });
                i++;
                continue;
            }
            if (c === ')') {
                out.push({ kind: 'rparen', text: c });
                i++;
                continue;
            }
            if (c === ',') {
                out.push({ kind: 'comma', text: c });
                i++;
                continue;
            }
            if (c === ':') {
                out.push({ kind: 'colon', text: c });
                i++;
                continue;
            }
            if (c === '+') {
                out.push({ kind: 'plus', text: c });
                i++;
                continue;
            }
            if (c === '-') {
                out.push({ kind: 'minus', text: c });
                i++;
                continue;
            }
            if (c === '*') {
                out.push({ kind: 'star', text: c });
                i++;
                continue;
            }
            if (c === '/') {
                out.push({ kind: 'slash', text: c });
                i++;
                continue;
            }
            if (c === '^') {
                out.push({ kind: 'caret', text: c });
                i++;
                continue;
            }
            if (c === '%') {
                out.push({ kind: 'percent', text: c });
                i++;
                continue;
            }
            if (c === '&') {
                out.push({ kind: 'amp', text: c });
                i++;
                continue;
            }
            if (c === '=') {
                out.push({ kind: 'eq', text: c });
                i++;
                continue;
            }
            if (c === '<') {
                if (src[i + 1] === '=') {
                    out.push({ kind: 'lte', text: '<=' });
                    i += 2;
                    continue;
                }
                if (src[i + 1] === '>') {
                    out.push({ kind: 'neq', text: '<>' });
                    i += 2;
                    continue;
                }
                out.push({ kind: 'lt', text: '<' });
                i++;
                continue;
            }
            if (c === '>') {
                if (src[i + 1] === '=') {
                    out.push({ kind: 'gte', text: '>=' });
                    i += 2;
                    continue;
                }
                out.push({ kind: 'gt', text: '>' });
                i++;
                continue;
            }
            i++;
        }
        out.push({ kind: 'eof', text: '' });
        return out;
    }
    class Parser {
        constructor(toks) {
            this.toks = toks;
            this.pos = 0;
        }
        peek() { return this.toks[this.pos]; }
        eat(kind) {
            if (this.toks[this.pos].kind === kind)
                return this.toks[this.pos++];
            return null;
        }
        expect(kind) {
            const t = this.eat(kind);
            if (!t)
                throw new Error(`expected ${kind}, got ${this.toks[this.pos].kind}`);
            return t;
        }
        parseExpr() { return this.parseCompare(); }
        parseCompare() {
            let left = this.parseConcat();
            while (true) {
                const t = this.peek();
                if (t.kind === 'eq' || t.kind === 'neq' || t.kind === 'lt' ||
                    t.kind === 'lte' || t.kind === 'gt' || t.kind === 'gte') {
                    this.pos++;
                    const right = this.parseConcat();
                    left = { kind: 'bin', op: t.kind, left, right };
                }
                else
                    break;
            }
            return left;
        }
        parseConcat() {
            let left = this.parseAdd();
            while (this.peek().kind === 'amp') {
                this.pos++;
                const right = this.parseAdd();
                left = { kind: 'bin', op: '&', left, right };
            }
            return left;
        }
        parseAdd() {
            let left = this.parseMul();
            while (true) {
                const t = this.peek();
                if (t.kind === 'plus' || t.kind === 'minus') {
                    this.pos++;
                    const right = this.parseMul();
                    left = { kind: 'bin', op: t.kind === 'plus' ? '+' : '-', left, right };
                }
                else
                    break;
            }
            return left;
        }
        parseMul() {
            let left = this.parsePow();
            while (true) {
                const t = this.peek();
                if (t.kind === 'star' || t.kind === 'slash') {
                    this.pos++;
                    const right = this.parsePow();
                    left = { kind: 'bin', op: t.kind === 'star' ? '*' : '/', left, right };
                }
                else
                    break;
            }
            return left;
        }
        parsePow() {
            const left = this.parseUnary();
            if (this.peek().kind === 'caret') {
                this.pos++;
                const right = this.parsePow();
                return { kind: 'bin', op: '^', left, right };
            }
            return left;
        }
        parseUnary() {
            const t = this.peek();
            if (t.kind === 'minus') {
                this.pos++;
                return { kind: 'unary', op: '-', arg: this.parseUnary() };
            }
            if (t.kind === 'plus') {
                this.pos++;
                return { kind: 'unary', op: '+', arg: this.parseUnary() };
            }
            return this.parsePostfix();
        }
        parsePostfix() {
            let arg = this.parsePrimary();
            while (this.peek().kind === 'percent') {
                this.pos++;
                arg = { kind: 'postfix', op: '%', arg };
            }
            return arg;
        }
        parsePrimary() {
            const t = this.peek();
            if (t.kind === 'num') {
                this.pos++;
                return { kind: 'num', value: t.numeric };
            }
            if (t.kind === 'str') {
                this.pos++;
                return { kind: 'str', value: t.text };
            }
            if (t.kind === 'bool') {
                this.pos++;
                return { kind: 'bool', value: t.text === 'TRUE' };
            }
            if (t.kind === 'err') {
                this.pos++;
                return { kind: 'err', value: t.text };
            }
            if (t.kind === 'lparen') {
                this.pos++;
                const inner = this.parseExpr();
                this.expect('rparen');
                return inner;
            }
            if (t.kind === 'ref') {
                this.pos++;
                const { col, row } = t.ref;
                if (this.peek().kind === 'colon') {
                    this.pos++;
                    const t2 = this.expect('ref');
                    const { col: col2, row: row2 } = t2.ref;
                    return {
                        kind: 'range',
                        col1: Math.min(col, col2), row1: Math.min(row, row2),
                        col2: Math.max(col, col2), row2: Math.max(row, row2),
                    };
                }
                return { kind: 'ref', col, row };
            }
            if (t.kind === 'ident') {
                this.pos++;
                if (this.peek().kind === 'lparen') {
                    this.pos++;
                    const args = [];
                    if (this.peek().kind !== 'rparen') {
                        args.push(this.parseExpr());
                        while (this.peek().kind === 'comma') {
                            this.pos++;
                            args.push(this.parseExpr());
                        }
                    }
                    this.expect('rparen');
                    return { kind: 'call', name: t.text, args };
                }
                return { kind: 'err', value: '#NAME?' };
            }
            throw new Error(`unexpected token ${t.kind} (${t.text})`);
        }
    }
    function parseFormula(src) {
        const body = src.startsWith('=') ? src.slice(1) : src;
        const toks = lex(body);
        const p = new Parser(toks);
        const ast = p.parseExpr();
        if (p.peek().kind !== 'eof') {
            throw new Error(`trailing tokens after expr: ${p.peek().kind}`);
        }
        return ast;
    }
    function toNumber(v) {
        if (v === ERROR)
            return ERROR;
        if (v === null)
            return 0;
        if (typeof v === 'number')
            return v;
        if (typeof v === 'boolean')
            return v ? 1 : 0;
        if (typeof v === 'string') {
            if (v === '')
                return 0;
            if (v.startsWith('#'))
                return ERROR;
            const n = Number(v);
            if (Number.isNaN(n))
                return ERROR;
            return n;
        }
        return ERROR;
    }
    function toBool(v) {
        if (v === ERROR)
            return ERROR;
        if (v === null)
            return false;
        if (typeof v === 'boolean')
            return v;
        if (typeof v === 'number')
            return v !== 0;
        if (typeof v === 'string') {
            const u = v.toUpperCase();
            if (u === 'TRUE')
                return true;
            if (u === 'FALSE')
                return false;
            if (v.startsWith('#'))
                return ERROR;
            const n = Number(v);
            if (Number.isNaN(n))
                return ERROR;
            return n !== 0;
        }
        return ERROR;
    }
    function flatten(v) {
        if (v === ERROR)
            return [ERROR];
        if (Array.isArray(v)) {
            const out = [];
            for (const x of v)
                for (const y of flatten(x))
                    out.push(y);
            return out;
        }
        return [v];
    }
    const FUNCTIONS = {
        SUM: (args) => {
            let total = 0;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    if (typeof v === 'string') {
                        if (v === '')
                            continue;
                        const n = Number(v);
                        if (Number.isNaN(n))
                            continue;
                        total += n;
                        continue;
                    }
                    if (typeof v === 'boolean') {
                        total += v ? 1 : 0;
                        continue;
                    }
                    total += v;
                }
            }
            return total;
        },
        AVERAGE: (args) => {
            let total = 0;
            let n = 0;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    if (typeof v === 'string') {
                        if (v === '')
                            continue;
                        const num = Number(v);
                        if (Number.isNaN(num))
                            continue;
                        total += num;
                        n++;
                        continue;
                    }
                    if (typeof v === 'boolean') {
                        total += v ? 1 : 0;
                        n++;
                        continue;
                    }
                    total += v;
                    n++;
                }
            }
            if (n === 0)
                return ERROR;
            return total / n;
        },
        MIN: (args) => {
            let m = null;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    let num;
                    if (typeof v === 'number')
                        num = v;
                    else if (typeof v === 'boolean')
                        num = v ? 1 : 0;
                    else if (typeof v === 'string') {
                        if (v === '')
                            continue;
                        const parsed = Number(v);
                        if (Number.isNaN(parsed))
                            continue;
                        num = parsed;
                    }
                    else
                        continue;
                    m = m === null ? num : Math.min(m, num);
                }
            }
            return m === null ? 0 : m;
        },
        MAX: (args) => {
            let m = null;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    let num;
                    if (typeof v === 'number')
                        num = v;
                    else if (typeof v === 'boolean')
                        num = v ? 1 : 0;
                    else if (typeof v === 'string') {
                        if (v === '')
                            continue;
                        const parsed = Number(v);
                        if (Number.isNaN(parsed))
                            continue;
                        num = parsed;
                    }
                    else
                        continue;
                    m = m === null ? num : Math.max(m, num);
                }
            }
            return m === null ? 0 : m;
        },
        COUNT: (args) => {
            let n = 0;
            for (const a of args) {
                const flat = flatten(a);
                const fromRange = Array.isArray(a);
                for (const v of flat) {
                    if (v === ERROR)
                        continue;
                    if (typeof v === 'number')
                        n++;
                    else if (!fromRange && typeof v === 'string' && v !== '') {
                        const parsed = Number(v);
                        if (!Number.isNaN(parsed))
                            n++;
                    }
                }
            }
            return n;
        },
        COUNTA: (args) => {
            let n = 0;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR) {
                        n++;
                        continue;
                    }
                    if (v === null)
                        continue;
                    if (typeof v === 'string' && v === '')
                        continue;
                    n++;
                }
            }
            return n;
        },
        IF: (args) => {
            if (args.length < 2)
                return ERROR;
            const cond = toBool(args[0]);
            if (cond === ERROR)
                return ERROR;
            if (cond)
                return args[1];
            return args.length >= 3 ? args[2] : false;
        },
        AND: (args) => {
            let seen = false;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    const b = toBool(v);
                    if (b === ERROR)
                        return ERROR;
                    if (!b)
                        return false;
                    seen = true;
                }
            }
            return seen;
        },
        OR: (args) => {
            let seen = false;
            for (const a of args) {
                for (const v of flatten(a)) {
                    if (v === ERROR)
                        return ERROR;
                    if (v === null)
                        continue;
                    const b = toBool(v);
                    if (b === ERROR)
                        return ERROR;
                    if (b)
                        return true;
                    seen = true;
                }
            }
            return seen ? false : ERROR;
        },
        NOT: (args) => {
            if (args.length !== 1)
                return ERROR;
            const b = toBool(args[0]);
            if (b === ERROR)
                return ERROR;
            return !b;
        },
    };
    function cmp(a, b) {
        if (a === null)
            a = 0;
        if (b === null)
            b = 0;
        if (typeof a === 'number' && typeof b === 'number')
            return a < b ? -1 : a > b ? 1 : 0;
        if (typeof a === 'string' && typeof b === 'string') {
            const la = a.toLowerCase();
            const lb = b.toLowerCase();
            return la < lb ? -1 : la > lb ? 1 : 0;
        }
        if (typeof a === 'boolean' && typeof b === 'boolean')
            return a === b ? 0 : a ? 1 : -1;
        const rank = (v) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
        const ra = rank(a);
        const rb = rank(b);
        return ra < rb ? -1 : ra > rb ? 1 : 0;
    }
    function evalAst(ast, resolver, depth = 0) {
        if (depth > 1024)
            return ERROR;
        switch (ast.kind) {
            case 'num': return ast.value;
            case 'str': return ast.value;
            case 'bool': return ast.value;
            case 'err': return ERROR;
            case 'ref': {
                const v = resolver(ast.col, ast.row);
                if (typeof v === 'string' && v.startsWith('#'))
                    return ERROR;
                return v;
            }
            case 'range': {
                const out = [];
                for (let r = ast.row1; r <= ast.row2; r++) {
                    for (let c = ast.col1; c <= ast.col2; c++) {
                        const v = resolver(c, r);
                        if (typeof v === 'string' && v.startsWith('#'))
                            return ERROR;
                        out.push(v);
                    }
                }
                return out;
            }
            case 'unary': {
                const v = evalAst(ast.arg, resolver, depth + 1);
                if (v === ERROR)
                    return ERROR;
                const n = toNumber(v);
                if (n === ERROR)
                    return ERROR;
                return ast.op === '-' ? -n : n;
            }
            case 'postfix': {
                const v = evalAst(ast.arg, resolver, depth + 1);
                if (v === ERROR)
                    return ERROR;
                const n = toNumber(v);
                if (n === ERROR)
                    return ERROR;
                return n / 100;
            }
            case 'bin': {
                const l = evalAst(ast.left, resolver, depth + 1);
                const r = evalAst(ast.right, resolver, depth + 1);
                if (l === ERROR || r === ERROR)
                    return ERROR;
                switch (ast.op) {
                    case '+':
                    case '-':
                    case '*':
                    case '/':
                    case '^': {
                        const ln = toNumber(l);
                        const rn = toNumber(r);
                        if (ln === ERROR || rn === ERROR)
                            return ERROR;
                        switch (ast.op) {
                            case '+': return ln + rn;
                            case '-': return ln - rn;
                            case '*': return ln * rn;
                            case '/': return rn === 0 ? ERROR : ln / rn;
                            case '^': return Math.pow(ln, rn);
                        }
                        return ERROR;
                    }
                    case '&': {
                        const stringify = (v) => {
                            if (v === null)
                                return '';
                            if (typeof v === 'boolean')
                                return v ? 'TRUE' : 'FALSE';
                            if (Array.isArray(v))
                                return stringify(v.length > 0 ? v[0] : null);
                            return String(v);
                        };
                        return stringify(l) + stringify(r);
                    }
                    case 'eq':
                    case 'neq':
                    case 'lt':
                    case 'lte':
                    case 'gt':
                    case 'gte': {
                        const c = cmp(l, r);
                        if (c === ERROR)
                            return ERROR;
                        switch (ast.op) {
                            case 'eq': return c === 0;
                            case 'neq': return c !== 0;
                            case 'lt': return c < 0;
                            case 'lte': return c <= 0;
                            case 'gt': return c > 0;
                            case 'gte': return c >= 0;
                        }
                        return ERROR;
                    }
                }
                return ERROR;
            }
            case 'call': {
                const fn = FUNCTIONS[ast.name];
                if (!fn)
                    return ERROR;
                const args = ast.args.map((a) => evalAst(a, resolver, depth + 1));
                return fn(args);
            }
        }
    }
    function evaluateFormula(source, resolver) {
        let ast;
        try {
            ast = parseFormula(source);
        }
        catch {
            return { value: '#ERROR!', kind: 'error' };
        }
        const result = evalAst(ast, resolver);
        if (result === ERROR)
            return { value: '#ERROR!', kind: 'error' };
        if (Array.isArray(result)) {
            const first = result.length > 0 ? result[0] : null;
            return formatScalar(first);
        }
        return formatScalar(result);
    }
    function formatScalar(v) {
        if (v === null)
            return { value: '', kind: 'string' };
        if (typeof v === 'number') {
            if (!Number.isFinite(v))
                return { value: '#ERROR!', kind: 'error' };
            return { value: String(v), kind: 'number' };
        }
        if (typeof v === 'boolean')
            return { value: v ? 'TRUE' : 'FALSE', kind: 'boolean' };
        if (typeof v === 'string') {
            if (v.startsWith('#'))
                return { value: v, kind: 'error' };
            return { value: v, kind: 'string' };
        }
        return { value: '#ERROR!', kind: 'error' };
    }
    function makeSheetResolver(sheet) {
        return (col, row) => {
            const rowArr = sheet.rows[row];
            if (!rowArr)
                return null;
            const cell = rowArr[col];
            if (!cell)
                return null;
            if (cell.kind === 'empty')
                return null;
            if (cell.kind === 'number') {
                if (cell.value === '')
                    return null;
                const n = Number(cell.value);
                return Number.isNaN(n) ? null : n;
            }
            if (cell.kind === 'boolean')
                return cell.value === 'TRUE' || cell.value === '1' || cell.value.toLowerCase() === 'true';
            if (cell.kind === 'error')
                return cell.value || '#ERROR!';
            return cell.value;
        };
    }
    function evaluateSheetFormulas(sheet, opts) {
        const force = opts?.force === true;
        const resolver = makeSheetResolver(sheet);
        const updates = [];
        for (const row of sheet.rows) {
            if (!row)
                continue;
            for (const cell of row) {
                if (!cell || !cell.formula)
                    continue;
                if (!force && cell.value !== '' && cell.value != null)
                    continue;
                const r = evaluateFormula(cell.formula, resolver);
                updates.push({ cell, value: r.value, kind: r.kind });
            }
        }
        for (const u of updates) {
            u.cell.value = u.value;
            u.cell.kind = u.kind;
        }
    }

    const defaultOptions = {
        className: 'xlsx',
        inWrapper: true,
        debug: false,
        showFormulas: false,
        formulaNotation: 'a1',
        inlineEmbeddings: false,
        parseOleCfb: false,
        renderCharts: true,
        interactiveFormControls: false,
        interactiveSlicers: false,
        smartArtLayout: 'tree',
        evaluateFormulas: false,
        evaluateFormulasForce: false,
        h,
    };
    function mergeOptions(userOptions) {
        return { ...defaultOptions, ...userOptions };
    }
    async function parseAsync(data, userOptions) {
        const ops = mergeOptions(userOptions);
        const wb = await Workbook.load(data, new WorkbookParser(ops));
        if (ops.evaluateFormulas && wb.parsed) {
            for (const sheet of wb.parsed.sheets) {
                evaluateSheetFormulas(sheet, { force: ops.evaluateFormulasForce });
            }
        }
        return wb;
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

    exports.MAX_EMBEDDING_BYTES = MAX_EMBEDDING_BYTES;
    exports.XlsxEncryptedError = XlsxEncryptedError;
    exports.a1ToR1c1 = a1ToR1c1;
    exports.applyFormControlUpdate = applyFormControlUpdate;
    exports.applyTint = applyTint;
    exports.bytesToDataUrl = bytesToDataUrl;
    exports.defaultOptions = defaultOptions;
    exports.emuToPx = emuToPx;
    exports.evalAst = evalAst;
    exports.evaluateFormula = evaluateFormula;
    exports.evaluateRule = evaluateRule;
    exports.evaluateSheetFormulas = evaluateSheetFormulas;
    exports.formatNumber = formatNumber;
    exports.indexedColor = indexedColor;
    exports.interpolateColorScale = interpolateColorScale;
    exports.isSafeHyperlinkHref = isSafeHyperlinkHref;
    exports.lookupNumberFormat = lookupNumberFormat;
    exports.makeSheetResolver = makeSheetResolver;
    exports.parseAsync = parseAsync;
    exports.parseCfb = parseCfb;
    exports.parseChart = parseChart;
    exports.parseColorElement = parseColorElement;
    exports.parseConditionalFormatting = parseConditionalFormatting;
    exports.parseFormula = parseFormula;
    exports.parseSmartArt = parseSmartArt;
    exports.parseStyles = parseStyles;
    exports.parseTheme = parseTheme;
    exports.parseThreadedComments = parseThreadedComments;
    exports.r1c1ToA1 = r1c1ToA1;
    exports.renderAsync = renderAsync;
    exports.renderChart = renderChart;
    exports.renderSmartArtSvg = renderSmartArtSvg;
    exports.renderWorkbook = renderWorkbook;
    exports.resolveCfvo = resolveCfvo;
    exports.resolveColor = resolveColor;
    exports.resolveEffectiveXf = resolveEffectiveXf;
    exports.sanitizeFontFamily = sanitizeFontFamily;
    exports.sanitizeHexColor = sanitizeHexColor;
    exports.sanitizeMediaMime = sanitizeMediaMime;

}));
//# sourceMappingURL=xlsx-preview.js.map
