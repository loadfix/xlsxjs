// Workbook: loads an XLSX package with JSZip, extracts the sheet-relevant
// XML parts, and hands them to WorkbookParser. Mirrors the split in docxjs
// between WordDocument (zip/relationships) and DocumentParser (XML→tree).

import { WorkbookParser, Workbook as ParsedWorkbook } from './workbook-parser';

declare const JSZip: any;

// Password-protected .xlsx files are OLE Compound File Binary containers,
// not zips. Detecting this up front lets callers distinguish "bad zip" from
// "encrypted" without having to interpret JSZip's error text.
export class XlsxEncryptedError extends Error {
    constructor() {
        super('xlsx-preview: this file is encrypted (OLE CFB container). xlsxjs does not decrypt — remove the password in Excel and re-save.');
        this.name = 'XlsxEncryptedError';
    }
}

const OLE_CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];

// DoS guardrail: cap a single inlined media file at 32 MiB. Anything larger
// is skipped (not inlined, not surfaced on `wb.media`). A normal embedded
// image in an .xlsx is well under a megabyte; the cap is generous enough to
// cover legitimate photography but stops a crafted 1 GiB "image" from
// exhausting the tab via base64 expansion (+33 % overhead) and DOM allocation.
// The chosen sinks are `<img src=data:…>` so the decoded bytes land in the
// renderer process too.
const MAX_MEDIA_BYTES = 32 * 1024 * 1024;

async function bytesOf(data: Blob | ArrayBuffer | Uint8Array): Promise<Uint8Array | null> {
    if (data instanceof Uint8Array) return data;
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    // Blob → ArrayBuffer. Feature-detect because Node Buffer / jsdom Blob
    // both satisfy the interface but via different paths.
    if (typeof (data as Blob)?.arrayBuffer === 'function') {
        return new Uint8Array(await (data as Blob).arrayBuffer());
    }
    return null;
}

function isOleCfb(bytes: Uint8Array | null): boolean {
    if (!bytes || bytes.length < 4) return false;
    return OLE_CFB_MAGIC.every((b, i) => bytes[i] === b);
}

export class Workbook {
    parts: Record<string, string> = {};
    // Binary media parts (xl/media/*). Kept as data: URLs so the renderer
    // can embed them without a separate fetch.
    media: Record<string, string> = {};
    parsed: ParsedWorkbook | null = null;

    static async load(data: Blob | ArrayBuffer | Uint8Array, parser: WorkbookParser): Promise<Workbook> {
        const wb = new Workbook();
        // Cheap pre-flight: detect OLE CFB magic before JSZip chokes on it.
        const bytes = await bytesOf(data);
        if (isOleCfb(bytes)) throw new XlsxEncryptedError();
        const zip = await JSZip.loadAsync(data);

        const readIfPresent = async (path: string): Promise<string | null> => {
            const f = zip.file(path);
            return f ? await f.async('string') : null;
        };

        const workbookXml = await readIfPresent('xl/workbook.xml');
        if (!workbookXml) throw new Error('xlsx-preview: missing xl/workbook.xml');
        wb.parts['xl/workbook.xml'] = workbookXml;

        // Workbook relationships — xl/_rels/workbook.xml.rels. The parser uses
        // this to bind <sheet r:id="rIdN"/> to the actual sheet xml path,
        // rather than assuming xl/worksheets/sheet{N}.xml.
        const workbookRels = await readIfPresent('xl/_rels/workbook.xml.rels');
        if (workbookRels) wb.parts['xl/_rels/workbook.xml.rels'] = workbookRels;

        const sharedStrings = await readIfPresent('xl/sharedStrings.xml');
        if (sharedStrings) wb.parts['xl/sharedStrings.xml'] = sharedStrings;

        const styles = await readIfPresent('xl/styles.xml');
        if (styles) wb.parts['xl/styles.xml'] = styles;

        // xl/metadata.xml — Excel 365 cell metadata. Describes dynamic-array
        // spill anchors (via XLDAPR type + fDynamic flag) and rich data types;
        // individual cells point into the <cellMetadata>/<valueMetadata>
        // blocks via their c/@cm and c/@vm attributes.
        const metadata = await readIfPresent('xl/metadata.xml');
        if (metadata) wb.parts['xl/metadata.xml'] = metadata;

        // Theme. Excel / LibreOffice / python-xlsx all write xl/theme/theme1.xml,
        // but the spec allows multiple theme parts — keep whatever we find.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/theme\/theme\d+\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // Every worksheet part. We used to scan xl/worksheets/sheet{N}.xml only;
        // now we keep any xml under xl/worksheets/ so we can resolve by path
        // from the workbook rels.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/worksheets\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // Per-sheet rels (xl/worksheets/_rels/sheetN.xml.rels) bind tables,
        // drawings, and comments to the sheet. Keep them all — the parser
        // resolves whichever it needs.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/worksheets\/_rels\/.*\.xml\.rels$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/tables/tableN.xml — referenced via sheet rels. Keep them all
        // so the parser can resolve them on demand.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/tables\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/drawings/drawingN.xml + their rels. Drawings anchor images
        // and charts to cell positions on a sheet. We also keep
        // vmlDrawingN.vml here — it carries the comment-bubble layout for
        // classic comments — but xlsxjs doesn't parse VML (comments render
        // as inline markers, not floating boxes).
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/drawings\/.*\.xml$/i.test(p) ||
                /^xl\/drawings\/_rels\/.*\.xml\.rels$/i.test(p) ||
                /^xl\/drawings\/.*\.vml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/ctrlProps/ctrlProp*.xml — form-control property parts. Each
        // `<xdr:sp>` wrapped in `<mc:AlternateContent>` for a form control
        // carries a `<xdr:clientData/>` and a rel (type .../ctrlProp) pointing
        // at the matching xml under xl/ctrlProps/. xlsxjs surfaces these on
        // the sheet model as `Sheet.formControls[]` but renders only an
        // informational `<aside>` — no interactive widget is painted.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/ctrlProps\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/commentsN.xml — classic (non-threaded) comments. Bound to a
        // sheet via its rels; the parser resolves the mapping. Threaded
        // comments (xl/threadedComments/*) are a separate part type that
        // xlsxjs does not yet surface.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/comments\d*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/charts/*.xml — both classic (c:chartSpace) and chartEx
        // (cx:chartSpace). The renderer does not draw the chart content; the
        // parser peeks at the chart XML to surface a chart-type name and the
        // loader exposes the raw XML for any consumers that want it.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/charts\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // Pivot tables, slicers, timelines, and their caches. xlsxjs does not
        // re-compute pivot values (the sheet xml already carries them), but we
        // surface the anchored range + name on the model so consumers can
        // render their own affordance, and so the smoke tool stops flagging
        // these parts as "dropped".
        //
        // pivotCacheRecords*.xml are intentionally skipped — they're the
        // materialised records and can be large.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/pivotTables\/.*\.xml$/i.test(p) ||
                /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(p) ||
                /^xl\/slicers\/.*\.xml$/i.test(p) ||
                /^xl\/slicerCaches\/.*\.xml$/i.test(p) ||
                /^xl\/timelines\/.*\.xml$/i.test(p) ||
                /^xl\/timelineCaches\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/threadedComments/threadedCommentN.xml — Excel 365's modern
        // comment threads, bound to sheets via the sheet rels. The author
        // registry is packaged separately under xl/persons/.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/threadedComments\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // xl/persons/person.xml — workbook-wide author registry referenced
        // from threaded comments via personId GUID.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/persons\/.*\.xml$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // Media binaries — embed as data: URLs so the renderer can use them
        // without any runtime fetch. Kept separate from `parts` (strings).
        // Oversized media (>MAX_MEDIA_BYTES) is skipped rather than inlined;
        // drawings that reference it will fail to resolve a dataUrl and the
        // image entry is quietly dropped. This caps attacker-controlled
        // memory pressure at the load boundary.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/media\/[^/]+$/i.test(p)) {
                const bin = zip.file(p);
                if (!bin) continue;
                // Size-check before base64 expansion. The base64 projection
                // is ~1.33x the raw byte length, so capping on raw size keeps
                // the resulting data URL bounded. Oversized media is skipped
                // rather than truncated so a half-image doesn't render.
                const buf = await bin.async('uint8array');
                if (buf.byteLength > MAX_MEDIA_BYTES) continue;
                const base64 = bytesToBase64(buf);
                const mime = sanitizeMediaMime(guessMime(p));
                wb.media[p] = `data:${mime};base64,${base64}`;
            }
        }

        wb.parsed = parser.parse(wb.parts, wb.media);
        return wb;
    }
}

function guessMime(path: string): string {
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    switch (ext) {
        case 'png':  return 'image/png';
        case 'jpg':
        case 'jpeg': return 'image/jpeg';
        case 'gif':  return 'image/gif';
        case 'svg':  return 'image/svg+xml';
        case 'webp': return 'image/webp';
        case 'bmp':  return 'image/bmp';
        case 'tif':
        case 'tiff': return 'image/tiff';
        default: return 'application/octet-stream';
    }
}

// Strict allowlist of MIME types we'll emit on a `data:` URL. The `<img>`
// element treats SVG as a passive image (no scripts, no external refs), so
// the whole allowlist is safe for that sink — but any MIME outside the list
// decays to `application/octet-stream`, which renders as a broken image
// rather than a potentially active resource. Belt-and-braces: guessMime()
// already produces these strings, but a future refactor that forwards a
// path-derived MIME must stay within the set.
const MEDIA_MIME_ALLOWLIST = new Set<string>([
    'image/png',
    'image/jpeg',
    'image/gif',
    'image/svg+xml',
    'image/webp',
    'image/bmp',
    'image/tiff',
]);

function sanitizeMediaMime(mime: string): string {
    return MEDIA_MIME_ALLOWLIST.has(mime) ? mime : 'application/octet-stream';
}

// Convert a Uint8Array to base64. We can't rely on Buffer (Node-only) or
// FileReader (browser-only) and need a single path that works in both jsdom
// and real browsers. `btoa` is on both; we feed it a binary-string chunked
// so we don't hit the V8 argument-count limit on very large inputs.
function bytesToBase64(bytes: Uint8Array): string {
    // atob/btoa round-trip through the "binary string" representation.
    const CHUNK = 0x8000;
    let binary = '';
    for (let i = 0; i < bytes.length; i += CHUNK) {
        const slice = bytes.subarray(i, Math.min(i + CHUNK, bytes.length));
        binary += String.fromCharCode.apply(null, slice as unknown as number[]);
    }
    return btoa(binary);
}
