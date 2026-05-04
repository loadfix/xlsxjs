// Workbook: loads an XLSX package with JSZip, extracts the sheet-relevant
// XML parts, and hands them to WorkbookParser. Mirrors the split in docxjs
// between WordDocument (zip/relationships) and DocumentParser (XML→tree).

import { WorkbookParser, Workbook as ParsedWorkbook } from './workbook-parser';

declare const JSZip: any;

export class Workbook {
    parts: Record<string, string> = {};
    // Binary media parts (xl/media/*). Kept as data: URLs so the renderer
    // can embed them without a separate fetch.
    media: Record<string, string> = {};
    parsed: ParsedWorkbook | null = null;

    static async load(data: Blob | ArrayBuffer | Uint8Array, parser: WorkbookParser): Promise<Workbook> {
        const wb = new Workbook();
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

        // xl/drawings/drawingN.xml + their rels. Drawings anchor images (and
        // charts — charts are out of scope) to cell positions on a sheet.
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/drawings\/.*\.xml$/i.test(p) ||
                /^xl\/drawings\/_rels\/.*\.xml\.rels$/i.test(p)) {
                const xml = await readIfPresent(p);
                if (xml) wb.parts[p] = xml;
            }
        }

        // Media binaries — embed as data: URLs so the renderer can use them
        // without any runtime fetch. Kept separate from `parts` (strings).
        for (const p of Object.keys(zip.files)) {
            if (/^xl\/media\/[^/]+$/i.test(p)) {
                const bin = zip.file(p);
                if (!bin) continue;
                const base64 = await bin.async('base64');
                const mime = guessMime(p);
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
