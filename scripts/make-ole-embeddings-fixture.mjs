// Build tests/render-test/ole-embeddings/workbook.xlsx. Exercises the
// detect-only OLE / package embedding path introduced in Wave 8:
//
//   · Sheet "Embeds" carries two <oleObject> entries in its <oleObjects>
//     block:
//       - r:id=rId2, progId="Word.Document.12", anchored at B2:D4,
//         targeting xl/embeddings/oleObject1.bin (raw OLE CFB payload →
//         SheetEmbedding kind='ole').
//       - r:id=rId3, progId="Package",          anchored at B6:D8,
//         targeting xl/embeddings/embedded.txt (a packaged generic file →
//         SheetEmbedding kind='package').
//
//   · The OLE .bin payload is NOT a real CFB container — xlsxjs is detect-
//     only, so byte-level validity doesn't matter. We write a plausible
//     header (OLE CFB magic 0xD0 0xCF 0x11 0xE0 …) padded with zeros to a
//     small size so the SheetEmbedding.size assertion has a stable target.
//
// Run:
//   node scripts/make-ole-embeddings-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/ole-embeddings');
mkdirSync(outDir, { recursive: true });

// Minimum-plausible OLE CFB header: magic + enough zero padding so the byte
// count is stable (= 512, one header sector). Attackers can't exercise CFB
// parsing through xlsxjs — we never decode these bytes — but the header
// keeps the file recognisable to external tools.
const OLE_BIN = (() => {
    const buf = new Uint8Array(512);
    buf.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    return buf;
})();

// Packaged real file: plain text. Arbitrary body so the
// contentType (text/plain) is distinguishable from the OLE .bin payload.
const EMBEDDED_TXT = new TextEncoder().encode('hello from an embedded package\n');

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="txt" ContentType="text/plain"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Embeds" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>OLE Word:</t></si>
  <si><t>Package text:</t></si>
</sst>`;

// Sheet xml carries a label in A2 / A6 next to each embedding anchor cell,
// plus the <oleObjects> block at the sheet root. The block uses the
// xdr-namespace anchor shape (same <xdr:from>/<xdr:to>/<xdr:col>/<xdr:row>
// markers the drawing code already groks) so our parser picks up col/row
// without a separate code path.
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <sheetData>
    <row r="2"><c r="A2" t="s"><v>0</v></c></row>
    <row r="6"><c r="A6" t="s"><v>1</v></c></row>
  </sheetData>
  <oleObjects>
    <oleObject progId="Word.Document.12" shapeId="1025" r:id="rId2">
      <objectPr defaultSize="0" r:id="rId10">
        <anchor>
          <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        </anchor>
      </objectPr>
    </oleObject>
    <oleObject progId="Package" shapeId="1026" r:id="rId3">
      <objectPr defaultSize="0" r:id="rId11">
        <anchor>
          <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        </anchor>
      </objectPr>
    </oleObject>
  </oleObjects>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embedded.txt"/>
</Relationships>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/sharedStrings.xml', sharedStringsXml);
zip.file('xl/worksheets/sheet1.xml', sheetXml);
zip.file('xl/worksheets/_rels/sheet1.xml.rels', sheetRels);
zip.file('xl/embeddings/oleObject1.bin', OLE_BIN);
zip.file('xl/embeddings/embedded.txt', EMBEDDED_TXT);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
