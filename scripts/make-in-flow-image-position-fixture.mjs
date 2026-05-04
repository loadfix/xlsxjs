// Build tests/render-test/in-flow-image-position/workbook.xlsx. Targets the
// image-overlay layout in html-renderer.ts: a tiny sheet with a known row
// height on row 2, a known column width on col B, and a twoCellAnchor image
// at C1 → D2 so the figure's left offset lands at gutter + A_width + B_width
// (± a pixel of rounding) and the top offset lands at 0.
//
// Layout:
//   - Sheet "Grid": 3×3, row 2 custom height 30pt (= 40px), col B width 14
//     char-units (= 7*14 + 5 = 103px).
//   - Drawing 1 (twoCellAnchor) at C1 → D2 with zero offsets. Renders as a
//     single-cell overlay that pins to the top-left corner of C1.
//
// The PNG payload is reused from the other hand-built fixtures: a minimal
// 1×1 red PNG encoded inline so this fixture has no external asset.
//
// Run:
//   node scripts/make-in-flow-image-position-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/in-flow-image-position');
mkdirSync(outDir, { recursive: true });

// Minimal 1×1 red PNG (same bytes as scripts/make-drawing-anchors-plus-fixture.mjs).
const RED_1x1_PNG_BASE64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
const PNG_BYTES = Buffer.from(RED_1x1_PNG_BASE64, 'base64');

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
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
  <sheets><sheet name="Grid" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

// Two labels so the image has something to sit next to in the rendered DOM.
const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>row 1</t></si>
  <si><t>row 2</t></si>
</sst>`;

// <cols> gives col B (min/max=2) a custom width of 14 character units
// (= 7*14 + 5 = 103px). Row 2 carries ht=30pt (= 40px at 96 DPI) with
// customHeight="1" so the renderer applies it. Cell values in A1/A2 give
// the rows some visible content.
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>
    <col min="2" max="2" width="14" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2" ht="30" customHeight="1"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"/>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

// twoCellAnchor from C1 to D2 with zero fractional offsets at both ends,
// so the image covers exactly cell C1. Renderer's left offset arithmetic
// then lands at gutter + A_width + B_width (col A default 64px, col B 103px,
// gutter ≈ 30px → ~197px).
const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="AnchoredImage" descr="two-cell at C1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/sharedStrings.xml', sharedStringsXml);
zip.file('xl/worksheets/sheet1.xml', sheetXml);
zip.file('xl/worksheets/_rels/sheet1.xml.rels', sheetRels);
zip.file('xl/drawings/drawing1.xml', drawingXml);
zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRels);
zip.file('xl/media/image1.png', PNG_BYTES);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
