// Build tests/render-test/fills-and-borders/workbook.xlsx. Exercises the
// three style features added in feat/fills-and-borders:
//
//   A1: a cell with both diagonal borders (diagonalUp + diagonalDown) in
//       red. Renderer paints a linear-gradient overlay per diagonal.
//   A2: a linear gradient fill (red → green, 0deg).
//   A3: a non-solid pattern fill (darkHorizontal) with red fg + white bg.
//
// The fixture is self-contained — hand-written XML, minimal styles.xml
// that lines up with the cellXf indices used on the sheet. No shared
// strings (all cells are empty or hold an inline label).
//
// Run:
//   node scripts/make-fills-and-borders-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/fills-and-borders');
mkdirSync(outDir, { recursive: true });

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Fills" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>diag</t></si>
  <si><t>grad</t></si>
  <si><t>patt</t></si>
</sst>`;

// Styles:
//   fills[0] = <patternFill patternType="none"/>   (Excel default)
//   fills[1] = <patternFill patternType="gray125"/>(Excel default, we skip)
//   fills[2] = gradientFill (linear, 0deg, red → green) — used by A2
//   fills[3] = patternFill patternType="darkHorizontal" fg=red bg=white — A3
//
//   borders[0] = empty/default
//   borders[1] = diagonal border (up + down, red) with no orthogonal sides —
//                used by A1
//
//   cellXfs:
//     0 = plain (no fills / borders / fonts)
//     1 = borderId=1 (A1: diagonals)
//     2 = fillId=2   (A2: gradient)
//     3 = fillId=3   (A3: darkHorizontal)
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <gradientFill type="linear" degree="0">
        <stop position="0"><color rgb="FFFF0000"/></stop>
        <stop position="1"><color rgb="FF00FF00"/></stop>
      </gradientFill>
    </fill>
    <fill>
      <patternFill patternType="darkHorizontal">
        <fgColor rgb="FFFF0000"/>
        <bgColor rgb="FFFFFFFF"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1" diagonalDown="1">
      <left/><right/><top/><bottom/>
      <diagonal style="thin"><color rgb="FFFF0000"/></diagonal>
    </border>
  </borders>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>`;

const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s" s="2"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s" s="3"><v>2</v></c></row>
  </sheetData>
</worksheet>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/sharedStrings.xml', sharedStringsXml);
zip.file('xl/styles.xml', stylesXml);
zip.file('xl/worksheets/sheet1.xml', sheet1Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
