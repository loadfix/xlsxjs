// Build tests/render-test/cf-expression-rules/workbook.xlsx. Exercises the
// expanded conditional-format <expression> interpreter:
//
//   A2:A6 — numeric column. Rule: =AND(A2>5, A2<10) fires for values in
//           the open interval (5, 10). Values 3, 6, 7, 10, 12.
//   B2:B6 — mixed column (blank / error / number). Rule:
//           =OR(ISBLANK(B2), ISERROR(B2)) fires on blank & error cells.
//   C2:C6 — row banding. Rule: =MOD(ROW(), 2) = 0 fires on even rows.
//   D2:D6 — text column. Rule: =SEARCH("flag", D2) fires when the cell
//           text contains "flag" (case-insensitive).
//
// Styles:
//   dxfs[0] = bold red fill           — AND rule (A)
//   dxfs[1] = yellow fill              — OR rule  (B)
//   dxfs[2] = blue fill                — banding  (C)
//   dxfs[3] = green fill               — SEARCH   (D)
//
// Run:
//   node scripts/make-cf-expression-rules-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/cf-expression-rules');
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
    <sheet name="Exprs" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

// Shared strings:
//   0: "alpha"          (D2)
//   1: "red flag"       (D3)
//   2: "beta"           (D4)
//   3: "flag it"        (D5)
//   4: "gamma"          (D6)
const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>alpha</t></si>
  <si><t>red flag</t></si>
  <si><t>beta</t></si>
  <si><t>flag it</t></si>
  <si><t>gamma</t></si>
</sst>`;

// Styles: one base xf + four dxfs (one per CF rule).
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="4">
    <dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFF9999"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FFFFFF99"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FFCCE5FF"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FFCCFFCC"/></patternFill></fill></dxf>
  </dxfs>
</styleSheet>`;

// Sheet content:
//   A: 3, 6, 7, 10, 12        (AND rule fires on 6 and 7)
//   B: 1, <blank>, #DIV/0!, 4, 5   (OR rule fires on blank B3 and error B4)
//   C: 10, 20, 30, 40, 50    (banding fires on rows 2, 4, 6 → even excel rows)
//   D: alpha, "red flag", beta, "flag it", gamma  (SEARCH fires on D3 + D5)
const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>1</v></c>
      <c r="C2"><v>10</v></c>
      <c r="D2" t="s"><v>0</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>6</v></c>
      <c r="B3"/>
      <c r="C3"><v>20</v></c>
      <c r="D3" t="s"><v>1</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>7</v></c>
      <c r="B4" t="e"><v>#DIV/0!</v></c>
      <c r="C4"><v>30</v></c>
      <c r="D4" t="s"><v>2</v></c>
    </row>
    <row r="5">
      <c r="A5"><v>10</v></c>
      <c r="B5"><v>4</v></c>
      <c r="C5"><v>40</v></c>
      <c r="D5" t="s"><v>3</v></c>
    </row>
    <row r="6">
      <c r="A6"><v>12</v></c>
      <c r="B6"><v>5</v></c>
      <c r="C6"><v>50</v></c>
      <c r="D6" t="s"><v>4</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A2:A6">
    <cfRule type="expression" dxfId="0" priority="1">
      <formula>AND(A2&gt;5, A2&lt;10)</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="B2:B6">
    <cfRule type="expression" dxfId="1" priority="2">
      <formula>OR(ISBLANK(B2), ISERROR(B2))</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C2:C6">
    <cfRule type="expression" dxfId="2" priority="3">
      <formula>MOD(ROW(), 2) = 0</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="D2:D6">
    <cfRule type="expression" dxfId="3" priority="4">
      <formula>SEARCH("flag", D2)</formula>
    </cfRule>
  </conditionalFormatting>
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
