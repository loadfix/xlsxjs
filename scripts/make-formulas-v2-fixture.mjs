// Build the formulas-v2 fixture: a hand-authored XLSX that exercises the
// wave-12 formula surface — IFERROR / SWITCH / VLOOKUP / ROUND / SUMIF plus
// the text family, the new error taxonomy, and the number-format routing
// path (0.1+0.2 in both the default General format and a 2-decimal format).
//
// The fixture ships without cached `<v>` values for the v2 formula cells so
// `evaluateFormulas: true` actually has work to do at parse time. A handful
// of cells carry seed data (numeric inputs, a small lookup table).

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');

const outDir = resolve(repo, 'tests/render-test/formulas-v2');
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
    <sheet name="FormulasV2" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

// Shared strings: lookup keys ("alpha"/"beta"/"gamma"), user-facing labels,
// and a padded string we can TRIM.
const sharedStrings = [
    'alpha', 'beta', 'gamma',     // 0..2: lookup keys
    'first', 'second', 'third',   // 3..5: lookup values
    '  hello   world  ',          // 6: TRIM source
    'mixed case',                 // 7: UPPER/LOWER source
    'foo', 'bar',                 // 8..9: SUMIF/COUNTIF criteria
];

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
  ${sharedStrings.map((s) => `<si><t xml:space="preserve">${s}</t></si>`).join('\n  ')}
</sst>`;

// Styles: default xf (General), plus xf#1 with numFmtId=2 (built-in "0.00").
// numFmt built-ins: 0 = General, 2 = "0.00".
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`;

// Sheet layout:
//   Col A: numeric seed data  (10, 20, 30, 5, 15)
//   Col B: lookup keys        ("alpha", "beta", "gamma")
//   Col C: lookup values      ("first", "second", "third")
//   Col D: category labels    ("foo", "bar", "foo", "baz", "foo")
//   Col E: formula cells — no cached <v>, so evaluator writes the result.
//
// Row 6+ holds number-format routing probes:
//   E6 (General) =0.1+0.2      → 0.3
//   E7 (fmt #1, 0.00) =0.1+0.2 → 0.30
const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1" t="s"><v>0</v></c>
      <c r="C1" t="s"><v>3</v></c>
      <c r="D1" t="s"><v>8</v></c>
      <c r="E1"><f>IFERROR(1/0, 42)</f></c>
    </row>
    <row r="2">
      <c r="A2"><v>20</v></c>
      <c r="B2" t="s"><v>1</v></c>
      <c r="C2" t="s"><v>4</v></c>
      <c r="D2" t="s"><v>9</v></c>
      <c r="E2"><f>SWITCH(2, 1, "one", 2, "two", "other")</f></c>
    </row>
    <row r="3">
      <c r="A3"><v>30</v></c>
      <c r="B3" t="s"><v>2</v></c>
      <c r="C3" t="s"><v>5</v></c>
      <c r="D3" t="s"><v>8</v></c>
      <c r="E3"><f>VLOOKUP("beta", B1:C3, 2)</f></c>
    </row>
    <row r="4">
      <c r="A4"><v>5</v></c>
      <c r="D4" t="s"><v>9</v></c>
      <c r="E4"><f>ROUND(1.2345, 2)</f></c>
    </row>
    <row r="5">
      <c r="A5"><v>15</v></c>
      <c r="D5" t="s"><v>8</v></c>
      <c r="E5"><f>SUMIF(D1:D5, "foo", A1:A5)</f></c>
    </row>
    <row r="6">
      <c r="B6" t="s"><v>6</v></c>
      <c r="C6" t="s"><v>7</v></c>
      <c r="E6"><f>0.1+0.2</f></c>
    </row>
    <row r="7">
      <c r="E7" s="1"><f>0.1+0.2</f></c>
    </row>
    <row r="8">
      <c r="E8"><f>TRIM(B6)</f></c>
    </row>
    <row r="9">
      <c r="E9"><f>UPPER(C6)</f></c>
    </row>
    <!-- Error taxonomy probes -->
    <row r="10"><c r="E10"><f>5/0</f></c></row>                   <!-- #DIV/0! -->
    <row r="11"><c r="E11"><f>"x"+1</f></c></row>                 <!-- #VALUE! -->
    <row r="12"><c r="E12"><f>NOSUCHFUNC(1)</f></c></row>          <!-- #NAME? -->
    <row r="13"><c r="E13"><f>SQRT(-1)</f></c></row>               <!-- #NUM! -->
    <row r="14"><c r="E14"><f>VLOOKUP("z", B1:C3, 2)</f></c></row> <!-- #N/A -->
    <row r="15"><c r="E15"><f>#REF!</f></c></row>                  <!-- #REF! -->
    <row r="16"><c r="E16"><f>#NULL!</f></c></row>                 <!-- #NULL! -->
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
