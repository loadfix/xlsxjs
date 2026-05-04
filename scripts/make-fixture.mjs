// Build fixture XLSX packages programmatically so the repo can ship working
// tests without requiring Excel/LibreOffice to generate binary fixtures.
// Produces:
//   tests/render-test/basic/workbook.xlsx   — headers + numbers + booleans
//   tests/render-test/merged/workbook.xlsx  — merged header + column widths

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');

function escapeXml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
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

function workbookXml(sheetName) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

function sharedStringsXml(strings) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
  ${strings.map((s) => `<si><t>${escapeXml(s)}</t></si>`).join('\n  ')}
</sst>`;
}

async function writeFixture(name, { sheetName, shared, sheetBody }) {
    const outDir = resolve(repo, 'tests/render-test', name);
    mkdirSync(outDir, { recursive: true });

    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    zip.file('xl/workbook.xml', workbookXml(sheetName));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(shared));
    zip.file('xl/worksheets/sheet1.xml', sheetBody);

    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── basic ──────────────────────────────────────────────────────────────────
//   A1=Name  B1=Score  C1=Passed
//   A2=Alice B2=90     C2=TRUE
//   A3=Bob   B3=72     C3=FALSE
await writeFixture('basic', {
    sheetName: 'People',
    shared: ['Name', 'Score', 'Passed', 'Alice', 'Bob'],
    sheetBody: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
    <row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>90</v></c><c r="C2" t="b"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3"><v>72</v></c><c r="C3" t="b"><v>0</v></c></row>
  </sheetData>
</worksheet>`,
});

// ── richtext ───────────────────────────────────────────────────────────────
// Hand-built XLSX: writeFixture() only supports plain shared strings, so we
// inline the shared-strings + styles + theme parts directly. One cell carries
// three runs (plain / bold-red / italic) via the rich-text path; a second
// cell uses a font whose colour is a theme reference with tint to verify
// theme colour resolution.
{
    const outDir = resolve(repo, 'tests/render-test/richtext');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('Rich'));

    // Custom shared strings: three entries.
    //   0: rich "Hello " + bold red "world" + italic "!"
    //   1: plain "plain text"
    //   2: theme-coloured "accent1 tinted"
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si>
    <r><t xml:space="preserve">Hello </t></r>
    <r><rPr><b/><color rgb="FFFF0000"/></rPr><t>world</t></r>
    <r><rPr><i/></rPr><t>!</t></r>
  </si>
  <si><t>plain text</t></si>
  <si>
    <r><rPr><color theme="4" tint="-0.25"/></rPr><t>accent1 tinted</t></r>
  </si>
</sst>`);

    // Minimal styles.xml: one font, one fill (none), one border (none), two
    // cellXfs. xf[1] uses fontId 1 (a theme-coloured font) — distinct from
    // the run-level theme colour so we can verify both paths work.
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><color theme="5" tint="0"/><name val="Calibri"/></font>
  </fonts>
  <fills count="1">
    <fill><patternFill patternType="none"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`);

    // A theme that maps accent1 → red, accent2 → green so we can verify the
    // resolver. Use a bare 4-colour scheme plus accents, wrapped in the
    // minimal a:theme / a:clrScheme skeleton Excel accepts.
    zip.file('xl/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
      <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
      <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
      <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Custom"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`);

    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s" s="1"><v>2</v></c></row>
  </sheetData>
</worksheet>`);

    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── named-styles ───────────────────────────────────────────────────────────
// Verifies cellStyleXfs inheritance: a cellXf with xfId=1 and no apply-flags
// must still pick up the base xf's font / fill. The named style here is
// "Heading 1" (bold, coloured). A1 uses it; A2 uses a plain xf.
{
    const outDir = resolve(repo, 'tests/render-test/named-styles');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('Styles'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['heading cell', 'normal cell']));
    // fonts[0] = plain, fonts[1] = bold red.
    // cellStyleXfs[0] = "Normal" (default), cellStyleXfs[1] = "Heading 1"
    //    (applyFont=1, fontId=1).
    // cellXfs[0] = plain; cellXfs[1] = xfId=1 with *no* apply-flags — the
    //    effective font should still come from the base cellStyleXfs[1].
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFCC0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Heading 1" xfId="1" builtinId="16"/>
  </cellStyles>
</styleSheet>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s" s="0"><v>1</v></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── dimensions ─────────────────────────────────────────────────────────────
// Custom row heights, a hidden row, and a hidden column. Column B is hidden
// (so cells in col 1 should render with display:none); row 3 (index 2) is
// hidden; row 1 (index 0) carries a 40pt custom height, row 2 (index 1) a
// 24pt height.
await writeFixture('dimensions', {
    sheetName: 'Dims',
    shared: ['A', 'B', 'C', 'tall', 'short', 'hidden', 'normal'],
    sheetBody: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" hidden="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="40" customHeight="1"><c r="A1" t="s"><v>3</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
    <row r="2" ht="24" customHeight="1"><c r="A2" t="s"><v>4</v></c><c r="B2" t="s"><v>1</v></c><c r="C2" t="s"><v>2</v></c></row>
    <row r="3" hidden="1"><c r="A3" t="s"><v>5</v></c><c r="B3" t="s"><v>1</v></c><c r="C3" t="s"><v>2</v></c></row>
    <row r="4"><c r="A4" t="s"><v>6</v></c><c r="B4" t="s"><v>1</v></c><c r="C4" t="s"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
});

// ── multisheet ─────────────────────────────────────────────────────────────
// Three sheets with non-sequential rId→target mapping, to verify the parser
// uses workbook.xml.rels rather than positional resolution. Also exercises
// out-of-order rIds (rId3 → sheet1, rId1 → sheet2, rId2 → sheet3).
{
    const outDir = resolve(repo, 'tests/render-test/multisheet');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId3"/>
    <sheet name="Beta"  sheetId="2" r:id="rId1"/>
    <sheet name="Gamma" sheetId="3" r:id="rId2"/>
  </sheets>
</workbook>`);
    const sheetBody = (label, n) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>${label}</t></is></c><c r="B1"><v>${n}</v></c></row>
  </sheetData>
</worksheet>`;
    // Deliberately name the xml files so that rId mapping matters:
    //   rId3 → sheet1.xml  ("Alpha", n=1)
    //   rId1 → sheet2.xml  ("Beta",  n=2)
    //   rId2 → sheet3.xml  ("Gamma", n=3)
    zip.file('xl/worksheets/sheet1.xml', sheetBody('Alpha', 1));
    zip.file('xl/worksheets/sheet2.xml', sheetBody('Beta',  2));
    zip.file('xl/worksheets/sheet3.xml', sheetBody('Gamma', 3));

    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── formulas ───────────────────────────────────────────────────────────────
// Exercises the four formula states a consumer can hit:
//   A1=10       B1=20       C1=30   (plain numeric inputs)
//   A2="=SUM(A1:C1)" with cached <v>60>      → renders 60
//   A3="=A1+A2"      with no cached value    → renders empty
//   A4="=A1&\" rows\"" string result "10 rows" (t="str" + cached)
//   A5 error formula, t="e", <v>#DIV/0!</v>
await writeFixture('formulas', {
    sheetName: 'Formulas',
    shared: [],
    sheetBody: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c><c r="C1"><v>30</v></c></row>
    <row r="2"><c r="A2"><f>SUM(A1:C1)</f><v>60</v></c></row>
    <row r="3"><c r="A3"><f>A1+A2</f></c></row>
    <row r="4"><c r="A4" t="str"><f>A1&amp;" rows"</f><v>10 rows</v></c></row>
    <row r="5"><c r="A5" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>
  </sheetData>
</worksheet>`,
});

// ── threaded-comments ──────────────────────────────────────────────────────
// Excel 365's modern comment model: a workbook-wide person registry under
// xl/persons/person.xml + per-sheet xl/threadedComments/threadedComment{N}.xml.
// Two authors (Alice, Bob) and two comments at A1 forming a parent+reply
// thread, plus the sheet rels pointing the sheet at the threadedComments
// part and the workbook rels pointing at the persons part.
{
    const outDir = resolve(repo, 'tests/render-test/threaded-comments');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  <Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    // Workbook rels: sheet + sharedStrings + person registry.
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('Thread'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['Review this']));
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`);
    // Sheet rels → threadedComments part.
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`);
    // Two authors.
    zip.file('xl/persons/person.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <person displayName="Alice" id="{AAAAAAAA-1111-2222-3333-444444444444}" userId="alice@example.com" providerId="AD"/>
  <person displayName="Bob" id="{BBBBBBBB-1111-2222-3333-444444444444}" userId="bob@example.com" providerId="AD"/>
</personList>`);
    // Parent+reply thread at A1.
    zip.file('xl/threadedComments/threadedComment1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
                   xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <threadedComment ref="A1" dT="2024-01-15T10:00:00" personId="{AAAAAAAA-1111-2222-3333-444444444444}" id="{C0000000-0000-0000-0000-000000000001}">
    <text>Please double-check this figure.</text>
  </threadedComment>
  <threadedComment ref="A1" dT="2024-01-15T10:05:00" personId="{BBBBBBBB-1111-2222-3333-444444444444}" id="{C0000000-0000-0000-0000-000000000002}" parentId="{C0000000-0000-0000-0000-000000000001}">
    <text>Checked — looks correct.</text>
  </threadedComment>
</ThreadedComments>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── merged ─────────────────────────────────────────────────────────────────
// Horizontal merge on the title banner, custom column widths, a vertical
// merge in the "Region" column.
//   A1:C1 = "Q4 Report"  (horizontal 3-way merge, anchored at A1)
//   A2=Name  B2=Region      C2=Score
//   A3=Alice A3:A4=merged   B3=North C3=90
//   A4=                     B4=North C4=85
await writeFixture('merged', {
    sheetName: 'Report',
    shared: ['Q4 Report', 'Name', 'Region', 'Score', 'Alice', 'North'],
    sheetBody: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="18" customWidth="1"/>
    <col min="2" max="2" width="10" customWidth="1"/>
    <col min="3" max="3" width="8" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c></row>
    <row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3" t="s"><v>5</v></c><c r="C3"><v>90</v></c></row>
    <row r="4"><c r="B4" t="s"><v>5</v></c><c r="C4"><v>85</v></c></row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:C1"/>
    <mergeCell ref="A3:A4"/>
  </mergeCells>
</worksheet>`,
});

// ── font-extras ────────────────────────────────────────────────────────────
// Hand-built XLSX that exercises the 2026-Q2 font-extras slice: strike,
// sub/superscript (via rich-text runs); double underline (via a cell-level
// font); a safe font family ("Calibri"); and an injection-attempt font name
// ("Arial; display:block") that sanitizeFontFamily must reject.
//
// Shared strings:
//   0: rich text → "struck" (strike) + "H" (plain) + "2" (subscript) +
//      "O" (plain) + " " + "sup" (superscript)
//   1: "double underline"        (cell-level font, cellXf s=1)
//   2: "calibri"                 (cell-level font with name="Calibri", s=2)
//   3: "injected"                (cell-level font with attacker-provided
//                                 name, s=3 → sanitizer rejects)
{
    const outDir = resolve(repo, 'tests/render-test/font-extras');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('FontExtras'));

    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si>
    <r><rPr><strike/></rPr><t>struck</t></r>
    <r><t xml:space="preserve"> H</t></r>
    <r><rPr><vertAlign val="subscript"/></rPr><t>2</t></r>
    <r><t xml:space="preserve">O </t></r>
    <r><rPr><vertAlign val="superscript"/></rPr><t>sup</t></r>
  </si>
  <si><t>double underline</t></si>
  <si><t>calibri</t></si>
  <si><t>injected</t></si>
</sst>`);

    // fonts:
    //   0 = default
    //   1 = double underline (<u val="double"/>)
    //   2 = name="Calibri" (safe, sanitizer wraps in quotes)
    //   3 = name="Arial; display:block" (injection attempt, sanitizer rejects)
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><u val="double"/><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Arial; display:block"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`);

    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s" s="1"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s" s="2"><v>2</v></c></row>
    <row r="4"><c r="A4" t="s" s="3"><v>3</v></c></row>
  </sheetData>
</worksheet>`);

    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}
