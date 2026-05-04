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

// ── alignment-flags ───────────────────────────────────────────────────────
// Hand-built XLSX exercising the extended <alignment/> surface. Seven cells
// each pick up a distinct flag through their own cellXf:
//   A1 (s=1) wrapText=1      — multi-line string wraps inside the cell
//   A2 (s=2) textRotation=90 — rotated header
//   A3 (s=3) textRotation=255 — stacked vertical (CJK convention)
//   A4 (s=4) indent=3 + horizontal=left     — paddingLeft ≈ 1.5em
//   A5 (s=5) readingOrder=2 + horizontal=right — direction:rtl
//   A6 (s=6) shrinkToFit=1   — tagged with .xlsx-shrink-to-fit
//   A7 (s=7) horizontal=distributed, vertical=justify — widened enums
{
    const outDir = resolve(repo, 'tests/render-test/alignment-flags');
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
    zip.file('xl/workbook.xml', workbookXml('Align'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml([
        'line one\nline two',  // 0: wrapText
        'ROT90',               // 1: textRotation=90
        'STACK',               // 2: textRotation=255
        'indented',            // 3: indent=3
        'rtl text',            // 4: readingOrder=2
        'shrinkme text that is quite long',  // 5: shrinkToFit
        'distrib',             // 6: distributed/justify
    ]));
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="90"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment textRotation="255"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="left" indent="3"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="right" readingOrder="2"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="distributed" vertical="justify"/></xf>
  </cellXfs>
</styleSheet>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s" s="2"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s" s="3"><v>2</v></c></row>
    <row r="4"><c r="A4" t="s" s="4"><v>3</v></c></row>
    <row r="5"><c r="A5" t="s" s="5"><v>4</v></c></row>
    <row r="6"><c r="A6" t="s" s="6"><v>5</v></c></row>
    <row r="7"><c r="A7" t="s" s="7"><v>6</v></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── font-extras ────────────────────────────────────────────────────────────
// Hand-built XLSX that exercises the font-extras slice: strike,
// sub/superscript (via rich-text runs); double underline (via a cell-level
// font); a safe font family ("Calibri"); and an injection-attempt font name
// ("Arial; display:block") that sanitizeFontFamily must reject.
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

// ── sheet-view-state ───────────────────────────────────────────────────────
// Hand-built XLSX exercising the per-sheet display state:
//   Alpha — visible, default view.
//   Beta  — state="hidden"; renderer drops the section entirely.
//   Gamma — visible, rightToLeft + showGridLines=0 + <sheetPr><tabColor rgb=FF0000/>.
// The workbook rels pair each <sheet r:id> with the correct worksheet file so
// the parser's rId→target resolution is also exercised.
{
    const outDir = resolve(repo, 'tests/render-test/sheet-view-state');
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId1"/>
    <sheet name="Beta"  sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="Gamma" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>`);
    // Alpha — default view, one cell.
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c></row>
  </sheetData>
</worksheet>`);
    // Beta — hidden via workbook.xml's state attribute.
    zip.file('xl/worksheets/sheet2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>beta</t></is></c></row>
  </sheetData>
</worksheet>`);
    // Gamma — RTL + gridlines off + red tab colour. <sheetPr> must come
    // before <sheetViews> per the worksheet schema sequence.
    zip.file('xl/worksheets/sheet3.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><tabColor rgb="FFFF0000"/></sheetPr>
  <sheetViews>
    <sheetView rightToLeft="1" showGridLines="0" workbookViewId="0"/>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>gamma</t></is></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── outlines-and-names ─────────────────────────────────────────────────────
// Exercises the outline-levels slice (on row + col) together with the
// workbook-scoped defined-names registry. Row outline levels go 0/1/2/1/0
// so the middle rows form a two-level group; column B carries outlineLevel=1.
// sheetFormatPr advertises the max observed levels; sheetPr/outlinePr flips
// summaryRight to false so the struct mirrors what Excel persists.
// Two definedNames: one workbook-scoped, one print-area scoped to sheet 0.
{
    const outDir = resolve(repo, 'tests/render-test/outlines-and-names');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="TotalRange">Sheet1!$A$1:$C$5</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0" hidden="1">Sheet1!$A$1:$C$3</definedName>
  </definedNames>
</workbook>`);
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['a', 'b', 'c', 'd', 'e']));
    // Row outline levels 0, 1, 2, 1, 0; column B carries outlineLevel=1.
    // sheetFormatPr advertises the max levels; sheetPr/outlinePr flips
    // summaryRight to "0" so the parser roundtrips a non-default value.
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><outlinePr summaryBelow="1" summaryRight="0"/></sheetPr>
  <sheetFormatPr defaultRowHeight="15" outlineLevelRow="2" outlineLevelCol="1"/>
  <cols>
    <col min="2" max="2" outlineLevel="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>0</v></c></row>
    <row r="2" outlineLevel="1"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>1</v></c><c r="C2" t="s"><v>1</v></c></row>
    <row r="3" outlineLevel="2"><c r="A3" t="s"><v>2</v></c><c r="B3" t="s"><v>2</v></c><c r="C3" t="s"><v>2</v></c></row>
    <row r="4" outlineLevel="1"><c r="A4" t="s"><v>3</v></c><c r="B4" t="s"><v>3</v></c><c r="C4" t="s"><v>3</v></c></row>
    <row r="5"><c r="A5" t="s"><v>4</v></c><c r="B5" t="s"><v>4</v></c><c r="C5" t="s"><v>4</v></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── hyperlinks-and-validation ─────────────────────────────────────────────
// Exercises both the hyperlink path (URL allowlist at render) and the
// type="list" data-validation surface.
//   A2 = "click me" (shared string), hyperlink rel → https://example.com
//   A3 = "danger"   (shared string), hyperlink rel → javascript:alert(1)
//                                                     (renderer MUST reject)
//   B2..B4 = type="list" data-validation with options "Red,Green,Blue".
{
    const outDir = resolve(repo, 'tests/render-test/hyperlinks-and-validation');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    zip.file('xl/workbook.xml', workbookXml('Links'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['click me', 'danger']));
    // Sheet rels: two hyperlink rels with TargetMode="External". rId1 →
    // safe https target; rId2 → attacker javascript: URL that must be
    // stripped at render time by isSafeHyperlinkHref.
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
</Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="2"><c r="A2" t="s"><v>0</v></c></row>
    <row r="3"><c r="A3" t="s"><v>1</v></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A2" r:id="rId1" tooltip="visit example"/>
    <hyperlink ref="A3" r:id="rId2"/>
  </hyperlinks>
  <dataValidations count="1">
    <dataValidation type="list" sqref="B2:B4" allowBlank="1">
      <formula1>"Red,Green,Blue"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── cf-new-rules ──────────────────────────────────────────────────────────
// Hand-built XLSX exercising the new cf rule types (containsBlanks,
// notContainsErrors, aboveAverage). A1:A5 = [1, 2, 3, '', 5]. Three dxfs:
//   0: red fill       → containsBlanks → only A4 matches
//   1: green fill     → notContainsErrors → all cells match (no error cells)
//   2: bold font      → aboveAverage → cells above mean of {1,2,3,5} (2.75)
// Priorities are ordered so the test can disentangle which dxf landed where.
{
    const outDir = resolve(repo, 'tests/render-test/cf-new-rules');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('CfNewRules'));
    // Three dxfs:
    //   0: solid red fill         (FFFF9999)
    //   1: solid green fill       (FF99FF99)
    //   2: bold font
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="3">
    <dxf><fill><patternFill><bgColor rgb="FFFF9999"/></patternFill></fill></dxf>
    <dxf><fill><patternFill><bgColor rgb="FF99FF99"/></patternFill></fill></dxf>
    <dxf><font><b/></font></dxf>
  </dxfs>
</styleSheet>`);
    // Rules chosen so each cell gets a single unambiguous dxf match:
    //   priority 1 (lowest = highest priority): aboveAverage → bold on A5.
    //   priority 2:                             containsBlanks → red on A4.
    //   priority 3:                             notContainsErrors → green
    //     on everything else (A1, A2, A3). A4 and A5 are already claimed
    //     by higher-priority rules so the green rule doesn't overwrite them.
    //
    // A1:A5 = 1, 2, 3, "", 5  — A4 is genuinely blank (no <c> cell at all).
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c></row>
    <row r="4"/>
    <row r="5"><c r="A5"><v>5</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A5">
    <cfRule type="aboveAverage" dxfId="2" priority="1"/>
    <cfRule type="containsBlanks" dxfId="0" priority="2">
      <formula>LEN(TRIM(A1))=0</formula>
    </cfRule>
    <cfRule type="notContainsErrors" dxfId="1" priority="3">
      <formula>NOT(ISERROR(A1))</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── cf-ext-databar ────────────────────────────────────────────────────────
// Exercises the x14 ext attributes for data bars (negative fill colour,
// middle axis, border + border colour) and the dxf strike/numFmt extras.
//
//   A1 = -5   B1 = "flag"    (containsText "flag" → dxf with strike + 0.00)
//   A2 =  0   B2 = "no"
//   A3 =  3   B3 = "flag"
//   A4 =  5   B4 = "quiet"
//
// Column C holds numeric cells that match a cellIs >= 0 rule using the
// same dxf so the dxf.numFmtCode path is exercised end-to-end.
{
    const outDir = resolve(repo, 'tests/render-test/cf-ext-databar');
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
    zip.file('xl/workbook.xml', workbookXml('ExtBar'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['flag', 'no', 'quiet']));
    // dxfs[0] = strike + numFmtCode "0.00". Fires via containsText (col B)
    // and cellIs >= 0 (col C). The numFmt path matters on C.
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="1">
    <dxf>
      <font><strike val="1"/></font>
      <numFmt numFmtId="164" formatCode="0.00"/>
    </dxf>
  </dxfs>
</styleSheet>`);
    const GUID = '{11111111-2222-3333-4444-555555555555}';
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <sheetData>
    <row r="1"><c r="A1"><v>-5</v></c><c r="B1" t="s"><v>0</v></c><c r="C1"><v>-2</v></c></row>
    <row r="2"><c r="A2"><v>0</v></c><c r="B2" t="s"><v>1</v></c><c r="C2"><v>4</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3" t="s"><v>0</v></c><c r="C3"><v>7</v></c></row>
    <row r="4"><c r="A4"><v>5</v></c><c r="B4" t="s"><v>2</v></c><c r="C4"><v>9</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A4">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
      <extLst>
        <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <x14:id>${GUID}</x14:id>
        </ext>
      </extLst>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="B1:B4">
    <cfRule type="containsText" priority="2" operator="containsText" dxfId="0" text="flag">
      <formula>NOT(ISERROR(SEARCH("flag",B1)))</formula>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C1:C4">
    <cfRule type="cellIs" priority="3" operator="greaterThanOrEqual" dxfId="0">
      <formula>0</formula>
    </cfRule>
  </conditionalFormatting>
  <extLst>
    <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
          <x14:cfRule type="dataBar" id="${GUID}">
            <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarColorSameAsPositive="0" axisPosition="middle">
              <x14:cfvo type="autoMin"/>
              <x14:cfvo type="autoMax"/>
              <x14:borderColor rgb="FF0000FF"/>
              <x14:negativeFillColor rgb="FFFF0000"/>
              <x14:negativeBorderColor rgb="FFFF0000"/>
              <x14:axisColor rgb="FF000000"/>
            </x14:dataBar>
          </x14:cfRule>
          <xm:sqref>A1:A4</xm:sqref>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
  </extLst>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── numfmt-r2 ─────────────────────────────────────────────────────────────
// Hand-built XLSX exercising the number-format round-2 surface:
//   A1 (s=1) numFmtId=164 "[Red]#,##0"          value=100   → "100", red text
//   A2 (s=2) numFmtId=165 "[>100][Red]#,##0;[<0][Blue]#,##0;#,##0"
//                                                 value=150   → "150", red
//   A3 (s=2) same predicate format               value=-50   → "-50", blue
//   A4 (s=2) same predicate format               value=25    → "25",  default
// The cellXf chain uses applyNumberFormat="1" so the format code wins.
{
    const outDir = resolve(repo, 'tests/render-test/numfmt-r2');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
    zip.file('xl/workbook.xml', workbookXml('NumFmtR2'));
    // Two custom numFmts:
    //   164: [Red]#,##0                                    — cell-level colour
    //   165: [>100][Red]#,##0;[<0][Blue]#,##0;#,##0       — conditional
    // formatCode attribute uses &gt; / &lt; since the XML attribute value
    // can't carry the raw < / > characters.
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="[Red]#,##0"/>
    <numFmt numFmtId="165" formatCode="[&gt;100][Red]#,##0;[&lt;0][Blue]#,##0;#,##0"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="1"><v>100</v></c></row>
    <row r="2"><c r="A2" s="2"><v>150</v></c></row>
    <row r="3"><c r="A3" s="2"><v>-50</v></c></row>
    <row r="4"><c r="A4" s="2"><v>25</v></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── page-layout ───────────────────────────────────────────────────────────
// Exercises the page-layout metadata surface: manual row/column page
// breaks, the `_xlnm.Print_Area` defined name resolving to a cell range,
// and a <headerFooter> with &L/&C/&R zones plus substitution codes.
// Two manual row breaks (XML ids 10 and 20 → 0-based indices 9 and 19)
// and one manual col break (id 5 → idx 4). Print area "Sheet1!$A$1:$C$5"
// resolves to {col:0, row:0, endCol:2, endRow:4}. The header's &D code
// is substituted with today's date; &P stays literal (no pagination).
{
    const outDir = resolve(repo, 'tests/render-test/page-layout');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$5</definedName>
  </definedNames>
</workbook>`);
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['a', 'b', 'c']));
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c></row>
    <row r="3"><c r="A3"><v>4</v></c><c r="B3"><v>5</v></c><c r="C3"><v>6</v></c></row>
  </sheetData>
  <headerFooter>
    <oddHeader>&amp;LMy Report&amp;C&amp;D&amp;RPage &amp;P</oddHeader>
    <oddFooter>&amp;CFooter</oddFooter>
  </headerFooter>
  <rowBreaks count="2" manualBreakCount="2">
    <brk id="10" man="1"/>
    <brk id="20" man="1"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="5" man="1"/>
  </colBreaks>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── theme-fontscheme ──────────────────────────────────────────────────────
// Hand-built XLSX exercising the theme fontScheme resolution path. The
// theme declares majorFont latin="Cambria" and minorFont latin="Calibri";
// A1 uses a cellXf whose font carries <scheme val="major"/> with no <name/>
// so the renderer must fall back to the theme's majorFont typeface.
{
    const outDir = resolve(repo, 'tests/render-test/theme-fontscheme');
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
    zip.file('xl/workbook.xml', workbookXml('SchemeFont'));
    zip.file('xl/sharedStrings.xml', sharedStringsXml(['heading', 'body']));
    // fonts:
    //   0 = default (no scheme, no name → CSS fontFamily untouched)
    //   1 = <scheme val="major"/> no name → theme majorFont (Cambria)
    //   2 = <scheme val="minor"/> no name → theme minorFont (Calibri)
    zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/></font>
    <font><sz val="14"/><scheme val="major"/></font>
    <font><sz val="11"/><scheme val="minor"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`);
    // Theme with distinct major / minor typefaces so each scheme path can be
    // verified independently. Minimal clrScheme so parseTheme still fills the
    // 12-entry colour table.
    zip.file('xl/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s" s="1"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s" s="2"><v>1</v></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}

// ── phonetics ─────────────────────────────────────────────────────────────
// Hand-built XLSX exercising the phonetic-ruby (<rPh>) path. One shared
// string carries base text "漢字" with a single phonetic annotation
// "かんじ" covering characters [0..2). A second cell is inline-str with the
// same pattern so the inlineStr path also gets a rPh.
{
    const outDir = resolve(repo, 'tests/render-test/phonetics');
    mkdirSync(outDir, { recursive: true });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    zip.file('xl/workbook.xml', workbookXml('Phon'));
    // One shared <si> with:
    //   <t>漢字</t>
    //   <rPh sb="0" eb="2"><t>かんじ</t></rPh>
    //   <phoneticPr fontId="1" type="noConversion"/>
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>漢字</t>
    <rPh sb="0" eb="2"><t>かんじ</t></rPh>
    <phoneticPr fontId="1" type="noConversion"/>
  </si>
</sst>`);
    // A1 uses the shared string (shared-string path).
    // A2 is an inline string carrying its own <rPh> (inlineStr path).
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh></is></c></row>
  </sheetData>
</worksheet>`);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const out = resolve(outDir, 'workbook.xlsx');
    writeFileSync(out, buf);
    console.log(`wrote ${out} (${buf.length} bytes)`);
}
