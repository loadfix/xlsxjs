// Build tests/render-test/slicers-pivot/workbook.xlsx. Hand-written XLSX
// covering the Wave 12 pivot re-materialisation slice:
//
//   - A single sheet ("Dash") with a materialised pivot table: one header
//     row ("Region | Sales") and four data rows (North / South / East /
//     West → 100 / 200 / 300 / 400). Unlike the Wave 8 slicers-timelines
//     fixture, the pivot cells are really written to the sheet so the
//     renderer emits real <tbody><tr> that a re-materialisation handler
//     can hide on slicer change.
//   - One slicer on the Region column. The cache enumerates all four
//     items and flags every one as selected via s="1", so the default
//     DOM starts with every row visible.
//   - One timeline is intentionally omitted; scenario 115 focuses on
//     slicer re-materialisation. A timeline would need a date column
//     whose textContent parses via `new Date(…)`, which the default
//     formatter doesn't produce without explicit numFmt wiring.
//
// Run:
//   node scripts/make-slicers-pivot-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/slicers-pivot');
mkdirSync(outDir, { recursive: true });

// ── [Content_Types].xml ──────────────────────────────────────────────────
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
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
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
</Relationships>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Dash" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId4"/>
  </pivotCaches>
</workbook>`;

// Shared-strings table: header labels (Region, Sales) + the four region
// names (North, South, East, West). The pivot's data cells reference these
// via <c t="s"><v>idx</v></c>.
const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>Region</t></si>
  <si><t>Sales</t></si>
  <si><t>North</t></si>
  <si><t>South</t></si>
  <si><t>East</t></si>
  <si><t>West</t></si>
</sst>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`;

// Sheet1: the materialised pivot cells. A1:B1 is the header row
// (Region | Sales); A2..A5 carry the four region labels and B2..B5 the
// corresponding totals. The row count of 4 gives us something to assert
// against before/after slicer filtering.
const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><v>200</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>4</v></c>
      <c r="B4"><v>300</v></c>
    </row>
    <row r="5">
      <c r="A5" t="s"><v>5</v></c>
      <c r="B5"><v>400</v></c>
    </row>
  </sheetData>
  <extLst>
    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
         uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}">
      <x14:slicerList>
        <x14:slicer r:id="rId2"/>
      </x14:slicerList>
    </ext>
  </extLst>
</worksheet>`;

const sheet1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
</Relationships>`;

// A minimal pivotTable part. Location A1:B5 lines up with the materialised
// cells on Sheet1 (the renderer doesn't actually consult this; it's here
// so parsePivotTable still produces a SheetPivot entry).
const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      name="PivotTable1" cacheId="1" applyNumberFormats="0"
                      dataCaption="Values">
  <location ref="A1:B5" firstHeaderRow="1" firstDataRow="1" firstDataCol="0"/>
  <pivotFields count="2">
    <pivotField name="Region" axis="axisRow" showAll="0"/>
    <pivotField name="Sales" dataField="1" showAll="0"/>
  </pivotFields>
</pivotTableDefinition>`;

const pivotCacheDefinitionXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      recordCount="0" createdVersion="6" refreshedVersion="6">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Dash" ref="A1:B5"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Sales" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>`;

// One slicer: all four regions start selected. Scenario 115 unchecks one
// chip and asserts the corresponding row hides.
const slicer1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="SlicerRegion" cache="Slicer_Region" caption="Region" columnCount="1" style="SlicerStyleLight1"/>
</slicers>`;

const slicerCache1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                       name="Slicer_Region" sourceName="Region">
  <pivotTables>
    <pivotTable tabId="1" name="PivotTable1"/>
  </pivotTables>
  <data>
    <tabular pivotCacheId="1">
      <items count="4">
        <i x="0" s="1" n="North"/>
        <i x="1" s="1" n="South"/>
        <i x="2" s="1" n="East"/>
        <i x="3" s="1" n="West"/>
      </items>
    </tabular>
  </data>
</slicerCacheDefinition>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/sharedStrings.xml', sharedStringsXml);
zip.file('xl/styles.xml', stylesXml);
zip.file('xl/worksheets/sheet1.xml', sheet1Xml);
zip.file('xl/worksheets/_rels/sheet1.xml.rels', sheet1Rels);
zip.file('xl/pivotTables/pivotTable1.xml', pivotTableXml);
zip.file('xl/pivotCache/pivotCacheDefinition1.xml', pivotCacheDefinitionXml);
zip.file('xl/slicers/slicer1.xml', slicer1Xml);
zip.file('xl/slicerCaches/slicerCache1.xml', slicerCache1Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
