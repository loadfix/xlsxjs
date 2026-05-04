// Build tests/render-test/slicers-timelines/workbook.xlsx. Hand-written
// XLSX covering the Wave 8 detect-only slice:
//
//   - One slicer anchored on a trivial pivot table (`Region` field), with
//     two items flagged as currently selected. The cache sits under
//     xl/slicerCaches/slicerCache1.xml and exposes a `<tabular><items>`
//     block with `@n` item names (inline, no pivot-cache lookup needed).
//   - One timeline on the same pivot (`Date` field), declaring a months-
//     level view with a concrete start/end selection range. The cache
//     sits under xl/timelineCaches/timelineCache1.xml.
//
// The pivot table itself is a stub — slicers/timelines reference it by
// tabId but xlsxjs doesn't re-compute pivot values. We still write the
// pivotCacheDefinition + a minimal cacheSource because Excel's schema
// validators (and our own parsePivotTable) expect a resolvable chain.
// Run:
//   node scripts/make-slicers-timelines-fixture.mjs
//
// The slicer's caption contains an ASCII double quote (Region "A") so the
// renderer scenario can spot-check attribute escaping via setAttribute.

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/slicers-timelines');
mkdirSync(outDir, { recursive: true });

// ── [Content_Types].xml ──────────────────────────────────────────────────
// Slicer + slicer-cache + timeline + timeline-cache parts each need an
// <Override> MIME binding; without them Excel (and some open-source
// consumers) refuses to recognise the part type.
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
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

// Workbook rels. The slicer/timeline *caches* are workbook-scoped (same
// pattern Excel uses — caches ride on the workbook, widgets ride on a
// sheet). rId1=sheet, rId2=sharedStrings, rId3=styles, rId4=pivotCache,
// rId5=slicerCache, rId6=timelineCache.
const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/>
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

// A tiny shared-strings table. The slicer caption carries an ASCII double
// quote so the renderer tests can spot-check attribute escaping. The
// strings here are referenced from the sheet's <c t="s"> cells below.
const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Region</t></si>
  <si><t>Date</t></si>
  <si><t>Sales</t></si>
</sst>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`;

// Sheet1: a tiny 1-row header. The slicer + timeline attach here via
// sheet1.xml.rels (TargetMode="Internal"). No pivot output is rendered on
// the sheet itself — the pivot-table part is a stub referenced only by
// the slicer/timeline caches.
const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
  </sheetData>
  <extLst>
    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
         uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}">
      <x14:slicerList>
        <x14:slicer r:id="rId2"/>
      </x14:slicerList>
    </ext>
    <ext xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
         uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}">
      <x15:timelineRefs>
        <x15:timelineRef r:id="rId3"/>
      </x15:timelineRefs>
    </ext>
  </extLst>
</worksheet>`;

// Sheet rels: rId1 = pivotTable, rId2 = slicer, rId3 = timeline.
const sheet1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`;

// A minimal pivotTable part. We declare a ref anchor so parsePivotTable
// finds the location; no cell output is required for detect-only slicer/
// timeline parsing. cacheId matches the workbook's <pivotCache>.
const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      name="PivotTable1" cacheId="1" applyNumberFormats="0"
                      dataCaption="Values">
  <location ref="A3:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
  <pivotFields count="3">
    <pivotField name="Region" axis="axisRow" showAll="0"/>
    <pivotField name="Date" axis="axisRow" showAll="0"/>
    <pivotField name="Sales" dataField="1" showAll="0"/>
  </pivotFields>
</pivotTableDefinition>`;

// Stub pivot cache. xlsxjs ignores everything here — it only matters that
// the file exists so the workbook rels chain stays resolvable. The
// cacheSource uses a worksheet reference pointing at the sheet's A1:C1
// header (avoiding any real data rows). count=0 keeps it empty.
const pivotCacheDefinitionXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      recordCount="0" createdVersion="6" refreshedVersion="6">
  <cacheSource type="worksheet">
    <worksheetSource sheet="Dash" ref="A1:C1"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Date" numFmtId="14"/>
    <cacheField name="Sales" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>`;

// One slicer: caption includes an ASCII double quote so scenario 89 can
// spot-check that setAttribute HTML-encodes it (DOM serialisation uses
// `&quot;`). cache="Slicer_Region" binds to slicerCache1.xml below.
const slicer1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="SlicerRegion" cache="Slicer_Region" caption="Region &quot;A&quot;" columnCount="1" rowHeight="241300" style="SlicerStyleLight1"/>
</slicers>`;

// One slicer cache. sourceName="Region" identifies the pivot column this
// slicer drives; two items (North, South) are flagged selected via s="1".
// We include a third unselected item (East) to confirm the parser only
// emits the selected ones.
const slicerCache1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                       name="Slicer_Region" sourceName="Region">
  <pivotTables>
    <pivotTable tabId="1" name="PivotTable1"/>
  </pivotTables>
  <data>
    <tabular pivotCacheId="1">
      <items count="3">
        <i x="0" s="1" n="North"/>
        <i x="1" s="1" n="South"/>
        <i x="2" n="East"/>
      </items>
    </tabular>
  </data>
</slicerCacheDefinition>`;

// One timeline: months-level view, concrete start/end selection. The
// @level="months" form (rather than the numeric "2") exercises the
// normaliseTimelineLevel path that takes the string lexical form
// directly. startDate/endDate sit on the timeline element itself so the
// parser can read them without descending into the cache's <state>/<bounds>.
const timeline1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="TimelineDate" cache="NativeTimeline_Date" caption="Date" level="months" startDate="2023-01-01T00:00:00" endDate="2023-12-31T00:00:00" style="TimeSlicerStyleLight1"/>
</timelines>`;

const timelineCache1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         name="NativeTimeline_Date" sourceName="Date">
  <pivotTables>
    <pivotTable tabId="1" name="PivotTable1"/>
  </pivotTables>
  <state pivotCacheId="1" filterType="monthly">
    <bounds startDate="2023-01-01T00:00:00" endDate="2023-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`;

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
zip.file('xl/timelines/timeline1.xml', timeline1Xml);
zip.file('xl/timelineCaches/timelineCache1.xml', timelineCache1Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
