// Build tests/render-test/charts/workbook.xlsx — three hand-built classic
// c:chartSpace charts anchored on a single sheet:
//
//   · chart1 — cluster bar / column: "Revenue by Quarter"
//       2 series (2023 / 2024) × 4 categories (Q1-Q4)
//   · chart2 — line: "Website Traffic"
//       1 series × 6 categories (Jan-Jun)
//   · chart3 — pie: "Market Share"
//       1 series × 4 slices (Apple, Google, Microsoft, Others)
//
// The chart XML is hand-written to the SpreadsheetML chart schema (ECMA-376
// Part 1 §21.2). Only the elements xlsxjs renders are populated — we skip
// axId axes, data-label styling, spPr overlays beyond srgbClr, and the
// printer-settings paraphernalia. Excel / LibreOffice will still open the
// file, and xlsxjs's parser walks exactly the elements we emit.
//
// Run:
//   node scripts/make-charts-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/charts');
mkdirSync(outDir, { recursive: true });

function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Chart data ────────────────────────────────────────────────────────────
const bar = {
    title: 'Revenue by Quarter',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
        { name: '2023', values: [120, 150, 180, 210], color: '5B9BD5' },
        { name: '2024', values: [140, 170, 200, 260], color: 'ED7D31' },
    ],
};

const line = {
    title: 'Website Traffic',
    categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
    series: [
        { name: 'Visitors', values: [1200, 1450, 1610, 1580, 1900, 2300], color: '70AD47' },
    ],
};

const pie = {
    title: 'Market Share',
    categories: ['Apple', 'Google', 'Microsoft', 'Others'],
    series: [
        { name: 'Share', values: [40, 30, 20, 10] },
    ],
};

// Helpers for the c:chartSpace XML. We produce one <c:ser> per series and
// cache the cat / val in strCache / numCache; that's what Excel writes and
// what xlsxjs's parser reads.
function strCache(pts) {
    const lines = pts.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('');
    return `<c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="${pts.length}"/>${lines}</c:strCache></c:strRef>`;
}
function numCache(pts) {
    const lines = pts.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
    return `<c:numRef><c:f>Sheet1!$A$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${pts.length}"/>${lines}</c:numCache></c:numRef>`;
}
function seriesTx(name) {
    return `<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
}
function seriesSpPr(color) {
    if (!color) return '';
    return `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
}

// Wrap the full chartSpace with the necessary namespaces.
const NS_ATTRS =
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function titleXml(title) {
    return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}
function legendXml(pos) {
    return `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/></c:legend>`;
}

function makeBarChart({ title, categories, series }) {
    const sers = series.map((s, i) => `
      <c:ser>
        <c:idx val="${i}"/>
        <c:order val="${i}"/>
        ${seriesTx(s.name)}
        ${seriesSpPr(s.color)}
        <c:cat>${strCache(categories)}</c:cat>
        <c:val>${numCache(s.values)}</c:val>
      </c:ser>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS_ATTRS}>
  <c:chart>
    ${titleXml(title)}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        ${sers}
        <c:gapWidth val="150"/>
        <c:overlap val="-20"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    ${legendXml('r')}
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function makeLineChart({ title, categories, series }) {
    const sers = series.map((s, i) => `
      <c:ser>
        <c:idx val="${i}"/>
        <c:order val="${i}"/>
        ${seriesTx(s.name)}
        ${seriesSpPr(s.color)}
        <c:cat>${strCache(categories)}</c:cat>
        <c:val>${numCache(s.values)}</c:val>
        <c:smooth val="0"/>
      </c:ser>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS_ATTRS}>
  <c:chart>
    ${titleXml(title)}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        ${sers}
        <c:marker val="1"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    ${legendXml('r')}
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function makePieChart({ title, categories, series }) {
    const s = series[0];
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS_ATTRS}>
  <c:chart>
    ${titleXml(title)}
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          ${seriesTx(s.name)}
          <c:cat>${strCache(categories)}</c:cat>
          <c:val>${numCache(s.values)}</c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    ${legendXml('r')}
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

const chart1Xml = makeBarChart(bar);
const chart2Xml = makeLineChart(line);
const chart3Xml = makePieChart(pie);

// ── Drawing XML anchoring the three chart parts ──────────────────────────
// Anchor each chart at a distinct column so they stack top-to-bottom when
// the SVG fills 100% width in the live viewer.
const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>29</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 2"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>30</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>44</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="Chart 3"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId3"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

// ── Container XML: sheet + workbook + rels ───────────────────────────────
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
  <sheets><sheet name="Charts" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Charts sheet</t></si>
</sst>`;

const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
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
zip.file('xl/charts/chart1.xml', chart1Xml);
zip.file('xl/charts/chart2.xml', chart2Xml);
zip.file('xl/charts/chart3.xml', chart3Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
