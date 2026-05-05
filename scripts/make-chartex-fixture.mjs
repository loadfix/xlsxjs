// Build tests/render-test/chartex/workbook.xlsx — a hand-built workbook
// carrying TWO chartEx (cx:chartSpace, Office 2014) charts anchored on a
// single sheet:
//
//   · chartEx1 — treemap:   "Budget by Region"
//       5 leaf cells (Alpha, Beta, Gamma, Delta, Epsilon) with allocations
//   · chartEx2 — waterfall: "Q1 Revenue Bridge"
//       4 categories (Opening, Sales, Returns, Closing) × 1 numeric series
//
// chartEx differs from classic chart XML in two material ways: (1) the
// root element sits in the cx: namespace (Office 2014 extensions), and
// (2) the drawing relationship type is the
// http://schemas.microsoft.com/office/2014/relationships/chartEx variant,
// not the classic /chart one. Excel / LibreOffice will both open the file;
// xlsxjs's parser routes the root-element namespace through the chartEx
// branch and surfaces the subtype via the `<cx:series layoutId="…">`
// attribute.
//
// Run:
//   node scripts/make-chartex-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/chartex');
mkdirSync(outDir, { recursive: true });

function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Chart data ────────────────────────────────────────────────────────────
const treemap = {
    title: 'Budget by Region',
    layoutId: 'treemap',
    categories: ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'],
    values: [120, 80, 60, 40, 25],
};

const waterfall = {
    title: 'Q1 Revenue Bridge',
    layoutId: 'waterfall',
    categories: ['Opening', 'Sales', 'Returns', 'Closing'],
    values: [1000, 450, -120, 1330],
};

// Wrap the full cx:chartSpace with the necessary namespaces. `cx` is the
// 2014 chartEx root; `a` pulls in DrawingML for any rich-text titles; `r`
// covers relationship attributes on series/data references (unused here
// because we embed literal data, but included for validator politeness).
const NS_ATTRS =
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function titleXml(title) {
    return `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich>`
        + `<a:bodyPr spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>`
        + `<a:lstStyle/>`
        + `<a:p><a:pPr algn="ctr"/><a:r><a:t>${esc(title)}</a:t></a:r></a:p>`
        + `</cx:rich></cx:tx></cx:title>`;
}

// chartEx data lives at <cx:chartData>/<cx:data id="0">. strDim carries
// the category strings; numDim carries the numeric values. Our parser
// walks both and surfaces categories[] + series[0].values[].
function dataXml(categories, values) {
    const catPts = categories.map((c, i) => `<cx:pt idx="${i}">${esc(c)}</cx:pt>`).join('');
    const valPts = values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('');
    return `<cx:chartData>`
        + `<cx:data id="0">`
        + `<cx:strDim type="cat"><cx:f dir="col">Sheet1!$A$1:$A$${categories.length}</cx:f>${catPts}</cx:strDim>`
        + `<cx:numDim type="val"><cx:f dir="col">Sheet1!$B$1:$B$${values.length}</cx:f>${valPts}</cx:numDim>`
        + `</cx:data>`
        + `</cx:chartData>`;
}

// One <cx:series> per chart. layoutId picks the subtype; dataId points at
// the data block we just emitted (id="0"). The tx/txData/v pair carries
// the series name, which xlsxjs lifts into ChartSeries.name.
function seriesXml({ layoutId, name }) {
    return `<cx:series layoutId="${esc(layoutId)}" hidden="0" ownerIdx="0">`
        + `<cx:tx><cx:txData><cx:v>${esc(name)}</cx:v></cx:txData></cx:tx>`
        + `<cx:dataId val="0"/>`
        + `</cx:series>`;
}

function buildChartEx({ title, layoutId, categories, values }) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace ${NS_ATTRS}>
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat">
        <cx:f dir="col">Sheet1!$A$1:$A$${categories.length}</cx:f>
        ${categories.map((c, i) => `<cx:pt idx="${i}">${esc(c)}</cx:pt>`).join('\n        ')}
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f dir="col">Sheet1!$B$1:$B$${values.length}</cx:f>
        ${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('\n        ')}
      </cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart>
    ${titleXml(title)}
    <cx:plotArea>
      <cx:plotAreaRegion>
        ${seriesXml({ layoutId, name: title })}
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;
    // dataXml is defined above but inlined here so the one-and-only data
    // block sits before <cx:chart> (ECMA-376 schema order).
    void dataXml;
}

const chartEx1Xml = buildChartEx(treemap);
const chartEx2Xml = buildChartEx(waterfall);

// ── Drawing XML anchoring the two chartEx parts ──────────────────────────
// The chartEx graphicData uri is the same drawingml/2006/chart uri (shared
// with classic charts — Excel distinguishes by the inner element's
// namespace: <cx:chart r:id="…"/>). A matching mc:AlternateContent is
// optional in practice; xlsxjs's detector walks either shape via
// getElementsByTagNameNS, so we skip the AlternateContent dance and emit
// the cx:chart directly.
const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>29</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 2"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
          <cx:chart r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

// ── Container XML ────────────────────────────────────────────────────────
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>
  <Override PartName="/xl/charts/chartEx2.xml" ContentType="application/vnd.ms-office.chartex+xml"/>
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
  <sheets><sheet name="ChartEx" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>ChartEx sheet</t></si>
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

// chartEx uses its own relationship type (the 2014 namespace variant) —
// NOT the classic /chart type. This is what xlsxjs looks at to decide
// whether to dispatch the chart part through the classic or chartEx
// parser branch.
const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx2.xml"/>
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
zip.file('xl/charts/chartEx1.xml', chartEx1Xml);
zip.file('xl/charts/chartEx2.xml', chartEx2Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
