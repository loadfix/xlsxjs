// Build tests/render-test/smartart-matrix-pyramid/workbook.xlsx. Exercises
// the Wave 12 SmartArt native layout strategies: a matrix and a pyramid
// diagram anchored on a single sheet. Same two-diagram `<dgm:relIds>`
// layout as make-smartart-layouts-fixture.mjs — one drawing1.xml with two
// <xdr:twoCellAnchor> entries, and a drawing1.xml.rels carrying 8 rIds
// (4 per diagram, data/layout/quickStyle/colors).
//
// Diagram 1 — matrix (4 quadrants):
//
//   Revenue | Growth
//   --------+--------
//   Quality | Speed
//
// Diagram 2 — pyramid (4 stacked bands, narrow at top):
//
//         Vision
//       Strategy
//     Execution
//   Review
//
// layoutDef uniqueIds carry "matrix1" and "pyramid1" respectively so the
// renderer's regex dispatch picks the right strategy.
//
// Run:
//   node scripts/make-smartart-matrix-pyramid-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/smartart-matrix-pyramid');
mkdirSync(outDir, { recursive: true });

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.ms-office.drawingml.diagramData+xml"/>
  <Override PartName="/xl/diagrams/layout1.xml" ContentType="application/vnd.ms-office.drawingml.diagramLayout+xml"/>
  <Override PartName="/xl/diagrams/quickStyle1.xml" ContentType="application/vnd.ms-office.drawingml.diagramStyle+xml"/>
  <Override PartName="/xl/diagrams/colors1.xml" ContentType="application/vnd.ms-office.drawingml.diagramColors+xml"/>
  <Override PartName="/xl/diagrams/data2.xml" ContentType="application/vnd.ms-office.drawingml.diagramData+xml"/>
  <Override PartName="/xl/diagrams/layout2.xml" ContentType="application/vnd.ms-office.drawingml.diagramLayout+xml"/>
  <Override PartName="/xl/diagrams/quickStyle2.xml" ContentType="application/vnd.ms-office.drawingml.diagramStyle+xml"/>
  <Override PartName="/xl/diagrams/colors2.xml" ContentType="application/vnd.ms-office.drawingml.diagramColors+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Diagrams" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

// A tiny A1 cell keeps the sheet non-empty for the renderer's short-circuit
// check; the two SmartArt diagrams themselves are anchored above / below.
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>SmartArt matrix + pyramid</t></is></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

// Two twoCellAnchor entries. The first (matrix) anchors above B2:E8; the
// second (pyramid) anchors at B10:E16. Both use <dgm:relIds> pointing at
// their own set of 4 rIds in drawing1.xml.rels.
const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Diagram 1 (matrix)"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="Diagram 2 (pyramid)"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rId5" r:lo="rId6" r:qs="rId7" r:cs="rId8"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

// 8 rels — 4 per diagram.
const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout2.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle2.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors2.xml"/>
</Relationships>`;

// ── Diagram 1: matrix ─────────────────────────────────────────────────────
// 4 root-level nodes under a synthetic doc point: Revenue / Growth /
// Quality / Speed. The matrix renderer picks up the first 4 roots and
// lays them out as a 2×2 grid.
const data1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc"/>
    <dgm:pt modelId="1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Revenue</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="2"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Growth</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="3"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Quality</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="4"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Speed</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="100" type="parOf" srcId="0" destId="1" srcOrd="0" destOrd="0"/>
    <dgm:cxn modelId="101" type="parOf" srcId="0" destId="2" srcOrd="1" destOrd="0"/>
    <dgm:cxn modelId="102" type="parOf" srcId="0" destId="3" srcOrd="2" destOrd="0"/>
    <dgm:cxn modelId="103" type="parOf" srcId="0" destId="4" srcOrd="3" destOrd="0"/>
  </dgm:cxnLst>
</dgm:dataModel>`;

// layout1.xml — uniqueId contains "matrix1" so the renderer's /matrix/i
// regex dispatches to the matrix strategy. The parser only reads @uniqueId.
const layout1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/matrix1">
  <dgm:title val="Basic Matrix"/>
  <dgm:desc val="Basic matrix layout"/>
  <dgm:layoutNode name="matrix"/>
</dgm:layoutDef>`;

// ── Diagram 2: pyramid ────────────────────────────────────────────────────
// 4 root-level nodes under a synthetic doc point: Vision / Strategy /
// Execution / Review (top → bottom). The pyramid renderer stacks them as
// N trapezoids whose widths interpolate from narrow at top to wide at bottom.
const data2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc"/>
    <dgm:pt modelId="1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Vision</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="2"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Strategy</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="3"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Execution</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="4"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Review</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="200" type="parOf" srcId="0" destId="1" srcOrd="0" destOrd="0"/>
    <dgm:cxn modelId="201" type="parOf" srcId="0" destId="2" srcOrd="1" destOrd="0"/>
    <dgm:cxn modelId="202" type="parOf" srcId="0" destId="3" srcOrd="2" destOrd="0"/>
    <dgm:cxn modelId="203" type="parOf" srcId="0" destId="4" srcOrd="3" destOrd="0"/>
  </dgm:cxnLst>
</dgm:dataModel>`;

const layout2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/pyramid1">
  <dgm:title val="Basic Pyramid"/>
  <dgm:desc val="Basic pyramid layout"/>
  <dgm:layoutNode name="pyramid"/>
</dgm:layoutDef>`;

// Style / colours stubs (same for both diagrams).
const quickStyleXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
              uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1">
  <dgm:title val="Simple"/>
  <dgm:desc val="Simple style"/>
  <dgm:styleLbl name="node0"><dgm:style/></dgm:styleLbl>
</dgm:styleDef>`;

const colorsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_1">
  <dgm:title val="Colorful"/>
  <dgm:desc val="Colorful palette"/>
</dgm:colorsDef>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/worksheets/sheet1.xml', sheetXml);
zip.file('xl/worksheets/_rels/sheet1.xml.rels', sheetRels);
zip.file('xl/drawings/drawing1.xml', drawingXml);
zip.file('xl/drawings/_rels/drawing1.xml.rels', drawingRels);
zip.file('xl/diagrams/data1.xml', data1Xml);
zip.file('xl/diagrams/layout1.xml', layout1Xml);
zip.file('xl/diagrams/quickStyle1.xml', quickStyleXml);
zip.file('xl/diagrams/colors1.xml', colorsXml);
zip.file('xl/diagrams/data2.xml', data2Xml);
zip.file('xl/diagrams/layout2.xml', layout2Xml);
zip.file('xl/diagrams/quickStyle2.xml', quickStyleXml);
zip.file('xl/diagrams/colors2.xml', colorsXml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
