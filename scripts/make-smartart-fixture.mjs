// Build tests/render-test/smartart/workbook.xlsx. Targets the SmartArt
// detect-only surface in workbook-parser.ts + html-renderer.ts: a tiny
// workbook with a single SmartArt diagram (hierarchy1 layout) anchored at
// B2:E8 containing a 3-level tree:
//
//   Project (root)
//     Backend
//       API
//       Database
//     Frontend
//       UI
//
// The four diagram parts (data1.xml, layout1.xml, quickStyle1.xml,
// colors1.xml) are minimal-but-valid. Excel tolerates missing style /
// colour detail, but data1.xml's <dgm:dataModel>/<dgm:ptLst>/<dgm:cxnLst>
// are the load-bearing block for the hierarchy. layout1.xml is a stub with
// <dgm:layoutDef uniqueId="…hierarchy1"/> so the parser's layout name
// lookup succeeds.
//
// The outer <xdr:graphicFrame> is the standard DrawingML wrapping used by
// Excel: <a:graphic>/<a:graphicData uri=".../diagram"> with a
// <dgm:relIds r:dm r:lo r:qs r:cs/> binding the four diagram files.
//
// Run:
//   node scripts/make-smartart-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/smartart');
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
  <sheets><sheet name="Diagram" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

// A single A1 cell keeps the sheet visible in the rendered DOM (renderer
// short-circuits empty sheets). The SmartArt diagram itself is anchored at
// B2:E8; the A1 cell is only scaffolding so the drawing layer lands.
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>SmartArt below</t></is></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

// twoCellAnchor from B2 (col=1, row=1) to E8 (col=4, row=7). The
// <xdr:graphicFrame> wraps an <a:graphic>/<a:graphicData uri="…/diagram">
// with a <dgm:relIds> binding the four diagram files via rIds.
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
        <xdr:cNvPr id="2" name="Diagram 1"/>
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
</xdr:wsDr>`;

const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
</Relationships>`;

// data1.xml — the hierarchy model. Ids are integers; a single doc point
// anchors the tree, and parOf connections link:
//   doc → Project (srcOrd 0)
//   Project → Backend (srcOrd 0)
//   Project → Frontend (srcOrd 1)
//   Backend → API (srcOrd 0)
//   Backend → Database (srcOrd 1)
//   Frontend → UI (srcOrd 0)
const data1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc"/>
    <dgm:pt modelId="1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Project</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="2"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Backend</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="3"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Frontend</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="4"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>API</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="5"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Database</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="6"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>UI</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="100" type="parOf" srcId="0" destId="1" srcOrd="0" destOrd="0"/>
    <dgm:cxn modelId="101" type="parOf" srcId="1" destId="2" srcOrd="0" destOrd="0"/>
    <dgm:cxn modelId="102" type="parOf" srcId="1" destId="3" srcOrd="1" destOrd="0"/>
    <dgm:cxn modelId="103" type="parOf" srcId="2" destId="4" srcOrd="0" destOrd="0"/>
    <dgm:cxn modelId="104" type="parOf" srcId="2" destId="5" srcOrd="1" destOrd="0"/>
    <dgm:cxn modelId="105" type="parOf" srcId="3" destId="6" srcOrd="0" destOrd="0"/>
  </dgm:cxnLst>
</dgm:dataModel>`;

// layout1.xml — a stub pointing at the built-in hierarchy1 layout. Excel
// tolerates missing layoutNode content; only the uniqueId attribute matters
// for xlsxjs to surface the layout name on the model.
const layout1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
               uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1">
  <dgm:title val="Hierarchy"/>
  <dgm:desc val="Hierarchy layout"/>
  <dgm:layoutNode name="hierChild"/>
</dgm:layoutDef>`;

// quickStyle1.xml / colors1.xml — minimal stubs. Excel is happy with an
// empty styleDef / colorsDef as long as the root element is valid.
const quickStyle1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
              uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1">
  <dgm:title val="Simple"/>
  <dgm:desc val="Simple style"/>
  <dgm:styleLbl name="node0"><dgm:style/></dgm:styleLbl>
</dgm:styleDef>`;

const colors1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
zip.file('xl/diagrams/quickStyle1.xml', quickStyle1Xml);
zip.file('xl/diagrams/colors1.xml', colors1Xml);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
