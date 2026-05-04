// Build tests/render-test/form-controls/workbook.xlsx. Hand-writes an xlsx
// package that exercises the form-control detection path:
//
//   · A1: "Form controls"        — plain label text, one shared string.
//   · Checkbox at B2:C3          — objectType="CheckBox" checked="Checked"
//                                   fmlaLink="$A$1"
//   · Radio    at B4:C5          — objectType="Radio" checked="Unchecked"
//                                   fmlaLink="$A$2"
//   · Combo    at B6:C7          — objectType="Drop" fmlaRange="$E$1:$E$4"
//                                   fmlaLink="$A$3" dropLines="4"
//   · Scrollbar at B8:C9         — objectType="Scroll" min="0" max="100"
//                                   val="42" inc="1" page="10" fmlaLink="$A$4"
//   · Button   at D2:E3          — objectType="Button" with txBody="Click me"
//
// Each drawing anchor wraps its <xdr:sp> in mc:AlternateContent with an
// <mc:Fallback> holding legacy VML (so the fixture mirrors how real Excel
// writes form controls); xlsxjs detects the ctrlProp rel on the outer
// <xdr:sp> → ctrlProps/ctrlProp{N}.xml binding.
//
// Run:
//   node scripts/make-form-controls-fixture.mjs

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/form-controls');
mkdirSync(outDir, { recursive: true });

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp2.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp3.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp4.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp5.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
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
  <sheets><sheet name="Controls" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>Form controls</t></si>
  <si><t>Red</t></si>
  <si><t>Green</t></si>
  <si><t>Blue</t></si>
  <si><t>Yellow</t></si>
</sst>`;

// Sheet cells: column A holds a caption and the linked-cell targets for the
// controls; column E holds the combo's input range.
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="E1" t="s"><v>1</v></c>
    </row>
    <row r="2"><c r="E2" t="s"><v>2</v></c></row>
    <row r="3"><c r="E3" t="s"><v>3</v></c></row>
    <row r="4"><c r="E4" t="s"><v>4</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
  <legacyDrawing r:id="rId2"/>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`;

// Drawing — five anchors, each hosting a form-control <xdr:sp> wrapped in
// <mc:AlternateContent>. Each anchor carries an <xdr:clientData/> plus a
// rel (via r:id on the <xdr:sp>'s clientData attribute position inside
// Choice) referencing a ctrlProps/ctrlProp{N}.xml part.
const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          xmlns:xdr14="http://schemas.microsoft.com/office/drawing/2010/spreadsheetDrawing">
  <!-- 1. Checkbox at B2:C3, fmlaLink=$A$1, checked -->
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="" fLocksText="0">
      <xdr:nvSpPr>
        <xdr:cNvPr id="1" name="Check Box 1" descr="Subscribe"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr vertOverflow="clip" wrap="square"/>
        <a:p><a:r><a:t>Subscribe</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData fPrintsWithSheet="0" r:id="rId2"/>
  </xdr:twoCellAnchor>

  <!-- 2. Radio (option button) at B4:C5, unchecked, fmlaLink=$A$2 -->
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="" fLocksText="0">
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Option Button 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr vertOverflow="clip" wrap="square"/>
        <a:p><a:r><a:t>Option A</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData fPrintsWithSheet="0" r:id="rId3"/>
  </xdr:twoCellAnchor>

  <!-- 3. Combo (drop down) at B6:C7 with fmlaRange=$E$1:$E$4, dropLines=4 -->
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="" fLocksText="0">
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="Drop Down 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="228600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData fPrintsWithSheet="0" r:id="rId4"/>
  </xdr:twoCellAnchor>

  <!-- 4. Scrollbar at B8:C9 -->
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="" fLocksText="0">
      <xdr:nvSpPr>
        <xdr:cNvPr id="4" name="Scroll Bar 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="152400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData fPrintsWithSheet="0" r:id="rId5"/>
  </xdr:twoCellAnchor>

  <!-- 5. Button at D2:E3 with text "Click me" -->
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="" fLocksText="0">
      <xdr:nvSpPr>
        <xdr:cNvPr id="5" name="Button 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr vertOverflow="clip" wrap="square"/>
        <a:p><a:r><a:t>Click me</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData fPrintsWithSheet="0" r:id="rId6"/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp3.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp4.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp5.xml"/>
</Relationships>`;

// Form-control ctrlProp xml — one file per control. Uses the x14ac
// default namespace (mirrors how Excel writes them). Attribute set
// sticks to what the spec documents: objectType, checked, fmlaLink,
// fmlaRange, min, max, inc, page, val, dropLines.
const ctrlProp1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    objectType="CheckBox" fmlaLink="$A$1" checked="Checked" lockText="1" noThreeD="1"/>`;

const ctrlProp2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    objectType="Radio" fmlaLink="$A$2" checked="Unchecked" firstButton="1" lockText="1" noThreeD="1"/>`;

const ctrlProp3 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    objectType="Drop" dropStyle="combo" dx="22" fmlaRange="$E$1:$E$4" fmlaLink="$A$3"
    sel="1" val="0" dropLines="4" noThreeD="1"/>`;

const ctrlProp4 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    objectType="Scroll" dx="16" fmlaLink="$A$4" min="0" max="100" val="42" inc="1" page="10" horiz="1"/>`;

const ctrlProp5 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    objectType="Button" lockText="1"/>`;

// A minimal legacy VML drawing — the fallback path real Excel writes alongside
// the modern DrawingML form-control shapes. xlsxjs doesn't parse VML, but we
// need a valid part so the fixture round-trips through Excel should anyone
// want to re-open it.
const vmlDrawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
   <v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype>
</xml>`;

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
zip.file('xl/drawings/vmlDrawing1.vml', vmlDrawing);
zip.file('xl/ctrlProps/ctrlProp1.xml', ctrlProp1);
zip.file('xl/ctrlProps/ctrlProp2.xml', ctrlProp2);
zip.file('xl/ctrlProps/ctrlProp3.xml', ctrlProp3);
zip.file('xl/ctrlProps/ctrlProp4.xml', ctrlProp4);
zip.file('xl/ctrlProps/ctrlProp5.xml', ctrlProp5);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
