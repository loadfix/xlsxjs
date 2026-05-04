// Build tests/render-test/ole-cfb/workbook.xlsx. Exercises the Wave-11 OLE
// CFB reader: the embedded `.bin` payload under xl/embeddings/oleObject1.bin
// is a real, minimal-but-valid OLE Compound File Binary container with three
// streams (\x01Ole, CONTENTS, meta) so the parser's directory walk has
// something to enumerate and the per-stream byte round-trip has stable
// expectations.
//
// Outer XLSX plumbing mirrors scripts/make-ole-embeddings-fixture.mjs — one
// sheet ("CFB") with a single <oleObject r:id="rId2" progId="Word.Document.12"
// shapeId="1025"/> anchored at B2:D4. The outer package is byte-stable when
// the inner CFB is byte-stable; JSZip writes stable headers for fixed input.
//
// Run:
//   node scripts/make-ole-cfb-fixture.mjs
//
// ── CFB layout (hand-crafted, 4096 bytes, 512-byte sectors = major=3) ──────
// The file is 8 sectors (header + 7 × 512):
//   sector -1: header (bytes 0x000-0x1FF)
//   sector  0: FAT   (covers all 8 data sectors + header)
//   sector  1: directory stream (4 entries × 128 bytes = 512 bytes)
//   sector  2: mini FAT (lists the mini-sectors we use)
//   sector  3: mini stream (carries CONTENTS + meta payloads at 64 B each)
//   sector  4: \x01Ole regular-FAT stream (512 B, zeroed)
//   sector  5..: unused / free
//
// Only three streams are actually surfaced:
//   · "\x01Ole"   — 20 bytes ([0x01, 0x00, 0x00, 0x00, 0x00, …] — the
//                  canonical OLE v1 header). Small-stream path → mini-FAT.
//                  Per MS-CFB §2.6, any stream below the 4096-byte cutoff
//                  MUST live in the mini-stream; storing \x01Ole there
//                  matches real-world `.bin` files that Word / Excel emit.
//   · "CONTENTS"  — the ASCII string "hello cfb" (9 bytes), in a mini
//                  sector. Small-stream path exercises the mini FAT walk.
//   · "meta"      — 5 fixed bytes [0x01,0x02,0x03,0x04,0x05], mini sector.
//
// The extra sector 4 (once used for a full-sized \x01Ole) is now a free
// spare — the reader ignores free sectors so we don't need to shrink the
// file. Keeping 4096 bytes makes it easy to extend the fixture later.
//
// The root entry advertises a fixed CLSID so `cfb.clsid` round-trips. The
// chosen value is deliberately NOT a registered Word/Excel CLSID — just a
// stable placeholder (`{0001F5EE-0000-0000-C000-000000000046}`) so the test
// assertion isn't sensitive to Microsoft product identity drift.

import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const outDir = resolve(repo, 'tests/render-test/ole-cfb');
mkdirSync(outDir, { recursive: true });

// ── Hand-crafted CFB ──────────────────────────────────────────────────────
const SECTOR_SIZE = 512;
const MINI_SECTOR_SIZE = 64;
const NUM_SECTORS = 7; // data sectors (in addition to the 512-byte header)
const FILE_SIZE = SECTOR_SIZE * (1 + NUM_SECTORS); // header + 7 × 512 = 4096
// FAT sector index (in the data-sector space, 0-based where 0 is "first
// sector after the header").
const FAT_SECTOR   = 0;
const DIR_SECTOR   = 1;
const MFAT_SECTOR  = 2;
const MINI_SECTOR  = 3; // the sector holding the mini-stream bytes

// Sector allocation sentinels (MS-CFB §2.1).
const FREESECT   = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT    = 0xfffffffd;

const file = new Uint8Array(FILE_SIZE);
const view = new DataView(file.buffer);

// ── Header (sector -1, bytes 0x000-0x1FF) ─────────────────────────────────
file.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);   // magic
// CLSID of root (bytes 8-23): left zero (we put the root CLSID in the dir entry).
view.setUint16(0x18, 0x003e, true);  // minor version
view.setUint16(0x1a, 0x0003, true);  // major version  (3 → 512-byte sectors)
view.setUint16(0x1c, 0xfffe, true);  // byte order (little-endian marker)
view.setUint16(0x1e, 9, true);       // sector shift  (2^9  = 512)
view.setUint16(0x20, 6, true);       // mini-sector shift (2^6 = 64)
// Reserved 6 bytes (0x22-0x27) stay zero.
view.setUint32(0x28, 0, true);        // number of directory sectors (0 for v3)
view.setUint32(0x2c, 1, true);        // number of FAT sectors
view.setUint32(0x30, DIR_SECTOR, true);  // first directory sector
view.setUint32(0x34, 0, true);        // transaction sig
view.setUint32(0x38, 4096, true);     // mini-stream cutoff (standard)
view.setUint32(0x3c, MFAT_SECTOR, true);  // first mini-FAT sector
view.setUint32(0x40, 1, true);        // number of mini-FAT sectors
view.setUint32(0x44, ENDOFCHAIN, true);   // first DIFAT sector
view.setUint32(0x48, 0, true);        // number of DIFAT sectors
// DIFAT[0] = FAT_SECTOR; DIFAT[1..108] = FREESECT (already 0? no — 0 is a
// valid sector id, so fill explicitly with FREESECT).
view.setUint32(0x4c, FAT_SECTOR, true);
for (let i = 1; i < 109; i++) view.setUint32(0x4c + i * 4, FREESECT, true);

// ── FAT sector (data sector index FAT_SECTOR) ────────────────────────────
// Each FAT entry maps data sector id → next sector in chain (or a sentinel).
const fatOff = SECTOR_SIZE * (FAT_SECTOR + 1); // +1 to skip header
for (let i = 0; i < SECTOR_SIZE / 4; i++) view.setUint32(fatOff + i * 4, FREESECT, true);
// Sector 0 (the FAT itself) is flagged with FATSECT.
view.setUint32(fatOff + FAT_SECTOR   * 4, FATSECT, true);
// Sector 1 (directory) — single-sector chain; terminate.
view.setUint32(fatOff + DIR_SECTOR   * 4, ENDOFCHAIN, true);
// Sector 2 (mini-FAT) — single sector; terminate.
view.setUint32(fatOff + MFAT_SECTOR  * 4, ENDOFCHAIN, true);
// Sector 3 (mini-stream payload) — single sector; terminate.
view.setUint32(fatOff + MINI_SECTOR  * 4, ENDOFCHAIN, true);
// Sectors 4+ stay FREESECT (already initialised above).

// ── Directory stream (sector DIR_SECTOR) ─────────────────────────────────
// Four 128-byte entries: Root, \x01Ole, CONTENTS, meta.
const dirOff = SECTOR_SIZE * (DIR_SECTOR + 1);
function writeDirEntry(index, { name, type, startSector, size, clsid, leftSib = 0xffffffff, rightSib = 0xffffffff, child = 0xffffffff }) {
    const base = dirOff + index * 128;
    // Zero-fill the entry first (names + sibling IDs + clsid default to 0xff…
    // or 0x00 per the spec; we set the fields we care about explicitly).
    for (let i = 0; i < 128; i++) file[base + i] = 0;
    // Name: UTF-16LE, null-terminated.
    const nameBytes = new Uint8Array(64);
    let bytesWritten = 0;
    for (let i = 0; i < name.length; i++) {
        const cp = name.charCodeAt(i);
        nameBytes[bytesWritten]     = cp & 0xff;
        nameBytes[bytesWritten + 1] = (cp >> 8) & 0xff;
        bytesWritten += 2;
    }
    // Terminator (already zero via allocation).
    bytesWritten += 2;
    file.set(nameBytes.subarray(0, 64), base);
    view.setUint16(base + 0x40, bytesWritten, true);
    file[base + 0x42] = type;
    file[base + 0x43] = 0x01; // color flag (black); unused by our reader
    view.setUint32(base + 0x44, leftSib, true);
    view.setUint32(base + 0x48, rightSib, true);
    view.setUint32(base + 0x4c, child, true);
    if (clsid) {
        const clsidBytes = parseClsid(clsid);
        file.set(clsidBytes, base + 0x50);
    }
    view.setUint32(base + 0x74, startSector, true);
    view.setUint32(base + 0x78, size, true);        // size low 32 bits
    view.setUint32(base + 0x7c, 0, true);            // size high (v3: reserved zero)
}

// Parse "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" into the 16 raw bytes the
// CFB stores (groups 1-3 little-endian, groups 4-5 big-endian).
function parseClsid(str) {
    const clean = str.replace(/[{}-]/g, '');
    if (clean.length !== 32) throw new Error('bad clsid');
    const raw = new Uint8Array(16);
    for (let i = 0; i < 16; i++) raw[i] = parseInt(clean.slice(i * 2, i * 2 + 2), 16);
    // Swap groups 1-3 (first 8 bytes) for little-endian storage:
    //   group1 = raw[0..4]    → reverse first 4
    //   group2 = raw[4..6]    → reverse next 2
    //   group3 = raw[6..8]    → reverse next 2
    const out = new Uint8Array(16);
    out[0] = raw[3]; out[1] = raw[2]; out[2] = raw[1]; out[3] = raw[0];
    out[4] = raw[5]; out[5] = raw[4];
    out[6] = raw[7]; out[7] = raw[6];
    for (let i = 8; i < 16; i++) out[i] = raw[i];
    return out;
}

// Root entry — index 0. startSector = MINI_SECTOR (first sector of the
// mini-stream). size = total mini-stream size = one sector (64 × 8 = 512
// bytes — we only use the first two mini-sectors).
writeDirEntry(0, {
    name: 'Root Entry',
    type: 0x05,                          // root
    startSector: MINI_SECTOR,
    size: SECTOR_SIZE,                   // mini-stream spans one sector
    clsid: '{0001F5EE-0000-0000-C000-000000000046}',
    child: 1,                             // first child = \x01Ole
});
// \x01Ole — index 1. Small stream (20 bytes) → mini-sector 0.
writeDirEntry(1, {
    name: '\x01Ole',
    type: 0x02,                          // stream
    startSector: 0,                      // mini-sector 0
    size: 20,
    rightSib: 2,                         // red-black ordering — reader walks linearly; ignored.
});
// CONTENTS — index 2. Small stream → mini-sector 1. Size = 9 ("hello cfb").
writeDirEntry(2, {
    name: 'CONTENTS',
    type: 0x02,
    startSector: 1,                      // mini-sector 1
    size: 9,
    rightSib: 3,
});
// meta — index 3. Small stream → mini-sector 2. Size = 5 bytes.
writeDirEntry(3, {
    name: 'meta',
    type: 0x02,
    startSector: 2,                      // mini-sector 2
    size: 5,
});

// ── Mini-FAT sector (sector MFAT_SECTOR) ─────────────────────────────────
// Entry per mini-sector. We use mini-sectors 0 (\x01Ole), 1 (CONTENTS),
// 2 (meta). All three are single-mini-sector streams → ENDOFCHAIN.
const mfatOff = SECTOR_SIZE * (MFAT_SECTOR + 1);
for (let i = 0; i < SECTOR_SIZE / 4; i++) view.setUint32(mfatOff + i * 4, FREESECT, true);
view.setUint32(mfatOff + 0, ENDOFCHAIN, true);  // \x01Ole  mini-sector 0
view.setUint32(mfatOff + 4, ENDOFCHAIN, true);  // CONTENTS mini-sector 1
view.setUint32(mfatOff + 8, ENDOFCHAIN, true);  // meta     mini-sector 2

// ── Mini-stream payload (sector MINI_SECTOR) ─────────────────────────────
// Mini-sectors are 64 bytes each. Place \x01Ole at mini-sector 0, CONTENTS
// at mini-sector 1, meta at mini-sector 2.
const miniOff = SECTOR_SIZE * (MINI_SECTOR + 1);
// \x01Ole header: the OLE v1 "native" header Word / Excel write. Byte 0 is
// version (0x01), bytes 4..7 are flags (0 = embedded object). The remaining
// 12 bytes are reserved zero. Real producers write the linked-source moniker
// here; we leave it zero — consumers walking this stream should see a valid
// OLE v1 embedded-object header that decodes to "embedded, unlinked".
file.set(new Uint8Array([
    0x01, 0x00, 0x00, 0x02,  // version (1, 0x02000001 little-endian, flags in upper 3 bytes)
    0x00, 0x00, 0x00, 0x00,  // link-update options (0)
    0x00, 0x00, 0x00, 0x00,  // reserved
    0x00, 0x00, 0x00, 0x00,  // moniker stream length (0)
    0x00, 0x00, 0x00, 0x00,  // moniker stream clsid (partial)
]), miniOff + 0 * MINI_SECTOR_SIZE);
const contents = new TextEncoder().encode('hello cfb');
file.set(contents, miniOff + 1 * MINI_SECTOR_SIZE);
file.set(new Uint8Array([0x01, 0x02, 0x03, 0x04, 0x05]), miniOff + 2 * MINI_SECTOR_SIZE);

const OLE_BIN = file;

// Sanity: the hand-crafted buffer must round-trip through parseCfb at
// build-time if the caller wants to verify. We don't invoke parseCfb here
// (it lives under src/), but any assertion failure will surface in the
// test-render scenarios.

// ── Outer XLSX package — mirrors make-ole-embeddings-fixture.mjs ─────────
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
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

const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="CFB" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>CFB Word:</t></si>
</sst>`;

const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <sheetData>
    <row r="2"><c r="A2" t="s"><v>0</v></c></row>
  </sheetData>
  <oleObjects>
    <oleObject progId="Word.Document.12" shapeId="1025" r:id="rId2">
      <objectPr defaultSize="0" r:id="rId10">
        <anchor>
          <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        </anchor>
      </objectPr>
    </oleObject>
  </oleObjects>
</worksheet>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', contentTypes);
zip.file('_rels/.rels', rootRels);
zip.file('xl/_rels/workbook.xml.rels', workbookRels);
zip.file('xl/workbook.xml', workbookXml);
zip.file('xl/sharedStrings.xml', sharedStringsXml);
zip.file('xl/worksheets/sheet1.xml', sheetXml);
zip.file('xl/worksheets/_rels/sheet1.xml.rels', sheetRels);
zip.file('xl/embeddings/oleObject1.bin', OLE_BIN);

const buf = await zip.generateAsync({ type: 'nodebuffer' });
const out = resolve(outDir, 'workbook.xlsx');
writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes; OLE .bin = ${OLE_BIN.length} bytes)`);
