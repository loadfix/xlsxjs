// Minimal OLE Compound File Binary (CFB) reader.
//
// CFB (a.k.a. "structured storage", MS-CFB) is the mini-filesystem format
// OLE uses to bundle multiple "streams" (files) into a single blob. An
// embedded `.bin` payload attached to a `<oleObject r:id="…"/>` inside an
// XLSX is a CFB container: its streams carry the actual Word / Excel /
// equation / formula content. xlsxjs does not decode those streams — this
// reader only enumerates them so consumers can route the payload to a
// downstream renderer (python-docx, Excel equation renderer, etc.).
//
// Layout (all little-endian):
//   · 512-byte header
//       0x00  8   signature      0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1
//       0x18  2   minor version
//       0x1A  2   major version  (3 for 512-byte sectors, 4 for 4096-byte)
//       0x1C  2   byte order     (0xFFFE, little-endian marker)
//       0x1E  2   sector shift   (9 → 512-byte, 12 → 4096-byte)
//       0x20  2   mini sector shift (usually 6 → 64 bytes)
//       0x2C  4   number of directory sectors
//       0x30  4   number of FAT sectors
//       0x34  4   first directory sector
//       0x38  4   transaction signature
//       0x3C  4   mini stream cutoff (usually 4096)
//       0x40  4   first mini FAT sector
//       0x44  4   number of mini FAT sectors
//       0x48  4   first DIFAT sector (or ENDOFCHAIN)
//       0x4C  4   number of DIFAT sectors
//       0x50 436  DIFAT[109] — the first 109 FAT sector indices
//
//   · FAT: one 4-byte entry per sector in the file. Chain by reading
//     FAT[sectorId] repeatedly until the entry is 0xFFFFFFFE (ENDOFCHAIN).
//     Special values: FREESECT=0xFFFFFFFF, DIFSECT=0xFFFFFFFC,
//     FATSECT=0xFFFFFFFD.
//
//   · Directory stream: fixed 128-byte entries.
//       0x00 64   UTF-16LE name, '\0'-terminated
//       0x40  2   name length in bytes (including terminator)
//       0x42  1   object type (0=unused, 1=storage, 2=stream, 5=root)
//       0x43  1   color flag (red/black tree — ignored)
//       0x44  4   left sibling dir id
//       0x48  4   right sibling dir id
//       0x4C  4   child dir id
//       0x50 16   CLSID (only meaningful on the root / storage entries)
//       0x64  4   state bits
//       0x68  8   creation time (FILETIME)
//       0x70  8   modified time (FILETIME)
//       0x78  4   starting sector
//       0x7C  8   stream size (bytes)
//
//   · Mini-FAT: a parallel FAT but for streams smaller than the cutoff
//     (≤ 4096 bytes by default). The mini-stream itself is a regular FAT-
//     chained stream whose first sector lives in the root directory entry.
//
// Scope / guardrails:
//   · Hard cap on stream count (MAX_STREAMS) to stop an adversarial CFB
//     from looping billions of directory entries.
//   · Hard cap on per-stream size (MAX_STREAM_BYTES). 16 MiB is well past
//     any legitimate Word document table and still safely below
//     MAX_EMBEDDING_BYTES (32 MiB, the outer buffer cap).
//   · Reject the whole buffer up-front if it exceeds MAX_EMBEDDING_BYTES.
//   · Never throw — every malformed path returns null.
//   · Walk the FAT / miniFAT with a visited set so a chain that loops back
//     to a sector we've already seen terminates instead of spinning.

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

// Sector allocation sentinels (MS-CFB §2.1).
const FREESECT     = 0xffffffff;
const ENDOFCHAIN   = 0xfffffffe;
// Used in FAT to mark a sector that holds FAT data itself. We don't interpret
// these specially — we just stop the chain walk when we see something that is
// not a valid forward pointer.
const FATSECT      = 0xfffffffd;
const DIFSECT      = 0xfffffffc;

// Directory entry object types (MS-CFB §2.6.1).
const DIR_TYPE_UNUSED  = 0x00;
const DIR_TYPE_STORAGE = 0x01;
const DIR_TYPE_STREAM  = 0x02;
const DIR_TYPE_ROOT    = 0x05;

// Hard caps — see guardrail note above.
const MAX_STREAMS = 256;
const MAX_STREAM_BYTES = 16 * 1024 * 1024; // 16 MiB
const MAX_BUFFER_BYTES = 32 * 1024 * 1024; // 32 MiB — mirrors MAX_EMBEDDING_BYTES in workbook.ts
// Cap chain walks so a pathological loop / runaway chain can never exceed
// the possible sector count of a 32 MiB container.
const MAX_CHAIN_SECTORS = 1 << 20; // 1,048,576 sectors

export interface CfbStream {
    name: string;
    bytes: Uint8Array;
}

export interface CfbModel {
    // Class ID of the root directory entry, formatted as
    // "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}". Null when the root CLSID is
    // all zeros (the common "no class" case) or the container doesn't have a
    // well-formed root entry.
    clsid: string | null;
    streams: CfbStream[];
}

export function parseCfb(bytes: Uint8Array): CfbModel | null {
    if (!bytes || bytes.byteLength < 512) return null;
    if (bytes.byteLength > MAX_BUFFER_BYTES) return null;
    // Magic check. CFB signature is at the very start of the header sector.
    for (let i = 0; i < CFB_MAGIC.length; i++) {
        if (bytes[i] !== CFB_MAGIC[i]) return null;
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    // Sector shift: 9 (512-byte sectors) or 12 (4096-byte sectors).
    const sectorShift = view.getUint16(0x1e, true);
    if (sectorShift !== 9 && sectorShift !== 12) return null;
    const sectorSize = 1 << sectorShift;
    if (sectorSize < 128) return null;

    const miniSectorShift = view.getUint16(0x20, true);
    if (miniSectorShift < 1 || miniSectorShift > 16) return null;
    const miniSectorSize = 1 << miniSectorShift;

    const numFatSectors = view.getUint32(0x2c, true);
    const firstDirSector = view.getUint32(0x30, true);
    const miniStreamCutoff = view.getUint32(0x38, true);
    const firstMiniFatSector = view.getUint32(0x3c, true);
    const numMiniFatSectors = view.getUint32(0x40, true);
    const firstDifatSector = view.getUint32(0x44, true);
    const numDifatSectors = view.getUint32(0x48, true);

    // Sanity caps on header-declared sector counts. Keeps us from trying to
    // allocate an 8 GiB FAT for a tiny malformed file.
    if (numFatSectors > MAX_CHAIN_SECTORS) return null;
    if (numMiniFatSectors > MAX_CHAIN_SECTORS) return null;
    if (numDifatSectors > MAX_CHAIN_SECTORS) return null;

    // The file's total sector count (excluding the 512-byte header). For a
    // major=3 (sectorSize=512) file the header occupies its own sector; for
    // major=4 (sectorSize=4096) the header still occupies only 512 bytes and
    // the rest of sector 0 is reserved. Either way sector 0's DATA starts at
    // offset `sectorSize` from the file start (for major=4 there's 3584 bytes
    // of reserved header padding we never read).
    const sectorDataStart = sectorSize;
    const availableBytes = bytes.byteLength - sectorDataStart;
    if (availableBytes < 0) return null;
    const totalSectors = Math.floor(availableBytes / sectorSize);
    if (totalSectors <= 0) return null;

    // Helper: bounds-checked read of one sector into a subarray. Returns null
    // if the sector index is out of range.
    const sectorBytes = (sid: number): Uint8Array | null => {
        if (sid >= totalSectors) return null;
        const off = sectorDataStart + sid * sectorSize;
        if (off + sectorSize > bytes.byteLength) return null;
        return bytes.subarray(off, off + sectorSize);
    };

    // Assemble the DIFAT — the list of sector IDs that hold FAT. The first
    // 109 entries sit in the header (bytes 0x4C..0x200); additional entries
    // live in DIFAT sectors chained via the last 4 bytes of each DIFAT sector.
    const difat: number[] = [];
    for (let i = 0; i < 109; i++) {
        const sid = view.getUint32(0x4c + i * 4, true);
        if (sid === FREESECT) break;
        difat.push(sid);
    }
    // Chase DIFAT sectors. Cap the walk so a cyclic chain can't spin forever.
    {
        let next = firstDifatSector;
        let guard = 0;
        const seen = new Set<number>();
        while (next !== ENDOFCHAIN && next !== FREESECT && guard < numDifatSectors + 4) {
            if (seen.has(next)) break; // cycle guard
            seen.add(next);
            const sec = sectorBytes(next);
            if (!sec) break;
            const secView = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
            const entriesPerDifat = (sectorSize / 4) - 1;
            for (let i = 0; i < entriesPerDifat; i++) {
                const sid = secView.getUint32(i * 4, true);
                if (sid === FREESECT) break;
                difat.push(sid);
            }
            next = secView.getUint32(sectorSize - 4, true);
            guard++;
            if (difat.length > MAX_CHAIN_SECTORS) return null;
        }
    }

    // Materialise the FAT into a flat array indexed by sector id. Each entry
    // of `difat` names a sector that contains `(sectorSize / 4)` FAT entries
    // laid out sequentially.
    const fatEntriesPerSector = sectorSize / 4;
    const fatLength = difat.length * fatEntriesPerSector;
    if (fatLength > MAX_CHAIN_SECTORS) return null;
    const fat = new Uint32Array(fatLength);
    for (let i = 0; i < difat.length; i++) {
        const sec = sectorBytes(difat[i]);
        if (!sec) return null;
        const secView = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
        for (let j = 0; j < fatEntriesPerSector; j++) {
            fat[i * fatEntriesPerSector + j] = secView.getUint32(j * 4, true);
        }
    }

    // Walk a sector chain through the FAT. Returns the concatenated bytes.
    // Stops if the stated `size` is reached, if the chain terminates, or if
    // a sector repeats (loop defence).
    function readFatChain(firstSid: number, size: number): Uint8Array | null {
        if (firstSid === ENDOFCHAIN || firstSid === FREESECT) return new Uint8Array(0);
        const cap = Math.min(size, MAX_STREAM_BYTES);
        const out = new Uint8Array(cap);
        let written = 0;
        let sid = firstSid;
        const seen = new Set<number>();
        let hops = 0;
        while (sid !== ENDOFCHAIN && sid !== FREESECT && sid !== FATSECT && sid !== DIFSECT) {
            if (hops++ > MAX_CHAIN_SECTORS) return null;
            if (seen.has(sid)) return null;
            seen.add(sid);
            const sec = sectorBytes(sid);
            if (!sec) return null;
            const copyLen = Math.min(sec.length, cap - written);
            if (copyLen <= 0) break;
            out.set(sec.subarray(0, copyLen), written);
            written += copyLen;
            if (written >= cap) break;
            if (sid >= fat.length) return null;
            sid = fat[sid];
        }
        return out.subarray(0, Math.min(written, size));
    }

    // Directory stream — chained off `firstDirSector` through the FAT.
    // Size is unknown ahead of time so we pass a generous cap equal to the
    // max sector count * sector size (but bounded to the buffer length).
    const dirBytes = readFatChain(firstDirSector, bytes.byteLength);
    if (!dirBytes) return null;

    // Each directory entry is exactly 128 bytes. Cap enumeration at
    // MAX_STREAMS to block degenerate / adversarial inputs.
    const entryCount = Math.min(Math.floor(dirBytes.length / 128), MAX_STREAMS);
    if (entryCount === 0) return null;
    const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);

    interface DirEntry {
        name: string;
        type: number;
        startSector: number;
        size: number;
        clsid: string | null;
    }
    const entries: DirEntry[] = [];
    for (let i = 0; i < entryCount; i++) {
        const base = i * 128;
        if (base + 128 > dirBytes.length) break;
        const type = dirBytes[base + 0x42];
        if (type === DIR_TYPE_UNUSED) { entries.push({ name: '', type, startSector: 0, size: 0, clsid: null }); continue; }
        const nameLenBytes = dirView.getUint16(base + 0x40, true);
        // nameLenBytes is inclusive of the UTF-16LE null terminator. A value
        // of 0 or >64 is malformed; clamp to the 64-byte field.
        const clampedNameLen = Math.max(0, Math.min(nameLenBytes, 64));
        const name = decodeUtf16LeName(dirBytes.subarray(base, base + clampedNameLen));
        const startSector = dirView.getUint32(base + 0x74, true);
        // Stream size is an 8-byte field. For major=3 files the high 4 bytes
        // are reserved (should be zero); for major=4 they're the true upper
        // half of a 64-bit size. We only ever honour sizes up to
        // MAX_STREAM_BYTES anyway, so clamping to 32 bits is safe.
        const sizeLo = dirView.getUint32(base + 0x78, true);
        const sizeHi = dirView.getUint32(base + 0x7c, true);
        // If sizeHi is non-zero we're well past our 16 MiB per-stream cap;
        // synthesise MAX_STREAM_BYTES + 1 so the size check below rejects.
        const size = sizeHi !== 0 ? MAX_STREAM_BYTES + 1 : sizeLo;
        const clsid = readClsid(dirBytes, base + 0x50);
        entries.push({ name, type, startSector, size, clsid });
    }

    // Root entry is always at index 0 per the CFB spec. Its `startSector`
    // names the first sector of the mini-stream; its `size` is the total
    // mini-stream size.
    const root = entries[0];
    if (!root || root.type !== DIR_TYPE_ROOT) return null;

    // Assemble the mini-stream (bytes that hold the mini-sectors). Only
    // needed if any stream is below the cutoff — we read it lazily on first
    // use.
    let miniStream: Uint8Array | null = null;
    let miniFat: Uint32Array | null = null;
    const ensureMini = (): boolean => {
        if (miniStream && miniFat) return true;
        // Mini-stream lives in normal sectors, chained off root.startSector.
        miniStream = readFatChain(root.startSector, Math.min(root.size, MAX_STREAM_BYTES));
        if (!miniStream) return false;
        // Mini-FAT lives in normal sectors, chained off firstMiniFatSector.
        if (firstMiniFatSector === ENDOFCHAIN || firstMiniFatSector === FREESECT) {
            miniFat = new Uint32Array(0);
            return true;
        }
        const mfatBytes = readFatChain(firstMiniFatSector, numMiniFatSectors * sectorSize);
        if (!mfatBytes) return false;
        const entries = Math.floor(mfatBytes.length / 4);
        miniFat = new Uint32Array(entries);
        const mfatView = new DataView(mfatBytes.buffer, mfatBytes.byteOffset, mfatBytes.byteLength);
        for (let i = 0; i < entries; i++) miniFat[i] = mfatView.getUint32(i * 4, true);
        return true;
    };

    // Walk a mini-sector chain. Same shape as readFatChain but slices out of
    // the mini-stream using miniSectorSize-byte steps.
    function readMiniChain(firstSid: number, size: number): Uint8Array | null {
        if (!ensureMini() || !miniStream || !miniFat) return null;
        if (firstSid === ENDOFCHAIN || firstSid === FREESECT) return new Uint8Array(0);
        const cap = Math.min(size, MAX_STREAM_BYTES);
        const out = new Uint8Array(cap);
        let written = 0;
        let sid = firstSid;
        const seen = new Set<number>();
        let hops = 0;
        while (sid !== ENDOFCHAIN && sid !== FREESECT && sid !== FATSECT && sid !== DIFSECT) {
            if (hops++ > MAX_CHAIN_SECTORS) return null;
            if (seen.has(sid)) return null;
            seen.add(sid);
            const off = sid * miniSectorSize;
            if (off + miniSectorSize > miniStream.length) return null;
            const copyLen = Math.min(miniSectorSize, cap - written);
            if (copyLen <= 0) break;
            out.set(miniStream.subarray(off, off + copyLen), written);
            written += copyLen;
            if (written >= cap) break;
            if (sid >= miniFat.length) return null;
            sid = miniFat[sid];
        }
        return out.subarray(0, Math.min(written, size));
    }

    // Build the streams list. Skip the root (entry 0) and any storage /
    // unused entries — consumers only care about actual file payloads.
    const streams: CfbStream[] = [];
    for (let i = 1; i < entries.length && streams.length < MAX_STREAMS; i++) {
        const e = entries[i];
        if (e.type !== DIR_TYPE_STREAM) continue;
        if (e.size > MAX_STREAM_BYTES) continue;
        if (!e.name) continue;
        const useMini = e.size > 0 && e.size < miniStreamCutoff;
        const streamBytes = useMini
            ? readMiniChain(e.startSector, e.size)
            : readFatChain(e.startSector, e.size);
        if (!streamBytes) continue;
        streams.push({ name: e.name, bytes: streamBytes });
    }

    return {
        clsid: root.clsid,
        streams,
    };
}

// Decode a UTF-16LE name buffer, strip the trailing NUL, and drop any
// remaining C0/C1 control characters EXCEPT the `\x01` / `\x05` prefixes
// Word / Excel use to flag "control" streams (`\x01Ole`,
// `\x05SummaryInformation`, etc.). Preserving those lets consumers recognise
// them by name.
function decodeUtf16LeName(buf: Uint8Array): string {
    // nameLenBytes includes the null terminator; step in UTF-16 code units.
    let out = '';
    for (let i = 0; i + 1 < buf.length; i += 2) {
        const cu = buf[i] | (buf[i + 1] << 8);
        if (cu === 0) break;
        out += String.fromCharCode(cu);
    }
    // Strip any dangling control chars except the leading \x01/\x05 markers.
    if (out.length === 0) return '';
    const head = out[0];
    const rest = out.slice(1).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
    if (head === '\x01' || head === '\x05') {
        return head + rest;
    }
    // For non-marker streams, also drop a leading control char so the name
    // is human-readable.
    if (head >= '\x00' && head < ' ' && head !== '\t' && head !== '\n' && head !== '\r') {
        return rest;
    }
    return head + rest;
}

// Format a 16-byte CLSID field as the canonical registry form:
//   {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
// The first three groups are stored little-endian in the file; the fourth
// group (2 bytes) and the remainder (6 bytes) are big-endian per RFC-4122
// "GUID" encoding (which CFB follows).
function readClsid(buf: Uint8Array, offset: number): string | null {
    if (offset + 16 > buf.length) return null;
    const b = buf.subarray(offset, offset + 16);
    // All-zeros CLSID = "no class assigned". Skip.
    let allZero = true;
    for (let i = 0; i < 16; i++) { if (b[i] !== 0) { allZero = false; break; } }
    if (allZero) return null;
    const hex = (n: number) => n.toString(16).padStart(2, '0');
    // Groups 1-3 are little-endian.
    const g1 = hex(b[3]) + hex(b[2]) + hex(b[1]) + hex(b[0]);
    const g2 = hex(b[5]) + hex(b[4]);
    const g3 = hex(b[7]) + hex(b[6]);
    // Group 4 and the final group are big-endian.
    const g4 = hex(b[8]) + hex(b[9]);
    const g5 = hex(b[10]) + hex(b[11]) + hex(b[12]) + hex(b[13]) + hex(b[14]) + hex(b[15]);
    return `{${g1}-${g2}-${g3}-${g4}-${g5}}`.toUpperCase();
}
