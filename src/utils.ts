// Shared helpers. Excel column letters ⇄ 0-based indices, cell-ref parsing,
// and safe-ident / CSS-value guards. Small and dependency-free so they can be
// unit-tested against malicious XLSX input without DOM.

export function columnLettersToIndex(letters: string): number {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
        const c = letters.charCodeAt(i);
        if (c < 0x41 || c > 0x5a) return -1;
        n = n * 26 + (c - 0x40);
    }
    return n - 1;
}

export function indexToColumnLetters(index: number): string {
    let n = index + 1;
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(0x41 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

const CELL_REF = /^([A-Z]+)([1-9][0-9]*)$/;

export function parseCellRef(ref: string): { col: number; row: number } | null {
    const m = CELL_REF.exec(ref);
    if (!m) return null;
    const col = columnLettersToIndex(m[1]);
    if (col < 0) return null;
    return { col, row: Number(m[2]) - 1 };
}

// English Metric Units → device-independent pixels at 96 DPI. Excel uses EMU
// (914400 per inch, i.e. 9525 per pixel) for drawing extents and offsets.
// Non-finite / negative inputs return 0 so callers can pipe the result
// straight into a `px(...)` formatter without extra guards.
export function emuToPx(emu: number): number {
    if (!Number.isFinite(emu)) return 0;
    return Math.round(emu / 9525);
}
