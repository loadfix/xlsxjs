// A1 ↔ R1C1 formula-notation converters. Excel supports both: A1 is the
// default; R1C1 is an option in Preferences that expresses every cell ref
// relative (bracketed) or absolute (unbracketed) with respect to the anchor
// cell carrying the formula.
//
// Conversion rules (per ECMA-376 / Excel behaviour):
//   A1-style:   Column letters, row numbers. `$` pins an axis.
//     A1                → relative col, relative row
//     $A$1              → absolute col, absolute row
//     $A1  / A$1        → mixed
//   R1C1-style: R[rowDelta]C[colDelta] for relative; R<row>C<col> for absolute.
//     R[-1]C[-1]        → relative -1 row, -1 col (from the anchor)
//     R1C1              → absolute row 1, col 1
//     R[-1]C1           → relative row, absolute col
//     R1C[-1]           → absolute row, relative col
//
// This module is pure string math: it does not evaluate formulas or track
// range boundaries. It handles the common A1:B2 range token by converting
// each endpoint independently.

import { columnLettersToIndex, indexToColumnLetters } from './utils';

// ── A1 → R1C1 ────────────────────────────────────────────────────────────

// Matches a single A1 cell reference with optional $ on column and row.
const A1_REF = /(\$?)([A-Z]+)(\$?)([1-9][0-9]*)/g;

export function a1ToR1c1(formula: string, anchorRow: number, anchorCol: number): string {
    // Replace each A1 ref in-place. 0-based anchor (row, col).
    return formula.replace(A1_REF, (_, colAbs, col, rowAbs, rowStr) => {
        const colIdx = columnLettersToIndex(col);
        const rowIdx = Number(rowStr) - 1;
        if (colIdx < 0 || !Number.isFinite(rowIdx)) return _;
        const rowPart = rowAbs ? `R${rowIdx + 1}` : `R[${rowIdx - anchorRow}]`;
        const colPart = colAbs ? `C${colIdx + 1}` : `C[${colIdx - anchorCol}]`;
        return rowPart + colPart;
    });
}

// ── R1C1 → A1 ────────────────────────────────────────────────────────────

// Matches R[delta]C[delta] / R<n>C<n>, with each axis independently bracketed
// or absolute. The two axes must appear in R-then-C order.
const R1C1_REF = /R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?/g;

export function r1c1ToA1(formula: string, anchorRow: number, anchorCol: number): string {
    return formula.replace(R1C1_REF, (match, rowPart, colPart) => {
        if (rowPart === undefined && colPart === undefined) return match;
        const { index: rowIdx, abs: rowAbs } = decodeAxis(rowPart, anchorRow);
        const { index: colIdx, abs: colAbs } = decodeAxis(colPart, anchorCol);
        if (rowIdx === null || colIdx === null) return match;
        const col = indexToColumnLetters(colIdx);
        const prefixCol = colAbs ? '$' : '';
        const prefixRow = rowAbs ? '$' : '';
        return `${prefixCol}${col}${prefixRow}${rowIdx + 1}`;
    });
}

function decodeAxis(part: string | undefined, anchor: number): { index: number | null; abs: boolean } {
    if (part === undefined) return { index: anchor, abs: false };
    if (part.startsWith('[') && part.endsWith(']')) {
        const delta = Number(part.slice(1, -1));
        if (!Number.isFinite(delta)) return { index: null, abs: false };
        return { index: anchor + delta, abs: false };
    }
    const n = Number(part);
    if (!Number.isFinite(n) || n < 1) return { index: null, abs: false };
    return { index: n - 1, abs: true };
}
