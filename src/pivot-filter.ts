// Opt-in pivot re-materialisation helpers. Wave 10 shipped the interactive
// slicer + timeline aside (`interactiveSlicers: true`) that fire
// `xlsx:slicer-change` / `xlsx:timeline-change` CustomEvents on the aside.
// This module provides the default event-handler body that walks the
// pivot's rendered `<table>` and hides the rows whose cell at the slicer's
// column doesn't match the current selection.
//
// All three helpers are pure, DOM-aware functions — they take a
// `<table>` element + plain arguments and toggle `tr.style.display`. They
// don't touch the parsed workbook model, and they only READ textContent
// from cells (never innerHTML). Reverting a filter is just calling
// `applySlicerFilter` again with a superset of items, or
// `clearPivotFilters` to reset every row.
//
// The column index is 0-based relative to the table's rendered <td> cells
// (the row-number `<th>` that xlsxjs emits as first child of every tbody
// row is NOT counted). Wiring from the html-renderer translates a slicer's
// `sourceName` into this index by walking the first visible tbody row's
// <td>s and matching textContent case-insensitively.

// Hide every tbody row whose cell at `column` is not one of `selectedItems`.
// Rows whose cell-at-column textContent is `null`/undefined (e.g. merged
// cells or a shorter `<tr>`) are left alone so the filter doesn't accidentally
// hide a row that doesn't participate in the column at all. Empty
// `selectedItems` hides every row that DOES have a cell at the column —
// that's the "all unselected" state and mirrors the chips' aria-pressed
// truth table.
//
// `headerRowIndex` optionally pins the tbody row that carries column
// labels (Excel-style row 1 land in tbody under xlsxjs); when set, that
// row is always kept visible so filtering a pivot doesn't hide its own
// header. Omit or pass -1 to treat every row as data.
export function applySlicerFilter(
    pivotTable: HTMLTableElement,
    column: number,
    selectedItems: string[],
    headerRowIndex: number = -1,
): void {
    const tbody = pivotTable.querySelector('tbody');
    if (!tbody) return;
    // Set lookup is O(1) per row, so large pivots stay linear in row count.
    // Use a Set (not a plain object) so attacker-controlled item strings
    // can't collide with Object.prototype keys.
    const selected = new Set(selectedItems);
    const rows = tbody.querySelectorAll(':scope > tr');
    for (let i = 0; i < rows.length; i++) {
        if (i === headerRowIndex) continue;
        const tr = rows[i] as HTMLTableRowElement;
        const cell = cellAt(tr, column);
        if (!cell) continue;
        const text = (cell.textContent ?? '').trim();
        // Empty text means either a blank cell or a cell that hasn't been
        // rendered with a value — don't hide those: they're not expressing
        // the column's domain.
        if (text === '') continue;
        tr.style.display = selected.has(text) ? '' : 'none';
    }
}

// Hide every tbody row whose cell at `column` parses to a Date outside the
// inclusive range `[start, end]`. Unparseable cells (textContent that
// `new Date(…)` rejects) are left alone — same principle as applySlicerFilter.
// `headerRowIndex` follows the same convention as applySlicerFilter (see
// that doc-comment).
export function applyTimelineFilter(
    pivotTable: HTMLTableElement,
    column: number,
    start: Date,
    end: Date,
    headerRowIndex: number = -1,
): void {
    const tbody = pivotTable.querySelector('tbody');
    if (!tbody) return;
    const startMs = start.getTime();
    const endMs = end.getTime();
    // A NaN bound means the caller handed us a non-Date — clear-filter
    // semantics so we don't hide every row.
    if (!Number.isFinite(startMs) || !Number.isFinite(endMs)) return;
    const lo = Math.min(startMs, endMs);
    const hi = Math.max(startMs, endMs);
    const rows = tbody.querySelectorAll(':scope > tr');
    for (let i = 0; i < rows.length; i++) {
        if (i === headerRowIndex) continue;
        const tr = rows[i] as HTMLTableRowElement;
        const cell = cellAt(tr, column);
        if (!cell) continue;
        const text = (cell.textContent ?? '').trim();
        if (text === '') continue;
        const t = Date.parse(text);
        // Null-guard: skip rows whose cell doesn't parse. Consumers who want
        // "unparseable = hide" can post-process, but the default contract is
        // "only filter rows whose cell participates in the timeline's field".
        if (!Number.isFinite(t)) continue;
        tr.style.display = t >= lo && t <= hi ? '' : 'none';
    }
}

// Reset `display` on every tbody row so the pivot re-shows its full row set.
// Used when a consumer wires "clear all filters" UI; also useful when a
// slicer is destroyed and its filter effect needs to roll back.
export function clearPivotFilters(pivotTable: HTMLTableElement): void {
    const tbody = pivotTable.querySelector('tbody');
    if (!tbody) return;
    const rows = tbody.querySelectorAll(':scope > tr');
    for (let i = 0; i < rows.length; i++) {
        (rows[i] as HTMLTableRowElement).style.display = '';
    }
}

// Resolve the n-th <td> of a row, skipping the row-number <th> that
// xlsxjs emits as the first child of every tbody row. Returns null when
// the row doesn't have enough <td>s — caller skips those rows so a short
// merged row doesn't accidentally get hidden.
function cellAt(tr: HTMLTableRowElement, column: number): Element | null {
    const tds = tr.querySelectorAll(':scope > td');
    if (column < 0 || column >= tds.length) return null;
    return tds[column];
}

// Locate the <td>/<th> column index whose textContent matches `sourceName`
// case-insensitively. Walks (a) the <thead> rows' <th>/<td> cells and
// (b) the <tbody> rows' cells — pivot producers may put the header
// label in either place, and xlsxjs renders Excel's "Row 1" into tbody
// so a sourceName like "Region" is typically in tbody row 0. Returns -1
// when no match is found so the caller can skip wiring gracefully.
export function findColumnByHeaderText(
    pivotTable: HTMLTableElement,
    sourceName: string,
): number {
    return findHeaderCell(pivotTable, sourceName).column;
}

// Richer variant of `findColumnByHeaderText` used by the renderer's
// wiring layer. Returns both the column index AND the tbody-relative
// row index where the header was found (or -1 when the label was
// resolved via <thead> only, in which case no tbody row needs to be
// pinned visible). Returning {-1, -1} means "source not resolved".
export function findHeaderCell(
    pivotTable: HTMLTableElement,
    sourceName: string,
): { column: number; tbodyRowIndex: number } {
    const target = sourceName.trim().toLowerCase();
    if (target === '') return { column: -1, tbodyRowIndex: -1 };
    const thead = pivotTable.querySelector(':scope > thead');
    if (thead) {
        const headerRows = thead.querySelectorAll(':scope > tr');
        for (let i = 0; i < headerRows.length; i++) {
            const idx = matchRowCells(headerRows[i] as HTMLTableRowElement, target);
            if (idx !== -1) return { column: idx, tbodyRowIndex: -1 };
        }
    }
    const tbody = pivotTable.querySelector(':scope > tbody');
    if (tbody) {
        const bodyRows = tbody.querySelectorAll(':scope > tr');
        for (let i = 0; i < bodyRows.length; i++) {
            const idx = matchRowCells(bodyRows[i] as HTMLTableRowElement, target);
            if (idx !== -1) return { column: idx, tbodyRowIndex: i };
        }
    }
    return { column: -1, tbodyRowIndex: -1 };
}

// Scan a row's cells (skipping the leading row-number <th>) for the first
// cell whose trimmed textContent matches `target` case-insensitively. The
// row-number <th> is emitted by xlsxjs with a plain integer label (e.g.
// "1"/"2"/…) so we intentionally skip it to keep the column index
// consistent with applySlicerFilter / applyTimelineFilter.
function matchRowCells(tr: HTMLTableRowElement, target: string): number {
    const cells = tr.querySelectorAll(':scope > td, :scope > th');
    // The first cell is either the row-number <th> (when present) or a
    // genuine data cell. Detect the former by checking for a numeric
    // textContent and a `<th>` tag; otherwise treat the 0-th cell as
    // column 0.
    let offset = 0;
    const first = cells[0];
    if (first && first.tagName === 'TH' && /^\d+$/.test((first.textContent ?? '').trim())) {
        offset = 1;
    }
    for (let i = offset; i < cells.length; i++) {
        const txt = (cells[i].textContent ?? '').trim().toLowerCase();
        if (txt === target) return i - offset;
    }
    return -1;
}
