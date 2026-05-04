// Node harness that loads an XLSX fixture through the built library in jsdom
// and asserts the renderer produces a sensible table. Analogous to docxjs's
// test-track-changes.mjs but for the initial rendering slice.

import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');

const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    url: 'http://localhost/',
});
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.Node = dom.window.Node;
globalThis.Element = dom.window.Element;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;

await import('jszip').then((m) => { globalThis.JSZip = m.default; });
const umd = readFileSync(`${repo}/dist/xlsx-preview.js`, 'utf8');
new Function('require', umd)(() => ({}));
const { parseAsync, renderWorkbook } = globalThis.xlsx;

const failures = [];
const warnings = [];
function assert(cond, msg) { if (!cond) failures.push(msg); }
function note(msg) { warnings.push(msg); }

async function renderFixture(path, options) {
    const buf = readFileSync(resolve(repo, 'tests/render-test', path, 'workbook.xlsx'));
    const wb = await parseAsync(buf, options);
    const nodes = await renderWorkbook(wb, options);
    const container = document.createElement('div');
    for (const n of nodes) container.appendChild(n);
    return { wb, container };
}

// ── 1. Basic fixture renders a table with the expected cells ──────────────
{
    const { wb, container } = await renderFixture('basic');
    assert(wb.parsed.sheets.length >= 1, '1a: workbook should have at least one sheet');
    const sheet = wb.parsed.sheets[0];
    note(`1·: sheet "${sheet.name}" → ${sheet.maxRow + 1} rows × ${sheet.maxCol + 1} cols`);

    const table = container.querySelector('section.xlsx table');
    assert(!!table, '1b: a <table> should be rendered');

    const tbodyRows = container.querySelectorAll('section.xlsx tbody tr');
    assert(tbodyRows.length === sheet.maxRow + 1, `1c: tbody rows should match maxRow+1 (got ${tbodyRows.length})`);

    // Row 0 should be the header: Name | Score | Passed.
    const firstDataRow = tbodyRows[0];
    const tds = firstDataRow.querySelectorAll('td');
    assert(tds[0].textContent === 'Name', `1d: A1 should be "Name" (got ${tds[0].textContent})`);
    assert(tds[1].textContent === 'Score', `1e: B1 should be "Score" (got ${tds[1].textContent})`);
    assert(tds[2].textContent === 'Passed', `1f: C1 should be "Passed" (got ${tds[2].textContent})`);

    // Row 1: Alice | 90 | TRUE. 90 should render as numeric (aligned).
    const row2 = tbodyRows[1].querySelectorAll('td');
    assert(row2[0].textContent === 'Alice', '1g: A2 should be "Alice"');
    assert(row2[1].textContent === '90', `1h: B2 should be "90" (got ${row2[1].textContent})`);
    assert(row2[1].classList.contains('xlsx-numeric'), '1i: numeric cell should have .xlsx-numeric');
    assert(row2[2].textContent === 'TRUE', `1j: C2 should be "TRUE" (got ${row2[2].textContent})`);
}

// ── 2. Header row shows column letters (A, B, C) ──────────────────────────
{
    const { container } = await renderFixture('basic');
    const headers = container.querySelectorAll('section.xlsx thead th');
    // First <th> is the empty corner, then A, B, C…
    assert(headers[0].textContent === '', '2a: corner cell should be empty');
    assert(headers[1].textContent === 'A', `2b: first column header should be "A" (got ${headers[1].textContent})`);
    assert(headers[2].textContent === 'B', `2c: second column header should be "B"`);
    assert(headers[3].textContent === 'C', `2d: third column header should be "C"`);
}

// ── 3. <style> emitted once ────────────────────────────────────────────────
{
    const { container } = await renderFixture('basic');
    const styles = container.querySelectorAll('style[data-xlsxjs]');
    assert(styles.length === 1, `3a: exactly one scoped <style> expected (got ${styles.length})`);
}

// ── 4. Merged cells: horizontal + vertical spans, suppressed covered cells ─
{
    const { wb, container } = await renderFixture('merged');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.merges.length === 2, `4a: expected 2 merges in fixture (got ${sheet.merges.length})`);
    note(`4·: merges = ${JSON.stringify(sheet.merges)}`);

    const rows = container.querySelectorAll('section.xlsx tbody tr');
    // Row 1 (0-indexed): anchor A1 → colspan=3, text "Q4 Report".
    // Layout per row: <th row#> + <td anchor colspan=3>  → 2 children total.
    const row1Cells = rows[0].children;
    assert(row1Cells.length === 2, `4b: merged header row should have 2 children (th + td), got ${row1Cells.length}`);
    const titleTd = row1Cells[1];
    assert(titleTd.getAttribute('colspan') === '3', `4c: title td should have colspan=3 (got ${titleTd.getAttribute('colspan')})`);
    assert(titleTd.textContent === 'Q4 Report', `4d: title td text should be "Q4 Report" (got ${titleTd.textContent})`);
    assert(titleTd.classList.contains('xlsx-merged'), '4e: merged anchor should carry .xlsx-merged');

    // Row 3 (0-indexed 2, rendered row 3): anchor A3 with rowspan=2.
    const row3Cells = rows[2].children;
    // Expected children: <th 3>, <td Alice rowspan=2>, <td North>, <td 90>.
    assert(row3Cells.length === 4, `4f: row 3 should have 4 children (got ${row3Cells.length})`);
    const aliceTd = row3Cells[1];
    assert(aliceTd.getAttribute('rowspan') === '2', `4g: Alice cell should have rowspan=2 (got ${aliceTd.getAttribute('rowspan')})`);
    assert(aliceTd.textContent === 'Alice', `4h: Alice cell text (got ${aliceTd.textContent})`);

    // Row 4 (0-indexed 3): the A4 slot is suppressed by the vertical merge.
    // Expected children: <th 4>, <td North>, <td 85> — only 3 children.
    const row4Cells = rows[3].children;
    assert(row4Cells.length === 3, `4i: row 4 should have 3 children after suppression (got ${row4Cells.length})`);
    assert(row4Cells[1].textContent === 'North', `4j: first td in row 4 should be "North" (got ${row4Cells[1].textContent})`);
}

// ── 5. Column widths from <cols> land on the <colgroup> ────────────────────
{
    const { container } = await renderFixture('merged');
    const colgroup = container.querySelector('section.xlsx colgroup');
    assert(!!colgroup, '5a: a <colgroup> should be rendered');
    const cols = colgroup.querySelectorAll('col');
    // 1 for gutter + 3 data columns = 4.
    assert(cols.length === 4, `5b: colgroup should have 4 <col>s (got ${cols.length})`);
    // Gutter has no width; data columns have widths 18, 10, 8 → px 131, 75, 61.
    assert(cols[0].style.width === '', '5c: gutter <col> should have no width');
    assert(cols[1].style.width === '131px', `5d: col A width should be 131px (got ${cols[1].style.width})`);
    assert(cols[2].style.width === '75px', `5e: col B width should be 75px (got ${cols[2].style.width})`);
    assert(cols[3].style.width === '61px', `5f: col C width should be 61px (got ${cols[3].style.width})`);
}

// ── 6. Basic fixture has no merges and no <cols> — behaviour unchanged ────
{
    const { wb, container } = await renderFixture('basic');
    assert(wb.parsed.sheets[0].merges.length === 0, '6a: basic fixture should have no merges');
    assert(wb.parsed.sheets[0].columns.length === 0, '6b: basic fixture should have no explicit column widths');
    const cols = container.querySelectorAll('section.xlsx colgroup col');
    // Still rendered, but every <col> except the gutter has no width.
    assert(cols.length === 4, `6c: basic should still have colgroup with 4 cols (got ${cols.length})`);
    for (let i = 1; i < cols.length; i++) {
        assert(cols[i].style.width === '', `6d·${i}: basic col ${i} should have no width (got ${cols[i].style.width})`);
    }
}

// ── 7. python-xlsx fixture: styles parse + apply ───────────────────────────
// This fixture comes out of the real python-xlsx library (see
// scripts/make-python-xlsx-fixture.py) and exercises the styles slice end
// to end: custom number formats, fonts, fills, borders, alignment.
{
    const { wb, container } = await renderFixture('python-xlsx');
    assert(!!wb.parsed.styles, '7a: styles should be present on the parsed workbook');

    const tbodyRows = container.querySelectorAll('section.xlsx tbody tr');

    // Row 0 (the title banner): bold, colspan=5.
    const titleTd = tbodyRows[0].querySelectorAll('td')[0];
    assert(titleTd.textContent === 'Q4 Sales Report', `7b: title text (got ${titleTd.textContent})`);
    assert(titleTd.style.fontWeight === 'bold', `7c: title should be bold (got ${titleTd.style.fontWeight})`);
    assert(titleTd.getAttribute('colspan') === '5', `7d: title colspan should be 5 (got ${titleTd.getAttribute('colspan')})`);

    // Row 2: header row with blue fill + bold white text across 5 cells.
    const headerTds = tbodyRows[2].querySelectorAll('td');
    assert(headerTds.length === 5, `7e: header row should have 5 data cells (got ${headerTds.length})`);
    for (const td of headerTds) {
        assert(td.style.backgroundColor !== '', `7f: header ${td.textContent} should have a fill`);
        assert(td.style.fontWeight === 'bold', `7g: header ${td.textContent} should be bold`);
    }
    assert(headerTds[0].textContent === 'Date', '7h: first header cell is Date');

    // Row 3: first data row. Date rendered from serial, percent + currency
    // rendered from their number formats, integer left un-decimalised.
    const row3 = tbodyRows[3].querySelectorAll('td');
    assert(row3[0].textContent === '2026-01-14', `7i: date serial 46036 should render as 2026-01-14 (got ${row3[0].textContent})`);
    assert(row3[2].textContent === '120', `7j: integer units stay as "120" (got ${row3[2].textContent})`);
    assert(row3[3].textContent === '22.0%', `7k: margin 0.22 should render as "22.0%" (got ${row3[3].textContent})`);
    assert(row3[4].textContent === '$1,560.00', `7l: revenue 1560 should render as "$1,560.00" (got ${row3[4].textContent})`);

    // Numeric class still applied for date + percent + currency + integer.
    for (const idx of [0, 2, 3, 4]) {
        assert(row3[idx].classList.contains('xlsx-numeric'), `7m·${idx}: row3 col ${idx} should have .xlsx-numeric`);
    }

    // Borders on data rows — any one side non-empty is enough evidence the
    // border pipeline ran. Excel writes rgb(153,153,153) thin borders here.
    const sampleDataCell = row3[2];
    assert(
        sampleDataCell.style.borderTop !== '' || sampleDataCell.style.borderLeft !== '',
        `7n: data cells should carry a border (got top="${sampleDataCell.style.borderTop}" left="${sampleDataCell.style.borderLeft}")`,
    );

    // Row 7: total row. E8 is a SUM formula with a cached value (6250.5)
    // and the currency number format, so it renders as "$6,250.50". B8
    // has "Total" in italic.
    const row7 = tbodyRows[7].querySelectorAll('td');
    assert(row7[1].textContent === 'Total', '7o: B8 text should be "Total"');
    assert(row7[1].style.fontStyle === 'italic', `7p: "Total" should be italic (got ${row7[1].style.fontStyle})`);
    assert(row7[4].textContent === '$6,250.50', `7q: E8 SUM should render cached value "$6,250.50" (got ${row7[4].textContent})`);
    assert(row7[4].style.fontWeight === 'bold', `7r: E8 total should be bold (got ${row7[4].style.fontWeight})`);

    // And the parser model should report the formula text on the Cell.
    const e8 = wb.parsed.sheets[0].rows[7].find((c) => c.col === 4);
    assert(e8.formula === 'SUM(E4:E6)', `7s: E8 formula text (got ${JSON.stringify(e8.formula)})`);
}

// ── 8. Formula cells: cached value rendered, formula text preserved ────────
{
    const { wb, container } = await renderFixture('formulas');
    const sheet = wb.parsed.sheets[0];

    // Formula detection lands on the Cell model regardless of whether a
    // cached value is present.
    const a2 = sheet.rows[1].find((c) => c.col === 0);
    assert(a2.formula === 'SUM(A1:C1)', `8a: A2 formula (got ${JSON.stringify(a2.formula)})`);
    assert(a2.value === '60', `8b: A2 should carry its cached value "60" (got ${a2.value})`);

    const a3 = sheet.rows[2].find((c) => c.col === 0);
    assert(a3.formula === 'A1+A2', `8c: A3 formula (got ${JSON.stringify(a3.formula)})`);
    assert(a3.value === '', `8d: A3 has no cached value → empty (got ${JSON.stringify(a3.value)})`);
    assert(a3.kind === 'empty', `8e: A3 kind should be "empty" without <v> (got ${a3.kind})`);

    const a4 = sheet.rows[3].find((c) => c.col === 0);
    assert(a4.formula === 'A1&" rows"', `8f: A4 formula (got ${JSON.stringify(a4.formula)})`);
    assert(a4.value === '10 rows', `8g: A4 cached string value (got ${a4.value})`);
    assert(a4.kind === 'string', `8h: A4 kind should be "string" for t="str" (got ${a4.kind})`);

    const a5 = sheet.rows[4].find((c) => c.col === 0);
    assert(a5.kind === 'error', `8i: A5 should be error kind (got ${a5.kind})`);
    assert(a5.value === '#DIV/0!', `8j: A5 should carry the #DIV/0! error text (got ${a5.value})`);

    // And now in the rendered DOM: cached value shows, no cached value is
    // blank, error class applied, string result renders as text (not
    // numeric-aligned).
    const tbodyRows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (rIdx, cIdx) => tbodyRows[rIdx].querySelectorAll('td')[cIdx];

    // A2 is in row index 1, col 0.
    assert(tdAt(1, 0).textContent === '60', `8k: A2 should render as "60" (got ${tdAt(1, 0).textContent})`);
    assert(tdAt(1, 0).classList.contains('xlsx-numeric'), '8l: A2 cached number should be right-aligned');

    // A3 renders empty.
    assert(tdAt(2, 0).textContent === '', `8m: A3 should render empty (got ${JSON.stringify(tdAt(2, 0).textContent)})`);
    assert(!tdAt(2, 0).classList.contains('xlsx-numeric'), '8n: A3 has no value → no numeric class');

    // A4 is a t="str" cached string, not numeric.
    assert(tdAt(3, 0).textContent === '10 rows', `8o: A4 should render as "10 rows" (got ${tdAt(3, 0).textContent})`);
    assert(!tdAt(3, 0).classList.contains('xlsx-numeric'), '8p: A4 string result should not be numeric-aligned');

    // A5 is an error cell.
    assert(tdAt(4, 0).textContent === '#DIV/0!', `8q: A5 should render as "#DIV/0!" (got ${tdAt(4, 0).textContent})`);
    assert(tdAt(4, 0).classList.contains('xlsx-error'), '8r: A5 should have .xlsx-error class');
}

// ── 9. Rich text: <si> with <r>/<rPr> runs renders spans with styling ────
{
    const { wb, container } = await renderFixture('richtext');
    assert(!!wb.parsed.theme, '9a: theme should be parsed');

    // Theme mapping: index 4 → accent1 ("FF0000" in this fixture's theme1.xml).
    assert(wb.parsed.theme.colors[4] === '#ff0000', `9b: accent1 (theme idx 4) should be #ff0000 (got ${wb.parsed.theme.colors[4]})`);

    const sheet = wb.parsed.sheets[0];

    // A1: rich cell. value is the concatenated plain text; runs carries 3.
    const a1 = sheet.rows[0].find((c) => c.col === 0);
    assert(a1.value === 'Hello world!', `9c: A1 flat value (got ${JSON.stringify(a1.value)})`);
    assert(Array.isArray(a1.runs) && a1.runs.length === 3, `9d: A1 should have 3 runs (got ${a1.runs?.length})`);
    assert(a1.runs[0].bold === false && a1.runs[0].text === 'Hello ', '9e: run 0 is plain "Hello "');
    assert(a1.runs[1].bold === true && a1.runs[1].text === 'world', '9f: run 1 is bold "world"');
    assert(a1.runs[1].color && a1.runs[1].color.kind === 'rgb' && a1.runs[1].color.value === '#ff0000',
        `9g: run 1 colour should be rgb #ff0000 (got ${JSON.stringify(a1.runs[1].color)})`);
    assert(a1.runs[2].italic === true && a1.runs[2].text === '!', '9h: run 2 is italic "!"');

    // A2: plain cell. No runs; single text-content path.
    const a2 = sheet.rows[1].find((c) => c.col === 0);
    assert(a2.runs === null, `9i: plain-string cell should have no runs (got ${JSON.stringify(a2.runs)})`);

    // DOM: A1 has 3 <span>s; plain cell has none.
    const tdAt = (rIdx) => container.querySelectorAll('section.xlsx tbody tr')[rIdx].querySelectorAll('td')[0];
    const a1Td = tdAt(0);
    const spans = a1Td.querySelectorAll('span');
    assert(spans.length === 3, `9j: A1 td should contain 3 spans (got ${spans.length})`);
    assert(spans[0].textContent === 'Hello ' && spans[0].style.fontWeight === '', `9k: span 0 plain (got weight="${spans[0].style.fontWeight}")`);
    assert(spans[1].style.fontWeight === 'bold', `9l: span 1 bold (got ${spans[1].style.fontWeight})`);
    assert(spans[1].style.color !== '', `9m: span 1 should carry colour (got "${spans[1].style.color}")`);
    assert(spans[2].style.fontStyle === 'italic', `9n: span 2 italic (got ${spans[2].style.fontStyle})`);

    // A2 plain: no spans, just text.
    const a2Td = tdAt(1);
    assert(a2Td.querySelectorAll('span').length === 0, '9o: plain cell should have no spans');
    assert(a2Td.textContent === 'plain text', `9p: plain cell text (got ${a2Td.textContent})`);
}

// ── 10. Theme colour resolution + tint on a cell's xf ──────────────────────
{
    const { wb, container } = await renderFixture('richtext');
    const tdAt = (rIdx) => container.querySelectorAll('section.xlsx tbody tr')[rIdx].querySelectorAll('td')[0];

    // A3 has styleIndex=1; styles.xml xf[1] uses fontId=1, whose colour is
    // theme=5 (accent2 → "00FF00" in this fixture's theme). No tint on the
    // font itself, but the run inside the <si> declares theme=4 tint=-0.25
    // which should darken accent1 red → rgb(191, 0, 0).
    const a3 = wb.parsed.sheets[0].rows[2].find((c) => c.col === 0);
    assert(a3.styleIndex === 1, `10a: A3 styleIndex (got ${a3.styleIndex})`);
    assert(a3.runs?.[0]?.color?.kind === 'theme', '10b: A3 run colour should be a theme ref');

    const a3Td = tdAt(2);
    // The cell-level xf font colour (accent2 green) is applied on the <td>;
    // the run span overrides it with accent1 red + tint.
    assert(a3Td.style.color.includes('rgb(0, 255, 0)') || a3Td.style.color === '#00ff00',
        `10c: A3 td should carry accent2 green from xf font (got "${a3Td.style.color}")`);

    const span = a3Td.querySelector('span');
    assert(!!span, '10d: A3 should render the run in a span');
    // Tint -0.25 on #ff0000 → roughly rgb(191, 0, 0). Accept either the
    // browser's canonical "rgb(R, G, B)" form or the hex string if the DOM
    // happens to echo back what we set.
    const spanColor = span.style.color;
    assert(/rgb\(191,\s*0,\s*0\)/.test(spanColor) || spanColor === '#bf0000',
        `10e: A3 span colour should be tinted accent1 ≈ rgb(191, 0, 0) (got "${spanColor}")`);
}

// ── 11. applyTint unit coverage ────────────────────────────────────────────
{
    const { applyTint } = globalThis.xlsx;
    assert(typeof applyTint === 'function', '11a: applyTint should be re-exported');
    // tint=0 is identity.
    assert(applyTint('#ff0000', 0) === '#ff0000', `11b: tint 0 identity (got ${applyTint('#ff0000', 0)})`);
    // tint=-1 should drive luminance → 0 (black).
    assert(applyTint('#ff0000', -1) === '#000000', `11c: tint -1 → black (got ${applyTint('#ff0000', -1)})`);
    // tint=1 should drive luminance → 1 (white).
    assert(applyTint('#ff0000', 1) === '#ffffff', `11d: tint +1 → white (got ${applyTint('#ff0000', 1)})`);
    // tint=-0.25 on pure red should darken to ~#bf0000 (within ±1 channel).
    const darker = applyTint('#ff0000', -0.25);
    const m = /^#([0-9a-f]{2})0000$/.exec(darker);
    assert(m && Math.abs(parseInt(m[1], 16) - 0xbf) <= 1, `11e: tint -0.25 on red ≈ #bf0000 (got ${darker})`);
}

// ── 12. python-xlsx fixture: theme-coloured subtitle + rich-text warning ──
{
    const { wb, container } = await renderFixture('python-xlsx');
    const sheet = wb.parsed.sheets[0];

    // Row 9 (1-indexed row 10): A10 is the theme-coloured subtitle.
    const subtitle = sheet.rows[9].find((c) => c.col === 0);
    assert(subtitle && subtitle.value === 'Generated by xlsxjs demo build', '12a: A10 subtitle text');

    // Row 11 (row 12): A12 is the rich-text warning — concatenated plain text
    // on `value`, 4 runs on `runs`.
    const warn = sheet.rows[11].find((c) => c.col === 0);
    assert(warn && warn.value === 'Warning: margins include discount category pricing.',
        `12b: A12 plain text (got ${JSON.stringify(warn?.value)})`);
    assert(Array.isArray(warn.runs) && warn.runs.length === 4, `12c: A12 should have 4 runs (got ${warn?.runs?.length})`);
    assert(warn.runs[0].bold === true && warn.runs[0].text === 'Warning: ', '12d: run 0 bold "Warning: "');
    assert(warn.runs[2].italic === true && warn.runs[2].text === 'discount', '12e: run 2 italic "discount"');

    // DOM: A12 renders 4 <span>s.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a12Td = rows[11].querySelectorAll('td')[0];
    const spans = a12Td.querySelectorAll('span');
    assert(spans.length === 4, `12f: A12 should render 4 spans (got ${spans.length})`);
    assert(spans[0].style.fontWeight === 'bold', '12g: span 0 bold');
    assert(spans[2].style.fontStyle === 'italic', '12h: span 2 italic');
}

// ── 13. Multi-sheet: rel resolution binds <sheet r:id> to the right xml ──
{
    const { wb, container } = await renderFixture('multisheet');
    const sheets = wb.parsed.sheets;
    assert(sheets.length === 3, `13a: expected 3 sheets (got ${sheets.length})`);
    assert(sheets[0].name === 'Alpha', `13b: first sheet should be "Alpha" (got ${sheets[0].name})`);
    assert(sheets[1].name === 'Beta', `13c: second sheet should be "Beta"`);
    assert(sheets[2].name === 'Gamma', `13d: third sheet should be "Gamma"`);

    // The fixture's rel map is Alpha→sheet1.xml (n=1), Beta→sheet2.xml (n=2),
    // Gamma→sheet3.xml (n=3). Each sheet's B1 cell carries `n` so we can
    // verify the rel binding. If the parser positionally resolved, the names
    // and numbers would still line up — the non-trivial evidence is that the
    // rId→target map is out of order (rId1→sheet2, rId2→sheet3, rId3→sheet1),
    // which means positional resolution would pair "Alpha" → sheet2.xml
    // (n=2) instead of sheet1.xml (n=1).
    const cellN = (sheetIdx) => sheets[sheetIdx].rows[0].find((c) => c.col === 1).value;
    assert(cellN(0) === '1', `13e: Alpha should carry n=1 (got ${cellN(0)})`);
    assert(cellN(1) === '2', `13f: Beta should carry n=2 (got ${cellN(1)})`);
    assert(cellN(2) === '3', `13g: Gamma should carry n=3 (got ${cellN(2)})`);

    // Each sheet renders as its own <section class="xlsx">.
    const sections = container.querySelectorAll('section.xlsx');
    assert(sections.length === 3, `13h: should render 3 sheet sections (got ${sections.length})`);
    assert(sections[0].getAttribute('data-sheet-name') === 'Alpha', '13i: first section is Alpha');
    assert(sections[2].getAttribute('data-sheet-name') === 'Gamma', '13j: last section is Gamma');
}

// ── 14. Conditional formatting: cellIs / containsText / duplicateValues ──
{
    const { wb, container } = await renderFixture('conditional-format');
    const sheet = wb.parsed.sheets[0];
    assert(wb.parsed.styles.dxfs.length === 3, `14a: expected 3 dxfs (got ${wb.parsed.styles.dxfs.length})`);
    assert(sheet.conditionalFormatting.length >= 1, '14b: sheet should carry conditionalFormatting blocks');

    const tds = [...container.querySelectorAll('td.xlsx-cf')];
    const byText = (t) => tds.filter((td) => td.textContent === t);

    // Scores ≥ 90 → bold + red fill. "95" (Alice) and "90" (Carol).
    const highs = byText('95').concat(byText('90'));
    assert(highs.length === 2, `14c: expected 95 and 90 to match cellIs rule (got ${highs.length})`);
    for (const td of highs) {
        assert(td.style.fontWeight === 'bold', `14d: high-score cell should be bold (got ${td.style.fontWeight})`);
        assert(/rgb\(255,\s*153,\s*153\)/.test(td.style.backgroundColor), `14e: high-score cell should have red fill (got "${td.style.backgroundColor}")`);
    }
    // Scores < 90 (72, 45, 62, 88) must NOT carry the cellIs class / colour.
    for (const t of ['72', '45', '62', '88']) {
        const matching = byText(t);
        assert(matching.length === 0, `14f: "${t}" should not match the cellIs rule (found ${matching.length})`);
    }

    // containsText "North" → yellow fill, priority 2 beats duplicateValues
    // (priority 3). Three "North" rows.
    const norths = byText('North');
    assert(norths.length === 3, `14g: expected 3 North cells matching (got ${norths.length})`);
    for (const td of norths) {
        assert(/rgb\(255,\s*255,\s*153\)/.test(td.style.backgroundColor), `14h: North cell should carry yellow fill (got "${td.style.backgroundColor}")`);
    }

    // duplicateValues → blue fill for "South" (appears twice in C2:C7).
    const souths = byText('South');
    assert(souths.length === 2, `14i: expected 2 South cells matching duplicate rule (got ${souths.length})`);
    for (const td of souths) {
        assert(/rgb\(204,\s*229,\s*255\)/.test(td.style.backgroundColor), `14j: South cell should carry blue fill (got "${td.style.backgroundColor}")`);
    }
    // "West" appears only once, so should NOT match the duplicate rule.
    assert(byText('West').length === 0, '14k: "West" is unique → no cf match');
}

// ── 15. Indexed colour resolution (direct call surface) ───────────────────
{
    const { resolveColor } = globalThis.xlsx;
    // Index 2 = red (#ff0000) per ECMA-376 §18.8.27.
    assert(resolveColor({ kind: 'indexed', index: 2, tint: 0 }, null) === '#ff0000', `15a: indexed 2 → red`);
    // Index 3 = green, index 4 = blue.
    assert(resolveColor({ kind: 'indexed', index: 3, tint: 0 }, null) === '#00ff00', `15b: indexed 3 → green`);
    assert(resolveColor({ kind: 'indexed', index: 4, tint: 0 }, null) === '#0000ff', `15c: indexed 4 → blue`);
    // Index 64/65 are "system" entries — return null so the renderer falls back.
    assert(resolveColor({ kind: 'indexed', index: 64, tint: 0 }, null) === null, `15d: indexed 64 (system fg) → null`);
    // Tint still applies: index 2 with tint -1 → black.
    assert(resolveColor({ kind: 'indexed', index: 2, tint: -1 }, null) === '#000000', `15e: indexed+tint`);
}

// ── 16. Row dimensions + hidden rows/columns ──────────────────────────────
{
    const { wb, container } = await renderFixture('dimensions');
    const sheet = wb.parsed.sheets[0];

    // Parser side: three rowDimensions (rows 0, 1 with custom heights; row 2
    // hidden) and one hidden column.
    const rd = sheet.rowDimensions.reduce((acc, d) => { acc[d.row] = d; return acc; }, {});
    assert(rd[0]?.height === 40, `16a: row 0 height 40 (got ${rd[0]?.height})`);
    assert(rd[1]?.height === 24, `16b: row 1 height 24 (got ${rd[1]?.height})`);
    assert(rd[2]?.hidden === true, `16c: row 2 hidden (got ${rd[2]?.hidden})`);
    assert(sheet.columns.length === 1 && sheet.columns[0].hidden, '16d: one hidden column record');
    assert(sheet.columns[0].min === 1 && sheet.columns[0].max === 1, '16e: hidden column covers col index 1 (B)');

    // DOM: rows have height styling; hidden row has display:none; hidden
    // column's cells all carry display:none on the td.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    assert(rows[0].style.height && rows[0].style.height !== '', `16f: row 0 should carry height style (got "${rows[0].style.height}")`);
    // 40pt → 53.33px (40 * 4/3).
    assert(/^53\.33px$/.test(rows[0].style.height), `16g: row 0 height should be 53.33px (got ${rows[0].style.height})`);
    assert(rows[2].style.display === 'none', `16h: row 2 should be display:none (got ${rows[2].style.display})`);
    // Every tbody row's column-index-1 td should be hidden.
    for (let i = 0; i < rows.length; i++) {
        const tds = rows[i].querySelectorAll('td');
        assert(tds[1].style.display === 'none', `16i·${i}: col B td should be hidden`);
    }
}

// ── 17. Named styles: cellStyleXfs inheritance (xfId without apply-flag) ──
{
    const { wb, container } = await renderFixture('named-styles');
    const styles = wb.parsed.styles;
    assert(styles.cellStyleXfs.length === 2, `17a: cellStyleXfs count (got ${styles.cellStyleXfs.length})`);
    assert(styles.cellXfs.length === 2, `17b: cellXfs count (got ${styles.cellXfs.length})`);
    // cellXfs[1].xfId should point at cellStyleXfs[1] (fontId=1).
    assert(styles.cellXfs[1].xfId === 1, `17c: cellXfs[1].xfId (got ${styles.cellXfs[1].xfId})`);
    assert(styles.cellStyleXfs[1].fontId === 1, '17d: cellStyleXfs[1].fontId → 1');

    // Rendered A1 (s=1) must inherit the base font (bold, red). The cell's
    // xf has no applyFont flag set — if inheritance is broken, A1 would be
    // plain.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a1 = rows[0].querySelectorAll('td')[0];
    assert(a1.style.fontWeight === 'bold', `17e: A1 should be bold via inheritance (got "${a1.style.fontWeight}")`);
    assert(/rgb\(204,\s*0,\s*0\)/.test(a1.style.color), `17f: A1 should carry base red (got "${a1.style.color}")`);

    // A2 uses xfId=0 (Normal) → no bold, no custom colour.
    const a2 = rows[1].querySelectorAll('td')[0];
    assert(a2.style.fontWeight === '', `17g: A2 should NOT be bold (got "${a2.style.fontWeight}")`);
    assert(a2.style.color === '', `17h: A2 should NOT carry a colour (got "${a2.style.color}")`);
}

// ── 18. Graphical cf: 3-colour scale + data bar ───────────────────────────
{
    const { wb, container } = await renderFixture('cf-graphical');
    const sheet = wb.parsed.sheets[0];

    // Two cf blocks, one colourScale (priority 1), one dataBar (priority 2).
    const allRules = sheet.conditionalFormatting.flatMap((b) => b.rules);
    const scale = allRules.find((r) => r.type === 'colorScale');
    const bar = allRules.find((r) => r.type === 'dataBar');
    assert(!!scale?.colorScale, '18a: colourScale rule parsed');
    assert(scale.colorScale.cfvos.length === 3, `18b: 3-stop scale (got ${scale.colorScale.cfvos.length})`);
    assert(!!bar?.dataBar, '18c: dataBar rule parsed');
    assert(bar.dataBar.color && bar.dataBar.color.kind === 'rgb', '18d: dataBar colour is rgb');

    // Rendered tds: five in col A (colourScale), five in col B (dataBar).
    const scaleTds = [...container.querySelectorAll('td.xlsx-cf-colorscale')];
    const barTds = [...container.querySelectorAll('td.xlsx-cf-databar')];
    assert(scaleTds.length === 5, `18e: 5 colour-scale tds (got ${scaleTds.length})`);
    assert(barTds.length === 5, `18f: 5 data-bar tds (got ${barTds.length})`);

    // Colour-scale endpoints: "0" → red #ff0000, "100" → green #00ff00.
    const scaleByText = scaleTds.reduce((acc, td) => { acc[td.textContent] = td; return acc; }, {});
    assert(/rgb\(255,\s*0,\s*0\)/.test(scaleByText['0'].style.backgroundColor),
        `18g: "0" should be red (got "${scaleByText['0'].style.backgroundColor}")`);
    assert(/rgb\(0,\s*255,\s*0\)/.test(scaleByText['100'].style.backgroundColor),
        `18h: "100" should be green (got "${scaleByText['100'].style.backgroundColor}")`);
    // Middle stop: 50th percentile → yellow #ffff00.
    assert(/rgb\(255,\s*255,\s*0\)/.test(scaleByText['50'].style.backgroundColor),
        `18i: "50" should be yellow (got "${scaleByText['50'].style.backgroundColor}")`);

    // Data bar: "0" is at the minLength (10%), "100" is at maxLength (90%).
    // The gradient is emitted on the `background` shorthand — read the raw
    // style attribute since jsdom's style.background rolls backgroundColor
    // into the same property.
    const barByText = barTds.reduce((acc, td) => { acc[td.textContent] = td; return acc; }, {});
    const style0 = barByText['0'].getAttribute('style');
    const style100 = barByText['100'].getAttribute('style');
    assert(/linear-gradient\([^)]+\) 0 10%,/.test(style0), `18j: "0" gradient at 10% (got "${style0}")`);
    assert(/linear-gradient\([^)]+\) 0 90%,/.test(style100), `18k: "100" gradient at 90% (got "${style100}")`);

    // Mid-range: "50" should land at 50% (lenMin=10 + 0.5*(90-10) = 50).
    const style50 = barByText['50'].getAttribute('style');
    assert(/linear-gradient\([^)]+\) 0 50%,/.test(style50), `18l: "50" gradient at 50% (got "${style50}")`);
}

// ── 19. interpolateColorScale / resolveCfvo unit coverage ─────────────────
{
    const { interpolateColorScale } = globalThis.xlsx;
    assert(typeof interpolateColorScale === 'function', '19a: interpolateColorScale exported');
    // Two-stop red→green, midpoint → olive-ish.
    const stops = [{ threshold: 0, hex: '#ff0000' }, { threshold: 100, hex: '#00ff00' }];
    assert(interpolateColorScale(0, stops) === '#ff0000', '19b: value=0 → stop a');
    assert(interpolateColorScale(100, stops) === '#00ff00', '19c: value=100 → stop b');
    assert(interpolateColorScale(50, stops) === '#808000', `19d: value=50 → #808000 (got ${interpolateColorScale(50, stops)})`);
    // Out-of-range clamps to endpoints.
    assert(interpolateColorScale(-10, stops) === '#ff0000', '19e: below min clamps to stop a');
    assert(interpolateColorScale(200, stops) === '#00ff00', '19f: above max clamps to stop b');
}

// ── 20. R1C1 ↔ A1 converters + showFormulas option ────────────────────────
{
    const { a1ToR1c1, r1c1ToA1 } = globalThis.xlsx;
    assert(typeof a1ToR1c1 === 'function', '20a: a1ToR1c1 exported');
    assert(typeof r1c1ToA1 === 'function', '20b: r1c1ToA1 exported');
    // Anchor B2 (row=1, col=1). "=A1" is one left one up relative.
    assert(a1ToR1c1('=A1', 1, 1) === '=R[-1]C[-1]', `20c: A1 relative to B2 (got "${a1ToR1c1('=A1', 1, 1)}")`);
    assert(a1ToR1c1('=$A$1', 1, 1) === '=R1C1', '20d: absolute $A$1');
    assert(a1ToR1c1('=$B2', 1, 1) === '=R[0]C2', '20e: mixed $B2');
    // Round-trip.
    assert(r1c1ToA1('=R[-1]C[-1]', 1, 1) === '=A1', '20f: R[-1]C[-1] → A1');
    assert(r1c1ToA1('=R1C1', 1, 1) === '=$A$1', '20g: R1C1 → $A$1');
    // SUM range.
    assert(a1ToR1c1('=SUM(A1:A3)', 3, 0) === '=SUM(R[-3]C[0]:R[-1]C[0])', `20h: SUM range (got "${a1ToR1c1('=SUM(A1:A3)', 3, 0)}")`);

    // showFormulas option on the `formulas` fixture: A2 has a SUM formula.
    // With showFormulas:true + notation:'a1', A2 should render as "=SUM(A1:C1)".
    const { container } = await renderFixture('formulas', { showFormulas: true, formulaNotation: 'a1' });
    const tds = container.querySelectorAll('section.xlsx tbody tr');
    const a2Td = tds[1].querySelectorAll('td')[0];
    assert(a2Td.textContent === '=SUM(A1:C1)', `20i: A2 should show formula text (got "${a2Td.textContent}")`);
    assert(a2Td.classList.contains('xlsx-formula'), '20j: formula cell should have .xlsx-formula class');

    // R1C1 notation: A2 anchor is row=1 col=0. =SUM(A1:C1) → col+0..2, row-1..-1.
    const { container: r1c1Container } = await renderFixture('formulas', { showFormulas: true, formulaNotation: 'r1c1' });
    const r1c1Tds = r1c1Container.querySelectorAll('section.xlsx tbody tr');
    const a2R1c1 = r1c1Tds[1].querySelectorAll('td')[0];
    assert(a2R1c1.textContent === '=SUM(R[-1]C[0]:R[-1]C[2])', `20k: A2 R1C1 (got "${a2R1c1.textContent}")`);
}

// ── 21. iconSet cf: SVG icons prepended to matching cells ─────────────────
{
    const { wb, container } = await renderFixture('cf-icons');
    const sheet = wb.parsed.sheets[0];
    const iconRule = sheet.conditionalFormatting.flatMap((b) => b.rules).find((r) => r.type === 'iconSet');
    assert(!!iconRule?.iconSet, '21a: iconSet rule parsed');
    assert(iconRule.iconSet.iconSet === '3TrafficLights1', `21b: iconSet name (got ${iconRule.iconSet.iconSet})`);

    const iconTds = [...container.querySelectorAll('td.xlsx-cf-iconset')];
    assert(iconTds.length === 5, `21c: 5 icon tds (got ${iconTds.length})`);
    // Each should contain a prepended <span class="xlsx-cf-icon"> with an SVG.
    for (const td of iconTds) {
        const span = td.querySelector('span.xlsx-cf-icon');
        assert(!!span, '21d: icon span should be present');
        assert(span.querySelector('svg'), '21e: icon span should contain an SVG');
    }
    // Lowest value (10) → red (#d13438); highest (100) → green (#107c10).
    const byText = iconTds.reduce((acc, td) => {
        // First text node after the icon span.
        const span = td.querySelector('span.xlsx-cf-icon');
        acc[td.textContent.slice(span ? 0 : 0)] = td;
        return acc;
    }, {});
    // Instead of parsing the text post-span, look up by dataset value.
    const low = iconTds.find((td) => td.textContent === '10');
    const high = iconTds.find((td) => td.textContent === '100');
    assert(!!low && !!high, '21f: 10 and 100 tds should be present');
    const lowSvg = low.querySelector('svg circle');
    const highSvg = high.querySelector('svg circle');
    assert(lowSvg && lowSvg.getAttribute('fill') === '#d13438', `21g: lowest-tier circle should be red (got ${lowSvg?.getAttribute('fill')})`);
    assert(highSvg && highSvg.getAttribute('fill') === '#107c10', `21h: highest-tier circle should be green (got ${highSvg?.getAttribute('fill')})`);
}

// ── 22. Frozen panes + autoFilter + table caption ─────────────────────────
{
    const { wb, container } = await renderFixture('tables');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.frozenPanes !== null, '22a: frozenPanes should be present');
    assert(sheet.frozenPanes.xSplit === 1 && sheet.frozenPanes.ySplit === 1, `22b: freeze B2 → xSplit=1 ySplit=1 (got ${JSON.stringify(sheet.frozenPanes)})`);
    assert(sheet.autoFilter !== null, '22c: autoFilter should be present');
    assert(sheet.autoFilter.col === 0 && sheet.autoFilter.endCol === 3, `22d: autoFilter A1:D5 span (got ${JSON.stringify(sheet.autoFilter)})`);
    assert(sheet.tables.length === 1, `22e: one table (got ${sheet.tables.length})`);
    assert(sheet.tables[0].displayName === 'InventoryTable', `22f: table displayName (got ${sheet.tables[0].displayName})`);

    // DOM: frozen cells have sticky classes; first row header cells carry
    // .xlsx-autofilter; a table caption is emitted.
    const frozenRow = container.querySelectorAll('td.xlsx-frozen-row, td.xlsx-frozen-both');
    const frozenCol = container.querySelectorAll('td.xlsx-frozen-col, td.xlsx-frozen-both');
    assert(frozenRow.length > 0, '22g: frozen-row tds should exist');
    assert(frozenCol.length > 0, '22h: frozen-col tds should exist');

    const afCells = container.querySelectorAll('td.xlsx-autofilter');
    assert(afCells.length === 4, `22i: autofilter tds should span 4 header cells (got ${afCells.length})`);

    const caption = container.querySelector('.xlsx-table-caption');
    assert(!!caption, '22j: table caption should be rendered');
    assert(caption.textContent.includes('InventoryTable'), `22k: caption should name the table (got "${caption.textContent}")`);
    assert(caption.getAttribute('data-table-name') === 'InventoryTable', '22l: caption should carry data-table-name');
}

// ── 23. Images: anchor + dataURL ──────────────────────────────────────────
{
    const { wb, container } = await renderFixture('image');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.images.length === 1, `23a: one image (got ${sheet.images.length})`);
    const img = sheet.images[0];
    assert(img.col === 1 && img.row === 2, `23b: anchor B3 → col=1 row=2 (got ${img.col},${img.row})`);
    assert(img.dataUrl.startsWith('data:image/png;base64,'), `23c: dataUrl should be a PNG data URL (got "${img.dataUrl.slice(0, 30)}...")`);

    const figs = container.querySelectorAll('figure.xlsx-image');
    assert(figs.length === 1, `23d: one figure (got ${figs.length})`);
    const imgEl = figs[0].querySelector('img');
    assert(imgEl.src.startsWith('data:image/png;base64,'), '23e: img src should be embedded data URL');
    // 32×32 PNG → 32x32 rendered dimensions (304800 EMU / 9525 = 32 px).
    assert(imgEl.width === 32 && imgEl.height === 32, `23f: img dimensions should be 32x32 (got ${imgEl.width}x${imgEl.height})`);
    assert(figs[0].getAttribute('data-anchor-col') === '1', '23g: figure should carry data-anchor-col');
    assert(figs[0].getAttribute('data-anchor-row') === '2', '23h: figure should carry data-anchor-row');
}

// ── 24. Number-format unit coverage (direct call surface) ─────────────────
// These guard the number-format formatter against corner cases that aren't
// easy to hit via a fixture (explicit formatCode strings rather than cell
// IDs), so regressions surface here rather than as rendering drift.
{
    const fn = globalThis.xlsx.formatNumber;
    assert(typeof fn === 'function', '24a: formatNumber should be re-exported from the UMD');
    assert(fn('1234.5', '#,##0.00').text === '1,234.50', '24b: thousands + decimals');
    assert(fn('0.5', '0%').text === '50%', '24c: percent no decimals');
    assert(fn('0.125', '0.00%').text === '12.50%', '24d: percent with decimals');
    assert(fn('46036', 'yyyy-mm-dd').text === '2026-01-14', '24e: date yyyy-mm-dd');
    assert(fn('46036', 'd-mmm-yy').text === '14-Jan-26', '24f: date d-mmm-yy');
    assert(fn('-50', '#,##0;(#,##0)').text === '(50)', '24g: negative section in parens');
    assert(fn('hello', '@').text === 'hello', '24h: @ echoes text');
    assert(fn('1234', 'General').text === '1234', '24i: General integer');
}

// ── 25. Chart detection + placeholder rendering ─────────────────────────
// Uses a chart-bearing workbook copied in from the 365 corpus
// (chart_sunburst.xlsx) so we exercise a real chartEx part and the
// mc:AlternateContent wrapping Excel emits.
{
    const { wb, container } = await renderFixture('chart-detect');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.charts.length >= 1, `25a: sheet should detect ≥1 chart (got ${sheet.charts.length})`);
    const chart = sheet.charts[0];
    assert(chart.kind === 'chartex' || chart.kind === 'classic', `25b: chart.kind enum (got ${chart.kind})`);
    // Anchor should come off the drawing's <from>. chart_sunburst anchors at
    // F11 (col=5, row=10); we check finite numbers rather than exact match
    // so the assertion survives cosmetic fixture drift.
    assert(Number.isFinite(chart.col) && Number.isFinite(chart.row), '25c: chart anchor should be finite');
    // Placeholder emitted after the table.
    const ph = container.querySelector('.xlsx-chart-placeholder');
    assert(!!ph, '25d: a .xlsx-chart-placeholder element should be rendered');
    assert(ph.getAttribute('data-chart-kind') === chart.kind, `25e: placeholder data-chart-kind should match (got "${ph.getAttribute('data-chart-kind')}")`);
    assert(/\[chart: /.test(ph.textContent ?? ''), `25f: placeholder text should contain "[chart: …]" (got "${ph.textContent}")`);
    // extLst rollup should surface on the sheet model.
    assert(Array.isArray(sheet.extensions), '25g: sheet.extensions should be an array');
}

// ── 26. Classic comments: xl/comments*.xml parsed + rendered as markers ──
{
    const { wb, container } = await renderFixture('comments');
    const sheet = wb.parsed.sheets[0];

    // Parser: two comments at A1 (Alice) and B2 (Bob), both plain text.
    assert(sheet.comments.length === 2, `26a: expected 2 comments (got ${sheet.comments.length})`);
    const byKey = sheet.comments.reduce((acc, c) => {
        acc[`${c.row},${c.col}`] = c;
        return acc;
    }, {});
    const a1 = byKey['0,0'];
    const b2 = byKey['1,1'];
    assert(a1 && a1.author === 'Alice' && a1.text === 'look here',
        `26b: A1 comment (got ${JSON.stringify(a1)})`);
    assert(b2 && b2.author === 'Bob' && b2.text === 'total count',
        `26c: B2 comment (got ${JSON.stringify(b2)})`);
    // Plain-text <t> body → no runs surface.
    assert(a1.runs === null, `26d: plain-text comment should have no runs (got ${JSON.stringify(a1.runs)})`);

    // DOM: each anchor cell has a .xlsx-comment-marker whose title attribute
    // concatenates "author: text" and whose role is "note".
    const markers = container.querySelectorAll('span.xlsx-comment-marker');
    assert(markers.length === 2, `26e: expected 2 comment markers rendered (got ${markers.length})`);

    const tbodyRows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (rIdx, cIdx) => tbodyRows[rIdx].querySelectorAll('td')[cIdx];

    const a1Marker = tdAt(0, 0).querySelector('span.xlsx-comment-marker');
    assert(!!a1Marker, '26f: A1 should carry a comment marker');
    assert(a1Marker.getAttribute('role') === 'note', `26g: A1 marker role=note (got ${a1Marker.getAttribute('role')})`);
    const a1Title = a1Marker.getAttribute('title') ?? '';
    assert(a1Title.includes('Alice') && a1Title.includes('look here'),
        `26h: A1 title should contain author + text (got "${a1Title}")`);

    const b2Marker = tdAt(1, 1).querySelector('span.xlsx-comment-marker');
    assert(!!b2Marker, '26i: B2 should carry a comment marker');
    const b2Title = b2Marker.getAttribute('title') ?? '';
    assert(b2Title.includes('Bob') && b2Title.includes('total count'),
        `26j: B2 title should contain author + text (got "${b2Title}")`);

    // A1's text content keeps its original value plus the visible marker
    // glyph — the cell content itself isn't destroyed.
    assert(tdAt(0, 0).textContent.startsWith('hello'),
        `26k: A1 cell text should still start with "hello" (got "${tdAt(0, 0).textContent}")`);
}

// ── 27. Threaded comments: parent + reply thread, 💬 marker title ─────────
{
    const { wb, container } = await renderFixture('threaded-comments');
    const sheet = wb.parsed.sheets[0];

    // Workbook persons registry parsed: two authors.
    assert(wb.parsed.persons.size === 2, `27a: persons registry should have 2 entries (got ${wb.parsed.persons.size})`);
    assert([...wb.parsed.persons.values()].includes('Alice'), '27b: Alice should be registered');
    assert([...wb.parsed.persons.values()].includes('Bob'), '27c: Bob should be registered');

    // Sheet.threadedComments populated with the parent + reply.
    assert(sheet.threadedComments.length === 2, `27d: expected 2 threaded comments (got ${sheet.threadedComments.length})`);
    const byAuthor = sheet.threadedComments.reduce((acc, c) => { acc[c.author ?? 'null'] = c; return acc; }, {});
    assert(byAuthor['Alice'] && byAuthor['Alice'].parentId === null, '27e: Alice is the thread starter');
    assert(byAuthor['Bob'] && byAuthor['Bob'].parentId === byAuthor['Alice'].id,
        `27f: Bob's comment should reference Alice's id as parentId (got ${byAuthor['Bob']?.parentId})`);
    assert(byAuthor['Alice'].col === 0 && byAuthor['Alice'].row === 0, '27g: anchored at A1');
    assert(byAuthor['Alice'].text === 'Please double-check this figure.', `27h: Alice text (got ${byAuthor['Alice']?.text})`);
    assert(byAuthor['Bob'].text === 'Checked — looks correct.', `27i: Bob text (got ${byAuthor['Bob']?.text})`);

    // DOM: one 💬 marker on A1, its title contains both authors + both texts
    // in chronological (parent-first) order.
    const markers = container.querySelectorAll('span.xlsx-threaded');
    assert(markers.length === 1, `27j: expected 1 threaded marker (got ${markers.length})`);
    const marker = markers[0];
    assert(marker.classList.contains('xlsx-comment-marker'), '27k: marker carries .xlsx-comment-marker');
    assert(marker.getAttribute('role') === 'note', '27l: marker role=note');
    assert(marker.textContent === '💬', `27m: marker glyph (got ${marker.textContent})`);
    const title = marker.getAttribute('title') ?? '';
    assert(title.includes('Alice'), `27n: title should mention Alice (got "${title}")`);
    assert(title.includes('Bob'), `27o: title should mention Bob (got "${title}")`);
    assert(title.includes('double-check'), `27p: title should include Alice's text (got "${title}")`);
    assert(title.includes('Checked'), `27q: title should include Bob's reply (got "${title}")`);
    assert(title.indexOf('Alice') < title.indexOf('Bob'), `27r: parent (Alice) should precede reply (Bob) in title`);
}

// ── 28. Alignment flags: wrapText + textRotation parse + apply ───────────
{
    const { wb, container } = await renderFixture('alignment-flags');
    const styles = wb.parsed.styles;
    // 8 cellXfs (index 0 = default, 1..7 carry each new flag).
    assert(styles.cellXfs.length === 8, `28a: expected 8 cellXfs (got ${styles.cellXfs.length})`);

    // Parser: wrapText flag on xf[1].
    assert(styles.cellXfs[1].alignment.wrapText === true, `28b: xf[1] wrapText (got ${styles.cellXfs[1].alignment.wrapText})`);
    // textRotation=90 on xf[2], textRotation=255 (stacked) on xf[3].
    assert(styles.cellXfs[2].alignment.textRotation === 90, `28c: xf[2] textRotation 90 (got ${styles.cellXfs[2].alignment.textRotation})`);
    assert(styles.cellXfs[3].alignment.textRotation === 255, `28d: xf[3] textRotation 255 (stacked) (got ${styles.cellXfs[3].alignment.textRotation})`);

    // DOM: A1 wrapText → white-space:normal + word-break.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (r) => rows[r].querySelectorAll('td')[0];
    const a1 = tdAt(0);
    assert(a1.style.whiteSpace === 'normal', `28e: wrapText td should set white-space:normal (got "${a1.style.whiteSpace}")`);
    assert(a1.style.wordBreak === 'break-word', `28f: wrapText td should set word-break (got "${a1.style.wordBreak}")`);

    // A2 textRotation=90 → transform:rotate(-90deg) + inline-block.
    const a2 = tdAt(1);
    assert(/rotate\(-90deg\)/.test(a2.style.transform), `28g: rotated td should carry rotate(-90deg) (got "${a2.style.transform}")`);

    // A3 textRotation=255 → writing-mode:vertical-lr.
    const a3 = tdAt(2);
    assert(a3.style.writingMode === 'vertical-lr', `28h: stacked td should set writing-mode:vertical-lr (got "${a3.style.writingMode}")`);
}

// ── 29. Alignment flags: indent + readingOrder + shrinkToFit ──────────────
{
    const { wb, container } = await renderFixture('alignment-flags');
    const styles = wb.parsed.styles;
    // Parser: indent, readingOrder, shrinkToFit flags.
    assert(styles.cellXfs[4].alignment.indent === 3, `29a: xf[4] indent 3 (got ${styles.cellXfs[4].alignment.indent})`);
    assert(styles.cellXfs[4].alignment.horizontal === 'left', `29b: xf[4] horizontal left (got ${styles.cellXfs[4].alignment.horizontal})`);
    assert(styles.cellXfs[5].alignment.readingOrder === 2, `29c: xf[5] readingOrder=2 (got ${styles.cellXfs[5].alignment.readingOrder})`);
    assert(styles.cellXfs[6].alignment.shrinkToFit === true, `29d: xf[6] shrinkToFit (got ${styles.cellXfs[6].alignment.shrinkToFit})`);

    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (r) => rows[r].querySelectorAll('td')[0];

    // A4: indent=3 with horizontal=left → paddingLeft ≈ 1.5em.
    const a4 = tdAt(3);
    // jsdom normalises '1.50em' to '1.5em' — accept either form.
    assert(a4.style.paddingLeft === '1.5em' || a4.style.paddingLeft === '1.50em',
        `29e: indent td paddingLeft should be ~1.5em (got "${a4.style.paddingLeft}")`);
    assert(a4.style.textAlign === 'left', `29f: horizontal left should land on text-align (got "${a4.style.textAlign}")`);

    // A5: readingOrder=2 → direction:rtl, and horizontal=right stays right.
    const a5 = tdAt(4);
    assert(a5.style.direction === 'rtl', `29g: readingOrder=2 should set direction:rtl (got "${a5.style.direction}")`);
    assert(a5.style.textAlign === 'right', `29h: horizontal right stays right under RTL (got "${a5.style.textAlign}")`);

    // A6: shrinkToFit → .xlsx-shrink-to-fit class.
    const a6 = tdAt(5);
    assert(a6.classList.contains('xlsx-shrink-to-fit'), `29i: shrinkToFit td should have .xlsx-shrink-to-fit (got classes: "${a6.className}")`);
}

// ── 30. Alignment flags: widened horizontal + vertical enums ──────────────
{
    const { wb, container } = await renderFixture('alignment-flags');
    const styles = wb.parsed.styles;

    // Parser: the widened enums are preserved verbatim on the model.
    assert(styles.cellXfs[7].alignment.horizontal === 'distributed',
        `30a: xf[7] horizontal=distributed parsed (got ${styles.cellXfs[7].alignment.horizontal})`);
    assert(styles.cellXfs[7].alignment.vertical === 'justify',
        `30b: xf[7] vertical=justify parsed (got ${styles.cellXfs[7].alignment.vertical})`);

    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a7 = rows[6].querySelectorAll('td')[0];
    // distributed → text-align:justify + text-align-last:justify.
    assert(a7.style.textAlign === 'justify',
        `30c: distributed horizontal → text-align:justify (got "${a7.style.textAlign}")`);
    assert(a7.style.textAlignLast === 'justify',
        `30d: distributed horizontal → text-align-last:justify (got "${a7.style.textAlignLast}")`);
    // vertical=justify → vertical-align:middle (CSS has no direct equivalent).
    assert(a7.style.verticalAlign === 'middle',
        `30e: vertical=justify should approximate as middle (got "${a7.style.verticalAlign}")`);

    // And the default xf[0] carries neutral alignment defaults.
    assert(styles.cellXfs[0].alignment.wrapText === false, '30f: default alignment wrapText=false');
    assert(styles.cellXfs[0].alignment.indent === 0, '30g: default alignment indent=0');
    assert(styles.cellXfs[0].alignment.textRotation === null, '30h: default alignment textRotation=null');
    assert(styles.cellXfs[0].alignment.readingOrder === 0, '30i: default alignment readingOrder=0');
}

// ── 31. Font extras: strike + sub/superscript via rich-text runs ──────────
{
    const { wb, container } = await renderFixture('font-extras');
    const sheet = wb.parsed.sheets[0];

    // Parser: A1 has 5 runs — struck / " H" / subscript "2" / "O " / super "sup".
    const a1 = sheet.rows[0].find((c) => c.col === 0);
    assert(Array.isArray(a1.runs) && a1.runs.length === 5, `31a: A1 should have 5 runs (got ${a1.runs?.length})`);
    assert(a1.runs[0].strike === true, `31b: run 0 should carry strike=true (got ${a1.runs[0].strike})`);
    assert(a1.runs[2].vertAlign === 'subscript', `31c: run 2 vertAlign=subscript (got ${a1.runs[2].vertAlign})`);
    assert(a1.runs[4].vertAlign === 'superscript', `31d: run 4 vertAlign=superscript (got ${a1.runs[4].vertAlign})`);

    // DOM: struck run carries line-through; sub/super runs carry verticalAlign.
    const tbodyRows = container.querySelectorAll('section.xlsx tbody tr');
    const a1Td = tbodyRows[0].querySelectorAll('td')[0];
    const spans = a1Td.querySelectorAll('span');
    assert(spans.length === 5, `31e: A1 should render 5 spans (got ${spans.length})`);
    assert(/line-through/.test(spans[0].style.textDecoration),
        `31f: struck run should have text-decoration line-through (got "${spans[0].style.textDecoration}")`);
    assert(spans[2].style.verticalAlign === 'sub',
        `31g: subscript run verticalAlign=sub (got "${spans[2].style.verticalAlign}")`);
    assert(spans[2].style.fontSize === '0.8em',
        `31h: subscript run font-size should be 0.8em (got "${spans[2].style.fontSize}")`);
    assert(spans[4].style.verticalAlign === 'super',
        `31i: superscript run verticalAlign=super (got "${spans[4].style.verticalAlign}")`);
}

// ── 32. Font extras: double underline on a cell-level font ────────────────
{
    const { wb, container } = await renderFixture('font-extras');
    const styles = wb.parsed.styles;

    // Parser: fonts[1].underline should be 'double'.
    assert(styles.fonts[1].underline === 'double',
        `32a: fonts[1].underline should be 'double' (got ${JSON.stringify(styles.fonts[1].underline)})`);
    // Default font is not underlined.
    assert(styles.fonts[0].underline === null,
        `32b: fonts[0].underline should be null (got ${JSON.stringify(styles.fonts[0].underline)})`);

    // DOM: A2 carries the double-underline font. td should have
    // text-decoration: underline and text-decoration-style: double.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a2 = rows[1].querySelectorAll('td')[0];
    assert(/underline/.test(a2.style.textDecoration),
        `32c: A2 td text-decoration should include underline (got "${a2.style.textDecoration}")`);
    assert(a2.style.textDecorationStyle === 'double',
        `32d: A2 td text-decoration-style=double (got "${a2.style.textDecorationStyle}")`);
}

// ── 33. sanitizeFontFamily: safe "Calibri" applied, injection blocked ─────
{
    const { sanitizeFontFamily } = globalThis.xlsx;
    assert(typeof sanitizeFontFamily === 'function', '33a: sanitizeFontFamily should be re-exported from the UMD');

    // Accept paths.
    assert(sanitizeFontFamily('Calibri') === '"Calibri"', `33b: "Calibri" should round-trip as "\"Calibri\"" (got ${sanitizeFontFamily('Calibri')})`);
    assert(sanitizeFontFamily('Times New Roman') === '"Times New Roman"', `33c: spaces allowed (got ${sanitizeFontFamily('Times New Roman')})`);
    assert(sanitizeFontFamily('  Arial  ') === '"Arial"', `33d: trims whitespace (got ${sanitizeFontFamily('  Arial  ')})`);

    // Rejection paths.
    const hostile = [
        null,
        undefined,
        '',
        '   ',
        'Arial;display:none',
        'Arial; display: block',
        '"><script>alert(1)</script>',
        'Arial\n{display:none}',
        '}body{display:none;',
        'Arial,sans-serif',          // commas aren't allowed (blocks fallback injection)
        'Arial\\20',                 // backslash escape attempt
        'url(x)',
    ];
    for (const v of hostile) {
        const out = sanitizeFontFamily(v);
        assert(out === null, `33e: hostile input ${JSON.stringify(v)} should be rejected (got ${JSON.stringify(out)})`);
    }

    // Rendered side: A3 uses a safe "Calibri" font → td.style.fontFamily = "\"Calibri\"".
    // A4 uses an attacker font "Arial; display:block" → td.style.fontFamily empty.
    const { container } = await renderFixture('font-extras');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a3 = rows[2].querySelectorAll('td')[0];
    const a4 = rows[3].querySelectorAll('td')[0];
    // jsdom normalises quoted font names — either the literal "\"Calibri\"" or "Calibri" is acceptable.
    const a3Fam = a3.style.fontFamily;
    assert(a3Fam === '"Calibri"' || a3Fam === 'Calibri',
        `33f: A3 fontFamily should be Calibri (quoted) (got "${a3Fam}")`);
    assert(a4.style.fontFamily === '',
        `33g: A4 attacker font name should be rejected → no fontFamily (got "${a4.style.fontFamily}")`);
}

// ── 34. Elapsed-time markers: [h], [mm], [ss] accumulate past their modulo ─
// `[h]:mm` for serial 1.5 means 36 hours (= 24 from the whole day + 12 from
// the fractional half) and the remaining minutes are 0. Same shape for `[mm]`
// (total-minute accumulator) and `[ss]` (total-second accumulator).
{
    const fn = globalThis.xlsx.formatNumber;
    assert(fn('1.5', '[h]:mm').text === '36:00',
        `34a: [h]:mm of 1.5 should be "36:00" (got "${fn('1.5', '[h]:mm').text}")`);
    assert(fn('36.5', '[mm]:ss').text === '52560:00',
        `34b: [mm]:ss of 36.5 should be "52560:00" (got "${fn('36.5', '[mm]:ss').text}")`);
    // A sub-day fraction still works: 0.25 → 6 hours.
    assert(fn('0.25', '[h]').text === '6',
        `34c: [h] of 0.25 should be "6" (got "${fn('0.25', '[h]').text}")`);
    // [hh] pads to two digits.
    assert(fn('0.25', '[hh]').text === '06',
        `34d: [hh] of 0.25 should be "06" (got "${fn('0.25', '[hh]').text}")`);
    // And [ss] alone gives the total seconds (2 minutes 30 seconds).
    assert(fn(String(150 / 86400), '[ss]').text === '150',
        `34e: [ss] of 150s should be "150" (got "${fn(String(150 / 86400), '[ss]').text}")`);
}

// ── 35. Accounting padding (_( _) _-) emits a space per _<char> ───────────
// `_("$"* #,##0_)` is the canonical Excel Accounting positive section: a
// leading `_(` (space for the matching `(` in the negative section), the
// `$` literal, a `*` fill with space, then the digits, then `_)` to align
// the trailing paren. We can't render the `*` fill (see file-header note),
// but the `_(` / `_)` padding should emit as single spaces.
{
    const fn = globalThis.xlsx.formatNumber;
    const res = fn('1234', '_("$"* #,##0_)');
    // Exact shape: [space][$][space from *-stripped / literal][1,234][space].
    // We don't pin the exact whitespace count since the "*<char>" collapse
    // is a documented compromise; instead we check the digits and currency.
    assert(res.text.includes('$1,234'), `35a: accounting form should include "$1,234" (got "${res.text}")`);
    assert(res.text.startsWith(' '), `35b: accounting form should start with leading _ space (got "${res.text}")`);
    assert(res.text.endsWith(' '), `35c: accounting form should end with trailing _ space (got "${res.text}")`);
    // Negative section with _(…-_): the `_-` pair should also render as one space.
    const neg = fn('-50', '#,##0_);(#,##0)');
    // -50 hits the second section (…) → "(50) "? Actually pattern 2 is
    // "(#,##0)" with no `_`. We only want to check this doesn't regress
    // the existing behaviour — "(50)" (and no extraneous underscore).
    assert(!neg.text.includes('_'), `35d: negative section should not surface the underscore (got "${neg.text}")`);
    assert(neg.text.includes('(50)'), `35e: negative section should render "(50)" (got "${neg.text}")`);
}

// ── 36. Locale currency: [$€-2] prefixes euro symbol, [$-409] drops locale ─
// `[$<symbol>-<localeHex>]` carries a currency symbol plus a locale id. We
// want the symbol to survive the strip and emit in place; a locale-only tag
// (empty symbol) should render nothing.
{
    const fn = globalThis.xlsx.formatNumber;
    assert(fn('1234', '[$€-2]#,##0').text === '€1,234',
        `36a: [$€-2]#,##0 of 1234 should be "€1,234" (got "${fn('1234', '[$€-2]#,##0').text}")`);
    assert(fn('1234', '[$¥-411]#,##0').text === '¥1,234',
        `36b: [$¥-411]#,##0 of 1234 should be "¥1,234" (got "${fn('1234', '[$¥-411]#,##0').text}")`);
    // No symbol (locale-only): the bracket is consumed silently.
    assert(fn('1234', '[$-409]#,##0').text === '1,234',
        `36c: [$-409]#,##0 of 1234 should be "1,234" (got "${fn('1234', '[$-409]#,##0').text}")`);
    // Multi-char symbol: "USD" as a prefix.
    assert(fn('1234', '[$USD-409]#,##0').text === 'USD1,234',
        `36d: [$USD-409]#,##0 of 1234 should be "USD1,234" (got "${fn('1234', '[$USD-409]#,##0').text}")`);
    // Existing behaviour still works: bare [Red] / [>0] are still dropped.
    assert(fn('5', '[Red]#,##0').text === '5', '36e: [Red] is still dropped');
}

// ── 37. 1904 date system threads through the formatter + renderer ─────────
// `formatNumber` takes an options bag; the `date1904` flag flips the epoch
// so serial 0 is 1904-01-01 (no leap bug). The workbook parser reads the
// flag from `<workbookPr date1904="1"/>` and the renderer threads it through.
{
    const fn = globalThis.xlsx.formatNumber;
    // Serial 0 under the 1904 system = 1904-01-01 exactly.
    assert(fn('0', 'yyyy-mm-dd', { date1904: true }).text === '1904-01-01',
        `37a: serial 0 with date1904 → "1904-01-01" (got "${fn('0', 'yyyy-mm-dd', { date1904: true }).text}")`);
    // Same serial under the 1900 system → 1899-12-30 (epoch sentinel).
    assert(fn('0', 'yyyy-mm-dd').text === '1899-12-30',
        `37b: serial 0 without date1904 → "1899-12-30" (got "${fn('0', 'yyyy-mm-dd').text}")`);
    // A real date: serial 31 in the 1904 system → 1904-02-01 (no fudge).
    assert(fn('31', 'yyyy-mm-dd', { date1904: true }).text === '1904-02-01',
        `37c: serial 31 with date1904 → "1904-02-01" (got "${fn('31', 'yyyy-mm-dd', { date1904: true }).text}")`);
    // Default flag on existing workbooks is false (regression guard).
    // The python-xlsx fixture was authored under the 1900 system and its
    // date cells still render as before (scenario 7i).
    const { wb } = await renderFixture('python-xlsx');
    assert(wb.parsed.date1904 === false,
        `37d: fixture should expose date1904=false by default (got ${wb.parsed.date1904})`);
}

// ── 38. Sheet state: hidden/veryHidden drop the section; visible renders ──
{
    const { wb, container } = await renderFixture('sheet-view-state');
    const sheets = wb.parsed.sheets;
    assert(sheets.length === 3, `38a: expected 3 sheets in the parsed model (got ${sheets.length})`);
    assert(sheets[0].name === 'Alpha' && sheets[0].state === 'visible',
        `38b: Alpha should be visible (got ${sheets[0].state})`);
    assert(sheets[1].name === 'Beta' && sheets[1].state === 'hidden',
        `38c: Beta should be hidden (got ${sheets[1].state})`);
    assert(sheets[2].name === 'Gamma' && sheets[2].state === 'visible',
        `38d: Gamma should be visible (got ${sheets[2].state})`);

    // Renderer drops hidden sheets entirely — only Alpha + Gamma emit sections.
    const sections = container.querySelectorAll('section.xlsx');
    assert(sections.length === 2, `38e: expected 2 rendered sections (Alpha + Gamma), got ${sections.length}`);
    assert(sections[0].getAttribute('data-sheet-name') === 'Alpha', '38f: first section is Alpha');
    assert(sections[1].getAttribute('data-sheet-name') === 'Gamma', '38g: second section is Gamma');
}

// ── 39. Sheet view: RTL + gridlines off on Gamma ───────────────────────────
{
    const { wb, container } = await renderFixture('sheet-view-state');
    const gamma = wb.parsed.sheets[2];
    assert(gamma.view.rightToLeft === true, `39a: Gamma view.rightToLeft should be true (got ${gamma.view.rightToLeft})`);
    assert(gamma.view.showGridLines === false, `39b: Gamma view.showGridLines should be false (got ${gamma.view.showGridLines})`);
    assert(gamma.view.showRowColHeaders === true, `39c: Gamma view.showRowColHeaders should default to true (got ${gamma.view.showRowColHeaders})`);
    // Alpha carries the defaults.
    const alpha = wb.parsed.sheets[0];
    assert(alpha.view.rightToLeft === false, `39d: Alpha view.rightToLeft default false (got ${alpha.view.rightToLeft})`);
    assert(alpha.view.showGridLines === true, `39e: Alpha view.showGridLines default true (got ${alpha.view.showGridLines})`);
    assert(alpha.view.tabColor === null, `39f: Alpha view.tabColor default null (got ${JSON.stringify(alpha.view.tabColor)})`);

    // DOM: section 1 = Gamma (sections array after Beta is dropped).
    const sections = container.querySelectorAll('section.xlsx');
    const gammaSection = sections[1];
    assert(gammaSection.getAttribute('dir') === 'rtl',
        `39g: Gamma section should have dir="rtl" (got "${gammaSection.getAttribute('dir')}")`);
    assert(gammaSection.classList.contains('xlsx-no-gridlines'),
        `39h: Gamma section should carry .xlsx-no-gridlines (got classes "${gammaSection.className}")`);
    // Alpha section should not carry these flags.
    const alphaSection = sections[0];
    assert(alphaSection.getAttribute('dir') !== 'rtl',
        `39i: Alpha section should NOT be rtl (got "${alphaSection.getAttribute('dir')}")`);
    assert(!alphaSection.classList.contains('xlsx-no-gridlines'),
        '39j: Alpha section should NOT carry .xlsx-no-gridlines');
}

// ── 40. Sheet view: tabColor surfaces as data-tab-color ───────────────────
{
    const { wb, container } = await renderFixture('sheet-view-state');
    const gamma = wb.parsed.sheets[2];
    assert(gamma.view.tabColor !== null, '40a: Gamma view.tabColor should be parsed');
    assert(gamma.view.tabColor.kind === 'rgb' && gamma.view.tabColor.value === '#ff0000',
        `40b: Gamma tabColor should be rgb #ff0000 (got ${JSON.stringify(gamma.view.tabColor)})`);

    const sections = container.querySelectorAll('section.xlsx');
    const gammaSection = sections[1];
    assert(gammaSection.getAttribute('data-tab-color') === '#ff0000',
        `40c: Gamma data-tab-color should be "#ff0000" (got "${gammaSection.getAttribute('data-tab-color')}")`);
    // Alpha has no tab colour set.
    const alphaSection = sections[0];
    assert(alphaSection.getAttribute('data-tab-color') === null,
        `40d: Alpha should have no data-tab-color (got "${alphaSection.getAttribute('data-tab-color')}")`);
}

// ── 44. Row outline levels: parsed model + data attributes on <tr> ────────
// The fixture has five rows at outline levels 0, 1, 2, 1, 0. Row 3 (index 2)
// sits at the deepest level; rows 2 + 4 bracket it at level 1.
{
    const { wb, container } = await renderFixture('outlines-and-names');
    const sheet = wb.parsed.sheets[0];

    // Parser: rowDimensions carries every row that has an outlineLevel, even
    // without height / hidden flags. Expect 3 entries (rows 1..3, indices 1..3).
    const byRow = sheet.rowDimensions.reduce((acc, d) => { acc[d.row] = d; return acc; }, {});
    assert(byRow[1]?.outlineLevel === 1, `44a: row 1 outlineLevel 1 (got ${byRow[1]?.outlineLevel})`);
    assert(byRow[2]?.outlineLevel === 2, `44b: row 2 outlineLevel 2 (got ${byRow[2]?.outlineLevel})`);
    assert(byRow[3]?.outlineLevel === 1, `44c: row 3 outlineLevel 1 (got ${byRow[3]?.outlineLevel})`);
    // Rows 0 + 4 carry outline level 0 and no other meaningful flags, so
    // rowDimensions doesn't bother listing them.
    assert(byRow[0] === undefined, `44d: row 0 should not be in rowDimensions (got ${JSON.stringify(byRow[0])})`);
    assert(byRow[4] === undefined, `44e: row 4 should not be in rowDimensions (got ${JSON.stringify(byRow[4])})`);

    // Sheet outline struct.
    assert(sheet.outline.maxRowLevel === 2, `44f: maxRowLevel 2 (got ${sheet.outline.maxRowLevel})`);
    assert(sheet.outline.maxColLevel === 1, `44g: maxColLevel 1 (got ${sheet.outline.maxColLevel})`);
    assert(sheet.outline.summaryBelow === true, `44h: summaryBelow default true (got ${sheet.outline.summaryBelow})`);
    assert(sheet.outline.summaryRight === false, `44i: summaryRight flipped false (got ${sheet.outline.summaryRight})`);

    // DOM: <tr data-outline-level="2"> + .xlsx-outline-2 on the deepest row.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    assert(rows[2].getAttribute('data-outline-level') === '2',
        `44j: deepest row should have data-outline-level="2" (got "${rows[2].getAttribute('data-outline-level')}")`);
    assert(rows[2].classList.contains('xlsx-outline-2'),
        `44k: deepest row should carry .xlsx-outline-2 (classes: "${rows[2].className}")`);
    assert(rows[1].getAttribute('data-outline-level') === '1',
        `44l: row 1 should have data-outline-level="1" (got "${rows[1].getAttribute('data-outline-level')}")`);
    assert(rows[0].getAttribute('data-outline-level') === null,
        `44m: row 0 should have no outline-level attr (got "${rows[0].getAttribute('data-outline-level')}")`);
}

// ── 45. Column outline levels: colgroup + header th data attributes ───────
{
    const { wb, container } = await renderFixture('outlines-and-names');
    const sheet = wb.parsed.sheets[0];

    // Parser: one column entry (col B = index 1) with outlineLevel=1.
    assert(sheet.columns.length === 1, `45a: one column entry (got ${sheet.columns.length})`);
    assert(sheet.columns[0].min === 1 && sheet.columns[0].max === 1,
        `45b: column B covers idx 1..1 (got min=${sheet.columns[0].min} max=${sheet.columns[0].max})`);
    assert(sheet.columns[0].outlineLevel === 1,
        `45c: column B outlineLevel 1 (got ${sheet.columns[0].outlineLevel})`);

    // DOM: the <col> for col B (gutter + A = idx 2 in colgroup) carries
    // data-outline-level + .xlsx-outline-1; same for the thead <th> at idx 2.
    const cols = container.querySelectorAll('section.xlsx colgroup col');
    assert(cols[2].getAttribute('data-outline-level') === '1',
        `45d: colgroup col B should have data-outline-level (got "${cols[2].getAttribute('data-outline-level')}")`);
    assert(cols[2].classList.contains('xlsx-outline-1'),
        `45e: colgroup col B should carry .xlsx-outline-1 (classes: "${cols[2].className}")`);

    const headers = container.querySelectorAll('section.xlsx thead th');
    // headers[0] = corner, [1] = A, [2] = B, [3] = C.
    assert(headers[2].getAttribute('data-outline-level') === '1',
        `45f: thead th B should have data-outline-level (got "${headers[2].getAttribute('data-outline-level')}")`);
    assert(headers[1].getAttribute('data-outline-level') === null,
        `45g: thead th A should have no outline-level attr`);
}

// ── 46. Workbook defined names: parse + model surface ─────────────────────
{
    const { wb } = await renderFixture('outlines-and-names');
    const names = wb.parsed.definedNames;
    assert(Array.isArray(names), '46a: definedNames should be an array');
    assert(names.length === 2, `46b: expected 2 definedNames (got ${names.length})`);

    const byName = names.reduce((acc, n) => { acc[n.name] = n; return acc; }, {});
    const total = byName['TotalRange'];
    assert(!!total, '46c: workbook-scoped "TotalRange" should be present');
    assert(total.localSheetId === null, `46d: TotalRange localSheetId should be null (got ${total.localSheetId})`);
    assert(total.formula === 'Sheet1!$A$1:$C$5', `46e: TotalRange formula (got "${total.formula}")`);
    assert(total.hidden === false, `46f: TotalRange not hidden (got ${total.hidden})`);

    const printArea = byName['_xlnm.Print_Area'];
    assert(!!printArea, '46g: print-area defined name should be present');
    assert(printArea.localSheetId === 0, `46h: print-area localSheetId should be 0 (got ${printArea.localSheetId})`);
    assert(printArea.formula === 'Sheet1!$A$1:$C$3', `46i: print-area formula (got "${printArea.formula}")`);
    assert(printArea.hidden === true, `46j: print-area hidden=true (got ${printArea.hidden})`);
}

// ── 41. Hyperlinks: model + rels → target URL for safe + unsafe entries ──
{
    const { wb } = await renderFixture('hyperlinks-and-validation');
    const sheet = wb.parsed.sheets[0];

    assert(sheet.hyperlinks.length === 2, `41a: expected 2 hyperlinks (got ${sheet.hyperlinks.length})`);
    const byKey = sheet.hyperlinks.reduce((acc, h) => { acc[`${h.row},${h.col}`] = h; return acc; }, {});
    const a2 = byKey['1,0'];
    const a3 = byKey['2,0'];
    assert(a2 && a2.target === 'https://example.com',
        `41b: A2 hyperlink target should be https://example.com (got ${JSON.stringify(a2?.target)})`);
    assert(a2.tooltip === 'visit example',
        `41c: A2 tooltip should be "visit example" (got ${JSON.stringify(a2?.tooltip)})`);
    assert(a3 && a3.target === 'javascript:alert(1)',
        `41d: A3 hyperlink target should carry the attacker URL on the model (got ${JSON.stringify(a3?.target)})`);

    // isSafeHyperlinkHref is re-exported from the UMD; verify directly.
    const { isSafeHyperlinkHref } = globalThis.xlsx;
    assert(typeof isSafeHyperlinkHref === 'function', '41e: isSafeHyperlinkHref should be re-exported');
    assert(isSafeHyperlinkHref('https://example.com') === true, '41f: https accepted');
    assert(isSafeHyperlinkHref('http://example.com') === true, '41g: http accepted');
    assert(isSafeHyperlinkHref('mailto:a@b.example') === true, '41h: mailto accepted');
    assert(isSafeHyperlinkHref('tel:+1-555-0100') === true, '41i: tel accepted');
    assert(isSafeHyperlinkHref('#anchor') === true, '41j: fragment accepted');
    assert(isSafeHyperlinkHref('foo/bar.html') === true, '41k: relative path accepted');
    // Attacker URLs: every hostile scheme should be rejected.
    for (const hostile of [
        'javascript:alert(1)',
        'JAVASCRIPT:alert(1)',
        ' javascript:alert(1)',
        'data:text/html,<script>alert(1)</script>',
        'vbscript:msgbox(1)',
        'file:///etc/passwd',
        'blob:https://evil.example/abc',
    ]) {
        assert(isSafeHyperlinkHref(hostile) === false,
            `41l: hostile URL ${JSON.stringify(hostile)} must be rejected`);
    }
}

// ── 42. Hyperlink render: safe A2 gets an anchor; unsafe A3 stays text ────
{
    const { container } = await renderFixture('hyperlinks-and-validation');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    // sheet rows: r=1 (A2) is rendered row index 1, r=2 (A3) is index 2.
    const a2Td = rows[1].querySelectorAll('td')[0];
    const a3Td = rows[2].querySelectorAll('td')[0];

    // A2: wrapped in an <a href="https://example.com">click me</a>.
    const a2Anchor = a2Td.querySelector('a.xlsx-hyperlink');
    assert(!!a2Anchor, '42a: A2 should contain an <a.xlsx-hyperlink>');
    assert(a2Anchor.getAttribute('href') === 'https://example.com',
        `42b: A2 anchor href (got ${a2Anchor?.getAttribute('href')})`);
    assert(a2Anchor.textContent === 'click me',
        `42c: A2 anchor text should be "click me" (got ${JSON.stringify(a2Anchor?.textContent)})`);
    assert(a2Anchor.getAttribute('target') === '_blank', '42d: A2 anchor target=_blank');
    assert(/noopener/.test(a2Anchor.getAttribute('rel') ?? ''), '42e: A2 anchor rel carries noopener');
    assert(a2Anchor.getAttribute('title') === 'visit example',
        `42f: A2 anchor title from tooltip (got ${a2Anchor?.getAttribute('title')})`);

    // A3: attacker URL should be REJECTED — no anchor, just the cell text.
    const a3Anchor = a3Td.querySelector('a');
    assert(a3Anchor === null, `42g: A3 attacker URL should NOT be wrapped in an anchor (got ${a3Anchor?.outerHTML})`);
    assert(a3Td.textContent === 'danger',
        `42h: A3 should render as inert text "danger" (got ${JSON.stringify(a3Td.textContent)})`);
    // Extra belt-and-braces: no element in the whole sheet carries the
    // javascript: URL, no matter where.
    const allAnchors = container.querySelectorAll('a');
    for (const a of allAnchors) {
        assert(!/javascript:/i.test(a.getAttribute('href') ?? ''),
            `42i: no anchor should carry a javascript: URL (got "${a.getAttribute('href')}")`);
    }
}

// ── 43. Data validation: B2..B4 carry list class + pipe-delimited options ─
{
    const { wb, container } = await renderFixture('hyperlinks-and-validation');
    const sheet = wb.parsed.sheets[0];

    assert(sheet.dataValidationLists.length === 1,
        `43a: expected 1 list validation range (got ${sheet.dataValidationLists.length})`);
    const v = sheet.dataValidationLists[0];
    assert(v.col === 1 && v.row === 1 && v.endCol === 1 && v.endRow === 3,
        `43b: validation range should cover B2:B4 (got ${JSON.stringify(v)})`);
    assert(Array.isArray(v.options) && v.options.join(',') === 'Red,Green,Blue',
        `43c: options should be ["Red","Green","Blue"] (got ${JSON.stringify(v.options)})`);

    // DOM: each of B2, B3, B4 should carry the list class + options attribute.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    for (const rowIdx of [1, 2, 3]) {
        const td = rows[rowIdx].querySelectorAll('td')[1];
        assert(td.classList.contains('xlsx-validation-list'),
            `43d·${rowIdx}: B${rowIdx + 1} td should carry .xlsx-validation-list (classes: "${td.className}")`);
        assert(td.getAttribute('data-validation-options') === 'Red|Green|Blue',
            `43e·${rowIdx}: B${rowIdx + 1} td should carry data-validation-options="Red|Green|Blue" (got "${td.getAttribute('data-validation-options')}")`);
    }

    // A2 / A3 / cells outside the validation range: no class applied.
    const a2Td = rows[1].querySelectorAll('td')[0];
    assert(!a2Td.classList.contains('xlsx-validation-list'),
        `43f: A2 should NOT carry the validation class`);
}

// ── 47. cf-new-rules fixture: parse containsBlanks / notContainsErrors /
//       aboveAverage rules off the sheet ─────────────────────────────────
{
    const { wb } = await renderFixture('cf-new-rules');
    const sheet = wb.parsed.sheets[0];
    assert(wb.parsed.styles.dxfs.length === 3, `47a: expected 3 dxfs (got ${wb.parsed.styles.dxfs.length})`);
    const rules = sheet.conditionalFormatting.flatMap((b) => b.rules);
    assert(rules.length === 3, `47b: expected 3 cfRules (got ${rules.length})`);
    const byType = rules.reduce((acc, r) => { acc[r.type] = r; return acc; }, {});
    assert(byType.aboveAverage, '47c: aboveAverage rule parsed');
    assert(byType.containsBlanks, '47d: containsBlanks rule parsed');
    assert(byType.notContainsErrors, '47e: notContainsErrors rule parsed');
    // Default attribute values on the parsed rule.
    assert(byType.aboveAverage.aboveAverage === true,
        `47f: aboveAverage default true (got ${byType.aboveAverage.aboveAverage})`);
    assert(byType.aboveAverage.equalAverage === false,
        `47g: equalAverage default false (got ${byType.aboveAverage.equalAverage})`);
    assert(byType.aboveAverage.stdDev === null,
        `47h: stdDev default null (got ${byType.aboveAverage.stdDev})`);
}

// ── 48. cf-new-rules: DOM wiring for containsBlanks + aboveAverage ────────
{
    const { container } = await renderFixture('cf-new-rules');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (rIdx) => rows[rIdx].querySelectorAll('td')[0];

    const a1 = tdAt(0); // val=1
    const a2 = tdAt(1); // val=2
    const a3 = tdAt(2); // val=3 (also above mean 2.75 — bold)
    const a4 = tdAt(3); // blank
    const a5 = tdAt(4); // val=5 (above mean — bold)

    // A4 (blank) should carry the red fill from the containsBlanks dxf.
    assert(a4.classList.contains('xlsx-cf'),
        `48a: A4 should carry .xlsx-cf from containsBlanks (classes: "${a4.className}")`);
    assert(/rgb\(255,\s*153,\s*153\)/.test(a4.style.backgroundColor),
        `48b: A4 should have red fill rgb(255,153,153) (got "${a4.style.backgroundColor}")`);

    // A5 matched aboveAverage → bold dxf.
    assert(a5.style.fontWeight === 'bold', `48c: A5 should be bold (got "${a5.style.fontWeight}")`);
    assert(a5.classList.contains('xlsx-cf'), `48d: A5 should carry .xlsx-cf`);
    // A3 also matched aboveAverage (3 > 2.75) → bold.
    assert(a3.style.fontWeight === 'bold', `48e: A3 should be bold (got "${a3.style.fontWeight}")`);

    // Negative: A1 (value 1) is below mean → NOT bold, and gets the
    // notContainsErrors green fill instead.
    assert(a1.style.fontWeight !== 'bold',
        `48f: A1 should NOT be bold (got "${a1.style.fontWeight}")`);
    // A2 (value 2) also below mean, should not carry the above-average bold.
    assert(a2.style.fontWeight !== 'bold',
        `48g: A2 should NOT carry above-average bold (got "${a2.style.fontWeight}")`);
    assert(/rgb\(153,\s*255,\s*153\)/.test(a1.style.backgroundColor),
        `48h: A1 should have green notContainsErrors fill (got "${a1.style.backgroundColor}")`);
    assert(/rgb\(153,\s*255,\s*153\)/.test(a2.style.backgroundColor),
        `48i: A2 should have green notContainsErrors fill (got "${a2.style.backgroundColor}")`);
}

// ── 49. evaluateRule direct: hostile / edge-case inputs ────────────────────
// These exercise the evaluator surface the harness can't hit naturally
// through rendered DOM — null cells, error cells, stdDev shift on
// aboveAverage, timePeriod against a known epoch.
{
    const { evaluateRule } = globalThis.xlsx;
    assert(typeof evaluateRule === 'function', '49a: evaluateRule re-exported');

    const mkCell = (over) => ({
        col: 0, row: 0, value: '', kind: 'number', styleIndex: -1, formula: null, runs: null,
        ...over,
    });
    const range = { col: 0, row: 0, endCol: 0, endRow: 0 };
    const emptyCtx = { cellsInRange: () => [] };

    // containsBlanks: null, empty-kind, and empty-string cells all match.
    const blanksRule = { type: 'containsBlanks', priority: 1, dxfId: 0, formulas: [], stopIfTrue: false, aboveAverage: true, equalAverage: false, stdDev: null, timePeriod: null };
    assert(evaluateRule(blanksRule, null, range, emptyCtx) === true,
        '49b: containsBlanks should match a null cell');
    assert(evaluateRule(blanksRule, mkCell({ kind: 'empty', value: '' }), range, emptyCtx) === true,
        '49c: containsBlanks should match kind=empty');
    assert(evaluateRule(blanksRule, mkCell({ kind: 'string', value: '' }), range, emptyCtx) === true,
        '49d: containsBlanks should match empty-string string cell');
    assert(evaluateRule(blanksRule, mkCell({ kind: 'string', value: 'x' }), range, emptyCtx) === false,
        '49e: containsBlanks should NOT match a non-empty cell');

    // notContainsBlanks inverts.
    const notBlanksRule = { ...blanksRule, type: 'notContainsBlanks' };
    assert(evaluateRule(notBlanksRule, mkCell({ kind: 'string', value: 'x' }), range, emptyCtx) === true,
        '49f: notContainsBlanks matches non-empty');
    assert(evaluateRule(notBlanksRule, null, range, emptyCtx) === false,
        '49g: notContainsBlanks does NOT match null');

    // containsErrors: only fires on error-kind cells. Null is not an error.
    const errRule = { ...blanksRule, type: 'containsErrors' };
    assert(evaluateRule(errRule, mkCell({ kind: 'error', value: '#DIV/0!' }), range, emptyCtx) === true,
        '49h: containsErrors matches error cell');
    assert(evaluateRule(errRule, mkCell({ kind: 'number', value: '1' }), range, emptyCtx) === false,
        '49i: containsErrors does NOT match number cell');
    assert(evaluateRule(errRule, null, range, emptyCtx) === false,
        '49j: containsErrors does NOT match null cell');

    // aboveAverage with stdDev shift: range values [1,2,3,5], mean=2.75,
    // population stddev ≈ 1.479. threshold at stdDev=1 = 4.229. Value 5
    // should match (> 4.229), value 4 should not.
    const cells = [1, 2, 3, 5].map((v, i) => ({ col: 0, row: i, cell: mkCell({ kind: 'number', value: String(v), row: i }) }));
    const ctx = { cellsInRange: () => cells };
    const avgRule = { ...blanksRule, type: 'aboveAverage', aboveAverage: true, equalAverage: false, stdDev: 1 };
    assert(evaluateRule(avgRule, mkCell({ kind: 'number', value: '5' }), range, ctx) === true,
        '49k: aboveAverage with stdDev=1 matches value 5');
    assert(evaluateRule(avgRule, mkCell({ kind: 'number', value: '4' }), range, ctx) === false,
        '49l: aboveAverage with stdDev=1 does NOT match value 4');
    // equalAverage flips to inclusive.
    const avgIncl = { ...avgRule, stdDev: null, equalAverage: true };
    assert(evaluateRule(avgIncl, mkCell({ kind: 'number', value: '2.75' }), range, ctx) === true,
        '49m: aboveAverage equalAverage=true matches mean');
    // belowAverage: aboveAverage=false.
    const belowRule = { ...avgRule, stdDev: null, aboveAverage: false };
    assert(evaluateRule(belowRule, mkCell({ kind: 'number', value: '1' }), range, ctx) === true,
        '49n: belowAverage matches value below mean');
    assert(evaluateRule(belowRule, mkCell({ kind: 'number', value: '5' }), range, ctx) === false,
        '49o: belowAverage does NOT match value above mean');

    // timePeriod=today: encode today's Excel serial from the 1900 epoch.
    const today = new Date();
    const todayUtc = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
    const epoch1900 = Date.UTC(1899, 11, 30);
    let serial = Math.floor((todayUtc - epoch1900) / 86400000);
    // 1900 leap-year fudge: add 1 when serial ≥ 60.
    if (serial >= 60) serial += 1;
    const timeCtx = { cellsInRange: () => [], date1904: false };
    const timeRule = { ...blanksRule, type: 'timePeriod', timePeriod: 'today' };
    assert(evaluateRule(timeRule, mkCell({ kind: 'number', value: String(serial) }), range, timeCtx) === true,
        `49p: timePeriod=today matches today's serial (${serial})`);
    // Unknown timePeriod → false.
    const badRule = { ...timeRule, timePeriod: 'nextMillennium' };
    assert(evaluateRule(badRule, mkCell({ kind: 'number', value: String(serial) }), range, timeCtx) === false,
        '49q: unknown timePeriod returns false');
}

// ── 50. cf-ext-databar: x14 ext attrs parsed onto the DataBar model ───────
// The sheet-level <extLst> on the cf-ext-databar fixture carries post-2010
// dataBar attributes (axisPosition, negativeFillColor, border, borderColor).
// The parser should pick up the <x14:id> on the classic cfRule's <extLst>,
// then splice the matching <x14:cfRule>'s attributes onto the DataBar.
{
    const { wb } = await renderFixture('cf-ext-databar');
    const sheet = wb.parsed.sheets[0];
    const allRules = sheet.conditionalFormatting.flatMap((b) => b.rules);
    const bar = allRules.find((r) => r.type === 'dataBar');
    assert(!!bar?.dataBar, '50a: dataBar rule should be parsed');
    // Classic dataBar bits still honoured.
    assert(bar.dataBar.color && bar.dataBar.color.kind === 'rgb', '50b: main fill colour survives ext merge');
    // x14 ext attributes.
    assert(bar.dataBar.axisPosition === 'middle',
        `50c: axisPosition should be "middle" (got ${bar.dataBar.axisPosition})`);
    assert(bar.dataBar.border === true, `50d: border flag should be true (got ${bar.dataBar.border})`);
    assert(bar.dataBar.borderColor && bar.dataBar.borderColor.kind === 'rgb' && bar.dataBar.borderColor.value === '#0000ff',
        `50e: borderColor should be rgb #0000ff (got ${JSON.stringify(bar.dataBar.borderColor)})`);
    assert(bar.dataBar.negativeFillColor && bar.dataBar.negativeFillColor.kind === 'rgb' && bar.dataBar.negativeFillColor.value === '#ff0000',
        `50f: negativeFillColor should be rgb #ff0000 (got ${JSON.stringify(bar.dataBar.negativeFillColor)})`);
    assert(bar.dataBar.negativeBorderColor && bar.dataBar.negativeBorderColor.kind === 'rgb' && bar.dataBar.negativeBorderColor.value === '#ff0000',
        `50g: negativeBorderColor should be rgb #ff0000 (got ${JSON.stringify(bar.dataBar.negativeBorderColor)})`);
    assert(bar.dataBar.axisColor && bar.dataBar.axisColor.kind === 'rgb' && bar.dataBar.axisColor.value === '#000000',
        `50h: axisColor should be rgb #000000 (got ${JSON.stringify(bar.dataBar.axisColor)})`);
    // Defaults for fields the fixture didn't declare.
    assert(bar.dataBar.gradient === true, `50i: gradient default true (got ${bar.dataBar.gradient})`);
    assert(bar.dataBar.direction === 'context', `50j: direction default "context" (got ${bar.dataBar.direction})`);
}

// ── 51. cf-ext-databar: negative fill wins on A1; border applied ──────────
// A1 holds -5 so the rendered bar on the td should land on the negative
// fill colour (red), not the main fill (#638EC6). The td should also carry
// the 1px border that the ext block declared, and the cell should be
// tagged with axis data-attributes.
{
    const { container } = await renderFixture('cf-ext-databar');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a1 = rows[0].querySelectorAll('td')[0];
    assert(a1.classList.contains('xlsx-cf-databar'), '51a: A1 should carry .xlsx-cf-databar');
    // Negative-value bar uses the red negativeFillColor. Read the style
    // attribute directly (jsdom rolls style.background into multiple sub-
    // properties, so the raw attribute is the reliable source).
    const style = a1.getAttribute('style') ?? '';
    assert(/linear-gradient\([^)]*(?:#ff0000|rgb\(255,\s*0,\s*0\))/i.test(style),
        `51b: A1 gradient should use the red negativeFillColor (got "${style}")`);
    // Border applied with the declared borderColor (blue).
    assert(/1px solid/.test(style),
        `51c: A1 should carry a 1px solid border (got "${style}")`);
    assert(/#0000ff|rgb\(0,\s*0,\s*255\)/i.test(style),
        `51d: A1 border colour should be blue (got "${style}")`);
    // Axis-position data attribute.
    assert(a1.getAttribute('data-cf-databar-axis') === 'middle',
        `51e: A1 should carry data-cf-databar-axis="middle" (got "${a1.getAttribute('data-cf-databar-axis')}")`);

    // A3 holds 3 (positive) — main fill colour, not the negative variant.
    const a3 = rows[2].querySelectorAll('td')[0];
    const a3Style = a3.getAttribute('style') ?? '';
    assert(/#638ec6|rgb\(99,\s*142,\s*198\)/i.test(a3Style),
        `51f: A3 positive bar should use the main blue fill (got "${a3Style}")`);
}

// ── 52. cf-ext-databar: dxf strike + numFmtCode "0.00" apply to matches ────
// The containsText "flag" rule fires on B1 / B3 ("flag") with dxfId=0, which
// carries <strike/> + numFmtCode "0.00". B2 / B4 don't match. A parallel
// cellIs rule on C1..C4 fires the same dxf so the numFmt re-format path
// runs against numeric values.
{
    const { wb, container } = await renderFixture('cf-ext-databar');
    const styles = wb.parsed.styles;
    assert(styles.dxfs.length === 1, `52a: expected 1 dxf (got ${styles.dxfs.length})`);
    assert(styles.dxfs[0].font?.strike === true,
        `52b: dxf[0] font.strike should be true (got ${styles.dxfs[0].font?.strike})`);
    assert(styles.dxfs[0].numFmtCode === '0.00',
        `52c: dxf[0] numFmtCode should be "0.00" (got ${JSON.stringify(styles.dxfs[0].numFmtCode)})`);

    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const tdAt = (r, c) => rows[r].querySelectorAll('td')[c];

    // B1 / B3 match the containsText rule → line-through text decoration.
    const b1 = tdAt(0, 1);
    const b3 = tdAt(2, 1);
    assert(/line-through/.test(b1.style.textDecoration),
        `52d: B1 ("flag") should have line-through (got "${b1.style.textDecoration}")`);
    assert(/line-through/.test(b3.style.textDecoration),
        `52e: B3 ("flag") should have line-through (got "${b3.style.textDecoration}")`);
    // B2 / B4 don't match → no dxf styling.
    const b2 = tdAt(1, 1);
    assert(!/line-through/.test(b2.style.textDecoration ?? ''),
        `52f: B2 ("no") should NOT have line-through (got "${b2.style.textDecoration}")`);

    // C1..C4 match cellIs >= 0 except C1 (-2). So C2..C4 should have the
    // numFmt re-applied to show "4.00", "7.00", "9.00".
    const c2 = tdAt(1, 2);
    const c3 = tdAt(2, 2);
    const c4 = tdAt(3, 2);
    assert(c2.textContent === '4.00',
        `52g: C2 should be re-formatted via dxf numFmt to "4.00" (got "${c2.textContent}")`);
    assert(c3.textContent === '7.00',
        `52h: C3 should be re-formatted via dxf numFmt to "7.00" (got "${c3.textContent}")`);
    assert(c4.textContent === '9.00',
        `52i: C4 should be re-formatted via dxf numFmt to "9.00" (got "${c4.textContent}")`);
    // C1 doesn't match cellIs >= 0 (-2 < 0) → original text preserved.
    const c1 = tdAt(0, 2);
    assert(c1.textContent === '-2',
        `52j: C1 (-2) should NOT be re-formatted (got "${c1.textContent}")`);
}

// ── 53. Format-code colour modifiers: [Red], [Blue], [Color N] ───────────
// formatNumber's FormatResult now carries an optional `color: '#rrggbb'`
// picked off the section's bracket prefix. Positive-only `[Red]…` colours
// every section; split-form `0;[Red]-0` only colours negative numbers.
{
    const fn = globalThis.xlsx.formatNumber;
    const redPos = fn('100', '[Red]#,##0');
    assert(redPos.text === '100', `53a: [Red]#,##0 of 100 should be "100" (got "${redPos.text}")`);
    assert(redPos.color === '#ff0000', `53b: [Red] should resolve to #ff0000 (got ${redPos.color})`);

    // Split format: positive section has no colour, negative carries [Red].
    // The matched section dictates the colour.
    const negOnly = fn('-100', '[Blue]0;[Red]-0');
    assert(negOnly.color === '#ff0000',
        `53c: [Red] on negative section should fire for -100 (got ${negOnly.color})`);
    const posOnly = fn('100', '[Blue]0;[Red]-0');
    assert(posOnly.color === '#0000ff',
        `53d: [Blue] on positive section should fire for 100 (got ${posOnly.color})`);

    // [Color N] uses the indexed palette (1-based in Excel). Index 3 in the
    // legacy palette (ECMA-376) is #ff0000, so [Color 3] → red.
    const palette = fn('1', '[Color 3]0');
    assert(palette.color === '#ff0000',
        `53e: [Color 3] should resolve via indexed palette to red (got ${palette.color})`);

    // Plain numeric format leaves color null.
    const noColor = fn('1', '#,##0');
    assert(noColor.color === null, `53f: plain format should have color null (got ${noColor.color})`);

    // Renderer path: the numfmt-r2 fixture's A1 uses [Red]#,##0; the td's
    // style.color must land on #ff0000 (jsdom may normalise to rgb(…)).
    const { container } = await renderFixture('numfmt-r2');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a1 = rows[0].querySelectorAll('td')[0];
    assert(a1.textContent === '100', `53g: A1 text should be "100" (got "${a1.textContent}")`);
    assert(/rgb\(255,\s*0,\s*0\)/.test(a1.style.color) || a1.style.color === '#ff0000',
        `53h: A1 td.style.color should be red (got "${a1.style.color}")`);
}

// ── 54. Conditional section predicates: [>N], [<N], [=N] ─────────────────
// The first bracket of a section can be a predicate that decides which
// section fires. Predicates win over the positional (positive/negative/zero)
// default when ANY section carries one.
{
    const fn = globalThis.xlsx.formatNumber;

    // Three-way split with [>3]"big"; [<3]"small"; "ok" fallback.
    const code = '[>3]"big";[<3]"small";"ok"';
    assert(fn('5', code).text === 'big', `54a: 5 matches [>3] (got "${fn('5', code).text}")`);
    assert(fn('1', code).text === 'small', `54b: 1 matches [<3] (got "${fn('1', code).text}")`);
    assert(fn('3', code).text === 'ok', `54c: 3 falls through to fallback (got "${fn('3', code).text}")`);

    // Equality + comparison ops.
    assert(fn('10', '[=10]"match";0').text === 'match', `54d: [=10] match`);
    assert(fn('11', '[=10]"match";0').text === '11', `54e: [=10] miss → 11`);
    assert(fn('5', '[>=5]"hi";"lo"').text === 'hi', `54f: [>=5] inclusive`);
    assert(fn('4', '[<>5]"not5";"is5"').text === 'not5', `54g: [<>5] excluded match`);
    // Predicate section preserves colour.
    const colored = fn('150', '[>100][Red]#,##0;[<0][Blue]#,##0;#,##0');
    assert(colored.text === '150', `54h: predicate + color section text (got "${colored.text}")`);
    assert(colored.color === '#ff0000', `54i: predicate section colour (got ${colored.color})`);
    // Fallback section: value 50 should land on "#,##0" with no colour.
    const fallback = fn('50', '[>100][Red]#,##0;[<0][Blue]#,##0;#,##0');
    assert(fallback.text === '50', `54j: fallback text (got "${fallback.text}")`);
    assert(fallback.color === null, `54k: fallback colour null (got ${fallback.color})`);
    // Negative section.
    const negSect = fn('-50', '[>100][Red]#,##0;[<0][Blue]#,##0;#,##0');
    assert(negSect.color === '#0000ff', `54l: negative predicate section blue (got ${negSect.color})`);

    // Renderer: numfmt-r2 A2 (value 150) → red; A3 (−50) → blue; A4 (25) default.
    const { container } = await renderFixture('numfmt-r2');
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a2 = rows[1].querySelectorAll('td')[0];
    const a3 = rows[2].querySelectorAll('td')[0];
    const a4 = rows[3].querySelectorAll('td')[0];
    assert(/rgb\(255,\s*0,\s*0\)/.test(a2.style.color) || a2.style.color === '#ff0000',
        `54m: A2 (150) should be red (got "${a2.style.color}")`);
    assert(/rgb\(0,\s*0,\s*255\)/.test(a3.style.color) || a3.style.color === '#0000ff',
        `54n: A3 (−50) should be blue (got "${a3.style.color}")`);
    assert(a4.style.color === '', `54o: A4 (25) should have no format-code colour (got "${a4.style.color}")`);
}

// ── 55. Fraction formats: # ?/? and # ??/?? best-fit fractions ────────────
// Single-digit denominators allow up to 9; double-digit up to 99. Exact
// rationals render cleanly; irrationals land on the best continued-fraction
// approximation within the denominator cap.
{
    const fn = globalThis.xlsx.formatNumber;
    assert(fn('0.25', '# ?/?').text === '0 1/4',
        `55a: 0.25 under # ?/? → "0 1/4" (got "${fn('0.25', '# ?/?').text}")`);
    assert(fn('0.5', '# ?/?').text === '0 1/2',
        `55b: 0.5 under # ?/? → "0 1/2" (got "${fn('0.5', '# ?/?').text}")`);
    // 0.333 nearest under `??` denom is 1/3.
    assert(fn('0.333', '# ??/??').text === '0 1/3',
        `55c: 0.333 under # ??/?? → "0 1/3" (got "${fn('0.333', '# ??/??').text}")`);
    // Whole + fraction: 2.75 = 2 3/4.
    assert(fn('2.75', '# ?/?').text === '2 3/4',
        `55d: 2.75 under # ?/? → "2 3/4" (got "${fn('2.75', '# ?/?').text}")`);
    // Exact integer renders without fraction.
    assert(fn('3', '# ?/?').text === '3',
        `55e: integer under fraction code → "3" (got "${fn('3', '# ?/?').text}")`);
    // Negative: magnitude fraction, leading minus.
    assert(fn('-0.5', '# ?/?').text === '-0 1/2',
        `55f: -0.5 under # ?/? (got "${fn('-0.5', '# ?/?').text}")`);
}

// ── 56. Scientific + engineering notation ─────────────────────────────────
// Classic scientific: mantissa with one integer digit. Engineering: mantissa
// with up to 3 integer digits, exponent snapped to a multiple of 3.
{
    const fn = globalThis.xlsx.formatNumber;
    assert(fn('12345.678', '0.00E+00').text === '1.23E+04',
        `56a: 12345.678 under 0.00E+00 → "1.23E+04" (got "${fn('12345.678', '0.00E+00').text}")`);
    assert(fn('12345.678', '##0.0E+0').text === '12.3E+3',
        `56b: 12345.678 under ##0.0E+0 → "12.3E+3" (got "${fn('12345.678', '##0.0E+0').text}")`);
    // Small number: 0.00045 → "4.50E-04".
    assert(fn('0.00045', '0.00E+00').text === '4.50E-04',
        `56c: 0.00045 under 0.00E+00 → "4.50E-04" (got "${fn('0.00045', '0.00E+00').text}")`);
    // Exact 1 → "1.00E+00".
    assert(fn('1', '0.00E+00').text === '1.00E+00',
        `56d: 1 under 0.00E+00 → "1.00E+00" (got "${fn('1', '0.00E+00').text}")`);
    // Zero → "0.00E+00" (special-cased).
    assert(fn('0', '0.00E+00').text === '0.00E+00',
        `56e: 0 under 0.00E+00 → "0.00E+00" (got "${fn('0', '0.00E+00').text}")`);
    // E- explicit-negative format: exponent sign only when negative.
    assert(fn('12345', '0.0E-0').text === '1.2E4',
        `56f: 12345 under 0.0E-0 → positive exponent renders with no sign (got "${fn('12345', '0.0E-0').text}")`);
}

// ── 57. Page-layout: manual row/col breaks parse + data attrs on section ──
{
    const { wb, container } = await renderFixture('page-layout');
    const sheet = wb.parsed.sheets[0];

    // Parser: two row breaks (ids 10, 20 → 0-based 9, 19), one col break
    // (id 5 → 0-based 4). Automatic breaks aren't present in this fixture
    // but the parser would filter them out.
    assert(Array.isArray(sheet.pageBreaks?.rows), '57a: pageBreaks.rows should be an array');
    assert(sheet.pageBreaks.rows.includes(9),
        `57b: pageBreaks.rows should contain 9 (got ${JSON.stringify(sheet.pageBreaks.rows)})`);
    assert(sheet.pageBreaks.rows.includes(19),
        `57c: pageBreaks.rows should contain 19 (got ${JSON.stringify(sheet.pageBreaks.rows)})`);
    assert(sheet.pageBreaks.cols.includes(4),
        `57d: pageBreaks.cols should contain 4 (got ${JSON.stringify(sheet.pageBreaks.cols)})`);

    // DOM: data-page-break-rows / data-page-break-cols on the section.
    const section = container.querySelector('section.xlsx');
    const rowAttr = section.getAttribute('data-page-break-rows');
    assert(rowAttr !== null && /\b9\b/.test(rowAttr),
        `57e: section should carry data-page-break-rows with 9 (got "${rowAttr}")`);
    assert(/\b19\b/.test(rowAttr ?? ''),
        `57f: section should carry data-page-break-rows with 19 (got "${rowAttr}")`);
    const colAttr = section.getAttribute('data-page-break-cols');
    assert(colAttr !== null && /\b4\b/.test(colAttr),
        `57g: section should carry data-page-break-cols with 4 (got "${colAttr}")`);
}

// ── 58. Page-layout: Print_Area defined name resolves to a cell range ─────
{
    const { wb } = await renderFixture('page-layout');
    const sheet = wb.parsed.sheets[0];

    // The fixture declares a localSheetId=0 `_xlnm.Print_Area` whose formula
    // is "Sheet1!$A$1:$C$5" — should resolve to a single PrintAreaRange with
    // col=0, row=0, endCol=2, endRow=4.
    assert(Array.isArray(sheet.printArea),
        `58a: printArea should be an array (got ${JSON.stringify(sheet.printArea)})`);
    assert(sheet.printArea.length === 1, `58b: printArea should have 1 entry (got ${sheet.printArea.length})`);
    const pa = sheet.printArea[0];
    assert(pa.col === 0 && pa.row === 0 && pa.endCol === 2 && pa.endRow === 4,
        `58c: printArea should resolve to {col:0,row:0,endCol:2,endRow:4} (got ${JSON.stringify(pa)})`);
}

// ── 59. Page-layout: header/footer zones + substituted &D + DOM render ────
{
    const { wb, container } = await renderFixture('page-layout');
    const sheet = wb.parsed.sheets[0];

    // Parser: oddHeader split into left/center/right; center carries today's
    // date (from the &D code). oddFooter center-only.
    assert(sheet.headerFooter !== null, '59a: headerFooter should be parsed');
    assert(sheet.headerFooter.oddHeader !== null, '59b: oddHeader should be present');
    assert(sheet.headerFooter.oddHeader.left === 'My Report',
        `59c: oddHeader.left should be "My Report" (got ${JSON.stringify(sheet.headerFooter.oddHeader.left)})`);
    // &D → today's ISO date: YYYY-MM-DD.
    const today = new Date();
    const pad = (n) => (n < 10 ? `0${n}` : String(n));
    const iso = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
    assert(sheet.headerFooter.oddHeader.center.includes(iso),
        `59d: oddHeader.center should include today's date "${iso}" (got ${JSON.stringify(sheet.headerFooter.oddHeader.center)})`);
    // &P stays literal (no pagination) — renderer emits "(page)".
    assert(sheet.headerFooter.oddHeader.right.includes('(page)'),
        `59e: oddHeader.right should include literal "(page)" for &P (got ${JSON.stringify(sheet.headerFooter.oddHeader.right)})`);
    // oddFooter: only the center zone is non-empty.
    assert(sheet.headerFooter.oddFooter !== null, '59f: oddFooter should be present');
    assert(sheet.headerFooter.oddFooter.center === 'Footer',
        `59g: oddFooter.center should be "Footer" (got ${JSON.stringify(sheet.headerFooter.oddFooter.center)})`);

    // DOM: header + footer div emitted after the table.
    const headerDiv = container.querySelector('div.xlsx-header');
    const footerDiv = container.querySelector('div.xlsx-footer');
    assert(!!headerDiv, '59h: a .xlsx-header div should be rendered');
    assert(!!footerDiv, '59i: a .xlsx-footer div should be rendered');
    const headerZones = headerDiv.querySelectorAll('[data-zone]');
    assert(headerZones.length === 3, `59j: header should have 3 zone divs (got ${headerZones.length})`);
    const byZone = (parent) => Array.from(parent.querySelectorAll('[data-zone]')).reduce(
        (acc, z) => { acc[z.getAttribute('data-zone')] = z; return acc; }, {});
    const hz = byZone(headerDiv);
    assert(hz.left.textContent === 'My Report', `59k: header left textContent (got "${hz.left.textContent}")`);
    assert(hz.center.textContent.includes(iso), `59l: header center textContent should include date (got "${hz.center.textContent}")`);
}

// ── 60. Split panes without freeze: kind='split' surfaces on FrozenPanes ──
// Regression guard for the widened FrozenPanes.kind field. The existing
// "tables" fixture uses `state="frozen"` → kind='frozen' + existing
// sticky-cell classes still apply (see scenario 22).
{
    const { wb } = await renderFixture('tables');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.frozenPanes !== null, '60a: frozenPanes should be parsed on tables fixture');
    assert(sheet.frozenPanes.kind === 'frozen',
        `60b: frozen-state fixture should set kind='frozen' (got ${JSON.stringify(sheet.frozenPanes.kind)})`);
}

// ── 65. theme-fontscheme: parseTheme lifts majorFont / minorFont latin ────
// The theme declares <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
// and <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>; the parser
// surfaces both on Theme so the renderer can resolve <font scheme="major|minor"/>.
{
    const { wb } = await renderFixture('theme-fontscheme');
    const theme = wb.parsed.theme;
    assert(theme !== null, '65a: theme should be parsed on theme-fontscheme fixture');
    assert(theme.majorFont === 'Cambria', `65b: theme.majorFont should be "Cambria" (got ${JSON.stringify(theme?.majorFont)})`);
    assert(theme.minorFont === 'Calibri', `65c: theme.minorFont should be "Calibri" (got ${JSON.stringify(theme?.minorFont)})`);
}

// ── 66. theme-fontscheme: FontStyle.scheme parsed + renderer resolves via theme
// fonts[1] has <scheme val="major"/>, fonts[2] has <scheme val="minor"/>.
// A1 uses xf[1] → majorFont ("Cambria"), A2 uses xf[2] → minorFont ("Calibri").
// sanitizeFontFamily wraps the typeface in double quotes before assigning to
// td.style.fontFamily.
{
    const { wb, container } = await renderFixture('theme-fontscheme');
    const fonts = wb.parsed.styles.fonts;
    assert(fonts[1].scheme === 'major', `66a: fonts[1].scheme should be "major" (got ${fonts[1].scheme})`);
    assert(fonts[2].scheme === 'minor', `66b: fonts[2].scheme should be "minor" (got ${fonts[2].scheme})`);
    assert(fonts[0].scheme === null, `66c: fonts[0].scheme should be null (got ${fonts[0].scheme})`);
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a1 = rows[0].querySelectorAll('td')[0];
    const a2 = rows[1].querySelectorAll('td')[0];
    assert(a1.style.fontFamily === '"Cambria"', `66d: A1 fontFamily should be '"Cambria"' (got ${JSON.stringify(a1.style.fontFamily)})`);
    assert(a2.style.fontFamily === '"Calibri"', `66e: A2 fontFamily should be '"Calibri"' (got ${JSON.stringify(a2.style.fontFamily)})`);
}

// ── 67. phonetics: parseSi lifts <rPh> + renderer emits HTML5 <ruby> markup ─
// The shared string carries base "漢字" with phonetic "かんじ" over [0..2),
// and an inline-str cell carries "東京" with "とうきょう" over [0..2). The
// model exposes both phonetic annotations; the rendered td contains a <ruby>
// with an <rt> holding the phonetic reading.
{
    const { wb, container } = await renderFixture('phonetics');
    const si = wb.parsed; // just to anchor a name; the shared strings live on the sheet's cells
    // Inspect the sheet's first cell for its PhoneticRun[] (mirrors the source si).
    const sheet = wb.parsed.sheets[0];
    const a1 = sheet.rows[0][0];
    assert(a1.phonetics !== null, '67a: A1 should carry phonetics from the shared string');
    assert(a1.phonetics.length === 1, `67b: A1 should have one phonetic entry (got ${a1.phonetics?.length})`);
    assert(a1.phonetics[0].base === '漢字', `67c: phonetic base should be "漢字" (got ${JSON.stringify(a1.phonetics?.[0].base)})`);
    assert(a1.phonetics[0].phonetic === 'かんじ', `67d: phonetic reading should be "かんじ" (got ${JSON.stringify(a1.phonetics?.[0].phonetic)})`);
    assert(a1.phonetics[0].startIdx === 0 && a1.phonetics[0].endIdx === 2,
        `67e: phonetic span should be [0..2) (got [${a1.phonetics?.[0].startIdx}..${a1.phonetics?.[0].endIdx}))`);

    // A2 (inline-string) should have phonetics too.
    const a2 = sheet.rows[1][0];
    assert(a2.phonetics !== null && a2.phonetics.length === 1, '67f: A2 (inlineStr) should carry phonetics');
    assert(a2.phonetics[0].phonetic === 'とうきょう',
        `67g: A2 phonetic reading should be "とうきょう" (got ${JSON.stringify(a2.phonetics?.[0].phonetic)})`);

    // Rendered DOM: first row's td should contain <ruby>漢字<rt>かんじ</rt></ruby>.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const td1 = rows[0].querySelectorAll('td')[0];
    const ruby1 = td1.querySelector('ruby');
    assert(!!ruby1, '67h: A1 td should contain a <ruby> element');
    const rt1 = ruby1?.querySelector('rt');
    assert(!!rt1, '67i: A1 <ruby> should contain an <rt> element');
    assert(rt1?.textContent === 'かんじ', `67j: <rt> text should be "かんじ" (got ${JSON.stringify(rt1?.textContent)})`);
    // The base kanji characters must also be inside the <ruby> (before the <rt>).
    assert(ruby1?.textContent.startsWith('漢字'),
        `67k: <ruby> should start with the base text "漢字" (got ${JSON.stringify(ruby1?.textContent)})`);

    // Inline-str row: same structure.
    const td2 = rows[1].querySelectorAll('td')[0];
    const ruby2 = td2.querySelector('ruby');
    assert(!!ruby2, '67l: A2 (inlineStr) td should contain a <ruby> element');
    assert(ruby2?.querySelector('rt')?.textContent === 'とうきょう',
        `67m: A2 <rt> text should be "とうきょう" (got ${JSON.stringify(ruby2?.querySelector('rt')?.textContent)})`);
    void si;
}

// ── 61. shapes-and-textboxes: xdr:sp + xdr:cxnSp parse into Sheet.shapes ──
// The fixture carries two twoCellAnchor entries:
//   · cxnSp at A1:A3 with prstGeom="line"            → kind='connector'
//   · sp    at B3:D5 with prstGeom="rect" + txBody   → kind='shape',
//     text = "Line 1\nLine 2"
// Sheet.shapes must carry exactly these two entries; images/charts must be
// unaffected (fixture has none). We assert kind / preset / text / anchor
// for each one rather than a single summary — future reorderings shouldn't
// hide a parser regression.
{
    const { wb } = await renderFixture('shapes-and-textboxes');
    const sheet = wb.parsed.sheets[0];
    assert(Array.isArray(sheet.shapes), '61a: Sheet.shapes should be an array');
    assert(sheet.shapes.length === 2,
        `61b: expected 2 shapes in fixture (got ${sheet.shapes.length})`);
    assert(sheet.images.length === 0,
        `61c: fixture has no images — images[] should stay empty (got ${sheet.images.length})`);
    assert(sheet.charts.length === 0,
        `61d: fixture has no charts — charts[] should stay empty (got ${sheet.charts.length})`);

    const textBox = sheet.shapes.find((s) => s.kind === 'shape');
    assert(textBox, '61e: a shape with kind="shape" should be present');
    if (textBox) {
        assert(textBox.preset === 'rect',
            `61f: text-box preset should be "rect" (got ${JSON.stringify(textBox.preset)})`);
        assert(textBox.text === 'Line 1\nLine 2',
            `61g: text-box text should flatten paragraphs joined by \\n (got ${JSON.stringify(textBox.text)})`);
        assert(textBox.col === 1 && textBox.row === 2,
            `61h: text-box anchor should be (col=1, row=2) (got col=${textBox.col}, row=${textBox.row}))`);
        assert(textBox.endCol === 3 && textBox.endRow === 4,
            `61i: text-box end anchor should be (endCol=3, endRow=4) (got endCol=${textBox.endCol}, endRow=${textBox.endRow})`);
        assert(textBox.alt === 'callout',
            `61j: text-box alt should resolve from cNvPr/@descr (got ${JSON.stringify(textBox.alt)})`);
    }

    const connector = sheet.shapes.find((s) => s.kind === 'connector');
    assert(connector, '61k: a shape with kind="connector" should be present');
    if (connector) {
        assert(connector.preset === 'line',
            `61l: connector preset should be "line" (got ${JSON.stringify(connector.preset)})`);
        assert(connector.text === null,
            `61m: connector should carry no text body (got ${JSON.stringify(connector.text)})`);
    }
}

// ── 62. shapes-and-textboxes renderer: <aside class="xlsx-shape"> emits ──
// Both shapes must surface in the DOM after the <table>: two <aside> with
// matching data-kind + data-preset + anchor attrs. The text-box must carry
// a <pre> with the paragraph-joined body (white-space:pre-line preserves
// the newline visually). The connector's border is dashed via CSS when
// data-kind="connector" — we assert the attribute, not the computed style.
{
    const { container } = await renderFixture('shapes-and-textboxes');
    const asides = container.querySelectorAll('section.xlsx aside.xlsx-shape');
    assert(asides.length === 2,
        `62a: expected 2 <aside class="xlsx-shape"> elements (got ${asides.length})`);

    // Every aside must come after the <table> in the section.
    const section = container.querySelector('section.xlsx');
    const table = section?.querySelector('table');
    if (table && asides[0]) {
        // compareDocumentPosition: DOCUMENT_POSITION_FOLLOWING = 4
        const pos = table.compareDocumentPosition(asides[0]);
        assert((pos & 4) === 4, '62b: <aside> elements should follow the <table> in the section');
    }

    const shapeAside = Array.from(asides).find((el) => el.getAttribute('data-kind') === 'shape');
    assert(shapeAside, '62c: at least one aside should have data-kind="shape"');
    if (shapeAside) {
        assert(shapeAside.getAttribute('data-preset') === 'rect',
            `62d: text-box aside should have data-preset="rect" (got ${JSON.stringify(shapeAside.getAttribute('data-preset'))})`);
        const pre = shapeAside.querySelector('pre');
        assert(!!pre, '62e: text-box aside should contain a <pre> for the text body');
        assert(pre?.textContent === 'Line 1\nLine 2',
            `62f: <pre> textContent should preserve paragraph break (got ${JSON.stringify(pre?.textContent)})`);
        assert(shapeAside.getAttribute('data-anchor-col') === '1',
            `62g: text-box data-anchor-col should be "1" (got ${JSON.stringify(shapeAside.getAttribute('data-anchor-col'))})`);
        assert(shapeAside.getAttribute('data-anchor-end-col') === '3',
            `62h: text-box data-anchor-end-col should be "3" (got ${JSON.stringify(shapeAside.getAttribute('data-anchor-end-col'))})`);
    }

    const connectorAside = Array.from(asides).find((el) => el.getAttribute('data-kind') === 'connector');
    assert(connectorAside, '62i: at least one aside should have data-kind="connector"');
    if (connectorAside) {
        assert(connectorAside.getAttribute('data-preset') === 'line',
            `62j: connector aside should have data-preset="line" (got ${JSON.stringify(connectorAside.getAttribute('data-preset'))})`);
        // Connectors in this fixture carry no text → no <pre>.
        assert(!connectorAside.querySelector('pre'),
            '62k: connector aside should not contain a <pre> (no text body)');
    }
}

// ── 63. Drawing anchors: oneCell + absolute + emuToPx helper ─────────────
{
    const { wb, container } = await renderFixture('drawing-anchors-plus');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.images.length === 3, `63a: three anchored images (got ${sheet.images.length})`);

    const byMode = sheet.images.reduce((acc, im) => { acc[im.anchorMode] = im; return acc; }, {});
    assert(byMode.oneCell, '63b: a oneCellAnchor image should be present');
    assert(byMode.twoCell, '63c: a twoCellAnchor image should be present');
    assert(byMode.absolute, '63d: an absoluteAnchor image should be present');

    // oneCellAnchor at B2 with cx=cy=914400 (1" → 96 px).
    const one = byMode.oneCell;
    assert(one.col === 1 && one.row === 1, `63e: oneCell anchor should be at B2 (got ${one.col},${one.row})`);
    assert(one.endCol === null && one.endRow === null,
        `63f: oneCell endCol/endRow should be null (got ${one.endCol},${one.endRow})`);
    assert(one.widthEmu === 914400 && one.heightEmu === 914400,
        `63g: oneCell widthEmu/heightEmu should be 914400 (got ${one.widthEmu}/${one.heightEmu})`);
    assert(one.absoluteX === null && one.absoluteY === null,
        '63h: oneCell absoluteX/Y should be null');

    // absoluteAnchor carries pos + ext.
    const abs = byMode.absolute;
    assert(abs.absoluteX === 1905000 && abs.absoluteY === 952500,
        `63i: absoluteAnchor position (EMU) should be (1905000,952500) (got ${abs.absoluteX},${abs.absoluteY})`);
    assert(abs.endCol === null && abs.endRow === null,
        '63j: absoluteAnchor endCol/endRow should be null');

    // DOM: three figures, each with data-anchor-mode. The absolute one
    // gets position:absolute + left/top in px.
    const figs = [...container.querySelectorAll('figure.xlsx-image')];
    assert(figs.length === 3, `63k: three rendered figures (got ${figs.length})`);
    const modes = figs.map((f) => f.getAttribute('data-anchor-mode')).sort();
    assert(JSON.stringify(modes) === JSON.stringify(['absolute', 'oneCell', 'twoCell']),
        `63l: data-anchor-mode should enumerate the three modes (got ${JSON.stringify(modes)})`);
    const absFig = figs.find((f) => f.getAttribute('data-anchor-mode') === 'absolute');
    assert(absFig.style.position === 'absolute', '63m: absolute figure should have CSS position:absolute');
    // 1905000 / 9525 = 200; 952500 / 9525 = 100.
    assert(absFig.style.left === '200px', `63n: absolute figure left should be 200px (got ${absFig.style.left})`);
    assert(absFig.style.top === '100px', `63o: absolute figure top should be 100px (got ${absFig.style.top})`);
    const oneFig = figs.find((f) => f.getAttribute('data-anchor-mode') === 'oneCell');
    const oneImg = oneFig.querySelector('img');
    assert(oneImg.width === 96 && oneImg.height === 96,
        `63p: oneCell img size (914400 EMU) should render as 96×96 px (got ${oneImg.width}×${oneImg.height})`);

    // emuToPx re-exported from the UMD should match the renderer's output.
    const { emuToPx } = globalThis.xlsx;
    assert(typeof emuToPx === 'function', '63q: emuToPx should be exported from the UMD');
    assert(emuToPx(914400) === 96, `63r: emuToPx(914400) === 96 (got ${emuToPx(914400)})`);
    assert(emuToPx(9525) === 1, `63s: emuToPx(9525) === 1 (got ${emuToPx(9525)})`);
    assert(emuToPx(0) === 0, `63t: emuToPx(0) === 0 (got ${emuToPx(0)})`);
}

// ── 64. Decorative images: alt="" + aria-hidden="true" ───────────────────
{
    const { wb, container } = await renderFixture('drawing-anchors-plus');
    const sheet = wb.parsed.sheets[0];

    // The twoCellAnchor image carries <adec:decorative val="1"/> in its
    // cNvPr extLst — it should flip decorative=true on the model.
    const deco = sheet.images.find((im) => im.decorative);
    assert(!!deco, `64a: one image should be flagged decorative (got ${sheet.images.map((i) => i.decorative).join(',')})`);
    assert(deco.anchorMode === 'twoCell', `64b: decorative image in fixture is the twoCellAnchor (got ${deco.anchorMode})`);

    // The non-decorative images keep their descr as alt.
    const nonDeco = sheet.images.filter((im) => !im.decorative);
    assert(nonDeco.length === 2, `64c: two non-decorative images (got ${nonDeco.length})`);
    assert(nonDeco.every((im) => typeof im.alt === 'string' && im.alt.length > 0),
        `64d: non-decorative images should carry alt text (got ${JSON.stringify(nonDeco.map((i) => i.alt))})`);

    // DOM: the decorative <img> has alt="" and aria-hidden="true". Matching
    // by data-anchor-mode lets us key it without depending on source order.
    const figs = [...container.querySelectorAll('figure.xlsx-image')];
    const decoFig = figs.find((f) => f.getAttribute('data-anchor-mode') === 'twoCell');
    const decoImg = decoFig.querySelector('img');
    assert(decoImg.getAttribute('alt') === '',
        `64e: decorative img should render alt="" (got ${JSON.stringify(decoImg.getAttribute('alt'))})`);
    assert(decoImg.getAttribute('aria-hidden') === 'true',
        `64f: decorative img should carry aria-hidden="true" (got ${JSON.stringify(decoImg.getAttribute('aria-hidden'))})`);

    // The other figures should NOT have aria-hidden set.
    const other = figs.filter((f) => f !== decoFig);
    for (const f of other) {
        const im = f.querySelector('img');
        assert(im.getAttribute('aria-hidden') === null,
            `64g: non-decorative img should not carry aria-hidden (mode=${f.getAttribute('data-anchor-mode')})`);
    }
}

// ── 71. Sheet protection: <sheetProtection> state parsed + DOM attr set ──
{
    const { wb, container } = await renderFixture('sheet-protection-and-alt');
    const sheet = wb.parsed.sheets[0];
    const p = sheet.protection;
    assert(p !== null, '71a: sheet.protection should be populated when <sheetProtection> is present');
    assert(p.enabled === true, `71b: protection.enabled should be true (got ${p?.enabled})`);
    // Explicit @selectLockedCells="1" → locked.
    assert(p.selectLockedCells === true, `71c: selectLockedCells should be true (got ${p?.selectLockedCells})`);
    // Explicit @formatCells="0" → unlocked.
    assert(p.formatCells === false, `71d: formatCells should be false (got ${p?.formatCells})`);
    // Absent @sort defaults to locked per ECMA-376 §18.3.1.85.
    assert(p.sort === true, `71e: sort should default to true when attribute absent (got ${p?.sort})`);
    // No @password → passwordHashed false.
    assert(p.passwordHashed === false, `71f: passwordHashed should be false when no hash attrs (got ${p?.passwordHashed})`);

    // DOM: section gets data-sheet-protected="true".
    const section = container.querySelector('section.xlsx');
    assert(section.getAttribute('data-sheet-protected') === 'true',
        `71g: section should carry data-sheet-protected="true" (got ${JSON.stringify(section.getAttribute('data-sheet-protected'))})`);
}

// ── 72. Table alt text: TableDef carries altText + caption gets aria-label ──
{
    const { wb, container } = await renderFixture('sheet-protection-and-alt');
    const sheet = wb.parsed.sheets[0];
    assert(sheet.tables.length === 1, `72a: one table parsed (got ${sheet.tables.length})`);
    const t = sheet.tables[0];
    assert(t.altText === 'Q4 stock', `72b: altText should be "Q4 stock" (got ${JSON.stringify(t.altText)})`);
    assert(t.altTextSummary === 'Weekly stock by SKU',
        `72c: altTextSummary should be "Weekly stock by SKU" (got ${JSON.stringify(t.altTextSummary)})`);

    // DOM: the caption carries aria-label="Q4 stock" (short altText wins over
    // the longer summary when both are present).
    const caption = container.querySelector('.xlsx-table-caption');
    assert(caption, '72d: table caption should be rendered');
    assert(caption.getAttribute('aria-label') === 'Q4 stock',
        `72e: caption aria-label should be "Q4 stock" (got ${JSON.stringify(caption?.getAttribute('aria-label'))})`);
}

// ── 73. cell-metadata: xl/metadata.xml parses into Workbook.metadata ──────
{
    const { wb } = await renderFixture('cell-metadata');
    const md = wb.parsed.metadata;
    assert(md !== null, '73a: Workbook.metadata should be non-null when xl/metadata.xml is present');
    assert(Array.isArray(md?.cellMetadata), '73b: metadata.cellMetadata should be an array');
    assert(md?.cellMetadata.length === 1,
        `73c: one cellMetadata block expected (got ${md?.cellMetadata.length})`);
    assert(md?.cellMetadata[0].dynamicArray === true,
        `73d: cellMetadata[0] should resolve to dynamicArray=true (got ${md?.cellMetadata[0].dynamicArray})`);
    // The type index should be 0 (XLDAPR is the only metadataType → 1-based 1 → 0-based 0).
    assert(md?.cellMetadata[0].typeIndex === 0,
        `73e: cellMetadata[0].typeIndex should be 0 (got ${md?.cellMetadata[0].typeIndex})`);
    // Sheet A1 references cm=1 → cellMetadataIndex=0 + isSpillAnchor=true.
    const sheet = wb.parsed.sheets[0];
    const row1 = sheet.rows[0];
    const a1 = row1[0];
    const b1 = row1[1];
    assert(a1.cellMetadataIndex === 0,
        `73f: A1.cellMetadataIndex should be 0 (got ${a1.cellMetadataIndex})`);
    assert(a1.isSpillAnchor === true, `73g: A1.isSpillAnchor should be true (got ${a1.isSpillAnchor})`);
    assert(a1.valueMetadataIndex === null,
        `73h: A1.valueMetadataIndex should be null (got ${a1.valueMetadataIndex})`);
    assert(b1.cellMetadataIndex === null,
        `73i: B1.cellMetadataIndex should be null (got ${b1.cellMetadataIndex})`);
    assert(b1.isSpillAnchor === false,
        `73j: B1.isSpillAnchor should be false (got ${b1.isSpillAnchor})`);
}

// ── 74. cell-metadata renderer: .xlsx-spill-anchor tags the A1 td ─────────
{
    const { container } = await renderFixture('cell-metadata');
    const tds = container.querySelectorAll('section.xlsx tbody tr:first-child td');
    // First tbody row: [A1, B1]. A1 carries the spill-anchor class; B1 doesn't.
    assert(tds.length >= 2, `74a: first row should have at least two tds (got ${tds.length})`);
    assert(tds[0].classList.contains('xlsx-spill-anchor'),
        `74b: A1 td should carry .xlsx-spill-anchor (classList=${tds[0].className})`);
    assert(!tds[1].classList.contains('xlsx-spill-anchor'),
        `74c: B1 td should NOT carry .xlsx-spill-anchor (classList=${tds[1].className})`);
    // Text content sanity: "anchor" / "plain".
    assert(tds[0].textContent === 'anchor', `74d: A1 should render "anchor" (got ${tds[0].textContent})`);
    assert(tds[1].textContent === 'plain', `74e: B1 should render "plain" (got ${tds[1].textContent})`);
}

// ── 75. Multi-step cellStyleXfs chain: A1 picks up deep style's bold ─────
// Fixture: cellStyleXfs[1].xfId=2 → cellStyleXfs[2] (bold). cellXfs[1] has
// xfId=1 and applyFont=0, so the walker must follow the chain two hops to
// reach the bold font. A single-hop resolver would have left A1 non-bold
// (and missed any styling, since the mid style also has applyFont=0).
{
    const { wb, container } = await renderFixture('cellstyle-chain');
    const styles = wb.parsed.styles;
    assert(styles.cellStyleXfs.length === 3, `75a: cellStyleXfs should have 3 entries (got ${styles.cellStyleXfs.length})`);
    assert(styles.cellStyleXfs[1].xfId === 2, `75b: cellStyleXfs[1].xfId should be 2 (got ${styles.cellStyleXfs[1].xfId})`);
    assert(styles.cellStyleXfs[2].xfId === -1, `75c: cellStyleXfs[2].xfId should be -1 (got ${styles.cellStyleXfs[2].xfId})`);
    assert(styles.cellXfs[1].xfId === 1, `75d: cellXfs[1].xfId should be 1 (got ${styles.cellXfs[1].xfId})`);
    assert(styles.cellXfs[1].applyFont === false, '75e: cellXfs[1].applyFont should be false');

    // Unit assertion: resolveEffectiveXf walks the chain to the bold font.
    const { resolveEffectiveXf } = globalThis.xlsx;
    assert(typeof resolveEffectiveXf === 'function', '75f: resolveEffectiveXf should be exported');
    const eff = resolveEffectiveXf(styles, styles.cellXfs[1]);
    assert(eff.fontId === 2, `75g: effective fontId should chain to 2 (got ${eff.fontId})`);

    // DOM assertion: the rendered td must be bold (fonts[2]), not italic
    // (fonts[1]) — the chain walker went one hop past the italic mid style.
    const rows = container.querySelectorAll('section.xlsx tbody tr');
    const a1 = rows[0].querySelectorAll('td')[0];
    assert(a1.style.fontWeight === 'bold', `75h: A1 should be bold via 2-hop chain (got "${a1.style.fontWeight}")`);
    assert(a1.style.fontStyle !== 'italic', `75i: A1 should NOT be italic (got "${a1.style.fontStyle}")`);

    // Cycle detection: forging a looped cellStyleXfs chain (1→2→1) must
    // terminate. Without the `seen` guard, the walker would blow the stack
    // or spin forever. Build a synthetic Styles object directly.
    const cycleStyles = {
        numFmts: new Map(),
        fonts: styles.fonts,
        fills: styles.fills,
        borders: styles.borders,
        dxfs: [],
        cellXfs: styles.cellXfs,
        cellStyleXfs: [
            { ...styles.cellStyleXfs[0] },
            { ...styles.cellStyleXfs[1], xfId: 2 },
            { ...styles.cellStyleXfs[2], xfId: 1 }, // loops back!
        ],
    };
    const start = Date.now();
    const cycleEff = resolveEffectiveXf(cycleStyles, { ...styles.cellXfs[1], xfId: 1 });
    const elapsed = Date.now() - start;
    assert(elapsed < 100, `75j: cycle detection should terminate quickly (got ${elapsed}ms)`);
    assert(typeof cycleEff.fontId === 'number', '75k: cycle walk should still return a CellXf');
}

// ── 76. Custom iconSet rule lists: per-position <cfIcon/> overrides ───────
// Fixture: one iconSet rule with custom="1" and three <cfIcon/> children
// that swap each position to a glyph from a different set. Values 10/50/90
// hit the three positions in turn.
{
    const { wb, container } = await renderFixture('cf-custom-icons');
    const sheet = wb.parsed.sheets[0];
    const iconRule = sheet.conditionalFormatting.flatMap((b) => b.rules).find((r) => r.type === 'iconSet');
    assert(!!iconRule?.iconSet, '76a: iconSet rule parsed');
    assert(Array.isArray(iconRule.iconSet.customIcons), '76b: customIcons should be an array when custom="1"');
    assert(iconRule.iconSet.customIcons.length === 3, `76c: customIcons should have 3 entries (got ${iconRule.iconSet.customIcons.length})`);
    assert(iconRule.iconSet.customIcons[0]?.iconSet === '3Arrows' && iconRule.iconSet.customIcons[0]?.iconId === 0,
        `76d: position 0 → 3Arrows/0 (got ${JSON.stringify(iconRule.iconSet.customIcons[0])})`);
    assert(iconRule.iconSet.customIcons[1]?.iconSet === '3TrafficLights1' && iconRule.iconSet.customIcons[1]?.iconId === 1,
        `76e: position 1 → 3TrafficLights1/1 (got ${JSON.stringify(iconRule.iconSet.customIcons[1])})`);
    assert(iconRule.iconSet.customIcons[2]?.iconSet === '3Symbols' && iconRule.iconSet.customIcons[2]?.iconId === 2,
        `76f: position 2 → 3Symbols/2 (got ${JSON.stringify(iconRule.iconSet.customIcons[2])})`);

    // DOM: each row should get the overridden glyph.
    //   10 → 3Arrows[0]        = down red arrow (a <path> inside a <g>)
    //   50 → 3TrafficLights1[1] = amber circle (#ffb900)
    //   90 → 3Symbols[2]        = green tick (a <path> stroke #107c10)
    const iconTds = [...container.querySelectorAll('td.xlsx-cf-iconset')];
    assert(iconTds.length === 3, `76g: 3 icon tds (got ${iconTds.length})`);

    const low = iconTds.find((td) => td.textContent === '10');
    const mid = iconTds.find((td) => td.textContent === '50');
    const high = iconTds.find((td) => td.textContent === '90');
    assert(low && mid && high, '76h: 10 / 50 / 90 tds all present');

    // Position 0 override: 3Arrows[0] → path wrapped in a <g> with rotate(180…).
    const lowG = low.querySelector('svg g');
    assert(!!lowG && /rotate\(180/.test(lowG.getAttribute('transform') ?? ''),
        `76i: value 10 should render 3Arrows[0] (rotated path; got transform="${lowG?.getAttribute('transform')}")`);

    // Position 1 override: 3TrafficLights1[1] → amber circle (#ffb900).
    const midCircle = mid.querySelector('svg circle');
    assert(!!midCircle && midCircle.getAttribute('fill') === '#ffb900',
        `76j: value 50 should render amber 3TrafficLights1[1] circle (got fill="${midCircle?.getAttribute('fill')}")`);

    // Position 2 override: 3Symbols[2] → tick path stroke #107c10, no <circle>.
    const highPath = high.querySelector('svg path');
    assert(!!highPath && highPath.getAttribute('stroke') === '#107c10',
        `76k: value 90 should render 3Symbols[2] tick (got stroke="${highPath?.getAttribute('stroke')}")`);
    assert(!high.querySelector('svg circle'),
        '76l: value 90 should NOT contain a <circle> (confirming the traffic-light default was overridden)');
}

// ── report ────────────────────────────────────────────────────────────────
console.log('--- xlsxjs render harness ---');
for (const w of warnings) console.log(`  · ${w}`);
if (failures.length) {
    console.error(`\n${failures.length} FAILURE(S):`);
    for (const f of failures) console.error(`  ✗ ${f}`);
    process.exit(1);
} else {
    console.log(`\n✓ all scenarios passed`);
}
