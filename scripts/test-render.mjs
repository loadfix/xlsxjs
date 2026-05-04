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
