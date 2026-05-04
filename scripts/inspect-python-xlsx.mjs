// One-shot inspector: load the python-xlsx fixture through the built library
// and dump the parsed model + rendered HTML skeleton, so we can see how much
// of the fixture the current renderer already captures.

import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');

const dom = new JSDOM('<!doctype html><html><body></body></html>', { url: 'http://localhost/' });
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

const buf = readFileSync(resolve(repo, 'tests/render-test/python-xlsx/workbook.xlsx'));
const wb = await parseAsync(buf);
const sheet = wb.parsed.sheets[0];

console.log(`sheet: "${sheet.name}"  maxRow=${sheet.maxRow}  maxCol=${sheet.maxCol}`);
console.log(`merges: ${JSON.stringify(sheet.merges)}`);
console.log(`columns: ${JSON.stringify(sheet.columns)}`);
console.log('cells (row,col -> kind:value):');
for (let r = 0; r <= sheet.maxRow; r++) {
    const row = sheet.rows[r];
    if (!row) continue;
    for (const cell of row) {
        console.log(`  ${r},${cell.col}  ${cell.kind.padEnd(9)} ${JSON.stringify(cell.value)}`);
    }
}

const nodes = await renderWorkbook(wb);
const container = document.createElement('div');
for (const n of nodes) container.appendChild(n);
const table = container.querySelector('section.xlsx table');
console.log('\nrendered table shape:');
console.log(`  thead rows: ${table.querySelectorAll('thead tr').length}`);
console.log(`  tbody rows: ${table.querySelectorAll('tbody tr').length}`);
console.log(`  merged tds (.xlsx-merged): ${table.querySelectorAll('td.xlsx-merged').length}`);
console.log(`  numeric tds: ${table.querySelectorAll('td.xlsx-numeric').length}`);

console.log('\nrendered cells (row → [text | bg | bold]):');
const bodyRows = table.querySelectorAll('tbody tr');
for (let i = 0; i < bodyRows.length; i++) {
    const row = bodyRows[i];
    const tds = row.querySelectorAll('td');
    const cells = [...tds].map((td) => {
        const bg = td.style.backgroundColor || '-';
        const fw = td.style.fontWeight || '-';
        const txt = td.textContent || '';
        return `${txt.padEnd(18)}(bg=${bg.padEnd(8)} fw=${fw})`;
    });
    console.log(`  r${i}: ${cells.join(' | ')}`);
}
