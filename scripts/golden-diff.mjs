// Golden-HTML-diff harness for the XLSX render pipeline.
//
// Two modes:
//   --capture         render every fixture under tests/render-test/*/workbook.xlsx,
//                     serialise the rendered container to HTML, and write it to
//                     tests/render-test/<name>/result.html. Overwrites existing
//                     snapshots. Used after an intentional renderer change.
//   --diff  (default) render every fixture and compare against the stored
//                     result.html, printing a unified-style diff on drift. Exits
//                     1 when any fixture diverges.
//
// Determinism:
//   • Rendered DOM is serialised, then broken across tag boundaries (`></`)
//     onto separate lines so diffs are readable and whitespace-stable.
//   • Data URLs on <img src=…> are kept as-is — they encode PNG bytes that
//     come straight from the zip, so same fixture → same bytes.
//   • ISO dates (`YYYY-MM-DD`) inside .xlsx-header / .xlsx-footer text come
//     from `&D` substitutions that call `new Date()` at render time; they are
//     replaced with `YYYY-MM-DD` before comparing so the snapshot is stable
//     across days.
//
// If future rendering work introduces other time- or machine-dependent
// attributes, extend `normaliseHtml` below (and keep this comment current).
//
// The diff is a tiny LCS-free line diff — adequate for spotting regressions
// without pulling in the `diff` npm package (not in package.json).

import { readFileSync, writeFileSync, readdirSync, statSync, existsSync } from 'node:fs';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const fixturesRoot = resolve(repo, 'tests/render-test');

const argv = new Set(process.argv.slice(2));
const CAPTURE = argv.has('--capture');
const DIFF = !CAPTURE; // default

// ── jsdom + UMD setup (mirrors scripts/test-render.mjs) ────────────────────
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

// ── Fixture discovery ─────────────────────────────────────────────────────
function listFixtures(root) {
    if (!existsSync(root)) return [];
    return readdirSync(root)
        .filter((name) => {
            const p = join(root, name);
            if (!statSync(p).isDirectory()) return false;
            return existsSync(join(p, 'workbook.xlsx'));
        })
        .sort();
}

// ── HTML normalisation ────────────────────────────────────────────────────
// Split between adjacent tags so one tag per line → readable diffs.
// We also trim inter-tag whitespace so differing serialiser whitespace can't
// cause false diffs.
function splitTagsToLines(html) {
    // Collapse runs of whitespace between `>` and `<` to nothing, then insert a
    // newline between every tag boundary.
    const compact = html.replace(/>\s+</g, '><');
    return compact.replace(/></g, '>\n<');
}

// Replace any ISO-like date (`YYYY-MM-DD`) that sits inside .xlsx-header /
// .xlsx-footer text nodes. Those come from `&D` expansions in
// workbook-parser.ts which invoke `new Date()` at render time — they're
// not stable across days. We only rewrite the date substring itself; the
// surrounding zone text stays intact so any real regression is still caught.
//
// We operate on the DOM (not the string) because the serialised HTML for
// a header contains nested <div data-zone> children and a non-greedy
// regex on the outer div stops at the first inner </div>.
function normaliseHeaderFooterDates(container) {
    const zones = container.querySelectorAll('.xlsx-header, .xlsx-footer');
    const ISO = /\d{4}-\d{2}-\d{2}/g;
    for (const el of zones) {
        const walker = document.createTreeWalker(el, dom.window.NodeFilter.SHOW_TEXT);
        let node;
        while ((node = walker.nextNode())) {
            if (ISO.test(node.nodeValue)) {
                node.nodeValue = node.nodeValue.replace(ISO, 'YYYY-MM-DD');
            }
        }
    }
}

// Normalisation applied to both captured + observed output before comparing
// or writing. Keep this list narrow + documented — every rule here hides a
// real bit of non-determinism, so future contributors should think twice
// before adding another.
//
// Today we only strip two things:
//   1. Dates inside xlsx-header/xlsx-footer (new Date() at render time).
//      Applied on the DOM before serialisation, so we correctly scope to
//      header/footer text nodes without trying to balance HTML tags.
//   2. Whitespace between adjacent tags (purely serialiser variance).
// No attribute scrubbing is needed — the renderer doesn't inject timestamps
// or random ids into attributes. If that ever changes, filter attributes here.
function normaliseHtml(container) {
    normaliseHeaderFooterDates(container);
    const broken = splitTagsToLines(container.innerHTML);
    return broken.trim() + '\n';
}

async function renderFixture(name) {
    const buf = readFileSync(join(fixturesRoot, name, 'workbook.xlsx'));
    const wb = await parseAsync(buf);
    const nodes = await renderWorkbook(wb);
    const container = document.createElement('div');
    for (const n of nodes) container.appendChild(n);
    return normaliseHtml(container);
}

// ── Minimal unified-style diff (5 lines of context) ───────────────────────
// Not an LCS — a line-by-line walk over the longer of the two. Good enough
// for snapshot drift where most lines match up positionally. Output mimics
// `diff -u` format so developers recognise it.
function unifiedDiff(expectedText, actualText, label, context = 5) {
    const exp = expectedText.split('\n');
    const act = actualText.split('\n');
    const maxLen = Math.max(exp.length, act.length);

    // Find changed line indices.
    const changed = [];
    for (let i = 0; i < maxLen; i++) {
        if (exp[i] !== act[i]) changed.push(i);
    }
    if (changed.length === 0) return '';

    // Group changed lines into hunks, expanded by `context` lines of surround.
    const hunks = [];
    let start = Math.max(0, changed[0] - context);
    let end = changed[0];
    for (let k = 1; k < changed.length; k++) {
        const i = changed[k];
        if (i - end > context * 2) {
            hunks.push([start, Math.min(maxLen - 1, end + context)]);
            start = Math.max(0, i - context);
        }
        end = i;
    }
    hunks.push([start, Math.min(maxLen - 1, end + context)]);

    const lines = [`--- ${label} (expected)`, `+++ ${label} (actual)`];
    for (const [from, to] of hunks) {
        lines.push(`@@ lines ${from + 1}-${to + 1} @@`);
        for (let i = from; i <= to; i++) {
            const e = exp[i];
            const a = act[i];
            if (e === a) {
                if (e !== undefined) lines.push(` ${e}`);
            } else {
                if (e !== undefined) lines.push(`-${e}`);
                if (a !== undefined) lines.push(`+${a}`);
            }
        }
    }
    return lines.join('\n');
}

// ── Main ──────────────────────────────────────────────────────────────────
const fixtures = listFixtures(fixturesRoot);
if (fixtures.length === 0) {
    console.error(`No fixtures found under ${fixturesRoot}`);
    process.exit(1);
}

let captured = 0;
let passed = 0;
let failed = 0;
let missing = 0;
const failures = [];

for (const name of fixtures) {
    const snapshotPath = join(fixturesRoot, name, 'result.html');
    let rendered;
    try {
        rendered = await renderFixture(name);
    } catch (err) {
        console.error(`✗ ${name}: render threw — ${err.message}`);
        failures.push(name);
        failed++;
        continue;
    }

    if (CAPTURE) {
        writeFileSync(snapshotPath, rendered);
        console.log(`  wrote ${name}/result.html  (${rendered.length} chars)`);
        captured++;
        continue;
    }

    // --diff mode
    if (!existsSync(snapshotPath)) {
        console.error(`! ${name}: no snapshot at ${snapshotPath} — run with --capture first`);
        missing++;
        failed++;
        failures.push(name);
        continue;
    }

    const expected = readFileSync(snapshotPath, 'utf8');
    if (expected === rendered) {
        console.log(`✓ ${name}`);
        passed++;
    } else {
        console.log(`✗ ${name}`);
        const d = unifiedDiff(expected, rendered, `tests/render-test/${name}/result.html`, 5);
        if (d) console.log(d);
        failed++;
        failures.push(name);
    }
}

if (CAPTURE) {
    console.log(`\ncaptured ${captured} snapshot(s) under ${fixturesRoot}`);
    process.exit(0);
}

console.log(`\n${passed} passed, ${failed} failed${missing ? ` (${missing} missing)` : ''}  of ${fixtures.length}`);
if (failed) {
    console.log('failures:');
    for (const f of failures) console.log(`  ${f}`);
    process.exit(1);
}
process.exit(0);
