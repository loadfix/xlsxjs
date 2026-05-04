// Smoke test: load every .xlsx under a given directory (default
// /mnt/data/Temp/365) through the built library and report anomalies.
//
// Reports per file:
//   - parse status (ok / error with message)
//   - render status (ok / error with message)
//   - sheets, cell count, merges, rich-text cells, formulas, images, tables
//   - conditional formatting: counts by rule type, including "unsupported"
//   - theme colour refs that failed to resolve
//   - cells with unsupported pattern fills
//   - any XML part kept by the loader that the parser didn't touch (dropped)
//
// This is a triage tool, not a pass/fail test. Output is a summary table
// plus a per-file detail section; a second pass rolls up "things xlsxjs
// ignored" across the corpus so we can see which gaps are most common.

import { readFileSync, readdirSync, statSync } from 'node:fs';
import { resolve, dirname, join, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '..');
const sampleDir = process.argv[2] ?? '/mnt/data/Temp/365';

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

// Collect all xlsx-ish files in the sample dir.
function listXlsxFiles(dir) {
    try { statSync(dir); }
    catch { return []; }
    return readdirSync(dir)
        .filter((f) => /\.(xlsx|xlsm)$/i.test(f))
        .map((f) => join(dir, f));
}

const files = listXlsxFiles(sampleDir);
if (files.length === 0) {
    console.error(`No .xlsx / .xlsm files found in ${sampleDir}`);
    process.exit(1);
}

// Corpus-level tallies for the roll-up at the bottom.
const tallies = {
    unsupportedCfRuleTypes: new Map(),
    unknownIconSets: new Map(),
    unresolvedColors: 0,
    unknownPatternFills: new Map(),
    droppedPartPrefixes: new Map(),
    errors: [],
};

function bump(map, key) { map.set(key, (map.get(key) ?? 0) + 1); }

const reports = [];

for (const file of files) {
    const report = { file: basename(file), bytes: statSync(file).size };
    const buf = readFileSync(file);

    let wb, container;
    try {
        wb = await parseAsync(buf);
        report.parse = 'ok';
    } catch (err) {
        report.parse = `ERROR: ${err.message}`;
        tallies.errors.push({ file: report.file, phase: 'parse', err: err.message });
        reports.push(report);
        continue;
    }

    try {
        const nodes = await renderWorkbook(wb);
        container = document.createElement('div');
        for (const n of nodes) container.appendChild(n);
        report.render = 'ok';
    } catch (err) {
        report.render = `ERROR: ${err.message}`;
        tallies.errors.push({ file: report.file, phase: 'render', err: err.message });
        reports.push(report);
        continue;
    }

    // Walk the parsed workbook and collect stats.
    const parsed = wb.parsed;
    report.sheetCount = parsed.sheets.length;
    let cells = 0, richCells = 0, formulaCells = 0, merges = 0, images = 0, tables = 0;
    const cfTypeCounts = new Map();
    const unknownIconSetsHere = new Set();
    let numericCfHits = 0;
    for (const sheet of parsed.sheets) {
        for (const row of sheet.rows) {
            if (!row) continue;
            for (const cell of row) {
                cells++;
                if (cell.runs) richCells++;
                if (cell.formula != null) formulaCells++;
            }
        }
        merges += sheet.merges.length;
        images += sheet.images.length;
        tables += sheet.tables.length;
        for (const block of sheet.conditionalFormatting) {
            for (const rule of block.rules) {
                bump(cfTypeCounts, rule.type);
                if (rule.type === 'unsupported') {
                    bump(tallies.unsupportedCfRuleTypes, rule.type);
                }
                if (rule.type === 'iconSet' && rule.iconSet) {
                    const name = rule.iconSet.iconSet;
                    if (!['3TrafficLights1', '3Arrows', '3Symbols', '3Symbols2'].includes(name)) {
                        unknownIconSetsHere.add(name);
                        bump(tallies.unknownIconSets, name);
                    }
                }
            }
        }
    }
    report.cells = cells;
    report.richCells = richCells;
    report.formulaCells = formulaCells;
    report.merges = merges;
    report.images = images;
    report.tables = tables;
    report.cfTypes = Object.fromEntries(cfTypeCounts);
    report.unknownIconSets = [...unknownIconSetsHere];

    // Count patternFill types we didn't honour (only 'solid' lands today).
    // We look at the raw styles.xml by reading wb.parts directly to avoid
    // losing the pattern type info in the parsed model.
    //
    // `gray125` is Excel's default fills[1] on every workbook; we render
    // nothing for it (correct), and surfacing it in the report drowns out
    // real signal. Skip it.
    const stylesXml = wb.parts['xl/styles.xml'];
    if (stylesXml) {
        const fillMatches = [...stylesXml.matchAll(/patternType="([^"]+)"/g)];
        const patternCounts = new Map();
        for (const m of fillMatches) {
            const pt = m[1];
            if (pt === 'solid' || pt === 'none' || pt === 'gray125') continue;
            bump(patternCounts, pt);
            bump(tallies.unknownPatternFills, pt);
        }
        if (patternCounts.size) report.unknownPatternFills = Object.fromEntries(patternCounts);
    }

    // Colour-reference misses: parsed cells whose styles resolve to a null
    // colour when it was declared. We sample the styles for visibility.
    if (parsed.styles) {
        let unresolved = 0;
        for (const font of parsed.styles.fonts) {
            if (font.color && font.color.kind === 'theme') {
                const base = parsed.theme?.colors[font.color.index];
                if (!base) unresolved++;
            }
        }
        for (const fill of parsed.styles.fills) {
            if (fill.fgColor && fill.fgColor.kind === 'theme') {
                const base = parsed.theme?.colors[fill.fgColor.index];
                if (!base) unresolved++;
            }
        }
        if (unresolved) {
            report.unresolvedColors = unresolved;
            tallies.unresolvedColors += unresolved;
        }
    }

    // "Dropped parts" — scan the raw zip contents. Anything inside xl/ that
    // the loader didn't pull into wb.parts or wb.media is a whole feature
    // category we skipped (threaded comments, pivots, slicers, chartEx,
    // sparklines' extLst children, etc.). This is the list that tells us
    // what's worth building next.
    const zip = await JSZip.loadAsync(buf);
    const kept = new Set(Object.keys(wb.parts).concat(Object.keys(wb.media)));
    // Render-irrelevant parts. calcChain is a formula dependency graph;
    // printerSettings is binary printer config. Intentionally dropped.
    const RENDER_IRRELEVANT = /^xl\/(calcChain\.xml|printerSettings\/|customProperty|connections\.xml)/;
    for (const p of Object.keys(zip.files)) {
        if (zip.files[p].dir) continue;
        if (!p.startsWith('xl/')) continue;          // skip [Content_Types].xml etc.
        if (p.endsWith('.rels')) continue;           // we don't parse rels we don't need
        if (kept.has(p)) continue;
        if (RENDER_IRRELEVANT.test(p)) continue;
        // Categorise by first sub-directory under xl/ (or the filename when
        // the part sits in xl/ directly). e.g. "xl/pivotCache/..." → "pivotCache".
        const rest = p.slice('xl/'.length);
        const prefix = rest.includes('/') ? rest.split('/')[0] : rest;
        bump(tallies.droppedPartPrefixes, prefix);
        report.droppedParts ??= [];
        report.droppedParts.push(p);
    }

    reports.push(report);
}

// ── Per-file report ────────────────────────────────────────────────────────
console.log('smoke test — per file\n');
for (const r of reports) {
    console.log(`• ${r.file} (${r.bytes.toLocaleString()} B)`);
    console.log(`    parse: ${r.parse}    render: ${r.render ?? '—'}`);
    if (r.parse !== 'ok' || r.render !== 'ok') continue;
    console.log(`    sheets=${r.sheetCount}  cells=${r.cells}  rich=${r.richCells}  formulas=${r.formulaCells}  merges=${r.merges}  images=${r.images}  tables=${r.tables}`);
    if (Object.keys(r.cfTypes).length) console.log(`    cf rules: ${JSON.stringify(r.cfTypes)}`);
    if (r.unknownIconSets?.length) console.log(`    ⚠  unknown iconSet(s): ${r.unknownIconSets.join(', ')}`);
    if (r.unknownPatternFills) console.log(`    ⚠  non-solid fills: ${JSON.stringify(r.unknownPatternFills)}`);
    if (r.unresolvedColors) console.log(`    ⚠  unresolved theme colours: ${r.unresolvedColors}`);
    if (r.droppedParts) {
        // Only list the first few — it's easier to eyeball with the roll-up.
        const show = r.droppedParts.slice(0, 6).join(', ');
        const more = r.droppedParts.length > 6 ? ` (+${r.droppedParts.length - 6} more)` : '';
        console.log(`    dropped parts: ${show}${more}`);
    }
}

// ── Corpus-level roll-up ───────────────────────────────────────────────────
console.log('\nroll-up across corpus\n');
function printMap(label, map) {
    if (!map.size) return;
    console.log(`${label}:`);
    const entries = [...map.entries()].sort((a, b) => b[1] - a[1]);
    for (const [k, v] of entries) console.log(`  ${k.padEnd(30)} ${v}`);
}

printMap('unsupported cf rule types',  tallies.unsupportedCfRuleTypes);
printMap('unknown iconSet names',      tallies.unknownIconSets);
printMap('unsupported pattern fills',  tallies.unknownPatternFills);
printMap('dropped part directories',   tallies.droppedPartPrefixes);
if (tallies.unresolvedColors) {
    console.log(`unresolved theme colour refs: ${tallies.unresolvedColors} (summed across styles / files)`);
}
if (tallies.errors.length) {
    console.log('\nerrors:');
    for (const e of tallies.errors) console.log(`  ${e.phase} · ${e.file} · ${e.err}`);
}

const failures = tallies.errors.length;
if (failures) process.exit(1);
