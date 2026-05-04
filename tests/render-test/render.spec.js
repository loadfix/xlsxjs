// @ts-check
import { test, expect } from '@playwright/test';

// Playwright in-browser smoke suite. The original Karma configuration
// (karma.conf.cjs) pointed at `tests/**/*spec.js`, but no Karma spec files
// ever existed in this repo — the scaffold was never populated. The
// authoritative render test suite lives in `scripts/test-render.mjs`, a
// 67-scenario JSDOM-based harness run via `npm run test:render`.
//
// This spec ports the smoke-level slice of that harness to a real browser:
// for every fixture under tests/render-test/, we load the .xlsx through the
// built UMD bundle and assert the renderer produced a plausible sheet. The
// per-scenario deep assertions in test-render.mjs remain in Node/JSDOM where
// they were written.

const fixtures = [
    'alignment-flags',
    'basic',
    'cellstyle-chain',
    'cf-custom-icons',
    'cf-ext-databar',
    'cf-graphical',
    'cf-icons',
    'cf-new-rules',
    'chart-detect',
    'comments',
    'conditional-format',
    'dimensions',
    'drawing-anchors-plus',
    'font-extras',
    'formulas',
    'hyperlinks-and-validation',
    'image',
    'merged',
    'multisheet',
    'named-styles',
    'numfmt-r2',
    'outlines-and-names',
    'page-layout',
    'phonetics',
    'python-xlsx',
    'richtext',
    'shapes-and-textboxes',
    'sheet-protection-and-alt',
    'sheet-view-state',
    'tables',
    'theme-fontscheme',
    'threaded-comments',
];

test.describe('Render workbook', () => {
    for (const fixture of fixtures) {
        test(`renders ${fixture} fixture`, async ({ page }) => {
            await page.goto('/tests/harness.html');

            const result = await page.evaluate(async (name) => {
                const res = await fetch(`/tests/render-test/${name}/workbook.xlsx`);
                if (!res.ok) throw new Error(`fetch ${name}: ${res.status}`);
                const buf = await res.arrayBuffer();

                // xlsx-preview exposes parseAsync + renderWorkbook as UMD globals.
                // @ts-ignore — global injected by dist/xlsx-preview.js
                const wb = await xlsx.parseAsync(buf);
                // @ts-ignore
                const nodes = await xlsx.renderWorkbook(wb);

                const container = document.createElement('div');
                for (const n of nodes) container.appendChild(n);
                document.body.appendChild(container);

                const sheets = wb.parsed?.sheets ?? [];
                const sections = container.querySelectorAll('section.xlsx');
                // Some fixtures include hidden sheets that the renderer drops;
                // assert at least one visible section exists.
                return {
                    sheetCount: sheets.length,
                    sectionCount: sections.length,
                    firstSectionHasTable: sections.length > 0 && !!sections[0].querySelector('table'),
                };
            }, fixture);

            expect(result.sheetCount).toBeGreaterThan(0);
            // sheet-view-state deliberately hides all but one sheet; others
            // always render at least one section.
            expect(result.sectionCount).toBeGreaterThan(0);
            expect(result.firstSectionHasTable).toBe(true);
        });
    }
});

test.describe('Library surface', () => {
    test('exposes parseAsync, renderWorkbook, renderAsync globals', async ({ page }) => {
        await page.goto('/tests/harness.html');
        const api = await page.evaluate(() => ({
            // @ts-ignore
            parseAsync: typeof xlsx.parseAsync,
            // @ts-ignore
            renderWorkbook: typeof xlsx.renderWorkbook,
            // @ts-ignore
            renderAsync: typeof xlsx.renderAsync,
        }));
        expect(api.parseAsync).toBe('function');
        expect(api.renderWorkbook).toBe('function');
        expect(api.renderAsync).toBe('function');
    });

    test('renderAsync mounts into a container', async ({ page }) => {
        await page.goto('/tests/harness.html');
        const result = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/basic/workbook.xlsx').then((r) => r.arrayBuffer());
            const host = document.createElement('div');
            document.body.appendChild(host);
            // @ts-ignore
            await xlsx.renderAsync(buf, host, null, {});
            return {
                hasSection: !!host.querySelector('section.xlsx'),
                hasTable: !!host.querySelector('section.xlsx table'),
            };
        });
        expect(result.hasSection).toBe(true);
        expect(result.hasTable).toBe(true);
    });
});
