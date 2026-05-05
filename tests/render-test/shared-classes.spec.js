// @ts-check
//
// W9-F shared-class assertions — xlsxjs side.
//
// The renderer now emits a cross-format `oox-*` class alongside each
// format-specific class so manifests can write a single selector that
// works for docxjs/pptxjs/xlsxjs output. This spec verifies the classes
// are present on the expected elements for the basic workbook fixture.
// See src/shared-classes.ts for the concept → class table.
import { test, expect } from '@playwright/test';

test.describe('Shared oox-* classes', () => {
    test('basic workbook carries oox-page / oox-table / oox-table-row / oox-table-cell', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const counts = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/basic/workbook.xlsx').then(r => r.arrayBuffer());
            const host = document.createElement('div');
            document.body.appendChild(host);
            // @ts-ignore — UMD global
            await xlsx.renderAsync(buf, host, null, {});
            const q = (sel) => host.querySelectorAll(sel).length;
            const result = {
                page:      q('.oox-page'),
                table:     q('.oox-table'),
                tableRow:  q('.oox-table-row'),
                tableCell: q('.oox-table-cell'),
                // Legacy format-specific selectors still present.
                xlsxSection: q('section.xlsx'),
            };
            host.remove();
            return result;
        });

        expect(counts.page).toBeGreaterThanOrEqual(1);
        expect(counts.table).toBeGreaterThanOrEqual(1);
        expect(counts.tableRow).toBeGreaterThanOrEqual(1);
        expect(counts.tableCell).toBeGreaterThanOrEqual(1);
        // Legacy class still emitted.
        expect(counts.xlsxSection).toBeGreaterThanOrEqual(1);
        // Each <section class="xlsx"> now also carries .oox-page.
        expect(counts.page).toBe(counts.xlsxSection);
    });

    test('image fixture carries oox-image on each figure', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const counts = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/image/workbook.xlsx').then(r => r.arrayBuffer());
            const host = document.createElement('div');
            document.body.appendChild(host);
            // @ts-ignore — UMD global
            await xlsx.renderAsync(buf, host, null, {});
            const result = {
                sharedImages: host.querySelectorAll('.oox-image').length,
                xlsxImages: host.querySelectorAll('figure.xlsx-image').length,
            };
            host.remove();
            return result;
        });

        expect(counts.xlsxImages).toBeGreaterThanOrEqual(1);
        // Every <figure class="xlsx-image"> also carries .oox-image.
        expect(counts.sharedImages).toBe(counts.xlsxImages);
    });
});
