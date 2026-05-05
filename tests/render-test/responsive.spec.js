// @ts-check
import { test, expect } from '@playwright/test';

// Exercises the `responsive: true` render option added in Wave 9 (W9-E).
//
// Default mode is byte-stable (no data-responsive attribute, no sticky /
// overflow CSS); responsive mode marks every sheet section with
// data-responsive="true" and emits the sticky + overflow-x CSS block.

test.describe('responsive mode', () => {
    test('default (responsive: false): section has no data-responsive hook', async ({ page }) => {
        await page.goto('/tests/harness.html');
        const result = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/basic/workbook.xlsx').then((r) => r.arrayBuffer());
            // @ts-ignore — UMD global
            const wb = await xlsx.parseAsync(buf);
            // @ts-ignore
            const nodes = await xlsx.renderWorkbook(wb);
            const host = document.createElement('div');
            for (const n of nodes) host.appendChild(n);
            document.body.appendChild(host);

            const section = /** @type {HTMLElement | null} */ (host.querySelector('section.xlsx'));
            const cssText = [...host.querySelectorAll('style')].map((s) => s.textContent ?? '').join('\n');
            const out = {
                hasResponsiveAttr: section?.getAttribute('data-responsive'),
                cssHasResponsiveRule: cssText.includes('data-responsive="true"'),
            };
            host.remove();
            return out;
        });
        expect(result.hasResponsiveAttr).toBe(null);
        expect(result.cssHasResponsiveRule).toBe(false);
    });

    test('responsive: true — section is hook-marked + CSS contains sticky + overflow rules', async ({ page }) => {
        await page.goto('/tests/harness.html');
        const result = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/basic/workbook.xlsx').then((r) => r.arrayBuffer());
            // @ts-ignore — UMD global
            const wb = await xlsx.parseAsync(buf);
            // @ts-ignore
            const nodes = await xlsx.renderWorkbook(wb, { responsive: true });
            const host = document.createElement('div');
            for (const n of nodes) host.appendChild(n);
            document.body.appendChild(host);

            const section = /** @type {HTMLElement | null} */ (host.querySelector('section.xlsx'));
            const cssText = [...host.querySelectorAll('style')].map((s) => s.textContent ?? '').join('\n');
            const out = {
                hasResponsiveAttr: section?.getAttribute('data-responsive'),
                cssHasResponsiveRule: cssText.includes('data-responsive="true"'),
                cssHasOverflowX: cssText.includes('overflow-x: auto'),
                cssHasStickyLeft: cssText.includes('position: sticky; left: 0'),
                cssHasStickyTop: cssText.includes('position: sticky; top: 0'),
                cssHasTouchScrolling: cssText.includes('-webkit-overflow-scrolling'),
                cssHasMediaQuery: cssText.includes('@media (max-width: 768px)'),
            };
            host.remove();
            return out;
        });
        expect(result.hasResponsiveAttr).toBe('true');
        expect(result.cssHasResponsiveRule).toBe(true);
        expect(result.cssHasOverflowX).toBe(true);
        expect(result.cssHasStickyLeft).toBe(true);
        expect(result.cssHasStickyTop).toBe(true);
        expect(result.cssHasTouchScrolling).toBe(true);
        expect(result.cssHasMediaQuery).toBe(true);
    });
});
