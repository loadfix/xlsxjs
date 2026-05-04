// Render-assertion conformance spec for xlsxjs.
//
// For every manifest under ../ooxml-reference-corpus/features/xlsx/*.json
// that declares a `render_assertions` block:
//   1. Load the committed machine fixture (../ooxml-reference-corpus/
//      fixtures/xlsx/<name>.xlsx).
//   2. Render it through the built UMD bundle inside a real browser.
//   3. Evaluate each render_assertion (css_selector / computed_style /
//      visual_ssim) and collect per-assertion verdicts.
//   4. Write a FeatureResult-shaped JSON to
//      ../ooxml-validate/conformance/results/xlsxjs/xlsx/<name>.json.
//
// The Playwright test body asserts only that aggregate status != "error";
// a `fail` is a recorded-but-non-blocking signal (the matrix page surfaces
// pass/fail/error columns). This matches the sibling docxjs/pptxjs model
// and avoids turning every manifest-author regression into a blocking CI
// failure in the rendering library.

import { test, expect } from '@playwright/test';
import { readFileSync, readdirSync, mkdirSync, writeFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import {
    browserEvaluateCssSelector,
    browserEvaluateComputedStyle,
    evaluateVisualSsim,
    aggregateStatus,
    isoNowZ,
    type RenderAssertion,
    type AssertionVerdict,
    type FeatureResult,
} from './evaluator';

// Playwright transpiles this spec to CJS, so __dirname is available and we
// avoid `import.meta.url` (which would flip the module into ESM-only
// interpretation and break the CJS-shimmed pngjs/pixelmatch imports in
// evaluator.ts).
const here = __dirname;
const repo = resolve(here, '../..');
const corpus = resolve(repo, '../ooxml-reference-corpus');
const validate = resolve(repo, '../ooxml-validate');
const resultsDir = resolve(validate, 'conformance/results/xlsxjs/xlsx');

const LIBRARY = 'xlsxjs';
const TOOL_VERSION = readToolVersion();

function readToolVersion(): string {
    try {
        const pkg = JSON.parse(readFileSync(resolve(repo, 'package.json'), 'utf8'));
        return String(pkg.version ?? '');
    } catch {
        return '';
    }
}

interface Manifest {
    id: string;
    format: string;
    fixtures: { machine: string; office?: string };
    render_assertions?: RenderAssertion[];
}

function loadManifests(): { manifest: Manifest; path: string }[] {
    const dir = resolve(corpus, 'features/xlsx');
    if (!existsSync(dir)) return [];
    const entries = readdirSync(dir).filter((f) => f.endsWith('.json'));
    const out: { manifest: Manifest; path: string }[] = [];
    for (const f of entries) {
        const p = resolve(dir, f);
        const raw = readFileSync(p, 'utf8');
        const m = JSON.parse(raw) as Manifest;
        if (!m.render_assertions || m.render_assertions.length === 0) continue;
        if (m.format !== 'xlsx') continue;
        out.push({ manifest: m, path: p });
    }
    return out;
}

function featureBaseName(id: string): string {
    // "xlsx/cell-bold" → "cell-bold"
    const slash = id.lastIndexOf('/');
    return slash >= 0 ? id.slice(slash + 1) : id;
}

function fixturePathFor(manifest: Manifest): string {
    // fixtures.machine is a logical name like "xlsx/cell-bold"; the corpus
    // stores the file as fixtures/xlsx/cell-bold.xlsx.
    const name = featureBaseName(manifest.fixtures.machine);
    return resolve(corpus, 'fixtures/xlsx', `${name}.xlsx`);
}

function referencePngPathFor(assertion: RenderAssertion, manifest: Manifest): string {
    if (assertion.reference_png) {
        return resolve(corpus, assertion.reference_png);
    }
    const name = featureBaseName(manifest.fixtures.machine);
    return resolve(corpus, 'refs/xlsx', `${name}-page1.png`);
}

function writeResult(result: FeatureResult): string {
    mkdirSync(resultsDir, { recursive: true });
    const name = featureBaseName(result.feature_id);
    const out = resolve(resultsDir, `${name}.json`);
    writeFileSync(out, JSON.stringify(result, null, 2) + '\n', 'utf8');
    return out;
}

// --- Test registration --------------------------------------------------

const manifests = loadManifests();

if (manifests.length === 0) {
    test('xlsx corpus not available — skipping conformance run', () => {
        test.skip(true, `No manifests found under ${resolve(corpus, 'features/xlsx')}.`);
    });
} else {
    test.describe('xlsxjs render-assertion conformance', () => {
        for (const { manifest } of manifests) {
            const featureId = manifest.id;
            const name = featureBaseName(featureId);

            test(`${featureId}`, async ({ page }) => {
                const fixtureAbs = fixturePathFor(manifest);
                if (!existsSync(fixtureAbs)) {
                    const result: FeatureResult = {
                        feature_id: featureId,
                        library: LIBRARY,
                        status: 'error',
                        fixture_path: fixtureAbs,
                        run_at: isoNowZ(),
                        tool_version: TOOL_VERSION,
                        assertions: [
                            {
                                id: '<open>',
                                status: 'error',
                                detail: `Fixture not found: ${fixtureAbs}`,
                            },
                        ],
                    };
                    writeResult(result);
                    expect(result.status, `fixture missing: ${fixtureAbs}`).not.toBe('error');
                    return;
                }

                // 1. Navigate to harness and render the workbook.
                await page.goto('/tests/harness.html');

                const fixtureBytes = readFileSync(fixtureAbs);
                // Playwright serializes Uint8Array over the CDP bridge, so pass
                // the buffer contents as a plain number[] and reconstruct in
                // the page. Number[] is large but every xlsx fixture here is
                // small (cell-bold is ~5 KB).
                const fixtureArray = Array.from(new Uint8Array(fixtureBytes));

                const renderError = await page.evaluate(async (bytes) => {
                    const u8 = new Uint8Array(bytes);
                    try {
                        // @ts-ignore — UMD globals injected by dist/xlsx-preview.js
                        const wb = await xlsx.parseAsync(u8.buffer);
                        // @ts-ignore
                        const nodes = await xlsx.renderWorkbook(wb);
                        const container = document.createElement('div');
                        for (const n of nodes) container.appendChild(n);
                        document.body.appendChild(container);
                        return null;
                    } catch (e: unknown) {
                        return (e as Error).message || String(e);
                    }
                }, fixtureArray);

                const verdicts: AssertionVerdict[] = [];

                if (renderError) {
                    verdicts.push({
                        id: '<render>',
                        status: 'error',
                        detail: `xlsxjs render threw: ${renderError}`,
                    });
                } else {
                    // 2. Evaluate each render assertion.
                    for (const assertion of manifest.render_assertions ?? []) {
                        if (assertion.kind === 'css_selector') {
                            const v = await page.evaluate(browserEvaluateCssSelector, assertion);
                            verdicts.push(v);
                        } else if (assertion.kind === 'computed_style') {
                            const v = await page.evaluate(browserEvaluateComputedStyle, assertion);
                            verdicts.push(v);
                        } else if (assertion.kind === 'visual_ssim') {
                            const refPath = referencePngPathFor(assertion, manifest);
                            let screenshot: Buffer;
                            try {
                                // Screenshot only the rendered workbook region. If
                                // the first section.xlsx is missing, fall back to
                                // the full page.
                                const locator = page.locator('section.xlsx').first();
                                if ((await locator.count()) > 0) {
                                    screenshot = await locator.screenshot();
                                } else {
                                    screenshot = await page.screenshot({ fullPage: true });
                                }
                            } catch (e) {
                                verdicts.push({
                                    id: assertion.id,
                                    status: 'error',
                                    detail: `Screenshot failed: ${(e as Error).message}`,
                                });
                                continue;
                            }
                            verdicts.push(evaluateVisualSsim(assertion, screenshot, refPath));
                        } else {
                            verdicts.push({
                                id: assertion.id,
                                status: 'error',
                                detail: `Unknown assertion kind: ${(assertion as RenderAssertion).kind}`,
                            });
                        }
                    }
                }

                // 3. Build + write result JSON.
                const result: FeatureResult = {
                    feature_id: featureId,
                    library: LIBRARY,
                    status: aggregateStatus(verdicts),
                    fixture_path: fixtureAbs,
                    run_at: isoNowZ(),
                    tool_version: TOOL_VERSION,
                    assertions: verdicts,
                };
                const writtenTo = writeResult(result);
                console.log(`[xlsxjs conformance] ${featureId} → ${result.status} (${writtenTo})`);

                // 4. Record-but-don't-block: fail the Playwright test only on
                //    "error" status. A fail-status run still produces a result
                //    JSON the matrix can consume.
                expect(
                    result.status,
                    `assertions=${JSON.stringify(verdicts, null, 2)}`,
                ).not.toBe('error');
            });
        }
    });
}
