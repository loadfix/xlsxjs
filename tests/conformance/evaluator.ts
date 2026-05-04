// Render-assertion evaluator for xlsxjs conformance.
//
// Consumes the feature manifest entries described in
// ooxml-reference-corpus/features/manifest.schema.json (render_assertion $def)
// and returns per-assertion verdicts matching the shape produced by
// ooxml_validate.conformance.FeatureResult.to_dict():
//
//   { id, status: "pass" | "fail" | "error", detail }
//
// Three assertion kinds are supported:
//
//   css_selector  – DOM match-count / match-text check via querySelectorAll.
//   computed_style – getComputedStyle(...).getPropertyValue(style_property).
//   visual_ssim   – pixelmatch-backed similarity approximation against a
//                   reference PNG. Runs Node-side from an already-taken
//                   Playwright screenshot.
//
// css_selector / computed_style checks are evaluated inside the browser
// (page.evaluate) because they need a live DOM and live computed styles.
// visual_ssim runs Node-side against saved buffers.

import { readFileSync, existsSync } from 'node:fs';

// pngjs + pixelmatch are CommonJS packages with no ESM build. Playwright's
// TypeScript loader transpiles our spec to CJS, so the idiomatic `import
// pixelmatch from 'pixelmatch'` resolves through the standard CJS path.
// Using `require` directly (legally available because Playwright's loader
// injects a CommonJS scope) avoids any `import.meta` machinery that would
// otherwise flip the module into an ESM-only interpretation.
type PNGType = { sync: { read: (buf: Buffer) => { width: number; height: number; data: Buffer } } };
type PixelmatchFn = (
    a: Buffer | Uint8Array,
    b: Buffer | Uint8Array,
    out: Buffer | Uint8Array | null,
    w: number,
    h: number,
    opts?: { threshold?: number },
) => number;
// eslint-disable-next-line @typescript-eslint/no-require-imports
const { PNG } = require('pngjs') as { PNG: PNGType };
// eslint-disable-next-line @typescript-eslint/no-require-imports
const pixelmatch: PixelmatchFn = require('pixelmatch');

export interface RenderAssertion {
    id: string;
    kind: 'css_selector' | 'computed_style' | 'visual_ssim';
    selector?: string;
    must?: 'exist' | 'absent' | 'equal-count' | 'match-text';
    count?: number;
    style_property?: string;
    value?: string;
    min_ssim?: number;
    reference_png?: string;
    description?: string;
}

export interface AssertionVerdict {
    id: string;
    status: 'pass' | 'fail' | 'error';
    detail: string;
}

// --- Browser-side evaluators --------------------------------------------
//
// These functions are serialized via fn.toString() by page.evaluate() and
// executed inside the browser. They must be self-contained: no imports,
// no closure captures, no references to symbols outside the body.

/** Runs in the browser. Evaluates a css_selector assertion. */
export function browserEvaluateCssSelector(a: RenderAssertion): AssertionVerdict {
    const sel = a.selector || '';
    if (!sel) {
        return { id: a.id, status: 'error', detail: "Missing 'selector'." };
    }
    const nodes = document.querySelectorAll(sel);
    const count = nodes.length;
    const must = a.must || 'exist';

    if (must === 'exist') {
        return count >= 1
            ? { id: a.id, status: 'pass', detail: `${count} node(s) matched.` }
            : { id: a.id, status: 'fail', detail: `No nodes matched "${sel}".` };
    }
    if (must === 'absent') {
        return count === 0
            ? { id: a.id, status: 'pass', detail: 'No nodes matched (as required).' }
            : {
                  id: a.id,
                  status: 'fail',
                  detail: `${count} node(s) matched but none expected.`,
              };
    }
    if (must === 'equal-count') {
        const expected = a.count ?? -1;
        return count === expected
            ? { id: a.id, status: 'pass', detail: `Matched ${count} node(s).` }
            : {
                  id: a.id,
                  status: 'fail',
                  detail: `Expected ${expected} nodes, got ${count}.`,
              };
    }
    if (must === 'match-text') {
        if (count === 0) {
            return { id: a.id, status: 'fail', detail: 'No node to match text against.' };
        }
        const value = a.value;
        if (value === undefined) {
            return { id: a.id, status: 'error', detail: "'match-text' requires 'value'." };
        }
        // Each match is a candidate. Pass if ANY candidate matches — xlsxjs
        // emits many <td> cells and the fixture text sits in exactly one.
        const re = new RegExp(value);
        for (const node of Array.from(nodes)) {
            const txt = (node.textContent ?? '').trim();
            if (re.test(txt)) {
                return {
                    id: a.id,
                    status: 'pass',
                    detail: `text=${JSON.stringify(txt)} matches /${value}/.`,
                };
            }
        }
        const first = (nodes[0].textContent ?? '').trim();
        return {
            id: a.id,
            status: 'fail',
            detail: `No match for /${value}/ among ${count} node(s); first=${JSON.stringify(first)}.`,
        };
    }
    return { id: a.id, status: 'error', detail: `Unknown 'must' mode: ${must}.` };
}

/** Runs in the browser. Evaluates a computed_style assertion. */
export function browserEvaluateComputedStyle(a: RenderAssertion): AssertionVerdict {
    const sel = a.selector || '';
    const prop = a.style_property || '';
    const value = a.value;
    if (!sel) return { id: a.id, status: 'error', detail: "Missing 'selector'." };
    if (!prop) return { id: a.id, status: 'error', detail: "Missing 'style_property'." };
    if (value === undefined) {
        return { id: a.id, status: 'error', detail: "Missing 'value'." };
    }
    const nodes = document.querySelectorAll(sel);
    if (nodes.length === 0) {
        return { id: a.id, status: 'fail', detail: `No element matched ${sel}.` };
    }
    const re = new RegExp(value);
    // If the selector matches multiple elements (as it does for cell-bold —
    // every td/th is matched), pass if ANY element's computed style satisfies
    // the pattern. This matches the manifest author's intent: "at least one
    // cell has the expected style".
    let lastSeen = '';
    for (const node of Array.from(nodes)) {
        const style = getComputedStyle(node as Element).getPropertyValue(prop).trim();
        lastSeen = style;
        if (re.test(style)) {
            return {
                id: a.id,
                status: 'pass',
                detail: `${prop}=${JSON.stringify(style)} matches /${value}/.`,
            };
        }
    }
    return {
        id: a.id,
        status: 'fail',
        detail: `None of ${nodes.length} element(s) had ${prop} matching /${value}/; last seen ${JSON.stringify(lastSeen)}.`,
    };
}

// --- Node-side evaluators -----------------------------------------------

/**
 * Compare a screenshot buffer against the reference PNG and verdict against
 * `min_ssim`. Uses pixelmatch, which computes a per-pixel anti-aliasing-
 * tolerant diff; we convert its "fraction of differing pixels" into a
 * coarse similarity score (1 - differing_fraction). This is not Wang's
 * structural similarity, but it's the standard approximation the sibling
 * render suites use and the manifest thresholds were calibrated against.
 */
export function evaluateVisualSsim(
    assertion: RenderAssertion,
    screenshotBuf: Buffer,
    referencePngPath: string,
): AssertionVerdict {
    const threshold = assertion.min_ssim ?? 0.95;
    if (!existsSync(referencePngPath)) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Reference PNG not found: ${referencePngPath}`,
        };
    }
    let refPng: ReturnType<PNGType['sync']['read']>;
    let shotPng: ReturnType<PNGType['sync']['read']>;
    try {
        refPng = PNG.sync.read(readFileSync(referencePngPath));
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Could not read reference PNG: ${(e as Error).message}`,
        };
    }
    try {
        shotPng = PNG.sync.read(screenshotBuf);
    } catch (e) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Could not decode screenshot: ${(e as Error).message}`,
        };
    }
    // pixelmatch requires matching dimensions. If the sizes diverge clip
    // both to the overlap — a partial similarity is better than an error,
    // and full-page alignment is sensitive to DPR anyway.
    const width = Math.min(refPng.width, shotPng.width);
    const height = Math.min(refPng.height, shotPng.height);
    if (width === 0 || height === 0) {
        return {
            id: assertion.id,
            status: 'error',
            detail: `Zero-dimension overlap between screenshot (${shotPng.width}x${shotPng.height}) and reference (${refPng.width}x${refPng.height}).`,
        };
    }
    const clip = (src: { width: number; height: number; data: Buffer }): Buffer => {
        const out = Buffer.alloc(width * height * 4);
        for (let y = 0; y < height; y++) {
            const srcOff = y * src.width * 4;
            const dstOff = y * width * 4;
            src.data.copy(out, dstOff, srcOff, srcOff + width * 4);
        }
        return out;
    };
    const a = clip(refPng);
    const b = clip(shotPng);
    const diff = pixelmatch(a, b, null, width, height, { threshold: 0.1 });
    const total = width * height;
    const ssimApprox = 1 - diff / total;
    if (ssimApprox >= threshold) {
        return {
            id: assertion.id,
            status: 'pass',
            detail: `similarity=${ssimApprox.toFixed(4)} >= ${threshold} (diff ${diff}/${total} px).`,
        };
    }
    return {
        id: assertion.id,
        status: 'fail',
        detail: `similarity=${ssimApprox.toFixed(4)} < ${threshold} (diff ${diff}/${total} px).`,
    };
}

// --- Result writer ------------------------------------------------------
//
// Mirrors ooxml_validate.conformance.FeatureResult.to_dict() exactly so the
// matrix page (which also consumes Python-emitted results) can read both
// interchangeably.

export interface FeatureResult {
    feature_id: string;
    library: string;
    status: 'pass' | 'fail' | 'error';
    fixture_path: string;
    run_at: string;
    tool_version: string;
    assertions: AssertionVerdict[];
}

export function aggregateStatus(verdicts: AssertionVerdict[]): 'pass' | 'fail' | 'error' {
    let worst: 'pass' | 'fail' | 'error' = 'pass';
    for (const v of verdicts) {
        if (v.status === 'error') return 'error';
        if (v.status === 'fail') worst = 'fail';
    }
    return worst;
}

export function isoNowZ(): string {
    // Strip milliseconds from the ISO-8601 timestamp so the output matches
    // the %Y-%m-%dT%H:%M:%SZ format used by the Python conformance runner.
    return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}
