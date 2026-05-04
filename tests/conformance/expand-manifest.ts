// Port of `expand_manifest` from ooxml-validate's Python
// implementation (see `src/ooxml_validate/conformance.py`).
//
// A feature manifest in the reference corpus can be either:
//
//   kind: "literal"        — single concrete case. The manifest file IS
//                            the test case. This is the default when
//                            `kind` is absent.
//   kind: "parameterised"  — a family of cases. One manifest file
//                            declares `parameters` (named axes with a
//                            list of value records each), plus
//                            `assertions_template` / `render_assertions_template`
//                            blocks that reference the axes via
//                            `{axis.field}` placeholders. At runtime we
//                            walk the Cartesian product of axes (axes
//                            sorted alphabetically for determinism,
//                            per-axis values keep their authored order)
//                            and produce one literal manifest per
//                            combination.
//
// The expanded manifest's id / fixtures.machine both gain a
// `--<axis-id>[--<axis-id>...]` suffix so they line up with the
// committed fixture files that python-docx generated (e.g.
// `fixtures/docx/font-color--red.docx`).
//
// Keep this in lockstep with the Python version. When either side
// changes shape, update both; downstream result JSONs must be
// indistinguishable regardless of which runner produced them.

export interface Manifest {
    id?: string;
    kind?: 'literal' | 'parameterised';
    fixtures?: { machine?: string; office?: string; [k: string]: unknown };
    parameters?: Record<string, Array<Record<string, unknown>>>;
    assertions?: unknown[];
    render_assertions?: unknown[];
    assertions_template?: unknown[];
    render_assertions_template?: unknown[];
    _expansion?: { parent_id: string; bindings: Record<string, string> };
    [k: string]: unknown;
}

/**
 * Expand a manifest into its concrete literal cases.
 *
 * `kind: "literal"` (or omitted) returns `[manifest]` unchanged.
 *
 * `kind: "parameterised"` walks the Cartesian product of every axis in
 * `parameters`, substitutes `{axis.field}` placeholders in every
 * `*_template` block, and returns one concrete manifest per
 * combination with:
 *
 *   - `id`              suffixed `<id>--<axis-id>[--<axis-id>...]`
 *   - `fixtures.machine` suffixed the same way
 *   - `assertions`      populated from `assertions_template`
 *   - `render_assertions` populated from `render_assertions_template`
 *   - `_expansion`      diagnostic dict carrying parent id + bindings
 *
 * Expansion is deterministic: axes are sorted alphabetically; per-axis
 * values keep their authored order.
 */
export function expandManifest(manifest: Manifest): Manifest[] {
    const kind = manifest.kind ?? 'literal';
    if (kind === 'literal') return [manifest];
    if (kind !== 'parameterised') {
        throw new Error(`Unknown manifest kind: ${JSON.stringify(kind)}`);
    }

    const parameters = manifest.parameters ?? {};
    const axes = Object.keys(parameters).sort();
    if (axes.length === 0) {
        throw new Error(
            `Parameterised manifest ${JSON.stringify(manifest.id)} has no parameters.`,
        );
    }

    const axisValues = axes.map((axis) => parameters[axis]);
    const expanded: Manifest[] = [];

    for (const combo of cartesian(axisValues)) {
        const bindings: Record<string, Record<string, unknown>> = {};
        axes.forEach((axis, i) => {
            bindings[axis] = combo[i];
        });

        const caseManifest = deepCopy(manifest) as Manifest;
        caseManifest.kind = 'literal';
        delete caseManifest.parameters;

        const suffix = axes
            .map((axis) => String(bindings[axis].id))
            .join('--');
        caseManifest.id = `${manifest.id}--${suffix}`;

        const bindingIds: Record<string, string> = {};
        for (const axis of axes) bindingIds[axis] = String(bindings[axis].id);
        caseManifest._expansion = {
            parent_id: String(manifest.id),
            bindings: bindingIds,
        };

        const fixtures = (caseManifest.fixtures ??= {});
        const machineName = String(fixtures.machine ?? manifest.id ?? '');
        fixtures.machine = `${machineName}--${suffix}`;
        if (typeof fixtures.office === 'string') {
            fixtures.office = `${fixtures.office}--${suffix}`;
        }

        const assertionsTemplate = caseManifest.assertions_template;
        delete caseManifest.assertions_template;
        if (Array.isArray(assertionsTemplate)) {
            caseManifest.assertions = assertionsTemplate.map((a) =>
                substitute(a, bindings),
            ) as unknown[];
        }

        const renderTemplate = caseManifest.render_assertions_template;
        delete caseManifest.render_assertions_template;
        if (Array.isArray(renderTemplate)) {
            caseManifest.render_assertions = renderTemplate.map((a) =>
                substitute(a, bindings),
            ) as unknown[];
        }

        expanded.push(caseManifest);
    }

    return expanded;
}

/**
 * Walk `obj` recursively. On every string leaf, replace `{axis.field}`
 * placeholders with the corresponding value in `bindings`. Unknown
 * placeholders (unknown axis, or axis without that field) pass through
 * unchanged so authoring mistakes are visible in the rendered
 * assertion rather than silently dropped.
 */
function substitute(
    obj: unknown,
    bindings: Record<string, Record<string, unknown>>,
): unknown {
    // Field portion accepts [A-Za-z][A-Za-z0-9_]* to handle camelCase
    // OOXML attribute names (e.g. `themeTint`, `numFmtId`). Must match
    // the Python reference in ooxml-validate's conformance.py.
    const placeholder = /\{([a-z][a-z0-9_]*)\.([A-Za-z][A-Za-z0-9_]*)\}/g;
    const apply = (text: string): string =>
        text.replace(placeholder, (full, axis: string, field: string) => {
            const record = bindings[axis];
            if (!record) return full;
            if (!(field in record)) return full;
            return String(record[field]);
        });

    if (typeof obj === 'string') return apply(obj);
    if (Array.isArray(obj)) return obj.map((v) => substitute(v, bindings));
    if (obj !== null && typeof obj === 'object') {
        const out: Record<string, unknown> = {};
        for (const [k, v] of Object.entries(obj)) {
            out[k] = substitute(v, bindings);
        }
        return out;
    }
    return obj;
}

function* cartesian<T>(arrays: T[][]): Generator<T[]> {
    if (arrays.length === 0) {
        yield [];
        return;
    }
    const [first, ...rest] = arrays;
    for (const item of first) {
        for (const tail of cartesian(rest)) {
            yield [item, ...tail];
        }
    }
}

function deepCopy<T>(value: T): T {
    // Manifests are pure JSON so structuredClone is overkill; JSON
    // round-trip is sufficient and avoids Node version gating.
    return JSON.parse(JSON.stringify(value)) as T;
}
