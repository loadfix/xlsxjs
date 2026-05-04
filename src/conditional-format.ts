// Parse <conditionalFormatting> blocks inside a worksheet XML into a
// renderer-friendly model, plus evaluate each rule against a cell value.
//
// Scope:
//   - cellIs                 with common operators
//   - containsText           / notContainsText / beginsWith / endsWith
//   - duplicateValues        / uniqueValues
//   - top10                  top/bottom N, optional percent
//   - expression             only the narrow "=$C2>X" form is interpreted;
//                            anything else falls back to "no match"
//   - containsBlanks         / notContainsBlanks
//   - containsErrors         / notContainsErrors
//   - aboveAverage           / belowAverage (equalAverage + stdDev shift)
//   - timePeriod             today / yesterday / tomorrow / last7Days /
//                            thisWeek / lastWeek / nextWeek /
//                            thisMonth / lastMonth / nextMonth
//
// Out of scope (rule types we ignore for now):
//   - colorScale / dataBar / iconSet (gradient / graphical rules)

import { parseCellRef } from './utils';
import { sanitizeHexColor } from './styles';
import type { ColorRef } from './theme';
import type { Cell } from './workbook-parser';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export type CfOperator =
    | 'equal' | 'notEqual'
    | 'greaterThan' | 'greaterThanOrEqual'
    | 'lessThan' | 'lessThanOrEqual'
    | 'between' | 'notBetween';

export type CfRuleType =
    | 'cellIs'
    | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
    | 'duplicateValues' | 'uniqueValues'
    | 'top10'
    | 'expression'
    | 'containsBlanks' | 'notContainsBlanks'
    | 'containsErrors' | 'notContainsErrors'
    | 'aboveAverage'
    | 'timePeriod'
    | 'colorScale' | 'dataBar' | 'iconSet'
    | 'unsupported';

// ── Graphical rule shapes ─────────────────────────────────────────────────
// Excel's "cfvo" (conditional-format value object) describes a stop or
// threshold: a position in the range (min / max / percent / percentile /
// num / formula) and an accompanying value.
export interface Cfvo {
    type: 'min' | 'max' | 'percent' | 'percentile' | 'num' | 'formula';
    value: string | null;
}

export interface ColorScale {
    cfvos: Cfvo[];          // 2 or 3 entries
    colors: (ColorRef | null)[]; // matching length
}

export interface DataBar {
    cfvos: Cfvo[];          // exactly 2 (min, max)
    color: ColorRef | null;
    minLength: number;      // 0..100 in %
    maxLength: number;      // 0..100 in %
    showValue: boolean;
}

export interface IconSet {
    iconSet: string;        // e.g. "3TrafficLights1"
    cfvos: Cfvo[];
    showValue: boolean;
    reverse: boolean;
}

export interface CfRule {
    type: CfRuleType;
    priority: number;
    dxfId: number;
    // type-specific payload:
    operator?: CfOperator;
    formulas: string[];       // raw formula text from <formula> children
    text?: string;            // containsText / beginsWith / endsWith
    rank?: number;            // top10 N
    percent?: boolean;        // top10 top/bottom by percent
    bottom?: boolean;         // top10 bottom instead of top
    stopIfTrue: boolean;
    // aboveAverage rule attributes. `aboveAverage` defaults to true (matches
    // above-mean); when false the rule matches below-mean. `equalAverage`
    // flips the inequality to inclusive (>= / <=). `stdDev` shifts the
    // threshold by N * population-stddev from the mean (may be negative).
    aboveAverage: boolean;
    equalAverage: boolean;
    stdDev: number | null;
    // timePeriod rule attribute — one of today / yesterday / tomorrow /
    // last7Days / thisWeek / lastWeek / nextWeek / thisMonth / lastMonth
    // / nextMonth. Null for rules that don't use it.
    timePeriod: string | null;
    // Graphical payloads. Present on exactly one of these rule types.
    colorScale?: ColorScale;
    dataBar?: DataBar;
    iconSet?: IconSet;
}

export interface ConditionalFormatting {
    // Set of cells covered by this block, as a list of inclusive ranges in
    // zero-based coordinates. Same sheet coverage as sqref="A1 B2:C4".
    ranges: CellRange[];
    rules: CfRule[];
}

export interface CellRange {
    col: number; row: number;
    endCol: number; endRow: number;
}

export function parseConditionalFormatting(doc: Document): ConditionalFormatting[] {
    const out: ConditionalFormatting[] = [];
    const blocks = doc.getElementsByTagNameNS(NS_MAIN, 'conditionalFormatting');
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        const sqref = block.getAttribute('sqref') ?? '';
        const ranges = parseSqref(sqref);
        const rules: CfRule[] = [];
        const ruleEls = block.getElementsByTagNameNS(NS_MAIN, 'cfRule');
        for (let j = 0; j < ruleEls.length; j++) {
            const rule = parseCfRule(ruleEls[j]);
            if (rule) rules.push(rule);
        }
        out.push({ ranges, rules });
    }
    return out;
}

export function parseSqref(sqref: string): CellRange[] {
    const parts = sqref.trim().split(/\s+/).filter((p) => p.length);
    const out: CellRange[] = [];
    for (const p of parts) {
        const r = parseRangeToken(p);
        if (r) out.push(r);
    }
    return out;
}

function parseRangeToken(tok: string): CellRange | null {
    const parts = tok.split(':');
    if (parts.length === 1) {
        const a = parseCellRef(parts[0].replace(/\$/g, ''));
        if (!a) return null;
        return { col: a.col, row: a.row, endCol: a.col, endRow: a.row };
    }
    const a = parseCellRef(parts[0].replace(/\$/g, ''));
    const b = parseCellRef(parts[1].replace(/\$/g, ''));
    if (!a || !b) return null;
    return {
        col: Math.min(a.col, b.col),
        row: Math.min(a.row, b.row),
        endCol: Math.max(a.col, b.col),
        endRow: Math.max(a.row, b.row),
    };
}

function parseCfRule(el: Element): CfRule | null {
    const rawType = el.getAttribute('type') ?? '';
    const type = asRuleType(rawType);
    const priority = Number(el.getAttribute('priority')) || 0;
    // Graphical rules (colorScale / dataBar / iconSet) have no dxfId — their
    // formatting lives on nested elements, not on a dxf record.
    const dxfIdAttr = el.getAttribute('dxfId');
    const dxfId = dxfIdAttr != null && Number.isFinite(Number(dxfIdAttr)) ? Number(dxfIdAttr) : -1;

    const formulas: string[] = [];
    const fEls = el.getElementsByTagNameNS(NS_MAIN, 'formula');
    for (let i = 0; i < fEls.length; i++) formulas.push(fEls[i].textContent ?? '');

    const operator = asOperator(el.getAttribute('operator'));
    const text = el.getAttribute('text') ?? undefined;
    const rank = el.getAttribute('rank') ? Number(el.getAttribute('rank')) : undefined;
    const percent = el.getAttribute('percent') === '1';
    const bottom = el.getAttribute('bottom') === '1';
    const stopIfTrue = el.getAttribute('stopIfTrue') === '1';

    // aboveAverage defaults to true per the schema. Only "0" flips it.
    const aboveAverageAttr = el.getAttribute('aboveAverage');
    const aboveAverage = aboveAverageAttr !== '0';
    const equalAverage = el.getAttribute('equalAverage') === '1';
    const stdDevAttr = el.getAttribute('stdDev');
    const stdDevN = stdDevAttr != null ? Number(stdDevAttr) : NaN;
    const stdDev = Number.isFinite(stdDevN) ? stdDevN : null;
    const timePeriod = el.getAttribute('timePeriod');

    const rule: CfRule = {
        type, priority, dxfId, operator, formulas, text, rank, percent, bottom, stopIfTrue,
        aboveAverage, equalAverage, stdDev,
        timePeriod: timePeriod ?? null,
    };

    if (type === 'colorScale') rule.colorScale = parseColorScale(el);
    else if (type === 'dataBar') rule.dataBar = parseDataBar(el);
    else if (type === 'iconSet') rule.iconSet = parseIconSet(el);

    return rule;
}

function parseCfvos(parent: Element): Cfvo[] {
    const out: Cfvo[] = [];
    const els = parent.getElementsByTagNameNS(NS_MAIN, 'cfvo');
    for (let i = 0; i < els.length; i++) {
        const t = els[i].getAttribute('type') ?? 'num';
        const v = els[i].getAttribute('val');
        const type = (t === 'min' || t === 'max' || t === 'percent' || t === 'percentile' || t === 'num' || t === 'formula')
            ? t : 'num';
        out.push({ type, value: v });
    }
    return out;
}

// Graphical rule <color> elements don't go through parseColorElement in
// styles.ts because the main-ns imports would be cyclic. Re-implement the
// subset we need: rgb + theme+tint (+ indexed + tint for completeness).
function parseCfColor(el: Element | null): ColorRef | null {
    if (!el) return null;
    const tintAttr = el.getAttribute('tint');
    const tint = tintAttr ? Number(tintAttr) : 0;
    const safeTint = Number.isFinite(tint) ? tint : 0;
    const rgb = el.getAttribute('rgb');
    if (rgb) {
        const sanitized = sanitizeHexColor(rgb);
        return sanitized ? { kind: 'rgb', value: sanitized } : null;
    }
    const themeAttr = el.getAttribute('theme');
    if (themeAttr != null) {
        const index = Number(themeAttr);
        if (!Number.isFinite(index) || index < 0) return null;
        return { kind: 'theme', index, tint: safeTint };
    }
    const idxAttr = el.getAttribute('indexed');
    if (idxAttr != null) {
        const index = Number(idxAttr);
        if (!Number.isFinite(index) || index < 0) return null;
        return { kind: 'indexed', index, tint: safeTint };
    }
    return null;
}

function parseColorScale(ruleEl: Element): ColorScale {
    const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'colorScale').item(0);
    if (!root) return { cfvos: [], colors: [] };
    const cfvos = parseCfvos(root);
    const colors: (ColorRef | null)[] = [];
    const colorEls = root.getElementsByTagNameNS(NS_MAIN, 'color');
    for (let i = 0; i < colorEls.length; i++) colors.push(parseCfColor(colorEls[i]));
    return { cfvos, colors };
}

function parseDataBar(ruleEl: Element): DataBar {
    const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'dataBar').item(0);
    if (!root) return { cfvos: [], color: null, minLength: 10, maxLength: 90, showValue: true };
    return {
        cfvos: parseCfvos(root),
        color: parseCfColor(root.getElementsByTagNameNS(NS_MAIN, 'color').item(0)),
        minLength: Number(root.getAttribute('minLength') ?? '10') || 10,
        maxLength: Number(root.getAttribute('maxLength') ?? '90') || 90,
        showValue: root.getAttribute('showValue') !== '0',
    };
}

function parseIconSet(ruleEl: Element): IconSet {
    const root = ruleEl.getElementsByTagNameNS(NS_MAIN, 'iconSet').item(0);
    if (!root) return { iconSet: '3TrafficLights1', cfvos: [], showValue: true, reverse: false };
    return {
        iconSet: root.getAttribute('iconSet') ?? '3TrafficLights1',
        cfvos: parseCfvos(root),
        showValue: root.getAttribute('showValue') !== '0',
        reverse: root.getAttribute('reverse') === '1',
    };
}

function asRuleType(s: string): CfRuleType {
    switch (s) {
        case 'cellIs':
        case 'containsText':
        case 'notContainsText':
        case 'beginsWith':
        case 'endsWith':
        case 'duplicateValues':
        case 'uniqueValues':
        case 'top10':
        case 'expression':
        case 'containsBlanks':
        case 'notContainsBlanks':
        case 'containsErrors':
        case 'notContainsErrors':
        case 'aboveAverage':
        case 'timePeriod':
        case 'colorScale':
        case 'dataBar':
        case 'iconSet':
            return s;
        default:
            return 'unsupported';
    }
}

function asOperator(s: string | null): CfOperator | undefined {
    if (!s) return undefined;
    switch (s) {
        case 'equal':
        case 'notEqual':
        case 'greaterThan':
        case 'greaterThanOrEqual':
        case 'lessThan':
        case 'lessThanOrEqual':
        case 'between':
        case 'notBetween':
            return s;
        default:
            return undefined;
    }
}

// ── rule evaluation ───────────────────────────────────────────────────────

// Cached "column of values" lookups for the top10 / duplicate rule families.
// Built lazily once per sheet render.
export interface CfContext {
    // rangeKey → list of (col,row,Cell). Used by top10 / duplicateValues to
    // know which cells they're ranking against.
    cellsInRange(range: CellRange): { col: number; row: number; cell: Cell | null }[];
    // Workbook-level date epoch. The timePeriod evaluator needs it to turn
    // a cell's numeric serial into a calendar date using the same rules the
    // number-format formatter uses elsewhere. Optional so the context can
    // still be built from code that predates the timePeriod evaluator.
    date1904?: boolean;
}

export function evaluateRule(rule: CfRule, cell: Cell | null, range: CellRange, ctx: CfContext): boolean {
    // Blank / error rules have to fire even when the cell is empty or null,
    // so branch before the early null return. Every other rule still needs
    // a cell to compare against.
    if (rule.type === 'containsBlanks')    return evalContainsBlanks(cell);
    if (rule.type === 'notContainsBlanks') return !evalContainsBlanks(cell);
    if (rule.type === 'containsErrors')    return evalContainsErrors(cell);
    if (rule.type === 'notContainsErrors') return !evalContainsErrors(cell);

    if (!cell) return false;
    switch (rule.type) {
        case 'cellIs':         return evalCellIs(rule, cell);
        case 'containsText':   return containsText(cell, rule.text, false);
        case 'notContainsText':return !containsText(cell, rule.text, false);
        case 'beginsWith':     return startsOrEnds(cell, rule.text, 'begin');
        case 'endsWith':       return startsOrEnds(cell, rule.text, 'end');
        case 'duplicateValues':return evalDuplicate(cell, range, ctx, true);
        case 'uniqueValues':   return evalDuplicate(cell, range, ctx, false);
        case 'top10':          return evalTop10(rule, cell, range, ctx);
        case 'expression':     return evalExpression(rule, cell, range, ctx);
        case 'aboveAverage':   return evalAboveAverage(rule, cell, range, ctx);
        case 'timePeriod':     return evalTimePeriod(rule, cell, ctx);
        case 'unsupported':    return false;
    }
    return false;
}

// containsBlanks matches cells that are genuinely empty: no cell at all,
// kind === 'empty', or a present cell whose value string is empty.
function evalContainsBlanks(cell: Cell | null): boolean {
    if (!cell) return true;
    if (cell.kind === 'empty') return true;
    if (cell.value === '' || cell.value == null) return true;
    return false;
}

function evalContainsErrors(cell: Cell | null): boolean {
    if (!cell) return false;
    return cell.kind === 'error';
}

function numericValue(cell: Cell): number | null {
    if (cell.kind !== 'number' && cell.kind !== 'boolean') {
        const n = Number(cell.value);
        return Number.isFinite(n) ? n : null;
    }
    if (cell.kind === 'boolean') return cell.value === 'TRUE' ? 1 : 0;
    const n = Number(cell.value);
    return Number.isFinite(n) ? n : null;
}

function parseFormulaLiteral(text: string | undefined): number | string | null {
    if (text === undefined) return null;
    const t = text.trim();
    if (t.startsWith('"') && t.endsWith('"')) return t.slice(1, -1);
    const n = Number(t);
    if (Number.isFinite(n)) return n;
    return null;
}

function evalCellIs(rule: CfRule, cell: Cell): boolean {
    const a = parseFormulaLiteral(rule.formulas[0]);
    const b = parseFormulaLiteral(rule.formulas[1]);
    if (a === null) return false;

    if (typeof a === 'number') {
        const v = numericValue(cell);
        if (v === null) return false;
        switch (rule.operator) {
            case 'equal':              return v === a;
            case 'notEqual':           return v !== a;
            case 'greaterThan':        return v > a;
            case 'greaterThanOrEqual': return v >= a;
            case 'lessThan':           return v < a;
            case 'lessThanOrEqual':    return v <= a;
            case 'between':            return typeof b === 'number' && v >= a && v <= b;
            case 'notBetween':         return typeof b === 'number' && (v < a || v > b);
            default:                   return false;
        }
    }
    // String comparisons — Excel is case-insensitive here.
    const sv = (cell.value ?? '').toLocaleLowerCase();
    const sa = (a as string).toLocaleLowerCase();
    switch (rule.operator) {
        case 'equal':    return sv === sa;
        case 'notEqual': return sv !== sa;
        default:         return false;
    }
}

function containsText(cell: Cell, text: string | undefined, _not: boolean): boolean {
    if (!text) return false;
    return (cell.value ?? '').toLocaleLowerCase().includes(text.toLocaleLowerCase());
}

function startsOrEnds(cell: Cell, text: string | undefined, side: 'begin' | 'end'): boolean {
    if (!text) return false;
    const v = (cell.value ?? '').toLocaleLowerCase();
    const t = text.toLocaleLowerCase();
    return side === 'begin' ? v.startsWith(t) : v.endsWith(t);
}

function evalDuplicate(cell: Cell, range: CellRange, ctx: CfContext, findDuplicate: boolean): boolean {
    const cells = ctx.cellsInRange(range);
    let seen = 0;
    for (const entry of cells) {
        const other = entry.cell;
        if (!other) continue;
        if (other.value === cell.value) seen++;
        if (seen > 1) break;
    }
    return findDuplicate ? seen > 1 : seen === 1;
}

function evalTop10(rule: CfRule, cell: Cell, range: CellRange, ctx: CfContext): boolean {
    const cells = ctx.cellsInRange(range);
    const values: { cell: Cell; value: number }[] = [];
    for (const entry of cells) {
        if (!entry.cell) continue;
        const n = numericValue(entry.cell);
        if (n === null) continue;
        values.push({ cell: entry.cell, value: n });
    }
    if (values.length === 0) return false;
    values.sort((a, b) => rule.bottom ? a.value - b.value : b.value - a.value);
    const rank = rule.rank ?? 10;
    const n = rule.percent ? Math.max(1, Math.ceil(values.length * rank / 100)) : Math.min(rank, values.length);
    const threshold = values[n - 1]?.value;
    if (threshold === undefined) return false;
    const v = numericValue(cell);
    if (v === null) return false;
    return rule.bottom ? v <= threshold : v >= threshold;
}

// aboveAverage: compute the mean of every numeric value in the range, then
// test the cell's numeric value against the mean (optionally shifted by
// stdDev * population-stddev). `aboveAverage` flips the comparison to
// below-mean; `equalAverage` flips the inequality to inclusive (>= / <=).
function evalAboveAverage(rule: CfRule, cell: Cell, range: CellRange, ctx: CfContext): boolean {
    const cells = ctx.cellsInRange(range);
    const values: number[] = [];
    for (const entry of cells) {
        if (!entry.cell) continue;
        const n = numericValue(entry.cell);
        if (n === null) continue;
        values.push(n);
    }
    if (values.length === 0) return false;

    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    let threshold = mean;
    if (rule.stdDev !== null && rule.stdDev !== undefined) {
        // Population stddev: sqrt(sum((x-mean)^2) / n).
        let variance = 0;
        for (const v of values) {
            const d = v - mean;
            variance += d * d;
        }
        variance /= values.length;
        const stddev = Math.sqrt(variance);
        threshold = mean + rule.stdDev * stddev;
    }

    const v = numericValue(cell);
    if (v === null) return false;
    const above = rule.aboveAverage;
    const incl = rule.equalAverage;
    if (above) return incl ? v >= threshold : v > threshold;
    return incl ? v <= threshold : v < threshold;
}

// Excel date epochs. 1900 system: serial 1 = 1900-01-01, but Excel also
// treats the (non-existent) 1900-02-29 as serial 60 — hence the "fudge"
// that the number-format formatter applies. 1904 system: serial 0 =
// 1904-01-01, no fudge. We don't have access to the formatter's helpers
// here (would be a circular import), so replicate the core conversion.
const MS_PER_DAY = 86400000;
const EPOCH_MS_1900 = Date.UTC(1899, 11, 30); // 1899-12-30 sentinel
const EPOCH_MS_1904 = Date.UTC(1904, 0, 1);

function serialToDate(serial: number, date1904: boolean): Date {
    let days = Math.floor(serial);
    // 1900 leap-year fudge: serials ≥ 60 are off by one day because Excel
    // counts 1900-02-29 as a real date.
    if (!date1904 && days >= 60) days -= 1;
    const epoch = date1904 ? EPOCH_MS_1904 : EPOCH_MS_1900;
    return new Date(epoch + days * MS_PER_DAY);
}

// Normalise a Date to midnight UTC so day-level comparisons don't wobble
// over timezone boundaries. The cell's serial is already in UTC-space so
// the comparison today value must live in the same space.
function startOfDayUtc(d: Date): number {
    return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

// Start of the ISO-ish week containing `d`, anchored on Sunday (Excel's
// default "first day of week"). Returns a UTC-midnight ms timestamp.
function startOfWeekUtc(d: Date): number {
    const dow = d.getUTCDay(); // 0 = Sun … 6 = Sat
    return startOfDayUtc(d) - dow * MS_PER_DAY;
}

function evalTimePeriod(rule: CfRule, cell: Cell, ctx: CfContext): boolean {
    if (!rule.timePeriod) return false;
    const n = numericValue(cell);
    if (n === null) return false;
    // Note: we use the workbook-level 1900/1904 epoch from the context,
    // falling back to 1900 when the ambient context didn't pass it. The
    // docstring on makeCfContext in html-renderer threads this through.
    const date1904 = ctx.date1904 === true;
    const cellDate = serialToDate(n, date1904);
    const cellDay = startOfDayUtc(cellDate);

    const now = new Date();
    const today = startOfDayUtc(now);
    const yesterday = today - MS_PER_DAY;
    const tomorrow = today + MS_PER_DAY;
    const thisWeekStart = startOfWeekUtc(now);
    const lastWeekStart = thisWeekStart - 7 * MS_PER_DAY;
    const nextWeekStart = thisWeekStart + 7 * MS_PER_DAY;
    const last7Start = today - 6 * MS_PER_DAY; // inclusive of today
    const thisMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1);
    const thisMonthEnd = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1);
    const lastMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1);
    const nextMonthStart = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1);
    const nextMonthEnd = Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 2, 1);

    switch (rule.timePeriod) {
        case 'today':     return cellDay === today;
        case 'yesterday': return cellDay === yesterday;
        case 'tomorrow':  return cellDay === tomorrow;
        case 'last7Days': return cellDay >= last7Start && cellDay <= today;
        case 'thisWeek':  return cellDay >= thisWeekStart && cellDay < nextWeekStart;
        case 'lastWeek':  return cellDay >= lastWeekStart && cellDay < thisWeekStart;
        case 'nextWeek':  return cellDay >= nextWeekStart && cellDay < nextWeekStart + 7 * MS_PER_DAY;
        case 'thisMonth': return cellDay >= thisMonthStart && cellDay < thisMonthEnd;
        case 'lastMonth': return cellDay >= lastMonthStart && cellDay < thisMonthStart;
        case 'nextMonth': return cellDay >= nextMonthStart && cellDay < nextMonthEnd;
        default: return false;
    }
}

// Resolve a cfvo against the numeric values in a range. Returns the
// threshold number, or null when the cfvo can't be resolved (e.g. formula
// type, which we don't evaluate).
export function resolveCfvo(cfvo: Cfvo, values: number[]): number | null {
    if (values.length === 0) return null;
    const min = Math.min(...values);
    const max = Math.max(...values);
    switch (cfvo.type) {
        case 'min': return min;
        case 'max': return max;
        case 'num': {
            const n = Number(cfvo.value);
            return Number.isFinite(n) ? n : null;
        }
        case 'percent': {
            const p = Number(cfvo.value);
            if (!Number.isFinite(p)) return null;
            return min + (max - min) * (p / 100);
        }
        case 'percentile': {
            const p = Number(cfvo.value);
            if (!Number.isFinite(p)) return null;
            const sorted = values.slice().sort((a, b) => a - b);
            const rank = (p / 100) * (sorted.length - 1);
            const lo = Math.floor(rank);
            const hi = Math.ceil(rank);
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (rank - lo);
        }
        default:
            return null;
    }
}

// Interpolate a value through a list of (threshold, hex color) stops.
// Returns '#rrggbb'. Assumes stops are sorted by threshold ascending.
export function interpolateColorScale(
    value: number,
    stops: { threshold: number; hex: string }[],
): string | null {
    if (stops.length === 0) return null;
    if (value <= stops[0].threshold) return stops[0].hex;
    if (value >= stops[stops.length - 1].threshold) return stops[stops.length - 1].hex;
    for (let i = 0; i < stops.length - 1; i++) {
        const a = stops[i], b = stops[i + 1];
        if (value >= a.threshold && value <= b.threshold) {
            const t = b.threshold === a.threshold ? 0 : (value - a.threshold) / (b.threshold - a.threshold);
            return lerpHex(a.hex, b.hex, t);
        }
    }
    return null;
}

function lerpHex(a: string, b: string, t: number): string {
    const pa = parseHex(a), pb = parseHex(b);
    if (!pa || !pb) return a;
    const mix = (x: number, y: number) => Math.round(x + (y - x) * t);
    return `#${toHex(mix(pa[0], pb[0]))}${toHex(mix(pa[1], pb[1]))}${toHex(mix(pa[2], pb[2]))}`;
}

function parseHex(h: string): [number, number, number] | null {
    const m = /^#([0-9a-f]{6})$/i.exec(h);
    if (!m) return null;
    return [parseInt(m[1].slice(0, 2), 16), parseInt(m[1].slice(2, 4), 16), parseInt(m[1].slice(4, 6), 16)];
}

function toHex(n: number): string {
    return Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0');
}

// Narrow "expression" support — only the common form where the formula is
// "=<cellRef> <op> <literal>" (e.g. "=$C2>=100") or "=<cellRef>". Anything
// more complex (AND/OR/MATCH/INDIRECT/…) returns false.
function evalExpression(rule: CfRule, cell: Cell, _range: CellRange, _ctx: CfContext): boolean {
    const raw = rule.formulas[0] ?? '';
    const src = raw.trim().replace(/^=/, '');
    const m = /^\$?([A-Z]+)\$?([1-9][0-9]*)\s*(<=|>=|<>|=|<|>)\s*(.+)$/.exec(src);
    if (!m) {
        // Bare reference → truthy if the referenced cell is truthy. Only
        // honour it when the formula points at the cell under evaluation.
        return false;
    }
    const [, , , op, rhs] = m;
    const v = numericValue(cell);
    const rhsLit = parseFormulaLiteral(rhs);
    if (v === null || typeof rhsLit !== 'number') return false;
    switch (op) {
        case '<':  return v < rhsLit;
        case '<=': return v <= rhsLit;
        case '>':  return v > rhsLit;
        case '>=': return v >= rhsLit;
        case '=':  return v === rhsLit;
        case '<>': return v !== rhsLit;
        default: return false;
    }
}
