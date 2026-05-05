// Formula evaluator v2. Recursive-descent parser + tree-walk evaluator with
// a first-class Excel error taxonomy. Opt-in via Options.evaluateFormulas.
//
// Design:
//   · Lexer splits the formula body (without leading `=`) into tokens.
//   · Parser is recursive-descent: expr → or → and → not → compare → add →
//     mul → pow → unary → primary.
//   · Primary handles numbers, strings, booleans, errors, cell refs,
//     cell ranges (returned as an array), and function calls.
//   · Evaluator resolves cell refs against a CellResolver that the caller
//     populates from the parsed Sheet.
//
// Error taxonomy (Wave-12 upgrade): errors are tagged with an Excel error
// code so the renderer can surface the matching `#DIV/0!`, `#VALUE!`,
// `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!` sentinel rather than a
// generic `#ERROR!`. Errors propagate through arithmetic (any op on an
// error returns that same error) and IFERROR / IFNA catch them.
//
// Deliberate limitations (documented in TODO.md):
//   · A1 / absolute ($A$1) only — no sheet-qualified refs (Sheet2!A1),
//     no structured references ([Table1[col]]), no named ranges.
//   · No array formulas, no spill semantics — a cell range always reduces
//     to a flat list of cells.
//   · Function coverage is still a subset; VLOOKUP supports exact-match
//     only (approximate-match requests surface as #VALUE! with a comment).
//   · Circular refs abort after a 1024-hop depth cap.

import { parseCellRef, columnLettersToIndex } from './utils';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

// A value at evaluation time. Arrays arise from cell ranges (`A1:B5` expands
// to a flat list in row-major order). `null` stands for an empty cell; it
// folds to 0 in numeric contexts and to the empty string when coerced.
export type FormulaValue = number | string | boolean | null | FormulaValue[];

// Provides cell values for ref / range resolution. The evaluator calls this
// once per referenced cell. Returning `null` means "empty cell" (not an
// error). Return a `string` starting with `#` to surface an Excel-style
// error that the evaluator will propagate.
export interface CellResolver {
    (col: number, row: number): FormulaValue;
}

// Excel error codes. Every error the evaluator surfaces carries one of
// these. Back-compat: the top-level evaluateFormula() collapses them to
// the bare `#<CODE>!` / `#<CODE>?` strings that Excel actually prints.
export type ErrorCode = '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#NULL!';

// Internal error carrier. Tagged so the evaluator can distinguish a legit
// string result starting with `#` (unlikely but possible) from an error.
export interface FormulaError {
    __xlsxError: true;
    code: ErrorCode;
}

export function makeError(code: ErrorCode): FormulaError {
    return { __xlsxError: true, code };
}

export function isFormulaError(v: unknown): v is FormulaError {
    return !!(v && typeof v === 'object' && (v as FormulaError).__xlsxError === true);
}

// Back-compat: callers that pattern-matched against the old `ERROR` symbol
// can still import it. Internally we now pass around `FormulaError` objects.
export const ERROR = Symbol('formula-error');

export type EvalResult = FormulaValue | FormulaError;

// Map a raw sentinel string (as it appears in a cached <v> element or an
// `=#DIV/0!` literal) to an ErrorCode. Returns #VALUE! as a last-ditch
// fallback when the string starts with `#` but isn't a recognised code.
function parseErrorLiteral(text: string): ErrorCode {
    const up = text.toUpperCase();
    if (up.startsWith('#DIV')) return '#DIV/0!';
    if (up.startsWith('#VALUE')) return '#VALUE!';
    if (up.startsWith('#REF')) return '#REF!';
    if (up.startsWith('#NAME')) return '#NAME?';
    if (up.startsWith('#NUM')) return '#NUM!';
    if (up.startsWith('#N/A') || up === '#NA') return '#N/A';
    if (up.startsWith('#NULL')) return '#NULL!';
    return '#VALUE!';
}

// ---------------------------------------------------------------------------
// Lexer
// ---------------------------------------------------------------------------

type TokenKind =
    | 'num' | 'str' | 'ident' | 'ref' | 'bool' | 'err'
    | 'lparen' | 'rparen' | 'comma' | 'colon'
    | 'plus' | 'minus' | 'star' | 'slash' | 'caret' | 'percent'
    | 'amp'
    | 'eq' | 'neq' | 'lt' | 'lte' | 'gt' | 'gte'
    | 'eof';

interface Token {
    kind: TokenKind;
    // Raw text of the token; for 'num' use numeric field.
    text: string;
    // Populated for numeric literals.
    numeric?: number;
    // Populated for cell refs: parsed {col,row}.
    ref?: { col: number; row: number };
}

const CELL_REF_TOKEN = /^\$?([A-Z]+)\$?([1-9][0-9]*)/;

function lex(src: string): Token[] {
    const out: Token[] = [];
    let i = 0;
    while (i < src.length) {
        const c = src[i];
        if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
        // Strings: double-quoted, "" escapes embedded quote.
        if (c === '"') {
            let j = i + 1;
            let s = '';
            while (j < src.length) {
                if (src[j] === '"') {
                    if (src[j + 1] === '"') { s += '"'; j += 2; continue; }
                    break;
                }
                s += src[j]; j++;
            }
            out.push({ kind: 'str', text: s });
            i = j + 1;
            continue;
        }
        // Numbers: digits optionally with decimal, no sign (unary minus handled in parser).
        if ((c >= '0' && c <= '9') || (c === '.' && src[i + 1] >= '0' && src[i + 1] <= '9')) {
            let j = i;
            while (j < src.length && ((src[j] >= '0' && src[j] <= '9') || src[j] === '.')) j++;
            // Scientific notation
            if (j < src.length && (src[j] === 'e' || src[j] === 'E')) {
                j++;
                if (src[j] === '+' || src[j] === '-') j++;
                while (j < src.length && src[j] >= '0' && src[j] <= '9') j++;
            }
            const text = src.slice(i, j);
            out.push({ kind: 'num', text, numeric: Number(text) });
            i = j;
            continue;
        }
        // Cell ref: optional $ + letters + optional $ + digits.
        // Letters must be followed (possibly through $) by a digit; otherwise
        // the letters are an identifier (function name, TRUE/FALSE, etc.).
        if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '$') {
            // Try cell ref first (uppercased slice).
            const upper = src.slice(i, i + 24).toUpperCase();
            const m = CELL_REF_TOKEN.exec(upper);
            if (m) {
                const refText = src.slice(i, i + m[0].length);
                const col = columnLettersToIndex(m[1]);
                const row = Number(m[2]) - 1;
                if (col >= 0) {
                    out.push({ kind: 'ref', text: refText, ref: { col, row } });
                    i += m[0].length;
                    continue;
                }
            }
            // Identifier (function name or TRUE/FALSE).
            if (c === '$') {
                // Dangling $: treat as error-ish; parser will barf.
                i++;
                continue;
            }
            let j = i;
            while (j < src.length && (
                (src[j] >= 'A' && src[j] <= 'Z') ||
                (src[j] >= 'a' && src[j] <= 'z') ||
                (src[j] >= '0' && src[j] <= '9') ||
                src[j] === '_' || src[j] === '.'
            )) j++;
            const id = src.slice(i, j).toUpperCase();
            if (id === 'TRUE') out.push({ kind: 'bool', text: 'TRUE' });
            else if (id === 'FALSE') out.push({ kind: 'bool', text: 'FALSE' });
            else out.push({ kind: 'ident', text: id });
            i = j;
            continue;
        }
        // Error literals: #NAME?, #DIV/0!, #VALUE!, #REF!, #NULL!, #N/A, #NUM!
        if (c === '#') {
            let j = i + 1;
            while (j < src.length && src[j] !== ' ' && src[j] !== ',' && src[j] !== ')' && src[j] !== '(' ) j++;
            out.push({ kind: 'err', text: src.slice(i, j) });
            i = j;
            continue;
        }
        // Operators / punctuation.
        if (c === '(') { out.push({ kind: 'lparen', text: c }); i++; continue; }
        if (c === ')') { out.push({ kind: 'rparen', text: c }); i++; continue; }
        if (c === ',') { out.push({ kind: 'comma', text: c }); i++; continue; }
        if (c === ':') { out.push({ kind: 'colon', text: c }); i++; continue; }
        if (c === '+') { out.push({ kind: 'plus', text: c }); i++; continue; }
        if (c === '-') { out.push({ kind: 'minus', text: c }); i++; continue; }
        if (c === '*') { out.push({ kind: 'star', text: c }); i++; continue; }
        if (c === '/') { out.push({ kind: 'slash', text: c }); i++; continue; }
        if (c === '^') { out.push({ kind: 'caret', text: c }); i++; continue; }
        if (c === '%') { out.push({ kind: 'percent', text: c }); i++; continue; }
        if (c === '&') { out.push({ kind: 'amp', text: c }); i++; continue; }
        if (c === '=') { out.push({ kind: 'eq', text: c }); i++; continue; }
        if (c === '<') {
            if (src[i + 1] === '=') { out.push({ kind: 'lte', text: '<=' }); i += 2; continue; }
            if (src[i + 1] === '>') { out.push({ kind: 'neq', text: '<>' }); i += 2; continue; }
            out.push({ kind: 'lt', text: '<' }); i++; continue;
        }
        if (c === '>') {
            if (src[i + 1] === '=') { out.push({ kind: 'gte', text: '>=' }); i += 2; continue; }
            out.push({ kind: 'gt', text: '>' }); i++; continue;
        }
        // Unrecognised: skip to avoid infinite loops; parser will report an error
        // when it hits EOF before expected.
        i++;
    }
    out.push({ kind: 'eof', text: '' });
    return out;
}

// ---------------------------------------------------------------------------
// AST
// ---------------------------------------------------------------------------

type Ast =
    | { kind: 'num'; value: number }
    | { kind: 'str'; value: string }
    | { kind: 'bool'; value: boolean }
    | { kind: 'err'; code: ErrorCode }
    | { kind: 'ref'; col: number; row: number }
    | { kind: 'range'; col1: number; row1: number; col2: number; row2: number }
    | { kind: 'unary'; op: '+' | '-'; arg: Ast }
    | { kind: 'postfix'; op: '%'; arg: Ast }
    | { kind: 'bin'; op: string; left: Ast; right: Ast }
    | { kind: 'call'; name: string; args: Ast[] };

// ---------------------------------------------------------------------------
// Parser (recursive-descent)
// ---------------------------------------------------------------------------

class Parser {
    pos = 0;
    constructor(private toks: Token[]) {}

    peek(): Token { return this.toks[this.pos]; }
    eat(kind: TokenKind): Token | null {
        if (this.toks[this.pos].kind === kind) return this.toks[this.pos++];
        return null;
    }
    expect(kind: TokenKind): Token {
        const t = this.eat(kind);
        if (!t) throw new Error(`expected ${kind}, got ${this.toks[this.pos].kind}`);
        return t;
    }

    parseExpr(): Ast { return this.parseCompare(); }

    parseCompare(): Ast {
        let left = this.parseConcat();
        while (true) {
            const t = this.peek();
            if (t.kind === 'eq' || t.kind === 'neq' || t.kind === 'lt' ||
                t.kind === 'lte' || t.kind === 'gt' || t.kind === 'gte') {
                this.pos++;
                const right = this.parseConcat();
                left = { kind: 'bin', op: t.kind, left, right };
            } else break;
        }
        return left;
    }

    parseConcat(): Ast {
        let left = this.parseAdd();
        while (this.peek().kind === 'amp') {
            this.pos++;
            const right = this.parseAdd();
            left = { kind: 'bin', op: '&', left, right };
        }
        return left;
    }

    parseAdd(): Ast {
        let left = this.parseMul();
        while (true) {
            const t = this.peek();
            if (t.kind === 'plus' || t.kind === 'minus') {
                this.pos++;
                const right = this.parseMul();
                left = { kind: 'bin', op: t.kind === 'plus' ? '+' : '-', left, right };
            } else break;
        }
        return left;
    }

    parseMul(): Ast {
        let left = this.parsePow();
        while (true) {
            const t = this.peek();
            if (t.kind === 'star' || t.kind === 'slash') {
                this.pos++;
                const right = this.parsePow();
                left = { kind: 'bin', op: t.kind === 'star' ? '*' : '/', left, right };
            } else break;
        }
        return left;
    }

    parsePow(): Ast {
        // right-associative per Excel (2^3^2 == 2^(3^2) == 512).
        const left = this.parseUnary();
        if (this.peek().kind === 'caret') {
            this.pos++;
            const right = this.parsePow();
            return { kind: 'bin', op: '^', left, right };
        }
        return left;
    }

    parseUnary(): Ast {
        const t = this.peek();
        if (t.kind === 'minus') { this.pos++; return { kind: 'unary', op: '-', arg: this.parseUnary() }; }
        if (t.kind === 'plus')  { this.pos++; return { kind: 'unary', op: '+', arg: this.parseUnary() }; }
        return this.parsePostfix();
    }

    parsePostfix(): Ast {
        let arg = this.parsePrimary();
        while (this.peek().kind === 'percent') {
            this.pos++;
            arg = { kind: 'postfix', op: '%', arg };
        }
        return arg;
    }

    parsePrimary(): Ast {
        const t = this.peek();
        if (t.kind === 'num') { this.pos++; return { kind: 'num', value: t.numeric! }; }
        if (t.kind === 'str') { this.pos++; return { kind: 'str', value: t.text }; }
        if (t.kind === 'bool') { this.pos++; return { kind: 'bool', value: t.text === 'TRUE' }; }
        if (t.kind === 'err') { this.pos++; return { kind: 'err', code: parseErrorLiteral(t.text) }; }
        if (t.kind === 'lparen') {
            this.pos++;
            const inner = this.parseExpr();
            this.expect('rparen');
            return inner;
        }
        if (t.kind === 'ref') {
            this.pos++;
            const { col, row } = t.ref!;
            // Range?
            if (this.peek().kind === 'colon') {
                this.pos++;
                const t2 = this.expect('ref');
                const { col: col2, row: row2 } = t2.ref!;
                return {
                    kind: 'range',
                    col1: Math.min(col, col2), row1: Math.min(row, row2),
                    col2: Math.max(col, col2), row2: Math.max(row, row2),
                };
            }
            return { kind: 'ref', col, row };
        }
        if (t.kind === 'ident') {
            this.pos++;
            // Function call.
            if (this.peek().kind === 'lparen') {
                this.pos++;
                const args: Ast[] = [];
                if (this.peek().kind !== 'rparen') {
                    args.push(this.parseExpr());
                    while (this.peek().kind === 'comma') {
                        this.pos++;
                        args.push(this.parseExpr());
                    }
                }
                this.expect('rparen');
                return { kind: 'call', name: t.text, args };
            }
            // Bare identifier (no parens): unknown name → #NAME?
            return { kind: 'err', code: '#NAME?' };
        }
        throw new Error(`unexpected token ${t.kind} (${t.text})`);
    }
}

export function parseFormula(src: string): Ast {
    const body = src.startsWith('=') ? src.slice(1) : src;
    const toks = lex(body);
    const p = new Parser(toks);
    const ast = p.parseExpr();
    if (p.peek().kind !== 'eof') {
        throw new Error(`trailing tokens after expr: ${p.peek().kind}`);
    }
    return ast;
}

// ---------------------------------------------------------------------------
// Evaluator helpers
// ---------------------------------------------------------------------------

// Coerce a single value to a number. null → 0, true → 1, false → 0, number
// passes through, string parses (empty → 0, non-numeric → #VALUE!).
function toNumber(v: FormulaValue | FormulaError): number | FormulaError {
    if (isFormulaError(v)) return v;
    if (v === null) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
        if (v === '') return 0;
        if (v.startsWith('#')) return makeError(parseErrorLiteral(v));
        const n = Number(v);
        if (Number.isNaN(n)) return makeError('#VALUE!');
        return n;
    }
    // Arrays shouldn't reach here — callers should flatten first.
    return makeError('#VALUE!');
}

function toBool(v: FormulaValue | FormulaError): boolean | FormulaError {
    if (isFormulaError(v)) return v;
    if (v === null) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
        const u = v.toUpperCase();
        if (u === 'TRUE') return true;
        if (u === 'FALSE') return false;
        if (v.startsWith('#')) return makeError(parseErrorLiteral(v));
        const n = Number(v);
        if (Number.isNaN(n)) return makeError('#VALUE!');
        return n !== 0;
    }
    return makeError('#VALUE!');
}

function toStr(v: FormulaValue | FormulaError): string | FormulaError {
    if (isFormulaError(v)) return v;
    if (v === null) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return cleanNumberString(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (Array.isArray(v)) return toStr(v.length > 0 ? v[0] : null);
    return makeError('#VALUE!');
}

// Flatten an EvalResult to an array of scalars. Arrays flatten recursively.
function flatten(v: FormulaValue | FormulaError): (FormulaValue | FormulaError)[] {
    if (isFormulaError(v)) return [v];
    if (Array.isArray(v)) {
        const out: (FormulaValue | FormulaError)[] = [];
        for (const x of v) for (const y of flatten(x)) out.push(y);
        return out;
    }
    return [v];
}

// ---------------------------------------------------------------------------
// Number-format routing: evaluator-side float-noise cleanup.
//
// JavaScript's default `(0.1 + 0.2).toString()` returns '0.30000000000000004'.
// Excel's General format rounds to ~15 significant digits, which hides the
// noise. We replicate that behaviour here so numeric results land in
// `cell.value` as a clean string. This keeps the downstream renderer path
// consistent: a cached `<v>0.3</v>` cell and an evaluated `=0.1+0.2` cell
// now both render identically through formatNumber().
//
// We only apply the cleanup to non-integer finite numbers. Integers, +/-Inf,
// and NaN fall through unchanged so we don't accidentally trim trailing zeros
// the caller explicitly cached (e.g. `=SUM(...)` → 60).
export function cleanNumberString(n: number): string {
    if (!Number.isFinite(n)) return String(n);
    if (Number.isInteger(n)) return String(n);
    // toPrecision(15) matches Excel's 15-significant-digit cap. Then parseFloat
    // round-trips to strip artefacts like trailing zeros and redundant exponents
    // while preserving value.
    const cleaned = parseFloat(n.toPrecision(15));
    return String(cleaned);
}

// ---------------------------------------------------------------------------
// Built-in functions.
// ---------------------------------------------------------------------------

// Built-in functions. Each receives the raw evaluated args (one per
// positional argument; ranges stay as arrays). Returns EvalResult.
type Fn = (args: EvalResult[]) => EvalResult;

// Criteria matching for SUMIF / COUNTIF. Excel criteria forms supported:
//   ">10"  "<=20"  "=foo"  "<>x"  bare literal (exact equality).
// Numeric criteria coerce the cell to a number before comparing; string
// criteria use case-insensitive string equality.
function buildPredicate(criteria: FormulaValue | FormulaError): ((v: FormulaValue) => boolean) | FormulaError {
    if (isFormulaError(criteria)) return criteria;
    // Number criteria: treat as "= number".
    if (typeof criteria === 'number') {
        return (v) => {
            if (typeof v === 'number') return v === criteria;
            if (typeof v === 'string' && v !== '') {
                const n = Number(v); return !Number.isNaN(n) && n === criteria;
            }
            return false;
        };
    }
    if (typeof criteria === 'boolean') {
        return (v) => v === criteria;
    }
    if (criteria === null) return (v) => v === null || v === '' || v === 0;
    if (Array.isArray(criteria)) return buildPredicate(criteria.length > 0 ? criteria[0] : null);
    // String criteria — strip operator prefix.
    const raw = String(criteria);
    const m = /^\s*(<=|>=|<>|=|>|<)\s*(.*)$/.exec(raw);
    const op = m ? m[1] : '=';
    const rest = m ? m[2] : raw;
    // Try to coerce `rest` to a number; if it works, do numeric compare,
    // else compare as string.
    const restNum = rest === '' ? null : Number(rest);
    const numericRhs = rest !== '' && !Number.isNaN(restNum as number);
    const cmpNum = (a: number, b: number): boolean => {
        switch (op) {
            case '>': return a > b;
            case '<': return a < b;
            case '>=': return a >= b;
            case '<=': return a <= b;
            case '<>': return a !== b;
            case '=': default: return a === b;
        }
    };
    const cmpStr = (a: string, b: string): boolean => {
        const la = a.toLowerCase(); const lb = b.toLowerCase();
        switch (op) {
            case '>': return la > lb;
            case '<': return la < lb;
            case '>=': return la >= lb;
            case '<=': return la <= lb;
            case '<>': return la !== lb;
            case '=': default: return la === lb;
        }
    };
    return (v) => {
        if (numericRhs) {
            const rhsN = restNum as number;
            if (typeof v === 'number') return cmpNum(v, rhsN);
            if (typeof v === 'string' && v !== '') {
                const n = Number(v);
                if (!Number.isNaN(n)) return cmpNum(n, rhsN);
            }
            // Non-numeric cell vs numeric criteria: only `<>` can be true.
            return op === '<>';
        }
        // String criteria.
        if (v === null) return rest === '' && (op === '=' || op === '<=' || op === '>=');
        if (typeof v === 'boolean') return cmpStr(v ? 'TRUE' : 'FALSE', rest);
        if (typeof v === 'number') {
            // Numeric cell vs string criteria: only `<>` can be true.
            return op === '<>';
        }
        return cmpStr(String(v), rest);
    };
}

const FUNCTIONS: Record<string, Fn> = {
    SUM: (args) => {
        let total = 0;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                // SUM skips strings and booleans that came from ranges (Excel
                // behaviour). For direct scalar args, booleans/strings coerce.
                if (v === null) continue;
                if (typeof v === 'string') {
                    if (v === '') continue;
                    const n = Number(v);
                    if (Number.isNaN(n)) continue; // skip non-numeric strings
                    total += n;
                    continue;
                }
                if (typeof v === 'boolean') { total += v ? 1 : 0; continue; }
                total += v as number;
            }
        }
        return total;
    },
    AVERAGE: (args) => {
        let total = 0; let n = 0;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                if (v === null) continue;
                if (typeof v === 'string') {
                    if (v === '') continue;
                    const num = Number(v);
                    if (Number.isNaN(num)) continue;
                    total += num; n++; continue;
                }
                if (typeof v === 'boolean') { total += v ? 1 : 0; n++; continue; }
                total += v as number; n++;
            }
        }
        if (n === 0) return makeError('#DIV/0!');
        return total / n;
    },
    MIN: (args) => {
        let m: number | null = null;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                if (v === null) continue;
                let num: number;
                if (typeof v === 'number') num = v;
                else if (typeof v === 'boolean') num = v ? 1 : 0;
                else if (typeof v === 'string') {
                    if (v === '') continue;
                    const parsed = Number(v);
                    if (Number.isNaN(parsed)) continue;
                    num = parsed;
                } else continue;
                m = m === null ? num : Math.min(m, num);
            }
        }
        return m === null ? 0 : m;
    },
    MAX: (args) => {
        let m: number | null = null;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                if (v === null) continue;
                let num: number;
                if (typeof v === 'number') num = v;
                else if (typeof v === 'boolean') num = v ? 1 : 0;
                else if (typeof v === 'string') {
                    if (v === '') continue;
                    const parsed = Number(v);
                    if (Number.isNaN(parsed)) continue;
                    num = parsed;
                } else continue;
                m = m === null ? num : Math.max(m, num);
            }
        }
        return m === null ? 0 : m;
    },
    // COUNT counts numeric values only (including numeric strings when passed
    // as direct scalar args — but not when they come from ranges, mirroring
    // Excel's behaviour).
    COUNT: (args) => {
        let n = 0;
        for (const a of args) {
            const flat = flatten(a);
            const fromRange = Array.isArray(a);
            for (const v of flat) {
                if (isFormulaError(v)) continue;
                if (typeof v === 'number') n++;
                else if (!fromRange && typeof v === 'string' && v !== '') {
                    const parsed = Number(v);
                    if (!Number.isNaN(parsed)) n++;
                }
            }
        }
        return n;
    },
    // COUNTA counts non-empty cells (anything that isn't null/empty-string).
    COUNTA: (args) => {
        let n = 0;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) { n++; continue; } // errors count as non-empty
                if (v === null) continue;
                if (typeof v === 'string' && v === '') continue;
                n++;
            }
        }
        return n;
    },
    IF: (args) => {
        if (args.length < 2) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        const cond = toBool(args[0] as FormulaValue);
        if (isFormulaError(cond)) return cond;
        if (cond) return args[1];
        return args.length >= 3 ? args[2] : false;
    },
    AND: (args) => {
        let seen = false;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                if (v === null) continue; // skip empties
                const b = toBool(v as FormulaValue);
                if (isFormulaError(b)) return b;
                if (!b) return false;
                seen = true;
            }
        }
        return seen;
    },
    OR: (args) => {
        let seen = false;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (isFormulaError(v)) return v;
                if (v === null) continue;
                const b = toBool(v as FormulaValue);
                if (isFormulaError(b)) return b;
                if (b) return true;
                seen = true;
            }
        }
        return seen ? false : makeError('#VALUE!');
    },
    NOT: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        const b = toBool(args[0] as FormulaValue);
        if (isFormulaError(b)) return b;
        return !b;
    },
    // IFERROR(value, fallback): if value is any error, return fallback;
    // otherwise pass value through. Since we eagerly evaluate both args,
    // this differs from Excel's short-circuit behaviour only if `fallback`
    // itself errors — then the fallback error still surfaces.
    IFERROR: (args) => {
        if (args.length !== 2) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[1];
        // A ref to a cached-error cell lands here as a plain '#DIV/0!' string;
        // coerce defensively.
        if (typeof args[0] === 'string' && args[0].startsWith('#')) return args[1];
        return args[0];
    },
    // IFNA(value, fallback): narrower IFERROR; only swallows #N/A.
    IFNA: (args) => {
        if (args.length !== 2) return makeError('#VALUE!');
        if (isFormulaError(args[0]) && args[0].code === '#N/A') return args[1];
        if (typeof args[0] === 'string' && args[0].toUpperCase().startsWith('#N/A')) return args[1];
        return args[0];
    },
    // IFS(cond1, r1, cond2, r2, …): returns the first r_i whose cond_i is
    // truthy. Errors propagate. Returns #N/A if no condition is true.
    IFS: (args) => {
        if (args.length < 2 || args.length % 2 !== 0) return makeError('#VALUE!');
        for (let i = 0; i < args.length; i += 2) {
            if (isFormulaError(args[i])) return args[i];
            const b = toBool(args[i] as FormulaValue);
            if (isFormulaError(b)) return b;
            if (b) return args[i + 1];
        }
        return makeError('#N/A');
    },
    // SWITCH(expr, match1, r1, match2, r2, …, default?):
    // Compare expr to each match_i; return r_i for the first hit. A lone
    // trailing argument (odd count after the initial expr) is the default.
    // Returns #N/A when no match + no default.
    SWITCH: (args) => {
        if (args.length < 3) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        const expr = args[0] as FormulaValue;
        const tail = args.slice(1);
        const hasDefault = tail.length % 2 === 1;
        const pairsEnd = hasDefault ? tail.length - 1 : tail.length;
        for (let i = 0; i < pairsEnd; i += 2) {
            if (isFormulaError(tail[i])) return tail[i];
            if (equalValues(expr, tail[i] as FormulaValue)) return tail[i + 1];
        }
        return hasDefault ? tail[tail.length - 1] : makeError('#N/A');
    },
    // ── Math ─────────────────────────────────────────────────────────────
    ROUND: (args) => roundToDigits(args, 'round'),
    ROUNDUP: (args) => roundToDigits(args, 'up'),
    ROUNDDOWN: (args) => roundToDigits(args, 'down'),
    ABS: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const n = toNumber(args[0] as FormulaValue);
        if (isFormulaError(n)) return n;
        return Math.abs(n);
    },
    SQRT: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const n = toNumber(args[0] as FormulaValue);
        if (isFormulaError(n)) return n;
        if (n < 0) return makeError('#NUM!');
        return Math.sqrt(n);
    },
    POWER: (args) => {
        if (args.length !== 2) return makeError('#VALUE!');
        const b = toNumber(args[0] as FormulaValue);
        const e = toNumber(args[1] as FormulaValue);
        if (isFormulaError(b)) return b;
        if (isFormulaError(e)) return e;
        const r = Math.pow(b, e);
        if (!Number.isFinite(r) || Number.isNaN(r)) return makeError('#NUM!');
        return r;
    },
    MOD: (args) => {
        if (args.length !== 2) return makeError('#VALUE!');
        const n = toNumber(args[0] as FormulaValue);
        const d = toNumber(args[1] as FormulaValue);
        if (isFormulaError(n)) return n;
        if (isFormulaError(d)) return d;
        if (d === 0) return makeError('#DIV/0!');
        // Excel MOD takes the sign of the divisor; JS `%` takes the sign of
        // the dividend. Adjust.
        return n - Math.floor(n / d) * d;
    },
    INT: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const n = toNumber(args[0] as FormulaValue);
        if (isFormulaError(n)) return n;
        return Math.floor(n);
    },
    // ── Statistical with criteria ────────────────────────────────────────
    SUMIF: (args) => {
        if (args.length < 2 || args.length > 3) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        if (isFormulaError(args[1])) return args[1];
        const range = args[0];
        const crit = args[1] as FormulaValue;
        const sumRange = args.length === 3 ? args[2] : range;
        if (isFormulaError(sumRange)) return sumRange;
        const pred = buildPredicate(crit);
        if (isFormulaError(pred)) return pred;
        const left = Array.isArray(range) ? range : [range as FormulaValue];
        const right = Array.isArray(sumRange) ? sumRange : [sumRange as FormulaValue];
        let total = 0;
        for (let i = 0; i < left.length; i++) {
            const probe = left[i];
            if (isFormulaError(probe)) return probe;
            if (!pred(probe)) continue;
            const s = right[i];
            if (s === undefined) continue;
            if (isFormulaError(s)) return s;
            if (s === null) continue;
            if (typeof s === 'number') { total += s; continue; }
            if (typeof s === 'string') {
                if (s === '') continue;
                const n = Number(s);
                if (!Number.isNaN(n)) total += n;
                continue;
            }
            if (typeof s === 'boolean') { total += s ? 1 : 0; continue; }
        }
        return total;
    },
    COUNTIF: (args) => {
        if (args.length !== 2) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        if (isFormulaError(args[1])) return args[1];
        const pred = buildPredicate(args[1] as FormulaValue);
        if (isFormulaError(pred)) return pred;
        let count = 0;
        const range = args[0];
        const flat = Array.isArray(range) ? range : [range as FormulaValue];
        for (const v of flat) {
            if (isFormulaError(v)) continue;
            if (pred(v)) count++;
        }
        return count;
    },
    // ── Text ─────────────────────────────────────────────────────────────
    LEFT: (args) => {
        if (args.length < 1 || args.length > 2) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        const nArg = args.length >= 2 ? toNumber(args[1] as FormulaValue) : 1;
        if (isFormulaError(nArg)) return nArg;
        if (nArg < 0) return makeError('#VALUE!');
        return s.slice(0, Math.floor(nArg));
    },
    RIGHT: (args) => {
        if (args.length < 1 || args.length > 2) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        const nArg = args.length >= 2 ? toNumber(args[1] as FormulaValue) : 1;
        if (isFormulaError(nArg)) return nArg;
        if (nArg < 0) return makeError('#VALUE!');
        const n = Math.floor(nArg);
        return n === 0 ? '' : s.slice(-n);
    },
    MID: (args) => {
        if (args.length !== 3) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        const startArg = toNumber(args[1] as FormulaValue); if (isFormulaError(startArg)) return startArg;
        const nArg = toNumber(args[2] as FormulaValue); if (isFormulaError(nArg)) return nArg;
        if (startArg < 1 || nArg < 0) return makeError('#VALUE!');
        const start = Math.floor(startArg) - 1;
        const n = Math.floor(nArg);
        return s.slice(start, start + n);
    },
    LEN: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        return s.length;
    },
    TRIM: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        // Excel TRIM collapses internal runs of spaces to a single space and
        // strips leading/trailing.
        return s.replace(/^ +| +$/g, '').replace(/ +/g, ' ');
    },
    UPPER: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        return s.toUpperCase();
    },
    LOWER: (args) => {
        if (args.length !== 1) return makeError('#VALUE!');
        const s = toStr(args[0] as FormulaValue); if (isFormulaError(s)) return s;
        return s.toLowerCase();
    },
    CONCATENATE: (args) => concatArgs(args),
    CONCAT: (args) => concatArgs(args),
    // ── Lookup ───────────────────────────────────────────────────────────
    // VLOOKUP(lookup, tableRange, colIndex, [exact]):
    //   The 4th arg is this evaluator's `exact` flag: TRUE / omitted =
    //   exact-match, FALSE = approximate. Approximate-match is punted to a
    //   future wave; when FALSE is passed we surface #VALUE! with a comment
    //   so the gap is visible to downstream consumers.
    VLOOKUP: (args) => {
        if (args.length < 3 || args.length > 4) return makeError('#VALUE!');
        if (isFormulaError(args[0])) return args[0];
        if (isFormulaError(args[1])) return args[1];
        if (isFormulaError(args[2])) return args[2];
        const lookup = args[0] as FormulaValue;
        const colIndexN = toNumber(args[2] as FormulaValue);
        if (isFormulaError(colIndexN)) return colIndexN;
        const colIndex = Math.floor(colIndexN);
        if (colIndex < 1) return makeError('#VALUE!');
        let exact = true;
        if (args.length === 4) {
            if (isFormulaError(args[3])) return args[3];
            const b = toBool(args[3] as FormulaValue);
            if (isFormulaError(b)) return b;
            exact = b;
        }
        if (!exact) {
            // Approximate match not implemented. TODO (future wave):
            // approximate-match VLOOKUP requires a binary search over a
            // sorted first column plus `<=` semantics.
            return makeError('#VALUE!');
        }
        // Pull the 2-D row shape attached by the 'range' AST node.
        const rangeVal = args[1];
        const rows = Array.isArray(rangeVal)
            ? (rangeVal as unknown as { __rows?: FormulaValue[][] }).__rows
            : undefined;
        if (!rows) return makeError('#VALUE!');
        if (!Array.isArray(rows) || rows.length === 0) return makeError('#N/A');
        if (rows[0].length < colIndex) return makeError('#REF!');
        for (const r of rows) {
            const probe = r[0];
            if (probe === undefined) continue;
            if (isFormulaError(probe)) return probe;
            if (equalValues(lookup, probe)) {
                const cell = r[colIndex - 1];
                return cell === undefined ? makeError('#N/A') : cell;
            }
        }
        return makeError('#N/A');
    },
};

function concatArgs(args: EvalResult[]): EvalResult {
    let out = '';
    for (const a of args) {
        for (const v of flatten(a)) {
            if (isFormulaError(v)) return v;
            if (v === null) continue;
            if (typeof v === 'string') { out += v; continue; }
            if (typeof v === 'number') { out += cleanNumberString(v); continue; }
            if (typeof v === 'boolean') { out += v ? 'TRUE' : 'FALSE'; continue; }
        }
    }
    return out;
}

function roundToDigits(args: EvalResult[], mode: 'round' | 'up' | 'down'): EvalResult {
    if (args.length !== 2) return makeError('#VALUE!');
    const n = toNumber(args[0] as FormulaValue); if (isFormulaError(n)) return n;
    const dArg = toNumber(args[1] as FormulaValue); if (isFormulaError(dArg)) return dArg;
    const d = Math.floor(dArg);
    const f = Math.pow(10, d);
    if (!Number.isFinite(f) || f === 0) return makeError('#NUM!');
    let r: number;
    if (mode === 'round') r = Math.round(n * f) / f;
    else if (mode === 'up') r = (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f;
    else r = (n >= 0 ? Math.floor(n * f) : Math.ceil(n * f)) / f;
    return r;
}

// Excel's `=` on scalars: case-insensitive string equality, strict numeric /
// boolean equality, null treated as 0 / ''.
function equalValues(a: FormulaValue, b: FormulaValue): boolean {
    if (a === null && b === null) return true;
    if (a === null) return b === 0 || b === '';
    if (b === null) return a === 0 || a === '';
    if (typeof a === 'number' && typeof b === 'number') return a === b;
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
    if (typeof a === 'number' && typeof b === 'string') {
        const n = Number(b); return !Number.isNaN(n) && a === n;
    }
    if (typeof a === 'string' && typeof b === 'number') {
        const n = Number(a); return !Number.isNaN(n) && n === b;
    }
    return false;
}

// Compare two scalars using Excel's ordering: numbers < strings < booleans,
// and within each type the natural order (numeric, case-insensitive string,
// false < true). Returns -1/0/1 or a FormulaError when either side is an error.
function cmp(a: FormulaValue, b: FormulaValue): number | FormulaError {
    if (a === null) a = 0;
    if (b === null) b = 0;
    // Same type: natural comparison.
    if (typeof a === 'number' && typeof b === 'number') return a < b ? -1 : a > b ? 1 : 0;
    if (typeof a === 'string' && typeof b === 'string') {
        const la = a.toLowerCase(); const lb = b.toLowerCase();
        return la < lb ? -1 : la > lb ? 1 : 0;
    }
    if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;
    // Mixed: number < string < boolean.
    const rank = (v: FormulaValue) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
    const ra = rank(a); const rb = rank(b);
    return ra < rb ? -1 : ra > rb ? 1 : 0;
}

// Evaluator. The outer `depth` guard catches runaway recursion (mutually
// recursive formulas / deeply nested IFs). Cell refs out of the resolver's
// range surface as #REF! via the resolver's null / string path.
export function evalAst(ast: Ast, resolver: CellResolver, depth = 0): EvalResult {
    if (depth > 1024) return makeError('#REF!');
    switch (ast.kind) {
        case 'num': return ast.value;
        case 'str': return ast.value;
        case 'bool': return ast.value;
        case 'err': return makeError(ast.code);
        case 'ref': {
            const v = resolver(ast.col, ast.row);
            if (typeof v === 'string' && v.startsWith('#')) return makeError(parseErrorLiteral(v));
            return v;
        }
        case 'range': {
            const out: FormulaValue[] = [];
            const cols = ast.col2 - ast.col1 + 1;
            for (let r = ast.row1; r <= ast.row2; r++) {
                for (let c = ast.col1; c <= ast.col2; c++) {
                    const v = resolver(c, r);
                    if (typeof v === 'string' && v.startsWith('#')) return makeError(parseErrorLiteral(v));
                    out.push(v);
                }
            }
            // Attach the shaped 2-D form so VLOOKUP can see the row structure.
            // The attached property is hidden from standard array behaviour —
            // the flat array still iterates row-major.
            const shaped: FormulaValue[][] = [];
            for (let i = 0; i < out.length; i += cols) shaped.push(out.slice(i, i + cols));
            (out as unknown as { __rows: FormulaValue[][] }).__rows = shaped;
            return out;
        }
        case 'unary': {
            const v = evalAst(ast.arg, resolver, depth + 1);
            if (isFormulaError(v)) return v;
            const n = toNumber(v as FormulaValue);
            if (isFormulaError(n)) return n;
            return ast.op === '-' ? -n : n;
        }
        case 'postfix': {
            const v = evalAst(ast.arg, resolver, depth + 1);
            if (isFormulaError(v)) return v;
            const n = toNumber(v as FormulaValue);
            if (isFormulaError(n)) return n;
            return n / 100;
        }
        case 'bin': {
            const l = evalAst(ast.left, resolver, depth + 1);
            const r = evalAst(ast.right, resolver, depth + 1);
            if (isFormulaError(l)) return l;
            if (isFormulaError(r)) return r;
            switch (ast.op) {
                case '+': case '-': case '*': case '/': case '^': {
                    const ln = toNumber(l as FormulaValue);
                    const rn = toNumber(r as FormulaValue);
                    if (isFormulaError(ln)) return ln;
                    if (isFormulaError(rn)) return rn;
                    switch (ast.op) {
                        case '+': return ln + rn;
                        case '-': return ln - rn;
                        case '*': return ln * rn;
                        case '/': return rn === 0 ? makeError('#DIV/0!') : ln / rn;
                        case '^': {
                            const v = Math.pow(ln, rn);
                            if (!Number.isFinite(v) || Number.isNaN(v)) return makeError('#NUM!');
                            return v;
                        }
                    }
                    return makeError('#VALUE!');
                }
                case '&': {
                    const stringify = (v: FormulaValue): string => {
                        if (v === null) return '';
                        if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
                        if (typeof v === 'number') return cleanNumberString(v);
                        if (Array.isArray(v)) return stringify(v.length > 0 ? v[0] : null);
                        return String(v);
                    };
                    return stringify(l as FormulaValue) + stringify(r as FormulaValue);
                }
                case 'eq': case 'neq': case 'lt': case 'lte': case 'gt': case 'gte': {
                    const c = cmp(l as FormulaValue, r as FormulaValue);
                    if (isFormulaError(c)) return c;
                    switch (ast.op) {
                        case 'eq': return c === 0;
                        case 'neq': return c !== 0;
                        case 'lt': return c < 0;
                        case 'lte': return c <= 0;
                        case 'gt': return c > 0;
                        case 'gte': return c >= 0;
                    }
                    return makeError('#VALUE!');
                }
            }
            return makeError('#VALUE!');
        }
        case 'call': {
            const fn = FUNCTIONS[ast.name];
            if (!fn) return makeError('#NAME?');
            const args = ast.args.map((a) => evalAst(a, resolver, depth + 1));
            return fn(args);
        }
    }
}

// Top-level entry: parse + evaluate + convert FormulaError → sentinel string.
// Returns { value, kind } mirroring the Cell.kind enum that the renderer
// expects, so the caller can splat into the Cell record.
export function evaluateFormula(
    source: string,
    resolver: CellResolver,
): { value: string; kind: 'number' | 'string' | 'boolean' | 'error' } {
    let ast: Ast;
    try {
        ast = parseFormula(source);
    } catch {
        return { value: '#VALUE!', kind: 'error' };
    }
    const result = evalAst(ast, resolver);
    if (isFormulaError(result)) return { value: result.code, kind: 'error' };
    if (Array.isArray(result)) {
        // An unaggregated range reached the top — Excel would display the
        // first cell (implicit intersection) or spill. Return the first.
        const first = result.length > 0 ? result[0] : null;
        return formatScalar(first as FormulaValue);
    }
    return formatScalar(result);
}

function formatScalar(v: FormulaValue): { value: string; kind: 'number' | 'string' | 'boolean' | 'error' } {
    if (v === null) return { value: '', kind: 'string' };
    if (typeof v === 'number') {
        if (!Number.isFinite(v)) return { value: '#NUM!', kind: 'error' };
        // Number-format routing: write a clean 15-significant-digit string so
        // `0.1+0.2` lands as '0.3' and the renderer's formatNumber() sees the
        // same shape it does for a cached <v>0.3</v>.
        return { value: cleanNumberString(v), kind: 'number' };
    }
    if (typeof v === 'boolean') return { value: v ? 'TRUE' : 'FALSE', kind: 'boolean' };
    if (typeof v === 'string') {
        if (v.startsWith('#')) return { value: v, kind: 'error' };
        return { value: v, kind: 'string' };
    }
    return { value: '#VALUE!', kind: 'error' };
}

// ---------------------------------------------------------------------------
// Sheet wiring
// ---------------------------------------------------------------------------

// A parsed Sheet-like object (just enough fields from workbook-parser Cell /
// Sheet to drive the resolver). Kept structurally typed so this module stays
// decoupled from workbook-parser's larger type soup.
export interface SheetLike {
    rows: CellLike[][];
}
export interface CellLike {
    col: number;
    row: number;
    value: string;
    kind: 'string' | 'number' | 'boolean' | 'inlineStr' | 'error' | 'empty';
    formula: string | null;
}

// Build a resolver closure that reads the sheet's cell grid. Empty cells
// return null; non-empty string cells return their string; numeric cells
// return the parsed number; boolean cells return the boolean; errors pass
// through as '#<code>!' strings so the evaluator can short-circuit.
//
// Each row is a *packed* array of cells (sparse columns are collapsed out),
// so we scan for the matching `col` rather than indexing directly. For wide
// rows we build a per-sheet col-index cache on first access to keep the
// lookup O(1).
export function makeSheetResolver(sheet: SheetLike): CellResolver {
    const cache = new Map<number, Map<number, CellLike>>();
    function indexRow(row: number): Map<number, CellLike> | null {
        let m = cache.get(row);
        if (m) return m;
        const rowArr = sheet.rows[row];
        if (!rowArr) return null;
        m = new Map();
        for (const cell of rowArr) if (cell) m.set(cell.col, cell);
        cache.set(row, m);
        return m;
    }
    return (col: number, row: number): FormulaValue => {
        const m = indexRow(row);
        if (!m) return null;
        const cell = m.get(col);
        if (!cell) return null;
        if (cell.kind === 'empty') return null;
        if (cell.kind === 'number') {
            if (cell.value === '') return null;
            const n = Number(cell.value);
            return Number.isNaN(n) ? null : n;
        }
        if (cell.kind === 'boolean') return cell.value === 'TRUE' || cell.value === '1' || cell.value.toLowerCase() === 'true';
        if (cell.kind === 'error') return cell.value || '#VALUE!';
        // String / inlineStr — try numeric coercion? No, Excel keeps strings
        // as strings and lets the evaluator coerce in numeric contexts.
        return cell.value;
    };
}

// Walk every formula cell on the sheet and evaluate it in-place. Two-pass:
// pass 1 records the result into a parallel map so mid-walk cell refs read
// the pre-evaluation value (avoids order sensitivity for forward refs);
// pass 2 writes results back. Cells with no `formula` are untouched.
//
// For cells that already have a non-empty cached value, the cached value is
// kept (Excel's own computation is authoritative — we only fill in blanks).
// Flip `force` to re-evaluate every formula regardless of the cache.
export function evaluateSheetFormulas(sheet: SheetLike, opts?: { force?: boolean }): void {
    const force = opts?.force === true;
    const resolver = makeSheetResolver(sheet);
    const updates: { cell: CellLike; value: string; kind: CellLike['kind'] }[] = [];
    for (const row of sheet.rows) {
        if (!row) continue;
        for (const cell of row) {
            if (!cell || !cell.formula) continue;
            if (!force && cell.value !== '' && cell.value != null) continue;
            const r = evaluateFormula(cell.formula, resolver);
            // Map the evaluator's kind into the Cell.kind enum. 'number' and
            // 'boolean' and 'error' all match; 'string' rides in as 'string'
            // (not inlineStr — the cached value path doesn't have a shared-
            // string table entry to point at).
            updates.push({ cell, value: r.value, kind: r.kind });
        }
    }
    for (const u of updates) {
        u.cell.value = u.value;
        u.cell.kind = u.kind;
    }
}
