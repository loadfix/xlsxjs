// Minimal formula evaluator. POC — covers 10 Excel functions + arithmetic
// + cell-range resolution. Opt-in via Options.evaluateFormulas.
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
// Deliberate limitations (documented in TODO.md):
//   · A1 / absolute ($A$1) only — no sheet-qualified refs (Sheet2!A1),
//     no structured references ([Table1[col]]), no named ranges.
//   · No array formulas, no spill semantics — a cell range always reduces
//     to a flat list of cells.
//   · Only 10 functions (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, AND,
//     OR, NOT) plus arithmetic + comparison + basic string equality.
//   · Errors surface as the string "#ERROR!" — no #DIV/0!, #VALUE!,
//     #REF!, #NAME? differentiation.
//   · Circular refs abort with #ERROR! after a 1024-hop depth cap.

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

// Error sentinel. Any function / op that needs to signal an error returns
// this same token; the top-level evaluator converts it to a display string.
export const ERROR = Symbol('formula-error');
export type EvalResult = FormulaValue | typeof ERROR;

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
    | { kind: 'err'; value: string }
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
        if (t.kind === 'err') { this.pos++; return { kind: 'err', value: t.text }; }
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
            // Bare identifier: treat as #NAME?
            return { kind: 'err', value: '#NAME?' };
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
// Evaluator
// ---------------------------------------------------------------------------

// Coerce a single value to a number. null → 0, true → 1, false → 0, number
// passes through, string parses (empty → 0, non-numeric → ERROR).
function toNumber(v: FormulaValue | typeof ERROR): number | typeof ERROR {
    if (v === ERROR) return ERROR;
    if (v === null) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
        if (v === '') return 0;
        if (v.startsWith('#')) return ERROR;
        const n = Number(v);
        if (Number.isNaN(n)) return ERROR;
        return n;
    }
    // Arrays shouldn't reach here — callers should flatten first.
    return ERROR;
}

function toBool(v: FormulaValue | typeof ERROR): boolean | typeof ERROR {
    if (v === ERROR) return ERROR;
    if (v === null) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
        const u = v.toUpperCase();
        if (u === 'TRUE') return true;
        if (u === 'FALSE') return false;
        if (v.startsWith('#')) return ERROR;
        const n = Number(v);
        if (Number.isNaN(n)) return ERROR;
        return n !== 0;
    }
    return ERROR;
}

// Flatten an EvalResult to an array of scalars. Arrays flatten recursively.
function flatten(v: FormulaValue | typeof ERROR): (FormulaValue | typeof ERROR)[] {
    if (v === ERROR) return [ERROR];
    if (Array.isArray(v)) {
        const out: (FormulaValue | typeof ERROR)[] = [];
        for (const x of v) for (const y of flatten(x)) out.push(y);
        return out;
    }
    return [v];
}

// Built-in functions. Each receives the raw evaluated args (one per
// positional argument; ranges stay as arrays). Returns EvalResult.
type Fn = (args: EvalResult[]) => EvalResult;

const FUNCTIONS: Record<string, Fn> = {
    SUM: (args) => {
        let total = 0;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (v === ERROR) return ERROR;
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
                if (v === ERROR) return ERROR;
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
        if (n === 0) return ERROR; // #DIV/0!
        return total / n;
    },
    MIN: (args) => {
        let m: number | null = null;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (v === ERROR) return ERROR;
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
                if (v === ERROR) return ERROR;
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
                if (v === ERROR) continue;
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
                if (v === ERROR) { n++; continue; } // errors count as non-empty
                if (v === null) continue;
                if (typeof v === 'string' && v === '') continue;
                n++;
            }
        }
        return n;
    },
    IF: (args) => {
        if (args.length < 2) return ERROR;
        const cond = toBool(args[0] as FormulaValue);
        if (cond === ERROR) return ERROR;
        if (cond) return args[1];
        return args.length >= 3 ? args[2] : false;
    },
    AND: (args) => {
        let seen = false;
        for (const a of args) {
            for (const v of flatten(a)) {
                if (v === ERROR) return ERROR;
                if (v === null) continue; // skip empties
                const b = toBool(v as FormulaValue);
                if (b === ERROR) return ERROR;
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
                if (v === ERROR) return ERROR;
                if (v === null) continue;
                const b = toBool(v as FormulaValue);
                if (b === ERROR) return ERROR;
                if (b) return true;
                seen = true;
            }
        }
        return seen ? false : ERROR;
    },
    NOT: (args) => {
        if (args.length !== 1) return ERROR;
        const b = toBool(args[0] as FormulaValue);
        if (b === ERROR) return ERROR;
        return !b;
    },
};

// Compare two scalars using Excel's ordering: numbers < strings < booleans,
// and within each type the natural order (numeric, case-insensitive string,
// false < true). Returns -1/0/1 or ERROR when either side is an error.
function cmp(a: FormulaValue, b: FormulaValue): number | typeof ERROR {
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

export function evalAst(ast: Ast, resolver: CellResolver, depth = 0): EvalResult {
    if (depth > 1024) return ERROR;
    switch (ast.kind) {
        case 'num': return ast.value;
        case 'str': return ast.value;
        case 'bool': return ast.value;
        case 'err': return ERROR;
        case 'ref': {
            const v = resolver(ast.col, ast.row);
            if (typeof v === 'string' && v.startsWith('#')) return ERROR;
            return v;
        }
        case 'range': {
            const out: FormulaValue[] = [];
            for (let r = ast.row1; r <= ast.row2; r++) {
                for (let c = ast.col1; c <= ast.col2; c++) {
                    const v = resolver(c, r);
                    if (typeof v === 'string' && v.startsWith('#')) return ERROR;
                    out.push(v);
                }
            }
            return out;
        }
        case 'unary': {
            const v = evalAst(ast.arg, resolver, depth + 1);
            if (v === ERROR) return ERROR;
            const n = toNumber(v as FormulaValue);
            if (n === ERROR) return ERROR;
            return ast.op === '-' ? -n : n;
        }
        case 'postfix': {
            const v = evalAst(ast.arg, resolver, depth + 1);
            if (v === ERROR) return ERROR;
            const n = toNumber(v as FormulaValue);
            if (n === ERROR) return ERROR;
            return n / 100;
        }
        case 'bin': {
            const l = evalAst(ast.left, resolver, depth + 1);
            const r = evalAst(ast.right, resolver, depth + 1);
            if (l === ERROR || r === ERROR) return ERROR;
            switch (ast.op) {
                case '+': case '-': case '*': case '/': case '^': {
                    const ln = toNumber(l as FormulaValue);
                    const rn = toNumber(r as FormulaValue);
                    if (ln === ERROR || rn === ERROR) return ERROR;
                    switch (ast.op) {
                        case '+': return ln + rn;
                        case '-': return ln - rn;
                        case '*': return ln * rn;
                        case '/': return rn === 0 ? ERROR : ln / rn;
                        case '^': return Math.pow(ln, rn);
                    }
                    return ERROR;
                }
                case '&': {
                    const stringify = (v: FormulaValue): string => {
                        if (v === null) return '';
                        if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
                        if (Array.isArray(v)) return stringify(v.length > 0 ? v[0] : null);
                        return String(v);
                    };
                    return stringify(l as FormulaValue) + stringify(r as FormulaValue);
                }
                case 'eq': case 'neq': case 'lt': case 'lte': case 'gt': case 'gte': {
                    const c = cmp(l as FormulaValue, r as FormulaValue);
                    if (c === ERROR) return ERROR;
                    switch (ast.op) {
                        case 'eq': return c === 0;
                        case 'neq': return c !== 0;
                        case 'lt': return c < 0;
                        case 'lte': return c <= 0;
                        case 'gt': return c > 0;
                        case 'gte': return c >= 0;
                    }
                    return ERROR;
                }
            }
            return ERROR;
        }
        case 'call': {
            const fn = FUNCTIONS[ast.name];
            if (!fn) return ERROR; // #NAME?
            const args = ast.args.map((a) => evalAst(a, resolver, depth + 1));
            // Short-circuit for IF: caller evaluated both branches already;
            // this is less efficient than true short-circuit but preserves
            // error semantics (a #DIV/0! in the untaken branch still surfaces
            // in Excel when the cell is referenced directly, but not through
            // IF — accept the gap for the POC).
            return fn(args);
        }
    }
}

// Top-level entry: parse + evaluate + convert ERROR → '#ERROR!' string.
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
        return { value: '#ERROR!', kind: 'error' };
    }
    const result = evalAst(ast, resolver);
    if (result === ERROR) return { value: '#ERROR!', kind: 'error' };
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
        if (!Number.isFinite(v)) return { value: '#ERROR!', kind: 'error' };
        return { value: String(v), kind: 'number' };
    }
    if (typeof v === 'boolean') return { value: v ? 'TRUE' : 'FALSE', kind: 'boolean' };
    if (typeof v === 'string') {
        if (v.startsWith('#')) return { value: v, kind: 'error' };
        return { value: v, kind: 'string' };
    }
    return { value: '#ERROR!', kind: 'error' };
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
// through as '#ERROR!' strings so the evaluator can short-circuit.
export function makeSheetResolver(sheet: SheetLike): CellResolver {
    return (col: number, row: number): FormulaValue => {
        const rowArr = sheet.rows[row];
        if (!rowArr) return null;
        const cell = rowArr[col];
        if (!cell) return null;
        if (cell.kind === 'empty') return null;
        if (cell.kind === 'number') {
            if (cell.value === '') return null;
            const n = Number(cell.value);
            return Number.isNaN(n) ? null : n;
        }
        if (cell.kind === 'boolean') return cell.value === 'TRUE' || cell.value === '1' || cell.value.toLowerCase() === 'true';
        if (cell.kind === 'error') return cell.value || '#ERROR!';
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
