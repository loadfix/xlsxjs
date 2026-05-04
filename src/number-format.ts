// Tiny Excel number-format formatter. Given a raw cell value (as a string,
// since that's what the XML carries) and an Excel format code, produce the
// display string.
//
// Coverage is pragmatic rather than exhaustive. The goal is to make fixtures
// written by Excel / LibreOffice / python-xlsx render sensibly:
//   - "General"                    → number or integer, no thousand separator
//   - Percent codes                → value × 100 + "%"
//   - Currency codes ("$"...)      → literal currency + digits + thousands
//   - Thousands separator (#,##0)  → insert commas in integer part
//   - Decimals (0.00, 0.0%)        → fixed-digit rounding
//   - Dates (yyyy, yy, m, d, mmm)  → from Excel serial to Y/M/D components
//   - Times (h, hh, mm, ss, AM/PM) → from fractional day component
//   - Elapsed time ([h], [mm], [ss]) → accumulated units past the modulo
//   - Accounting padding (_( _) _-) → one space per _<char> sequence
//   - Fill character (*<char>)      → stripped (can't measure column width)
//   - Locale currency ([$€-2])      → symbol preserved, locale id dropped
//   - Colour modifiers ([Red], [Color 14]) → return the colour on FormatResult
//   - Conditional sections ([>100]) → predicate-based section dispatch
//   - Fractions (# ?/?, # ??/??)    → best-fit fraction via continued fractions
//   - Scientific (0.00E+00)         → mantissa × 10^exponent; engineering too
//   - Text ("@")                   → the raw cell text unchanged
//
// Excel's format-code grammar supports positive;negative;zero;text sections.
// We honour the split. When the first bracket of a section is a predicate
// (e.g. `[>100]`), the section is conditional and only fires when the
// predicate matches the value; otherwise positional (positive/negative/zero)
// semantics apply.
//
// Compromises worth flagging for downstream consumers:
//   - `*<char>` is a "fill the remaining column width with <char>" marker in
//     Excel. A web renderer can't measure remaining cell width without a
//     layout pass, so we silently drop the marker rather than emit a bogus
//     single repeat character. Accounting columns that rely on the fill to
//     visually line up their currency symbol and digits will look slightly
//     tighter than in Excel, but the digits themselves are correct.

import { indexedColor } from './theme';

export interface FormatResult {
    text: string;
    // When true, right-align the cell (numbers, dates, times). When false,
    // treat like a string (text format, @).
    numeric: boolean;
    // Optional colour modifier from the format code's bracket prefix
    // (e.g. `[Red]`, `[Blue]`, `[Color 14]`). The renderer applies this
    // as the cell's text colour, overriding any xf-level colour. Null
    // when the matched section carries no colour modifier.
    color: string | null;
}

// Options passed through from the parsed workbook. `date1904` flips the
// date-serial epoch from 1899-12-30 (PC convention with the historical leap
// bug) to 1904-01-01 (Mac Office pre-2011 convention with no leap bug).
export interface FormatOptions {
    date1904?: boolean;
}

const NUMERIC_SECTION_INDEX = {
    positive: 0,
    negative: 1,
    zero: 2,
} as const;

// Named colours per ECMA-376 §18.8.30. Keys are lower-cased on lookup.
const NAMED_COLORS: Record<string, string> = {
    black: '#000000',
    blue: '#0000ff',
    cyan: '#00ffff',
    green: '#00ff00',
    magenta: '#ff00ff',
    red: '#ff0000',
    white: '#ffffff',
    yellow: '#ffff00',
};

export function formatNumber(value: string, formatCode: string, options?: FormatOptions): FormatResult {
    if (formatCode === '' || formatCode.toLowerCase() === 'general') {
        return formatGeneral(value);
    }

    // @ sections are "text" and should echo the input. If that's the only
    // section, treat as a non-numeric result.
    if (formatCode.trim() === '@') {
        return { text: value, numeric: false, color: null };
    }

    const num = Number(value);
    if (!Number.isFinite(num)) {
        // Not actually a number — fall back to raw value. Excel itself does
        // this for error cells.
        return { text: value, numeric: false, color: null };
    }

    const sections = splitSections(formatCode);
    const section = selectSection(sections, num);
    // [$<symbol>-<localeHex>] is a locale-tagged currency marker. Pull out the
    // symbol (if any) BEFORE stripSquareBracketModifiers eats the whole thing,
    // then splice it back in at the same spot as a quoted literal so the rest
    // of the pipeline treats it as an ordinary literal prefix.
    const { code: sectionWithSymbol } = extractLocaleCurrency(section.code);
    // Pull colour modifier out of the section's bracket prefix, if any.
    const color = extractColorModifier(sectionWithSymbol);
    const cleaned = stripSquareBracketModifiers(sectionWithSymbol);
    // Negative numbers: the negative section format itself typically includes
    // the sign or parentheses. If we fall back to the positive section, prefix
    // a minus.
    const magnitude = Math.abs(num);
    const body = applyFormat(cleaned, magnitude, options);
    const text = (num < 0 && section.prefixMinus) ? '-' + body : body;
    return { text, numeric: true, color };
}

// Section dispatch. Returns the chosen section code and a flag: when the
// value is negative but the selected section is the positive/default one,
// the caller should prefix a minus sign.
function selectSection(sections: string[], num: number): { code: string; prefixMinus: boolean } {
    // Conditional section semantics. If any section has a predicate prefix,
    // dispatch by predicate. Fall back to positional.
    const predicates = sections.map(parseSectionPredicate);
    const anyConditional = predicates.some((p) => p.predicate !== null);
    if (anyConditional) {
        // Walk sections in order; pick the first whose predicate matches.
        // An un-predicated section is the "else" fallback.
        let fallback: string | null = null;
        for (let i = 0; i < sections.length; i++) {
            const p = predicates[i];
            if (p.predicate === null) {
                // Only the first non-conditional section acts as fallback;
                // later ones are text sections that don't fire on numbers.
                if (fallback === null && !isTextSection(sections[i])) fallback = sections[i];
                continue;
            }
            if (evaluatePredicate(p.predicate, num)) {
                return { code: p.body, prefixMinus: false };
            }
        }
        if (fallback !== null) return { code: fallback, prefixMinus: false };
        // No match, no fallback — render the raw number via first section.
        return { code: sections[0], prefixMinus: false };
    }
    // Positional semantics.
    let sectionIndex: number;
    if (num > 0) sectionIndex = NUMERIC_SECTION_INDEX.positive;
    else if (num < 0) sectionIndex = sections[NUMERIC_SECTION_INDEX.negative] ? NUMERIC_SECTION_INDEX.negative : NUMERIC_SECTION_INDEX.positive;
    else sectionIndex = sections[NUMERIC_SECTION_INDEX.zero] ? NUMERIC_SECTION_INDEX.zero : NUMERIC_SECTION_INDEX.positive;

    const code = sections[sectionIndex] ?? sections[0] ?? '';
    const prefixMinus = num < 0 && sectionIndex === NUMERIC_SECTION_INDEX.positive;
    return { code, prefixMinus };
}

// Does this section contain only text-format tokens (no digit placeholders,
// no date tokens)? Used to decide whether an unpredicated section at index 3
// should be considered as a numeric fallback.
function isTextSection(code: string): boolean {
    const stripped = stripQuotedLiterals(code);
    return /@/.test(stripped) && !/[0#?]/.test(stripped);
}

type Predicate = { op: string; value: number };

// Parse a leading predicate prefix like `[>100]` off a section body. Returns
// the remaining body plus the parsed predicate. Non-predicate brackets
// ([Red], [$€-2], [h]) are left at the front of the body and predicate is null.
function parseSectionPredicate(code: string): { predicate: Predicate | null; body: string } {
    const m = /^\[([<>=]+)(-?\d+(?:\.\d+)?)\]/.exec(code);
    if (!m) return { predicate: null, body: code };
    const op = m[1];
    // Validate operator — only the six comparison ops are predicates;
    // anything else (e.g. unexpected punctuation) is left as a literal.
    if (!['>', '<', '=', '>=', '<=', '<>'].includes(op)) {
        return { predicate: null, body: code };
    }
    const value = Number(m[2]);
    if (!Number.isFinite(value)) return { predicate: null, body: code };
    return { predicate: { op, value }, body: code.slice(m[0].length) };
}

function evaluatePredicate(p: Predicate, num: number): boolean {
    switch (p.op) {
        case '>': return num > p.value;
        case '<': return num < p.value;
        case '=': return num === p.value;
        case '>=': return num >= p.value;
        case '<=': return num <= p.value;
        case '<>': return num !== p.value;
    }
    return false;
}

// Find any colour-modifier bracket (`[Red]`, `[Color 14]`) and return the
// resolved hex. Non-colour brackets are ignored. We only look at brackets
// whose content is a recognised colour name or `Color N`; everything else
// (locale tags, elapsed time, predicates) passes through untouched.
function extractColorModifier(code: string): string | null {
    let match: RegExpExecArray | null;
    const re = /\[([^\]]*)\]/g;
    while ((match = re.exec(code)) !== null) {
        const inner = match[1].trim();
        const named = NAMED_COLORS[inner.toLowerCase()];
        if (named) return named;
        const colorM = /^color\s*(\d+)$/i.exec(inner);
        if (colorM) {
            const idx = Number(colorM[1]);
            // ECMA indexes [Color N] as 1-based into the legacy palette.
            // indexedColor() is 0-based, so subtract one before looking up.
            const hex = indexedColor(idx - 1);
            if (hex) return hex;
        }
    }
    return null;
}

// Find any [$<symbol>-<locale>] modifier and rewrite it in place as a quoted
// literal carrying only the symbol. `[$-409]` (locale-only, no symbol) is
// rewritten to empty. Non-locale brackets ([Red], [>100], [h]) are left alone.
function extractLocaleCurrency(code: string): { code: string } {
    return {
        code: code.replace(/\[\$([^\]]*)\]/g, (_, inner: string) => {
            // Inner layout: <symbol>-<localeHex>, or just <symbol>, or -<localeHex>.
            const dashIdx = inner.indexOf('-');
            const symbol = dashIdx >= 0 ? inner.slice(0, dashIdx) : inner;
            if (!symbol) return '';
            // Quote it so a `$` inside doesn't trip up downstream parsing, and
            // so a stray number placeholder inside the symbol (unlikely but
            // possible, e.g. "¥") renders verbatim.
            return `"${symbol.replace(/"/g, '')}"`;
        }),
    };
}

function splitSections(code: string): string[] {
    // Split on ; but not on \; or "literal ;".
    const out: string[] = [];
    let current = '';
    let inQuote = false;
    for (let i = 0; i < code.length; i++) {
        const ch = code[i];
        if (ch === '"') {
            inQuote = !inQuote;
            current += ch;
            continue;
        }
        if (ch === '\\' && i + 1 < code.length) {
            current += ch + code[i + 1];
            i++;
            continue;
        }
        if (ch === ';' && !inQuote) {
            out.push(current);
            current = '';
            continue;
        }
        current += ch;
    }
    out.push(current);
    return out;
}

function stripSquareBracketModifiers(code: string): string {
    // Drop [Red], [Blue], [>100] etc. Interior of brackets may not contain a
    // closing bracket, so a lazy match is fine. Elapsed-time markers
    // ([h], [hh], [m], [mm], [s], [ss]) are preserved — formatDateTime
    // reads them directly to accumulate units past the modulo boundary.
    return code.replace(/\[([^\]]*)\]/g, (match, inner: string) => {
        if (/^(hh?|mm?|ss?)$/i.test(inner)) return match;
        return '';
    });
}

function applyFormat(code: string, value: number, options?: FormatOptions): string {
    // Fast path: a "General" literal inside a section.
    if (code.trim().toLowerCase() === 'general') {
        return formatGeneralNumber(value);
    }
    // Elapsed-time markers forcibly put us on the date/time path even if the
    // non-bracketed body has no date tokens (e.g. "[s]" alone).
    if (/\[(hh?|mm?|ss?)\]/.test(code)) {
        return formatDateTime(code, value, options);
    }
    // Fraction pattern: `# ?/?`, `# ??/??`, etc. Take priority over the
    // generic numeric path so the `?/?` placeholders aren't confused with
    // digit placeholders above.
    if (isFractionCode(code)) {
        return formatFraction(code, value);
    }
    // Scientific / engineering: explicit `E+` or `E-` exponent marker.
    if (/[eE][+-]/.test(stripQuotedLiterals(code))) {
        return formatScientific(code, value);
    }
    // If any date/time token is present, treat the code as a date/time format.
    if (/[yMdhsmAP]/.test(stripQuotedLiterals(code))) {
        // Heuristic: "m" is month when adjacent to y/d, minute when adjacent
        // to h/s. The dedicated date path handles that.
        if (/[yMdAPhs]/.test(stripQuotedLiterals(code))) {
            return formatDateTime(code, value, options);
        }
    }
    return formatNumeric(code, value);
}

function stripQuotedLiterals(code: string): string {
    // For token detection only: remove \x escapes and "quoted" runs so we
    // don't misread a literal "d" inside "Date" as a day token.
    return code
        .replace(/\\./g, '')
        .replace(/"([^"]*)"/g, '')
        .replace(/\[[^\]]*\]/g, '');
}

// ── Numeric formatting ────────────────────────────────────────────────────

function formatNumeric(code: string, value: number): string {
    // Detect percent: a literal % in the code (outside quotes). Multiply the
    // value by 100^n where n is the count of percent signs.
    const percentCount = (stripQuotedLiterals(code).match(/%/g) ?? []).length;
    let working = value;
    for (let i = 0; i < percentCount; i++) working *= 100;

    // Decide decimals. Look at digits after the decimal point in the code.
    const dotIdx = firstUnquotedIndexOf(code, '.');
    let decimals = 0;
    if (dotIdx >= 0) {
        for (let i = dotIdx + 1; i < code.length; i++) {
            const c = code[i];
            if (c === '0' || c === '#' || c === '?') decimals++;
            else if (c === '%' || c === ',' || c === ' ') continue;
            else break;
        }
    }

    const thousands = /#,##0|0,000/.test(stripQuotedLiterals(code));
    const fixed = Math.abs(working).toFixed(decimals);
    const [intPart, fracPart] = fixed.split('.');
    const withSep = thousands ? insertThousands(intPart) : intPart;
    const body = fracPart != null ? `${withSep}.${fracPart}` : withSep;
    const signed = working < 0 ? '-' + body : body;

    // Replace numeric placeholders with the computed body, then render the
    // remaining literal characters (., $, etc.) as-is. We walk the code and
    // substitute the first contiguous run of placeholders with `body`.
    return renderLiteralsAroundNumber(code, signed);
}

function firstUnquotedIndexOf(code: string, ch: string): number {
    let inQuote = false;
    for (let i = 0; i < code.length; i++) {
        const c = code[i];
        if (c === '"') { inQuote = !inQuote; continue; }
        if (c === '\\') { i++; continue; }
        if (c === ch && !inQuote) return i;
    }
    return -1;
}

function insertThousands(intPart: string): string {
    return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function renderLiteralsAroundNumber(code: string, numberText: string): string {
    // Walk the code producing either the number (at the first number run)
    // or the literal character. This keeps "$1,234.00" and "1,234.00 %"
    // working without needing a full parser.
    // Sections that contain NO numeric placeholders are pure text literals
    // (`"big"`, `"none"`) and render without injecting the number — that's
    // how Excel's conditional sections ([>100]"big") work too.
    const hasPlaceholder = /[0#?]/.test(stripQuotedLiterals(code));
    let out = '';
    let inserted = false;
    let i = 0;
    while (i < code.length) {
        const c = code[i];
        if (c === '"') {
            // Quoted literal: copy content without quotes.
            i++;
            while (i < code.length && code[i] !== '"') { out += code[i]; i++; }
            i++; // skip closing "
            continue;
        }
        if (c === '\\') {
            if (i + 1 < code.length) out += code[i + 1];
            i += 2;
            continue;
        }
        if (c === '_') {
            // `_<char>` = leave a space equal to the width of <char>. We emit
            // a single ASCII space as a reasonable proxy; this powers
            // accounting-style `_( _) _-` alignment even though we can't match
            // the glyph width exactly.
            if (i + 1 < code.length) out += ' ';
            i += 2;
            continue;
        }
        if (c === '*') {
            // `*<char>` = fill the remaining column width with <char>. Without
            // a layout pass we can't know how much to emit, so drop both
            // characters. See the file-header compromise note.
            i += 2;
            continue;
        }
        if (c === '0' || c === '#' || c === '?' || c === '.' || c === ',') {
            if (!inserted) {
                out += numberText;
                inserted = true;
            }
            // Skip the rest of this contiguous number run.
            while (i < code.length && '0#?.,'.includes(code[i])) i++;
            continue;
        }
        if (c === '%') {
            out += '%';
            i++;
            continue;
        }
        out += c;
        i++;
    }
    if (!inserted && hasPlaceholder) out = numberText + out;
    return out;
}

function formatGeneral(value: string): FormatResult {
    const n = Number(value);
    if (!Number.isFinite(n)) return { text: value, numeric: false, color: null };
    return { text: formatGeneralNumber(n), numeric: true, color: null };
}

function formatGeneralNumber(n: number): string {
    // "General" trims trailing zeros and falls back to scientific for very
    // large/small magnitudes. toString() is close enough for the common case.
    return n.toString();
}

// ── Fractions ──────────────────────────────────────────────────────────────

// Detect `# ?/?` or `# ??/??` style fraction markers. Needs a digit-or-#
// placeholder, whitespace, then `?/?` (the fraction marker). `?` is the
// canonical placeholder for the fraction numerator / denominator; Excel
// allows fixed-denominator variants like `# ?/8` too.
function isFractionCode(code: string): boolean {
    return /[0#]\s+\?+\/(\?+|\d+)/.test(code);
}

function formatFraction(code: string, value: number): string {
    // Determine the max denominator from the `?` count on the denominator
    // side of the slash. `# ?/?` → 10^1 = 10; `# ??/??` → 10^2 = 100.
    // Literal-denominator variants like `# ?/8` pin the denominator to 8.
    const m = /(\?+)\/(\?+|\d+)/.exec(code);
    if (!m) return String(value); // shouldn't happen given isFractionCode
    const denomRaw = m[2];
    const fixedDenom = /^\d+$/.test(denomRaw) ? Number(denomRaw) : null;
    const denomPlaces = denomRaw.length;
    const maxDenom = fixedDenom ?? Math.pow(10, denomPlaces);

    const sign = value < 0 ? '-' : '';
    const abs = Math.abs(value);
    const whole = Math.floor(abs);
    const frac = abs - whole;

    // Best-fit fraction via continued fractions (Stern-Brocot style). The
    // iterative form converges on the closest rational with denom ≤ maxDenom.
    // See https://en.wikipedia.org/wiki/Continued_fraction#Best_rational_approximations.
    // Fixed-denominator variants (`# ?/8`) use the exact denom and round
    // the numerator to the nearest integer.
    let num: number, den: number;
    if (fixedDenom !== null && fixedDenom > 0) {
        den = fixedDenom;
        num = Math.round(frac * fixedDenom);
    } else {
        ({ num, den } = bestFraction(frac, maxDenom));
    }

    // If the numerator rounded up to the denominator, the fraction carries
    // into the whole part (e.g. 0.9999 with max denom 2 → 1/1 → roll to 1 0/2).
    // We keep that roll-up so the rendered whole matches the value.
    let outWhole = whole;
    let outNum = num;
    if (outNum === den && den !== 0) {
        outWhole += 1;
        outNum = 0;
    }

    if (outNum === 0) {
        // No fractional part after rounding — render "N" (no fraction).
        return sign + String(outWhole);
    }
    return `${sign}${outWhole} ${outNum}/${den}`;
}

function bestFraction(x: number, maxDenom: number): { num: number; den: number } {
    // Standard continued-fraction convergents: track (h_{-1},k_{-1})=(1,0)
    // and (h_0,k_0)=(a0,1). At each step, a = floor(x). When the next
    // convergent's denom would exceed maxDenom, stop.
    if (x === 0) return { num: 0, den: 1 };
    let h1 = 1, k1 = 0;
    let h = Math.floor(x), k = 1;
    let rem = x - h;
    // If x is already ≥ 1 we've folded its integer part into h; for the
    // fractional-only case here x < 1 so h0 = 0, k0 = 1.
    while (rem > 1e-12) {
        const inv = 1 / rem;
        const a = Math.floor(inv);
        const newH = a * h + h1;
        const newK = a * k + k1;
        if (newK > maxDenom) break;
        h1 = h; k1 = k;
        h = newH; k = newK;
        rem = inv - a;
    }
    return { num: h, den: k };
}

// ── Scientific / engineering notation ─────────────────────────────────────

function formatScientific(code: string, value: number): string {
    // Split the format into mantissa + exponent parts around the `E[+-]`
    // marker. Walk the un-quoted code so literal E's inside quotes don't
    // confuse the parser.
    const split = findExponentSplit(code);
    if (!split) return formatNumeric(code, value);
    const { mantissaCode, expSign, exponentCode, before, after } = split;

    // Mantissa decimal count is the number of `0`/`#`/`?` chars after the
    // decimal point in the mantissa pattern. Integer digits before the
    // decimal matter for engineering notation (##0.0E+0 → exponent rounded
    // to multiple of 3 so the mantissa fits the integer-digit pattern).
    const intPlaceholders = countIntegerPlaceholders(mantissaCode);
    const decimalPlaces = countDecimalPlaceholders(mantissaCode);

    // Engineering heuristic: integer part uses `##0` (i.e. ≥3 placeholders)
    // → round exponent to a multiple of 3. Otherwise keep the canonical
    // single-integer-digit scientific form.
    const engineering = intPlaceholders >= 3;

    let mantissa: number;
    let exponent: number;
    if (value === 0) {
        mantissa = 0;
        exponent = 0;
    } else {
        exponent = Math.floor(Math.log10(Math.abs(value)));
        if (engineering) {
            // Round exponent DOWN to the nearest multiple of 3 so the
            // mantissa gains the appropriate integer digits.
            exponent = Math.floor(exponent / 3) * 3;
        }
        mantissa = value / Math.pow(10, exponent);
        // Rounding the mantissa may push it across the next power-of-10
        // boundary (e.g. 9.995 → 10.00 with 2 decimal places). When that
        // happens, divide through and bump the exponent.
        const rounded = Number(mantissa.toFixed(decimalPlaces));
        const boundary = engineering ? Math.pow(10, intPlaceholders) : 10;
        if (Math.abs(rounded) >= boundary) {
            mantissa = rounded / 10;
            exponent += 1;
            if (engineering) {
                // Keep exponent aligned to multiple of 3 if we bumped past.
                const misalign = exponent % 3;
                if (misalign !== 0) {
                    mantissa *= Math.pow(10, misalign);
                    exponent -= misalign;
                }
            }
        }
    }

    // Format mantissa through the existing numeric path so thousands etc.
    // still apply if the mantissa format requested them.
    const mantissaText = formatNumeric(mantissaCode, mantissa);

    // Format the exponent. `E+00` means "two digits with explicit sign";
    // `E+0` means "one digit with sign". Always emit an explicit sign when
    // the code used `E+`; emit a minus-only when it used `E-`.
    const expAbs = Math.abs(exponent);
    const expDigits = exponentCode.replace(/[^0#?]/g, '').length;
    const expBody = String(expAbs).padStart(expDigits, '0');
    const sign = exponent < 0 ? '-' : (expSign === '+' ? '+' : '');
    const expText = sign + expBody;

    return `${before}${mantissaText}E${expText}${after}`;
}

// Split `…<mantissa>E[+-]<exponent>…` into (before, mantissa, sign, exponent, after).
// Returns null if no unquoted `E[+-]` marker is found.
function findExponentSplit(code: string): {
    before: string;
    mantissaCode: string;
    expSign: '+' | '-';
    exponentCode: string;
    after: string;
} | null {
    // Walk the code to find an `E` outside quotes/brackets followed by `+` or `-`.
    let inQuote = false;
    let inBracket = false;
    let eIdx = -1;
    let sign: '+' | '-' = '+';
    for (let i = 0; i < code.length - 1; i++) {
        const c = code[i];
        if (c === '"') { inQuote = !inQuote; continue; }
        if (!inQuote && c === '[') { inBracket = true; continue; }
        if (!inQuote && c === ']') { inBracket = false; continue; }
        if (inQuote || inBracket) continue;
        if (c === '\\') { i++; continue; }
        if ((c === 'E' || c === 'e') && (code[i + 1] === '+' || code[i + 1] === '-')) {
            eIdx = i;
            sign = code[i + 1] as '+' | '-';
            break;
        }
    }
    if (eIdx < 0) return null;

    // Walk back from E to the start of the mantissa placeholder block.
    let mStart = eIdx;
    while (mStart > 0) {
        const c = code[mStart - 1];
        if ('0#?.,'.includes(c)) { mStart--; continue; }
        break;
    }
    // Walk forward from E+sign to the end of the exponent placeholder block.
    let eEnd = eIdx + 2;
    while (eEnd < code.length) {
        const c = code[eEnd];
        if ('0#?'.includes(c)) { eEnd++; continue; }
        break;
    }
    return {
        before: code.slice(0, mStart),
        mantissaCode: code.slice(mStart, eIdx),
        expSign: sign,
        exponentCode: code.slice(eIdx + 2, eEnd),
        after: code.slice(eEnd),
    };
}

function countIntegerPlaceholders(code: string): number {
    // Count 0/#/? chars before the first decimal point.
    const dotIdx = code.indexOf('.');
    const slice = dotIdx < 0 ? code : code.slice(0, dotIdx);
    return (slice.match(/[0#?]/g) ?? []).length;
}

function countDecimalPlaceholders(code: string): number {
    const dotIdx = code.indexOf('.');
    if (dotIdx < 0) return 0;
    let n = 0;
    for (let i = dotIdx + 1; i < code.length; i++) {
        const c = code[i];
        if (c === '0' || c === '#' || c === '?') n++;
        else break;
    }
    return n;
}

// ── Date / time formatting ────────────────────────────────────────────────

// Excel's 1900 date system treats 1900-02-29 as a valid day (historical MS
// bug for Lotus compatibility). Serials ≥ 60 are off by one from true dates
// counted from 1900-01-01; the conventional fix is to treat serial 0 as
// 1899-12-30 so serial 1 = 1900-01-01 and serial 60 = 1900-02-28 (we skip
// 29th, matching Excel).
//
// The 1904 date system (Mac Office pre-2011) has no such leap bug: serial 0
// is 1904-01-01 exactly, and every serial counts forward from there with no
// fudge. Workbook flag `workbookPr/@date1904` selects the mode.
const EPOCH_MS_1900 = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
const EPOCH_MS_1904 = Date.UTC(1904, 0, 1);   // 1904-01-01 UTC
const MS_PER_DAY = 86400000;

function serialToDate(serial: number, date1904: boolean): Date {
    // Round the ms-within-day to nearest second to avoid floating-point
    // artifacts (e.g. 0.75 * 86400 = 64799.999… → 23:59:60).
    const whole = Math.floor(serial);
    const frac = serial - whole;
    const epoch = date1904 ? EPOCH_MS_1904 : EPOCH_MS_1900;
    const ms = epoch + whole * MS_PER_DAY + Math.round(frac * MS_PER_DAY);
    return new Date(ms);
}

function formatDateTime(code: string, value: number, options?: FormatOptions): string {
    const date1904 = options?.date1904 === true;
    const date = serialToDate(value, date1904);
    const Y = date.getUTCFullYear();
    const M = date.getUTCMonth() + 1;
    const D = date.getUTCDate();
    const h24 = date.getUTCHours();
    const m = date.getUTCMinutes();
    const s = date.getUTCSeconds();
    const ampm = /AM\/PM|am\/pm/.test(stripQuotedLiterals(code));
    const h12 = ((h24 + 11) % 12) + 1;
    const hour = ampm ? h12 : h24;

    // Elapsed-time accumulators: total hours / minutes / seconds since the
    // epoch, floored to integer units. When the format code carries `[h]`
    // etc. these replace the wall-clock value, and any remaining non-bracketed
    // minute/second tokens render the remainder inside the next-smaller unit.
    // We use the raw serial here (not the rounded Date) so that e.g. serial
    // 1.5 with `[h]:mm` is 36 hours exactly, not 35:59.
    const totalHours = Math.floor(value * 24);
    const totalMinutes = Math.floor(value * 1440);
    const totalSeconds = Math.floor(value * 86400);

    // Token walker. "m" is month when adjacent to y or d, minute when adjacent
    // to h or s. Track whether we just saw an h/hh token to disambiguate.
    let out = '';
    let i = 0;
    let lastSawHour = false;
    while (i < code.length) {
        const c = code[i];
        if (c === '"') {
            i++;
            while (i < code.length && code[i] !== '"') { out += code[i]; i++; }
            i++;
            continue;
        }
        if (c === '\\' && i + 1 < code.length) {
            out += code[i + 1];
            i += 2;
            continue;
        }
        if (c === '[' ) {
            // [h] / [hh] / [m] / [mm] / [s] / [ss] are elapsed-time markers —
            // the accumulated count of that unit since serial 0, not the
            // modulo wall-clock value. Any other bracket (e.g. [Red] or a
            // stripped locale tag) is skipped silently.
            const end = code.indexOf(']', i);
            if (end < 0) { i++; continue; }
            const inner = code.slice(i + 1, end).toLowerCase();
            i = end + 1;
            if (inner === 'h' || inner === 'hh') {
                out += inner.length >= 2 ? pad2(totalHours) : String(totalHours);
                lastSawHour = true;
            } else if (inner === 'm' || inner === 'mm') {
                out += inner.length >= 2 ? pad2(totalMinutes) : String(totalMinutes);
                lastSawHour = false;
            } else if (inner === 's' || inner === 'ss') {
                out += inner.length >= 2 ? pad2(totalSeconds) : String(totalSeconds);
                lastSawHour = false;
            }
            continue;
        }
        const run = readRun(code, i, c.toLowerCase());
        if (run > 0) {
            const token = code.slice(i, i + run).toLowerCase();
            switch (token[0]) {
                case 'y':
                    out += token.length >= 4 ? String(Y).padStart(4, '0') : String(Y % 100).padStart(2, '0');
                    lastSawHour = false;
                    break;
                case 'd':
                    out += token.length >= 4 ? weekdayName(date, true)
                         : token.length === 3 ? weekdayName(date, false)
                         : token.length === 2 ? pad2(D)
                         : String(D);
                    lastSawHour = false;
                    break;
                case 'h':
                    out += token.length >= 2 ? pad2(hour) : String(hour);
                    lastSawHour = true;
                    break;
                case 's':
                    out += token.length >= 2 ? pad2(s) : String(s);
                    lastSawHour = false;
                    break;
                case 'm':
                    if (lastSawHour) {
                        out += token.length >= 2 ? pad2(m) : String(m);
                    } else {
                        // Month. mmmm = full name, mmm = abbrev, mm = 2-digit, m = 1.
                        out += token.length >= 4 ? monthName(M, true)
                             : token.length === 3 ? monthName(M, false)
                             : token.length === 2 ? pad2(M)
                             : String(M);
                    }
                    break;
                case 'a':
                case 'p':
                    if (token === 'am/pm') {
                        out += h24 < 12 ? 'AM' : 'PM';
                    } else {
                        out += code.slice(i, i + run);
                    }
                    break;
            }
            i += run;
            continue;
        }
        out += c;
        i++;
    }
    return out;
}

function readRun(code: string, start: number, token: string): number {
    // Match am/pm as a single multi-char token.
    const lower = code.slice(start).toLowerCase();
    if (lower.startsWith('am/pm')) return 5;
    if (!'ymdhs'.includes(token)) return 0;
    let n = 0;
    while (start + n < code.length && code[start + n].toLowerCase() === token) n++;
    return n;
}

function pad2(n: number): string {
    return n < 10 ? '0' + n : String(n);
}

const MONTHS_FULL = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
];
const MONTHS_ABBR = MONTHS_FULL.map((m) => m.slice(0, 3));

function monthName(m: number, full: boolean): string {
    const idx = m - 1;
    if (idx < 0 || idx > 11) return String(m);
    return (full ? MONTHS_FULL : MONTHS_ABBR)[idx];
}

const DAYS_FULL = [
    'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
];
const DAYS_ABBR = DAYS_FULL.map((d) => d.slice(0, 3));

function weekdayName(date: Date, full: boolean): string {
    return (full ? DAYS_FULL : DAYS_ABBR)[date.getUTCDay()];
}
