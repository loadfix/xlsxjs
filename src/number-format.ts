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
//   - Text ("@")                   → the raw cell text unchanged
//
// Excel's format-code grammar supports positive;negative;zero;text sections.
// We honour the split but only apply the matched section to the absolute
// value. Colour codes ([Red], [Blue]) and conditional sections ([>100]) are
// parsed but not rendered — they're dropped silently.
//
// Compromises worth flagging for downstream consumers:
//   - `*<char>` is a "fill the remaining column width with <char>" marker in
//     Excel. A web renderer can't measure remaining cell width without a
//     layout pass, so we silently drop the marker rather than emit a bogus
//     single repeat character. Accounting columns that rely on the fill to
//     visually line up their currency symbol and digits will look slightly
//     tighter than in Excel, but the digits themselves are correct.

export interface FormatResult {
    text: string;
    // When true, right-align the cell (numbers, dates, times). When false,
    // treat like a string (text format, @).
    numeric: boolean;
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

export function formatNumber(value: string, formatCode: string, options?: FormatOptions): FormatResult {
    if (formatCode === '' || formatCode.toLowerCase() === 'general') {
        return formatGeneral(value);
    }

    // @ sections are "text" and should echo the input. If that's the only
    // section, treat as a non-numeric result.
    if (formatCode.trim() === '@') {
        return { text: value, numeric: false };
    }

    const num = Number(value);
    if (!Number.isFinite(num)) {
        // Not actually a number — fall back to raw value. Excel itself does
        // this for error cells.
        return { text: value, numeric: false };
    }

    const sections = splitSections(formatCode);
    let sectionIndex: number;
    if (num > 0) sectionIndex = NUMERIC_SECTION_INDEX.positive;
    else if (num < 0) sectionIndex = sections[NUMERIC_SECTION_INDEX.negative] ? NUMERIC_SECTION_INDEX.negative : NUMERIC_SECTION_INDEX.positive;
    else sectionIndex = sections[NUMERIC_SECTION_INDEX.zero] ? NUMERIC_SECTION_INDEX.zero : NUMERIC_SECTION_INDEX.positive;

    const section = sections[sectionIndex] ?? formatCode;
    // [$<symbol>-<localeHex>] is a locale-tagged currency marker. Pull out the
    // symbol (if any) BEFORE stripSquareBracketModifiers eats the whole thing,
    // then splice it back in at the same spot as a quoted literal so the rest
    // of the pipeline treats it as an ordinary literal prefix.
    const { code: sectionWithSymbol } = extractLocaleCurrency(section);
    const cleaned = stripSquareBracketModifiers(sectionWithSymbol);
    // Negative numbers: the negative section format itself typically includes
    // the sign or parentheses. If we fall back to the positive section, prefix
    // a minus.
    let magnitude = Math.abs(num);
    if (num < 0 && sectionIndex === NUMERIC_SECTION_INDEX.positive) {
        return { text: '-' + applyFormat(cleaned, magnitude, options), numeric: true };
    }
    return { text: applyFormat(cleaned, magnitude, options), numeric: true };
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
    if (!inserted) out = numberText + out;
    return out;
}

function formatGeneral(value: string): FormatResult {
    const n = Number(value);
    if (!Number.isFinite(n)) return { text: value, numeric: false };
    return { text: formatGeneralNumber(n), numeric: true };
}

function formatGeneralNumber(n: number): string {
    // "General" trims trailing zeros and falls back to scientific for very
    // large/small magnitudes. toString() is close enough for the common case.
    return n.toString();
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
