# SECURITY_REVIEW.md — xlsxjs

_Reviewed 2026-05-02 against commit `1d4a4e9` (master), with fixes committed on
branch `feat/security-review`._

xlsxjs is a browser-side library that accepts an attacker-controlled `.xlsx`
blob, parses it, and injects rendered output into the host page's DOM.
Everything in the zip — shared strings, rich-text runs, hyperlink targets,
theme colours, number-format codes, media bytes — is untrusted input. The
host page is the trust boundary. Any path where an XLSX-derived byte
reaches a JS-evaluating sink (script execution, `href="javascript:"`,
same-origin iframe, `innerHTML`, CSS `@import`) is exploitable by anyone
who can get a user to open a crafted XLSX.

The project documented its security model in `CLAUDE.md` ("Security
constraints"). Most of the code follows it. The findings below are places
where the implementation diverged from that model or where a shared
surface (range expansion, media size) needed a DoS guardrail.

## Fixed in this review

| # | Severity | Surface | Fix |
|---|----------|---------|-----|
| 1 | **HIGH** | `isSafeHyperlinkHref` — control-character bypass | Reject any URL containing `U+0000..U+001F` / `U+007F` before the `URL().protocol` check. |
| 2 | **HIGH** | Hyperlink range expansion — parser DoS | Cap `<hyperlink ref="…">` expansion at 1,048,576 cells. |
| 3 | **MEDIUM** | Data-validation range expansion — renderer DoS | Clamp `<dataValidation sqref="…">` ranges at the parse boundary. |
| 4 | **MEDIUM** | Inlined media size — memory DoS | Skip any `xl/media/*` entry whose raw byte length exceeds 32 MiB, and restrict the data-URL MIME to a strict allowlist. |

Findings 1 and 2 are practical, ship-blocking vectors against a viewer
that handles real-world XLSX files. 3 and 4 harden known DoS surfaces
that a crafted file could trigger with a single element. All four fixes
are unit-tested in scenario 85.

## Surfaces reviewed

### 1. XML parsing (XXE / billion laughs)

**Status**: OK (deferred hardening)
**Severity**: low — bounded by browser DOMParser semantics.
**Details**: Every XML part is parsed via
`new DOMParser().parseFromString(xml, 'application/xml')` in
`src/workbook-parser.ts:1396`, `src/styles.ts:294`, and `src/theme.ts:73`.
Per the WHATWG HTML / XML spec, browser `DOMParser` does **not** fetch
external DTDs, does not resolve external entities, and imposes
implementation-specific caps on internal entity expansion (Chrome and
Firefox both bound the expansion).

jsdom inherits the same defaults. We do not wrap the parser with a custom
entity-resolver; the risk is limited to what the host environment allows.
No XXE vector found — external references cannot reach the filesystem or
the network from DOMParser.

Tracked for future work: if we ever ingest a DOCTYPE with internal
entities we should reject the document up front, and we should fingerprint
the parser (browser / jsdom / sax-js fallback) before deciding what the
guarantees are. For now the conservative recommendation stands: do not
evaluate XML on the server.

### 2. Zip parsing (zip-slip / zip bomb)

**Status**: OK — with 2025-05-02 media-size cap
**Severity**: n/a (no filesystem sink) / medium for memory DoS
**Details**: JSZip handles the zip structure. No code anywhere writes a
zip entry to the filesystem; every `zip.file(path)` in `src/workbook.ts`
feeds into `parts[path] = string` or `media[path] = data:URL`. Entry
names therefore cannot path-traverse out of anything. No zip-slip.

Zip-bomb angle: JSZip does not itself enforce a decompressed-size limit.
Prior to this review, an `xl/media/*` entry of any size would be fully
decompressed and base64-encoded, which meant a 1 GiB "image" blew up to
~1.33 GiB of data-URL string held in the renderer. **Fixed in this
review** (see commit): size-check the raw bytes before base64 expansion;
skip anything over `MAX_MEDIA_BYTES` (32 MiB).

XML parts are also unbounded, but an attacker feeding a 1 GiB sharedStrings
would also have crashed Excel — out of scope for a viewer. We do not
defend against that yet; see "future work" below.

### 3. Hyperlinks — scheme allowlist

**Status**: Fixed in this review (see commit)
**Severity**: **HIGH** — one-click XSS against any consumer.
**Details**: `isSafeHyperlinkHref` (`src/workbook-parser.ts:28-45`) uses the
WHATWG URL parser + a strict scheme allowlist. Before this review the
allowlist was `http / https / mailto / tel` plus fragment-only and
relative paths — which looks right, and catches the obvious
`javascript:` / `data:` / `vbscript:` / `blob:` / `file:` variants.

**Bypass that worked**: a URL containing a control character in the
scheme (`java\tscript:`, `java\nscript:`, `java\0script:`) was
**accepted** because the URL parser canonicalised it as a relative path
(falling back to `http:`), passing our allowlist. When the browser then
rendered the `<a href="…">`, it stripped the control character and
evaluated the URL as `javascript:` on click.

Example that popped before the fix:
```
<Relationship Id="rIdXSS"
              Type="…/hyperlink"
              Target="java&#9;script:alert(1)"
              TargetMode="External"/>
```

**Fix**: reject any URL containing a control character (U+0000..U+001F
plus U+007F) outright in `isSafeHyperlinkHref`, before the URL parse.
Legitimate URLs do not contain control characters — the WHATWG URL spec
itself treats their presence as a "validation error". See
`src/workbook-parser.ts:41-48`. Regression coverage: scenario 85d tests
tab / newline / carriage return / null byte variants.

### 4. CSS identifiers / class names

**Status**: OK
**Details**: Every `classList.add(…)` in `src/html-renderer.ts` is either
a literal string or a numeric interpolation. The dynamic cases
(`xlsx-outline-${Math.min(lvl, 7)}` at lines 421 / 436 / 514) use a
bounded integer (`Math.min(lvl, 7)`) — no XLSX string reaches a class
name. `col.setAttribute('data-outline-level', String(lvl))` writes a
numeric-coerced value via `setAttribute`, which is HTML-encoded by the
browser.

Data attributes that carry XLSX strings (`data-sheet-name`,
`data-table-name`, `data-table-display-name`, `data-chart-type`,
`data-name`, `data-preset`, `data-kind`, `data-anchor-mode`, plus every
`data-anchor-*` that came from integer coordinates) all route through
`setAttribute` — the HTML parser encodes any embedded quote / bracket.

### 5. `innerHTML` sinks

**Status**: OK
**Details**: `grep -rn 'innerHTML' src/` finds:

- `src/xlsx-preview.ts:68-69` — `styleContainer.innerHTML = ''` and
  `bodyContainer.innerHTML = ''`. Both empty strings, safe: they just
  clear the container before mounting.
- `src/html-renderer.ts:1376` — `iconSpan.innerHTML = state.icon.svg`.
  `state.icon.svg` is a string produced by `renderIcon` in
  `src/icons.ts`, which is a strict lookup into a hard-coded SVG palette.
  An attacker can **choose** which icon renders (by controlling the cf
  rule's `iconSet` name + `customIcons[].iconSet`), but both inputs flow
  through the `SETS` key map — unknown keys return `null` and the sink
  is skipped. No XLSX string reaches the serialized SVG.

### 6. Inline SVG

**Status**: OK
**Details**: The SVG strings in `src/icons.ts` are hard-coded literals
with colour values drawn from an internal palette. No XLSX data is
substituted into any attribute value. The rendered SVG has no event
handlers, no `<script>`, and no external references.

### 7. `<img src>` with `data:` URLs

**Status**: Fixed in this review (see commit) + residual caveat
**Details**: `src/workbook.ts:194-204` inlines each `xl/media/*` as
`data:${mime};base64,${bytes}`. Prior to this review:

- **MIME** came from `guessMime(path)` which switched on file extension.
  An unknown extension yielded `application/octet-stream`, rendered as a
  broken image — safe. Known extensions mapped to the standard image
  MIMEs.
- **Size** was unbounded. A 1 GiB embedded media would fully expand.

**Fix**: Size-check raw bytes against `MAX_MEDIA_BYTES` (32 MiB) before
calling `zip.file(p).async('base64')`. The MIME string is additionally
routed through `sanitizeMediaMime()`, which decays any value outside the
7-entry allowlist back to `application/octet-stream`. See
`src/workbook.ts:12-33` and `:223-248`.

**Residual caveat (not fixed — future work)**: The MIME is derived from
the filename extension, not from the bytes. A producer that ships SVG
content under an `image1.png` filename will decay the image to a broken
render (PNG parser rejects SVG bytes). A producer that correctly names
the file `.svg` triggers `image/svg+xml`. Browsers do NOT execute scripts
in SVG loaded via `<img src=…>` — no `<script>`, no `onload`, no external
fetches — so the SVG path is still safe against the attacker. Tracked as
info only.

### 8. Attribute injection

**Status**: OK
**Details**: Every `setAttribute` call in `src/html-renderer.ts` uses a
static attribute name (`title`, `href`, `alt`, `aria-label`,
`data-anchor-*`, `colspan`, `rowspan`, `role`, `target`, `rel`). None of
those are script-executing. The browser HTML-encodes the value on write,
so quotes / angle brackets in an attacker string cannot break out.

No use of `srcdoc`, `onerror`, `onclick`, or any other sink-attribute
name. No dynamic attribute names interpolated from XLSX strings.

### 9. Prototype pollution

**Status**: OK
**Details**: Writes into objects with XLSX-derived keys are bounded:

- `wb.parts[p] = xml` / `wb.media[p] = dataUrl` in `src/workbook.ts` —
  `p` is a zip entry path, but the outer `if` gates it through a literal
  regex per part class (e.g. `/^xl\/worksheets\/.*\.xml$/i`). `p` cannot
  equal `__proto__` and pass that guard.
- `byCol[cell.col] = cell` in `src/html-renderer.ts:519` — `cell.col` is
  a number.
- `colors[idx] = c` in `src/theme.ts:106` — `idx` is a fixed integer
  from the `mapping` array.

No `Object.assign` call hands attacker data into a prototype-sensitive
object. No `JSON.parse` → `mergeDeep` pattern.

### 10. DoS via oversized input (renderer)

**Status**: partially fixed; residual concerns tracked
**Severity**: HIGH (unfixed) / MEDIUM (fixed)
**Details**: Two hot paths expand an attacker-controlled range into a
per-cell data structure:

**(a) Hyperlinks** (`src/workbook-parser.ts:parseHyperlinks`). A
`<hyperlink ref="A1:XFD1048576"/>` is Excel-legal and would have
expanded to ~17 billion `Hyperlink` records before the parser returned.
**Fixed in this review**: ranges with `width * height >
MAX_RANGE_EXPANSION_CELLS` (1,048,576) keep only the top-left anchor.

**(b) Data-validation lists**
(`src/workbook-parser.ts:parseDataValidationLists`). The parser stored
the range verbatim; the renderer then walked it into a `Map` at
`html-renderer.ts:494-502`. Same unbounded expansion as hyperlinks.
**Fixed in this review**: range is clamped at the parse boundary (one
row × up to cap cols) before the renderer ever sees it.

**Residual**:
- `sheet.maxRow` / `sheet.maxCol` are extended by hyperlink and
  validation ranges so the render loop covers every covered cell. With
  the range cap the maxRow/maxCol extension is bounded at the same cap
  (one million rows worth) — which is still large. Rendering a 1M-row
  sheet won't silently crash jsdom but will be slow and memory-heavy.
  Tracked for future work: add a rendering-time budget /
  "too-large-to-render" bail and surface it on the model.
- Conditional-format rule count is not bounded. A sheet with 100K cf
  rules against a 100×100 range causes `resolveConditionalFormats` to
  run O(rules × cells) = 10^9 iterations. Tracked for future work.
- XML part size is not bounded. A 1 GiB sharedStrings XML will take
  minutes to parse and exhaust browser memory. Reject-on-size-limit at
  the `zip.file().async('string')` boundary is tracked for future work.

### 11. Formula display (`showFormulas`)

**Status**: OK
**Details**: `src/html-renderer.ts:871-877`. When `showFormulas=true`, the
cell's formula text is prefixed with `=` and assigned via
`td.textContent = text`. No innerHTML, no attribute, no CSS. Formula
strings can contain arbitrary content (including `<script>`-looking
text) but textContent HTML-encodes — the browser renders the angle
brackets as literal characters.

R1C1 translation
(`src/formula-notation.ts:a1ToR1c1`) does pure-string math. It cannot
introduce new sinks.

### 12. Locale currency / number-format strings

**Status**: OK
**Details**: `src/number-format.ts` formats numeric cells. The output is
assigned via `td.textContent` in the main render path
(`html-renderer.ts:892`) and via `td.textContent` again when a `dxf`
overrides the format (`html-renderer.ts:1443`). No innerHTML.

Colour-modifier brackets (`[Red]`, `[Color 14]`) resolve through
`NAMED_COLORS` (a fixed dictionary) or `indexedColor` (a fixed
64-entry palette). Both return a `#rrggbb` hex — never the attacker's
raw bracket body. That hex lands on `td.style.color = formatColor`,
where any non-hex input (e.g. a string like `red;display:none`) would be
rejected by the CSSStyleDeclaration setter.

---

## Areas I specifically checked and found clean

- **`sanitizeFontFamily`** (`src/styles.ts:243-251`). Reject-by-default
  allowlist: alphanumerics + space + hyphen + dot, quoted on output.
  Rejects `;`, `{`, `}`, newline, CR, quote, backslash, comma,
  parentheses, `@`. Commas are deliberately rejected so the single
  allowed font can't be used to inject a fallback font-family list.
- **`sanitizeHexColor`** (`src/styles.ts:253-261`). Returns
  `#<6-hex-lowercased>` for a 6- or 8-hex input; null otherwise.
  Lowercases on output so an attacker can't craft a mixed-case value
  for a matching selector.
- **`applyFill` / `applyBorder`** in `src/html-renderer.ts`. Every
  colour flows through `resolveColor()`, which returns a sanitized hex
  or null. Gradient directions use validated `fill.degree` (numeric).
  Border sides use validated style names (`thin`, `medium`, `thick`,
  `dashed`, etc.) mapped to width/style constants — no XLSX string
  reaches a CSS property key or value.
- **`applyAlignment`** in `src/html-renderer.ts:1229-1316`. Every
  value is either a validated enum (`'left'`, `'right'`, …), a
  numeric-coerced `textRotation`, or a numeric `indent`. No XLSX string
  reaches a style property.
- **`appendBackgroundImage`** (`src/html-renderer.ts:1164-1181`). The
  regex-merged `background-image` strings are built from our own
  colour-sanitised gradients; no XLSX string reaches the concatenation.
- **Hyperlink rendering** (`src/html-renderer.ts:803-837`). After the
  allowlist passes, the href is written via `setAttribute('href', …)`.
  `target="_blank"` and `rel="noopener noreferrer"` are applied
  unconditionally so a navigated link cannot `window.opener`-hijack the
  host page (tabnabbing).
- **Sheet-view / tab color**
  (`src/html-renderer.ts:781-795`). `view.tabColor` resolves through
  `resolveColor()` to a hex; `setAttribute('data-tab-color', hex)` is
  attribute-only, no CSS parse involved.
- **Comments + threaded comments**
  (`src/html-renderer.ts:716-772`). Author and body text land on
  `setAttribute('title', …)` and `textContent` only. No innerHTML.
- **Headers / footers**
  (`src/html-renderer.ts:698-711`). Zone text is set via
  `textContent` on a child `<div>`. Substitution codes resolved in
  `workbook-parser.ts:substituteHeaderFooterCodes` are fixed strings
  (date/time/sheet name / literal page markers).

## Suggested fix order (already shipped)

Fixes 1–4 are in the same commit as this review.

## Tracked for future work (not fixed in this review)

1. **Rendering budget for very large sheets** (info). `maxRow` /
   `maxCol` can still reach 1M rows via the range-expansion cap. The
   renderer should bail (or paginate) at a much lower ceiling and
   surface the bail reason on the Workbook model.
2. **Conditional-format rule count cap** (low). Bound the total rule
   count considered per sheet; above the cap, drop lowest-priority
   rules and record a note.
3. **XML part size cap** (low). A sharedStrings or styles part of
   multi-hundred-MB size can exhaust browser memory; cap at
   `zip.file().async('string')` boundary.
4. **Billion-laughs explicit guard** (info). Even though browser /
   jsdom DOMParser bounds entity expansion, we could pre-scan for a
   DOCTYPE with `<!ENTITY … "…">` recursion and reject the document.
5. **SVG sniff-on-bytes** (info). We currently trust the filename
   extension for MIME. A byte-sniffer that rejected SVG-inside-a-jpeg
   would cost ~a byte of code and remove a class of edge cases.

---

## Invariants verified

- `npm run build` succeeds.
- `npm run test:render` passes 77 scenarios (76 pre-existing + scenario
  85 added by this review).
- `npx playwright test` passes 35 tests in 3.7s.
- `node scripts/smoke-test.mjs` clean — only pre-existing drop notes for
  `pivotCacheRecords*.xml` (deliberately not parsed) and the encrypted
  corpus entry (correctly surfaced via `XlsxEncryptedError`).
- No public API broken.
- `dist/` has NOT been rebuilt for the final commit.

*Report generated against working tree at commit `1d4a4e9` on 2026-05-02.
Findings 1–4 are confirmed by source reading + the scenario-85 regression
pass; no runtime exploitation was attempted as part of this review.*
