import { Workbook } from './workbook';
export { XlsxEncryptedError, sanitizeMediaMime, bytesToDataUrl, MAX_EMBEDDING_BYTES } from './workbook';
import { WorkbookParser } from './workbook-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';

export type { Workbook as ParsedWorkbook, Sheet, SheetView, Cell, RichTextRun, SharedString, PhoneticRun, FrozenPanes, AutoFilter, TableDef, SheetChart, SheetPivot, SheetExtensionUri, SheetImage, SheetEmbedding, SheetComment, ThreadedCommentEntry, SheetOutline, DefinedName, ColumnWidth, RowDimension, Hyperlink, DataValidationList, PageBreaks, PrintAreaRange, HeaderFooter, HeaderFooterZones } from './workbook-parser';
export { parseThreadedComments, isSafeHyperlinkHref } from './workbook-parser';
export type { Styles, CellXf, FontStyle, FillStyle, BorderStyle, Dxf, UnderlineStyle } from './styles';
export type { Theme, ColorRef } from './theme';
export type { ConditionalFormatting, CfRule, CfRuleType, CfOperator, CellRange, Cfvo, ColorScale, DataBar, IconSet } from './conditional-format';
export { parseStyles, lookupNumberFormat, sanitizeHexColor, sanitizeFontFamily, parseColorElement, resolveEffectiveXf } from './styles';
export { parseTheme, applyTint, resolveColor, indexedColor } from './theme';
export { parseConditionalFormatting, evaluateRule, resolveCfvo, interpolateColorScale } from './conditional-format';
export { formatNumber } from './number-format';
export { a1ToR1c1, r1c1ToA1 } from './formula-notation';
export { emuToPx } from './utils';

export interface Options {
    className: string;
    inWrapper: boolean;
    debug: boolean;
    // When true, formula cells render their formula text (with leading "=")
    // instead of the cached value. Useful for audit views.
    showFormulas: boolean;
    // Formula notation used for display. Has no effect unless showFormulas
    // is true. 'a1' is Excel's default; 'r1c1' is an author-relative form
    // popular in accounting contexts.
    formulaNotation: 'a1' | 'r1c1';
    // Opt-in inlining of embedded OLE / package payloads as `data:` URLs on
    // the rendered `<aside class="xlsx-embedding">`. Detection + metadata
    // surfacing happen unconditionally; only the payload projection is gated
    // on this flag, since a naive `dataUrl` lands the full embedding in both
    // model and DOM memory. Oversized payloads (>32 MiB) or MIMEs outside
    // the sanitizer's allowlist keep `dataUrl` null even when this is true.
    inlineEmbeddings: boolean;
    h: typeof h;
}

export const defaultOptions: Options = {
    className: 'xlsx',
    inWrapper: true,
    debug: false,
    showFormulas: false,
    formulaNotation: 'a1',
    inlineEmbeddings: false,
    h,
};

function mergeOptions(userOptions?: Partial<Options>): Options {
    return { ...defaultOptions, ...userOptions };
}

export async function parseAsync(data: Blob | ArrayBuffer | Uint8Array, userOptions?: Partial<Options>): Promise<Workbook> {
    const ops = mergeOptions(userOptions);
    return Workbook.load(data, new WorkbookParser(ops));
}

export async function renderWorkbook(workbook: Workbook, userOptions?: Partial<Options>): Promise<Node[]> {
    const ops = mergeOptions(userOptions);
    const renderer = new HtmlRenderer();
    if (!workbook.parsed) throw new Error('xlsx-preview: workbook is not parsed');
    return await renderer.render(workbook.parsed, ops);
}

export async function renderAsync(
    data: Blob | ArrayBuffer | Uint8Array,
    bodyContainer: HTMLElement,
    styleContainer?: HTMLElement,
    userOptions?: Partial<Options>,
): Promise<Workbook> {
    const wb = await parseAsync(data, userOptions);
    const nodes = await renderWorkbook(wb, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = '';
    bodyContainer.innerHTML = '';

    for (const n of nodes) {
        const c = n.nodeName === 'STYLE' ? styleContainer : bodyContainer;
        c.appendChild(n);
    }
    return wb;
}
