// SmartArt data model → hierarchy tree. Parses `xl/diagrams/data1.xml` into a
// flat list of points (`<dgm:pt>`) connected by parent-of relationships from
// `<dgm:cxnLst>`. Root nodes are points not referenced as `destId` in any
// parOf connection (or those referenced by the doc-root pt, type="doc").
//
// We only surface the hierarchy + text. Layout / colour / style definitions
// (layoutDef / colorsDef / styleDef) are NOT interpreted — the renderer emits
// a detect-only indented tree. `layoutDef/@uniqueId` is surfaced for
// consumers who want to branch on the layout variant.
//
// Security contract: every string reaches the model verbatim. The renderer
// pushes them through textContent / setAttribute. No regex-or-map keys are
// derived from attacker-controlled strings here — the only Map keys we build
// are modelId strings, but lookups against those Maps are not fed into CSS
// selectors or innerHTML.

const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

// Maximum depth for the tree walk. Malformed / circular parent-of links
// would otherwise infinite-loop; we bail after this depth so a crafted
// data1.xml can't hang the parser.
const MAX_TREE_DEPTH = 32;

export interface SmartArtNode {
    // Raw @modelId off the <dgm:pt>. We surface it as an opaque string so
    // consumers can match entries across renders without xlsxjs needing to
    // interpret Excel's integer / GUID scheme. Not used in a CSS selector.
    id: string;
    // Flattened text body: every <a:t> under <dgm:t> joined with no
    // separator. '' when the point carries no text.
    text: string;
    children: SmartArtNode[];
    // Depth from the diagram root. Root nodes are level 0; their immediate
    // children are level 1, etc. Computed after the tree has been built so
    // the level is consistent under defensive cycle-break handling.
    level: number;
}

export interface SmartArtModel {
    // Synthetic diagram id. Empty string when the data1.xml declared no
    // <dgm:dataModel> id attribute (the schema doesn't define one — we
    // reserve the field in case producers add a custom id via extLst).
    id: string;
    rootNodes: SmartArtNode[];
    // <dgm:layoutDef uniqueId="…"> from the companion layout1.xml, when
    // available. null when layoutXml is absent / malformed / carries no
    // uniqueId. Consumers can use this to discriminate the diagram layout
    // (hierarchy1 / cycle1 / pyramid1 / etc.) without xlsxjs interpreting
    // the layout graph itself.
    layout: string | null;
}

/**
 * Parse a SmartArt diagram's `data1.xml` (and optionally its `layout1.xml`)
 * into a hierarchy tree. The result is stable: points in modelId order,
 * children in srcOrd order. Malformed inputs degrade gracefully:
 *   - missing dataModel / ptLst                       → rootNodes = []
 *   - cxn entries pointing at unknown modelIds        → ignored
 *   - cycles in the parent-of graph                   → broken at depth 32
 */
export function parseSmartArt(dataXml: string, layoutXml?: string | null): SmartArtModel {
    const empty: SmartArtModel = { id: '', rootNodes: [], layout: parseLayoutName(layoutXml ?? null) };
    if (!dataXml) return empty;

    let doc: Document;
    try {
        doc = new DOMParser().parseFromString(dataXml, 'application/xml');
    } catch {
        return empty;
    }

    const dataModel = doc.getElementsByTagNameNS(DGM_NS, 'dataModel').item(0);
    if (!dataModel) return empty;

    const ptLst = dataModel.getElementsByTagNameNS(DGM_NS, 'ptLst').item(0);
    if (!ptLst) return empty;

    // Build modelId → { text, type } map. We intentionally use a Map rather
    // than a plain object so attacker-controlled keys can't collide with
    // Object.prototype properties.
    interface RawPoint {
        id: string;
        type: string;     // 'node' | 'doc' | 'asst' | 'pres' | 'parTrans' | 'sibTrans' | '' when absent
        text: string;
    }
    const points = new Map<string, RawPoint>();
    const ptEls = ptLst.getElementsByTagNameNS(DGM_NS, 'pt');
    for (let i = 0; i < ptEls.length; i++) {
        const pt = ptEls[i];
        const id = pt.getAttribute('modelId');
        if (!id) continue;
        const type = pt.getAttribute('type') ?? '';
        // Skip presentation / transition points — they carry layout plumbing,
        // not user-visible hierarchy. We keep 'doc' (diagram root), 'node'
        // (standard tree nodes), and 'asst' (org-chart assistants, which
        // Excel shows on the tree).
        if (type === 'pres' || type === 'parTrans' || type === 'sibTrans') continue;
        const text = flattenPtText(pt);
        points.set(id, { id, type, text });
    }

    // Build parent-of relationships. cxn/@type defaults to "parOf" per the
    // RNC; we accept either an absent type or an explicit "parOf". Other
    // types (presOf, presParOf, unknownRelationship) are ignored — they
    // describe layout, not the node hierarchy.
    interface ParentOf {
        src: string;
        dest: string;
        ord: number;
    }
    const parOfs: ParentOf[] = [];
    const cxnLst = dataModel.getElementsByTagNameNS(DGM_NS, 'cxnLst').item(0);
    if (cxnLst) {
        const cxnEls = cxnLst.getElementsByTagNameNS(DGM_NS, 'cxn');
        for (let i = 0; i < cxnEls.length; i++) {
            const cxn = cxnEls[i];
            const type = cxn.getAttribute('type') ?? 'parOf';
            if (type !== 'parOf') continue;
            const src = cxn.getAttribute('srcId');
            const dest = cxn.getAttribute('destId');
            if (!src || !dest) continue;
            // Only accept edges whose src + dest both resolve to a known
            // point. Entries pointing at pres/parTrans/sibTrans points are
            // dropped by construction (those ids aren't in the map).
            if (!points.has(src) || !points.has(dest)) continue;
            const ordAttr = cxn.getAttribute('srcOrd');
            const ord = ordAttr !== null && Number.isFinite(Number(ordAttr))
                ? Number(ordAttr)
                : 0;
            parOfs.push({ src, dest, ord });
        }
    }

    // Build the children-of map (src → ordered list of dest ids). Sort by
    // srcOrd so the final tree reads in the order Excel presents it.
    const childrenOf = new Map<string, string[]>();
    // Track everything referenced as a destId so we can identify roots by
    // absence below.
    const referencedAsDest = new Set<string>();
    const grouped = new Map<string, ParentOf[]>();
    for (const edge of parOfs) {
        const list = grouped.get(edge.src) ?? [];
        list.push(edge);
        grouped.set(edge.src, list);
        referencedAsDest.add(edge.dest);
    }
    for (const [src, list] of grouped) {
        list.sort((a, b) => a.ord - b.ord);
        childrenOf.set(src, list.map((e) => e.dest));
    }

    // Root detection. Excel emits a "doc" point as the synthetic diagram
    // root; every user-visible root is a child of that point. If a doc
    // point exists, use its children. Otherwise fall back to any point not
    // referenced as a destId.
    const docPt = Array.from(points.values()).find((p) => p.type === 'doc');
    let rootIds: string[];
    if (docPt && childrenOf.has(docPt.id)) {
        rootIds = childrenOf.get(docPt.id) ?? [];
    } else {
        rootIds = Array.from(points.values())
            .filter((p) => p.type !== 'doc' && !referencedAsDest.has(p.id))
            .map((p) => p.id);
    }

    // Walk the tree depth-first, breaking cycles after MAX_TREE_DEPTH and
    // skipping ids that are already on the ancestor stack (a direct cycle
    // is broken immediately without blowing the recursion).
    const rootNodes: SmartArtNode[] = [];
    const onStack = new Set<string>();
    for (const rootId of rootIds) {
        const root = buildNode(rootId, points, childrenOf, onStack, 0);
        if (root) rootNodes.push(root);
    }

    return { id: '', rootNodes, layout: parseLayoutName(layoutXml ?? null) };
}

function buildNode(
    id: string,
    points: Map<string, { id: string; type: string; text: string }>,
    childrenOf: Map<string, string[]>,
    onStack: Set<string>,
    depth: number,
): SmartArtNode | null {
    const pt = points.get(id);
    if (!pt) return null;
    if (depth >= MAX_TREE_DEPTH) {
        // Defensive cut-off. A legitimately deep hierarchy diagram rarely
        // exceeds 4-5 levels; 32 is well beyond realistic inputs.
        return { id: pt.id, text: pt.text, children: [], level: depth };
    }
    if (onStack.has(id)) {
        // Direct cycle — break without recursing. The caller still sees a
        // leaf node for the id, so the tree shape stays well-formed.
        return { id: pt.id, text: pt.text, children: [], level: depth };
    }
    onStack.add(id);
    const children: SmartArtNode[] = [];
    const childIds = childrenOf.get(id) ?? [];
    for (const childId of childIds) {
        const child = buildNode(childId, points, childrenOf, onStack, depth + 1);
        if (child) children.push(child);
    }
    onStack.delete(id);
    return { id: pt.id, text: pt.text, children, level: depth };
}

// Flatten <dgm:t>/<a:r>/<a:t> — the SmartArt point text body. Runs within a
// paragraph concatenate without a separator; multiple paragraphs (<a:p>)
// join with a single newline so a consumer can tell them apart if they
// care, but they rarely appear in practice.
function flattenPtText(pt: Element): string {
    const t = pt.getElementsByTagNameNS(DGM_NS, 't').item(0);
    if (!t) return '';
    const pEls = t.getElementsByTagNameNS(A_NS, 'p');
    if (pEls.length === 0) {
        // Some producers write <dgm:t><a:t>…</a:t></dgm:t> without a <a:p>
        // wrapper. Flatten every descendant <a:t>.
        const tEls = t.getElementsByTagNameNS(A_NS, 't');
        let out = '';
        for (let i = 0; i < tEls.length; i++) out += tEls[i].textContent ?? '';
        return out;
    }
    const paragraphs: string[] = [];
    for (let i = 0; i < pEls.length; i++) {
        const p = pEls[i];
        const tEls = p.getElementsByTagNameNS(A_NS, 't');
        let line = '';
        for (let j = 0; j < tEls.length; j++) line += tEls[j].textContent ?? '';
        paragraphs.push(line);
    }
    return paragraphs.join('\n');
}

function parseLayoutName(layoutXml: string | null): string | null {
    if (!layoutXml) return null;
    let doc: Document;
    try {
        doc = new DOMParser().parseFromString(layoutXml, 'application/xml');
    } catch {
        return null;
    }
    const def = doc.getElementsByTagNameNS(DGM_NS, 'layoutDef').item(0);
    if (!def) return null;
    return def.getAttribute('uniqueId') || null;
}
