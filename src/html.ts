// Minimal hyperscript helper, mirroring docxjs's `h`. Kept simple: accepts a
// tag name, an optional attribute bag, and an optional children list.

export type HChild = Node | string | null | undefined | false;

export function h(tag: string, attrs?: Record<string, string | number | null | undefined> | null, children?: HChild[]): HTMLElement {
    const el = document.createElement(tag);
    if (attrs) {
        for (const [k, v] of Object.entries(attrs)) {
            if (v === null || v === undefined) continue;
            el.setAttribute(k, String(v));
        }
    }
    if (children) {
        for (const c of children) {
            if (c === null || c === undefined || c === false) continue;
            el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
        }
    }
    return el;
}
