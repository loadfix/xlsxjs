// Inline-SVG icons for the common Excel iconSets. The palette here covers
// the three 3-icon sets consumers run into most often; anything else falls
// through to null and the renderer does nothing (the cf rule's icon-set
// name is still exposed on the model so a caller can plug in their own art).
//
// SVGs are 12×12 viewBox, monochrome fills, text-baseline aligned. They are
// emitted inline as serialized strings so the caller has no runtime
// dependency on an icon font.

export type IconSetName = '3TrafficLights1' | '3Arrows' | '3Symbols' | '3Symbols2';

// Named pieces, indexed 0..n-1 (position in the sorted set, matching Excel).
// Red/warning variants use Excel's stock colour palette.
const TL_COLORS = ['#d13438', '#ffb900', '#107c10']; // red / amber / green
const ARROW_COLORS = ['#d13438', '#ffb900', '#107c10'];

function circle(color: string): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><circle cx="6" cy="6" r="5" fill="${color}"/></svg>`;
}

function arrow(color: string, angleDeg: number): string {
    // Arrow pointing up by default, rotated per call: 0=up, 180=down, 90=right.
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><g transform="rotate(${angleDeg} 6 6)"><path d="M6 1 L10 7 L7 7 L7 11 L5 11 L5 7 L2 7 Z" fill="${color}"/></g></svg>`;
}

function cross(): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M3 3 L9 9 M9 3 L3 9" stroke="#d13438" stroke-width="2" stroke-linecap="round"/></svg>`;
}
function bang(): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M6 1 L11 11 L1 11 Z" fill="#ffb900" stroke="#8a6300" stroke-width="0.5"/><path d="M6 5 L6 8" stroke="#202020" stroke-width="1" stroke-linecap="round"/><circle cx="6" cy="10" r="0.6" fill="#202020"/></svg>`;
}
function tick(): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M2 7 L5 10 L10 3" stroke="#107c10" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
}

// Icon tables per set. Indexed 0..n-1 by position (not by colour name).
// Excel's convention in these 3-icon sets is "lowest first" → icons[0] is
// the "bad" state; icons[2] is the "good" state.
const SETS: Record<IconSetName, string[]> = {
    '3TrafficLights1': [circle(TL_COLORS[0]), circle(TL_COLORS[1]), circle(TL_COLORS[2])],
    '3Arrows':         [arrow(ARROW_COLORS[0], 180), arrow(ARROW_COLORS[1], 90), arrow(ARROW_COLORS[2], 0)],
    '3Symbols':        [cross(), bang(), tick()],
    '3Symbols2':       [cross(), bang(), tick()],
};

export function renderIcon(set: string, index: number, reverse: boolean): string | null {
    const icons = SETS[set as IconSetName];
    if (!icons) return null;
    const effective = reverse ? icons.length - 1 - index : index;
    if (effective < 0 || effective >= icons.length) return null;
    return icons[effective];
}
