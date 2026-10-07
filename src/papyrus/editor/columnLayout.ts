// Text columns on the page (persona-300, 雙語公告; w:cols): a section Word lays out in columns is
// shown and printed in columns too. How (see pagination.ts):
//   - every block of the section is as wide as a column (sectionColumns' padding, column 1);
//   - pagination fills the columns one after another: a block that does not fit in the column
//     moves to the top of the next one (a column break sends what follows there), after the last
//     column to the next page. A block in another column is moved there on screen (relative top
//     and its padding), which changes nothing about the text flow the layout is measured on;
//   - the line between the columns (w:sep) is drawn on the page.
// What differs from Word: a paragraph or table moves to the next column as a whole (Word breaks a
// paragraph between its lines and a table between its rows there); the columns of a section's last
// page are not balanced before a continuous section break; a column break inside a paragraph sends
// the next paragraph to the next column.
import { twipsToPx } from '../units';
import './columnLayout.css';

/** A section's columns in px, from the left edge of its text area. */
export interface ColumnGeometry {
  lefts: number[];
  widths: number[];
  separator: boolean;
}

const cache = new Map<string, ColumnGeometry | null>();

/**
 * The columns of a w:sectPr for a text area `textWidth` px wide; null for one column. Equal
 * columns share the width left by w:space (720 twips when missing, as Word reads it); a section
 * with its own widths (w:equalWidth="0") uses each w:col's w:w and w:space, scaled to the text area
 * when they do not add up to it (Word keeps them in step).
 */
export function columnGeometry(sectPr: string | null, textWidth: number): ColumnGeometry | null {
  const el = sectPr ? /<(?:[\w.-]+:)?cols\b([^>]*?)(\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?cols>)/.exec(sectPr) : null;
  if (!el || !(textWidth > 0)) return null;
  const key = `${el[0]}|${textWidth}`;
  const hit = cache.get(key);
  if (hit !== undefined) return hit;
  const attrs = el[1];
  const num = (a: string, s: string) => {
    const v = new RegExp(`\\bw:${a}="(-?\\d+(?:\\.\\d+)?)"`).exec(s)?.[1];
    return v == null ? null : Number(v);
  };
  const onOff = (a: string) => {
    const v = new RegExp(`\\bw:${a}="([^"]*)"`).exec(attrs)?.[1];
    return v == null ? null : v === '1' || v === 'true' || v === 'on';
  };
  const separator = onOff('sep') === true;
  const own = [...(el[3] ?? '').matchAll(/<(?:[\w.-]+:)?col\b([^>]*)\/?>/g)].map((m) => ({ w: num('w', m[1]) ?? 0, space: num('space', m[1]) ?? 0 }));
  let out: ColumnGeometry | null = null;
  if (onOff('equalWidth') === false && own.length > 1) {
    const total = own.reduce((s, c, i) => s + c.w + (i < own.length - 1 ? c.space : 0), 0);
    const k = total > 0 ? textWidth / twipsToPx(total) : 1;
    const lefts: number[] = [];
    const widths: number[] = [];
    let x = 0;
    own.forEach((c, i) => {
      lefts.push(x);
      widths.push(twipsToPx(c.w) * k);
      x += (twipsToPx(c.w) + (i < own.length - 1 ? twipsToPx(c.space) : 0)) * k;
    });
    out = { lefts, widths, separator };
  } else {
    const count = Math.min(45, Math.max(1, Math.trunc(num('num', attrs) ?? 1)));
    if (count > 1) {
      const space = twipsToPx(Math.max(0, num('space', attrs) ?? 720));
      const width = (textWidth - space * (count - 1)) / count;
      if (width > 8) out = { lefts: Array.from({ length: count }, (_, i) => i * (width + space)), widths: Array(count).fill(width), separator };
    }
  }
  if (cache.size > 200) cache.clear();
  cache.set(key, out);
  return out;
}

/** Where a column break (w:br w:type="column") sends the text: this block, the next one, or none. */
export function columnBreakIn(node: { forEach: (f: (c: any) => void) => void; type: { name: string } }): 'before' | 'after' | null {
  if (node.type.name !== 'paragraph') return null;
  let seenVisible = false;
  let at: 'before' | 'after' | null = null;
  node.forEach((c) => {
    if (at) return;
    if (c.type.name === 'hard_break' && c.attrs.type === 'column') {
      at = seenVisible ? 'after' : 'before';
      return;
    }
    if (!(c.type.name === 'raw_inline' && c.attrs.hidden)) seenVisible = true;
  });
  return at;
}

/** The lines between columns on a page (page px), drawn by drawColumnRules. */
export interface ColumnRule {
  x: number;
  top: number;
  bottom: number;
}

/** Draws a page's lines between columns (w:sep) as children of its element; old ones go. */
export function drawColumnRules(pageEl: HTMLElement, rules: ColumnRule[] | undefined): void {
  const old = Array.from(pageEl.querySelectorAll<HTMLElement>(':scope > .dx-col-rule'));
  const want = rules ?? [];
  want.forEach((r, i) => {
    const el = old[i] ?? pageEl.appendChild(Object.assign(document.createElement('div'), { className: 'dx-col-rule' }));
    const css = `left:${r.x}px;top:${r.top}px;height:${Math.max(0, r.bottom - r.top)}px`;
    if (el.style.cssText !== css) el.style.cssText = css;
  });
  for (const el of old.slice(want.length)) el.remove();
}
