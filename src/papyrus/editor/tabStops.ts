// Right-aligned tab stops (persona-300, 目錄): in a paragraph with a right tab stop (w:tab
// w:val="right", as a table of contents line has one at the right margin), the text after its last
// tab ends at that stop, and the gap is drawn with the stop's leader (dots …), as Word shows it.
// Every other tab keeps the editor's fixed width. The widths are measured on the page and set as
// decorations on the tab nodes, so nothing about the document changes.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { twipsToPx } from '../units';
import './tabStops.css';

/** A paragraph's right tab stop: where it is (twips from the text column's left edge) and its leader. */
export interface RightStop {
  pos: number;
  leader: string | null;
}

const TAB = /<(?:[\w.-]+:)?tab\b([^>]*)\/?>/g;
const cache = new Map<string, RightStop | null>();

/** The rightmost right-aligned tab stop in a paragraph's w:pPr (none cleared). Null when there is none. */
export function rightStop(pPr: string | null): RightStop | null {
  if (!pPr || !/<(?:[\w.-]+:)?tabs\b/.test(pPr)) return null;
  const hit = cache.get(pPr);
  if (hit !== undefined) return hit;
  let best: RightStop | null = null;
  const tabs = /<(?:[\w.-]+:)?tabs\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?tabs>/.exec(pPr)?.[1] ?? '';
  for (const m of tabs.matchAll(TAB)) {
    const attrs = m[1];
    const val = /\bw:val="([^"]*)"/.exec(attrs)?.[1];
    const pos = Number(/\bw:pos="(-?\d+)"/.exec(attrs)?.[1]);
    if ((val !== 'right' && val !== 'end') || !Number.isFinite(pos)) continue;
    const leader = /\bw:leader="([^"]*)"/.exec(attrs)?.[1] ?? null;
    if (!best || pos > best.pos) best = { pos, leader: leader === 'none' ? null : leader };
  }
  if (cache.size > 500) cache.clear();
  cache.set(pPr, best);
  return best;
}

export const tabStopsKey = new PluginKey<DecorationSet>('dx-tab-stops');

/** The last tab of each paragraph with a right tab stop: its position and the stop. */
function candidates(doc: PMNode): { pos: number; para: number; stop: RightStop }[] {
  const out: { pos: number; para: number; stop: RightStop }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    const stop = rightStop(node.attrs.pPr as string | null);
    if (!stop) return false;
    let last = -1;
    node.forEach((c, offset) => {
      if (c.type.name === 'tab') last = pos + 1 + offset;
    });
    if (last >= 0) out.push({ pos: last, para: pos, stop });
    return false;
  });
  return out;
}

const LEADER_CLASS: Record<string, string> = { dot: 'dx-tab-dot', middleDot: 'dx-tab-dot', hyphen: 'dx-tab-hyphen', underscore: 'dx-tab-line', heavy: 'dx-tab-line' };

/**
 * Widths for the tabs (px, unzoomed), measured on the page as it is, with each tab at the
 * editor's own width: a width set here never feeds into the next measurement, so a width that
 * makes a line wrap cannot take itself away again and again (the page would keep resizing).
 */
export function measureTabs(view: EditorView): Map<number, { width: number; leader: string | null }> {
  const out = new Map<number, { width: number; leader: string | null }>();
  const root = view.dom as HTMLElement;
  const rootRect = root.getBoundingClientRect();
  const scale = root.offsetWidth > 0 && rootRect.width > 0 ? rootRect.width / root.offsetWidth : 1;
  const range = document.createRange();
  const found: { c: ReturnType<typeof candidates>[number]; tab: HTMLElement; para: HTMLElement; set: string }[] = [];
  for (const c of candidates(view.state.doc)) {
    const tab = view.nodeDOM(c.pos) as HTMLElement | null;
    const para = view.nodeDOM(c.para) as HTMLElement | null;
    if (!(tab instanceof HTMLElement) || !(para instanceof HTMLElement) || !para.contains(tab)) continue;
    found.push({ c, tab, para, set: tab.style.width });
  }
  // (A tab is a leaf node: the editor ignores this change to its element.)
  for (const f of found) if (f.set) f.tab.style.width = '';
  try {
    for (const { c, tab, para } of found) {
      const pr = para.getBoundingClientRect();
      const cs = getComputedStyle(para);
      const px = (v: string) => (parseFloat(v) || 0) * scale;
      // Tab stops are measured from the text column's left edge, not from the paragraph's indent;
      // a section's own margins or text column (its padding, see pagination.ts) move that edge.
      const columnLeft = pr.left + px(cs.paddingLeft) - px(cs.marginLeft);
      // Never past the paragraph's right edge (a right indent, a table cell, a column): the text
      // after the tab would wrap.
      const contentRight = pr.right - px(cs.paddingRight) - px(cs.borderRightWidth);
      const target = Math.min(columnLeft + twipsToPx(c.stop.pos) * scale, contentRight);
      const tr = tab.getBoundingClientRect();
      // What follows the tab to the end of the paragraph (the page number), when it is on the same line.
      range.setStartAfter(tab);
      range.setEnd(para, para.childNodes.length);
      const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);
      let trailing = 0;
      if (rects.length) {
        // (Measured from the tab's line: the text after it may have wrapped to the next one whole.)
        const top = Math.min(tr.top, ...rects.map((r) => r.top));
        if (rects.some((r) => r.top > top + (r.height || 1) * 0.8)) continue; // it wraps: leave the tab alone
        trailing = Math.max(...rects.map((r) => r.right)) - Math.min(...rects.map((r) => r.left));
      }
      // Rounded down, so the text after it never ends past the edge by a fraction of a pixel.
      const width = Math.floor(((target - trailing - tr.left) / scale) * 2) / 2;
      out.set(c.pos, { width: Math.max(4, width), leader: c.stop.leader });
    }
  } finally {
    for (const f of found) if (f.set) f.tab.style.width = f.set;
  }
  return out;
}

/**
 * Keeps each right tab's width so the text after it ends at its tab stop. Measured after the
 * document changed or the page was laid out again (the pagination plugin's spacers move lines,
 * not their widths), and redrawn only when a width changed.
 */
export function tabStops(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: tabStopsKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set) {
        const next = tr.getMeta(tabStopsKey) as DecorationSet | undefined;
        if (next) return next;
        return tr.docChanged ? set.map(tr.mapping, tr.doc) : set;
      },
    },
    props: { decorations: (state) => tabStopsKey.getState(state) },
    view(view) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      let shown = '';
      const run = () => {
        timer = undefined;
        if (view.isDestroyed) return;
        const widths = measureTabs(view);
        const key = JSON.stringify([...widths]);
        if (key === shown) return;
        shown = key;
        const decos = [...widths].map(([pos, w]) =>
          Decoration.node(pos, pos + 1, {
            class: `dx-tab-right ${w.leader ? LEADER_CLASS[w.leader] ?? '' : ''}`.trim(),
            style: `width:${w.width}px`,
          }),
        );
        view.dispatch(view.state.tr.setMeta(tabStopsKey, DecorationSet.create(view.state.doc, decos)).setMeta('addToHistory', false));
      };
      const soon = () => {
        clearTimeout(timer);
        timer = setTimeout(run, 60);
      };
      // Late fonts and pictures change where the text before a tab ends.
      const resize = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(soon) : null;
      resize?.observe(view.dom);
      document.fonts?.ready.then(soon);
      soon();
      return {
        update: (v, prev) => {
          if (prev.doc !== v.state.doc) soon();
        },
        destroy: () => {
          clearTimeout(timer);
          resize?.disconnect();
        },
      };
    },
  });
}
