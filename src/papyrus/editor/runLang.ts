// Languages of runs for screen readers (persona-300): the page is 中文 (lang="zh-TW" on the editor),
// and a stretch of Latin text whose run says another language (w:lang w:val="en-US" …) gets
// that lang, so it is read with an English voice. Only text without CJK characters: Word files
// set w:val="en-US" on nearly every run, Chinese text included (its language is w:eastAsia).
// Decorations only: nothing is written to the document.
//
// While an input method is composing (注音 typed at the end of an English run), the decorations
// are only mapped, growing over the text typed next to them: taking the lang off the text node
// being composed would redraw it and end the composition (Chrome, Safari). The blocks typed in
// are worked out again once the composition is over.

import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';

const CJK = /[⺀-鿿가-힯豈-﫿＀-￯]/;
const LANG = /<w:lang\b[^>]*?\bw:val="([A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*)"/;
const langCache = new Map<string, string | null>();

/** The Latin-text language a run's w:rPr names (null: none, or Chinese). */
export function runLang(rPr: string | null | undefined): string | null {
  if (!rPr || !rPr.includes('lang')) return null;
  let hit = langCache.get(rPr);
  if (hit === undefined) {
    const v = LANG.exec(rPr)?.[1] ?? null;
    hit = v && !/^zh\b/i.test(v) ? v : null;
    if (langCache.size > 2000) langCache.clear();
    langCache.set(rPr, hit);
  }
  return hit;
}

function blockDecos(block: PMNode, pos: number, out: Decoration[]): void {
  block.forEach((child, offset) => {
    if (!child.isText || CJK.test(child.text!)) return;
    const run = child.marks.find((m) => m.type.name === 'run');
    const lang = runLang(run?.attrs.rPr);
    // Inclusive: mapped during a composition, it covers the characters typed at its ends.
    if (lang) out.push(Decoration.inline(pos + 1 + offset, pos + 1 + offset + child.nodeSize, { lang }, { inclusiveStart: true, inclusiveEnd: true }));
  });
}

function build(doc: PMNode): DecorationSet {
  const out: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    blockDecos(node, pos, out);
    return false;
  });
  return out.length ? DecorationSet.create(doc, out) : DecorationSet.empty;
}

/** The ranges of tr.doc a transaction touched (null: too many to work out one by one). */
function touchedRanges(tr: Transaction): [number, number][] | null {
  const ranges: [number, number][] = [];
  tr.mapping.maps.forEach((map, i) => {
    for (const r of ranges) {
      r[0] = map.map(r[0], -1);
      r[1] = map.map(r[1], 1);
    }
    map.forEach((_a, _b, from, to) => ranges.push([from, to]));
    // Mark steps move nothing: their range, mapped through the later steps.
    const s = tr.steps[i] as unknown as { from?: number; to?: number };
    if (typeof s.from === 'number' && typeof s.to === 'number') ranges.push([map.map(s.from, -1), map.map(s.to, 1)]);
  });
  return ranges.length > 100 ? null : ranges;
}

/** The textblocks of `doc` around `ranges`. */
function blocksAt(doc: PMNode, ranges: [number, number][]): { node: PMNode; pos: number }[] {
  const seen = new Set<number>();
  const out: { node: PMNode; pos: number }[] = [];
  for (const [a, b] of ranges) {
    doc.nodesBetween(Math.max(0, Math.min(a, b) - 1), Math.min(doc.content.size, Math.max(a, b) + 1), (node, pos) => {
      if (!node.isTextblock) return true;
      if (!seen.has(pos)) {
        seen.add(pos);
        out.push({ node, pos });
      }
      return false;
    });
  }
  return out;
}

interface LangState {
  set: DecorationSet;
  /** Ranges typed in during a composition, to work out again after it (null: the whole document). */
  pending: [number, number][] | null | undefined;
}

const key = new PluginKey<LangState>('dx-run-lang');
/** Meta of the transaction that works the composed text out again once the composition ended. */
const FLUSH = 'dx-run-lang-flush';

/** Soon, once no composition is going on: a transaction that works the composed text out again (not an edit, no undo step). */
function flushSoon(view: EditorView, delay: number): void {
  setTimeout(() => {
    if (view.isDestroyed || view.composing || key.getState(view.state)?.pending === undefined) return;
    view.dispatch(view.state.tr.setMeta(FLUSH, true).setMeta('addToHistory', false));
  }, delay);
}

export function runLanguages(): Plugin<LangState> {
  return new Plugin<LangState>({
    key,
    state: {
      init: (_c, state) => ({ set: build(state.doc), pending: undefined }),
      apply(tr, old) {
        const flush = !!tr.getMeta(FLUSH);
        if (!tr.docChanged && !flush) return old;
        let ranges = tr.docChanged ? touchedRanges(tr) : [];
        let set = old.set.map(tr.mapping, tr.doc);
        let pending = old.pending;
        if (pending) pending = pending.map(([a, b]) => [tr.mapping.map(a, -1), tr.mapping.map(b, 1)] as [number, number]);
        if (tr.getMeta('composition')) {
          // Composing: only mapped (see the top of this file).
          const more = pending === null || ranges === null ? null : [...(pending ?? []), ...ranges];
          return { set, pending: more && more.length > 100 ? null : more };
        }
        if (pending !== undefined) ranges = pending === null || ranges === null ? null : [...pending, ...ranges];
        if (!ranges) return { set: build(tr.doc), pending: undefined };
        const add: Decoration[] = [];
        for (const { node, pos } of blocksAt(tr.doc, ranges)) {
          set = set.remove(set.find(pos, pos + node.nodeSize));
          blockDecos(node, pos, add);
        }
        return { set: add.length ? set.add(tr.doc, add) : set, pending: undefined };
      },
    },
    view: () => ({
      // A composition ended with a transaction (its last characters): work it out again.
      update: (view) => {
        if (!view.composing && key.getState(view.state)?.pending !== undefined) flushSoon(view, 0);
      },
    }),
    props: {
      decorations: (state: EditorState) => key.getState(state)?.set,
      handleDOMEvents: {
        // ProseMirror ends a composition a moment after compositionend (20 ms): check again after that.
        compositionend: (view) => {
          flushSoon(view, 50);
          return false;
        },
      },
    },
  });
}
