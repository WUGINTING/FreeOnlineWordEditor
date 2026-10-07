// The document as an AI assistant reads and changes it.
//
// Reading: the body's paragraphs (table cells included), numbered from 1 in document order, each
// as one line of text (assistantContext). Things that are not text keep a place in the line: a tab
// is a tab character, a line break 「↵」, a picture or field 「￼」; invisible markers (bookmarks,
// comment ranges, tracked deletions) are left out.
// A long document is not sent whole: the paragraphs the user is paying attention to (selected, on
// screen, at the cursor, just edited: attentionPlugin remembers where) with an outline of the rest.
//
// Changing: a list of actions about those numbers (AssistantAction), applied as ONE transaction
// through the body's own dispatch (applyAssistantActions), so they are one undo step, are recorded
// as tracked changes while 追蹤修訂 is on, and are refused by locked fields like any other edit.
// Nothing that is not text is ever removed by a text change: rewriting a paragraph replaces only
// the characters that differ, and pictures, fields and markers inside a replaced stretch stay.
// An action about a paragraph whose text is no longer what the assistant was shown is not applied.
//
// The actions come from a language model (the page's 小助手) or, later, from another program
// calling the same functions; who decides them is not this file's business.

import { Fragment, type Mark, type Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { schema } from './schema';
import { afterParagraphSplit } from './review';
import { documentOutline } from './outline';
import { splitParagraphPPr } from '../docx/revisions';
import { inheritedAlign, inheritedToggle, type StyleInheritance } from '../docx/inheritance';
import type { ParagraphStyleInfo } from '../docx/model';
import type { DocxEditor } from './core';

const TAB = '\t';
const BREAK = '↵';
const OBJECT = '￼';
/** The characters standing for things that are not text. */
const ATOMS = /[\t↵￼]/;
const ATOMS_ALL = /[\t↵￼]/g;

export interface AssistantParagraph {
  /** From 1, in document order. */
  n: number;
  text: string;
  /** Its style's name (標題 1 …), when it names one. */
  style: string | null;
}

export interface AssistantSelection {
  /** The first and last paragraph the selection touches. */
  from: number;
  to: number;
  text: string;
}

/** What the assistant is shown of the document. */
export interface AssistantContext {
  /** The paragraphs sent: all of them, or part of them (not necessarily next to each other), see assistantContext. */
  paragraphs: AssistantParagraph[];
  /** How many paragraphs the document has. */
  total: number;
  /** Null with only a cursor. */
  selection: AssistantSelection | null;
  /** The paragraph styles that can be applied, by name. */
  styles: string[];
  /** Paragraphs sent only in part (longer than the whole budget): never rewritten. */
  partial: number[];
  /** When only part of the document is sent: its headings and the paragraphs that start a part, as a map of the rest. */
  outline: OutlineEntry[];
  /**
   * Set when the paragraphs were picked by what the user is paying attention to (scope 'focus', and
   * the document did not fit): how many were taken because they are selected, asked for, on screen, just edited.
   */
  focus: { selected: number; asked: number; onScreen: number; edited: number } | null;
}

export interface OutlineEntry {
  n: number;
  /** 0 = a level 1 heading; null: not a heading, a paragraph that starts a part (主旨、一、…). */
  level: number | null;
  text: string;
}

export type AssistantAction =
  | { type: 'replace_text'; find: string; replace: string }
  | { type: 'set_paragraph_text'; paragraph: number; text: string }
  | { type: 'insert_paragraph'; after: number; text: string }
  | { type: 'delete_paragraphs'; from: number; to: number }
  | { type: 'format_paragraphs'; from: number; to: number; align?: 'left' | 'center' | 'right' | 'justify'; style?: string }
  | {
      type: 'format_text';
      from: number;
      to: number;
      text: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      color?: string;
    };

/** What became of one action: done (and what, in words) or not (and why). */
export interface ActionOutcome {
  ok: boolean;
  message: string;
}

// ----- reading -----

/** The body's paragraphs in document order, table cells included. */
export function bodyParagraphs(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    out.push({ node, pos });
    return false;
  });
  return out;
}

/** A stretch of a paragraph's text between two things that are not text. */
interface Segment {
  /** Offset in the paragraph's content where it starts. */
  start: number;
  text: string;
  /** The content offset of each character. */
  at: number[];
}

/** Markers nobody sees: the text goes on across them (as in search.ts). */
const invisible = (node: PMNode): boolean => node.type === schema.nodes.raw_inline && !!node.attrs.hidden;

/** A paragraph's text as segments, and the character standing for what ends each but the last. */
function segmentsOf(para: PMNode): { segments: Segment[]; atoms: string[] } {
  const segments: Segment[] = [{ start: 0, text: '', at: [] }];
  const atoms: string[] = [];
  para.forEach((child, offset) => {
    const seg = segments[segments.length - 1];
    if (child.isText) {
      seg.text += child.text!;
      for (let i = 0; i < child.text!.length; i++) seg.at.push(offset + i);
    } else if (!invisible(child)) {
      atoms.push(child.type === schema.nodes.tab ? TAB : child.type === schema.nodes.hard_break ? BREAK : OBJECT);
      segments.push({ start: offset + child.nodeSize, text: '', at: [] });
    }
  });
  return { segments, atoms };
}

/** A paragraph as the one line the assistant reads. */
export function paragraphLine(para: PMNode): string {
  const { segments, atoms } = segmentsOf(para);
  return segments.map((s, i) => s.text + (atoms[i] ?? '')).join('');
}

const styleName = (styles: ParagraphStyleInfo[], id: string | null): string | null =>
  (id && styles.find((s) => s.id === id)?.name) || null;

// ----- what the user is paying attention to -----
// Nobody watches the user's eyes; what they are busy with shows in four things the editor knows:
// what is selected, what is on screen, where the cursor is, and what was edited a moment ago.

interface Touched {
  /** A position in the paragraph that was edited (mapped through later edits). */
  pos: number;
  /** When (ms). */
  at: number;
}

const attentionKey = new PluginKey<Touched[]>('papyrus-assistant-attention');
/** Edits older than this no longer say where the user's mind is. */
export const RECENT_EDIT_MS = 10 * 60_000;
const MAX_TOUCHED = 30;

/**
 * Remembers where the body was edited lately (the editor has it among its plugins), for
 * assistantContext. Only the places: a few positions, mapped through the edits that follow.
 * Changes other plugins append (list numbers, page references) are not the user's doing.
 */
export function attentionPlugin(now: () => number = Date.now): Plugin<Touched[]> {
  return new Plugin<Touched[]>({
    key: attentionKey,
    state: {
      init: () => [],
      apply(tr, value) {
        if (!tr.docChanged) return value;
        const next = value.map((t) => ({ pos: tr.mapping.map(t.pos), at: t.at }));
        if (tr.getMeta('appendedTransaction')) return next;
        const at = now();
        // Where each step changed something, as positions in the document after the whole transaction;
        // a few are enough (全部取代 changes hundreds of places).
        const places: number[] = [];
        tr.steps.forEach((step, i) => {
          step.getMap().forEach((_from, _to, newFrom) => {
            if (places.length < 5) places.push(tr.mapping.slice(i + 1).map(newFrom));
          });
        });
        for (const pos of places) next.push({ pos, at });
        return next.filter((t) => at - t.at < RECENT_EDIT_MS).slice(-MAX_TOUCHED);
      },
    },
  });
}

/** The positions edited lately, latest first. */
function recentEdits(state: EditorState, now: number): number[] {
  const touched = attentionKey.getState(state) ?? [];
  return touched.filter((t) => now - t.at < RECENT_EDIT_MS).map((t) => t.pos).reverse();
}

/** The first and last document position on screen in the body; null when that cannot be told (not laid out). */
function visibleRange(view: EditorView): { from: number; to: number } | null {
  try {
    const dom = view.dom as HTMLElement;
    const rect = dom.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    let top = Math.max(rect.top, 0);
    let bottom = Math.min(rect.bottom, window.innerHeight || rect.bottom);
    // Less what the scrolling areas around it cut off.
    for (let el = dom.parentElement; el; el = el.parentElement) {
      if (!/(auto|scroll|hidden)/.test(getComputedStyle(el).overflowY)) continue;
      const r = el.getBoundingClientRect();
      top = Math.max(top, r.top);
      bottom = Math.min(bottom, r.bottom);
    }
    if (bottom - top < 16) return null;
    const x = rect.left + rect.width / 2;
    // Between two pages there is no text: a little further in, then.
    const probe = (y: number, step: number): number | null => {
      for (let i = 0; i < 12 && y > top && y < bottom; i++, y += step) {
        const hit = view.posAtCoords({ left: x, top: y });
        if (hit) return hit.pos;
      }
      return null;
    };
    const from = probe(top + 4, 24);
    const to = probe(bottom - 4, -24);
    return from == null || to == null ? null : { from: Math.min(from, to), to: Math.max(from, to) };
  } catch {
    return null;
  }
}

/** How 公文 mark their parts without heading styles: 主旨、說明、一、（一）、第三條 … */
const STRUCTURAL_START = /^\s*(主旨|說明|辦法|擬辦|依據|公告事項|會議名稱|討論事項|報告事項|決議|附件|正本|副本|[一二三四五六七八九十]{1,3}、|[（(][一二三四五六七八九十]{1,3}[）)]|第\s*[一二三四五六七八九十百\d]+\s*[條章節項])/;
const OUTLINE_MAX = 80;
const OUTLINE_TEXT = 30;

/** A map of the whole document for an assistant that is shown only part of it: its headings, and the paragraphs that start a part. */
function outlineOf(doc: PMNode, paras: { node: PMNode; pos: number }[], lines: AssistantParagraph[], styles?: StyleInheritance): OutlineEntry[] {
  const levels = new Map(documentOutline(doc, styles).map((h) => [h.pos, h.level]));
  const out: OutlineEntry[] = [];
  const short = (text: string) => (text.length > OUTLINE_TEXT ? text.slice(0, OUTLINE_TEXT) + '…' : text);
  paras.forEach((p, i) => {
    const text = lines[i].text.replace(ATOMS_ALL, ' ').trim();
    if (!text) return;
    const level = levels.get(p.pos);
    if (level != null) out.push({ n: i + 1, level, text: short(text) });
    else if (STRUCTURAL_START.test(text)) out.push({ n: i + 1, level: null, text: short(text) });
  });
  if (out.length <= OUTLINE_MAX) return out;
  // Too many: the headings first, then as many of the others as fit, in document order.
  const headings = out.filter((o) => o.level != null).slice(0, OUTLINE_MAX);
  const keep = new Set([...headings, ...out.filter((o) => o.level == null).slice(0, OUTLINE_MAX - headings.length)]);
  return out.filter((o) => keep.has(o));
}

/** A part of the document the user can name: from its heading (paragraph `n`) to the paragraph before the next part (`to`). */
export interface DocumentSection extends OutlineEntry {
  to: number;
}

/** The longest part that is sent because it was named (paragraphs). */
const SECTION_MAX = 60;

/**
 * The document's parts, for naming one to the assistant (「@說明」): each heading, and each
 * paragraph that starts a part as 公文 mark them, with where it ends. A heading's part goes up to
 * the next heading of its level or above; the others up to the next entry of any kind.
 */
export function documentSections(editor: DocxEditor): DocumentSection[] {
  const view = editor.view;
  if (!view) return [];
  const paras = bodyParagraphs(view.state.doc);
  const lines = paras.map((p, i): AssistantParagraph => ({ n: i + 1, text: paragraphLine(p.node), style: null }));
  const entries = outlineOf(view.state.doc, paras, lines, editor.model?.styles);
  return entries.map((entry, i) => {
    const next = entries.slice(i + 1).find((o) => entry.level == null || (o.level != null && o.level <= entry.level));
    return { ...entry, to: Math.min(next ? next.n - 1 : paras.length, entry.n + SECTION_MAX - 1) };
  });
}

export interface ContextOptions {
  /**
   * 'all' (the default): the whole document when it fits in `maxChars`, else what fits around the selection.
   * 'focus': only what the user is paying attention to, up to `focusChars`.
   */
  scope?: 'all' | 'focus';
  /** How much text 'focus' sends (the selection, and paragraphs asked for with `include`, may take it up to `maxChars`). */
  focusChars?: number;
  /** Paragraphs to send as well (the assistant asked to read them, or the user named them), from / to as numbers. */
  include?: { from: number; to: number }[];
  /** With only a cursor, the paragraph it is in counts as selected (a request made in the text is about that paragraph). */
  atCursor?: boolean;
  /** The positions on screen, for tests (as visibleRange finds them in a browser); null: unknown. */
  visible?: { from: number; to: number } | null;
  /** The time, for tests. */
  now?: number;
}

/**
 * What to show the assistant of the body.
 *
 * Scope 'all': every paragraph when their text fits in `maxChars`, otherwise those around the
 * selection (or the cursor), as many as fit.
 *
 * Scope 'focus': what the user is paying attention to, in this order while it fits in
 * `focusChars`: the selected paragraphs and those asked for (`include`; these two up to
 * `maxChars`), the paragraph with the cursor, the paragraphs on screen, those edited in the last
 * minutes (latest first), then the neighbours of all these. The paragraphs sent need not be next
 * to each other; an outline of the whole document goes along as a map. A document that fits in
 * `focusChars` anyway is sent whole.
 */
export function assistantContext(editor: DocxEditor, maxChars: number, options: ContextOptions = {}): AssistantContext {
  const view = editor.view;
  const styles = editor.model?.paragraphStyles ?? [];
  const out: AssistantContext = {
    paragraphs: [], total: 0, selection: null, styles: styles.map((s) => s.name).slice(0, 40), partial: [], outline: [], focus: null,
  };
  if (!view) return out;
  const { doc, selection } = view.state;
  const paras = bodyParagraphs(doc);
  const all = paras.map((p, i): AssistantParagraph => ({ n: i + 1, text: paragraphLine(p.node), style: styleName(styles, p.node.attrs.styleId) }));
  out.total = all.length;
  // The paragraph a position is in (or the last one before it).
  const indexAt = (pos: number) => {
    let lo = 0;
    let hi = paras.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (paras[mid].pos <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const first = indexAt(selection.from);
  const last = Math.max(first, indexAt(selection.empty ? selection.from : Math.max(selection.from, selection.to - 1)));
  if (!selection.empty && all.length) {
    out.selection = { from: first + 1, to: last + 1, text: doc.textBetween(selection.from, selection.to, '\n', '') };
  } else if (options.atCursor && all.length) {
    out.selection = { from: first + 1, to: first + 1, text: all[first].text };
  }
  const size = (p: AssistantParagraph) => p.text.length + 8;
  const whole = all.reduce((sum, p) => sum + size(p), 0);
  const focus = options.scope === 'focus';
  if (!all.length || whole <= (focus ? Math.min(maxChars, options.focusChars ?? maxChars) : maxChars)) {
    out.paragraphs = all;
    return out;
  }
  /** One paragraph longer than everything that may be sent: its beginning, and never rewritten. */
  const cut = (p: AssistantParagraph): AssistantParagraph => {
    if (p.text.length <= maxChars) return p;
    out.partial.push(p.n);
    return { ...p, text: p.text.slice(0, maxChars) + '…' };
  };
  out.outline = outlineOf(doc, paras, all, editor.model?.styles);

  if (!focus) {
    // Too long: from the selection outwards, both ways in turn, while the next one fits.
    let lo = first;
    let hi = Math.min(last, all.length - 1);
    let used = 0;
    for (let i = lo; i <= hi; i++) used += size(all[i]);
    if (used > maxChars) {
      // The selection alone does not fit: its beginning, the last one cut short.
      hi = lo;
      used = size(all[lo]);
      while (hi < last && used + size(all[hi + 1]) <= maxChars) used += size(all[++hi]);
    } else {
      for (let grew = true; grew; ) {
        grew = false;
        if (hi + 1 < all.length && used + size(all[hi + 1]) <= maxChars) {
          used += size(all[++hi]);
          grew = true;
        }
        if (lo > 0 && used + size(all[lo - 1]) <= maxChars) {
          used += size(all[--lo]);
          grew = true;
        }
      }
    }
    out.paragraphs = all.slice(lo, hi + 1).map(cut);
    return out;
  }

  // What the user is paying attention to.
  const budget = Math.min(maxChars, Math.max(200, options.focusChars ?? maxChars));
  const picked = new Set<number>();
  let used = 0;
  /** Takes paragraph `i` when it still fits in `limit`; the very first one is taken whatever its size. */
  const take = (i: number, limit: number): boolean => {
    if (i < 0 || i >= all.length) return false;
    if (picked.has(i)) return true;
    if (picked.size && used + size(all[i]) > limit) return false;
    picked.add(i);
    used += size(all[i]);
    return true;
  };
  const why = { selected: 0, asked: 0, onScreen: 0, edited: 0 };
  const count = (key: keyof typeof why, i: number, limit: number) => {
    const had = picked.has(i);
    if (take(i, limit) && !had) why[key]++;
  };
  if (out.selection) for (let i = first; i <= last; i++) count('selected', i, maxChars);
  for (const r of options.include ?? []) for (let n = Math.max(1, r.from); n <= Math.min(all.length, r.to); n++) count('asked', n - 1, maxChars);
  take(first, budget);
  const screen = options.visible !== undefined ? options.visible : visibleRange(view);
  if (screen) for (let i = indexAt(screen.from), end = indexAt(screen.to); i <= end; i++) count('onScreen', i, budget);
  const edited = [...new Set(recentEdits(view.state, options.now ?? Date.now()).map(indexAt))].slice(0, 5);
  for (const i of edited) count('edited', i, budget);
  // Their neighbours: a paragraph is read in the light of the ones around it. One step further at a time.
  for (let reach = 1, grew = true; reach <= 5 && grew; reach++) {
    grew = false;
    for (const i of [...picked].sort((a, b) => a - b)) {
      for (const j of [i + 1, i - 1]) {
        if (j < 0 || j >= all.length || picked.has(j)) continue;
        if (take(j, budget)) grew = true;
      }
    }
  }
  out.paragraphs = [...picked].sort((a, b) => a - b).map((i) => cut(all[i]));
  out.focus = why;
  return out;
}

/**
 * What a completion at the cursor goes on from: the text up to the cursor (the paragraphs before
 * it, at most `beforeChars` in all) and a little of what follows. Null unless the cursor (no
 * selection) is at the end of a paragraph that has some text.
 */
export function completionContext(editor: DocxEditor, beforeChars = 2000, afterChars = 300): { before: string; after: string; at: number } | null {
  const view = editor.view;
  if (!view || !editor.editable) return null;
  const { selection, doc } = view.state;
  if (!selection.empty) return null;
  const $pos = selection.$from;
  if ($pos.parent.type !== schema.nodes.paragraph || $pos.parentOffset !== $pos.parent.content.size) return null;
  const own = paragraphLine($pos.parent);
  if (own.replace(ATOMS_ALL, '').trim().length < 2) return null;
  const paras = bodyParagraphs(doc);
  const index = paras.findIndex((p) => p.pos === $pos.before());
  if (index < 0) return null;
  let before = own;
  for (let i = index - 1; i >= 0 && before.length < beforeChars; i--) before = paragraphLine(paras[i].node) + '\n' + before;
  let after = '';
  for (let i = index + 1; i < paras.length && after.length < afterChars; i++) after += (after ? '\n' : '') + paragraphLine(paras[i].node);
  return { before: before.slice(-beforeChars), after: after.slice(0, afterChars), at: $pos.pos };
}

/** Paragraph numbers as a short list of runs: 3～9、15. */
export function numberRuns(numbers: number[]): string {
  const runs: string[] = [];
  const sorted = [...numbers].sort((a, b) => a - b);
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    runs.push(j > i ? `${sorted[i]}～${sorted[j]}` : String(sorted[i]));
    i = j + 1;
  }
  return runs.join('、');
}

// ----- changing -----

/** Text as it may go into a paragraph: one line, without the characters that stand for other things. */
const plain = (text: string): string => text.replace(/\r\n?|\n/g, ' ').replace(ATOMS_ALL, ' ').replace(/[\u0000-\u001f\u007f]/g, '');

const STRUCTURAL = new Set(['inlineWrap', 'link', 'fieldResult']);
const NEXT_IS_NORMAL = /^(heading\d|title|subtitle)$/i;

/**
 * Replaces the text in [from, to) of one paragraph (positions in `tr.doc`) with `text`, in the
 * formatting of the first character replaced; what is not text in there (markers) stays, after it.
 */
function replaceChars(tr: Transaction, from: number, to: number, text: string): void {
  if (from === to) {
    if (text) tr.insert(from, schema.text(text, tr.doc.resolve(from).marks()));
    return;
  }
  const $from = tr.doc.resolve(from);
  const start = $from.start();
  let marks: readonly Mark[] | null = null;
  const kept: PMNode[] = [];
  $from.parent.forEach((child, offset) => {
    const cs = start + offset;
    const ce = cs + child.nodeSize;
    if (ce <= from || cs >= to) return;
    if (child.isText) marks ??= child.marks;
    else kept.push(child);
  });
  const nodes = [...(text ? [schema.text(text, marks ?? $from.marks())] : []), ...kept];
  if (nodes.length) tr.replaceWith(from, to, Fragment.fromArray(nodes));
  else tr.delete(from, to);
}

/** One place where two texts differ: characters [from, to) of the old one become `text`. */
export interface TextEdit {
  from: number;
  to: number;
  text: string;
}

/** Texts whose differing middles are longer than this (product of their lengths) are changed as one stretch. */
const DIFF_CELLS = 4_000_000;
/** Two changes with fewer than this many unchanged characters between them are one change: only those right next to each other. */
const DIFF_GAP = 1;

/**
 * The places where `text` differs from `old`, in order: a corrected character is one small change,
 * not the whole text from the first difference to the last. Character by character (a longest
 * common subsequence); changes right next to each other are one change. When less than half of the
 * text stays (a rewrite, not a correction), the whole differing stretch is one change: many
 * scattered characters that happen to match would not read as anything.
 */
export function textEdits(old: string, text: string): TextEdit[] {
  if (old === text) return [];
  let head = 0;
  while (head < old.length && head < text.length && old[head] === text[head]) head++;
  let tail = 0;
  while (tail < old.length - head && tail < text.length - head && old[old.length - 1 - tail] === text[text.length - 1 - tail]) tail++;
  const a = old.slice(head, old.length - tail);
  const b = text.slice(head, text.length - tail);
  const whole: TextEdit[] = [{ from: head, to: head + a.length, text: b }];
  const n = a.length;
  const m = b.length;
  if (!n || !m || n * m > DIFF_CELLS) return whole;
  // lcs[i][j]: the longest common subsequence of a[i..] and b[j..].
  const w = m + 1;
  const lcs = new Uint16Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i * w + j] = a[i] === b[j] ? lcs[(i + 1) * w + j + 1] + 1 : Math.max(lcs[(i + 1) * w + j], lcs[i * w + j + 1]);
    }
  }
  // A rewrite: less than half of the whole text stays (its unchanged beginning and end count).
  if ((head + tail + lcs[0]) * 2 < Math.max(old.length, text.length)) return whole;
  const edits: TextEdit[] = [];
  let open: TextEdit | null = null;
  /** Unchanged characters since the last change. */
  let kept = '';
  let i = 0;
  let j = 0;
  const change = (at: number, removed: number, added: string) => {
    if (open && kept.length < DIFF_GAP) {
      // Too close to the change before: one change, the characters between them included.
      open.text += kept + added;
      open.to = at + removed;
    } else {
      open = { from: at, to: at + removed, text: added };
      edits.push(open);
    }
    kept = '';
  };
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) {
      kept += a[i];
      i++;
      j++;
    } else if (j >= m || (i < n && lcs[(i + 1) * w + j] >= lcs[i * w + j + 1])) {
      change(head + i, 1, '');
      i++;
    } else {
      change(head + i, 0, b[j]);
      j++;
    }
  }
  return edits;
}

/** Makes `edits` (in order, positions in the segment's text) in `seg`, of the paragraph whose content starts at `base` in `tr.doc`. */
function editSegment(tr: Transaction, base: number, seg: Segment, edits: TextEdit[]): void {
  // Where character `i` of the segment is; past its last character: right after it.
  const posOf = (i: number) => base + (i < seg.at.length ? seg.at[i] : seg.at.length ? seg.at[seg.at.length - 1] + 1 : seg.start);
  // From the last to the first: the positions before a change stay as they were.
  for (const edit of [...edits].reverse()) {
    const from = edit.from < edit.to ? posOf(edit.from) : edit.from > 0 ? base + seg.at[edit.from - 1] + 1 : posOf(edit.from);
    const to = edit.from < edit.to ? base + seg.at[edit.to - 1] + 1 : from;
    replaceChars(tr, from, to, edit.text);
  }
}

/** Changes `seg` (of the paragraph whose content starts at `base` in `tr.doc`) into `text`: only the places that differ. */
function rewriteSegment(tr: Transaction, base: number, seg: Segment, text: string): boolean {
  const edits = textEdits(seg.text, text);
  editSegment(tr, base, seg, edits);
  return edits.length > 0;
}

/** Rewrites the paragraph at `pos` to read `line`; false when it already does. */
function rewriteParagraph(tr: Transaction, pos: number, line: string): boolean {
  const para = tr.doc.nodeAt(pos)!;
  const { segments } = segmentsOf(para);
  const parts = line.replace(/\r\n?|\n/g, ' ').split(ATOMS).map((t) => t.replace(/[\u0000-\u001f\u007f]/g, ''));
  let changed = false;
  if (parts.length === segments.length) {
    // The things that are not text are where they were: each stretch between them on its own, last first.
    for (let i = segments.length - 1; i >= 0; i--) changed = rewriteSegment(tr, pos + 1, segments[i], parts[i]) || changed;
    return changed;
  }
  // Some were added, dropped or moved in the new text: they all stay, and the text is changed as one stretch.
  const whole: Segment = { start: 0, text: segments.map((s) => s.text).join(''), at: segments.flatMap((s) => s.at) };
  return rewriteSegment(tr, pos + 1, whole, parts.join(''));
}

/** A new paragraph after the one at `pos`, as Enter at its end makes it; its position. */
function paragraphAfter(tr: Transaction, pos: number, text: string): number {
  const para = tr.doc.nodeAt(pos)!;
  const end = pos + para.nodeSize - 1;
  const heading = !!para.attrs.styleId && NEXT_IS_NORMAL.test(para.attrs.styleId);
  // Typed after Enter, the text would be formatted like the end of the paragraph before (not after a heading).
  const marks = heading ? [] : tr.doc.resolve(end).marks().filter((m) => !STRUCTURAL.has(m.type.name));
  const attrs: Record<string, unknown> = { ...para.attrs, pageBreakBefore: false, pAttrs: null, brGroup: null };
  attrs.pPr = splitParagraphPPr(attrs.pPr as string | null).last;
  if (heading) attrs.styleId = null;
  tr.split(end, 1, [{ type: schema.nodes.paragraph, attrs }]);
  afterParagraphSplit(tr, pos);
  const next = pos + tr.doc.nodeAt(pos)!.nodeSize;
  if (text) tr.insert(next + 1, schema.text(text, marks));
  return next;
}

/** A new paragraph before the one at `pos` (the document's first), formatted like it; its position (`pos`). */
function paragraphBefore(tr: Transaction, pos: number, text: string): number {
  const para = tr.doc.nodeAt(pos)!;
  const heading = !!para.attrs.styleId && NEXT_IS_NORMAL.test(para.attrs.styleId);
  const attrs = {
    ...para.attrs,
    pageBreakBefore: false,
    pAttrs: null,
    brGroup: null,
    sectPr: null,
    pPr: splitParagraphPPr(para.attrs.pPr as string | null).first,
    ...(heading ? { styleId: null } : {}),
  };
  tr.insert(pos, schema.nodes.paragraph.create(attrs, text ? schema.text(text) : null));
  return pos;
}

/** Where the occurrences of `text` are in the paragraph at `pos` (positions in `doc`); the whole content when `text` is empty. */
function textRanges(doc: PMNode, pos: number, text: string): { from: number; to: number }[] {
  const para = doc.nodeAt(pos)!;
  if (!text) return para.content.size ? [{ from: pos + 1, to: pos + 1 + para.content.size }] : [];
  const out: { from: number; to: number }[] = [];
  for (const seg of segmentsOf(para).segments) {
    for (let i = seg.text.indexOf(text); i >= 0; i = seg.text.indexOf(text, i + text.length)) {
      out.push({ from: pos + 1 + seg.at[i], to: pos + 1 + seg.at[i + text.length - 1] + 1 });
    }
  }
  return out;
}

class Refused extends Error {}

/**
 * What `actions` would do to the body: one transaction on the document as it is now (not sent to
 * the editor; null when none of them changes anything), and one outcome per action, in order.
 * `shown` is what the assistant was shown when it decided them: an action about a paragraph that
 * reads differently now (edited since, or renumbered) is left out, and the others still count.
 */
export function planAssistantActions(
  editor: DocxEditor, actions: AssistantAction[], shown: AssistantContext,
): { tr: Transaction | null; outcomes: ActionOutcome[] } {
  const view = editor.view;
  if (!view || !editor.editable) return { tr: null, outcomes: actions.map(() => ({ ok: false, message: '文件是唯讀的，無法修改。' })) };
  const state = view.state;
  let tr = state.tr;
  const styles: StyleInheritance | undefined = editor.model?.styles;
  const styleList = editor.model?.paragraphStyles ?? [];
  const original = bodyParagraphs(state.doc);
  const shownText = new Map(shown.paragraphs.map((p) => [p.n, p.text]));
  const partial = new Set(shown.partial);
  /** New paragraphs put after paragraph n so far: the next one goes after the last of them. */
  const tails = new Map<number, { pos: number; steps: number }>();
  /** Paragraphs an earlier action removed, and how many steps there were before that. */
  const gone = new Map<number, number>();

  /** Paragraph `n` as it is in `tr.doc` now. */
  const locate = (n: number): number => {
    const p = original[n - 1];
    if (!p || !shownText.has(n)) throw new Refused(`沒有第 ${n} 段。`);
    if (!partial.has(n) && paragraphLine(p.node) !== shownText.get(n)) throw new Refused(`第 ${n} 段在小助手讀過之後改過了，請再問一次。`);
    // After what was put right before it (a new first paragraph). Not mapResult().deleted: changing a
    // paragraph's attributes replaces its opening token, which counts as deleted there.
    const pos = tr.mapping.map(p.pos, 1);
    if (gone.has(n) || tr.doc.nodeAt(pos)?.type !== schema.nodes.paragraph) throw new Refused(`第 ${n} 段已經不在了。`);
    return pos;
  };
  const range = (from: number, to: number): number[] => {
    if (!(from >= 1 && to >= from)) throw new Refused('段落範圍不對。');
    const out: number[] = [];
    for (let n = from; n <= to; n++) out.push(n);
    return out;
  };

  const run = (action: AssistantAction): string => {
    switch (action.type) {
      case 'replace_text': {
        const find = action.find;
        if (!find || find === action.replace) throw new Refused('沒有要取代的文字。');
        // Exactly as written: 「台北」 does not find 「臺北」, 114 does not find １１４. Where it is found, only the
        // characters that differ between the two are changed: 「其二為癮蔽的邏輯錯誤」 to 「其二為隱蔽的邏輯錯誤」 is one character.
        const inner = textEdits(find, plain(action.replace));
        const paras = bodyParagraphs(tr.doc);
        let found = 0;
        // Last paragraph first, last stretch first: the positions before a change stay as they were.
        for (let k = paras.length - 1; k >= 0; k--) {
          const { segments } = segmentsOf(paras[k].node);
          for (let s = segments.length - 1; s >= 0; s--) {
            const edits: TextEdit[] = [];
            for (let i = segments[s].text.indexOf(find); i >= 0; i = segments[s].text.indexOf(find, i + find.length)) {
              found++;
              for (const e of inner) edits.push({ from: i + e.from, to: i + e.to, text: e.text });
            }
            if (edits.length) editSegment(tr, paras[k].pos + 1, segments[s], edits);
          }
        }
        if (!found) throw new Refused(`文件裡找不到「${find}」。`);
        return `已取代 ${found} 處`;
      }
      case 'set_paragraph_text': {
        if (partial.has(action.paragraph)) throw new Refused(`第 ${action.paragraph} 段太長，小助手只讀了一部分，不能整段改寫。`);
        return rewriteParagraph(tr, locate(action.paragraph), action.text) ? `已修改第 ${action.paragraph} 段` : `第 ${action.paragraph} 段沒有變化`;
      }
      case 'insert_paragraph': {
        const lines = action.text.split(/\r\n?|\n/).map(plain).slice(0, 200);
        const tail = tails.get(action.after);
        let at: number | null = tail ? tr.mapping.slice(tail.steps).map(tail.pos, 1) : null;
        for (const line of lines) {
          if (at != null) at = paragraphAfter(tr, at, line);
          else if (action.after === 0) {
            if (!original.length) throw new Refused('文件沒有段落。');
            at = paragraphBefore(tr, tr.mapping.map(original[0].pos, 1), line);
          } else at = paragraphAfter(tr, locate(action.after), line);
        }
        if (at != null) tails.set(action.after, { pos: at, steps: tr.steps.length });
        return lines.length > 1 ? `已新增 ${lines.length} 段` : '已新增一段';
      }
      case 'delete_paragraphs': {
        const targets = range(action.from, action.to).map((n) => ({ n, pos: locate(n) }));
        for (const { n, pos } of targets.reverse()) {
          const para = tr.doc.nodeAt(pos)!;
          // The only paragraph of its cell (or of the document), or the one ending a section: emptied, not removed.
          if (tr.doc.resolve(pos).parent.childCount > 1 && !para.attrs.sectPr) {
            gone.set(n, tr.steps.length);
            tr.delete(pos, pos + para.nodeSize);
          } else if (para.content.size) tr.delete(pos + 1, pos + 1 + para.content.size);
        }
        return targets.length > 1 ? `已刪除 ${targets.length} 段` : `已刪除第 ${action.from} 段`;
      }
      case 'format_paragraphs': {
        let styleId: string | null | undefined;
        if (action.style) {
          const wanted = action.style.trim().toLowerCase();
          const found = styleList.find((s) => s.name.trim().toLowerCase() === wanted);
          if (!found) throw new Refused(`這份文件沒有「${action.style}」樣式。`);
          // The default style is a paragraph naming none (as the ribbon's style box does).
          styleId = found.name.toLowerCase() === 'normal' ? null : found.id;
        }
        let changed = 0;
        for (const pos of range(action.from, action.to).map(locate)) {
          const para = tr.doc.nodeAt(pos)!;
          const attrs: Record<string, unknown> = { ...para.attrs };
          if (styleId !== undefined) attrs.styleId = styleId;
          if (action.align) {
            const inherited = inheritedAlign(styles, (attrs.styleId as string | null) ?? null);
            attrs.align = action.align === inherited ? null : action.align;
          }
          if (attrs.styleId === para.attrs.styleId && attrs.align === para.attrs.align) continue;
          tr.setNodeMarkup(pos, undefined, attrs, para.marks);
          changed++;
        }
        return changed ? `已調整 ${changed} 段的格式` : '格式已經是這樣了';
      }
      case 'format_text': {
        const part = action.from === action.to ? action.text : '';
        let ranges = 0;
        for (const pos of range(action.from, action.to).map(locate)) {
          const para = tr.doc.nodeAt(pos)!;
          const found = textRanges(tr.doc, pos, part);
          if (part && !found.length) throw new Refused(`第 ${action.from} 段找不到「${part}」。`);
          for (const { from, to } of found) {
            for (const name of ['bold', 'italic'] as const) {
              const want = action[name];
              if (want == null) continue;
              tr.removeMark(from, to, schema.marks[name]);
              // Only said where the styles do not already make it so.
              if (want !== inheritedToggle(styles, name, para.attrs.styleId ?? null, null)) tr.addMark(from, to, schema.marks[name].create({ on: want }));
            }
            if (action.underline === true) tr.addMark(from, to, schema.marks.underline.create());
            if (action.underline === false) tr.removeMark(from, to, schema.marks.underline);
            if (action.fontSize) tr.addMark(from, to, schema.marks.fontSize.create({ pt: action.fontSize }));
            if (action.color && /^[0-9a-f]{6}$/i.test(action.color)) tr.addMark(from, to, schema.marks.color.create({ color: '#' + action.color.toUpperCase() }));
            ranges++;
          }
        }
        return ranges ? '已調整文字格式' : '沒有文字可以調整';
      }
      default:
        throw new Refused('小助手不會這種修改。');
    }
  };

  const outcomes: ActionOutcome[] = [];
  for (const action of actions) {
    const steps = tr.steps.length;
    try {
      outcomes.push({ ok: true, message: run(action) });
    } catch (err) {
      // What this action already did is taken back (the ones before it stay): the steps up to it, made again.
      if (tr.steps.length > steps) {
        const again = state.tr;
        for (const step of tr.steps.slice(0, steps)) again.step(step);
        tr = again;
        for (const [n, at] of gone) if (at >= steps) gone.delete(n);
        for (const [n, tail] of tails) if (tail.steps > steps) tails.delete(n);
      }
      outcomes.push({ ok: false, message: err instanceof Refused ? err.message : '這項修改無法套用。' });
    }
  }
  return { tr: tr.docChanged ? tr : null, outcomes };
}

/**
 * Applies `actions` to the body (see planAssistantActions), all in one undo step. One outcome
 * per action, in order.
 */
export function applyAssistantActions(editor: DocxEditor, actions: AssistantAction[], shown: AssistantContext): ActionOutcome[] {
  const { tr, outcomes } = planAssistantActions(editor, actions, shown);
  const view = editor.view;
  if (!tr || !view) return outcomes;
  const before = view.state.doc;
  view.dispatch(closeHistory(tr));
  // Refused as a whole (a locked field, or something 追蹤修訂 cannot record): the editor said why.
  if (view.state.doc === before) return outcomes.map((o) => (o.ok ? { ok: false, message: '文件沒有接受這項修改（例如鎖定的欄位，或追蹤修訂無法記錄的變更）。' } : o));
  return outcomes;
}

/**
 * An action in words, for the user to decide on before it is applied. `shown`: what the assistant
 * was shown, so a rewritten paragraph can be said as what changes in it.
 */
export function describeAction(action: AssistantAction, shown?: AssistantContext): string {
  const span = (from: number, to: number) => (from === to ? `第 ${from} 段` : `第 ${from}～${to} 段`);
  const quote = (text: string, max = 60) => `「${text.length > max ? text.slice(0, max) + '…' : text}」`;
  switch (action.type) {
    case 'replace_text': {
      // A long phrase replaced by a corrected one: the characters that change, with a little of what is around each.
      const edits = action.find.length > 6 ? textEdits(action.find, action.replace) : [];
      if (edits.length && edits.length <= 4 && edits.every((e) => e.to - e.from <= 10 && e.text.length <= 10)) {
        const around = (e: TextEdit, middle: string) => `「${action.find.slice(Math.max(0, e.from - 1), e.from)}${middle}${action.find.slice(e.to, e.to + 1)}」`;
        return `${quote(action.find, 30)}裡的${edits.map((e) => `${around(e, action.find.slice(e.from, e.to))}改成${around(e, e.text)}`).join('、')}`;
      }
      return `把所有的${quote(action.find, 30)}改成${quote(action.replace, 30)}`;
    }
    case 'set_paragraph_text': {
      // A correction is said as the few places it changes (with a little of what is around each), not as the whole new paragraph.
      const before = shown?.paragraphs.find((p) => p.n === action.paragraph)?.text;
      const edits = before == null ? [] : textEdits(before, action.text.replace(/\r\n?|\n/g, ' '));
      if (before != null && edits.length && edits.length <= 5 && edits.every((e) => e.to - e.from <= 20 && e.text.length <= 20)) {
        const around = (e: TextEdit, middle: string) => `「${before.slice(Math.max(0, e.from - 2), e.from)}${middle}${before.slice(e.to, e.to + 2)}」`;
        return `第 ${action.paragraph} 段：${edits.map((e) => `${around(e, before.slice(e.from, e.to))}→${around(e, e.text)}`).join('、')}`;
      }
      return `第 ${action.paragraph} 段改為：${quote(action.text, 120)}`;
    }
    case 'insert_paragraph':
      return `${action.after === 0 ? '在最前面' : `在第 ${action.after} 段後面`}新增：${quote(action.text, 120)}`;
    case 'delete_paragraphs':
      return `刪除${span(action.from, action.to)}`;
    case 'format_paragraphs': {
      const ALIGN = { left: '靠左對齊', center: '置中', right: '靠右對齊', justify: '左右對齊' };
      const what = [action.style ? `套用「${action.style}」樣式` : '', action.align ? ALIGN[action.align] : ''].filter(Boolean).join('、');
      return `${span(action.from, action.to)}${what}`;
    }
    case 'format_text': {
      const what = [
        action.bold == null ? '' : action.bold ? '粗體' : '取消粗體',
        action.italic == null ? '' : action.italic ? '斜體' : '取消斜體',
        action.underline == null ? '' : action.underline ? '底線' : '取消底線',
        action.fontSize ? `${action.fontSize} 點` : '',
        action.color ? `顏色 #${action.color}` : '',
      ].filter(Boolean).join('、');
      const where = action.from === action.to && action.text ? `第 ${action.from} 段的${quote(action.text, 30)}` : span(action.from, action.to);
      return `${where}設為${what}`;
    }
    default:
      return '（不支援的修改）';
  }
}
