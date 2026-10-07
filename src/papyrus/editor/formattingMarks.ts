import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import { AttrStep } from 'prosemirror-transform';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';

/**
 * 顯示／隱藏編輯標記 (Word's Show/Hide ¶, Home › 段落, Ctrl+Shift+8): Word's non-printing
 * formatting marks. ¶ ends every paragraph (table cells and headers/footers too), · is a space,
 * ° a non-breaking space, □ a full-width (ideographic) space, → a tab and ↵ a manual line break.
 * Page and section breaks and column breaks keep the labels they always show.
 *
 * The marks are what the viewer chose to see, not part of the document: they are decorations
 * only (no step, no undo entry, nothing saved) and each glyph is CSS generated content drawn
 * over the text (editor.css), so it takes no room — lines break and pages end exactly as with
 * the marks hidden, which pagination relies on, since it measures the text. Generated content
 * is also left out of copied text, and the widgets are aria-hidden, so screen readers read the
 * text as usual. The printout and the PDF never show them (stripFormattingMarks), as Word does
 * not print them by default.
 *
 * Cost: a span per space makes the browser's layout slow once there are tens of thousands of
 * them (measured on 200 pages of mixed Chinese/English text: ~120 ms per keystroke instead of
 * ~8 ms, and 3 s to show them). Spaces are therefore marked only in a window around what is on
 * screen (MarksWindow; see visibleWindow), which follows the scrolling; ¶, → and ↵ (one or two
 * per paragraph) are marked everywhere. Edits re-mark only the paragraphs they touch.
 *
 * Known differences from Word: Word also shows ¤ at the end of each table cell and row (here a
 * cell's last paragraph shows ¶), and draws ¶ in the paragraph mark's own font and size (here
 * the paragraph's, i.e. its style's).
 */

/** What a mark stands for (the decoration's spec.kind). */
export type MarkKind = 'paragraph' | 'space' | 'nbsp' | 'ideographicSpace' | 'tab' | 'lineBreak';

/** The document range whose paragraphs have their spaces marked. */
export interface MarksWindow {
  from: number;
  to: number;
}

/** The plugin's state: whether the marks are shown, and the decorations drawing them. */
export interface FormattingMarksState {
  on: boolean;
  set: DecorationSet;
  /** How many paragraphs the last transaction (re)decorated: all on showing, few while typing. */
  rebuilt: number;
  /** Where spaces are marked; null: the whole document. */
  window: MarksWindow | null;
}

/**
 * A transaction's formattingMarksKey meta: show or hide the marks (`on`), and/or move the window
 * where spaces are marked (`window`; null for the whole document).
 */
export interface FormattingMarksMeta {
  on?: boolean;
  window?: MarksWindow | null;
}

/**
 * The window a view starts with when the marks are already on (the remembered choice): the start
 * of the document, which is what a newly opened document shows, until the view is laid out.
 */
const INITIAL_WINDOW = 30000;

export const formattingMarksKey = new PluginKey<FormattingMarksState>('dx-formatting-marks');

/** Where the choice is remembered (per browser, like Word keeps it per installation). */
export const SHOW_MARKS_STORAGE = 'papyrus.showMarks';

/** The class that turns the glyphs on, on an editor's .dx-doc (or a static header/footer copy). */
export const SHOW_MARKS_CLASS = 'dx-show-marks';

/** The class of each kind of space, and its kind. */
const SPACES: Record<string, { cls: string; kind: MarkKind }> = {
  ' ': { cls: 'dx-mk-sp', kind: 'space' },
  '\u00a0': { cls: 'dx-mk-nbsp', kind: 'nbsp' },
  '\u3000': { cls: 'dx-mk-idsp', kind: 'ideographicSpace' },
};
/** The spaces that get a mark (one shared pattern; `g`, so reset lastIndex before a loop). */
const SPACE_RE = /[ \u00a0\u3000]/g;
/** Whether `text` has a space that gets a mark. */
function hasSpace(text: string): boolean {
  SPACE_RE.lastIndex = 0;
  return SPACE_RE.test(text);
}
/** Classes that only mark text (unwrapped or dropped in the printout). */
const TEXT_MARK_CLASSES = ['dx-mk-sp', 'dx-mk-nbsp', 'dx-mk-idsp'];

/**
 * An empty, zero-width span for ¶ or ↵: the glyph is its ::after (editor.css). Screen readers
 * skip it (aria-hidden); ProseMirror makes widgets non-editable.
 */
function markElement(doc: Document, cls: string): HTMLElement {
  const span = doc.createElement('span');
  span.className = `dx-mk ${cls}`;
  span.setAttribute('aria-hidden', 'true');
  return span;
}
// One function per kind, so redrawn widgets are recognised as the same and their DOM is kept.
const paragraphMark = (view: EditorView) => markElement(view.dom.ownerDocument, 'dx-mk-p');
const breakMark = (view: EditorView) => markElement(view.dom.ownerDocument, 'dx-mk-br');

/**
 * Whether Word shows ¶ at the end of this textblock. Not where the paragraph ends a section (Word
 * shows the section break mark there, see sectionMarks.ts), nor on a piece of a paragraph that
 * a page break splits when the paragraph goes on after the break (its mark is at the end of the
 * last piece; see paragraph.brGroup).
 */
function endsParagraph(block: PMNode, parent: PMNode | null, index: number): boolean {
  if (block.attrs.sectPr) return false;
  const next = parent?.maybeChild(index + 1);
  return !(block.attrs.brGroup && next?.type.name === 'page_break' && next.attrs.group === block.attrs.brGroup);
}

/** Whether the block at `pos` has its spaces marked. */
function inWindow(pos: number, block: PMNode, window: MarksWindow | null): boolean {
  return !window || (pos < window.to && pos + block.nodeSize > window.from);
}

/** The marks of one textblock at `pos` (its spaces only when `spaces`), added to `out`. */
function blockMarks(block: PMNode, pos: number, parent: PMNode | null, index: number, spaces: boolean, out: Decoration[]): void {
  block.forEach((child, offset) => {
    const at = pos + 1 + offset;
    if (child.isText) {
      if (!spaces) return;
      const text = child.text!;
      SPACE_RE.lastIndex = 0;
      for (let m = SPACE_RE.exec(text); m; m = SPACE_RE.exec(text)) {
        const space = SPACES[m[0]];
        out.push(Decoration.inline(at + m.index, at + m.index + 1, { class: space.cls }, { kind: space.kind }));
      }
    } else if (child.type.name === 'tab') {
      out.push(Decoration.node(at, at + 1, { class: 'dx-mk-tab' }, { kind: 'tab' }));
    } else if (child.type.name === 'hard_break' && child.attrs.type !== 'column') {
      // Before the break, so ↵ ends the line it breaks (a column break shows its 分欄符號 label).
      out.push(Decoration.widget(at, breakMark, { side: -1, key: 'dx-mk-br', kind: 'lineBreak', ignoreSelection: true }));
    }
  });
  if (endsParagraph(block, parent, index)) {
    // After the last character, and after the cursor at the paragraph's end.
    const end = pos + block.nodeSize - 1;
    out.push(Decoration.widget(end, paragraphMark, { side: 1, key: 'dx-mk-p', kind: 'paragraph', ignoreSelection: true }));
  }
}

/** The marks of the whole document (spaces in `window`). */
function decorateAll(doc: PMNode, window: MarksWindow | null): FormattingMarksState {
  const out: Decoration[] = [];
  let rebuilt = 0;
  doc.descendants((node, pos, parent, index) => {
    if (!node.isTextblock) return true;
    blockMarks(node, pos, parent, index, inWindow(pos, node, window), out);
    rebuilt++;
    return false;
  });
  return { on: true, set: DecorationSet.create(doc, out), rebuilt, window };
}

type Blocks = Map<number, { node: PMNode; parent: PMNode | null; index: number }>;

/** The textblocks between `from` and `to` (clamped to the document) that `keep`, added to `blocks`. */
function collectBlocks(doc: PMNode, from: number, to: number, blocks: Blocks, keep?: (pos: number, node: PMNode) => boolean): void {
  doc.nodesBetween(Math.max(0, from), Math.min(doc.content.size, to), (node, pos, parent, index) => {
    if (!node.isTextblock) return true;
    if (!keep || keep(pos, node)) blocks.set(pos, { node, parent, index });
    return false;
  });
}

/** `set` with the marks of `blocks` made again (their spaces when in `window`). */
function remark(doc: PMNode, set: DecorationSet, blocks: Blocks, window: MarksWindow | null): DecorationSet {
  if (!blocks.size) return set;
  const stale: Decoration[] = [];
  const fresh: Decoration[] = [];
  for (const [pos, b] of blocks) {
    stale.push(...set.find(pos, pos + b.node.nodeSize));
    blockMarks(b.node, pos, b.parent, b.index, inWindow(pos, b.node, window), fresh);
  }
  return set.remove(stale).add(doc, fresh);
}

/** Above this many steps in one transaction, the whole document is marked again. */
const MAX_STEPS = 200;

/** The ranges of `tr.doc` a transaction changed (replaced content, changed node attributes). */
function changedRanges(tr: Transaction): [number, number][] {
  const out: [number, number][] = [];
  tr.steps.forEach((step, i) => {
    const rest = tr.mapping.slice(i + 1);
    tr.mapping.maps[i].forEach((_oldFrom, _oldTo, from, to) => out.push([rest.map(from, -1), rest.map(to, 1)]));
    // setNodeAttribute changes no positions, but the node may now show a mark or not.
    if (step instanceof AttrStep) {
      const at = rest.map(step.pos, 1);
      out.push([at, at + 1]);
    }
  });
  return out;
}

/**
 * The marks after an edit: the previous ones mapped, with the paragraphs the edit touched (and
 * their neighbours, whose ¶ may depend on what follows them) decorated again. Typing in a long
 * document redoes one paragraph, not all of them.
 */
function decorateChanges(tr: Transaction, value: FormattingMarksState): FormattingMarksState {
  const doc = tr.doc;
  // The window keeps covering the same text (and grows with what is typed at its ends).
  const w = value.window;
  const window = w && { from: tr.mapping.map(w.from, -1), to: tr.mapping.map(w.to, 1) };
  // Finding what each step changed costs steps × steps (see changedRanges): a transaction of
  // very many steps (a large replace-all, a paste of many runs) marks the document again instead.
  if (tr.steps.length > MAX_STEPS) return decorateAll(doc, window);
  const blocks: Blocks = new Map();
  for (const [from, to] of changedRanges(tr)) collectBlocks(doc, from - 1, to + 1, blocks);
  return { on: true, set: remark(doc, value.set.map(tr.mapping, doc), blocks, window), rebuilt: blocks.size, window };
}

/** Spaces marked in `window` instead: only the paragraphs entering or leaving it change. */
function moveWindow(doc: PMNode, value: FormattingMarksState, window: MarksWindow | null): FormattingMarksState {
  const old = value.window;
  const blocks: Blocks = new Map();
  const changes = (pos: number, node: PMNode) => inWindow(pos, node, old) !== inWindow(pos, node, window);
  if (!old || !window) collectBlocks(doc, 0, doc.content.size, blocks, changes);
  else for (const w of [old, window]) collectBlocks(doc, w.from, w.to, blocks, changes);
  return { on: true, set: remark(doc, value.set, blocks, window), rebuilt: blocks.size, window };
}

const HIDDEN: FormattingMarksState = { on: false, set: DecorationSet.empty, rebuilt: 0, window: null };

const sameWindow = (a: MarksWindow | null, b: MarksWindow | null) => a === b || (!!a && !!b && a.from === b.from && a.to === b.to);

/**
 * The formatting marks of one editor view (the body, or a header/footer being edited).
 * `initial` says whether they show when the view starts; showFormattingMarks changes it.
 */
export function formattingMarks(initial: () => boolean): Plugin<FormattingMarksState> {
  return new Plugin<FormattingMarksState>({
    key: formattingMarksKey,
    state: {
      init: (_config, state) =>
        initial() ? decorateAll(state.doc, { from: 0, to: Math.min(state.doc.content.size, INITIAL_WINDOW) }) : HIDDEN,
      apply: (tr, value) => {
        const meta = tr.getMeta(formattingMarksKey) as FormattingMarksMeta | undefined;
        if (meta?.on !== undefined && meta.on !== value.on) return meta.on ? decorateAll(tr.doc, meta.window ?? null) : HIDDEN;
        if (!value.on) return value;
        let next = tr.docChanged ? decorateChanges(tr, value) : value.rebuilt ? { ...value, rebuilt: 0 } : value;
        if (meta?.window !== undefined && !sameWindow(meta.window, next.window)) next = moveWindow(tr.doc, next, meta.window);
        return next;
      },
    },
    view: (view) => new WindowTracker(view),
    props: {
      decorations: (state) => formattingMarksKey.getState(state)?.set,
      // The glyphs are drawn only under this class (a copy of the text elsewhere shows none).
      attributes: (state): Record<string, string> => (formattingMarksKey.getState(state)?.on ? { class: SHOW_MARKS_CLASS } : {}),
    },
  });
}

/**
 * Show or hide the marks in a view: a transaction that changes nothing in the document and is
 * kept out of the undo history. Nothing is dispatched when the view already shows what is asked.
 */
export function showFormattingMarks(view: EditorView, on: boolean): void {
  const current = formattingMarksKey.getState(view.state);
  if (!current || current.on === on) return;
  const meta: FormattingMarksMeta = { on, window: on ? visibleWindow(view, WINDOW_MARGIN) : null };
  view.dispatch(view.state.tr.setMeta(formattingMarksKey, meta).setMeta('addToHistory', false));
}

/**
 * Word's Ctrl+Shift+8, which is Ctrl+* on a US keyboard, the keypad's * too; on layouts whose
 * Shift+8 is another character, found by the key's code. On a Mac ⌘ instead of Ctrl, and also
 * ⌘8 without Shift (Word for Mac's shortcut). Not with Alt (AltGr on Windows is Ctrl+Alt).
 * Checked by the editor itself, since ProseMirror handles no keys when the document is read-only.
 */
export function isShowMarksKey(e: KeyboardEvent): boolean {
  const mac = typeof navigator !== 'undefined' && /Mac|iP(hone|[oa]d)/.test(navigator.platform);
  if (e.altKey || !(mac ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey)) return false;
  if (e.key === '*') return true;
  const eight = e.code === 'Digit8' || e.keyCode === 56;
  return eight && (e.shiftKey || mac);
}

/**
 * Check again soon where spaces are marked, after something moved the text without a scroll or
 * resize event (zooming the page view).
 */
export function recheckFormattingMarks(view: EditorView): void {
  trackers.get(view)?.schedule();
}

/** Spaces are marked this many screen heights above and below the screen … */
const WINDOW_MARGIN = 2;
/** … and marked again when less than this is left on either side. */
const WINDOW_SLACK = 0.5;
/** How soon after scrolling (or an edit) the window is checked, at most once per this many ms. */
const CHECK_DELAY = 100;

/**
 * The top-level blocks on screen, and `margin` screen heights above and below, as a document
 * range; null when the view is not laid out (hidden, or no layout at all), which means the
 * whole document. A binary search over the blocks' boxes: a few measurements, however long the
 * document is. Being top-level blocks, a long table on screen is one block: all its cells get
 * their spaces marked.
 */
export function visibleWindow(view: EditorView, margin: number): MarksWindow | null {
  const rect = view.dom.getBoundingClientRect();
  const height = view.dom.ownerDocument.defaultView?.innerHeight ?? 0;
  if (!rect.height || !height) return null;
  const top = -margin * height;
  const bottom = height + margin * height;
  const doc = view.state.doc;
  if (rect.bottom < top || rect.top > bottom || !doc.childCount) return { from: 0, to: 0 };
  const starts: number[] = [];
  doc.forEach((_node, offset) => starts.push(offset));
  const box = (i: number) => (view.nodeDOM(starts[i]) as Element | null)?.getBoundingClientRect?.();
  // The first block reaching below `top`, then the last one starting above `bottom`.
  let lo = 0;
  let hi = starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if ((box(mid)?.bottom ?? top) < top) lo = mid + 1;
    else hi = mid;
  }
  const first = lo;
  hi = starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if ((box(mid)?.top ?? bottom) > bottom) hi = mid - 1;
    else lo = mid;
  }
  return { from: starts[first], to: starts[lo] + doc.child(lo).nodeSize };
}

/**
 * Moves the window where spaces are marked as the view scrolls, is resized or changes: shortly
 * after (at most every CHECK_DELAY ms), and only when the screen gets near the window's edge.
 */
/** Each view's tracker (see recheckFormattingMarks). */
const trackers = new WeakMap<EditorView, WindowTracker>();

class WindowTracker {
  private timer: ReturnType<typeof setTimeout> | undefined;
  private readonly win: Window | null;

  constructor(private view: EditorView) {
    trackers.set(view, this);
    this.win = view.dom.ownerDocument.defaultView;
    // Any scrolling container (capture: scroll events do not bubble).
    view.dom.ownerDocument.addEventListener('scroll', this.schedule, { capture: true, passive: true });
    this.win?.addEventListener('resize', this.schedule);
    if (formattingMarksKey.getState(view.state)?.on) this.schedule();
  }

  update(view: EditorView, prev: { doc: PMNode }): void {
    this.view = view;
    if (formattingMarksKey.getState(view.state)?.on && view.state.doc !== prev.doc) this.schedule();
  }

  schedule = (): void => {
    if (this.timer === undefined) this.timer = setTimeout(this.check, CHECK_DELAY);
  };

  private check = (): void => {
    this.timer = undefined;
    const view = this.view;
    const state = formattingMarksKey.getState(view.state);
    if (!state?.on || view.isDestroyed) return;
    // Re-marking paragraphs under an input method's composition would disturb it: later.
    if (view.composing) return this.schedule();
    const needed = visibleWindow(view, WINDOW_SLACK);
    if (!needed) return;
    const w = state.window;
    if (w && w.from <= needed.from && w.to >= needed.to) return;
    const meta: FormattingMarksMeta = { window: visibleWindow(view, WINDOW_MARGIN) };
    view.dispatch(view.state.tr.setMeta(formattingMarksKey, meta).setMeta('addToHistory', false));
  };

  destroy(): void {
    trackers.delete(this.view);
    clearTimeout(this.timer);
    this.view.dom.ownerDocument.removeEventListener('scroll', this.schedule, { capture: true });
    this.win?.removeEventListener('resize', this.schedule);
  }
}

/** The remembered choice; hidden when nothing is remembered or the storage can't be read. */
export function readShowMarks(): boolean {
  try {
    return globalThis.localStorage?.getItem(SHOW_MARKS_STORAGE) === '1';
  } catch {
    return false; // storage blocked (private mode, site data refused …)
  }
}

/** Remember the choice for the next document; silently not remembered when storage fails. */
export function writeShowMarks(on: boolean): void {
  try {
    globalThis.localStorage?.setItem(SHOW_MARKS_STORAGE, on ? '1' : '0');
  } catch {
    // storage blocked or full: the marks still toggle for this page
  }
}

/**
 * Adds the marks to a static copy of a header/footer (the ones drawn on the pages, made by
 * DOMSerializer, which draws no decorations): the same elements and classes the editor's
 * decorations make. Fields (their values are filled in per page) and non-editable chips are
 * left alone. The copy's container needs SHOW_MARKS_CLASS for the glyphs to show.
 */
export function markStaticContent(root: DocumentFragment | Element): void {
  const doc = root.ownerDocument ?? document;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const texts: Text[] = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const t = n as Text;
    if (hasSpace(t.data) && !t.parentElement?.closest('[contenteditable="false"], [data-field], .dx-tab')) texts.push(t);
  }
  for (const t of texts) {
    const pieces = doc.createDocumentFragment();
    let last = 0;
    SPACE_RE.lastIndex = 0;
    for (let m = SPACE_RE.exec(t.data); m; m = SPACE_RE.exec(t.data)) {
      if (m.index > last) pieces.append(t.data.slice(last, m.index));
      const span = doc.createElement('span');
      span.className = SPACES[m[0]].cls;
      span.textContent = m[0];
      pieces.append(span);
      last = m.index + 1;
    }
    if (last < t.data.length) pieces.append(t.data.slice(last));
    t.replaceWith(pieces);
  }
  for (const tab of Array.from(root.querySelectorAll('.dx-tab'))) tab.classList.add('dx-mk-tab');
  for (const br of Array.from(root.querySelectorAll('br'))) br.before(markElement(doc, 'dx-mk-br'));
  for (const p of Array.from(root.querySelectorAll('p.dx-p'))) {
    if (!p.classList.contains('dx-sect-end')) p.append(markElement(doc, 'dx-mk-p'));
  }
}

/**
 * Takes the marks out of a copy of the pages (the printout and the PDF): the ¶/↵ elements go,
 * the spans around spaces give back their text, and nothing is left to turn glyphs on. The text
 * and its layout stay as they were.
 */
export function stripFormattingMarks(root: Element): void {
  for (const m of Array.from(root.querySelectorAll('.dx-mk'))) m.remove();
  for (const s of Array.from(root.querySelectorAll(TEXT_MARK_CLASSES.map((c) => `span.${c}`).join(',')))) {
    s.replaceWith(...Array.from(s.childNodes));
  }
  for (const t of Array.from(root.querySelectorAll('.dx-mk-tab'))) t.classList.remove('dx-mk-tab');
  root.classList.remove(SHOW_MARKS_CLASS);
  for (const e of Array.from(root.querySelectorAll(`.${SHOW_MARKS_CLASS}`))) e.classList.remove(SHOW_MARKS_CLASS);
}
