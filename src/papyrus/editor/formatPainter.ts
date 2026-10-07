import type { Command, EditorState } from 'prosemirror-state';
import { Plugin, TextSelection } from 'prosemirror-state';
import type { Mark, Node as PMNode, ResolvedPos } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { columnResizingPluginKey } from 'prosemirror-tables';
import { schema } from './schema';
import { APPEARANCE_MARKS, selectedParagraphs } from './commands';
import { MODEL_KEYS } from './trackChanges';

// Word's 複製格式 (Format Painter, Home › Clipboard): copy the formatting of the selection and paste
// it onto other text. The character formatting is the direct formatting the editor models (the
// appearance marks, character style included); the paragraph formatting is the modelled paragraph
// settings. Numbering, section breaks, page break before and other raw XML are not copied: Word's
// painter carries a paragraph's list only through its style, and a section or a page break is
// not formatting. Everything goes through ordinary mark and paragraph-attribute steps, so 追蹤修訂
// records it (w:rPrChange / w:pPrChange) and one paste is one undo step.
// Known gap from Word: run properties the editor does not model (w:spacing, caps, shading ...)
// live only in the raw w:rPr and are not copied; the target keeps its own.

/**
 * The paragraph attributes 複製格式 copies: the modelled ones that 追蹤修訂 records (style,
 * alignment, indents, spacing, line spacing), less list membership and page break before.
 */
export const PARAGRAPH_FORMAT_KEYS: readonly string[] = MODEL_KEYS.filter(
  (k) => k !== 'numId' && k !== 'ilvl' && k !== 'pageBreakBefore',
);

/** Formatting copied by 複製格式 (the toolbar button or Ctrl+Shift+C). */
export interface CopiedFormat {
  /** The character formatting: the appearance marks of the copied text (none = plain text). */
  marks: readonly Mark[];
  /** The paragraph formatting, when the copy included a paragraph mark; otherwise null. */
  paragraph: Record<string, unknown> | null;
}

const appearanceTypes = () => APPEARANCE_MARKS.map((n) => schema.marks[n]);

/**
 * The formatting to copy from the selection, as Word decides it: the character formatting of
 * the text at the cursor or the start of the selection; the paragraph formatting as well when
 * nothing is selected, or when the selection holds a whole paragraph or reaches into the next
 * (it includes a paragraph mark).
 */
export function copyFormat(state: EditorState): CopiedFormat {
  const { empty, from, to, $from } = state.selection;
  const types = appearanceTypes();
  let marks: readonly Mark[] = empty ? state.storedMarks ?? $from.marks() : $from.marks();
  if (!empty) {
    let first: readonly Mark[] | null = null;
    state.doc.nodesBetween(from, to, (node) => {
      if (first) return false;
      if (node.isText) first = node.marks;
      return !node.isInline;
    });
    if (first) marks = first;
  }
  const paras = selectedParagraphs(state);
  let paragraph: Record<string, unknown> | null = null;
  const p = paras[0];
  if (p && (empty || paras.length > 1 || (from <= p.pos + 1 && to >= p.pos + p.node.nodeSize - 1))) {
    paragraph = Object.fromEntries(PARAGRAPH_FORMAT_KEYS.map((k) => [k, p.node.attrs[k] ?? null]));
  }
  return { marks: marks.filter((m) => types.includes(m.type)), paragraph };
}

/** Word boundaries that know Chinese words, where the browser has them (Intl.Segmenter). */
const SEGMENTER: { segment(text: string): Iterable<{ segment: string; index: number; isWordLike?: boolean }> } | null =
  typeof Intl !== 'undefined' && (Intl as any).Segmenter ? new (Intl as any).Segmenter('zh-Hant', { granularity: 'word' }) : null;

/** Where `at` falls in a word of `text` (Chinese words included, where the browser can tell them). */
function wordAround(text: string, at: number): { from: number; to: number } | null {
  if (SEGMENTER) {
    let before: { from: number; to: number } | null = null;
    for (const s of SEGMENTER.segment(text)) {
      const end = s.index + s.segment.length;
      if (!s.isWordLike) continue;
      if (s.index <= at && at < end) return { from: s.index, to: end };
      if (end === at) before = { from: s.index, to: end };
    }
    return before;
  }
  const isWord = (c: string | undefined) => !!c && /[\p{L}\p{N}_]/u.test(c);
  let from = at;
  let to = at;
  while (isWord(text[from - 1])) from--;
  while (isWord(text[to])) to++;
  return from < to ? { from, to } : null;
}

/** The word at a cursor position (Word paints the word under a click), as document positions. */
function wordAt($pos: ResolvedPos): { from: number; to: number } | null {
  const parent = $pos.parent;
  if (!parent.isTextblock) return null;
  // One character per position: pictures, tabs and fields (inline leaves) end a word.
  let text = '';
  parent.forEach((n) => (text += n.isText ? n.text! : '\uFFFC'.repeat(n.nodeSize)));
  const word = wordAround(text, $pos.parentOffset);
  return word && { from: $pos.start() + word.from, to: $pos.start() + word.to };
}

/**
 * Paste copied formatting onto the selection, or onto the word at the cursor when nothing is
 * selected: the text's own appearance marks are replaced by the copied ones (other run
 * properties, such as language, stay), and copied paragraph formatting goes to every paragraph
 * the selection touches. Outside a word, text typed next gets the formatting. One undo step; no
 * transaction when nothing changes.
 */
export function pasteFormat(format: CopiedFormat): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const types = appearanceTypes();
    const { empty, $from } = state.selection;
    // Every range of the selection: a cell selection has one per cell.
    const word = empty ? wordAt($from) : null;
    const ranges = empty ? (word ? [word] : []) : state.selection.ranges.map((r) => ({ from: r.$from.pos, to: r.$to.pos }));
    const tr = state.tr;
    for (const range of ranges) {
      for (const type of types) tr.removeMark(range.from, range.to, type);
      for (const mark of format.marks) tr.addMark(range.from, range.to, mark);
    }
    if (format.paragraph) {
      for (const { node, pos } of selectedParagraphs(state)) {
        if (PARAGRAPH_FORMAT_KEYS.every((k) => (node.attrs[k] ?? null) === format.paragraph![k])) continue;
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...format.paragraph }, node.marks);
      }
    }
    const changed = !tr.doc.eq(state.doc);
    if (!ranges.length) {
      // Stored marks last: a step clears them.
      const here = (state.storedMarks ?? $from.marks()).filter((m) => !types.includes(m.type));
      tr.setStoredMarks(format.marks.reduce<readonly Mark[]>((set, m) => m.addToSet(set), here));
    } else if (!changed) {
      return true;
    }
    // Its own undo step, even right after typing or another paste (Word undoes each one alone).
    dispatch(changed ? closeHistory(tr).scrollIntoView() : tr);
    return true;
  };
}

/**
 * The browser's selection in `view` as document positions (null when it is elsewhere). Read at
 * mouseup, when it is what the user just selected: ProseMirror takes it over only on the
 * `selectionchange` event, which may come later (Chrome delays it while the window is not
 * drawing), and by the time the painter runs a new click may already have moved it.
 */
function domSelection(view: EditorView): { anchor: number; head: number } | null {
  if (!view.hasFocus()) return null;
  const sel = view.dom.ownerDocument.getSelection();
  if (!sel?.anchorNode || !sel.focusNode || !view.dom.contains(sel.anchorNode) || !view.dom.contains(sel.focusNode)) return null;
  try {
    return { anchor: view.posAtDOM(sel.anchorNode, sel.anchorOffset), head: view.posAtDOM(sel.focusNode, sel.focusOffset) };
  } catch {
    return null; // not a position in the document (a widget, a node view's inside)
  }
}

/** Whether 複製格式 is painting: once (one click on the button) or until stopped (double-click). */
export type FormatPainterMode = 'once' | 'sticky';

/**
 * The 複製格式 state of one editor, shared by the body and the header/footer being edited (so a
 * format copied in one can be pasted in the other). While painting, the next mouse selection
 * in any of the editor's views receives the button's format; Escape stops. Ctrl+Shift+V pastes
 * only what Ctrl+Shift+C copied, and only until the next Ctrl+C / Ctrl+X: otherwise the key is
 * left to the browser, whose Ctrl+Shift+V is 「貼上為純文字」.
 */
export class FormatPainter {
  /** What Ctrl+Shift+C copied, for Ctrl+Shift+V; null after a native copy or cut. */
  copied: CopiedFormat | null = null;
  /** What the button copied, for painting with the mouse. */
  private brush: CopiedFormat | null = null;
  /** Painting with the mouse, and how long; null when not painting. */
  mode: FormatPainterMode | null = null;
  /** The view the mouse button went down in while painting: its mouseup paints. */
  private pressed: EditorView | null = null;

  /** @param changed called when painting starts or stops (the toolbar and the cursor follow). */
  constructor(private changed: () => void) {}

  /** Ctrl+Shift+C: copy the formatting of the selection, without starting to paint. */
  copy(view: EditorView): boolean {
    this.copied = copyFormat(view.state);
    return true;
  }

  /** Ctrl+Shift+V: paste the copied formatting onto the selection; false when nothing was copied. */
  paste(view: EditorView): boolean {
    if (!this.copied) return false;
    return pasteFormat(this.copied)(view.state, view.dispatch, view);
  }

  /** The button: copy the selection's formatting and paint the next mouse selection(s). */
  start(view: EditorView, sticky: boolean): void {
    this.brush = copyFormat(view.state);
    this.mode = sticky ? 'sticky' : 'once';
    this.pressed = null;
    this.changed();
  }

  /** Stop painting (Escape, the button again, or another command). */
  stop(): void {
    if (!this.mode) return;
    this.mode = null;
    this.pressed = null;
    this.changed();
  }

  /** A new document: stop painting and forget what was copied. */
  reset(): void {
    this.copied = null;
    this.brush = null;
    this.stop();
  }

  /**
   * The mouse was released after a click or drag in `view`: paint what was selected then (`at`,
   * read from the browser while the document was `doc`), else ProseMirror's selection. A picture
   * or cells ProseMirror selected on mouseup (a node or cell selection) are left as they are.
   */
  private paint(view: EditorView, at: { anchor: number; head: number } | null, doc: PMNode): void {
    if (!this.mode || !this.brush || view.isDestroyed) return;
    const { state } = view;
    const sel = state.selection;
    if (at && state.doc === doc && sel instanceof TextSelection && (at.anchor !== sel.anchor || at.head !== sel.head)) {
      const next = TextSelection.between(doc.resolve(at.anchor), doc.resolve(at.head));
      view.dispatch(state.tr.setSelection(next).setMeta('pointer', true));
    }
    pasteFormat(this.brush)(view.state, view.dispatch, view);
    if (this.mode === 'once') this.stop();
  }

  /** The plugin each of the editor's views gets (before the keymaps: Escape stops painting first). */
  plugin(): Plugin {
    return new Plugin({
      props: {
        handleKeyDown: (_view, event) => {
          if (!this.mode || event.key !== 'Escape') return false;
          this.stop();
          return true;
        },
        handleDOMEvents: {
          mousedown: (view, event) => {
            if (!this.mode || event.button !== 0) return false;
            // Dragging a table column border is not a selection to paint.
            const onHandle = (event.target as Element | null)?.closest?.('.column-resize-handle');
            const resizing = (columnResizingPluginKey.getState(view.state)?.activeHandle ?? -1) > -1;
            if (!onHandle && !resizing) this.pressed = view;
            return false;
          },
        },
      },
      view: (view) => {
        const doc = view.dom.ownerDocument;
        // On the document: a drag may end outside the text.
        const up = () => {
          if (this.pressed !== view) return;
          this.pressed = null;
          const at = domSelection(view);
          const now = view.state.doc;
          // After ProseMirror has handled the mouseup (a click on a picture selects it).
          setTimeout(() => this.paint(view, at, now), 0);
        };
        // Ctrl+C / Ctrl+X: the user now means the clipboard, so Ctrl+Shift+V is the browser's
        // 「貼上為純文字」 again.
        const forget = () => (this.copied = null);
        doc.addEventListener('mouseup', up);
        doc.addEventListener('copy', forget);
        doc.addEventListener('cut', forget);
        return {
          destroy: () => {
            doc.removeEventListener('mouseup', up);
            doc.removeEventListener('copy', forget);
            doc.removeEventListener('cut', forget);
          },
        };
      },
    });
  }
}
