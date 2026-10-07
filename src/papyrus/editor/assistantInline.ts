// What an AI assistant shows in the text itself, without changing the document:
//
//  - a suggestion: the changes it proposes, drawn where they would be made (what would go struck
//    through in red, what would come in green, what would be reformatted highlighted), until the
//    user accepts (the changes are then made, see assistant.ts) or declines them;
//  - a completion ("ghost text"): grey text after the cursor that Tab makes real and Escape, or
//    anything else typed, makes go away.
//
// Both are decorations: the document, its undo history and the saved file know nothing of them,
// and they are not printed. An edit of the document ends them: they were about the text as it was.
// suggestionPlugin and completionPlugin are among the editor's body plugins; who asks for the
// suggestions and when is the page's business (it is told about typing by an event, below).

import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { AddMarkStep, Mapping, RemoveMarkStep, ReplaceAroundStep, ReplaceStep } from 'prosemirror-transform';
import type { Slice } from 'prosemirror-model';
import { schema } from './schema';
import { isTyping } from './steps';
import './assistantInline.css';

// ----- a suggestion -----

const suggestionKey = new PluginKey<DecorationSet>('papyrus-assistant-suggestion');

/** The text a slice would bring in, one line per block. */
function sliceText(slice: Slice): string {
  return slice.content.textBetween(0, slice.content.size, '\n', (leaf) =>
    leaf.type === schema.nodes.tab ? '\t' : leaf.type === schema.nodes.hard_break ? '↵' : '');
}

function inserted(text: string, block: boolean): HTMLElement {
  const el = document.createElement('ins');
  el.className = block ? 'dx-ai-ins dx-ai-ins-block' : 'dx-ai-ins';
  el.textContent = text;
  // Not part of the text: the cursor and the selection pass it by.
  el.contentEditable = 'false';
  return el;
}

/**
 * How `tr` (made on `state.doc`, not applied) would change the document, as decorations on the
 * document as it is: removed text struck through, new text as a widget where it would go, a new
 * paragraph as a line of its own, reformatted text and paragraphs highlighted.
 */
export function previewDecorations(state: EditorState, tr: Transaction): Decoration[] {
  const out: Decoration[] = [];
  const size = state.doc.content.size;
  /** Positions that already start a new line: the text put there next belongs to that new paragraph. */
  const newLines = new Set<number>();
  tr.steps.forEach((step, i) => {
    // Where this step works, in the document as it is now: back through the steps before it, last first.
    // (Not mapping.slice(0, i).invert(): inverting a slice takes all the maps, the later steps' too.)
    const back = new Mapping();
    for (let j = i - 1; j >= 0; j--) back.appendMap(tr.mapping.maps[j].invert());
    const at = (pos: number, assoc: number) => Math.max(0, Math.min(size, back.map(pos, assoc)));
    if (step instanceof ReplaceStep) {
      const s = step as unknown as { from: number; to: number; slice: Slice };
      const from = at(s.from, 1);
      const to = Math.max(from, at(s.to, -1));
      if (to > from) out.push(Decoration.inline(from, to, { class: 'dx-ai-del' }));
      // A paragraph split in two (Enter): what is put there next is a new paragraph (nothing to show yet).
      const split = s.slice.openStart > 0 && s.slice.content.childCount > 1;
      if (split) newLines.add(to);
      const text = split ? '' : sliceText(s.slice);
      const wholeBlock = !s.slice.openStart && s.slice.content.firstChild?.isBlock;
      if (text) {
        const block = !!wholeBlock || newLines.has(to);
        // In the order the steps were made: a new paragraph's text after the one before it.
        out.push(Decoration.widget(to, () => inserted(text, block), { side: i + 1, key: `ai-${i}-${text}`, ignoreSelection: true }));
      }
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const s = step as unknown as { from: number; to: number };
      const from = at(s.from, 1);
      const to = at(s.to, -1);
      if (to > from) out.push(Decoration.inline(from, to, { class: 'dx-ai-fmt' }));
    } else if (step instanceof ReplaceAroundStep && (step as unknown as { structure: boolean }).structure) {
      // A paragraph's attributes (alignment, style). A paragraph only split (its mark made new) looks
      // the same: said only when what shows of it changes.
      const s = step as unknown as { from: number };
      const pos = at(s.from, 1);
      const before = state.doc.nodeAt(pos);
      const after = tr.docs[i + 1]?.nodeAt(s.from) ?? (i + 1 === tr.steps.length ? tr.doc.nodeAt(s.from) : null);
      if (before && after && before.type === schema.nodes.paragraph && after.type === before.type
        && (before.attrs.align !== after.attrs.align || before.attrs.styleId !== after.attrs.styleId)) {
        out.push(Decoration.node(pos, pos + before.nodeSize, { class: 'dx-ai-fmt-block' }));
      }
    }
  });
  return out;
}

/** The suggestion being shown (drawn over the body), kept until the document changes or it is cleared. */
export function suggestionPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: suggestionKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, value) {
        const meta = tr.getMeta(suggestionKey) as { set?: Decoration[]; clear?: boolean } | undefined;
        if (meta?.set) return DecorationSet.create(tr.doc, meta.set);
        if (meta?.clear || tr.docChanged) return DecorationSet.empty;
        return value;
      },
    },
    props: {
      decorations: (state) => suggestionKey.getState(state),
    },
  });
}

/** Whether a suggestion is being shown. */
export function hasSuggestion(state: EditorState): boolean {
  return (suggestionKey.getState(state) ?? DecorationSet.empty) !== DecorationSet.empty;
}

/** Shows what `tr` (a transaction on the view's document, not applied) would change; false when it would change nothing shown. */
export function showSuggestion(view: EditorView, tr: Transaction): boolean {
  const decorations = previewDecorations(view.state, tr);
  if (!decorations.length) return false;
  view.dispatch(view.state.tr.setMeta(suggestionKey, { set: decorations }).setMeta('addToHistory', false));
  return true;
}

/**
 * Marks the text a request is about (positions in the view's document), while the keyboard is
 * elsewhere (in the box the request is typed in) and the browser no longer shows the selection.
 */
export function showTarget(view: EditorView, from: number, to: number): void {
  const size = view.state.doc.content.size;
  const a = Math.max(0, Math.min(size, from));
  const b = Math.max(a, Math.min(size, to));
  view.dispatch(view.state.tr.setMeta(suggestionKey, { set: b > a ? [Decoration.inline(a, b, { class: 'dx-ai-target' })] : [] }).setMeta('addToHistory', false));
}

export function clearSuggestion(view: EditorView): void {
  if (hasSuggestion(view.state)) view.dispatch(view.state.tr.setMeta(suggestionKey, { clear: true }).setMeta('addToHistory', false));
}

// ----- a completion -----

interface Ghost {
  pos: number;
  text: string;
}

interface CompletionState {
  ghost: Ghost | null;
  /** Goes up with every keystroke that typed text: the page asks for a completion when it settles. */
  typed: number;
}

const completionKey = new PluginKey<CompletionState>('papyrus-assistant-completion');

/** Sent (bubbling) from the editor's text when the user typed: the page may ask for a completion. */
export const TYPED_EVENT = 'dx-assistant-typed';

/**
 * Grey text after the cursor. Tab types it; Escape, moving the cursor or any edit drops it.
 * Before the editor's own key bindings, so Tab and Escape are the completion's while one shows.
 */
export function completionPlugin(): Plugin<CompletionState> {
  return new Plugin<CompletionState>({
    key: completionKey,
    state: {
      init: () => ({ ghost: null, typed: 0 }),
      apply(tr, value, old, next) {
        const meta = tr.getMeta(completionKey) as { ghost?: Ghost | null } | undefined;
        const typed = isTyping(tr) && !tr.getMeta(completionKey) ? value.typed + 1 : value.typed;
        if (meta && 'ghost' in meta) return { ghost: meta.ghost ?? null, typed };
        // It was about the text and the cursor as they were. (A selection set again to where it was is no move.)
        if (value.ghost && (tr.docChanged || !next.selection.eq(old.selection))) return { ghost: null, typed };
        return typed === value.typed ? value : { ghost: value.ghost, typed };
      },
    },
    props: {
      decorations(state) {
        const ghost = completionKey.getState(state)?.ghost;
        if (!ghost) return null;
        return DecorationSet.create(state.doc, [
          Decoration.widget(ghost.pos, () => {
            const el = document.createElement('span');
            el.className = 'dx-ai-ghost';
            el.textContent = ghost.text;
            el.contentEditable = 'false';
            el.setAttribute('aria-hidden', 'true');
            return el;
          }, { side: 1, key: `ghost-${ghost.text}`, ignoreSelection: true }),
        ]);
      },
      handleKeyDown(view, event) {
        const ghost = completionKey.getState(view.state)?.ghost;
        if (!ghost || event.isComposing || event.keyCode === 229) return false;
        if (event.key === 'Tab' && !event.shiftKey && !event.ctrlKey && !event.altKey && !event.metaKey) {
          event.preventDefault();
          acceptCompletion(view);
          return true;
        }
        if (event.key === 'Escape') {
          clearCompletion(view);
          return true;
        }
        return false;
      },
    },
    view: () => ({
      update(view, prev) {
        const now = completionKey.getState(view.state)?.typed ?? 0;
        const before = completionKey.getState(prev)?.typed ?? 0;
        if (now !== before) view.dom.dispatchEvent(new CustomEvent(TYPED_EVENT, { bubbles: true }));
      },
    }),
  });
}

/** The completion being shown, if any. */
export function completionOf(state: EditorState): Ghost | null {
  return completionKey.getState(state)?.ghost ?? null;
}

/**
 * Where a completion may be offered: the cursor (no selection) at the end of a paragraph's text.
 * Returns that position, or null.
 */
export function completionPoint(state: EditorState): number | null {
  const { selection } = state;
  if (!selection.empty) return null;
  const $pos = selection.$from;
  if ($pos.parent.type !== schema.nodes.paragraph || $pos.parentOffset !== $pos.parent.content.size) return null;
  return $pos.pos;
}

/** Shows `text` after the cursor, when the cursor is still where a completion may be offered (`at`: where it was asked for). */
export function showCompletion(view: EditorView, text: string, at: number): boolean {
  if (!text || completionPoint(view.state) !== at) return false;
  view.dispatch(view.state.tr.setMeta(completionKey, { ghost: { pos: at, text } }).setMeta('addToHistory', false));
  return true;
}

export function clearCompletion(view: EditorView): void {
  if (completionOf(view.state)) view.dispatch(view.state.tr.setMeta(completionKey, { ghost: null }).setMeta('addToHistory', false));
}

/** Types the completion being shown (in the formatting typed text gets there); false when there is none. */
export function acceptCompletion(view: EditorView): boolean {
  const ghost = completionOf(view.state);
  if (!ghost) return false;
  // Through the usual dispatch: one edit like typed text (a tracked insertion while 追蹤修訂 is on).
  view.dispatch(view.state.tr.insertText(ghost.text, ghost.pos).setMeta(completionKey, { ghost: null }).scrollIntoView());
  return true;
}
