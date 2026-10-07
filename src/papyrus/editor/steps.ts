import type { Transaction } from 'prosemirror-state';
import { isHistoryTransaction } from 'prosemirror-history';
import {
  AddMarkStep, AddNodeMarkStep, AttrStep, DocAttrStep, RemoveMarkStep, RemoveNodeMarkStep, ReplaceAroundStep, ReplaceStep,
} from 'prosemirror-transform';
import type { Fragment, Mark, Node as PMNode } from 'prosemirror-model';

/**
 * What a transaction's steps add, remove or change, to decide whether a plugin can keep (map)
 * what it computed for the previous document instead of computing it again.
 */
export interface NodeTest {
  node: (node: PMNode) => boolean;
  /** Marks that count too (added or removed by mark steps, or on inserted content). */
  mark?: (mark: Mark) => boolean;
}

function fragmentHas(content: Fragment, test: NodeTest): boolean {
  let found = false;
  content.descendants((node) => {
    if (found) return false;
    if (test.node(node) || (test.mark && node.marks.some(test.mark))) found = true;
    return !found;
  });
  return found;
}

/**
 * Whether a matching node in `doc` is removed or changed by replacing `from`..`to`: nodes the
 * range lies strictly inside of (the paragraph being typed in) stay as they are.
 */
function rangeHits(doc: PMNode, from: number, to: number, test: NodeTest): boolean {
  if (to > doc.content.size) return true;
  let found = false;
  doc.nodesBetween(from, to, (node, pos) => {
    if (found) return false;
    const inside = pos + 1 <= from && to <= pos + node.nodeSize - 1;
    // Removed text keeps its marks out of the document (e.g. a field's wrapper).
    if (!inside && (test.node(node) || (test.mark && from < to && node.marks.some(test.mark)))) found = true;
    return !found;
  });
  return found;
}

/**
 * Whether the transaction adds, removes or changes a node (or mark) that passes `test`.
 * Unknown step types count as a change.
 */
export function touches(tr: Transaction, test: NodeTest): boolean {
  for (let i = 0; i < tr.steps.length; i++) {
    const step = tr.steps[i];
    const before = tr.docs[i];
    const after = i + 1 < tr.docs.length ? tr.docs[i + 1] : tr.doc;
    if (step instanceof ReplaceStep) {
      const s = step as unknown as { from: number; to: number; slice: { content: Fragment } };
      if (fragmentHas(s.slice.content, test) || rangeHits(before, s.from, s.to, test)) return true;
    } else if (step instanceof ReplaceAroundStep) {
      const s = step as unknown as { from: number; to: number; gapFrom: number; gapTo: number; slice: { content: Fragment } };
      if (fragmentHas(s.slice.content, test) || rangeHits(before, s.from, s.gapFrom, test) || rangeHits(before, s.gapTo, s.to, test)) return true;
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep || step instanceof AddNodeMarkStep || step instanceof RemoveNodeMarkStep) {
      const mark = (step as unknown as { mark: Mark }).mark;
      if (test.mark?.(mark)) return true;
    } else if (step instanceof AttrStep) {
      const pos = (step as unknown as { pos: number }).pos;
      const a = before.nodeAt(pos);
      const b = after.nodeAt(pos);
      if (!a || !b || test.node(a) || test.node(b)) return true;
    } else if (step instanceof DocAttrStep) {
      if (test.node(before) || test.node(after)) return true;
    } else {
      return true;
    }
  }
  return false;
}

/**
 * Whether the transaction only types or deletes text inside single paragraphs (no paste,
 * no undo, no new or removed blocks): what a keystroke does.
 */
export function isTyping(tr: Transaction): boolean {
  if (!tr.docChanged || isHistoryTransaction(tr) || tr.getMeta('paste') || tr.getMeta('uiEvent') === 'paste' || tr.getMeta('uiEvent') === 'drop' || tr.getMeta('uiEvent') === 'cut') return false;
  let replaced = false;
  for (let i = 0; i < tr.steps.length; i++) {
    const step = tr.steps[i];
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) continue; // e.g. a tracked insertion's mark
    if (!(step instanceof ReplaceStep)) return false;
    const s = step as unknown as { from: number; to: number; slice: { content: Fragment; openStart: number; openEnd: number } };
    if (s.slice.openStart || s.slice.openEnd) return false;
    // Text only: a picture, a field or a tab is an edit of its own (measured right away).
    let text = true;
    s.slice.content.forEach((n) => {
      if (!n.isText) text = false;
    });
    if (!text) return false;
    const doc = tr.docs[i];
    if (s.to > doc.content.size) return false;
    const $from = doc.resolve(s.from);
    if (!$from.parent.isTextblock || !$from.sameParent(doc.resolve(s.to))) return false;
    replaced = true;
  }
  return replaced;
}
