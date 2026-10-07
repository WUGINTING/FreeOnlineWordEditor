// Editing a text box's text (double-click it, or select it and press Enter), like editing a
// header: an editor of its own over the text box, with the same schema, styles, lists, undo and
// shortcuts as the body. Every change is written into the shape's model in the editor that holds
// the shape (the body, or the header being edited): one undo step there per editing session, so
// Ctrl+Z in the body after closing it undoes the whole edit. Esc, or a click outside, closes it.
// On save only that text box's w:txbxContent is written again (docx/shapes.ts rewriteTextBoxes).

import { EditorState, NodeSelection, Plugin, TextSelection, type Transaction } from 'prosemirror-state';
import { ReplaceStep, ReplaceAroundStep } from 'prosemirror-transform';
import { EditorView } from 'prosemirror-view';
import { keymap } from 'prosemirror-keymap';
import { closeHistory } from 'prosemirror-history';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from './schema';
import { NO_TRACK } from './trackChanges';
import { frameDoc, WORD_ONLY } from './shapeView';
import { newShapeKey, textFrames, withText, type ShapeModel, type TextFrame } from '../docx/shapes';
import { drawingIds, reidentified } from '../docx/shapeOps';
import { tl } from '../i18n';

/** The shape with model key `key` in a document. */
export function findShape(doc: PMNode, key: string): { pos: number; node: PMNode; shape: ShapeModel } | null {
  let out: { pos: number; node: PMNode; shape: ShapeModel } | null = null;
  doc.descendants((node, pos) => {
    if (out) return false;
    const shape = node.attrs?.shape as ShapeModel | null | undefined;
    if (node.type === schema.nodes.raw_inline && shape?.key === key) out = { pos, node, shape };
    return !out;
  });
  return out;
}

/**
 * The text boxes of a document's shapes that can be edited here (with `all`, also those of
 * shapes that can't be drawn, whose text still counts), in document order.
 */
export function editableTextBoxes(doc: PMNode, all = false): { key: string; frame: TextFrame; shape: ShapeModel }[] {
  const out: { key: string; frame: TextFrame; shape: ShapeModel }[] = [];
  doc.descendants((node) => {
    const shape = node.attrs?.shape as ShapeModel | null | undefined;
    if (node.type === schema.nodes.raw_inline && shape && (all || shape.drawable)) {
      for (const frame of textFrames(shape)) out.push({ key: shape.key, frame, shape });
    }
    return true;
  });
  return out;
}

/**
 * A transaction on `state` that gives text box `index` of shape `key` the text `doc` (never
 * tracked as a revision: the text box's own editor tracks its text). Text back as it was gives
 * `original` back, so an untouched file is saved as it was. Null when the shape is gone.
 */
export function setTextBoxTr(state: EditorState, key: string, index: number, doc: PMNode, original?: ShapeModel): Transaction | null {
  const found = findShape(state.doc, key);
  if (!found) return null;
  const was = original ?? found.shape;
  const frame = textFrames(was).find((f) => f.i === index);
  if (!frame) return null;
  const shape = !textFrames(was).some((f) => f.edited) && doc.eq(frameDoc(frame)) ? was : withText(found.shape, index, doc.toJSON() as Record<string, unknown>);
  if (shape === found.shape) return null;
  return state.tr.setNodeMarkup(found.pos, undefined, { ...found.node.attrs, shape }, found.node.marks).setMeta(NO_TRACK, true);
}

export interface ShapeEditHost {
  /** The editor holding the shape. */
  view: EditorView;
  /** Where the text box's editor goes: a layer over the pages (canvas coordinates). */
  layer: HTMLElement;
  /** The editor's plugins (undo, lists, track changes ...), as for a header. */
  plugins(): Plugin[];
  nodeViews: NonNullable<ConstructorParameters<typeof EditorView>[1]>['nodeViews'];
  attributes(): Record<string, string>;
  /** Dispatches in the text box's editor (tracked changes, errors caught). */
  dispatch(view: EditorView, tr: Transaction, after: (tr: Transaction) => void): void;
  /** After the text changed (written into the shape). */
  onChange(): void;
  /** The session ended (Esc, a click outside, the shape went away). */
  onClose(session: ShapeTextSession): void;
}

let sessions = 0;

/** A text box being edited. */
export class ShapeTextSession {
  readonly view!: EditorView;
  readonly box: HTMLElement;
  private readonly style: HTMLStyleElement;
  /** The shape as it was when editing started (text put back as it was gives it back). */
  private readonly original: ShapeModel;
  /** The first change of this session written into the holding editor (later ones join its undo step). */
  private first: Transaction | null = null;
  private closed = false;
  private frameId = 0;

  constructor(readonly host: ShapeEditHost, readonly key: string, readonly index: number) {
    const found = findShape(host.view.state.doc, key);
    const frame = found && textFrames(found.shape).find((f) => f.i === index);
    if (!found || !frame) throw new Error('no such text box');
    this.original = found.shape;
    this.box = document.createElement('div');
    this.box.className = 'dx-shape-edit-box';
    this.box.dataset.session = String(++sessions);
    // The drawn text stays hidden while this editor shows it.
    this.style = document.createElement('style');
    this.style.textContent = `.dx-shape[data-shape-key="${key}"] .dx-shape-text[data-text="${index}"]{visibility:hidden}`;
    const mount = document.createElement('div');
    this.box.append(this.style, mount);
    host.layer.append(this.box);
    try {
      this.view = this.makeView(mount, frame);
    } catch (err) {
      this.box.remove(); // nothing of a failed editor stays on the page
      throw err;
    }
    this.place();
    document.addEventListener('mousedown', this.onMouseDown, true);
  }

  private makeView(mount: HTMLElement, frame: TextFrame): EditorView {
    const host = this.host;
    const close: Plugin = keymap({
      Escape: () => {
        this.close(true);
        return true;
      },
    });
    const view: EditorView = new EditorView(mount, {
      state: EditorState.create({ schema, doc: frameDoc(frame), plugins: [close, ...host.plugins()] }),
      attributes: () => ({ ...host.attributes(), class: 'dx-doc dx-shape-doc', role: 'textbox', 'aria-multiline': 'true', 'aria-label': tl('文字方塊') }),
      nodeViews: host.nodeViews,
      dispatchTransaction: (tr) => host.dispatch(view, tr, (done) => {
        if (done.docChanged) this.commit();
      }),
    });
    return view;
  }

  /** A click outside the text box ends the editing (toolbar and menus aside). */
  private onMouseDown = (e: MouseEvent) => {
    const t = e.target as Element | null;
    if (!t || this.box.contains(t) || !this.host.layer.closest('.dx-root')?.contains(t)) return;
    if (t.closest?.(`.dx-shape[data-shape-key="${this.key}"]`)) return;
    this.close(false);
  };

  /** The text box's editor over the drawn text (again after the shape was drawn again). */
  place(): void {
    const target = this.host.view.dom.querySelector<HTMLElement>(`.dx-shape[data-shape-key="${this.key}"] .dx-shape-text[data-text="${this.index}"]`);
    if (!target) {
      // The shape went away (undone, deleted elsewhere): its text can't be edited any more.
      if (!this.host.view.isDestroyed && !findShape(this.host.view.state.doc, this.key)) this.close(false);
      return;
    }
    const layer = this.host.layer.getBoundingClientRect();
    const zoom = this.host.layer.offsetWidth ? layer.width / this.host.layer.offsetWidth || 1 : 1;
    // Placed from the text box's centre, which turning the shape does not move; turned the same way.
    const holder = target.closest<HTMLElement>('.dx-shape-textbox') ?? target;
    const hr = holder.getBoundingClientRect();
    const left = ((hr.left + hr.right) / 2 - layer.left) / zoom - holder.offsetWidth / 2 + (holder === target ? 0 : target.offsetLeft);
    const top = ((hr.top + hr.bottom) / 2 - layer.top) / zoom - holder.offsetHeight / 2 + (holder === target ? 0 : target.offsetTop);
    const cs = getComputedStyle(target);
    const turn = holder.style.transform;
    this.box.style.cssText =
      `left:${left}px;top:${top}px;width:${target.offsetWidth}px;min-height:${target.offsetHeight}px;justify-content:${cs.justifyContent || 'flex-start'}` +
      (cs.writingMode && cs.writingMode !== 'horizontal-tb' ? `;writing-mode:${cs.writingMode}` : '') +
      (turn ? `;transform:${turn};transform-origin:${holder.offsetWidth / 2 - target.offsetLeft}px ${holder.offsetHeight / 2 - target.offsetTop}px` : '');
  }

  /** Write the text into the shape, in the holding editor. */
  private commit(): void {
    const host = this.host.view;
    const tr = setTextBoxTr(host.state, this.key, this.index, this.view.state.doc, this.original);
    if (!tr) {
      if (!findShape(host.state.doc, this.key)) this.close(false);
      return;
    }
    // One undo step in the holding editor for the whole session.
    if (this.first) tr.setMeta('appendedTransaction', this.first);
    else this.first = closeHistory(tr);
    host.dispatch(tr);
    this.host.onChange();
    cancelAnimationFrame(this.frameId);
    if (typeof requestAnimationFrame === 'function') this.frameId = requestAnimationFrame(() => this.place());
  }

  /** Focus the text, with the cursor where `at` (client coordinates, the double-click) is when given. */
  focus(at?: { x: number; y: number }): void {
    const hit = at ? this.view.posAtCoords({ left: at.x, top: at.y }) : null;
    if (hit) this.view.dispatch(this.view.state.tr.setSelection(TextSelection.near(this.view.state.doc.resolve(hit.pos))).setMeta('addToHistory', false));
    this.view.focus();
  }

  close(focusHost: boolean): void {
    if (this.closed) return;
    this.closed = true;
    document.removeEventListener('mousedown', this.onMouseDown, true);
    cancelAnimationFrame(this.frameId);
    this.view.destroy();
    this.box.remove();
    this.host.onClose(this);
    if (focusHost && !this.host.view.isDestroyed) {
      // Back on the shape, selected.
      const found = findShape(this.host.view.state.doc, this.key);
      if (found) this.host.view.dispatch(this.host.view.state.tr.setSelection(NodeSelection.create(this.host.view.state.doc, found.pos)).setMeta('addToHistory', false));
      this.host.view.focus();
    }
  }
}

/** The text box a double-click (or Enter on a selected shape) edits: the one hit, else the shape's first. */
export function textBoxAt(target: Element | null): { key: string; index: number } | { wordOnly: true } | null {
  const el = target?.closest?.<HTMLElement>('.dx-shape');
  if (!el) return null;
  if (el.classList.contains('dx-shape-fixed')) return { wordOnly: true };
  const key = el.dataset.shapeKey;
  if (!key) return null;
  const hit = target!.closest<HTMLElement>('.dx-shape-text');
  const index = hit ? Number(hit.dataset.text) : Number(el.querySelector<HTMLElement>('.dx-shape-text')?.dataset.text ?? NaN);
  return Number.isFinite(index) ? { key, index } : null;
}

export { WORD_ONLY };

/**
 * Shapes copied in the document (pasted, dragged with Ctrl) get a key and drawing ids of their
 * own, in the same step: edits of the copy must not go to the original, and Word wants each
 * drawing's id once.
 */
export function shapeKeys(): Plugin {
  const insertsShapes = (tr: Transaction) =>
    tr.steps.some((s) => {
      if (!(s instanceof ReplaceStep || s instanceof ReplaceAroundStep)) return false;
      let found = false;
      (s as unknown as { slice: { content: { descendants(f: (n: PMNode) => boolean): void } } }).slice.content.descendants((n) => {
        if (n.attrs?.shape) found = true;
        return !found;
      });
      return found;
    });
  return new Plugin({
    appendTransaction(trs, _old, state) {
      if (!trs.some((t) => t.docChanged && insertsShapes(t))) return null;
      const keys = new Set<string>();
      const ids = new Set<number>();
      const dup: { pos: number; node: PMNode }[] = [];
      state.doc.descendants((node, pos) => {
        const shape = node.attrs?.shape as ShapeModel | null | undefined;
        if (node.type !== schema.nodes.raw_inline || !shape) return true;
        if (keys.has(shape.key) || (shape.docId != null && ids.has(shape.docId))) dup.push({ pos, node });
        keys.add(shape.key);
        if (shape.docId != null) ids.add(shape.docId);
        return true;
      });
      if (!dup.length) return null;
      const used = drawingIds(Array.from(ids, (id) => `<wp:docPr id="${id}"/>`));
      state.doc.descendants((n) => {
        if (typeof n.attrs?.xml === 'string' && n.attrs.shape) for (const id of drawingIds([n.attrs.xml])) used.add(id);
        return true;
      });
      const tr = state.tr;
      for (const { pos, node } of dup) {
        const { xml, model } = reidentified(node.attrs.xml, node.attrs.shape as ShapeModel, used);
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, xml, shape: { ...model, key: newShapeKey() } }, node.marks);
      }
      return tr.setMeta(NO_TRACK, true);
    },
  });
}
