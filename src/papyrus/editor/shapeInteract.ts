// Working with shapes on the page (phase 8B): select (click, Shift+click for more, Tab from one
// to the next), drag to move, drag a handle to resize (Shift keeps the proportions), arrow keys
// to nudge (Ctrl+arrow finer), drag a connector's end onto a shape's connection site, a shape
// on a canvas moved with its connectors, 插入 › 圖案 placed with a click or a drag, align,
// distribute, restack, group and ungroup. Each change is one undo step in the editor holding the
// shape (the body, or the header being edited); what changes is written back on save
// (docx/shapeWrite.ts). While 追蹤修訂 is on, moving and changing shapes is refused, as it can't be
// recorded (Word applies such changes without recording them); inserting and deleting are recorded.

import { NodeSelection, Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from './schema';
import { NO_TRACK } from './trackChanges';
import { areaStart, measureFrame, placeShapes, type Rect } from './shapeView';
import { findShape } from './shapeEdit';
import { connectionSites, isLinePreset, routeConnector } from '../docx/shapeGeometry';
import {
  EMU_CM, EMU_PX, SHAPE_KINDS, changedKid, drawingIds, filled, groupXml, moved, newDrawingId, newShapeXml, outlined, readRunShape, resized,
  routed, siteAt, stacked, ungroupXml, wrapOf, wrapped, type ShapeKind, type WrapChoice,
} from '../docx/shapeOps';
import { changedModel } from '../docx/shapeWrite';
import { currentShapeXml } from '../docx/writer';
import { textFrames, type Connection, type ShapeModel, type ShapeNode, type SpNode } from '../docx/shapes';
import type { ThemeColors } from '../docx/theme';
import { tl } from '../i18n';

export const TRACKING_REFUSED = '追蹤修訂時不能移動或變更圖形，請先關閉追蹤修訂';

/** Arrow keys move a shape this far (cm); with Ctrl, a tenth. */
const NUDGE_CM = 0.2;

export interface ShapeToolsOptions {
  editable(): boolean;
  tracking(): boolean;
  notice(message: string): void;
  /** Said to screen readers (a live region). */
  announce(message: string): void;
  /** Edit a text box's text (see DocxEditor.editShapeText). */
  editText(key: string, index: number): boolean;
  theme(): ThemeColors | undefined;
  /** Every document holding shapes (body, headers, footers): new ids avoid theirs. */
  documents(): PMNode[];
}

interface ToolsState {
  /** Selected shapes (the NodeSelection's first), by model key. */
  keys: string[];
  /** A shape on the primary shape's canvas (or in its group): its index among the top group's children. */
  kid: number | null;
}

export const shapeToolsKey = new PluginKey<ToolsState>('dx-shape-tools');

const SELECT = 'dx-shapes-select';

/** The shape keys selected in a state: the selected shape and those Shift+clicked with it. */
export function selectedShapeKeys(state: EditorState): string[] {
  return shapeToolsKey.getState(state)?.keys ?? [];
}

function primaryKey(state: EditorState): string | null {
  const sel = state.selection;
  return sel instanceof NodeSelection && sel.node.attrs.shape ? (sel.node.attrs.shape as ShapeModel).key : null;
}

// ----- measuring on the page -----

interface Placed {
  key: string;
  pos: number;
  node: PMNode;
  shape: ShapeModel;
  el: HTMLElement;
  /** Its box on the page (page px: client / zoom). */
  box: Rect;
  zoom: number;
}

const w = (r: Rect) => r.right - r.left;
const h = (r: Rect) => r.bottom - r.top;
const cm = (px: number) => Math.round(((px * EMU_PX) / EMU_CM) * 100) / 100;

function shapeElement(view: EditorView, key: string): HTMLElement | null {
  return view.dom.querySelector<HTMLElement>(`.dx-shape[data-shape-key="${key}"]`);
}

function zoomOf(el: Element): number {
  const canvas = el.closest<HTMLElement>('.dx-canvas');
  return canvas && canvas.offsetWidth ? canvas.getBoundingClientRect().width / canvas.offsetWidth || 1 : 1;
}

function pageRect(el: Element, zoom: number): Rect {
  const r = el.getBoundingClientRect();
  return { left: r.left / zoom, top: r.top / zoom, right: r.right / zoom, bottom: r.bottom / zoom };
}

function placed(view: EditorView, key: string): Placed | null {
  const found = findShape(view.state.doc, key);
  const el = shapeElement(view, key);
  const box = el?.querySelector<HTMLElement>(':scope > .dx-shape-box');
  if (!found || !el || !box) return null;
  const zoom = zoomOf(el);
  return { key, pos: found.pos, node: found.node, shape: found.shape, el, box: pageRect(box, zoom), zoom };
}

/** Where an anchored shape's references start on the page (page px), from placeShapes' note on it. */
function refOrigin(p: Placed): { x: number; y: number } | null {
  const [hx, vy] = (p.el.dataset.ref ?? '').split(',').map(Number);
  if (!Number.isFinite(hx) || !Number.isFinite(vy)) return null;
  const o = pageRect(p.el, p.zoom);
  return { x: o.left + hx, y: o.top + vy };
}

/** A shape's position as offsets from its references (EMU), where its box is now. */
function offsetsNow(p: Placed): { h: number; v: number } | null {
  const o = refOrigin(p);
  return o ? { h: (p.box.left - o.x) * EMU_PX, v: (p.box.top - o.y) * EMU_PX } : null;
}

/** The top group's children of a drawing placed on the page (page px), with the scale EMU → page px. */
function kidBoxes(p: Placed): { node: ShapeNode; box: Rect; sx: number; sy: number }[] {
  const g = p.shape.root;
  if (g.t !== 'grp') return [];
  const sx = g.ch.w ? w(p.box) / g.ch.w : 0;
  const sy = g.ch.h ? h(p.box) / g.ch.h : 0;
  return g.kids.map((k) => ({
    node: k,
    box: { left: p.box.left + (k.x - g.ch.x) * sx, top: p.box.top + (k.y - g.ch.y) * sy, right: p.box.left + (k.x - g.ch.x + k.w) * sx, bottom: p.box.top + (k.y - g.ch.y + k.h) * sy },
    sx,
    sy,
  }));
}

/** A shape node as a box on the page, for siteAt. */
function asNode(n: SpNode, box: Rect): SpNode {
  return { ...n, x: box.left, y: box.top, w: w(box), h: h(box) };
}

// ----- the selection overlay -----

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
type Handle = (typeof HANDLES)[number];

type Drag =
  | { mode: 'move'; view: EditorView; keys: string[]; kid: number | null; x: number; y: number; zoom: number; moved: boolean; copy: boolean }
  | { mode: 'resize'; view: EditorView; key: string; kid: number | null; handle: Handle; x: number; y: number; zoom: number; box: Rect; moved: boolean }
  | { mode: 'end'; view: EditorView; key: string; kid: number | null; end: 0 | 1; x: number; y: number; zoom: number; moved: boolean }
  | { mode: 'place'; view: EditorView; kind: ShapeKind; x: number; y: number; zoom: number; moved: boolean };

/**
 * The shape tools of one editor (body or header): selection, dragging, the ribbon's commands.
 * DocxEditor makes one and puts its plugin in the body's and headers' editors.
 */
export class ShapeTools {
  private layer: HTMLElement | null = null;
  private drag: Drag | null = null;
  private ghost: HTMLElement[] = [];
  /** 插入 › 圖案 waiting for a click or a drag on the page. */
  placing: { kind: ShapeKind; view: EditorView } | null = null;
  /** The editors this is in (body, header being edited). */
  private views = new Set<EditorView>();

  constructor(private opts: ShapeToolsOptions) {}

  // ----- the plugin -----

  plugin(): Plugin<ToolsState> {
    return new Plugin<ToolsState>({
      key: shapeToolsKey,
      state: {
        init: (_c, state) => ({ keys: primaryKey(state) ? [primaryKey(state)!] : [], kid: null }),
        apply: (tr, value, _old, state) => {
          const meta = tr.getMeta(SELECT) as ToolsState | undefined;
          const primary = primaryKey(state);
          if (meta) return { keys: meta.keys.filter((k) => !!findShape(state.doc, k)), kid: meta.kid };
          if (!primary) return value.keys.length ? { keys: [], kid: null } : value;
          if (value.keys[0] === primary) return tr.docChanged ? { ...value, keys: value.keys.filter((k) => !!findShape(state.doc, k)) } : value;
          return { keys: [primary], kid: null };
        },
      },
      props: {
        handleKeyDown: (view, e) => this.onKey(view, e),
        handleDOMEvents: {
          mousedown: (view, e) => this.onMouseDown(view, e as MouseEvent),
        },
      },
      view: (view) => {
        this.views.add(view);
        return {
          update: () => this.drawSoon(),
          destroy: () => {
            this.views.delete(view);
            this.draw();
          },
        };
      },
    });
  }

  /** The editor whose shapes are selected (the one with a selection), or null. */
  private selectedView(): EditorView | null {
    for (const v of this.views) if (selectedShapeKeys(v.state).length) return v;
    return null;
  }

  selected(): { view: EditorView; items: Placed[]; kid: number | null } | null {
    const view = this.selectedView();
    if (!view) return null;
    // Measured where they are drawn now (a shape drawn again is placed at the next frame).
    placeShapes(view.dom);
    const items = selectedShapeKeys(view.state).map((k) => placed(view, k)).filter((p): p is Placed => !!p);
    return items.length ? { view, items, kid: shapeToolsKey.getState(view.state)?.kid ?? null } : null;
  }

  private select(view: EditorView, keys: string[], kid: number | null = null): void {
    const first = keys[0] ? findShape(view.state.doc, keys[0]) : null;
    const tr = view.state.tr.setMeta(SELECT, { keys, kid }).setMeta('addToHistory', false);
    if (first) tr.setSelection(NodeSelection.create(view.state.doc, first.pos));
    view.dispatch(tr);
  }

  // ----- drawing the selection -----

  private layerFor(view: EditorView): HTMLElement | null {
    const canvas = view.dom.closest<HTMLElement>('.dx-canvas');
    if (!canvas) return null;
    if (!this.layer || this.layer.parentElement !== canvas) {
      this.layer?.remove();
      this.layer = document.createElement('div');
      this.layer.className = 'dx-shape-sel-layer';
      this.layer.setAttribute('aria-hidden', 'true');
      canvas.append(this.layer);
    }
    return this.layer;
  }

  private frame = 0;

  /** Draw after the shapes are placed (their placing waits for the next frame too). */
  drawSoon(): void {
    if (typeof requestAnimationFrame !== 'function') return this.draw();
    cancelAnimationFrame(this.frame);
    this.frame = requestAnimationFrame(() => this.draw());
  }

  /** No handles on the page (printing: the printout copies the page). */
  clearOverlay(): void {
    this.layer?.replaceChildren();
  }

  /** The handles round the selected shapes (called after each update and each layout). */
  draw(): void {
    const sel = this.selected();
    if (!sel) {
      this.layer?.replaceChildren();
      return;
    }
    const layer = this.layerFor(sel.view);
    if (!layer) return;
    const canvas = layer.parentElement!;
    const z = zoomOf(canvas);
    const cr = pageRect(canvas, z);
    const out: HTMLElement[] = [];
    const frame = (r: Rect, cls: string) => {
      const f = document.createElement('div');
      f.className = cls;
      f.style.cssText = `left:${r.left - cr.left}px;top:${r.top - cr.top}px;width:${w(r)}px;height:${h(r)}px`;
      return f;
    };
    sel.items.forEach((p, i) => {
      const kid = i === 0 && sel.kid != null ? kidBoxes(p)[sel.kid] : null;
      const node = kid ? kid.node : p.shape.root;
      const r = kid ? kid.box : p.box;
      const f = frame(r, 'dx-shape-sel' + (i === 0 ? ' dx-shape-sel-main' : ''));
      if (kid) out.push(frame(p.box, 'dx-shape-sel dx-shape-sel-parent'));
      const editable = p.shape.drawable && i === 0 && sel.items.length === 1;
      if (editable && node.t === 'sp' && node.cxn) {
        // A line or connector: its two ends.
        const ends = this.endsOnPage(p, kid);
        ends.forEach((e, k) => {
          const hd = document.createElement('div');
          hd.className = 'dx-shape-handle dx-shape-end';
          hd.dataset.end = String(k);
          hd.style.cssText = `left:${e.x - r.left - 4}px;top:${e.y - r.top - 4}px`;
          this.handleEvents(hd, sel.view, { end: k as 0 | 1 });
          f.append(hd);
        });
      } else if (editable) {
        for (const handle of HANDLES) {
          const hd = document.createElement('div');
          hd.className = `dx-shape-handle dx-h-${handle}`;
          this.handleEvents(hd, sel.view, { handle });
          f.append(hd);
        }
      }
      out.push(f);
    });
    layer.replaceChildren(...out, ...this.ghost);
  }

  private handleEvents(el: HTMLElement, view: EditorView, what: { handle?: Handle; end?: 0 | 1 }): void {
    el.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const sel = this.selected();
      if (!sel) return;
      e.preventDefault();
      e.stopPropagation();
      const p = sel.items[0];
      const z = p.zoom;
      const kid = sel.kid;
      const box = kid != null ? kidBoxes(p)[kid].box : p.box;
      this.drag =
        what.handle != null
          ? { mode: 'resize', view, key: p.key, kid, handle: what.handle, x: e.clientX, y: e.clientY, zoom: z, box, moved: false }
          : { mode: 'end', view, key: p.key, kid, end: what.end!, x: e.clientX, y: e.clientY, zoom: z, moved: false };
      this.listen();
    });
  }

  // ----- mouse -----

  private onMouseDown(view: EditorView, e: MouseEvent): boolean {
    if (e.button !== 0 || !this.opts.editable()) return false;
    if (this.placing) return false;
    const el = (e.target as Element | null)?.closest?.<HTMLElement>('.dx-shape[data-shape-key]');
    if (!el || !view.dom.contains(el)) return false;
    const key = el.dataset.shapeKey!;
    const found = findShape(view.state.doc, key);
    if (!found) return false;
    const keys = selectedShapeKeys(view.state);
    const kidEl = (e.target as Element).closest('[data-kid]');
    const kidIndex = kidEl ? Number(kidEl.getAttribute('data-kid')) : null;
    e.preventDefault();
    view.focus();
    if (e.shiftKey && keys.length) {
      // Shift+click: one more (or one less) of the shapes floating on the page.
      const next = keys.includes(key) ? keys.filter((k) => k !== key) : [...keys, key];
      this.select(view, next.length ? [next[0], ...next.slice(1)] : []);
      return true;
    }
    // A shape on the selected canvas (or group): that one; else the drawing.
    const kid = keys[0] === key && kidIndex != null && found.shape.root.t === 'grp' ? kidIndex : null;
    if (!(keys.includes(key) && kid === (shapeToolsKey.getState(view.state)?.kid ?? null))) this.select(view, keys.includes(key) && kid == null ? [key, ...keys.filter((k) => k !== key)] : [key], kid);
    const now = selectedShapeKeys(view.state);
    if (found.shape.drawable && (found.shape.anchor || kid != null)) {
      this.drag = { mode: 'move', view, keys: now, kid, x: e.clientX, y: e.clientY, zoom: zoomOf(el), moved: false, copy: e.ctrlKey };
      this.listen();
    }
    return true;
  }

  private listen(): void {
    document.addEventListener('mousemove', this.onMove, true);
    document.addEventListener('mouseup', this.onUp, true);
  }

  private unlisten(): void {
    document.removeEventListener('mousemove', this.onMove, true);
    document.removeEventListener('mouseup', this.onUp, true);
  }

  private onMove = (e: MouseEvent) => {
    const d = this.drag;
    if (!d) return;
    const dx = (e.clientX - d.x) / d.zoom;
    const dy = (e.clientY - d.y) / d.zoom;
    if (!d.moved && Math.hypot(dx, dy) < 3) return;
    d.moved = true;
    e.preventDefault();
    const layer = this.layerFor(d.view);
    if (!layer) return;
    const canvas = layer.parentElement!;
    const cr = pageRect(canvas, zoomOf(canvas));
    const rect = (r: Rect, cls = 'dx-shape-ghost') => {
      const g = document.createElement('div');
      g.className = cls;
      g.style.cssText = `left:${r.left - cr.left}px;top:${r.top - cr.top}px;width:${Math.max(1, w(r))}px;height:${Math.max(1, h(r))}px`;
      return g;
    };
    const sel = this.selected();
    this.ghost = [];
    if (d.mode === 'move' && sel) {
      for (const [i, p] of sel.items.entries()) {
        const r = i === 0 && d.kid != null ? kidBoxes(p)[d.kid].box : p.box;
        this.ghost.push(rect({ left: r.left + dx, top: r.top + dy, right: r.right + dx, bottom: r.bottom + dy }));
      }
    } else if (d.mode === 'resize') {
      this.ghost.push(rect(resizeBox(d.box, d.handle, dx, dy, e.shiftKey)));
    } else if (d.mode === 'end' && sel) {
      const p = sel.items[0];
      const ends = this.endsOnPage(p, d.kid != null ? kidBoxes(p)[d.kid] : null);
      const moving = { x: ends[d.end].x + dx, y: ends[d.end].y + dy };
      const snap = this.snap(d.view, p, d.kid, moving);
      const a = d.end === 0 ? snap?.at ?? moving : ends[0];
      const b = d.end === 1 ? snap?.at ?? moving : ends[1];
      const line = rect({ left: Math.min(a.x, b.x), top: Math.min(a.y, b.y), right: Math.max(a.x, b.x), bottom: Math.max(a.y, b.y) }, 'dx-shape-ghost dx-shape-ghost-line');
      this.ghost.push(line);
      if (snap) this.ghost.push(rect({ left: snap.at.x - 5, top: snap.at.y - 5, right: snap.at.x + 5, bottom: snap.at.y + 5 }, 'dx-shape-snap'));
    } else if (d.mode === 'place') {
      const x0 = d.x / d.zoom;
      const y0 = d.y / d.zoom;
      this.ghost.push(rect({ left: Math.min(x0, x0 + dx), top: Math.min(y0, y0 + dy), right: Math.max(x0, x0 + dx), bottom: Math.max(y0, y0 + dy) }));
    }
    this.draw();
  };

  private onUp = (e: MouseEvent) => {
    const d = this.drag;
    this.drag = null;
    this.unlisten();
    this.ghost = [];
    if (!d) return;
    const dx = (e.clientX - d.x) / d.zoom;
    const dy = (e.clientY - d.y) / d.zoom;
    if (d.mode === 'place') {
      this.finishPlacing(d, e.clientX, e.clientY);
      return;
    }
    if (!d.moved) {
      this.draw();
      return;
    }
    if (d.mode === 'move' && d.copy && d.kid == null) this.duplicate(d.view, dx, dy);
    else if (d.mode === 'move') this.moveBy(d.view, dx, dy);
    else if (d.mode === 'resize') this.resizeTo(d.view, resizeBox(d.box, d.handle, dx, dy, e.shiftKey));
    else if (d.mode === 'end') this.moveEnd(d.view, d.end, dx, dy);
    this.draw();
  };

  // ----- keyboard -----

  private onKey(view: EditorView, e: KeyboardEvent): boolean {
    if (this.placing && e.key === 'Escape') {
      this.cancelPlacing();
      return true;
    }
    if (this.placing && e.key === 'Enter') {
      this.placeAtCursor();
      return true;
    }
    const keys = selectedShapeKeys(view.state);
    if (!keys.length || !this.opts.editable()) return false;
    const arrows: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    if (arrows[e.key] && !e.altKey && !e.metaKey) {
      const sel = this.selected();
      if (!sel || (!sel.items[0].shape.anchor && sel.kid == null)) return false;
      const step = ((e.ctrlKey ? NUDGE_CM / 10 : NUDGE_CM) * EMU_CM) / EMU_PX;
      const [ax, ay] = arrows[e.key];
      this.moveBy(view, ax * step, ay * step);
      return true;
    }
    if (e.key === 'Tab' && !e.ctrlKey && !e.altKey) {
      this.selectNext(view, e.shiftKey ? -1 : 1);
      return true;
    }
    if (e.key === 'Escape') {
      const found = findShape(view.state.doc, keys[0]);
      const tr = view.state.tr.setMeta(SELECT, { keys: [], kid: null }).setMeta('addToHistory', false);
      if (found) tr.setSelection(TextSelection.near(view.state.doc.resolve(found.pos + 1)));
      view.dispatch(tr);
      this.opts.announce(tl('已回到文字。'));
      return true;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && keys.length > 1) {
      this.deleteSelected(view);
      return true;
    }
    if ((e.key === 'g' || e.key === 'G') && e.ctrlKey && !e.altKey) {
      if (e.shiftKey) this.ungroup();
      else this.group();
      return true;
    }
    return false;
  }

  /** Select the next (previous) shape after the cursor or the selected one (Tab / 「選取圖形」). */
  selectNext(view: EditorView, dir: 1 | -1 = 1): boolean {
    const all: { pos: number; shape: ShapeModel }[] = [];
    view.state.doc.descendants((n, pos) => {
      if (n.type === schema.nodes.raw_inline && n.attrs.shape) all.push({ pos, shape: n.attrs.shape as ShapeModel });
      return true;
    });
    if (!all.length) {
      this.opts.announce(tl('這裡沒有圖形。'));
      return false;
    }
    const at = view.state.selection.from;
    const current = selectedShapeKeys(view.state)[0];
    let i: number;
    if (current) {
      const k = all.findIndex((s) => s.shape.key === current);
      i = (k + dir + all.length) % all.length;
    } else {
      i = dir > 0 ? all.findIndex((s) => s.pos >= at) : all.map((s) => s.pos < at).lastIndexOf(true);
      if (i < 0) i = dir > 0 ? 0 : all.length - 1;
    }
    const s = all[i];
    this.select(view, [s.shape.key]);
    view.dispatch(view.state.tr.scrollIntoView().setMeta('addToHistory', false));
    const p = placed(view, s.shape.key);
    this.opts.announce(tl('已選取{0}。{1}{2}Esc 回到文字。', describe(s.shape, p), s.shape.anchor ? tl('方向鍵移動，') : '', textFrames(s.shape).length && s.shape.drawable ? tl('Enter 編輯文字，') : ''));
    return true;
  }

  // ----- changes -----

  /** Refused while tracking, with the notice. */
  private refused(): boolean {
    if (!this.opts.tracking()) return false;
    this.opts.notice(tl('{0}。', tl(TRACKING_REFUSED)));
    return true;
  }

  /** Put new models (by key) and new nodes in one undo step of `view`. */
  private commit(view: EditorView, models: Map<string, ShapeModel>, message?: string, anchors: Map<string, number> = new Map()): boolean {
    if (!models.size) return false;
    const tr = view.state.tr;
    for (const [key, m] of models) {
      const found = findShape(tr.doc, key);
      if (!found) continue;
      const to = anchors.get(key);
      if (to == null) {
        tr.setNodeMarkup(found.pos, undefined, { ...found.node.attrs, shape: m }, found.node.marks);
        continue;
      }
      // Anchored to another paragraph: its run moves to the start of that one, as Word's anchor does.
      // (`to` is a position of the document before this step.)
      tr.delete(found.pos, found.pos + 1);
      tr.insert(tr.mapping.map(to, -1), found.node.type.create({ ...found.node.attrs, shape: m }, null, found.node.marks));
    }
    if (!tr.docChanged) return false;
    // The same shapes stay selected (a changed node is a new node).
    const st = shapeToolsKey.getState(view.state);
    const first = st?.keys[0] ? findShape(tr.doc, st.keys[0]) : null;
    if (first) tr.setSelection(NodeSelection.create(tr.doc, first.pos)).setMeta(SELECT, { keys: st!.keys, kid: st!.kid });
    view.dispatch(closeHistory(tr).setMeta(NO_TRACK, true));
    if (message) this.opts.announce(message);
    return true;
  }

  /**
   * The top-level connectors attached to the shapes whose boxes change (`boxes`, page px by
   * docId), going again from site to site: their new models.
   */
  private rerouted(view: EditorView, boxes: Map<number, { box: Rect; node: SpNode }>): Map<string, ShapeModel> {
    const out = new Map<string, ShapeModel>();
    if (!boxes.size) return out;
    view.state.doc.descendants((n) => {
      const m = n.attrs?.shape as ShapeModel | undefined;
      const c = m?.root;
      if (!m || !m.anchor || !c || c.t !== 'sp' || !c.cxn) return true;
      if (!((c.st && boxes.has(c.st.id)) || (c.end && boxes.has(c.end.id)))) return true;
      const p = placed(view, m.key);
      if (!p) return true;
      const ends = this.endsOnPage(p, null);
      const site = (conn: Connection | undefined, fallback: { x: number; y: number }) => {
        const target = conn ? boxes.get(conn.id) ?? this.shapeBoxById(view, conn.id) : null;
        return target ? siteAt(asNode(target.node, target.box), conn!.idx) ?? { ...fallback, dir: undefined } : { ...fallback, dir: undefined };
      };
      const a = site(c.st, ends[0]);
      const b = site(c.end, ends[1]);
      out.set(m.key, this.connectorAt(p, a, b, a.dir));
      return true;
    });
    return out;
  }

  /** A top-level shape's box on the page and its node, by its drawing id. */
  private shapeBoxById(view: EditorView, docId: number): { box: Rect; node: SpNode } | null {
    let hit: { box: Rect; node: SpNode } | null = null;
    view.state.doc.descendants((n) => {
      const m = n.attrs?.shape as ShapeModel | undefined;
      if (hit || !m || m.docId !== docId || m.root.t !== 'sp') return !hit;
      const p = placed(view, m.key);
      if (p) hit = { box: p.box, node: m.root };
      return false;
    });
    return hit;
  }

  /** A top-level connector's model going from a to b (page px). */
  private connectorAt(p: Placed, a: { x: number; y: number }, b: { x: number; y: number }, da?: number): ShapeModel {
    const m = p.shape;
    const root = m.root as SpNode;
    const box = routeConnector(root.geom, a, b, da);
    const o = refOrigin(p) ?? { x: p.box.left, y: p.box.top };
    const cw = Math.round(box.w * EMU_PX);
    const ch = Math.round(box.h * EMU_PX);
    const r: SpNode = { ...root, x: 0, y: 0, w: cw, h: ch };
    delete r.rot;
    delete r.flipH;
    delete r.flipV;
    if (box.rot) r.rot = box.rot;
    if (box.flipH) r.flipH = true;
    if (box.flipV) r.flipV = true;
    if (box.av && root.geom.startsWith('bentConnector')) r.av = box.av;
    const anchor = { ...m.anchor!, h: { rel: m.anchor!.h.rel, offset: Math.round((box.x - o.x) * EMU_PX) }, v: { rel: m.anchor!.v.rel, offset: Math.round((box.y - o.y) * EMU_PX) } };
    return changedModel(m, { ...m, w: cw, h: ch, root: r, anchor });
  }

  /** A connector's ends on the page (page px): a top-level one, or one on the selected canvas. */
  private endsOnPage(p: Placed, kid: { node: ShapeNode; box: Rect } | null): { x: number; y: number }[] {
    const n = (kid ? kid.node : p.shape.root) as SpNode;
    const box = kid ? kid.box : p.box;
    const at = (s: { x: number; y: number }) => {
      let x = n.flipH ? 1 - s.x : s.x;
      let y = n.flipV ? 1 - s.y : s.y;
      const t = ((n.rot ?? 0) * Math.PI) / 180;
      x = (x - 0.5) * w(box);
      y = (y - 0.5) * h(box);
      return { x: box.left + w(box) / 2 + x * Math.cos(t) - y * Math.sin(t), y: box.top + h(box) / 2 + x * Math.sin(t) + y * Math.cos(t) };
    };
    return [at({ x: 0, y: 0 }), at({ x: 1, y: 1 })];
  }

  /** The connection site near a point (page px) to attach a connector's end to. */
  private snap(view: EditorView, p: Placed, kid: number | null, at: { x: number; y: number }): { at: { x: number; y: number; dir: number }; conn: Connection } | null {
    const radius = 10;
    let best: { at: { x: number; y: number; dir: number }; conn: Connection } | null = null;
    let dist = radius;
    const consider = (node: SpNode, box: Rect, id: number | undefined) => {
      if (id == null || node.cxn || isLinePreset(node.geom)) return;
      connectionSites(node.geom, 1, 1).forEach((_s, idx) => {
        const site = siteAt(asNode(node, box), idx);
        if (!site) return;
        const d = Math.hypot(site.x - at.x, site.y - at.y);
        if (d <= dist) {
          dist = d;
          best = { at: site, conn: { id, idx } };
        }
      });
    };
    if (kid != null) {
      // On a canvas: its other shapes.
      kidBoxes(p).forEach((k, i) => {
        if (i !== kid && k.node.t === 'sp') consider(k.node, k.box, k.node.id);
      });
    } else {
      view.state.doc.descendants((n) => {
        const m = n.attrs?.shape as ShapeModel | undefined;
        if (!m || m.key === p.key || !m.anchor || m.root.t !== 'sp' || !m.drawable) return true;
        const q = placed(view, m.key);
        if (q) consider(m.root, q.box, m.docId);
        return true;
      });
    }
    return best;
  }

  /** Move the selected shapes (or the selected shape on a canvas) by (dx, dy) page px. */
  moveBy(view: EditorView, dx: number, dy: number): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const models = new Map<string, ShapeModel>();
    const anchors = new Map<string, number>();
    const boxes = new Map<number, { box: Rect; node: SpNode }>();
    if (sel.kid != null) {
      const p = sel.items[0];
      const k = kidBoxes(p)[sel.kid];
      if (!k || !k.sx || !k.sy) return false;
      models.set(p.key, changedKid(p.shape, sel.kid, (n) => ({ ...n, x: Math.round(n.x + dx / k.sx), y: Math.round(n.y + dy / k.sy) })));
    } else {
      for (const p of sel.items) {
        const at = p.shape.anchor ? offsetsNow(p) : null;
        if (!at) continue;
        const box = shift(p.box, dx, dy);
        // As in Word, a shape dragged elsewhere takes the paragraph nearest its new place as its
        // anchor (unless its anchor is locked), placed from that paragraph.
        const again = p.shape.anchor!.locked ? null : this.reanchor(view, p, box);
        if (again) {
          models.set(p.key, moved(p.shape, again.h, again.v));
          anchors.set(p.key, again.to);
        } else models.set(p.key, moved(p.shape, at.h + dx * EMU_PX, at.v + dy * EMU_PX));
        if (p.shape.docId != null && p.shape.root.t === 'sp') boxes.set(p.shape.docId, { box, node: p.shape.root });
      }
      for (const [k, m] of this.rerouted(view, boxes)) if (!models.has(k)) models.set(k, m);
    }
    const p = sel.items[0];
    return this.commit(view, models, sel.kid != null ? tl('已移動畫布上的圖形。') : tl('已移動圖形：{0}', position(p, dx, dy)), anchors);
  }

  /**
   * The paragraph nearest a point on the page (page px): its element and where its content
   * starts in the document; null where there is none (or no layout to tell).
   */
  paragraphAt(view: EditorView, x: number, y: number): { el: HTMLElement; start: number } | null {
    const z = zoomOf(view.dom);
    let hit: { pos: number; inside: number } | null = null;
    try {
      hit = view.posAtCoords({ left: x * z, top: y * z });
    } catch {
      return null;
    }
    if (!hit) return null;
    const $p = view.state.doc.resolve(hit.inside >= 0 && view.state.doc.nodeAt(hit.inside)?.isInline ? hit.inside : hit.pos);
    if ($p.parent.type !== schema.nodes.paragraph) return null;
    const el = view.nodeDOM($p.before());
    return el instanceof HTMLElement ? { el, start: $p.start() } : null;
  }

  /** A shape moved to `box` (page px): the paragraph to anchor it to now and its offsets from that one (EMU); null to keep its anchor. */
  private reanchor(view: EditorView, p: Placed, box: Rect): { to: number; h: number; v: number } | null {
    const target = this.paragraphAt(view, box.left, box.top);
    if (!target) return null;
    const $now = view.state.doc.resolve(p.pos);
    if ($now.parent.type === schema.nodes.paragraph && $now.start() === target.start) return null;
    const a = p.shape.anchor!;
    const { frame } = measureFrame(target.el);
    return {
      to: target.start,
      h: Math.round((box.left - areaStart(a.h.rel, frame, true)) * EMU_PX),
      v: Math.round((box.top - areaStart(a.v.rel, frame, false)) * EMU_PX),
    };
  }

  /** Ctrl+drag: a copy of each selected floating shape, (dx, dy) page px away, selected (recorded while 追蹤修訂 is on). */
  duplicate(view: EditorView, dx: number, dy: number): boolean {
    const sel = this.selected();
    if (!sel) return false;
    const tr = view.state.tr;
    const keys: string[] = [];
    let first: number | null = null;
    for (const p of sel.items) {
      const at = p.shape.anchor ? offsetsNow(p) : null;
      const model = at ? readRunShape(currentShapeXml(p.node), this.opts.theme()) : null;
      if (!model || !at) continue;
      const xml = currentShapeXml(p.node);
      const copy = moved(model, at.h + dx * EMU_PX, at.v + dy * EMU_PX);
      const pos = tr.mapping.map(p.pos + 1);
      tr.insert(pos, schema.nodes.raw_inline.create({ ...p.node.attrs, xml, shape: copy }, null, p.node.marks));
      keys.push(copy.key);
      first ??= pos;
    }
    if (first == null) return false;
    tr.setSelection(NodeSelection.create(tr.doc, first)).setMeta(SELECT, { keys, kid: null });
    view.dispatch(closeHistory(tr));
    this.opts.announce(tl('已複製 {0} 個圖形。', keys.length));
    return true;
  }

  /** Resize the selected shape (or the selected one on a canvas) to a box on the page (page px). */
  resizeTo(view: EditorView, box: Rect): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const p = sel.items[0];
    const models = new Map<string, ShapeModel>();
    if (sel.kid != null) {
      const k = kidBoxes(p)[sel.kid];
      if (!k || !k.sx || !k.sy) return false;
      const g = p.shape.root as { ch: { x: number; y: number } };
      models.set(p.key, changedKid(p.shape, sel.kid, (n) => ({ ...n, x: Math.round(g.ch.x + (box.left - p.box.left) / k.sx), y: Math.round(g.ch.y + (box.top - p.box.top) / k.sy), w: Math.round(w(box) / k.sx), h: Math.round(h(box) / k.sy) })));
    } else {
      let m = resized(p.shape, w(box) * EMU_PX, h(box) * EMU_PX);
      const at = p.shape.anchor ? offsetsNow(p) : null;
      if (at && (box.left !== p.box.left || box.top !== p.box.top)) m = moved(m, at.h + (box.left - p.box.left) * EMU_PX, at.v + (box.top - p.box.top) * EMU_PX);
      models.set(p.key, m);
      if (p.shape.docId != null && p.shape.root.t === 'sp') {
        for (const [k, c] of this.rerouted(view, new Map([[p.shape.docId, { box, node: p.shape.root }]]))) models.set(k, c);
      }
    }
    return this.commit(view, models, tl('已調整圖形大小：寬 {0} 公分，高 {1} 公分。', cm(w(box)), cm(h(box))));
  }

  /** Drag a connector's end by (dx, dy) page px, attaching it to a connection site it reaches. */
  private moveEnd(view: EditorView, end: 0 | 1, dx: number, dy: number): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const p = sel.items[0];
    const kid = sel.kid != null ? kidBoxes(p)[sel.kid] : null;
    const ends = this.endsOnPage(p, kid);
    const moving = { x: ends[end].x + dx, y: ends[end].y + dy };
    const snap = this.snap(view, p, sel.kid, moving);
    const a = end === 0 ? snap?.at ?? moving : ends[0];
    const b = end === 1 ? snap?.at ?? moving : ends[1];
    const conn = snap?.conn;
    const attach = (n: SpNode): SpNode => {
      const next = { ...n };
      if (end === 0) {
        if (conn) next.st = conn;
        else delete next.st;
      } else if (conn) next.end = conn;
      else delete next.end;
      return next;
    };
    const models = new Map<string, ShapeModel>();
    if (kid) {
      const g = p.shape.root as { ch: { x: number; y: number } };
      const toKid = (q: { x: number; y: number }) => ({ x: g.ch.x + (q.x - p.box.left) / kid.sx, y: g.ch.y + (q.y - p.box.top) / kid.sy });
      models.set(p.key, changedKid(p.shape, sel.kid!, (n) => attach(routed(n as SpNode, toKid(a), toKid(b), end === 0 ? snap?.at.dir : undefined))));
    } else {
      const next = this.connectorAt(p, a, b, end === 0 ? snap?.at.dir : undefined);
      models.set(p.key, { ...next, root: attach(next.root as SpNode) });
    }
    return this.commit(view, models, snap ? tl('已連接到圖形。') : tl('已移動線條的端點。'));
  }

  /** Set the selected shape's position (cm from its references) and size (cm). */
  setGeometry(g: { x?: number; y?: number; w?: number; h?: number }): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const p = sel.items[0];
    let m = p.shape;
    if (g.w != null || g.h != null) m = resized(m, (g.w ?? cm(w(p.box))) * EMU_CM, (g.h ?? cm(h(p.box))) * EMU_CM);
    if ((g.x != null || g.y != null) && m.anchor) {
      const at = offsetsNow(p) ?? { h: m.anchor.h.offset ?? 0, v: m.anchor.v.offset ?? 0 };
      m = moved(m, g.x != null ? g.x * EMU_CM : at.h, g.y != null ? g.y * EMU_CM : at.v);
    }
    return this.commit(sel.view, new Map([[p.key, m]]), tl('已變更圖形的位置與大小。'));
  }

  /** Change every selected shape with `f` (fill, outline, wrapping ...). */
  changeSelected(f: (m: ShapeModel, p: Placed) => ShapeModel, message: string): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const models = new Map<string, ShapeModel>();
    for (const p of sel.items) if (p.shape.drawable) models.set(p.key, f(p.shape, p));
    return this.commit(sel.view, models, message);
  }

  setFill(color: string | null): boolean {
    return this.changeSelected((m) => filled(m, color ? color.replace('#', '').toUpperCase() : null), color ? tl('已變更填滿色彩。') : tl('已移除填滿。'));
  }

  setOutline(color: string | null, widthPt?: number): boolean {
    return this.changeSelected((m) => outlined(m, color ? color.replace('#', '').toUpperCase() : null, widthPt != null ? Math.round(widthPt * 12700) : undefined), color ? tl('已變更外框。') : tl('已移除外框。'));
  }

  setWrap(choice: WrapChoice): boolean {
    const z = this.topZ() + 1;
    return this.changeSelected((m, p) => {
      if (m.anchor || choice === 'inline') return wrapped(m, choice);
      // Out of the line of text: placed where it is, from the column and its paragraph.
      const { frame } = measureFrame(p.el);
      return wrapped(m, choice, { h: (p.box.left - areaStart('column', frame, true)) * EMU_PX, v: (p.box.top - areaStart('paragraph', frame, false)) * EMU_PX }, z);
    }, tl('文繞圖：{0}。', tl(WRAP_LABEL[choice])));
  }

  /** The highest stacking order of the shapes in the selected editor's document. */
  private topZ(view = this.selected()?.view ?? null): number {
    let z = 251659264;
    view?.state.doc.descendants((n) => {
      const a = (n.attrs?.shape as ShapeModel | undefined)?.anchor;
      if (a) z = Math.max(z, a.z);
      return true;
    });
    return z;
  }

  /** 上移一層 (dir 1) / 下移一層 (-1): trade places in the stacking order with the next shape. */
  restack(dir: 1 | -1): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel || sel.items.length !== 1 || !sel.items[0].shape.anchor) return false;
    const p = sel.items[0];
    const all: ShapeModel[] = [];
    sel.view.state.doc.descendants((n) => {
      const m = n.attrs?.shape as ShapeModel | undefined;
      if (m?.anchor && m.anchor.behind === p.shape.anchor!.behind) all.push(m);
      return true;
    });
    all.sort((a, b) => a.anchor!.z - b.anchor!.z);
    const i = all.findIndex((m) => m.key === p.key);
    const other = all[i + dir];
    if (!other) {
      this.opts.announce(dir > 0 ? tl('已在最上層。') : tl('已在最下層。'));
      return false;
    }
    const mine = p.shape.anchor!.z;
    const theirs = other.anchor!.z === mine ? mine + dir : other.anchor!.z;
    return this.commit(sel.view, new Map([[p.key, stacked(p.shape, theirs)], [other.key, stacked(other, mine)]]), dir > 0 ? tl('已上移一層。') : tl('已下移一層。'));
  }

  /** 對齊: the selected shapes to one another (one shape: to the margins). */
  align(how: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const items = sel.items.filter((p) => p.shape.anchor);
    if (!items.length) return false;
    let area: Rect;
    if (items.length === 1) area = measureFrame(items[0].el).frame.column;
    else area = { left: Math.min(...items.map((p) => p.box.left)), top: Math.min(...items.map((p) => p.box.top)), right: Math.max(...items.map((p) => p.box.right)), bottom: Math.max(...items.map((p) => p.box.bottom)) };
    const models = new Map<string, ShapeModel>();
    const boxes = new Map<number, { box: Rect; node: SpNode }>();
    for (const p of items) {
      const at = offsetsNow(p);
      if (!at) continue;
      let dx = 0;
      let dy = 0;
      if (how === 'left') dx = area.left - p.box.left;
      else if (how === 'right') dx = area.right - p.box.right;
      else if (how === 'center') dx = (area.left + area.right) / 2 - (p.box.left + p.box.right) / 2;
      else if (how === 'top') dy = area.top - p.box.top;
      else if (how === 'bottom') dy = area.bottom - p.box.bottom;
      else dy = (area.top + area.bottom) / 2 - (p.box.top + p.box.bottom) / 2;
      models.set(p.key, moved(p.shape, at.h + dx * EMU_PX, at.v + dy * EMU_PX));
      if (p.shape.docId != null && p.shape.root.t === 'sp') boxes.set(p.shape.docId, { box: shift(p.box, dx, dy), node: p.shape.root });
    }
    for (const [k, m] of this.rerouted(sel.view, boxes)) if (!models.has(k)) models.set(k, m);
    return this.commit(sel.view, models, tl('已{0}。', tl(ALIGN_LABEL[how])));
  }

  /** 均分: three or more selected shapes, the gaps between them made equal. */
  distribute(axis: 'h' | 'v'): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel) return false;
    const items = sel.items.filter((p) => p.shape.anchor).sort((a, b) => (axis === 'h' ? a.box.left - b.box.left : a.box.top - b.box.top));
    if (items.length < 3) {
      this.opts.notice(tl('請選取三個以上的圖形再均分。'));
      return false;
    }
    const size = (r: Rect) => (axis === 'h' ? w(r) : h(r));
    const start = axis === 'h' ? items[0].box.left : items[0].box.top;
    const end = axis === 'h' ? items[items.length - 1].box.right : items[items.length - 1].box.bottom;
    const gap = (end - start - items.reduce((s, p) => s + size(p.box), 0)) / (items.length - 1);
    const models = new Map<string, ShapeModel>();
    const boxes = new Map<number, { box: Rect; node: SpNode }>();
    let at = start;
    for (const p of items) {
      const o = offsetsNow(p);
      const d = at - (axis === 'h' ? p.box.left : p.box.top);
      at += size(p.box) + gap;
      if (!o || Math.abs(d) < 0.01) continue;
      models.set(p.key, moved(p.shape, o.h + (axis === 'h' ? d : 0) * EMU_PX, o.v + (axis === 'v' ? d : 0) * EMU_PX));
      if (p.shape.docId != null && p.shape.root.t === 'sp') boxes.set(p.shape.docId, { box: shift(p.box, axis === 'h' ? d : 0, axis === 'v' ? d : 0), node: p.shape.root });
    }
    for (const [k, m] of this.rerouted(sel.view, boxes)) if (!models.has(k)) models.set(k, m);
    return this.commit(sel.view, models, axis === 'h' ? tl('已水平均分。') : tl('已垂直均分。')) || true;
  }

  /** Whether the selection can be grouped / ungrouped. */
  canGroup(): boolean {
    const sel = this.selected();
    return !!sel && sel.items.length > 1 && sel.items.every((p) => p.shape.anchor && p.shape.drawable && p.shape.root.t === 'sp' && p.shape.src === 'dml');
  }

  canUngroup(): boolean {
    const sel = this.selected();
    return !!sel && sel.items.length === 1 && sel.items[0].shape.kind === 'group' && sel.items[0].shape.src === 'dml' && !!sel.items[0].shape.anchor;
  }

  /** 群組: the selected floating shapes become one group where the first one is anchored. */
  group(): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel || !this.canGroup()) {
      if (sel) this.opts.notice(tl('請選取兩個以上浮動的圖形再群組。'));
      return false;
    }
    const view = sel.view;
    const used = this.usedIds();
    const items = sel.items;
    const xml = groupXml(
      items.map((p) => ({ model: p.shape, x: p.box.left * EMU_PX, y: p.box.top * EMU_PX, xml: currentShapeXml(p.node) })),
      newDrawingId(used),
      items.map(() => newDrawingId(used)),
      (i) => innerText(currentShapeXml(items[i].node)),
    );
    const model = readRunShape(xml, this.opts.theme());
    if (!model) return false;
    const node = schema.nodes.raw_inline.create({ xml, label: '圖形', shape: model }, null, items[0].node.marks);
    const tr = view.state.tr;
    // The group where the first shape was; the others out.
    const positions = items.map((p) => p.pos).sort((a, b) => b - a);
    for (const pos of positions) if (pos !== items[0].pos) tr.delete(tr.mapping.map(pos), tr.mapping.map(pos) + 1);
    const at = tr.mapping.map(items[0].pos);
    tr.replaceWith(at, at + 1, node);
    tr.setSelection(NodeSelection.create(tr.doc, at)).setMeta(SELECT, { keys: [model.key], kid: null });
    view.dispatch(closeHistory(tr).setMeta(NO_TRACK, true));
    this.opts.announce(tl('已將 {0} 個圖形群組。', items.length));
    return true;
  }

  /** 取消群組: the selected group's shapes become shapes of their own, in place. */
  ungroup(): boolean {
    if (this.refused()) return false;
    const sel = this.selected();
    if (!sel || !this.canUngroup()) return false;
    const p = sel.items[0];
    const used = this.usedIds();
    const kids = p.shape.root.t === 'grp' ? p.shape.root.kids.length : 0;
    const parts = ungroupXml(currentShapeXml(p.node), p.shape, Array.from({ length: kids }, () => newDrawingId(used)));
    if (!parts) {
      this.opts.notice(tl('這個群組無法在網頁上取消群組，請在 Word 中進行。'));
      return false;
    }
    const nodes = parts.map((q) => schema.nodes.raw_inline.create({ xml: q.xml, label: '圖形', shape: readRunShape(q.xml, this.opts.theme()) }, null, p.node.marks));
    const tr = sel.view.state.tr.replaceWith(p.pos, p.pos + 1, nodes);
    const keys = nodes.map((n) => (n.attrs.shape as ShapeModel).key);
    tr.setSelection(NodeSelection.create(tr.doc, p.pos)).setMeta(SELECT, { keys, kid: null });
    sel.view.dispatch(closeHistory(tr).setMeta(NO_TRACK, true));
    this.opts.announce(tl('已取消群組：{0} 個圖形。', nodes.length));
    return true;
  }

  /** Delete the selected shapes (recorded while 追蹤修訂 is on). */
  deleteSelected(view = this.selected()?.view ?? null): boolean {
    if (!view) return false;
    const keys = selectedShapeKeys(view.state);
    const positions = keys.map((k) => findShape(view.state.doc, k)?.pos).filter((p): p is number => p != null).sort((a, b) => b - a);
    if (!positions.length) return false;
    const tr = view.state.tr;
    for (const pos of positions) tr.delete(pos, pos + 1);
    tr.setMeta(SELECT, { keys: [], kid: null });
    view.dispatch(tr);
    this.opts.announce(tl('已刪除 {0} 個圖形。', positions.length));
    return true;
  }

  private usedIds(): Set<number> {
    const xmls: string[] = [];
    for (const d of this.opts.documents()) {
      d.descendants((n) => {
        if (typeof n.attrs?.xml === 'string' && n.attrs.xml.includes('docPr')) xmls.push(n.attrs.xml);
        return true;
      });
    }
    return drawingIds(xmls);
  }

  // ----- 插入 › 圖案 -----

  /** Wait for a click or a drag on the page to place a new shape of `kind` (Esc cancels, Enter puts it at the cursor). */
  startPlacing(kindId: string, view: EditorView): boolean {
    const kind = SHAPE_KINDS.find((k) => k.id === kindId);
    if (!kind || !this.opts.editable()) return false;
    this.placing = { kind, view };
    const canvas = view.dom.closest<HTMLElement>('.dx-canvas');
    canvas?.classList.add('dx-placing');
    const cover = document.createElement('div');
    cover.className = 'dx-shape-place-cover';
    cover.addEventListener('mousedown', (e) => {
      if (e.button !== 0 || !this.placing) return;
      e.preventDefault();
      this.drag = { mode: 'place', view, kind, x: e.clientX, y: e.clientY, zoom: zoomOf(cover), moved: false };
      this.listen();
    });
    this.placeCover?.remove();
    this.placeCover = cover;
    canvas?.append(cover);
    view.focus();
    this.opts.notice(kind.line ? tl('請在頁面上拖曳畫出{0}（Esc 取消）。', tl(kind.label)) : tl('請在頁面上按一下或拖曳，放置{0}（Enter 放在游標處，Esc 取消）。', tl(kind.label)));
    return true;
  }

  private placeCover: HTMLElement | null = null;

  cancelPlacing(): void {
    this.placing?.view.dom.closest('.dx-canvas')?.classList.remove('dx-placing');
    this.placing = null;
    this.placeCover?.remove();
    this.placeCover = null;
  }

  private finishPlacing(d: Extract<Drag, { mode: 'place' }>, cx: number, cy: number): void {
    this.cancelPlacing();
    const a = { x: d.x / d.zoom, y: d.y / d.zoom };
    const b = { x: cx / d.zoom, y: cy / d.zoom };
    this.insert(d.view, d.kind, a, d.moved ? b : null);
  }

  private placeAtCursor(): void {
    const p = this.placing;
    if (!p) return;
    this.cancelPlacing();
    this.insert(p.view, p.kind, null, null);
  }

  /**
   * Insert a new shape of `kind`, anchored to the paragraph at the cursor: from `a` to `b` (page
   * px; a line from its start to its end), or its default size at `a`, or at the cursor.
   */
  insert(view: EditorView, kind: ShapeKind, a: { x: number; y: number } | null, b: { x: number; y: number } | null): boolean {
    const state = view.state;
    placeShapes(view.dom);
    // Anchored as Word does: to the paragraph nearest where it is put (its run at that
    // paragraph's start); put at the cursor, to the cursor's paragraph.
    const dropped = a ? this.paragraphAt(view, a.x, Math.min(a.y, b?.y ?? a.y)) : null;
    const pos = dropped ? dropped.start : state.selection.from;
    const $pos = state.doc.resolve(pos);
    const paraNode = dropped ? dropped.el : $pos.parent.type === schema.nodes.paragraph ? view.nodeDOM($pos.before()) : null;
    const para = paraNode instanceof HTMLElement && paraNode.classList.contains('dx-p') ? paraNode : null;
    if (!para) {
      this.opts.notice(tl('請先把游標放在要放圖案的段落。'));
      return false;
    }
    const { frame } = measureFrame(para);
    const col = areaStart('column', frame, true);
    const top = frame.paragraph.top;
    const [dw, dh] = kind.size.map((v) => (v * EMU_CM) / EMU_PX);
    let x0 = a?.x ?? col + 1 * (EMU_CM / EMU_PX);
    let y0 = a?.y ?? top;
    let x1 = b?.x ?? x0 + dw;
    let y1 = b?.y ?? y0 + (kind.line && !b ? 0 : dh);
    let st: Connection | undefined;
    let end: Connection | undefined;
    if (kind.line) {
      // A line's ends attach to the connection sites they are dropped on.
      const snapA = this.snapTop(view, { x: x0, y: y0 });
      const snapB = this.snapTop(view, { x: x1, y: y1 });
      if (snapA) {
        x0 = snapA.at.x;
        y0 = snapA.at.y;
        st = snapA.conn;
      }
      if (snapB) {
        x1 = snapB.at.x;
        y1 = snapB.at.y;
        end = snapB.conn;
      }
    }
    const flipH = x1 < x0;
    const flipV = y1 < y0;
    let box = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) };
    let rot: number | undefined;
    if (kind.prst.startsWith('bentConnector')) {
      const r = routeConnector(kind.prst, { x: x0, y: y0 }, { x: x1, y: y1 }, st ? this.siteDir(view, st) : undefined);
      box = { x: r.x, y: r.y, w: r.w, h: r.h };
      rot = r.rot;
    }
    const used = this.usedIds();
    const docId = newDrawingId(used);
    let count = 0;
    state.doc.descendants((n) => {
      if (n.attrs?.shape) count++;
      return true;
    });
    const name = `${kind.label.replace(/^流程圖：/, '')} ${count + 1}`;
    const xml = newShapeXml({
      kind,
      x: Math.round((box.x - col) * EMU_PX),
      y: Math.round((box.y - top) * EMU_PX),
      w: Math.round(box.w * EMU_PX),
      h: Math.round(box.h * EMU_PX),
      flipH: !rot && flipH,
      flipV: !rot && flipV,
      z: this.topZ(view) + 1,
      docId,
      name,
      st,
      end,
    });
    let model = readRunShape(xml, this.opts.theme());
    if (!model) return false;
    if (rot) model = changedModel(model, { ...model, root: { ...(model.root as SpNode), rot } });
    // No marks of the text there: a new shape's run is not inside a link or a field's result.
    const node = schema.nodes.raw_inline.create({ xml, label: '圖形', shape: model });
    const tr = state.tr.insert(pos, node);
    tr.setSelection(NodeSelection.create(tr.doc, pos)).setMeta(SELECT, { keys: [model.key], kid: null });
    // Not scrolled to the cursor: the shape is where it was put, in view.
    view.dispatch(closeHistory(tr));
    view.focus();
    this.opts.announce(tl('已插入{0}。方向鍵移動，{1}Esc 回到文字。', tl(kind.label), kind.line ? '' : tl('Enter 輸入文字，')));
    if (kind.textbox) this.opts.editText(model.key, 0);
    return true;
  }

  /** The connection site of a floating shape near a point (page px), for a new line's end. */
  private snapTop(view: EditorView, at: { x: number; y: number }): { at: { x: number; y: number; dir: number }; conn: Connection } | null {
    let best: { at: { x: number; y: number; dir: number }; conn: Connection } | null = null;
    let dist = 10;
    view.state.doc.descendants((n) => {
      const m = n.attrs?.shape as ShapeModel | undefined;
      if (!m || !m.anchor || m.root.t !== 'sp' || m.root.cxn || m.docId == null || !m.drawable) return true;
      const q = placed(view, m.key);
      if (!q) return true;
      const node = asNode(m.root, q.box);
      connectionSites(node.geom, 1, 1).forEach((_s, idx) => {
        const site = siteAt(node, idx);
        if (!site) return;
        const d = Math.hypot(site.x - at.x, site.y - at.y);
        if (d <= dist) {
          dist = d;
          best = { at: site, conn: { id: m.docId!, idx } };
        }
      });
      return true;
    });
    return best;
  }

  private siteDir(view: EditorView, c: Connection): number | undefined {
    const t = this.shapeBoxById(view, c.id);
    return t ? siteAt(asNode(t.node, t.box), c.idx)?.dir : undefined;
  }

  // ----- for the ribbon -----

  /** What 圖形格式 shows: the selected shape's position and size (cm), wrapping, colours. */
  info(): ShapeInfo | null {
    const sel = this.selected();
    if (!sel) return null;
    const p = sel.items[0];
    const at = p.shape.anchor ? offsetsNow(p) : null;
    const root = p.shape.root;
    const sp = root.t === 'sp' ? root : root.t === 'grp' ? (root.kids.find((k) => k.t === 'sp' && !k.cxn) as SpNode | undefined) : undefined;
    return {
      count: sel.items.length,
      name: describe(p.shape, p),
      drawable: p.shape.drawable,
      floating: !!p.shape.anchor,
      x: at ? cm(at.h / EMU_PX) : null,
      y: at ? cm(at.v / EMU_PX) : null,
      xFrom: p.shape.anchor ? tl(H_LABEL[p.shape.anchor.h.rel] ?? '欄') : null,
      yFrom: p.shape.anchor ? tl(V_LABEL[p.shape.anchor.v.rel] ?? '段落') : null,
      w: cm(w(p.box)),
      h: cm(h(p.box)),
      wrap: wrapOf(p.shape),
      fill: sp?.fill ? '#' + sp.fill.color.toLowerCase() : null,
      line: sp?.line ? '#' + sp.line.color.toLowerCase() : null,
      lineWidth: sp?.line ? Math.round((sp.line.width / 12700) * 100) / 100 : null,
      line1: !!sp && isLinePreset(sp.geom),
      canGroup: this.canGroup(),
      canUngroup: this.canUngroup(),
    };
  }
}

export interface ShapeInfo {
  count: number;
  name: string;
  drawable: boolean;
  floating: boolean;
  /** Position from its references (cm), and what they are (「欄」「段落」「頁面」 ...). */
  x: number | null;
  y: number | null;
  xFrom: string | null;
  yFrom: string | null;
  w: number;
  h: number;
  wrap: WrapChoice;
  fill: string | null;
  line: string | null;
  lineWidth: number | null;
  /** A line or connector (no fill). */
  line1: boolean;
  canGroup: boolean;
  canUngroup: boolean;
}

export const WRAP_LABEL: Record<WrapChoice, string> = { inline: '與文字排列', square: '矩形', topAndBottom: '上及下', front: '文字在前', behind: '文字在後' };
const ALIGN_LABEL = { left: '靠左對齊', center: '水平置中', right: '靠右對齊', top: '靠上對齊', middle: '垂直置中', bottom: '靠下對齊' };
const H_LABEL: Record<string, string> = { column: '欄', margin: '邊界', page: '頁面', character: '字元', leftMargin: '左邊界', rightMargin: '右邊界', insideMargin: '內側邊界', outsideMargin: '外側邊界' };
const V_LABEL: Record<string, string> = { paragraph: '段落', line: '行', margin: '邊界', page: '頁面', topMargin: '上邊界', bottomMargin: '下邊界', insideMargin: '內側邊界', outsideMargin: '外側邊界' };

function shift(r: Rect, dx: number, dy: number): Rect {
  return { left: r.left + dx, right: r.right + dx, top: r.top + dy, bottom: r.bottom + dy };
}

/** A box resized from one of its handles by (dx, dy); with `keep`, in proportion. */
export function resizeBox(r: Rect, handle: Handle, dx: number, dy: number, keep: boolean): Rect {
  let { left, top, right, bottom } = r;
  if (handle.includes('w')) left = Math.min(right - 1, left + dx);
  if (handle.includes('e')) right = Math.max(left + 1, right + dx);
  if (handle.includes('n')) top = Math.min(bottom - 1, top + dy);
  if (handle.includes('s')) bottom = Math.max(top + 1, bottom + dy);
  if (keep && w(r) > 0 && h(r) > 0) {
    const ratio = w(r) / h(r);
    const corner = handle.length === 2;
    if (corner || handle === 'e' || handle === 'w') {
      const nh = (right - left) / ratio;
      if (handle.includes('n')) top = bottom - nh;
      else bottom = top + nh;
    } else {
      const nw = (bottom - top) * ratio;
      right = left + nw;
    }
  }
  return { left, top, right, bottom };
}

function describe(m: ShapeModel, p: Placed | null): string {
  const what = m.kind === 'textbox' ? tl('文字方塊') : m.kind === 'canvas' ? tl('繪圖畫布') : m.kind === 'group' ? tl('群組') : tl('圖案');
  const name = m.title || m.name;
  const wrap = tl(WRAP_LABEL[wrapOf(m)]);
  const size = p ? tl('，寬 {0} 公分、高 {1} 公分', cm(w(p.box)), cm(h(p.box))) : '';
  return tl('{0}{1}（{2}{3}）', what, name ? tl('「{0}」', name) : '', wrap, size);
}

function position(p: Placed, dx: number, dy: number): string {
  const at = offsetsNow(p);
  if (!at) return '';
  return tl('水平 {0} 公分，垂直 {1} 公分。', cm(at.h / EMU_PX + dx), cm(at.v / EMU_PX + dy));
}

/** The text of the first text box in a run's XML (for the VML copy of a group). */
function innerText(xml: string): string {
  const m = /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(xml);
  return m ? m[1] : '<w:p/>';
}

export type { WrapChoice };
