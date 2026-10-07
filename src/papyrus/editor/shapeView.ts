// Drawing text boxes and shapes (docx/shapes.ts) at Word's position and size.
//
// A shape is a kept-as-is raw_inline node with a shape model; schema.ts asks shapeDOM for its
// DOM (so the body, the static header/footer copies and the printout all draw it the same way):
//  - the outline, fill and lines as SVG (docx/shapeGeometry.ts), connectors with arrowheads;
//  - each text box's text as the editor draws paragraphs (the document's styles apply);
//  - in the line of text (wp:inline) it takes its size in the line; anchored (wp:anchor), a
//    zero-width mark stays where it is anchored and the drawing is placed from the page, the
//    margins, the column or the paragraph by placeShapes, after layout. Text wraps around a
//    shape set to wrap (square, tight, through, top and bottom) that is placed from its paragraph.
// A shape that can't be drawn keeps Word's preview picture (or its label) and says so.

import { DOMSerializer, Node as PMNode } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema, setShapeParser, setShapeRenderer } from './schema';
import { NS, parseFragment } from '../docx/xml';
import { readRunShape } from '../docx/shapeOps';
import { followLink } from './linkFollow';
import { customPaths, isLinePreset, presetPaths, presetTextRect } from '../docx/shapeGeometry';
import { shapeLabel, type Anchor, type AxisPos, type ShapeModel, type ShapeNode, type SpNode, type TextFrame } from '../docx/shapes';
import { repaginate } from './pagination';
import { ListCounter, levelOf } from '../docx/numbering';
import type { Numbering } from '../docx/model';
import { twipsToPx } from '../units';
import { tl } from '../i18n';
import './shapes.css';

const SVG = 'http://www.w3.org/2000/svg';
/** EMU per CSS pixel. */
const EMU_PX = 9525;
const px = (emu: number) => emu / EMU_PX;
const r2 = (n: number) => Math.round(n * 100) / 100;

/** Said where a shape can't be drawn or edited here. */
export const WORD_ONLY = '這個圖形只能在 Word 修改';

/** The editor document of a text box, made once per text (versions of the document share it). */
const frameDocs = new WeakMap<object, PMNode>();
export function frameDoc(f: TextFrame): PMNode {
  let d = frameDocs.get(f.doc);
  if (!d) {
    try {
      d = PMNode.fromJSON(schema, f.doc);
    } catch {
      d = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
    }
    frameDocs.set(f.doc, d);
  }
  return d;
}

/** A shape of the drawing with its box in px, in the drawing's coordinates. */
export interface PlacedNode {
  node: Exclude<ShapeNode, { t: 'grp' }>;
  /** Which child of the drawing's top group it is (or is in); absent for a lone shape. */
  kid?: number;
  x: number;
  y: number;
  w: number;
  h: number;
}

/** The drawing's shapes in drawing order, their boxes worked out through their groups. */
export function placedNodes(m: ShapeModel): PlacedNode[] {
  const out: PlacedNode[] = [];
  // Child EMU → drawing EMU: x' = sx * x + tx.
  const visit = (n: ShapeNode, sx: number, sy: number, tx: number, ty: number, kid?: number) => {
    const x = sx * n.x + tx;
    const y = sy * n.y + ty;
    const w = sx * n.w;
    const h = sy * n.h;
    if (n.t === 'grp') {
      const cx = n.ch.w ? w / n.ch.w : 1;
      const cy = n.ch.h ? h / n.ch.h : 1;
      n.kids.forEach((k, i) => visit(k, cx, cy, x - cx * n.ch.x, y - cy * n.ch.y, kid ?? (n === m.root ? i : undefined)));
      return;
    }
    const placed: PlacedNode = { node: n, x: px(x), y: px(y), w: px(w), h: px(h) };
    if (kid != null) placed.kid = kid;
    out.push(placed);
  };
  visit(m.root, 1, 1, 0, 0);
  return out;
}

function svgEl(tag: string, attrs: Record<string, string | number>): SVGElement {
  const e = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
}

/** stroke-dasharray of a:prstDash (and VML dashstyle), in line widths. */
const DASHES: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], sysDot: [1, 1], dot: [1, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
  shortdash: [3, 1], shortdot: [1, 1], shortdashdot: [3, 1, 1, 1], longdash: [8, 3], dashdot: [4, 3, 1, 3], longdashdot: [8, 3, 1, 3],
};

const END_SIZE: Record<string, number> = { sm: 2.5, med: 4, lg: 6 };

/** An arrowhead marker (in line widths, pointing along the line). */
function marker(id: string, type: string, w: string | undefined, len: string | undefined, color: string): SVGElement {
  const mw = END_SIZE[w ?? 'med'] ?? 3;
  const ml = END_SIZE[len ?? 'med'] ?? 3;
  const m = svgEl('marker', { id, viewBox: '0 0 10 10', refX: type === 'oval' || type === 'diamond' ? 5 : 9, refY: 5, markerWidth: ml, markerHeight: mw, orient: 'auto-start-reverse', markerUnits: 'strokeWidth' });
  const d = type === 'oval' ? 'M0,5A5,5 0 1 0 10,5A5,5 0 1 0 0,5Z' : type === 'diamond' ? 'M0,5L5,0L10,5L5,10Z' : type === 'arrow' ? 'M0,0L10,5L0,10' : type === 'stealth' ? 'M0,0L10,5L0,10L3,5Z' : 'M0,0L10,5L0,10Z';
  const open = type === 'arrow';
  m.append(svgEl('path', open ? { d, fill: 'none', stroke: color, 'stroke-width': 1.5 } : { d, fill: color }));
  return m;
}

let markerIds = 0;

/** The SVG of one shape at (x, y) in the drawing. */
function drawSp(p: PlacedNode, n: SpNode, defs: SVGElement): SVGElement | null {
  if (n.art) return drawArt(p, n);
  const paths = n.geom === 'custom' ? (n.paths ? customPaths(n.paths, p.w, p.h) : null) : presetPaths(n.geom, p.w, p.h, n.av);
  if (!paths) return null;
  const g = svgEl('g', { transform: `translate(${r2(p.x)},${r2(p.y)})` });
  const inner = svgEl('g', {});
  const t: string[] = [];
  if (n.rot) t.push(`rotate(${r2(n.rot)} ${r2(p.w / 2)} ${r2(p.h / 2)})`);
  if (n.flipH || n.flipV) t.push(`translate(${n.flipH ? r2(p.w) : 0},${n.flipV ? r2(p.h) : 0}) scale(${n.flipH ? -1 : 1},${n.flipV ? -1 : 1})`);
  if (t.length) inner.setAttribute('transform', t.join(' '));
  const line = n.line;
  const width = line ? Math.max(0.75, px(line.width)) : 0;
  const color = line ? '#' + line.color : 'none';
  let head = '';
  let tail = '';
  if (line?.head) {
    const id = `dx-mk${++markerIds}`;
    defs.append(marker(id, line.head.type, line.head.w, line.head.len, color));
    head = `url(#${id})`;
  }
  if (line?.tail) {
    const id = `dx-mk${++markerIds}`;
    defs.append(marker(id, line.tail.type, line.tail.w, line.tail.len, color));
    tail = `url(#${id})`;
  }
  for (const path of paths) {
    const a: Record<string, string | number> = { d: path.d, 'stroke-linejoin': 'miter', 'stroke-miterlimit': 8 };
    a.fill = path.fill && n.fill ? '#' + n.fill.color : 'none';
    if (path.fill && n.fill?.alpha != null) a['fill-opacity'] = r2(n.fill.alpha);
    if (path.stroke && line) {
      a.stroke = color;
      a['stroke-width'] = r2(width);
      if (line.alpha != null) a['stroke-opacity'] = r2(line.alpha);
      const dash = line.dash ? DASHES[line.dash] : null;
      if (dash) a['stroke-dasharray'] = dash.map((d) => r2(d * width)).join(' ');
      // Arrowheads on open lines (connectors, lines, arcs).
      if (!path.fill || isLinePreset(n.geom)) {
        if (head) a['marker-start'] = head;
        if (tail) a['marker-end'] = tail;
      }
    } else a.stroke = 'none';
    inner.append(svgEl('path', a));
  }
  g.append(inner);
  return g;
}

/** WordArt (a watermark): its text stretched over the box, in the fill colour (faint, as Word shows it). */
function drawArt(p: PlacedNode, n: SpNode): SVGElement {
  const g = svgEl('g', { transform: `translate(${r2(p.x)},${r2(p.y)})${n.rot ? ` rotate(${r2(n.rot)} ${r2(p.w / 2)} ${r2(p.h / 2)})` : ''}` });
  const a: Record<string, string | number> = {
    x: r2(p.w / 2), y: r2(p.h / 2), 'text-anchor': 'middle', 'dominant-baseline': 'central', 'font-size': r2(p.h * 0.9),
    textLength: r2(p.w), lengthAdjust: 'spacingAndGlyphs', fill: n.fill ? '#' + n.fill.color : '#c0c0c0', 'fill-opacity': r2(n.fill?.alpha ?? 1),
  };
  if (n.art!.font) a['font-family'] = n.art!.font;
  const t = svgEl('text', a);
  t.textContent = n.art!.text;
  g.append(t);
  return g;
}

/** A text box's text, drawn over its shape (turned with it, never flipped). */
function drawText(p: PlacedNode, n: SpNode, f: TextFrame): HTMLElement {
  const holder = document.createElement('div');
  holder.className = 'dx-shape-textbox';
  if (p.kid != null) holder.dataset.kid = String(p.kid);
  holder.style.cssText = `left:${r2(p.x)}px;top:${r2(p.y)}px;width:${r2(p.w)}px;height:${r2(p.h)}px` + (n.rot ? `;transform:rotate(${r2(n.rot)}deg)` : '');
  const tr = presetTextRect(n.geom, p.w, p.h, n.av);
  const [l, t, r, b] = f.inset.map(px);
  const frame = document.createElement('div');
  // .dx-doc: the document's default paragraph and run formatting, not the anchor paragraph's.
  frame.className = 'dx-doc dx-shape-text';
  frame.dataset.text = String(f.i);
  frame.style.cssText =
    `left:${r2(tr.l + l)}px;top:${r2(tr.t + t)}px;width:${r2(Math.max(0, tr.r - tr.l - l - r))}px;height:${r2(Math.max(0, tr.b - tr.t - t - b))}px;` +
    `justify-content:${f.anchor === 'ctr' ? 'center' : f.anchor === 'b' ? 'flex-end' : 'flex-start'}` +
    (f.vert ? `;writing-mode:${f.vert === 'vert270' ? 'vertical-lr' : 'vertical-rl'}` : '');
  const content = document.createElement('div');
  content.className = 'dx-shape-content';
  content.append(DOMSerializer.fromSchema(schema).serializeFragment(frameDoc(f).content));
  frame.append(content);
  holder.append(frame);
  return holder;
}

/** The drawing: its SVG and its text boxes' text, `w` × `h` px. */
function drawing(m: ShapeModel): HTMLElement {
  const box = document.createElement('span');
  box.className = 'dx-shape-box';
  const w = px(m.w);
  const h = px(m.h);
  box.style.width = `${r2(w)}px`;
  box.style.height = `${r2(h)}px`;
  const svg = svgEl('svg', { width: r2(Math.max(w, 1)), height: r2(Math.max(h, 1)), viewBox: `0 0 ${r2(Math.max(w, 1))} ${r2(Math.max(h, 1))}`, 'aria-hidden': 'true', focusable: 'false' });
  svg.setAttribute('class', 'dx-shape-svg');
  const defs = svgEl('defs', {});
  svg.append(defs);
  if (m.root.t === 'grp' && m.root.bg) svg.append(svgEl('rect', { x: 0, y: 0, width: r2(w), height: r2(h), fill: '#' + m.root.bg.color }));
  const texts: HTMLElement[] = [];
  for (const p of placedNodes(m)) {
    if (p.node.t === 'pic') {
      if (p.node.src) svg.append(svgEl('image', { href: p.node.src, x: r2(p.x), y: r2(p.y), width: r2(p.w), height: r2(p.h), preserveAspectRatio: 'none' }));
      continue;
    }
    const g = drawSp(p, p.node, defs);
    if (g && p.kid != null) g.setAttribute('data-kid', String(p.kid));
    if (g) svg.append(g);
    if (p.node.text) texts.push(drawText(p, p.node, p.node.text));
  }
  box.append(svg, ...texts);
  return box;
}

/** What placeShapes needs of an anchored shape, kept on its element (static copies of headers have no model). */
interface AnchorData {
  h: AxisPos;
  v: AxisPos;
  w: number;
  hgt: number;
  wrap: Anchor['wrap'];
  side?: string;
  /** Distances from the text in px [top, bottom, left, right]. */
  dist: number[];
  behind: boolean;
}

/** The DOM of a raw_inline holding a shape model (see schema.ts). */
export function shapeDOM(node: PMNode): HTMLElement | null {
  const m = node.attrs.shape as ShapeModel | null;
  if (!m?.root) return null;
  const label = shapeLabel(m);
  const outer = document.createElement('span');
  outer.dataset.shapeKey = m.key;
  outer.contentEditable = 'false';
  let box: HTMLElement;
  if (!m.drawable) {
    // Where Word has it and as big, with Word's preview picture when there is one; Word only.
    outer.title = tl(WORD_ONLY);
    outer.setAttribute('role', 'img');
    outer.setAttribute('aria-label', tl('{0}（{1}）', label, tl(WORD_ONLY)));
    box = document.createElement('span');
    box.className = 'dx-shape-box dx-shape-unknown';
    box.style.width = `${r2(px(m.w))}px`;
    box.style.height = `${r2(px(m.h))}px`;
    const { src } = node.attrs;
    if (src) {
      const img = document.createElement('img');
      img.src = src;
      img.alt = '';
      img.className = 'dx-shape-preview';
      box.append(img);
    }
    const note = document.createElement('span');
    note.className = 'dx-shape-note';
    note.textContent = tl(WORD_ONLY);
    box.append(note);
  } else {
    const hasText = placedNodes(m).some((p) => p.node.t === 'sp' && p.node.text);
    if (hasText) {
      outer.setAttribute('role', 'group');
      outer.setAttribute('aria-label', label);
      outer.title = tl('按兩下可編輯文字');
    } else if (m.title || m.descr) {
      outer.setAttribute('role', 'img');
      outer.setAttribute('aria-label', label);
    } else outer.setAttribute('aria-hidden', 'true');
    box = drawing(m);
  }
  if (!m.anchor) {
    outer.className = 'dx-shape dx-shape-inline' + (m.drawable ? '' : ' dx-shape-fixed');
    const [l, t, r, b] = (m.effect ?? [0, 0, 0, 0]).map(px);
    outer.style.cssText = `width:${r2(px(m.w) + l + r)}px;height:${r2(px(m.h) + t + b)}px`;
    box.style.left = `${r2(l)}px`;
    box.style.top = `${r2(t)}px`;
    outer.append(box);
    return outer;
  }
  const a = m.anchor;
  outer.className = 'dx-shape dx-shape-float' + (a.behind ? ' dx-shape-behind' : '') + (m.drawable ? '' : ' dx-shape-fixed');
  // Word's stacking order (relativeHeight); behind the text below it (see shapes.css).
  const z = Math.min(Math.max(0, Math.round(a.z)), 1e9);
  box.style.zIndex = String(a.behind ? z - 1e9 - 1 : z + 1);
  const data: AnchorData = { h: a.h, v: a.v, w: px(m.w), hgt: px(m.h), wrap: a.wrap, dist: a.dist.map(px), behind: a.behind };
  if (a.side) data.side = a.side;
  outer.dataset.anchor = JSON.stringify(data);
  // Where it is anchored, to measure from. Empty: a wrap float after it then starts the line
  // (with any text before it on the line, the text would stay beside a full-width float).
  const at = document.createElement('span');
  at.className = 'dx-shape-at';
  outer.append(at);
  if (wraps(data)) {
    const wrap = document.createElement('span');
    wrap.className = 'dx-shape-wrap';
    wrap.setAttribute('aria-hidden', 'true');
    outer.append(wrap);
  }
  outer.append(box);
  return outer;
}

/** Text wraps around it here: set to wrap and placed from its paragraph (not from the page). */
function wraps(a: AnchorData): boolean {
  return a.wrap !== 'none' && !a.behind && (a.v.rel === 'paragraph' || a.v.rel === 'line');
}

// ----- placing anchored shapes -----

export interface Rect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** What an anchored shape is placed from, in one coordinate space (px). */
export interface PlaceFrame {
  /** The page, and its margins [top, right, bottom, left]; null when there is no page (e.g. in a test). */
  page: Rect | null;
  margins: [number, number, number, number];
  /** The text column (the page's text area, or the table cell). */
  column: Rect;
  paragraph: Rect;
  /** The anchor mark (its line, and the character). */
  mark: Rect;
}

/** Where an axis starts and ends for a `relativeFrom`. */
function area(rel: string, f: PlaceFrame, horizontal: boolean): [number, number] {
  const p = f.page;
  const [mt, mr, mb, ml] = f.margins;
  if (horizontal) {
    switch (rel) {
      case 'page':
        return p ? [p.left, p.right] : [f.column.left, f.column.right];
      case 'margin':
        return p ? [p.left + ml, p.right - mr] : [f.column.left, f.column.right];
      case 'character':
        return [f.mark.left, f.mark.left];
      case 'leftMargin':
      case 'insideMargin':
        return p ? [p.left, p.left + ml] : [f.column.left, f.column.left];
      case 'rightMargin':
      case 'outsideMargin':
        return p ? [p.right - mr, p.right] : [f.column.right, f.column.right];
      default:
        return [f.column.left, f.column.right];
    }
  }
  switch (rel) {
    case 'page':
      return p ? [p.top, p.bottom] : [f.paragraph.top, f.paragraph.bottom];
    case 'margin':
    case 'insideMargin':
    case 'outsideMargin':
      return p ? [p.top + mt, p.bottom - mb] : [f.paragraph.top, f.paragraph.bottom];
    case 'topMargin':
      return p ? [p.top, p.top + mt] : [f.paragraph.top, f.paragraph.top];
    case 'bottomMargin':
      return p ? [p.bottom - mb, p.bottom] : [f.paragraph.bottom, f.paragraph.bottom];
    case 'line':
      return [f.mark.top, f.mark.bottom];
    default:
      return [f.paragraph.top, f.paragraph.bottom];
  }
}

function along(axis: AxisPos, [a0, a1]: [number, number], size: number): number {
  if (axis.offset != null) return a0 + px(axis.offset);
  if (axis.pct != null) return a0 + ((a1 - a0) * axis.pct) / 100000;
  switch (axis.align) {
    case 'right':
    case 'bottom':
    case 'outside':
      return a1 - size;
    case 'center':
      return (a0 + a1) / 2 - size / 2;
    default:
      return a0;
  }
}

/** The top left corner of an anchored shape (w × h px) in the frame's coordinates, as Word places it. */
export function anchorPosition(h: AxisPos, v: AxisPos, w: number, hgt: number, f: PlaceFrame): { x: number; y: number } {
  return { x: along(h, area(h.rel, f, true), w), y: along(v, area(v.rel, f, false), hgt) };
}

/**
 * The wrap exclusion (a float the text flows around) for a shape at (x, y) px in the frame,
 * relative to its column and paragraph: which side, its margins and size, and the part above
 * the shape that text may still use (shape-outside). Null when there is nothing to keep clear.
 */
export function wrapBox(a: AnchorData, x: number, y: number, f: PlaceFrame): { side: 'left' | 'right'; css: string } | null {
  const [dt, db, dl, dr] = a.dist;
  const colW = f.column.right - f.column.left;
  const ox = x - f.column.left;
  const oy = y - f.paragraph.top;
  const height = oy + a.hgt + db;
  if (height <= 0) return null;
  const inset = `shape-outside:inset(${r2(Math.max(0, oy - dt))}px 0 0 0)`;
  if (a.wrap === 'topAndBottom') return { side: 'left', css: `display:block;float:left;width:100%;height:${r2(height)}px;margin:0;${inset}` };
  let side: 'left' | 'right' = ox + a.w / 2 <= colW / 2 ? 'left' : 'right';
  if (a.side === 'left') side = 'right'; // text only on the left: the shape keeps the right
  else if (a.side === 'right') side = 'left';
  else if (a.side === 'largest') side = ox >= colW - ox - a.w ? 'right' : 'left';
  const width = a.w + dl + dr;
  const margin = side === 'left' ? `margin:0 0 0 ${r2(ox - dl)}px` : `margin:0 ${r2(colW - ox - a.w - dr)}px 0 0`;
  return { side, css: `display:block;float:${side};width:${r2(width)}px;height:${r2(height)}px;${margin};${inset}` };
}

function rectOf(el: Element): Rect {
  const r = el.getBoundingClientRect();
  return { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
}

/** The content box of an element (without its padding and border), client coordinates. */
function contentBox(el: HTMLElement, zoom: number): Rect {
  const r = el.getBoundingClientRect();
  const cs = getComputedStyle(el);
  const n = (v: string) => (parseFloat(v) || 0) * zoom;
  return {
    left: r.left + n(cs.borderLeftWidth) + n(cs.paddingLeft),
    right: r.right - n(cs.borderRightWidth) - n(cs.paddingRight),
    top: r.top + n(cs.borderTopWidth) + n(cs.paddingTop),
    bottom: r.bottom - n(cs.borderBottomWidth) - n(cs.paddingBottom),
  };
}

/** The pages of a canvas, top to bottom, measured once per placing (see placeShapes). */
type PageList = { el: HTMLElement; top: number; bottom: number }[];

function pageList(canvas: Element | null, cache: Map<Element, PageList>): PageList {
  if (!canvas) return [];
  let list = cache.get(canvas);
  if (!list) {
    list = Array.from(canvas.querySelectorAll<HTMLElement>(':scope > .dx-pages > .dx-page')).map((el) => {
      const r = el.getBoundingClientRect();
      return { el, top: r.top, bottom: r.bottom };
    });
    cache.set(canvas, list);
  }
  return list;
}

/** The page an element is on: the page element holding it (a header copy, the header editor), else the one level with `y` (client). */
export function pageOf(el: HTMLElement, y: number, cache: Map<Element, PageList> = new Map()): HTMLElement | null {
  const own = el.closest<HTMLElement>('.dx-page, .dx-hf-box');
  if (own) return own;
  const pages = pageList(el.closest('.dx-canvas'), cache);
  if (!pages.length) return null;
  // The last page starting above y (pages are drawn top to bottom); the gap below a page counts to it.
  let lo = 0;
  let hi = pages.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (pages[mid].top <= y) lo = mid;
    else hi = mid - 1;
  }
  return pages[lo].el;
}

/**
 * What a shape (or anything in a paragraph) is placed from: its page, the margins, the column
 * (the table cell, else the page's text area), its paragraph and its own mark, in the page's own
 * px (client coordinates divided by the zoom); and the zoom.
 */
export function measureFrame(el: HTMLElement, pages: Map<Element, PageList> = new Map()): { frame: PlaceFrame; zoom: number } {
  const canvas = el.closest<HTMLElement>('.dx-canvas');
  const zoom = canvas && canvas.offsetWidth ? canvas.getBoundingClientRect().width / canvas.offsetWidth || 1 : 1;
  const unz = (r: Rect): Rect => ({ left: r.left / zoom, right: r.right / zoom, top: r.top / zoom, bottom: r.bottom / zoom });
  const para = (el.classList.contains('dx-p') ? el : el.parentElement?.closest<HTMLElement>('.dx-p, p') ?? el.parentElement ?? el)!;
  const mark = unz(rectOf(el.querySelector?.(':scope > .dx-shape-at') ?? el));
  const paragraph = unz(rectOf(para));
  const page = pageOf(el, paragraph.top * zoom, pages);
  const margins = (page?.dataset.margins ?? '').split(',').map(Number).filter(Number.isFinite);
  const m: [number, number, number, number] = margins.length === 4 ? (margins as [number, number, number, number]) : [0, 0, 0, 0];
  const pageRect = page ? unz(rectOf(page)) : null;
  // The column: the table cell holding the paragraph, else the page's text area (else the container).
  const inText = el.parentElement?.closest('.dx-shape-text') ?? null;
  const cell = para.closest<HTMLElement>('td, th');
  const container = para.parentElement as HTMLElement | null;
  let column: Rect;
  if (cell && (cell.closest('.dx-shape-text') ?? null) === inText) column = unz(contentBox(cell, zoom));
  else if (pageRect && margins.length === 4 && !inText) column = { left: pageRect.left + m[3], right: pageRect.right - m[1], top: pageRect.top + m[0], bottom: pageRect.bottom - m[2] };
  else column = container ? unz(contentBox(container, zoom)) : paragraph;
  return { frame: { page: pageRect, margins: m, column, paragraph, mark }, zoom };
}

/** Where an axis's reference area starts (see anchorPosition), in the frame's coordinates. */
export function areaStart(rel: string, f: PlaceFrame, horizontal: boolean): number {
  return area(rel, f, horizontal)[0];
}

/**
 * The list numbers and bullets of drawn text boxes (the editor's lists get theirs as decorations,
 * see listMarkers.ts, which a drawn text box doesn't have), each text box counting on its own.
 */
export function textBoxListMarkers(scope: ParentNode, numbering: Numbering): void {
  for (const frame of Array.from(scope.querySelectorAll<HTMLElement>('.dx-shape-text:not([data-marked])'))) {
    frame.dataset.marked = '';
    const counter = new ListCounter(numbering);
    for (const p of Array.from(frame.querySelectorAll<HTMLElement>(':scope > .dx-shape-content > p[data-num-id]'))) {
      const numId = p.dataset.numId!;
      const ilvl = Number(p.dataset.ilvl ?? 0) || 0;
      const text = counter.next(numId, ilvl);
      const lvl = levelOf(numbering, numId, ilvl);
      if (text == null || !lvl) continue;
      // The level's indentation, unless the paragraph sets its own.
      const ownLeft = !!p.style.marginLeft;
      if (!ownLeft && lvl.indLeft != null) p.style.marginLeft = `${twipsToPx(lvl.indLeft)}px`;
      if (!p.style.textIndent && lvl.hanging != null) p.style.textIndent = `${-twipsToPx(lvl.hanging)}px`;
      const hanging = lvl.hanging ?? 0;
      const marker = document.createElement('span');
      marker.className = 'dx-marker';
      marker.textContent = text;
      marker.style.minWidth = `${hanging > 0 ? twipsToPx(hanging) : ownLeft || lvl.indLeft ? 0 : 24}px`;
      p.prepend(marker);
    }
  }
}

/**
 * Place the anchored shapes under `scope` (after layout: the body, a header or footer copy, the
 * header editor) and size their wrap exclusions; with `numbering`, their lists get their numbers.
 * Returns whether an exclusion changed (the text then flows differently: the pages need measuring again).
 */
export function placeShapes(scope: ParentNode, numbering?: Numbering | null): boolean {
  if (numbering) textBoxListMarkers(scope, numbering);
  let changed = false;
  const pages = new Map<Element, PageList>();
  for (const el of Array.from(scope.querySelectorAll<HTMLElement>('.dx-shape-float[data-anchor]'))) {
    let a: AnchorData;
    try {
      a = JSON.parse(el.dataset.anchor!);
    } catch {
      continue;
    }
    const box = el.querySelector<HTMLElement>(':scope > .dx-shape-box');
    if (!box || !el.isConnected) continue;
    const { frame, zoom } = measureFrame(el, pages);
    const unz = (r: Rect): Rect => ({ left: r.left / zoom, right: r.right / zoom, top: r.top / zoom, bottom: r.bottom / zoom });
    const pos = anchorPosition(a.h, a.v, a.w, a.hgt, frame);
    const origin = unz(rectOf(el));
    const left = `${r2(pos.x - origin.left)}px`;
    const top = `${r2(pos.y - origin.top)}px`;
    if (box.style.left !== left) box.style.left = left;
    if (box.style.top !== top) box.style.top = top;
    // Where its references start, from the mark (so a move can be turned into offsets).
    const ref = `${r2(area(a.h.rel, frame, true)[0] - origin.left)},${r2(area(a.v.rel, frame, false)[0] - origin.top)}`;
    if (el.dataset.ref !== ref) el.dataset.ref = ref;
    const wrapEl = el.querySelector<HTMLElement>(':scope > .dx-shape-wrap');
    if (wrapEl) {
      const wb = wrapBox(a, pos.x, pos.y, frame);
      const css = wb?.css ?? 'display:none';
      if (wrapEl.dataset.css !== css) {
        wrapEl.dataset.css = css;
        wrapEl.style.cssText = css;
        changed = true;
      }
    }
  }
  return changed;
}

/**
 * Keeps an editor's anchored shapes placed as its content changes (the pages being drawn again
 * places them too, see DocxEditor.renderPages). A changed wrap exclusion measures the pages again.
 */
export function shapeLayout(numbering: () => Numbering | null = () => null): Plugin {
  return new Plugin({
    view(view: EditorView) {
      let frame: number | null = null;
      const place = () => {
        frame = null;
        if (view.isDestroyed) return;
        if (placeShapes(view.dom, numbering())) repaginate(view);
      };
      const schedule = () => {
        if (frame != null || typeof requestAnimationFrame !== 'function') return;
        frame = requestAnimationFrame(place);
      };
      schedule();
      return {
        update: (_v, prev) => {
          if (prev.doc !== view.state.doc) schedule();
        },
        destroy: () => {
          if (frame != null) cancelAnimationFrame(frame);
        },
      };
    },
  });
}

setShapeRenderer(shapeDOM);
setShapeParser(pastedShape);

/**
 * A shape pasted from the editor's clipboard (its run's XML): only a run holding a drawing or VML
 * shape and nothing that refers to the package (pictures, links: they would point at nothing
 * in another document), not too big. Null otherwise (it is left out of the paste).
 */
export function pastedShape(xml: string | null): Record<string, unknown> | null {
  if (!xml || xml.length > 2_000_000) return null;
  let run: Element;
  try {
    run = parseFragment(xml);
  } catch {
    return null;
  }
  if (!run || run.namespaceURI !== NS.w || run.localName !== 'r') return null;
  if (!Array.from(run.children).every((c) => (c.namespaceURI === NS.w && ['rPr', 'drawing', 'pict'].includes(c.localName)) || (c.namespaceURI === NS.mc && c.localName === 'AlternateContent'))) return null;
  for (const e of Array.from(run.getElementsByTagName('*'))) {
    if (Array.from(e.attributes).some((a) => a.namespaceURI === NS.r || a.name.startsWith('r:'))) return null;
  }
  const shape = readRunShape(xml);
  return shape ? { xml, label: '圖形', shape } : null;
}

/**
 * A link in a drawn text box (not being edited) opens only after asking, like one in the text
 * (a drawn text box is not editable, so the browser would follow it on a single click).
 */
function onLinkClick(e: MouseEvent): void {
  const a = (e.target as Element | null)?.closest?.<HTMLAnchorElement>('.dx-shape-text a[href]');
  if (!a) return;
  e.preventDefault();
  e.stopPropagation();
  // The same rules and question as links in the text (linkFollow.ts); a place in the document is not gone to from here.
  followLink(a.getAttribute('href') ?? '', { confirm: (m) => window.confirm(m), goToBookmark: () => false });
}
if (typeof document !== 'undefined') document.addEventListener('click', onLinkClick, true);
