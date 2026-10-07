// Changing shapes (phase 8B): moving, resizing, recolouring, wrapping, stacking, connecting,
// grouping; and making new ones. Pure functions on the shape model (docx/shapes.ts): the editor
// puts the result on the shape's raw_inline, and the save writes what changed (docx/shapeWrite.ts).
// New shapes and groups get their XML made at once and their model read back from it, so the
// model always says what the XML says.

import { Converter } from './convert';
import { schema } from '../editor/schema';
import { NS, child, parseFragment, serializeXml } from './xml';
import { DEFAULT_THEME, type ThemeColors } from './theme';
import { connectionSites, isLinePreset, routeConnector } from './shapeGeometry';
import { changedModel, drawingXml, newShapeRun, vmlXml } from './shapeWrite';
import {
  SHAPE_NS, readShape, shapeSource, textFrames, type Anchor, type Box, type Connection, type Emu, type GrpNode, type Paint, type ShapeModel,
  type ShapeNode, type SpNode, type Stroke, type TextFrame,
} from './shapes';

/** EMU per cm and per CSS px. */
export const EMU_CM = 360000;
export const EMU_PX = 9525;

// ----- reading a run made here -----

/** The model of a run's shape (made here or pasted), read as a file's would be. */
export function readRunShape(runXml: string, theme: ThemeColors = DEFAULT_THEME): ShapeModel | null {
  let run: Element;
  try {
    run = parseFragment(runXml);
  } catch {
    return null;
  }
  const holder = Array.from(run.children).find((c) => shapeSource(c));
  if (!holder) return null;
  const conv = () => new Converter(new Map(), new Map(), theme);
  return readShape(holder, {
    theme,
    image: () => undefined,
    blocks: (el) => conv().blocks(el),
    doc: (blocks) => schema.nodes.doc.create(null, blocks.length ? blocks : [schema.nodes.paragraph.create()]),
  });
}

// ----- new shapes (插入 › 圖案) -----

export interface ShapeKind {
  id: string;
  /** 插入 › 圖案's name for it. */
  label: string;
  prst: string;
  /** A line or connector: drawn from a start point to an end point. */
  line?: boolean;
  /** Arrowhead at the end of a line. */
  arrow?: boolean;
  textbox?: boolean;
  /** Default size, cm. */
  size: [number, number];
}

export const SHAPE_KINDS: ShapeKind[] = [
  { id: 'process', label: '流程圖：程序', prst: 'flowChartProcess', size: [3, 1.5] },
  { id: 'decision', label: '流程圖：決策', prst: 'flowChartDecision', size: [3.5, 2] },
  { id: 'terminator', label: '流程圖：開始或結束', prst: 'flowChartTerminator', size: [3, 1.2] },
  { id: 'data', label: '流程圖：資料', prst: 'flowChartInputOutput', size: [3, 1.5] },
  { id: 'document', label: '流程圖：文件', prst: 'flowChartDocument', size: [3, 1.8] },
  { id: 'predefined', label: '流程圖：預設程序', prst: 'flowChartPredefinedProcess', size: [3, 1.5] },
  { id: 'rect', label: '矩形', prst: 'rect', size: [3, 1.5] },
  { id: 'roundRect', label: '圓角矩形', prst: 'roundRect', size: [3, 1.5] },
  { id: 'ellipse', label: '橢圓', prst: 'ellipse', size: [3, 2] },
  { id: 'rightArrow', label: '向右箭號', prst: 'rightArrow', size: [2.5, 1.2] },
  { id: 'leftArrow', label: '向左箭號', prst: 'leftArrow', size: [2.5, 1.2] },
  { id: 'downArrow', label: '向下箭號', prst: 'downArrow', size: [1.2, 2.5] },
  { id: 'upArrow', label: '向上箭號', prst: 'upArrow', size: [1.2, 2.5] },
  { id: 'line', label: '直線', prst: 'line', line: true, size: [3, 0] },
  { id: 'arrowLine', label: '箭號（直線接點）', prst: 'straightConnector1', line: true, arrow: true, size: [3, 0] },
  { id: 'elbow', label: '肘形接點（箭號）', prst: 'bentConnector3', line: true, arrow: true, size: [3, 2] },
  { id: 'textbox', label: '文字方塊', prst: 'rect', textbox: true, size: [5, 2.5] },
];

const BLACK: Stroke = { color: '000000', width: 12700 };
const EMPTY_TEXT = '<w:p/>';
const CENTRED_TEXT = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>';

/** Ids of drawings (wp:docPr, wps:cNvPr) in XML, to give new ones others. */
export function drawingIds(xmls: Iterable<string>): Set<number> {
  const out = new Set<number>();
  for (const xml of xmls) for (const m of xml.matchAll(/<(?:wp:docPr|wps:cNvPr|pic:cNvPr|wpg:cNvPr)\b[^>]*?\sid="(\d+)"/g)) out.add(Number(m[1]));
  return out;
}

/** A new id not in `used` (and then in it): large, as Word gives them, so a picture added later can't take it. */
export function newDrawingId(used: Set<number>): number {
  let id: number;
  do id = 100000000 + Math.floor(Math.random() * 1900000000);
  while (used.has(id));
  used.add(id);
  return id;
}

export interface NewShape {
  kind: ShapeKind;
  /** Where, from the column (h) and the anchor paragraph (v), EMU; for a line its start. */
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  /** A line going up / left from its start. */
  flipH?: boolean;
  flipV?: boolean;
  /** Its stacking order (above the others). */
  z: number;
  docId: number;
  name: string;
  /** Connector ends attached to shapes. */
  st?: Connection;
  end?: Connection;
}

/** The run (XML) of a new shape as Word 2010+ writes it. */
export function newShapeXml(s: NewShape): string {
  const k = s.kind;
  const root: SpNode = { t: 'sp', x: 0, y: 0, w: s.w, h: s.h, geom: k.prst, fill: k.line ? null : { color: 'FFFFFF' }, line: { ...BLACK, width: k.textbox ? 9525 : 12700 } };
  if (s.flipH) root.flipH = true;
  if (s.flipV) root.flipV = true;
  if (k.arrow) root.line = { ...root.line!, tail: { type: 'triangle' } };
  if (k.line) {
    root.cxn = true;
    if (s.st) root.st = s.st;
    if (s.end) root.end = s.end;
  } else root.text = { i: 0, doc: {}, inset: [91440, 45720, 91440, 45720], anchor: k.textbox ? 't' : 'ctr' };
  const m: ShapeModel = {
    key: '',
    kind: k.textbox ? 'textbox' : 'shape',
    src: 'dml',
    w: s.w,
    h: s.h,
    anchor: { h: { rel: 'column', offset: s.x }, v: { rel: 'paragraph', offset: s.y }, wrap: 'none', dist: [0, 0, 114300, 114300], behind: false, z: s.z },
    root,
    drawable: true,
    docId: s.docId,
    name: s.name,
  };
  return newShapeRun(m, () => (k.textbox ? EMPTY_TEXT : CENTRED_TEXT));
}

// ----- changes -----

/** Placed at these offsets from its references (EMU); an alignment gives way to the offset. */
export function moved(m: ShapeModel, h: Emu, v: Emu): ShapeModel {
  if (!m.anchor) return m;
  const a: Anchor = { ...m.anchor, h: { rel: m.anchor.h.rel, offset: Math.round(h) }, v: { rel: m.anchor.v.rel, offset: Math.round(v) } };
  return changedModel(m, { ...m, anchor: a });
}

/** Resized to w × h (EMU): a group's children with it; a canvas grows around its shapes. */
export function resized(m: ShapeModel, w: Emu, h: Emu): ShapeModel {
  w = Math.max(0, Math.round(w));
  h = Math.max(0, Math.round(h));
  let root: ShapeNode = m.root;
  if (root.t === 'grp' && m.kind === 'canvas') root = { ...root, w, h, ch: { ...root.ch, w, h } };
  else if (root.t === 'grp') root = { ...root, w, h };
  else root = { ...root, w, h };
  return changedModel(m, { ...m, w, h, root });
}

const mapSps = (n: ShapeNode, f: (sp: SpNode) => SpNode): ShapeNode => (n.t === 'grp' ? { ...n, kids: n.kids.map((k) => mapSps(k, f)) } : n.t === 'sp' ? f(n) : n);

/** Filled with a colour (null: no fill); lines and connectors have none. */
export function filled(m: ShapeModel, color: string | null): ShapeModel {
  const root = mapSps(m.root, (sp) => (isLinePreset(sp.geom) ? sp : { ...sp, fill: color ? { color } : null }));
  return changedModel(m, { ...m, root });
}

/** Outlined in a colour and width (EMU) (colour null: no line). */
export function outlined(m: ShapeModel, color: string | null, width?: Emu): ShapeModel {
  const root = mapSps(m.root, (sp) => ({
    ...sp,
    line: color ? { ...(sp.line ?? { color, width: 12700 }), color, width: width ?? sp.line?.width ?? 12700 } : null,
  }));
  return changedModel(m, { ...m, root });
}

/** Word's 文繞圖 choices. */
export type WrapChoice = 'inline' | 'square' | 'topAndBottom' | 'front' | 'behind';

export function wrapOf(m: ShapeModel): WrapChoice {
  if (!m.anchor) return 'inline';
  if (m.anchor.behind) return 'behind';
  if (m.anchor.wrap === 'none') return 'front';
  return m.anchor.wrap === 'topAndBottom' ? 'topAndBottom' : 'square';
}

/**
 * Wrapped another way. Taken out of the line of text, it is placed at `at` (EMU from the column
 * and the anchor paragraph: where it was) above the other shapes (`z`).
 */
export function wrapped(m: ShapeModel, choice: WrapChoice, at: { h: Emu; v: Emu } = { h: 0, v: 0 }, z = 251659264): ShapeModel {
  if (choice === 'inline') return m.anchor ? changedModel(m, { ...m, anchor: null }) : m;
  const a: Anchor = m.anchor
    ? { ...m.anchor }
    : { h: { rel: 'column', offset: Math.round(at.h) }, v: { rel: 'paragraph', offset: Math.round(at.v) }, wrap: 'none', dist: [0, 0, 114300, 114300], behind: false, z };
  a.behind = choice === 'behind';
  a.wrap = choice === 'square' ? 'square' : choice === 'topAndBottom' ? 'topAndBottom' : 'none';
  if (choice === 'square') a.side = m.anchor?.side ?? 'bothSides';
  else delete a.side;
  if (choice === 'topAndBottom' || choice === 'square') a.dist = choice === 'topAndBottom' ? [0, 0, 0, 0] : [0, 0, 114300, 114300];
  return changedModel(m, { ...m, anchor: a });
}

/** Stacked at `z` (relativeHeight). */
export function stacked(m: ShapeModel, z: number): ShapeModel {
  return m.anchor ? changedModel(m, { ...m, anchor: { ...m.anchor, z } }) : m;
}

// ----- inside a group or canvas -----

/** The children of the drawing's top group (a canvas's shapes), with their index. */
export function kidsOf(m: ShapeModel): ShapeNode[] {
  return m.root.t === 'grp' ? m.root.kids : [];
}

/** A child of the top group changed (`f`), connectors attached to it rerouted. */
export function changedKid(m: ShapeModel, index: number, f: (n: ShapeNode) => ShapeNode): ShapeModel {
  if (m.root.t !== 'grp') return m;
  const kids = m.root.kids.map((k, i) => (i === index ? f(k) : k));
  const moved = kids[index];
  const root: GrpNode = { ...m.root, kids: moved.t === 'sp' && moved.id ? rerouteKids(kids, new Set([moved.id])) : kids };
  return changedModel(m, { ...m, root });
}

/** Where connection site `idx` of a node is, in its parent's coordinates. */
export function siteAt(n: SpNode, idx: number): { x: number; y: number; dir: number } | null {
  const sites = connectionSites(n.geom, n.w, n.h, n.av);
  const s = sites[idx];
  if (!s) return null;
  // Flipped and turned about the centre, as the shape is drawn.
  let x = n.flipH ? n.w - s.x : s.x;
  let y = n.flipV ? n.h - s.y : s.y;
  let dir = n.flipH ? 180 - s.dir : s.dir;
  if (n.flipV) dir = -dir;
  const t = ((n.rot ?? 0) * Math.PI) / 180;
  x -= n.w / 2;
  y -= n.h / 2;
  const rx = x * Math.cos(t) - y * Math.sin(t);
  const ry = x * Math.sin(t) + y * Math.cos(t);
  return { x: n.x + n.w / 2 + rx, y: n.y + n.h / 2 + ry, dir: (((dir + (n.rot ?? 0)) % 360) + 360) % 360 };
}

/** A connector's box going between two points, the direction it leaves the first by (see routeConnector). */
export function routed(c: SpNode, a: { x: number; y: number }, b: { x: number; y: number }, da?: number): SpNode {
  const box = routeConnector(c.geom, a, b, da);
  const next: SpNode = { ...c, x: Math.round(box.x), y: Math.round(box.y), w: Math.round(box.w), h: Math.round(box.h) };
  delete next.rot;
  delete next.flipH;
  delete next.flipV;
  if (box.rot) next.rot = box.rot;
  if (box.flipH) next.flipH = true;
  if (box.flipV) next.flipV = true;
  if (c.geom.startsWith('bentConnector')) next.av = box.av;
  return next;
}

/**
 * The connectors among `kids` attached to the shapes `ids` go again from site to site (the ends
 * not attached stay where they are).
 */
export function rerouteKids(kids: ShapeNode[], ids: Set<number>): ShapeNode[] {
  const byId = new Map<number, SpNode>();
  for (const k of kids) if (k.t === 'sp' && k.id != null) byId.set(k.id, k);
  return kids.map((k) => {
    if (k.t !== 'sp' || !k.cxn || !((k.st && ids.has(k.st.id)) || (k.end && ids.has(k.end.id)))) return k;
    const ends = connectorEndsOf(k);
    const sa = k.st && byId.get(k.st.id) ? siteAt(byId.get(k.st.id)!, k.st.idx) : null;
    const sb = k.end && byId.get(k.end.id) ? siteAt(byId.get(k.end.id)!, k.end.idx) : null;
    return routed(k, sa ?? ends.a, sb ?? ends.b, sa?.dir);
  });
}

/** A connector's two ends, in its parent's coordinates. */
export function connectorEndsOf(c: SpNode): { a: { x: number; y: number }; b: { x: number; y: number } } {
  const sites = connectionSites(c.geom, c.w, c.h);
  const at = (s: { x: number; y: number }) => {
    let x = c.flipH ? c.w - s.x : s.x;
    let y = c.flipV ? c.h - s.y : s.y;
    const t = ((c.rot ?? 0) * Math.PI) / 180;
    x -= c.w / 2;
    y -= c.h / 2;
    return { x: c.x + c.w / 2 + x * Math.cos(t) - y * Math.sin(t), y: c.y + c.h / 2 + x * Math.sin(t) + y * Math.cos(t) };
  };
  return { a: at(sites[0]), b: at(sites[1]) };
}

/** The connection site of one of `shapes` within `radius` of `p` (all in one coordinate space). */
export function nearestSite(shapes: SpNode[], p: { x: number; y: number }, radius: number): { node: SpNode; idx: number; x: number; y: number; dir: number } | null {
  let best: { node: SpNode; idx: number; x: number; y: number; dir: number } | null = null;
  let dist = radius;
  for (const n of shapes) {
    if (n.cxn || n.id == null || isLinePreset(n.geom)) continue;
    connectionSites(n.geom, n.w, n.h, n.av).forEach((_s, idx) => {
      const at = siteAt(n, idx);
      if (!at) return;
      const d = Math.hypot(at.x - p.x, at.y - p.y);
      if (d <= dist) {
        dist = d;
        best = { node: n, idx, ...at };
      }
    });
  }
  return best;
}

// ----- groups -----

export interface Placed {
  model: ShapeModel;
  /** Its box on the page, EMU, from a common origin (e.g. the column's left, the first paragraph's top). */
  x: Emu;
  y: Emu;
  /** Its run's XML as it would be saved now (see writer.ts shapeRun). */
  xml: string;
}

/** The wps:wsp of a single shape's run, its box set to `box` and given `id` (as a group's child). */
function wspOf(xml: string, box: Box, id: number, name: string): string {
  const run = parseFragment(xml);
  const holder = Array.from(run.children).find((c) => shapeSource(c));
  const src = holder ? shapeSource(holder) : null;
  const wsp = src?.kind === 'dml' ? src.primary.getElementsByTagNameNS(SHAPE_NS.wps, 'wsp')[0] : null;
  if (!wsp) throw new Error('only DrawingML shapes are grouped');
  const doc = wsp.ownerDocument;
  const old = child(wsp, 'cNvPr', SHAPE_NS.wps);
  if (old) wsp.removeChild(old);
  const nv = doc.createElementNS(SHAPE_NS.wps, 'wps:cNvPr');
  nv.setAttribute('id', String(id));
  nv.setAttribute('name', name);
  wsp.insertBefore(nv, wsp.firstChild);
  const spPr = child(wsp, 'spPr', SHAPE_NS.wps)!;
  const xfrm = child(spPr, 'xfrm', NS.a);
  const flip = xfrm ? `${xfrm.getAttribute('rot') ? ` rot="${xfrm.getAttribute('rot')}"` : ''}${xfrm.getAttribute('flipH') === '1' ? ' flipH="1"' : ''}${xfrm.getAttribute('flipV') === '1' ? ' flipV="1"' : ''}` : '';
  const next = parseFragment(`<a:xfrm${flip}><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></a:xfrm>`);
  const imported = doc.importNode(next, true);
  if (xfrm) spPr.replaceChild(imported, xfrm);
  else spPr.insertBefore(imported, spPr.firstChild);
  return serializeXml(wsp);
}

/**
 * The run of a group of single shapes (their boxes on the page, EMU from one origin), anchored
 * where the first is: a wpg:wgp whose children keep their size and place, with a VML copy.
 */
export function groupXml(items: Placed[], docId: number, childIds: number[], texts: (item: number, f: TextFrame) => string): string {
  const first = items[0].model;
  const x0 = Math.min(...items.map((p) => p.x));
  const y0 = Math.min(...items.map((p) => p.y));
  const x1 = Math.max(...items.map((p) => p.x + p.model.w));
  const y1 = Math.max(...items.map((p) => p.y + p.model.h));
  const w = x1 - x0;
  const h = y1 - y0;
  const kids = items.map((p, i) => wspOf(p.xml, { x: p.x - x0, y: p.y - y0, w: p.model.w, h: p.model.h }, childIds[i], p.model.name ?? `圖案 ${childIds[i]}`));
  // Where the first one's references put the group's top left corner.
  const ax = (first.anchor!.h.offset ?? 0) - (items[0].x - x0);
  const ay = (first.anchor!.v.offset ?? 0) - (items[0].y - y0);
  const kidNodes: ShapeNode[] = items.map((p, i) => ({ ...(p.model.root as SpNode), x: p.x - x0, y: p.y - y0, w: p.model.w, h: p.model.h, id: childIds[i], name: p.model.name }));
  const root: GrpNode = { t: 'grp', x: 0, y: 0, w, h, ch: { x: 0, y: 0, w, h }, kids: kidNodes };
  const model: ShapeModel = {
    key: '',
    kind: 'group',
    src: 'dml',
    w,
    h,
    anchor: { ...first.anchor!, h: { rel: first.anchor!.h.rel, offset: Math.round(ax) }, v: { rel: first.anchor!.v.rel, offset: Math.round(ay) }, z: Math.max(...items.map((p) => p.model.anchor?.z ?? 0)) },
    root,
    drawable: true,
    docId,
    name: `群組 ${docId % 1000}`,
  };
  // The drawing written by hand around the children's own wps:wsp (their text and styles kept).
  const shell = drawingXml({ ...model, root: { ...root, kids: [] } }, () => EMPTY_TEXT);
  const drawing = shell.replace('</wpg:grpSpPr></wpg:wgp>', () => `</wpg:grpSpPr>${kids.join('')}</wpg:wgp>`);
  let n = 0;
  const vml = vmlXml(model, (f) => texts(kidIndexOf(kidNodes, f, n++), f));
  return `<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wpg">${drawing}</mc:Choice><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent></w:r>`;
}

function kidIndexOf(kids: ShapeNode[], f: TextFrame, fallback: number): number {
  const i = kids.findIndex((k) => k.t === 'sp' && k.text === f);
  return i >= 0 ? i : fallback;
}

/**
 * A group taken apart: each child as a shape of its own, placed where it is (EMU from the
 * group's references), stacked as the group was. Only groups of single shapes (wps) come apart.
 */
export function ungroupXml(xml: string, m: ShapeModel, docIds: number[]): { xml: string; model: ShapeModel }[] | null {
  if (m.kind !== 'group' || m.root.t !== 'grp' || !m.anchor) return null;
  const g = m.root;
  const run = parseFragment(xml);
  const holder = Array.from(run.children).find((c) => shapeSource(c));
  const src = holder ? shapeSource(holder) : null;
  const wgp = src?.kind === 'dml' ? src.primary.getElementsByTagNameNS(SHAPE_NS.wpg, 'wgp')[0] : null;
  if (!wgp) return null;
  const wsps = Array.from(wgp.children).filter((c) => c.namespaceURI === SHAPE_NS.wps && c.localName === 'wsp');
  if (wsps.length !== g.kids.length || g.kids.some((k) => k.t !== 'sp')) return null;
  const sx = g.ch.w ? m.w / g.ch.w : 1;
  const sy = g.ch.h ? m.h / g.ch.h : 1;
  const rPr = Array.from(run.children).find((c) => c.namespaceURI === NS.w && c.localName === 'rPr');
  return (g.kids as SpNode[]).map((k, i) => {
    const wsp = wsps[i].cloneNode(true) as Element;
    const nv = child(wsp, 'cNvPr', SHAPE_NS.wps);
    const name = nv?.getAttribute('name') ?? k.name ?? `圖案 ${docIds[i]}`;
    const descr = nv?.getAttribute('descr');
    if (nv) wsp.removeChild(nv);
    const w = k.w * sx;
    const h = k.h * sy;
    const spPr = child(wsp, 'spPr', SHAPE_NS.wps)!;
    const xfrm = child(spPr, 'xfrm', NS.a);
    const turn = xfrm ? ['rot', 'flipH', 'flipV'].filter((a) => xfrm.getAttribute(a)).map((a) => ` ${a}="${xfrm.getAttribute(a)}"`).join('') : '';
    const next = wsp.ownerDocument.importNode(parseFragment(`<a:xfrm${turn}><a:off x="0" y="0"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm>`), true);
    if (xfrm) spPr.replaceChild(next, xfrm);
    else spPr.insertBefore(next, spPr.firstChild);
    const model: ShapeModel = {
      key: '',
      kind: k.text && !k.fill ? 'textbox' : 'shape',
      src: 'dml',
      w: Math.round(w),
      h: Math.round(h),
      anchor: {
        ...m.anchor!,
        h: { rel: m.anchor!.h.rel, offset: Math.round((m.anchor!.h.offset ?? 0) + (k.x - g.ch.x) * sx) },
        v: { rel: m.anchor!.v.rel, offset: Math.round((m.anchor!.v.offset ?? 0) + (k.y - g.ch.y) * sy) },
        z: m.anchor!.z + i,
      },
      root: { ...k, x: 0, y: 0, w: Math.round(w), h: Math.round(h) },
      drawable: true,
      docId: docIds[i],
      name,
    };
    if (descr) model.descr = descr;
    const shell = drawingXml({ ...model, root: { ...model.root, text: undefined } as SpNode }, () => EMPTY_TEXT);
    const drawing = shell.replace(/<wps:wsp>[\s\S]*<\/wps:wsp>/, () => serializeXml(wsp));
    const txbx = Array.from(wsp.getElementsByTagNameNS(NS.w, 'txbxContent'))[0];
    const inner = txbx ? Array.from(txbx.childNodes).map((c) => serializeXml(c)).join('') : EMPTY_TEXT;
    const vml = vmlXml(model, () => inner);
    return {
      xml: `<w:r>${rPr ? serializeXml(rPr) : '<w:rPr><w:noProof/></w:rPr>'}<mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent></w:r>`,
      model,
    };
  });
}

export type { Paint };
export { textFrames };

// ----- copies -----

/**
 * A copy of a shape (pasted, dragged with Ctrl) with ids of its own: its docPr, the ids of the
 * shapes inside and what its connectors are attached to by them. `used` holds the document's ids.
 */
export function reidentified(xml: string, m: ShapeModel, used: Set<number>): { xml: string; model: ShapeModel } {
  const map = new Map<number, number>();
  const idOf = (old: number) => {
    let n = map.get(old);
    if (n == null) map.set(old, (n = newDrawingId(used)));
    return n;
  };
  const out = xml.replace(/(<(?:wp:docPr|wps:cNvPr|pic:cNvPr|wpg:cNvPr)\b[^>]*?\sid=")(\d+)"/g, (_all, head: string, id: string) => `${head}${idOf(Number(id))}"`)
    .replace(/(<a:(?:stCxn|endCxn)\b[^>]*?\sid=")(\d+)"/g, (_all, head: string, id: string) => `${head}${map.get(Number(id)) ?? id}"`);
  const node = (n: ShapeNode): ShapeNode => {
    if (n.t === 'grp') return { ...n, kids: n.kids.map(node) };
    if (n.t !== 'sp') return n;
    const next: SpNode = { ...n };
    if (n.id != null) next.id = map.get(n.id) ?? n.id;
    if (n.st) next.st = { ...n.st, id: map.get(n.st.id) ?? n.st.id };
    if (n.end) next.end = { ...n.end, id: map.get(n.end.id) ?? n.end.id };
    return next;
  };
  const model: ShapeModel = { ...m, root: node(m.root) };
  if (m.docId != null) model.docId = map.get(m.docId) ?? idOf(m.docId);
  return { xml: out, model };
}
