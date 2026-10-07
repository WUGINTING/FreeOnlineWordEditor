// Text boxes and drawn shapes: a model read from the file, for drawing them and editing their text.
//
// A run holding a shape stays a kept-as-is node (raw_inline) whose XML is written back
// byte for byte; the model (raw_inline.attrs.shape) is what the editor shows and edits:
//  - where it is: in the line (wp:inline) or anchored to a paragraph (wp:anchor: relative to
//    the page, margin, column, paragraph ..., with an offset or an alignment, and its wrapping),
//  - what it looks like: a tree of shapes (a text box, a shape, a group, a drawing canvas) with
//    their boxes in EMU in their parent's coordinates, outline, fill and line,
//  - the text of each text box (w:txbxContent) as an editor document.
// Word 2010+ writes a shape twice (mc:AlternateContent): as DrawingML (mc:Choice, read here) and
// as VML for older Word (mc:Fallback). A file with only VML (w:pict) is read from the VML.
//
// Saving: only a text box whose text was edited is written again, into both copies (see
// rewriteTextBoxes); everything else in the run keeps its XML.
//
// Phase 8B (moving, resizing, inserting shapes and connectors) changes the model's boxes and
// anchor and writes them back the same way: the numbers here are the file's own (EMU, the
// DrawingML element each came from), so they can be put back where they were read.

import type { Node as PMNode } from 'prosemirror-model';
import { NS, child, children, parseFragment, serializeXml } from './xml';
import { colorIn, readColor, vmlColor, type Color, type ThemeColors } from './theme';
import { isKnownPreset, isLinePreset, type CustomPath } from './shapeGeometry';
import { tl } from '../i18n';

export const SHAPE_NS = {
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
} as const;

/** English Metric Units: 914400 per inch, 12700 per point, 9525 per CSS pixel. */
export type Emu = number;

export interface Paint {
  /** RRGGBB. */
  color: string;
  /** 0..1, absent when opaque. */
  alpha?: number;
}

export interface LineEnd {
  /** triangle, stealth, arrow, diamond, oval. */
  type: string;
  w?: string;
  len?: string;
}

export interface Stroke extends Paint {
  width: Emu;
  /** a:prstDash (dash, sysDot ...), absent for a solid line. */
  dash?: string;
  /** Arrowheads at the start (head) and the end (tail) of the line. */
  head?: LineEnd;
  tail?: LineEnd;
}

export interface TextFrame {
  /** Which w:txbxContent of the drawing this is: document order in the copy that was read. */
  i: number;
  /** The text, as an editor document (ProseMirror JSON). */
  doc: Record<string, unknown>;
  /** Changed in the editor: written into the file on save (see rewriteTextBoxes). */
  edited?: boolean;
  /** Space between the shape's text rectangle and the text [left, top, right, bottom]. */
  inset: [Emu, Emu, Emu, Emu];
  /** Where the text sits vertically. */
  anchor: 't' | 'ctr' | 'b';
  /** Text direction (a:bodyPr/@vert: vert, eaVert, vert270 ...); absent for horizontal text. */
  vert?: string;
}

/** A box in its parent's coordinates, turned (degrees, clockwise, about its centre) and flipped. */
export interface Box {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface SpNode extends Box {
  t: 'sp';
  /** Preset geometry (a:prstGeom/@prst), or 'custom' (a:custGeom, see paths). */
  geom: string;
  /** Adjust values (a:avLst), 1/100000. */
  av?: Record<string, number>;
  paths?: CustomPath[];
  fill: Paint | null;
  line: Stroke | null;
  text?: TextFrame;
  /** A connector (wps:cNvCnPr). */
  cxn?: boolean;
  /** What a connector's start and end are attached to (a:stCxn / a:endCxn: a shape's id and connection site). */
  st?: Connection;
  end?: Connection;
  /** Its id (wps:cNvPr/@id, in a group or on a canvas): what connectors are attached by. */
  id?: number;
  /** Something about it can't be drawn (a picture fill, an unknown shape ...). */
  unknown?: boolean;
  /** WordArt (VML v:textpath, a watermark): its text, stretched over the box in its fill colour. */
  art?: { text: string; font?: string };
  name?: string;
  descr?: string;
}

export interface Connection {
  id: number;
  idx: number;
}

export interface PicNode extends Box {
  t: 'pic';
  src: string | null;
  descr?: string;
}

export interface GrpNode extends Box {
  t: 'grp';
  /** The coordinates the children are given in (a:chOff / a:chExt), mapped onto the box. */
  ch: { x: Emu; y: Emu; w: Emu; h: Emu };
  /** A drawing canvas's background. */
  bg?: Paint | null;
  kids: ShapeNode[];
}

export type ShapeNode = SpNode | PicNode | GrpNode;

/** One axis of an anchored shape's position (wp:positionH / wp:positionV). */
export interface AxisPos {
  /** page, margin, column, character, leftMargin ... / page, margin, paragraph, line, topMargin ... */
  rel: string;
  offset?: Emu;
  /** left, center, right, inside, outside / top, center, bottom, inside, outside. */
  align?: string;
  /** wp14:pctPosHOffset / pctPosVOffset, 1/1000 of a percent of the reference area. */
  pct?: number;
}

export interface Anchor {
  h: AxisPos;
  v: AxisPos;
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** wrapText: bothSides, left, right, largest. */
  side?: string;
  /** Distance from the text [top, bottom, left, right]. */
  dist: [Emu, Emu, Emu, Emu];
  /** Behind the text (behindDoc). */
  behind: boolean;
  /** Stacking order (relativeHeight / z-index). */
  z: number;
  /** 鎖定錨點 (wp:anchor/@locked): the anchor stays with its paragraph when the shape is dragged. */
  locked?: boolean;
}

export interface ShapeModel {
  /** Unique on the page: styles the editor adds while its text is edited find it by this. */
  key: string;
  kind: 'textbox' | 'shape' | 'group' | 'canvas';
  /** Read from DrawingML (wps / wpg / wpc) or from VML (w:pict). */
  src: 'dml' | 'vml';
  /** Size of the drawing (wp:extent). */
  w: Emu;
  h: Emu;
  /** Room around it for its line and effects (wp:effectExtent) [left, top, right, bottom]. */
  effect?: [Emu, Emu, Emu, Emu];
  /** Anchored (floating); null when it sits in the line of text. */
  anchor: Anchor | null;
  root: ShapeNode;
  /** wp:docPr: name, title and description (alternative text). */
  name?: string;
  title?: string;
  descr?: string;
  /** Everything in it can be drawn; otherwise Word's preview picture is shown. */
  drawable: boolean;
  /** wp:docPr/@id: unique in the document; what connectors outside a canvas are attached by. */
  docId?: number;
  /**
   * The model as read (or as made), without the text boxes' text, once it was changed in the editor
   * (moved, resized, recoloured ...): the save writes what differs from it (docx/shapeWrite.ts).
   * Absent while unchanged.
   */
  base?: string;
}

export interface ShapeContext {
  theme: ThemeColors;
  /** The picture of a relationship id (data: URL). */
  image(rId: string): string | undefined;
  /** A w:txbxContent's paragraphs and tables as editor blocks. */
  blocks(txbx: Element): PMNode[];
  /** The editor document for those blocks. */
  doc(blocks: PMNode[]): PMNode;
}

let keys = 0;
export function newShapeKey(): string {
  return 'S' + ++keys;
}

const W = NS.w;
const isEl = (el: Element | null | undefined, ns: string, ...locals: string[]): boolean =>
  !!el && el.namespaceURI === ns && locals.includes(el.localName);
const num = (v: string | null | undefined, d = 0) => {
  const n = Number(v);
  return v != null && v !== '' && Number.isFinite(n) ? n : d;
};
const on = (v: string | null | undefined) => v === '1' || v === 'true' || v === 'on' || v === 't';

// ----- which part of the run holds the drawing -----

/**
 * The element a shape's run is read from: the w:drawing of the mc:Choice Word 2010+ reads (wps,
 * wpg, wpc), else the VML w:pict (a file with VML only, or the mc:Fallback). `fallback` is the
 * other copy (VML), when there is one.
 */
export function shapeSource(el: Element): { primary: Element; kind: 'dml' | 'vml'; fallback: Element | null } | null {
  if (isEl(el, W, 'drawing')) return { primary: el, kind: 'dml', fallback: null };
  if (isEl(el, W, 'pict')) return { primary: el, kind: 'vml', fallback: null };
  if (!isEl(el, NS.mc, 'AlternateContent')) return null;
  const fallback = children(el, 'Fallback', NS.mc)[0] ?? null;
  for (const choice of children(el, 'Choice', NS.mc)) {
    const drawing = child(choice, 'drawing');
    if (drawing) return { primary: drawing, kind: 'dml', fallback };
  }
  const pict = fallback ? child(fallback, 'pict') : null;
  return pict ? { primary: pict, kind: 'vml', fallback: null } : null;
}

/** The w:txbxContent elements of a drawing that belong to it (not those inside its text boxes' text), in document order. */
export function ownTextBoxes(root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(W, 'txbxContent')).filter((t) => {
    for (let p = t.parentElement; p && p !== root; p = p.parentElement) if (isEl(p, W, 'txbxContent')) return false;
    return true;
  });
}

// ----- reading -----

/**
 * The shape model of a run's w:drawing, w:pict or mc:AlternateContent; null when it holds no
 * shape (a picture, chart, SmartArt, OLE object ...).
 */
export function readShape(el: Element, ctx: ShapeContext): ShapeModel | null {
  const src = shapeSource(el);
  if (!src) return null;
  try {
    const boxes = ownTextBoxes(src.primary);
    return src.kind === 'dml' ? readDrawing(src.primary, ctx, boxes) : readPict(src.primary, ctx, boxes);
  } catch {
    return null; // shown as today: kept as-is
  }
}

class Reader {
  drawable = true;
  constructor(readonly ctx: ShapeContext, readonly boxes: Element[]) {}

  text(txbx: Element | null, bodyPr: Element | null): TextFrame | undefined {
    const content = txbx ? child(txbx, 'txbxContent') : null;
    const i = content ? this.boxes.indexOf(content) : -1;
    if (!content || i < 0) return undefined;
    const lIns = num(bodyPr?.getAttribute('lIns'), 91440);
    const tIns = num(bodyPr?.getAttribute('tIns'), 45720);
    const rIns = num(bodyPr?.getAttribute('rIns'), 91440);
    const bIns = num(bodyPr?.getAttribute('bIns'), 45720);
    const a = bodyPr?.getAttribute('anchor');
    const vert = bodyPr?.getAttribute('vert');
    const frame: TextFrame = {
      i,
      doc: this.ctx.doc(this.ctx.blocks(content)).toJSON() as Record<string, unknown>,
      inset: [lIns, tIns, rIns, bIns],
      anchor: a === 'ctr' ? 'ctr' : a === 'b' ? 'b' : 't',
    };
    if (vert && vert !== 'horz') frame.vert = vert;
    return frame;
  }
}

function readDrawing(drawing: Element, ctx: ShapeContext, boxes: Element[]): ShapeModel | null {
  const frame = child(drawing, 'inline', NS.wp) ?? child(drawing, 'anchor', NS.wp);
  if (!frame) return null;
  const data = frame.getElementsByTagNameNS(NS.a, 'graphicData')[0];
  const uri = data?.getAttribute('uri') ?? '';
  const extent = child(frame, 'extent', NS.wp);
  const w = num(extent?.getAttribute('cx'));
  const h = num(extent?.getAttribute('cy'));
  const r = new Reader(ctx, boxes);
  let root: ShapeNode | null = null;
  let kind: ShapeModel['kind'] = 'shape';
  if (uri === SHAPE_NS.wps) {
    const wsp = child(data, 'wsp', SHAPE_NS.wps);
    if (!wsp) return null;
    // A lone shape fills the drawing (its own a:xfrm says the same, at 0, 0).
    const sp = readSp(wsp, r, { x: 0, y: 0, w, h });
    root = sp;
    kind = sp.text && child(wsp, 'cNvSpPr', SHAPE_NS.wps)?.getAttribute('txBox') === '1' ? 'textbox' : 'shape';
  } else if (uri === SHAPE_NS.wpg) {
    const wgp = child(data, 'wgp', SHAPE_NS.wpg);
    if (!wgp) return null;
    root = readGroup(wgp, r, { x: 0, y: 0, w, h });
    kind = 'group';
  } else if (uri === SHAPE_NS.wpc) {
    const wpc = child(data, 'wpc', SHAPE_NS.wpc);
    if (!wpc) return null;
    // A canvas's shapes are placed from its top left corner, in EMU.
    const bg = child(wpc, 'bg', SHAPE_NS.wpc);
    root = { t: 'grp', x: 0, y: 0, w, h, ch: { x: 0, y: 0, w, h }, bg: bg ? readFill(bg, r) : null, kids: readKids(wpc, r) };
    kind = 'canvas';
  } else return null;

  const docPr = child(frame, 'docPr', NS.wp);
  const ee = child(frame, 'effectExtent', NS.wp);
  const model: ShapeModel = {
    key: newShapeKey(),
    kind,
    src: 'dml',
    w,
    h,
    anchor: frame.localName === 'anchor' ? readAnchor(frame) : null,
    root,
    drawable: r.drawable && drawableTree(root),
  };
  if (ee) {
    const e: [Emu, Emu, Emu, Emu] = [num(ee.getAttribute('l')), num(ee.getAttribute('t')), num(ee.getAttribute('r')), num(ee.getAttribute('b'))];
    if (e.some((v) => v)) model.effect = e;
  }
  setAlt(model, docPr?.getAttribute('name'), docPr?.getAttribute('title'), docPr?.getAttribute('descr'));
  if (on(docPr?.getAttribute('hidden'))) model.drawable = false;
  const docId = Number(docPr?.getAttribute('id'));
  if (Number.isInteger(docId) && docId > 0) model.docId = docId;
  return model;
}

function setAlt(m: ShapeModel, name?: string | null, title?: string | null, descr?: string | null): void {
  if (name) m.name = name;
  if (title?.trim()) m.title = title.trim();
  if (descr?.trim()) m.descr = descr.trim();
}

function drawableTree(n: ShapeNode): boolean {
  if (n.t === 'grp') return !n.rot && !n.flipH && !n.flipV && n.kids.every(drawableTree);
  if (n.t === 'pic') return !!n.src;
  return !n.unknown;
}

function readAnchor(a: Element): Anchor {
  const axis = (el: Element | null, dflt: string): AxisPos => {
    const out: AxisPos = { rel: el?.getAttribute('relativeFrom') || dflt };
    const off = child(el, 'posOffset', NS.wp);
    const align = child(el, 'align', NS.wp);
    const pct = el ? (el.getElementsByTagNameNS(SHAPE_NS.wp14, 'pctPosHOffset')[0] ?? el.getElementsByTagNameNS(SHAPE_NS.wp14, 'pctPosVOffset')[0]) : null;
    if (off) out.offset = num(off.textContent);
    else if (align) out.align = (align.textContent ?? '').trim();
    else if (pct) out.pct = num(pct.textContent);
    else out.offset = 0;
    return out;
  };
  let h = axis(child(a, 'positionH', NS.wp), 'column');
  let v = axis(child(a, 'positionV', NS.wp), 'paragraph');
  if (on(a.getAttribute('simplePos'))) {
    const sp = child(a, 'simplePos', NS.wp);
    h = { rel: 'page', offset: num(sp?.getAttribute('x')) };
    v = { rel: 'page', offset: num(sp?.getAttribute('y')) };
  }
  let wrap: Anchor['wrap'] = 'none';
  let side: string | undefined;
  for (const [local, kind] of [['wrapSquare', 'square'], ['wrapTight', 'tight'], ['wrapThrough', 'through'], ['wrapTopAndBottom', 'topAndBottom']] as const) {
    const el = child(a, local, NS.wp);
    if (el) {
      wrap = kind;
      side = el.getAttribute('wrapText') ?? undefined;
    }
  }
  const out: Anchor = {
    h,
    v,
    wrap,
    dist: [num(a.getAttribute('distT')), num(a.getAttribute('distB')), num(a.getAttribute('distL')), num(a.getAttribute('distR'))],
    behind: on(a.getAttribute('behindDoc')),
    z: num(a.getAttribute('relativeHeight')),
  };
  if (on(a.getAttribute('locked'))) out.locked = true;
  if (side) out.side = side;
  return out;
}

function readXfrm(xfrm: Element | null, dflt: Box): Box {
  if (!xfrm) return { ...dflt };
  const off = child(xfrm, 'off', NS.a);
  const ext = child(xfrm, 'ext', NS.a);
  const box: Box = {
    x: off ? num(off.getAttribute('x')) : dflt.x,
    y: off ? num(off.getAttribute('y')) : dflt.y,
    w: ext ? num(ext.getAttribute('cx')) : dflt.w,
    h: ext ? num(ext.getAttribute('cy')) : dflt.h,
  };
  const rot = num(xfrm.getAttribute('rot')) / 60000;
  if (rot) box.rot = rot;
  if (on(xfrm.getAttribute('flipH'))) box.flipH = true;
  if (on(xfrm.getAttribute('flipV'))) box.flipV = true;
  return box;
}

function readKids(parent: Element, r: Reader): ShapeNode[] {
  const out: ShapeNode[] = [];
  for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
    if (isEl(c, SHAPE_NS.wps, 'wsp')) out.push(readSp(c, r, null));
    else if (isEl(c, SHAPE_NS.wpg, 'grpSp', 'wgp')) out.push(readGroup(c, r, null));
    else if (isEl(c, NS.pic, 'pic')) out.push(readPic(c, r));
    else if (isEl(c, SHAPE_NS.wpc, 'bg', 'whole') || isEl(c, SHAPE_NS.wpg, 'cNvPr', 'cNvGrpSpPr', 'grpSpPr')) continue;
    else if (isEl(c, SHAPE_NS.wps, 'cNvPr')) continue;
    else r.drawable = false; // ink, a graphic frame (chart ...) inside a group
  }
  return out;
}

function readGroup(g: Element, r: Reader, fill: Box | null): GrpNode {
  const xfrm = child(child(g, 'grpSpPr', SHAPE_NS.wpg), 'xfrm', NS.a);
  const box = readXfrm(xfrm, fill ?? { x: 0, y: 0, w: 0, h: 0 });
  const chOff = child(xfrm, 'chOff', NS.a);
  const chExt = child(xfrm, 'chExt', NS.a);
  const ch = {
    x: num(chOff?.getAttribute('x'), box.x),
    y: num(chOff?.getAttribute('y'), box.y),
    w: num(chExt?.getAttribute('cx'), box.w),
    h: num(chExt?.getAttribute('cy'), box.h),
  };
  // The drawing's own group fills the drawing: its a:off / a:ext may say where it was on the page.
  const at = fill ? { ...box, x: fill.x, y: fill.y, w: fill.w, h: fill.h } : box;
  return { t: 'grp', ...at, ch, kids: readKids(g, r) };
}

function readPic(p: Element, r: Reader): PicNode {
  const spPr = child(p, 'spPr', NS.pic);
  const box = readXfrm(child(spPr, 'xfrm', NS.a), { x: 0, y: 0, w: 0, h: 0 });
  const blip = p.getElementsByTagNameNS(NS.a, 'blip')[0];
  const id = blip?.getAttributeNS(NS.r, 'embed');
  const descr = child(child(p, 'nvPicPr', NS.pic), 'cNvPr', NS.pic)?.getAttribute('descr');
  const out: PicNode = { t: 'pic', ...box, src: (id && r.ctx.image(id)) || null };
  if (descr) out.descr = descr;
  return out;
}

/** a:solidFill / a:gradFill / a:pattFill / a:noFill inside `parent`: undefined when there is none. */
function readFill(parent: Element | null, r: Reader, placeholder?: Color | null): Paint | null | undefined {
  for (let c = parent?.firstElementChild ?? null; c; c = c.nextElementSibling) {
    if (c.namespaceURI !== NS.a) continue;
    switch (c.localName) {
      case 'noFill':
        return null;
      case 'solidFill':
        return paint(colorIn(c, r.ctx.theme, placeholder));
      case 'gradFill': {
        // The first stop's colour stands for the gradient.
        const gs = c.getElementsByTagNameNS(NS.a, 'gs')[0];
        return paint(colorIn(gs ?? null, r.ctx.theme, placeholder));
      }
      case 'pattFill':
        return paint(colorIn(child(c, 'fgClr', NS.a), r.ctx.theme, placeholder));
      case 'blipFill':
      case 'grpFill':
        r.drawable = false;
        return null;
    }
  }
  return undefined;
}

function paint(c: Color | null): Paint | null {
  if (!c) return null;
  const out: Paint = { color: c.hex };
  if (c.alpha != null) out.alpha = c.alpha;
  return out;
}

/** Line widths of the theme's line styles (Office theme: 1/2, 1 and 1.5 pt). */
const STYLE_LINE_WIDTH = [0, 6350, 12700, 19050];

function readSp(sp: Element, r: Reader, fill: Box | null): SpNode {
  const spPr = child(sp, 'spPr', SHAPE_NS.wps);
  const box = readXfrm(child(spPr, 'xfrm', NS.a), fill ?? { x: 0, y: 0, w: 0, h: 0 });
  const at = fill ? { ...box, x: fill.x, y: fill.y, w: fill.w, h: fill.h } : box;
  const node: SpNode = { t: 'sp', ...at, geom: 'rect', fill: null, line: null };

  const prst = child(spPr, 'prstGeom', NS.a);
  const cust = child(spPr, 'custGeom', NS.a);
  if (prst) {
    node.geom = prst.getAttribute('prst') || 'rect';
    const av = readAv(child(prst, 'avLst', NS.a));
    if (av) node.av = av;
    if (!isKnownPreset(node.geom)) node.unknown = true;
  } else if (cust) {
    node.geom = 'custom';
    const paths = readCustom(cust);
    if (paths) node.paths = paths;
    else node.unknown = true;
  }

  // Style references (wps:style): the colours and line a shape has unless its spPr says otherwise.
  const style = child(sp, 'style', SHAPE_NS.wps);
  const lnRef = child(style, 'lnRef', NS.a);
  const fillRef = child(style, 'fillRef', NS.a);
  const styleLine = lnRef && num(lnRef.getAttribute('idx')) > 0 ? readColor(lnRef.firstElementChild, r.ctx.theme) : null;
  const styleFill = fillRef && num(fillRef.getAttribute('idx')) > 0 ? readColor(fillRef.firstElementChild, r.ctx.theme) : null;

  const f = readFill(spPr, r, styleFill);
  node.fill = f === undefined ? paint(styleFill) : f;
  if (isLinePreset(node.geom)) node.fill = null;

  const ln = child(spPr, 'ln', NS.a);
  const styleWidth = STYLE_LINE_WIDTH[Math.min(3, num(lnRef?.getAttribute('idx')))] ?? 9525;
  if (ln || styleLine) {
    const lf = ln ? readFill(ln, r, styleLine) : undefined;
    const color = lf === undefined ? paint(styleLine) : lf;
    if (color) {
      const stroke: Stroke = { ...color, width: num(ln?.getAttribute('w'), styleWidth || 9525) };
      const dash = child(ln, 'prstDash', NS.a)?.getAttribute('val');
      if (dash && dash !== 'solid') stroke.dash = dash;
      for (const end of ['headEnd', 'tailEnd'] as const) {
        const e = child(ln, end, NS.a);
        const type = e?.getAttribute('type');
        if (e && type && type !== 'none') {
          const le: LineEnd = { type };
          if (e.getAttribute('w')) le.w = e.getAttribute('w')!;
          if (e.getAttribute('len')) le.len = e.getAttribute('len')!;
          stroke[end === 'headEnd' ? 'head' : 'tail'] = le;
        }
      }
      node.line = stroke;
    }
  }

  const nv = child(sp, 'cNvPr', SHAPE_NS.wps);
  if (nv?.getAttribute('name')) node.name = nv.getAttribute('name')!;
  if (nv?.getAttribute('descr')) node.descr = nv.getAttribute('descr')!;
  const id = Number(nv?.getAttribute('id'));
  if (nv && Number.isInteger(id) && id > 0) node.id = id;
  const cnv = child(sp, 'cNvCnPr', SHAPE_NS.wps);
  if (cnv) {
    node.cxn = true;
    const conn = (el: Element | null): Connection | undefined => {
      const cid = Number(el?.getAttribute('id'));
      const idx = Number(el?.getAttribute('idx'));
      return el && Number.isInteger(cid) && Number.isInteger(idx) ? { id: cid, idx } : undefined;
    };
    const st = conn(child(cnv, 'stCxn', NS.a));
    const end = conn(child(cnv, 'endCxn', NS.a));
    if (st) node.st = st;
    if (end) node.end = end;
  }

  const text = r.text(child(sp, 'txbx', SHAPE_NS.wps), child(sp, 'bodyPr', SHAPE_NS.wps));
  if (text) node.text = text;
  return node;
}

function readAv(avLst: Element | null): Record<string, number> | undefined {
  let out: Record<string, number> | undefined;
  for (const gd of avLst ? children(avLst, 'gd', NS.a) : []) {
    const m = /^val\s+(-?\d+)/.exec(gd.getAttribute('fmla') ?? '');
    if (m) (out ??= {})[gd.getAttribute('name') ?? ''] = Number(m[1]);
  }
  return out;
}

/** a:custGeom's paths when every point is a number (guide formulas are not worked out). */
function readCustom(cust: Element): CustomPath[] | null {
  const lst = child(cust, 'pathLst', NS.a);
  if (!lst) return null;
  const paths: CustomPath[] = [];
  const pt = (el: Element | null): [number, number] | null => {
    const x = Number(el?.getAttribute('x'));
    const y = Number(el?.getAttribute('y'));
    return el && Number.isFinite(x) && Number.isFinite(y) && el.getAttribute('x') !== '' ? [x, y] : null;
  };
  for (const p of children(lst, 'path', NS.a)) {
    const cmds: (string | number)[][] = [];
    for (let c = p.firstElementChild; c; c = c.nextElementSibling) {
      const pts = children(c, 'pt', NS.a).map(pt);
      if (pts.some((q) => !q)) return null;
      const flat = (pts as [number, number][]).flat();
      switch (c.localName) {
        case 'moveTo':
          cmds.push(['M', ...flat]);
          break;
        case 'lnTo':
          cmds.push(['L', ...flat]);
          break;
        case 'cubicBezTo':
          cmds.push(['C', ...flat]);
          break;
        case 'quadBezTo':
          cmds.push(['Q', ...flat]);
          break;
        case 'arcTo': {
          const vals = ['wR', 'hR', 'stAng', 'swAng'].map((a) => Number(c!.getAttribute(a)));
          if (vals.some((v) => !Number.isFinite(v))) return null;
          cmds.push(['A', vals[0], vals[1], vals[2] / 60000, vals[3] / 60000]);
          break;
        }
        case 'close':
          cmds.push(['Z']);
          break;
        default:
          return null;
      }
    }
    paths.push({
      w: num(p.getAttribute('w')),
      h: num(p.getAttribute('h')),
      fill: p.getAttribute('fill') !== 'none',
      stroke: p.getAttribute('stroke') !== '0' && p.getAttribute('stroke') !== 'false',
      cmds,
    });
  }
  return paths.length ? paths : null;
}

// ----- VML (w:pict) -----

/** VML shape types (o:spt of v:shapetype, the number in type="#_x0000_tN") that match a preset. */
const VML_TYPES: Record<number, string> = {
  1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid',
  9: 'hexagon', 10: 'octagon', 11: 'plus', 13: 'rightArrow', 15: 'homePlate', 20: 'line', 32: 'straightConnector1',
  33: 'bentConnector2', 34: 'bentConnector3', 35: 'bentConnector4', 36: 'bentConnector5', 38: 'curvedConnector3',
  55: 'chevron', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow',
  109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput', 112: 'flowChartPredefinedProcess',
  113: 'flowChartInternalStorage', 114: 'flowChartDocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation',
  118: 'flowChartManualInput', 119: 'flowChartManualOperation', 120: 'flowChartConnector', 121: 'flowChartPunchedCard',
  127: 'flowChartExtract', 128: 'flowChartMerge', 176: 'flowChartAlternateProcess', 177: 'flowChartOffpageConnector',
  202: 'rect',
};

const VML_SHAPES = ['shape', 'rect', 'roundrect', 'oval', 'line', 'polyline', 'arc', 'group', 'image', 'curve'];

/** CSS-like `a:b;c:d` of a VML style attribute. */
export function vmlStyle(style: string | null | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of (style ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

/** A VML length (12pt, 1in, 2cm, 5mm, 40px, 40) in EMU; bare numbers are px unless `unit` says otherwise. */
export function vmlLength(v: string | null | undefined, unit: Emu = 9525): Emu | null {
  const m = /^\s*(-?[\d.]+)\s*(pt|in|cm|mm|px|pc|emu)?\s*$/i.exec(v ?? '');
  if (!m) return null;
  const n = Number(m[1]);
  switch ((m[2] ?? '').toLowerCase()) {
    case 'pt':
      return n * 12700;
    case 'in':
      return n * 914400;
    case 'cm':
      return n * 360000;
    case 'mm':
      return n * 36000;
    case 'pc':
      return n * 152400;
    case 'px':
      return n * 9525;
    case 'emu':
      return n;
    default:
      return n * unit;
  }
}

function readPict(pict: Element, ctx: ShapeContext, boxes: Element[]): ShapeModel | null {
  const types = new Map<string, Element>();
  for (const t of Array.from(pict.getElementsByTagNameNS(SHAPE_NS.v, 'shapetype'))) types.set(t.getAttribute('id') ?? '', t);
  const top = Array.from(pict.children).find((c) => c.namespaceURI === SHAPE_NS.v && VML_SHAPES.includes(c.localName));
  if (!top) return null;
  // A picture (v:shape holding only v:imagedata) is shown as one already.
  if (top.localName !== 'group' && top.getElementsByTagNameNS(SHAPE_NS.v, 'imagedata').length && !top.getElementsByTagNameNS(W, 'txbxContent').length) return null;
  const r = new Reader(ctx, boxes);
  const style = vmlStyle(top.getAttribute('style'));
  const box = vmlBox(top, style, 12700);
  const root = top.localName === 'group' ? vmlGroup(top, r, types, { ...box, x: 0, y: 0 }) : vmlShape(top, r, types, { ...box, x: 0, y: 0 }, 12700);
  const anchored = style.position === 'absolute';
  const model: ShapeModel = {
    key: newShapeKey(),
    kind: top.localName === 'group' ? (top.getAttribute('editas') === 'canvas' ? 'canvas' : 'group') : root.t === 'sp' && root.text ? (vmlType(top, types) === 202 ? 'textbox' : 'shape') : 'shape',
    src: 'vml',
    w: box.w,
    h: box.h,
    anchor: anchored ? vmlAnchor(top, style, box) : null,
    root,
    drawable: r.drawable && drawableTree(root) && style.visibility !== 'hidden',
  };
  setAlt(model, top.getAttribute('id'), top.getAttribute('title') ?? top.getAttributeNS(SHAPE_NS.o, 'title'), top.getAttribute('alt'));
  return model;
}

/** A VML element's box: left/top (or margin-left/-top) and width/height of its style, in `unit`s when bare. */
function vmlBox(el: Element, style: Record<string, string>, unit: Emu): Box {
  let x = vmlLength(style.left ?? style['margin-left'], unit) ?? 0;
  let y = vmlLength(style.top ?? style['margin-top'], unit) ?? 0;
  let w = vmlLength(style.width, unit) ?? 0;
  let h = vmlLength(style.height, unit) ?? 0;
  if (el.localName === 'line') {
    // v:line: from / to (x,y), in the same units.
    const pt = (v: string | null, d: string) => (v ?? d).split(',').map((s) => vmlLength(s, unit) ?? 0);
    const [x1, y1] = pt(el.getAttribute('from'), '0,0');
    const [x2, y2] = pt(el.getAttribute('to'), '10,10');
    x += Math.min(x1, x2);
    y += Math.min(y1, y2);
    w = Math.abs(x2 - x1);
    h = Math.abs(y2 - y1);
    const b: Box = { x, y, w, h };
    if (x2 < x1) b.flipH = true;
    if (y2 < y1) b.flipV = true;
    return b;
  }
  const b: Box = { x, y, w, h };
  const rot = Number(style.rotation);
  if (rot) b.rot = rot;
  const flip = style.flip ?? '';
  if (/x/.test(flip)) b.flipH = true;
  if (/y/.test(flip)) b.flipV = true;
  return b;
}

function vmlType(el: Element, types: Map<string, Element>): number | null {
  const own = el.getAttributeNS(SHAPE_NS.o, 'spt');
  if (own) return Number(own);
  const ref = (el.getAttribute('type') ?? '').replace(/^#/, '');
  const t = types.get(ref);
  const spt = t?.getAttributeNS(SHAPE_NS.o, 'spt');
  if (spt) return Number(spt);
  const m = /_x0000_t(\d+)$/.exec(ref);
  return m ? Number(m[1]) : null;
}

function vmlGroup(g: Element, r: Reader, types: Map<string, Element>, box: Box): GrpNode {
  const size = (g.getAttribute('coordsize') ?? '1000,1000').split(',').map(Number);
  const origin = (g.getAttribute('coordorigin') ?? '0,0').split(',').map(Number);
  const ch = { x: origin[0] || 0, y: origin[1] || 0, w: size[0] || 1000, h: size[1] || 1000 };
  const kids: ShapeNode[] = [];
  for (let c = g.firstElementChild; c; c = c.nextElementSibling) {
    if (c.namespaceURI !== SHAPE_NS.v || !VML_SHAPES.includes(c.localName)) continue;
    const style = vmlStyle(c.getAttribute('style'));
    // Children are placed in the group's coordinates (bare numbers).
    const b = vmlBox(c, style, 1);
    if (c.localName === 'group') kids.push(vmlGroup(c, r, types, b));
    else if (c.localName === 'shape' && vmlType(c, types) === 75 && !c.getElementsByTagNameNS(SHAPE_NS.v, 'imagedata').length) continue; // a canvas's background
    else kids.push(vmlShape(c, r, types, b, 1));
  }
  return { t: 'grp', ...box, ch, kids };
}

function vmlShape(el: Element, r: Reader, types: Map<string, Element>, box: Box, unit: Emu): ShapeNode {
  const data = el.getElementsByTagNameNS(SHAPE_NS.v, 'imagedata')[0];
  if (el.localName === 'image' || (data && !el.getElementsByTagNameNS(W, 'txbxContent').length)) {
    const id = data?.getAttributeNS(NS.r, 'id') ?? el.getAttributeNS(NS.r, 'id');
    return { t: 'pic', ...box, src: (id && r.ctx.image(id)) || null };
  }
  const node: SpNode = { t: 'sp', ...box, geom: 'rect', fill: null, line: null };
  switch (el.localName) {
    case 'rect':
      break;
    case 'roundrect': {
      node.geom = 'roundRect';
      const a = el.getAttribute('arcsize');
      if (a) {
        const v = a.endsWith('f') ? Number(a.slice(0, -1)) / 65536 : a.endsWith('%') ? Number(a.slice(0, -1)) / 100 : Number(a);
        // VML arcsize is of the shorter side; DrawingML's adj is half of that.
        if (Number.isFinite(v)) node.av = { adj: Math.round(v * 50000) };
      }
      break;
    }
    case 'oval':
      node.geom = 'ellipse';
      break;
    case 'line':
      node.geom = 'line';
      break;
    case 'shape': {
      const spt = vmlType(el, types);
      const prst = spt != null ? VML_TYPES[spt] : undefined;
      // WordArt text (v:textpath), e.g. a 「草稿」 or 「密」 watermark: drawn as its text.
      const path = child(el, 'textpath', SHAPE_NS.v);
      const text = path?.getAttribute('string');
      if (path && text && path.getAttribute('on') !== 'f' && path.getAttribute('on') !== 'false') {
        node.art = { text, font: vmlStyle(path.getAttribute('style'))['font-family']?.replace(/^&quot;|&quot;$|^"|"$/g, '') };
      } else if (prst) node.geom = prst;
      else node.unknown = true;
      break;
    }
    default:
      node.unknown = true;
  }
  const typeEl = types.get((el.getAttribute('type') ?? '').replace(/^#/, ''));
  const get = (a: string) => el.getAttribute(a) ?? typeEl?.getAttribute(a) ?? null;
  const fillEl = child(el, 'fill', SHAPE_NS.v);
  const strokeEl = child(el, 'stroke', SHAPE_NS.v);
  const filled = get('filled');
  if (!isLinePreset(node.geom) && !(filled === 'f' || filled === 'false') && fillEl?.getAttribute('on') !== 'f' && fillEl?.getAttribute('on') !== 'false') {
    const color = vmlColor(fillEl?.getAttribute('color') ?? get('fillcolor') ?? '#ffffff') ?? 'FFFFFF';
    const p: Paint = { color };
    const op = fillEl?.getAttribute('opacity');
    if (op) {
      const v = op.endsWith('f') ? Number(op.slice(0, -1)) / 65536 : Number(op);
      if (Number.isFinite(v) && v < 1) p.alpha = v;
    }
    if (fillEl?.getAttribute('type') && !['solid', 'gradient', 'gradientRadial'].includes(fillEl.getAttribute('type')!)) r.drawable = false;
    node.fill = p;
  }
  const stroked = get('stroked');
  if (!(stroked === 'f' || stroked === 'false') && strokeEl?.getAttribute('on') !== 'f' && strokeEl?.getAttribute('on') !== 'false') {
    const s: Stroke = {
      color: vmlColor(strokeEl?.getAttribute('color') ?? get('strokecolor') ?? '#000000') ?? '000000',
      width: vmlLength(strokeEl?.getAttribute('weight') ?? get('strokeweight') ?? '0.75pt', 12700) ?? 9525,
    };
    const dash = strokeEl?.getAttribute('dashstyle');
    if (dash && dash !== 'solid') s.dash = dash;
    const arrow = (a: string | null | undefined): LineEnd | undefined =>
      a && a !== 'none' ? { type: a === 'classic' || a === 'block' ? 'triangle' : a === 'open' ? 'arrow' : a } : undefined;
    const head = arrow(strokeEl?.getAttribute('startarrow'));
    const tail = arrow(strokeEl?.getAttribute('endarrow'));
    if (head) s.head = head;
    if (tail) s.tail = tail;
    node.line = s;
  }
  const alt = el.getAttribute('alt');
  if (alt) node.descr = alt;
  const tb = child(el, 'textbox', SHAPE_NS.v);
  if (tb) {
    const ins = (tb.getAttribute('inset') ?? '').split(',');
    const dflt = ['0.1in', '0.05in', '0.1in', '0.05in'];
    const inset = dflt.map((d, k) => vmlLength(ins[k]?.trim() || d, 12700) ?? 0) as [Emu, Emu, Emu, Emu];
    const content = child(tb, 'txbxContent');
    const i = content ? r.boxes.indexOf(content) : -1;
    if (content && i >= 0) {
      const va = vmlStyle(el.getAttribute('style'))['v-text-anchor'] ?? '';
      const flow = vmlStyle(tb.getAttribute('style'))['layout-flow'];
      node.text = {
        i,
        doc: r.ctx.doc(r.ctx.blocks(content)).toJSON() as Record<string, unknown>,
        inset,
        anchor: va.startsWith('middle') ? 'ctr' : va.startsWith('bottom') ? 'b' : 't',
      };
      if (flow && flow.startsWith('vertical')) node.text.vert = 'eaVert';
    }
  }
  return node;
}

function vmlAnchor(el: Element, style: Record<string, string>, box: Box): Anchor {
  const hRel: Record<string, string> = { margin: 'margin', page: 'page', text: 'column', char: 'character', 'left-margin-area': 'leftMargin', 'right-margin-area': 'rightMargin', 'inner-margin-area': 'insideMargin', 'outer-margin-area': 'outsideMargin' };
  const vRel: Record<string, string> = { margin: 'margin', page: 'page', text: 'paragraph', line: 'line', 'top-margin-area': 'topMargin', 'bottom-margin-area': 'bottomMargin', 'inner-margin-area': 'insideMargin', 'outer-margin-area': 'outsideMargin' };
  const axis = (pos: string | undefined, rel: string | undefined, map: Record<string, string>, dflt: string, offset: Emu): AxisPos => {
    const out: AxisPos = { rel: map[rel ?? ''] ?? dflt };
    if (pos && pos !== 'absolute') out.align = pos;
    else out.offset = offset;
    return out;
  };
  const wrapEl = child(el, 'wrap', SHAPE_NS.w10);
  const type = wrapEl?.getAttribute('type');
  const wrap: Anchor['wrap'] = type === 'square' ? 'square' : type === 'tight' ? 'tight' : type === 'through' ? 'through' : type === 'topAndBottom' ? 'topAndBottom' : 'none';
  const dist = (k: string, d: Emu) => vmlLength(style[`mso-wrap-distance-${k}`], 12700) ?? d;
  const z = Number(style['z-index']) || 0;
  return {
    h: axis(style['mso-position-horizontal'], style['mso-position-horizontal-relative'], hRel, 'column', box.x),
    v: axis(style['mso-position-vertical'], style['mso-position-vertical-relative'], vRel, 'paragraph', box.y),
    wrap,
    dist: [dist('top', 0), dist('bottom', 0), dist('left', 114300), dist('right', 114300)],
    behind: z < 0,
    z: Math.abs(z),
  };
}

// ----- the model -----

/** Every text box of the drawing, in order (the index of each is TextFrame.i). */
export function textFrames(m: ShapeModel): TextFrame[] {
  const out: TextFrame[] = [];
  const visit = (n: ShapeNode) => {
    if (n.t === 'grp') n.kids.forEach(visit);
    else if (n.t === 'sp' && n.text) out.push(n.text);
  };
  visit(m.root);
  return out;
}

/** A copy of the model whose text box `i` has the text `doc` (edited). */
export function withText(m: ShapeModel, i: number, doc: Record<string, unknown>): ShapeModel {
  const visit = (n: ShapeNode): ShapeNode => {
    if (n.t === 'grp') return { ...n, kids: n.kids.map(visit) };
    if (n.t === 'sp' && n.text?.i === i) return { ...n, text: { ...n.text, doc, edited: true } };
    return n;
  };
  return { ...m, root: visit(m.root) };
}

/** Whether a text box's text was changed in the editor. */
export function hasEditedText(m: ShapeModel | null | undefined): boolean {
  return !!m && textFrames(m).some((f) => f.edited);
}

/** What a screen reader says for the shape: its title and description, else what it is. */
export function shapeLabel(m: ShapeModel): string {
  const what = m.kind === 'textbox' ? tl('文字方塊') : m.kind === 'canvas' ? tl('繪圖畫布') : m.kind === 'group' ? tl('圖案群組') : tl('圖案');
  const alt = [m.title, m.descr].filter(Boolean).join(tl('：'));
  return alt ? tl('{0}：{1}', what, alt) : what;
}

// ----- writing -----

/**
 * The run's XML with the edited text boxes written in: in the copy the model was read from and
 * in the other one (the VML mc:Fallback, found by having the same text box XML as the first, or
 * else by its place), so Word 2007 and newer Word show the same text. Nothing else changes; a
 * model without edits gives `xml` back as it is. `write` gives a text box's new content (w:p ...).
 */
export function rewriteTextBoxes(xml: string, m: ShapeModel, write: (frame: TextFrame) => string): string {
  const frames = textFrames(m).filter((f) => f.edited);
  if (!frames.length) return xml;
  const run = parseFragment(xml);
  const el = Array.from(run.children).find((c) => shapeSource(c));
  const src = el ? shapeSource(el) : null;
  if (!src) return xml;
  const primary = ownTextBoxes(src.primary);
  const fallback = src.fallback ? ownTextBoxes(src.fallback) : [];
  const originals = fallback.map((b) => serializeXml(b));
  const used = new Set<number>();
  for (const f of frames) {
    const target = primary[f.i];
    if (!target) continue;
    const before = serializeXml(target);
    const inner = write(f);
    replaceContent(target, inner);
    // The one in the same place with the same content, else one with the same content (the VML
    // copy may list them in another order), else the one in the same place: two text boxes with
    // the same text (no paragraph ids) pair up by place.
    let twin = fallback.length === primary.length && !used.has(f.i) && originals[f.i] === before ? f.i : -1;
    if (twin < 0) twin = originals.findIndex((x, k) => !used.has(k) && x === before);
    if (twin < 0 && fallback.length === primary.length && !used.has(f.i)) twin = f.i;
    if (twin >= 0) {
      used.add(twin);
      replaceContent(fallback[twin], inner);
    }
  }
  return serializeXml(run);
}

function replaceContent(target: Element, inner: string): void {
  const holder = parseFragment(`<w:txbxContent>${inner}</w:txbxContent>`);
  while (target.firstChild) target.removeChild(target.firstChild);
  for (let c = holder.firstChild; c; c = c.nextSibling) target.appendChild(target.ownerDocument.importNode(c, true));
}
