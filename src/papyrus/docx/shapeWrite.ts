// Writing shapes back (phase 8B): a changed shape's XML patched from its model, and new shapes.
//
// A shape changed in the editor (moved, resized, recoloured, wrapped otherwise, restacked, its
// connectors rerouted) keeps its XML as read; on save only what differs between the model as it
// was (ShapeModel.base) and as it is now is written into it, in the DrawingML copy (wp:anchor /
// wp:inline, wp:extent, a:xfrm, fill and line, a:stCxn / a:endCxn) and in the VML copy for older
// Word (style margins, size, stacking, w10:wrap, colours). A VML copy that can't be patched to
// match (the inside of a group or canvas changed) is written again from the model. A model back
// as it was writes the XML as read.
//
// New shapes (插入 › 圖案, 文字方塊, a group) are written as Word 2010+ does: DrawingML (wps / wpg)
// in mc:AlternateContent's mc:Choice, with a VML mc:Fallback.

import { NS, child, children, escapeAttr, parseFragment, serializeXml } from './xml';
import { isLinePreset, presetPaths } from './shapeGeometry';
import {
  SHAPE_NS, ownTextBoxes, shapeSource, textFrames, type Anchor, type AxisPos, type Box, type Emu, type GrpNode, type LineEnd, type Paint,
  type ShapeModel, type ShapeNode, type SpNode, type Stroke, type TextFrame,
} from './shapes';

// ----- what changed -----

/** The model without its text (text boxes are written by rewriteTextBoxes) nor its session key. */
function geometry(m: ShapeModel): unknown {
  const strip = (n: ShapeNode): unknown => {
    if (n.t === 'grp') return { ...n, kids: n.kids.map(strip) };
    if (n.t === 'sp' && n.text) {
      const { doc: _doc, edited: _e, ...frame } = n.text;
      return { ...n, text: frame };
    }
    return n;
  };
  const { key: _k, base: _b, ...rest } = m;
  return { ...rest, root: strip(m.root) };
}

/** What a model's XML is written from: comparing two says whether the XML must change. */
export function geomKey(m: ShapeModel): string {
  return JSON.stringify(geometry(m));
}

/** `next`, a change of `m` made in the editor: it remembers the model as read (see ShapeModel.base). */
export function changedModel(m: ShapeModel, next: ShapeModel): ShapeModel {
  return { ...next, base: m.base ?? geomKey(m) };
}

/** Whether the model differs from the one its XML was read (or made) from, other than in its text. */
export function hasChangedShape(m: ShapeModel | null | undefined): boolean {
  return !!m?.base && m.base !== geomKey(m);
}

// ----- units -----

const pt = (emu: Emu) => `${Math.round((emu / 12700) * 100) / 100}pt`;
const int = (n: number) => String(Math.round(n));

const H_VML: Record<string, string> = {
  column: 'text', page: 'page', margin: 'margin', character: 'char', leftMargin: 'left-margin-area', rightMargin: 'right-margin-area',
  insideMargin: 'inner-margin-area', outsideMargin: 'outer-margin-area',
};
const V_VML: Record<string, string> = {
  paragraph: 'text', line: 'line', page: 'page', margin: 'margin', topMargin: 'top-margin-area', bottomMargin: 'bottom-margin-area',
  insideMargin: 'inner-margin-area', outsideMargin: 'outer-margin-area',
};

// ----- DrawingML -----

function axisXml(tag: 'positionH' | 'positionV', a: AxisPos): string {
  const inner = a.offset != null ? `<wp:posOffset>${int(a.offset)}</wp:posOffset>` : a.align ? `<wp:align>${escapeAttr(a.align)}</wp:align>` : `<wp:posOffset>0</wp:posOffset>`;
  return `<wp:${tag} relativeFrom="${escapeAttr(a.rel)}">${inner}</wp:${tag}>`;
}

function wrapXml(a: Anchor): string {
  switch (a.wrap) {
    case 'square':
      return `<wp:wrapSquare wrapText="${a.side ?? 'bothSides'}"/>`;
    case 'topAndBottom':
      return '<wp:wrapTopAndBottom/>';
    case 'tight':
    case 'through':
      // A new polygon would be the box: square wrapping is what that is.
      return `<wp:wrapSquare wrapText="${a.side ?? 'bothSides'}"/>`;
    default:
      return '<wp:wrapNone/>';
  }
}

function frameAttrs(m: ShapeModel): string {
  const [t, b, l, r] = m.anchor?.dist ?? [0, 0, 0, 0];
  const dist = `distT="${int(t)}" distB="${int(b)}" distL="${int(l)}" distR="${int(r)}"`;
  if (!m.anchor) return dist;
  return `${dist} simplePos="0" relativeHeight="${int(m.anchor.z)}" behindDoc="${m.anchor.behind ? 1 : 0}" locked="${m.anchor.locked ? 1 : 0}" layoutInCell="1" allowOverlap="1"`;
}

const effectXml = (m: ShapeModel) => {
  const [l, t, r, b] = m.effect ?? [0, 0, 0, 0];
  return `<wp:effectExtent l="${int(l)}" t="${int(t)}" r="${int(r)}" b="${int(b)}"/>`;
};

/** The start of a wp:anchor or wp:inline for `m`, up to (not including) wp:docPr. */
function frameHead(m: ShapeModel): string {
  if (!m.anchor) return `<wp:inline ${frameAttrs(m)}><wp:extent cx="${int(m.w)}" cy="${int(m.h)}"/>${effectXml(m)}`;
  return (
    `<wp:anchor ${frameAttrs(m)}><wp:simplePos x="0" y="0"/>${axisXml('positionH', m.anchor.h)}${axisXml('positionV', m.anchor.v)}` +
    `<wp:extent cx="${int(m.w)}" cy="${int(m.h)}"/>${effectXml(m)}${wrapXml(m.anchor)}`
  );
}

function colorXml(p: Paint): string {
  return `<a:srgbClr val="${p.color}">${p.alpha != null ? `<a:alpha val="${int(p.alpha * 100000)}"/>` : ''}</a:srgbClr>`;
}

function fillXml(p: Paint | null): string {
  return p ? `<a:solidFill>${colorXml(p)}</a:solidFill>` : '<a:noFill/>';
}

function endXml(tag: 'headEnd' | 'tailEnd', e: LineEnd | undefined): string {
  if (!e) return '';
  return `<a:${tag} type="${escapeAttr(e.type)}"${e.w ? ` w="${escapeAttr(e.w)}"` : ''}${e.len ? ` len="${escapeAttr(e.len)}"` : ''}/>`;
}

function lineXml(s: Stroke | null): string {
  if (!s) return '<a:ln><a:noFill/></a:ln>';
  return `<a:ln w="${int(s.width)}">${fillXml(s)}${s.dash ? `<a:prstDash val="${escapeAttr(s.dash)}"/>` : ''}${endXml('headEnd', s.head)}${endXml('tailEnd', s.tail)}</a:ln>`;
}

function xfrmXml(b: Box, group?: GrpNode): string {
  const attrs = `${b.rot ? ` rot="${int(b.rot * 60000)}"` : ''}${b.flipH ? ' flipH="1"' : ''}${b.flipV ? ' flipV="1"' : ''}`;
  const ch = group ? `<a:chOff x="${int(group.ch.x)}" y="${int(group.ch.y)}"/><a:chExt cx="${int(group.ch.w)}" cy="${int(group.ch.h)}"/>` : '';
  return `<a:xfrm${attrs}><a:off x="${int(b.x)}" y="${int(b.y)}"/><a:ext cx="${int(b.w)}" cy="${int(b.h)}"/>${ch}</a:xfrm>`;
}

function avXml(av: Record<string, number> | undefined): string {
  return av && Object.keys(av).length ? `<a:avLst>${Object.entries(av).map(([k, v]) => `<a:gd name="${escapeAttr(k)}" fmla="val ${int(v)}"/>`).join('')}</a:avLst>` : '<a:avLst/>';
}

function bodyPrXml(f: TextFrame | undefined): string {
  const [l, t, r, b] = f?.inset ?? [91440, 45720, 91440, 45720];
  const anchor = f?.anchor ?? 'ctr';
  return `<wps:bodyPr rot="0" vert="${f?.vert ?? 'horz'}" wrap="square" lIns="${int(l)}" tIns="${int(t)}" rIns="${int(r)}" bIns="${int(b)}" anchor="${anchor}" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
}

/**
 * A wps:wsp for a shape node. `inGroup`: in a group or on a canvas (it then has its own
 * wps:cNvPr with its id). `text` gives a text box's content (w:p ...).
 */
export function wspXml(n: SpNode, text: (f: TextFrame) => string, inGroup: boolean, box: Box = n, textBox = false): string {
  const nv = inGroup ? `<wps:cNvPr id="${int(n.id ?? 0)}" name="${escapeAttr(n.name ?? `圖案 ${n.id ?? ''}`)}"${n.descr ? ` descr="${escapeAttr(n.descr)}"` : ''}/>` : '';
  const conn = n.cxn
    ? `<wps:cNvCnPr>${n.st ? `<a:stCxn id="${int(n.st.id)}" idx="${int(n.st.idx)}"/>` : ''}${n.end ? `<a:endCxn id="${int(n.end.id)}" idx="${int(n.end.idx)}"/>` : ''}</wps:cNvCnPr>`
    : `<wps:cNvSpPr${n.text && textBox ? ' txBox="1"' : ''}/>`;
  const geom = `<a:prstGeom prst="${escapeAttr(n.geom)}">${avXml(n.av)}</a:prstGeom>`;
  const fill = isLinePreset(n.geom) ? '' : fillXml(n.fill);
  const txbx = n.text ? `<wps:txbx><w:txbxContent>${text(n.text)}</w:txbxContent></wps:txbx>` : '';
  return `<wps:wsp>${nv}${conn}<wps:spPr>${xfrmXml(box)}${geom}${fill}${lineXml(n.line)}</wps:spPr>${txbx}${n.cxn ? '<wps:bodyPr/>' : bodyPrXml(n.text)}</wps:wsp>`;
}

function kidXml(n: ShapeNode, text: (f: TextFrame) => string): string {
  if (n.t === 'sp') return wspXml(n, text, true);
  if (n.t === 'grp') return `<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrmXml(n, n)}</wpg:grpSpPr>${n.kids.map((k) => kidXml(k, text)).join('')}</wpg:grpSp>`;
  throw new Error('pictures are not made here');
}

/** The w:drawing of a new shape or group (DrawingML), from its model. */
export function drawingXml(m: ShapeModel, text: (f: TextFrame) => string): string {
  const docPr = `<wp:docPr id="${int(m.docId ?? 1)}" name="${escapeAttr(m.name ?? `圖案 ${m.docId ?? ''}`)}"${m.descr ? ` descr="${escapeAttr(m.descr)}"` : ''}${m.title ? ` title="${escapeAttr(m.title)}"` : ''}/>`;
  const r = m.root;
  let data: string;
  let uri: string;
  if (r.t === 'sp') {
    uri = SHAPE_NS.wps;
    data = wspXml(r, text, false, { ...r, x: 0, y: 0, w: m.w, h: m.h }, m.kind === 'textbox');
  } else if (r.t === 'grp' && m.kind === 'canvas') {
    uri = SHAPE_NS.wpc;
    data = `<wpc:wpc><wpc:bg>${fillXml(r.bg ?? null)}</wpc:bg><wpc:whole><a:ln><a:noFill/></a:ln></wpc:whole>${r.kids.map((k) => kidXml(k, text)).join('')}</wpc:wpc>`;
  } else if (r.t === 'grp') {
    uri = SHAPE_NS.wpg;
    data = `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrmXml({ ...r, x: 0, y: 0, w: m.w, h: m.h }, r)}</wpg:grpSpPr>${r.kids.map((k) => kidXml(k, text)).join('')}</wpg:wgp>`;
  } else throw new Error('pictures are not made here');
  const locks = r.t === 'sp' ? '' : '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>';
  return (
    `<w:drawing>${frameHead(m)}${docPr}<wp:cNvGraphicFramePr>${locks}</wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic>` +
    `</wp:${m.anchor ? 'anchor' : 'inline'}></w:drawing>`
  );
}

/** The run of a new shape: DrawingML for Word 2010+ with a VML copy for older Word. */
export function newShapeRun(m: ShapeModel, text: (f: TextFrame) => string, rPr = '<w:rPr><w:noProof/></w:rPr>'): string {
  const requires = m.kind === 'canvas' ? 'wpc' : m.root.t === 'grp' ? 'wpg' : 'wps';
  return `<w:r>${rPr}<mc:AlternateContent><mc:Choice Requires="${requires}">${drawingXml(m, text)}</mc:Choice><mc:Fallback>${vmlXml(m, text)}</mc:Fallback></mc:AlternateContent></w:r>`;
}

// ----- VML -----

/** An SVG path (M L H V C A Z, as shapeGeometry makes them) as a VML path. */
export function svgToVml(d: string, filled: boolean): string {
  const tokens = d.match(/[MLHVCAZ]|-?[\d.]+(?:e-?\d+)?/gi) ?? [];
  let out = filled ? '' : 'nf';
  let i = 0;
  let cmd = '';
  let x = 0;
  let y = 0;
  let sx = 0;
  let sy = 0;
  const num = () => Number(tokens[i++]);
  const v = (n: number) => String(Math.round(n));
  while (i < tokens.length) {
    if (/[a-z]/i.test(tokens[i])) cmd = tokens[i++].toUpperCase();
    switch (cmd) {
      case 'M':
        x = num();
        y = num();
        sx = x;
        sy = y;
        out += `m${v(x)},${v(y)}`;
        cmd = 'L';
        break;
      case 'L':
        x = num();
        y = num();
        out += `l${v(x)},${v(y)}`;
        break;
      case 'H':
        x = num();
        out += `l${v(x)},${v(y)}`;
        break;
      case 'V':
        y = num();
        out += `l${v(x)},${v(y)}`;
        break;
      case 'C': {
        const p = [num(), num(), num(), num(), num(), num()];
        out += `c${p.map(v).join(',')}`;
        x = p[4];
        y = p[5];
        break;
      }
      case 'A': {
        const [rx, ry, , large, sweep, ex, ey] = [num(), num(), num(), num(), num(), num(), num()];
        for (const c of arcToBeziers(x, y, rx, ry, !!large, !!sweep, ex, ey)) out += `c${c.map(v).join(',')}`;
        x = ex;
        y = ey;
        break;
      }
      case 'Z':
        out += 'x';
        x = sx;
        y = sy;
        break;
      default:
        i++;
    }
  }
  return out + 'e';
}

/** An SVG elliptical arc (no x-axis rotation) as cubic Béziers [x1,y1,x2,y2,x,y]. */
function arcToBeziers(x1: number, y1: number, rx: number, ry: number, large: boolean, sweep: boolean, x2: number, y2: number): number[][] {
  if (!rx || !ry) return [[x1, y1, x2, y2, x2, y2]];
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  let lambda = (dx * dx) / (rx * rx) + (dy * dy) / (ry * ry);
  if (lambda > 1) {
    lambda = Math.sqrt(lambda);
    rx *= lambda;
    ry *= lambda;
  }
  const sign = large === sweep ? -1 : 1;
  const num = rx * rx * ry * ry - rx * rx * dy * dy - ry * ry * dx * dx;
  const k = sign * Math.sqrt(Math.max(0, num / (rx * rx * dy * dy + ry * ry * dx * dx)));
  const cx = (k * rx * dy) / ry + (x1 + x2) / 2;
  const cy = (-k * ry * dx) / rx + (y1 + y2) / 2;
  const ang = (ux: number, uy: number) => Math.atan2(uy, ux);
  const t1 = ang((x1 - cx) / rx, (y1 - cy) / ry);
  let dt = ang((x2 - cx) / rx, (y2 - cy) / ry) - t1;
  if (sweep && dt < 0) dt += 2 * Math.PI;
  else if (!sweep && dt > 0) dt -= 2 * Math.PI;
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
  const step = dt / n;
  const out: number[][] = [];
  const kappa = (4 / 3) * Math.tan(step / 4);
  for (let j = 0; j < n; j++) {
    const a = t1 + j * step;
    const b = a + step;
    const p = (t: number) => [cx + rx * Math.cos(t), cy + ry * Math.sin(t)];
    const d = (t: number) => [-rx * Math.sin(t), ry * Math.cos(t)];
    const [ax, ay] = p(a);
    const [bx, by] = p(b);
    const [dax, day] = d(a);
    const [dbx, dby] = d(b);
    out.push([ax + kappa * dax, ay + kappa * day, bx - kappa * dbx, by - kappa * dby, bx, by]);
  }
  return out;
}

/** The VML style of a top-level shape: where it is, how big, stacked and turned. */
function vmlTopStyle(m: ShapeModel, root: Box): string[] {
  const out: string[] = [];
  if (m.anchor) {
    const a = m.anchor;
    out.push('position:absolute');
    out.push(`margin-left:${pt(a.h.offset ?? 0)}`, `margin-top:${pt(a.v.offset ?? 0)}`);
  }
  out.push(`width:${pt(m.w)}`, `height:${pt(m.h)}`);
  if (m.anchor) {
    const a = m.anchor;
    out.push(`z-index:${a.behind ? -Math.max(1, Math.round(a.z)) : Math.max(1, Math.round(a.z))}`);
    if (root.rot) out.push(`rotation:${Math.round(root.rot * 100) / 100}`);
    if (root.flipH || root.flipV) out.push(`flip:${[root.flipH ? 'x' : '', root.flipV ? 'y' : ''].filter(Boolean).join(' ')}`);
    out.push('visibility:visible', 'mso-wrap-style:square');
    const [t, b, l, r] = a.dist;
    out.push(`mso-wrap-distance-left:${pt(l)}`, `mso-wrap-distance-top:${pt(t)}`, `mso-wrap-distance-right:${pt(r)}`, `mso-wrap-distance-bottom:${pt(b)}`);
    out.push(`mso-position-horizontal:${a.h.align ?? 'absolute'}`, `mso-position-horizontal-relative:${H_VML[a.h.rel] ?? 'text'}`);
    out.push(`mso-position-vertical:${a.v.align ?? 'absolute'}`, `mso-position-vertical-relative:${V_VML[a.v.rel] ?? 'text'}`);
  } else {
    if (root.rot) out.push(`rotation:${Math.round(root.rot * 100) / 100}`);
  }
  return out;
}

function vmlWrap(m: ShapeModel): string {
  const w = m.anchor?.wrap;
  if (!m.anchor || w === 'none') return '';
  return `<w10:wrap type="${w === 'topAndBottom' ? 'topAndBottom' : 'square'}"/>`;
}

const ARROWS: Record<string, string> = { triangle: 'block', stealth: 'classic', arrow: 'open', oval: 'oval', diamond: 'diamond' };

function vmlPaint(n: SpNode): string {
  const line = isLinePreset(n.geom);
  let out = line || !n.fill ? ' filled="f"' : ` fillcolor="#${n.fill.color}"`;
  out += n.line ? ` strokecolor="#${n.line.color}" strokeweight="${pt(n.line.width)}"` : ' stroked="f"';
  return out;
}

function vmlStroke(n: SpNode): string {
  const s = n.line;
  if (!s || (!s.dash && !s.head && !s.tail && s.alpha == null)) return n.fill?.alpha != null ? `<v:fill opacity="${n.fill.alpha}"/>` : '';
  const a = [
    s.dash ? ` dashstyle="${s.dash === 'sysDot' ? 'dot' : s.dash === 'sysDash' ? 'shortdash' : s.dash.toLowerCase()}"` : '',
    s.head ? ` startarrow="${ARROWS[s.head.type] ?? 'block'}"` : '',
    s.tail ? ` endarrow="${ARROWS[s.tail.type] ?? 'block'}"` : '',
    s.alpha != null ? ` opacity="${s.alpha}"` : '',
  ].join('');
  return `${n.fill?.alpha != null ? `<v:fill opacity="${n.fill.alpha}"/>` : ''}<v:stroke${a}/>`;
}

function vmlText(n: SpNode, text: (f: TextFrame) => string): string {
  if (!n.text) return '';
  const [l, t, r, b] = n.text.inset;
  return `<v:textbox inset="${pt(l)},${pt(t)},${pt(r)},${pt(b)}"${n.text.vert ? ' style="layout-flow:vertical"' : ''}><w:txbxContent>${text(n.text)}</w:txbxContent></v:textbox>`;
}

const VML_ANCHOR: Record<string, string> = { t: 'top', ctr: 'middle', b: 'bottom' };

/** A VML element for a shape node with the given style (position and size in the parent's terms). */
function vmlShapeXml(n: SpNode, style: string[], text: (f: TextFrame) => string, id: string): string {
  const s = [...style];
  if (n.text) s.push(`v-text-anchor:${VML_ANCHOR[n.text.anchor] ?? 'top'}`);
  const inner = vmlStroke(n) + vmlText(n, text);
  const attrs = `id="${escapeAttr(id)}"${n.descr ? ` alt="${escapeAttr(n.descr)}"` : ''}`;
  if (n.geom === 'rect' || n.geom === 'flowChartProcess') return `<v:rect ${attrs} style="${s.join(';')}"${vmlPaint(n)}>${inner}</v:rect>`;
  if (n.geom === 'ellipse' || n.geom === 'flowChartConnector') return `<v:oval ${attrs} style="${s.join(';')}"${vmlPaint(n)}>${inner}</v:oval>`;
  if (n.geom === 'roundRect' || n.geom === 'flowChartAlternateProcess') {
    const adj = (n.av?.adj ?? 16667) / 50000;
    return `<v:roundrect ${attrs} style="${s.join(';')}" arcsize="${Math.round(adj * 65536)}f"${vmlPaint(n)}>${inner}</v:roundrect>`;
  }
  // Anything else: its outline as a path in a 21600 × 21600 box (turned and flipped by its style).
  const paths = presetPaths(n.geom, 21600, 21600, n.av) ?? presetPaths('rect', 21600, 21600)!;
  const path = paths.map((p) => svgToVml(p.d, p.fill)).join('');
  return `<v:shape ${attrs} coordsize="21600,21600" path="${path}" style="${s.join(';')}"${vmlPaint(n)}>${inner}</v:shape>`;
}

let vmlIds = 0;

/** The VML copy (w:pict) of a shape, from its model, for Word 2007 and older. */
export function vmlXml(m: ShapeModel, text: (f: TextFrame) => string): string {
  const r = m.root;
  const id = `_x0000_s${2000 + (++vmlIds % 8000)}`;
  const alt = `${m.descr ? ` alt="${escapeAttr(m.descr)}"` : ''}${m.title ? ` title="${escapeAttr(m.title)}"` : ''}`;
  if (r.t === 'sp') {
    const xml = vmlShapeXml({ ...r, descr: undefined }, vmlTopStyle(m, r), text, m.name ?? id);
    return `<w:pict>${xml.replace(/^(<v:\w+ )/, `$1o:spid="${id}"${alt} `).replace(/(<\/v:\w+>)$/, `${vmlWrap(m)}$1`)}</w:pict>`;
  }
  if (r.t !== 'grp') return '<w:pict/>';
  // Children in hundredths of an EMU of the group's own coordinates.
  const u = (v: Emu) => Math.round(v / 100);
  const kids = (g: GrpNode): string =>
    g.kids
      .map((k, i) => {
        const style = [`position:absolute`, `left:${u(k.x)}`, `top:${u(k.y)}`, `width:${u(k.w)}`, `height:${u(k.h)}`];
        if (k.rot) style.push(`rotation:${Math.round(k.rot * 100) / 100}`);
        if (k.flipH || k.flipV) style.push(`flip:${[k.flipH ? 'x' : '', k.flipV ? 'y' : ''].filter(Boolean).join(' ')}`);
        if (k.t === 'sp') return vmlShapeXml(k, style, text, k.name ?? `${id}_${i}`);
        if (k.t === 'grp') return `<v:group style="${style.join(';')}" coordorigin="${u(k.ch.x)},${u(k.ch.y)}" coordsize="${u(k.ch.w)},${u(k.ch.h)}">${kids(k)}</v:group>`;
        return '';
      })
      .join('');
  const canvas = m.kind === 'canvas' ? ' editas="canvas"' : '';
  const bg = m.kind === 'canvas' && r.bg ? `<v:rect style="position:absolute;left:0;top:0;width:${u(r.w)};height:${u(r.h)}" fillcolor="#${r.bg.color}" stroked="f"/>` : '';
  return (
    `<w:pict><v:group id="${escapeAttr(m.name ?? id)}" o:spid="${id}"${alt}${canvas} style="${vmlTopStyle(m, r).join(';')}" coordorigin="${u(r.ch.x)},${u(r.ch.y)}" coordsize="${u(r.ch.w)},${u(r.ch.h)}">` +
    `${bg}${kids(r)}${vmlWrap(m)}</v:group></w:pict>`
  );
}

// ----- patching a shape's XML -----

const A = NS.a;
const WP = NS.wp;

function el(doc: Document, xml: string): Element {
  const frag = parseFragment(xml);
  return doc.importNode(frag, true) as Element;
}

/** Put `xml` where `old` is. */
function replaceWith(old: Element, xml: string): Element {
  const next = el(old.ownerDocument, xml);
  old.parentNode!.replaceChild(next, old);
  return next;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** The shape elements of a group or canvas in the order readKids reads them. */
function kidElements(parent: Element): Element[] {
  return Array.from(parent.children).filter(
    (c) => (c.namespaceURI === SHAPE_NS.wps && c.localName === 'wsp') || (c.namespaceURI === SHAPE_NS.wpg && (c.localName === 'grpSp' || c.localName === 'wgp')) || (c.namespaceURI === NS.pic && c.localName === 'pic'),
  );
}

/** Set an a:xfrm (off, ext, turn, flips) to a box; creates it where missing. */
function patchXfrm(spPr: Element, b: Box, before: Box, group?: GrpNode, groupBefore?: GrpNode): void {
  if (same(b, before) && same(group?.ch, groupBefore?.ch)) return;
  const xfrm = child(spPr, 'xfrm', A);
  const xml = xfrmXml(b, group);
  if (xfrm) replaceWith(xfrm, xml);
  else spPr.insertBefore(el(spPr.ownerDocument, xml), spPr.firstChild);
}

const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];

function patchSp(sp: Element, n: SpNode, was: SpNode, box: Box, boxBefore: Box): void {
  const spPr = child(sp, 'spPr', SHAPE_NS.wps);
  if (!spPr) return;
  patchXfrm(spPr, box, boxBefore);
  if (!same(n.av, was.av) || n.geom !== was.geom) {
    const prst = child(spPr, 'prstGeom', A);
    if (prst) replaceWith(prst, `<a:prstGeom prst="${escapeAttr(n.geom)}">${avXml(n.av)}</a:prstGeom>`);
  }
  if (!same(n.fill, was.fill) && !isLinePreset(n.geom)) {
    const old = Array.from(spPr.children).find((c) => c.namespaceURI === A && FILLS.includes(c.localName));
    const xml = fillXml(n.fill);
    if (old) replaceWith(old, xml);
    else {
      const geom = child(spPr, 'prstGeom', A) ?? child(spPr, 'custGeom', A);
      spPr.insertBefore(el(spPr.ownerDocument, xml), geom ? geom.nextSibling : null);
    }
  }
  if (!same(n.line, was.line)) {
    const ln = child(spPr, 'ln', A);
    if (ln && n.line && was.line) {
      // Keep what the editor doesn't model (caps, joins, compound lines).
      ln.setAttribute('w', int(n.line.width));
      const old = Array.from(ln.children).find((c) => c.namespaceURI === A && FILLS.includes(c.localName));
      const fill = el(ln.ownerDocument, fillXml(n.line));
      if (old) ln.replaceChild(fill, old);
      else ln.insertBefore(fill, ln.firstChild);
      for (const [tag, e] of [['headEnd', n.line.head], ['tailEnd', n.line.tail]] as const) {
        const cur = child(ln, tag, A);
        if (cur) ln.removeChild(cur);
        if (e) ln.appendChild(el(ln.ownerDocument, endXml(tag, e)));
      }
      const dash = child(ln, 'prstDash', A);
      if (!same(n.line.dash, was.line.dash)) {
        if (dash) ln.removeChild(dash);
        if (n.line.dash) ln.insertBefore(el(ln.ownerDocument, `<a:prstDash val="${escapeAttr(n.line.dash)}"/>`), child(ln, 'headEnd', A) ?? child(ln, 'tailEnd', A));
      }
    } else {
      const xml = lineXml(n.line);
      if (ln) replaceWith(ln, xml);
      else {
        const after = Array.from(spPr.children).filter((c) => c.namespaceURI === A && (FILLS.includes(c.localName) || c.localName === 'prstGeom' || c.localName === 'custGeom')).pop();
        spPr.insertBefore(el(spPr.ownerDocument, xml), after ? after.nextSibling : null);
      }
    }
  }
  if (n.cxn && (!same(n.st, was.st) || !same(n.end, was.end))) {
    const cnv = child(sp, 'cNvCnPr', SHAPE_NS.wps);
    if (cnv) {
      for (const tag of ['stCxn', 'endCxn']) {
        const cur = child(cnv, tag, A);
        if (cur) cnv.removeChild(cur);
      }
      const locks = child(cnv, 'cxnSpLocks', A);
      const xml = `${n.st ? `<a:stCxn id="${int(n.st.id)}" idx="${int(n.st.idx)}"/>` : ''}${n.end ? `<a:endCxn id="${int(n.end.id)}" idx="${int(n.end.idx)}"/>` : ''}`;
      if (xml) {
        const holder = parseFragment(`<wps:cNvCnPr>${xml}</wps:cNvCnPr>`);
        for (const c of Array.from(holder.children)) cnv.insertBefore(cnv.ownerDocument.importNode(c, true), locks ? locks.nextSibling : null);
      }
    }
  }
}

/** Whether two trees have the same shapes in the same order (only their properties may differ). */
function sameStructure(a: ShapeNode, b: ShapeNode): boolean {
  if (a.t !== b.t) return false;
  if (a.t === 'grp' && b.t === 'grp') return a.kids.length === b.kids.length && a.kids.every((k, i) => sameStructure(k, b.kids[i]));
  return true;
}

function patchKids(parent: Element, now: GrpNode, was: GrpNode): void {
  const els = kidElements(parent);
  now.kids.forEach((k, i) => {
    const w = was.kids[i];
    const e = els[i];
    if (!e || same(k, w)) return;
    if (k.t === 'sp' && w.t === 'sp') patchSp(e, k, w, k, w);
    else if (k.t === 'grp' && w.t === 'grp') {
      const grpSpPr = child(e, 'grpSpPr', SHAPE_NS.wpg);
      if (grpSpPr) patchXfrm(grpSpPr, k, w, k, w);
      patchKids(e, k, w);
    } else if (k.t === 'pic' && w.t === 'pic') {
      const spPr = child(e, 'spPr', NS.pic);
      if (spPr) patchXfrm(spPr, k, w);
    }
  });
}

function patchDrawing(drawing: Element, now: ShapeModel, was: ShapeModel): void {
  let frame = child(drawing, 'anchor', WP) ?? child(drawing, 'inline', WP);
  if (!frame) return;
  if (!!now.anchor !== !!was.anchor) {
    // In the line ↔ floating: a new frame around the same docPr and graphic.
    const keep = ['docPr', 'cNvGraphicFramePr', 'graphic'].map((l) => child(frame!, l, WP) ?? child(frame!, l, A)).filter(Boolean) as Element[];
    const tail = Array.from(frame.children).filter((c) => c.namespaceURI !== WP || !['simplePos', 'positionH', 'positionV', 'extent', 'effectExtent', 'wrapNone', 'wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'docPr', 'cNvGraphicFramePr'].includes(c.localName)).filter((c) => !keep.includes(c));
    const next = el(frame.ownerDocument, `${frameHead(now)}</wp:${now.anchor ? 'anchor' : 'inline'}>`);
    for (const c of [...keep, ...(now.anchor ? tail : [])]) next.appendChild(c);
    frame.parentNode!.replaceChild(next, frame);
    frame = next;
  } else if (now.anchor && was.anchor) {
    const a = now.anchor;
    const b = was.anchor;
    if (a.behind !== b.behind) frame.setAttribute('behindDoc', a.behind ? '1' : '0');
    if (a.z !== b.z) frame.setAttribute('relativeHeight', int(a.z));
    if (!same(a.dist, b.dist)) ['distT', 'distB', 'distL', 'distR'].forEach((k, i) => frame!.setAttribute(k, int(a.dist[i])));
    if (!same(a.h, b.h) || !same(a.v, b.v)) {
      frame.setAttribute('simplePos', '0');
      for (const [tag, axis] of [['positionH', a.h], ['positionV', a.v]] as const) {
        const cur = child(frame, tag, WP);
        const xml = axisXml(tag, axis);
        if (cur) replaceWith(cur, xml);
        else frame.insertBefore(el(frame.ownerDocument, xml), child(frame, 'extent', WP));
      }
    }
    if (a.wrap !== b.wrap || a.side !== b.side) {
      const wraps = Array.from(frame.children).filter((c) => c.namespaceURI === WP && c.localName.startsWith('wrap'));
      const at = wraps.length ? wraps[wraps.length - 1].nextSibling : child(frame, 'docPr', WP);
      for (const w of wraps) frame.removeChild(w);
      frame.insertBefore(el(frame.ownerDocument, wrapXml(a)), at);
    }
  }
  if (now.w !== was.w || now.h !== was.h) {
    const ext = child(frame, 'extent', WP);
    ext?.setAttribute('cx', int(now.w));
    ext?.setAttribute('cy', int(now.h));
  }
  const data = frame.getElementsByTagNameNS(A, 'graphicData')[0];
  if (!data) return;
  const r = now.root;
  const w = was.root;
  if (r.t === 'sp' && w.t === 'sp') {
    const wsp = child(data, 'wsp', SHAPE_NS.wps);
    if (wsp) patchSp(wsp, r, w, { ...r, x: 0, y: 0, w: now.w, h: now.h }, { ...w, x: 0, y: 0, w: was.w, h: was.h });
  } else if (r.t === 'grp' && w.t === 'grp') {
    const wgp = child(data, 'wgp', SHAPE_NS.wpg);
    const wpc = child(data, 'wpc', SHAPE_NS.wpc);
    if (wgp) {
      const grpSpPr = child(wgp, 'grpSpPr', SHAPE_NS.wpg);
      if (grpSpPr) patchXfrm(grpSpPr, { ...r, x: 0, y: 0, w: now.w, h: now.h }, { ...w, x: 0, y: 0, w: was.w, h: was.h }, r, w);
      patchKids(wgp, r, w);
    } else if (wpc) {
      if (!same(r.bg, w.bg)) {
        const bg = child(wpc, 'bg', SHAPE_NS.wpc);
        if (bg) replaceWith(bg, `<wpc:bg>${fillXml(r.bg ?? null)}</wpc:bg>`);
      }
      patchKids(wpc, r, w);
    }
  }
}

/** Patch the top VML element of a single shape (or a group whose inside is unchanged). */
function patchVml(pict: Element, now: ShapeModel, was: ShapeModel): boolean {
  const top = Array.from(pict.children).find((c) => c.namespaceURI === SHAPE_NS.v && c.localName !== 'shapetype');
  if (!top) return false;
  const style = new Map<string, string>();
  for (const part of (top.getAttribute('style') ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) style.set(part.slice(0, i).trim(), part.slice(i + 1).trim());
  }
  const set = (k: string, v: string | null) => (v == null ? style.delete(k) : style.set(k, v));
  const a = now.anchor;
  if (!!a !== !!was.anchor) set('position', a ? 'absolute' : null);
  if (a) {
    if (!same(a.h, was.anchor?.h)) {
      set('margin-left', pt(a.h.offset ?? 0));
      set('mso-position-horizontal', a.h.align ?? 'absolute');
      set('mso-position-horizontal-relative', H_VML[a.h.rel] ?? 'text');
      style.delete('left');
    }
    if (!same(a.v, was.anchor?.v)) {
      set('margin-top', pt(a.v.offset ?? 0));
      set('mso-position-vertical', a.v.align ?? 'absolute');
      set('mso-position-vertical-relative', V_VML[a.v.rel] ?? 'text');
      style.delete('top');
    }
    if (a.z !== was.anchor?.z || a.behind !== was.anchor?.behind) set('z-index', String(a.behind ? -Math.max(1, Math.round(a.z)) : Math.max(1, Math.round(a.z))));
  } else {
    for (const k of ['margin-left', 'margin-top', 'left', 'top', 'z-index', 'mso-position-horizontal', 'mso-position-horizontal-relative', 'mso-position-vertical', 'mso-position-vertical-relative']) style.delete(k);
  }
  if (top.localName === 'line') {
    if (now.w !== was.w || now.h !== was.h || !same(now.root, was.root)) {
      const r = now.root as SpNode;
      const from = `${r.flipH ? pt(now.w) : '0'},${r.flipV ? pt(now.h) : '0'}`;
      const to = `${r.flipH ? '0' : pt(now.w)},${r.flipV ? '0' : pt(now.h)}`;
      top.setAttribute('from', from);
      top.setAttribute('to', to);
    }
  } else if (now.w !== was.w || now.h !== was.h) {
    set('width', pt(now.w));
    set('height', pt(now.h));
  }
  const r = now.root;
  const w = was.root;
  if (r.t === 'sp' && w.t === 'sp') {
    if (r.rot !== w.rot) set('rotation', r.rot ? String(Math.round(r.rot * 100) / 100) : null);
    if (r.flipH !== w.flipH || r.flipV !== w.flipV) set('flip', r.flipH || r.flipV ? [r.flipH ? 'x' : '', r.flipV ? 'y' : ''].filter(Boolean).join(' ') : null);
    if (!same(r.fill, w.fill) && !isLinePreset(r.geom)) {
      if (r.fill) {
        top.setAttribute('fillcolor', `#${r.fill.color}`);
        top.removeAttribute('filled');
      } else top.setAttribute('filled', 'f');
      const f = child(top, 'fill', SHAPE_NS.v);
      if (f && (f.getAttribute('on') === 'f' || f.getAttribute('on') === 'false')) f.removeAttribute('on');
    }
    if (!same(r.line, w.line)) {
      if (r.line) {
        top.setAttribute('strokecolor', `#${r.line.color}`);
        top.setAttribute('strokeweight', pt(r.line.width));
        top.removeAttribute('stroked');
        const s = child(top, 'stroke', SHAPE_NS.v);
        if (s) {
          s.removeAttribute('color');
          s.removeAttribute('weight');
        }
      } else top.setAttribute('stroked', 'f');
    }
  }
  top.setAttribute('style', [...style].map(([k, v]) => `${k}:${v}`).join(';'));
  if (!same(a?.wrap, was.anchor?.wrap)) {
    const cur = child(top, 'wrap', SHAPE_NS.w10);
    if (cur) top.removeChild(cur);
    const xml = vmlWrap(now);
    if (xml) top.appendChild(el(top.ownerDocument, xml));
  }
  return true;
}

/**
 * A shape's run with the changes from `was` to `now` (non-text: position, size, wrapping,
 * stacking, colours, children's boxes, connections) written into both of its copies. Text boxes'
 * text is not touched here (rewriteTextBoxes). Throws when the shapes inside differ in number or
 * kind (a structural change makes its XML anew instead).
 */
export function patchShapeXml(xml: string, was: ShapeModel, now: ShapeModel): string {
  if (!sameStructure(was.root, now.root)) throw new Error('the shapes inside changed');
  const run = parseFragment(xml);
  const holder = Array.from(run.children).find((c) => shapeSource(c));
  const src = holder ? shapeSource(holder) : null;
  if (!src) return xml;
  // The text boxes' content as it is in the file, for a VML copy written again.
  const boxes = ownTextBoxes(src.primary).map((b) => Array.from(b.childNodes).map((c) => serializeXml(c)).join(''));
  const text = (f: TextFrame) => boxes[f.i] ?? '<w:p/>';
  if (src.kind === 'dml') patchDrawing(src.primary, now, was);
  const pict = src.kind === 'vml' ? src.primary : src.fallback ? child(src.fallback, 'pict') : null;
  if (pict) {
    const inside = was.root.t === 'grp' && now.root.t === 'grp' && (!same(was.root.kids, now.root.kids) || (now.kind === 'canvas' && (now.w !== was.w || now.h !== was.h)));
    if (inside || !patchVml(pict, now, was)) replaceWith(pict, vmlXml(now, text));
  }
  return serializeXml(run);
}

export { textFrames };
