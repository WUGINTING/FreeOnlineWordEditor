// Lossless property handling.
//
// Every paragraph / run / table cell keeps its original OOXML property element
// (w:pPr, w:rPr, w:tcPr) verbatim. The editor only models a subset of those
// properties. On save we re-read the modeled subset from the original element,
// compare it with what the editor has now, and patch only the properties that
// actually changed. Untouched elements are written back byte-for-byte.

import { NS, WORD_NAMESPACES, attr, child, find, numAttr, onOff, parseFragment, serializeXml } from './xml';
import { halfPointsToPt, ptToHalfPoints } from '../units';
import { Lru, deepFreeze, memo, memoKey } from './memo';
import { clampLevel } from './numbering';

// Child order defined by ECMA-376 (CT_PPrBase + CT_PPr, CT_RPr, CT_TcPr).
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
  'sectPr', 'pPrChange',
];
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
  'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
];
const TCPR_ORDER = [
  'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
  'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge',
  'tcPrChange',
];
// CT_TcBorders; start/end are the bidi-aware names of left/right.
const TCBORDERS_ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
// CT_TrPrBase is a choice group; this is the order Word writes, then CT_TrPr's tail.
const TRPR_ORDER = [
  'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
  'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange',
];

/** Sort w: children into schema order; non-w children (w14:...) keep their place at the end. */
function sortChildren(el: Element, order: string[]): void {
  const kids = Array.from(el.children);
  const rank = (c: Element) => {
    const i = c.namespaceURI === NS.w ? order.indexOf(c.localName) : -1;
    return i < 0 ? order.length : i;
  };
  const sorted = kids.map((c, i) => ({ c, i })).sort((a, b) => rank(a.c) - rank(b.c) || a.i - b.i);
  for (const { c } of sorted) el.appendChild(c);
}

function ensure(el: Element, local: string): Element {
  let c = child(el, local);
  if (!c) {
    c = el.ownerDocument.createElementNS(NS.w, 'w:' + local);
    el.appendChild(c);
  }
  return c;
}

function remove(el: Element, ...locals: string[]): void {
  for (const l of locals) {
    const c = child(el, l);
    if (c) el.removeChild(c);
  }
}

function setAttr(el: Element, local: string, value: string | number | null): void {
  if (value == null) el.removeAttributeNS(NS.w, local);
  else el.setAttributeNS(NS.w, 'w:' + local, String(value));
}

function removeAttrs(el: Element, ...locals: string[]): void {
  for (const l of locals) el.removeAttributeNS(NS.w, l);
}

function toElement(xml: string | null, local: string): Element {
  if (xml) return parseFragment(xml);
  return parseFragment(`<w:${local}/>`);
}

const same = (a: unknown, b: unknown) => (a ?? null) === (b ?? null);

// ---------------------------------------------------------------------------
// Paragraph properties

export interface ParaModel {
  styleId: string | null;
  align: string | null;
  indLeft: number | null;
  indRight: number | null;
  indFirst: number | null;
  indLeftChars: number | null;
  indRightChars: number | null;
  indFirstChars: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  line: number | null;
  lineRule: string | null;
  numId: string | null;
  ilvl: number;
  pageBreakBefore: boolean;
}

export function readParaModel(pPr: Element | null): ParaModel {
  const m: ParaModel = {
    styleId: null, align: null, indLeft: null, indRight: null, indFirst: null,
    indLeftChars: null, indRightChars: null, indFirstChars: null,
    spaceBefore: null, spaceAfter: null, line: null, lineRule: null,
    numId: null, ilvl: 0, pageBreakBefore: false,
  };
  if (!pPr) return m;
  m.styleId = attr(child(pPr, 'pStyle'), 'val');
  const jc = attr(child(pPr, 'jc'), 'val');
  if (jc) m.align = jc === 'both' || jc === 'distribute' ? 'justify' : jc === 'start' ? 'left' : jc === 'end' ? 'right' : jc;
  const ind = child(pPr, 'ind');
  if (ind) {
    m.indLeft = numAttr(ind, 'left') ?? numAttr(ind, 'start');
    m.indRight = numAttr(ind, 'right') ?? numAttr(ind, 'end');
    const hanging = numAttr(ind, 'hanging');
    m.indFirst = hanging != null ? -hanging : numAttr(ind, 'firstLine');
    m.indLeftChars = numAttr(ind, 'leftChars') ?? numAttr(ind, 'startChars');
    m.indRightChars = numAttr(ind, 'rightChars') ?? numAttr(ind, 'endChars');
    const hangingChars = numAttr(ind, 'hangingChars');
    m.indFirstChars = hangingChars != null ? -hangingChars : numAttr(ind, 'firstLineChars');
  }
  const sp = child(pPr, 'spacing');
  if (sp) {
    m.spaceBefore = numAttr(sp, 'before');
    m.spaceAfter = numAttr(sp, 'after');
    m.line = numAttr(sp, 'line');
    m.lineRule = attr(sp, 'lineRule');
  }
  const numPr = child(pPr, 'numPr');
  if (numPr) {
    const numId = attr(child(numPr, 'numId'), 'val');
    if (numId != null) {
      m.numId = numId;
      // Word has levels 0..8; a larger value would make list counting walk a huge array.
      m.ilvl = clampLevel(numAttr(child(numPr, 'ilvl'), 'val') ?? 0);
    }
  }
  m.pageBreakBefore = onOff(child(pPr, 'pageBreakBefore')) === true;
  return m;
}

/** Results of the property writers, by their inputs (see memo.ts). */
const written = new Lru<string>(4000);

/** Original w:pPr (or null) patched so that it expresses `want`. */
export function writePPr(raw: string | null, want: ParaModel, sectPr: string | null): string {
  return memo(written, memoKey('p', raw, want, sectPr), () => patchPPr(raw, want, sectPr));
}

function patchPPr(raw: string | null, want: ParaModel, sectPr: string | null): string {
  const el = toElement(raw, 'pPr');
  const have = readParaModel(raw ? el : null);
  let changed = false;

  if (!same(have.styleId, want.styleId)) {
    changed = true;
    if (want.styleId) setAttr(ensure(el, 'pStyle'), 'val', want.styleId);
    else remove(el, 'pStyle');
  }
  if (!same(have.align, want.align)) {
    changed = true;
    if (want.align) setAttr(ensure(el, 'jc'), 'val', want.align === 'justify' ? 'both' : want.align);
    else remove(el, 'jc');
  }
  const indKeys = ['indLeft', 'indRight', 'indFirst', 'indLeftChars', 'indRightChars', 'indFirstChars'] as const;
  if (indKeys.some((k) => !same(have[k], want[k]))) {
    changed = true;
    const ind = ensure(el, 'ind');
    const put = (twips: string, alt: string, value: number | null, prev: number | null) => {
      if (same(value, prev)) return;
      removeAttrs(ind, alt);
      setAttr(ind, twips, value);
    };
    put('left', 'start', want.indLeft, have.indLeft);
    put('right', 'end', want.indRight, have.indRight);
    put('leftChars', 'startChars', want.indLeftChars, have.indLeftChars);
    put('rightChars', 'endChars', want.indRightChars, have.indRightChars);
    if (!same(want.indFirst, have.indFirst)) {
      removeAttrs(ind, 'firstLine', 'hanging');
      if (want.indFirst != null) setAttr(ind, want.indFirst < 0 ? 'hanging' : 'firstLine', Math.abs(want.indFirst));
    }
    if (!same(want.indFirstChars, have.indFirstChars)) {
      removeAttrs(ind, 'firstLineChars', 'hangingChars');
      if (want.indFirstChars != null) setAttr(ind, want.indFirstChars < 0 ? 'hangingChars' : 'firstLineChars', Math.abs(want.indFirstChars));
    }
    if (!ind.attributes.length) el.removeChild(ind);
  }
  const spKeys = ['spaceBefore', 'spaceAfter', 'line', 'lineRule'] as const;
  if (spKeys.some((k) => !same(have[k], want[k]))) {
    changed = true;
    const sp = ensure(el, 'spacing');
    if (!same(have.spaceBefore, want.spaceBefore)) {
      // An explicit value replaces "lines" / auto spacing, as in Word's dialog.
      removeAttrs(sp, 'beforeLines', 'beforeAutospacing');
      setAttr(sp, 'before', want.spaceBefore);
    }
    if (!same(have.spaceAfter, want.spaceAfter)) {
      removeAttrs(sp, 'afterLines', 'afterAutospacing');
      setAttr(sp, 'after', want.spaceAfter);
    }
    if (!same(have.line, want.line)) setAttr(sp, 'line', want.line);
    if (!same(have.lineRule, want.lineRule)) setAttr(sp, 'lineRule', want.lineRule);
    if (!sp.attributes.length) el.removeChild(sp);
  }
  if (!same(have.numId, want.numId) || (want.numId && have.ilvl !== want.ilvl)) {
    changed = true;
    if (want.numId) {
      const numPr = ensure(el, 'numPr');
      remove(numPr, 'ilvl', 'numId');
      const ilvl = el.ownerDocument.createElementNS(NS.w, 'w:ilvl');
      setAttr(ilvl, 'val', want.ilvl);
      const numId = el.ownerDocument.createElementNS(NS.w, 'w:numId');
      setAttr(numId, 'val', want.numId);
      numPr.insertBefore(numId, numPr.firstChild);
      numPr.insertBefore(ilvl, numId);
    } else {
      remove(el, 'numPr');
    }
  }
  if (have.pageBreakBefore !== want.pageBreakBefore) {
    changed = true;
    remove(el, 'pageBreakBefore');
    if (want.pageBreakBefore) ensure(el, 'pageBreakBefore');
  }

  // The section break lives inside pPr but is stored separately.
  const oldSect = child(el, 'sectPr');
  const oldSectXml = oldSect ? serializeXml(oldSect) : null;
  if (oldSectXml !== sectPr) {
    changed = true;
    if (oldSect) el.removeChild(oldSect);
    if (sectPr) el.appendChild(parseFragment(sectPr));
  }

  if (!changed && raw) return raw;
  if (changed) sortChildren(el, PPR_ORDER);
  return el.children.length || el.attributes.length ? serializeXml(el) : '';
}

// ---------------------------------------------------------------------------
// Run properties

export interface RunModel {
  charStyle: string | null;
  /** true / false when the run says so (w:b, w:b w:val="0"); null to inherit from styles. */
  bold: boolean | null;
  italic: boolean | null;
  underline: boolean;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
  color: string | null; // "#RRGGBB"
  highlight: string | null; // "#rrggbb" (w:highlight) or fill color (w:shd)
  fontFamily: string | null;
  fontEastAsia: string | null;
  fontSize: number | null; // pt
}

export const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

const HEX6 = /^[0-9A-Fa-f]{6}$/;

/** A w:highlight value as "#rrggbb": a named colour (not an inherited property name such as "constructor") or a hex colour. */
function highlightColor(val: string): string | null {
  if (Object.prototype.hasOwnProperty.call(HIGHLIGHT_COLORS, val)) return HIGHLIGHT_COLORS[val];
  const h = /^#?([0-9A-Fa-f]{6})$/.exec(val);
  return h ? '#' + h[1].toLowerCase() : null;
}

export function readRunModel(rPr: Element | null): RunModel {
  const m: RunModel = {
    charStyle: null, bold: null, italic: null, underline: false, strike: false,
    superscript: false, subscript: false, color: null, highlight: null,
    fontFamily: null, fontEastAsia: null, fontSize: null,
  };
  if (!rPr) return m;
  m.charStyle = attr(child(rPr, 'rStyle'), 'val');
  m.bold = onOff(child(rPr, 'b'));
  m.italic = onOff(child(rPr, 'i'));
  const u = child(rPr, 'u');
  m.underline = !!u && attr(u, 'val') !== 'none';
  m.strike = onOff(child(rPr, 'strike')) === true || onOff(child(rPr, 'dstrike')) === true;
  const va = attr(child(rPr, 'vertAlign'), 'val');
  m.superscript = va === 'superscript';
  m.subscript = va === 'subscript';
  // Only well-formed values are modeled (they end up in CSS); the element itself stays as written.
  const color = attr(child(rPr, 'color'), 'val');
  if (color && HEX6.test(color)) m.color = '#' + color;
  const hl = attr(child(rPr, 'highlight'), 'val');
  if (hl && hl !== 'none') m.highlight = highlightColor(hl);
  const fonts = child(rPr, 'rFonts');
  m.fontFamily = attr(fonts, 'ascii') ?? attr(fonts, 'hAnsi');
  m.fontEastAsia = attr(fonts, 'eastAsia');
  const sz = numAttr(child(rPr, 'sz'), 'val');
  if (sz != null) m.fontSize = halfPointsToPt(sz);
  return m;
}

/** "#rrggbb" / "rgb(r, g, b)" -> "RRGGBB". */
export function hex(color: string): string {
  const c = color.trim();
  if (c.startsWith('#')) {
    const h = c.slice(1);
    return (h.length === 3 ? h.split('').map((x) => x + x).join('') : h).toUpperCase();
  }
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase();
  return 'auto';
}

/** Original w:rPr (or null) patched so that it expresses `want`. */
export function writeRPr(raw: string | null, want: RunModel): string {
  return memo(written, memoKey('r', raw, want), () => patchRPr(raw, want));
}

function patchRPr(raw: string | null, want: RunModel): string {
  const el = toElement(raw, 'rPr');
  const have = readRunModel(raw ? el : null);
  let changed = false;
  const diff = <K extends keyof RunModel>(k: K) => !same(have[k], want[k]);

  if (diff('charStyle')) {
    changed = true;
    if (want.charStyle) setAttr(ensure(el, 'rStyle'), 'val', want.charStyle);
    else remove(el, 'rStyle');
  }
  if (diff('fontFamily') || diff('fontEastAsia')) {
    changed = true;
    const f = ensure(el, 'rFonts');
    if (diff('fontFamily')) {
      removeAttrs(f, 'asciiTheme', 'hAnsiTheme');
      setAttr(f, 'ascii', want.fontFamily);
      setAttr(f, 'hAnsi', want.fontFamily);
    }
    if (diff('fontEastAsia')) {
      removeAttrs(f, 'eastAsiaTheme');
      setAttr(f, 'eastAsia', want.fontEastAsia);
    }
    if (!f.attributes.length) el.removeChild(f);
  }
  const boolProp = (k: 'bold' | 'italic', local: string, cs: string) => {
    if (!diff(k)) return;
    changed = true;
    // Keep the complex-script twin in step only if the document already used it.
    const pair = child(el, cs) || !raw ? cs : undefined;
    remove(el, local);
    if (pair) remove(el, pair);
    const value = want[k];
    if (value == null) return; // inherit from the styles
    for (const name of pair ? [local, pair] : [local]) {
      const e = ensure(el, name);
      if (!value) setAttr(e, 'val', '0'); // explicitly off, e.g. inside a bold style
    }
  };
  boolProp('bold', 'b', 'bCs');
  boolProp('italic', 'i', 'iCs');
  if (diff('strike')) {
    changed = true;
    remove(el, 'strike', 'dstrike');
    if (want.strike) ensure(el, 'strike');
  }
  if (diff('color')) {
    changed = true;
    const c = ensure(el, 'color');
    removeAttrs(c, 'themeColor', 'themeTint', 'themeShade');
    if (want.color) setAttr(c, 'val', hex(want.color));
    else el.removeChild(c);
  }
  if (diff('fontSize')) {
    changed = true;
    if (want.fontSize != null) {
      setAttr(ensure(el, 'sz'), 'val', ptToHalfPoints(want.fontSize));
      setAttr(ensure(el, 'szCs'), 'val', ptToHalfPoints(want.fontSize));
    } else remove(el, 'sz', 'szCs');
  }
  if (diff('highlight')) {
    changed = true;
    remove(el, 'highlight');
    const name = want.highlight
      ? Object.entries(HIGHLIGHT_COLORS).find(([, v]) => v === '#' + hex(want.highlight!).toLowerCase())?.[0]
      : undefined;
    if (name) setAttr(ensure(el, 'highlight'), 'val', name);
    else if (want.highlight) {
      const shd = ensure(el, 'shd');
      setAttr(shd, 'val', 'clear');
      setAttr(shd, 'color', 'auto');
      setAttr(shd, 'fill', hex(want.highlight));
    }
  }
  if (diff('underline')) {
    changed = true;
    remove(el, 'u');
    if (want.underline) setAttr(ensure(el, 'u'), 'val', 'single');
  }
  if (diff('superscript') || diff('subscript')) {
    changed = true;
    remove(el, 'vertAlign');
    if (want.superscript) setAttr(ensure(el, 'vertAlign'), 'val', 'superscript');
    else if (want.subscript) setAttr(ensure(el, 'vertAlign'), 'val', 'subscript');
  }

  if (!changed && raw) return raw;
  if (changed) sortChildren(el, RPR_ORDER);
  return el.children.length ? serializeXml(el) : '';
}

/**
 * The w:rPr (or none) with w:rFonts/@w:hint="eastAsia", which Word writes on a run of symbols
 * entered in Chinese text: characters that both Latin and CJK fonts have (①, ※, ○, ±, ×, →,
 * ℃ …) are then drawn in the run's East Asian font (標楷體), not its Latin one (Times New
 * Roman). Everything else in the w:rPr stays; an rPr that already has the hint is returned as-is.
 */
export function withEastAsiaHint(raw: string | null): string {
  const el = toElement(raw, 'rPr');
  const fonts = ensure(el, 'rFonts');
  if (raw && fonts.getAttributeNS(NS.w, 'hint') === 'eastAsia') return raw;
  setAttr(fonts, 'hint', 'eastAsia');
  sortChildren(el, RPR_ORDER);
  return serializeXml(el);
}

const EAST_ASIA_HINT = /<w:rFonts\b[^>]*\sw:hint="eastAsia"/;

/** Whether a raw w:rPr carries w:rFonts/@w:hint="eastAsia" (see withEastAsiaHint). */
export function hasEastAsiaHint(raw: string | null | undefined): boolean {
  return !!raw && EAST_ASIA_HINT.test(raw);
}

/**
 * Run properties that only change how text looks. "Clear formatting" removes these and
 * keeps the rest: language, proofing, hidden text, complex-script / right-to-left flags,
 * and tracked formatting changes (w:rPrChange).
 */
const RUN_APPEARANCE = new Set([
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
  'shadow', 'emboss', 'imprint', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
  'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'em',
]);
/** Word 2010+ text effects (w14:textFill, w14:glow ...) are appearance too. */
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** The w:rPr without its appearance properties; null when nothing is left. Unchanged input is returned as-is. */
export function stripRunAppearance(raw: string): string | null {
  const el = parseFragment(raw);
  const drop = Array.from(el.children).filter(
    (c) => (c.namespaceURI === NS.w && RUN_APPEARANCE.has(c.localName)) || c.namespaceURI === W14,
  );
  if (!drop.length) return raw;
  for (const c of drop) el.removeChild(c);
  return el.children.length ? serializeXml(el) : null;
}

// ---------------------------------------------------------------------------
// Table cell properties

export type BorderSide = 'top' | 'left' | 'bottom' | 'right';
export const BORDER_SIDES: BorderSide[] = ['top', 'left', 'bottom', 'right'];
/** One w:tcBorders side. val "nil" / "none" = no border; sz in eighths of a point; color "RRGGBB" or "auto". */
export interface Border {
  val: string;
  sz: number | null;
  color: string | null;
}
/** Sides the cell sets itself; a missing side comes from the table. */
export type CellBorders = Partial<Record<BorderSide, Border>>;

export interface CellModel {
  width: number | null; // twips (w:tcW type=dxa)
  background: string | null;
  vAlign: string | null; // top | middle | bottom
  borders: CellBorders | null;
}

/** The w:tcBorders element of a side (left/right may be written as start/end). */
function borderEl(tcBorders: Element | null, side: BorderSide): Element | null {
  const alt = side === 'left' ? 'start' : side === 'right' ? 'end' : null;
  return child(tcBorders, side) ?? (alt ? child(tcBorders, alt) : null);
}

export function readBorders(tcBorders: Element | null): CellBorders | null {
  if (!tcBorders) return null;
  const out: CellBorders = {};
  for (const side of BORDER_SIDES) {
    const b = borderEl(tcBorders, side);
    if (!b) continue;
    // Only well-formed values are modeled (they end up in CSS); the element itself stays as written.
    const val = attr(b, 'val') ?? 'nil';
    const color = attr(b, 'color');
    out[side] = { val: /^[A-Za-z]+$/.test(val) ? val : 'single', sz: numAttr(b, 'sz'), color: color && isColorValue(color) ? color : null };
  }
  return Object.keys(out).length ? out : null;
}

/** An ST_HexColor value: six hex digits or "auto". */
const isColorValue = (v: string): boolean => v === 'auto' || /^[0-9A-Fa-f]{6}$/.test(v);

/** ST_VerticalJc values. */
const V_ALIGN = new Set(['top', 'center', 'bottom', 'both']);

export const sameBorder = (a: Border | undefined | null, b: Border | undefined | null): boolean =>
  (!a && !b) || (!!a && !!b && a.val === b.val && (a.sz ?? null) === (b.sz ?? null) && (a.color ?? null) === (b.color ?? null));

/** Most columns a Word table can have: the cap for w:gridSpan / w:gridBefore / w:gridAfter. */
export const MAX_GRID_COLUMNS = 63;

/**
 * A cell's w:gridSpan as used for layout: a whole number from 1 to 63; anything else (0,
 * negative, fractional, missing) is 1. A span that doesn't move the column on would stall
 * the table code. The element itself stays as written.
 */
export function gridSpanOf(tcPr: Element | null): number {
  const v = numAttr(child(tcPr, 'gridSpan'), 'val');
  return v != null && Number.isInteger(v) && v >= 1 ? Math.min(v, MAX_GRID_COLUMNS) : 1;
}

/** A row's w:gridBefore / w:gridAfter as used for layout: a whole number from 0 to 63. */
export function gridSkipOf(trPr: Element | null, local: 'gridBefore' | 'gridAfter'): number {
  const v = numAttr(child(trPr, local), 'val');
  return v != null && Number.isInteger(v) && v >= 0 ? Math.min(v, MAX_GRID_COLUMNS) : 0;
}

const cellInfos = new Lru<{ span: number; model: CellModel }>(2000);
const EMPTY_CELL = deepFreeze({ span: 1, model: { width: null, background: null, vAlign: null, borders: null } as CellModel });

/** gridSpanOf and readCellModel of a kept w:tcPr (cached, frozen: don't change it). */
export function cellInfo(tcPr: string | null): { readonly span: number; readonly model: CellModel } {
  if (!tcPr) return EMPTY_CELL;
  return memo(cellInfos, tcPr, () => {
    const el = parseFragment(tcPr);
    return deepFreeze({ span: gridSpanOf(el), model: readCellModel(el) });
  });
}

export function readCellModel(tcPr: Element | null): CellModel {
  const m: CellModel = { width: null, background: null, vAlign: null, borders: null };
  if (!tcPr) return m;
  const w = child(tcPr, 'tcW');
  if (w && (attr(w, 'type') ?? 'dxa') === 'dxa') m.width = numAttr(w, 'w');
  const fill = attr(child(tcPr, 'shd'), 'fill');
  if (fill && fill !== 'auto' && isColorValue(fill)) m.background = '#' + fill;
  const va = attr(child(tcPr, 'vAlign'), 'val');
  if (va && V_ALIGN.has(va)) m.vAlign = va === 'center' ? 'middle' : va;
  m.borders = readBorders(child(tcPr, 'tcBorders'));
  return m;
}

/** Patch the sides of w:tcBorders that differ between `have` and `want`; other sides stay as written. */
function patchBorders(tcPr: Element, have: CellBorders | null, want: CellBorders | null): void {
  let box = child(tcPr, 'tcBorders');
  for (const side of BORDER_SIDES) {
    const next = want?.[side];
    if (sameBorder(have?.[side], next)) continue;
    const old = borderEl(box, side);
    if (!next) {
      if (old) box!.removeChild(old);
      continue;
    }
    box ??= ensure(tcPr, 'tcBorders');
    const b = old ?? box.appendChild(tcPr.ownerDocument.createElementNS(NS.w, 'w:' + side));
    setAttr(b, 'val', next.val);
    if (next.val === 'nil' || next.val === 'none') {
      removeAttrs(b, 'sz', 'space', 'color', 'themeColor', 'themeTint', 'themeShade', 'shadow', 'frame');
      continue;
    }
    setAttr(b, 'sz', next.sz);
    if (!attr(b, 'space')) setAttr(b, 'space', 0);
    if ((next.color ?? null) !== attr(b, 'color')) {
      removeAttrs(b, 'themeColor', 'themeTint', 'themeShade');
      setAttr(b, 'color', next.color);
    }
  }
  if (!box) return;
  if (!box.children.length) tcPr.removeChild(box);
  else sortChildren(box, TCBORDERS_ORDER);
}

/**
 * Original w:tcPr patched for the cell's current layout (span / vertical merge,
 * always recomputed) and any changed modeled properties.
 */
export function writeTcPr(
  raw: string | null,
  want: CellModel,
  layout: { gridSpan: number; vMerge: 'restart' | 'continue' | null },
): string {
  return memo(written, memoKey('c', raw, want, layout), () => patchTcPr(raw, want, layout));
}

function patchTcPr(
  raw: string | null,
  want: CellModel,
  layout: { gridSpan: number; vMerge: 'restart' | 'continue' | null },
): string {
  const el = toElement(raw, 'tcPr');
  // The original as serializeXml writes it, to tell whether anything changed.
  const before = raw ? serializeXml(el) : null;
  const have = readCellModel(raw ? el : null);

  if ((!same(have.width, want.width) || !raw) && want.width != null) {
    const w = ensure(el, 'tcW');
    setAttr(w, 'w', want.width);
    setAttr(w, 'type', 'dxa');
  }
  if (!same(have.background, want.background)) {
    if (want.background) {
      const shd = ensure(el, 'shd');
      removeAttrs(shd, 'themeFill', 'themeFillTint', 'themeFillShade');
      // A plain fill, as Word's shading button writes it (a pattern would hide the new color).
      setAttr(shd, 'val', 'clear');
      setAttr(shd, 'color', 'auto');
      setAttr(shd, 'fill', hex(want.background));
    } else remove(el, 'shd');
  }
  if (!same(have.vAlign, want.vAlign)) {
    if (want.vAlign) setAttr(ensure(el, 'vAlign'), 'val', want.vAlign === 'middle' ? 'center' : want.vAlign);
    else remove(el, 'vAlign');
  }
  patchBorders(el, have.borders, want.borders ?? null);

  // Layout: only touch it when it differs from what the original says.
  const span = gridSpanOf(el);
  const vm = child(el, 'vMerge');
  const haveVMerge = vm ? (attr(vm, 'val') === 'restart' ? 'restart' : 'continue') : null;
  // Legacy w:hMerge (read as separate cells) stays while the cell's layout is unchanged.
  if (span !== layout.gridSpan || haveVMerge !== layout.vMerge) remove(el, 'hMerge');
  if (span !== layout.gridSpan) {
    remove(el, 'gridSpan');
    if (layout.gridSpan > 1) setAttr(ensure(el, 'gridSpan'), 'val', layout.gridSpan);
  }
  if (haveVMerge !== layout.vMerge) {
    remove(el, 'vMerge');
    if (layout.vMerge === 'restart') setAttr(ensure(el, 'vMerge'), 'val', 'restart');
    else if (layout.vMerge === 'continue') ensure(el, 'vMerge');
  }

  if (raw && serializeXml(el) === before) return raw;
  sortChildren(el, TCPR_ORDER);
  return serializeXml(el);
}

// ---------------------------------------------------------------------------
// Table row properties

export interface RowModel {
  height: number | null; // twips (w:trHeight/@w:val)
  /** w:trHeight/@w:hRule as written: atLeast | exact | auto; null when omitted (Word reads it as atLeast). */
  heightRule: string | null;
  /** w:tblHeader: repeat as header row at the top of each page. */
  header: boolean;
  /** w:cantSplit: the row may not break across pages. */
  cantSplit: boolean;
}

export function readRowModel(trPr: Element | null): RowModel {
  const h = child(trPr, 'trHeight');
  return {
    height: numAttr(h, 'val'),
    heightRule: attr(h, 'hRule'),
    header: onOff(child(trPr, 'tblHeader')) === true,
    cantSplit: onOff(child(trPr, 'cantSplit')) === true,
  };
}

/** Original w:trPr (or null) patched so that it expresses `want`; unchanged input is returned as-is. */
export function writeTrPr(raw: string | null, want: RowModel): string {
  return memo(written, memoKey('t', raw, want), () => patchTrPr(raw, want));
}

function patchTrPr(raw: string | null, want: RowModel): string {
  const el = toElement(raw, 'trPr');
  const have = readRowModel(raw ? el : null);
  let changed = false;

  if (!same(have.height, want.height) || !same(have.heightRule, want.heightRule)) {
    changed = true;
    if (want.height == null) remove(el, 'trHeight');
    else {
      const h = ensure(el, 'trHeight');
      setAttr(h, 'val', want.height);
      setAttr(h, 'hRule', want.heightRule);
    }
  }
  const flag = (k: 'header' | 'cantSplit', local: string) => {
    if (have[k] === want[k]) return;
    changed = true;
    remove(el, local);
    if (want[k]) ensure(el, local);
  };
  flag('cantSplit', 'cantSplit');
  flag('header', 'tblHeader');

  if (!changed) return raw ?? '';
  sortChildren(el, TRPR_ORDER);
  return el.children.length || el.attributes.length ? serializeXml(el) : '';
}

// ---------------------------------------------------------------------------
// Pictures (w:drawing)

/** What the editor models of a picture's drawing. */
export interface DrawingInfo {
  cx: number | null; // EMU (wp:extent)
  cy: number | null;
  descr: string;
  embed: string | null; // relationship id of the picture (a:blip/@r:embed)
  /** Word shows the picture turned or flipped (pic:spPr/a:xfrm @rot, @flipH, @flipV). */
  turned: boolean;
  /** Word shows only part of the picture (a:srcRect): what is cut off each side, as a fraction of it; null when not cropped. */
  crop: { left: number; top: number; right: number; bottom: number } | null;
}

/** Parse kept XML even when it uses a prefix declared only on the original part's root. */
export function parseKept(xml: string): Element {
  const extra = new Set<string>();
  for (const m of xml.matchAll(/[<\s/]([A-Za-z_][\w.-]*):[A-Za-z_]/g)) {
    if (!(m[1] in WORD_NAMESPACES) && m[1] !== 'xmlns' && m[1] !== 'xml') extra.add(m[1]);
  }
  if (!extra.size) return parseFragment(xml);
  const decl = [...extra].map((p) => `xmlns:${p}="urn:papyrus:undeclared:${p}"`).join(' ');
  return parseFragment(`<px-wrap ${decl}>${xml}</px-wrap>`).firstElementChild!;
}

const drawings = new Lru<DrawingInfo>(500);

/** What the editor models of a kept w:drawing (cached, frozen). */
export function readDrawing(xml: string): DrawingInfo {
  return memo(drawings, xml, () => deepFreeze(parseDrawing(xml)));
}

function parseDrawing(xml: string): DrawingInfo {
  const d = parseKept(xml);
  const extent = find(d, 'extent', NS.wp);
  const docPr = find(d, 'docPr', NS.wp);
  return {
    cx: numAttr(extent, 'cx', ''),
    cy: numAttr(extent, 'cy', ''),
    descr: docPr?.getAttribute('descr') ?? '',
    embed: find(d, 'blip', NS.a)?.getAttributeNS(NS.r, 'embed') ?? null,
    turned: drawingTurned(d),
    crop: drawingCrop(d),
  };
}

function drawingTurned(d: Element): boolean {
  const xfrm = find(find(d, 'spPr', NS.pic), 'xfrm', NS.a);
  const on = (name: string) => ['1', 'true', 'on'].includes(xfrm?.getAttribute(name) ?? '');
  return !!xfrm && ((Number(xfrm.getAttribute('rot')) || 0) % 21600000 !== 0 || on('flipH') || on('flipV'));
}

function drawingCrop(d: Element): DrawingInfo['crop'] {
  const rect = find(find(d, 'blipFill', NS.pic), 'srcRect', NS.a);
  // In thousandths of a percent; a negative value is space added around the picture (not cut off).
  const side = (name: string) => Math.min(1, Math.max(0, (Number(rect?.getAttribute(name)) || 0) / 100000));
  const crop = { left: side('l'), top: side('t'), right: side('r'), bottom: side('b') };
  if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) return null;
  return crop.left || crop.top || crop.right || crop.bottom ? crop : null;
}

export interface DrawingPatch {
  /** New size in EMU; a dimension left out keeps its original value. */
  cx?: number;
  cy?: number;
  /** New alt text (wp:docPr/@descr). */
  descr?: string;
  /** A replaced picture: the relationship id of the new media (null: leave the reference out). */
  embed?: string | null;
}

const SVG_BLIP = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

/**
 * The original w:drawing with only the given changes: size (wp:extent and the picture's
 * a:xfrm), alt text, or the picture it shows. Wrapping, effects, ids and everything else stay.
 */
export function patchDrawing(xml: string, p: DrawingPatch): string {
  const d = parseKept(xml);
  const extent = find(d, 'extent', NS.wp);
  const xfrmExt = find(find(d, 'spPr', NS.pic), 'xfrm', NS.a)?.getElementsByTagNameNS(NS.a, 'ext')[0] ?? null;
  const resize = (key: 'cx' | 'cy', value: number | undefined) => {
    if (value == null || !extent) return;
    const old = Number(extent.getAttribute(key));
    extent.setAttribute(key, String(value));
    if (!xfrmExt) return;
    // The picture normally fills the frame; keep any other proportion it had.
    const inner = Number(xfrmExt.getAttribute(key));
    xfrmExt.setAttribute(key, String(inner === old || !old ? value : Math.round((inner * value) / old)));
  };
  resize('cx', p.cx);
  resize('cy', p.cy);

  if (p.descr !== undefined) {
    const docPr = find(d, 'docPr', NS.wp);
    const cNvPr = find(d, 'cNvPr', NS.pic);
    for (const e of [docPr, cNvPr]) {
      if (!e || (e === cNvPr && !e.hasAttribute('descr'))) continue;
      if (p.descr) e.setAttribute('descr', p.descr);
      else e.removeAttribute('descr');
    }
  }

  if (p.embed !== undefined) {
    const blip = find(d, 'blip', NS.a);
    if (blip) {
      blip.removeAttributeNS(NS.r, 'link');
      if (p.embed == null) blip.removeAttributeNS(NS.r, 'embed');
      else if (blip.hasAttributeNS(NS.r, 'embed')) blip.getAttributeNodeNS(NS.r, 'embed')!.value = p.embed;
      else blip.setAttributeNS(NS.r, 'r:embed', p.embed);
      // An SVG version of the old picture would still be shown by newer Word.
      for (const ext of Array.from(blip.getElementsByTagNameNS(NS.a, 'ext'))) {
        if (ext.getAttribute('uri') === SVG_BLIP) ext.parentNode!.removeChild(ext);
      }
      const extLst = child(blip, 'extLst', NS.a);
      if (extLst && !extLst.children.length) blip.removeChild(extLst);
      // Cropping was measured on the old picture.
      const srcRect = child(blip.parentElement, 'srcRect', NS.a);
      if (srcRect) srcRect.parentNode!.removeChild(srcRect);
    }
  }
  return serializeXml(d);
}
