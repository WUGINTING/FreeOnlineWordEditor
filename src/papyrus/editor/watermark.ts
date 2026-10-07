// Word's 浮水印 in the editor: finding the watermark shapes kept in a header (raw_inline nodes),
// replacing or removing them, keeping them from being deleted by typing in the header, and the
// layer that draws one on a page (see DocxEditor.renderPages; print.ts copies it into the printout).

import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { isHistoryTransaction } from 'prosemirror-history';
import { cssColor, isSafeImageSrc, schema } from './schema';
import { isWatermarkRun, readWatermark, type TextMetrics, type WatermarkShape } from '../docx/watermark';
import { parseLayers } from '../docx/wrappers';
import { cssFontName } from '../docx/fontNames';
import { localNames } from '../docx/eastAsiaFonts';
import { Lru, memo } from '../docx/memo';
import { sectionHeaderFooter, type Section } from '../docx/sections';
import type { DocxModel, HeaderFooterPart, HeaderFooterType } from '../docx/model';

/**
 * Set on the transactions that change a header's watermarks on purpose, which keepWatermarks
 * leaves alone: 'set' by 浮水印 (replaceWatermarkTr), 'keep' by keepWatermarks itself.
 */
export const WATERMARK_META = 'papyrusWatermark';

/** A watermark shape kept in a document (a header), with the picture shown for a picture watermark. */
export interface DocWatermark {
  shape: WatermarkShape;
  src: string | null;
}

/** A watermark itself (a run holding the shape), not a tracked deletion holding one. */
export const isWatermarkNode = (n: PMNode) => n.type === schema.nodes.raw_inline && isWatermarkRun(n.attrs.xml);

/** The watermark shapes of a document, in order. */
export function watermarkNodes(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (isWatermarkNode(node)) out.push({ node, pos });
    return !node.isInline;
  });
  return out;
}

const shown = new WeakMap<PMNode, DocWatermark | null>();
/** Whether a picture can be shown, by its source (the check decodes the picture). */
const safeSources = new Lru<boolean>(4);
const safeSource = (src: string | null): boolean => !!src && memo(safeSources, src, () => isSafeImageSrc(src));

/** The watermark a header shows (its first watermark shape Word would draw); null for none. */
export function docWatermark(doc: PMNode): DocWatermark | null {
  if (shown.has(doc)) return shown.get(doc)!;
  let found: DocWatermark | null = null;
  for (const { node } of watermarkNodes(doc)) {
    const shape = readWatermark(node.attrs.xml);
    if (!shape) continue;
    const src = shape.kind === 'picture' ? (node.attrs.src as string | null) : null;
    // A picture that can't be shown (not in the package, or not a picture) is no watermark here.
    if (shape.kind === 'picture' && !safeSource(src)) continue;
    found = { shape, src };
    break;
  }
  shown.set(doc, found);
  return found;
}

/** Everything a document keeps as-is (to see which shape types a part already defines). */
export function keptXml(doc: PMNode, except: (node: PMNode) => boolean = () => false): string {
  const out: string[] = [];
  doc.descendants((node) => {
    if ((node.type === schema.nodes.raw_inline || node.type === schema.nodes.raw_block) && !except(node)) out.push(node.attrs.xml ?? '');
    return true;
  });
  return out.join('');
}

/**
 * The header parts the sections show (each its own, or an earlier section's when linked), each
 * with the first section showing it. A part the file has but no section uses is not among them.
 */
export function sectionHeaders(model: DocxModel, sections: Section[]): Map<HeaderFooterPart, Section> {
  const out = new Map<HeaderFooterPart, Section>();
  for (const section of sections) {
    for (const type of TYPES) {
      const relId = section.refs.header[type];
      const hf = relId
        ? model.headerFooters.find((h) => h.relId === relId && h.kind === 'header')
        : sectionHeaderFooter(model, sections, section.index, 'header', type);
      if (hf && !hf.pending && !out.has(hf)) out.set(hf, section);
    }
  }
  return out;
}

const TYPES = ['default', 'first', 'even'] as const;

/** Where a new watermark goes (see watermarkTargets). */
export interface WatermarkTargets {
  /** The header parts to hold it, each with a section showing it (for its page size). */
  parts: Map<HeaderFooterPart, Section>;
  /**
   * Headers to make: a section that uses a header of this type and shows none. Later sections
   * without one of their own show it too, so it is made once, for the first of them.
   */
  missing: { section: Section; type: HeaderFooterType }[];
}

/**
 * Word's 浮水印 goes into every header part the sections refer to (default, first page and even
 * pages, used or not) and into the one a section shows linked to an earlier section. A section
 * that uses a type of header (default always, first page with 不同的首頁, even pages with 奇偶頁
 * 不同) and shows none needs one made. A reference to a part that could not be read is left alone.
 */
export function watermarkTargets(model: DocxModel, sections: Section[]): WatermarkTargets {
  const parts = new Map<HeaderFooterPart, Section>();
  const missing: WatermarkTargets['missing'] = [];
  const made = new Set<HeaderFooterType>();
  for (const section of sections) {
    for (const type of TYPES) {
      const relId = section.refs.header[type];
      if (relId) {
        const own = model.headerFooters.find((h) => h.relId === relId && h.kind === 'header');
        if (own && !parts.has(own)) parts.set(own, section);
        continue;
      }
      const used = type === 'default' || (type === 'first' && section.titlePage) || (type === 'even' && model.evenAndOdd);
      if (!used) continue;
      const shows = sectionHeaderFooter(model, sections, section.index, 'header', type);
      if (shows && !shows.pending) {
        if (!parts.has(shows)) parts.set(shows, section);
      } else if (!made.has(type)) {
        made.add(type);
        missing.push({ section, type });
      }
    }
  }
  return { parts, missing };
}

/** Shape ids and watermark numbers the documents use, so new ones are unique. */
export function usedShapeIds(docs: PMNode[]): { spids: number[]; numbers: Set<string> } {
  const spids: number[] = [];
  const numbers = new Set<string>();
  for (const doc of docs) {
    const xml = keptXml(doc);
    for (const m of xml.matchAll(/o:spid="_x0000_s(\d+)"/g)) spids.push(Number(m[1]));
    for (const m of xml.matchAll(/(?:PowerPlusWaterMarkObject|WordPictureWatermark)(\d+)/g)) numbers.add(m[1]);
  }
  return { spids, numbers };
}

/** A paragraph in the content control Word's 浮水印 gallery inserts (w:docPartGallery "Watermarks"). */
function inWatermarkGallery(node: PMNode): boolean {
  if (node.type !== schema.nodes.paragraph || !node.attrs.wrap) return false;
  try {
    return parseLayers(node.attrs.wrap).some((l) => /<w:docPartGallery\b[^>]*w:val="Watermarks"/.test(l.open));
  } catch {
    return false;
  }
}

/** The v:shapetype definitions in kept XML, with their ids. */
function shapetypes(xml: string): { id: string; xml: string }[] {
  return [...xml.matchAll(/<v:shapetype\b[^>]*?\bid="([^"]+)"[^>]*?(?:\/>|>[\s\S]*?<\/v:shapetype>)/g)].map((m) => ({ id: m[1], xml: m[0] }));
}

/**
 * A transaction that takes the document's watermark shapes out and, with `node`, puts that one
 * where the first of them was, or else at the start of the first paragraph (where Word puts it).
 * A paragraph of Word's 浮水印 gallery control left empty goes, and with its last paragraph the
 * control (w:sdt) itself; the header keeps at least one paragraph. A shape type a removed
 * watermark defined stays (as an invisible run of its own) when another shape in the part uses
 * it. Not an undo step of the part (see DocxEditor.setWatermark). Null when nothing changes.
 */
export function replaceWatermarkTr(state: EditorState, node: PMNode | null): Transaction | null {
  const old = watermarkNodes(state.doc);
  if (!old.length && !node) return null;
  const tr = state.tr;
  const going = new Set(old.map((w) => w.node));
  let rest = keptXml(state.doc, (n) => going.has(n)) + (node?.attrs.xml ?? '');
  for (const { node: n, pos } of [...old].reverse()) {
    const needed = shapetypes(n.attrs.xml).filter((t) => rest.includes(`"#${t.id}"`) && !rest.includes(`id="${t.id}"`));
    if (!needed.length) {
      tr.delete(pos, pos + n.nodeSize);
      continue;
    }
    const kept = needed.map((t) => t.xml).join('');
    rest += kept;
    tr.replaceWith(pos, pos + n.nodeSize, schema.nodes.raw_inline.create(
      { xml: `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>${kept}</w:pict></w:r>`, label: '圖形', hidden: true },
      null,
      n.marks,
    ));
  }
  const emptied: { from: number; to: number }[] = [];
  tr.doc.forEach((child, offset) => {
    if (inWatermarkGallery(child) && child.content.size === 0) emptied.push({ from: offset, to: offset + child.nodeSize });
  });
  for (const r of emptied.reverse()) tr.delete(r.from, r.to);
  if (tr.doc.childCount === 0) tr.insert(0, schema.nodes.paragraph.create());
  if (node) {
    let at = old.length ? tr.mapping.map(old[0].pos) : null;
    if (at == null || !tr.doc.resolve(at).parent.inlineContent) at = firstParagraphStart(tr);
    tr.insert(at, node);
  }
  return tr.setMeta(WATERMARK_META, 'set').setMeta('addToHistory', false);
}

/** The start of the document's first top-level paragraph (one is added at the end when there is none). */
function firstParagraphStart(tr: Transaction): number {
  let at = -1;
  tr.doc.forEach((child, offset) => {
    if (at < 0 && child.type === schema.nodes.paragraph) at = offset + 1;
  });
  if (at >= 0) return at;
  tr.insert(tr.doc.content.size, schema.nodes.paragraph.create());
  return tr.doc.content.size - 1;
}

/** What 設計 › 浮水印 did to a header: the watermarks it placed, and those it took out since. */
interface WatermarkHistory {
  placed: Set<string>;
  replaced: Set<string>;
}
const watermarkHistoryKey = new PluginKey<WatermarkHistory>('papyrusWatermarks');
const xmlsOf = (doc: PMNode) => watermarkNodes(doc).map((w) => w.node.attrs.xml as string);

/**
 * A header's watermarks change with 設計 › 浮水印 (DocxEditor.setWatermark), which is not an undo
 * step, and with edits of the watermarks themselves that are (rejecting or accepting a tracked
 * deletion of one). Anything else keeps them as they are:
 * - one deleted by editing the header around it (typing over a selection that holds it,
 *   Backspace, cutting) is put back where it was, and the user is told how to remove it
 *   (設計 › 浮水印 › 移除浮水印), as Word keeps a watermark out of the way of the header's text;
 * - Ctrl+Z / Ctrl+Y can't undo 浮水印: a watermark it placed stays, and one it took out that the
 *   header's undo history brings back (it still holds the header as it was before) goes again.
 *   Other changes Ctrl+Z / Ctrl+Y make to watermarks (undoing a rejected tracked deletion) stand.
 */
export function keepWatermarks(onNotice: (message: string) => void): Plugin<WatermarkHistory> {
  return new Plugin<WatermarkHistory>({
    key: watermarkHistoryKey,
    state: {
      init: () => ({ placed: new Set(), replaced: new Set() }),
      apply(tr, value, oldState, newState) {
        if (tr.getMeta(WATERMARK_META) !== 'set') return value;
        const after = xmlsOf(newState.doc);
        const replaced = new Set(value.replaced);
        for (const xml of xmlsOf(oldState.doc)) if (!after.includes(xml)) replaced.add(xml);
        for (const xml of after) replaced.delete(xml);
        return { placed: new Set(after), replaced };
      },
    },
    appendTransaction(trs, oldState, newState) {
      if (!trs.some((tr) => tr.docChanged) || trs.some((tr) => tr.getMeta(WATERMARK_META))) return null;
      const before = watermarkNodes(oldState.doc);
      const after = watermarkNodes(newState.doc);
      if (!before.length && !after.length) return null;
      // Matched by their XML: what is left over on each side was lost or came in.
      const fresh = [...after];
      const lost = before.filter(({ node }) => {
        const i = fresh.findIndex((w) => w.node.attrs.xml === node.attrs.xml);
        if (i < 0) return true;
        fresh.splice(i, 1);
        return false;
      });
      const history = trs.some((t) => isHistoryTransaction(t));
      const made = watermarkHistoryKey.getState(newState) ?? { placed: new Set<string>(), replaced: new Set<string>() };
      const stale = history ? fresh.filter((w) => made.replaced.has(w.node.attrs.xml) && !made.placed.has(w.node.attrs.xml)) : [];
      const restore = history ? lost.filter((w) => made.placed.has(w.node.attrs.xml)) : lost;
      if (!stale.length && !restore.length) return null;
      const tr = newState.tr;
      for (const { node, pos } of [...stale].reverse()) tr.delete(pos, pos + node.nodeSize);
      for (const { node, pos } of restore) {
        let at = trs.reduce((p, t) => t.mapping.map(p, -1), pos);
        at = tr.mapping.map(at, -1);
        if (at > tr.doc.content.size || !tr.doc.resolve(at).parent.inlineContent) at = firstParagraphStart(tr);
        tr.insert(at, node);
      }
      if (restore.length && !history) onNotice('浮水印不會因編輯頁首而刪除；要移除請用「設計」›「浮水印」›「移除浮水印」。');
      return tr.setMeta(WATERMARK_META, 'keep');
    },
  });
}

/** Pasted or dropped content in a header loses its watermark shapes: a second copy would share the first's ids. */
export function withoutWatermarks(slice: Slice): Slice {
  let found = false;
  slice.content.descendants((n) => {
    if (isWatermarkNode(n)) found = true;
    return !found;
  });
  if (!found) return slice;
  const strip = (fragment: Fragment): Fragment => {
    const out: PMNode[] = [];
    fragment.forEach((child) => {
      if (isWatermarkNode(child)) return;
      out.push(child.isLeaf ? child : child.copy(strip(child.content)));
    });
    return Fragment.from(out);
  };
  return new Slice(strip(slice.content), slice.openStart, slice.openEnd);
}

type Measurer = { font: string; measureText(text: string): globalThis.TextMetrics };
let measurer: Measurer | null = null;
/** What `measurer` was made with (an OffscreenCanvas class, or 'canvas'). */
let measurerSource: unknown = null;
/** Texts measured in their font, by font and text (one whose font is not loaded yet is measured again later). */
const measured = new Lru<TextMetrics>(64);

/**
 * A 2D context to measure text with: an OffscreenCanvas, else a canvas element; none in jsdom
 * (tests), whose canvas can't measure (and complains when asked).
 */
function measuringContext(): Measurer | null {
  const source = typeof OffscreenCanvas !== 'undefined'
    ? OffscreenCanvas
    : typeof document !== 'undefined' && !/\bjsdom\b/i.test(globalThis.navigator?.userAgent ?? '') ? 'canvas' : null;
  if (source !== measurerSource) {
    // Measured with something else before (tests swap it): measured again.
    measurerSource = source;
    measured.clear();
    measurer = null;
    if (source === 'canvas') measurer = document.createElement('canvas').getContext('2d');
    else if (source) measurer = new OffscreenCanvas(1, 1).getContext('2d') as Measurer | null;
  }
  return measurer;
}

/**
 * The text's width and its font's height (ascent + descent) in em, measured in `font` on a
 * canvas; null where there is none (tests) or the font is not loaded yet.
 */
export function measureWatermarkText(text: string, font: string): TextMetrics | null {
  if (!text) return null;
  const family = localNames(font || 'serif').map(cssFontName).filter(Boolean).map((f) => `"${f}"`).join(', ') || 'serif';
  const spec = `100px ${family}`;
  const key = `${spec}\n${text}`;
  try {
    const ctx = measuringContext();
    if (!ctx || typeof ctx.measureText !== 'function') return null;
    const known = measured.get(key);
    if (known) return known;
    const fonts = typeof document === 'undefined' ? null : document.fonts;
    if (fonts && typeof fonts.check === 'function' && !fonts.check(spec, text)) return null;
    ctx.font = spec;
    const m = ctx.measureText(text);
    const ascent = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent;
    const descent = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent;
    if (!(m.width > 0) || !(ascent + descent > 0)) return null;
    const out = { width: m.width / 100, height: (ascent + descent) / 100, ascent: ascent / (ascent + descent) };
    measured.set(key, out);
    return out;
  } catch {
    return null;
  }
}

/** Where a page's watermark goes, in px: the page's size and margins. */
export interface WatermarkPage {
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
}

const SVG = 'http://www.w3.org/2000/svg';
const round = (n: number) => Math.round(n * 100) / 100;

/**
 * The page layer drawing a watermark as Word does: centred between the margins (or on the page),
 * turned, behind the header and the text, not selectable. Plain markup and CSS (an SVG text
 * stretched to the shape, or a picture), so the printout and the server's PDF, where no script
 * runs, show it the same. Null when there is nothing to draw.
 */
export function watermarkLayer(wm: DocWatermark, page: WatermarkPage, metrics?: TextMetrics | null): HTMLElement | null {
  const { shape } = wm;
  const width = (shape.width * 96) / 72;
  const height = (shape.height * 96) / 72;
  if (!(width > 0 && height > 0)) return null;
  const centreX = shape.pageRelative.horizontal ? page.width / 2 : page.left + (page.width - page.left - page.right) / 2;
  const centreY = shape.pageRelative.vertical ? page.height / 2 : page.top + (page.height - page.top - page.bottom) / 2;

  const layer = document.createElement('div');
  layer.className = 'dx-wm-layer';
  layer.setAttribute('aria-hidden', 'true');
  const box = document.createElement('div');
  box.className = 'dx-wm';
  box.style.left = `${round(centreX - width / 2)}px`;
  box.style.top = `${round(centreY - height / 2)}px`;
  box.style.width = `${round(width)}px`;
  box.style.height = `${round(height)}px`;
  if (shape.rotation) box.style.transform = `rotate(${shape.rotation}deg)`;

  if (shape.kind === 'picture') {
    if (!safeSource(wm.src)) return null;
    const img = document.createElement('img');
    img.className = shape.washout ? 'dx-wm-img dx-wm-washout' : 'dx-wm-img';
    img.alt = '';
    img.draggable = false;
    img.src = wm.src!;
    box.append(img);
  } else {
    if (!shape.text) return null;
    // WordArt fits the text to the shape (Word's fitshape): across its width (textLength), and
    // its font's height (ascent + descent, measured as when it was written) over the shape's
    // height; unmeasured, 0.9 of the height for the font size (a full-width character fills it).
    const svg = document.createElementNS(SVG, 'svg');
    svg.setAttribute('viewBox', `0 0 ${round(width)} ${round(height)}`);
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.setAttribute('preserveAspectRatio', 'none');
    const text = document.createElementNS(SVG, 'text');
    const m = metrics === undefined ? measureWatermarkText(shape.text, shape.font) : metrics;
    const size = m ? height / m.height : height * 0.9;
    const baseline = m ? size * m.height * (m.ascent ?? 0.8) : height / 2 + size * 0.38;
    text.setAttribute('x', '0');
    text.setAttribute('y', String(round(baseline)));
    text.setAttribute('font-size', String(round(size)));
    text.setAttribute('textLength', String(round(width)));
    text.setAttribute('lengthAdjust', 'spacingAndGlyphs');
    const fonts = shape.font ? localNames(shape.font).map(cssFontName).filter(Boolean).map((f) => `"${f}"`) : [];
    text.setAttribute('font-family', [...fonts, 'serif'].join(', '));
    text.setAttribute('fill', cssColor(shape.color) ?? '#c0c0c0');
    if (shape.opacity < 1) text.setAttribute('fill-opacity', String(round(shape.opacity)));
    text.textContent = shape.text;
    svg.append(text);
    box.append(svg);
  }
  layer.append(box);
  return layer;
}
