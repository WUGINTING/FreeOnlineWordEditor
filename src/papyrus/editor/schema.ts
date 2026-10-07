import { Schema, type NodeSpec, type MarkSpec, type DOMOutputSpec, type Node as PMNode } from 'prosemirror-model';
import { tableNodes } from 'prosemirror-tables';
import { twipsToPx } from '../units';
import { hasEastAsianText, wordLineRatio } from '../docx/fontMetrics';
import { deletionDOM, insertionAttrs } from '../docx/revisions';
import { tcMarginCss } from '../docx/cellMargins';
import { isSafeHref, safeHref } from '../docx/links';
import { HIGHLIGHT_COLORS, hasEastAsiaHint } from '../docx/props';
import { useEastAsiaFont } from '../docx/eastAsiaFonts';
import { eastAsiaAlias } from '../docx/fontNames';
import { fieldNumberFormat } from '../docx/pageNumbers';
import { fontStack } from '../docx/fonts';
import { pastedBorders } from './pasteExcel';
import { isWatermarkRun } from '../docx/watermark';
import { tl } from '../i18n';

// ----- pictures: what a pasted <img> may show -----

/** Picture formats a pasted picture may have (as a data: URL), and how each one starts. */
const IMAGE_SIGNATURES: Record<string, (head: string) => boolean> = {
  png: (b) => b.startsWith('\x89PNG'),
  jpeg: (b) => b.startsWith('\xFF\xD8\xFF'),
  gif: (b) => b.startsWith('GIF8'),
  bmp: (b) => b.startsWith('BM'),
  webp: (b) => b.startsWith('RIFF') && b.slice(8, 12) === 'WEBP',
  'svg+xml': (b) => /<svg[\s>]/i.test(b),
};
const DATA_IMAGE = /^data:image\/(png|jpeg|gif|bmp|webp|svg\+xml)((?:;[\w.+-]+=[^;,]*)*)(;base64)?,/i;
const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;

/** A data: URL of a PNG, JPEG, GIF, BMP, WebP or SVG picture whose content really is one. */
function decodesAsImage(src: string): boolean {
  const m = DATA_IMAGE.exec(src);
  if (!m) return false;
  const payload = src.slice(m[0].length);
  if (!payload) return false;
  let head: string;
  try {
    if (m[3]) {
      if (payload.length % 4 !== 0 || !BASE64.test(payload)) return false;
      head = atob(payload.slice(0, 1368)); // the first kilobyte
    } else {
      head = decodeURIComponent(payload.slice(0, 3072).replace(/%[0-9a-f]?$/i, ''));
    }
  } catch {
    return false;
  }
  return IMAGE_SIGNATURES[m[1].toLowerCase()](head);
}

/** Pictures of the documents open in an editor (any format the file has), by editor. */
const packageImages = new Map<object, Set<string>>();

/**
 * The pictures of an opened document: a copy of one pasted back in is kept as it is, whatever
 * its format (EMF, TIFF ...). Replaces what `owner` registered before.
 */
export function registerImageSources(owner: object, sources: Iterable<string>): void {
  const set = new Set<string>();
  for (const src of sources) if (src && !decodesAsImage(src)) set.add(src);
  packageImages.set(owner, set);
}

export function releaseImageSources(owner: object): void {
  packageImages.delete(owner);
}

/**
 * Whether a picture's src may be shown and saved: a PNG, JPEG, GIF, BMP, WebP or SVG data: URL
 * that decodes as one, or a picture of an open document. Never a remote or script URL.
 */
export function isSafeImageSrc(src: unknown): boolean {
  if (typeof src !== 'string' || !src) return false;
  for (const set of packageImages.values()) if (set.has(src)) return true;
  return decodesAsImage(src);
}

const IMG_TAG = /<img\b(?:[^>"']|"[^"]*"|'[^']*')*>/gi;
const escapeText = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * Pasted HTML with every picture isSafeImageSrc refuses (a Word clipboard's file:///…clip_image,
 * a web page's https:// picture) replaced by its alt text, or by what `replace` gives for it
 * (its index among the refused ones, from 0).
 */
export function replaceUnsafeImages(html: string, replace?: (img: Element, index: number) => string): string {
  const probe = document.createElement('template');
  let n = 0;
  return html.replace(IMG_TAG, (tag) => {
    probe.innerHTML = tag; // inert: a template's content loads nothing
    const img = probe.content.firstElementChild;
    if (!img || isSafeImageSrc(img.getAttribute('src'))) return tag;
    return replace ? replace(img, n++) : (n++, escapeText(img.getAttribute('alt') ?? ''));
  });
}

/** How many pictures of pasted HTML replaceUnsafeImages replaces. */
export function unsafeImageCount(html: string): number {
  if (!/<img\b/i.test(html)) return 0;
  let n = 0;
  replaceUnsafeImages(html, () => (n++, ''));
  return n;
}

/** A placeholder picture standing for picture `index` of paste `token` until its file is read. */
export function pasteSlotSrc(token: string, index: number): string {
  return `data:image/gif;dxslot=${token}-${index};base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7`;
}

/** Which picture of paste `token` a placeholder stands for (-1: not one of them). */
export function pasteSlotIndex(src: unknown, token: string): number {
  const m = typeof src === 'string' ? /^data:image\/gif;dxslot=([\w-]+)-(\d+);/.exec(src) : null;
  return m && m[1] === token ? Number(m[2]) : -1;
}

// Paragraph geometry is stored in OOXML units (twips) so it round-trips exactly;
// it is converted to CSS only when rendering.
const paragraphAttrs = {
  styleId: { default: null as string | null },
  align: { default: null as string | null }, // left | center | right | justify
  indLeft: { default: null as number | null }, // twips
  indRight: { default: null as number | null },
  indFirst: { default: null as number | null }, // twips; negative = hanging
  // Indents in hundredths of a character (w:leftChars ...). Word lets these override the twips values.
  indLeftChars: { default: null as number | null },
  indRightChars: { default: null as number | null },
  indFirstChars: { default: null as number | null },
  spaceBefore: { default: null as number | null },
  spaceAfter: { default: null as number | null },
  line: { default: null as number | null }, // w:spacing/@w:line
  lineRule: { default: null as string | null }, // auto | exact | atLeast
  numId: { default: null as string | null },
  ilvl: { default: 0 },
  pageBreakBefore: { default: false },
  sectPr: { default: null as string | null }, // raw XML of an in-paragraph section break
  // Lossless round-trip: the original w:pPr and the w:p element's own attributes.
  pPr: { default: null as string | null },
  pAttrs: { default: null as string | null },
  /** Pieces of one w:p split around page breaks share this id with the breaks. */
  brGroup: { default: null as string | null },
  /** Enclosing block wrappers (w:sdt, w:customXml), outermost first; see docx/wrappers.ts. */
  wrap: { default: null as string | null },
};

const wrapAttr = { wrap: { default: null as string | null } };

export function paragraphStyle(attrs: Record<string, any>): string {
  const s: string[] = [];
  // Only known values: the attribute comes from the file and ends up in a style attribute.
  const align = attrs.align === 'both' ? 'justify' : attrs.align;
  if (['left', 'center', 'right', 'justify'].includes(align)) s.push(`text-align:${align}`);
  // A value in characters wins over twips, 0 included (Word lays w:leftChars="0" out as no indent).
  const side = (chars: number | null, twips: number | null): string | null =>
    chars != null ? `${chars / 100}em` : twips != null ? `${twipsToPx(twips)}px` : null;
  const left = side(attrs.indLeftChars, attrs.indLeft);
  if (left) s.push(`margin-left:${left}`);
  const right = side(attrs.indRightChars, attrs.indRight);
  if (right) s.push(`margin-right:${right}`);
  const first = side(attrs.indFirstChars, attrs.indFirst);
  if (first) s.push(`text-indent:${first}`);
  if (attrs.spaceBefore != null) s.push(`margin-top:${twipsToPx(attrs.spaceBefore)}px`);
  if (attrs.spaceAfter != null) s.push(`margin-bottom:${twipsToPx(attrs.spaceAfter)}px`);
  const lh = lineHeightCss(attrs.line, attrs.lineRule);
  if (lh) s.push(`line-height:${lh}`);
  const grid = lineGridVars(attrs.line, attrs.lineRule);
  if (grid) s.push(...Object.entries(grid).map(([k, v]) => `${k}:${v}`));
  return s.join(';');
}

/**
 * The line spacing again, as the document grid needs it (editor/docGrid.ts): in a section with a
 * line grid Word lays a line out as a whole number of grid lines, times "auto" spacing
 * (--dx-lm), at least an "at least" height (--dx-lmin); an exact height (--dx-lfix) stays.
 * Each spacing sets all three, so one from a style never mixes with the paragraph's own.
 */
export function lineGridVars(line: number | null, rule: string | null): Record<string, string> | null {
  if (line == null) return null;
  if (!rule || rule === 'auto') return { '--dx-lm': String(+(line / 240).toFixed(3)), '--dx-lmin': '0px', '--dx-lfix': 'initial' };
  if (rule === 'atLeast') return { '--dx-lm': '1', '--dx-lmin': `${twipsToPx(line)}px`, '--dx-lfix': 'initial' };
  return { '--dx-lfix': `${twipsToPx(line)}px` };
}

export function lineHeightCss(line: number | null, rule: string | null): string | null {
  if (line == null) return null;
  // "Auto" spacing multiplies Word's single line height for the paragraph's fonts (--dx-lh).
  if (!rule || rule === 'auto') return `calc(var(--dx-lh, 1.15) * ${+(line / 240).toFixed(3)})`;
  return `${twipsToPx(line)}px`;
}

// ----- Lists in pasted HTML -----
//
// Parse rules cannot see the document's numbering, so a pasted list item gets a
// placeholder numId; the editor turns placeholders into real Word lists when the paste
// is inserted (DocxEditor.adoptPastedLists).

/** Placeholder numId prefix for an HTML list (<ol>/<ul>) met while parsing a paste. */
export const PASTED_LIST = 'paste:';
/**
 * Placeholder numId prefix for a list paragraph copied from an editor: its numId there, then
 * "@" and the id of the document it was copied from (data-doc; see pasteLists.ts).
 */
export const COPIED_LIST = 'copy:';

export interface PastedList {
  /** Numbering format per level: bullet, decimal, lowerLetter ... */
  formats: string[];
  /** First number of the top level (<ol start>). */
  start: number;
  /** The number's text per level, Word style ("(%1)", "%1、"); "%1." when not given. */
  texts?: (string | undefined)[];
}

/** Lists seen by the parser, by placeholder id; taken out when the paste is inserted. */
export const pastedLists = new Map<string, PastedList>();
/**
 * How a list copied from an editor looks (data-num-def), by COPIED_LIST placeholder: a paste
 * into another document makes a list like it. Taken out when the paste is inserted.
 */
export const copiedLists = new Map<string, PastedList>();

/** The list definition an editor put on a copied list paragraph (data-num-def), checked. */
function copiedListDef(json: string | null): PastedList | null {
  if (!json) return null;
  try {
    const d = JSON.parse(json) as { formats?: unknown; texts?: unknown; start?: unknown };
    const formats = Array.isArray(d.formats) ? d.formats.slice(0, 9).map((f) => (typeof f === 'string' && /^[A-Za-z]{1,40}$/.test(f) ? f : 'decimal')) : [];
    if (!formats.length) return null;
    const texts = Array.isArray(d.texts) ? d.texts.slice(0, 9).map((t) => (typeof t === 'string' && t.length <= 40 ? t : undefined)) : [];
    const start = typeof d.start === 'number' && Number.isFinite(d.start) ? Math.min(32767, Math.max(0, Math.floor(d.start))) : 1;
    return { formats, texts, start };
  } catch {
    return null;
  }
}
const listIds = new WeakMap<Element, string>();
let listSeq = 0;

const isList = (el: Element | null): el is Element => !!el && (el.tagName === 'OL' || el.tagName === 'UL');

function listFormat(list: Element): string {
  if (list.tagName === 'UL') return 'bullet';
  switch (list.getAttribute('type')) {
    case 'a': return 'lowerLetter';
    case 'A': return 'upperLetter';
    case 'i': return 'lowerRoman';
    case 'I': return 'upperRoman';
    default: return 'decimal';
  }
}

/** numId / ilvl for an <li>: one list per outermost <ol>/<ul>, nesting depth as the level. */
function pastedListAttrs(li: Element): { numId: string; ilvl: number } | null {
  const list = li.parentElement;
  if (!isList(list)) return null;
  let top = list;
  let depth = 0;
  for (let e = list.parentElement; e; e = e.parentElement) {
    if (isList(e)) {
      top = e;
      depth++;
    }
  }
  let id = listIds.get(top);
  // The same HTML parsed again after its list was used: count it as a new list.
  if (!id || !pastedLists.has(id)) {
    id = PASTED_LIST + ++listSeq;
    listIds.set(top, id);
    pastedLists.set(id, { formats: [], start: Number(top.getAttribute('start')) || 1 });
  }
  const info = pastedLists.get(id)!;
  info.formats[depth] ??= listFormat(list);
  return { numId: id, ilvl: Math.min(depth, 8) };
}

// Lists copied from Word: the clipboard has no <ol>/<ul>. Each item is a paragraph styled
// `mso-list:l0 level2 lfo1` (list l0, level 2, list instance lfo1), and its number is typed-out
// text in a span styled `mso-list:Ignore` (dropped when parsing; see ignoreWordListNumber).
const MSO_LIST = /mso-list:\s*(l\d+)\s+level(\d+)\s+(lfo\d+)/i;
const MSO_IGNORE = /mso-list:\s*Ignore/i;
const msoListIds = new Map<string, string>();

const CJK_DIGITS = '〇一二三四五六七八九';
const STEMS = '甲乙丙丁戊己庚辛壬癸';

/** 一 → 1, 十二 → 12, 二十 → 20 (up to 99). */
function cjkNumber(s: string): number | null {
  const m = /^([一二三四五六七八九]?)(十?)([一二三四五六七八九]?)$/.exec(s);
  if (!m || !s) return null;
  if (!m[2]) return m[1] && !m[3] ? CJK_DIGITS.indexOf(m[1]) : null;
  return (m[1] ? CJK_DIGITS.indexOf(m[1]) : 1) * 10 + (m[3] ? CJK_DIGITS.indexOf(m[3]) : 0);
}

function romanNumber(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  let total = 0;
  const t = s.toLowerCase();
  for (let i = 0; i < t.length; i++) {
    const a = v[t[i]];
    const b = v[t[i + 1]] ?? 0;
    total += a < b ? -a : a;
  }
  return total;
}

/** A Word list number's text ("3.", "(二)", "c)", "·") as a format, start and level text. */
export function wordListNumber(text: string, level: number): { fmt: string; start: number | null; text: string } {
  const t = text.replace(/[\s ]+/g, '');
  const m = /^([(（]?)([0-9]+|[a-zA-Z]+|[一二三四五六七八九十]+|[甲乙丙丁戊己庚辛壬癸])([.)）、．]?)$/.exec(t);
  if (!m) return { fmt: 'bullet', start: null, text: '' };
  const [, open, core, close] = m;
  const lvl = `${open}%${level + 1}${close}`;
  if (/^[0-9]+$/.test(core)) return { fmt: 'decimal', start: Number(core), text: lvl };
  const cjk = cjkNumber(core);
  if (cjk != null) return { fmt: 'taiwaneseCountingThousand', start: cjk, text: lvl };
  if (STEMS.includes(core)) return { fmt: 'ideographTraditional', start: STEMS.indexOf(core) + 1, text: lvl };
  // Roman when it can only be roman ("ii", "iv") or is Word's usual first roman item ("i").
  if (/^[ivxlcdm]+$/i.test(core) && (core.length > 1 || /^i$/i.test(core))) {
    return { fmt: core === core.toLowerCase() ? 'lowerRoman' : 'upperRoman', start: romanNumber(core), text: lvl };
  }
  if (/^[a-zA-Z]$/.test(core)) {
    const lower = core.toLowerCase();
    return { fmt: core === lower ? 'lowerLetter' : 'upperLetter', start: lower.charCodeAt(0) - 96, text: lvl };
  }
  return { fmt: 'bullet', start: null, text: '' };
}

/** numId / ilvl for a paragraph copied from Word's own list (mso-list). */
function wordListAttrs(el: HTMLElement): { numId: string; ilvl: number } | null {
  const m = MSO_LIST.exec(el.getAttribute('style') ?? '');
  if (!m) return null;
  const ilvl = Math.min(8, Math.max(0, Number(m[2]) - 1));
  const key = `${m[1]}|${m[3]}`;
  let id = msoListIds.get(key);
  // The same list pasted again after it was used: a new list.
  if (!id || !pastedLists.has(id)) {
    id = PASTED_LIST + ++listSeq;
    msoListIds.set(key, id);
    pastedLists.set(id, { formats: [], start: 1, texts: [] });
  }
  const info = pastedLists.get(id)!;
  if (info.formats[ilvl] === undefined) {
    const marker = Array.from(el.querySelectorAll('span')).find((s) => MSO_IGNORE.test(s.getAttribute('style') ?? ''));
    const n = wordListNumber(marker?.textContent ?? '', ilvl);
    info.formats[ilvl] = n.fmt;
    (info.texts ??= [])[ilvl] = n.text || undefined;
    if (ilvl === 0 && n.start != null) info.start = n.start;
  }
  return { numId: id, ilvl };
}

/** Parse rule: the typed-out number of a Word list item is not text of the paragraph. */
export const ignoreWordListNumber = {
  tag: 'span',
  priority: 100,
  ignore: true,
  getAttrs: (dom: HTMLElement) => (MSO_IGNORE.test(dom.getAttribute('style') ?? '') ? null : false),
};

const BLOCK_TAGS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'TABLE']);

/** List attributes for a pasted <p>: an editor list paragraph, a Word list item, or the first block of an <li>. */
function listAttrsOfBlock(el: HTMLElement): { numId: string; ilvl: number } | Record<string, never> {
  const copied = el.getAttribute('data-num-id');
  if (copied) {
    const numId = `${COPIED_LIST}${copied}@${el.getAttribute('data-doc') ?? ''}`;
    const def = copiedListDef(el.getAttribute('data-num-def'));
    if (def) copiedLists.set(numId, def);
    return { numId, ilvl: Math.min(8, Math.max(0, Number(el.getAttribute('data-ilvl')) || 0)) };
  }
  const word = wordListAttrs(el);
  if (word) return word;
  const li = el.parentElement;
  if (li?.tagName === 'LI' && li.firstElementChild === el) return pastedListAttrs(li) ?? {};
  return {};
}

const NON_EAST_ASIAN = /[^\s⺀-鿿가-힯豈-﫿＀-￯]/;

/**
 * Word's single line height from the fonts named directly on the paragraph's runs: the tallest
 * of them, and whether they cover all its text. Word sizes a line by the fonts its text really
 * uses (and the paragraph mark's), so when the runs name every font the style's fonts don't count.
 */
function runFontsLine(node: PMNode): { ratio: number | null; covered: boolean } {
  let max: number | null = null;
  let covered = true;
  const take = (r: number | null) => {
    if (r != null && (max == null || r > max)) max = r;
  };
  node.forEach((child) => {
    const font = child.marks.find((m) => m.type.name === 'font');
    if (!child.isText) {
      if (child.type.name !== 'image' && !font?.attrs.family) covered = false;
      if (font) take(wordLineRatio(font.attrs.family));
      return;
    }
    const text = child.text ?? '';
    const eastAsian = hasEastAsianText(text);
    const latin = NON_EAST_ASIAN.test(text);
    if (!font || (latin && !font.attrs.family) || (eastAsian && !font.attrs.eastAsia)) covered = false;
    if (!font) return;
    if (latin) take(wordLineRatio(font.attrs.family));
    if (eastAsian) take(wordLineRatio(font.attrs.eastAsia));
  });
  // The paragraph mark's own font (w:pPr/w:rPr) sizes the last line, and alone an empty paragraph.
  const markRPr = /<w:rPr\b[\s\S]*?<\/w:rPr>/.exec((node.attrs.pPr as string | null) ?? '')?.[0] ?? '';
  const markFont = /<w:rFonts\b[^>]*\bw:ascii="([^"]+)"/.exec(markRPr)?.[1];
  if (markFont) take(wordLineRatio(markFont));
  else if (node.childCount === 0) covered = false;
  return { ratio: max, covered: covered && max != null };
}

/** Word's name for a section break mark, by how the section after it starts (its w:type). */
export const SECTION_LABEL: Record<string, string> = {
  nextPage: '分節符號（下一頁）',
  continuous: '分節符號（接續本頁）',
  evenPage: '分節符號（自偶數頁起）',
  oddPage: '分節符號（自奇數頁起）',
  nextColumn: '分節符號（下一欄）',
};

const headingLevels: Record<string, string> = {
  H1: 'Heading1', H2: 'Heading2', H3: 'Heading3', H4: 'Heading4', H5: 'Heading5', H6: 'Heading6',
};

const nodes: Record<string, NodeSpec> = {
  doc: {
    content: 'block+',
    // The body-level w:sectPr (the last section). Kept on the document so page setup and
    // "different first page" of the last section are undoable like any other edit.
    // The comments (docx/comments.ts DocComment[]; null: not loaded, the file's are kept as they
    // are). On the document so adding, replying, resolving, editing and deleting are undoable.
    attrs: { sectPr: { default: null as string | null }, comments: { default: null as unknown[] | null } },
  },

  paragraph: {
    group: 'block',
    content: 'inline*',
    attrs: paragraphAttrs,
    parseDOM: [
      ignoreWordListNumber,
      {
        tag: 'p',
        getAttrs: (dom) => {
          const el = dom as HTMLElement;
          const align = el.style.textAlign || null;
          return { styleId: el.getAttribute('data-style-id'), align: align === 'start' ? null : align, ...listAttrsOfBlock(el) };
        },
      },
      ...Object.entries(headingLevels).map(([tag, styleId]) => ({ tag, attrs: { styleId } })),
      {
        tag: 'li',
        // An <li> holding paragraphs is not a paragraph itself: its first <p> carries the list.
        getAttrs: (dom) => {
          const li = dom as HTMLElement;
          if (Array.from(li.children).some((c) => BLOCK_TAGS.has(c.tagName))) return false;
          return pastedListAttrs(li) ?? {};
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const a = node.attrs;
      const cls = ['dx-p'];
      if (a.styleId) cls.push('dx-ps-' + cssIdent(a.styleId));
      const out: Record<string, string> = { class: cls.join(' ') };
      if (a.styleId) out['data-style-id'] = a.styleId;
      // A paragraph ending a section shows Word's section-break mark.
      if (a.sectPr) {
        out.class += ' dx-sect-end';
        const type = /<w:type\b[^>]*w:val="(\w+)"/.exec(a.sectPr)?.[1] ?? 'nextPage';
        out['data-sect-label'] = tl(SECTION_LABEL[type] ?? SECTION_LABEL.nextPage);
      }
      // Lets a list paragraph copied inside the editor keep its list when pasted.
      if (a.numId) {
        out['data-num-id'] = a.numId;
        out['data-ilvl'] = String(a.ilvl ?? 0);
      }
      // Lines holding East Asian text use the taller of the Latin and East Asian fonts (fontMetrics.ts).
      if (hasEastAsianText(node.textContent)) out.class += ' dx-cjk';
      const run = runFontsLine(node);
      const runLine = run.ratio == null ? '' : run.covered ? `--dx-lh:${run.ratio}` : `--dx-lh-run:${run.ratio}`;
      const style = [paragraphStyle(a), runLine].filter(Boolean).join(';');
      if (style) out.style = style;
      return ['p', out, 0];
    },
  },

  page_break: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: {
      // Word keeps a page break inside a paragraph (w:br w:type="page"). We show it as its
      // own block; these remember how to put it back exactly.
      /** Same value on the paragraphs and breaks that came from one w:p (see paragraph.brGroup). */
      group: { default: null as string | null },
      para: { default: null as string | null }, // JSON paragraph attrs when no text piece of the w:p survived
      rPr: { default: null as string | null }, // raw rPr of the run holding the w:br
      runAttrs: { default: null as string | null }, // that run's own attributes
      /** Inline wrappers around that run (w:ins, w:hyperlink ...), JSON Layer[] as in inlineWrap. */
      layers: { default: null as string | null },
      /** Target of a hyperlink wrapping the run (its layer is rebuilt from it, like the link mark). */
      href: { default: null as string | null },
      ...wrapAttr,
    },
    parseDOM: [{ tag: 'div.dx-page-break' }],
    toDOM: () => ['div', { class: 'dx-page-break', contenteditable: 'false' }],
  },

  /** Body-level OOXML we do not understand; kept verbatim so saving loses nothing. */
  raw_block: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: { xml: { default: '' }, label: { default: 'content' }, hidden: { default: false }, ...wrapAttr },
    parseDOM: [],
    toDOM: (node) =>
      node.attrs.hidden
        ? ['div', { class: 'dx-raw-hidden', contenteditable: 'false' }]
        : ['div', { class: 'dx-raw', contenteditable: 'false' }, `[${tl(node.attrs.label)}]`],
  },

  text: { group: 'inline' },

  /** Inline OOXML we can't edit (charts, shapes, footnote marks ...); written back verbatim. */
  raw_inline: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    attrs: {
      xml: { default: '' },
      label: { default: '' },
      src: { default: null as string | null },
      width: { default: null as number | null },
      height: { default: null as number | null },
      /** Invisible marker (bookmarks, field codes, comment ranges ...). */
      hidden: { default: false },
      /** Show `label` as ordinary text (symbols, special hyphens). */
      plain: { default: false },
      /**
       * A kept deletion: the relationships its content refers to ({ rels, images } by id), so
       * rejecting it gives back working pictures and links.
       */
      refs: { default: null as { rels: Record<string, { target: string; external: boolean }>; images: Record<string, string> } | null },
      /** A text box or shape: what the editor draws and edits (docx/shapes.ts ShapeModel); the XML stays as read. */
      shape: { default: null as unknown },
    },
    // Kept-as-is content can't travel through the clipboard (its XML isn't in the HTML): a copy
    // must not paste its on-screen label ("註解", ...) as ordinary text (GOV-184 / run-187).
    parseDOM: [
      // A shape copied in an editor carries its run (see pasteLists.ts); checked by shapeView.ts.
      { tag: 'span[data-dx-shape-run]', getAttrs: (dom) => shapeParser((dom as HTMLElement).getAttribute('data-dx-shape-run')) ?? false },
      { tag: 'span.dx-raw-inline', ignore: true },
      { tag: 'span.dx-hidden-mark', ignore: true },
      // A tracked deletion (its text is in the page for screen readers): never pasted as text.
      { tag: 'span.dx-rev-del', ignore: true },
      { tag: 'span.dx-shape', ignore: true },
    ],
    toDOM(node): DOMOutputSpec {
      const { src, width, height, label, hidden, plain } = node.attrs;
      // Tracked deletions are shown struck through (see docx/revisions.ts).
      if (label === 'del' || label === 'moveFrom') {
        const del = deletionDOM(node.attrs.xml);
        if (del) return del;
      }
      // A watermark (浮水印) is drawn on the pages (DocxEditor.renderPages): in the header it is a
      // chip naming it, shown while the header is edited and taking no room on the page.
      if (isWatermarkRun(node.attrs.xml)) {
        return ['span', { class: 'dx-raw-inline dx-wm-chip', title: tl('浮水印：要變更或移除，請用「設計」›「浮水印」') }, tl('浮水印')];
      }
      if (hidden) return ['span', { class: 'dx-hidden-mark' }];
      if (plain) return ['span', { class: 'dx-raw-plain' }, label];
      // Text boxes and shapes are drawn by editor/shapeView.ts.
      const shape = node.attrs.shape ? shapeRenderer(node) : null;
      if (shape) return shape;
      if (src) {
        const a: Record<string, string> = { src, class: 'dx-img dx-raw-img', title: tl(label), alt: tl(label) };
        if (width) a.width = String(Math.round(width));
        if (height) a.height = String(Math.round(height));
        return ['img', a];
      }
      return ['span', { class: 'dx-raw-inline', title: tl('此內容無法在這裡編輯，存檔時會原樣保留') }, tl(label)];
    },
  },

  hard_break: {
    inline: true,
    group: 'inline',
    selectable: false,
    attrs: { type: { default: null as string | null }, clear: { default: null as string | null } },
    parseDOM: [{ tag: 'span.dx-col-break', attrs: { type: 'column' } }, { tag: 'br' }],
    // A column break (w:br w:type="column") ends the line and shows Word's 分欄符號 mark: the
    // editor shows one column (in Word the text after it starts the next column).
    toDOM: (node) =>
      node.attrs.type === 'column'
        ? ['span', { class: 'dx-col-break', contenteditable: 'false', 'data-label': tl('分欄符號') }]
        : ['br'],
  },

  tab: {
    inline: true,
    group: 'inline',
    selectable: false,
    parseDOM: [{ tag: 'span.dx-tab' }],
    toDOM: () => ['span', { class: 'dx-tab' }, '\t'],
  },

  image: {
    inline: true,
    group: 'inline',
    draggable: true,
    attrs: {
      src: { default: '' },
      width: { default: null as number | null }, // px
      height: { default: null as number | null },
      alt: { default: '' },
      /** Original w:drawing, written back as-is while the picture and its size are unchanged. */
      xml: { default: null as string | null },
      origSrc: { default: null as string | null },
      origWidth: { default: null as number | null },
      origHeight: { default: null as number | null },
    },
    parseDOM: [
      {
        tag: 'img[src]',
        getAttrs: (dom) => {
          const el = dom as HTMLImageElement;
          // A remote, script or unreadable src is not a picture: dropped (see replaceUnsafeImages).
          if (!isSafeImageSrc(el.getAttribute('src'))) return false;
          return {
            src: el.getAttribute('src'),
            alt: el.getAttribute('alt') || '',
            width: Number(el.getAttribute('width')) || null,
            height: Number(el.getAttribute('height')) || null,
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const { src, alt, width, height } = node.attrs;
      const a: Record<string, string> = { src, alt, class: 'dx-img' };
      if (width) a.width = String(Math.round(width));
      if (height) a.height = String(Math.round(height));
      return ['img', a];
    },
  },

  /** A Word field such as PAGE or NUMPAGES; shows its cached result in the body. */
  field: {
    inline: true,
    group: 'inline',
    atom: true,
    attrs: {
      instr: { default: '' },
      text: { default: '' },
      /** Original w:fldSimple, written back as-is while instr/text are unchanged. */
      xml: { default: null as string | null },
    },
    parseDOM: [
      {
        tag: 'span.dx-field',
        getAttrs: (dom) => ({
          instr: (dom as HTMLElement).getAttribute('data-instr') || '',
          text: (dom as HTMLElement).textContent || '',
        }),
      },
    ],
    toDOM: (node) => [
      'span',
      { class: 'dx-field', 'data-instr': node.attrs.instr, 'data-field': fieldKind(node.attrs.instr), 'data-format': fieldNumberFormat(node.attrs.instr) ?? undefined },
      node.attrs.text || ' ',
    ],
  },

  ...(() => {
    const t = tableNodes({
      tableGroup: 'block',
      cellContent: 'block+',
      cellAttributes: {
        background: {
          default: null,
          getFromDOM: (dom) => pastedFill((dom as HTMLElement).style.backgroundColor || (dom as HTMLElement).getAttribute('bgcolor')),
          setDOMAttr: (value, attrs) => {
            const color = cssColor(value);
            if (color) attrs.style = (attrs.style || '') + `background-color:${color};`;
          },
        },
        vAlign: {
          default: null,
          getFromDOM: (dom) => {
            // Only what Word has (top / center / bottom); baseline, sub ... are left out.
            const v = ((dom as HTMLElement).style.verticalAlign || (dom as HTMLElement).getAttribute('valign') || '').toLowerCase();
            return v === 'top' || v === 'bottom' ? v : v === 'middle' || v === 'center' ? 'middle' : null;
          },
          setDOMAttr: (value, attrs) => {
            if (value === 'top' || value === 'middle' || value === 'bottom') attrs.style = (attrs.style || '') + `vertical-align:${value};`;
          },
        },
        /** Sides set on the cell itself (w:tcBorders); see CellBorders in docx/props.ts. */
        borders: {
          default: null,
          // Only an Excel cell's borders, as inlineExcelStyles found them (pasteExcel.ts).
          getFromDOM: (dom) => pastedBorders((dom as HTMLElement).getAttribute('data-dx-borders')),
          setDOMAttr: (value, attrs) => {
            const css = bordersCss(value as Record<string, BorderAttr> | null);
            if (css) attrs.style = (attrs.style || '') + css;
          },
        },
        // Lossless round-trip: original w:tcPr, and those of the cells merged into this one.
        tcPr: {
          default: null,
          getFromDOM: () => null,
          // The cell's own margins (w:tcMar) replace the table's.
          setDOMAttr: (value, attrs) => {
            const css = tcMarginCss(value as string | null);
            if (css) attrs.style = (attrs.style || '') + css;
          },
        },
        /** Raw w:tc of the cells vertically merged into this one (rows below), in order. */
        mergedTc: { default: null, getFromDOM: () => null, setDOMAttr: () => {} },
        /** An empty cell we added to fill a ragged row; not written back unless edited. */
        filler: { default: false, getFromDOM: () => false, setDOMAttr: () => {} },
      },
    });
    // Extra table-level attributes so the table's look survives a round trip.
    t.table = {
      ...t.table,
      attrs: {
        styleId: { default: null as string | null },
        tblPr: { default: null as string | null }, // raw w:tblPr XML
        grid: { default: null as number[] | null }, // original w:gridCol widths (twips)
        tblGridXml: { default: null as string | null }, // original w:tblGrid, reused when widths are unchanged
        ...wrapAttr,
      },
      toDOM: () => ['table', { class: 'dx-table' }, ['tbody', 0]],
    };
    t.table_row = {
      ...t.table_row,
      attrs: {
        trPr: { default: null as string | null },
        trAttrs: { default: null as string | null },
        /** Raw tblPrEx, kept verbatim. */
        tblPrEx: { default: null as string | null },
        // Modeled trPr properties (docx/props.ts RowModel); written by patching trPr.
        height: { default: null as number | null }, // twips
        heightRule: { default: null as string | null }, // atLeast | exact | auto; null = omitted (at least)
        header: { default: false }, // w:tblHeader
        cantSplit: { default: false }, // w:cantSplit
      },
      toDOM(node): DOMOutputSpec {
        const { height, heightRule, header, cantSplit } = node.attrs;
        const out: Record<string, string> = {};
        if (height && heightRule !== 'auto') out.style = `height:${twipsToPx(height)}px`;
        const cls = [header && 'dx-tr-header', heightRule === 'exact' && 'dx-tr-exact', cantSplit && 'dx-tr-keep'].filter(Boolean);
        if (cls.length) out.class = cls.join(' ');
        if (header) out.title = tl('標題列（跨頁重複）');
        return ['tr', out, 0];
      },
    };
    return t;
  })(),
};

interface BorderAttr {
  val: string;
  sz: number | null;
  color: string | null;
}

const BORDER_CSS_STYLE: Record<string, string> = {
  double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed',
  dotDotDash: 'dotted', inset: 'inset', outset: 'outset',
};

/** CSS for the border sides a cell sets itself (w:tcBorders). */
export function bordersCss(b: Record<string, BorderAttr> | null): string {
  if (!b) return '';
  let css = '';
  for (const side of ['top', 'left', 'bottom', 'right']) {
    const s = b[side];
    if (!s) continue;
    if (s.val === 'nil' || s.val === 'none') {
      css += `border-${side}:none;`;
      continue;
    }
    // sz is in eighths of a point. Only known styles, numbers and hex colours reach the CSS.
    const style = Object.prototype.hasOwnProperty.call(BORDER_CSS_STYLE, s.val) ? BORDER_CSS_STYLE[s.val] : 'solid';
    const sz = typeof s.sz === 'number' && Number.isFinite(s.sz) ? Math.min(Math.max(s.sz, 0), 96) : 4;
    const px = Math.max(style === 'double' ? 3 : 1, Math.round((sz / 8) * (96 / 72)));
    const color = s.color && /^[0-9A-Fa-f]{6}$/.test(s.color) ? '#' + s.color : '#000';
    css += `border-${side}:${px}px ${style} ${color};`;
  }
  return css;
}

/** A colour safe to put in a style attribute (#hex, rgb()/rgba(), a colour name); null otherwise. */
export function cssColor(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const v = value.trim();
  return /^#(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(v) ||
    /^rgba?\(\s*[\d.]+%?\s*,\s*[\d.]+%?\s*,\s*[\d.]+%?\s*(?:,\s*[\d.]+%?\s*)?\)$/.test(v) ||
    /^[a-zA-Z]{3,30}$/.test(v)
    ? v
    : null;
}

/** A pasted cell background as "#RRGGBB"; transparent (alpha 0) or unreadable is no shading. */
function pastedFill(value: string | null): string | null {
  const c = (value || '').trim().toLowerCase();
  const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(c);
  if (h) return '#' + (h[1].length === 3 ? h[1].replace(/./g, '$&$&') : h[1]).toUpperCase();
  const m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)(%?)\s*)?\)$/.exec(c);
  if (!m || (m[4] != null && Number(m[4]) === 0)) return null;
  return '#' + [m[1], m[2], m[3]].map((x) => Math.min(255, Number(x)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Word's highlight colour names, lower case (what docx/props.ts accepts from a file). */
const WORD_COLOR_NAMES = new Map(Object.entries(HIGHLIGHT_COLORS).map(([name, hex]) => [name.toLowerCase(), hex]));

/**
 * A pasted text or highlight colour as "#rrggbb", as the reader takes colours from a file: a
 * hex or rgb() colour, or one of Word's highlight colour names (yellow, darkBlue ...). Anything
 * else (transparent, inherit, system colours, other names, broken values) is no colour.
 */
export function pastedColor(value: string | null | undefined): string | null {
  const c = (value || '').trim().toLowerCase();
  const named = WORD_COLOR_NAMES.get(c);
  if (named) return named;
  const fill = pastedFill(c);
  return fill ? fill.toLowerCase() : null;
}

/** A font name that cannot leave its quotes in a style attribute. */
const cssFontName = (name: string): string => String(name).replace(/["'\\;<>{}\u0000-\u001f]/g, '');

/**
 * CSS per run properties, least recently used first: documents repeat a few hundred of them,
 * so the cache stays small, and never grows past RUN_CSS_MAX whatever is opened.
 */
const runCssCache = new Map<string, string>();
const RUN_CSS_MAX = 4000;
/** CSS for run properties the editor doesn't model as marks (caps, spacing, hidden ...). */
/** A copied shape's run as raw_inline attrs (null when it isn't one that may be pasted); set by editor/shapeView.ts. */
let shapeParser: (xml: string | null) => Record<string, unknown> | null = () => null;
export function setShapeParser(fn: (xml: string | null) => Record<string, unknown> | null): void {
  shapeParser = fn;
}

/** Draws a text box or shape (raw_inline with a shape model); set by editor/shapeView.ts. */
let shapeRenderer: (node: PMNode) => HTMLElement | null = () => null;
export function setShapeRenderer(fn: (node: PMNode) => HTMLElement | null): void {
  shapeRenderer = fn;
}

export let runExtraCss: (rPr: string) => string = () => '';
export function setRunExtraCss(fn: (rPr: string) => string): void {
  runExtraCss = fn;
  runCssCache.clear();
}

function runCss(rPr: string): string {
  let css = runCssCache.get(rPr);
  if (css != null) {
    runCssCache.delete(rPr);
  } else {
    css = runExtraCss(rPr);
    if (runCssCache.size >= RUN_CSS_MAX) runCssCache.delete(runCssCache.keys().next().value!);
  }
  runCssCache.set(rPr, css);
  return css;
}

/** How many run properties have their CSS cached (for tests). */
export const runCssCacheSize = (): number => runCssCache.size;

const marks: Record<string, MarkSpec> = {
  /**
   * The original w:r it came from: raw w:rPr and the run's own attributes.
   * Every other formatting mark is compared against this on save.
   */
  run: {
    attrs: { rPr: { default: null as string | null }, attrs: { default: null as string | null } },
    parseDOM: [],
    toDOM: (m) => {
      const rPr = m.attrs.rPr as string | null;
      if (!rPr) return ['span', 0];
      const css = runCss(rPr);
      // A run with w:hint="eastAsia" is marked: its shared symbols take the East Asian font
      // (docx/eastAsiaFonts.ts HINT_RULE). Only the run's own w:rPr is read, not its styles'.
      const attrs: Record<string, string> = {};
      if (css) attrs.style = css;
      if (hasEastAsiaHint(rPr)) attrs['data-ea-hint'] = '';
      return Object.keys(attrs).length ? ['span', attrs, 0] : ['span', 0];
    },
  },
  /** Text that is the displayed result of a Word field (PAGE, TOC entry ...). */
  fieldResult: {
    attrs: { instr: { default: '' } },
    parseDOM: [],
    // data-format: a \* switch (PAGE \* roman) that overrides the section's page number format.
    toDOM: (m) => ['span', { class: 'dx-field-result', 'data-field': fieldKind(m.attrs.instr), 'data-format': fieldNumberFormat(m.attrs.instr) ?? undefined }, 0],
  },
  /** Inline wrappers (w:sdt, w:ins, w:smartTag ...), outermost first; see docx/wrappers.ts. */
  inlineWrap: {
    attrs: { layers: { default: '[]' } },
    parseDOM: [],
    toDOM: (m) => {
      const ins = insertionAttrs(m.attrs.layers);
      return ins ? ['span', ins, 0] : ['span', 0];
    },
  },
  link: {
    attrs: { href: { default: '' }, attrs: { default: null as string | null } },
    inclusive: false,
    // A pasted link's refused target (javascript: ... see docx/links.ts) is not taken in: the
    // link keeps no target, shows as "#" and is saved as plain text.
    parseDOM: [
      {
        tag: 'a[href]',
        getAttrs: (d) => {
          const href = (d as HTMLElement).getAttribute('href');
          return { href: isSafeHref(href) ? href : '' };
        },
      },
    ],
    toDOM: (m) => ['a', { href: safeHref(m.attrs.href), rel: 'noopener noreferrer', target: '_blank' }, 0],
  },
  // Mark order is nesting order in the page: the character style wraps the direct
  // formatting below it, so a direct "not bold" inside a bold character style wins, as in Word.
  /** A Word character style (w:rStyle). */
  charStyle: {
    attrs: { id: {} },
    parseDOM: [{ tag: 'span[data-char-style]', getAttrs: (d) => ({ id: (d as HTMLElement).getAttribute('data-char-style') }) }],
    toDOM: (m) => ['span', { 'data-char-style': m.attrs.id, class: 'dx-cs-' + cssIdent(m.attrs.id) }, 0],
  },
  /** `on: false` is Word's explicit "not bold" (w:b w:val="0"), used to cancel a bold style. */
  bold: {
    attrs: { on: { default: true } },
    parseDOM: [
      { tag: 'strong' },
      { tag: 'b', getAttrs: (d) => (d as HTMLElement).style.fontWeight !== 'normal' && null },
      { style: 'font-weight', getAttrs: (v) => /^(bold(er)?|[6-9]\d\d)$/.test(v as string) && null },
    ],
    toDOM: (m) => (m.attrs.on === false ? ['span', { style: 'font-weight:normal' }, 0] : ['strong', 0]),
  },
  italic: {
    attrs: { on: { default: true } },
    parseDOM: [{ tag: 'i' }, { tag: 'em' }, { style: 'font-style=italic' }],
    toDOM: (m) => (m.attrs.on === false ? ['span', { style: 'font-style:normal' }, 0] : ['em', 0]),
  },
  underline: {
    parseDOM: [{ tag: 'u' }, { style: 'text-decoration-line=underline' }],
    toDOM: () => ['u', 0],
  },
  strike: {
    parseDOM: [{ tag: 's' }, { tag: 'del' }, { style: 'text-decoration-line=line-through' }],
    toDOM: () => ['s', 0],
  },
  superscript: {
    excludes: 'subscript superscript',
    parseDOM: [{ tag: 'sup' }],
    toDOM: () => ['sup', 0],
  },
  subscript: {
    excludes: 'subscript superscript',
    parseDOM: [{ tag: 'sub' }],
    toDOM: () => ['sub', 0],
  },
  color: {
    attrs: { color: {} },
    parseDOM: [
      {
        style: 'color',
        getAttrs: (v) => {
          const color = pastedColor(v as string);
          return color ? { color } : false;
        },
      },
    ],
    toDOM: (m) => {
      const color = cssColor(m.attrs.color);
      return color ? ['span', { style: `color:${color}` }, 0] : ['span', 0];
    },
  },
  highlight: {
    attrs: { color: {} },
    parseDOM: [
      {
        tag: 'mark',
        getAttrs: (d) => {
          // A <mark> without a colour is Word's default yellow; an unreadable colour is no highlight.
          const color = pastedColor((d as HTMLElement).style.backgroundColor || 'yellow');
          return color ? { color } : false;
        },
      },
    ],
    toDOM: (m) => {
      const color = cssColor(m.attrs.color);
      return color ? ['mark', { style: `background-color:${color}` }, 0] : ['mark', 0];
    },
  },
  /** `family` is the Latin font (w:ascii), `eastAsia` the CJK font (w:eastAsia). */
  font: {
    attrs: { family: { default: null }, eastAsia: { default: null } },
    parseDOM: [
      {
        style: 'font-family',
        getAttrs: (v) => {
          // "var(--dx-font-…)" in our own output stands for a font the run does not set.
          const list = (v as string).split(',').map((f) => f.replace(/["']/g, '').trim()).map((f) => (f.startsWith('var(') ? '' : f));
          if (!list[0] && !list[1]) return false;
          return { family: list[0] || null, eastAsia: list[1] && !/serif|monospace/.test(list[1]) ? list[1] : null };
        },
      },
    ],
    toDOM: (m) => {
      // Word uses w:ascii for Latin text and w:eastAsia for CJK text, each inherited on its
      // own: a run naming one font keeps the paragraph's other one (docx/styles.ts FONT_FAMILY).
      const latin = m.attrs.family ? cssFontName(m.attrs.family) : '';
      const eastAsia = m.attrs.eastAsia ? cssFontName(m.attrs.eastAsia) : '';
      if (!latin && !eastAsia) return ['span', 0];
      // The Chinese font comes with its other names and similar fonts, for computers without it
      // (docx/fonts.ts), and its alias for hinted runs (docx/eastAsiaFonts.ts).
      if (eastAsia) useEastAsiaFont(eastAsia);
      const eaStack = eastAsia ? fontStack(eastAsia, true) : '';
      const style = [
        latin ? `--dx-font-l:"${latin}"` : '',
        eastAsia ? `--dx-font-e:${eaStack};--dx-font-h:"${eastAsiaAlias(eastAsia)}"` : '',
        `font-family:${latin ? `"${latin}"` : 'var(--dx-font-l)'},${eastAsia ? eaStack : 'var(--dx-font-e)'}`,
      ];
      return ['span', { style: style.filter(Boolean).join(';') }, 0];
    },
  },
  fontSize: {
    attrs: { pt: {} },
    parseDOM: [
      {
        style: 'font-size',
        getAttrs: (v) => {
          const m = /^([\d.]+)pt$/.exec(v as string);
          return m ? { pt: Number(m[1]) } : false;
        },
      },
    ],
    toDOM: (m) => ['span', { style: `font-size:${m.attrs.pt}pt` }, 0],
  },
};

export const schema = new Schema({ nodes, marks });

export function cssIdent(s: string): string {
  return s.replace(/[^a-zA-Z0-9_-]/g, '_');
}

export function fieldKind(instr: string): string {
  return (instr.trim().split(/\s+/)[0] || '').toUpperCase();
}

/** Only let harmless URL schemes become clickable in the page (shared allowlist: docx/links.ts). */
export { safeHref };
