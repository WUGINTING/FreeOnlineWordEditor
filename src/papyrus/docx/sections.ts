// Document sections (w:sectPr). A section break is a w:sectPr inside the last paragraph of a
// section; the body-level w:sectPr describes the last section. Each section has its own page
// size, margins, "different first page" and header/footer references; a section without a
// reference of some type shows the previous section's (Word's "link to previous").
import type { Node as PMNode } from 'prosemirror-model';
import { NS, attr, child, children, numAttr, onOff, parseFragment, serializeXml } from './xml';
import {
  A4, type DocxModel, type HeaderFooterKind, type HeaderFooterPart, type HeaderFooterType, type PageSetup,
} from './model';

export type SectionStart = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

export type Refs = Record<HeaderFooterKind, Partial<Record<HeaderFooterType, string>>>;

/**
 * A section's text columns (w:cols), as Word lays the section out. The editor itself shows one
 * column (GOV-FINDING-014): these are set for Word, which shows and prints them.
 */
export interface SectionColumns {
  /** How many columns (1: no columns). */
  count: number;
  /** Space between columns in twips (w:space; Word reads a missing value as 720). */
  space: number;
  /** A line between the columns (w:sep). */
  separator: boolean;
  /** All columns the same width; false when the file sets each column's width (w:equalWidth="0"). */
  equalWidth: boolean;
}

/** Word's reading of w:cols/@w:space when it is missing (36 pt; read from Word 2016+). */
export const DEFAULT_COLUMN_SPACE = 720;
/** Most columns Word allows in a section. */
export const MAX_COLUMNS = 45;

interface ParsedSection {
  page: PageSetup;
  titlePage: boolean;
  start: SectionStart;
  /** w:pgNumType/@w:start: page numbering restarts here. */
  pageNumberStart: number | null;
  /** w:pgNumType/@w:fmt: how PAGE fields show this section's page numbers (null: decimal). */
  pageNumberFormat: string | null;
  /** w:cols: the section's text columns (in Word; the editor shows one column). */
  columns: SectionColumns;
  /** This section's own header/footer references (relationship ids). */
  refs: Refs;
}

export interface Section extends ParsedSection {
  index: number;
  /** Position of the paragraph ending this section; null for the last (body-level) section. */
  pos: number | null;
  /** Top-level blocks of the section: indices into the document's children. */
  firstBlock: number;
  lastBlock: number;
  /** The w:sectPr XML (null only for a document without one). */
  xml: string | null;
}

/** Word's page size range: 0.1" to 22" (twips). */
export const MIN_PAGE = 144;
export const MAX_PAGE = 31680;

/**
 * Page setup of a w:sectPr as used for layout. Values are clamped to what Word allows: page
 * size 144..31680 twips, margins and header/footer distances not negative, and the margins
 * leave some of the page for text (a negative or zero-height page would stall pagination).
 * Only the layout values are clamped: the w:sectPr itself is saved as written unless the user
 * changes the page setup.
 */
export function parsePage(sectPr: Element | null): PageSetup {
  if (!sectPr) return { ...A4 };
  const sz = child(sectPr, 'pgSz');
  const mar = child(sectPr, 'pgMar');
  const size = (v: number) => Math.min(MAX_PAGE, Math.max(MIN_PAGE, v));
  const len = (v: number, max: number) => Math.min(max, Math.max(0, v));
  const width = size(numAttr(sz, 'w') ?? A4.width);
  const height = size(numAttr(sz, 'h') ?? A4.height);
  const [marginTop, marginBottom] = fitMargins(
    len(Math.abs(numAttr(mar, 'top') ?? A4.marginTop), height),
    len(Math.abs(numAttr(mar, 'bottom') ?? A4.marginBottom), height),
    height,
  );
  const [marginLeft, marginRight] = fitMargins(
    len(numAttr(mar, 'left') ?? A4.marginLeft, width),
    len(numAttr(mar, 'right') ?? A4.marginRight, width),
    width,
  );
  return {
    width,
    height,
    marginTop,
    marginBottom,
    marginLeft,
    marginRight,
    header: len(numAttr(mar, 'header') ?? A4.header, height),
    footer: len(numAttr(mar, 'footer') ?? A4.footer, height),
  };
}

/** Two opposite margins, shrunk in proportion when they leave less than 0.1" of the page for text. */
function fitMargins(a: number, b: number, page: number): [number, number] {
  const room = page - MIN_PAGE;
  if (a + b <= room) return [a, b];
  const k = room / (a + b);
  return [Math.floor(a * k), Math.floor(b * k)];
}

/** An ST_OnOff attribute: 1 / true / on (null when missing). */
function onOffAttr(el: Element | null, local: string): boolean | null {
  const v = attr(el, local);
  return v == null ? null : v === '1' || v === 'true' || v === 'on';
}

/** A w:cols element as Word reads it: w:num columns, or one per w:col when the widths are the file's own. */
function parseColumns(cols: Element | null): SectionColumns {
  const equalWidth = onOffAttr(cols, 'equalWidth') !== false;
  const own = cols ? children(cols, 'col').length : 0;
  const num = Math.trunc(numAttr(cols, 'num') ?? 1);
  const count = !equalWidth && own > 0 ? own : num;
  return {
    count: Math.min(MAX_COLUMNS, Math.max(1, Number.isFinite(count) ? count : 1)),
    space: Math.max(0, Math.round(numAttr(cols, 'space') ?? DEFAULT_COLUMN_SPACE)),
    separator: onOffAttr(cols, 'sep') === true,
    equalWidth,
  };
}

const parsed = new Map<string, ParsedSection>();

function parseSection(xml: string | null): ParsedSection {
  const key = xml ?? '';
  let p = parsed.get(key);
  if (p) return p;
  let el: Element | null = null;
  try {
    el = xml ? parseFragment(xml) : null;
  } catch {
    el = null;
  }
  const refs: Refs = { header: {}, footer: {} };
  for (const kind of ['header', 'footer'] as const) {
    for (const r of el ? children(el, kind + 'Reference') : []) {
      const type = (attr(r, 'type') ?? 'default') as HeaderFooterType;
      const id = attr(r, 'id', NS.r);
      if (id && (type === 'default' || type === 'first' || type === 'even')) refs[kind][type] = id;
    }
  }
  p = {
    page: parsePage(el),
    titlePage: onOff(child(el, 'titlePg')) === true,
    start: (attr(child(el, 'type'), 'val') as SectionStart | null) ?? 'nextPage',
    pageNumberStart: numAttr(child(el, 'pgNumType'), 'start'),
    pageNumberFormat: attr(child(el, 'pgNumType'), 'fmt'),
    columns: parseColumns(child(el, 'cols')),
    refs,
  };
  if (parsed.size > 500) parsed.clear();
  parsed.set(key, p);
  return p;
}

/** A w:sectPr's own header/footer references (relationship ids by kind and type). */
export function sectionRefs(xml: string | null): Refs {
  return parseSection(xml).refs;
}

/** The sections of a document, in order; the last one is the body-level w:sectPr (model.sectPr). */
export function documentSections(doc: PMNode, model: DocxModel): Section[] {
  const out: Section[] = [];
  let first = 0;
  doc.forEach((node, offset, index) => {
    const xml = node.type.name === 'paragraph' ? (node.attrs.sectPr as string | null) : null;
    if (!xml) return;
    out.push({ ...parseSection(xml), index: out.length, pos: offset, firstBlock: first, lastBlock: index, xml });
    first = index + 1;
  });
  // The last section: its page and "different first page" live on the model, which the
  // editor changes directly (and the writer patches into model.sectPr).
  const last = parseSection(model.sectPr);
  out.push({
    ...last,
    page: model.page,
    titlePage: model.titlePage,
    index: out.length,
    pos: null,
    firstBlock: first,
    lastBlock: doc.childCount - 1,
    xml: model.sectPr,
  });
  return out;
}

/** The section a top-level block (by child index) belongs to. */
export function sectionOfBlock(sections: Section[], block: number): Section {
  return sections.find((s) => block <= s.lastBlock) ?? sections[sections.length - 1];
}

/**
 * The header/footer a section shows for a variant: its own reference, else the nearest earlier
 * section's (link to previous). Parts made in the editor for the last section have no
 * relationship id until saved.
 */
export function sectionHeaderFooter(
  model: DocxModel,
  sections: Section[],
  index: number,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): HeaderFooterPart | undefined {
  const last = sections.length - 1;
  for (let i = Math.min(index, last); i >= 0; i--) {
    const relId = sections[i].refs[kind][type];
    if (relId) {
      const hf = model.headerFooters.find((h) => h.relId === relId && h.kind === kind);
      if (hf) return hf;
    }
    if (i === last) {
      const made = model.headerFooters.find((h) => !h.relId && !h.pending && h.kind === kind && h.type === type);
      if (made) return made;
    }
  }
  return undefined;
}

/** Which variant Word shows on a page: `inSection` is the page's index inside its section, `number` its page number. */
export function variantFor(section: Section, inSection: number, number: number, evenAndOdd: boolean): HeaderFooterType {
  if (inSection === 0 && section.titlePage) return 'first';
  if (evenAndOdd && number % 2 === 0) return 'even';
  return 'default';
}

/** Whether a section begins on a new page (a continuous section with the same page size flows on). */
export function startsNewPage(section: Section, previous: Section): boolean {
  if (section.start !== 'continuous') return true;
  const a = section.page;
  const b = previous.page;
  return a.width !== b.width || a.height !== b.height;
}

// ----- editing a w:sectPr -----

/** Word's order of w:sectPr children, to insert new ones in the right place. */
const SECT_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc',
  'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
  'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange',
];

function ensureChild(sect: Element, local: string): Element {
  const existing = child(sect, local);
  if (existing) return existing;
  const el = sect.ownerDocument.createElementNS(NS.w, 'w:' + local);
  const rank = SECT_ORDER.indexOf(local);
  const after = Array.from(sect.children).find((c) => c.namespaceURI === NS.w && SECT_ORDER.indexOf(c.localName) > rank);
  sect.insertBefore(el, after ?? null);
  return el;
}

const EMPTY_SECT = `<w:sectPr xmlns:w="${NS.w}" xmlns:r="${NS.r}"/>`;

/** The section properties with a new page size, orientation and margins; everything else kept. */
export function withPageSetup(xml: string | null, page: PageSetup): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const sz = ensureChild(sect, 'pgSz');
  sz.setAttributeNS(NS.w, 'w:w', String(Math.round(page.width)));
  sz.setAttributeNS(NS.w, 'w:h', String(Math.round(page.height)));
  if (page.width > page.height) sz.setAttributeNS(NS.w, 'w:orient', 'landscape');
  else sz.removeAttributeNS(NS.w, 'orient');
  const mar = ensureChild(sect, 'pgMar');
  const set = (name: string, v: number) => mar.setAttributeNS(NS.w, 'w:' + name, String(Math.round(v)));
  set('top', page.marginTop);
  set('right', page.marginRight);
  set('bottom', page.marginBottom);
  set('left', page.marginLeft);
  set('header', page.header);
  set('footer', page.footer);
  if (!mar.hasAttributeNS(NS.w, 'gutter')) set('gutter', 0);
  return serializeXml(sect);
}

/** The section properties with "different first page" switched on or off. */
export function withTitlePage(xml: string | null, on: boolean): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const t = child(sect, 'titlePg');
  if (on && !t) ensureChild(sect, 'titlePg');
  else if (!on && t) sect.removeChild(t);
  return serializeXml(sect);
}

/** A section's page numbering: its format, and where it restarts (null: continues from the previous section). */
export interface PageNumbering {
  format: string | null;
  start: number | null;
}

/**
 * The section properties with a new page number format and/or start (Word's 頁碼格式): only the
 * given values change, the rest of w:pgNumType (chapter numbering ...) is kept. A format equal to
 * the current one (null and "decimal" are the same) leaves the attribute as written.
 */
export function withPageNumbering(xml: string | null, numbering: Partial<PageNumbering>): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const current = child(sect, 'pgNumType');
  const fmtNow = attr(current, 'fmt') ?? 'decimal';
  const startNow = numAttr(current, 'start');
  const fmt = numbering.format === undefined ? undefined : numbering.format ?? 'decimal';
  const setFormat = fmt !== undefined && fmt !== fmtNow;
  const setStart = numbering.start !== undefined && numbering.start !== startNow;
  // Nothing to change: the section's properties stay exactly as written.
  if (!setFormat && !setStart) return xml ?? serializeXml(sect);
  const el = ensureChild(sect, 'pgNumType');
  if (setFormat) {
    if (fmt === 'decimal') el.removeAttributeNS(NS.w, 'fmt');
    else el.setAttributeNS(NS.w, 'w:fmt', fmt!);
  }
  if (setStart) {
    if (numbering.start == null) el.removeAttributeNS(NS.w, 'start');
    else el.setAttributeNS(NS.w, 'w:start', String(Math.max(0, Math.trunc(numbering.start))));
  }
  if (!el.attributes.length && !el.childNodes.length) sect.removeChild(el);
  return serializeXml(sect);
}

/**
 * The section properties with a new start (Word's 版面設定 > 節的起始位置, w:type): where this
 * section begins after the one before it. 新頁 (nextPage) is Word's default and has no w:type.
 * A section already starting that way keeps its w:sectPr as written.
 */
export function withSectionStart(xml: string | null, start: SectionStart): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const type = child(sect, 'type');
  if (((attr(type, 'val') as SectionStart | null) ?? 'nextPage') === start) return xml ?? serializeXml(sect);
  if (start === 'nextPage') sect.removeChild(type!);
  else ensureChild(sect, 'type').setAttributeNS(NS.w, 'w:val', start);
  return serializeXml(sect);
}

/**
 * The section properties with new text columns (Word's 版面配置 > 欄): `count` columns of equal
 * width, `space` twips apart, with or without a line between them (w:cols w:num / w:space / w:sep).
 * Only the given values change and the rest of w:cols is kept; columns set to the values they
 * already have leave the w:sectPr as written. Changing the count or spacing of columns whose
 * widths the file sets one by one makes them equal widths (the editor offers no other).
 */
export function withColumns(xml: string | null, columns: Partial<Omit<SectionColumns, 'equalWidth'>>): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const now = parseColumns(child(sect, 'cols'));
  const count = columns.count === undefined ? undefined : Math.min(MAX_COLUMNS, Math.max(1, Math.trunc(columns.count)));
  const space = columns.space === undefined ? undefined : Math.max(0, Math.round(columns.space));
  const setCount = count !== undefined && count !== now.count;
  const setSpace = space !== undefined && space !== now.space;
  // One column has no line between columns.
  const separator = (count ?? now.count) === 1 ? false : columns.separator;
  const setSeparator = separator !== undefined && separator !== now.separator;
  if (!setCount && !setSpace && !setSeparator) return xml ?? serializeXml(sect);
  const el = ensureChild(sect, 'cols');
  const toEqual = (setCount || setSpace) && !now.equalWidth;
  if (toEqual) {
    el.removeAttributeNS(NS.w, 'equalWidth');
    for (const c of children(el, 'col')) el.removeChild(c);
  }
  const n = count ?? now.count;
  if (setCount || toEqual) {
    if (n > 1) el.setAttributeNS(NS.w, 'w:num', String(n));
    else el.removeAttributeNS(NS.w, 'num');
  }
  // Written out whenever there are columns, as Word does (a missing w:space reads as 720).
  if (setSpace || (n > 1 && !el.hasAttributeNS(NS.w, 'space'))) el.setAttributeNS(NS.w, 'w:space', String(space ?? now.space));
  if (setSeparator) {
    if (separator) el.setAttributeNS(NS.w, 'w:sep', '1');
    else el.removeAttributeNS(NS.w, 'sep');
  }
  if (!el.attributes.length && !el.childNodes.length) sect.removeChild(el);
  return serializeXml(sect);
}

/** The section properties pointing a header/footer variant at a relationship id. */
export function withReference(xml: string | null, kind: HeaderFooterKind, type: HeaderFooterType, relId: string): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const name = kind + 'Reference';
  let ref = children(sect, name).find((r) => (attr(r, 'type') ?? 'default') === type);
  if (!ref) {
    ref = sect.ownerDocument.createElementNS(NS.w, 'w:' + name);
    // headerReference* then footerReference* come first.
    const allowed = kind === 'header' ? ['headerReference'] : ['headerReference', 'footerReference'];
    const last = Array.from(sect.children).filter((c) => c.namespaceURI === NS.w && allowed.includes(c.localName)).pop();
    sect.insertBefore(ref, last ? last.nextSibling : sect.firstChild);
  }
  for (const a of Array.from(ref.attributes)) if (a.namespaceURI === NS.r) ref.removeAttributeNode(a);
  ref.setAttributeNS(NS.w, 'w:type', type);
  ref.setAttributeNS(NS.r, 'r:id', relId);
  return serializeXml(sect);
}

/** A w:sectPr without its own header/footer of this variant: the section shows the previous section's (link to previous). */
export function withoutReference(xml: string | null, kind: HeaderFooterKind, type: HeaderFooterType): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  for (const ref of children(sect, kind + 'Reference')) {
    if ((attr(ref, 'type') ?? 'default') === type) sect.removeChild(ref);
  }
  return serializeXml(sect);
}

/**
 * The properties of a section that now starts at a new section break: it starts on a new
 * page (no w:type = next page) and continues the page numbering of the section before it.
 */
export function startingAtBreak(xml: string | null): string {
  const sect = parseFragment(xml || EMPTY_SECT);
  const type = child(sect, 'type');
  if (type) sect.removeChild(type);
  const numbers = child(sect, 'pgNumType');
  if (numbers) {
    numbers.removeAttributeNS(NS.w, 'start');
    if (!numbers.attributes.length) sect.removeChild(numbers);
  }
  return serializeXml(sect);
}

/** The section's page setup and "different first page" written into its w:sectPr. */
export function lastSectionXml(xml: string | null, page: PageSetup, titlePage: boolean): string {
  return withTitlePage(withPageSetup(xml, page), titlePage);
}

/** Page setup and "different first page" read back from a w:sectPr. */
export function readSectPr(xml: string): { page: PageSetup; titlePage: boolean } {
  const p = parseSection(xml);
  return { page: { ...p.page }, titlePage: p.titlePage };
}

/** Relationship ids that some section refers to (in the body's and paragraphs' w:sectPr). */
export function referencedIds(doc: PMNode, lastSectPr: string | null): Set<string> {
  const ids = new Set<string>();
  const add = (xml: string | null) => {
    if (!xml) return;
    for (const m of xml.matchAll(/="([^"]+)"/g)) ids.add(m[1]);
  };
  add(lastSectPr);
  doc.forEach((node) => add(node.attrs.sectPr ?? null));
  return ids;
}

/** A section break's properties: the given section's, so both halves keep its page and headers. */
export function sectionBreakXml(section: Section): string {
  const sect = parseFragment(section.xml || EMPTY_SECT);
  if (section.pos == null) {
    // The last section's page / first-page setting live on the model: write them in.
    return withTitlePage(withPageSetup(serializeXml(sect), section.page), section.titlePage);
  }
  return serializeXml(sect);
}
