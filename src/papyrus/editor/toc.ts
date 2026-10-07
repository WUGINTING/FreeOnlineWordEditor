// 參考資料 › 目錄 (persona-300): a table of contents as Word writes one, and 「更新目錄」 that
// rebuilds its entries from the headings, as Word's F9 「更新整個目錄」 does.
//
// What Word writes (and what is written here): a TOC field `TOC \o "1-3" \h \z \u` whose result
// is one paragraph per heading of levels 1-3, in the paragraph styles 「目錄 1」-「目錄 3」 (w:name
// "toc 1" …, added to styles.xml when the file has none, see docx/tocStyles.ts) with a right tab
// with a dot leader at the right margin; each entry is a link to a bookmark `_Toc…` around its
// heading, and its page number a PAGEREF field to that bookmark. The field's end is in a paragraph
// of its own after the entries. Page numbers come from the editor's page layout.
import type { Node as PMNode } from 'prosemirror-model';
import { Fragment } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import { schema } from './schema';
import { Converter } from '../docx/convert';
import { NS, child, escapeXml, parseXml } from '../docx/xml';
import type { Numbering, ParagraphStyleInfo } from '../docx/model';
import type { StyleInheritance } from '../docx/inheritance';
import { ListCounter } from '../docx/numbering';
import { documentOutline } from './outline';

/** The field Word's 參考資料 › 目錄 inserts: heading levels 1-3, links, no page numbers on the web, outline levels. */
export const TOC_INSTR = 'TOC \\o "1-3" \\h \\z \\u';
/** Word's result when the document has no headings. */
export const NO_ENTRIES = '找不到目錄項目。';
/** The title above a new table of contents (Word's 自動目錄 has one, in the style 「目錄標題」). */
export const TOC_TITLE = '目錄';

const FLD_CHAR = /<(?:[\w.-]+:)?fldChar\b[^>]*?\bw:fldCharType="(begin|separate|end)"/;
const INSTR_TEXT = /<(?:[\w.-]+:)?instrText\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?instrText>/;
const BOOKMARK_START = /<(?:[\w.-]+:)?bookmarkStart\b[^>]*\bw:name="([^"]+)"/;

/** A complex field (w:fldChar begin … end) of the body, by the positions of its field characters. */
export interface BodyField {
  instr: string;
  begin: number;
  separate: number | null;
  end: number;
  /** Nested in another field. */
  nested: boolean;
}

/** The complex fields of a document, in the order they begin. */
export function bodyFields(doc: PMNode): BodyField[] {
  const out: BodyField[] = [];
  const open: { f: Partial<BodyField> & { instr: string; begin: number }; phase: 'instr' | 'result' }[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.raw_inline) return true;
    const xml = node.attrs.xml as string;
    const ch = FLD_CHAR.exec(xml)?.[1];
    if (ch === 'begin') {
      const f = { instr: '', begin: pos, separate: null, nested: open.length > 0 };
      open.push({ f, phase: 'instr' });
      return false;
    }
    const top = open[open.length - 1];
    if (!top) return false;
    if (ch === 'separate') {
      top.f.separate = pos;
      top.phase = 'result';
    } else if (ch === 'end') {
      open.pop();
      out.push({ ...(top.f as BodyField), end: pos });
    } else if (top.phase === 'instr') {
      const t = INSTR_TEXT.exec(xml);
      if (t) top.f.instr += unescape(t[1]);
    }
    return false;
  });
  return out.sort((a, b) => a.begin - b.begin);
}

function unescape(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

/** The switches of a field code (o, h … for \o, \h, lower case), outside its quoted arguments. */
export function fieldSwitches(instr: string): Set<string> {
  const out = new Set<string>();
  for (const m of instr.replace(/"[^"]*"/g, '""').matchAll(/\\([a-z])/gi)) out.add(m[1].toLowerCase());
  return out;
}

/**
 * Whether a TOC field lists the headings, so that 更新目錄 rebuilds it from them: it takes the
 * heading styles (\o) or the outline levels (\u), and none of \c / \a (a 圖目錄 or 表目錄 of the
 * captions), \f (TC fields), \t (other styles) or \b (only part of the document). Any other TOC
 * field (a table of figures …) keeps its entries; only its page numbers (PAGEREF) are updated.
 */
export function isHeadingToc(instr: string): boolean {
  if (!/^\s*TOC(\s|$)/i.test(instr)) return false;
  const sw = fieldSwitches(instr);
  return (sw.has('o') || sw.has('u')) && !['c', 'a', 'f', 't', 'b'].some((k) => sw.has(k));
}

const isToc = (f: BodyField) => !f.nested && isHeadingToc(f.instr);

/** The document's tables of contents of the headings (not inside another field; see isHeadingToc). */
export function tocFields(doc: PMNode): BodyField[] {
  return bodyFields(doc).filter(isToc);
}

/** The heading levels a TOC field lists (\o "1-3"; Word's default without a range is 1-9). */
export function tocLevels(instr: string): { from: number; to: number } {
  const m = /\\o\s+"?(\d)\s*-\s*(\d)"?/i.exec(instr);
  if (!m) return { from: 1, to: 9 };
  const a = Number(m[1]);
  const b = Number(m[2]);
  return { from: Math.max(1, Math.min(a, b)), to: Math.min(9, Math.max(a, b)) };
}

/** One entry of a table of contents to write. */
export interface TocEntry {
  /** 1 = heading 1. */
  level: number;
  text: string;
  /** The heading's list number ("一、"), shown before its text with a tab, as Word does. */
  number: string | null;
  bookmark: string;
  /** Position of the heading paragraph. */
  pos: number;
  /** The heading has no _Toc bookmark yet: one is added around its text. */
  newBookmark: boolean;
}

/** Everything a table of contents is built from. */
export interface TocContext {
  styles?: StyleInheritance;
  paragraphStyles: ParagraphStyleInfo[];
  numbering: Numbering;
  /** The printed page number at a body position ("iii" in a section numbered so); null before layout. */
  pageLabel: (pos: number) => string | null;
  /** Width of the text column in twips (for the right tab), where the table of contents goes. */
  textWidth: number;
  /** Makes new bookmark names (tests pass a fixed one). */
  bookmarkName?: (taken: Set<string>) => string;
}

/** Word names the bookmarks _Toc and nine digits; any free one will do. */
function newTocName(taken: Set<string>): string {
  let n = 100000000 + Math.floor(Math.random() * 800000000);
  while (taken.has(`_Toc${n}`)) n++;
  return `_Toc${n}`;
}

/** The headings a table of contents with `levels` lists, with their bookmarks (existing or new). */
export function planEntries(doc: PMNode, ctx: TocContext, levels: { from: number; to: number }, skip?: { from: number; to: number }): TocEntry[] {
  const taken = new Set<string>();
  doc.descendants((node) => {
    if (node.type === schema.nodes.raw_inline) {
      const m = BOOKMARK_START.exec(node.attrs.xml as string);
      if (m) taken.add(m[1]);
    }
    return true;
  });
  const numbers = listNumbers(doc, ctx.numbering);
  const out: TocEntry[] = [];
  for (const h of documentOutline(doc, ctx.styles)) {
    const level = h.level + 1;
    if (level < levels.from || level > levels.to) continue;
    if (skip && h.pos >= skip.from && h.pos < skip.to) continue;
    const para = doc.nodeAt(h.pos)!;
    let bookmark: string | null = null;
    para.forEach((c) => {
      const m = !bookmark && c.type === schema.nodes.raw_inline ? BOOKMARK_START.exec(c.attrs.xml as string) : null;
      if (m && /^_Toc/i.test(m[1])) bookmark = m[1];
    });
    const fresh = !bookmark;
    if (!bookmark) {
      bookmark = (ctx.bookmarkName ?? newTocName)(taken);
      taken.add(bookmark);
    }
    out.push({ level, text: h.text, number: numbers.get(h.pos) ?? null, bookmark, pos: h.pos, newBookmark: fresh });
  }
  return out;
}

/** The list number shown before each numbered paragraph (by position), counted through the document. */
function listNumbers(doc: PMNode, numbering: Numbering): Map<number, string> {
  const counter = new ListCounter(numbering);
  const out = new Map<number, string>();
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    if (node.attrs.numId) {
      const text = counter.next(node.attrs.numId, node.attrs.ilvl ?? 0);
      if (text) out.set(pos, text);
    }
    return false;
  });
  return out;
}

/** The id of the paragraph style Word calls `name` ("toc 1"), else the English built-in id. */
export function builtInStyleId(paragraphStyles: ParagraphStyleInfo[], name: string, fallback: string): string {
  const want = name.toLowerCase();
  return paragraphStyles.find((s) => s.name.toLowerCase() === want)?.id ?? fallback;
}

export const tocStyleId = (styles: ParagraphStyleInfo[], level: number) => builtInStyleId(styles, `toc ${level}`, `TOC${level}`);

const RUN_PR = '<w:rPr><w:noProof/></w:rPr>';
const HIDDEN_PR = '<w:rPr><w:noProof/><w:webHidden/></w:rPr>';
const run = (inner: string, rPr = RUN_PR) => `<w:r>${rPr}${inner}</w:r>`;
const text = (t: string) => `<w:t xml:space="preserve">${escapeXml(t)}</w:t>`;

/** The runs Word writes for the TOC field's start (before the first entry). */
export function tocBeginXml(instr = TOC_INSTR): string {
  return run('<w:fldChar w:fldCharType="begin"/>', '') + run(`<w:instrText xml:space="preserve"> ${escapeXml(instr)} </w:instrText>`, '') + run('<w:fldChar w:fldCharType="separate"/>', '');
}

/** One entry's paragraph properties: its 目錄 N style and the right tab with a dot leader. */
export function entryPPr(styleId: string, textWidth: number): string {
  return `<w:pPr><w:pStyle w:val="${escapeXml(styleId)}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${Math.max(0, Math.round(textWidth) - 10)}"/></w:tabs>${RUN_PR}</w:pPr>`;
}

/**
 * One entry's runs: its text, a tab and the PAGEREF field, in a link to its heading (the TOC
 * field's \h, which Word's 參考資料 › 目錄 always sets).
 */
export function entryRunsXml(e: TocEntry, page: string, link = true): string {
  const bm = escapeXml(e.bookmark);
  return (
    (link ? `<w:hyperlink w:anchor="${bm}" w:history="1">` : '') +
    (e.number ? run(text(e.number)) + run('<w:tab/>') : '') +
    run(text(e.text)) +
    run('<w:tab/>', HIDDEN_PR) +
    run('<w:fldChar w:fldCharType="begin"/>', HIDDEN_PR) +
    run(`<w:instrText xml:space="preserve"> PAGEREF ${bm} \\h </w:instrText>`, HIDDEN_PR) +
    run('<w:fldChar w:fldCharType="separate"/>', HIDDEN_PR) +
    run(text(page), HIDDEN_PR) +
    run('<w:fldChar w:fldCharType="end"/>', HIDDEN_PR) +
    (link ? '</w:hyperlink>' : '')
  );
}

/** Editor blocks from WordprocessingML body content (paragraphs), as the reader makes them. */
export function blocksFromXml(xml: string): PMNode[] {
  const doc = parseXml(`<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body>${xml}</w:body></w:document>`);
  return new Converter(new Map(), new Map()).blocks(child(doc.documentElement, 'body')!);
}

const pageOf = (e: TocEntry, ctx: TocContext) => ctx.pageLabel(e.pos) ?? '1';

/** A whole new table of contents: its title, the entries, and the paragraph ending the field. */
export function newTocXml(entries: TocEntry[], ctx: TocContext): string {
  const title = `<w:p><w:pPr><w:pStyle w:val="${escapeXml(builtInStyleId(ctx.paragraphStyles, 'toc heading', 'TOCHeading'))}"/></w:pPr>${run(text(TOC_TITLE), '')}</w:p>`;
  if (!entries.length) {
    return title + `<w:p>${tocBeginXml()}${run(text(NO_ENTRIES), '<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>')}${run('<w:fldChar w:fldCharType="end"/>', '')}</w:p>`;
  }
  const paras = entries.map(
    (e, i) => `<w:p>${entryPPr(tocStyleId(ctx.paragraphStyles, e.level), ctx.textWidth)}${i === 0 ? tocBeginXml() : ''}${entryRunsXml(e, pageOf(e, ctx))}</w:p>`,
  );
  return title + paras.join('') + `<w:p>${run('<w:fldChar w:fldCharType="end"/>', '')}</w:p>`;
}

// ----- bookmarks around the headings -----

/** The largest w:id in the document's kept XML (bookmarks, comments, revisions share the range here). */
function maxAnnotationId(doc: PMNode): number {
  let max = 0;
  const scan = (v: unknown) => {
    if (typeof v === 'string' && v.includes('w:id')) for (const m of v.matchAll(/w:id=\\?"(\d+)\\?"/g)) max = Math.max(max, Number(m[1]));
  };
  doc.descendants((n) => {
    Object.values(n.attrs).forEach(scan);
    for (const m of n.marks) Object.values(m.attrs).forEach(scan);
    return true;
  });
  return max;
}

/**
 * Adds a bookmark around each heading that gets a new one (Word's _Toc bookmarks: from before
 * the heading's text to its end). Positions are mapped through `tr`.
 */
export function addHeadingBookmarks(tr: Transaction, entries: TocEntry[]): void {
  let id = maxAnnotationId(tr.doc);
  const raw = (xml: string, label: string) => schema.nodes.raw_inline.create({ xml, label, hidden: true });
  // From the last heading back, so the earlier positions stay valid.
  for (const e of [...entries].filter((x) => x.newBookmark).sort((a, b) => b.pos - a.pos)) {
    const para = tr.doc.nodeAt(e.pos);
    if (!para || para.type !== schema.nodes.paragraph) continue;
    id++;
    tr.insert(e.pos + 1 + para.content.size, raw(`<w:bookmarkEnd w:id="${id}"/>`, 'bookmarkEnd'));
    tr.insert(e.pos + 1, raw(`<w:bookmarkStart w:id="${id}" w:name="${escapeXml(e.bookmark)}"/>`, 'bookmarkStart'));
  }
}

// ----- the entries of an existing table of contents -----

/** Level of a table of contents entry paragraph from its style (目錄 N / toc N / TOCN). */
function entryLevel(node: PMNode, styles: ParagraphStyleInfo[]): number | null {
  const id = node.attrs.styleId as string | null;
  if (!id) return null;
  const name = styles.find((s) => s.id === id)?.name ?? id;
  const m = /^(?:toc|目錄)\s*(\d)$/i.exec(name.trim()) ?? /^toc(\d)$/i.exec(id);
  return m ? Number(m[1]) : null;
}

/** What an entry paragraph says: its link target and its text up to the page number. */
function entryOf(node: PMNode, styles: ParagraphStyleInfo[]): { level: number | null; bookmark: string | null; text: string } {
  let bookmark: string | null = null;
  let out = '';
  node.forEach((c) => {
    const link = c.marks.find((m) => m.type === schema.marks.link);
    if (!bookmark && link && String(link.attrs.href).startsWith('#')) bookmark = String(link.attrs.href).slice(1);
    const pageRef = c.marks.find((m) => m.type === schema.marks.fieldResult && /^\s*PAGEREF\s/i.test(m.attrs.instr));
    // An entry without a link (no \h) names its heading's bookmark in its PAGEREF only.
    if (pageRef) {
      bookmark ??= String(pageRef.attrs.instr).trim().split(/\s+/)[1] ?? null;
      return;
    }
    if (c.isText) out += c.text;
    else if (c.type === schema.nodes.tab) out += '\t';
  });
  return { level: entryLevel(node, styles), bookmark, text: out.replace(/\t+$/, '') };
}

const norm = (s: string) => s.replace(/\s+/g, ' ').trim();

/** An existing table of contents, as blocks of the body. */
interface TocBlocks {
  field: BodyField;
  /** Indices of the top-level blocks from the field's start to its end. */
  first: number;
  last: number;
  /** Position of the first block. */
  from: number;
  to: number;
}

function tocBlocks(doc: PMNode, field: BodyField): TocBlocks | null {
  const $b = doc.resolve(field.begin);
  const $e = doc.resolve(field.end);
  // Only a table of contents in the body's own paragraphs (not in a table) is rebuilt.
  if ($b.depth !== 1 || $e.depth !== 1 || field.separate == null || doc.resolve(field.separate).depth !== 1) return null;
  if ($b.index(0) !== doc.resolve(field.separate).index(0)) return null;
  const first = $b.index(0);
  const last = $e.index(0);
  return { field, first, last, from: $b.before(1), to: $e.after(1) };
}

export interface RebuildResult {
  /** The table of contents was rebuilt (its entries changed). */
  rebuilt: boolean;
  entries: number;
  /** The paragraph styles of the rebuilt entries (the 目錄 styles the file lacks are added on save). */
  styleIds: Set<string>;
}

/**
 * Rebuilds the entries of every table of contents from the headings, when they differ from the
 * headings now (added, removed, renamed or other levels): the field's code and its start, its end
 * paragraph and the look of each level's entries (their paragraph properties) stay as they were.
 * Headings get their bookmarks. A table of contents that lists the same headings is left alone
 * (its page numbers are updated separately).
 */
export function rebuildTocs(tr: Transaction, ctx: TocContext): RebuildResult {
  let rebuilt = false;
  let count = 0;
  const styleIds = new Set<string>();
  const total = tocFields(tr.doc).length;
  // From the last one back, each found again in the document as it is now: rebuilding one adds
  // bookmarks to headings before the others too, which moves them.
  for (let k = total - 1; k >= 0; k--) {
    const fields = tocFields(tr.doc);
    if (fields.length !== total) break;
    const field = fields[k];
    const blocks = tocBlocks(tr.doc, field);
    if (!blocks) continue;
    const levels = tocLevels(field.instr);
    const entries = planEntries(tr.doc, ctx, levels, { from: blocks.from, to: blocks.to });
    count += entries.length;
    const olds: PMNode[] = [];
    for (let i = blocks.first; i <= blocks.last; i++) olds.push(tr.doc.child(i));
    const oldEntries = olds.map((n) => (n.type === schema.nodes.paragraph ? entryOf(n, ctx.paragraphStyles) : null)).filter((e) => e && e.bookmark);
    const same =
      oldEntries.length === entries.length &&
      entries.every((e, i) => {
        const o = oldEntries[i]!;
        return o.bookmark === e.bookmark && (o.level == null || o.level === e.level) && norm(o.text) === norm((e.number ? e.number + ' ' : '') + e.text);
      });
    if (same) continue;
    const replacement = rebuiltBlocks(tr.doc, blocks, olds, entries, ctx);
    if (!replacement) continue;
    // Bookmarks go into headings after the table of contents too: map its range afterwards.
    const mapFrom = tr.mapping.maps.length;
    addHeadingBookmarks(tr, entries);
    const map = tr.mapping.slice(mapFrom);
    tr.replaceWith(map.map(blocks.from, -1), map.map(blocks.to, 1), replacement);
    for (const p of replacement) if (p.attrs.styleId) styleIds.add(p.attrs.styleId as string);
    rebuilt = true;
  }
  return { rebuilt, entries: count, styleIds };
}

function rebuiltBlocks(doc: PMNode, blocks: TocBlocks, olds: PMNode[], entries: TocEntry[], ctx: TocContext): PMNode[] | null {
  const beginPara = olds[0];
  const endPara = olds[olds.length - 1];
  if (beginPara.type !== schema.nodes.paragraph || endPara.type !== schema.nodes.paragraph) return null;
  const beginStart = blocks.from + 1;
  const endStart = doc.resolve(blocks.field.end).start();
  // The field's start: everything in its first paragraph up to its separator.
  const prefix = beginPara.content.cut(0, blocks.field.separate! + 1 - beginStart);
  // Its end, and what follows it in the paragraph.
  const endOffset = blocks.field.end - endStart;
  const suffix = endPara.content.cut(endOffset);
  const before = endPara.content.cut(0, endOffset);
  // Word ends the field in a paragraph of its own: kept as it is.
  const endAlone = olds.length > 1 && !before.textBetween(0, before.size).trim() && !hasVisible(before);
  // Each level's paragraph properties as the entries had them (Word's style ids, tab position).
  const byLevel = new Map<number, Record<string, unknown>>();
  for (const n of olds) {
    const level = entryLevel(n, ctx.paragraphStyles);
    if (level != null && !byLevel.has(level)) byLevel.set(level, n.attrs);
  }
  const wrap = beginPara.attrs.wrap ?? null;
  const link = /\\h\b/i.test(blocks.field.instr);
  const fresh = entries.length
    ? blocksFromXml(entries.map((e) => `<w:p>${entryPPr(tocStyleId(ctx.paragraphStyles, e.level), ctx.textWidth)}${entryRunsXml(e, pageOf(e, ctx), link)}</w:p>`).join(''))
    : blocksFromXml(`<w:p>${run(text(NO_ENTRIES), '<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>')}</w:p>`);
  const out = fresh.map((p, i) => {
    const kept = entries.length ? byLevel.get(entries[i].level) : beginPara.attrs;
    const attrs = { ...(kept ?? p.attrs), sectPr: null, brGroup: null, wrap };
    return schema.nodes.paragraph.create(attrs, p.content);
  });
  out[0] = out[0].type.create(out[0].attrs, prefix.append(out[0].content));
  if (endAlone) {
    out.push(endPara);
  } else {
    const lastIndex = out.length - 1;
    const last = out[lastIndex];
    out[lastIndex] = last.type.create({ ...last.attrs, sectPr: endPara.attrs.sectPr }, last.content.append(suffix));
  }
  return out;
}

function hasVisible(f: Fragment): boolean {
  let found = false;
  f.forEach((n) => {
    if (!(n.type === schema.nodes.raw_inline && n.attrs.hidden)) found = true;
  });
  return found;
}

export { Fragment };
