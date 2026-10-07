// Comments (word/comments.xml, resolved state and replies from word/commentsExtended.xml, UTC
// times from word/commentsIds.xml + word/commentsExtensible.xml).
// The document keeps the comment range markers and reference marks as raw nodes (see
// convert.ts); the list of comments travels on the editor document (doc attribute `comments`,
// so undo covers comment edits) and is written back by writeComments: a package whose
// comments nobody touched keeps every comment part byte for byte, and an edit rewrites only
// the entries it changed.

import type JSZip from 'jszip';
import { NS, children, escapeAttr, parseFragment, parseXml, rootAttributes, serializeXml, xmlSafe } from './xml';
import { findMainPart, readRels } from './reader';

export interface DocComment {
  id: string;
  author: string;
  initials: string | null;
  /** w:date as written (Word writes local time there, with a "Z"). */
  date: string | null;
  /** The real UTC time (w16cex:dateUtc), when Word wrote one. */
  dateUtc: string | null;
  /** Plain text, paragraphs separated by newlines. */
  text: string;
  /** Id of the comment this one replies to. */
  parentId: string | null;
  /** Marked as resolved in Word. */
  done: boolean;
  /** Made in the web editor (not in the opened file), even when it reuses a deleted comment's id. */
  created?: boolean;
}

export const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const COMMENTS_IDS_REL = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
export const COMMENTS_EXTENSIBLE_REL = 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

function paragraphText(p: Element): string {
  let out = '';
  const walk = (e: Element) => {
    for (const c of Array.from(e.children)) {
      const w = c.namespaceURI === NS.w ? c.localName : '';
      if (w === 't') out += c.textContent ?? '';
      else if (w === 'tab') out += '\t';
      else if (w === 'br' || w === 'cr') out += '\n';
      else if (w === 'delText' || w === 'instrText' || w === 'rPr' || w === 'pPr') continue;
      else walk(c);
    }
  };
  walk(p);
  return out;
}

/** Optional parts that add to comments.xml. */
export interface CommentParts {
  extended?: string | null;
  ids?: string | null;
  extensible?: string | null;
}

/** Parse comments.xml (and optionally commentsExtended / commentsIds / commentsExtensible). */
export function parseComments(commentsXml: string, extra: string | null | CommentParts = null): DocComment[] {
  const parts: CommentParts = typeof extra === 'string' || extra == null ? { extended: extra } : extra;
  const root = parseXml(commentsXml).documentElement;
  const byPara = new Map<string, DocComment>();
  const out: DocComment[] = [];
  for (const c of children(root, 'comment')) {
    const paras = Array.from(c.getElementsByTagNameNS(NS.w, 'p'));
    const comment: DocComment = {
      id: c.getAttributeNS(NS.w, 'id') ?? '',
      author: c.getAttributeNS(NS.w, 'author') ?? '',
      initials: c.getAttributeNS(NS.w, 'initials') || null,
      date: c.getAttributeNS(NS.w, 'date') || null,
      dateUtc: null,
      text: paras.map(paragraphText).join('\n').trim(),
      parentId: null,
      done: false,
    };
    // commentsExtended / commentsIds refer to a comment by the paraId of its last paragraph.
    const last = paras[paras.length - 1];
    const paraId = last?.getAttributeNS(W14, 'paraId');
    if (paraId) byPara.set(paraId, comment);
    out.push(comment);
  }
  if (parts.extended) {
    const ext = parseXml(parts.extended).documentElement;
    for (const e of Array.from(ext.getElementsByTagNameNS(W15, 'commentEx'))) {
      const c = byPara.get(e.getAttributeNS(W15, 'paraId') ?? '');
      if (!c) continue;
      const done = e.getAttributeNS(W15, 'done');
      c.done = done === '1' || done === 'true';
      const parent = byPara.get(e.getAttributeNS(W15, 'paraIdParent') ?? '');
      if (parent && parent !== c) c.parentId = parent.id;
    }
  }
  if (parts.ids && parts.extensible) {
    // paraId -> durableId (commentsIds), durableId -> UTC time (commentsExtensible).
    const utc = new Map<string, string>();
    for (const e of Array.from(parseXml(parts.extensible).documentElement.getElementsByTagNameNS(W16CEX, 'commentExtensible'))) {
      const time = e.getAttributeNS(W16CEX, 'dateUtc');
      if (time) utc.set(e.getAttributeNS(W16CEX, 'durableId') ?? '', time);
    }
    for (const e of Array.from(parseXml(parts.ids).documentElement.getElementsByTagNameNS(W16CID, 'commentId'))) {
      const c = byPara.get(e.getAttributeNS(W16CID, 'paraId') ?? '');
      const time = utc.get(e.getAttributeNS(W16CID, 'durableId') ?? '');
      if (c && time) c.dateUtc = time;
    }
  }
  return out;
}

/** The thread a comment belongs to: its top-level comment (see commentThreads). */
export function threadRoot(comments: readonly DocComment[], id: string): DocComment | null {
  const byId = new Map<string, DocComment>();
  for (const c of comments) if (!byId.has(c.id)) byId.set(c.id, c);
  let c = byId.get(id);
  if (!c) return null;
  const seen = new Set<DocComment>([c]);
  for (let p = c.parentId != null ? byId.get(c.parentId) : undefined; p; p = p.parentId != null ? byId.get(p.parentId) : undefined) {
    if (seen.has(p)) return byId.get(id)!; // a loop: the comment stands on its own
    seen.add(p);
    c = p;
  }
  return c;
}

/** A comment and every reply under it, however deep (ids). */
export function withReplies(comments: readonly DocComment[], id: string): Set<string> {
  const out = new Set([id]);
  for (let grew = true; grew; ) {
    grew = false;
    for (const c of comments) {
      if (c.parentId != null && out.has(c.parentId) && !out.has(c.id)) {
        out.add(c.id);
        grew = true;
      }
    }
  }
  return out;
}

export interface CommentThread {
  /** Stable key for lists (comment ids may repeat in a damaged file). */
  key: string;
  comment: DocComment;
  /** Every reply in the thread (replies to replies included), in file order. */
  replies: DocComment[];
}

/**
 * Comments grouped as Word shows them: each top-level comment with all the replies under it,
 * however deep. A reply whose parent is missing, or that is part of a parent loop, starts a
 * thread of its own, so no comment is ever hidden.
 */
export function commentThreads(comments: DocComment[]): CommentThread[] {
  const byId = new Map<string, DocComment>();
  for (const c of comments) if (!byId.has(c.id)) byId.set(c.id, c);
  const rootOf = new Map<DocComment, DocComment>();
  const root = (c: DocComment): DocComment => {
    const known = rootOf.get(c);
    if (known) return known;
    const seen = new Set<DocComment>([c]);
    let top = c;
    for (let p = c.parentId != null ? byId.get(c.parentId) : undefined; p && p !== top; p = p.parentId != null ? byId.get(p.parentId) : undefined) {
      if (seen.has(p)) {
        top = c; // a loop: the comment stands on its own
        break;
      }
      seen.add(p);
      top = p;
    }
    rootOf.set(c, top);
    return top;
  };
  const threads = new Map<DocComment, CommentThread>();
  comments.forEach((c, i) => {
    const top = root(c);
    if (top === c) threads.set(c, { key: `${i}:${c.id}`, comment: c, replies: [] });
  });
  for (const c of comments) {
    const top = root(c);
    if (top !== c) threads.get(top)?.replies.push(c);
  }
  return [...threads.values()];
}

/** The comments of an opened package; empty when it has none. */
export async function readComments(zip: JSZip | null): Promise<DocComment[]> {
  if (!zip) return [];
  const main = await findMainPart(zip);
  const rels = [...(await readRels(zip, main)).values()];
  const part = rels.find((r) => r.type === COMMENTS_REL && !r.external);
  const file = part && zip.file(part.target);
  if (!file) return [];
  const text = async (type: string) => {
    const r = rels.find((x) => x.type === type && !x.external);
    const f = r && zip.file(r.target);
    return f ? f.async('string') : null;
  };
  try {
    return parseComments(await file.async('string'), {
      extended: await text(COMMENTS_EXTENDED_REL),
      ids: await text(COMMENTS_IDS_REL),
      extensible: await text(COMMENTS_EXTENSIBLE_REL),
    });
  } catch {
    return [];
  }
}

// ----- writing -----

export const COMMENT_CONTENT_TYPES = {
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  extended: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
  ids: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml',
  extensible: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml',
} as const;

/** Who writes new comments (the signed-in user, from the host page). */
export interface CommentAuthor {
  name: string;
  initials?: string | null;
}

/** Author of new comments when the host page gives none. */
export const DEFAULT_AUTHOR = '使用者';

/** Initials as Word makes them from a name: the first character of each part ("灌庭 吳" → "灌吳"). */
export function authorInitials(name: string): string {
  return name
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((p) => Array.from(p)[0])
    .join('');
}

/**
 * The time of a new comment as Word writes it: w:date holds the local time's digits with a "Z"
 * (not a UTC time), the real UTC time goes to commentsExtensible.xml; both to the minute.
 */
export function commentDate(now: Date = new Date()): { date: string; dateUtc: string } {
  const p = (n: number) => String(n).padStart(2, '0');
  const date = `${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}T${p(now.getHours())}:${p(now.getMinutes())}:00Z`;
  const dateUtc = `${now.getUTCFullYear()}-${p(now.getUTCMonth() + 1)}-${p(now.getUTCDate())}T${p(now.getUTCHours())}:${p(now.getUTCMinutes())}:00Z`;
  return { date, dateUtc };
}

/** Ids of the styles Word gives comment text and comment marks ("annotation text" / "annotation reference"). */
export interface CommentStyles {
  text: string | null;
  reference: string | null;
}

export function commentStyleIds(stylesXml: string | null): CommentStyles {
  const out: CommentStyles = { text: null, reference: null };
  if (!stylesXml) return out;
  try {
    for (const s of children(parseXml(stylesXml).documentElement, 'style')) {
      const name = (children(s, 'name')[0]?.getAttributeNS(NS.w, 'val') ?? '').toLowerCase();
      const id = s.getAttributeNS(NS.w, 'styleId');
      if (!id) continue;
      if (name === 'annotation text' && !out.text) out.text = id;
      if (name === 'annotation reference' && !out.reference) out.reference = id;
    }
  } catch {
    // unreadable styles: no style references
  }
  return out;
}

const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The comment styles of an opened package. */
export async function readCommentStyles(zip: JSZip | null): Promise<CommentStyles> {
  if (!zip) return { text: null, reference: null };
  try {
    const main = await findMainPart(zip);
    const styles = [...(await readRels(zip, main)).values()].find((r) => r.type === STYLES_REL && !r.external);
    const f = styles && zip.file(styles.target);
    return commentStyleIds(f ? await f.async('string') : null);
  } catch {
    return { text: null, reference: null };
  }
}

/** Whether two comment lists say the same (nothing to write). */
export function sameComments(a: readonly DocComment[], b: readonly DocComment[]): boolean {
  return (
    a.length === b.length &&
    a.every((x, i) => {
      const y = b[i];
      return (
        !x.created && !y.created && x.id === y.id && x.author === y.author && x.initials === y.initials &&
        x.date === y.date && x.dateUtc === y.dateUtc && x.text === y.text && x.parentId === y.parentId && x.done === y.done
      );
    })
  );
}

// A part is edited as text, so every entry nobody changed keeps its exact bytes.
interface Span {
  start: number;
  end: number;
  local: string;
  xml: string;
}
interface SplitPart {
  text: string;
  root: { start: number; end: number; name: string; selfClosing: boolean };
  /** Where the root's end tag starts (the root's end when it is self-closing). */
  close: number;
  children: Span[];
}

const TAG =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<!DOCTYPE[^>]*>|<(\/?)([\w.:-]+)((?:\s+[\w.:-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;

/** The root and the top-level elements of a part, by position; null when it can't be split. */
function splitPart(text: string): SplitPart | null {
  let depth = 0;
  let root: SplitPart['root'] | null = null;
  let close = -1;
  let open: { start: number; local: string } | null = null;
  const kids: Span[] = [];
  for (const m of text.matchAll(TAG)) {
    if (!m[2]) continue; // comment, processing instruction, CDATA
    const at = m.index!;
    const end = at + m[0].length;
    const local = m[2].replace(/^.*:/, '');
    if (m[1] === '/') {
      depth--;
      if (depth === 1 && open) {
        kids.push({ start: open.start, end, local: open.local, xml: text.slice(open.start, end) });
        open = null;
      } else if (depth === 0) {
        close = at;
        break;
      }
      continue;
    }
    const selfClosing = m[4] === '/';
    if (depth === 0) {
      root = { start: at, end, name: m[2], selfClosing };
      if (selfClosing) {
        close = end;
        break;
      }
      depth = 1;
    } else if (depth === 1) {
      if (selfClosing) kids.push({ start: at, end, local, xml: m[0] });
      else {
        open = { start: at, local };
        depth = 2;
      }
    } else if (!selfClosing) depth++;
  }
  return root && close >= 0 ? { text, root, close, children: kids } : null;
}

/** The root start tag with the namespace declarations it lacks. */
function withNamespaces(tag: string, ns: Record<string, string>): string {
  let add = '';
  for (const [p, uri] of Object.entries(ns)) if (!new RegExp(`\\sxmlns:${p}=`).test(tag)) add += ` xmlns:${p}="${uri}"`;
  return add ? tag.replace(/\s*(\/?)>$/, `${add}$1>`) : tag;
}

/**
 * A part rebuilt from its text: `edit` gives each top-level element's replacement (undefined:
 * kept as is, '': removed), `append` goes before the root's end tag.
 */
function rebuild(p: SplitPart, ns: Record<string, string>, edit: (span: Span) => string | undefined, append: string): string {
  const { text, root } = p;
  const tag = withNamespaces(text.slice(root.start, root.end), ns);
  if (root.selfClosing) {
    const body = append ? tag.replace(/\s*\/>$/, '>') + append + `</${root.name}>` : tag;
    return text.slice(0, root.start) + body + text.slice(root.end);
  }
  let out = text.slice(0, root.start) + tag;
  let at = root.end;
  for (const s of p.children) {
    out += text.slice(at, s.start);
    out += edit(s) ?? s.xml;
    at = s.end;
  }
  return out + text.slice(at, p.close) + append + text.slice(p.close);
}

/** An attribute of an element's start tag, by its qualified name. */
function attrOf(xml: string, name: string): string | null {
  const m = new RegExp(`^<[^>]*?\\s${name.replace(/[.:]/g, '\\$&')}="([^"]*)"`).exec(xml);
  return m ? unescapeAttr(m[1]) : null;
}

/** An attribute value as the DOM reads it, so it compares equal to ids read with getAttributeNS. */
function unescapeAttr(v: string): string {
  if (!v.includes('&')) return v;
  const named: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  return v.replace(/&(#x[0-9A-Fa-f]+|#[0-9]+|[a-z]+);/g, (whole, e: string) => {
    if (e[0] !== '#') return named[e] ?? whole;
    const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return code <= 0x10ffff ? String.fromCodePoint(code) : whole;
  });
}

/**
 * A paragraph id for comparing: hex ids (all Word writes) in upper case; anything else a damaged
 * or hostile file holds is kept exactly, so it still matches the paragraph it names.
 */
const normId = (id: string): string => (/^[0-9A-Fa-f]+$/.test(id) ? id.toUpperCase() : id);

/** A new 8-digit hex id below 0x80000000 (Word's paraId / durableId range), unique in `used`. */
function hex8(used: Set<string>): string {
  let id: string;
  do id = Math.floor(1 + Math.random() * 0x7ffffffe).toString(16).toUpperCase().padStart(8, '0');
  while (used.has(id));
  used.add(id);
  return id;
}

/** Escaped comment text; characters XML forbids are dropped (see xmlSafe). */
const textEsc = (s: string) => xmlSafe(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A run for one line of comment text (tabs become w:tab, a vertical tab a line break as in Word). */
function textRun(line: string): string {
  if (!line) return '';
  const inner = line
    .split(/([\t\u000B])/)
    .map((t) => (t === '\t' ? '<w:tab/>' : t === '\u000B' ? '<w:br/>' : t ? `<w:t xml:space="preserve">${textEsc(t)}</w:t>` : ''))
    .join('');
  return `<w:r>${inner}</w:r>`;
}

/** What an edited comment keeps from its element: the start tag, the paragraph properties, the mark's run. */
interface Keep {
  startTag: string;
  pPr: string | null;
  refRun: string | null;
}

/** A w:comment with the text as plain paragraphs, like Word's; the last paragraph gets `paraId`. */
function commentXml(c: DocComment, paraId: string, usedParaIds: Set<string>, styles: CommentStyles, keep: Keep | null): string {
  const lines = c.text.split('\n');
  const start =
    keep?.startTag ??
    `<w:comment w:id="${escapeAttr(c.id)}" w:author="${escapeAttr(c.author)}"` +
      `${c.date ? ` w:date="${escapeAttr(c.date)}"` : ''}${c.initials ? ` w:initials="${escapeAttr(c.initials)}"` : ''}>`;
  const pPr = keep ? keep.pPr ?? '' : styles.text ? `<w:pPr><w:pStyle w:val="${escapeAttr(styles.text)}"/></w:pPr>` : '';
  const ref =
    keep?.refRun ?? `<w:r>${styles.reference ? `<w:rPr><w:rStyle w:val="${escapeAttr(styles.reference)}"/></w:rPr>` : ''}<w:annotationRef/></w:r>`;
  const paras = lines.map((line, i) => {
    const id = i === lines.length - 1 ? paraId : hex8(usedParaIds);
    return `<w:p w14:paraId="${escapeAttr(id)}" w14:textId="77777777">${pPr}${i === 0 ? ref : ''}${textRun(line)}</w:p>`;
  });
  return `${start}${paras.join('')}</w:comment>`;
}

/** What an original w:comment element is made of. */
function commentParts(xml: string): { paraId: string | null; keep: Keep; paragraphs: number } {
  const startTag = /^<[^>]*>/.exec(xml)?.[0] ?? '';
  const keep: Keep = { startTag, pPr: null, refRun: null };
  try {
    const el = parseFragment(xml);
    const paras = Array.from(el.getElementsByTagNameNS(NS.w, 'p'));
    const first = paras[0];
    const pPr = first && children(first, 'pPr')[0];
    if (pPr) keep.pPr = serializeXml(pPr);
    const run = el.getElementsByTagNameNS(NS.w, 'annotationRef')[0]?.parentElement;
    if (run && run.namespaceURI === NS.w && run.localName === 'r') keep.refRun = serializeXml(run);
    const paraId = paras[paras.length - 1]?.getAttributeNS(W14, 'paraId') || null;
    return { paraId: paraId && normId(paraId), keep, paragraphs: paras.length };
  } catch {
    return { paraId: null, keep, paragraphs: 0 };
  }
}

type Rel = { id: string; type: string; target: string; external: boolean };

/** The package writeComments adds to (see writeDocx). */
export interface CommentPackage {
  zip: JSZip;
  /** The main document part and its relationships as read. */
  main: string;
  mainRels: Map<string, Rel>;
  /** A relationship id the main part doesn't use yet. */
  newRelId: (prefix: string) => string;
  /** Relationships to add to the main part (targets are package paths). */
  extraRels: Rel[];
  /** Content type overrides to add, by part name ("/word/comments.xml"). */
  overrides: Record<string, string>;
}

/** A new part next to the main document ("word/comments.xml", with a number when the name is taken). */
function newPart(zip: JSZip, main: string, name: string): string {
  const dir = main.includes('/') ? main.slice(0, main.lastIndexOf('/') + 1) : '';
  let path = `${dir}${name}.xml`;
  for (let n = 1; zip.file(path); n++) path = `${dir}${name}${n}.xml`;
  return path;
}

/**
 * Write the editor's comments into the package. Nothing is written when they are what the
 * package already has; otherwise only the entries of the comments that were added, edited,
 * replied to, resolved / reopened or deleted change, and every other entry keeps its bytes.
 * A package without comments gets comments.xml, commentsExtended.xml, commentsIds.xml and
 * commentsExtensible.xml, with their relationships and content types, as Word makes them.
 */
export async function writeComments(current: readonly DocComment[] | null | undefined, pkg: CommentPackage): Promise<void> {
  if (!current) return;
  const { zip, main } = pkg;
  const rels = [...pkg.mainRels.values()];
  const pathOf = (type: string) => rels.find((r) => r.type === type && !r.external)?.target ?? null;
  const paths = {
    comments: pathOf(COMMENTS_REL),
    extended: pathOf(COMMENTS_EXTENDED_REL),
    ids: pathOf(COMMENTS_IDS_REL),
    extensible: pathOf(COMMENTS_EXTENSIBLE_REL),
  };
  const read = async (p: string | null) => {
    const f = p ? zip.file(p) : null;
    return f ? f.async('string') : null;
  };
  const texts = {
    comments: await read(paths.comments),
    extended: await read(paths.extended),
    ids: await read(paths.ids),
    extensible: await read(paths.extensible),
  };
  let original: DocComment[] = [];
  try {
    if (texts.comments) original = parseComments(texts.comments, { extended: texts.extended, ids: texts.ids, extensible: texts.extensible });
  } catch {
    original = [];
  }
  if (sameComments(current, original)) return;

  // A package that had no comments gets every part Word 365 writes.
  const fresh = !texts.comments;
  const commentsPart = texts.comments ? splitPart(texts.comments) : null;
  const spans = commentsPart?.children.filter((s) => s.local === 'comment') ?? [];
  // Each original comment's element (none when they can't be matched: all are written again).
  const elements = spans.length === original.length ? spans : [];
  const info = elements.map((s) => commentParts(s.xml));

  // Current comment -> its original (null: made in the editor).
  const kept = new Set<number>();
  const origOf = current.map((c) => {
    if (c.created) return null;
    const j = original.findIndex((o, k) => !kept.has(k) && o.id === c.id);
    if (j < 0) return null;
    kept.add(j);
    return j;
  });

  // Paragraph ids must be unique in the package; durable ids in the comment parts.
  const usedParaIds = new Set<string>();
  const usedDurable = new Set<string>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir || !/\.xml$/i.test(name)) continue;
    const xml = await file.async('string');
    for (const m of xml.matchAll(/:paraId="([0-9A-Fa-f]{8})"/g)) usedParaIds.add(m[1].toUpperCase());
    for (const m of xml.matchAll(/:durableId="([0-9A-Fa-f]{8})"/g)) usedDurable.add(m[1].toUpperCase());
  }

  const parents = new Set(current.map((c) => c.parentId).filter((x): x is string => x != null));
  /** Whether commentsExtended must say something about the comment. */
  const extendedInfo = (c: DocComment) => c.parentId != null || c.done || parents.has(c.id);
  const rewritten = current.map((c, i) => {
    const j = origOf[i];
    return j == null || !elements.length || info[j].paragraphs === 0 || c.text !== original[j].text;
  });
  const paraOf: (string | null)[] = current.map((c, i) => {
    const j = origOf[i];
    const had = j != null && elements.length ? info[j].paraId : null;
    if (had) return had;
    return j == null || rewritten[i] || extendedInfo(c) ? hex8(usedParaIds) : null;
  });
  const indexById = new Map<string, number>();
  current.forEach((c, i) => {
    if (!indexById.has(c.id)) indexById.set(c.id, i);
  });
  const parentPara = (c: DocComment) => {
    const k = c.parentId != null ? indexById.get(c.parentId) : undefined;
    return k != null ? paraOf[k] : null;
  };
  const byPara = new Map<string, number>(); // paraId of a comment still there -> current index
  current.forEach((_, i) => {
    if (paraOf[i]) byPara.set(paraOf[i]!, i);
  });

  const addPart = (path: string | null, name: string, type: string, contentType: string, xml: string) => {
    const at = path ?? newPart(zip, main, name);
    zip.file(at, xml);
    if (!path) pkg.extraRels.push({ id: pkg.newRelId('rIdPxCm'), type, target: at, external: false });
    pkg.overrides['/' + at] = contentType;
  };

  // comments.xml
  const styles = await readCommentStyles(zip);
  const replaced = new Map<number, string>(); // original index -> its new element
  let added = '';
  current.forEach((c, i) => {
    const j = origOf[i];
    if (j == null || !elements.length) {
      added += commentXml(c, paraOf[i]!, usedParaIds, styles, null);
    } else if (rewritten[i]) {
      replaced.set(j, commentXml(c, paraOf[i]!, usedParaIds, styles, info[j].keep));
    } else if (!info[j].paraId && paraOf[i]) {
      // It needs a paragraph id now (a reply, or resolved): given to its last paragraph.
      const xml = elements[j].xml;
      let last = -1;
      for (const m of xml.matchAll(/<(?:[\w.-]+:)?p(?=[\s>/])/g)) last = m.index! + m[0].length;
      replaced.set(j, xml.slice(0, last) + ` w14:paraId="${escapeAttr(paraOf[i]!)}" w14:textId="77777777"` + xml.slice(last));
    }
  });
  const index = new Map(elements.map((s, k) => [s.start, k]));
  const W_NS = { w: NS.w, w14: W14 };
  addPart(
    paths.comments,
    'comments',
    COMMENTS_REL,
    COMMENT_CONTENT_TYPES.comments,
    commentsPart
      ? rebuild(
          commentsPart,
          W_NS,
          (s) => {
            if (s.local !== 'comment') return undefined;
            const k = index.get(s.start);
            if (k == null) return ''; // unmatched: written again above
            return kept.has(k) ? replaced.get(k) : ''; // '': deleted
          },
          added,
        )
      : `${XML_HEAD}<w:comments ${rootAttributes(null)}>${added}</w:comments>`,
  );

  // commentsExtended.xml: replies (w15:paraIdParent) and resolved threads (w15:done).
  if (texts.extended || fresh || current.some((c) => c.parentId != null || c.done)) {
    const entry = (i: number) => {
      const pp = parentPara(current[i]);
      // Paragraph ids read from the file are written back escaped (a file can hold any text there).
      return `<w15:commentEx w15:paraId="${escapeAttr(paraOf[i]!)}"${pp ? ` w15:paraIdParent="${escapeAttr(pp)}"` : ''} w15:done="${current[i].done ? 1 : 0}"/>`;
    };
    const part = texts.extended ? splitPart(texts.extended) : null;
    const listed = new Set(
      (part?.children ?? []).filter((s) => s.local === 'commentEx').map((s) => normId(attrOf(s.xml, 'w15:paraId') ?? '')),
    );
    let append = '';
    current.forEach((c, i) => {
      const para = paraOf[i];
      // Word lists every comment; one from the file that never had an entry gets one only when needed.
      if (para && !listed.has(para) && (origOf[i] == null || !texts.extended || extendedInfo(c))) append += entry(i);
    });
    const edit = (s: Span) => {
      if (s.local !== 'commentEx') return undefined;
      const i = byPara.get(normId(attrOf(s.xml, 'w15:paraId') ?? ''));
      if (i == null) return ''; // its comment was deleted
      const j = origOf[i];
      const o = j != null ? original[j] : null;
      return o && o.done === current[i].done && o.parentId === current[i].parentId ? undefined : entry(i);
    };
    addPart(
      paths.extended,
      'commentsExtended',
      COMMENTS_EXTENDED_REL,
      COMMENT_CONTENT_TYPES.extended,
      part ? rebuild(part, { w15: W15 }, edit, append) : `${XML_HEAD}<w15:commentsEx ${rootAttributes(null)}>${append}</w15:commentsEx>`,
    );
  }

  // commentsIds.xml + commentsExtensible.xml: durable ids and UTC times (Word 2019 and later).
  if (texts.ids || texts.extensible || fresh) {
    const idsPart = texts.ids ? splitPart(texts.ids) : null;
    const listed = new Set<string>();
    const dropped = new Set<string>(); // durable ids of deleted comments
    for (const s of idsPart?.children ?? []) {
      if (s.local !== 'commentId') continue;
      const para = normId(attrOf(s.xml, 'w16cid:paraId') ?? '');
      if (byPara.has(para)) listed.add(para);
      else dropped.add((attrOf(s.xml, 'w16cid:durableId') ?? '').toUpperCase());
    }
    let appendIds = '';
    let appendTimes = '';
    current.forEach((c, i) => {
      const para = paraOf[i];
      if (!para || listed.has(para)) return;
      const durable = hex8(usedDurable);
      appendIds += `<w16cid:commentId w16cid:paraId="${escapeAttr(para)}" w16cid:durableId="${durable}"/>`;
      if (c.dateUtc) appendTimes += `<w16cex:commentExtensible w16cex:durableId="${durable}" w16cex:dateUtc="${escapeAttr(c.dateUtc)}"/>`;
    });
    const dropIds = (s: Span) =>
      s.local === 'commentId' && !byPara.has(normId(attrOf(s.xml, 'w16cid:paraId') ?? '')) ? '' : undefined;
    addPart(
      paths.ids,
      'commentsIds',
      COMMENTS_IDS_REL,
      COMMENT_CONTENT_TYPES.ids,
      idsPart
        ? rebuild(idsPart, { w16cid: W16CID }, dropIds, appendIds)
        : `${XML_HEAD}<w16cid:commentsIds xmlns:mc="${MC}" xmlns:w16cid="${W16CID}" mc:Ignorable="w16cid">${appendIds}</w16cid:commentsIds>`,
    );
    const exPart = texts.extensible ? splitPart(texts.extensible) : null;
    const dropTimes = (s: Span) =>
      s.local === 'commentExtensible' && dropped.has((attrOf(s.xml, 'w16cex:durableId') ?? '').toUpperCase()) ? '' : undefined;
    addPart(
      paths.extensible,
      'commentsExtensible',
      COMMENTS_EXTENSIBLE_REL,
      COMMENT_CONTENT_TYPES.extensible,
      exPart
        ? rebuild(exPart, { w16cex: W16CEX }, dropTimes, appendTimes)
        : `${XML_HEAD}<w16cex:commentsExtensible xmlns:mc="${MC}" xmlns:w16cex="${W16CEX}" mc:Ignorable="w16cex">${appendTimes}</w16cex:commentsExtensible>`,
    );
  }
}
