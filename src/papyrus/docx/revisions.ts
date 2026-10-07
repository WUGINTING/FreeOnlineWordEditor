// Tracked changes as the editor keeps them (see convert.ts):
//  - w:ins / w:moveTo: an inlineWrap layer around the inserted content,
//  - w:del / w:moveFrom: a raw_inline node holding the whole element,
//  - w:rPrChange: inside a run's w:rPr; w:pPrChange and paragraph-mark revisions
//    (w:pPr/w:rPr/w:ins|w:del): inside a paragraph's w:pPr.
// Pure XML helpers: reading who/when, and the property rewrites for accept / reject.

import type { DOMOutputSpec } from 'prosemirror-model';
import { NS, child, parseFragment, parseFragments, serializeXml } from './xml';

export interface RevisionInfo {
  author: string;
  /** w:date as written: Word writes local time there, with a "Z". */
  date: string | null;
  /** w16du:dateUtc (Word 365): the real UTC time, when present. */
  dateUtc?: string | null;
}

/** Put a value in a module-level cache, forgetting the oldest entries past `limit`. */
export function remember<K, V>(cache: Map<K, V>, key: K, value: V, limit = 4000): V {
  if (cache.size >= limit) {
    let drop = Math.max(1, limit >> 2);
    for (const k of cache.keys()) {
      cache.delete(k);
      if (--drop <= 0) break;
    }
  }
  cache.set(key, value);
  return value;
}

function unescape(s: string): string {
  return s
    .replace(/&#10;/g, '\n')
    .replace(/&#13;/g, '\r')
    .replace(/&#9;/g, '\t')
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function tagAttr(tag: string, local: string): string | null {
  const m = new RegExp(`\\s(?:[\\w.-]+:)?${local}="([^"]*)"`).exec(tag);
  return m ? unescape(m[1]) : null;
}

function infoOf(el: Element): RevisionInfo {
  // w16du may be declared on the part root only, so it is found by local name.
  const utc = Array.from(el.attributes).find((a) => a.localName === 'dateUtc')?.value || null;
  return { author: el.getAttributeNS(NS.w, 'author') || '', date: el.getAttributeNS(NS.w, 'date') || null, dateUtc: utc };
}

const layerCache = new Map<string, { kind: 'ins' | 'moveTo'; info: RevisionInfo } | null>();

/** An inline wrapper's start tag, when it is a tracked insertion (w:ins) or move destination (w:moveTo). */
export function insertionOfLayer(open: string): { kind: 'ins' | 'moveTo'; info: RevisionInfo } | null {
  let hit = layerCache.get(open);
  if (hit === undefined) {
    const m = /^<(?:[\w.-]+:)?(ins|moveTo)[\s>/]/.exec(open);
    const tag = /^<[^>]*>/.exec(open)?.[0] ?? '';
    hit = m
      ? { kind: m[1] as 'ins' | 'moveTo', info: { author: tagAttr(tag, 'author') ?? '', date: tagAttr(tag, 'date'), dateUtc: tagAttr(tag, 'dateUtc') } }
      : null;
    remember(layerCache, open, hit);
  }
  return hit;
}

export interface DeletedInfo extends RevisionInfo {
  kind: 'del' | 'moveFrom';
  /** What the deleted content read as, for showing it struck through. */
  text: string;
}

const deletedCache = new Map<string, DeletedInfo | null>();

/** A kept w:del / w:moveFrom element: who deleted it, and its text. */
export function deletedInfo(xml: string): DeletedInfo | null {
  let hit = deletedCache.get(xml);
  if (hit === undefined) {
    hit = null;
    try {
      const el = parseFragment(xml);
      if (el.namespaceURI === NS.w && (el.localName === 'del' || el.localName === 'moveFrom')) {
        hit = { kind: el.localName, ...infoOf(el), text: contentText(el) };
      }
    } catch {
      hit = null;
    }
    remember(deletedCache, xml, hit, 2000);
  }
  return hit;
}

function contentText(el: Element): string {
  let out = '';
  const walk = (e: Element) => {
    for (const c of Array.from(e.children)) {
      const w = c.namespaceURI === NS.w ? c.localName : '';
      if (w === 'delText' || w === 't') out += c.textContent ?? '';
      else if (w === 'tab' || w === 'ptab') out += '\t';
      else if (w === 'br' || w === 'cr') out += '\n';
      else if (w === 'noBreakHyphen') out += '‑';
      else if (w === 'drawing' || w === 'pict' || w === 'object') out += '[圖形]';
      else if (w === 'rPr' || w === 'instrText' || w === 'delInstrText' || w === 'fldChar') continue;
      else walk(c);
    }
  };
  walk(el);
  return out;
}

/**
 * The content of a rejected deletion as ordinary content: w:delText becomes w:t,
 * w:delInstrText becomes w:instrText, and runs lose their "deleted in" rsid.
 * For w:moveFrom the content already is ordinary.
 */
export function restoredContent(xml: string): Element {
  const renamed = renamedDeletion(xml);
  if (renamed != null) return parseFragment(renamed);
  const el = parseFragment(xml);
  const doc = el.ownerDocument;
  const rename: Record<string, string> = { delText: 't', delInstrText: 'instrText' };
  for (const local of Object.keys(rename)) {
    for (const old of Array.from(el.getElementsByTagNameNS(NS.w, local))) {
      const prefix = old.prefix ? old.prefix + ':' : '';
      const repl = doc.createElementNS(NS.w, prefix + rename[local]);
      for (const a of Array.from(old.attributes)) repl.setAttributeNode(a.cloneNode() as Attr);
      while (old.firstChild) repl.appendChild(old.firstChild);
      old.parentNode!.replaceChild(repl, old);
    }
  }
  for (const r of Array.from(el.getElementsByTagNameNS(NS.w, 'r'))) r.removeAttributeNS(NS.w, 'rsidDel');
  return el;
}

/**
 * restoredContent's renames done on the text (our XML strings are written by serializeXml);
 * null when an element to rename declares namespaces itself (then it is done on the elements).
 */
function renamedDeletion(xml: string): string | null {
  if (/<\/?[\w.-]*:?del(Text|InstrText)\b[^>]*\sxmlns/.test(xml)) return null;
  return xml
    .replace(/<(\/?)([\w.-]+:)?delText\b/g, '<$1$2t')
    .replace(/<(\/?)([\w.-]+:)?delInstrText\b/g, '<$1$2instrText')
    .replace(/<([\w.-]+:)?r\s[^>]*>/g, (tag) => tag.replace(/\s([\w.-]+:)?rsidDel="[^"]*"/, ''));
}

/**
 * restoredContent of many deletions at once (Reject All), parsed a few dozen per document: by
 * the deletion's XML. One left out (not well-formed, or renamed on its elements) goes through
 * restoredContent itself. Each element is its own (one per distinct XML).
 */
export function restoredContents(xmls: Iterable<string>): Map<string, Element> {
  const originals: string[] = [];
  const renamed: string[] = [];
  for (const xml of new Set(xmls)) {
    const r = renamedDeletion(xml);
    if (r == null) continue;
    originals.push(xml);
    renamed.push(r);
  }
  const out = new Map<string, Element>();
  parseFragments(renamed).forEach((el, i) => {
    if (el) out.set(originals[i], el);
  });
  return out;
}

// ----- formatting changes -----

function hasChild(xml: string | null, local: string): boolean {
  return !!xml && new RegExp(`<(?:[\\w.-]+:)?${local}[\\s>/]`).test(xml);
}

const formatCache = new Map<string, RevisionInfo | null>();

/** The w:rPrChange inside a run's w:rPr. */
export function runFormatChange(rPr: string | null): RevisionInfo | null {
  if (!hasChild(rPr, 'rPrChange')) return null;
  let hit = formatCache.get(rPr!);
  if (hit === undefined) {
    const change = child(parseFragment(rPr!), 'rPrChange');
    hit = remember(formatCache, rPr!, change ? infoOf(change) : null);
  }
  return hit;
}

function emptyToNull(el: Element): string | null {
  return el.children.length || el.attributes.length ? serializeXml(el) : null;
}

/**
 * A w:rPr string (as serializeXml writes it) split into its start tag, its w:rPrChange and the
 * rest of its children; null when it doesn't have that plain shape.
 */
function splitRunChange(rPr: string): { open: string; close: string; attrs: boolean; change: string; rest: string } | null {
  const m = /^<(([\w.-]+:)?rPr)(\s[^>]*)?>([\s\S]*)<\/\1>$/.exec(rPr);
  if (!m) return null;
  const inner = m[4];
  const c = /<(([\w.-]+:)?rPrChange)(\s[^>]*)?(\/>|>[\s\S]*?<\/\1>)/.exec(inner);
  if (!c || /rPrChange/.test(inner.slice(c.index + c[0].length))) return null;
  return {
    open: `<${m[1]}${m[3] ?? ''}`,
    close: `</${m[1]}>`,
    attrs: !!m[3],
    change: c[0],
    rest: inner.slice(0, c.index) + inner.slice(c.index + c[0].length),
  };
}

function joinRunProps(parts: { open: string; close: string; attrs: boolean }, inner: string): string | null {
  if (inner) return `${parts.open}>${inner}${parts.close}`;
  return parts.attrs ? `${parts.open}/>` : null;
}

/** Accepting a formatting change keeps the current properties and drops the record. */
export function acceptRunFormat(rPr: string): string | null {
  const fast = splitRunChange(rPr);
  if (fast) return joinRunProps(fast, fast.rest);
  const el = parseFragment(rPr);
  const change = child(el, 'rPrChange');
  if (change) el.removeChild(change);
  return emptyToNull(el);
}

/** Rejecting one puts back the properties recorded in it. */
export function rejectRunFormat(rPr: string): string | null {
  const fast = splitRunChange(rPr);
  if (fast) {
    // The recorded properties: the children of the w:rPr inside w:rPrChange.
    const old = /^<([\w.-]+:)?rPrChange\b[^>]*>\s*(?:<(([\w.-]+:)?rPr)\b[^>]*?(?:\/>|>([\s\S]*)<\/\2>))?\s*<\/\1rPrChange>$/.exec(fast.change);
    // Revision marks of a paragraph mark (w:ins / w:del ...) are not formatting; they stay.
    const marks = fast.rest.match(/<([\w.-]+:)?(ins|del|moveFrom|moveTo)\b[^>]*\/>/g) ?? [];
    const complex = /<([\w.-]+:)?(ins|del|moveFrom|moveTo)\b[^>]*[^/]>/.test(fast.rest);
    if (old && !complex) return joinRunProps(fast, marks.join('') + (old[4] ?? ''));
  }
  const el = parseFragment(rPr);
  const change = child(el, 'rPrChange');
  if (!change) return rPr;
  const old = child(change, 'rPr');
  // Revision marks of a paragraph mark (w:ins / w:del ...) are not formatting; they stay.
  const keep = Array.from(el.children).filter((c) => c.namespaceURI === NS.w && MARK_REVISIONS.includes(c.localName));
  while (el.firstChild) el.removeChild(el.firstChild);
  for (const k of keep) el.appendChild(k);
  if (old) for (const c of Array.from(old.childNodes)) el.appendChild(c);
  return emptyToNull(el);
}

// ----- paragraphs -----

const MARK_REVISIONS = ['ins', 'del', 'moveFrom', 'moveTo'];

export interface ParagraphRevisions {
  /** A tracked insertion or deletion of the paragraph mark (w:pPr/w:rPr/w:ins|del|moveTo|moveFrom). */
  mark: { kind: 'ins' | 'del'; info: RevisionInfo } | null;
  /** A paragraph formatting change (w:pPrChange), or one of the paragraph mark's formatting (w:pPr/w:rPr/w:rPrChange). */
  format: RevisionInfo | null;
}

const paraCache = new Map<string, ParagraphRevisions>();
const NONE: ParagraphRevisions = { mark: null, format: null };

export function paragraphRevisions(pPr: string | null): ParagraphRevisions {
  if (!pPr || !/(pPrChange|rPrChange|<(?:[\w.-]+:)?(ins|del|moveFrom|moveTo)[\s>/])/.test(pPr)) return NONE;
  let hit = paraCache.get(pPr);
  if (!hit) {
    const el = parseFragment(pPr);
    const rPr = child(el, 'rPr');
    let mark: ParagraphRevisions['mark'] = null;
    for (const local of MARK_REVISIONS) {
      const m = child(rPr, local);
      if (m) mark = { kind: local === 'ins' || local === 'moveTo' ? 'ins' : 'del', info: infoOf(m) };
    }
    const change = child(el, 'pPrChange') ?? child(rPr, 'rPrChange');
    hit = remember(paraCache, pPr, { mark, format: change ? infoOf(change) : null });
  }
  return hit;
}

/**
 * Paragraph properties for the two halves of a paragraph split with Enter (tracking off), as
 * Word does it: the first half ends with a new, untracked mark (no mark revision, no mark
 * formatting change); the last half keeps the original mark. Neither keeps w:pPrChange.
 */
export function splitParagraphPPr(pPr: string | null): { first: string | null; last: string | null } {
  if (!pPr || !/(pPrChange|rPrChange|<(?:[\w.-]+:)?(ins|del|moveFrom|moveTo)[\s>/])/.test(pPr)) return { first: pPr, last: pPr };
  const drop = (el: Element | null, local: string) => {
    const c = child(el, local);
    if (c) el!.removeChild(c);
  };
  const last = parseFragment(pPr);
  drop(last, 'pPrChange');
  const first = parseFragment(serializeXml(last));
  const rPr = child(first, 'rPr');
  if (rPr) {
    for (const local of [...MARK_REVISIONS, 'rPrChange']) drop(rPr, local);
    if (!rPr.children.length && !rPr.attributes.length) first.removeChild(rPr);
  }
  return { first: emptyToNull(first), last: emptyToNull(last) };
}

/** Markers that delimit a range, and the one ending (or starting) it. */
const RANGE_PAIRS: Record<string, string> = {
  bookmarkStart: 'bookmarkEnd', bookmarkEnd: 'bookmarkStart',
  commentRangeStart: 'commentRangeEnd', commentRangeEnd: 'commentRangeStart',
  permStart: 'permEnd', permEnd: 'permStart',
};

/**
 * What stays of an accepted deletion: range markers whose other end is outside it (Word keeps
 * the bookmark or comment, now starting or ending where the deletion was). A range wholly
 * inside goes with it, and so does a comment whose range starts inside it (with its reference
 * mark). Returns the deletion element holding only what stays, or null when nothing does.
 */
export function keptMarkers(xml: string): Element | null {
  if (!/(bookmark|commentR|perm)/.test(xml)) return null;
  const el = parseFragment(xml);
  const w = (e: Element) => (e.namespaceURI === NS.w ? e.localName : '');
  const idOf = (e: Element) => e.getAttributeNS(NS.w, 'id') ?? '';
  const all = Array.from(el.getElementsByTagName('*'));
  const inside = new Set(all.filter((e) => RANGE_PAIRS[w(e)]).map((e) => `${w(e)}:${idOf(e)}`));
  const kept: Element[] = [];
  for (const e of all) {
    const local = w(e);
    if (RANGE_PAIRS[local]) {
      if (!inside.has(`${RANGE_PAIRS[local]}:${idOf(e)}`)) kept.push(e.cloneNode(true) as Element);
    } else if (local === 'commentReference' && !inside.has(`commentRangeStart:${idOf(e)}`)) {
      // The reference mark stays in a run of its own, with the run's properties.
      const run = e.parentElement!;
      const r = run.cloneNode(false) as Element;
      const rPr = child(run, 'rPr');
      if (rPr) r.appendChild(rPr.cloneNode(true));
      r.appendChild(e.cloneNode(true));
      kept.push(r);
    }
  }
  if (!kept.length) return null;
  const box = el.cloneNode(false) as Element;
  for (const k of kept) box.appendChild(k);
  return box;
}

/** w:pPr without the paragraph mark's insertion/deletion record. */
export function withoutMarkRevision(pPr: string): string | null {
  const el = parseFragment(pPr);
  const rPr = child(el, 'rPr');
  if (rPr) {
    for (const local of MARK_REVISIONS) {
      const m = child(rPr, local);
      if (m) rPr.removeChild(m);
    }
    if (!rPr.children.length && !rPr.attributes.length) el.removeChild(rPr);
  }
  return emptyToNull(el);
}

export function acceptParagraphFormat(pPr: string): string | null {
  const el = parseFragment(pPr);
  const change = child(el, 'pPrChange');
  if (change) el.removeChild(change);
  const rPr = child(el, 'rPr');
  const rChange = child(rPr, 'rPrChange');
  if (rPr && rChange) {
    rPr.removeChild(rChange);
    if (!rPr.children.length && !rPr.attributes.length) el.removeChild(rPr);
  }
  return emptyToNull(el);
}

/**
 * The paragraph properties recorded in w:pPrChange, plus what that record never covers:
 * the paragraph mark's run properties (themselves restored when they carry a change) and
 * the section break.
 */
export function rejectParagraphFormat(pPr: string): string | null {
  const el = parseFragment(pPr);
  const change = child(el, 'pPrChange');
  const rPr = child(el, 'rPr');
  const sectPr = child(el, 'sectPr');
  if (change) {
    const old = child(change, 'pPr');
    while (el.firstChild) el.removeChild(el.firstChild);
    if (old) for (const c of Array.from(old.childNodes)) el.appendChild(c);
    if (rPr) el.appendChild(rPr);
    if (sectPr) el.appendChild(sectPr);
  }
  if (rPr && child(rPr, 'rPrChange')) {
    const restored = rejectRunFormat(serializeXml(rPr));
    const next = restored ? parseFragment(restored) : null;
    if (next) el.replaceChild(el.ownerDocument.importNode(next, true), rPr);
    else el.removeChild(rPr);
  }
  return emptyToNull(el);
}

/** Paragraph properties element of a w:pPr string, for re-reading the modeled values. */
export function pPrElement(pPr: string | null): Element | null {
  return pPr ? parseFragment(pPr) : null;
}

/** A stable color per author, like Word's "by author" revision colors. */
const PALETTE = ['#c62828', '#1565c0', '#2e7d32', '#6a1b9a', '#ef6c00', '#00838f', '#ad1457', '#4e342e'];
export function authorColor(author: string): string {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) | 0;
  return PALETTE[Math.abs(h) % PALETTE.length];
}

/** "王小明，2026/01/02 10:30" for tooltips. */
export function describeRevision(verb: string, info: RevisionInfo): string {
  const who = info.author || '未知作者';
  const when = formatDate(info.date, info.dateUtc);
  return `${verb}：${who}${when ? '，' + when : ''}`;
}

/**
 * When a revision or comment was made, as Word shows it. Word writes its local time into
 * w:date with a "Z" suffix, so those digits are shown as written (converting them would be off
 * by the time zone); a real UTC time (w16du:dateUtc, w16cex:dateUtc) is shown in local time.
 */
export function formatDate(date: string | null, dateUtc?: string | null): string {
  const p = (n: number) => String(n).padStart(2, '0');
  if (dateUtc) {
    const d = new Date(dateUtc);
    if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
  if (!date) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(date.trim());
  if (!m) return date;
  return m[4] ? `${m[1]}/${m[2]}/${m[3]} ${m[4]}:${m[5]}` : `${m[1]}/${m[2]}/${m[3]}`;
}

/** The same time in 民國 years, as the UI shows dates (115/01/02 10:30); '' when there is none. */
export function rocDate(date: string | null, dateUtc?: string | null): string {
  const s = formatDate(date, dateUtc);
  const m = /^(\d{4})(\/.*)$/.exec(s);
  return m && Number(m[1]) > 1911 ? `${Number(m[1]) - 1911}${m[2]}` : s;
}

// ----- how revisions look in the page -----

/**
 * A kept deletion shows its text struck through. The text is in the page (not editable: the node
 * is an atom; not copied: the clipboard is written from the document) so screen readers read it,
 * after a hidden 「刪除：」 and before a hidden 「（刪除結束）」 (persona-300): said once, so no
 * role="deletion" as well (read as a second 「刪除」). The title (who, when) is for the mouse.
 */
export function deletionDOM(xml: string): DOMOutputSpec | null {
  const d = deletedInfo(xml);
  if (!d) return null;
  const verb = d.kind === 'moveFrom' ? '移出' : '刪除';
  return ['span', {
    class: `dx-rev-del dx-rev-${d.kind}`,
    'data-text': d.text || ' ',
    title: describeRevision(verb, d),
    style: `--dx-rev:${authorColor(d.author)}`,
  }, ['span', { class: 'dx-sr' }, `${verb}：`], ['span', { class: 'dx-rev-del-text' }, d.text || ' '], ['span', { class: 'dx-sr' }, `（${verb}結束）`]];
}

const wrapCache = new Map<string, Record<string, string> | null>();

/** Attributes for the inline wrapper span: tracked insertions are colored and underlined. */
export function insertionAttrs(layers: string): Record<string, string> | null {
  let hit = wrapCache.get(layers);
  if (hit === undefined) {
    hit = null;
    let list: { open: string }[] = [];
    try {
      list = JSON.parse(layers);
    } catch {
      list = [];
    }
    // The innermost insertion is the one that shows.
    for (let i = list.length - 1; i >= 0 && !hit; i--) {
      const ins = insertionOfLayer(list[i].open);
      if (ins) {
        hit = {
          class: `dx-rev-ins dx-rev-${ins.kind}`,
          // Said by screen readers that know it, besides the colour and underline (persona-300).
          role: 'insertion',
          title: describeRevision(ins.kind === 'moveTo' ? '移入' : '插入', ins.info),
          style: `--dx-rev:${authorColor(ins.info.author)}`,
        };
      }
    }
    remember(wrapCache, layers, hit, 2000);
  }
  return hit;
}
