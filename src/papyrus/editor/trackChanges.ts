// Tracked changes made in the editor: Word's 追蹤修訂 (Review > Track Changes).
//
// While tracking is on, every edit the user makes is recorded as Word records it, in the same
// form the editor reads tracked changes from a file (docx/revisions.ts), so review.ts shows,
// accepts and rejects them like any other and the writer saves them as Word does:
//  - typed / pasted / dropped content gets a w:ins wrapper layer (inlineWrap mark) with the
//    author and the date (local time with a "Z", as Word writes it); typing on from an insertion
//    made here by the same author extends it, so it stays one w:ins;
//  - deleted content stays, as a kept w:del node shown struck through (the runs as they would
//    be saved, with w:delText); Backspace leaves the cursor before it, Delete after it. Content
//    that is the author's own tracked insertion is simply removed, and so is a field's
//    placeholder text;
//  - a new paragraph mark (Enter, pasted paragraphs) is marked inserted, a deleted one deleted
//    (w:pPr/w:rPr/w:ins|w:del); deleting a paragraph mark the author inserted joins the paragraphs;
//  - character formatting records w:rPrChange with the properties before, paragraph formatting
//    w:pPrChange (only the first change is recorded, as Word keeps the original properties).
// What can't be recorded that way (table structure, section breaks, page breaks, pictures not
// from the file ...) is refused with a notice, never done untracked.
//
// How: the editor passes every transaction it dispatches through trackTransaction, which
// rebuilds it from the old document step by step with the edit tracked (one transaction, so the
// edit and its record are one undo step), or refuses it. Accept / reject, comments, undo / redo
// and transactions plugins append are not tracked. Doing it at dispatch (not in a plugin's
// filterTransaction) keeps other plugins' appended transactions out: ProseMirror filters those
// before it marks them as appended.
// The pending insertion is also put into the stored marks at the cursor (see trackingMarks), so
// typed text (and IME composition) already carries it and doesn't have to be rewrapped.

import { Plugin, PluginKey, Selection, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import {
  AddMarkStep, AttrStep, DocAttrStep, Mapping, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, Transform, type Step,
} from 'prosemirror-transform';
import { Fragment, Slice, type Mark, type Node as PMNode } from 'prosemirror-model';
import { isHistoryTransaction } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { NUMBERING_CHANGED } from './listMarkers';
import { newLayerId, parseLayers, type Layer } from '../docx/wrappers';
import { acceptRunFormat, insertionOfLayer } from '../docx/revisions';
import { commentDate } from '../docx/comments';
import { deletedRunXml, groupId, paraModel, runModel } from '../docx/writer';
import { writePPr, writeRPr } from '../docx/props';
import { escapeAttr, parseFragment, serializeXml, NS, child } from '../docx/xml';
import { isWatermarkRun } from '../docx/watermark';

/** Transactions with this meta are never tracked (accept / reject, comments ...). */
export const NO_TRACK = 'dx-no-track';

export interface TrackOptions {
  /** Whether tracking is on (and editing allowed). */
  enabled: () => boolean;
  /** Who makes the changes (w:author). */
  author: () => string;
  /** Why an edit was refused. */
  onNotice?: (message: string) => void;
  /** Revision ids given out so far; shared by the body and header / footer editors. */
  ids?: RevisionIds;
  /** The time a change is made (tests). */
  now?: () => Date;
}

/** New w:id values for revisions: after every id in the documents seen, never twice. */
export class RevisionIds {
  private next = 1;
  private scanned = new WeakSet<PMNode>();
  private started = false;

  /** `sources`: every part of the document (body, headers, footers), read before the first id is given. */
  constructor(private sources: () => (PMNode | null | undefined)[] = () => []) {}

  /** Make sure ids already in `doc` are never given out (read once per document). */
  see(doc: PMNode): void {
    if (this.scanned.has(doc)) return;
    this.scanned.add(doc);
    let max = 0;
    const scan = (v: unknown) => {
      if (typeof v === 'string') {
        if (v.includes(':id=')) for (const m of v.matchAll(/:id=\\?"(\d+)\\?"/g)) max = Math.max(max, Number(m[1]));
      } else if (Array.isArray(v)) v.forEach(scan);
    };
    const visit = (n: PMNode) => {
      Object.values(n.attrs).forEach(scan);
      for (const m of n.marks) Object.values(m.attrs).forEach(scan);
      n.forEach(visit);
    };
    visit(doc);
    this.next = Math.max(this.next, max + 1);
  }

  take(): number {
    if (!this.started) {
      this.started = true;
      for (const doc of this.sources()) if (doc) this.see(doc);
    }
    return this.next++;
  }
}

/** What the plugin keeps while the editor is open. */
interface Session {
  opts: TrackOptions;
  ids: RevisionIds;
  /** Wrapper layers of insertions made here (typing on from one extends it). */
  layers: Set<string>;
  /** w:id of deletions made here (deleting next to one extends it). */
  dels: Set<string>;
  /** The insertion the next typed text goes into, when it doesn't extend one. */
  pending: Layer | null;
  /** Whether the document's own revision ids were read. */
  seen: boolean;
}

export const trackKey = new PluginKey<Session>('dx-track');

const MESSAGES = {
  table: '追蹤修訂開啟時，無法插入、刪除或調整表格的結構，請先關閉「追蹤修訂」。',
  cell: '追蹤修訂開啟時，無法變更表格或儲存格的格式，請先關閉「追蹤修訂」。',
  pageBreak: '追蹤修訂開啟時，無法刪除分頁符號，請先關閉「追蹤修訂」。',
  section: '追蹤修訂開啟時，無法變更分節符號或版面設定，請先關閉「追蹤修訂」。',
  image: '追蹤修訂開啟時，無法變更或刪除這張圖片，請先關閉「追蹤修訂」。',
  field: '追蹤修訂開啟時，無法插入或刪除這個功能變數，請先關閉「追蹤修訂」。',
  object: '追蹤修訂開啟時，無法刪除這個物件，請先關閉「追蹤修訂」。',
  link: '追蹤修訂開啟時，無法新增或移除連結，請先關閉「追蹤修訂」。',
  other: '追蹤修訂開啟時，無法記錄這項變更，請先關閉「追蹤修訂」。',
} as const;

class Refused extends Error {
  constructor(readonly notice: string) {
    super(notice);
  }
}

// ----- layers and revision XML -----

const isRaw = (n: PMNode | null | undefined): n is PMNode => n?.type === schema.nodes.raw_inline;
const isDeletion = (n: PMNode | null | undefined) => isRaw(n) && (n.attrs.label === 'del' || n.attrs.label === 'moveFrom');

function layersOf(node: PMNode): Layer[] {
  return node.type === schema.nodes.page_break ? parseLayers(node.attrs.layers) : markLayers(node.marks);
}

function markLayers(marks: readonly Mark[]): Layer[] {
  const wrap = marks.find((m) => m.type === schema.marks.inlineWrap);
  return wrap ? parseLayers(wrap.attrs.layers) : [];
}

/** A new revision id; the document's own ids are read the first time. */
function takeId(s: Session, doc: PMNode): number {
  if (!s.seen) {
    s.ids.see(doc);
    s.seen = true;
  }
  return s.ids.take();
}

function withLayers(marks: readonly Mark[], layers: Layer[]): readonly Mark[] {
  const rest = schema.marks.inlineWrap.removeFromSet(marks);
  return layers.length ? schema.marks.inlineWrap.create({ layers: JSON.stringify(layers) }).addToSet(rest) : rest;
}

const isInsertionLayer = (l: Layer) => !!insertionOfLayer(l.open);

/** The innermost tracked insertion (not a move) around a node, with its author. */
function insertionAround(node: PMNode): { layer: Layer; author: string } | null {
  const layers = layersOf(node);
  for (let i = layers.length - 1; i >= 0; i--) {
    const hit = insertionOfLayer(layers[i].open);
    if (hit) return hit.kind === 'ins' ? { layer: layers[i], author: hit.info.author } : null;
  }
  return null;
}

/** A content control still showing its placeholder text (fields.ts): that text isn't content. */
const PLACEHOLDER = /^<(?:\w+:)?sdt[\s>][\s\S]*?<(?:\w+:)?showingPlcHdr\b(?![^>]*w:val="(?:0|false|off)")/;
const inPlaceholder = (node: PMNode) => layersOf(node).some((l) => PLACEHOLDER.test(l.open));

function revisionAttrs(id: number, author: string, date: string): string {
  return `w:id="${id}" w:author="${escapeAttr(author)}" w:date="${date}"`;
}

const OWN_DEL = /^<w:del w:id="(\d+)" w:author="([^"]*)" w:date="([^"]*)">([\s\S]*)<\/w:del>$/;

// ----- the plugin -----

/**
 * The tracking plugin: keeps what tracking needs while the editor is open and refreshes the
 * pending insertion's date before typing. The edits themselves are tracked where they are
 * dispatched (trackTransaction), so only what the user does is tracked, never what plugins
 * append to it.
 */
export function trackChanges(opts: TrackOptions): Plugin<Session> {
  return new Plugin<Session>({
    key: trackKey,
    state: {
      init: () => ({ opts, ids: opts.ids ?? new RevisionIds(), layers: new Set(), dels: new Set(), pending: null, seen: false }),
      apply: (_tr, s) => s,
    },
    props: {
      // The date of the insertion typed text goes into is the time of typing, not of clicking.
      handleDOMEvents: {
        keydown: (view) => refreshPending(view),
        compositionstart: (view) => refreshPending(view),
      },
    },
  });
}

/**
 * What to apply for a transaction about to be dispatched: the transaction itself when it
 * isn't tracked (tracking off, accept / reject, undo ...) or tracking changes nothing, the
 * tracked version of it, or null when it can't be tracked (refused, with a notice).
 */
export function trackTransaction(state: EditorState, tr: Transaction): Transaction | null {
  const session = trackKey.getState(state);
  if (!session || !shouldTrack(tr, session)) return tr;
  let res: Tracked;
  try {
    res = track(state, tr, session);
  } catch (e) {
    if (!(e instanceof Refused)) throw e;
    session.opts.onNotice?.(e.notice);
    return null;
  }
  if (res.doc.eq(tr.doc) && res.anchor === tr.selection.anchor && res.head === tr.selection.head) return tr;
  const out = state.tr;
  for (const step of res.steps) out.step(step);
  if (!out.doc.eq(res.doc)) replaceDifference(out, res.doc);
  const $anchor = out.doc.resolve(Math.min(res.anchor, out.doc.content.size));
  const $head = out.doc.resolve(Math.min(res.head, out.doc.content.size));
  out.setSelection(res.anchor === res.head ? Selection.near($head, res.assoc) : TextSelection.between($anchor, $head));
  // The original's meta (paste, composition, list numbers ...), time and scrolling.
  const meta = (tr as unknown as { meta: Record<string, unknown> }).meta ?? {};
  for (const k of Object.keys(meta)) out.setMeta(k, meta[k]);
  out.setTime(tr.time);
  if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks);
  if (tr.scrolledIntoView) out.scrollIntoView();
  if (res.paragraphs) out.setMeta(NUMBERING_CHANGED, true);
  return out;
}

function shouldTrack(tr: Transaction, s: Session): boolean {
  return (
    tr.docChanged &&
    s.opts.enabled() &&
    !tr.getMeta(NO_TRACK) &&
    !tr.getMeta('appendedTransaction') &&
    tr.getMeta('addToHistory') !== false &&
    !isHistoryTransaction(tr)
  );
}

/** Turn `tr.doc` into `target` with one step over the part that differs. */
function replaceDifference(tr: Transaction, target: PMNode): void {
  const doc = tr.doc;
  const a = doc.content.findDiffStart(target.content);
  if (a == null) return;
  let { a: endA, b: endB } = doc.content.findDiffEnd(target.content)!;
  const overlap = a - Math.min(endA, endB);
  if (overlap > 0) {
    endA += overlap;
    endB += overlap;
  }
  tr.replace(a, endA, target.slice(a, endB));
  if (!tr.doc.eq(target)) tr.replaceWith(0, tr.doc.content.size, target.content);
}

// ----- the insertion typed text goes into -----

function nowDate(s: Session): string {
  return commentDate(s.opts.now?.() ?? new Date()).date;
}

/** An insertion made here by the current author that the text at `pos` touches. */
function sessionLayerAt(doc: PMNode, pos: number, s: Session): Layer | null {
  const $pos = doc.resolve(pos);
  if (!$pos.parent.inlineContent) return null;
  const author = s.opts.author();
  for (const n of [$pos.nodeBefore, $pos.nodeAfter]) {
    if (!n) continue;
    const ins = insertionAround(n);
    if (ins && s.layers.has(ins.layer.id) && ins.author === author) return ins.layer;
  }
  return null;
}

/** The insertion for new text that extends none: made once, dated now. */
function pendingLayer(s: Session, doc: PMNode): Layer {
  const date = nowDate(s);
  const author = s.opts.author();
  const p = s.pending;
  if (p) {
    const info = insertionOfLayer(p.open)?.info;
    if (info && info.date === date && info.author === author) return p;
  }
  s.pending = { id: newLayerId(), open: `<w:ins ${revisionAttrs(takeId(s, doc), author, date)}>`, close: '</w:ins>' };
  return s.pending;
}

/**
 * The marks new text at the cursor gets while tracking is on: `marks` (without insertion
 * layers) plus the insertion it goes into. Used by review.ts for the stored marks.
 */
export function trackingMarks(state: EditorState, marks: readonly Mark[]): readonly Mark[] {
  const s = trackKey.getState(state);
  if (!s || !s.opts.enabled()) return marks;
  const { $from } = state.selection;
  if (!$from.parent.inlineContent) return marks;
  const layer = sessionLayerAt(state.doc, state.selection.from, s) ?? pendingLayer(s, state.doc);
  const layers = markLayers(marks).filter((l) => !isInsertionLayer(l));
  return withLayers(marks, [...layers, layer]);
}

/** Before a key or IME input: an insertion dated earlier than now gets the current time. */
function refreshPending(view: EditorView): boolean {
  const s = trackKey.getState(view.state);
  const stored = view.state.storedMarks;
  if (!s || !s.pending || !stored || !s.opts.enabled() || !view.editable) return false;
  const old = s.pending;
  if (!markLayers(stored).some((l) => l.id === old.id)) return false;
  const fresh = pendingLayer(s, view.state.doc);
  if (fresh === old) return false;
  const layers = markLayers(stored).map((l) => (l.id === old.id ? fresh : l));
  view.dispatch(view.state.tr.setStoredMarks(withLayers(stored, layers)).setMeta('addToHistory', false));
  return false;
}

// ----- tracking one transaction -----

interface Tracked {
  doc: PMNode;
  steps: readonly Step[];
  anchor: number;
  head: number;
  assoc: 1 | -1;
  /** Paragraphs joined or marked (list numbers may change). */
  paragraphs: boolean;
}

/** A position in `out` recorded after step `at` of it. */
interface Recorded {
  pos: number;
  at: number;
}

class Tracker {
  readonly out: Transform;
  readonly author: string;
  readonly date: string;
  /** The insertion used for this edit (one per transaction). */
  private layer: Layer | null = null;
  /** Paragraphs (start positions) whose mark was inserted / deleted. */
  readonly insMarks: Recorded[] = [];
  readonly delMarks: Recorded[] = [];
  paragraphs = false;

  constructor(doc: PMNode, readonly s: Session) {
    this.out = new Transform(doc);
    this.author = s.opts.author();
    this.date = nowDate(s);
  }

  newId(): number {
    return takeId(this.s, this.out.doc);
  }

  record(list: Recorded[], pos: number): void {
    list.push({ pos, at: this.out.steps.length });
  }

  mapRecorded(r: Recorded): number {
    return this.out.mapping.slice(r.at).map(r.pos, 1);
  }

  // ----- deleting -----

  /** Mark [a, b) of `out` deleted. */
  deleteRange(a: number, b: number): void {
    if (b <= a) return;
    const doc = this.out.doc;
    type Seg = { from: number; to: number; node: PMNode; remove: boolean; xml?: string };
    const segs: Seg[] = [];
    const paras: number[] = [];
    doc.nodesBetween(a, b, (node, pos, parent, index) => {
      const end = pos + node.nodeSize;
      if (node.isTextblock) {
        // Its mark is deleted when the range runs past the end of its content.
        const e = end - 1;
        if (e >= a && e < b && parent && index < parent.childCount - 1) paras.push(pos);
        return true;
      }
      if (node.isInline) {
        const from = Math.max(a, pos);
        const to = Math.min(b, end);
        if (from < to) {
          const seg = this.classify(node, pos, from, to);
          if (seg) segs.push(seg);
        }
        return false;
      }
      if (pos < a && end > b) return true; // holds the whole range (a table cell ...)
      if (node.type === schema.nodes.raw_block && node.attrs.hidden) return false;
      if (node.type === schema.nodes.page_break) throw new Refused(MESSAGES.pageBreak);
      if (node.type.spec.tableRole || node.type === schema.nodes.table) throw new Refused(MESSAGES.table);
      throw new Refused(MESSAGES.object);
    });
    // Group neighbouring deleted pieces with the same marks into one w:del, last first.
    const groups: Seg[][] = [];
    for (const seg of segs) {
      const last = groups[groups.length - 1];
      const prev = last?.[last.length - 1];
      if (prev && !prev.remove && !seg.remove && prev.to === seg.from && doc.resolve(seg.from).parent === doc.resolve(prev.from).parent &&
        sameMarks(prev.node.marks, seg.node.marks)) last.push(seg);
      else groups.push([seg]);
    }
    const at = this.out.steps.length;
    for (const g of groups.reverse()) {
      const from = g[0].from;
      const to = g[g.length - 1].to;
      if (g[0].remove) this.out.delete(from, to);
      else this.putDeletion(from, to, g[0].node.marks, g.map((x) => x.xml!).join(''), refsOf(g.map((x) => x.node)));
    }
    for (const p of paras) this.delMarks.push({ pos: p, at });
  }

  /** What deleting (part of) an inline node does; null: it stays as it is. */
  private classify(node: PMNode, pos: number, from: number, to: number) {
    if (isDeletion(node)) return null;
    // A header's watermark (浮水印) is not deleted by editing around it, tracked or not: it
    // changes only with 設計 › 浮水印 (see keepWatermarks in editor/watermark.ts).
    if (isRaw(node) && isWatermarkRun(node.attrs.xml)) return null;
    if (isRaw(node)) {
      const xml = node.attrs.xml as string;
      // Range markers, comment marks and empty insertions stay where they are.
      if (node.attrs.label === '註解' || !/^<(?:[\w.-]+:)?r[\s>/]/.test(xml)) {
        if (node.attrs.hidden || node.attrs.label === '註解' || !xml) return null;
        throw new Refused(MESSAGES.object);
      }
    }
    const own = insertionAround(node);
    if ((own && own.author === this.author) || inPlaceholder(node)) return { from, to, node, remove: true };
    if (isRaw(node) && /<(?:[\w.-]+:)?(footnoteReference|endnoteReference)\b/.test(node.attrs.xml)) throw new Refused(MESSAGES.object);
    if (node.type === schema.nodes.field) throw new Refused(MESSAGES.field);
    const piece = node.isText ? schema.text(node.text!.slice(from - pos, to - pos), node.marks) : node;
    const xml = deletedRunXml(piece);
    if (xml == null) throw new Refused(node.type === schema.nodes.image ? MESSAGES.image : MESSAGES.object);
    return { from, to, node, remove: false, xml };
  }

  /** Replace [from, to) with a w:del holding `runs`, joined with a deletion made here next to it. */
  private putDeletion(from: number, to: number, marks: readonly Mark[], runs: string, refs: Refs | null): void {
    const doc = this.out.doc;
    const before = doc.resolve(from).nodeBefore;
    const after = doc.resolve(to).nodeAfter;
    const own = (n: PMNode | null) => (n && isDeletion(n) && sameMarks(n.marks, marks) ? this.ownDeletion(n) : null);
    const b = own(before);
    const f = own(after);
    let start = from;
    let end = to;
    let open = `<w:del ${revisionAttrs(0, this.author, this.date)}>`;
    let inner = runs;
    let allRefs = refs;
    if (b) {
      start -= before!.nodeSize;
      open = b.open;
      inner = b.inner + inner;
      allRefs = mergeRefs(before!.attrs.refs, allRefs);
    }
    if (f) {
      end += after!.nodeSize;
      if (!b) open = f.open;
      inner += f.inner;
      allRefs = mergeRefs(allRefs, after!.attrs.refs);
    }
    if (!b && !f) {
      const id = this.newId();
      this.s.dels.add(String(id));
      open = `<w:del ${revisionAttrs(id, this.author, this.date)}>`;
    }
    const node = schema.nodes.raw_inline.create({ xml: `${open}${inner}</w:del>`, label: 'del', hidden: true, refs: allRefs }, null, marks);
    this.out.replaceWith(start, end, node);
  }

  private ownDeletion(n: PMNode): { open: string; inner: string } | null {
    const m = OWN_DEL.exec(n.attrs.xml);
    if (!m || !this.s.dels.has(m[1]) || m[2] !== escapeAttr(this.author)) return null;
    return { open: n.attrs.xml.slice(0, n.attrs.xml.indexOf('>') + 1), inner: m[4] };
  }

  // ----- inserting -----

  /** The insertion for this edit: the one the text extends, the pending one, or a new one. */
  private insertionLayer(pos: number, slice: Slice): Layer {
    if (this.layer) return this.layer;
    let found: Layer | null = null;
    slice.content.descendants((n) => {
      if (found) return false;
      const ins = n.isInline || n.type === schema.nodes.page_break ? insertionAround(n) : null;
      if (ins && ins.author === this.author && (this.s.layers.has(ins.layer.id) || ins.layer.id === this.s.pending?.id)) found = ins.layer;
      return !found;
    });
    let layer = found ?? sessionLayerAt(this.out.doc, pos, this.s);
    if (!layer) {
      const p = this.s.pending;
      const info = p ? insertionOfLayer(p.open)?.info : null;
      layer = p && info?.author === this.author && info.date === this.date ? p : null;
    }
    if (!layer) layer = { id: newLayerId(), open: `<w:ins ${revisionAttrs(this.newId(), this.author, this.date)}>`, close: '</w:ins>' };
    else if (layer.id === this.s.pending?.id) {
      // Typed into (or pasted with) the pending insertion: it is used now, dated now.
      const info = insertionOfLayer(layer.open)!.info;
      if (info.date !== this.date) layer = { ...layer, open: layer.open.replace(/w:date="[^"]*"/, `w:date="${this.date}"`) };
    }
    if (layer.id === this.s.pending?.id) this.s.pending = null;
    this.s.layers.add(layer.id);
    this.layer = layer;
    return layer;
  }

  /** Insert `slice` at `pos` of `out` as a tracked insertion; returns the end of what went in. */
  insert(pos: number, slice: Slice): number {
    slice.content.descendants((n) => {
      if (n.type === schema.nodes.table || n.type.spec.tableRole) throw new Refused(MESSAGES.table);
      if (n.type === schema.nodes.raw_block && !n.attrs.hidden) throw new Refused(MESSAGES.object);
      if (n.type === schema.nodes.field) throw new Refused(MESSAGES.field);
      return true;
    });
    const layer = this.insertionLayer(pos, slice);
    const content = markInserted(slice.content, layer);
    const before = this.out.steps.length;
    this.out.replace(pos, pos, new Slice(content, slice.openStart, slice.openEnd));
    if (this.out.steps.length === before) return pos;
    let start = pos;
    let end = pos;
    this.out.steps[this.out.steps.length - 1].getMap().forEach((_a, _b, newStart, newEnd) => {
      start = newStart;
      end = newEnd;
    });
    // Paragraphs whose mark came with the insertion (Enter, pasted paragraphs).
    this.out.doc.nodesBetween(start, end, (node, p) => {
      if (!node.isTextblock) return true;
      const e = p + node.nodeSize - 1;
      if (e >= start && e < end) this.record(this.insMarks, p);
      return false;
    });
    return end;
  }

  // ----- formatting -----

  /** The document before the edit's first formatting step, and where that step is in `out`. */
  private formatBase: { doc: PMNode; at: number } | null = null;
  /** Ranges formatting steps changed (in `out` after step `at`). */
  private formatRanges: (Recorded & { to: number })[] = [];

  /** Before a formatting (mark) step is applied to `out`. */
  beforeFormat(): void {
    this.formatBase ??= { doc: this.out.doc, at: this.out.steps.length };
  }

  /** After a formatting step over [from, to) of `out`. */
  noteFormat(from: number, to: number): void {
    if (from < to) this.formatRanges.push({ pos: from, to, at: this.out.steps.length });
  }

  /**
   * At the end of the edit: text whose run properties changed records those it had before
   * (w:rPrChange), unless it already has a record (Word keeps the first) or is the author's own
   * insertion. Compared once, with the document before the edit's formatting, so commands that
   * format in several steps (clear formatting ...) make one record.
   */
  finishFormats(): void {
    const base = this.formatBase;
    if (!base) return;
    const forward = this.out.mapping.slice(base.at);
    const back = forward.invert();
    const done = new Set<number>();
    const changes: { from: number; to: number; mark: Mark }[] = [];
    const isRunContent = (n: PMNode) => n.isText || n.type === schema.nodes.tab || n.type === schema.nodes.hard_break;
    for (const r of this.formatRanges) {
      const map = this.out.mapping.slice(r.at);
      const from = map.map(r.pos, 1);
      const to = map.map(r.to, -1);
      if (from >= to) continue;
      this.out.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isInline) return true;
        if (!isRunContent(node)) return false;
        const own = insertionAround(node);
        if (own && own.author === this.author) return false;
        const run = node.marks.find((m) => m.type === schema.marks.run);
        if (/rPrChange/.test(run?.attrs.rPr ?? '')) return false; // the first change is kept
        const now = writeRPr(run?.attrs.rPr ?? null, runModel(node.marks));
        const start = Math.max(pos, from);
        const end = Math.min(pos + node.nodeSize, to);
        // Piece by piece as the text was before (it may not all have been formatted alike).
        base.doc.nodesBetween(back.map(start, 1), back.map(end, -1), (old, oldPos) => {
          if (!old.isInline) return true;
          const s = Math.max(start, forward.map(oldPos, 1));
          const e = Math.min(end, forward.map(oldPos + old.nodeSize, -1));
          if (s >= e || done.has(s) || !isRunContent(old)) return false;
          done.add(s);
          const oldRun = old.marks.find((m) => m.type === schema.marks.run);
          if (/rPrChange/.test(oldRun?.attrs.rPr ?? '')) return false;
          const was = writeRPr(oldRun?.attrs.rPr ?? null, runModel(old.marks));
          if (was === now) return false;
          const rPr = withRunChange(run?.attrs.rPr ?? null, was, revisionAttrs(this.newId(), this.author, this.date));
          changes.push({ from: s, to: e, mark: schema.marks.run.create({ rPr, attrs: run?.attrs.attrs ?? null }) });
          return false;
        });
        return false;
      });
    }
    for (const c of changes) {
      this.out.removeMark(c.from, c.to, schema.marks.run);
      this.out.addMark(c.from, c.to, c.mark);
    }
  }

  /** New attributes for a paragraph whose formatting changes: w:pPrChange with the properties before. */
  paragraphChange(node: PMNode, attrs: Record<string, any>): Record<string, any> {
    if (sameModel(node.attrs, attrs)) return attrs;
    const pPr = attrs.pPr as string | null;
    if (pPr && /pPrChange/.test(pPr)) return attrs; // the first change is kept
    if (markInsertedBy(node.attrs.pPr, this.author)) return attrs; // a paragraph this author added
    const was = writePPr(node.attrs.pPr ?? null, paraModel(node.attrs), null);
    this.paragraphs = true;
    return { ...attrs, pPr: withParaChange(pPr, was, revisionAttrs(this.newId(), this.author, this.date)) };
  }

  // ----- paragraph marks, at the end -----

  finishMarks(): void {
    const out = this.out;
    for (const r of this.insMarks) {
      const pos = this.mapRecorded(r);
      const node = out.doc.nodeAt(pos);
      if (!node || node.type !== schema.nodes.paragraph) continue;
      const next = out.doc.resolve(pos + node.nodeSize).nodeAfter;
      // Pieces of one w:p split at a page break: no paragraph mark between them.
      if (next && groupId(node) && groupId(next) === groupId(node)) continue;
      this.setMark(pos, 'ins');
    }
    const dels = this.delMarks.map((r) => this.mapRecorded(r)).sort((x, y) => y - x);
    for (const pos of dels) {
      const node = out.doc.nodeAt(pos);
      if (!node || node.type !== schema.nodes.paragraph) continue;
      const after = pos + node.nodeSize;
      const next = out.doc.resolve(after).nodeAfter;
      if (markInsertedBy(node.attrs.pPr, this.author) && next?.type === schema.nodes.paragraph && !groupId(node) && !groupId(next)) {
        // The author's own new paragraph mark: removed, the paragraphs are one again (with the next one's mark).
        this.paragraphs = true;
        const attrs = next.attrs;
        out.join(after);
        out.setNodeMarkup(pos, undefined, attrs);
        continue;
      }
      if (markRevision(node.attrs.pPr) === 'del') continue;
      this.setMark(pos, 'del');
    }
  }

  /** Mark the paragraph mark at `pos` (every piece of its w:p) inserted or deleted. */
  private setMark(pos: number, kind: 'ins' | 'del'): void {
    this.paragraphs = true;
    const tag = `<w:${kind} ${revisionAttrs(this.newId(), this.author, this.date)}/>`;
    for (const { node, pos: p } of groupPieces(this.out.doc, pos)) {
      this.out.setNodeMarkup(p, undefined, { ...node.attrs, pPr: withMarkRevision(node.attrs.pPr, kind, tag) }, node.marks);
    }
  }
}

/** Inline content (and page breaks) of an insertion get its layer; other insertions and deletions don't come along. */
function markInserted(content: Fragment, layer: Layer): Fragment {
  const nodes: PMNode[] = [];
  content.forEach((node) => {
    if (isDeletion(node)) return;
    if (node.isInline) {
      let marks = node.marks;
      const run = marks.find((m) => m.type === schema.marks.run);
      // Text typed next to a formatting change doesn't take the change along.
      if (run && /rPrChange/.test(run.attrs.rPr ?? '')) {
        marks = schema.marks.run.removeFromSet(marks);
        const rPr = acceptRunFormat(run.attrs.rPr);
        if (rPr || run.attrs.attrs) marks = schema.marks.run.create({ rPr, attrs: run.attrs.attrs }).addToSet(marks);
      }
      const keep = layersOf(node.mark(marks)).filter((l) => !isInsertionLayer(l));
      nodes.push(node.mark(withLayers(marks, [...keep, layer])));
    } else if (node.type === schema.nodes.page_break) {
      const keep = parseLayers(node.attrs.layers).filter((l) => !isInsertionLayer(l));
      nodes.push(node.type.create({ ...node.attrs, layers: JSON.stringify([...keep, layer]) }, null, node.marks));
    } else {
      nodes.push(node.copy(markInserted(node.content, layer)));
    }
  });
  return Fragment.fromArray(nodes);
}

interface Refs {
  rels: Record<string, { target: string; external: boolean }>;
  images: Record<string, string>;
}

/** Pictures a deletion holds (by relationship id), so rejecting it gives them back. */
function refsOf(nodes: PMNode[]): Refs | null {
  let out: Refs | null = null;
  for (const n of nodes) {
    if (n.type !== schema.nodes.image || !n.attrs.xml) continue;
    for (const m of (n.attrs.xml as string).matchAll(/\br:(?:embed|link|id)="([^"]+)"/g)) {
      out ??= { rels: {}, images: {} };
      out.images[m[1]] = n.attrs.src;
    }
  }
  return out;
}

function mergeRefs(a: Refs | null, b: Refs | null): Refs | null {
  if (!a || !b) return a ?? b;
  return { rels: { ...a.rels, ...b.rels }, images: { ...a.images, ...b.images } };
}

function sameMarks(a: readonly Mark[], b: readonly Mark[]): boolean {
  return a.length === b.length && a.every((m, i) => m.eq(b[i]));
}

/**
 * The paragraph attributes that are modelled paragraph properties: a change to one of them is
 * recorded as w:pPrChange. 複製格式 copies these, less list membership and page break before.
 */
export const MODEL_KEYS = [
  'styleId', 'align', 'indLeft', 'indRight', 'indFirst', 'indLeftChars', 'indRightChars', 'indFirstChars',
  'spaceBefore', 'spaceAfter', 'line', 'lineRule', 'numId', 'ilvl', 'pageBreakBefore',
];
function sameModel(a: Record<string, any>, b: Record<string, any>): boolean {
  return MODEL_KEYS.every((k) => (a[k] ?? null) === (b[k] ?? null) || (k === 'ilvl' && !a.numId && !b.numId));
}

// ----- property XML -----

const MARK_TAGS = ['ins', 'del', 'moveFrom', 'moveTo'];

/** Whether a paragraph mark is inserted / deleted (w:pPr/w:rPr/w:ins|w:del). */
function markRevision(pPr: string | null): 'ins' | 'del' | null {
  if (!pPr || !/<(?:[\w.-]+:)?(ins|del|moveFrom|moveTo)\b/.test(pPr)) return null;
  const rPr = child(parseFragment(pPr), 'rPr');
  if (child(rPr, 'del') || child(rPr, 'moveFrom')) return 'del';
  if (child(rPr, 'ins') || child(rPr, 'moveTo')) return 'ins';
  return null;
}

function markInsertedBy(pPr: string | null, author: string): boolean {
  if (!pPr || !/<(?:[\w.-]+:)?ins\b/.test(pPr)) return false;
  const ins = child(child(parseFragment(pPr), 'rPr'), 'ins');
  return !!ins && ins.getAttributeNS(NS.w, 'author') === author && !child(child(parseFragment(pPr), 'rPr'), 'del');
}

/** w:pPr with the paragraph mark marked inserted or deleted (an inserted mark keeps nothing else). */
function withMarkRevision(pPr: string | null, kind: 'ins' | 'del', tag: string): string {
  const el = parseFragment(pPr || '<w:pPr/>');
  let rPr = child(el, 'rPr');
  if (!rPr) {
    rPr = el.ownerDocument.createElementNS(NS.w, 'w:rPr');
    const at = child(el, 'sectPr') ?? child(el, 'pPrChange');
    el.insertBefore(rPr, at);
  }
  if (kind === 'ins') {
    for (const local of MARK_TAGS) {
      const old = child(rPr, local);
      if (old) rPr.removeChild(old);
    }
  }
  const mark = parseFragment(tag);
  const ins = child(rPr, 'ins');
  // CT_ParaRPr: w:ins, w:del, w:moveFrom, w:moveTo come first, in that order.
  rPr.insertBefore(el.ownerDocument.importNode(mark, true), kind === 'del' && ins ? ins.nextSibling : rPr.firstChild);
  return serializeXml(el);
}

/** A run's w:rPr (`current`) with a w:rPrChange (last) recording `was` (as it was written) as the properties before. */
function withRunChange(current: string | null, was: string, attrs: string): string {
  const inner = was ? /^<(?:[\w.-]+:)?rPr\b[^>]*?(?:\/>|>([\s\S]*)<\/(?:[\w.-]+:)?rPr>)$/.exec(was)?.[1] ?? '' : '';
  const change = `<w:rPrChange ${attrs}>${inner ? `<w:rPr>${inner}</w:rPr>` : '<w:rPr/>'}</w:rPrChange>`;
  if (!current) return `<w:rPr>${change}</w:rPr>`;
  return current.replace(/(\/>|<\/(?:[\w.-]+:)?rPr>)$/, (end) => (end === '/>' ? `>${change}</w:rPr>` : change + end));
}

/** w:pPr with a w:pPrChange (last) recording `was` (without the mark's run properties and section break). */
function withParaChange(pPr: string | null, was: string, attrs: string): string {
  const old = parseFragment(was || '<w:pPr/>');
  for (const local of ['rPr', 'sectPr', 'pPrChange']) {
    const c = child(old, local);
    if (c) old.removeChild(c);
  }
  const el = parseFragment(pPr || '<w:pPr/>');
  const change = parseFragment(`<w:pPrChange ${attrs}>${serializeXml(old)}</w:pPrChange>`);
  el.appendChild(el.ownerDocument.importNode(change, true));
  return serializeXml(el);
}

/** The paragraph at `pos` and the other pieces of its w:p (split at page breaks). */
function groupPieces(doc: PMNode, pos: number): { node: PMNode; pos: number }[] {
  const $p = doc.resolve(pos);
  const parent = $p.parent;
  const index = $p.index();
  const node = parent.child(index);
  const id = groupId(node);
  if (!id) return [{ node, pos }];
  let i = index;
  let start = pos;
  while (i > 0 && groupId(parent.child(i - 1)) === id) start -= parent.child(--i).nodeSize;
  const out: { node: PMNode; pos: number }[] = [];
  for (let p = start; i < parent.childCount && groupId(parent.child(i)) === id; p += parent.child(i++).nodeSize) {
    if (parent.child(i).type === schema.nodes.paragraph) out.push({ node: parent.child(i), pos: p });
  }
  return out;
}

// ----- the steps of a transaction -----

/** A step that only changes one node's attributes: that node and its new attributes. */
function attrChange(step: Step, doc: PMNode): { pos: number; node: PMNode; attrs: Record<string, any> } | null {
  if (step instanceof AttrStep) {
    const s = step as unknown as { pos: number; attr: string; value: unknown };
    const node = doc.nodeAt(s.pos);
    return node ? { pos: s.pos, node, attrs: { ...node.attrs, [s.attr]: s.value } } : null;
  }
  const s = step as unknown as { from: number; to: number; gapFrom?: number; gapTo?: number; insert?: number; slice: Slice };
  if (step instanceof ReplaceAroundStep) {
    const repl = s.slice.content.firstChild;
    const node = doc.nodeAt(s.from);
    if (s.gapFrom === s.from + 1 && s.gapTo === s.to - 1 && s.insert === 1 && s.slice.content.childCount === 1 && repl &&
      !repl.content.size && node && node.nodeSize === s.to - s.from && repl.type === node.type) {
      return { pos: s.from, node, attrs: repl.attrs };
    }
    return null;
  }
  if (step instanceof ReplaceStep) {
    const node = doc.nodeAt(s.from);
    const repl = s.slice.content.firstChild;
    if (node && node.isLeaf && !node.isText && s.to === s.from + node.nodeSize && s.slice.openStart === 0 && s.slice.openEnd === 0 &&
      s.slice.content.childCount === 1 && repl && repl.type === node.type) {
      return { pos: s.from, node, attrs: repl.attrs };
    }
  }
  return null;
}

/** The document with transaction `tr` tracked, and where its selection goes; throws Refused. */
function track(state: EditorState, tr: Transaction, s: Session): Tracked {
  const t = new Tracker(state.doc, s);
  const out = t.out;
  let m = new Mapping();
  // Backspace: the cursor ends before what it deleted (Delete: after it).
  let assoc: 1 | -1 = 1;
  const sel = state.selection;
  tr.steps.forEach((step, i) => {
    const doc = tr.docs[i];
    const before = out.steps.length;
    let mirror = -1;
    const change = attrChange(step, doc);
    if (step instanceof DocAttrStep) {
      if ((step as unknown as { attr: string }).attr === 'sectPr') throw new Refused(MESSAGES.section);
      out.step(step);
    } else if (change) {
      const pos = m.map(change.pos, 1);
      const node = out.doc.nodeAt(pos);
      if (!node || node.type !== change.node.type) throw new Refused(MESSAGES.other);
      let attrs = change.attrs;
      const type = node.type;
      if (type === schema.nodes.paragraph) attrs = t.paragraphChange(node, attrs);
      else if (type === schema.nodes.image) {
        const own = insertionAround(node);
        if (!own || own.author !== t.author) throw new Refused(MESSAGES.image);
      } else if (type === schema.nodes.table || type.spec.tableRole) throw new Refused(MESSAGES.cell);
      else if (type !== schema.nodes.page_break && type !== schema.nodes.raw_inline && type !== schema.nodes.raw_block) {
        throw new Refused(MESSAGES.other);
      }
      out.setNodeMarkup(pos, undefined, attrs, node.marks);
    } else if (step instanceof ReplaceStep) {
      const r = step as unknown as { from: number; to: number; slice: Slice };
      const a = m.map(r.from, 1);
      const b = Math.max(a, m.map(r.to, -1));
      if (r.slice.size === 0 && sel.empty && r.to === sel.head && tr.steps.length === 1) assoc = -1;
      t.deleteRange(a, b);
      const at = out.mapping.slice(before).map(b, 1);
      const filler = a < b && !r.slice.openStart && !r.slice.openEnd && !hasContent(r.slice.content);
      if (r.slice.size && !filler) {
        const n = out.steps.length;
        t.insert(at, r.slice);
        // Positions inside the inserted content map to the same place in it.
        if (out.steps.length === n + 1 && sizeOf(out.steps[n]) === r.slice.size) mirror = n;
      }
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const mark = (step as unknown as { mark: Mark }).mark;
      if (mark.type === schema.marks.link) linkChange(step, doc, t.author);
      const mapped = step.map(m);
      const format = FORMAT_MARKS.has(mark.type.name);
      if (mapped) {
        if (format) t.beforeFormat();
        out.step(mapped);
        const r = mapped as unknown as { from: number; to: number };
        if (format) t.noteFormat(r.from, r.to);
      }
    } else if (step instanceof ReplaceAroundStep) {
      throw new Refused(MESSAGES.other);
    } else {
      const mapped = step.map(m);
      if (mapped && out.maybeStep(mapped).failed) throw new Refused(MESSAGES.other);
    }
    // Positions after this step, in `out`.
    const next = new Mapping();
    next.appendMap(step.getMap().invert());
    next.appendMapping(m);
    for (let k = before; k < out.steps.length; k++) {
      next.appendMap(out.steps[k].getMap(), k === mirror ? 0 : undefined);
    }
    m = next;
  });
  const end = out.steps.length;
  t.finishFormats();
  t.finishMarks();
  const map = (pos: number, side: 1 | -1) => out.mapping.slice(end).map(m.map(pos, side), side);
  const head = map(tr.selection.head, assoc);
  const anchor = tr.selection.empty ? head : map(tr.selection.anchor, tr.selection.anchor < tr.selection.head ? -1 : 1);
  return { doc: out.doc, steps: out.steps, anchor, head, assoc, paragraphs: t.paragraphs };
}

const FORMAT_MARKS = new Set(['run', 'charStyle', 'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'color', 'highlight', 'font', 'fontSize']);

function hasContent(content: Fragment): boolean {
  let found = false;
  content.descendants((n) => {
    if (n.isInline || n.type === schema.nodes.page_break) found = true;
    return !found;
  });
  return found;
}

function sizeOf(step: Step): number {
  let size = -1;
  step.getMap().forEach((a, b, c, d) => {
    size = a === b ? d - c : -1;
  });
  return size;
}

/** Adding or removing a link on text that isn't the author's own insertion can't be recorded. */
function linkChange(step: Step, doc: PMNode, author: string): void {
  const r = step as unknown as { from: number; to: number };
  doc.nodesBetween(r.from, r.to, (node) => {
    if (!node.isInline) return true;
    const own = insertionAround(node);
    if (!own || own.author !== author) throw new Refused(MESSAGES.link);
    return false;
  });
}
