// Review: tracked changes (show, navigate, accept / reject) and comment ranges.
//
// Accepting or rejecting must give exactly what Word gives:
//  - insertion (w:ins / w:moveTo layer): accept drops the wrapper, reject removes the content;
//  - deletion (kept w:del / w:moveFrom node): accept removes it (range markers whose other end
//    is outside it stay), reject turns it back into ordinary runs with their original properties;
//  - formatting change (w:rPrChange / w:pPrChange): accept drops the record, reject restores
//    the recorded properties;
//  - paragraph mark: accepting a deleted mark (or rejecting an inserted one) joins the
//    paragraph with the next one in document order (the first cell of a table that follows),
//    which keeps its own properties (as Word does); with nothing to join, an empty paragraph
//    goes, another one just loses the record.
// A paragraph split at page breaks (pieces sharing brGroup, see convert.ts) is one w:p: its
// paragraph revisions are listed and resolved once, for all its pieces.
// Every accept / reject is one transaction, so one undo step.
//
// With 追蹤修訂 (tracking) off, typing is not tracked: text typed next to or inside a tracked
// insertion is not part of it, as in Word with tracking off. With it on (trackChanges.ts), what
// the user types, pastes, deletes or formats is recorded in the same forms as above (w:ins
// layers, kept w:del nodes, paragraph mark revisions, w:rPrChange / w:pPrChange) with the
// author and date, so everything here shows, accepts and rejects it like changes read from a
// file; accept / reject and comment edits themselves are never tracked (NO_TRACK).

import {
  NodeSelection, Plugin, PluginKey, Selection, TextSelection,
  type Command, type EditorState, type Transaction,
} from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { Mapping, ReplaceAroundStep, ReplaceStep, Transform, canJoin, replaceStep } from 'prosemirror-transform';
import { Fragment, type Mark, type Node as PMNode } from 'prosemirror-model';
import { schema } from './schema';
import { NUMBERING_CHANGED } from './listMarkers';
import { Converter, modelMarks } from '../docx/convert';
import { readParaModel, readRunModel } from '../docx/props';
import { newLayerId, parseLayers, type Layer } from '../docx/wrappers';
import { NS, parseFragment, serializeXml } from '../docx/xml';
import { groupId, paragraphGroupSize } from '../docx/writer';
import {
  acceptParagraphFormat, acceptRunFormat, authorColor, deletedInfo, describeRevision, insertionOfLayer, keptMarkers,
  paragraphRevisions, pPrElement, rejectParagraphFormat, rejectRunFormat, remember, restoredContent, restoredContents, runFormatChange,
  splitParagraphPPr, withoutMarkRevision, rocDate, type RevisionInfo,
} from '../docx/revisions';
import { threadRoot, withReplies, type DocComment } from '../docx/comments';
import { closeHistory } from 'prosemirror-history';
import { NO_TRACK, trackingMarks } from './trackChanges';

export type RevisionKind = 'ins' | 'del' | 'format' | 'paraFormat' | 'paraMark';

export interface Revision extends RevisionInfo {
  key: string;
  kind: RevisionKind;
  /** Tracked moves are shown and handled as an insertion / deletion. */
  move: boolean;
  /** For paraMark: whether the mark was inserted or deleted. */
  mark?: 'ins' | 'del';
  /** Range in the document: the content, or the paragraph end for a paragraph mark. */
  from: number;
  to: number;
  /** The paragraph it is in: start and end of the nodes written as one w:p. */
  group: number;
  groupEnd: number;
  /** ins: the wrapper layer (null for an empty w:ins kept as a marker). */
  layerId?: string | null;
  /** format: the run properties carrying the change. */
  rPr?: string;
  /** paraFormat / paraMark: position of the paragraph (its first piece). */
  para?: number;
  /** paraMark: there is a paragraph to join with (else an empty paragraph goes, or only the record). */
  joinable?: boolean;
}

const KIND_LABEL: Record<RevisionKind, string> = {
  ins: '插入', del: '刪除', format: '格式變更', paraFormat: '段落格式變更', paraMark: '段落標記',
};

export function revisionLabel(r: Revision): string {
  if (r.kind === 'paraMark') return r.mark === 'ins' ? '插入段落標記' : '刪除段落標記';
  if (r.move) return r.kind === 'ins' ? '移入' : '移出';
  return KIND_LABEL[r.kind];
}

// ----- finding revisions -----

const layerCache = new Map<string, Layer[]>();
function parsedLayers(json: string | null | undefined): Layer[] {
  if (!json) return [];
  return layerCache.get(json) ?? remember(layerCache, json, parseLayers(json));
}

/** Inline wrappers around a node (its inlineWrap mark, or a page break's own). */
function layersOf(node: PMNode): Layer[] {
  if (node.type === schema.nodes.page_break) return parsedLayers(node.attrs.layers);
  const wrap = node.marks.find((m) => m.type === schema.marks.inlineWrap);
  return wrap ? parsedLayers(wrap.attrs.layers) : [];
}

function isDeletion(node: PMNode): boolean {
  return node.type === schema.nodes.raw_inline && (node.attrs.label === 'del' || node.attrs.label === 'moveFrom');
}

// Not type predicates: a node that is neither is still a node.
const isPara = (n: PMNode | null | undefined): boolean => n?.type === schema.nodes.paragraph;
const isBreak = (n: PMNode | null | undefined): boolean => n?.type === schema.nodes.page_break;

/** Paragraph attributes a page break keeps for a w:p that has no text piece. */
const paraJsonCache = new WeakMap<PMNode, Record<string, any> | null>();
function breakPara(node: PMNode): Record<string, any> | null {
  const known = paraJsonCache.get(node);
  if (known !== undefined) return known;
  const hit: Record<string, any> | null = node.attrs.para ? JSON.parse(node.attrs.para) : null;
  paraJsonCache.set(node, hit);
  return hit;
}

/** A revision inside one paragraph piece, positions relative to its content. */
interface LocalRevision {
  kind: 'ins' | 'del' | 'format';
  move: boolean;
  from: number;
  to: number;
  layerId?: string | null;
  rPr?: string;
  info: RevisionInfo;
}

// Per node, so paragraphs the user didn't touch are not read again on every keystroke.
const localCache = new WeakMap<PMNode, LocalRevision[]>();

function localRevisions(node: PMNode): LocalRevision[] {
  let out = localCache.get(node);
  if (out) return out;
  out = [];
  if (isBreak(node)) {
    for (const l of layersOf(node)) {
      const hit = insertionOfLayer(l.open);
      if (hit) out.push({ kind: 'ins', move: hit.kind === 'moveTo', from: 0, to: 1, layerId: l.id, info: hit.info });
    }
  } else {
    const list = out;
    const ins = new Map<string, LocalRevision>();
    let fmt: LocalRevision | null = null;
    node.forEach((child, p) => {
      const e = p + child.nodeSize;
      for (const l of layersOf(child)) {
        const hit = insertionOfLayer(l.open);
        if (!hit) continue;
        const open = ins.get(l.id);
        if (open && open.to === p) open.to = e;
        else {
          const r: LocalRevision = { kind: 'ins', move: hit.kind === 'moveTo', from: p, to: e, layerId: l.id, info: hit.info };
          ins.set(l.id, r);
          list.push(r);
        }
      }
      if (isDeletion(child)) {
        const d = deletedInfo(child.attrs.xml);
        if (d) list.push({ kind: 'del', move: d.kind === 'moveFrom', from: p, to: e, info: { author: d.author, date: d.date, dateUtc: d.dateUtc } });
      } else if (child.type === schema.nodes.raw_inline && child.attrs.hidden && child.attrs.label === 'ins') {
        // An empty w:ins, kept whole.
        const hit = insertionOfLayer(child.attrs.xml);
        if (hit) list.push({ kind: 'ins', move: false, from: p, to: e, layerId: null, info: hit.info });
      }
      const run = child.marks.find((m) => m.type === schema.marks.run);
      const change = run ? runFormatChange(run.attrs.rPr) : null;
      if (change) {
        if (fmt && fmt.to === p && fmt.rPr === run!.attrs.rPr) fmt.to = e;
        else {
          fmt = { kind: 'format', move: false, from: p, to: e, rPr: run!.attrs.rPr, info: change };
          list.push(fmt);
        }
      } else fmt = null;
    });
  }
  localCache.set(node, out);
  return out;
}

/** The nodes written as one w:p (see writer.ts paragraphGroupSize). */
interface Group {
  parent: PMNode;
  /** Index of its first node in the parent, and how many nodes. */
  index: number;
  count: number;
  start: number;
  end: number;
}

/** Every paragraph (group) in the document, in order; with a range, those in it. */
function eachGroup(node: PMNode, base: number, fn: (g: Group) => void, lo = -1, hi = Infinity): void {
  let i = 0;
  let pos = base;
  if (lo > base) {
    // Straight to the block holding `lo`, then back to the start of its w:p.
    while (i < node.childCount - 1 && pos + node.child(i).nodeSize <= lo) pos += node.child(i++).nodeSize;
    while (i > 0 && groupId(node.child(i)) && groupId(node.child(i - 1)) === groupId(node.child(i))) pos -= node.child(--i).nodeSize;
  }
  while (i < node.childCount && pos < hi) {
    const k = node.child(i);
    if (isPara(k) || isBreak(k)) {
      let n = 1;
      const id = groupId(k);
      if (id) {
        const run: PMNode[] = [];
        for (let j = i; j < node.childCount && groupId(node.child(j)) === id; j++) run.push(node.child(j));
        n = paragraphGroupSize(run, 0);
      }
      let end = pos;
      for (let j = 0; j < n; j++) end += node.child(i + j).nodeSize;
      if (end > lo) fn({ parent: node, index: i, count: n, start: pos, end });
      i += n;
      pos = end;
    } else {
      if (!k.isLeaf && !k.isTextblock && pos + k.nodeSize > lo && pos < hi) eachGroup(k, pos + 1, fn);
      pos += k.nodeSize;
      i++;
    }
  }
}

function firstParagraphIn(node: PMNode): boolean {
  let n: PMNode | null = node;
  while (n && !isPara(n)) n = n.isLeaf ? null : n.firstChild;
  return !!n;
}

function groupRevisions(g: Group, out: Revision[]): void {
  const where = { group: g.start, groupEnd: g.end };
  const first = g.parent.child(g.index);
  let para: PMNode | null = null;
  for (let j = 0; j < g.count && !para; j++) if (isPara(g.parent.child(g.index + j))) para = g.parent.child(g.index + j);
  const pr = paragraphRevisions(para ? para.attrs.pPr : breakPara(first)?.pPr ?? null);
  if (pr.format) {
    out.push({ key: `paraFormat:${g.start}`, kind: 'paraFormat', move: false, from: g.start + 1, to: g.end - 1, para: g.start, ...where, ...pr.format });
  }
  // Inline revisions; an insertion around a page break continues into the next piece.
  let carry: Revision[] = [];
  let lastEnd = g.start;
  let pos = g.start;
  for (let j = 0; j < g.count; j++) {
    const node = g.parent.child(g.index + j);
    const start = isPara(node) ? pos + 1 : pos;
    const end = isPara(node) ? pos + node.nodeSize - 1 : pos + 1;
    const next: Revision[] = [];
    for (const l of localRevisions(node)) {
      const from = start + l.from;
      const to = start + l.to;
      const cont = l.kind === 'ins' && l.layerId && l.from === 0 ? carry.find((r) => r.layerId === l.layerId) : undefined;
      const r: Revision = cont ?? { key: `${l.kind}:${from}`, kind: l.kind, move: l.move, from, to, layerId: l.layerId, rPr: l.rPr, ...where, ...l.info };
      if (cont) cont.to = to;
      else out.push(r);
      if (r.kind === 'ins' && r.layerId && to === end) next.push(r);
    }
    carry = next;
    if (isPara(node)) lastEnd = end;
    pos += node.nodeSize;
  }
  if (pr.mark) {
    const after = g.parent.maybeChild(g.index + g.count);
    const joinable = !!after && (isPara(after) || isBreak(after) || (after.type === schema.nodes.table && g.count === 1 && !!para && firstParagraphIn(after)));
    const at = para ? lastEnd : g.start;
    out.push({ key: `paraMark:${at}`, kind: 'paraMark', move: false, mark: pr.mark.kind, from: at, to: at, para: g.start, joinable, ...where, ...pr.mark.info });
  }
}

const revisionCache = new WeakMap<PMNode, Revision[]>();

/** Every tracked change in a document, in document order. */
/**
 * Whether a node may hold a revision: a paragraph or page break whose own content or paragraph
 * properties carry one, or a container (table, cell ...) with such a node. It never answers "no"
 * for a node that has one; per node, so after an edit only the changed blocks are looked at again.
 */
const mayHoldCache = new WeakMap<PMNode, boolean>();
function mayHoldRevisions(node: PMNode): boolean {
  let hit = mayHoldCache.get(node);
  if (hit !== undefined) return hit;
  if (isPara(node) || isBreak(node)) {
    const pr = paragraphRevisions(isPara(node) ? node.attrs.pPr : breakPara(node)?.pPr ?? null);
    hit = !!pr.format || !!pr.mark || localRevisions(node).length > 0;
  } else {
    hit = false;
    if (!node.isLeaf && !node.isTextblock) node.forEach((child) => { if (!hit && mayHoldRevisions(child)) hit = true; });
  }
  mayHoldCache.set(node, hit);
  return hit;
}

export function collectRevisions(doc: PMNode): Revision[] {
  const cached = revisionCache.get(doc);
  if (cached) return cached;
  // Most documents have none: no walk over every paragraph after each keystroke.
  if (!mayHoldRevisions(doc)) {
    revisionCache.set(doc, []);
    return [];
  }
  const list: Revision[] = [];
  eachGroup(doc, 0, (g) => groupRevisions(g, list));
  // Groups come in order; an insertion on a page break after the last text piece may not.
  if (list.some((r, i) => i > 0 && r.from < list[i - 1].from)) list.sort((a, b) => a.from - b.from);
  revisionCache.set(doc, list);
  return list;
}

// ----- plugin state: markup on/off, the revision reached by next/previous, comments -----

interface Focus {
  key: string;
  doc: PMNode;
  selection: Selection;
}

interface ReviewState {
  show: boolean;
  focus: Focus | null;
  comments: DocComment[];
  decorations: DecorationSet;
}

interface ReviewMeta {
  show?: boolean;
  focus?: string | null;
  comments?: DocComment[];
}

export const reviewKey = new PluginKey<ReviewState>('dx-review');

export function review(): Plugin<ReviewState> {
  return new Plugin<ReviewState>({
    key: reviewKey,
    state: {
      init: (_, state) => ({ show: true, focus: null, comments: [], decorations: decorate(state.doc, []) }),
      apply(tr, prev) {
        const meta = tr.getMeta(reviewKey) as ReviewMeta | undefined;
        let next = prev;
        if (meta) {
          next = { ...prev };
          if (meta.show != null) next.show = meta.show;
          if (meta.focus !== undefined) next.focus = meta.focus ? { key: meta.focus, doc: tr.doc, selection: tr.selection } : null;
          if (meta.comments) next.comments = meta.comments;
        }
        // The comments on the document (editable, undoable) win over those given with setComments.
        const list = (tr.doc.attrs.comments as DocComment[] | null) ?? next.comments;
        if (meta?.comments) next = { ...next, decorations: decorate(tr.doc, list) };
        else if (tr.docChanged) {
          let decorations = updateDecorations(prev.decorations, tr, list);
          if (tr.before.attrs.comments !== tr.doc.attrs.comments) decorations = refreshCommentDecorations(decorations, tr.doc, list);
          next = { ...next, decorations };
        }
        if (tr.docChanged && !meta?.focus && next.focus) next = { ...next, focus: null };
        return next;
      },
    },
    // Text typed or pasted as plain text must not join a tracked insertion: the marks it
    // would inherit (stored marks) lose their insertion layers. With tracking on they get the
    // insertion the text goes into instead (trackChanges.ts trackingMarks).
    appendTransaction(trs, _old, state) {
      if (!trs.some((t) => t.docChanged || t.selectionSet || t.storedMarksSet)) return null;
      const { selection } = state;
      let marks: readonly Mark[] | null = state.storedMarks;
      if (!marks) {
        const { $from, $to } = selection;
        if (selection.empty) marks = $from.marks();
        else if ($from.sameParent($to) && $from.parent.inlineContent) marks = $from.marksAcross($to);
      }
      if (!marks) return null;
      const clean = trackingMarks(state, withoutInsertionLayers(marks));
      return clean === marks || sameMarks(clean, marks) ? null : state.tr.setStoredMarks(clean);
    },
    props: {
      decorations: (state) => reviewKey.getState(state)?.decorations,
      attributes: (state): Record<string, string> => (reviewKey.getState(state)?.show === false ? { class: 'dx-rev-off' } : {}),
    },
  });
}

function sameMarks(a: readonly Mark[], b: readonly Mark[]): boolean {
  return a.length === b.length && a.every((m, i) => m.eq(b[i]));
}

function withoutInsertionLayers(marks: readonly Mark[]): readonly Mark[] {
  const wrap = marks.find((m) => m.type === schema.marks.inlineWrap);
  if (!wrap) return marks;
  const layers = parsedLayers(wrap.attrs.layers);
  const keep = layers.filter((l) => !insertionOfLayer(l.open));
  return keep.length === layers.length ? marks : withLayers(marks, keep);
}

const REV = { rev: true };
const COMMENT = { comment: true };

function decorate(doc: PMNode, comments: DocComment[]): DecorationSet {
  const decos: Decoration[] = [];
  revisionDecorations(doc, collectRevisions(doc), decos);
  commentDecorations(doc, comments, decos);
  return DecorationSet.create(doc, decos);
}

/** Paragraph pieces (and page breaks) of a group, with positions. */
function groupNodes(doc: PMNode, from: number, to: number, fn: (node: PMNode, pos: number) => void): void {
  doc.nodesBetween(from, to, (node, pos) => {
    if (pos >= from && pos + node.nodeSize <= to) {
      if (isPara(node) || isBreak(node)) fn(node, pos);
      return false;
    }
    return true;
  });
}

// Change bars beside paragraphs with a revision, like Word's: drawn by CSS for paragraphs
// holding a marked revision (editor.css, :has), so a long document doesn't need thousands of
// node decorations (which the view walks on every keystroke); paragraph formatting changes
// carry the bar on their own decoration.
function revisionDecorations(doc: PMNode, revs: Revision[], decos: Decoration[]): void {
  for (const r of revs) {
    const style = `--dx-rev:${authorColor(r.author)}`;
    if (r.kind === 'format') {
      decos.push(Decoration.inline(r.from, r.to, { class: 'dx-rev-fmt', title: describeRevision('格式變更', r), style }, REV));
    } else if (r.kind === 'paraFormat') {
      const title = describeRevision('段落格式變更', r);
      groupNodes(doc, r.group, r.groupEnd, (n, pos) => {
        if (isPara(n)) decos.push(Decoration.node(pos, pos + n.nodeSize, { class: 'dx-rev-pfmt dx-rev-bar', title }, REV));
      });
    } else if (r.kind === 'paraMark') {
      const title = describeRevision(r.mark === 'ins' ? '插入段落標記' : '刪除段落標記', r);
      decos.push(Decoration.widget(r.to, () => markWidget(r.mark!, title, style), { side: 1, key: `pm:${r.mark}:${r.author}`, ignoreSelection: true, rev: true }));
    } else if (r.kind === 'ins' && (isBreak(doc.nodeAt(r.from)) || doc.resolve(r.from).parent !== doc.resolve(r.to).parent)) {
      // A page break inside an insertion is shown as inserted too.
      const title = describeRevision(r.move ? '移入' : '插入', r);
      groupNodes(doc, r.from, r.to, (n, pos) => {
        if (isBreak(n) && layersOf(n).some((l) => l.id === r.layerId)) {
          decos.push(Decoration.node(pos, pos + 1, { class: 'dx-rev-ins dx-rev-break dx-rev-bar', title, style }, REV));
        }
      });
    }
  }
}

function commentDecorations(doc: PMNode, comments: DocComment[], decos: Decoration[]): void {
  const byId = new Map(comments.map((c) => [c.id, c]));
  for (const [id, range] of commentRanges(doc)) {
    const c = byId.get(id);
    const title = c ? `留言（${c.author || '未知作者'}）：${c.text}` : '留言';
    if (range.to > range.from) {
      // Text typed at either end is inside the markers, so inside the range.
      decos.push(Decoration.inline(range.from, range.to, { class: 'dx-comment', 'data-comment-id': id, title }, { ...COMMENT, inclusiveStart: true, inclusiveEnd: true }));
    }
    if (range.ref != null) decos.push(Decoration.node(range.ref, range.ref + 1, { class: 'dx-comment-ref', 'data-comment-id': id, title }, COMMENT));
  }
}

/** Comment decorations drawn again (a comment was added, edited, resolved or deleted). */
function refreshCommentDecorations(set: DecorationSet, doc: PMNode, comments: DocComment[]): DecorationSet {
  const add: Decoration[] = [];
  commentDecorations(doc, comments, add);
  return set.remove(set.find(undefined, undefined, (spec) => spec.comment)).add(doc, add);
}

/** Where a transaction changed the document, in its final positions (marks and attributes included). */
function changedRange(tr: Transaction): { from: number; to: number } | null {
  let from = Infinity;
  let to = -Infinity;
  tr.steps.forEach((step, i) => {
    const map = tr.mapping.maps[i];
    if (from <= to) {
      from = map.map(from, -1);
      to = map.map(to, 1);
    }
    const s = step as unknown as { from?: number; to?: number; pos?: number };
    let a: number;
    let b: number;
    if (typeof s.from === 'number' && typeof s.to === 'number') [a, b] = [s.from, s.to];
    else if (typeof s.pos === 'number') [a, b] = [s.pos, s.pos + 1];
    else return;
    from = Math.min(from, map.map(a, -1));
    to = Math.max(to, map.map(b, 1));
  });
  return from <= to ? { from, to } : null;
}

const COMMENT_LABELS = new Set(['commentRangeStart', 'commentRangeEnd', '註解']);
const isCommentMarker = (n: PMNode) =>
  (n.type === schema.nodes.raw_inline || n.type === schema.nodes.raw_block) && COMMENT_LABELS.has(n.attrs.label);

/** Whether a transaction added or removed comment markers. */
function touchesComments(tr: Transaction): boolean {
  return tr.steps.some((step, i) => {
    if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) return false;
    let hit = false;
    const check = (n: PMNode) => {
      if (isCommentMarker(n)) hit = true;
      return !hit;
    };
    const s = step as unknown as { from: number; to: number; slice: { content: { descendants(f: (n: PMNode) => boolean): void } } };
    tr.docs[i].nodesBetween(s.from, s.to, check);
    if (!hit) s.slice.content.descendants(check);
    return hit;
  });
}

/** Map the decorations and recompute those of the paragraphs a transaction changed. */
function updateDecorations(set: DecorationSet, tr: Transaction, comments: DocComment[]): DecorationSet {
  const doc = tr.doc;
  const range = changedRange(tr);
  if (range && range.to - range.from > doc.content.size / 2) return decorate(doc, comments);
  let deco = set.map(tr.mapping, doc);
  if (!range) return deco;
  // Whole top-level blocks, and the whole w:p when it was split at page breaks.
  let a = Math.max(0, Math.min(range.from, doc.content.size));
  let b = Math.max(a, Math.min(range.to, doc.content.size));
  const $a = doc.resolve(a);
  if ($a.depth) a = $a.before(1);
  const $b = doc.resolve(b);
  if ($b.depth) b = $b.after(1);
  const sameGroup = (x: PMNode | null | undefined, y: PMNode | null | undefined) => !!x && !!y && !!groupId(x) && groupId(x) === groupId(y);
  for (let $x = doc.resolve(a); sameGroup($x.nodeBefore, $x.nodeAfter); $x = doc.resolve(a)) a -= $x.nodeBefore!.nodeSize;
  for (let $x = doc.resolve(b); sameGroup($x.nodeBefore, $x.nodeAfter); $x = doc.resolve(b)) b += $x.nodeAfter!.nodeSize;
  const stale = deco.find(a, b, (spec) => spec.rev).filter((d) => d.from >= a && d.to <= b);
  deco = deco.remove(stale);
  const revs: Revision[] = [];
  eachGroup(doc, 0, (g) => groupRevisions(g, revs), a, b);
  const add: Decoration[] = [];
  revisionDecorations(doc, revs, add);
  if (touchesComments(tr)) {
    deco = deco.remove(deco.find(undefined, undefined, (spec) => spec.comment));
    commentDecorations(doc, comments, add);
  }
  return deco.add(doc, add);
}

function markWidget(kind: 'ins' | 'del', title: string, style: string): HTMLElement {
  const el = document.createElement('span');
  el.className = `dx-rev-pmark dx-rev-pmark-${kind}`;
  el.textContent = '¶';
  el.title = title;
  el.setAttribute('style', style);
  el.contentEditable = 'false';
  return el;
}

// ----- which revisions a command acts on -----

function focused(state: EditorState): Revision | null {
  const f = reviewKey.getState(state)?.focus;
  if (!f || f.doc !== state.doc || !f.selection.eq(state.selection)) return null;
  return collectRevisions(state.doc).find((r) => r.key === f.key) ?? null;
}

const isInline = (r: Revision) => r.kind === 'ins' || r.kind === 'del' || r.kind === 'format';

/**
 * The revisions at the selection: the one reached with next/previous, those the selection
 * overlaps, or at a cursor the one it is in (else next to it), else its paragraph's own.
 */
export function revisionsAtSelection(state: EditorState): Revision[] {
  const f = focused(state);
  if (f) return [f];
  const revs = collectRevisions(state.doc);
  const { from, to, empty, $from } = state.selection;
  if (!empty) {
    return revs.filter((r) => {
      if (r.kind === 'paraMark') return from <= r.to && to > r.to; // the selection runs past the mark
      if (r.kind === 'paraFormat') return from <= r.to && to >= r.from;
      return r.from < to && r.to > from;
    });
  }
  const inline = revs.filter(isInline);
  const inside = inline.filter((r) => r.from < from && r.to > from);
  if (inside.length) return inside;
  const after = inline.find((r) => r.from === from);
  if (after) return [after];
  const before = inline.filter((r) => r.to === from).pop();
  if (before) return [before];
  if ($from.parent.type !== schema.nodes.paragraph) return [];
  const para = $from.before();
  return revs.filter(
    (r) => r.group <= para && para < r.groupEnd && (r.kind === 'paraFormat' || (r.kind === 'paraMark' && from === r.to)),
  );
}

// ----- accept / reject -----

const MODELED = new Set(['charStyle', 'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'color', 'highlight', 'font', 'fontSize']);
/** Markers that are not content: rejecting an insertion keeps them (outside the insertion). */
const KEPT_ON_REJECT = new Set([
  'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', '註解',
  'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
]);
const MOVE_RANGES = new Set(['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);

function withLayers(marks: readonly Mark[], layers: Layer[]): readonly Mark[] {
  const rest = schema.marks.inlineWrap.removeFromSet(marks);
  return layers.length ? schema.marks.inlineWrap.create({ layers: JSON.stringify(layers) }).addToSet(rest) : rest;
}

function setMarks(tr: Transform, from: number, to: number, marks: readonly Mark[]): void {
  tr.removeMark(from, to);
  for (const m of marks) tr.addMark(from, to, m);
}

/** Give a node (inline, or a page break) new wrapper layers. */
function setLayers(tr: Transform, pos: number, node: PMNode, layers: Layer[]): void {
  if (isBreak(node)) {
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, layers: layers.length ? JSON.stringify(layers) : null });
    return;
  }
  const end = pos + node.nodeSize;
  tr.removeMark(pos, end, schema.marks.inlineWrap);
  if (layers.length) tr.addMark(pos, end, schema.marks.inlineWrap.create({ layers: JSON.stringify(layers) }));
}

/** Inline nodes and page breaks in [from, to) of the current document. */
function contentNodes(tr: Transform, from: number, to: number): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  if (to <= from) return out;
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isInline || isBreak(node)) {
      if (pos >= from && pos + node.nodeSize <= to) out.push({ node, pos });
      return false;
    }
    return true;
  });
  return out;
}

/** Remove a page break from its w:p: the text pieces around it become one paragraph again. */
function removePageBreak(tr: Transform, pos: number): void {
  const node = tr.doc.nodeAt(pos)!;
  const id = node.attrs.group as string | null;
  const same = (n: PMNode | null | undefined) => !!n && !!id && groupId(n) === id;
  const before = tr.doc.resolve(pos).nodeBefore;
  const after = tr.doc.resolve(pos + node.nodeSize).nodeAfter;
  tr.delete(pos, pos + node.nodeSize);
  if (same(before) && same(after)) {
    if (isPara(before) && isPara(after) && canJoin(tr.doc, pos)) tr.join(pos);
  } else if (!same(before) && !same(after)) {
    // The w:p held only this break: its paragraph mark stays, as an empty paragraph.
    tr.insert(pos, schema.nodes.paragraph.create({ ...(breakPara(node) ?? {}), brGroup: null }));
  }
}

function resolveInsertion(tr: Transform, r: Revision, accept: boolean, m: Mapping): void {
  if (r.layerId === null) {
    const at = m.mapResult(r.from, 1);
    const node = at.deleted ? null : tr.doc.nodeAt(at.pos);
    if (node?.type === schema.nodes.raw_inline && node.attrs.label === 'ins') tr.delete(at.pos, at.pos + 1);
    return;
  }
  const from = m.map(r.from, 1);
  const to = m.map(r.to, -1);
  const hits = contentNodes(tr, from, to).filter(({ node }) => layersOf(node).some((l) => l.id === r.layerId));
  const drop = (node: PMNode) => layersOf(node).filter((l) => l.id !== r.layerId);
  if (accept) {
    for (const { node, pos } of hits) setLayers(tr, pos, node, drop(node));
    return;
  }
  // Reject: remove the content, last first so earlier positions stay valid.
  for (const { node, pos } of hits.reverse()) {
    if (isBreak(node)) removePageBreak(tr, pos);
    else if (node.type === schema.nodes.raw_inline && KEPT_ON_REJECT.has(node.attrs.label)) setLayers(tr, pos, node, drop(node));
    else tr.delete(pos, pos + node.nodeSize);
  }
}

/** A converter for the content of a kept deletion, with the pictures and links it refers to. */
function converterFor(node: PMNode): Converter {
  const refs = node.attrs.refs as { rels: Record<string, { target: string; external: boolean }>; images: Record<string, string> } | null;
  return new Converter(new Map(Object.entries(refs?.rels ?? {})), new Map(Object.entries(refs?.images ?? {})));
}

/** A deletion's content parsed beforehand, once (converting it may change the elements). */
function take(restored: Map<string, Element> | undefined, xml: string): Element | undefined {
  const el = restored?.get(xml);
  if (el) restored!.delete(xml);
  return el;
}

function resolveDeletion(tr: Transform, r: Revision, accept: boolean, m: Mapping, restored?: Map<string, Element>): void {
  // Skip it when it went with something resolved before (an insertion holding it).
  const at = m.mapResult(r.from, 1);
  if (at.deleted) return;
  const pos = at.pos;
  const node = tr.doc.nodeAt(pos);
  if (!node || !isDeletion(node)) return;
  const kept = accept ? keptMarkers(node.attrs.xml) : take(restored, node.attrs.xml) ?? restoredContent(node.attrs.xml);
  const content = kept ? converterFor(node).inlineContent(kept, node.marks) : [];
  if (content.length) tr.replaceWith(pos, pos + 1, content);
  else tr.delete(pos, pos + 1);
}

function resolveFormat(tr: Transform, r: Revision, accept: boolean, m: Mapping): void {
  const from = m.map(r.from, 1);
  const to = m.map(r.to, -1);
  for (const { node, pos } of contentNodes(tr, from, to)) {
    const run = node.marks.find((mk) => mk.type === schema.marks.run);
    if (!run || run.attrs.rPr !== r.rPr) continue;
    const rPr = accept ? acceptRunFormat(r.rPr!) : rejectRunFormat(r.rPr!);
    let marks = schema.marks.run.removeFromSet(node.marks);
    if (!accept) {
      // The recorded properties replace the formatting shown now.
      marks = marks.filter((mk) => !MODELED.has(mk.type.name));
      for (const mk of runMarks(rPr)) marks = mk.addToSet(marks);
    }
    if (rPr || run.attrs.attrs) marks = schema.marks.run.create({ rPr, attrs: run.attrs.attrs }).addToSet(marks);
    setMarks(tr, pos, pos + node.nodeSize, marks);
  }
}

const runMarkCache = new Map<string, Mark[]>();
function runMarks(rPr: string | null): Mark[] {
  if (!rPr) return [];
  return runMarkCache.get(rPr) ?? remember(runMarkCache, rPr, modelMarks(readRunModel(parseFragment(rPr))));
}

/** New paragraph properties for every piece of the w:p in [start, end). */
function setGroupPPr(tr: Transform, start: number, end: number, change: (pPr: string) => string | null): void {
  for (let pos = start; pos < end; ) {
    const node = tr.doc.nodeAt(pos);
    if (!node) break;
    if (isPara(node) && node.attrs.pPr) {
      const pPr = change(node.attrs.pPr);
      // Re-read what the editor models (alignment, indents, list ...) from the new properties.
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...readParaModel(pPrElement(pPr)), pPr }, node.marks);
    } else if (isBreak(node) && breakPara(node)?.pPr) {
      const a = breakPara(node)!;
      const pPr = change(a.pPr);
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, para: JSON.stringify({ ...a, ...readParaModel(pPrElement(pPr)), pPr }) });
    }
    pos += node.nodeSize;
  }
}

/** The nodes of the w:p starting at `pos` (a paragraph and the pieces split from it at page breaks). */
function groupAt(doc: PMNode, pos: number): { node: PMNode; pos: number }[] {
  const $p = doc.resolve(pos);
  const kids: PMNode[] = [];
  $p.parent.forEach((k) => kids.push(k));
  const i = $p.index();
  const n = paragraphGroupSize(kids, i);
  const out: { node: PMNode; pos: number }[] = [];
  for (let j = 0, p = pos; j < n; j++) {
    out.push({ node: kids[i + j], pos: p });
    p += kids[i + j].nodeSize;
  }
  return out;
}

/**
 * Join the w:p in [start, end) with the next one: one w:p holding both, with the next one's
 * properties (and section break). Pieces split at page breaks share one group id.
 */
function joinGroups(tr: Transaction, start: number, end: number): void {
  const g: { node: PMNode; pos: number }[] = [];
  for (let pos = start; pos < end; ) {
    const node = tr.doc.nodeAt(pos)!;
    g.push({ node, pos });
    pos += node.nodeSize;
  }
  const h = groupAt(tr.doc, end);
  const hPara = h.find(({ node }) => isPara(node))?.node;
  const props: Record<string, any> = hPara ? hPara.attrs : breakPara(h[0].node) ?? {};
  const lastG = g[g.length - 1];
  const physical = isPara(lastG.node) && isPara(h[0].node);
  const count = g.length + h.length - (physical ? 1 : 0);
  const id = count > 1 ? groupId(g[0].node) ?? groupId(h[0].node) ?? newLayerId() : null;
  const hasPara = !!hPara || g.some(({ node }) => isPara(node));
  const para = hasPara ? null : JSON.stringify(props);
  // Properties first (positions don't move), then the join.
  for (const { node, pos } of g) {
    if (isPara(node)) tr.setNodeMarkup(pos, undefined, { ...props, brGroup: id, sectPr: null }, node.marks);
    else tr.setNodeMarkup(pos, undefined, { ...node.attrs, group: id, para });
  }
  for (const { node, pos } of h) {
    if (isPara(node)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, brGroup: id }, node.marks);
    else tr.setNodeMarkup(pos, undefined, { ...node.attrs, group: id, para });
  }
  if (physical && canJoin(tr.doc, end)) {
    tr.join(end);
    tr.setNodeMarkup(lastG.pos, undefined, { ...h[0].node.attrs, brGroup: id }, h[0].node.marks);
  }
}

/** Word joins a paragraph followed by a table with the table's first paragraph (which keeps its properties). */
function joinIntoTable(tr: Transaction, start: number, end: number): boolean {
  const para = tr.doc.nodeAt(start)!;
  let pos = end;
  let node = tr.doc.nodeAt(pos);
  while (node && !isPara(node)) {
    if (node.isLeaf || !node.firstChild) return false;
    pos += 1;
    node = node.firstChild;
  }
  if (!node) return false;
  if (para.content.size) tr.insert(pos + 1, para.content);
  tr.delete(start, end);
  return true;
}

function resolveMark(tr: Transaction, r: Revision, accept: boolean, m: Mapping): void {
  const start = r.group;
  const end = m.map(r.groupEnd, -1);
  const join = (r.mark === 'del') === accept;
  if (!join) {
    setGroupPPr(tr, start, end, withoutMarkRevision);
    return;
  }
  const $end = tr.doc.resolve(end);
  const next = $end.nodeAfter;
  if (isPara(next) || isBreak(next)) return joinGroups(tr, start, end);
  const node = tr.doc.nodeAt(start);
  const single = !!node && isPara(node) && start + node.nodeSize === end;
  if (single && next?.type === schema.nodes.table && joinIntoTable(tr, start, end)) return;
  // Nothing to join with: an empty paragraph goes (never the last block, nor one a table would end), else only the record.
  const parent = $end.parent;
  const last = $end.index() === parent.childCount;
  const prev = tr.doc.resolve(start).nodeBefore;
  if (single && node!.content.size === 0 && parent.childCount > 1 && !(last && !isPara(prev) && parent.type === schema.nodes.doc)) {
    tr.delete(start, end);
    return;
  }
  setGroupPPr(tr, start, end, withoutMarkRevision);
}

/**
 * Resolve the revisions inside one paragraph (all but its mark) on a copy of just that
 * paragraph, then put back what changed with a single step: the work stays proportional to
 * the paragraph, however long the document (Accept All on thousands of paragraphs).
 */
function resolveInParagraph(tr: Transaction, start: number, end: number, revs: Revision[], accept: boolean): void {
  const before = tr.doc.slice(start, end).content;
  const after = resolvedParagraph(before, start, revs, accept);
  if (!after) return;
  const a = before.findDiffStart(after);
  if (a == null) return;
  let { a: endA, b: endB } = before.findDiffEnd(after)!;
  const overlap = a - Math.min(endA, endB);
  if (overlap > 0) {
    endA += overlap;
    endB += overlap;
  }
  // Only the part that changed is replaced (so the selection elsewhere stays put); the whole
  // paragraph when that part can't be put back as it is.
  const piece = schema.nodes.doc.create(null, after).slice(a, endB);
  const step = replaceStep(tr.doc, start + a, start + endA, piece);
  const res = step?.apply(tr.doc);
  if (step && res?.doc && res.doc.slice(start, start + after.size).content.eq(after)) tr.step(step);
  else tr.replaceWith(start, end, after);
}

/**
 * The nodes of the w:p `before` (starting at `start`) with its inner revisions resolved; null if
 * unchanged. `restored`: rejected deletions' content parsed beforehand (see restoredContents).
 */
function resolvedParagraph(before: Fragment, start: number, revs: Revision[], accept: boolean, restored?: Map<string, Element>): Fragment | null {
  const mini = new Transform(schema.nodes.doc.create(null, before));
  const local = (r: Revision): Revision => ({ ...r, from: r.from - start, to: r.to - start, group: 0, groupEnd: r.groupEnd - start });
  for (const r of revs) {
    // Positions map through what was done so far (an insertion rejected before may hold this one).
    if (r.kind === 'ins') resolveInsertion(mini, local(r), accept, mini.mapping);
    else if (r.kind === 'del') resolveDeletion(mini, local(r), accept, mini.mapping, restored);
    else if (r.kind === 'format') resolveFormat(mini, local(r), accept, mini.mapping);
  }
  if (revs.some((r) => r.kind === 'paraFormat')) {
    setGroupPPr(mini, 0, mini.doc.content.size, accept ? acceptParagraphFormat : rejectParagraphFormat);
  }
  return mini.docChanged ? mini.doc.content : null;
}

/** The XML of the kept deletions among `revs` (positions in `doc`). */
function* deletedXml(doc: PMNode, revs: Revision[]): Iterable<string> {
  for (const r of revs) {
    if (r.kind !== 'del') continue;
    const node = doc.nodeAt(r.from);
    if (node && isDeletion(node) && typeof node.attrs.xml === 'string') yield node.attrs.xml;
  }
}

/**
 * Accept / Reject All: every paragraph's inner revisions resolved on its own, and the
 * document rebuilt once (one step, not one per paragraph: the work stays linear in the size
 * of the document). Then the paragraph marks, last first. Returns whether paragraphs changed.
 */
function resolveEverything(tr: Transaction, revs: Revision[], accept: boolean): boolean {
  const groups = new Map<number, { end: number; inner: Revision[]; marks: Revision[] }>();
  for (const r of revs) {
    let g = groups.get(r.group);
    if (!g) groups.set(r.group, (g = { end: r.groupEnd, inner: [], marks: [] }));
    (r.kind === 'paraMark' ? g.marks : g.inner).push(r);
  }
  const doc = tr.doc;
  // Rejected deletions come back as runs: their XML is parsed all together first (thousands of
  // separate parses were most of Reject All's time).
  const restored = accept ? undefined : restoredContents(deletedXml(doc, revs));
  const replace = new Map<number, { end: number; content: Fragment }>();
  let paragraphs = false;
  for (const [start, g] of groups) {
    if (g.inner.some((r) => r.kind === 'paraFormat')) paragraphs = true;
    if (!g.inner.length) continue;
    const content = resolvedParagraph(doc.slice(start, g.end).content, start, g.inner, accept, restored);
    if (content) replace.set(start, { end: g.end, content });
  }
  // Where each paragraph with a mark revision is once the others have changed.
  const track = new Map<number, number>();
  for (const [start, g] of groups) if (g.marks.length) track.set(start, g.end);
  const moved = new Map<number, [number, number]>();
  if (replace.size) {
    const keys = [...new Set([...replace.keys(), ...track.keys()])].sort((a, b) => a - b);
    const next = rebuild(doc, 0, 0, replace, track, moved, keys);
    const a = doc.content.findDiffStart(next.content);
    if (a != null) {
      let { a: endA, b: endB } = doc.content.findDiffEnd(next.content)!;
      const overlap = a - Math.min(endA, endB);
      if (overlap > 0) {
        endA += overlap;
        endB += overlap;
      }
      const { from } = tr.selection;
      tr.replace(a, endA, next.slice(a, endB));
      if (!tr.doc.eq(next)) tr.replaceWith(0, tr.doc.content.size, next.content);
      // Keep the cursor about where it was, not at the end of the rebuilt part.
      if (from >= a && from <= endA) tr.setSelection(Selection.near(tr.doc.resolve(Math.min(from, tr.doc.content.size))));
    }
  } else for (const [start, end] of track) moved.set(start, [start, end]);
  // Paragraph marks, last first: joining takes the (already resolved) next paragraph's properties.
  const marks = [...groups].filter(([, g]) => g.marks.length).sort((x, y) => y[0] - x[0]);
  const none = new Mapping();
  for (const [start, g] of marks) {
    paragraphs = true;
    const [s, e] = moved.get(start) ?? [start, g.end];
    for (const r of g.marks) resolveMark(tr, { ...r, group: s, groupEnd: e }, accept, none);
  }
  return paragraphs;
}

/**
 * A copy of `node` (its content starting at `base`, at `nbase` in the copy) with the given
 * paragraphs replaced. Records where tracked paragraphs start and end in the copy.
 */
function rebuild(
  node: PMNode,
  base: number,
  nbase: number,
  replace: Map<number, { end: number; content: Fragment }>,
  track: Map<number, number>,
  moved: Map<number, [number, number]>,
  keys: number[],
): PMNode {
  // Whether any replaced or tracked paragraph starts in [from, to).
  const any = (from: number, to: number) => {
    let lo = 0;
    let hi = keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (keys[mid] < from) lo = mid + 1;
      else hi = mid;
    }
    return lo < keys.length && keys[lo] < to;
  };
  const out: PMNode[] = [];
  let pos = base;
  let npos = nbase;
  let changed = false;
  for (let i = 0; i < node.childCount; ) {
    const rep = replace.get(pos);
    if (rep) {
      if (track.has(pos)) moved.set(pos, [npos, npos + rep.content.size]);
      rep.content.forEach((n) => out.push(n));
      npos += rep.content.size;
      while (pos < rep.end && i < node.childCount) pos += node.child(i++).nodeSize;
      changed = true;
      continue;
    }
    const child = node.child(i);
    const end = track.get(pos);
    if (end != null) moved.set(pos, [npos, npos + (end - pos)]);
    let copy = child;
    if (!child.isLeaf && !child.isTextblock && any(pos + 1, pos + child.nodeSize)) {
      copy = rebuild(child, pos + 1, npos + 1, replace, track, moved, keys);
      if (copy !== child) changed = true;
    }
    out.push(copy);
    npos += copy.nodeSize;
    pos += child.nodeSize;
    i++;
  }
  return changed ? node.copy(Fragment.from(out)) : node;
}

/**
 * Resolve revisions (positions in tr.doc as it is now). Paragraphs are done last first, so
 * nothing done in one moves the ones before it. Returns whether paragraph properties or
 * structure changed.
 */
function resolvePass(tr: Transaction, revs: Revision[], accept: boolean): boolean {
  const byGroup = new Map<number, Revision[]>();
  for (const r of revs) {
    const list = byGroup.get(r.group);
    if (list) list.push(r);
    else byGroup.set(r.group, [r]);
  }
  let paragraphs = false;
  for (const g of [...byGroup.keys()].sort((a, b) => b - a)) {
    const list = byGroup.get(g)!;
    const base = tr.steps.length;
    const inner = list.filter((r) => r.kind !== 'paraMark');
    if (inner.length) resolveInParagraph(tr, g, list[0].groupEnd, inner, accept);
    if (inner.some((r) => r.kind === 'paraFormat')) paragraphs = true;
    // The paragraph mark last: joining takes the (already resolved) next paragraph's properties.
    for (const r of list) {
      if (r.kind !== 'paraMark') continue;
      paragraphs = true;
      resolveMark(tr, r, accept, tr.mapping.slice(base));
    }
  }
  return paragraphs;
}

function resolve(state: EditorState, revs: Revision[], accept: boolean, all: boolean): Transaction | null {
  if (!revs.length) return null;
  const tr = state.tr;
  let paragraphs = all ? resolveEverything(tr, revs, accept) : resolvePass(tr, revs, accept);
  if (all) {
    // Revisions inside others (a formatting change in a restored deletion ...) show up once
    // those are resolved; Word's Accept / Reject All resolves them too.
    for (let pass = 0; pass < 4; pass++) {
      const more = collectRevisions(tr.doc);
      if (!more.length) break;
      const steps = tr.steps.length;
      paragraphs = resolveEverything(tr, more, accept) || paragraphs;
      if (tr.steps.length === steps) break;
    }
  }
  if (revs.some((r) => r.move)) removeFinishedMoveRanges(tr);
  if (!tr.docChanged) return null;
  if (paragraphs) tr.setMeta(NUMBERING_CHANGED, true);
  return tr;
}

/** Once no move is left, the markers that delimited moves go too (Word removes them with the move). */
function removeFinishedMoveRanges(tr: Transaction): void {
  if (collectRevisions(tr.doc).some((r) => r.move)) return;
  const ranges: { pos: number; size: number }[] = [];
  tr.doc.descendants((node, pos) => {
    if ((node.type === schema.nodes.raw_inline || node.type === schema.nodes.raw_block) && MOVE_RANGES.has(node.attrs.label)) {
      ranges.push({ pos, size: node.nodeSize });
    }
    return true;
  });
  for (const { pos, size } of ranges.reverse()) {
    const $p = tr.doc.resolve(pos);
    // Never leave a table cell or the document without a block.
    if (tr.doc.nodeAt(pos)?.isBlock && $p.parent.childCount === 1) continue;
    tr.delete(pos, pos + size);
  }
}

function resolveCommand(which: 'selection' | 'all', accept: boolean): Command {
  return (state, dispatch) => {
    const revs = which === 'all' ? collectRevisions(state.doc) : revisionsAtSelection(state);
    const tr = resolve(state, revs, accept, which === 'all');
    if (!tr) return false;
    // Its own undo step, even right after typing; never tracked itself.
    dispatch?.(closeHistory(tr).setMeta(NO_TRACK, true).scrollIntoView());
    return true;
  };
}

export const acceptRevision = resolveCommand('selection', true);
export const rejectRevision = resolveCommand('selection', false);
export const acceptAllRevisions = resolveCommand('all', true);
export const rejectAllRevisions = resolveCommand('all', false);

// ----- Enter -----

/**
 * After Enter split the paragraph at `pos` (tracking off, as in Word): the first half ends
 * with a new, untracked paragraph mark, so it loses the mark's revision and the section
 * break, which stay with the last half; neither half keeps w:pPrChange. Pieces of the same w:p
 * (split at page breaks) before the split follow the first half, those after it the last half.
 */
export function afterParagraphSplit(tr: Transaction, pos: number): Transaction {
  const first = tr.doc.nodeAt(pos);
  if (!first || !isPara(first)) return tr;
  const $p = tr.doc.resolve(pos);
  const parent = $p.parent;
  const index = $p.index();
  const second = parent.maybeChild(index + 1);
  const { first: firstPPr, last: lastPPr } = splitParagraphPPr(first.attrs.pPr);
  tr.setNodeMarkup(pos, undefined, { ...first.attrs, pPr: firstPPr, sectPr: null }, first.marks);
  const id = first.attrs.brGroup as string | null;
  if (!id) return tr;
  for (let k = index - 1, p = pos; k >= 0; k--) {
    const n = parent.child(k);
    if (groupId(n) !== id) break;
    p -= n.nodeSize;
    if (isPara(n)) tr.setNodeMarkup(p, undefined, { ...n.attrs, pPr: firstPPr, sectPr: null }, n.marks);
  }
  if (!second || !isPara(second)) return tr;
  const later: { node: PMNode; pos: number }[] = [];
  for (let k = index + 2, p = pos + first.nodeSize + second.nodeSize; k < parent.childCount; k++) {
    const n = parent.child(k);
    if (groupId(n) !== id) break;
    later.push({ node: n, pos: p });
    p += n.nodeSize;
  }
  if (!later.length) return tr;
  const g = newLayerId();
  tr.setNodeMarkup(pos + first.nodeSize, undefined, { ...second.attrs, brGroup: g }, second.marks);
  for (const { node, pos: p } of later) {
    if (isPara(node)) tr.setNodeMarkup(p, undefined, { ...node.attrs, pPr: lastPPr, brGroup: g }, node.marks);
    else tr.setNodeMarkup(p, undefined, { ...node.attrs, group: g });
  }
  return tr;
}

// ----- navigation & markup -----

function selectionFor(doc: PMNode, r: Revision): Selection {
  if (r.kind === 'del' || (isBreak(doc.nodeAt(r.from)) && r.to === r.from + 1)) return NodeSelection.create(doc, r.from);
  if (r.kind === 'paraMark') return Selection.near(doc.resolve(r.to));
  if (r.kind === 'paraFormat') return Selection.near(doc.resolve(r.from));
  return TextSelection.between(doc.resolve(r.from), doc.resolve(r.to));
}

/** Go to the next (1) or previous (-1) revision, wrapping around like Word offers to. */
export function goToRevision(dir: 1 | -1): Command {
  return (state, dispatch) => {
    const revs = collectRevisions(state.doc);
    if (!revs.length) return false;
    const cur = focused(state);
    let target: Revision | undefined;
    if (cur) target = revs[(revs.indexOf(cur) + dir + revs.length) % revs.length];
    else {
      const at = state.selection.from;
      target = dir > 0 ? revs.find((r) => r.from >= at) ?? revs[0] : revs.filter((r) => r.from < at).pop() ?? revs[revs.length - 1];
    }
    if (dispatch) {
      const tr = state.tr.setSelection(selectionFor(state.doc, target)).setMeta(reviewKey, { focus: target.key } as ReviewMeta);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

// ----- the revisions pane (persona-300) -----

/** Select one revision (by its key), as 上一個 / 下一個 do. */
export function goToRevisionKey(key: string): Command {
  return (state, dispatch) => {
    const r = collectRevisions(state.doc).find((x) => x.key === key);
    if (!r) return false;
    dispatch?.(state.tr.setSelection(selectionFor(state.doc, r)).setMeta(reviewKey, { focus: r.key } as ReviewMeta).scrollIntoView());
    return true;
  };
}

/** Accept or reject one revision (by its key): one undo step, never tracked itself. */
export function resolveRevisionKey(key: string, accept: boolean): Command {
  return (state, dispatch) => {
    const r = collectRevisions(state.doc).find((x) => x.key === key);
    const tr = r ? resolve(state, [r], accept, false) : null;
    if (!tr) return false;
    dispatch?.(closeHistory(tr).setMeta(NO_TRACK, true).scrollIntoView());
    return true;
  };
}

/** What a revision is about, for the pane: the inserted / deleted text, or its paragraph's (at most `max` characters). */
export function revisionExcerpt(doc: PMNode, r: Revision, max = 40): string {
  let text = '';
  if (r.kind === 'del') {
    const node = doc.nodeAt(r.from);
    text = node ? deletedInfo(node.attrs.xml ?? '')?.text ?? '' : '';
  } else if (r.kind === 'ins' || r.kind === 'format') {
    text = doc.textBetween(r.from, Math.min(r.to, doc.content.size), ' ', '');
  } else {
    const at = r.para ?? r.group;
    const para = doc.nodeAt(at);
    text = para?.isTextblock ? para.textContent : '';
  }
  text = text.replace(/\s+/g, ' ').trim();
  return text.length > max ? text.slice(0, max) + '…' : text;
}

/** How a revision is said: 「插入，王小明，115/09/26 14:32」. */
export function revisionSpoken(r: Revision): string {
  const when = rocDate(r.date, r.dateUtc);
  return `${revisionLabel(r)}，${r.author || '未知作者'}${when ? '，' + when : ''}`;
}

/** Show or hide revision marks (hidden = the document as if every change were accepted). */
export const toggleRevisionMarks: Command = (state, dispatch) => {
  const s = reviewKey.getState(state);
  if (!s) return false;
  dispatch?.(state.tr.setMeta(reviewKey, { show: !s.show } as ReviewMeta).setMeta('addToHistory', false));
  return true;
};

export interface ReviewSummary {
  count: number;
  show: boolean;
  /** Revisions at the selection (what 接受 / 拒絕 act on). */
  here: number;
  current: Revision | null;
}

export function reviewSummary(state: EditorState | null | undefined): ReviewSummary {
  if (!state) return { count: 0, show: true, here: 0, current: null };
  const here = revisionsAtSelection(state);
  return {
    count: collectRevisions(state.doc).length,
    show: reviewKey.getState(state)?.show !== false,
    here: here.length,
    current: here[0] ?? null,
  };
}

// ----- loading -----

/**
 * Runs loads one after another; each can ask whether a newer one has started since, so an
 * editor whose document changes while the previous one is still opening only ever shows the
 * latest document's comments and notices.
 */
export function latestLoads(): <T>(task: (current: () => boolean) => Promise<T>) => Promise<T | undefined> {
  let seq = 0;
  let tail: Promise<unknown> = Promise.resolve();
  return <T>(task: (current: () => boolean) => Promise<T>) => {
    const n = ++seq;
    const current = () => n === seq;
    const run = tail.then(() => (current() ? task(current) : undefined));
    tail = run.catch(() => undefined);
    return run;
  };
}

// ----- comments -----

export interface CommentRange {
  /** After the start marker / at the end marker; equal when the comment marks a point. */
  from: number;
  to: number;
  /** Position of the comment's reference mark, if any. */
  ref: number | null;
}

const ID = /\s(?:[\w.-]+:)?id="([^"]*)"/;
const REF = /commentReference\b[^>]*?\s(?:[\w.-]+:)?id="([^"]*)"/;
const rangeCache = new WeakMap<PMNode, Map<string, CommentRange>>();

/** Where each comment's range and reference mark are, by comment id. */
export function commentRanges(doc: PMNode): Map<string, CommentRange> {
  const cached = rangeCache.get(doc);
  if (cached) return cached;
  const out = new Map<string, CommentRange>();
  const get = (id: string) => {
    let r = out.get(id);
    if (!r) out.set(id, (r = { from: -1, to: -1, ref: null }));
    return r;
  };
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.raw_inline && node.type !== schema.nodes.raw_block) return true;
    const { label, xml } = node.attrs;
    if (label === 'commentRangeStart' || label === 'commentRangeEnd') {
      const id = ID.exec(xml)?.[1];
      if (id != null) {
        if (label === 'commentRangeStart') get(id).from = pos + node.nodeSize;
        else get(id).to = pos;
      }
    } else if (label === '註解') {
      const id = REF.exec(xml)?.[1];
      if (id != null) get(id).ref = pos;
    }
    return false;
  });
  for (const r of out.values()) {
    if (r.from < 0) r.from = r.to >= 0 ? r.to : r.ref ?? 0;
    if (r.to < r.from) r.to = r.from;
  }
  rangeCache.set(doc, out);
  return out;
}

/** The comments the editor should show titles for (read from word/comments.xml). */
export function setComments(view: EditorView, comments: DocComment[]): void {
  view.dispatch(view.state.tr.setMeta(reviewKey, { comments } as ReviewMeta).setMeta('addToHistory', false));
}

/** Id of the comment whose range holds the cursor (innermost), if any. */
export function commentAtSelection(state: EditorState): string | null {
  const at = state.selection.from;
  let best: { id: string; size: number } | null = null;
  for (const [id, r] of commentRanges(state.doc)) {
    const inRange = (r.from <= at && at <= r.to) || (r.ref != null && (at === r.ref || at === r.ref + 1));
    if (inRange && (!best || r.to - r.from < best.size)) best = { id, size: r.to - r.from };
  }
  return best?.id ?? null;
}

/** Text a comment is attached to. */
export function commentedText(doc: PMNode, id: string): string {
  const r = commentRanges(doc).get(id);
  return r && r.to > r.from ? doc.textBetween(r.from, r.to, '\n', '') : '';
}

/** Select a comment's range and scroll it into view. */
export function revealComment(view: EditorView, id: string): boolean {
  const r = commentRanges(view.state.doc).get(id);
  if (!r) return false;
  const doc = view.state.doc;
  const sel = r.to > r.from ? TextSelection.create(doc, r.from, r.to) : TextSelection.near(doc.resolve(r.ref ?? r.from));
  view.dispatch(view.state.tr.setSelection(sel).scrollIntoView());
  view.focus();
  return true;
}

// ----- editing comments -----
//
// A comment is its entry in the document's comment list (doc attribute `comments`, written to
// comments.xml and the parts beside it by docx/comments.ts) plus, in the text, the markers Word
// writes: w:commentRangeStart / w:commentRangeEnd around what it comments on and a run with
// w:commentReference after the end. Each change is one transaction, so one undo step.

/** The document's comments: the list on the document, else those given with setComments. */
export function commentsOf(state: EditorState): DocComment[] {
  return (state.doc.attrs.comments as DocComment[] | null) ?? reviewKey.getState(state)?.comments ?? [];
}

const MARKER_ID = /<(?:[\w.-]+:)?comment(?:RangeStart|RangeEnd|Reference)\b[^>]*?\s(?:[\w.-]+:)?id="(\d+)"/g;

/** An id for a new comment: after every id in the comment list and in the text (kept deletions included). */
export function nextCommentId(state: EditorState): string {
  let max = -1;
  const see = (id: string) => {
    const n = Number(id);
    if (Number.isInteger(n) && n > max) max = n;
  };
  for (const c of commentsOf(state)) see(c.id);
  state.doc.descendants((node) => {
    if (node.type === schema.nodes.raw_inline || node.type === schema.nodes.raw_block) {
      for (const m of (node.attrs.xml as string).matchAll(MARKER_ID)) see(m[1]);
      return false;
    }
    return true;
  });
  return String(max + 1);
}

/** The markers of each comment in the text, by id: node positions. */
function commentMarkers(doc: PMNode): Map<string, { start: number | null; end: number | null; refs: number[]; all: { pos: number; size: number }[] }> {
  const out = new Map<string, { start: number | null; end: number | null; refs: number[]; all: { pos: number; size: number }[] }>();
  const get = (id: string) => {
    let m = out.get(id);
    if (!m) out.set(id, (m = { start: null, end: null, refs: [], all: [] }));
    return m;
  };
  doc.descendants((node, pos) => {
    if (!isCommentMarker(node)) return !node.isAtom;
    const { label, xml } = node.attrs;
    const id = (label === '註解' ? REF : ID).exec(xml)?.[1];
    if (id == null) return false;
    const m = get(id);
    if (label === 'commentRangeStart') m.start ??= pos;
    else if (label === 'commentRangeEnd') m.end ??= pos;
    else m.refs.push(pos);
    m.all.push({ pos, size: node.nodeSize });
    return false;
  });
  return out;
}

/** Wrappers (a hyperlink, a tracked insertion ...) the text on both sides of `$pos` is in: a marker there stays inside them. */
function wrapperMarks($pos: ReturnType<PMNode['resolve']>): readonly Mark[] {
  const before = $pos.nodeBefore;
  const after = $pos.nodeAfter;
  if (!before || !after) return [];
  return before.marks.filter((m) => (m.type === schema.marks.inlineWrap || m.type === schema.marks.link) && m.isInSet(after.marks));
}

/** A comment marker for position `pos`: inline in a paragraph, else (between blocks) a block. */
function markerNode(doc: PMNode, pos: number, label: string, xml: string): PMNode {
  const $pos = doc.resolve(pos);
  if ($pos.parent.inlineContent) {
    return schema.nodes.raw_inline.create({ xml, label, hidden: label !== '註解' }, null, wrapperMarks($pos));
  }
  return schema.nodes.raw_block.create({ xml, label, hidden: true });
}

const referenceXml = (id: string, style: string | null) =>
  `<w:r>${style ? `<w:rPr><w:rStyle w:val="${style.replace(/[<>&"]/g, '')}"/></w:rPr>` : ''}<w:commentReference w:id="${id}"/></w:r>`;

/** What a new comment or reply says, and who wrote it when. */
export interface NewComment {
  text: string;
  author: string;
  initials: string | null;
  date: string;
  dateUtc: string | null;
  /** Style id of Word's "annotation reference" (the comment mark), when the document has it. */
  referenceStyle?: string | null;
}

const isInlinePos = (doc: PMNode, pos: number) => pos >= 0 && pos <= doc.content.size && doc.resolve(pos).parent.inlineContent;

/** The characters of a paragraph with their positions; markers and other hidden things are skipped. */
function textWithPositions(parent: PMNode, start: number): { text: string; pos: number[] } {
  let text = '';
  const pos: number[] = [];
  parent.forEach((child, offset) => {
    if (child.isText) {
      for (let i = 0; i < child.text!.length; i++) {
        text += child.text![i];
        pos.push(start + offset + i);
      }
    } else if (!child.attrs.hidden) {
      text += String.fromCharCode(0xfffc); // object replacement character: a picture, a symbol ...
      pos.push(start + offset);
    }
  });
  return { text, pos };
}

/** The word at a cursor (as Word takes it for a new comment), or null when the cursor touches none. */
function wordAt(state: EditorState, at: number): { from: number; to: number } | null {
  const $at = state.doc.resolve(at);
  const { text, pos } = textWithPositions($at.parent, $at.start());
  if (!text) return null;
  // Index of the first character after the cursor.
  let idx = pos.findIndex((p) => p >= at);
  if (idx < 0) idx = text.length;
  type Seg = { index: number; segment: string; isWordLike?: boolean };
  let segs: Seg[];
  const Segmenter = (Intl as unknown as { Segmenter?: new (l: string, o: object) => { segment(t: string): Iterable<Seg> } }).Segmenter;
  if (Segmenter) segs = Array.from(new Segmenter('zh-Hant', { granularity: 'word' }).segment(text));
  else segs = Array.from(text.matchAll(/[\p{L}\p{N}_]+|[^\p{L}\p{N}_]+/gu), (m) => ({ index: m.index!, segment: m[0], isWordLike: /[\p{L}\p{N}_]/u.test(m[0]) }));
  const word = (i: number) => segs.find((s) => s.isWordLike && s.index <= i && i < s.index + s.segment.length);
  const hit = word(idx) ?? (idx > 0 ? word(idx - 1) : undefined);
  if (!hit) return null;
  const last = hit.index + hit.segment.length - 1;
  return { from: pos[hit.index], to: pos[last] + 1 };
}

/**
 * What a new comment would be on: the selection, or with nothing selected the word at the
 * cursor (as Word does), else the cursor itself. Null where comments can't go.
 */
export function commentTarget(state: EditorState): { from: number; to: number } | null {
  const sel = state.selection;
  const doc = state.doc;
  if (sel.empty) {
    if (!sel.$from.parent.inlineContent) return null;
    return wordAt(state, sel.from) ?? { from: sel.from, to: sel.from };
  }
  let { from, to } = sel;
  if (!isInlinePos(doc, from)) {
    const s = Selection.findFrom(doc.resolve(from), 1, true);
    if (!s || s.from > to) return null;
    from = s.from;
  }
  if (!isInlinePos(doc, to)) {
    const s = Selection.findFrom(doc.resolve(to), -1, true);
    if (!s || s.to < from) return null;
    to = s.to;
  }
  return from <= to ? { from, to } : null;
}

/** The comment list changed, as one undo step of its own. */
function withList(tr: Transaction, list: DocComment[]): Transaction {
  return closeHistory(tr.setDocAttribute('comments', list).setMeta(NO_TRACK, true));
}

/** Add a comment on `range` (see commentTarget); returns its id, null when it can't go there. */
export function insertComment(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  range: { from: number; to: number },
  c: NewComment,
): string | null {
  const doc = state.doc;
  const { from, to } = range;
  if (from > to || !isInlinePos(doc, from) || !isInlinePos(doc, to)) return null;
  const id = nextCommentId(state);
  const tr = state.tr;
  tr.insert(to, [
    markerNode(doc, to, 'commentRangeEnd', `<w:commentRangeEnd w:id="${id}"/>`),
    markerNode(doc, to, '註解', referenceXml(id, c.referenceStyle ?? null)),
  ]);
  tr.insert(from, markerNode(doc, from, 'commentRangeStart', `<w:commentRangeStart w:id="${id}"/>`));
  const entry: DocComment = {
    id, author: c.author, initials: c.initials, date: c.date, dateUtc: c.dateUtc, text: c.text, parentId: null, done: false, created: true,
  };
  dispatch?.(withList(tr, [...commentsOf(state), entry]));
  return id;
}

/**
 * Reply to a comment's thread (Word threads replies under the top comment). The reply covers
 * the same text: its start marker follows the thread's, its end and mark follow the thread's last mark.
 */
export function insertReply(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  parentId: string,
  c: NewComment,
): string | null {
  const list = commentsOf(state);
  const root = threadRoot(list, parentId);
  if (!root) return null;
  const id = nextCommentId(state);
  const doc = state.doc;
  const markers = commentMarkers(doc);
  const thread = list.filter((x) => threadRoot(list, x.id) === root).map((x) => x.id);
  const own = markers.get(root.id);
  let endAt: number | null = null;
  for (const t of thread) for (const r of markers.get(t)?.refs ?? []) endAt = Math.max(endAt ?? -1, r + 1);
  if (endAt == null && own?.end != null) endAt = own.end + 1;
  const startAt = own?.start != null ? own.start + 1 : endAt;
  const tr = state.tr;
  if (endAt != null && startAt != null && startAt <= endAt) {
    tr.insert(endAt, [
      markerNode(doc, endAt, 'commentRangeEnd', `<w:commentRangeEnd w:id="${id}"/>`),
      markerNode(doc, endAt, '註解', referenceXml(id, c.referenceStyle ?? null)),
    ]);
    tr.insert(startAt, markerNode(doc, startAt, 'commentRangeStart', `<w:commentRangeStart w:id="${id}"/>`));
  }
  const entry: DocComment = {
    id, author: c.author, initials: c.initials, date: c.date, dateUtc: c.dateUtc, text: c.text, parentId: root.id, done: false, created: true,
  };
  dispatch?.(withList(tr, [...list, entry]));
  return id;
}

/** Change a comment's text. */
export function setCommentText(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, id: string, text: string): boolean {
  const list = commentsOf(state);
  const i = list.findIndex((c) => c.id === id);
  if (i < 0 || list[i].text === text) return false;
  dispatch?.(withList(state.tr, list.map((c, k) => (k === i ? { ...c, text } : c))));
  return true;
}

/** Mark a comment's thread resolved (done) or open it again, as Word does on the top comment. */
export function setCommentDone(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, id: string, done: boolean): boolean {
  const list = commentsOf(state);
  const root = threadRoot(list, id);
  if (!root || root.done === done) return false;
  dispatch?.(withList(state.tr, list.map((c) => (c === root ? { ...c, done } : c))));
  return true;
}

/** Delete a comment, its replies (as Word does for the top comment of a thread) and their markers. */
export function removeComment(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, id: string): boolean {
  const list = commentsOf(state);
  if (!list.some((c) => c.id === id)) return false;
  const ids = withReplies(list, id);
  const markers = commentMarkers(state.doc);
  const ops: { pos: number; size: number; xml?: string }[] = [...ids].flatMap((x) => markers.get(x)?.all ?? []);
  // Markers inside something kept whole (a tracked deletion) go from its XML.
  state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.raw_inline && node.type !== schema.nodes.raw_block) return !node.isAtom;
    if (isCommentMarker(node)) return false;
    const xml = node.attrs.xml as string;
    if ([...xml.matchAll(MARKER_ID)].some((m) => ids.has(m[1]))) ops.push({ pos, size: node.nodeSize, xml: withoutCommentMarkers(xml, ids) });
    return false;
  });
  const tr = state.tr;
  for (const op of ops.sort((a, b) => b.pos - a.pos)) {
    if (op.xml == null) tr.delete(op.pos, op.pos + op.size);
    else tr.setNodeMarkup(op.pos, undefined, { ...tr.doc.nodeAt(op.pos)!.attrs, xml: op.xml });
  }
  dispatch?.(withList(tr, list.filter((c) => !ids.has(c.id))));
  return true;
}

/** XML without the range markers and marks (with their runs) of the given comments. */
function withoutCommentMarkers(xml: string, ids: Set<string>): string {
  try {
    const root = parseFragment(xml);
    const found = Array.from(root.getElementsByTagNameNS(NS.w, '*')).filter(
      (e) => ['commentRangeStart', 'commentRangeEnd', 'commentReference'].includes(e.localName) && ids.has(e.getAttributeNS(NS.w, 'id') ?? ''),
    );
    for (const e of found) {
      const run = e.parentElement;
      const alone = e.localName === 'commentReference' && run && run !== root && run.namespaceURI === NS.w && run.localName === 'r' &&
        Array.from(run.children).every((c) => c === e || c.localName === 'rPr');
      (alone ? run! : e).remove();
    }
    return serializeXml(root);
  } catch {
    return xml;
  }
}
