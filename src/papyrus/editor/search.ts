// Find and replace.
//
// Matches lie inside one paragraph (table cells included) and may span formatting runs,
// but not structural wrappers: a match stays inside one link, one field result and one
// set of inline wrappers (w:ins, w:sdt ...). Atoms (fields, field codes, kept XML, pictures,
// tabs, line breaks) end a stretch of text; invisible markers such as bookmarks are skipped
// over and survive a replacement. Replacement text takes the formatting of the first
// matched character.
//
// Big documents: after an edit only the paragraphs it touched are searched again, moving
// between matches only swaps the current highlight, and at most MAX_HIGHLIGHTS matches are
// highlighted (all are counted).
//
// The plugin searches one editor (the body, or the header/footer being edited). The find panel
// also searches the other parts of the document (SearchPart: the body and each header/footer
// part shown on a page, see DocxEditor.searchParts) with findMatches, moves between them with
// nextMatchIndex / selectMatch, and replaces in each with replaceMatchesTr (GOV-ISSUE-012).

import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { Fragment, Mark, type Node as PMNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import { schema } from './schema';
import './features.css';

export interface SearchMatch {
  from: number;
  to: number;
}

export interface SearchState {
  query: string;
  caseSensitive: boolean;
  /** 區分全形／半形: off by default, so 114 finds １１４ (persona-300). */
  widthSensitive: boolean;
  matches: SearchMatch[];
  /** Index of the current match, -1 when there is none. */
  current: number;
  decorations: DecorationSet;
}

interface SearchMeta {
  query?: string;
  caseSensitive?: boolean;
  widthSensitive?: boolean;
  /** Make this match index current. */
  current?: number;
}

export const searchKey = new PluginKey<SearchState>('papyrus-search');

/** Matches beyond this many are counted but not highlighted (the current one always is). */
export const MAX_HIGHLIGHTS = 5000;
/** An edit made of more changes than this is searched again as a whole. */
const MAX_RANGES = 200;

const STRUCTURAL = new Set(['inlineWrap', 'link', 'fieldResult']);
/** Hidden markers that are shown after all (tracked deletions) and so end a stretch of text. */
const VISIBLE_HIDDEN = new Set(['del', 'moveFrom']);

const structural = (marks: readonly Mark[]) => marks.filter((m) => STRUCTURAL.has(m.type.name));

/**
 * Case (and width) folding that keeps one UTF-16 unit per unit, so indexes map back to
 * positions. Final sigma folds like sigma; Turkish İ (whose lower case is two units) folds to i.
 * Unless `widthSensitive`, full-width letters, digits and signs (Ａ１２％, U+FF01–FF5E) and the
 * ideographic space fold to their half-width forms, as Word's 「全半形須相符」 off: 114 finds １１４.
 */
function fold(text: string, caseSensitive: boolean, widthSensitive = true): string {
  if (caseSensitive && widthSensitive) return text;
  let out = '';
  for (let ch of text) {
    if (ch.length !== 1) {
      out += ch; // surrogate pair: keep as is
      continue;
    }
    if (!widthSensitive) {
      const c = ch.charCodeAt(0);
      if (c >= 0xff01 && c <= 0xff5e) ch = String.fromCharCode(c - 0xfee0);
      else if (c === 0x3000) ch = ' ';
    }
    if (caseSensitive) out += ch;
    else if (ch === 'ς') out += 'σ';
    else if (ch === 'İ') out += 'i';
    else {
      const lower = ch.toLowerCase();
      out += lower.length === 1 ? lower : ch;
    }
  }
  return out;
}

/** Add the matches inside one textblock (at `pos`) to `out`. */
function scanBlock(node: PMNode, pos: number, needle: string, caseSensitive: boolean, widthSensitive: boolean, out: SearchMatch[]): void {
  let text = '';
  let at: number[] = [];
  let marks: readonly Mark[] | null = null;
  const flush = () => {
    for (let i = text.indexOf(needle); i >= 0; i = text.indexOf(needle, i + needle.length)) {
      out.push({ from: at[i], to: at[i + needle.length - 1] + 1 });
    }
    text = '';
    at = [];
    marks = null;
  };
  node.forEach((child, offset) => {
    const start = pos + 1 + offset;
    if (child.isText) {
      const s = structural(child.marks);
      if (marks && !Mark.sameSet(marks, s)) flush();
      marks = s;
      text += fold(child.text!, caseSensitive, widthSensitive);
      for (let i = 0; i < child.text!.length; i++) at.push(start + i);
    } else if (
      child.type === schema.nodes.raw_inline && child.attrs.hidden &&
      !VISIBLE_HIDDEN.has(child.attrs.label) && child.attrs.label !== 'field'
    ) {
      // Invisible marker (bookmark, comment range ...): the text continues across it.
      // Field codes are not: text before and after a field is not one word.
    } else flush();
  });
  flush();
}

/** Every match of `query` in the document, in document order. */
export function findMatches(doc: PMNode, query: string, caseSensitive = false, widthSensitive = false): SearchMatch[] {
  const out: SearchMatch[] = [];
  if (!query) return out;
  const needle = fold(query, caseSensitive, widthSensitive);
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    scanBlock(node, pos, needle, caseSensitive, widthSensitive, out);
    return false;
  });
  return out;
}

/** Ranges of `tr.doc` that `tr` changed (content, marks or attributes); null when there are too many. */
function changedRanges(tr: Transaction): [number, number][] | null {
  let ranges: [number, number][] = [];
  for (let i = 0; i < tr.steps.length; i++) {
    const map = tr.mapping.maps[i];
    ranges = ranges.map(([f, t]) => [map.map(f, -1), map.map(t, 1)]);
    let mapped = false;
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      ranges.push([newStart, newEnd]);
      mapped = true;
    });
    if (!mapped) {
      // Mark and attribute steps move nothing but still change what matches.
      const s = tr.steps[i] as unknown as { from?: number; to?: number; pos?: number };
      if (typeof s.from === 'number' && typeof s.to === 'number') ranges.push([s.from, s.to]);
      else if (typeof s.pos === 'number') ranges.push([s.pos, s.pos + 1]);
      else return null;
    }
    if (ranges.length > MAX_RANGES) return null;
  }
  return ranges;
}

/** Index of the first match at or after `pos` (wrapping to the first one). */
function indexFrom(matches: SearchMatch[], pos: number): number {
  if (!matches.length) return -1;
  const i = firstAtOrAfter(matches, pos);
  return i < matches.length ? i : 0;
}

/** Smallest index whose match starts at or after `pos` (matches.length when none). */
function firstAtOrAfter(matches: SearchMatch[], pos: number): number {
  let lo = 0;
  let hi = matches.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (matches[mid].from < pos) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

interface Update {
  matches: SearchMatch[];
  /** Textblocks searched again, sorted [start, end); null when the whole document was. */
  blocks: [number, number][] | null;
}

/** The matches after `tr`: kept (mapped) outside the textblocks it touched, searched again inside them. */
function updateMatches(matches: SearchMatch[], tr: Transaction, query: string, caseSensitive: boolean, widthSensitive: boolean): Update {
  const doc = tr.doc;
  const ranges = changedRanges(tr);
  if (!ranges) return { matches: findMatches(doc, query, caseSensitive, widthSensitive), blocks: null };
  const blocks: [number, number][] = [];
  const seen = new Set<number>();
  for (const [from, to] of ranges) {
    // One position wider, so a change at a paragraph edge counts for that paragraph.
    doc.nodesBetween(Math.max(0, Math.min(from, to) - 1), Math.min(doc.content.size, Math.max(from, to) + 1), (node, pos) => {
      if (!node.isTextblock) return true;
      if (!seen.has(pos)) {
        seen.add(pos);
        blocks.push([pos, pos + node.nodeSize]);
      }
      return false;
    });
  }
  blocks.sort((a, b) => a[0] - b[0]);
  const touched = (pos: number) => {
    let lo = 0;
    let hi = blocks.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (blocks[mid][1] <= pos) lo = mid + 1;
      else hi = mid;
    }
    return lo < blocks.length && blocks[lo][0] <= pos;
  };
  const kept: SearchMatch[] = [];
  for (const m of matches) {
    const a = tr.mapping.mapResult(m.from, 1);
    const b = tr.mapping.mapResult(m.to, -1);
    if (a.deletedAfter || b.deletedBefore || b.pos - a.pos !== m.to - m.from || touched(a.pos)) continue;
    // Unmoved matches stay the same objects.
    kept.push(a.pos === m.from ? m : { from: a.pos, to: b.pos });
  }
  const found: SearchMatch[] = [];
  const needle = fold(query, caseSensitive, widthSensitive);
  for (const [pos] of blocks) scanBlock(doc.nodeAt(pos)!, pos, needle, caseSensitive, widthSensitive, found);
  if (!found.length) return { matches: kept, blocks };
  // Both lists are sorted and never overlap: merge them.
  const out: SearchMatch[] = [];
  let i = 0;
  let j = 0;
  while (i < kept.length || j < found.length) {
    if (j >= found.length || (i < kept.length && kept[i].from < found[j].from)) out.push(kept[i++]);
    else out.push(found[j++]);
  }
  return { matches: out, blocks };
}

const hit = (m: SearchMatch, current: boolean) =>
  Decoration.inline(m.from, m.to, { class: current ? 'dx-search-hit dx-search-current' : 'dx-search-hit' }, { current });

function decorate(doc: PMNode, matches: SearchMatch[], current: number): DecorationSet {
  if (!matches.length) return DecorationSet.empty;
  const decos: Decoration[] = [];
  const n = Math.min(matches.length, MAX_HIGHLIGHTS);
  for (let i = 0; i < n; i++) decos.push(hit(matches[i], i === current));
  if (current >= n) decos.push(hit(matches[current], true));
  return DecorationSet.create(doc, decos);
}

/** The highlights with another match made current. */
function moveCurrent(set: DecorationSet, doc: PMNode, matches: SearchMatch[], from: number, to: number): DecorationSet {
  const remove: Decoration[] = [];
  const add: Decoration[] = [];
  const at = (m: SearchMatch) => set.find(m.from, m.to).filter((d) => d.from === m.from && d.to === m.to);
  const old = matches[from];
  if (old) {
    remove.push(...at(old));
    if (from < MAX_HIGHLIGHTS) add.push(hit(old, false));
  }
  const next = matches[to];
  if (next) {
    remove.push(...at(next));
    add.push(hit(next, true));
  }
  return set.remove(remove).add(doc, add);
}

const EMPTY: SearchState = { query: '', caseSensitive: false, widthSensitive: false, matches: [], current: -1, decorations: DecorationSet.empty };

const clampIndex = (i: number, n: number) => (n ? Math.max(0, Math.min(i, n - 1)) : -1);

export function searchPlugin(): Plugin<SearchState> {
  return new Plugin<SearchState>({
    key: searchKey,
    state: {
      init: () => EMPTY,
      apply(tr, value, _old, state) {
        const meta = tr.getMeta(searchKey) as SearchMeta | undefined;
        if (meta && (meta.query !== undefined || meta.caseSensitive !== undefined || meta.widthSensitive !== undefined)) {
          // A new search.
          const query = meta.query ?? value.query;
          const caseSensitive = meta.caseSensitive ?? value.caseSensitive;
          const widthSensitive = meta.widthSensitive ?? value.widthSensitive;
          if (!query) return { ...EMPTY, caseSensitive, widthSensitive };
          const matches = findMatches(state.doc, query, caseSensitive, widthSensitive);
          const current = meta.current != null ? clampIndex(meta.current, matches.length) : indexFrom(matches, state.selection.from);
          return { query, caseSensitive, widthSensitive, matches, current, decorations: decorate(state.doc, matches, current) };
        }
        if (!value.query) return value;
        if (!tr.docChanged) {
          if (meta?.current == null) return value;
          const current = clampIndex(meta.current, value.matches.length);
          if (current === value.current) return value;
          return { ...value, current, decorations: moveCurrent(value.decorations, state.doc, value.matches, value.current, current) };
        }
        const { matches, blocks } = updateMatches(value.matches, tr, value.query, value.caseSensitive, value.widthSensitive);
        let current: number;
        if (meta?.current != null) current = clampIndex(meta.current, matches.length);
        else {
          // Stay at (or just after) where the current match was.
          const was = value.matches[value.current];
          current = indexFrom(matches, was ? tr.mapping.map(was.from, -1) : state.selection.from);
        }
        let decorations: DecorationSet;
        if (!blocks || matches.length > MAX_HIGHLIGHTS || value.matches.length > MAX_HIGHLIGHTS) {
          decorations = decorate(state.doc, matches, current);
        } else {
          // Move the highlights along; redo them in the paragraphs searched again, and the current one.
          let set = value.decorations.map(tr.mapping, state.doc);
          const inBlocks = (d: Decoration) => blocks.some(([from, to]) => d.from >= from && d.to <= to);
          const stale = blocks.flatMap(([from, to]) => set.find(from, to).filter((d) => d.from >= from && d.to <= to));
          const add: Decoration[] = [];
          for (const d of set.find(undefined, undefined, (spec) => spec.current)) {
            stale.push(d);
            // The match that was current elsewhere becomes a plain highlight.
            const k = firstAtOrAfter(matches, d.from);
            if (!inBlocks(d) && k !== current && matches[k]?.from === d.from && matches[k].to === d.to) add.push(hit(matches[k], false));
          }
          set = set.remove(stale);
          for (const [from, to] of blocks) {
            for (let k = firstAtOrAfter(matches, from); k < matches.length && matches[k].to <= to; k++) {
              if (k !== current) add.push(hit(matches[k], false));
            }
          }
          const cur = matches[current];
          if (cur) {
            set = set.remove(set.find(cur.from, cur.to).filter((d) => d.from === cur.from && d.to === cur.to));
            add.push(hit(cur, true));
          }
          decorations = set.add(state.doc, add);
        }
        return { ...value, matches, current, decorations };
      },
    },
    props: {
      decorations: (state) => searchKey.getState(state)?.decorations,
    },
  });
}

export function searchState(state: EditorState): SearchState {
  return searchKey.getState(state) ?? EMPTY;
}

/**
 * Search for `query` (empty clears the search). The first match after the cursor becomes current.
 * Full-width and half-width forms match each other unless `widthSensitive`.
 */
export const setSearch = (query: string, caseSensitive = false, widthSensitive = false): Command => (state, dispatch) => {
  if (!searchKey.getState(state)) return false;
  dispatch?.(state.tr.setMeta(searchKey, { query, caseSensitive, widthSensitive } satisfies SearchMeta));
  return true;
};

export const clearSearch: Command = setSearch('');

/** Select a match and scroll to it. */
function goTo(tr: Transaction, m: SearchMatch, index: number): Transaction {
  return tr.setSelection(TextSelection.create(tr.doc, m.from, m.to)).setMeta(searchKey, { current: index } satisfies SearchMeta).scrollIntoView();
}

/**
 * Go to the next (1) or previous (-1) match. From the current match this steps along;
 * after the cursor was moved elsewhere it continues from the cursor, as Word does.
 */
export const findNext = (dir: 1 | -1 = 1): Command => (state, dispatch) => {
  const s = searchState(state);
  const n = s.matches.length;
  if (!n) return false;
  const { from, to } = state.selection;
  const cur = s.matches[s.current];
  let i: number;
  if (cur && cur.from === from && cur.to === to) i = (s.current + dir + n) % n;
  else if (dir > 0) i = indexFrom(s.matches, from);
  else {
    // The last match ending at or before the cursor (matches never overlap, so ends are sorted too).
    let lo = 0;
    let hi = n;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (s.matches[mid].to <= from) lo = mid + 1;
      else hi = mid;
    }
    i = lo > 0 ? lo - 1 : n - 1;
  }
  dispatch?.(goTo(state.tr, s.matches[i], i));
  return true;
};

/**
 * The content replacing [first match start, last match end] of one textblock: the matched
 * text becomes `text` (with the first matched character's marks); text between the matches
 * and markers inside them (bookmarks ...) stay.
 */
function replaced(doc: PMNode, group: SearchMatch[], text: string): Fragment {
  const from = group[0].from;
  const to = group[group.length - 1].to;
  const $from = doc.resolve(from);
  const start = $from.start();
  const nodes: PMNode[] = [];
  let k = 0;
  $from.parent.forEach((child, offset) => {
    const cs = start + offset;
    const ce = cs + child.nodeSize;
    if (ce <= from || cs >= to) return;
    if (!child.isText) {
      nodes.push(child);
      return;
    }
    const end = Math.min(ce, to);
    for (let p = Math.max(cs, from); p < end; ) {
      while (k < group.length && group[k].to <= p) k++;
      const m = group[k];
      if (m && m.from <= p) {
        if (p === m.from && text) nodes.push(schema.text(text, child.marks));
        p = Math.min(end, m.to);
      } else {
        const stop = Math.min(end, m ? m.from : end);
        nodes.push(child.cut(p - cs, stop - cs));
        p = stop;
      }
    }
  });
  return Fragment.fromArray(nodes);
}

/** Replace the given matches (sorted, positions in `tr.doc`): one step per textblock, back to front. */
function replaceMatches(tr: Transaction, matches: SearchMatch[], text: string): void {
  const doc = tr.doc;
  let end = matches.length;
  while (end > 0) {
    const blockStart = doc.resolve(matches[end - 1].from).start();
    let begin = end - 1;
    while (begin > 0 && doc.resolve(matches[begin - 1].from).start() === blockStart) begin--;
    const group = matches.slice(begin, end);
    tr.replaceWith(group[0].from, group[group.length - 1].to, replaced(doc, group, text));
    end = begin;
  }
}

/**
 * Replace the current match and move on to the next one (wrapping around). With `advance`
 * false the cursor stays just after the replacement, so the caller can decide where to go on
 * (the next match may be in another part of the document).
 */
export const replaceCurrent = (replacement: string, advance = true): Command => (state, dispatch) => {
  const s = searchState(state);
  const m = s.matches[s.current];
  if (!m) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  replaceMatches(tr, [m], replacement);
  const after = tr.mapping.map(m.to);
  const { matches: next } = updateMatches(s.matches, tr, s.query, s.caseSensitive, s.widthSensitive);
  if (next.length && advance) {
    const i = indexFrom(next, after);
    goTo(tr, next[i], i);
  } else tr.setSelection(TextSelection.create(tr.doc, after)).scrollIntoView();
  dispatch(tr);
  return true;
};

/** Replace every match in one transaction: a single undo brings them all back. */
export const replaceAll = (replacement: string): Command => (state, dispatch) => {
  const s = searchState(state);
  if (!s.matches.length) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  replaceMatches(tr, s.matches, replacement);
  dispatch(closeHistory(tr).scrollIntoView());
  return true;
};


// ----- searching every part of the document (find panel) -----

/** Where a match is. */
export type SearchArea = 'body' | 'header' | 'footer' | 'textbox';

/** A part of the document the find panel searches: the body, one header/footer part, or one text box. */
export interface SearchPart {
  /** Stable while the document is open: 'body', the header/footer part's key, or a text box's (see DocxEditor.searchParts). */
  id: string;
  area: SearchArea;
  /** For people: 正文, 頁首, 頁尾（第一頁）－第 2 節 ... */
  label: string;
  /** Its content now. */
  doc: PMNode;
}

export const AREA_LABEL: Record<SearchArea, string> = { body: '正文', header: '頁首', footer: '頁尾', textbox: '文字方塊' };

/**
 * The index of the match to go to from `state`'s selection in direction `dir`, without
 * wrapping: -1 when there is none that way in this part (the next one is in another part,
 * or back at the other end).
 */
export function nextMatchIndex(state: EditorState, dir: 1 | -1): number {
  const s = searchState(state);
  const n = s.matches.length;
  if (!n) return -1;
  const { from, to } = state.selection;
  const cur = s.matches[s.current];
  let i: number;
  if (cur && cur.from === from && cur.to === to) i = s.current + dir;
  else if (dir > 0) i = firstAtOrAfter(s.matches, from);
  else {
    let lo = 0;
    let hi = n;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (s.matches[mid].to <= from) lo = mid + 1;
      else hi = mid;
    }
    i = lo - 1;
  }
  return i >= 0 && i < n ? i : -1;
}

/** Select match `index` (-1: the last one) of the current search and scroll to it. */
export const selectMatch = (index: number): Command => (state, dispatch) => {
  const matches = searchState(state).matches;
  const i = index < 0 ? matches.length - 1 : index;
  const m = matches[i];
  if (!m) return false;
  dispatch?.(goTo(state.tr, m, i));
  return true;
};

/**
 * A transaction replacing `matches` (sorted, found in `state.doc`) with `text`: one undo step,
 * not merged with the typing before it.
 */
export function replaceMatchesTr(state: EditorState, matches: SearchMatch[], text: string): Transaction {
  const tr = state.tr;
  replaceMatches(tr, matches, text);
  return closeHistory(tr);
}

/** Text shown around a match in the replace-all review list (see matchContext). */
export interface MatchContext {
  before: string;
  text: string;
  after: string;
  /**
   * The match may be the start of a longer term (GOV-ISSUE-010): right after it comes 之
   * (第十二條 in 第十二條之一), or the number or word it ends with goes on (2025 in 20251,
   * 第十 in 第十二, "cat" in "category"). Chinese has no spaces between words, so other
   * characters after a match (第十二條規定, 2025年) are not taken as part of it.
   */
  longer: boolean;
}

const HAN = /\p{Script=Han}/u;
const DIGIT = /\p{Nd}/u;
const CJK_NUMERAL = /[〇零一二三四五六七八九十百千萬兩壹貳參肆伍陸柒捌玖拾佰仟]/u;
const LETTER = /\p{L}/u;

/** The kind of character, for telling whether the text after a match continues the same number or word. */
function charKind(ch: string | undefined): 'digit' | 'numeral' | 'letter' | null {
  if (!ch) return null;
  if (DIGIT.test(ch)) return 'digit';
  if (CJK_NUMERAL.test(ch)) return 'numeral';
  if (LETTER.test(ch) && !HAN.test(ch)) return 'letter';
  return null;
}

/** Whether `next`, right after a match ending with `last`, makes it part of a longer term. */
function continuesTerm(last: string | undefined, next: string | undefined): boolean {
  if (!last || !next) return false;
  if (next === '之' && (HAN.test(last) || DIGIT.test(last))) return true;
  const kind = charKind(last);
  return kind != null && kind === charKind(next);
}

/** What a leaf shows as text: a field its result, a tab or line break a space, anything else nothing. */
function leafText(node: PMNode): string {
  if (node.type === schema.nodes.field) return (node.attrs.text as string) || '';
  if (node.type.name === 'tab' || node.type.name === 'hard_break') return ' ';
  return '';
}

/** Up to `chars` characters of the match's paragraph on each side of it, and whether it may be part of a longer term. */
export function matchContext(doc: PMNode, m: SearchMatch, chars = 12): MatchContext {
  const $from = doc.resolve(m.from);
  const start = $from.start();
  const end = $from.end();
  const beforeAll = doc.textBetween(start, m.from, '', leafText);
  const afterAll = doc.textBetween(m.to, end, '', leafText);
  const text = doc.textBetween(m.from, m.to, '', leafText);
  // Whole characters (a surrogate pair is one).
  const b = Array.from(beforeAll);
  const a = Array.from(afterAll);
  return {
    before: b.slice(Math.max(0, b.length - chars)).join(''),
    text,
    after: a.slice(0, chars).join(''),
    longer: continuesTerm(Array.from(text).pop(), a[0]),
  };
}
