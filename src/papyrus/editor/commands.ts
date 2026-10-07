import type { Command, EditorState, Transaction } from 'prosemirror-state';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { Mark } from 'prosemirror-model';
import type { MarkType, Node as PMNode } from 'prosemirror-model';
import { splitBlockAs, toggleMark } from 'prosemirror-commands';
import { CellSelection, addRowAfter, isInTable } from 'prosemirror-tables';
import { schema } from './schema';
import { ensureList, isListKind, restartedList, type ListKind } from '../docx/numbering';
import { hex, stripRunAppearance, withEastAsiaHint } from '../docx/props';
import { inheritedAlign, inheritedToggle, inheritedValue, type StyleInheritance } from '../docx/inheritance';
import type { Numbering } from '../docx/model';
import { NUMBERING_CHANGED } from './listMarkers';
import { afterParagraphSplit } from './review';
import { splitParagraphPPr } from '../docx/revisions';
import { newLayerId } from '../docx/wrappers';
import { closeHistory } from 'prosemirror-history';
import { NO_TRACK } from './trackChanges';
import { isSymbol, needsEastAsiaHint } from './symbols';

/** Paragraphs touched by the selection (including those inside table cells). */
export function selectedParagraphs(state: EditorState): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  const { ranges } = state.selection;
  for (const r of ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos) => {
      if (node.type === schema.nodes.paragraph) {
        out.push({ node, pos });
        return false;
      }
      return true;
    });
  }
  return out;
}

function updateParagraphs(fn: (attrs: Record<string, any>, node: PMNode) => Record<string, any> | null): Command {
  return (state, dispatch) => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) {
        const next = fn({ ...node.attrs }, node);
        if (next) tr.setNodeMarkup(pos, undefined, next, node.marks);
      }
      // Nothing to change: no transaction (no edit, nothing to undo).
      if (tr.docChanged) dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export const setParagraphStyle = (styleId: string | null): Command =>
  updateParagraphs((a) => ({ ...a, styleId }));

/** The alignment a paragraph shows: its own, else its style's. */
export const effectiveAlign = (attrs: Record<string, any>, styles?: StyleInheritance): string =>
  attrs.align ?? inheritedAlign(styles, attrs.styleId ?? null);

/**
 * Align the selected paragraphs. When the style already gives that alignment the paragraph
 * just inherits it; otherwise it is set explicitly, e.g. "left" inside a centered style.
 */
export const setAlign = (align: 'left' | 'center' | 'right' | 'justify', styles?: StyleInheritance): Command =>
  updateParagraphs((a) => {
    if (effectiveAlign(a, styles) === align) return null;
    const inherited = inheritedAlign(styles, a.styleId ?? null);
    return { ...a, align: align === inherited ? null : align };
  });

/**
 * The 段落 dialog's settings for the selected paragraphs: only the attributes given change
 * (indents in twips or hundredths of a character, spacing in twips, line and its rule).
 */
export const setParagraphFormat = (patch: Record<string, number | string | null>): Command =>
  updateParagraphs((a) => (Object.keys(patch).every((k) => (a[k] ?? null) === patch[k]) ? null : { ...a, ...patch }));

export const setLineSpacing = (multiple: number | null): Command =>
  updateParagraphs((a) => ({ ...a, line: multiple == null ? null : Math.round(multiple * 240), lineRule: multiple == null ? null : 'auto' }));

/** Toggle a bullet, numbered or 公文 (一、（一）1.（1）…) list on the selected paragraphs. */
export function toggleList(kind: ListKind, numbering: Numbering): Command {
  return (state, dispatch) => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    const isKind = (numId: string | null) => !!numId && isListKind(numbering.abstracts[numbering.nums[numId]?.abstractId ?? ''], kind);
    const allOn = paras.every((p) => isKind(p.node.attrs.numId));
    if (!dispatch) return true;
    // Keep the existing list id if the selection starts inside a list of this kind.
    const numId = allOn ? null : paras.find((p) => isKind(p.node.attrs.numId))?.node.attrs.numId ?? ensureList(numbering, kind);
    const tr = state.tr;
    for (const { node, pos } of paras) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, numId, ilvl: numId ? node.attrs.ilvl ?? 0 : 0 }, node.marks);
    }
    dispatch(tr.setMeta(NUMBERING_CHANGED, true));
    return true;
  };
}

/**
 * The list paragraph at the cursor and the ones after it in the same list instance: Word's
 * 重新從 1 開始編號 / 接續編號 move them all to another list instance.
 */
function listFromCursor(state: EditorState): { numId: string; ilvl: number; paras: { node: PMNode; pos: number }[] } | null {
  const { $from } = state.selection;
  const para = $from.parent;
  if (para.type !== schema.nodes.paragraph || !para.attrs.numId) return null;
  const numId = para.attrs.numId as string;
  const start = $from.before();
  const paras: { node: PMNode; pos: number }[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    if (pos >= start && node.attrs.numId === numId) paras.push({ node, pos });
    return false;
  });
  return { numId, ilvl: para.attrs.ilvl ?? 0, paras };
}

function moveToList(state: EditorState, paras: { node: PMNode; pos: number }[], numId: string): Transaction {
  const tr = state.tr;
  for (const { node, pos } of paras) tr.setNodeMarkup(pos, undefined, { ...node.attrs, numId }, node.marks);
  return tr.setMeta(NUMBERING_CHANGED, true);
}

/**
 * 重新從 1 開始編號 (Word's Restart at 1): the cursor's list paragraph starts its level's numbers
 * again, and the paragraphs after it in the same list count on from there. A numbered list only.
 */
export function restartNumbering(numbering: Numbering): Command {
  return (state, dispatch) => {
    const at = listFromCursor(state);
    const abs = at && numbering.abstracts[numbering.nums[at.numId]?.abstractId ?? ''];
    if (!at || !abs || isListKind(abs, 'bullet')) return false;
    if (!dispatch) return true;
    const numId = restartedList(numbering, at.numId, at.ilvl);
    if (!numId) return false;
    dispatch(moveToList(state, at.paras, numId));
    return true;
  };
}

/** The numbered list before the cursor's that 接續編號 would join (the same kind of numbers); null when none. */
function previousList(state: EditorState, numbering: Numbering): string | null {
  const at = listFromCursor(state);
  if (!at) return null;
  const fmtOf = (numId: string) => {
    const lvl = numbering.abstracts[numbering.nums[numId]?.abstractId ?? '']?.levels[0];
    return lvl ? `${lvl.fmt}|${lvl.text}` : null;
  };
  const mine = fmtOf(at.numId);
  if (!mine || mine.startsWith('bullet|')) return null;
  const start = state.selection.$from.before();
  let found: string | null = null;
  state.doc.nodesBetween(0, start, (node) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const id = node.attrs.numId as string | null;
    if (id && id !== at.numId && fmtOf(id) === mine) found = id;
    return false;
  });
  return found;
}

/** 接續編號 (Word's Continue Numbering): the cursor's list counts on from the same kind of list before it. */
export function continueNumbering(numbering: Numbering): Command {
  return (state, dispatch) => {
    const prev = previousList(state, numbering);
    const at = listFromCursor(state);
    if (!prev || !at) return false;
    if (dispatch) dispatch(moveToList(state, at.paras, prev));
    return true;
  };
}

const INDENT_STEP = 480; // twips (about two CJK characters at 12pt)

/** Indent: in lists change the level, otherwise shift the left indent. */
export const indent = (dir: 1 | -1): Command =>
  updateParagraphs((a) => {
    if (a.numId) return { ...a, ilvl: Math.max(0, Math.min(8, (a.ilvl ?? 0) + dir)) };
    // Paragraphs indented in characters (Word's "2 字元") step by two characters.
    if (a.indLeftChars != null) return { ...a, indLeftChars: Math.max(0, a.indLeftChars + dir * 200) };
    const left = Math.max(0, (a.indLeft ?? 0) + dir * INDENT_STEP);
    return { ...a, indLeft: left || null };
  });

export function inList(state: EditorState): boolean {
  const { $from } = state.selection;
  const p = $from.parent;
  return p.type === schema.nodes.paragraph && !!p.attrs.numId;
}

const NEXT_IS_NORMAL = /^(heading\d|title|subtitle)$/i;

/** Enter: continue lists and paragraph formatting; leave headings for normal text. */
export const splitParagraph: Command = (state, dispatch, view) => {
  const { $from, empty } = state.selection;
  const para = $from.parent;
  if (para.type !== schema.nodes.paragraph) return false;
  // Enter on an empty list item ends the list.
  if (empty && para.attrs.numId && para.content.size === 0) {
    if (dispatch) {
      dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, numId: null, ilvl: 0 }));
    }
    return true;
  }
  const at = $from.before();
  return splitBlockAs((node, atEnd) => {
    // The new paragraph copies the formatting (like Word) but not the identity of the
    // original w:p: its ids or page-break grouping. It ends with the original paragraph mark,
    // so it keeps the section break and the mark's revision (see afterParagraphSplit).
    const attrs: Record<string, any> = { ...node.attrs, pageBreakBefore: false, pAttrs: null, brGroup: null };
    attrs.pPr = splitParagraphPPr(attrs.pPr).last;
    if (atEnd && attrs.styleId && NEXT_IS_NORMAL.test(attrs.styleId)) attrs.styleId = null;
    return { type: schema.nodes.paragraph, attrs };
  })(state, dispatch && ((tr) => dispatch(afterParagraphSplit(tr, tr.mapping.map(at, -1)))), view);
};

export const insertHardBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create()).scrollIntoView());
  return true;
};

/**
 * Tab in a table's last cell adds a row below and moves into its first cell, as in Word
 * (in any other cell, goToNextCell moves on; outside tables Tab inserts a tab).
 */
export const tabNewRow: Command = (state, dispatch) => {
  if (!isInTable(state)) return false;
  if (!dispatch) return true;
  let tr: Transaction | null = null;
  addRowAfter(state, (t) => (tr = t));
  if (!tr) return false;
  const added = tr as Transaction;
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type.name !== 'table_row') continue;
    // The new row starts where the current row ended.
    const start = added.mapping.map($from.after(d), -1);
    added.setSelection(TextSelection.near(added.doc.resolve(start + 1)));
    break;
  }
  dispatch(added.scrollIntoView());
  return true;
};

export const insertTab: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.tab.create()).scrollIntoView());
  return true;
};

/**
 * A column break (Word's 分欄符號, w:br w:type="column"): in Word the text after it starts the next
 * column of a section with columns (the next page when the section has one column). The editor
 * shows one column, so it is only a marked line here and doesn't start a new page. Only in a
 * body paragraph (not in a table cell).
 */
export const insertColumnBreak: Command = (state, dispatch) => {
  const { $from } = state.selection;
  if ($from.depth !== 1 || $from.parent.type !== schema.nodes.paragraph) return false;
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create({ type: 'column' })).scrollIntoView());
  return true;
};

export const insertPageBreak: Command = (state, dispatch) => {
  if (!dispatch) return true;
  const tr = state.tr.deleteSelection();
  const pos = tr.selection.from;
  const $pos = tr.doc.resolve(pos);
  if ($pos.depth === 1 && $pos.parent.type === schema.nodes.paragraph) {
    // Like Word, the break goes inside the paragraph (w:br w:type="page" in the same w:p): the
    // two halves and the break are pieces of one w:p, so its mark (section break, revision)
    // is written once, at the end of the last piece.
    const para = $pos.parent;
    const group = (para.attrs.brGroup as string | null) ?? newLayerId();
    const start = $pos.before();
    tr.split(pos);
    tr.setNodeMarkup(start, undefined, { ...para.attrs, brGroup: group, sectPr: null });
    tr.setNodeMarkup(pos + 1, undefined, { ...para.attrs, brGroup: group });
    tr.insert(pos + 1, schema.nodes.page_break.create({ group }));
    tr.setSelection(TextSelection.create(tr.doc, pos + 3));
  } else {
    tr.replaceSelectionWith(schema.nodes.page_break.create());
  }
  dispatch(tr.scrollIntoView());
  return true;
};

/**
 * Word's 插入 › 空白頁: two page breaks at the cursor, with the cursor left on the new blank page
 * between them (the text after the cursor moves to the page after it). One undo step.
 */
export const insertBlankPage: Command = (state, dispatch) => {
  if (!dispatch) return true;
  let first: Transaction | null = null;
  insertPageBreak(state, (t) => (first = t));
  if (!first) return false;
  const tr1 = first as Transaction;
  const after = state.apply(tr1);
  // The cursor sits right after the first break: the second break goes there too.
  const between = tr1.selection.from;
  let second: Transaction | null = null;
  insertPageBreak(after, (t) => (second = t));
  if (!second) return false;
  const tr2 = second as Transaction;
  const tr = state.tr;
  for (const step of tr1.steps) tr.step(step);
  for (const step of tr2.steps) tr.step(step);
  // Back between the two breaks: the empty piece that is the blank page.
  const pos = tr2.mapping.map(between, -1);
  tr.setSelection(TextSelection.near(tr.doc.resolve(pos), -1));
  dispatch(tr.scrollIntoView());
  return true;
};

/** Apply a mark with attributes (color, font ...) or remove it when attrs is null. */
export function setMark(type: MarkType, attrs: Record<string, unknown> | null): Command {
  return (state, dispatch) => {
    const { from, to, empty, $from } = state.selection;
    if (!dispatch) return true;
    const tr = state.tr;
    if (empty) {
      const marks = type.removeFromSet(state.storedMarks ?? $from.marks());
      tr.setStoredMarks(attrs ? type.create(attrs).addToSet(marks) : marks);
    } else {
      tr.removeMark(from, to, type);
      if (attrs) tr.addMark(from, to, type.create(attrs));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/**
 * Word's 字型 box (persona-300 B-3): an East Asian font (標楷體, 新細明體 …) is set for Chinese and
 * Latin text alike (w:eastAsia, w:ascii, w:hAnsi); a Latin font (Times New Roman, Arial …) for
 * Latin text only (w:ascii, w:hAnsi): each run keeps its own Chinese font. Null: the styles' fonts.
 */
export function setFont(name: string | null, eastAsian: boolean): Command {
  if (!name || eastAsian) return setMark(schema.marks.font, name ? { family: name, eastAsia: name } : null);
  const type = schema.marks.font;
  return (state, dispatch) => {
    if (!dispatch) return true;
    const tr = state.tr;
    if (state.selection.empty) {
      const marks = state.storedMarks ?? state.selection.$from.marks();
      const eastAsia = type.isInSet(marks)?.attrs.eastAsia ?? null;
      tr.setStoredMarks(type.create({ family: name, eastAsia }).addToSet(type.removeFromSet(marks)));
    } else {
      for (const r of state.selection.ranges) {
        const from = r.$from.pos;
        const to = r.$to.pos;
        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!node.isInline) return true;
          const eastAsia = type.isInSet(node.marks)?.attrs.eastAsia ?? null;
          tr.addMark(Math.max(from, pos), Math.min(to, pos + node.nodeSize), type.create({ family: name, eastAsia }));
          return false;
        });
      }
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export const toggle = (name: string): Command => toggleMark(schema.marks[name]);

const charStyleOf = (marks: readonly Mark[]): string | null => schema.marks.charStyle.isInSet(marks)?.attrs.id ?? null;

/** Whether text with these marks, in this paragraph, shows bold / italic (its own setting, else the styles'). */
export function effectiveToggle(name: 'bold' | 'italic', marks: readonly Mark[], para: PMNode, styles?: StyleInheritance): boolean {
  const own = schema.marks[name].isInSet(marks);
  if (own) return own.attrs.on !== false;
  return inheritedToggle(styles, name, para.attrs.styleId ?? null, charStyleOf(marks));
}

interface TextSpan {
  from: number;
  to: number;
  marks: readonly Mark[];
  para: PMNode;
}

function selectedText(state: EditorState): TextSpan[] {
  const out: TextSpan[] = [];
  for (const r of state.selection.ranges) {
    const from = r.$from.pos;
    const to = r.$to.pos;
    state.doc.nodesBetween(from, to, (node, pos, parent) => {
      if (!node.isInline) return true;
      if (node.isText && parent) out.push({ from: Math.max(from, pos), to: Math.min(to, pos + node.nodeSize), marks: node.marks, para: parent });
      return false;
    });
  }
  return out;
}

/** Whether the whole selection (or the cursor) shows bold / italic, styles included. */
export function selectionHas(state: EditorState, name: 'bold' | 'italic', styles?: StyleInheritance): boolean {
  const { empty, $from } = state.selection;
  if (empty) return effectiveToggle(name, state.storedMarks ?? $from.marks(), $from.parent, styles);
  const spans = selectedText(state);
  return spans.length > 0 && spans.every((s) => effectiveToggle(name, s.marks, s.para, styles));
}

/** A run value the toolbar shows: font size (pt), text colour ("#rrggbb"), font name. */
export type RunValue = 'fontSize' | 'color' | 'font';

/** The value text with these marks shows: its own formatting, else character style, paragraph style, defaults. */
export function effectiveValue(name: RunValue, marks: readonly Mark[], para: PMNode, styles?: StyleInheritance): string | number | null {
  const ps = para.attrs.styleId ?? null;
  const cs = charStyleOf(marks);
  if (name === 'fontSize') {
    return schema.marks.fontSize.isInSet(marks)?.attrs.pt ?? inheritedValue(styles, 'fontSize', ps, cs);
  }
  if (name === 'color') {
    const own = schema.marks.color.isInSet(marks)?.attrs.color as string | undefined;
    const color = own ?? inheritedValue(styles, 'color', ps, cs);
    return color ? '#' + hex(color).toLowerCase() : null;
  }
  // The toolbar names the CJK font when there is one, as Word's Chinese UI does.
  const font = schema.marks.font.isInSet(marks);
  return (
    font?.attrs.eastAsia ?? inheritedValue(styles, 'fontEastAsia', ps, cs) ??
    font?.attrs.family ?? inheritedValue(styles, 'fontFamily', ps, cs)
  );
}

/** Word's font size list, for Ctrl+Shift+> / < (one step up or down it). */
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** The size after one Ctrl+Shift+> (1) or < (-1): the next size of the list; past 72, by 10 pt. */
export function stepFontSize(pt: number, dir: 1 | -1): number {
  if (dir > 0) return FONT_SIZES.find((s) => s > pt) ?? Math.min(1638, Math.floor(pt / 10) * 10 + 10);
  const below = [...FONT_SIZES].reverse().find((s) => s < pt);
  return pt > 72 ? Math.max(72, Math.ceil(pt / 10) * 10 - 10) : below ?? 1;
}

/**
 * Grow or shrink the font of the selection (each run from its own size), as Word's Ctrl+] / Ctrl+[
 * (`by` 1 or -1 pt) and Ctrl+Shift+> / < (`by` 'step': along the size list). At the cursor, the
 * text typed next. Sizes stay within Word's 1–1638 pt; text without a size counts as 10 pt.
 */
export function growFont(by: 1 | -1 | 'up' | 'down', styles?: StyleInheritance): Command {
  const type = schema.marks.fontSize;
  const next = (pt: number) => {
    const n = by === 'up' ? stepFontSize(pt, 1) : by === 'down' ? stepFontSize(pt, -1) : pt + by;
    return Math.min(1638, Math.max(1, Math.round(n * 2) / 2));
  };
  const size = (marks: readonly Mark[], para: PMNode) => Number(effectiveValue('fontSize', marks, para, styles) ?? 10) || 10;
  return (state, dispatch) => {
    const { empty, $from } = state.selection;
    if (!dispatch) return true;
    const tr = state.tr;
    if (empty) {
      const marks = state.storedMarks ?? $from.marks();
      tr.setStoredMarks(type.create({ pt: next(size(marks, $from.parent)) }).addToSet(type.removeFromSet(marks)));
    } else {
      for (const s of selectedText(state)) tr.addMark(s.from, s.to, type.create({ pt: next(size(s.marks, s.para)) }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/**
 * The value shown for the selection: `mixed` when the selected text has more than one
 * (the toolbar then shows no single value).
 */
export function selectionValue(state: EditorState, name: RunValue, styles?: StyleInheritance): { value: string | number | null; mixed: boolean } {
  const { empty, $from } = state.selection;
  if (empty) return { value: effectiveValue(name, state.storedMarks ?? $from.marks(), $from.parent, styles), mixed: false };
  const values = new Set(selectedText(state).map((s) => effectiveValue(name, s.marks, s.para, styles)));
  if (values.size > 1) return { value: null, mixed: true };
  return { value: values.size ? [...values][0] : effectiveValue(name, $from.marks(), $from.parent, styles), mixed: false };
}

/**
 * Bold / italic that also works against styles: turning it off where a style makes text
 * bold writes an explicit "not bold"; where the styles already give the wanted result the
 * text simply inherits it.
 */
export function toggleFormat(name: 'bold' | 'italic', styles?: StyleInheritance): Command {
  return (state, dispatch) => {
    const type = schema.marks[name];
    const { empty, $from } = state.selection;
    const want = !selectionHas(state, name, styles);
    if (empty) {
      if (!dispatch) return true;
      const marks = state.storedMarks ?? $from.marks();
      const inherited = inheritedToggle(styles, name, $from.parent.attrs.styleId ?? null, charStyleOf(marks));
      let next = type.removeFromSet(marks);
      if (want !== inherited) next = type.create({ on: want }).addToSet(next);
      dispatch(state.tr.setStoredMarks(next));
      return true;
    }
    const spans = selectedText(state);
    if (!spans.length) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    for (const s of spans) {
      const inherited = inheritedToggle(styles, name, s.para.attrs.styleId ?? null, charStyleOf(s.marks));
      const own = type.isInSet(s.marks);
      const target = want === inherited ? null : want;
      if ((own ? own.attrs.on !== false : null) === target) continue;
      tr.removeMark(s.from, s.to, type);
      if (target != null) tr.addMark(s.from, s.to, type.create({ on: target }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Marks that are only about how text looks; "clear formatting" removes these and 複製格式 copies them. */
export const APPEARANCE_MARKS = [
  'charStyle', 'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'color', 'highlight', 'font', 'fontSize',
];

/** The run mark with its appearance properties taken out (null when nothing is left to keep). */
function plainRun(run: Mark): Mark | null {
  const rPr = run.attrs.rPr ? stripRunAppearance(run.attrs.rPr) : null;
  if (rPr === run.attrs.rPr) return run;
  return rPr || run.attrs.attrs ? schema.marks.run.create({ ...run.attrs, rPr }) : null;
}

/**
 * Remove how text looks (fonts, sizes, bold, colours, character styles) and keep what it
 * means: tracked changes and content controls (inlineWrap), field results, links, and run
 * properties such as language or hidden text.
 */
export const clearFormatting: Command = (state, dispatch) => {
  const { from, to, empty, $from } = state.selection;
  if (!dispatch) return true;
  const tr = state.tr;
  const appearance = new Set(APPEARANCE_MARKS.map((n) => schema.marks[n]));
  if (empty) {
    let marks: readonly Mark[] = (state.storedMarks ?? $from.marks()).filter((m) => !appearance.has(m.type));
    const run = schema.marks.run.isInSet(marks);
    if (run) {
      marks = run.removeFromSet(marks);
      const plain = plainRun(run);
      if (plain) marks = plain.addToSet(marks);
    }
    tr.setStoredMarks(marks);
  } else {
    for (const type of appearance) tr.removeMark(from, to, type);
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isInline) return true;
      const run = schema.marks.run.isInSet(node.marks);
      const plain = run && plainRun(run);
      if (!run || plain === run) return false;
      const start = Math.max(from, pos);
      const end = Math.min(to, pos + node.nodeSize);
      tr.removeMark(start, end, run);
      if (plain) tr.addMark(start, end, plain);
      return false;
    });
  }
  dispatch(tr);
  return true;
};

/**
 * Sets (or with null removes) the link on the selection. The target is not checked here: the
 * link dialog refuses unsafe targets, the page renders them as "#" and the DOCX writer drops
 * them (docx/links.ts).
 */
export function setLink(href: string | null): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const type = schema.marks.link;
    if (!dispatch) return true;
    const tr = state.tr;
    if (empty) {
      // Extend to the whole existing link around the cursor.
      const range = markRange(state, from, type);
      if (!range && href) {
        tr.insertText(href, from).addMark(from, from + href.length, type.create({ href }));
      } else if (range) {
        tr.removeMark(range.from, range.to, type);
        if (href) tr.addMark(range.from, range.to, type.create({ href }));
      }
    } else {
      tr.removeMark(from, to, type);
      if (href) tr.addMark(from, to, type.create({ href }));
    }
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** The extent of the mark of `type` around `pos` inside its textblock. */
export function markRange(state: EditorState, pos: number, type: MarkType): { from: number; to: number } | null {
  const $pos = state.doc.resolve(pos);
  const base = $pos.start();
  const kids: { node: PMNode; off: number }[] = [];
  $pos.parent.forEach((node, off) => kids.push({ node, off }));
  const at = $pos.parentOffset;
  const i = kids.findIndex((k) => at >= k.off && at <= k.off + k.node.nodeSize && type.isInSet(k.node.marks));
  if (i < 0) return null;
  const mark = type.isInSet(kids[i].node.marks)!;
  let a = i;
  let b = i;
  while (a > 0 && mark.isInSet(kids[a - 1].node.marks)) a--;
  while (b < kids.length - 1 && mark.isInSet(kids[b + 1].node.marks)) b++;
  return { from: base + kids[a].off, to: base + kids[b].off + kids[b].node.nodeSize };
}

/** An East Asian font or language on the run itself (w:rFonts/@w:eastAsia(Theme), w:lang/@w:eastAsia). */
const OWN_EAST_ASIAN = /<w:(?:rFonts\b[^>]*\sw:eastAsia(?:Theme)?|lang\b[^>]*\sw:eastAsia)="[^"]+"/;

/**
 * Whether text with these marks, in this paragraph, is set up for East Asian text: an East Asian
 * font or language of its own, or an East Asian font from its character or paragraph style or
 * the document defaults. (A language set only in a style is not looked at.)
 */
function eastAsianText(marks: readonly Mark[], para: PMNode, styles?: StyleInheritance): boolean {
  if (schema.marks.font.isInSet(marks)?.attrs.eastAsia) return true;
  if (OWN_EAST_ASIAN.test((schema.marks.run.isInSet(marks)?.attrs.rPr as string | null) ?? '')) return true;
  return !!inheritedValue(styles, 'fontEastAsia', para.attrs.styleId ?? null, charStyleOf(marks));
}

/**
 * `marks` for `ch` in `para`: with w:rFonts/@w:hint="eastAsia" in the run mark when the symbol is
 * one of those the hint moves to the East Asian font and the text has an East Asian font or
 * language (as Word writes the hint for symbols entered in Chinese text).
 */
function hinted(ch: string, marks: readonly Mark[], para: PMNode, styles?: StyleInheritance): readonly Mark[] {
  if (!needsEastAsiaHint(ch) || !eastAsianText(marks, para, styles)) return marks;
  const run = schema.marks.run.isInSet(marks);
  const rPr = withEastAsiaHint((run?.attrs.rPr as string | null) ?? null);
  return rPr === run?.attrs.rPr ? marks : schema.marks.run.create({ rPr, attrs: run?.attrs.attrs ?? null }).addToSet(marks);
}

/**
 * Word's 插入 › 符號: put one character in at the selection as text, as typing it would: it
 * replaces the selection (a picture, table cells) and takes the formatting of the text at the
 * cursor, the stored marks included, so 追蹤修訂 records it as an ordinary insertion. A symbol
 * Latin and CJK fonts share (①, ※, ±, →…), entered in text with an East Asian font or language
 * (its own or its styles', `styles`), gets w:rFonts/@w:hint="eastAsia" in its own run, as Word
 * writes it, so Word draws it in the East Asian font (標楷體) like the text around it; the text
 * typed after it doesn't take the hint. It is its own undo step, before and after: Word undoes
 * 插入符號 separately from typing. False for anything that is not one character that can be
 * inserted (see isSymbol).
 */
export const insertSymbol = (ch: string, styles?: StyleInheritance): Command => (state, dispatch, view) => {
  if (!isSymbol(ch)) return false;
  if (!dispatch) return true;
  const sel = state.selection;
  // The marks typed text gets (as Transaction.replaceSelectionWith gives them).
  const marks = state.storedMarks ?? (sel.empty ? sel.$from.marks() : sel.$from.marksAcross(sel.$to) ?? Mark.none);
  const tr = state.tr;
  /** The marks text typed after the symbol gets: those around it, without the symbol's hint. */
  let after = marks;
  if (sel instanceof CellSelection) {
    // Table cells selected: as when typing over them in Word, their text goes and the symbol is
    // written in the first one (a text deletion per cell, so 追蹤修訂 can record it).
    const cells: { node: PMNode; pos: number }[] = [];
    sel.forEachCell((node, pos) => {
      cells.push({ node, pos });
    });
    const at = cells[0].pos + 2; // in the first cell's first paragraph
    if (!state.doc.resolve(at).parent.inlineContent) return false;
    for (const { node, pos } of [...cells].reverse()) {
      const from = pos + 2;
      const to = pos + node.nodeSize - 2;
      if (to > from && tr.doc.resolve(from).parent.inlineContent && tr.doc.resolve(to).parent.inlineContent) tr.delete(from, to);
    }
    const $at = state.doc.resolve(at);
    after = $at.marks();
    tr.insert(at, schema.text(ch, hinted(ch, after, $at.parent, styles)));
    tr.setSelection(TextSelection.create(tr.doc, at + ch.length));
  } else {
    tr.replaceSelectionWith(schema.text(ch, hinted(ch, marks, sel.$from.parent, styles)), false);
  }
  if (needsEastAsiaHint(ch)) tr.setStoredMarks(after);
  dispatch(closeHistory(tr).scrollIntoView());
  // Typing right after it starts a new undo step too.
  if (view) view.dispatch(closeHistory(view.state.tr).setMeta('addToHistory', false));
  return true;
};

export function insertImage(src: string, width: number, height: number, maxWidth: number): Command {
  return (state, dispatch) => {
    const scale = width > maxWidth ? maxWidth / width : 1;
    const node = schema.nodes.image.create({ src, width: Math.round(width * scale), height: Math.round(height * scale) });
    if (dispatch) dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    return true;
  };
}

export function insertTable(rows: number, cols: number, contentWidthPx: number): Command {
  return (state, dispatch) => {
    const w = Math.floor(contentWidthPx / cols);
    const cell = () => schema.nodes.table_cell.create({ colwidth: [w] }, schema.nodes.paragraph.create());
    const row = () => schema.nodes.table_row.create(null, Array.from({ length: cols }, cell));
    const table = schema.nodes.table.create({ styleId: null }, Array.from({ length: rows }, row));
    if (!dispatch) return true;
    const tr = state.tr.replaceSelectionWith(table);
    // Put the cursor in the first cell (as Word does). The table may go before the paragraph the
    // cursor was in (cursor at its start), so it is looked up rather than worked out.
    const start = tr.mapping.map(state.selection.from, -1);
    let at = -1;
    tr.doc.nodesBetween(Math.max(0, start - table.nodeSize - 2), Math.min(tr.doc.content.size, start + table.nodeSize + 2), (node, pos) => {
      if (at < 0 && node === table) at = pos;
      return at < 0 && !node.isTextblock;
    });
    const $in = tr.doc.resolve(at >= 0 ? at + 3 : Math.min(tr.doc.content.size, start + 3));
    tr.setSelection(TextSelection.near($in));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

export function selectedImage(state: EditorState): PMNode | null {
  const sel = state.selection;
  return sel instanceof NodeSelection && sel.node.type === schema.nodes.image ? sel.node : null;
}

/**
 * Set the alt text (Word's 替代文字, saved as wp:docPr/@descr) of the picture at `pos`; a
 * picture that was selected stays selected. The alt box calls this on every keystroke with
 * the same `group` time and `first` only for the first one: the history then keeps the
 * whole typing as one undo step (closed before it, merged after). Word does not record alt
 * text as a tracked change, so it is applied as it is while 追蹤修訂 is on. GOV-ISSUE-011.
 */
export function setImageAlt(pos: number, alt: string, group?: { time: number; first: boolean }): Command {
  return (state, dispatch) => {
    const node = state.doc.nodeAt(pos);
    if (!node || node.type !== schema.nodes.image) return false;
    if ((node.attrs.alt ?? '') === alt) return true;
    if (dispatch) {
      const wasSelected = state.selection instanceof NodeSelection && state.selection.from === pos;
      // A leaf: this replaces the node, a step the history can merge with the next keystroke's.
      const tr = state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, alt }, node.marks).setMeta(NO_TRACK, true);
      if (wasSelected) tr.setSelection(NodeSelection.create(tr.doc, pos));
      if (group) tr.setTime(group.time);
      dispatch(!group || group.first ? closeHistory(tr) : tr);
    }
    return true;
  };
}

export type { Transaction };
