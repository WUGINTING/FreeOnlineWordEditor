// Table appearance: shading, borders, vertical alignment, row height, header rows,
// "don't split row", equal column widths. Each command is one transaction (one undo step)
// and only sets the attributes it is about; the writer patches w:tcPr / w:trPr from them.

import { NodeSelection, TextSelection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import {
  CellSelection, TableMap, cellAround, isInTable, selectedRect, selectionCell, type Rect, type TableRect,
} from 'prosemirror-tables';
import { cellInfo, hex, parseKept, type Border, type BorderSide, type CellBorders } from '../docx/props';
import { keptTcPr } from '../docx/tableLayout';
import { Converter } from '../docx/convert';
import { parseLayers } from '../docx/wrappers';
import { schema } from './schema';

export type BorderMode = 'all' | 'outer' | 'inner' | 'none';
/** w:trHeight/@w:hRule; "auto" (height ignored) only keeps what the file says. */
export type HeightRule = 'atLeast' | 'exact' | 'auto';

interface CellRef {
  pos: number;
  node: PMNode;
  rect: Rect;
}

/** A single ½ pt line, as Word's border buttons write it. */
const SINGLE: Border = { val: 'single', sz: 4, color: 'auto' };
const NONE: Border = { val: 'nil', sz: null, color: null };

const OPPOSITE: Record<BorderSide, BorderSide> = { top: 'bottom', bottom: 'top', left: 'right', right: 'left' };

function selection(state: EditorState): TableRect | null {
  return isInTable(state) ? selectedRect(state) : null;
}

/**
 * The cells in `rect`. Cells added only to square up ragged rows (w:gridBefore / w:gridAfter)
 * are left out: they are not saved, so formatting them would be lost.
 */
function cellsIn(sel: TableRect, rect: Rect): CellRef[] {
  return sel.map
    .cellsInRect(rect)
    .map((p) => ({ pos: sel.tableStart + p, node: sel.table.nodeAt(p)!, rect: sel.map.findCell(p) }))
    .filter((c) => !c.node.attrs.filler);
}

function rowsOf(sel: TableRect): { pos: number; node: PMNode }[] {
  const out: { pos: number; node: PMNode }[] = [];
  let pos = sel.tableStart;
  sel.table.forEach((row) => {
    out.push({ pos, node: row });
    pos += row.nodeSize;
  });
  return out;
}

function setAttrs(tr: Transaction, pos: number, node: PMNode, attrs: Record<string, unknown>): boolean {
  const next = { ...node.attrs, ...attrs };
  if (Object.keys(attrs).every((k) => JSON.stringify(node.attrs[k]) === JSON.stringify(next[k]))) return false;
  tr.setNodeMarkup(pos, undefined, next, node.marks);
  return true;
}

/** Apply `fn` to every selected cell (or the cell with the cursor). */
function updateCells(fn: (cell: CellRef, sel: TableRect) => Record<string, unknown> | null): Command {
  return (state, dispatch) => {
    const sel = selection(state);
    if (!sel) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const cell of cellsIn(sel, sel)) {
        const attrs = fn(cell, sel);
        if (attrs) setAttrs(tr, cell.pos, cell.node, attrs);
      }
      if (tr.docChanged) dispatch(tr);
    }
    return true;
  };
}

/** Cell shading (w:shd fill); null removes it. */
export const setCellBackground = (color: string | null): Command =>
  updateCells(() => ({ background: color ? '#' + hex(color) : null }));

/** Vertical alignment of the cell content (w:vAlign). */
export const setCellVAlign = (vAlign: 'top' | 'middle' | 'bottom'): Command => updateCells(() => ({ vAlign }));

/**
 * Borders of the selected cells, like Word's border menu: all, outside, inside or none.
 * Edges shared with cells outside the selection are set on those cells too, so both
 * sides of an edge agree (Word does the same).
 */
export function setCellBorders(mode: BorderMode): Command {
  return (state, dispatch) => {
    const sel = selection(state);
    if (!sel) return false;
    if (!dispatch) return true;
    // Wanted sides per cell position, merged over all edges that touch the cell.
    const wanted = new Map<number, { node: PMNode; sides: CellBorders }>();
    const want = (cell: CellRef, side: BorderSide, b: Border) => {
      const entry = wanted.get(cell.pos) ?? { node: cell.node, sides: {} };
      entry.sides[side] = b;
      wanted.set(cell.pos, entry);
    };
    const { map } = sel;
    for (const cell of cellsIn(sel, sel)) {
      const r = cell.rect;
      const outer: Record<BorderSide, boolean> = {
        top: r.top <= sel.top, bottom: r.bottom >= sel.bottom, left: r.left <= sel.left, right: r.right >= sel.right,
      };
      for (const side of ['top', 'left', 'bottom', 'right'] as BorderSide[]) {
        const b = mode === 'none' ? NONE
          : mode === 'all' ? SINGLE
          : mode === 'outer' ? (outer[side] ? SINGLE : null)
          : !outer[side] ? SINGLE : null;
        if (!b) continue;
        want(cell, side, b);
        if (!outer[side]) continue;
        // The cells on the other side of a selection edge.
        const beyond: Rect | null =
          side === 'top' ? (r.top > 0 ? { left: r.left, right: r.right, top: r.top - 1, bottom: r.top } : null)
          : side === 'bottom' ? (r.bottom < map.height ? { left: r.left, right: r.right, top: r.bottom, bottom: r.bottom + 1 } : null)
          : side === 'left' ? (r.left > 0 ? { left: r.left - 1, right: r.left, top: r.top, bottom: r.bottom } : null)
          : r.right < map.width ? { left: r.right, right: r.right + 1, top: r.top, bottom: r.bottom } : null;
        if (beyond) for (const n of cellsIn(sel, beyond)) want(n, OPPOSITE[side], b);
      }
    }
    const tr = state.tr;
    for (const [pos, { node, sides }] of wanted) {
      setAttrs(tr, pos, node, { borders: { ...(node.attrs.borders ?? {}), ...sides } });
    }
    if (tr.docChanged) dispatch(tr);
    return true;
  };
}

/** Apply `fn` to the rows touched by the selection. */
function updateRows(fn: (row: PMNode, index: number, sel: TableRect) => Record<string, unknown> | null, which?: (sel: TableRect, i: number) => boolean): Command {
  return (state, dispatch) => {
    const sel = selection(state);
    if (!sel) return false;
    if (dispatch) {
      const tr = state.tr;
      rowsOf(sel).forEach(({ pos, node }, i) => {
        if (!(which ? which(sel, i) : i >= sel.top && i < sel.bottom)) return;
        const attrs = fn(node, i, sel);
        if (attrs) setAttrs(tr, pos, node, attrs);
      });
      if (tr.docChanged) dispatch(tr);
    }
    return true;
  };
}

/** Row height (w:trHeight) in twips; null lets the row fit its content. */
export const setRowHeight = (height: number | null, rule: HeightRule = 'atLeast'): Command =>
  updateRows((row) => {
    if (height == null) return { height: null, heightRule: null };
    // An omitted rule already means "at least"; keep the file as it was written.
    const heightRule = rule === 'atLeast' && row.attrs.heightRule == null ? null : rule;
    return { height: Math.max(1, Math.round(height)), heightRule };
  });

/** Only the rule of the selected rows that have a height; the heights stay as they are. */
export const setRowHeightRule = (rule: HeightRule): Command =>
  updateRows((row) => {
    if (row.attrs.height == null) return null;
    return { heightRule: rule === 'atLeast' && row.attrs.heightRule == null ? null : rule };
  });

/**
 * "Repeat as header row" (w:tblHeader). Word only repeats rows from the top of the table,
 * so turning it on includes every row above the selection and turning it off every row below.
 */
export const setHeaderRow = (on: boolean): Command =>
  updateRows(() => ({ header: on }), (sel, i) => (on ? i < sel.bottom : i >= sel.top));

/** "Allow row to break across pages" off (w:cantSplit). */
export const setCantSplit = (on: boolean): Command => updateRows(() => ({ cantSplit: on }));

/** The current width of each column of the table, px (null where neither the cells nor the grid say). */
function columnWidths(sel: TableRect): (number | null)[] {
  const { map, table } = sel;
  const widths: (number | null)[] = new Array(map.width).fill(null);
  for (const p of new Set(map.map)) {
    const cell = table.nodeAt(p)!;
    const rect = map.findCell(p);
    const cw = cell.attrs.colwidth as number[] | null;
    for (let c = rect.left; c < rect.right; c++) if (widths[c] == null && cw?.[c - rect.left]) widths[c] = cw[c - rect.left];
  }
  const grid = table.attrs.grid as number[] | null;
  for (let c = 0; c < map.width; c++) widths[c] ??= grid?.[c] ? grid[c] / 15 : null;
  return widths;
}

/**
 * 欄寬 (Word's 表格內容 › 欄 › 慣用寬度): the selected columns (or the cursor's) get `px` wide,
 * the other columns keep theirs (the table gets wider or narrower, as in Word). One undo step.
 */
export function setColumnWidth(px: number): Command {
  return (state, dispatch) => {
    const sel = selection(state);
    if (!sel || !(px > 0)) return false;
    const { map, table } = sel;
    // The columns of a cell selection, else the cursor's cell's.
    const [from, to] = [sel.left, sel.right];
    const widths = columnWidths(sel);
    if (!dispatch) return true;
    const tr = state.tr;
    for (const p of new Set(map.map)) {
      const cell = table.nodeAt(p)!;
      const rect = map.findCell(p);
      if (rect.right <= from || rect.left >= to || cell.attrs.filler) continue;
      const colwidth: number[] = [];
      for (let c = rect.left; c < rect.right; c++) colwidth.push(c >= from && c < to ? px : (widths[c] ?? px));
      setAttrs(tr, sel.tableStart + p, cell, { colwidth });
    }
    if (tr.docChanged) dispatch(tr);
    return true;
  };
}

/** Give the selected columns (or all columns, with just a cursor) the same width, keeping their total. */
export const distributeColumns: Command = (state, dispatch) => {
  const sel = selection(state);
  if (!sel) return false;
  const { map, table } = sel;
  const [from, to] = state.selection instanceof CellSelection && sel.right - sel.left > 1 ? [sel.left, sel.right] : [0, map.width];
  if (to - from < 2) return false;
  const widths = columnWidths(sel);
  const known = widths.slice(from, to).filter((w): w is number => w != null);
  if (!known.length) return false;
  const avg = known.reduce((s, w) => s + w, 0) / known.length;
  const total = widths.slice(from, to).reduce<number>((s, w) => s + (w ?? avg), 0);
  // Exact in twips (1 px = 15 twips), so the saved grid adds up.
  const each = Math.round((total / (to - from)) * 15) / 15;
  if (!dispatch) return true;
  const tr = state.tr;
  for (const p of new Set(map.map)) {
    const cell = table.nodeAt(p)!;
    const rect = map.findCell(p);
    if (rect.right <= from || rect.left >= to) continue;
    const colwidth: number[] = [];
    for (let c = rect.left; c < rect.right; c++) colwidth.push(c >= from && c < to ? each : Math.round(widths[c] ?? avg));
    setAttrs(tr, sel.tableStart + p, cell, { colwidth });
  }
  if (tr.docChanged) dispatch(tr);
  return true;
};

/** What the table panel shows for the cell with the cursor (the selection's anchor cell). */
export interface TableInfo {
  /** Width of the cursor's (or the selected) column(s), px; null when unknown or they differ. */
  colWidth: number | null;
  background: string | null;
  vAlign: string | null;
  height: number | null; // twips
  heightRule: string | null;
  header: boolean;
  cantSplit: boolean;
}

export function tableInfo(state: EditorState): TableInfo | null {
  if (!isInTable(state)) return null;
  const $cell = selectionCell(state);
  const cell = $cell.nodeAfter;
  const row = $cell.parent;
  if (!cell) return null;
  const sel = selection(state);
  const cols = sel ? columnWidths(sel).slice(sel.left, sel.right) : [];
  return {
    colWidth: cols.length && cols.every((w) => w != null && Math.abs(w - cols[0]!) < 0.01) ? cols[0] : null,
    background: cell.attrs.background ?? null,
    vAlign: cell.attrs.vAlign ?? null,
    height: row.attrs.height ?? null,
    heightRule: row.attrs.heightRule ?? null,
    header: !!row.attrs.header,
    cantSplit: !!row.attrs.cantSplit,
  };
}

// ----- 分割表格 (Split Table) -----

/**
 * The innermost table at the selection (or the table selected as a whole), where it is, and
 * the row the selection starts in.
 */
export function tableAt(state: EditorState): { table: PMNode; pos: number; row: number } | null {
  const sel = state.selection;
  if (sel instanceof NodeSelection && sel.node.type.spec.tableRole === 'table') return { table: sel.node, pos: sel.from, row: 0 };
  if (sel instanceof CellSelection) {
    const rect = selectedRect(state);
    return { table: rect.table, pos: rect.tableStart - 1, row: rect.top };
  }
  // The selection's start, as Word takes it (cellAround gives the innermost cell).
  const $cell = cellAround(sel.$from);
  if (!$cell) return null;
  const table = $cell.node(-1);
  const start = $cell.start(-1);
  return { table, pos: start - 1, row: TableMap.get(table).findCell($cell.pos - start).top };
}

/**
 * The kept w:tc of the rows below a vertically merged cell's first one, when they still match
 * its rows (otherwise the writer doesn't use them either).
 */
function mergedPieces(cell: PMNode): string[] | null {
  const pieces = cell.attrs.mergedTc as string[] | null;
  return pieces && pieces.length === (cell.attrs.rowspan ?? 1) - 1 ? pieces : null;
}

/**
 * The content of a kept w:tc (a piece of a vertically merged cell), which Word keeps but doesn't
 * show while the cell is merged: usually an empty paragraph, sometimes bookmarks or comment
 * marks. Converted without the part's pictures and links (a piece holds none in practice; any
 * it had would stay as kept XML).
 */
function pieceContent(tc: string): PMNode[] {
  const blocks = new Converter(new Map(), new Map()).blocks(parseKept(tc));
  return blocks.length ? blocks : [schema.nodes.paragraph.create()];
}

/**
 * The rows of `table` above row `at` and from row `at` on, as two tables with the table's own
 * properties (w:tblPr, grid, style). A vertically merged cell that crosses the line is cut in
 * two, as Word does: the upper part keeps its content, the lower part starts a merge of its own
 * (w:vMerge restart) from its first row's piece: that piece's w:tcPr, shading, borders and
 * alignment, and content (an empty paragraph when there is no piece).
 * Horizontal merges are unaffected; rows keep their own w:trPr (a repeated header row is not
 * copied into the lower table, as in Word).
 */
function splitRows(table: PMNode, at: number): [PMNode, PMNode] {
  const map = TableMap.get(table);
  const rows: PMNode[] = [];
  table.forEach((row) => rows.push(row));
  const upper = rows.slice(0, at).map((row, r) => {
    const cells: PMNode[] = [];
    row.forEach((cell) => {
      const span = cell.attrs.rowspan ?? 1;
      if (r + span <= at) {
        cells.push(cell);
        return;
      }
      const kept = at - r;
      const pieces = mergedPieces(cell)?.slice(0, kept - 1);
      cells.push(cell.type.create({ ...cell.attrs, rowspan: kept, mergedTc: pieces?.length ? pieces : null }, cell.content, cell.marks));
    });
    return row.copy(Fragment.from(cells));
  });
  // Row `at` gets the lower parts of the cells above it that reach into it, in column order.
  const first: PMNode[] = [];
  for (let c = 0; c < map.width; ) {
    const p = map.map[at * map.width + c];
    const rect = map.findCell(p);
    const cell = table.nodeAt(p)!;
    if (rect.top === at) first.push(cell);
    else if (rect.left === c) {
      const all = mergedPieces(cell);
      const pieces = all?.slice(at - rect.top);
      const piece = all?.[at - rect.top - 1];
      const tcPr = piece ? keptTcPr(piece) : cell.attrs.tcPr;
      // The piece's own looks, not the top cell's: the writer would otherwise write those onto it.
      const looks = piece ? cellInfo(tcPr).model : null;
      first.push(
        cell.type.create(
          {
            ...cell.attrs,
            ...(looks && { background: looks.background, vAlign: looks.vAlign, borders: looks.borders ? { ...looks.borders } : null }),
            rowspan: rect.bottom - at,
            tcPr,
            mergedTc: pieces?.length ? pieces : null,
          },
          piece ? pieceContent(piece) : schema.nodes.paragraph.create(),
        ),
      );
    }
    c = rect.right;
  }
  const lower = [rows[at].copy(Fragment.from(first)), ...rows.slice(at + 1)];
  return [table.copy(Fragment.from(upper)), table.copy(Fragment.from(lower))];
}

/** Wrappers that hold blocks: content controls (w:sdt) and custom XML, whatever their prefix. */
const HOLDER = new Set(['sdt', 'customXml']);
const localName = (open: string): string => /^<(?:[\w.-]+:)?([\w.-]+)/.exec(open)?.[1] ?? '';

/**
 * The wrappers around a table that also go around a new paragraph beside it: the content
 * controls and custom XML, and whatever is outside them. A tracked insertion or move (w:ins,
 * w:moveTo) inside the innermost of them is the table's, not the new paragraph's (it would be
 * recorded as someone else's revision), so it is left off. One outside a control stays: taking
 * it off would close and reopen the control around the paragraph, which would then be saved
 * as two controls with the same w:id.
 */
function holders(wrap: string | null): string | null {
  const layers = parseLayers(wrap);
  let last = -1;
  layers.forEach((l, i) => {
    if (HOLDER.has(localName(l.open))) last = i;
  });
  const kept = layers.slice(0, last + 1);
  return kept.length === layers.length ? wrap : kept.length ? JSON.stringify(kept) : null;
}

/**
 * Word's 分割表格 (表格版面配置 › 合併 › 分割表格, Ctrl+Shift+Enter in a table): the table is
 * split above the row with the cursor (or where the selection starts) into two tables with an
 * empty paragraph between them, where the cursor goes. In the first row, Word puts an empty
 * paragraph before the table instead: the way to get text above a table at the top of a
 * document. The innermost table is split when tables are nested.
 *
 * The whole table is replaced in one step, so it is one undo step, and 追蹤修訂 refuses it as a
 * change of table structure (Word can't record it as a revision either).
 */
export const splitTable: Command = (state, dispatch) => {
  const found = tableAt(state);
  if (!found) return false;
  if (!dispatch) return true;
  const { table, pos, row } = found;
  // The paragraph stays inside whatever holds the table (a content control ...).
  const para = schema.nodes.paragraph.create({ wrap: holders(table.attrs.wrap ?? null) });
  const content = row === 0 ? [para, table] : (() => {
    const [upper, lower] = splitRows(table, row);
    return [upper, para, lower];
  })();
  const tr = state.tr.replaceWith(pos, pos + table.nodeSize, content);
  const at = row === 0 ? pos : pos + content[0].nodeSize;
  tr.setSelection(TextSelection.create(tr.doc, at + 1));
  dispatch(tr.scrollIntoView());
  return true;
};

// ----- sorting rows (Word's 表格 → 排序) -----

const collator = new Intl.Collator('zh-Hant-TW', { numeric: true, sensitivity: 'base' });

/** 1,200 / $30 / 45% / NT$ 1,000 元 → a number; null when the text isn't one. */
function numberValue(text: string): number | null {
  const s = text.replace(/NT\$|[,\s$＄元%％]/g, '');
  return /^[-+]?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : null;
}

/** 2026/9/24, 2026-09-24, 115年9月24日 (民國) → a time; null when the text isn't a date. */
function dateValue(text: string): number | null {
  const m = /^(\d{2,4})\s*[/.\-年]\s*(\d{1,2})\s*[/.\-月]\s*(\d{1,2})\s*日?$/.exec(text.trim());
  if (!m) return null;
  let year = Number(m[1]);
  if (year < 1000) year += 1911; // 民國年
  return Date.UTC(year, Number(m[2]) - 1, Number(m[3]));
}

/** Word's order: numbers by value, dates by time, other text in Traditional Chinese order; empty cells last. */
export function compareCellText(a: string, b: string): number {
  if (!a !== !b) return a ? -1 : 1;
  const na = numberValue(a);
  const nb = numberValue(b);
  if (na != null && nb != null) return na - nb;
  const da = dateValue(a);
  const db = dateValue(b);
  if (da != null && db != null) return da - db;
  return collator.compare(a, b);
}

/** Why the table at the cursor can't be sorted (merged cells), or null when it can. */
export function sortRefusal(state: EditorState): string | null {
  const sel = selection(state);
  if (!sel) return '游標不在表格中。';
  let merged = false;
  sel.table.forEach((row) =>
    row.forEach((cell) => {
      if ((cell.attrs.rowspan ?? 1) > 1 || (cell.attrs.colspan ?? 1) > 1) merged = true;
    }),
  );
  return merged ? '表格有合併儲存格，無法排序。' : null;
}

/**
 * Sorts the table's rows by the column with the cursor. Header rows (w:tblHeader) stay on top;
 * `keepFirst` also keeps the first row in place when it isn't marked as a header.
 */
export const sortTable = (descending: boolean, keepFirst = true): Command => (state, dispatch) => {
  const sel = selection(state);
  if (!sel || sortRefusal(state)) return false;
  const rows: PMNode[] = [];
  sel.table.forEach((row) => rows.push(row));
  let head = 0;
  while (head < rows.length && rows[head].attrs.header) head++;
  if (head === 0 && keepFirst && rows.length > 1) head = 1;
  const col = sel.left;
  const keyed = rows.slice(head).map((row, i) => ({ row, i, key: (row.maybeChild(col)?.textContent ?? '').trim() }));
  keyed.sort((x, y) => {
    // Empty cells go last either way; the order among equal keys is kept.
    if (!x.key !== !y.key) return x.key ? -1 : 1;
    const c = compareCellText(x.key, y.key);
    return (descending ? -c : c) || x.i - y.i;
  });
  if (keyed.every((k, i) => k.i === i)) return true; // already in that order
  if (dispatch) {
    const tablePos = sel.tableStart - 1;
    const table = sel.table.copy(Fragment.from([...rows.slice(0, head), ...keyed.map((k) => k.row)]));
    const tr = state.tr.replaceWith(tablePos, tablePos + sel.table.nodeSize, table);
    tr.setSelection(TextSelection.near(tr.doc.resolve(tablePos + 1)));
    dispatch(tr.scrollIntoView());
  }
  return true;
};
