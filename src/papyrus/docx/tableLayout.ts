// Table layout: the columns' saved widths, Word's 自動調整 (AutoFit), the widths it works out
// and the w:tblPr / w:tblGrid / w:tcW it writes for each choice, and the kept cell XML
// 分割表格 (Split Table) needs. Pure functions; the editor measures the cells
// (editor/tableAutoFit.ts) and calls these.

import type { Node as PMNode } from 'prosemirror-model';
import type { TableMap } from 'prosemirror-tables';
import { NS, child, children, escapeXml, parseFragment, serializeXml } from './xml';
import { parseKept } from './props';
import { pxToTwips, twipsToPx } from '../units';

/** Word's three 自動調整 choices. */
export type AutoFitMode = 'contents' | 'window' | 'fixed';

/** The narrowest and widest a column's content can be laid out (twips, cell margins included). */
export interface ColumnRange {
  /** The widest unbreakable piece (a word, one CJK character, a picture). */
  min: number;
  /** The content on one line each paragraph (no wrapping). */
  max: number;
}

/** A cell's measured range and the grid columns it covers. */
export interface CellRange extends ColumnRange {
  col: number;
  span: number;
}

// Child order defined by ECMA-376 (CT_TblPr: CT_TblPrBase, then tblPrChange).
const TBLPR_ORDER = [
  'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
  'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption',
  'tblDescription', 'tblPrChange',
];

/**
 * The w:tblPr the writer gives a table made in the editor (single borders, fixed layout, Word's
 * default look). AutoFit starts from it when the table has none yet.
 */
export function defaultTblPr(styleId: string | null): string {
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('');
  return (
    '<w:tblPr>' +
    (styleId ? `<w:tblStyle w:val="${escapeXml(styleId)}"/>` : '') +
    `<w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders>` +
    '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    '</w:tblPr>'
  );
}

/** The grid rectangle a table cell covers. */
export interface CellRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
  /** Index in map.map where the grid first meets the cell. */
  first: number;
}

/**
 * Every cell of a table and the grid rectangle it covers (as TableMap.findCell gives it), by
 * its position in the table, in one pass. findCell searches the whole map for each call and
 * nodeAt walks the rows, so calling them per grid slot was quadratic in the table's size.
 */
export function tableCells(table: PMNode, map: TableMap): { rects: Map<number, CellRect>; cells: Map<number, PMNode> } {
  const cells = new Map<number, PMNode>();
  table.forEach((row, rowOffset) => {
    row.forEach((cell, cellOffset) => cells.set(rowOffset + 1 + cellOffset, cell));
  });
  const rects = new Map<number, CellRect>();
  const { width, height } = map;
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i];
    if (rects.has(pos)) continue;
    // Same as TableMap.findCell: from the first slot, as far right / down as the cell reaches.
    const left = i % width;
    const top = (i / width) | 0;
    let right = left + 1;
    let bottom = top + 1;
    for (let j = 1; right < width && map.map[i + j] === pos; j++) right++;
    for (let j = 1; bottom < height && map.map[i + width * j] === pos; j++) bottom++;
    rects.set(pos, { left, top, right, bottom, first: i });
  }
  return { rects, cells };
}

/** Narrowest column (twips) the editor gives a column it knows nothing about (as columnResizing's cellMinWidth). */
export const MIN_COLUMN = 360;

/**
 * The columns' current widths in twips, as they are saved: the file's grid (w:gridCol) where a
 * column still has the width it was opened with, the cells' widths (px) where it was resized,
 * and the rest of `fillWidth` shared by columns with neither (at least MIN_COLUMN each).
 * `fromGrid` tells which columns are still the file's own.
 */
export function columnWidths(
  table: PMNode,
  map: TableMap,
  fillWidth: number,
  layout: { rects: Map<number, CellRect>; cells: Map<number, PMNode> } = tableCells(table, map),
): { widths: number[]; fromGrid: boolean[] } {
  const { rects, cells } = layout;
  const original = table.attrs.grid as number[] | null;
  const px: (number | null)[] = new Array(map.width).fill(null);
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i];
    const rect = rects.get(pos)!;
    if (rect.first !== i) continue; // each cell once, where the grid first meets it
    const widths = cells.get(pos)!.attrs.colwidth as number[] | null;
    if (!widths) continue;
    for (let c = rect.left; c < rect.right; c++) if (px[c] == null && widths[c - rect.left]) px[c] = widths[c - rect.left];
  }
  const fromGrid = px.map((w, c) =>
    !!original && original.length === map.width && (w == null || w === Math.round(twipsToPx(original[c]))),
  );
  const cols = px.map((w, c) => (fromGrid[c] ? original![c] : w != null ? pxToTwips(w) : null));
  const known = cols.reduce<number>((s, w) => s + (w ?? 0), 0);
  const unknown = cols.filter((w) => w == null).length;
  const fill = unknown ? Math.max(MIN_COLUMN, Math.round((fillWidth - known) / unknown)) : 0;
  return { widths: cols.map((w) => w ?? fill), fromGrid };
}

/**
 * Per grid column, the range its cells need. Cells spanning one column set it directly; a
 * merged cell spanning several only widens its columns when they are together narrower than
 * it needs (the shortfall shared in proportion to what they have, evenly when they have
 * nothing), as a browser's automatic table layout does.
 */
export function columnRanges(cells: readonly CellRange[], columns: number): ColumnRange[] {
  const out: ColumnRange[] = Array.from({ length: columns }, () => ({ min: 0, max: 0 }));
  for (const c of cells) {
    if (c.span !== 1 || c.col < 0 || c.col >= columns) continue;
    out[c.col].min = Math.max(out[c.col].min, c.min);
    out[c.col].max = Math.max(out[c.col].max, c.max, c.min);
  }
  const spanned = cells.filter((c) => c.span > 1 && c.col >= 0 && c.col < columns).sort((a, b) => a.span - b.span);
  for (const c of spanned) {
    const cols = out.slice(c.col, Math.min(columns, c.col + c.span));
    const widen = (key: keyof ColumnRange, need: number) => {
      const have = cols.reduce((s, r) => s + r[key], 0);
      if (need <= have) return;
      for (const r of cols) r[key] += have > 0 ? ((need - have) * r[key]) / have : (need - have) / cols.length;
    };
    widen('min', c.min);
    widen('max', Math.max(c.max, c.min));
    for (const r of cols) r.max = Math.max(r.max, r.min);
  }
  return out;
}

/**
 * Word's 自動調整成內容大小: each column as wide as its widest content. When that is wider than
 * the text area, the short columns (a number, a two-character heading) stay on one line and the
 * long ones are narrowed: a column whose content fits in an even share of the room left keeps
 * it, and the columns that don't share what is left in proportion to their content, none below
 * the narrowest its content allows (such a column keeps its minimum and the others share the
 * rest). When even the minimums don't fit, every column gets its minimum and the table runs
 * past the margin, as in Word. Whole twips; the total is exact when narrowed.
 */
export function autoFitWidths(cols: readonly ColumnRange[], available: number): number[] {
  const min = cols.map((c) => Math.max(0, c.min));
  const max = cols.map((c, i) => Math.max(min[i], c.max));
  const sum = (idx: number[], a: number[]) => idx.reduce((s, i) => s + a[i], 0);
  const all = cols.map((_, i) => i);
  if (sum(all, max) <= available) return max.map((w) => Math.round(w));
  if (sum(all, min) >= available) return min.map((w) => Math.round(w));
  // The short columns: those whose content fits in an even share of the room they leave.
  const kept = new Set<number>();
  for (;;) {
    const rest = all.filter((i) => !kept.has(i));
    const share = (available - sum([...kept], max)) / rest.length;
    const fits = rest.filter((i) => max[i] <= share);
    if (!fits.length) break;
    fits.forEach((i) => kept.add(i));
  }
  // Keeping them must leave the long columns their minimums; otherwise every column is narrowed.
  const long = all.filter((i) => !kept.has(i));
  if (sum([...kept], max) + sum(long, min) > available) kept.clear();
  const widths = all.map((i) => (kept.has(i) ? max[i] : 0));
  const shrink = all.filter((i) => !kept.has(i));
  const pinned = new Set<number>();
  for (;;) {
    const free = shrink.filter((i) => !pinned.has(i));
    const room = available - sum([...kept], max) - sum([...pinned], min);
    const k = sum(free, max) > 0 ? room / sum(free, max) : 0;
    for (const i of pinned) widths[i] = min[i];
    for (const i of free) widths[i] = max[i] * k;
    // Pin every column under its minimum at once; the others are worked out again.
    const under = free.filter((i) => widths[i] < min[i]);
    if (!under.length) break;
    under.forEach((i) => pinned.add(i));
  }
  return roundTo(widths, available);
}

/** Every width scaled so they add up to `total` (自動調整成視窗大小), whole twips. */
export function scaleWidths(widths: readonly number[], total: number): number[] {
  const have = widths.reduce((s, w) => s + w, 0);
  if (!widths.length) return [];
  const scaled = have > 0 ? widths.map((w) => (w * total) / have) : widths.map(() => total / widths.length);
  return roundTo(scaled, total);
}

/** Whole numbers that add up to `total`: the rounding difference goes to the widest column. */
function roundTo(widths: number[], total: number): number[] {
  const out = widths.map((w) => Math.round(w));
  const diff = Math.round(total) - out.reduce((s, w) => s + w, 0);
  if (diff && out.length) {
    const widest = out.indexOf(Math.max(...out));
    out[widest] += diff;
  }
  return out;
}

/** Insert `el` among `parent`'s children in schema order, leaving the others where they are. */
function insertInOrder(parent: Element, el: Element, order: string[]): void {
  const rank = order.indexOf(el.localName);
  for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
    const r = c.namespaceURI === NS.w ? order.indexOf(c.localName) : -1;
    if (r > rank || r < 0) {
      parent.insertBefore(el, c);
      return;
    }
  }
  parent.appendChild(el);
}

/** The w: child `local` of `parent`, created in schema order when missing. */
function ensureIn(parent: Element, local: string, order: string[]): Element {
  let el = child(parent, local);
  if (!el) {
    el = parent.ownerDocument.createElementNS(NS.w, 'w:' + local);
    insertInOrder(parent, el, order);
  }
  return el;
}

function setW(el: Element, local: string, value: string | number): void {
  el.setAttributeNS(NS.w, 'w:' + local, String(value));
}

/**
 * The table's w:tblPr as Word writes it for a 自動調整 choice; every other child stays as it
 * was, where it was:
 *  - contents: w:tblW auto (0) and w:tblLayout autofit;
 *  - window: w:tblW 100 % (pct 5000) and w:tblLayout autofit;
 *  - fixed: w:tblLayout fixed and w:tblW the columns' total (dxa).
 * Automatic layout is written, not left out: w:tblLayout is part of CT_TblPrBase, so the table
 * style (or a conditional w:tblStylePr) may say fixed, and without its own element the table
 * would stay fixed in Word.
 */
export function withTableAutoFit(tblPr: string, mode: AutoFitMode, total: number): string {
  return patchAutoFit(tblPr, mode, total, (el, local) => ensureIn(el, local, TBLPR_ORDER));
}

/**
 * A row's w:tblPrEx (table property exceptions for that row) for a 自動調整 choice: a w:tblW or
 * w:tblLayout it has is set as on the table, so the row doesn't keep the old width or layout.
 * Nothing is added; the rest stays as it was.
 */
export function withRowExceptionAutoFit(tblPrEx: string, mode: AutoFitMode, total: number): string {
  return patchAutoFit(tblPrEx, mode, total, (el, local) => child(el, local));
}

/** w:tblW and w:tblLayout of `xml` (a w:tblPr or w:tblPrEx) for `mode`; `get` finds or makes them. */
function patchAutoFit(xml: string, mode: AutoFitMode, total: number, get: (el: Element, local: string) => Element | null): string {
  const el = parseFragment(xml);
  const tblW = get(el, 'tblW');
  const [w, type] = mode === 'contents' ? [0, 'auto'] : mode === 'window' ? [5000, 'pct'] : [Math.round(total), 'dxa'];
  if (tblW) {
    setW(tblW, 'w', w);
    setW(tblW, 'type', type);
  }
  const layout = get(el, 'tblLayout');
  if (layout) setW(layout, 'type', mode === 'fixed' ? 'fixed' : 'autofit');
  const out = serializeXml(el);
  return out === serializeXml(parseFragment(xml)) ? xml : out;
}

/** A w:tblGrid with these column widths: the file's own one patched when it has as many columns. */
export function withGridWidths(tblGrid: string | null, widths: readonly number[]): string {
  if (tblGrid) {
    const el = parseFragment(tblGrid);
    const cols = children(el, 'gridCol');
    if (cols.length === widths.length) {
      cols.forEach((c, i) => setW(c, 'w', widths[i]));
      return serializeXml(el);
    }
  }
  return `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
}

/**
 * A w:tcPr (or null: a new one) with w:tcW set to `width` twips (dxa); the rest is kept. In
 * CT_TcPr, w:tcW comes first, after w:cnfStyle only.
 */
export function withCellWidth(tcPr: string | null, width: number): string {
  const el = parseFragment(tcPr ?? '<w:tcPr/>');
  const tcW = child(el, 'tcW') ?? el.ownerDocument.createElementNS(NS.w, 'w:tcW');
  if (!tcW.parentNode) {
    const cnf = child(el, 'cnfStyle');
    el.insertBefore(tcW, cnf ? cnf.nextSibling : el.firstChild);
  }
  setW(tcW, 'w', Math.round(width));
  setW(tcW, 'type', 'dxa');
  return serializeXml(el);
}

/** A kept w:tc (a piece of a vertically merged cell) with its w:tcW set to `width` twips. */
export function withMergedCellWidth(tc: string, width: number): string {
  const el = parseKept(tc);
  const old = child(el, 'tcPr');
  const next = parseFragment(withCellWidth(old ? serializeXml(old) : null, width));
  const node = el.ownerDocument.importNode(next, true);
  if (old) el.replaceChild(node, old);
  else el.insertBefore(node, el.firstChild);
  return serializeXml(el);
}

/** The w:tcPr of a kept w:tc (a piece of a vertically merged cell), or null when it has none. */
export function keptTcPr(tc: string): string | null {
  const tcPr = child(parseKept(tc), 'tcPr');
  return tcPr ? serializeXml(tcPr) : null;
}
