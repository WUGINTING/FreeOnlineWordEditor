// Word's 自動調整 (表格版面配置 › 儲存格大小 › 自動調整): 自動調整成內容大小, 自動調整成視窗大小 and
// 固定欄寬. The widths are worked out by the pure functions in docx/tableLayout.ts; this file
// measures the cells in the browser (the only part that needs a layout) and applies the result
// as one transaction: the table's w:tblPr / grid, its rows' w:tblPrEx and every cell's width.

import type { Command, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { TableMap } from 'prosemirror-tables';
import {
  MIN_COLUMN, autoFitWidths, columnRanges, columnWidths, defaultTblPr, scaleWidths, tableCells, withCellWidth,
  withGridWidths, withMergedCellWidth, withRowExceptionAutoFit, withTableAutoFit, type AutoFitMode, type CellRange,
} from '../docx/tableLayout';
import { pxToTwips, twipsToPx } from '../units';
import { tableAt } from './tableCommands';

export type { AutoFitMode } from '../docx/tableLayout';

/** Word's default left + right cell margins (0.19 cm each), taken off a cell's width for a table in it. */
const CELL_MARGINS = 216;

export interface AutoFitOptions {
  /** Width of the text area (page minus left and right margins) where the table is, twips. */
  textWidth: number;
  /**
   * The cells' content ranges for 自動調整成內容大小 (twips, cell margins included); by default
   * measured in the view (measureCells). Null when they can't be measured (no layout).
   */
  measure?: (view: EditorView | undefined, table: PMNode, pos: number) => CellRange[] | null;
}

/**
 * Word's 自動調整 for the table with the cursor (the innermost one, or the table selected as a
 * whole):
 *  - contents: each column as wide as its content, narrowed to fit the text area without going
 *    under the narrowest its content allows; automatic layout, table width auto;
 *  - window: the columns scaled in proportion to fill the text area; table width 100 %;
 *  - fixed: the columns keep their current widths, which are written as fixed (twips), with a
 *    fixed layout.
 * The grid and every cell's w:tcW are written in twips (dxa), as Word does. One undo step;
 * refused while 追蹤修訂 is on, as other table formatting.
 */
export function autoFitTable(mode: AutoFitMode, opts: AutoFitOptions): Command {
  return (state, dispatch, view) => {
    const found = tableAt(state);
    if (!found) return false;
    if (!dispatch) return true;
    const { table, pos } = found;
    const map = TableMap.get(table);
    const available = availableWidth(state.doc, pos, opts.textWidth);
    let widths: number[];
    if (mode === 'contents') {
      const cells = (opts.measure ?? measureCells)(view, table, pos);
      if (!cells) return false;
      // A column no cell was measured in (only merged cells over it) still gets a real width.
      const ranges = columnRanges(cells, map.width).map((r) => ({ min: Math.max(MIN_COLUMN, r.min), max: Math.max(MIN_COLUMN, r.max) }));
      widths = autoFitWidths(ranges, available);
    } else {
      const current = columnWidths(table, map, available).widths;
      widths = mode === 'window' ? scaleWidths(current, available) : current;
    }
    const tr = state.tr;
    applyWidths(tr, table, pos, map, widths, mode);
    if (tr.docChanged) dispatch(tr.scrollIntoView());
    return true;
  };
}

/**
 * The width a table may take: the text area, or for a table in a table cell, that cell's width
 * less its margins.
 */
function availableWidth(doc: PMNode, pos: number, textWidth: number): number {
  const $pos = doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    const cell = $pos.node(d);
    if (cell.type.spec.tableRole !== 'cell') continue;
    const outer = $pos.node(d - 2);
    const outerMap = TableMap.get(outer);
    const rect = outerMap.findCell($pos.before(d) - $pos.start(d - 2));
    const cols = columnWidths(outer, outerMap, textWidth).widths.slice(rect.left, rect.right);
    return Math.max(MIN_COLUMN, cols.reduce((s, w) => s + w, 0) - CELL_MARGINS);
  }
  return textWidth;
}

/**
 * Set the table's w:tblPr and grid, its rows' w:tblPrEx, and every cell's width (colwidth and
 * w:tcW) to `widths`.
 */
function applyWidths(tr: Transaction, table: PMNode, pos: number, map: TableMap, widths: number[], mode: AutoFitMode): void {
  const total = widths.reduce((s, w) => s + w, 0);
  const tblPr = withTableAutoFit((table.attrs.tblPr as string | null) || defaultTblPr(table.attrs.styleId), mode, total);
  // The grid attribute is what the writer compares the cells' widths with: the columns keep
  // exactly these twips, and the cells whole pixels of them (as a file is opened).
  tr.setNodeMarkup(pos, undefined, {
    ...table.attrs,
    tblPr,
    grid: widths,
    tblGridXml: withGridWidths(table.attrs.tblGridXml, widths),
  });
  const start = pos + 1;
  // A row's exceptions (w:tblPrEx) may give the table another width or layout for that row.
  table.forEach((row, offset) => {
    const ex = row.attrs.tblPrEx as string | null;
    const next = ex ? withRowExceptionAutoFit(ex, mode, total) : ex;
    if (next !== ex) tr.setNodeMarkup(start + offset, undefined, { ...row.attrs, tblPrEx: next }, row.marks);
  });
  const { rects, cells } = tableCells(table, map);
  for (const [p, cell] of cells) {
    const rect = rects.get(p);
    if (!rect) continue;
    const cols = widths.slice(rect.left, rect.right);
    const width = cols.reduce((s, w) => s + w, 0);
    const attrs: Record<string, unknown> = { ...cell.attrs, colwidth: cols.map((w) => Math.round(twipsToPx(w))) };
    if (cell.attrs.tcPr) attrs.tcPr = withCellWidth(cell.attrs.tcPr, width);
    const pieces = cell.attrs.mergedTc as string[] | null;
    if (pieces) attrs.mergedTc = pieces.map((tc) => withMergedCellWidth(tc, width));
    if (JSON.stringify(attrs) !== JSON.stringify(cell.attrs)) tr.setNodeMarkup(start + p, undefined, attrs, cell.marks);
  }
}

/** Elements the pagination adds inside cells (page-break spacers): not content. */
const NOT_CONTENT = '.dx-spacer, .dx-line-spacer, .dx-row-spacer';

/**
 * The content range of each cell of `table` as the browser lays it out: its content on one line
 * per paragraph (max-content) and at its narrowest (min-content: the longest word, one CJK
 * character, a picture at its size), plus the cell's left and right padding (its margins) and
 * borders, in twips. The cells are copied, each inside copies of its ancestors (so the
 * document's, table style's and cell's CSS still apply), into a hidden box beside the pages
 * (outside the zoomed canvas), measured all at once and removed. An empty cell measures as its
 * margins. Null without a layout (the cells have no width on the page).
 */
export function measureCells(view: EditorView | undefined, table: PMNode, pos: number): CellRange[] | null {
  if (!view) return null;
  const map = TableMap.get(table);
  const doc = view.dom;
  const root = (doc.closest('.dx-root') as HTMLElement | null) ?? doc.parentElement;
  if (!root) return null;
  const box = doc.cloneNode(false) as HTMLElement;
  box.removeAttribute('contenteditable');
  box.removeAttribute('role');
  box.setAttribute('aria-hidden', 'true');
  box.style.cssText = 'position:absolute;left:0;top:0;visibility:hidden;pointer-events:none;width:auto;height:auto;padding:0;margin:0;overflow:visible';
  const measured: { col: number; span: number; td: HTMLElement; inner: HTMLElement; padding: number }[] = [];
  const start = pos + 1;
  const { rects } = tableCells(table, map);
  for (const [p, rect] of rects) {
    const td = view.nodeDOM(start + p);
    if (!(td instanceof HTMLElement)) continue;
    // Copies of the cell's ancestors up to the document, outermost first.
    let copy: HTMLElement = box;
    const chain: HTMLElement[] = [];
    for (let n = td.parentElement; n && n !== doc; n = n.parentElement) chain.unshift(n);
    for (const n of chain) copy = copy.appendChild(n.cloneNode(false) as HTMLElement);
    const cellCopy = copy.appendChild(td.cloneNode(false) as HTMLElement);
    const inner = cellCopy.appendChild(document.createElement('div'));
    inner.style.cssText = 'width:max-content;display:block';
    td.childNodes.forEach((c) => inner.appendChild(c.cloneNode(true)));
    inner.querySelectorAll(NOT_CONTENT).forEach((e) => e.remove());
    // On the page a picture is at most as wide as its cell (max-width:100%), which would make it
    // count as nothing here: it is measured at its own size.
    inner.querySelectorAll<HTMLElement>('img, .dx-img-wrap').forEach((e) => (e.style.maxWidth = 'none'));
    // The cell's margins, and its side borders, which take room from the text on the page.
    const cs = getComputedStyle(td);
    const padding = ['paddingLeft', 'paddingRight', 'borderLeftWidth', 'borderRightWidth']
      .reduce((s, k) => s + (parseFloat(cs[k as 'paddingLeft']) || 0), 0);
    measured.push({ col: rect.left, span: rect.right - rect.left, td, inner, padding });
  }
  root.appendChild(box);
  try {
    if (!measured.some((m) => m.td.getBoundingClientRect().width > 0)) return null;
    const max = measured.map((m) => m.inner.getBoundingClientRect().width);
    measured.forEach((m) => (m.inner.style.width = 'min-content'));
    const min = measured.map((m) => m.inner.getBoundingClientRect().width);
    // Whole pixels, rounded up with half a pixel to spare: the columns are shown in whole pixels,
    // and a width a fraction short would wrap the text it was measured for.
    const twips = (px: number, padding: number) => Math.max(MIN_COLUMN, pxToTwips(Math.ceil(px + padding + 0.5)));
    return measured.map((m, i) => ({ col: m.col, span: m.span, min: twips(min[i], m.padding), max: twips(max[i], m.padding) }));
  } finally {
    box.remove();
  }
}
