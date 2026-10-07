// Dragging a table's column border, as Word does it. prosemirror-tables' columnResizing finds
// the border under the mouse and draws the handle; this plugin, placed before it, takes the
// press and the drag:
//  - a border between two columns moves between them: what one gains the other loses, so the
//    table keeps its width. (The library widened the one column and the table with it: a drag
//    that started on a border, such as one meant to select a row's text from the start of a
//    cell, pushed the table out of the page.)
//  - the table's right edge changes the last column, and the table's width with it;
//  - the border follows the mouse at any zoom (the page is scaled, the mouse's distance is not);
//  - only the left button drags.
// The columns are shown while dragging; the document changes on release, as one undo step.

import { Plugin, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { TableMap, columnResizingPluginKey } from 'prosemirror-tables';

export interface ColumnResizeOptions {
  /** The narrowest a column can be dragged to, px (as columnResizing's cellMinWidth). */
  cellMinWidth: number;
  /** The page's zoom (1 = 100 %). */
  zoom: () => number;
}

/**
 * The widths (px) of the columns left and right of a border dragged by `delta`: the left one in
 * whole pixels, the right one the rest of what the two had. Neither goes under `min` (a column
 * already narrower keeps at least what it has). `right` is null for the table's right edge:
 * only the last column changes. A border that wasn't moved leaves the widths as they are.
 */
export function moveBorder(left: number, right: number | null, delta: number, min: number): [number, number | null] {
  if (!delta) return [left, right];
  const narrowest = Math.min(min, left);
  if (right == null) return [Math.max(narrowest, Math.round(left + delta)), null];
  const total = left + right;
  const widest = total - Math.min(min, right);
  const moved = Math.min(widest, Math.max(narrowest, Math.round(left + delta)));
  return [moved, total - moved];
}

/** The table and the grid column whose right border the resize handle `handle` (a cell's position) is on. */
function borderAt(view: EditorView, handle: number): { table: PMNode; start: number; map: TableMap; col: number } | null {
  if (handle < 0 || handle > view.state.doc.content.size) return null;
  const $cell = view.state.doc.resolve(handle);
  const cell = $cell.nodeAfter;
  const table = $cell.node(-1);
  if (!cell || cell.type.spec.tableRole !== 'cell' || table.type.spec.tableRole !== 'table') return null;
  const start = $cell.start(-1);
  const map = TableMap.get(table);
  return { table, start, map, col: map.colCount($cell.pos - start) + (cell.attrs.colspan as number) - 1 };
}

/** A column's width, px: what its cells say, else as a cell of that column alone is laid out (0 when neither). */
function columnPx(view: EditorView, table: PMNode, start: number, map: TableMap, col: number): number {
  let measured = 0;
  for (let row = 0; row < map.height; row++) {
    const pos = map.map[row * map.width + col];
    const cell = table.nodeAt(pos)!;
    const width = (cell.attrs.colwidth as number[] | null)?.[col - map.colCount(pos)];
    if (width) return width;
    if (!measured && cell.attrs.colspan === 1) measured = (view.nodeDOM(start + pos) as HTMLElement | null)?.offsetWidth ?? 0;
  }
  return measured;
}

/** Give the grid columns in `widths` (column -> px) their width in every cell over them. */
function setColumns(tr: Transaction, table: PMNode, start: number, map: TableMap, widths: Map<number, number>): void {
  for (const pos of new Set(map.map)) {
    const cell = table.nodeAt(pos)!;
    const first = map.colCount(pos);
    const span = cell.attrs.colspan as number;
    const colwidth: number[] = (cell.attrs.colwidth as number[] | null)?.slice() ?? new Array(span).fill(0);
    let changed = false;
    for (const [col, width] of widths) {
      const i = col - first;
      if (i < 0 || i >= span || colwidth[i] === width) continue;
      colwidth[i] = width;
      changed = true;
    }
    if (changed) tr.setNodeMarkup(start + pos, null, { ...cell.attrs, colwidth });
  }
}

function startDrag(view: EditorView, event: MouseEvent, opts: ColumnResizeOptions): boolean {
  const resize = columnResizingPluginKey.getState(view.state);
  if (!view.editable || !resize || resize.activeHandle < 0 || resize.dragging) return false;
  // Another button on a border is not a drag (the library would start one), nor a click in the text.
  if (event.button !== 0) return true;
  const border = borderAt(view, resize.activeHandle);
  if (!border) return false;
  const { table, start, map, col } = border;
  const left = columnPx(view, table, start, map, col);
  const right = col === map.width - 1 ? null : columnPx(view, table, start, map, col + 1);
  // A column nothing gives a width to (only merged cells over it): left to the library.
  if (!left || right === 0) return false;

  let dom: Node | null = view.domAtPos(start).node;
  while (dom && dom.nodeName !== 'TABLE') dom = dom.parentNode;
  const tableDom = dom as HTMLElement | null;
  const cols = tableDom?.querySelector('colgroup')?.children;
  const colDom = (c: number) => cols?.[c] as HTMLElement | undefined;
  // What the table view set, put back on release: the view redraws from the document then.
  const was = { left: colDom(col)?.style.width, right: colDom(col + 1)?.style.width, table: tableDom?.style.width, min: tableDom?.style.minWidth };
  const show = ([l, r]: [number, number | null]) => {
    if (!tableDom || !colDom(col)) return;
    colDom(col)!.style.width = `${l}px`;
    if (r != null) {
      if (colDom(col + 1)) colDom(col + 1)!.style.width = `${r}px`;
    } else if (was.table) {
      tableDom.style.width = `${parseFloat(was.table) + l - left}px`;
    } else if (was.min) {
      tableDom.style.minWidth = `${parseFloat(was.min) + l - left}px`;
    }
  };
  const restore = () => {
    if (!tableDom || !colDom(col)) return;
    colDom(col)!.style.width = was.left ?? '';
    if (right != null && colDom(col + 1)) colDom(col + 1)!.style.width = was.right ?? '';
    tableDom.style.width = was.table ?? '';
    tableDom.style.minWidth = was.min ?? '';
  };

  const startX = event.clientX;
  // Where the mouse was while the button was last known to be down.
  let x = startX;
  const widths = () => moveBorder(left, right, (x - startX) / (opts.zoom() || 1), opts.cellMinWidth);
  const win = view.dom.ownerDocument.defaultView ?? window;
  const finish = () => {
    win.removeEventListener('mouseup', up);
    win.removeEventListener('mousemove', move);
    restore();
    if (view.isDestroyed) return;
    const now = columnResizingPluginKey.getState(view.state);
    if (!now?.dragging) return;
    const at = borderAt(view, now.activeHandle);
    const [l, r] = widths();
    if (at && (l !== left || r !== right)) {
      const changes = new Map([[at.col, l]]);
      if (r != null) changes.set(at.col + 1, r);
      const tr = view.state.tr;
      setColumns(tr, at.table, at.start, at.map, changes);
      view.dispatch(tr);
    }
    // On its own: the change above may be refused (追蹤修訂 can't record it), the drag still ends.
    view.dispatch(view.state.tr.setMeta(columnResizingPluginKey, { setDragging: null }));
  };
  const up = (e: MouseEvent) => {
    x = e.clientX;
    finish();
  };
  const move = (e: MouseEvent) => {
    // Let go where the page didn't see it (outside the window): the drag ended there.
    if (!(e.buttons & 1)) return finish();
    x = e.clientX;
    show(widths());
  };
  // The library's own state: the handle stays on this border and is drawn as dragged.
  view.dispatch(view.state.tr.setMeta(columnResizingPluginKey, { setDragging: { startX, startWidth: left } }));
  win.addEventListener('mouseup', up);
  win.addEventListener('mousemove', move);
  event.preventDefault();
  return true;
}

/** The plugin; goes before prosemirror-tables' columnResizing, which it takes the mouse press from. */
export function columnResize(opts: ColumnResizeOptions): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => startDrag(view, event, opts),
      },
    },
  });
}
