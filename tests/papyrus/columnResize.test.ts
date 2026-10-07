// Dragging a table's column border, as Word: a border between two columns moves between them
// (the table keeps its width, so it can't be dragged out of the page); the table's right edge
// changes the last column. The border follows the mouse at any zoom.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import { columnResizingPluginKey } from 'prosemirror-tables';
import { moveBorder } from '../../src/papyrus/editor/columnResize';
import { docxOf, domStubs, setup } from './p300Helpers';

domStubs();
// A press this plugin leaves alone goes on to the editor, which looks up what is under the mouse.
beforeAll(() => {
  (document as any).elementFromPoint ??= () => null;
});

const tc =(w: number, text: string, more = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${more}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
// 100 px, 200 px and 100 px; the second row's first cell covers the first two columns.
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/><w:gridCol w:w="1500"/></w:tblGrid>' +
  `<w:tr>${tc(1500, 'A')}${tc(3000, 'B')}${tc(1500, 'C')}</w:tr>` +
  `<w:tr>${tc(4500, 'D', '<w:gridSpan w:val="2"/>')}${tc(1500, 'E')}</w:tr></w:tbl><w:p/>`;

/** Position of the cell whose text is `text`. */
function cellPos(view: EditorView, text: string): number {
  let found = -1;
  view.state.doc.descendants((n, pos) => {
    if (n.type.name === 'table_cell' && n.textContent === text) found = pos;
  });
  return found;
}

/** The widths (px) of each row's cells. */
function widths(view: EditorView): number[][] {
  const table = view.state.doc.child(0);
  const rows: number[][] = [];
  table.forEach((row) => {
    const out: number[] = [];
    row.forEach((cell) => out.push(...(cell.attrs.colwidth as number[])));
    rows.push(out);
  });
  return rows;
}

const mouse = (type: string, x: number, init: MouseEventInit = {}) =>
  new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: 10, button: 0, buttons: type === 'mouseup' ? 0 : 1, ...init });

/** The mouse over the right border of the cell `text` (as the hover finds it), pressed there. */
function press(view: EditorView, text: string, x = 500, init: MouseEventInit = {}): MouseEvent {
  view.dispatch(view.state.tr.setMeta(columnResizingPluginKey, { setHandle: cellPos(view, text) }));
  const down = mouse('mousedown', x, init);
  view.dom.dispatchEvent(down);
  return down;
}

/** Drag the right border of the cell `text` by `dx` pixels on screen. */
function drag(view: EditorView, text: string, dx: number): void {
  press(view, text, 500);
  window.dispatchEvent(mouse('mousemove', 500 + dx));
  window.dispatchEvent(mouse('mouseup', 500 + dx));
}

describe('moveBorder', () => {
  it('moves a border between two columns: what one gains the other loses', () => {
    expect(moveBorder(100, 200, 50, 24)).toEqual([150, 150]);
    expect(moveBorder(100, 200, -30, 24)).toEqual([70, 230]);
  });

  it('stops where either column would be narrower than the narrowest allowed', () => {
    expect(moveBorder(100, 200, 500, 24)).toEqual([276, 24]);
    expect(moveBorder(100, 200, -500, 24)).toEqual([24, 276]);
  });

  it("the table's right edge (no next column) changes the last column only", () => {
    expect(moveBorder(100, null, 60, 24)).toEqual([160, null]);
    expect(moveBorder(100, null, -500, 24)).toEqual([24, null]);
  });

  it('whole pixels that add up to what the two columns had', () => {
    expect(moveBorder(100, 200, 33.4, 24)).toEqual([133, 167]);
  });
});

describe('dragging a column border', () => {
  it('between two columns: the border moves, the table keeps its width', async () => {
    const d = await setup(TABLE);
    drag(d.view, 'A', 50);
    expect(widths(d.view)).toEqual([[150, 150, 100], [150, 150, 100]]);
    d.done();
  });

  it('cannot push the table out of the page: it stops at the next column', async () => {
    const d = await setup(TABLE);
    drag(d.view, 'A', 412);
    expect(widths(d.view)).toEqual([[276, 24, 100], [276, 24, 100]]);
    drag(d.view, 'A', -1000);
    expect(widths(d.view)).toEqual([[24, 276, 100], [24, 276, 100]]);
    d.done();
  });

  it("the table's right edge changes the last column; the others keep their widths", async () => {
    const d = await setup(TABLE);
    drag(d.view, 'C', 60);
    expect(widths(d.view)).toEqual([[100, 200, 160], [100, 200, 160]]);
    d.done();
  });

  it('follows the mouse when zoomed: 150 px on screen at 150 % is 100 px on the page', async () => {
    const d = await setup(TABLE);
    d.editor.setZoom(1.5);
    drag(d.view, 'B', -150);
    expect(widths(d.view)).toEqual([[100, 100, 200], [100, 100, 200]]);
    d.done();
  });

  it('shows the columns while dragging; the document changes on release, in one undo step', async () => {
    const d = await setup(TABLE);
    const cols = () => Array.from(d.view.dom.querySelectorAll('table col')).map((c) => (c as HTMLElement).style.width);
    const table = d.view.dom.querySelector('table') as HTMLElement;
    expect(cols()).toEqual(['100px', '200px', '100px']);
    press(d.view, 'A');
    window.dispatchEvent(mouse('mousemove', 550));
    expect(cols()).toEqual(['150px', '150px', '100px']);
    expect(table.style.width).toBe('400px');
    expect(widths(d.view)[0]).toEqual([100, 200, 100]);
    window.dispatchEvent(mouse('mouseup', 550));
    expect(widths(d.view)[0]).toEqual([150, 150, 100]);
    expect(columnResizingPluginKey.getState(d.view.state)!.dragging).toBeFalsy();
    d.editor.undo();
    expect(widths(d.view)).toEqual([[100, 200, 100], [100, 200, 100]]);
    d.done();
  });

  it('a border released where it was changes nothing (and the columns show as before)', async () => {
    const d = await setup(TABLE);
    const before = d.view.state.doc;
    press(d.view, 'A');
    window.dispatchEvent(mouse('mousemove', 560));
    window.dispatchEvent(mouse('mousemove', 500));
    window.dispatchEvent(mouse('mouseup', 500));
    expect(d.view.state.doc.eq(before)).toBe(true);
    expect((d.view.dom.querySelector('table col') as HTMLElement).style.width).toBe('100px');
    d.done();
  });

  it('only the left button drags: a right-click on a border resizes nothing', async () => {
    const d = await setup(TABLE);
    const down = press(d.view, 'A', 500, { button: 2, buttons: 2 });
    expect(down.defaultPrevented).toBe(false);
    expect(columnResizingPluginKey.getState(d.view.state)!.dragging).toBeFalsy();
    window.dispatchEvent(mouse('mousemove', 700, { buttons: 0 }));
    window.dispatchEvent(mouse('mouseup', 700, { button: 2 }));
    expect(widths(d.view)).toEqual([[100, 200, 100], [100, 200, 100]]);
    d.done();
  });

  it('a release the page never saw (outside the window) ends the drag where the button was last down', async () => {
    const d = await setup(TABLE);
    press(d.view, 'A');
    window.dispatchEvent(mouse('mousemove', 520));
    // The button is up by now: the mouse comes back far from where it was let go.
    window.dispatchEvent(mouse('mousemove', 900, { buttons: 0 }));
    expect(widths(d.view)[0]).toEqual([120, 180, 100]);
    expect(columnResizingPluginKey.getState(d.view.state)!.dragging).toBeFalsy();
    window.dispatchEvent(mouse('mousemove', 950));
    expect(widths(d.view)[0]).toEqual([120, 180, 100]);
    d.done();
  });

  it('is saved as the two columns in twips; the third column and the table width are the file\'s own', async () => {
    const d = await setup(TABLE);
    drag(d.view, 'A', 50);
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:gridCol w:w="2250"/><w:gridCol w:w="2250"/><w:gridCol w:w="1500"/>');
    expect(xml).toContain('<w:tblW w:w="6000" w:type="dxa"/>');
    d.done();
  });

  it('with 追蹤修訂 on the change is refused with a notice (it cannot be recorded); the drag still ends', async () => {
    const d = await setup(TABLE);
    d.editor.setTrackChanges(true);
    drag(d.view, 'A', 50);
    expect(d.notices).toEqual(['追蹤修訂開啟時，無法變更表格或儲存格的格式，請先關閉「追蹤修訂」。']);
    expect(widths(d.view)).toEqual([[100, 200, 100], [100, 200, 100]]);
    expect((d.view.dom.querySelector('table col') as HTMLElement).style.width).toBe('100px');
    expect(columnResizingPluginKey.getState(d.view.state)!.dragging).toBeFalsy();
    // The next drag works as usual once tracking is off.
    d.editor.setTrackChanges(false);
    drag(d.view, 'A', 50);
    expect(widths(d.view)[0]).toEqual([150, 150, 100]);
    d.done();
  });

  it('does nothing in a read-only document', async () => {
    const { DocxEditor } = await import('../../src/papyrus/editor/core');
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host, { editable: false });
    await editor.open(await docxOf(TABLE));
    const view = editor.view!;
    const down = press(view, 'A');
    expect(down.defaultPrevented).toBe(false);
    window.dispatchEvent(mouse('mousemove', 550));
    window.dispatchEvent(mouse('mouseup', 550));
    expect(widths(view)).toEqual([[100, 200, 100], [100, 200, 100]]);
    editor.destroy();
    host.remove();
  });
});
