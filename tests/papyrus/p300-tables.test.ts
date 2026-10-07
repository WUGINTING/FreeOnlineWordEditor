// persona-300 表格: 插入表格… with any number of rows and columns (up to 63 columns, as Word),
// 欄寬（公分） for the cursor's column, and 合併儲存格 off (and saying why) with one cell.
import { describe, expect, it } from 'vitest';
import { nextTick } from 'vue';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { setColumnWidth, tableInfo } from '../../src/papyrus/editor/tableCommands';
import { domStubs, setup } from './p300Helpers';

domStubs();

const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';

describe('插入表格…', () => {
  it('asks for columns and rows: 12 × 30 is inserted', async () => {
    const d = await setup('<w:p/>', 'DocxToolbar.vue', { styles: [] });
    const btn = (text: string) => Array.from(d.panel.querySelectorAll('button')).find((b) => b.textContent!.trim().startsWith(text)) as HTMLButtonElement;
    btn('表格').click();
    await nextTick();
    btn('插入表格…').click();
    await nextTick();
    const [cols, rows] = Array.from(d.panel.querySelectorAll('.dx-tform input')) as HTMLInputElement[];
    const set = (el: HTMLInputElement, v: string) => {
      el.value = v;
      el.dispatchEvent(new Event('input'));
    };
    set(cols, '12');
    set(rows, '30');
    (d.panel.querySelector('.dx-tform') as HTMLFormElement).dispatchEvent(new Event('submit', { cancelable: true }));
    await nextTick();
    const table = d.view.state.doc.child(0);
    expect(table.type.name).toBe('table');
    expect(table.childCount).toBe(30);
    expect(table.child(0).childCount).toBe(12);
    expect(d.panel.querySelector('.dx-grid')).toBeNull();
    d.done();
  });
});

describe('欄寬（公分）', () => {
  it('shows and sets the cursor column width in cm; the other column keeps its width; saved in twips', async () => {
    const d = await setup(TABLE, 'TablePanel.vue');
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    const box = d.panel.querySelector('input[aria-label="欄寬（公分）"]') as HTMLInputElement;
    expect(box.value).toBe('2.65'); // 1500 twips
    box.value = '3';
    box.dispatchEvent(new Event('input'));
    box.dispatchEvent(new Event('change'));
    const t = d.view.state.doc.child(0);
    expect(t.child(0).child(0).attrs.colwidth[0]).toBeCloseTo(1701 / 15, 6);
    expect(t.child(1).child(0).attrs.colwidth[0]).toBeCloseTo(1701 / 15, 6);
    expect(t.child(0).child(1).attrs.colwidth[0]).toBe(200);
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:gridCol w:w="1701"/><w:gridCol w:w="3000"/>');
    expect(xml).toContain('<w:tcW w:w="1701" w:type="dxa"/>');
    // Undo in one step.
    d.editor.undo();
    expect(d.view.state.doc.child(0).child(0).child(0).attrs.colwidth[0]).toBe(100);
    d.done();
  });

  it('no command outside a table; differing selected columns show no single width', async () => {
    const d = await setup(TABLE);
    const view = d.view;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, view.state.doc.content.size - 1)));
    expect(setColumnWidth(100)(view.state)).toBe(false);
    const cells: number[] = [];
    view.state.doc.descendants((n, pos) => void (n.type.name === 'table_cell' && cells.push(pos)));
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cells[0], cells[1])));
    expect(tableInfo(view.state)!.colWidth).toBeNull();
    d.done();
  });
});

describe('合併儲存格', () => {
  it('is off with one cell, saying to select two or more; on with two selected', async () => {
    const d = await setup(TABLE, 'DocxToolbar.vue', { styles: [] });
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    const merge = () => Array.from(d.panel.querySelectorAll('button')).find((b) => b.textContent!.trim() === '合併儲存格') as HTMLButtonElement;
    expect(merge().disabled).toBe(true);
    expect(merge().getAttribute('aria-description')).toBe('請先選取兩個以上的儲存格');
    expect(merge().title).toBe('請先選取兩個以上的儲存格');
    const cells: number[] = [];
    d.view.state.doc.descendants((n, pos) => void (n.type.name === 'table_cell' && cells.push(pos)));
    d.view.dispatch(d.view.state.tr.setSelection(CellSelection.create(d.view.state.doc, cells[0], cells[1])));
    await d.refresh();
    expect(merge().disabled).toBe(false);
    expect(merge().getAttribute('aria-description')).toBeNull();
    merge().click();
    expect(d.view.state.doc.child(0).child(0).childCount).toBe(1);
    d.done();
  });
});
