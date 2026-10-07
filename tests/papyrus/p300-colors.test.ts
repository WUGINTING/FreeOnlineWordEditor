// persona-300 顏色: Word-like colour menus with named standard colours (aria-label and title),
// 自動 / 無色彩, and 其他色彩… (the browser's picker) for any other colour.
import { describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { schema } from '../../src/papyrus/editor/schema';
import { domStubs, setup } from './p300Helpers';

domStubs();

function selectAll(view: any) {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.child(0).nodeSize - 1)));
}
const wait = () => new Promise((r) => setTimeout(r, 0));

describe('colour menus', () => {
  it('text colour: a named swatch applies it, 自動 removes it; each swatch has a Chinese name', async () => {
    const d = await setup('<w:p><w:r><w:t>主旨</w:t></w:r></w:p>', 'DocxToolbar.vue', { styles: [] });
    selectAll(d.view);
    await d.refresh();
    const button = d.panel.querySelector('button[aria-label^="文字顏色："]') as HTMLButtonElement;
    expect(button.getAttribute('aria-haspopup')).toBe('dialog');
    button.click();
    await wait();
    const swatches = Array.from(d.panel.querySelectorAll<HTMLButtonElement>('.dx-cp-swatch'));
    expect(swatches.length).toBeGreaterThanOrEqual(10);
    for (const s of swatches) {
      expect(s.getAttribute('aria-label')).toMatch(/[一-鿿]/);
      expect(s.title).toBe(s.getAttribute('aria-label'));
    }
    swatches.find((s) => s.title === '深藍')!.click();
    await wait();
    const color = () => schema.marks.color.isInSet(d.view.state.doc.child(0).child(0).marks)?.attrs.color;
    expect(color()).toBe('#002060');
    expect(d.panel.querySelector('.dx-cp-menu')).toBeNull();
    selectAll(d.view);
    button.click();
    await wait();
    (Array.from(d.panel.querySelectorAll('.dx-cp-menu button')).find((b) => b.textContent === '自動') as HTMLButtonElement).click();
    expect(color()).toBeUndefined();
    // 其他色彩… is still the browser's picker.
    selectAll(d.view);
    button.click();
    await wait();
    const other = d.panel.querySelector('input[type="color"][aria-label="文字顏色：其他色彩"]') as HTMLInputElement;
    other.value = '#123456';
    other.dispatchEvent(new Event('change'));
    expect(color()).toBe('#123456');
    d.done();
  }, 20000);

  it('highlight: Word highlight colours by name, 無色彩 removes it; Escape closes the menu on its button', async () => {
    const d = await setup('<w:p><w:r><w:t>重點</w:t></w:r></w:p>', 'DocxToolbar.vue', { styles: [] });
    selectAll(d.view);
    await d.refresh();
    const button = d.panel.querySelector('button[aria-label^="螢光標記："]') as HTMLButtonElement;
    button.click();
    await wait();
    (d.panel.querySelector('.dx-cp-swatch[title="亮綠色"]') as HTMLButtonElement).click();
    const hl = () => schema.marks.highlight.isInSet(d.view.state.doc.child(0).child(0).marks)?.attrs.color;
    expect(hl()).toBe('#00ff00');
    button.click();
    await wait();
    const menu = d.panel.querySelector('.dx-cp-menu') as HTMLElement;
    menu.querySelector('button')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    await wait();
    expect(d.panel.querySelector('.dx-cp-menu')).toBeNull();
    selectAll(d.view);
    button.click();
    await wait();
    (Array.from(d.panel.querySelectorAll('.dx-cp-menu button')).find((b) => b.textContent === '無色彩') as HTMLButtonElement).click();
    expect(hl()).toBeUndefined();
    d.done();
  }, 20000);

  it('cell shading in the table tools uses the same menu', async () => {
    const TABLE =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const d = await setup(TABLE, 'TablePanel.vue');
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    (d.panel.querySelector('button[aria-label^="儲存格底色："]') as HTMLButtonElement).click();
    await wait();
    (d.panel.querySelector('.dx-cp-swatch[title="淺綠"]') as HTMLButtonElement).click();
    expect(d.view.state.doc.child(0).child(0).child(0).attrs.background).toBe('#92D050');
    d.done();
  }, 20000);
});
