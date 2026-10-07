// persona-300 code review, the controls: the colour menus' Escape and names, typing with an input
// method next to an English run, Ctrl+K with the ribbon hidden, Ctrl+Shift+. / , by the key's place.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { schema } from '../../src/papyrus/editor/schema';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { docxOf, domStubs, key, setup } from './p300Helpers';

domStubs();
afterEach(() => vi.restoreAllMocks());

const selectAll = (view: any) => view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1)));

describe('colour menus', () => {
  it('Escape with the menu closed is left alone; the button names its colour and opens a dialog', async () => {
    const d = await setup('<w:p><w:r><w:t>主旨</w:t></w:r></w:p>', 'DocxToolbar.vue', { styles: [] });
    selectAll(d.view);
    await d.refresh();
    const button = () => d.panel.querySelector('button[aria-label^="文字顏色："]') as HTMLButtonElement;
    expect(button().getAttribute('aria-label')).toBe('文字顏色：自動');
    expect(button().getAttribute('aria-haspopup')).toBe('dialog');
    expect(button().getAttribute('aria-expanded')).toBe('false');
    // It goes on to the ribbon (which takes the keyboard back to the text).
    let reachedRibbon = false;
    const around = button().closest('.dx-cp')!.parentElement!;
    const listener = () => (reachedRibbon = true);
    around.addEventListener('keydown', listener);
    key(button(), { key: 'Escape' });
    expect(reachedRibbon).toBe(true);
    around.removeEventListener('keydown', listener);
    button().click();
    await nextTick();
    const menu = d.panel.querySelector('.dx-cp-menu') as HTMLElement;
    expect(menu.getAttribute('role')).toBe('dialog');
    expect(button().getAttribute('aria-controls')).toBe(menu.id);
    // Escape of an input method choosing a character does not close it; a real Escape does.
    key(menu.querySelector('button')!, { key: 'Escape', ime: true });
    await nextTick();
    expect(d.panel.querySelector('.dx-cp-menu')).not.toBeNull();
    (d.panel.querySelector('.dx-cp-swatch[title="深藍"]') as HTMLButtonElement).click();
    await d.refresh();
    expect(button().getAttribute('aria-label')).toBe('文字顏色：深藍');
    d.done();
  }, 20000);
});

describe('an input method next to an English run', () => {
  const EN = '<w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Hello</w:t></w:r></w:p>';
  const langSpan = (d: any) => d.view.dom.querySelector('[lang="en-US"]') as HTMLElement | null;

  it('while composing, the lang stays around the text being composed; it is worked out again after', async () => {
    const d = await setup(EN);
    expect(langSpan(d)!.textContent).toBe('Hello');
    const end = d.view.state.doc.child(0).nodeSize - 1;
    d.view.dispatch(d.view.state.tr.insertText('ㄓ', end).setMeta('composition', 1));
    expect(langSpan(d)?.textContent).toBe('Helloㄓ');
    d.view.dispatch(d.view.state.tr.insertText('中', end, end + 1).setMeta('composition', 1));
    expect(langSpan(d)?.textContent).toBe('Hello中');
    // The composition is over: Chinese text is not English.
    await new Promise((r) => setTimeout(r, 80));
    expect(langSpan(d)).toBeNull();
    // No undo step of its own: one undo takes the typing back.
    expect(d.view.state.doc.textContent).toBe('Hello中');
    d.done();
  });

  it('typed without an input method, it is worked out at once', async () => {
    const d = await setup(EN);
    d.view.dispatch(d.view.state.tr.insertText('中', d.view.state.doc.child(0).nodeSize - 1));
    expect(langSpan(d)).toBeNull();
    d.done();
  });
});

describe('keys', () => {
  it('Ctrl+K asks for the link with the ribbon hidden too', async () => {
    const src = await docxOf('<w:p><w:r><w:t>公文系統</w:t></w:r></w:p>');
    const host = document.createElement('div');
    document.body.append(host);
    let editor: DocxEditorType | null = null;
    const app = createApp({ render: () => h(DocxEditorVue, { src, toolbar: false, onReady: (ed: DocxEditorType) => (editor = ed) }) });
    app.mount(host);
    for (let i = 0; i < 200 && !editor; i++) {
      await new Promise((r) => setTimeout(r, 5));
      await nextTick();
    }
    await nextTick();
    expect(host.querySelector('.dx-ribbon')).toBeNull();
    const view = editor!.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
    const prompt = vi.spyOn(window, 'prompt').mockReturnValue('https://example.gov.tw/');
    expect(key(view.dom, { key: 'k', ctrlKey: true }).defaultPrevented).toBe(true);
    expect(prompt).toHaveBeenCalledTimes(1);
    expect(schema.marks.link.isInSet(view.state.doc.child(0).child(0).marks)?.attrs.href).toBe('https://example.gov.tw/');
    app.unmount();
    host.remove();
  });

  it('Ctrl+Shift+. / , work by the key’s place where Shift+. is not ">" (another keyboard layout)', async () => {
    const d = await setup('<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>十二</w:t></w:r></w:p>');
    selectAll(d.view);
    const size = () => schema.marks.fontSize.isInSet(d.view.state.doc.child(0).child(0).marks)?.attrs.pt;
    key(d.view.dom, { key: ':', keyCode: 190, ctrlKey: true, shiftKey: true } as any);
    expect(size()).toBe(14);
    key(d.view.dom, { key: ';', keyCode: 188, ctrlKey: true, shiftKey: true } as any);
    expect(size()).toBe(12);
    d.done();
  });
});
