// persona-300: Word's shortcuts. Ctrl+] / Ctrl+[ grow / shrink the font by 1 pt, Ctrl+Shift+> / <
// by one step of the size list, Ctrl+K inserts a link, Ctrl+D is not the browser's bookmark.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { schema } from '../../src/papyrus/editor/schema';
import { stepFontSize } from '../../src/papyrus/editor/commands';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { docxOf, domStubs, key, setup } from './p300Helpers';

domStubs();
afterEach(() => vi.restoreAllMocks());

const size = (view: any, i = 0) => schema.marks.fontSize.isInSet(view.state.doc.child(0).child(i).marks)?.attrs.pt;
function selectAll(view: any) {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.child(0).nodeSize - 1)));
}

describe('font size keys', () => {
  it('steps along Word’s size list', () => {
    expect([stepFontSize(12, 1), stepFontSize(12, -1), stepFontSize(10.5, 1), stepFontSize(72, 1), stepFontSize(80, -1), stepFontSize(8, -1)]).toEqual([
      14, 11, 11, 80, 72, 1,
    ]);
  });

  it('Ctrl+] / Ctrl+[ change each run by 1 pt; Ctrl+Shift+> / < by a step', async () => {
    const d = await setup('<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>十二</w:t></w:r><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>十六</w:t></w:r></w:p>');
    selectAll(d.view);
    key(d.view.dom, { key: ']', ctrlKey: true });
    expect([size(d.view, 0), size(d.view, 1)]).toEqual([13, 17]);
    key(d.view.dom, { key: '[', ctrlKey: true });
    key(d.view.dom, { key: '[', ctrlKey: true });
    expect([size(d.view, 0), size(d.view, 1)]).toEqual([11, 15]);
    key(d.view.dom, { key: '>', ctrlKey: true, shiftKey: true });
    expect([size(d.view, 0), size(d.view, 1)]).toEqual([12, 16]);
    key(d.view.dom, { key: '<', ctrlKey: true, shiftKey: true });
    expect([size(d.view, 0), size(d.view, 1)]).toEqual([11, 14]);
    d.done();
  });
});

describe('Ctrl+K and Ctrl+D', () => {
  it('Ctrl+K in the document asks for the link address; Ctrl+D is kept from the browser', async () => {
    const src = await docxOf('<w:p><w:r><w:t>公文系統</w:t></w:r></w:p>');
    const host = document.createElement('div');
    document.body.append(host);
    let editor: DocxEditorType | null = null;
    const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditorType) => (editor = ed) }) });
    app.mount(host);
    for (let i = 0; i < 200 && !editor; i++) {
      await new Promise((r) => setTimeout(r, 5));
      await nextTick();
    }
    await nextTick();
    const view = editor!.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
    const prompt = vi.spyOn(window, 'prompt').mockReturnValue('https://example.gov.tw/');
    const k = key(view.dom, { key: 'k', ctrlKey: true });
    expect(k.defaultPrevented).toBe(true);
    expect(prompt).toHaveBeenCalledTimes(1);
    expect(schema.marks.link.isInSet(view.state.doc.child(0).child(0).marks)?.attrs.href).toBe('https://example.gov.tw/');
    expect(key(view.dom, { key: 'd', ctrlKey: true }).defaultPrevented).toBe(true);
    app.unmount();
    host.remove();
  });
});
