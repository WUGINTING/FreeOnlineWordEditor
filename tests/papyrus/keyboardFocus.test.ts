// @vitest-environment jsdom
// GOV-ISSUE-013: closing 尋找與取代 left the keyboard focus on the page (BODY), and the text was
// 43 Tab presses away, behind every toolbar button.
import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

async function mount() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>第一段 2025 年度</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const src = await zip.generateAsync({ type: 'uint8array' });
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditorType | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditorType) => (editor = ed) }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  return { host, editor: editor as DocxEditorType, flush, done: () => { app.unmount(); host.remove(); } };
}

describe('keyboard focus in the editor', () => {
  it('closing 尋找與取代 with Escape puts the focus back in the document', async () => {
    const ui = await mount();
    ui.editor.view!.focus();
    ui.editor.view!.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'h', ctrlKey: true, bubbles: true }));
    await ui.flush();
    const find = ui.host.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    expect(find).not.toBeNull();
    expect(document.activeElement).toBe(find);
    find.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await ui.flush();
    expect(ui.host.querySelector('input[aria-label="尋找"]')).toBeNull();
    expect(document.activeElement).toBe(ui.editor.view!.dom);
    ui.done();
  });

  it('the first Tab stop is 「跳到文件內容」, which goes straight to the text', async () => {
    const ui = await mount();
    const skip = ui.host.querySelector('a.dx-skip') as HTMLAnchorElement;
    expect(skip).not.toBeNull();
    expect(skip.textContent).toBe('跳到文件內容');
    // Before the toolbar in document (Tab) order.
    const toolbar = ui.host.querySelector('.dx-toolbar, [role="toolbar"]');
    if (toolbar) expect(skip.compareDocumentPosition(toolbar) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    skip.click();
    await ui.flush();
    expect(document.activeElement).toBe(ui.editor.view!.dom);
    ui.done();
  });
});
