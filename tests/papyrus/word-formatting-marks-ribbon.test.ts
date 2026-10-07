// @vitest-environment jsdom
// 顯示／隱藏編輯標記 in the ribbon (常用 › 段落 ¶). The button toggles the marks and shows its state,
// and Ctrl+Shift+8 works with the focus on the ribbon too, as in Word.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => { try { localStorage.clear(); } catch { /* storage unavailable */ } });

async function mount() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>一段 文字</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const src = await zip.generateAsync({ type: 'uint8array' });
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditor) => (editor = ed) }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  const button = host.querySelector('[role="toolbar"] button[aria-label="顯示／隱藏編輯標記"]') as HTMLButtonElement;
  return { host, button, editor: editor as DocxEditor, flush, done: () => { app.unmount(); host.remove(); } };
}

describe('顯示／隱藏編輯標記: ribbon', () => {
  it('the ¶ button toggles the marks and aria-pressed follows', async () => {
    const ui = await mount();
    expect(ui.button).not.toBeNull();
    expect(ui.button.getAttribute('aria-pressed')).toBe('false');
    ui.button.click();
    await ui.flush();
    expect(ui.editor.showMarks).toBe(true);
    expect(ui.button.getAttribute('aria-pressed')).toBe('true');
    ui.button.click();
    await ui.flush();
    expect(ui.editor.showMarks).toBe(false);
    expect(ui.button.getAttribute('aria-pressed')).toBe('false');
    ui.done();
  });

  it('Ctrl+Shift+8 with the focus on the ribbon toggles once', async () => {
    const ui = await mount();
    ui.button.focus();
    ui.button.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    await ui.flush();
    expect(ui.editor.showMarks).toBe(true);
    ui.button.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    await ui.flush();
    expect(ui.editor.showMarks).toBe(false);
    ui.done();
  });

  it('Ctrl+Shift+8 in the text toggles once, not twice', async () => {
    const ui = await mount();
    const text = ui.editor.view!.dom;
    text.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    await ui.flush();
    expect(ui.editor.showMarks).toBe(true);
    ui.done();
  });
});
