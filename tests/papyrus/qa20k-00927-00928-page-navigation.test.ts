import { beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} })) as any;
});

async function mount() {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src: null, onReady: (instance: DocxEditor) => { editor = instance; } }) });
  app.mount(host);
  for (let i = 0; i < 80 && !editor; i++) { await new Promise((resolve) => setTimeout(resolve, 0)); await nextTick(); }
  if (!editor) throw new Error('DocxEditor did not become ready');
  (editor as any).pageCount = 3;
  editor.setZoom(1); // publish the synthetic page count through the public update channel
  for (let i = 0; i < 5; i++) await nextTick();
  const input = host.querySelector<HTMLInputElement>('.dx-page-input')!;
  const scroll = vi.spyOn(editor, 'scrollToPage').mockImplementation(() => {});
  return { host, app, editor, input, scroll, done: () => { app.unmount(); host.remove(); } };
}

describe('page navigation in the status bar', () => {
  it('QA20K-00927 jumps to the requested one-based page number', async () => {
    const x = await mount();
    try {
      expect(x.host.querySelector('.dx-status [aria-live="polite"]')!.textContent).toContain('共 3 頁');
      x.input.value = '2'; x.input.dispatchEvent(new Event('change', { bubbles: true }));
      expect(x.scroll).toHaveBeenCalledTimes(1);
      expect(x.scroll).toHaveBeenCalledWith(1); // UI page 2 maps to zero-based page index 1
    } finally { x.done(); }
  });

  it('QA20K-00928 ignores a page number outside the available range', async () => {
    const x = await mount();
    try {
      x.input.value = '4'; x.input.dispatchEvent(new Event('change', { bubbles: true }));
      expect(x.scroll).not.toHaveBeenCalled();
      x.input.value = '0'; x.input.dispatchEvent(new Event('change', { bubbles: true }));
      expect(x.scroll).not.toHaveBeenCalled();
    } finally { x.done(); }
  });
});
