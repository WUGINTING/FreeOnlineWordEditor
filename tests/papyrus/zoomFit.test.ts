// @vitest-environment jsdom
// GOV-FINDING-025: in a narrow window the page (about 800 px) was wider than the space for it
// (about 300 px) and had to be scrolled sideways. Until the user picks a zoom, DocxEditor.vue
// shows the page smaller so it fits (「符合寬度」), never below 50 %, and back at 100 % once it
// fits. Only the view is scaled: the pages and page breaks stay the same.
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
  const paras = Array.from({ length: 5 }, (_, i) => `<w:p><w:r><w:t>第 ${i + 1} 段</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${paras}<w:p><w:r><w:br w:type="page"/><w:t>第二頁</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>`);
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
  // The element the editor is mounted in: the space for the page (jsdom has no layout).
  const area = host.querySelector('.dx-scroll > div') as HTMLElement;
  let width = 0;
  Object.defineProperty(area, 'clientWidth', { configurable: true, get: () => width });
  const resize = async (w: number) => {
    width = w;
    window.dispatchEvent(new Event('resize'));
    await flush();
  };
  const select = host.querySelector('select[aria-label="縮放"]') as HTMLSelectElement;
  const ed = editor as DocxEditorType;
  const pageWidth = () => Math.max(...ed.pages.map((p) => p.left + p.width));
  return { host, editor: ed, select, resize, flush, pageWidth, done: () => { app.unmount(); host.remove(); } };
}

describe('GOV-FINDING-025: the page fits a narrow window', () => {
  it('a wide window: nothing changes (100 %)', async () => {
    const ui = await mount();
    expect(ui.editor.pages.length).toBeGreaterThan(0);
    await ui.resize(1200);
    expect(ui.editor.zoom).toBe(1);
    expect(ui.select.value).toBe('1');
    ui.done();
  });

  it('a narrow window: 「符合寬度」, recomputed on resize, never below 50 %, back to 100 % when it fits', async () => {
    const ui = await mount();
    const page = ui.pageWidth();
    expect(page).toBeGreaterThan(700); // A4, about 794 px
    await ui.resize(600);
    expect(ui.editor.zoom).toBeCloseTo((600 - 32) / page, 6);
    expect(ui.select.value).toBe('fit');
    const fit = Array.from(ui.select.options).find((o) => o.value === 'fit')!;
    expect(fit.textContent).toBe(`符合寬度（${Math.round(((600 - 32) / page) * 100)}%）`);
    await ui.resize(360); // a phone: 50 % at least, the rest scrolls
    expect(ui.editor.zoom).toBe(0.5);
    expect(ui.select.value).toBe('fit');
    await ui.resize(1280);
    expect(ui.editor.zoom).toBe(1);
    expect(ui.select.value).toBe('1');
    ui.done();
  });

  it('only scales the view: the same pages and page breaks, the canvas at 100 %', async () => {
    const ui = await mount();
    await ui.resize(1200);
    const before = JSON.stringify(ui.editor.pages);
    const canvas = ui.host.querySelector('.dx-canvas') as HTMLElement;
    const stage = ui.host.querySelector('.dx-stage') as HTMLElement;
    const canvasWidth = canvas.style.width;
    await ui.resize(500);
    const z = ui.editor.zoom;
    expect(z).toBeLessThan(1);
    expect(JSON.stringify(ui.editor.pages)).toBe(before);
    expect(ui.editor.snapshot()?.pageCount).toBe(ui.editor.pages.length);
    expect(canvas.style.width).toBe(canvasWidth);
    expect(canvas.style.transform).toBe(`scale(${z})`);
    expect(stage.style.width).toBe(`${parseFloat(canvasWidth) * z}px`);
    ui.done();
  });

  it('a zoom the user picked is kept, whatever the window does', async () => {
    const ui = await mount();
    ui.select.value = '0.75';
    ui.select.dispatchEvent(new Event('change', { bubbles: true }));
    await ui.flush();
    await ui.resize(400);
    expect(ui.editor.zoom).toBe(0.75);
    expect(ui.select.value).toBe('0.75');
    await ui.resize(1400);
    expect(ui.editor.zoom).toBe(0.75);
    ui.done();
  });

  it('choosing 「符合寬度」 still asks the editor to fit the page width', async () => {
    const ui = await mount();
    await ui.resize(900);
    ui.select.value = 'fit';
    ui.select.dispatchEvent(new Event('change', { bubbles: true }));
    await ui.flush();
    // The editor's own fit mode (core.ts), measured on its root: not changed by this.
    expect(ui.select.value).toBe('fit');
    await ui.resize(300);
    expect(ui.select.value).toBe('fit');
    ui.done();
  });
});
