// persona-300 B-10: the zoom a user picks is remembered; the automatic fit never undoes a
// browser zoom; the ribbon and the comments panel follow the browser's font size (rem) and
// their buttons are at least 24 px high.
import { afterEach, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { createApp, h, nextTick } from 'vue';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { docxOf, domStubs } from './p300Helpers';

domStubs();
afterEach(() => {
  localStorage.clear();
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 1 });
  Object.defineProperty(window, 'outerWidth', { configurable: true, value: window.innerWidth });
});

async function mount() {
  const src = await docxOf('<w:p><w:r><w:t>內容</w:t></w:r></w:p>');
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditorType | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditorType) => (editor = ed) }) });
  app.mount(host);
  const flush = async () => {
    for (let i = 0; i < 7; i++) {
      await new Promise((r) => setTimeout(r, 0));
      await nextTick();
    }
  };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  const area = host.querySelector('.dx-scroll > div') as HTMLElement;
  let width = 0;
  Object.defineProperty(area, 'clientWidth', { configurable: true, get: () => width });
  const resize = async (w: number) => {
    width = w;
    window.dispatchEvent(new Event('resize'));
    await flush();
  };
  const select = host.querySelector('select[aria-label="縮放"]') as HTMLSelectElement;
  return { editor: editor!, select, resize, flush, done: () => (app.unmount(), host.remove()) };
}

describe('zoom', () => {
  it('the zoom picked is used again the next time a document opens', async () => {
    const a = await mount();
    a.select.value = '1.5';
    a.select.dispatchEvent(new Event('change', { bubbles: true }));
    await a.flush();
    expect(a.editor.zoom).toBe(1.5);
    a.done();
    expect(localStorage.getItem('papyrus.zoom')).toBe('1.5');
    const b = await mount();
    expect(b.editor.zoom).toBe(1.5);
    expect(b.select.value).toBe('1.5');
    // Picked: a narrow window does not change it.
    await b.resize(400);
    expect(b.editor.zoom).toBe(1.5);
    b.done();
  });

  it('a stored value that is not a zoom is ignored', async () => {
    localStorage.setItem('papyrus.zoom', '<script>');
    const a = await mount();
    await a.resize(1200);
    expect(a.editor.zoom).toBe(1);
    a.done();
  });

  it('without a choice, a window narrowed by zooming the browser in keeps the page at 100 %', async () => {
    const a = await mount();
    await a.resize(500);
    expect(a.editor.zoom).toBeLessThan(1); // a narrow window: fitted
    Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 }); // browser at 200 %
    await a.resize(501);
    expect(a.editor.zoom).toBe(1);
    expect(a.select.value).toBe('1');
    a.done();
  });

  it('a window narrowed by DevTools or a sidebar (outer width far above the page width) is still fitted', async () => {
    const a = await mount();
    Object.defineProperty(window, 'outerWidth', { configurable: true, value: window.innerWidth * 2 });
    await a.resize(500);
    expect(a.editor.zoom).toBeLessThan(1);
    expect(a.select.value).toBe('fit');
    a.done();
  });
});

describe('sizes follow the browser font size', () => {
  const style = (file: string) => readFileSync(`src/papyrus/vue/${file}`, 'utf8').split('<style scoped>')[1];
  const rule = (css: string, sel: string) => new RegExp(`(^|\\n)${sel.replace(/[.[\]]/g, '\\$&')} \\{([^}]*)\\}`).exec(css)?.[2] ?? '';
  const rem = (decl: string, prop: string) => Number(new RegExp(`(?:^|\\s|;)${prop}:\\s*([\\d.]+)rem`).exec(decl)?.[1] ?? NaN);

  it('ribbon buttons are sized in rem, small ones at least 1.5 rem (24 px)', () => {
    const css = style('DocxToolbar.vue');
    expect(rem(rule(css, '.dx-sm'), 'height')).toBeGreaterThanOrEqual(1.5);
    expect(rem(rule(css, '.dx-sm'), 'font-size')).toBeGreaterThanOrEqual(0.75);
    expect(rem(rule(css, '.dx-big'), 'height')).toBe(4);
    expect(rem(rule(css, 'button.dx-launch'), 'height')).toBeGreaterThanOrEqual(1.5);
    expect(rule(css, '.dx-ribbon')).toMatch(/font: 0\.8125rem/);
  });

  it('comments panel buttons are at least 24 px high with 14 px text', () => {
    const css = style('CommentsPanel.vue');
    expect(rule(css, '.dx-comments')).toMatch(/font: 0\.875rem/);
    for (const sel of ['.dx-comments-nav button', '.dx-comment-actions button', '.dx-comments-new', '.dx-comments-close']) {
      expect(rem(rule(css, sel), 'min-height')).toBeGreaterThanOrEqual(1.5);
      expect(rule(css, sel)).not.toMatch(/font-size:\s*1[0-3]px/);
    }
  });
});
