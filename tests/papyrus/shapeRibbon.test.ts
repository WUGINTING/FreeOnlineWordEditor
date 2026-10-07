// @vitest-environment jsdom
// The ribbon's shape tools (phase 8B): 插入 › 圖案 (flowchart symbols, arrows, lines) and 文字方塊,
// 常用 › 選取圖形, and the 圖形格式 tab while a shape is selected (fill, outline, 文繞圖, stacking,
// align, group, position and size in cm).
import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { placeShapes } from '../../src/papyrus/editor/shapeView';
import type { ShapeModel, SpNode } from '../../src/papyrus/docx/shapes';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => vi.restoreAllMocks());

async function mount(file: string) {
  const src = readFileSync(`tests/fixtures/shapes/${file}.docx`);
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditorType | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditorType) => (editor = ed) }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  const tabs = () => Array.from(host.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
  const tab = (label: string) => tabs().find((t) => t.textContent?.trim() === label);
  const button = (text: string) => Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((b) => b.textContent?.trim() === text);
  return { host, editor: editor as DocxEditorType, flush, tab, button, done: () => { app.unmount(); host.remove(); } };
}

const shapeNamed = (ed: DocxEditorType, name: string) => {
  let out: ShapeModel | null = null;
  ed.view!.state.doc.descendants((n) => {
    if (!out && n.attrs?.shape?.name === name) out = n.attrs.shape;
    return !out;
  });
  return out as ShapeModel | null;
};

describe('the ribbon’s shape tools', () => {
  it('插入 › 圖案 lists the flowchart symbols, arrows and lines; 文字方塊 and each kind start placing', async () => {
    const m = await mount('textboxes');
    m.tab('插入')!.click();
    await m.flush();
    m.button('圖案 ▾')!.click();
    await m.flush();
    const items = Array.from(m.host.querySelectorAll<HTMLButtonElement>('.dx-shape-menu [role="menuitem"]')).map((b) => b.textContent?.trim());
    expect(items).toEqual(expect.arrayContaining(['程序', '決策', '開始或結束', '資料', '文件', '預設程序', '向右箭號', '直線', '箭號（直線接點）', '肘形接點（箭號）']));
    const start = vi.spyOn(m.editor, 'startInsertShape');
    Array.from(m.host.querySelectorAll<HTMLButtonElement>('.dx-shape-menu [role="menuitem"]')).find((b) => b.textContent?.trim() === '決策')!.click();
    expect(start).toHaveBeenCalledWith('decision');
    m.button('文字方塊')!.click();
    expect(start).toHaveBeenCalledWith('textbox');
    // Placing waits for the page; Esc cancels it.
    expect(m.host.querySelector('.dx-shape-place-cover')).toBeTruthy();
    m.editor.view!.someProp('handleKeyDown', (f) => f(m.editor.view!, new KeyboardEvent('keydown', { key: 'Escape' })));
    expect(m.host.querySelector('.dx-shape-place-cover')).toBeNull();
    m.done();
  });

  it('圖形格式 appears with a selected shape (選取圖形) and changes it', async () => {
    const m = await mount('shapes');
    expect(m.tab('圖形格式')).toBeUndefined();
    m.tab('插入')!.click();
    await m.flush();
    m.button('選取圖形')!.click();
    await m.flush();
    placeShapes(m.editor.view!.dom);
    // Tab to the plain rectangle.
    for (let i = 0; i < 20 && !m.editor.shapes.info()?.name.includes('矩形 1'); i++) {
      m.editor.view!.someProp('handleKeyDown', (f) => f(m.editor.view!, new KeyboardEvent('keydown', { key: 'Tab' })));
      placeShapes(m.editor.view!.dom);
    }
    await m.flush();
    expect(m.tab('圖形格式')).toBeTruthy();
    m.tab('圖形格式')!.click();
    await m.flush();
    const panel = m.host.querySelector('.dx-shape-panel')!;
    expect(panel.textContent).toContain('矩形 1');
    const wrap = panel.querySelector<HTMLSelectElement>('select[aria-label="文繞圖"]')!;
    expect(wrap.value).toBe('front');
    wrap.value = 'square';
    wrap.dispatchEvent(new Event('change'));
    await m.flush();
    expect(shapeNamed(m.editor, '矩形 1')!.anchor!.wrap).toBe('square');
    placeShapes(m.editor.view!.dom);
    await m.flush();
    const width = panel.ownerDocument.querySelector<HTMLInputElement>('.dx-shape-panel input[aria-label="寬度（公分）"]')!;
    width.value = '5';
    width.dispatchEvent(new Event('input'));
    width.dispatchEvent(new Event('change'));
    await m.flush();
    expect(shapeNamed(m.editor, '矩形 1')!.w).toBe(5 * 360000);
    const weight = m.host.querySelector<HTMLSelectElement>('.dx-shape-panel select[aria-label="外框粗細（點）"]')!;
    weight.value = '3';
    weight.dispatchEvent(new Event('change'));
    await m.flush();
    expect((shapeNamed(m.editor, '矩形 1')!.root as SpNode).line!.width).toBe(3 * 12700);
    expect(m.host.querySelector<HTMLButtonElement>('.dx-shape-panel button[title^="群組"]')!.disabled).toBe(true);
    m.done();
  });
});
