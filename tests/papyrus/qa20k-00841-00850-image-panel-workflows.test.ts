import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import type { DocxEditor as DocxEditorType, EditorSnapshot } from '../../src/papyrus/editor/core';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const VUE = 'src/papyrus/vue/';
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
const HEIC = btoa('\x00\x00\x00\x18ftypheic\x00\x00\x00\x00mif1heic');

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });

async function setup(body = '<w:p/>') {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { model } = await readDocx(bytes);
  const host = document.createElement('div');
  const panelHost = document.createElement('div');
  document.body.append(host, panelHost);
  const snapshot = ref<EditorSnapshot | null>(null);
  const editor = new DocxEditor(host, { onUpdate: (s) => { snapshot.value = s; } });
  await editor.open(bytes);
  const Panel = await loadSfc(VUE + 'ImagePanel.vue');
  const app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value }) });
  app.mount(panelHost);
  const refresh = async () => { snapshot.value = editor.snapshot(); await nextTick(); };
  const done = () => { app.unmount(); editor.destroy(); host.remove(); panelHost.remove(); };
  return { editor, model, panel: panelHost, refresh, done };
}

function textStart(editor: DocxEditorType, text: string) {
  let start = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (start < 0 && node.isText && node.text!.includes(text)) start = pos + node.text!.indexOf(text);
  });
  if (start < 0) throw new Error(`missing text: ${text}`);
  return start;
}

function addImage(editor: DocxEditorType, pos: number, attrs: Record<string, unknown>) {
  const image = schema.nodes.image.create(attrs);
  const tr = editor.view!.state.tr.insert(pos, image);
  editor.view!.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
  editor.view!.dispatch(closeHistory(editor.view!.state.tr));
}

function imageEntries(editor: DocxEditorType) {
  const out: { pos: number; node: any }[] = [];
  editor.view!.state.doc.descendants((node, pos) => { if (node.type === schema.nodes.image) out.push({ pos, node }); });
  return out;
}

function textInput(panel: HTMLElement, label: string, value: string) {
  const box = panel.querySelector(`input[aria-label="${label}"]`) as HTMLInputElement;
  if (!box) throw new Error(`missing input ${label}`);
  box.value = value;
  box.dispatchEvent(new Event('input', { bubbles: true }));
  box.dispatchEvent(new Event('change', { bubbles: true }));
  return box;
}

function choosePx(panel: HTMLElement) {
  const unit = panel.querySelector('select[aria-label="尺寸單位"]') as HTMLSelectElement;
  unit.value = 'px'; unit.dispatchEvent(new Event('change', { bubbles: true }));
}

const fileFrom = (base64: string, name: string, type: string) => new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], name, { type });

function fakeImage(width: number, height: number, broken: string[] = []) {
  vi.stubGlobal('Image', class {
    naturalWidth = 0; naturalHeight = 0;
    onload: (() => void) | null = null; onerror: (() => void) | null = null;
    set src(value: string) {
      setTimeout(() => {
        if (broken.some((type) => value.startsWith(`data:${type}`))) this.onerror?.();
        else { this.naturalWidth = width; this.naturalHeight = height; this.onload?.(); }
      });
    }
  });
}

describe('QA20K-00841–00850 image editing workflows', () => {
  it('QA20K-00841 edits descriptive alt text from the picture panel and saves it', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: '' });
    await d.refresh();
    textInput(d.panel, '替代文字', '簽核流程圖，顯示主管核准步驟');
    await d.refresh();
    const image = imageEntries(d.editor)[0];
    expect(image.node.attrs.alt).toBe('簽核流程圖，顯示主管核准步驟');
    expect(d.editor.view!.dom.querySelector('img.dx-img')!.alt).toBe('簽核流程圖，顯示主管核准步驟');
    const zip = await JSZip.loadAsync(await writeDocx(d.editor.view!.state.doc, d.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('descr="簽核流程圖，顯示主管核准步驟"');
    d.done();
  });

  it('QA20K-00842 changing selection switches the panel to the selected picture before editing its description', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 80, height: 40, alt: '公司標誌' });
    addImage(d.editor, 2, { src: `data:image/gif;base64,${GIF}`, width: 90, height: 60, alt: '季度圖表' });
    const before = imageEntries(d.editor);
    expect(before).toHaveLength(2);
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(NodeSelection.create(d.editor.view!.state.doc, before[1].pos)));
    await d.refresh();
    const alt = d.panel.querySelector('input[aria-label="替代文字"]') as HTMLInputElement;
    expect(alt.value).toBe('季度圖表');
    textInput(d.panel, '替代文字', '第二季營收比較圖');
    await d.refresh();
    expect(imageEntries(d.editor).map(({ node }) => node.attrs.alt)).toEqual(['公司標誌', '第二季營收比較圖']);
    d.done();
  });

  it('QA20K-00843 turning off the proportion lock changes only the entered width', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: 'chart' });
    await d.refresh();
    choosePx(d.panel); await d.refresh();
    const lock = d.panel.querySelector('.dx-check input') as HTMLInputElement;
    expect(lock.checked).toBe(true);
    lock.click();
    textInput(d.panel, '圖片寬度', '160');
    await d.refresh();
    expect(imageEntries(d.editor)[0].node.attrs).toMatchObject({ width: 160, height: 50, alt: 'chart' });
    d.done();
  });

  it('QA20K-00844 editing height with the default proportion lock recalculates width', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: 'chart' });
    await d.refresh();
    choosePx(d.panel); await d.refresh();
    expect((d.panel.querySelector('.dx-check input') as HTMLInputElement).checked).toBe(true);
    textInput(d.panel, '圖片高度', '100');
    await d.refresh();
    expect(imageEntries(d.editor)[0].node.attrs).toMatchObject({ width: 200, height: 100 });
    d.done();
  });

  it('QA20K-00845 invalid nonpositive size input restores the current displayed size', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: 'chart' });
    await d.refresh();
    const original = { ...imageEntries(d.editor)[0].node.attrs };
    const box = textInput(d.panel, '圖片寬度', '0');
    await d.refresh();
    expect(imageEntries(d.editor)[0].node.attrs).toEqual(original);
    expect(box.value).toBe('2.65');
    d.done();
  });

  it('QA20K-00846 dragging a resize handle below the minimum clamps width without shrinking below eight pixels', async () => {
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: '' });
    const image = d.panel.parentElement!.querySelector('.dx-img-wrap') as HTMLElement;
    const east = image.querySelector('[data-handle="e"]') as HTMLElement;
    east.dispatchEvent(new MouseEvent('pointerdown', { clientX: 0, clientY: 0, bubbles: true, cancelable: true }));
    window.dispatchEvent(new MouseEvent('pointermove', { clientX: -500, clientY: 0 }));
    window.dispatchEvent(new MouseEvent('pointerup'));
    expect(imageEntries(d.editor)[0].node.attrs).toMatchObject({ width: 8, height: 50 });
    d.done();
  });

  it('QA20K-00847 inserts a selected image at the text caret and scales it to page content width', async () => {
    fakeImage(1200, 600);
    const d = await setup('<w:p><w:r><w:t xml:space="preserve">Use chart now</w:t></w:r></w:p>');
    const start = textStart(d.editor, 'Use chart now');
    const caret = start + 'Use '.length;
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(TextSelection.create(d.editor.view!.state.doc, caret)));
    await d.editor.insertImageFile(fileFrom(PNG, 'wide.png', 'image/png'));
    const paragraph = d.editor.view!.state.doc.firstChild!;
    expect(paragraph.childCount).toBe(3);
    expect(paragraph.child(0).textContent).toBe('Use ');
    expect(paragraph.child(1).type).toBe(schema.nodes.image);
    expect(paragraph.child(2).textContent).toBe('chart now');
    const max = d.editor.contentWidth;
    const width = Math.min(1200, max);
    expect(paragraph.child(1).attrs.width).toBe(Math.round(width));
    expect(paragraph.child(1).attrs.height).toBe(Math.round(600 * width / 1200));
    d.done();
  });

  it('QA20K-00848 replacing a selected picture through the file control keeps its width and uses the new aspect ratio', async () => {
    fakeImage(80, 20);
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 120, height: 60, alt: 'approved logo' });
    await d.refresh();
    const picker = d.panel.querySelector('input[type="file"]') as HTMLInputElement;
    Object.defineProperty(picker, 'files', { configurable: true, value: [fileFrom(GIF, 'new-banner.gif', 'image/gif')] });
    picker.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    await d.refresh();
    expect(imageEntries(d.editor)[0].node.attrs).toMatchObject({
      src: `data:image/gif;base64,${GIF}`, width: 120, height: 30, alt: 'approved logo',
    });
    expect(picker.value).toBe('');
    d.done();
  });

  it('QA20K-00849 unsupported replacement displays an error and leaves the current picture intact', async () => {
    fakeImage(0, 0, ['image/heic']);
    const d = await setup();
    addImage(d.editor, 1, { src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: 'current' });
    await d.refresh();
    const before = { ...imageEntries(d.editor)[0].node.attrs };
    const picker = d.panel.querySelector('input[type="file"]') as HTMLInputElement;
    Object.defineProperty(picker, 'files', { configurable: true, value: [fileFrom(HEIC, 'scan.heic', 'image/heic')] });
    picker.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    await d.refresh();
    expect(d.panel.querySelector('[role="alert"]')?.textContent).toContain('PNG、JPEG、GIF 或 BMP');
    expect(imageEntries(d.editor)[0].node.attrs).toEqual(before);
    expect(picker.value).toBe('');
    d.done();
  });

  it('QA20K-00850 deleting a picture and undoing restores it between the original surrounding words', async () => {
    const d = await setup('<w:p><w:r><w:t xml:space="preserve">Before picture after</w:t></w:r></w:p>');
    const start = textStart(d.editor, 'Before picture after');
    const pos = start + 'Before '.length;
    addImage(d.editor, pos, { src: `data:image/png;base64,${PNG}`, width: 75, height: 50, alt: 'receipt' });
    await d.refresh();
    (d.panel.querySelector('button[title="刪除圖片"]') as HTMLButtonElement).click();
    expect(imageEntries(d.editor)).toHaveLength(0);
    expect(d.editor.view!.state.doc.textContent).toBe('Before picture after');
    expect(d.editor.undo()).toBe(true);
    const paragraph = d.editor.view!.state.doc.firstChild!;
    expect(paragraph.childCount).toBe(3);
    expect(paragraph.child(0).textContent).toBe('Before ');
    expect(paragraph.child(1).type).toBe(schema.nodes.image);
    expect(paragraph.child(1).attrs).toMatchObject({ width: 75, height: 50, alt: 'receipt' });
    expect(paragraph.child(2).textContent).toBe('picture after');
    d.done();
  });
});
