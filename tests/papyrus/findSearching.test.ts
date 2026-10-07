// @vitest-environment jsdom
// run-188 (GOV-170): right after typing, the count said 「找不到」 for a moment and then 「1 / 1」,
// which reads like a miss. While the typing still owes a search, the panel says 「搜尋中…」.
import { afterEach, describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref, type App } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor, type EditorSnapshot } from '../../src/papyrus/editor/core';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
let app: App | null = null;
let editor: DocxEditor | null = null;
afterEach(() => { app?.unmount(); app = null; editor?.destroy(); editor = null; document.body.innerHTML = ''; });

describe('the find panel count while typing', () => {
  it('says 搜尋中… until the search runs, then the count', async () => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>會議日期 2026-10-05</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const host = document.createElement('div');
    const panelHost = document.createElement('div');
    document.body.append(host, panelHost);
    const snapshot = ref<EditorSnapshot | null>(null);
    editor = new DocxEditor(host, { onUpdate: (v) => { snapshot.value = v; } });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
    app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value, replace: false, canReplace: true, onClose: () => {} }) });
    app.mount(panelHost);
    await nextTick();
    const input = panelHost.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    input.value = '2026-10-05';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
    const count = () => panelHost.querySelector('.dx-count')?.textContent?.trim();
    expect(count()).toBe('搜尋中…');
    await new Promise((r) => setTimeout(r, 250));
    await nextTick();
    expect(count()).toBe('1 / 1');
  });

  it('closing with Escape leaves the found text selected, as in Word (GOV-113 / run-206)', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>第一段開頭</w:t></w:r></w:p><w:p><w:r><w:t>申請期限為十月五日</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const host = document.createElement('div');
    const panelHost = document.createElement('div');
    document.body.append(host, panelHost);
    const snapshot = ref<EditorSnapshot | null>(null);
    editor = new DocxEditor(host, { onUpdate: (v) => { snapshot.value = v; } });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
    let closed = false;
    app = createApp({ render: () => closed ? null : h(Panel, { editor, snapshot: snapshot.value, replace: false, canReplace: true, onClose: () => { closed = true; } }) });
    app.mount(panelHost);
    await nextTick();
    const input = panelHost.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    input.value = '十月五日';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    const { from, to } = editor.view!.state.selection;
    expect(editor.view!.state.doc.textBetween(from, to)).toBe('十月五日');
  });
});
