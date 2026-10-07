// @vitest-environment jsdom
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick, ref, type App } from 'vue';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor, type EditorSnapshot } from '../src/papyrus/editor/core';
import { searchState } from '../src/papyrus/editor/search';
import { loadSfc } from './papyrus/sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

let app: App | null = null;
let editor: DocxEditor | null = null;
let editorHost: HTMLElement | null = null;
let panelHost: HTMLElement | null = null;

function installLayoutShims() {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
}

async function mountPanel(body: string, options: { replace?: boolean; canReplace?: boolean } = {}) {
  installLayoutShims();
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p>${run(body)}</w:p>${SECT}</w:body></w:document>`);
  editorHost = document.createElement('div');
  panelHost = document.createElement('div');
  document.body.append(editorHost, panelHost);
  const snapshot = ref<EditorSnapshot | null>(null);
  editor = new DocxEditor(editorHost, { onUpdate: (value) => { snapshot.value = value; } });
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
  const closed = vi.fn();
  const visible = ref(true);
  app = createApp({ render: () => visible.value ? h(Panel, {
    editor, snapshot: snapshot.value, replace: options.replace ?? false,
    canReplace: options.canReplace ?? true, onClose: () => { closed(); visible.value = false; },
  }) : null });
  app.mount(panelHost);
  await nextTick();
  return { closed, input: (name = '尋找') => panelHost!.querySelector(`input[aria-label="${name}"]`) as HTMLInputElement | null };
}

async function typeInto(input: HTMLInputElement, value: string) {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  await nextTick();
}

async function settleSearch() {
  await new Promise((resolve) => setTimeout(resolve, 180));
  await nextTick();
}

function marks() { return [...editor!.view!.dom.querySelectorAll<HTMLElement>('.dx-search-hit')]; }
function count() { return panelHost!.querySelector('.dx-count')?.textContent?.trim() ?? ''; }
function text() { return editor!.view!.state.doc.textContent; }

afterEach(() => {
  app?.unmount(); app = null;
  editor?.destroy(); editor = null;
  editorHost?.remove(); panelHost?.remove(); editorHost = null; panelHost = null;
});

describe('QA20K find-panel result and UI workflows 01171–01180', () => {
  it('QA20K-01171 shows the current hit index and highlights every visible match after entering a query', async () => {
    const { input } = await mountPanel('Policy A. Policy B. Policy C.');
    await typeInto(input()!, 'Policy'); await settleSearch();
    expect(count()).toBe('1 / 3');
    expect(marks().map((mark) => mark.textContent)).toEqual(['Policy', 'Policy', 'Policy']);
    expect(editor!.view!.dom.querySelectorAll('.dx-search-current')).toHaveLength(1);
  });

  it('QA20K-01172 settles rapid query edits on the final text instead of leaving stale result highlights', async () => {
    const { input } = await mountPanel('draft approved draft final');
    await typeInto(input()!, 'draft');
    await typeInto(input()!, 'final');
    await settleSearch();
    expect(count()).toBe('1 / 1');
    expect(marks().map((mark) => mark.textContent)).toEqual(['final']);
    expect(searchState(editor!.view!.state).query).toBe('final');
  });

  it('QA20K-01173 reports a missing phrase and disables both result navigation controls', async () => {
    const { input } = await mountPanel('Quarterly report is ready');
    await typeInto(input()!, 'invoice'); await settleSearch();
    expect(count()).toBe('找不到');
    expect(panelHost!.querySelector('button[aria-label="上一個"]')!.hasAttribute('disabled')).toBe(true);
    expect(panelHost!.querySelector('button[aria-label="下一個"]')!.hasAttribute('disabled')).toBe(true);
    expect(marks()).toHaveLength(0);
  });

  it('QA20K-01174 lets a reviewer place the editor selection on the currently highlighted first hit', async () => {
    const { input } = await mountPanel('Clause 4. Clause 7. Clause 9.');
    await typeInto(input()!, 'Clause'); await settleSearch();
    (panelHost!.querySelector('button[aria-label="下一個"]') as HTMLButtonElement).click();
    await nextTick();
    const state = editor!.view!.state;
    const currentMatch = searchState(state).matches[0];
    expect(count()).toBe('1 / 3');
    expect(editor!.view!.dom.querySelector('.dx-search-current')?.textContent).toBe('Clause');
    expect(state.selection.from).toBe(currentMatch.from);
    expect(state.selection.to).toBe(currentMatch.to);
    expect(panelHost!.querySelector('.dx-count')?.getAttribute('aria-live')).toBe('polite');
  });

  it('QA20K-01175 moves backward from the first result to the final result and updates the live counter', async () => {
    const { input } = await mountPanel('Budget. Budget. Budget.');
    await typeInto(input()!, 'Budget'); await settleSearch();
    (panelHost!.querySelector('button[aria-label="上一個"]') as HTMLButtonElement).click();
    await nextTick();
    expect(count()).toBe('3 / 3');
    const current = editor!.view!.dom.querySelector('.dx-search-current') as HTMLElement;
    expect(current.textContent).toBe('Budget');
    const currentMatch = searchState(editor!.view!.state).matches[searchState(editor!.view!.state).current];
    expect(currentMatch.from).toBeGreaterThan(searchState(editor!.view!.state).matches[1].from);
    expect(editor!.view!.state.selection.from).toBe(currentMatch.from);
    expect(editor!.view!.state.selection.to).toBe(currentMatch.to);
  });

  it('QA20K-01176 narrows the displayed result count to exact capitalization when case sensitivity is enabled', async () => {
    const { input } = await mountPanel('Acme ACME acme');
    await typeInto(input()!, 'acme'); await settleSearch();
    expect(count()).toBe('1 / 3');
    (panelHost!.querySelector('input[type="checkbox"]') as HTMLInputElement).click();
    await nextTick();
    expect(count()).toBe('1 / 1');
    expect(marks().map((mark) => mark.textContent)).toEqual(['acme']);
  });

  it('QA20K-01177 replaces only the currently selected occurrence and leaves later hits available in the panel', async () => {
    const { input } = await mountPanel('TBD | TBD | OK', { replace: true });
    await typeInto(input()!, 'TBD'); await settleSearch();
    await typeInto(input('取代為')!, 'Done');
    (panelHost!.querySelector('button:not([aria-label])') as HTMLButtonElement).click();
    await nextTick();
    expect(text()).toBe('Done | TBD | OK');
    expect(count()).toBe('1 / 1');
    expect(marks().map((mark) => mark.textContent)).toEqual(['TBD']);
  });

  it('QA20K-01178 confirms the number of bulk replacements and removes all corresponding highlights', async () => {
    const { input } = await mountPanel('TBD / TBD / TBD', { replace: true });
    await typeInto(input()!, 'TBD'); await settleSearch();
    await typeInto(input('取代為')!, 'Ready');
    const button = (label: string) => [...panelHost!.querySelectorAll('button')].find((b) => b.textContent?.trim() === label)!;
    button('全部取代').click();
    await nextTick();
    // Every match is listed for review first (GOV-ISSUE-010); nothing is replaced yet.
    expect(text()).toBe('TBD / TBD / TBD');
    button('全部取代（3 處）').click();
    await nextTick();
    expect(text()).toBe('Ready / Ready / Ready');
    expect(panelHost!.querySelector('.dx-find-msg')?.textContent).toBe('已取代 3 處（正文 3）');
    expect(marks()).toHaveLength(0);
  });

  it('QA20K-01179 lets a read-only reviewer navigate matches without exposing any replacement controls', async () => {
    const { input } = await mountPanel('Review. Review. Approved.', { canReplace: false });
    await typeInto(input()!, 'Review'); await settleSearch();
    expect(count()).toBe('1 / 2');
    expect(panelHost!.querySelector('button[aria-label="下一個"]')!.hasAttribute('disabled')).toBe(false);
    expect(panelHost!.querySelector('button[aria-label*="取代"]')).toBeNull();
    expect(panelHost!.querySelector('input[aria-label="取代為"]')).toBeNull();
    expect(panelHost!.textContent).not.toContain('全部取代');
  });

  it('QA20K-01180 closes Find on Escape, clears visible search decorations and returns focus to the editor', async () => {
    const { input, closed } = await mountPanel('Approval pending. Approval received.');
    await typeInto(input()!, 'Approval'); await settleSearch();
    expect(marks()).toHaveLength(2);
    panelHost!.querySelector('[role="search"]')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    expect(closed).toHaveBeenCalledTimes(1);
    expect(panelHost!.querySelector('[role="search"]')).toBeNull();
    expect(marks()).toHaveLength(0);
    expect(document.activeElement).toBe(editor!.view!.dom);
  });
});
