import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { DocxEditor } from '../../src/papyrus/editor/core';
import ParagraphDialog from '../../src/papyrus/vue/ParagraphDialog.vue';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const p = (text: string, pPr = '') => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const fixture = async (body = [p('First note'), p('Target paragraph'), p('Last note')].join('')) => {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
};

async function setup(target = 'Target paragraph', body?: string) {
  const editorHost = document.createElement('div');
  const dialogHost = document.createElement('div');
  document.body.append(editorHost, dialogHost);
  const editor = new DocxEditor(editorHost);
  await editor.open(await fixture(body));
  let position = -1;
  editor.view!.state.doc.descendants((node, at) => {
    if (position < 0 && node.isText && node.text!.includes(target)) position = at + node.text!.indexOf(target);
  });
  if (position < 0) throw new Error(`Fixture text not found: ${target}`);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, position, position + target.length)));
  const isOpen = ref(true);
  const app = createApp({ setup: () => () => h(ParagraphDialog, { editor, open: isOpen.value, onClose: () => { isOpen.value = false; } }) });
  app.mount(dialogHost);
  await nextTick();
  const dialog = () => dialogHost.querySelector<HTMLElement>('[role="dialog"]');
  const flush = async () => { await nextTick(); await new Promise((resolve) => setTimeout(resolve, 0)); await nextTick(); };
  const textInput = async (fieldset: number, index: number, value: string) => {
    const input = dialog()!.querySelectorAll('fieldset')[fieldset].querySelectorAll<HTMLInputElement>('input[type="text"]')[index];
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
  };
  const choose = async (fieldset: number, index: number, value: string) => {
    const select = dialog()!.querySelectorAll('fieldset')[fieldset].querySelectorAll<HTMLSelectElement>('select')[index];
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
    await nextTick();
  };
  const setUnit = async (unit: 'cm' | 'char') => {
    const radio = dialog()!.querySelector<HTMLInputElement>(`input[name="dx-pd-unit"][value="${unit}"]`)!;
    radio.checked = true;
    radio.dispatchEvent(new Event('input', { bubbles: true }));
    radio.dispatchEvent(new Event('change', { bubbles: true }));
    await nextTick();
  };
  const apply = async () => { dialog()!.querySelector<HTMLButtonElement>('.dx-primary')!.click(); await flush(); };
  const saveXml = async () => {
    const out = await JSZip.loadAsync(await writeDocx(editor.view!.state.doc, editor.model));
    return out.file('word/document.xml')!.async('string');
  };
  const done = () => { app.unmount(); editor.destroy(); editorHost.remove(); dialogHost.remove(); };
  return { editor, dialog, textInput, choose, setUnit, apply, saveXml, flush, done, isOpen };
}

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} })) as any;
});

describe('paragraph dialog office workflows', () => {
  it('QA20K-00971 sets a right indent without replacing the existing left indent', async () => {
    const x = await setup('Target paragraph', p('First note') + p('Target paragraph', '<w:ind w:left="720"/>') + p('Last note'));
    try {
      await x.textInput(0, 1, '1.27');
      await x.apply();
      const paras = x.editor.view!.state.doc.content.content;
      expect(paras[1].attrs.indLeft).toBe(720);
      expect(paras[1].attrs.indRight).toBe(720);
      const xml = await x.saveXml();
      expect(xml).toMatch(/<w:ind [^>]*w:left="720"[^>]*w:right="720"|<w:ind [^>]*w:right="720"[^>]*w:left="720"/);
      expect(paras[0].attrs.indRight).toBeNull();
      expect(paras[2].attrs.indRight).toBeNull();
    } finally { x.done(); }
  });

  it('QA20K-00972 applies a two-character hanging indent in character units', async () => {
    const x = await setup();
    try {
      await x.setUnit('char');
      await x.choose(0, 0, 'hanging');
      await x.textInput(0, 2, '2');
      await x.apply();
      expect(x.editor.view!.state.doc.child(1).attrs.indFirstChars).toBe(-200);
      expect(x.editor.view!.state.doc.child(1).attrs.indFirst).toBeNull();
      expect(await x.saveXml()).toContain('w:hangingChars="200"');
    } finally { x.done(); }
  });

  it('QA20K-00973 removes the previous hanging indent when the user selects no special indent', async () => {
    const x = await setup('Target paragraph', p('First note') + p('Target paragraph', '<w:ind w:left="720" w:hanging="360"/>') + p('Last note'));
    try {
      await x.choose(0, 0, 'none');
      await x.apply();
      expect(x.editor.view!.state.doc.child(1).attrs.indFirst).toBeNull();
      const xml = await x.saveXml();
      expect(xml).not.toContain('w:hanging="360"');
      expect(xml).toContain('w:left="720"');
    } finally { x.done(); }
  });

  it('QA20K-00974 converts line spacing to a custom 1.25 multiple', async () => {
    const x = await setup();
    try {
      await x.choose(1, 0, 'multiple');
      await x.textInput(1, 2, '1.25');
      await x.apply();
      expect(x.editor.view!.state.doc.child(1).attrs.line).toBe(300);
      expect(x.editor.view!.state.doc.child(1).attrs.lineRule).toBe('auto');
      expect(await x.saveXml()).toMatch(/w:line="300"[^>]*w:lineRule="auto"|w:lineRule="auto"[^>]*w:line="300"/);
    } finally { x.done(); }
  });

  it('QA20K-00975 stores a minimum line height as atLeast instead of exact spacing', async () => {
    const x = await setup();
    try {
      await x.choose(1, 0, 'atLeast');
      await x.textInput(1, 2, '13.5');
      await x.apply();
      expect(x.editor.view!.state.doc.child(1).attrs.line).toBe(270);
      expect(x.editor.view!.state.doc.child(1).attrs.lineRule).toBe('atLeast');
      expect(await x.saveXml()).toMatch(/w:line="270"[^>]*w:lineRule="atLeast"|w:lineRule="atLeast"[^>]*w:line="270"/);
    } finally { x.done(); }
  });

  it('QA20K-00976 clears an explicit after-paragraph spacing override back to style default', async () => {
    const x = await setup('Target paragraph', p('First note') + p('Target paragraph', '<w:spacing w:before="120" w:after="240"/>') + p('Last note'));
    try {
      await x.textInput(1, 1, '');
      await x.apply();
      expect(x.editor.view!.state.doc.child(1).attrs.spaceBefore).toBe(120);
      expect(x.editor.view!.state.doc.child(1).attrs.spaceAfter).toBeNull();
      const xml = await x.saveXml();
      expect(xml).toContain('w:before="120"');
      expect(xml).not.toContain('w:after="240"');
    } finally { x.done(); }
  });

  it('QA20K-00977 rejects a negative spacing value without committing any paragraph change', async () => {
    const x = await setup();
    try {
      const before = x.editor.view!.state.doc;
      await x.textInput(1, 0, '-3');
      await x.apply();
      expect(x.dialog()).not.toBeNull();
      expect(x.dialog()!.querySelector('[role="alert"]')!.textContent).toContain('0 以上');
      expect(x.editor.view!.state.doc).toBe(before);
      expect(x.editor.isModified()).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-00978 cancel discards draft paragraph settings and keeps the document clean', async () => {
    const x = await setup();
    try {
      const before = x.editor.view!.state.doc;
      await x.textInput(0, 0, '2');
      x.dialog()!.querySelectorAll<HTMLButtonElement>('.dx-actions button')[0].click();
      await x.flush();
      expect(x.dialog()).toBeNull();
      expect(x.editor.view!.state.doc).toBe(before);
      expect(x.editor.isModified()).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-00979 applies a paragraph-format patch to every paragraph in the selection only', async () => {
    const x = await setup('First note');
    try {
      const doc = x.editor.view!.state.doc;
      const textPos = (text: string) => {
        let found = -1;
        doc.descendants((node, at) => { if (found < 0 && node.isText && node.text === text) found = at; });
        if (found < 0) throw new Error(`Fixture text not found: ${text}`);
        return found;
      };
      const start = textPos('First note');
      const end = textPos('Target paragraph') + 'Target paragraph'.length;
      x.editor.view!.dispatch(x.editor.view!.state.tr.setSelection(TextSelection.create(doc, start, end)));
      await x.textInput(1, 0, '8');
      await x.apply();
      const attrs = x.editor.view!.state.doc.content.content.map((node) => node.attrs.spaceBefore);
      expect(attrs).toEqual([160, 160, null]);
      const xml = await x.saveXml();
      const paras = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g)!;
      expect(paras.slice(0, 2).every((part) => part.includes('w:before="160"'))).toBe(true);
      expect(paras[2]).not.toContain('w:before="160"');
    } finally { x.done(); }
  });

  it('QA20K-00980 changes first-line indentation without disturbing paragraph spacing or its neighbor', async () => {
    const x = await setup('Target paragraph', p('First note') + p('Target paragraph', '<w:spacing w:after="200"/>') + p('Last note'));
    try {
      await x.choose(0, 0, 'first');
      await x.textInput(0, 2, '0.5');
      await x.apply();
      const target = x.editor.view!.state.doc.child(1).attrs;
      expect(target.indFirst).toBe(283);
      expect(target.spaceAfter).toBe(200);
      expect(x.editor.view!.state.doc.child(2).attrs.indFirst).toBeNull();
      const xml = await x.saveXml();
      expect(xml).toContain('w:firstLine="283"');
      expect(xml).toContain('w:after="200"');
    } finally { x.done(); }
  });
});
