import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, nextTick } from 'vue';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';

const SFC = resolve(__dirname, '../../src/papyrus/vue/PageSetupDialog.vue');
const OUT = resolve(__dirname, `.qa-page-setup-${process.pid}.tmp.js`);
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const setup = (w: number, h: number, left = 1440) => `<w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="${left}" w:header="720" w:footer="720"/>`;
const para = (text: string, sectPr = '') => `<w:p>${sectPr ? `<w:pPr><w:sectPr>${sectPr}</w:sectPr></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});
afterAll(() => rmSync(OUT, { force: true }));

async function compiledDialog(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'qa-page-setup', inlineTemplate: true });
  const code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText
    // The component's own imports (./locale) are found from this file's folder.
    .replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) => `from '${resolve(dirname(SFC), spec).replace(/\\/g, '/')}'`);
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ OUT.replace(/\\/g, '/'))).default;
}

async function openEditor(multi = false) {
  const zip = blankPackage();
  const body = multi
    ? para('A1', setup(11906, 16838)) + para('B1', setup(16838, 11906, 1800)) + para('C1') + `<w:sectPr>${setup(12240, 15840, 1600)}</w:sectPr>`
    : para('Memo') + `<w:sectPr>${setup(11906, 16838)}</w:sectPr>`;
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return { editor, host, done: () => { editor.destroy(); host.remove(); } };
}

async function mountDialog(editor: DocxEditor) {
  const component = await compiledDialog();
  const host = document.createElement('div'); document.body.append(host);
  const app = createApp(component, { editor, open: true, onClose: () => {} });
  app.mount(host);
  await nextTick();
  return { host, done: () => { app.unmount(); host.remove(); } };
}

function typeNumber(host: HTMLElement, index: number, value: number) {
  const input = host.querySelectorAll<HTMLInputElement>('input[type="number"]')[index];
  input.value = String(value);
  input.dispatchEvent(new Event('input', { bubbles: true }));
}
async function apply(host: HTMLElement) {
  host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
  await nextTick();
}
function selectText(editor: DocxEditor, text: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); });
  if (pos < 0) throw new Error(`missing ${text}`);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
}
async function savedSections(editor: DocxEditor) {
  const zip = await JSZip.loadAsync(await editor.save());
  const xml = await zip.file('word/document.xml')!.async('string');
  return xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
}

describe('QA20K-00951–00960 page setup dialog workflows', () => {
  it('QA20K-00951 applying the A3 paper preset writes A3 portrait dimensions', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const paper = d.host.querySelector('select')!; paper.value = 'A3'; paper.dispatchEvent(new Event('change', { bubbles: true }));
      await apply(d.host);
      const [sect] = await savedSections(x.editor);
      expect(sect).toContain('<w:pgSz w:w="16838" w:h="23811"/>');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00952 custom page dimensions convert centimetres to DOCX twips', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      typeNumber(d.host, 0, 20); typeNumber(d.host, 1, 28);
      await apply(d.host);
      const [sect] = await savedSections(x.editor);
      expect(sect).toContain('<w:pgSz w:w="11339" w:h="15874"/>');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00953 switching A4 to landscape swaps physical dimensions and records orientation', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const landscape = d.host.querySelectorAll<HTMLInputElement>('input[name="dx-ps-orient"]')[1];
      landscape.checked = true; landscape.dispatchEvent(new Event('change', { bubbles: true }));
      await apply(d.host);
      const [sect] = await savedSections(x.editor);
      expect(sect).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00954 adjusting print margins and header footer distances preserves paper size', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      [2.0, 2.1, 2.2, 2.3, 1.1, 1.2].forEach((value, i) => typeNumber(d.host, i + 2, value));
      await apply(d.host);
      const [sect] = await savedSections(x.editor);
      expect(sect).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
      expect(sect).toContain('w:top="1134"'); expect(sect).toContain('w:bottom="1191"');
      expect(sect).toContain('w:left="1247"'); expect(sect).toContain('w:right="1304"');
      expect(sect).toContain('w:header="624"'); expect(sect).toContain('w:footer="680"');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00955 applying a paper change to the cursor section leaves other section geometries intact', async () => {
    const x = await openEditor(true); selectText(x.editor, 'B1'); const d = await mountDialog(x.editor);
    try {
      const paper = d.host.querySelector('select')!; paper.value = 'A5'; paper.dispatchEvent(new Event('change', { bubbles: true }));
      await apply(d.host);
      const sects = await savedSections(x.editor);
      expect(sects).toHaveLength(3);
      expect(sects.map(s => /<w:pgSz w:w="(\d+)" w:h="(\d+)"/.exec(s)?.slice(1))).toEqual([
        ['11906', '16838'], ['11906', '8391'], ['12240', '15840'],
      ]);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00956 applying one page preset to all sections writes the geometry to every section', async () => {
    const x = await openEditor(true); const d = await mountDialog(x.editor);
    try {
      const paper = d.host.querySelector('select')!; paper.value = 'Letter'; paper.dispatchEvent(new Event('change', { bubbles: true }));
      d.host.querySelector<HTMLInputElement>('input[name="dx-ps-scope"][value="all"]')!.click();
      await apply(d.host);
      const sects = await savedSections(x.editor);
      expect(sects).toHaveLength(3);
      expect(sects.map(s => /<w:pgSz w:w="(\d+)" w:h="(\d+)"/.exec(s)?.slice(1))).toEqual([['12240','15840'],['12240','15840'],['12240','15840']]);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00957 margins that consume the printable width keep the dialog open without committing', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const before = (await savedSections(x.editor))[0];
      typeNumber(d.host, 4, 11.5); typeNumber(d.host, 5, 9.6);
      await apply(d.host);
      expect(d.host.querySelector('[role="alert"]')?.textContent).toContain('邊界太大');
      expect(x.editor.isModified()).toBe(false);
      expect((await savedSections(x.editor))[0]).toBe(before);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00958 a page width below five centimetres is rejected without changing the document', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const before = (await savedSections(x.editor))[0];
      typeNumber(d.host, 0, 4.9);
      await apply(d.host);
      expect(d.host.querySelector('[role="alert"]')?.textContent).toContain('介於 5 到 150 公分');
      expect(x.editor.isModified()).toBe(false);
      expect((await savedSections(x.editor))[0]).toBe(before);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00959 applying an unchanged section setup is a no-op that keeps a clean document', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const before = (await savedSections(x.editor))[0];
      await apply(d.host);
      expect(x.editor.isModified()).toBe(false);
      expect((await savedSections(x.editor))[0]).toBe(before);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-00960 opening page setup at the second section loads that section values and scope label', async () => {
    const x = await openEditor(true); selectText(x.editor, 'B1'); const d = await mountDialog(x.editor);
    try {
      const dimensions = [...d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')].slice(0, 2).map(input => Number(input.value));
      expect(dimensions).toEqual([29.7, 21]);
      expect(d.host.textContent).toContain('目前這一節（第 2 節，共 3 節）');
      expect(d.host.querySelector<HTMLInputElement>('input[name="dx-ps-scope"][value="section"]')!.checked).toBe(true);
    } finally { d.done(); x.done(); }
  });
});
