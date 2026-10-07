import { describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

async function setup(body: string, selectText: string, options: { open?: boolean; styles?: string } = {}) {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const emptyRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getClientRects ??= emptyRects;
  Range.prototype.getBoundingClientRect ??= (() => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect);
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  if (options.styles) zip.file('word/styles.xml', options.styles);
  const host = document.createElement('div');
  const dialogHost = document.createElement('div');
  document.body.append(host, dialogHost);
  const editor = new DocxEditor(host);
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  let position = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (position < 0 && node.isText && node.text?.includes(selectText)) position = pos + node.text.indexOf(selectText);
  });
  if (position < 0) throw new Error(`Fixture text not found: ${selectText}`);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, position + 1)));
  const open = ref(options.open ?? true);
  const closed = ref(false);
  const Dialog = await loadSfc('src/papyrus/vue/ParagraphDialog.vue');
  const app = createApp({ render: () => h(Dialog, { editor, open: open.value, onClose: () => { closed.value = true; open.value = false; } }) });
  app.mount(dialogHost);
  await nextTick();
  const input = async (selector: string, value: string) => {
    let el = dialogHost.querySelector<HTMLInputElement | HTMLSelectElement>(selector);
    const labelText = /aria-label="([^"]+)"/.exec(selector)?.[1];
    if (!el && labelText) {
      const label = Array.from(dialogHost.querySelectorAll('label')).find(candidate => candidate.textContent?.includes(labelText));
      el = label?.querySelector<HTMLInputElement | HTMLSelectElement>('input, select') ?? null;
    }
    if (!el) throw new Error(`Missing dialog control ${selector}`);
    el.value = value;
    el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }));
    await nextTick();
  };
  const apply = async () => { (dialogHost.querySelector('.dx-primary') as HTMLButtonElement).click(); await nextTick(); };
  const xml = async () => {
    return (await JSZip.loadAsync(await writeDocx(editor.view!.state.doc, editor.model))).file('word/document.xml')!.async('string');
  };
  const done = () => { app.unmount(); editor.destroy(); host.remove(); dialogHost.remove(); };
  return { editor, dialogHost, open, closed, input, apply, xml, done };
}

describe('QA20K-00981–00990 paragraph formatting dialog workflows', () => {
  it('QA20K-00981 changes only the selected paragraph after-spacing while retaining its other geometry', async () => {
    const d = await setup('<w:p><w:pPr><w:ind w:left="720"/><w:spacing w:before="120" w:after="80"/></w:pPr><w:r><w:t>Memo target</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:t>Next memo</w:t></w:r></w:p>', 'Memo target');
    await d.input('input[aria-label="與後段距離（點）"]', '18'); await d.apply();
    const xml = await d.xml();
    const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map(match => match[0]);
    expect(paragraphs[0]).toContain('<w:ind w:left="720"');
    expect(paragraphs[0]).toMatch(/<w:spacing[^>]*w:before="120"[^>]*w:after="360"/);
    expect(paragraphs[1]).toContain('<w:t>Next memo</w:t>');
    expect(paragraphs[1]).toContain('w:after="240"');
    d.done();
  });

  it('QA20K-00982 removes a direct hanging indent but leaves the paragraph left indent intact', async () => {
    const d = await setup('<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>Bulky note</w:t></w:r></w:p>', 'Bulky note');
    const method = d.dialogHost.querySelector('select')!;
    method.value = 'none'; method.dispatchEvent(new Event('change', { bubbles: true })); await nextTick();
    await d.apply();
    const xml = await d.xml();
    expect(xml).toMatch(/<w:ind[^>]*w:left="720"/);
    expect(xml).not.toMatch(/w:hanging="360"|w:firstLine="-?360"/);
    d.done();
  });

  it('QA20K-00983 changes paragraph indents from centimeters to character units without retaining stale twips', async () => {
    const d = await setup('<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>Chinese memo</w:t></w:r></w:p>', 'Chinese memo');
    (d.dialogHost.querySelector('input[name="dx-pd-unit"][value="char"]') as HTMLInputElement).click(); await nextTick();
    await d.input('input[aria-label="左"]', '2'); await d.apply();
    const xml = await d.xml();
    expect(xml).toContain('w:leftChars="200"');
    expect(xml).not.toMatch(/<w:ind[^>]*w:left="720"/);
    d.done();
  });

  it('QA20K-00984 stores a hanging indent as positive w:hanging twips', async () => {
    const d = await setup('<w:p><w:r><w:t>Reference note</w:t></w:r></w:p>', 'Reference note');
    const method = d.dialogHost.querySelector('select')!;
    method.value = 'hanging'; method.dispatchEvent(new Event('change', { bubbles: true })); await nextTick();
    await d.input('input[aria-label="位移"]', '0.5'); await d.apply();
    const xml = await d.xml();
    expect(xml).toMatch(/w:hanging="283"/);
    expect(xml).not.toMatch(/w:firstLine="-?283"/);
    d.done();
  });

  it('QA20K-00985 applies 1.15 multiple line spacing while retaining explicit paragraph gaps', async () => {
    const d = await setup('<w:p><w:pPr><w:spacing w:before="100" w:after="160"/></w:pPr><w:r><w:t>Review notes</w:t></w:r></w:p>', 'Review notes');
    const spacing = d.dialogHost.querySelectorAll('fieldset')[1];
    const line = spacing.querySelectorAll('select')[0];
    line.value = 'multiple'; line.dispatchEvent(new Event('change', { bubbles: true })); await nextTick();
    const multiple = spacing.querySelectorAll('input')[2];
    multiple.value = '1.15'; multiple.dispatchEvent(new Event('input', { bubbles: true })); await nextTick();
    await d.apply();
    const xml = await d.xml();
    expect(xml).toMatch(/w:line="276"[^>]*w:lineRule="auto"|w:lineRule="auto"[^>]*w:line="276"/);
    expect(xml).toMatch(/w:before="100"[^>]*w:after="160"/);
    d.done();
  });

  it('QA20K-00986 sets a minimum line height in points rather than treating it as a multiple', async () => {
    const d = await setup('<w:p><w:r><w:t>Signature notes</w:t></w:r></w:p>', 'Signature notes');
    const spacing = d.dialogHost.querySelectorAll('fieldset')[1];
    const line = spacing.querySelectorAll('select')[0];
    line.value = 'atLeast'; line.dispatchEvent(new Event('change', { bubbles: true })); await nextTick();
    const points = spacing.querySelectorAll('input')[2];
    points.value = '14'; points.dispatchEvent(new Event('input', { bubbles: true })); await nextTick();
    await d.apply();
    const xml = await d.xml();
    expect(xml).toMatch(/w:line="280"[^>]*w:lineRule="atLeast"|w:lineRule="atLeast"[^>]*w:line="280"/);
    d.done();
  });

  it('QA20K-00987 rejects a negative paragraph spacing value without changing the document', async () => {
    const d = await setup('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>Invoice note</w:t></w:r></w:p>', 'Invoice note');
    const before = d.editor.view!.state.doc;
    await d.input('input[aria-label="與後段距離（點）"]', '-3'); await d.apply();
    expect(d.dialogHost.querySelector('[role="alert"]')?.textContent).toContain('0 以上');
    expect(d.editor.view!.state.doc).toBe(before);
    expect(d.closed.value).toBe(false);
    d.done();
  });

  it('QA20K-00988 applies an unchanged dialog without creating a document transaction', async () => {
    const d = await setup('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>Stable note</w:t></w:r></w:p>', 'Stable note');
    const before = d.editor.view!.state.doc;
    await d.apply();
    expect(d.editor.view!.state.doc).toBe(before);
    expect(d.closed.value).toBe(true);
    d.done();
  });

  it('QA20K-00989 applies the selected spacing to two paragraphs and leaves the following paragraph alone', async () => {
    const d = await setup('<w:p><w:r><w:t>First item</w:t></w:r></w:p><w:p><w:r><w:t>Second item</w:t></w:r></w:p><w:p><w:r><w:t>Footer note</w:t></w:r></w:p>', 'First item');
    const doc = d.editor.view!.state.doc;
    const from = 1; const to = doc.child(0).nodeSize + doc.child(1).nodeSize - 1;
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(TextSelection.create(doc, from, to)));
    await d.input('input[aria-label="與前段距離（點）"]', '12'); await d.apply();
    const json = d.editor.view!.state.doc.toJSON();
    expect(json.content.slice(0, 2).map((p: any) => p.attrs.spaceBefore)).toEqual([240, 240]);
    expect(json.content[2].attrs.spaceBefore).toBeNull();
    d.done();
  });

  it('QA20K-00990 cancels pending paragraph-format edits with Escape and preserves the existing spacing', async () => {
    const d = await setup('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>Approval note</w:t></w:r></w:p>', 'Approval note');
    const before = d.editor.view!.state.doc;
    await d.input('input[aria-label="與後段距離（點）"]', '30');
    d.dialogHost.querySelector('[role="dialog"]')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    expect(d.editor.view!.state.doc).toBe(before);
    expect(d.closed.value).toBe(true);
    expect(d.editor.view!.state.doc.firstChild!.attrs.spaceAfter).toBe(120);
    d.done();
  });
});
