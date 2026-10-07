import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { TextSelection } from 'prosemirror-state';
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

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => vi.restoreAllMocks());

async function setup(text: string, extraBody = '') {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>${extraBody}<w:sectPr/></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { model } = await readDocx(bytes);
  const host = document.createElement('div');
  const toolbarHost = document.createElement('div');
  document.body.append(host, toolbarHost);
  const snapshot = ref<EditorSnapshot | null>(null);
  const editor = new DocxEditor(host, { onUpdate: (s) => { snapshot.value = s; } });
  await editor.open(bytes);
  const Toolbar = await loadSfc(VUE + 'DocxToolbar.vue');
  const app = createApp({ render: () => h(Toolbar, { editor, snapshot: snapshot.value, styles: editor.model.paragraphStyles }) });
  app.mount(toolbarHost);
  const refresh = async () => { snapshot.value = editor.snapshot(); await nextTick(); };
  const done = () => { app.unmount(); editor.destroy(); host.remove(); toolbarHost.remove(); };
  return { editor, model, toolbar: toolbarHost, refresh, done };
}

function textStart(editor: DocxEditorType, text: string) {
  let start = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (start < 0 && node.isText && node.text!.includes(text)) start = pos + node.text!.indexOf(text);
  });
  if (start < 0) throw new Error(`missing text: ${text}`);
  return start;
}

function select(editor: DocxEditorType, text: string, offset = 0, length = text.length) {
  const start = textStart(editor, text) + offset;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start, start + length)));
  return start;
}

function caret(editor: DocxEditorType, text: string, offset: number) {
  const pos = textStart(editor, text) + offset;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
  return pos;
}

function addLinkMark(editor: DocxEditorType, text: string, href: string, bold = false) {
  const from = textStart(editor, text);
  let tr = editor.view!.state.tr.addMark(from, from + text.length, schema.marks.link.create({ href }));
  if (bold) tr = tr.addMark(from, from + text.length, schema.marks.bold.create());
  editor.view!.dispatch(tr);
  editor.view!.dispatch(closeHistory(editor.view!.state.tr));
}

function linkRuns(editor: DocxEditorType) {
  const out: { text: string; href: string | null; marks: string[] }[] = [];
  editor.view!.state.doc.descendants((node) => {
    if (node.isText) out.push({
      text: node.text!,
      href: schema.marks.link.isInSet(node.marks)?.attrs.href ?? null,
      marks: node.marks.map((mark) => mark.type.name),
    });
  });
  return out;
}

function clickLink(toolbar: HTMLElement) {
  const button = toolbar.querySelector('button[title="插入連結"]') as HTMLButtonElement;
  if (!button) throw new Error('missing Insert Link toolbar button');
  button.click();
}

describe('QA20K-00861–00870 hyperlink editing workflows', () => {
  it('QA20K-00861 links selected invoice wording without changing adjacent text', async () => {
    const d = await setup('Open invoice details before approval');
    select(d.editor, 'invoice');
    vi.spyOn(window, 'prompt').mockReturnValue('https://intranet.example/invoice');
    clickLink(d.toolbar);
    expect(linkRuns(d.editor)).toEqual([
      { text: 'Open ', href: null, marks: [] },
      { text: 'invoice', href: 'https://intranet.example/invoice', marks: ['link'] },
      { text: ' details before approval', href: null, marks: [] },
    ]);
    d.done();
  });

  it('QA20K-00862 links the selected approver email with a mailto target', async () => {
    const d = await setup('Contact payables@example.com for approval');
    select(d.editor, 'payables@example.com');
    vi.spyOn(window, 'prompt').mockReturnValue('mailto:payables@example.com');
    clickLink(d.toolbar);
    expect(linkRuns(d.editor)).toEqual([
      { text: 'Contact ', href: null, marks: [] },
      { text: 'payables@example.com', href: 'mailto:payables@example.com', marks: ['link'] },
      { text: ' for approval', href: null, marks: [] },
    ]);
    d.done();
  });

  it('QA20K-00863 links an appendix label to a same-document anchor', async () => {
    const d = await setup(
      'See Appendix for the details',
      '<w:p><w:bookmarkStart w:id="7" w:name="Appendix"/><w:r><w:t>Appendix begins here</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
    );
    select(d.editor, 'Appendix');
    vi.spyOn(window, 'prompt').mockReturnValue('#Appendix');
    clickLink(d.toolbar);
    const zip = await JSZip.loadAsync(await writeDocx(d.editor.view!.state.doc, d.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:hyperlink w:anchor="Appendix">');
    expect(xml).toContain('<w:bookmarkStart w:id="7" w:name="Appendix"/>');
    expect(xml).toContain('<w:bookmarkEnd w:id="7"/>');
    expect(xml.indexOf('<w:bookmarkStart w:id="7" w:name="Appendix"/>')).toBeLessThan(xml.indexOf('<w:bookmarkEnd w:id="7"/>'));
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).not.toContain('hyperlink');
    d.done();
  });

  it('QA20K-00864 creates a clickable visible URL when inserting a link at a caret', async () => {
    const d = await setup('Reference: end');
    caret(d.editor, 'Reference: end', 'Reference: '.length);
    vi.spyOn(window, 'prompt').mockReturnValue('https://policy.example/2026');
    clickLink(d.toolbar);
    expect(d.editor.view!.state.doc.textContent).toBe('Reference: https://policy.example/2026end');
    expect(linkRuns(d.editor)).toEqual([
      { text: 'Reference: ', href: null, marks: [] },
      { text: 'https://policy.example/2026', href: 'https://policy.example/2026', marks: ['link'] },
      { text: 'end', href: null, marks: [] },
    ]);
    d.done();
  });

  it('QA20K-00865 edits an existing hyperlink destination from a cursor inside it', async () => {
    const d = await setup('Review policy today');
    addLinkMark(d.editor, 'policy', 'https://old.example/policy');
    caret(d.editor, 'policy', 3);
    vi.spyOn(window, 'prompt').mockReturnValue('https://handbook.example/policy');
    clickLink(d.toolbar);
    expect(linkRuns(d.editor)).toEqual([
      { text: 'Review ', href: null, marks: [] },
      { text: 'policy', href: 'https://handbook.example/policy', marks: ['link'] },
      { text: ' today', href: null, marks: [] },
    ]);
    d.done();
  });

  it('QA20K-00866 removes a hyperlink at the caret while keeping its display wording', async () => {
    const d = await setup('Open policy reference');
    addLinkMark(d.editor, 'policy reference', 'https://old.example/policy');
    caret(d.editor, 'policy reference', 5);
    vi.spyOn(window, 'prompt').mockReturnValue('');
    clickLink(d.toolbar);
    expect(d.editor.view!.state.doc.textContent).toBe('Open policy reference');
    expect(linkRuns(d.editor)).toEqual([{ text: 'Open policy reference', href: null, marks: [] }]);
    d.done();
  });

  it('QA20K-00867 canceling the link prompt leaves the selected phrase and history unchanged', async () => {
    const d = await setup('Please review the final draft');
    select(d.editor, 'final draft');
    const before = d.editor.view!.state.doc;
    const modified = d.editor.isModified;
    vi.spyOn(window, 'prompt').mockReturnValue(null);
    clickLink(d.toolbar);
    expect(d.editor.view!.state.doc.eq(before)).toBe(true);
    expect(d.editor.isModified).toBe(modified);
    expect(linkRuns(d.editor).every((run) => run.href == null)).toBe(true);
    d.done();
  });

  it('QA20K-00868 rejects a javascript link target without altering selected text', async () => {
    const d = await setup('Open the approval record');
    select(d.editor, 'approval record');
    const before = d.editor.view!.state.doc;
    vi.spyOn(window, 'prompt').mockReturnValue('javascript:alert(1)');
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {});
    clickLink(d.toolbar);
    expect(alert).toHaveBeenCalledWith('只支援 http://、https://、mailto: 或 # 開頭的連結');
    expect(d.editor.view!.state.doc.eq(before)).toBe(true);
    expect(linkRuns(d.editor).every((run) => run.href == null)).toBe(true);
    d.done();
  });

  it('QA20K-00869 changing a linked phrase selection leaves the unselected suffix on its old target', async () => {
    const d = await setup('Read policy section');
    addLinkMark(d.editor, 'policy section', 'https://old.example/guide');
    select(d.editor, 'policy section', 0, 'policy'.length);
    vi.spyOn(window, 'prompt').mockReturnValue('https://new.example/policy');
    clickLink(d.toolbar);
    expect(linkRuns(d.editor)).toEqual([
      { text: 'Read ', href: null, marks: [] },
      { text: 'policy', href: 'https://new.example/policy', marks: ['link'] },
      { text: ' section', href: 'https://old.example/guide', marks: ['link'] },
    ]);
    d.done();
  });

  it('QA20K-00870 saves and reopens a bold linked phrase with its external target intact', async () => {
    const d = await setup('The approved terms apply');
    const from = select(d.editor, 'approved terms');
    const tr = d.editor.view!.state.tr.addMark(from, from + 'approved terms'.length, schema.marks.bold.create());
    d.editor.view!.dispatch(tr);
    vi.spyOn(window, 'prompt').mockReturnValue('https://contract.example/terms');
    clickLink(d.toolbar);
    await d.refresh();
    const zip = await JSZip.loadAsync(await writeDocx(d.editor.view!.state.doc, d.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const relId = /<w:hyperlink r:id="([^"]+)"/.exec(xml)?.[1];
    expect(relId).toBeTruthy();
    expect(rels).toContain(`Id="${relId}"`);
    expect(rels).toContain('Target="https://contract.example/terms" TargetMode="External"');
    const reopened = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const marks: { text: string; names: string[] }[] = [];
    reopened.doc.descendants((node) => { if (node.isText) marks.push({ text: node.text!, names: node.marks.map((m) => m.type.name) }); });
    const linkedRun = marks.find((run) => run.text === 'approved terms');
    expect(linkedRun?.names).toContain('bold');
    expect(linkedRun?.names).toContain('link');
    expect(reopened.doc.textContent).toBe('The approved terms apply');
    d.done();
  });
});
