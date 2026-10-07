import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { readComments } from '../../src/papyrus/docx/comments';
import { commentedText, commentRanges } from '../../src/papyrus/editor/review';
import { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
async function bytes(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}
async function open(body: string, options: Record<string, unknown> = {}) {
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host, options); await editor.open(await bytes(body));
  return { editor, view: editor.view!, done: () => { editor.destroy(); host.remove(); } };
}
function select(editor: DocxEditor, text: string, start = 0, length = text.length) {
  let pos = -1;
  editor.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); return pos < 0; });
  if (pos < 0) throw new Error(`Text not found: ${text}`);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos + start, pos + start + length)));
  return pos + start;
}
async function save(editor: DocxEditor) { return JSZip.loadAsync(await editor.save()); }
function tableShape(xml: string) {
  const table = xml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>/)?.[0];
  if (!table) throw new Error('DOCX has no table');
  const grid = table.match(/<w:tblGrid\b[^>]*>([\s\S]*?)<\/w:tblGrid>/)?.[1] ?? '';
  const rows = [...table.matchAll(/<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g)].map(m => m[1]);
  return {
    rowCount: rows.length,
    cellsPerRow: rows.map(r => [...r.matchAll(/<w:tc\b/g)].length),
    gridWidths: [...grid.matchAll(/<w:gridCol\b[^>]*\bw:w="(\d+)"/g)].map(m => Number(m[1])),
  };
}
function textStart(doc: import('prosemirror-model').Node, text: string) {
  let pos = -1;
  doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); });
  if (pos < 0) throw new Error(`Text not found: ${text}`);
  return pos;
}
async function mountPanel(src: Uint8Array | null) {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, author: { name: '我' }, onReady: (ed: DocxEditor) => (editor = ed) }) });
  app.mount(host);
  for (let i = 0; i < 60 && !editor; i++) { await new Promise(r => setTimeout(r, 0)); await nextTick(); }
  const flush = async () => { for (let i = 0; i < 5; i++) { await new Promise(r => setTimeout(r, 0)); await nextTick(); } };
  return { host, editor: editor!, flush, done: () => { app.unmount(); host.remove(); } };
}

describe('QA20K comment workflows 00791–00800', () => {
  it('QA20K-00791 anchors a comment across two paragraphs and survives DOCX reopen', async () => {
    const { editor, view, done } = await open(para('Quarterly') + para('forecast'));
    const first = select(editor, 'Quarterly');
    let second = -1; view.state.doc.descendants((n, p) => { if (n.isText && n.text === 'forecast') second = p + n.nodeSize; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, first, second)));
    const id = editor.addComment('Review these two lines')!;
    expect(commentedText(view.state.doc, id)).toBe('Quarterly\nforecast');
    const reopened = await readDocx(await (await save(editor)).generateAsync({ type: 'uint8array' }));
    expect(commentedText(reopened.doc, id)).toBe('Quarterly\nforecast');
    done();
  });

  it('QA20K-00792 keeps a comment anchored when text is inserted before its range', async () => {
    const { editor, view, done } = await open(para('Hello world here'));
    const world = select(editor, 'world'); const id = editor.addComment('Check this word')!;
    view.dispatch(view.state.tr.insertText('PRE ', 1));
    expect(view.state.doc.textContent).toBe('PRE Hello world here');
    expect(commentedText(view.state.doc, id)).toBe('world');
    const reopened = await readDocx(await (await save(editor)).generateAsync({ type: 'uint8array' }));
    expect(commentedText(reopened.doc, id)).toBe('world');
    expect(reopened.doc.textContent).toContain('PRE Hello');
    expect(commentRanges(view.state.doc).get(id)!.from).toBeGreaterThan(world);
    done();
  });

  it('QA20K-00793 removes deleted words from the comment anchor but retains its remainder', async () => {
    const { editor, view, done } = await open(para('Please review these words'));
    select(editor, 'review these'); const id = editor.addComment('Clarify phrasing')!;
    const range = commentRanges(view.state.doc).get(id)!;
    view.dispatch(view.state.tr.delete(range.from, range.from + 'review '.length));
    expect(view.state.doc.textContent).toBe('Please these words');
    expect(commentedText(view.state.doc, id)).toBe('these');
    const reopened = await readDocx(await (await save(editor)).generateAsync({ type: 'uint8array' }));
    expect(commentedText(reopened.doc, id)).toBe('these');
    done();
  });

  it('QA20K-00794 edits comment text as one undoable change without moving its anchor', async () => {
    const { editor, view, done } = await open(para('Budget approved'));
    select(editor, 'Budget'); const id = editor.addComment('Check amount')!;
    expect(editor.editComment(id, 'Checked by finance')).toBe(true);
    expect(editor.comments()[0]).toMatchObject({ id, text: 'Checked by finance', author: '使用者' });
    expect(commentedText(view.state.doc, id)).toBe('Budget');
    editor.undo();
    expect(editor.comments()[0].text).toBe('Check amount');
    expect(commentedText(view.state.doc, id)).toBe('Budget');
    editor.redo();
    expect(editor.comments()[0].text).toBe('Checked by finance');
    expect(commentedText(view.state.doc, id)).toBe('Budget');
    done();
  });

  it('QA20K-00795 escapes comment XML metacharacters and reads the original text back', async () => {
    const { editor, done } = await open(para('Total 10'));
    select(editor, '10'); const id = editor.addComment('Use <net> & "gross" > 10')!;
    const zip = await save(editor); const xml = await zip.file('word/comments.xml')!.async('string');
    expect(xml).toContain('Use &lt;net&gt; &amp; "gross" &gt; 10');
    expect((await readComments(zip)).find(c => c.id === id)!.text).toBe('Use <net> & "gross" > 10');
    done();
  });

  it('QA20K-00796 cancels a comment draft with Escape without changing the document', async () => {
    const { host, editor, flush, done } = await mountPanel(await bytes(para('Please approve')));
    select(editor, 'approve');
    (Array.from(host.querySelectorAll('.dx-toolbar button')).find(b => b.textContent === '新增留言') as HTMLButtonElement).click();
    await flush();
    const draft = host.querySelector('.dx-comment-draft')!;
    const box = draft.querySelector('textarea') as HTMLTextAreaElement; box.value = 'Draft that should be discarded'; box.dispatchEvent(new Event('input')); await flush();
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); await flush();
    expect(host.querySelector('.dx-comment-draft')).toBeNull();
    expect(editor.comments()).toEqual([]); expect(editor.isModified()).toBe(false);
    done();
  });

  it('QA20K-00797 refuses a whitespace-only comment in the panel', async () => {
    const { host, editor, flush, done } = await mountPanel(await bytes(para('Approve')));
    select(editor, 'Approve');
    (Array.from(host.querySelectorAll('.dx-toolbar button')).find(b => b.textContent === '新增留言') as HTMLButtonElement).click(); await flush();
    const box = host.querySelector('.dx-comment-draft textarea') as HTMLTextAreaElement;
    box.value = '   \n  '; box.dispatchEvent(new Event('input')); await flush();
    const submit = host.querySelector('.dx-comment-draft button[type="submit"]') as HTMLButtonElement;
    expect(submit.disabled).toBe(true);
    expect(editor.comments()).toEqual([]); expect(editor.isModified()).toBe(false);
    done();
  });

  it('QA20K-00798 refuses comment creation while a header is being edited', async () => {
    const { editor, done } = await open(para('Body text'));
    (editor as any).pages = [{ ...(editor as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 }];
    editor.editHeaderFooter('header', 0);
    expect(editor.commentTarget()).toBeNull();
    expect(editor.addComment('Not allowed')).toBeNull();
    expect(editor.comments()).toEqual([]); expect(editor.isModified()).toBe(false);
    done();
  });

  it('QA20K-00799 anchors a comment to text inside a table cell after roundtrip', async () => {
    const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Q4 target</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { editor, view, done } = await open(table);
    select(editor, 'target'); const id = editor.addComment('Verify this cell')!;
    const beforeRange = commentRanges(view.state.doc).get(id)!;
    const beforeStart = textStart(view.state.doc, 'target');
    expect(commentedText(view.state.doc, id)).toBe('target');
    expect(beforeRange).toMatchObject({ from: beforeStart, to: beforeStart + 'target'.length });
    expect(view.state.doc.textBetween(beforeRange.from, beforeRange.to)).toBe('target');
    const zip = await save(editor);
    const beforeRoundTrip = tableShape(await zip.file('word/document.xml')!.async('string'));
    expect(beforeRoundTrip).toEqual({ rowCount: 1, cellsPerRow: [1], gridWidths: [3000] });
    const reopened = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(reopened.doc.firstChild!.type.name).toBe('table');
    expect(reopened.doc.firstChild!.childCount).toBe(beforeRoundTrip.rowCount);
    expect(reopened.doc.firstChild!.child(0).childCount).toBe(beforeRoundTrip.cellsPerRow[0]);
    const afterRange = commentRanges(reopened.doc).get(id)!;
    const afterStart = textStart(reopened.doc, 'target');
    expect(commentedText(reopened.doc, id)).toBe('target');
    expect(afterRange).toMatchObject({ from: afterStart, to: afterStart + 'target'.length });
    expect(reopened.doc.textBetween(afterRange.from, afterRange.to)).toBe('target');
    const afterXmlZip = await JSZip.loadAsync(await writeDocx(reopened.doc, reopened.model));
    const afterRoundTrip = tableShape(await afterXmlZip.file('word/document.xml')!.async('string'));
    expect(afterRoundTrip).toEqual(beforeRoundTrip);
    done();
  });

  it('QA20K-00800 keeps two comments on the same text range as separate threads', async () => {
    const { editor, view, done } = await open(para('Policy deadline'));
    select(editor, 'deadline'); const first = editor.addComment('Confirm source')!;
    select(editor, 'deadline'); const second = editor.addComment('Confirm date')!;
    expect(first).not.toBe(second);
    expect(editor.comments().map(c => c.id)).toEqual([first, second]);
    expect(commentedText(view.state.doc, first)).toBe('deadline');
    expect(commentedText(view.state.doc, second)).toBe('deadline');
    const zip = await save(editor); const reopened = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(commentedText(reopened.doc, first)).toBe('deadline');
    expect(commentedText(reopened.doc, second)).toBe('deadline');
    done();
  });
});
