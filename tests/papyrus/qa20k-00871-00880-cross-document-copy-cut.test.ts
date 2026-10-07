import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const para = (s: string) => `<w:p><w:r><w:t xml:space="preserve">${s}</w:t></w:r></w:p>`;

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const relPath = 'word/_rels/document.xml.rels';
  const rels = await zip.file(relPath)!.async('string');
  zip.file(relPath, rels.replace('</Relationships>', '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://policy.example/" TargetMode="External"/></Relationships>'));
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { model } = await readDocx(bytes);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host); await editor.open(bytes);
  return { editor, model, host, done: () => { editor.destroy(); host.remove(); } };
}

function textStart(editor: DocxEditor, text: string) {
  let found = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text!.includes(text)) found = pos + node.text!.indexOf(text);
  });
  if (found < 0) throw new Error(`missing text: ${text}`);
  return found;
}

function select(editor: DocxEditor, text: string, from = 0, length = text.length) {
  const start = textStart(editor, text) + from;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start, start + length)));
}

type Payload = { html: string; text: string };
function clipboardEvent(editor: DocxEditor, type: 'copy' | 'cut') {
  const values: Record<string, string> = {};
  const transfer = {
    clearData: () => { for (const key of Object.keys(values)) delete values[key]; },
    setData: (kind: string, value: string) => { values[kind] = value; },
    getData: (kind: string) => values[kind] ?? '',
  };
  const event = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: transfer });
  editor.view!.dom.dispatchEvent(event);
  return { event, payload: { html: values['text/html'] ?? '', text: values['text/plain'] ?? '' } as Payload };
}

function paste(editor: DocxEditor, payload: Payload) {
  const event = new Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { files: [], getData: (kind: string) => kind === 'text/html' ? payload.html : kind === 'text/plain' ? payload.text : '' } });
  editor.view!.dom.dispatchEvent(event);
  return event;
}

function setCaret(editor: DocxEditor, pos: number) {
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
}

function marks(editor: DocxEditor, text: string) {
  const found: string[][] = [];
  editor.view!.state.doc.descendants((n) => { if (n.isText && n.text === text) found.push(n.marks.map(m => m.type.name).sort()); });
  return found;
}

function hrefs(editor: DocxEditor, text: string) {
  const found: string[] = [];
  editor.view!.state.doc.descendants((n) => {
    if (n.isText && n.text === text) {
      const link = n.marks.find(m => m.type === schema.marks.link);
      if (link) found.push(link.attrs.href);
    }
  });
  return found;
}

function paragraphs(editor: DocxEditor) { const rows: any[] = []; editor.view!.state.doc.forEach(n => rows.push(n)); return rows; }

describe('QA20K-00871–00880 cross-document copy and cut workflows', () => {
  it('QA20K-00871 copies a bold approval phrase into another document with its mark intact', async () => {
    const source = await open('<w:p><w:r><w:t>Approve </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>budget</w:t></w:r><w:r><w:t> today</w:t></w:r></w:p>');
    const target = await open(para('Decision: '));
    select(source.editor, 'budget');
    const { event, payload } = clipboardEvent(source.editor, 'copy');
    expect(event.defaultPrevented).toBe(true);
    setCaret(target.editor, textStart(target.editor, 'Decision: ') + 'Decision: '.length);
    paste(target.editor, payload);
    expect(target.editor.view!.state.doc.textContent).toBe('Decision: budget');
    expect(marks(target.editor, 'budget')).toEqual([['bold']]);
    expect(source.editor.view!.state.doc.textContent).toBe('Approve budget today');
    source.done(); target.done();
  });

  it('QA20K-00872 copies only the selected middle words and leaves source neighbors untouched', async () => {
    const source = await open(para('Please update the vendor address before sending.'));
    const target = await open(para('Email: '));
    select(source.editor, 'Please update the vendor address before sending.', 7, 'update the vendor address'.length);
    const { payload } = clipboardEvent(source.editor, 'copy');
    setCaret(target.editor, textStart(target.editor, 'Email: ') + 7);
    paste(target.editor, payload);
    expect(target.editor.view!.state.doc.textContent).toBe('Email: update the vendor address');
    expect(source.editor.view!.state.doc.textContent).toBe('Please update the vendor address before sending.');
    source.done(); target.done();
  });

  it('QA20K-00873 copies a linked label across documents with the destination intact', async () => {
    const source = await open('<w:p><w:r><w:t>See </w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>policy</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>');
    const target = await open(para('Reference: '));
    select(source.editor, 'policy');
    const { payload } = clipboardEvent(source.editor, 'copy');
    expect(payload.html).toContain('href="https://policy.example/"');
    setCaret(target.editor, textStart(target.editor, 'Reference: ') + 11);
    paste(target.editor, payload);
    let linkAttrs: any[] = [];
    target.editor.view!.state.doc.descendants(n => { if (n.isText && n.text === 'policy') linkAttrs = n.marks.filter(m => m.type.name === 'link').map(m => m.attrs); });
    expect(target.editor.view!.state.doc.textContent).toBe('Reference: policy');
    expect(linkAttrs).toEqual([{ href: 'https://policy.example/', attrs: null }]);
    source.done(); target.done();
  });

  it('QA20K-00874 cuts a linked phrase from the source and inserts it as a link in the destination', async () => {
    const source = await open('<w:p><w:r><w:t>Open </w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>handbook</w:t></w:r></w:hyperlink><w:r><w:t> now</w:t></w:r></w:p>');
    const target = await open(para('For details: '));
    expect(hrefs(source.editor, 'handbook')).toEqual(['https://policy.example/']);
    select(source.editor, 'handbook');
    const { event, payload } = clipboardEvent(source.editor, 'cut');
    expect(event.defaultPrevented).toBe(true);
    expect(source.editor.view!.state.doc.textContent).toBe('Open  now');
    setCaret(target.editor, textStart(target.editor, 'For details: ') + 13);
    paste(target.editor, payload);
    expect(target.editor.view!.state.doc.textContent).toBe('For details: handbook');
    expect(marks(target.editor, 'handbook')).toEqual([['link']]);
    expect(hrefs(target.editor, 'handbook')).toEqual(['https://policy.example/']);
    source.done(); target.done();
  });

  it('QA20K-00875 restores cut wording and bold formatting with one source undo', async () => {
    const source = await open('<w:p><w:r><w:t>Final </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>approval</w:t></w:r><w:r><w:t> needed</w:t></w:r></w:p>');
    const target = await open(para('Archive: '));
    select(source.editor, 'approval');
    const { payload } = clipboardEvent(source.editor, 'cut');
    setCaret(target.editor, textStart(target.editor, 'Archive: ') + 9);
    paste(target.editor, payload);
    expect(source.editor.view!.state.doc.textContent).toBe('Final  needed');
    expect(source.editor.undo()).toBe(true);
    expect(source.editor.view!.state.doc.textContent).toBe('Final approval needed');
    expect(marks(source.editor, 'approval')).toEqual([['bold', 'run']]);
    expect(target.editor.view!.state.doc.textContent).toBe('Archive: approval');
    source.done(); target.done();
  });

  it('QA20K-00876 copies two selected paragraphs without merging their boundary', async () => {
    const source = await open(para('Agenda:') + para('1. Budget review') + para('2. Hiring plan') + para('Next steps'));
    const target = await open(para('Meeting notes:'));
    const start = textStart(source.editor, '1. Budget review');
    const end = textStart(source.editor, '2. Hiring plan') + '2. Hiring plan'.length;
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(TextSelection.create(source.editor.view!.state.doc, start, end)));
    const { payload } = clipboardEvent(source.editor, 'copy');
    setCaret(target.editor, textStart(target.editor, 'Meeting notes:') + 'Meeting notes:'.length);
    paste(target.editor, payload);
    expect(paragraphs(target.editor).map(n => n.textContent)).toEqual(['Meeting notes:1. Budget review', '2. Hiring plan']);
    expect(source.editor.view!.state.doc.textContent).toBe('Agenda:1. Budget review2. Hiring planNext steps');
    source.done(); target.done();
  });

  it('QA20K-00877 copies a whole comparison table as rows and cells into another document', async () => {
    const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Owner</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Budget</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Finance</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const source = await open(table + para('After table'));
    const target = await open(para('Copied report') );
    const tablePos = 0;
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, tablePos)));
    const { payload } = clipboardEvent(source.editor, 'copy');
    setCaret(target.editor, textStart(target.editor, 'Copied report') + 'Copied report'.length);
    paste(target.editor, payload);
    const inserted = paragraphs(target.editor).find(n => n.type === schema.nodes.table);
    expect(inserted).toBeTruthy();
    expect(inserted.childCount).toBe(2);
    expect(inserted.child(0).childCount).toBe(2);
    expect(inserted.child(1).child(0).textContent).toBe('Budget');
    expect(inserted.child(1).child(1).textContent).toBe('Finance');
    source.done(); target.done();
  });

  it('QA20K-00878 copies text around a hidden bookmark without removing the source marker', async () => {
    const body = '<w:p><w:r><w:t>See policy</w:t></w:r><w:bookmarkStart w:id="7" w:name="PolicyRef"/><w:r><w:t> section</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>';
    const source = await open(body);
    const target = await open(para('Reference:'));
    const bookmarkNodes = () => {
      const found: any[] = [];
      source.editor.view!.state.doc.descendants(n => { if (n.type === schema.nodes.raw_inline && /bookmark(Start|End)/.test(n.attrs.label)) found.push(n.attrs); });
      return found;
    };
    const beforeBookmarks = bookmarkNodes();
    const start = textStart(source.editor, 'See policy');
    let end = -1;
    source.editor.view!.state.doc.descendants((n, pos) => { if (end < 0 && n.isText && n.text!.includes(' section')) end = pos + n.text!.indexOf(' section') + ' section'.length; });
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(TextSelection.create(source.editor.view!.state.doc, start, end)));
    const { payload } = clipboardEvent(source.editor, 'copy');
    setCaret(target.editor, textStart(target.editor, 'Reference:') + 'Reference:'.length);
    paste(target.editor, payload);
    expect(bookmarkNodes()).toEqual(beforeBookmarks);
    expect(target.editor.view!.state.doc.textContent).toBe('Reference:See policy section');
    source.done(); target.done();
  });

  it('QA20K-00879 cuts a selected phrase and undoes without changing the adjacent hyperlink', async () => {
    const body = '<w:p><w:r><w:t>Draft </w:t></w:r><w:r><w:t>appendix</w:t></w:r><w:r><w:t>;</w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>source</w:t></w:r></w:hyperlink></w:p>';
    const source = await open(body);
    select(source.editor, 'appendix');
    clipboardEvent(source.editor, 'cut');
    expect(source.editor.view!.state.doc.textContent).toBe('Draft ;source');
    expect(source.editor.undo()).toBe(true);
    expect(source.editor.view!.state.doc.textContent).toBe('Draft appendix;source');
    let href = '';
    source.editor.view!.state.doc.descendants(n => { if (n.isText && n.text === 'source') href = n.marks.find(m => m.type.name === 'link')?.attrs.href ?? ''; });
    expect(href).toBe('https://policy.example/');
    source.done();
  });

  it('QA20K-00880 pastes a copied table and saves it with both rows and column boundaries', async () => {
    const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="3500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Task</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Review</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Open</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const source = await open(table);
    const target = await open(para('Appendix'));
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, 0)));
    const { payload } = clipboardEvent(source.editor, 'copy');
    setCaret(target.editor, textStart(target.editor, 'Appendix') + 8);
    paste(target.editor, payload);
    const zip = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    const copy = xml.slice(xml.lastIndexOf('<w:tbl>'), xml.lastIndexOf('</w:tbl>') + '</w:tbl>'.length);
    expect((copy.match(/<w:tr>/g) ?? [])).toHaveLength(2);
    expect((copy.match(/<w:tc>/g) ?? [])).toHaveLength(4);
    expect(copy).toContain('<w:t>Review</w:t>');
    expect(copy).toContain('<w:t>Open</w:t>');
    source.done(); target.done();
  });
});
