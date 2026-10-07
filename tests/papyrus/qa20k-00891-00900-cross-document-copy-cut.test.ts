import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { ensureList } from '../../src/papyrus/docx/numbering';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const para = (s: string) => `<w:p><w:r><w:t xml:space="preserve">${s}</w:t></w:r></w:p>`;
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { model } = await readDocx(bytes);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host); await editor.open(bytes);
  return { editor, model, host, done: () => { editor.destroy(); host.remove(); } };
}

function textStart(editor: DocxEditor, text: string) {
  let found = -1;
  editor.view!.state.doc.descendants((node, pos) => { if (found < 0 && node.isText && node.text!.includes(text)) found = pos + node.text!.indexOf(text); });
  if (found < 0) throw new Error(`missing text: ${text}`);
  return found;
}
function selectText(editor: DocxEditor, text: string, from = 0, size = text.length) {
  const start = textStart(editor, text) + from;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start, start + size)));
}
type Payload = { html: string; text: string };
function clipboard(editor: DocxEditor, type: 'copy' | 'cut') {
  const values: Record<string, string> = {};
  const transfer = { clearData: () => { for (const key of Object.keys(values)) delete values[key]; }, setData: (k: string, v: string) => { values[k] = v; } };
  const event = new Event(type, { bubbles: true, cancelable: true }); Object.defineProperty(event, 'clipboardData', { value: transfer });
  editor.view!.dom.dispatchEvent(event);
  return { event, payload: { html: values['text/html'] ?? '', text: values['text/plain'] ?? '' } as Payload };
}
function paste(editor: DocxEditor, payload: Payload) {
  const event = new Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { files: [], getData: (k: string) => k === 'text/html' ? payload.html : k === 'text/plain' ? payload.text : '' } });
  editor.view!.dom.dispatchEvent(event); return event;
}
function caret(editor: DocxEditor, pos: number) { editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos))); }
function children(editor: DocxEditor) { const out: any[] = []; editor.view!.state.doc.forEach(n => out.push(n)); return out; }
function markNames(editor: DocxEditor, text: string) {
  const out: string[][] = [];
  editor.view!.state.doc.descendants(n => { if (n.isText && n.text === text) out.push(n.marks.map(m => m.type.name).sort()); });
  return out;
}
async function expectEmbeddedImageBytes(zip: JSZip, documentXml: string, expectedBase64: string) {
  const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const embeds = [...documentXml.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/?\s*>/g)].map(m => m[1]);
  expect(embeds.length).toBeGreaterThan(0);
  const relationships = [...relsXml.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map(m => m[0]);
  for (const embed of embeds) {
    const relationship = relationships.find(rel => {
      const id = /\bId="([^"]+)"/.exec(rel)?.[1];
      const type = /\bType="([^"]+)"/.exec(rel)?.[1];
      return id === embed && type?.endsWith('/image');
    });
    expect(relationship, `image relationship for ${embed}`).toBeTruthy();
    const target = /\bTarget="([^"]+)"/.exec(relationship!)?.[1];
    expect(target).toBeTruthy();
    const mediaPath = target!.startsWith('/') ? target!.slice(1) : `word/${target}`;
    const media = zip.file(mediaPath);
    expect(media, `media part ${mediaPath} referenced by ${embed}`).toBeTruthy();
    expect(await media!.async('base64')).toBe(expectedBase64);
  }
}

function makeList(editor: DocxEditor, text: string, kind: 'bullet' | 'decimal', level = 0) {
  const numId = ensureList(editor.model.numbering, kind);
  let pos = -1;
  editor.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.type === schema.nodes.paragraph && node.textContent === text) pos = at; });
  if (pos < 0) throw new Error(`missing paragraph: ${text}`);
  editor.view!.dispatch(editor.view!.state.tr.setNodeMarkup(pos, undefined, { ...editor.view!.state.doc.nodeAt(pos)!.attrs, numId, ilvl: level }));
  return numId;
}

describe('QA20K-00891–00900 cross-document copy and cut workflows', () => {
  it('QA20K-00891 pastes a bullet item into a destination that already uses the same numbering id', async () => {
    const source = await open(para('Deliver signed contract'));
    const target = await open(para('Existing action') + para('Append here'));
    const srcNum = makeList(source.editor, 'Deliver signed contract', 'bullet');
    const targetNum = makeList(target.editor, 'Existing action', 'decimal');
    expect(srcNum).toBe(targetNum); // force a realistic cross-DOCX numId collision
    selectText(source.editor, 'Deliver signed contract');
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Append here') + 'Append here'.length);
    paste(target.editor, payload);
    const pasted = children(target.editor).find(n => n.textContent === 'Deliver signed contract');
    expect(pasted?.attrs.numId).toBeTruthy();
    const targetModel = target.editor.model;
    expect(targetModel.numbering.abstracts[targetModel.numbering.nums[pasted.attrs.numId].abstractId].levels[0].fmt).toBe('bullet');
    source.done(); target.done();
  });

  it('QA20K-00892 copies a nested numbered task and retains its nesting level in the destination', async () => {
    const source = await open(para('Prepare report') + para('Attach receipts'));
    const target = await open(para('Travel claim'));
    makeList(source.editor, 'Prepare report', 'decimal', 0);
    makeList(source.editor, 'Attach receipts', 'decimal', 1);
    selectText(source.editor, 'Attach receipts');
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Travel claim') + 'Travel claim'.length);
    paste(target.editor, payload);
    const pasted = children(target.editor).find(n => n.textContent === 'Attach receipts');
    expect(pasted?.attrs.numId).toBeTruthy();
    expect(pasted?.attrs.ilvl).toBe(1);
    expect(target.editor.model.numbering.abstracts[target.editor.model.numbering.nums[pasted!.attrs.numId].abstractId].levels[0].fmt).toBe('decimal');
    source.done(); target.done();
  });

  it('QA20K-00893 copies a heading paragraph with its heading style into the report', async () => {
    const source = await open('<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Risk summary</w:t></w:r></w:p>');
    const target = await open(para('Appendix'));
    const headingPos = 0;
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, headingPos)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Appendix') + 8);
    paste(target.editor, payload);
    const heading = children(target.editor).find(n => n.textContent === 'Risk summary');
    expect(heading?.attrs.styleId).toBe('Heading2');
    expect(children(target.editor).map(n => n.textContent)).toContain('Appendix');
    source.done(); target.done();
  });

  it('QA20K-00894 copies an inline PNG to another document and writes its media part', async () => {
    const source = await open(para('Logo area'));
    const target = await open(para('Appendix:'));
    const pos = textStart(source.editor, 'Logo area') + 5;
    const image = schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 42, height: 21, alt: 'Quarterly logo' });
    const tr = source.editor.view!.state.tr.insert(pos, image);
    source.editor.view!.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Appendix:') + 9);
    paste(target.editor, payload);
    const copied = children(target.editor)[0].content.content.find((n: any) => n.type === schema.nodes.image);
    expect(copied?.attrs).toMatchObject({ src: `data:image/png;base64,${PNG}`, width: 42, height: 21, alt: 'Quarterly logo' });
    const zip = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('descr="Quarterly logo"');
    await expectEmbeddedImageBytes(zip, xml, PNG);
    source.done(); target.done();
  });

  it('QA20K-00895 copies a page break between report sections as a DOCX page break', async () => {
    const source = await open('<w:p><w:r><w:t>Cover</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>Body</w:t></w:r></w:p>');
    const target = await open(para('Attachment'));
    const breakNode = children(source.editor).find(n => n.type === schema.nodes.page_break);
    expect(breakNode).toBeTruthy();
    let breakPos = -1; source.editor.view!.state.doc.descendants((n, p) => { if (breakPos < 0 && n.type === schema.nodes.page_break) breakPos = p; });
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, breakPos)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Attachment') + 10);
    paste(target.editor, payload);
    expect(children(target.editor).some(n => n.type === schema.nodes.page_break)).toBe(true);
    const xml = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model)).then(z => z.file('word/document.xml')!.async('string'));
    expect(xml).toMatch(/<w:br w:type="page"\s*\/>/);
    source.done(); target.done();
  });

  it('QA20K-00896 copies a vertically merged comparison table and retains the merge in DOCX', async () => {
    const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Plan</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Q1</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>Q2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const source = await open(table);
    const target = await open(para('Quarterly comparison'));
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, 0)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Quarterly comparison') + 'Quarterly comparison'.length);
    paste(target.editor, payload);
    const xml = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model)).then(z => z.file('word/document.xml')!.async('string'));
    const copied = xml.slice(xml.lastIndexOf('<w:tbl>'), xml.lastIndexOf('</w:tbl>'));
    expect((copied.match(/<w:vMerge/g) ?? [])).toHaveLength(2);
    expect(copied).toContain('<w:vMerge w:val="restart"');
    expect(copied).toMatch(/<w:vMerge\s*\/>/);
    expect(copied).toContain('<w:t>Q2</w:t>');
    source.done(); target.done();
  });

  it('QA20K-00897 carries table cell shading and grid widths into the saved destination DOCX', async () => {
    const table = '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:fill="DDEEFF"/></w:tcPr><w:p><w:r><w:t>Decision</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Approved</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const source = await open(table);
    const target = await open(para('Minutes'));
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, 0)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Minutes') + 7);
    paste(target.editor, payload);
    const xml = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model)).then(z => z.file('word/document.xml')!.async('string'));
    const copied = xml.slice(xml.lastIndexOf('<w:tbl>'), xml.lastIndexOf('</w:tbl>'));
    expect(copied).toMatch(/<w:shd[^>]*w:fill="DDEEFF"/);
    expect(copied).toContain('<w:gridCol w:w="2400"');
    expect(copied).toContain('<w:gridCol w:w="3600"');
    source.done(); target.done();
  });

  it('QA20K-00898 cuts a formatted phrase across runs while leaving neighboring text in place', async () => {
    const source = await open('<w:p><w:r><w:t>Use </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>approved</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t> wording</w:t></w:r><w:r><w:t> now.</w:t></w:r></w:p>');
    const target = await open(para('Replacement: '));
    const start = textStart(source.editor, 'Use ') + 4;
    const end = textStart(source.editor, ' wording') + ' wording'.length;
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(TextSelection.create(source.editor.view!.state.doc, start, end)));
    const { event, payload } = clipboard(source.editor, 'cut');
    expect(event.defaultPrevented).toBe(true);
    expect(source.editor.view!.state.doc.textContent).toBe('Use  now.');
    caret(target.editor, textStart(target.editor, 'Replacement:') + 'Replacement: '.length);
    paste(target.editor, payload);
    expect(target.editor.view!.state.doc.textContent).toBe('Replacement: approved wording');
    expect(markNames(target.editor, 'approved').some(ms => ms.includes('bold'))).toBe(true);
    expect(markNames(target.editor, ' wording').some(ms => ms.includes('italic'))).toBe(true);
    expect(source.editor.view!.state.doc.textContent).toContain('now.');
    source.done(); target.done();
  });

  it('QA20K-00899 copies a chart together with its caption paragraph and verifies DOCX picture description', async () => {
    const source = await open(para('Chart: — April sales'));
    const target = await open(para('Attachment:'));
    const pos = textStart(source.editor, 'Chart: — April sales') + 6;
    const node = schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 64, height: 32, alt: 'Monthly sales chart' });
    const tr = source.editor.view!.state.tr.insert(pos, node);
    source.editor.view!.dispatch(tr.setSelection(NodeSelection.create(tr.doc, 0)));
    const { payload } = clipboard(source.editor, 'copy');
    caret(target.editor, textStart(target.editor, 'Attachment:') + 'Attachment:'.length);
    paste(target.editor, payload);
    const copiedParagraph = children(target.editor).find(n => n.type === schema.nodes.paragraph && n.content.content.some((c: any) => c.type === schema.nodes.image));
    expect(copiedParagraph?.textContent).toBe('Chart: — April sales');
    expect(copiedParagraph?.content.content.map((c: any) => c.type.name)).toEqual(['text', 'image', 'text']);
    const copied = copiedParagraph?.content.content.find((n: any) => n.type === schema.nodes.image);
    expect(copied?.attrs).toMatchObject({ src: `data:image/png;base64,${PNG}`, width: 64, height: 32, alt: 'Monthly sales chart' });
    const zip = await JSZip.loadAsync(await writeDocx(target.editor.view!.state.doc, target.model));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('descr="Monthly sales chart"');
    await expectEmbeddedImageBytes(zip, xml, PNG);
    source.done(); target.done();
  });

  it('QA20K-00900 cuts a table, pastes its cells elsewhere and can undo the source cut', async () => {
    const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="3500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Department</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Owner</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Finance</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Friday</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const source = await open(table + para('After table'));
    const target = await open(para('Appendix'));
    source.editor.view!.dispatch(source.editor.view!.state.tr.setSelection(NodeSelection.create(source.editor.view!.state.doc, 0)));
    expect(source.editor.view!.state.selection.empty).toBe(false);
    const { event, payload } = clipboard(source.editor, 'cut');
    expect(event.defaultPrevented).toBe(true);
    expect(children(source.editor).some(n => n.type === schema.nodes.table)).toBe(true);
    const cutSourceTable = children(source.editor).find(n => n.type === schema.nodes.table);
    expect(cutSourceTable?.childCount).toBe(2);
    expect(cutSourceTable && [0, 1].map(row => cutSourceTable.child(row).childCount)).toEqual([2, 2]);
    expect(cutSourceTable && [0, 1].map(row => [0, 1].map(col => cutSourceTable.child(row).child(col).textContent))).toEqual([
      ['', ''],
      ['', ''],
    ]);
    expect(source.editor.view!.state.doc.childCount).toBe(2);
    expect(source.editor.view!.state.doc.child(1).textContent).toBe('After table');
    caret(target.editor, textStart(target.editor, 'Appendix') + 8);
    paste(target.editor, payload);
    const destTable = children(target.editor).find(n => n.type === schema.nodes.table);
    expect(destTable?.childCount).toBe(2);
    expect(destTable && [0, 1].map(row => destTable.child(row).childCount)).toEqual([2, 2]);
    expect(destTable && [0, 1].map(row => [0, 1].map(col => destTable.child(row).child(col).textContent))).toEqual([
      ['Department', 'Owner'],
      ['Finance', 'Friday'],
    ]);
    expect(source.editor.undo()).toBe(true);
    const restoredTable = children(source.editor).find(n => n.type === schema.nodes.table);
    expect(restoredTable?.childCount).toBe(2);
    expect(restoredTable && [0, 1].map(row => restoredTable.child(row).childCount)).toEqual([2, 2]);
    expect(restoredTable && [0, 1].map(row => [0, 1].map(col => restoredTable.child(row).child(col).textContent))).toEqual([
      ['Department', 'Owner'],
      ['Finance', 'Friday'],
    ]);
    expect(source.editor.view!.state.doc.childCount).toBe(2);
    expect(source.editor.view!.state.doc.child(1).textContent).toBe('After table');
    expect(source.editor.view!.state.doc.textContent).toBe('DepartmentOwnerFinanceFridayAfter table');
    source.done(); target.done();
  });
});
