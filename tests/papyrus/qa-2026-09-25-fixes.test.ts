// Fixes for the QA findings of 2026-09-25 (BUG-002 unsafe link targets, BUG-003 numbering left
// behind by an undone paste, BUG-005 a list copied between documents with colliding numIds).
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { ensureList } from '../../src/papyrus/docx/numbering';
import { isSafeHref, safeHref } from '../../src/papyrus/docx/links';
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

async function packageWith(body: string, rels: string[] = []) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (rels.length) {
    const path = 'word/_rels/document.xml.rels';
    const xml = (await zip.file(path)?.async('string')) ?? '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    zip.file(path, xml.replace('</Relationships>', rels.join('') + '</Relationships>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}
async function openEditor(body: string) {
  const bytes = await packageWith(body);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(bytes);
  return { editor, done: () => { editor.destroy(); host.remove(); } };
}
function paragraphAt(editor: DocxEditor, text: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((n, at) => { if (pos < 0 && n.type === schema.nodes.paragraph && n.textContent === text) pos = at; });
  if (pos < 0) throw new Error(`missing paragraph: ${text}`);
  return pos;
}
function makeList(editor: DocxEditor, text: string, kind: 'bullet' | 'decimal') {
  const numId = ensureList(editor.model.numbering, kind);
  const pos = paragraphAt(editor, text);
  editor.view!.dispatch(editor.view!.state.tr.setNodeMarkup(pos, undefined, { ...editor.view!.state.doc.nodeAt(pos)!.attrs, numId, ilvl: 0 }));
  return numId;
}
function copy(editor: DocxEditor, text: string) {
  const pos = paragraphAt(editor, text);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos + 1, pos + 1 + text.length)));
  const values: Record<string, string> = {};
  const event = new Event('copy', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { clearData: () => {}, setData: (k: string, v: string) => (values[k] = v) } });
  editor.view!.dom.dispatchEvent(event);
  return { html: values['text/html'] ?? '', text: values['text/plain'] ?? '' };
}
function pasteAtEnd(editor: DocxEditor, text: string, payload: { html: string; text: string }) {
  const pos = paragraphAt(editor, text) + 1 + text.length;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
  const event = new Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { files: [], getData: (k: string) => (k === 'text/html' ? payload.html : k === 'text/plain' ? payload.text : '') } });
  editor.view!.dom.dispatchEvent(event);
}
function listParagraphs(editor: DocxEditor) {
  const out: { text: string; numId: string }[] = [];
  editor.view!.state.doc.forEach((n) => n.attrs.numId && out.push({ text: n.textContent, numId: n.attrs.numId }));
  return out;
}

describe('BUG-002 one allowlist for link targets', () => {
  it('allows web, mail, phone, anchors and relative targets; refuses script, file and UNC targets', () => {
    for (const ok of ['https://example.com', 'HTTP://x', 'mailto:a@b.c', 'tel:+886', '#_Toc1', 'report.docx', '../a/b.pdf', 'a/b:c'])
      expect(isSafeHref(ok), ok).toBe(true);
    for (const bad of ['javascript:alert(1)', ' JavaScript:alert(1)', 'java\tscript:alert(1)', 'java\nscript:x', '\u0001javascript:x', 'vbscript:x', 'data:text/html,x', 'file:///C:/Windows/calc.exe', '\\\\server\\share', '//evil/x', 'C:\\x.exe', ''])
      expect(isSafeHref(bad), JSON.stringify(bad)).toBe(false);
    expect(safeHref('java\tscript:alert(1)')).toBe('#');
  });

  it('a pasted javascript: link keeps its text but not its target, and is saved as plain text', async () => {
    const div = document.createElement('div');
    div.innerHTML = '<p><a href="javascript:alert(1)">bad</a> <a href="https://example.com">good</a></p>';
    const doc = PMDOMParser.fromSchema(schema).parse(div);
    const links: string[] = [];
    doc.descendants((n) => { const m = schema.marks.link.isInSet(n.marks); if (m) links.push(m.attrs.href); });
    expect(doc.textContent).toBe('bad good');
    expect(links).toEqual(['', 'https://example.com']);
    const { model } = await readDocx(await packageWith(para('x')));
    const zip = await JSZip.loadAsync(await writeDocx(doc, model));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect((xml.match(/<w:hyperlink\b/g) ?? []).length).toBe(1);
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain('Target="https://example.com"');
  });

  it('a javascript: link in a loaded file is saved as plain text; safe links (external and anchors) stay links', async () => {
    const H = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const bytes = await packageWith(
      '<w:p><w:hyperlink r:id="rIdBad"><w:r><w:t>bad</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdMail"><w:r><w:t>mail</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="_Top"><w:r><w:t>top</w:t></w:r></w:hyperlink></w:p>',
      [`<Relationship Id="rIdBad" Type="${H}" Target="javascript:alert(1)" TargetMode="External"/>`, `<Relationship Id="rIdMail" Type="${H}" Target="mailto:a@b.c" TargetMode="External"/>`],
    );
    const { doc, model } = await readDocx(bytes);
    const zip = await JSZip.loadAsync(await writeDocx(doc, model));
    const xml = await zip.file('word/document.xml')!.async('string');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).not.toContain('javascript:');
    expect(rels).toContain('Target="mailto:a@b.c"');
    expect(xml).toContain('<w:hyperlink w:anchor="_Top">');
    expect((xml.match(/<w:hyperlink\b/g) ?? []).length).toBe(2);
    expect(xml).toMatch(/<w:r><w:t>bad<\/w:t><\/w:r>/);
  });

  it('links in a loaded file to a network share or a local file are saved as they were, but never clickable in the page', async () => {
    const H = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const unc = '\\\\fileserver\\公文\\114\\簽呈.docx';
    const file = 'file:///D:/公文/附件.pdf';
    const bytes = await packageWith(
      '<w:p><w:hyperlink r:id="rIdUnc"><w:r><w:t>share</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdFile"><w:r><w:t>file</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdVb"><w:r><w:t>vb</w:t></w:r></w:hyperlink></w:p>',
      [
        `<Relationship Id="rIdUnc" Type="${H}" Target="${unc}" TargetMode="External"/>`,
        `<Relationship Id="rIdFile" Type="${H}" Target="${file}" TargetMode="External"/>`,
        `<Relationship Id="rIdVb" Type="${H}" Target="vbscript:msgbox(1)" TargetMode="External"/>`,
      ],
    );
    const { doc, model } = await readDocx(bytes);
    const zip = await JSZip.loadAsync(await writeDocx(doc, model));
    const xml = await zip.file('word/document.xml')!.async('string');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain(`Target="${unc}"`);
    expect(rels).toContain(`Target="${file}"`);
    expect(rels).not.toContain('vbscript:');
    expect((xml.match(/<w:hyperlink\b/g) ?? []).length).toBe(2);
    expect(safeHref(unc)).toBe('#');
    expect(safeHref(file)).toBe('#');
  });
});

describe('BUG-003 / BUG-005 lists in the real editor', () => {
  it('a list copied inside one document continues that list', async () => {
    const d = await openEditor(para('one') + para('two') + para('end'));
    const numId = makeList(d.editor, 'one', 'bullet');
    makeList(d.editor, 'two', 'bullet');
    pasteAtEnd(d.editor, 'end', copy(d.editor, 'two'));
    expect(listParagraphs(d.editor).map((p) => p.numId)).toEqual([numId, numId, numId]);
    d.done();
  });

  it('a list copied from another document gets its own list looking like the source, and undo takes it back out', async () => {
    const source = await openEditor(para('Deliver signed contract'));
    const target = await openEditor(para('Existing action') + para('Append here'));
    makeList(source.editor, 'Deliver signed contract', 'bullet');
    const targetNum = makeList(target.editor, 'Existing action', 'decimal');
    const before = structuredClone(target.editor.model.numbering.nums);
    pasteAtEnd(target.editor, 'Append here', copy(source.editor, 'Deliver signed contract'));
    const pasted = listParagraphs(target.editor).find((p) => p.text === 'Deliver signed contract')!;
    expect(pasted.numId).not.toBe(targetNum);
    const n = target.editor.model.numbering;
    expect(n.abstracts[n.nums[pasted.numId].abstractId].levels[0].fmt).toBe('bullet');
    expect(n.abstracts[n.nums[targetNum].abstractId].levels[0].fmt).toBe('decimal');

    target.editor.undo();
    expect(target.editor.model.numbering.nums).toEqual(before);
    const zip = await JSZip.loadAsync(await target.editor.save());
    const numbering = await zip.file('word/numbering.xml')!.async('string');
    expect((numbering.match(/<w:num\b/g) ?? []).length).toBe(1);
    expect(numbering).not.toContain('w:val="bullet"');
    source.done(); target.done();
  });
});
