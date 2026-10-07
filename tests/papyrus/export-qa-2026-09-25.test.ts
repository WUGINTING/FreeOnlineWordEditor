// Writes documents for the QA 2026-09-25 fixes, so Word itself can confirm the result
// (scripts/inspect-qa-2026-09-25-in-word.ps1). Opt-in:
//   DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-qa-2026-09-25.test.ts
import { beforeAll, describe, it } from 'vitest';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { ensureList } from '../../src/papyrus/docx/numbering';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';
import { setLink } from '../../src/papyrus/editor/commands';

const out = process.env.DOCX_EXPORT;
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const H = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const para = (s: string) => `<w:p><w:r><w:t xml:space="preserve">${s}</w:t></w:r></w:p>`;

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function packageWith(body: string, rels: string[] = []) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (rels.length) {
    const path = 'word/_rels/document.xml.rels';
    const xml = (await zip.file(path)!.async('string'));
    zip.file(path, xml.replace('</Relationships>', rels.join('') + '</Relationships>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}
async function openEditor(body: string) {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await packageWith(body));
  return { editor, done: () => { editor.destroy(); host.remove(); } };
}
function paragraphAt(editor: DocxEditor, text: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((n, at) => { if (pos < 0 && n.type === schema.nodes.paragraph && n.textContent === text) pos = at; });
  return pos;
}
function makeList(editor: DocxEditor, text: string, kind: 'bullet' | 'decimal') {
  const numId = ensureList(editor.model.numbering, kind);
  const pos = paragraphAt(editor, text);
  editor.view!.dispatch(editor.view!.state.tr.setNodeMarkup(pos, undefined, { ...editor.view!.state.doc.nodeAt(pos)!.attrs, numId, ilvl: 0 }));
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

describe.skipIf(!out)('export documents for the QA 2026-09-25 fixes', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });

    // BUG-005: a bullet copied into a document whose list with the same numId is decimal.
    const source = await openEditor(para('Deliver signed contract'));
    const target = await openEditor(para('Existing action') + para('Append here'));
    makeList(source.editor, 'Deliver signed contract', 'bullet');
    makeList(target.editor, 'Existing action', 'decimal');
    pasteAtEnd(target.editor, 'Append here', copy(source.editor, 'Deliver signed contract'));
    writeFileSync(join(out!, 'xdoc-list.docx'), await writeDocx(target.editor.view!.state.doc, target.editor.model));

    // BUG-003: the same paste undone: no bullet list left in numbering.xml.
    target.editor.undo();
    writeFileSync(join(out!, 'xdoc-list-undone.docx'), await writeDocx(target.editor.view!.state.doc, target.editor.model));
    source.done(); target.done();

    // BUG-002: links set in the editor, one javascript: and one https:.
    const { doc, model } = await readDocx(await packageWith(para('bad link') + para('good link')));
    let state = EditorState.create({ schema, doc });
    const link = (from: number, to: number, href: string) => {
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
      setLink(href)(state, (tr) => (state = state.apply(tr)));
    };
    link(1, 9, 'javascript:alert(1)');
    link(11, 20, 'https://example.com/');
    writeFileSync(join(out!, 'unsafe-link.docx'), await writeDocx(state.doc, model));

    // BUG-002: a file that already has a javascript: hyperlink (and a mailto: one), saved again.
    const loaded = await readDocx(
      await packageWith(
        '<w:p><w:hyperlink r:id="rIdBad"><w:r><w:t>bad link</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink r:id="rIdMail"><w:r><w:t>mail link</w:t></w:r></w:hyperlink></w:p>',
        [`<Relationship Id="rIdBad" Type="${H}" Target="javascript:alert(1)" TargetMode="External"/>`, `<Relationship Id="rIdMail" Type="${H}" Target="mailto:qa@example.com" TargetMode="External"/>`],
      ),
    );
    const bytes = await writeDocx(loaded.doc, loaded.model);
    const rels = await (await JSZip.loadAsync(bytes)).file('word/_rels/document.xml.rels')!.async('string');
    if (rels.includes('javascript:')) throw new Error('javascript: target written');
    writeFileSync(join(out!, 'loaded-unsafe-link.docx'), bytes);
  });
});
