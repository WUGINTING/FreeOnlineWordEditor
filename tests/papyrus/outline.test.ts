// Navigation pane headings and table-of-contents page numbers (editor/outline.ts).
import { beforeAll, describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { documentOutline, pageReferences, updatePageReferences } from '../../src/papyrus/editor/outline';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>
  <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:basedOn w:val="1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
</w:styles>`;
const p = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
// A table of contents entry: its page number is a PAGEREF field result pointing at a bookmark.
const tocEntry = (text: string, bookmark: string, page: string) =>
  `<w:p><w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${bookmark} \\h </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
const heading = (text: string, bookmark: string, style = '1') =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:bookmarkStart w:id="0" w:name="${bookmark}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;

async function read(body: string) {
  const zip = blankPackage();
  zip.file('word/styles.xml', STYLES);
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const BODY =
  tocEntry('一、總則', '_Toc1', '9') + tocEntry('二、細則', '_Toc2', '9') +
  heading('一、總則', '_Toc1') + p('內文') + heading('（一）範圍', '_Toc3', '2') + p('第三層', '<w:outlineLvl w:val="2"/>') +
  heading('二、細則', '_Toc2') + p('   ', '<w:pStyle w:val="1"/>');

describe('navigation pane', () => {
  it('lists headings from styles (outline level, basedOn) and from the paragraph itself; not body text or empty headings', async () => {
    const { doc, model } = await readDocx(await read(BODY));
    expect(documentOutline(doc, model.styles).map((h) => [h.level, h.text])).toEqual([
      [0, '一、總則'], [1, '（一）範圍'], [2, '第三層'], [0, '二、細則'],
    ]);
  });

  it('the editor goes to a heading (cursor at its start)', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await read(BODY));
    const h = ed.outline()[3];
    ed.goToHeading(h.pos);
    expect(ed.view!.state.selection.$from.parent.textContent).toBe('二、細則');
    ed.destroy();
    host.remove();
  });
});

describe('table of contents page numbers', () => {
  it('updates each PAGEREF result from the page of its bookmark, keeping its formatting; one undoable step', async () => {
    const { doc } = await readDocx(await read(BODY));
    expect(pageReferences(doc).map((r) => [r.bookmark, doc.textBetween(r.from, r.to)])).toEqual([['_Toc1', '9'], ['_Toc2', '9']]);
    const state = EditorState.create({ doc });
    const tr = state.tr;
    const page = new Map([['一、總則', 1], ['二、細則', 12]]);
    const n = updatePageReferences(tr, (pos) => page.get(tr.doc.resolve(pos).parent.textContent) ?? null);
    expect(n).toBe(2);
    const refs = pageReferences(tr.doc);
    expect(refs.map((r) => tr.doc.textBetween(r.from, r.to))).toEqual(['1', '12']);
    // Still bold, still the field's result.
    expect(tr.doc.nodeAt(refs[1].from)!.marks.map((m) => m.type.name)).toEqual(expect.arrayContaining(['bold', 'fieldResult']));
    // Nothing to change the second time.
    expect(updatePageReferences(tr, (pos) => page.get(tr.doc.resolve(pos).parent.textContent) ?? null)).toBe(0);
  });
});
