// persona-300 code review, the editor: missing fonts only from what the text uses, one undo over
// several parts, screen readers (deletions said once, headers / footers readable), the clipboard
// and the printout without the hidden helper text.
import { afterEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { schema } from '../../src/papyrus/editor/schema';
import { findMatches, replaceMatchesTr } from '../../src/papyrus/editor/search';
import { listClipboardSerializer } from '../../src/papyrus/editor/pasteLists';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { docxOf, domStubs, setup } from './p300Helpers';

domStubs();
afterEach(() => vi.restoreAllMocks());

const P = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

describe('fonts this computer lacks: only those the text uses', () => {
  it('a font of a style no paragraph uses is not reported; used styles, character styles and header runs are', async () => {
    const styles = await blankPackage().file('word/styles.xml')!.async('string');
    const extra = (zip: JSZip) =>
      zip.file(
        'word/styles.xml',
        styles.replace(
            '</w:styles>',
            '<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/><w:rPr><w:rFonts w:ascii="NoSuchFontA" w:eastAsia="NoSuchFontA"/></w:rPr></w:style>' +
              '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:rFonts w:eastAsia="BaseFontB"/></w:rPr></w:style>' +
              '<w:style w:type="paragraph" w:styleId="Used"><w:name w:val="Used"/><w:basedOn w:val="Base"/><w:rPr><w:rFonts w:ascii="UsedFontC"/></w:rPr></w:style>' +
              '<w:style w:type="character" w:styleId="Cs"><w:name w:val="Cs"/><w:rPr><w:rFonts w:ascii="CharFontD"/></w:rPr></w:style>' +
              '<w:style w:type="character" w:styleId="CsUnused"><w:name w:val="CsUnused"/><w:rPr><w:rFonts w:ascii="NoSuchFontE"/></w:rPr></w:style>' +
              '</w:styles>',
        ),
      );
    const src = await docxOf('<w:p><w:pPr><w:pStyle w:val="Used"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Cs"/></w:rPr><w:t>字</w:t></w:r></w:p>', extra);
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(src);
    editor.editHeaderFooter('header');
    const hv = editor.activeView!;
    hv.dispatch(hv.state.tr.insert(1, schema.text('頁首', [schema.marks.font.create({ family: 'HeaderFontF', eastAsia: null })])));
    editor.closeHeaderFooter();
    const fonts = editor.documentFonts();
    expect(fonts).toEqual(expect.arrayContaining(['UsedFontC', 'BaseFontB', 'CharFontD', 'HeaderFontF', 'Times New Roman', '標楷體']));
    expect(fonts).not.toContain('NoSuchFontA');
    expect(fonts).not.toContain('NoSuchFontE');
    editor.destroy();
    host.remove();
  });
});

describe('one undo over several parts', () => {
  async function replaced() {
    const d = await setup('<w:p><w:r><w:t>甲 2025</w:t></w:r></w:p>');
    d.editor.editHeaderFooter('header');
    const hv = d.editor.activeView!;
    hv.dispatch(hv.state.tr.insertText('頁首 2025', 1));
    d.editor.closeHeaderFooter();
    d.editor.changeSearchParts(
      d.editor.searchParts().map((p) => ({ id: p.id, build: (s: EditorState) => replaceMatchesTr(s, findMatches(s.doc, '2025'), '2026') })),
    );
    const texts = () => d.editor.searchParts().map((p) => p.doc.textContent);
    return { d, texts };
  }

  it('asking whether undo can run changes nothing', async () => {
    const { d, texts } = await replaced();
    const undoCommand = (d.editor as any).historyCommand('undo');
    expect(d.editor.can(undoCommand)).toBe(true);
    expect(texts()).toEqual(['甲 2026', '頁首 2026']);
    d.editor.undo();
    expect(texts()).toEqual(['甲 2025', '頁首 2025']);
    d.done();
  });

  it('typing in the body and undoing it: the next undo still undoes the header too', async () => {
    const { d, texts } = await replaced();
    // Typed later (its own undo step).
    d.view.dispatch(closeHistory(d.view.state.tr.insertText('！', d.view.state.doc.content.size - 1)));
    d.editor.undo();
    expect(texts()).toEqual(['甲 2026', '頁首 2026']);
    d.editor.undo();
    expect(texts()).toEqual(['甲 2025', '頁首 2025']);
    d.done();
  });
});

const A = (id: number, author: string) => `w:id="${id}" w:author="${author}" w:date="2026-09-26T14:32:00Z"`;
const WITH_DEL = `<w:p><w:r><w:t xml:space="preserve">本案 </w:t></w:r><w:del ${A(2, '李大華')}><w:r><w:delText>不同意</w:delText></w:r></w:del><w:r><w:t>辦理。</w:t></w:r></w:p>`;

describe('screen readers, the clipboard and the printout', () => {
  it('the clipboard has no deleted text, not even in data-text or title', async () => {
    const d = await setup(WITH_DEL);
    const clip = listClipboardSerializer(() => d.editor.model.numbering).serializeFragment(d.view.state.doc.content);
    const div = document.createElement('div');
    div.append(clip);
    expect(div.innerHTML).not.toContain('不同意');
    expect(div.innerHTML).not.toContain('李大華');
    expect(div.querySelector('.dx-rev-del')).not.toBeNull(); // still ignored when pasted
    d.done();
  });

  it('the printout has no text for screen readers only (「刪除：」「（刪除結束）」)', async () => {
    const d = await setup(WITH_DEL);
    (d.editor as any).pages = [{ ...(d.editor as any).firstPage(), section: 0, inSection: 0, number: 1, top: -1 }];
    const html = d.editor.printHtml('公文')!;
    const printed = new DOMParser().parseFromString(html, 'text/html');
    expect(printed.body.textContent).toContain('本案 不同意辦理。'); // the deletion as shown, struck through
    expect(printed.body.textContent).not.toContain('刪除：');
    expect(printed.body.textContent).not.toContain('（刪除結束）');
    expect(printed.querySelector('.dx-sr')).toBeNull();
    d.done();
  });

  it('what the headers and footers say is readable once, before the text', async () => {
    const d = await setup(P('內文'));
    d.editor.editHeaderFooter('header');
    const hv = d.editor.activeView!;
    hv.dispatch(hv.state.tr.insertText('臺北市政府 函', 1));
    d.editor.closeHeaderFooter();
    (d.editor as any).renderPages((d.editor as any).pages);
    const summary = d.host.querySelector('.dx-hf-summary') as HTMLElement;
    expect(summary.textContent).toBe('頁首：臺北市政府 函。');
    expect(summary.closest('[aria-hidden="true"]')).toBeNull();
    expect(summary.compareDocumentPosition(d.view.dom) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    d.done();
  });
});
