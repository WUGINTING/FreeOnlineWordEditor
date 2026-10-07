// @vitest-environment jsdom
// 插入 › 空白頁 (asked for by the product owner): two page breaks at the cursor, the cursor left on
// the blank page between them, the text after the cursor on the page after it; one undo step.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { insertBlankPage } from '../../src/papyrus/editor/commands';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function open() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>第一段AB</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return editor;
}

describe('插入空白頁', () => {
  it('puts two page breaks at the cursor and leaves the cursor on the blank page', async () => {
    const editor = await open();
    const view = editor.view!;
    // Cursor between 「A」 and 「B」.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1 + '第一段A'.length)));
    editor.run(insertBlankPage);
    const doc = view.state.doc;
    const kinds: string[] = [];
    doc.forEach((n) => kinds.push(n.type.name === 'paragraph' ? `p:${n.textContent}` : n.type.name));
    expect(kinds).toEqual(['p:第一段A', 'page_break', 'p:', 'page_break', 'p:B']);
    // The cursor is in the empty piece between the two breaks.
    const $c = view.state.selection.$from;
    expect($c.parent.type.name).toBe('paragraph');
    expect($c.parent.textContent).toBe('');
    // Saved as Word writes a page break: two w:br w:type="page", 「A」 before and 「B」 after.
    const xml = await (await JSZip.loadAsync(await editor.save())).file('word/document.xml')!.async('string');
    expect(xml.match(/<w:br w:type="page"\/>/g)).toHaveLength(2);
    expect(xml.indexOf('第一段A')).toBeLessThan(xml.indexOf('w:type="page"'));
    expect(xml.lastIndexOf('w:type="page"')).toBeLessThan(xml.indexOf('>B<'));
    // One undo step takes it all back.
    undo(view.state, view.dispatch);
    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.textContent).toBe('第一段AB');
    editor.destroy();
  });

  it('typing after it puts the text on the blank page, between the breaks', async () => {
    const editor = await open();
    const view = editor.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1 + '第一段AB'.length)));
    editor.run(insertBlankPage);
    view.dispatch(view.state.tr.insertText('新的一頁'));
    const kinds: string[] = [];
    view.state.doc.forEach((n) => kinds.push(n.type.name === 'paragraph' ? `p:${n.textContent}` : n.type.name));
    expect(kinds).toEqual(['p:第一段AB', 'page_break', 'p:新的一頁', 'page_break', 'p:']);
    editor.destroy();
  });
});
