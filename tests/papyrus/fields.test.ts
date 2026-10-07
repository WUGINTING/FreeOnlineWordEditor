// Fields to fill in: a template's content controls (editor/fields.ts).
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PH = '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>';
const inlineField = (id: number, alias: string, tag: string, text: string, placeholder = true, kind = '<w:text/>') =>
  `<w:sdt><w:sdtPr>${placeholder ? PH : ''}<w:alias w:val="${alias}"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/>${placeholder ? '<w:showingPlcHdr/>' : ''}${kind}</w:sdtPr>` +
  `<w:sdtContent><w:r>${placeholder ? PH : ''}<w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;
const BODY =
  `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:t>目錄</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
  `<w:p><w:r><w:t>客戶：</w:t></w:r>${inlineField(1, '客戶名稱*', 'customer', '按一下輸入客戶名稱')}</w:p>` +
  `<w:p><w:r><w:t>日期：</w:t></w:r>${inlineField(2, '簽約日期', 'date', '選擇日期', true, '<w:date/>')}</w:p>` +
  `<w:p><w:r><w:t>備註：</w:t></w:r>${inlineField(3, '備註', 'note', '無', false)}</w:p>` +
  `<w:sdt><w:sdtPr><w:alias w:val="條款"/><w:tag w:val="required-terms"/><w:id w:val="4"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>第一條 服務內容。</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

async function open(body = BODY) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host);
  await ed.open(await zip.generateAsync({ type: 'uint8array' }));
  const saved = async () => (await JSZip.loadAsync(await writeDocx(ed.view!.state.doc, ed.model))).file('word/document.xml')!.async('string');
  return { ed, saved, done: () => (ed.destroy(), host.remove()) };
}

describe('fields to fill in', () => {
  it('lists the content controls with title, kind, required and filled; not the table of contents', async () => {
    const { ed, done } = await open();
    const fields = ed.fields();
    expect(fields.map((f) => [f.title, f.kind, f.required, f.placeholder, f.filled])).toEqual([
      ['客戶名稱*', 'text', true, true, false],
      ['簽約日期', 'date', false, true, false],
      ['備註', 'text', false, false, true],
      ['條款', 'richText', true, false, true],
    ]);
    expect(ed.unfilledRequired().map((f) => f.title)).toEqual(['客戶名稱*']);
    done();
  });

  it('typing over a placeholder makes it a real value: no w:showingPlcHdr, no placeholder style; undo brings it back', async () => {
    const { ed, saved, done } = await open();
    ed.goToField(ed.fields()[0].id);
    expect(ed.view!.state.doc.textBetween(ed.view!.state.selection.from, ed.view!.state.selection.to)).toBe('按一下輸入客戶名稱');
    ed.run((s, d) => (d?.(s.tr.insertText('台灣測試股份有限公司')), true));
    const f = ed.fields()[0];
    expect([f.text, f.placeholder, f.filled]).toEqual(['台灣測試股份有限公司', false, true]);
    expect(ed.unfilledRequired()).toEqual([]);
    const xml = await saved();
    const sdt = xml.match(/<w:sdt>(?:(?!<\/w:sdt>)[\s\S])*?customer[\s\S]*?<\/w:sdt>/)![0];
    expect(sdt).not.toContain('showingPlcHdr');
    expect(sdt.match(/<w:sdtContent>[\s\S]*<\/w:sdtContent>/)![0]).not.toContain('PlaceholderText');
    expect(sdt).toContain('台灣測試股份有限公司');
    // The other placeholder is untouched.
    expect(xml.match(/showingPlcHdr/g)).toHaveLength(1);
    ed.undo();
    expect(ed.fields()[0].placeholder).toBe(true);
    expect(ed.fields()[0].text).toBe('按一下輸入客戶名稱');
    done();
  });

  it('next unfilled field goes to the next empty one after the cursor, wrapping around', async () => {
    const { ed, done } = await open();
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.atStart(ed.view!.state.doc)));
    expect(ed.nextUnfilledField()?.title).toBe('客戶名稱*');
    expect(ed.nextUnfilledField()?.title).toBe('簽約日期');
    expect(ed.nextUnfilledField()?.title).toBe('客戶名稱*');
    expect(ed.nextUnfilledField(true)?.title).toBe('客戶名稱*');
    done();
  });

  it('a template saved without edits keeps its content controls exactly', async () => {
    const { saved, done } = await open();
    const xml = await saved();
    expect(xml.match(/showingPlcHdr/g)).toHaveLength(2);
    expect(xml).toContain('<w:alias w:val="客戶名稱*"/>');
    expect(xml).toContain('<w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>按一下輸入客戶名稱</w:t>');
    done();
  });
});

describe('locked fields (Word content control locks)', () => {
  const LOCKED =
    `<w:p><w:r><w:t>條款：</w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="固定條款"/><w:id w:val="10"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr>` +
    `<w:sdtContent><w:r><w:t>本契約不得轉讓。</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>（完）</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>名稱：</w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="名稱"/><w:id w:val="11"/><w:lock w:val="sdtLocked"/><w:text/></w:sdtPr>` +
    `<w:sdtContent><w:r><w:t>甲公司</w:t></w:r></w:sdtContent></w:sdt></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

  async function openLocked() {
    const notices: string[] = [];
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${LOCKED}</w:body></w:document>`);
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    const text = () => ed.view!.state.doc.textContent;
    return { ed, notices, text, done: () => (ed.destroy(), host.remove()) };
  }

  it('content-locked: no typing, deleting or formatting inside, not even at its edges; with a message', async () => {
    const { ed, notices, text, done } = await openLocked();
    const f = ed.fields()[0];
    const before = text();
    ed.run((s, d) => (d?.(s.tr.insertText('X', f.from + 2)), true));
    ed.run((s, d) => (d?.(s.tr.delete(f.from + 1, f.from + 3)), true));
    ed.run((s, d) => (d?.(s.tr.addMark(f.from, f.to, s.schema.marks.bold.create())), true));
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, f.to)));
    ed.run((s, d) => (d?.(s.tr.insertText('尾')), true)); // inherits the field's mark
    expect(text()).toBe(before);
    expect(notices[0]).toBe('「固定條款」已鎖定，內容不能修改。');
    // Outside the field, editing works.
    ed.run((s, d) => (d?.(s.tr.insertText('第一條', 1)), true));
    expect(text()).toContain('第一條條款：');
    done();
  });

  it('delete-locked: its content may change, but the field may not be removed', async () => {
    const { ed, notices, text, done } = await openLocked();
    const f = ed.fields()[1];
    ed.run((s, d) => (d?.(s.tr.insertText('乙', f.from, f.from + 1)), true));
    expect(ed.fields()[1].text).toBe('乙公司');
    const g = ed.fields()[1];
    ed.run((s, d) => (d?.(s.tr.delete(g.from - 1, g.to)), true));
    expect(text()).toContain('乙公司');
    expect(notices.pop()).toBe('「名稱」已鎖定，欄位不能刪除。');
    done();
  });
});
