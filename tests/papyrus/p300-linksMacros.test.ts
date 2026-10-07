// persona-300: a link to another site opens only once the user has seen its address and agreed;
// a link to a place in the document goes there. A document with macros says so (hasMacros).
import { afterEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { domStubs, docxOf } from './p300Helpers';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { externalLinkQuestion, followLink, MISSING_ANCHOR, relativeLinkNotice } from '../../src/papyrus/editor/linkFollow';
import { hasMacroParts, MACRO_NOTICE } from '../../src/papyrus/docx/macros';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';

domStubs();

const REL = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const BODY =
  `<w:p><w:hyperlink r:id="rIdX1" ${REL}><w:r><w:t>市府網站</w:t></w:r></w:hyperlink></w:p>` +
  '<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>第一章</w:t></w:r></w:hyperlink></w:p>' +
  '<w:p><w:hyperlink w:anchor="nowhere"><w:r><w:t>斷掉的連結</w:t></w:r></w:hyperlink></w:p>' +
  '<w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>第一章 總則</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>';

async function linkDoc(): Promise<Uint8Array> {
  return docxOf(BODY, (zip) => {
    zip.file(
      'word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdX1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.gov.taipei/" TargetMode="External"/>' +
        '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>',
    );
  });
}

let editor: DocxEditor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
  vi.restoreAllMocks();
});

async function open(editable: boolean, answer: boolean) {
  const host = document.createElement('div');
  document.body.append(host);
  const asked: string[] = [];
  const notices: string[] = [];
  editor = new DocxEditor(host, { editable, confirm: (m) => (asked.push(m), answer), onNotice: (m) => notices.push(m) });
  await editor.open(await linkDoc());
  const opened: unknown[][] = [];
  vi.spyOn(window, 'open').mockImplementation((...args: unknown[]) => (opened.push(args), null));
  const link = (text: string) => [...host.querySelectorAll('a')].find((a) => a.textContent === text)!;
  const click = (text: string, init: MouseEventInit = {}) => {
    const e = new MouseEvent('click', { bubbles: true, cancelable: true, button: 0, ...init });
    link(text).dispatchEvent(e);
    return e;
  };
  return { host, asked, notices, opened, click };
}

describe('clicking a link', () => {
  it('while viewing only: asks with the address, then opens it in a new tab', async () => {
    const d = await open(false, true);
    const e = d.click('市府網站');
    expect(e.defaultPrevented).toBe(true);
    expect(d.asked).toEqual(['即將開啟外部網站「www.gov.taipei」：\nhttps://www.gov.taipei/\n確定要開啟嗎？']);
    expect(d.opened).toEqual([['https://www.gov.taipei/', '_blank', 'noopener,noreferrer']]);
  });

  it('answering no opens nothing', async () => {
    const d = await open(false, false);
    expect(d.click('市府網站').defaultPrevented).toBe(true);
    expect(d.asked).toHaveLength(1);
    expect(d.opened).toEqual([]);
  });

  it('a link to a place in the document goes there, without asking', async () => {
    const d = await open(false, true);
    d.click('第一章');
    expect(d.asked).toEqual([]);
    expect(d.opened).toEqual([]);
    const sel = editor!.view!.state.selection;
    expect(editor!.view!.state.doc.resolve(sel.from).parent.textContent).toBe('第一章 總則');
    d.click('斷掉的連結');
    expect(d.notices).toContain(MISSING_ANCHOR);
  });

  it('while editing: a plain click only places the cursor; Ctrl+click follows (as in Word)', async () => {
    const d = await open(true, true);
    expect(d.click('市府網站').defaultPrevented).toBe(true);
    expect(d.asked).toEqual([]);
    d.click('市府網站', { ctrlKey: true });
    expect(d.asked).toHaveLength(1);
    expect(d.opened).toHaveLength(1);
  });

  it('followLink refuses targets the page never opens, and names the address it asks about', () => {
    const confirm = vi.fn(() => true);
    const open = vi.fn();
    expect(followLink('javascript:alert(1)', { confirm, open, goToBookmark: () => true })).toBe(false);
    expect(followLink('file:///C:/x.exe', { confirm, open, goToBookmark: () => true })).toBe(false);
    expect(confirm).not.toHaveBeenCalled();
    expect(followLink('mailto:a@b.gov.tw', { confirm, open, goToBookmark: () => true })).toBe(true);
    expect(confirm).toHaveBeenCalledWith('即將開啟郵件程式，寄信給：a@b.gov.tw，確定要開啟嗎？');
    expect(confirm).toHaveBeenCalledWith(externalLinkQuestion('mailto:a@b.gov.tw'));
    expect(open).toHaveBeenCalledWith('mailto:a@b.gov.tw');
  });

  it('a relative target (a file beside the document) is not opened against this site; it says so', () => {
    const confirm = vi.fn(() => true);
    const open = vi.fn();
    const notices: string[] = [];
    expect(followLink('附件/報告.docx', { confirm, open, goToBookmark: () => true, onNotice: (m) => notices.push(m) })).toBe(false);
    expect(followLink('report.docx', { confirm, open, goToBookmark: () => true, onNotice: (m) => notices.push(m) })).toBe(false);
    expect(confirm).not.toHaveBeenCalled();
    expect(open).not.toHaveBeenCalled();
    expect(notices).toEqual([relativeLinkNotice('附件/報告.docx'), relativeLinkNotice('report.docx')]);
    expect(notices[0]).toContain('網頁上無法開啟');
  });

  it('the question names the host as the browser contacts it (punycode) and the whole address', () => {
    const confirm = vi.fn(() => true);
    const open = vi.fn();
    followLink('https://www.gov.taipei@evil.example/login', { confirm, open, goToBookmark: () => true });
    expect(confirm).toHaveBeenLastCalledWith('即將開啟外部網站「evil.example」：\nhttps://www.gov.taipei@evil.example/login\n確定要開啟嗎？');
    followLink('https://臺北.tw/公告', { confirm, open, goToBookmark: () => true });
    expect(confirm).toHaveBeenLastCalledWith(expect.stringContaining('「xn--'));
    expect(open).toHaveBeenLastCalledWith(expect.stringMatching(/^https:\/\/xn--[^/]+\.tw\/%E5%85%AC%E5%91%8A$/));
  });

  it('a middle click is asked about too, and the browser never opens it by itself', async () => {
    const d = await open(true, false);
    const link = [...d.host.querySelectorAll('a')].find((a) => a.textContent === '市府網站')!;
    const e = new MouseEvent('auxclick', { bubbles: true, cancelable: true, button: 1 });
    link.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(d.asked).toHaveLength(1);
    expect(d.opened).toEqual([]);
    // The right button is the context menu's.
    const r = new MouseEvent('auxclick', { bubbles: true, cancelable: true, button: 2 });
    link.dispatchEvent(r);
    expect(d.asked).toHaveLength(1);
  });
});

describe('macros', () => {
  it('hasMacros is true for a package with a VBA project (and it is saved untouched)', async () => {
    const bytes = await docxOf('<w:p><w:r><w:t>巨集</w:t></w:r></w:p>', (zip) => {
      zip.file('word/vbaProject.bin', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]));
      zip.file('word/vbaData.xml', '<?xml version="1.0"?><wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>');
    });
    const host = document.createElement('div');
    editor = new DocxEditor(host);
    await editor.open(bytes);
    expect(editor.hasMacros()).toBe(true);
    const { doc, model } = await readDocx(bytes);
    const saved = await JSZip.loadAsync(await writeDocx(doc, model));
    expect(Array.from(await saved.file('word/vbaProject.bin')!.async('uint8array'))).toEqual([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]);
    expect(MACRO_NOTICE).toBe('這份文件含有巨集（Word 的自動程式），線上文件不會執行，下載後在 Word 開啟請小心');
  });

  it('hasMacros is true for ActiveX controls and for a document attached to a .dotm template', async () => {
    const activeX = await docxOf('<w:p/>', (zip) => zip.file('word/activeX/activeX1.bin', new Uint8Array([1, 2])));
    editor = new DocxEditor(document.createElement('div'));
    await editor.open(activeX);
    expect(editor.hasMacros()).toBe(true);
    editor.destroy();

    const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    const withTemplate = (target: string) =>
      docxOf('<w:p/>', (zip) => {
        zip.file(
          'word/_rels/settings.xml.rels',
          RELS + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="${target}" TargetMode="External"/></Relationships>`,
        );
      });
    const { hasMacroTemplate } = await import('../../src/papyrus/docx/macros');
    const dotm = await JSZip.loadAsync(await withTemplate('file:///C:\\Users\\a\\範本\\公文.dotm'));
    const settingsRel = /relationships\/settings/.test((await dotm.file('word/_rels/document.xml.rels')?.async('string')) ?? '');
    // (The blank package points at its settings part; the template relationship hangs off it.)
    expect(settingsRel).toBe(true);
    expect(await hasMacroTemplate(dotm)).toBe(true);
    expect(await hasMacroTemplate(await JSZip.loadAsync(await withTemplate('file:///C:\\範本\\Normal.dotx')))).toBe(false);
    editor = new DocxEditor(document.createElement('div'));
    await editor.open(await withTemplate('公文.dotm'));
    expect(editor.hasMacros()).toBe(true);
  });

  it('hasMacros is false for an ordinary document', async () => {
    const host = document.createElement('div');
    editor = new DocxEditor(host);
    await editor.open(await docxOf('<w:p/>'));
    expect(editor.hasMacros()).toBe(false);
    expect(hasMacroParts(null)).toBe(false);
  });
});
