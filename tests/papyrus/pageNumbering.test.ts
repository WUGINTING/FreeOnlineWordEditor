// Section page numbers (GOV-FINDING-016): w:pgNumType's start and format are read, shown like
// Word (PAGE fields in headers/footers, PAGEREF results) and edited per section from 版面設定,
// as one undo step, leaving the other sections' w:sectPr as written.
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, nextTick } from 'vue';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { documentSections, withPageNumbering } from '../../src/papyrus/docx/sections';
import { LETTER_OVERFLOW, fieldNumberFormat, formatPageNumber } from '../../src/papyrus/docx/pageNumbers';
import { computeBreaks, type PageBox } from '../../src/papyrus/editor/pagination';
import { DocxEditor } from '../../src/papyrus/editor/core';

const SFC = resolve(__dirname, '../../src/papyrus/vue/PageSetupDialog.vue');
const OUT = resolve(__dirname, `.pgnum-dialog-${process.pid}.tmp.js`);

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});
afterAll(() => rmSync(OUT, { force: true }));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const fld = (instr: string, cached = '1') =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (t: string, breakBefore = false, sectPr = '') =>
  `<w:p>${breakBefore || sectPr ? `<w:pPr>${breakBefore ? '<w:pageBreakBefore/>' : ''}${sectPr}</w:pPr>` : ''}${text(t)}</w:p>`;

/** Cover section 1 (1 page, 首頁不同, `coverSect` extra) then body sections, each page a paragraph. */
const COVER_SECT = `<w:sectPr><w:footerReference w:type="default" r:id="rIdF1"/>${PAGE}<w:titlePg/></w:sectPr>`;

async function pkg(body: string): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}><w:p>${text('第 ')}${fld('PAGE')}${text(' 頁，共 ')}` +
      `<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple>${text(' 頁；本節 ')}${fld('SECTIONPAGES')}${text(' 頁；')}${fld('PAGE \\* Arabic')}</w:p></w:ftr>`,
  );
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>',
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}

/** A cover page, then a body of three pages whose section has `bodyNumType`. */
const coverAndBody = (bodyNumType: string) =>
  pkg(para('封面', false, COVER_SECT) + para('正文一') + para('正文二', true) + para('正文三', true) + `<w:sectPr>${PAGE}${bodyNumType}</w:sectPr>`);

async function openEditor(bytes: Uint8Array) {
  const host = document.createElement('div');
  document.body.append(host);
  const notices: string[] = [];
  const editor = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
  await editor.open(bytes);
  return { editor, notices, done: () => { editor.destroy(); host.remove(); } };
}

/** Lays the pages out (page breaks and section starts only: jsdom has no text metrics) and draws them. */
function layOut(editor: DocxEditor): PageBox[] {
  const view = editor.view!;
  const { pages } = computeBreaks(view, (editor as any).layout(view.state.doc));
  (editor as any).renderPages(pages);
  return pages;
}

/** The footer text Word would print on each page ('' for none). */
function footers(editor: DocxEditor): string[] {
  return Array.from(document.querySelectorAll<HTMLElement>('.dx-page')).slice(0, editor.pages.length)
    .map((p) => p.querySelector('.dx-footer')?.textContent?.trim() ?? '');
}

function cursorIn(editor: DocxEditor, needle: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((n, at) => {
    if (pos < 0 && n.isText && n.text!.includes(needle)) pos = at;
  });
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
}

async function savedXml(editor: DocxEditor): Promise<string> {
  return (await JSZip.loadAsync(await editor.save())).file('word/document.xml')!.async('string');
}
const sectPrs = (xml: string) => xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];

describe('page number formats (as read from Word: PAGE fields in restarted sections)', () => {
  // [format, number, what Word 2016+ (zh-TW) shows]
  const WORD: [string, number, string][] = [
    ['decimal', 0, '0'], ['decimal', 7, '7'], ['decimal', 32767, '32767'],
    ['upperRoman', 0, ''], ['upperRoman', 4, 'IV'], ['upperRoman', 9, 'IX'], ['upperRoman', 14, 'XIV'], ['upperRoman', 99, 'XCIX'],
    ['upperRoman', 2024, 'MMXXIV'], ['upperRoman', 9999, 'MMMMMMMMMCMXCIX'],
    ['lowerRoman', 1, 'i'], ['lowerRoman', 3, 'iii'], ['lowerRoman', 28, 'xxviii'], ['lowerRoman', 999, 'cmxcix'],
    ['upperLetter', 0, ''], ['upperLetter', 1, 'A'], ['upperLetter', 26, 'Z'], ['upperLetter', 27, 'AA'], ['upperLetter', 28, 'BB'],
    ['upperLetter', 53, 'AAA'], ['upperLetter', 99, 'UUUU'], ['upperLetter', 780, 'Z'.repeat(30)], ['upperLetter', 781, LETTER_OVERFLOW],
    ['lowerLetter', 2, 'b'], ['lowerLetter', 52, 'zz'],
    ['taiwaneseCounting', 0, '○'], ['taiwaneseCounting', 1, '一'], ['taiwaneseCounting', 10, '十'], ['taiwaneseCounting', 11, '十一'],
    ['taiwaneseCounting', 20, '二十'], ['taiwaneseCounting', 21, '二十一'], ['taiwaneseCounting', 99, '九十九'],
    ['taiwaneseCounting', 100, '一○○'], ['taiwaneseCounting', 101, '一○一'], ['taiwaneseCounting', 2024, '二○二四'],
    ['chineseCounting', 12, '十二'], ['chineseCounting', 52, '五十二'], ['chineseCounting', 110, '一一○'],
    ['taiwaneseCountingThousand', 0, '○'], ['taiwaneseCountingThousand', 9, '九'], ['taiwaneseCountingThousand', 10, '一○'],
    ['taiwaneseCountingThousand', 21, '二一'], ['taiwaneseCountingThousand', 1001, '一○○一'],
    ['ideographTraditional', 0, '0'], ['ideographTraditional', 1, '甲'], ['ideographTraditional', 10, '癸'], ['ideographTraditional', 11, '11'],
    ['ideographZodiac', 1, '子'], ['ideographZodiac', 11, '戍'], ['ideographZodiac', 12, '亥'], ['ideographZodiac', 13, '13'],
    ['chineseCountingThousand', 0, '〇'], ['chineseCountingThousand', 10, '十'], ['chineseCountingThousand', 15, '十五'],
    ['chineseCountingThousand', 101, '一百〇一'], ['chineseCountingThousand', 110, '一百一十'], ['chineseCountingThousand', 1010, '一千〇一十'],
    ['chineseCountingThousand', 2024, '二千〇二十四'], ['chineseCountingThousand', 10001, '一万〇一'], ['chineseCountingThousand', 32767, '三万二千七百六十七'],
    ['ideographLegalTraditional', 0, '零'], ['ideographLegalTraditional', 10, '壹拾'], ['ideographLegalTraditional', 101, '壹佰零壹'],
    ['ideographLegalTraditional', 2024, '貳仟零貳拾肆'], ['ideographLegalTraditional', 10001, '壹萬零壹'], ['ideographLegalTraditional', 32767, '參萬貳仟柒佰陸拾柒'],
    ['ideographDigital', 10, '一〇'], ['taiwaneseDigital', 10, '一○'],
    ['decimalZero', 0, '00'], ['decimalZero', 9, '09'], ['decimalZero', 10, '10'],
    ['decimalFullWidth', 105, '１０５'], ['numberInDash', 12, '- 12 -'],
  ];
  it.each(WORD)('%s %i → %s', (fmt, n, shown) => {
    expect(formatPageNumber(n, fmt)).toBe(shown);
  });

  it('no format, "decimal" and formats the editor does not know show 1, 2, 3', () => {
    expect(formatPageNumber(12, null)).toBe('12');
    expect(formatPageNumber(12, 'decimal')).toBe('12');
    expect(formatPageNumber(12, 'hebrew1')).toBe('12');
    expect(formatPageNumber(Number.NaN, 'lowerRoman')).toBe('i');
  });

  it('a field’s \\* switch overrides the section (Word: PAGE \\* Arabic in a lowerRoman section shows 14)', () => {
    expect(fieldNumberFormat('PAGE')).toBeNull();
    expect(fieldNumberFormat(' PAGE  \\* MERGEFORMAT ')).toBeNull();
    expect(fieldNumberFormat('PAGE \\* roman')).toBe('lowerRoman');
    expect(fieldNumberFormat('PAGE \\* ROMAN')).toBe('upperRoman');
    expect(fieldNumberFormat('PAGE \\* alphabetic \\* MERGEFORMAT')).toBe('lowerLetter');
    expect(fieldNumberFormat('PAGE \\* ALPHABETIC')).toBe('upperLetter');
    expect(fieldNumberFormat('PAGE \\* Arabic')).toBe('decimal');
    expect(fieldNumberFormat('PAGE \\* ArabicDash')).toBe('numberInDash');
    expect(fieldNumberFormat('PAGEREF _Toc1 \\h')).toBeNull();
  });
});

describe('reading and writing w:pgNumType', () => {
  it('each section reports its start and format', async () => {
    const { doc, model } = await readDocx(await coverAndBody('<w:pgNumType w:fmt="lowerRoman" w:start="1"/>'));
    const s = documentSections(doc, model);
    expect(s.map((x) => [x.pageNumberStart, x.pageNumberFormat, x.titlePage])).toEqual([[null, null, true], [1, 'lowerRoman', false]]);
  });

  it('only the given values change; the rest of w:pgNumType and the order of w:sectPr stay', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:pgNumType w:fmt="upperRoman" w:chapStyle="1"/><w:cols w:space="425"/><w:titlePg/></w:sectPr>`;
    const started = withPageNumbering(xml, { start: 1 });
    expect(started).toMatch(/<w:pgNumType w:fmt="upperRoman" w:chapStyle="1" w:start="1"\/><w:cols/);
    const decimal = withPageNumbering(started, { format: 'decimal' });
    expect(decimal).toMatch(/<w:pgNumType w:chapStyle="1" w:start="1"\/>/);
    expect(withPageNumbering(decimal, { start: null })).toMatch(/<w:pgNumType w:chapStyle="1"\/>/);
    // New: after w:pgMar, before w:cols and w:titlePg (Word's order).
    const fresh = withPageNumbering(`<w:sectPr ${W}>${PAGE}<w:cols w:space="425"/><w:titlePg/></w:sectPr>`, { format: 'lowerRoman', start: 0 });
    expect(fresh).toMatch(/w:gutter="0"\/><w:pgNumType w:fmt="lowerRoman" w:start="0"\/><w:cols/);
    // Back to continuing in decimal: the empty element goes.
    expect(withPageNumbering(fresh, { format: null, start: null })).not.toContain('pgNumType');
  });

  it('settings already as asked leave the w:sectPr string as it was', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:pgNumType w:fmt="decimal" w:start="3"/></w:sectPr>`;
    expect(withPageNumbering(xml, { format: 'decimal', start: 3 })).toBe(xml);
    expect(withPageNumbering(xml, { format: null })).toBe(xml);
    expect(withPageNumbering(xml, {})).toBe(xml);
  });
});

describe('page numbers shown on the pages', () => {
  it('a cover with 首頁不同 shows no footer, and the body restarted at 1 in lowerRoman shows i, ii, iii', async () => {
    const x = await openEditor(await coverAndBody('<w:pgNumType w:fmt="lowerRoman" w:start="1"/>'));
    try {
      const pages = layOut(x.editor);
      expect(pages.map((p) => [p.section, p.number, p.numberFormat])).toEqual([[0, 1, null], [1, 1, 'lowerRoman'], [1, 2, 'lowerRoman'], [1, 3, 'lowerRoman']]);
      // NUMPAGES counts every page, SECTIONPAGES the section's; both in 1, 2, 3 (Word: fmt is for PAGE only).
      expect(footers(x.editor)).toEqual(['', '第 i 頁，共 4 頁；本節 3 頁；1', '第 ii 頁，共 4 頁；本節 3 頁；2', '第 iii 頁，共 4 頁；本節 3 頁；3']);
    } finally {
      x.done();
    }
  });

  it('a section without w:start continues from the previous one, in its own format', async () => {
    const x = await openEditor(await coverAndBody('<w:pgNumType w:fmt="upperRoman"/>'));
    try {
      expect(layOut(x.editor).map((p) => p.number)).toEqual([1, 2, 3, 4]);
      expect(footers(x.editor).map((f) => f.split('，')[0])).toEqual(['', '第 II 頁', '第 III 頁', '第 IV 頁']);
    } finally {
      x.done();
    }
  });

  it('Chinese formats: 一, 二, 三 in the body restarted at 1', async () => {
    const x = await openEditor(await coverAndBody('<w:pgNumType w:fmt="taiwaneseCounting" w:start="1"/>'));
    try {
      layOut(x.editor);
      expect(footers(x.editor).map((f) => f.split('，')[0])).toEqual(['', '第 一 頁', '第 二 頁', '第 三 頁']);
    } finally {
      x.done();
    }
  });

  it('PAGEREF results (a table of contents) show the target page in its section’s format', async () => {
    const toc = `<w:p>${text('第一章 … ')}${fld('PAGEREF _Toc1 \\h', '9')}</w:p><w:p>${text('（阿拉伯）')}${fld('PAGEREF _Toc1 \\h \\* Arabic', '9')}</w:p>`;
    const heading = '<w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>第一章</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>';
    const x = await openEditor(await pkg(toc + heading + `<w:sectPr>${PAGE}<w:pgNumType w:fmt="lowerRoman"/></w:sectPr>`));
    try {
      // The heading is on the third page of a lowerRoman section (jsdom can't lay out text).
      (x.editor as any).pageAt = () => ({ number: 3, numberFormat: 'lowerRoman' });
      expect(x.editor.pageLabelAt(1)).toBe('iii');
      expect(x.editor.updatePageReferences()).toBe(2);
      const all = () => x.editor.view!.state.doc.textBetween(0, x.editor.view!.state.doc.content.size, ' ');
      const shown = all();
      expect(shown).toContain('第一章 … iii');
      expect(shown).toContain('（阿拉伯）3');
      // One undo step brings the cached numbers back.
      x.editor.undo();
      expect(all()).toContain('第一章 … 9');
    } finally {
      x.done();
    }
  });
});

describe('editing a section’s page numbering', () => {
  const THREE = async () =>
    pkg(
      para('封面', false, COVER_SECT) +
        para('目錄', false, `<w:sectPr>${PAGE}<w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr>`) +
        para('正文一') + para('正文二', true) + `<w:sectPr>${PAGE}<w:pgNumType w:fmt="decimal"/><w:cols w:space="425"/></w:sectPr>`,
    );

  it('writes w:pgNumType into the cursor’s section only; the other sections keep their w:sectPr bytes', async () => {
    const x = await openEditor(await THREE());
    try {
      const before = sectPrs(await savedXml(x.editor));
      cursorIn(x.editor, '正文二');
      x.editor.setPageNumbering({ start: 1 });
      const after = sectPrs(await savedXml(x.editor));
      // As in the file that was opened, and as saved without the edit.
      expect(after[0]).toBe(COVER_SECT);
      expect(after[1]).toBe(`<w:sectPr>${PAGE}<w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr>`);
      expect(after[0]).toBe(before[0]);
      expect(after[1]).toBe(before[1]);
      expect(after[2]).toContain('<w:pgNumType w:fmt="decimal" w:start="1"/>');
      expect(after[2]).toContain('<w:cols w:space="425"/>');
      expect(layOut(x.editor).map((p) => [p.number, p.numberFormat ?? 'decimal'])).toEqual([[1, 'decimal'], [1, 'lowerRoman'], [1, 'decimal'], [2, 'decimal']]);

      // A paragraph's section (目錄): from roman restarted at 1 to continuing in 1, 2, 3.
      cursorIn(x.editor, '目錄');
      x.editor.setPageNumbering({ format: 'decimal', start: null });
      const again = sectPrs(await savedXml(x.editor));
      expect(again[0]).toBe(before[0]);
      expect(again[1]).not.toContain('pgNumType');
      expect(again[2]).toBe(after[2]);
      expect(layOut(x.editor).map((p) => p.number)).toEqual([1, 2, 1, 2]);
    } finally {
      x.done();
    }
  });

  it('the saved file reads back with the new numbering', async () => {
    const x = await openEditor(await coverAndBody(''));
    try {
      cursorIn(x.editor, '正文一');
      x.editor.setPageNumbering({ format: 'lowerRoman', start: 1 });
      const { doc, model } = await readDocx(await (await JSZip.loadAsync(await x.editor.save())).generateAsync({ type: 'uint8array' }));
      expect(documentSections(doc, model).map((s) => [s.pageNumberStart, s.pageNumberFormat])).toEqual([[null, null], [1, 'lowerRoman']]);
    } finally {
      x.done();
    }
  });

  it('is one undo step (format, start and 首頁不同 together), and redo brings it back', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '目錄');
      const original = x.editor.view!.state.doc;
      x.editor.setSectionSetup({ numbering: { format: 'upperRoman', start: 5 }, titlePage: true });
      const s = x.editor.cursorSection();
      expect([s.pageNumberFormat, s.pageNumberStart, s.titlePage]).toEqual(['upperRoman', 5, true]);
      expect(x.editor.undo()).toBe(true);
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
      expect(x.editor.redo()).toBe(true);
      expect(x.editor.cursorSection().pageNumberStart).toBe(5);
      expect(x.editor.undo()).toBe(true);
      // The last (body-level) section too.
      cursorIn(x.editor, '正文一');
      x.editor.setPageNumbering({ start: 1 });
      expect(x.editor.cursorSection().pageNumberStart).toBe(1);
      x.editor.undo();
      expect(x.editor.cursorSection().pageNumberStart).toBeNull();
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
    } finally {
      x.done();
    }
  });

  it('nothing to change: no edit, the document stays unmodified', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '目錄');
      x.editor.setPageNumbering({ format: 'lowerRoman', start: 1 });
      expect(x.editor.isModified()).toBe(false);
    } finally {
      x.done();
    }
  });

  it('is refused while 追蹤修訂 is on, with the same notice as other page setup changes', async () => {
    const x = await openEditor(await THREE());
    try {
      x.editor.setTrackChanges(true);
      cursorIn(x.editor, '正文一');
      const doc = x.editor.view!.state.doc;
      x.editor.setPageNumbering({ start: 1 });
      x.editor.setPageSetup({ ...x.editor.model.page, marginLeft: 1000 });
      expect(x.editor.view!.state.doc.eq(doc)).toBe(true);
      expect(x.notices).toHaveLength(2);
      expect(x.notices[0]).toBe(x.notices[1]);
      expect(x.notices[0]).toContain('追蹤修訂開啟時，無法變更分節符號或版面設定');
    } finally {
      x.done();
    }
  });

  it('整份文件 applies the format to every section, keeping each one’s start', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '目錄');
      x.editor.setPageNumbering({ format: 'upperLetter' }, 'all');
      expect(x.editor.sections().map((s) => [s.pageNumberFormat, s.pageNumberStart])).toEqual([['upperLetter', null], ['upperLetter', 1], ['upperLetter', null]]);
    } finally {
      x.done();
    }
  });
});

describe('版面設定 dialog: 頁碼', () => {
  async function compiledDialog(): Promise<any> {
    const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
    const script = compileScript(descriptor, { id: 'pgnum-dialog', inlineTemplate: true });
    const code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText
      // The component's own imports (./locale) are found from this file's folder.
      .replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) => `from '${resolve(dirname(SFC), spec).replace(/\\/g, '/')}'`);
    writeFileSync(OUT, code);
    return (await import(/* @vite-ignore */ OUT.replace(/\\/g, '/'))).default;
  }
  async function mount(editor: DocxEditor) {
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp(await compiledDialog(), { editor, open: true, onClose: () => {} });
    app.mount(host);
    await nextTick();
    return { host, done: () => { app.unmount(); host.remove(); } };
  }
  const formatSelect = (host: HTMLElement) => host.querySelector<HTMLSelectElement>('select[aria-label="頁碼格式"]')!;
  const radios = (host: HTMLElement) => host.querySelectorAll<HTMLInputElement>('input[name="dx-ps-restart"]');

  it('shows the cursor section’s numbering and sets the body to restart at 1 in lowerRoman', async () => {
    const x = await openEditor(await coverAndBody(''));
    cursorIn(x.editor, '正文二');
    layOut(x.editor);
    const d = await mount(x.editor);
    try {
      expect(d.host.textContent).toContain('頁碼格式');
      expect(d.host.textContent).toContain('封面不顯示頁碼、正文從 1 起算');
      expect(formatSelect(d.host).value).toBe('decimal');
      expect(radios(d.host)[0].checked).toBe(true); // 接續前一節
      expect(d.host.textContent).toContain('這一節第一頁的頁碼：2');
      const select = formatSelect(d.host);
      select.value = 'lowerRoman';
      select.dispatchEvent(new Event('change', { bubbles: true }));
      radios(d.host)[1].click();
      const start = d.host.querySelector<HTMLInputElement>('input[aria-label="起始頁碼"]')!;
      start.value = '1';
      start.dispatchEvent(new Event('input', { bubbles: true }));
      await nextTick();
      expect(d.host.textContent).toContain('這一節第一頁的頁碼：i');
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      const s = x.editor.sections();
      expect(s.map((v) => [v.pageNumberFormat, v.pageNumberStart])).toEqual([[null, null], ['lowerRoman', 1]]);
      // One undo step for the dialog.
      x.editor.undo();
      expect(x.editor.sections()[1].pageNumberStart).toBeNull();
    } finally {
      d.done();
      x.done();
    }
  });

  it('首頁不同 is in the dialog for the cursor’s section; only what changed is written', async () => {
    const x = await openEditor(await coverAndBody('<w:pgNumType w:fmt="lowerRoman" w:start="1"/>'));
    cursorIn(x.editor, '封面');
    const before = sectPrs(await savedXml(x.editor));
    const d = await mount(x.editor);
    try {
      const box = [...d.host.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].find((b) => b.parentElement?.textContent?.includes('首頁不同'))!;
      expect(box.checked).toBe(true);
      box.click();
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.sections().map((s) => s.titlePage)).toEqual([false, false]);
      const after = sectPrs(await savedXml(x.editor));
      expect(after[0]).not.toContain('titlePg');
      expect(after[1]).toBe(before[1]);
    } finally {
      d.done();
      x.done();
    }
  });

  it('a start number that is not a whole number from 0 is refused in the dialog', async () => {
    const x = await openEditor(await coverAndBody(''));
    cursorIn(x.editor, '正文一');
    const d = await mount(x.editor);
    try {
      radios(d.host)[1].click();
      const start = d.host.querySelector<HTMLInputElement>('input[aria-label="起始頁碼"]')!;
      start.value = '-2';
      start.dispatchEvent(new Event('input', { bubbles: true }));
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(d.host.querySelector('[role="alert"]')?.textContent).toContain('起始頁碼');
      expect(x.editor.isModified()).toBe(false);
    } finally {
      d.done();
      x.done();
    }
  });
});
