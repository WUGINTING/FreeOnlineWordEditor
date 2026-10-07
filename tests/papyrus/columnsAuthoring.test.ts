// Page columns authoring (GOV-FINDING-014, option D): 版面設定 sets a section's columns (w:cols
// w:num / w:space / w:sep) and 「分欄符號」 inserts a column break (w:br w:type="column"), so Word
// shows and prints the columns exactly. The editor keeps showing one column and says so: a note in
// the dialog, a badge on the section's pages, and hasMultiColumnSections() for the PDF warning.
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, h, nextTick, ref } from 'vue';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import * as sections from '../../src/papyrus/docx/sections';
import { scanCompat } from '../../src/papyrus/docx/compat';
import { computeBreaks, paginationKey, type PageBox } from '../../src/papyrus/editor/pagination';
import { columnGeometry } from '../../src/papyrus/editor/columnLayout';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { insertPageBreak, splitParagraph } from '../../src/papyrus/editor/commands';
import { loadSfc } from './sfc';

const SFC = resolve(__dirname, '../../src/papyrus/vue/PageSetupDialog.vue');
const OUT = resolve(__dirname, `.cols-dialog-${process.pid}.tmp.js`);
/** The compiled component's own imports (./locale) made absolute: it is written next to this file. */
const absoluteImports = (code: string) =>
  code.replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) => `from '${resolve(dirname(SFC), spec).replace(/\\/g, '/')}'`);

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});
afterAll(() => rmSync(OUT, { force: true }));

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (t: string, breakBefore = false, sectPr = '') =>
  `<w:p>${breakBefore || sectPr ? `<w:pPr>${breakBefore ? '<w:pageBreakBefore/>' : ''}${sectPr}</w:pPr>` : ''}${text(t)}</w:p>`;
const COLUMN_BREAK = '<w:r><w:br w:type="column"/></w:r>';

async function pkg(body: string): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

/** 封面 (section 1), 目錄 (section 2, continuous), then a two-page body (last section, w:cols space only). */
const COVER_SECT = `<w:sectPr>${PAGE}<w:titlePg/></w:sectPr>`;
const TOC_SECT = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:pgNumType w:fmt="lowerRoman"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
const LAST_SECT = (cols = '<w:cols w:space="425"/>') => `<w:sectPr>${PAGE}${cols}<w:docGrid w:linePitch="360"/></w:sectPr>`;
const THREE = (lastCols?: string) =>
  pkg(para('封面', false, COVER_SECT) + para('目錄', false, TOC_SECT) + para('正文一') + para('正文二', true) + LAST_SECT(lastCols));

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

function cursorIn(editor: DocxEditor, needle: string, offset = 0) {
  let pos = -1;
  editor.view!.state.doc.descendants((n, at) => {
    if (pos < 0 && n.isText && n.text!.includes(needle)) pos = at + n.text!.indexOf(needle) + offset;
  });
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
}

async function savedXml(editor: DocxEditor): Promise<string> {
  return (await JSZip.loadAsync(await editor.save())).file('word/document.xml')!.async('string');
}
const sectPrs = (xml: string) => xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? [];
const colsOf = (sectPr: string) => /<w:cols\b[^>]*\/>|<w:cols\b[\s\S]*?<\/w:cols>/.exec(sectPr)?.[0] ?? '';
const attrOf = (el: string, name: string) => new RegExp(`\\bw:${name}="([^"]*)"`).exec(el)?.[1] ?? null;

describe('reading w:cols', () => {
  it('each section reports its columns: count, spacing (twips), separator line, equal widths', async () => {
    const sect1 = `<w:sectPr>${PAGE}<w:cols w:num="2" w:sep="1" w:space="567"/></w:sectPr>`;
    const sect2 = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="3" w:space="425" w:equalWidth="0"><w:col w:w="2000" w:space="425"/><w:col w:w="3000" w:space="425"/><w:col w:w="2000"/></w:cols></w:sectPr>`;
    const { doc, model } = await readDocx(await pkg(para('一', false, sect1) + para('二', false, sect2) + para('三') + LAST_SECT('')));
    const s = sections.documentSections(doc, model);
    expect(s.map((x) => x.columns)).toEqual([
      { count: 2, space: 567, separator: true, equalWidth: true },
      { count: 3, space: 425, separator: false, equalWidth: false },
      // No w:cols: one column; Word reads a missing w:space as 720 twips (36 pt, checked with Word).
      { count: 1, space: 720, separator: false, equalWidth: true },
    ]);
  });
});

describe('writing w:cols (withColumns)', () => {
  const withColumns = (sections as any).withColumns as (xml: string | null, c: Partial<{ count: number; space: number; separator: boolean }>) => string;

  it('sets w:num, w:space and w:sep, in Word’s place among the w:sectPr children', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:pgNumType w:start="1"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    const out = withColumns(xml, { count: 2, space: 720, separator: true });
    const cols = colsOf(out);
    expect([attrOf(cols, 'num'), attrOf(cols, 'space'), attrOf(cols, 'sep')]).toEqual(['2', '720', '1']);
    // After w:pgNumType, before w:titlePg and w:docGrid (CT_SectPr order).
    expect(out).toMatch(/<w:pgNumType w:start="1"\/><w:cols [^>]*\/><w:titlePg\/><w:docGrid/);
  });

  it('keeps an existing w:cols’ spacing, and one column drops w:num and w:sep', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:cols w:space="425"/></w:sectPr>`;
    const two = withColumns(xml, { count: 2, separator: true });
    expect([attrOf(colsOf(two), 'num'), attrOf(colsOf(two), 'space'), attrOf(colsOf(two), 'sep')]).toEqual(['2', '425', '1']);
    const one = withColumns(two, { count: 1, separator: false });
    expect(colsOf(one)).toBe('<w:cols w:space="425"/>');
  });

  it('unequal widths become equal widths when the columns are changed', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:cols w:num="2" w:space="425" w:equalWidth="0"><w:col w:w="3000" w:space="425"/><w:col w:w="5000"/></w:cols></w:sectPr>`;
    const out = withColumns(xml, { count: 3 });
    expect(colsOf(out)).not.toContain('w:col ');
    expect(colsOf(out)).not.toContain('equalWidth');
    expect(attrOf(colsOf(out), 'num')).toBe('3');
  });

  it('settings already as asked leave the w:sectPr string exactly as it was', () => {
    const xml = `<w:sectPr ${W}>${PAGE}<w:cols w:num="2" w:sep="1" w:space="425"/></w:sectPr>`;
    expect(withColumns(xml, { count: 2, space: 425, separator: true })).toBe(xml);
    expect(withColumns(xml, {})).toBe(xml);
    const none = `<w:sectPr ${W}>${PAGE}</w:sectPr>`;
    expect(withColumns(none, { count: 1, separator: false })).toBe(none);
  });
});

describe('editing a section’s columns', () => {
  it('writes w:cols into the cursor’s section only; the other sections keep their w:sectPr bytes', async () => {
    const x = await openEditor(await THREE());
    try {
      const before = sectPrs(await savedXml(x.editor));
      cursorIn(x.editor, '正文二');
      (x.editor as any).setColumns({ count: 2, separator: true });
      const after = sectPrs(await savedXml(x.editor));
      expect(after[0]).toBe(before[0]);
      expect(after[1]).toBe(before[1]);
      expect(after[0]).toBe(COVER_SECT);
      expect(after[1]).toBe(TOC_SECT);
      const cols = colsOf(after[2]);
      expect([attrOf(cols, 'num'), attrOf(cols, 'space'), attrOf(cols, 'sep')]).toEqual(['2', '425', '1']);
      expect(after[2]).toContain('<w:docGrid w:linePitch="360"/>');
      // The saved file reads back with the columns.
      const { doc, model } = await readDocx(await x.editor.save());
      expect(sections.documentSections(doc, model).map((s) => s.columns.count)).toEqual([1, 1, 2]);
    } finally {
      x.done();
    }
  });

  it('a paragraph’s section (目錄) gets its own w:cols; 整份文件 sets every section', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '目錄');
      x.editor.setSectionSetup({ columns: { count: 3, space: 567, separator: false } } as any);
      const s = x.editor.sections();
      expect(s.map((v) => v.columns.count)).toEqual([1, 3, 1]);
      expect(s[1].columns.space).toBe(567);
      x.editor.setSectionSetup({ columns: { count: 2 } } as any, 'all');
      expect(x.editor.sections().map((v) => v.columns.count)).toEqual([2, 2, 2]);
    } finally {
      x.done();
    }
  });

  it('is one undo step (with the other page setup changes), and redo brings it back', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '目錄');
      const original = x.editor.view!.state.doc;
      x.editor.setSectionSetup({ numbering: { start: 1 }, columns: { count: 2, separator: true } } as any);
      expect(x.editor.cursorSection().columns).toEqual({ count: 2, space: 720, separator: true, equalWidth: true });
      expect(x.editor.undo()).toBe(true);
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
      expect(x.editor.redo()).toBe(true);
      expect(x.editor.cursorSection().columns.count).toBe(2);
      x.editor.undo();
      // The last (body-level) section too.
      cursorIn(x.editor, '正文一');
      (x.editor as any).setColumns({ count: 3 });
      expect(x.editor.cursorSection().columns.count).toBe(3);
      x.editor.undo();
      expect(x.editor.cursorSection().columns.count).toBe(1);
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
    } finally {
      x.done();
    }
  });

  it('nothing to change: no edit, the document stays unmodified', async () => {
    const x = await openEditor(await THREE('<w:cols w:num="2" w:space="425"/>'));
    try {
      cursorIn(x.editor, '正文一');
      (x.editor as any).setColumns({ count: 2, space: 425, separator: false });
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
      (x.editor as any).setColumns({ count: 2 });
      x.editor.setPageSetup({ ...x.editor.model.page, marginLeft: 1000 });
      expect(x.editor.view!.state.doc.eq(doc)).toBe(true);
      expect(x.notices).toHaveLength(2);
      expect(x.notices[0]).toBe(x.notices[1]);
      expect(x.notices[0]).toContain('追蹤修訂開啟時，無法變更分節符號或版面設定');
    } finally {
      x.done();
    }
  });
});

describe('hasMultiColumnSections (for the web PDF warning)', () => {
  it('is true while some section has more than one column', async () => {
    const x = await openEditor(await THREE());
    try {
      const ed = x.editor as any;
      expect(ed.hasMultiColumnSections()).toBe(false);
      cursorIn(x.editor, '目錄');
      ed.setColumns({ count: 2 });
      expect(ed.hasMultiColumnSections()).toBe(true);
      x.editor.undo();
      expect(ed.hasMultiColumnSections()).toBe(false);
    } finally {
      x.done();
    }
    const y = await openEditor(await THREE('<w:cols w:num="2" w:space="425"/>'));
    try {
      expect((y.editor as any).hasMultiColumnSections()).toBe(true);
    } finally {
      y.done();
    }
  });
});

describe('column breaks (w:br w:type="column")', () => {
  const TWO_COL_BODY = () =>
    pkg(`<w:p>${text('左欄')}${COLUMN_BREAK}${text('右欄')}</w:p>` + LAST_SECT('<w:cols w:num="2" w:space="425"/>'));

  it('round-trips: read as a column break, shown as 分欄符號, saved as the same w:br', async () => {
    const x = await openEditor(await TWO_COL_BODY());
    try {
      const breaks: any[] = [];
      x.editor.view!.state.doc.descendants((n) => {
        if (n.type.name === 'hard_break') breaks.push(n.attrs.type);
      });
      expect(breaks).toEqual(['column']);
      const marker = x.editor.view!.dom.querySelector<HTMLElement>('.dx-col-break');
      expect(marker).not.toBeNull();
      expect(marker!.getAttribute('data-label')).toBe('分欄符號');
      const xml = await savedXml(x.editor);
      expect(xml).toMatch(/左欄<\/w:t><\/w:r><w:r><w:br w:type="column"\/><\/w:r><w:r><w:t[^>]*>右欄/);
    } finally {
      x.done();
    }
  });

  it('「插入分欄符號」 inserts one at the cursor (one undo step) and it saves as w:br w:type="column"', async () => {
    const x = await openEditor(await THREE());
    try {
      cursorIn(x.editor, '正文一', 2);
      const original = x.editor.view!.state.doc;
      expect((x.editor as any).insertColumnBreak()).toBe(true);
      expect(x.editor.view!.dom.querySelectorAll('.dx-col-break')).toHaveLength(1);
      const xml = await savedXml(x.editor);
      expect(xml).toMatch(/正文<\/w:t><\/w:r><w:r><w:br w:type="column"\/><\/w:r><w:r><w:t[^>]*>一/);
      // Read back: still a column break.
      const { doc } = await readDocx(await x.editor.save());
      let found = 0;
      doc.descendants((n) => { if (n.type.name === 'hard_break' && n.attrs.type === 'column') found++; });
      expect(found).toBe(1);
      x.editor.undo();
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
    } finally {
      x.done();
    }
  });

  it('is not inserted in a table cell', async () => {
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>${text('儲存格')}</w:p></w:tc></w:tr></w:tbl>`;
    const x = await openEditor(await pkg(table + para('後') + LAST_SECT()));
    try {
      cursorIn(x.editor, '儲存格', 1);
      const doc = x.editor.view!.state.doc;
      expect((x.editor as any).insertColumnBreak()).toBe(false);
      expect(x.editor.view!.state.doc.eq(doc)).toBe(true);
    } finally {
      x.done();
    }
  });

  it('does not start a new page (a page break does)', async () => {
    const x = await openEditor(await TWO_COL_BODY());
    try {
      expect(layOut(x.editor)).toHaveLength(1);
      cursorIn(x.editor, '右欄');
      x.editor.run(insertPageBreak);
      expect(layOut(x.editor)).toHaveLength(2);
    } finally {
      x.done();
    }
  });
});

describe('on screen: the columns as Word lays them out (persona-300, columnLayout.ts)', () => {
  const COLS = '<w:cols w:num="2" w:space="720" w:sep="1"/>';
  const TWO = () =>
    pkg(
      para('標題', false, `<w:sectPr>${PAGE}</w:sectPr>`) + para('中文') + `<w:p>${COLUMN_BREAK}${text('ENGLISH')}</w:p>` +
        `<w:sectPr><w:type w:val="continuous"/>${PAGE}${COLS}</w:sectPr>`,
    );

  it("column geometry: equal columns share the text width; own widths are the file's", () => {
    expect(columnGeometry('<w:sectPr><w:cols w:num="2" w:space="720"/></w:sectPr>', 600)).toEqual({ lefts: [0, 324], widths: [276, 276], separator: false });
    expect(columnGeometry('<w:sectPr><w:cols w:num="2"/></w:sectPr>', 600)!.lefts[1]).toBe(324);
    const own = columnGeometry(
      '<w:sectPr><w:cols w:num="3" w:equalWidth="0" w:sep="1"><w:col w:w="1500" w:space="300"/><w:col w:w="3000" w:space="300"/><w:col w:w="1500"/></w:cols></w:sectPr>',
      440,
    )!;
    expect(own).toEqual({ lefts: [0, 120, 340], widths: [100, 200, 100], separator: true });
    expect(columnGeometry('<w:sectPr><w:cols w:space="425"/></w:sectPr>', 600)).toBeNull();
    expect(columnGeometry(null, 600)).toBeNull();
  });

  it('a column break sends what follows to the next column; the blocks are a column wide; a line between', async () => {
    const x = await openEditor(await TWO());
    try {
      const view = x.editor.view!;
      const layout = (x.editor as any).layout(view.state.doc);
      const g = layout.sections[1].columns;
      const geo = layout.sections[1].geometry;
      expect(g.lefts.length).toBe(2);
      const result = computeBreaks(view, layout);
      // 「ENGLISH」 (after the column break) is in the second column.
      const english = view.state.doc.child(2);
      expect(english.textContent).toBe('ENGLISH');
      const at = view.state.doc.child(0).nodeSize + view.state.doc.child(1).nodeSize;
      expect(result.placements).toHaveLength(1);
      const [pl] = result.placements!;
      expect([pl.pos, pl.size, pl.top]).toEqual([at, english.nodeSize, 0]);
      expect(pl.dl).toBeCloseTo(g.lefts[1], 1);
      // As wide as the first column: the right padding shrinks by as much as the left one grows.
      expect(pl.dr).toBeCloseTo(-pl.dl, 1);
      expect(result.pages[0].rules).toHaveLength(1);
      view.dispatch(view.state.tr.setMeta(paginationKey, result));
      (x.editor as any).renderPages(result.pages);
      const blocks = Array.from(view.dom.children) as HTMLElement[];
      // The title is not in columns; the section's blocks are a column wide (padding), the English one moved.
      expect(blocks[0].classList.contains('dx-colsec')).toBe(false);
      expect(blocks[1].classList.contains('dx-colsec')).toBe(true);
      expect(parseFloat(blocks[1].style.getPropertyValue('--dx-cpr'))).toBeCloseTo(geo.width - geo.marginLeft - geo.marginRight - g.widths[0], 3);
      expect(parseFloat(blocks[2].style.getPropertyValue('--dx-cdl'))).toBeCloseTo(g.lefts[1], 1);
      expect(document.querySelectorAll('.dx-page .dx-col-rule')).toHaveLength(1);
      // Printed with the line between the columns; no badge any more.
      const html = x.editor.printHtml('t')!;
      expect(html).toContain('dx-col-rule');
      expect(html).not.toContain('本節在 Word 中為');
    } finally {
      x.done();
    }
  });

  it('a document without columns is laid out as before (no column classes, no placements)', async () => {
    const x = await openEditor(await THREE('<w:cols w:space="425"/>'));
    try {
      const view = x.editor.view!;
      const result = computeBreaks(view, (x.editor as any).layout(view.state.doc));
      expect(result.placements).toBeUndefined();
      expect(view.dom.querySelectorAll('.dx-colsec')).toHaveLength(0);
      expect(result.pages.every((p) => !p.rules)).toBe(true);
    } finally {
      x.done();
    }
  });

  it('the compat notice says columns are kept for Word (also those set in the web)', async () => {
    const r = await scanCompat(await JSZip.loadAsync(await THREE('<w:cols w:num="2" w:space="425"/>')));
    const item = r.items.find((i) => i.id === 'columns')!;
    expect(item.effect).toContain('用 Word 開啟或列印時會分欄');
    expect(item.effect).toContain('網頁');
  });
});

describe('版面設定 dialog: 分欄', () => {
  async function compiledDialog(): Promise<any> {
    const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
    const script = compileScript(descriptor, { id: 'cols-dialog', inlineTemplate: true });
    const code = absoluteImports(ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText);
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
  const NOTE = '網頁上也分欄顯示與列印；段落與表格整段移到下一欄（Word 會在段落中間換欄），分欄位置可能和 Word 略有不同。';
  const count = (host: HTMLElement) => host.querySelector<HTMLSelectElement>('select[aria-label="欄數"]')!;
  const spacing = (host: HTMLElement) => host.querySelector<HTMLInputElement>('input[aria-label="欄間距"]')!;
  const box = (host: HTMLElement, label: string) =>
    [...host.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].find((b) => b.parentElement?.textContent?.includes(label))!;
  const choose = async (select: HTMLSelectElement, value: string) => {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
    await nextTick();
  };

  it('shows the section’s columns; 2 欄 + 分隔線 are written as one undo step, with the one-column note', async () => {
    const x = await openEditor(await THREE());
    cursorIn(x.editor, '目錄');
    const d = await mount(x.editor);
    try {
      expect(d.host.textContent).toContain('分欄');
      expect(count(d.host).value).toBe('1');
      expect(Number(spacing(d.host).value)).toBe(1.27);
      expect(box(d.host, '欄寬相等').checked).toBe(true);
      expect(box(d.host, '欄寬相等').disabled).toBe(true);
      expect(d.host.textContent).not.toContain(NOTE);
      const original = x.editor.view!.state.doc;
      await choose(count(d.host), '2');
      box(d.host, '分隔線').click();
      await nextTick();
      expect(d.host.textContent).toContain(NOTE);
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.sections().map((s) => s.columns)).toEqual([
        { count: 1, space: 720, separator: false, equalWidth: true },
        { count: 2, space: 720, separator: true, equalWidth: true },
        { count: 1, space: 425, separator: false, equalWidth: true },
      ]);
      x.editor.undo();
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
    } finally {
      d.done();
      x.done();
    }
  });

  it('a spacing left as shown keeps its exact twips; a new spacing in 公分 is written in twips', async () => {
    const x = await openEditor(await THREE());
    cursorIn(x.editor, '正文一');
    const d = await mount(x.editor);
    try {
      expect(Number(spacing(d.host).value)).toBe(0.75);
      await choose(count(d.host), '3');
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.cursorSection().columns).toMatchObject({ count: 3, space: 425 });
    } finally {
      d.done();
    }
    const d2 = await mount(x.editor);
    try {
      expect(count(d2.host).value).toBe('3');
      const input = spacing(d2.host);
      input.value = '1';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      await nextTick();
      d2.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.cursorSection().columns).toMatchObject({ count: 3, space: 567 });
    } finally {
      d2.done();
      x.done();
    }
  });

  it('refuses spacing that leaves no room for the columns', async () => {
    const x = await openEditor(await THREE());
    cursorIn(x.editor, '正文一');
    const d = await mount(x.editor);
    try {
      await choose(count(d.host), '3');
      const input = spacing(d.host);
      input.value = '9';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      await nextTick();
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(d.host.querySelector('.dx-error')?.textContent).toContain('欄');
      expect(x.editor.cursorSection().columns.count).toBe(1);
    } finally {
      d.done();
      x.done();
    }
  });

  it('while 追蹤修訂 is on, applying columns is refused with the page setup notice', async () => {
    const x = await openEditor(await THREE());
    x.editor.setTrackChanges(true);
    cursorIn(x.editor, '正文一');
    const d = await mount(x.editor);
    try {
      await choose(count(d.host), '2');
      d.host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.cursorSection().columns.count).toBe(1);
      expect(x.notices.at(-1)).toContain('追蹤修訂開啟時，無法變更分節符號或版面設定');
    } finally {
      d.done();
      x.done();
    }
  });
});

// A title over two columns needs a continuous section break (Word: 分節符號（接續本頁）): the
// body's section starts on the title's page. 版面設定 sets where the cursor's section starts.
describe('where a section starts (版面設定: 這一節的開始位置)', () => {
  const withSectionStart = (sections as any).withSectionStart as (xml: string | null, start: string) => string;

  it('writes w:type first in w:sectPr; 新頁 is Word’s default (no w:type); unchanged keeps the string', () => {
    const xml = `<w:sectPr ${W}><w:footerReference w:type="default" r:id="rId9"/>${PAGE}<w:cols w:space="425"/></w:sectPr>`;
    const cont = withSectionStart(xml, 'continuous');
    expect(cont).toMatch(/r:id="rId9"\/><w:type w:val="continuous"\/><w:pgSz/);
    expect(withSectionStart(cont, 'continuous')).toBe(cont);
    expect(withSectionStart(cont, 'nextPage')).not.toContain('w:type w:val');
    expect(withSectionStart(xml, 'nextPage')).toBe(xml);
  });

  it('a section break mark shows how the section after it starts, as Word labels it', async () => {
    const x = await openEditor(await THREE());
    try {
      const label = (needle: string) =>
        Array.from(x.editor.view!.dom.querySelectorAll<HTMLElement>('.dx-sect-end')).find((p) => p.textContent?.includes(needle))?.getAttribute('data-sect-label');
      // 目錄 starts on the cover's page; the body after it on a new page.
      expect(label('封面')).toBe('分節符號（接續本頁）');
      expect(label('目錄')).toBe('分節符號（下一頁）');
      cursorIn(x.editor, '正文一');
      x.editor.setSectionSetup({ start: 'continuous' } as any);
      expect(label('目錄')).toBe('分節符號（接續本頁）');
      x.editor.undo();
      expect(label('目錄')).toBe('分節符號（下一頁）');
    } finally {
      x.done();
    }
  });

  it('the dialog sets the cursor’s section to 接續本頁 (one undo step; other sections keep their bytes)', async () => {
    const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
    const script = compileScript(descriptor, { id: 'cols-dialog-start', inlineTemplate: true });
    writeFileSync(OUT, absoluteImports(ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText));
    const Dialog = (await import(/* @vite-ignore */ `${OUT.replace(/\\/g, '/')}?start`)).default;
    const x = await openEditor(await THREE());
    cursorIn(x.editor, '正文一');
    const before = sectPrs(await savedXml(x.editor));
    const original = x.editor.view!.state.doc;
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp(Dialog, { editor: x.editor, open: true, onClose: () => {} });
    app.mount(host);
    await nextTick();
    try {
      const select = host.querySelector<HTMLSelectElement>('select[aria-label="這一節的開始位置"]')!;
      expect(select.value).toBe('nextPage');
      select.value = 'continuous';
      select.dispatchEvent(new Event('change', { bubbles: true }));
      await nextTick();
      host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      expect(x.editor.sections().map((s) => s.start)).toEqual(['nextPage', 'continuous', 'continuous']);
      const after = sectPrs(await savedXml(x.editor));
      expect(after[0]).toBe(before[0]);
      expect(after[1]).toBe(before[1]);
      expect(after[2]).toBe(before[2].replace('<w:sectPr>', '<w:sectPr><w:type w:val="continuous"/>'));
      x.editor.undo();
      expect(x.editor.view!.state.doc.eq(original)).toBe(true);
    } finally {
      app.unmount();
      host.remove();
      x.done();
    }
  });
});

/** 版面配置 › 分隔設定 (as in Word): opens the menu and returns its items by their name. */
async function breaksMenu(host: HTMLElement) {
  const open = Array.from(host.querySelectorAll<HTMLButtonElement>('button')).find((b) => b.title.startsWith('分隔設定'))!;
  if (open.getAttribute('aria-expanded') !== 'true') open.click();
  await nextTick();
  const items = Array.from(host.querySelectorAll<HTMLButtonElement>('[role="menu"][aria-label="分隔設定"] [role="menuitem"]'));
  return { items, item: (label: string) => items.find((b) => b.querySelector('b')?.textContent === label)! };
}

describe('版面配置 › 分隔設定: 分欄符號 after 分頁符號, before 分節符號 (as in Word)', () => {
  async function toolbar(editor: DocxEditor, snapshot: { value: unknown }) {
    const Toolbar = await loadSfc('src/papyrus/vue/DocxToolbar.vue');
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(Toolbar, { editor, snapshot: snapshot.value, styles: [] }) });
    app.mount(host);
    await nextTick();
    return { host, menu: () => breaksMenu(host), done: () => { app.unmount(); host.remove(); } };
  }

  it('inserts a column break at the cursor; the item is off in tables', async () => {
    const snapshot = ref<unknown>(null);
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s) });
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>${text('儲存格')}</w:p></w:tc></w:tr></w:tbl>`;
    await editor.open(await pkg(para('甲乙') + table + LAST_SECT()));
    const t = await toolbar(editor, snapshot);
    try {
      const { items, item } = await t.menu();
      expect(items.map((b) => b.querySelector('b')?.textContent)).toEqual(['分頁符號', '分欄符號', '分節符號（下一頁）']);
      expect(item('分欄符號').title).toContain('Ctrl+Shift+Enter');
      cursorIn(editor, '甲乙', 1);
      await nextTick();
      (await t.menu()).item('分欄符號').click();
      expect(editor.view!.dom.querySelectorAll('.dx-col-break')).toHaveLength(1);
      cursorIn(editor, '儲存格', 1);
      t.done();
      const again = await toolbar(editor, snapshot);
      expect((await again.menu()).item('分欄符號').disabled).toBe(true);
      again.done();
    } finally {
      editor.destroy();
      host.remove();
    }
  });
});

/**
 * Word check (opt-in): COLS_EXPORT=<folder> writes cols-ui.docx made through the UI — a title in
 * section 1, 分節符號 (版面配置 › 分隔設定), then in 版面設定 the body's section set to 接續本頁 with 2 欄 and 分隔線,
 * and a 分欄符號 (分隔設定) before the English text — for scripts in the scratch folder to open in Word.
 */
describe.skipIf(!process.env.COLS_EXPORT)('document for the Word check', () => {
  it('is made through the toolbar and 版面設定', async () => {
    const snapshot = ref<unknown>(null);
    const host = document.createElement('div');
    document.body.append(host);
    const notices: string[] = [];
    const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s), onNotice: (m) => notices.push(m) });
    await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));
    const type = (t: string) => editor.view!.dispatch(editor.view!.state.tr.insertText(t));
    const Toolbar = await loadSfc('src/papyrus/vue/DocxToolbar.vue');
    const bar = document.createElement('div');
    document.body.append(bar);
    const app = createApp({ render: () => h(Toolbar, { editor, snapshot: snapshot.value, styles: [] }) });
    app.mount(bar);
    await nextTick();
    try {
      type('雙語公告 Bilingual Notice');
      (await breaksMenu(bar)).item('分節符號（下一頁）').click();
      expect(editor.sections()).toHaveLength(2);
      type('本公告以中文及英文並列。中文內容在左欄，英文內容在右欄。');
      editor.run(splitParagraph);
      type('ENGLISH-STARTS-HERE This notice is published in Chinese and English side by side.');
      // 版面設定 for the body's section.
      const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
      const script = compileScript(descriptor, { id: 'cols-word', inlineTemplate: true });
      writeFileSync(OUT, absoluteImports(ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText));
      const Dialog = (await import(/* @vite-ignore */ `${OUT.replace(/\\/g, '/')}?word`)).default;
      const dhost = document.createElement('div');
      document.body.append(dhost);
      const dialog = createApp(Dialog, { editor, open: true, onClose: () => {} });
      dialog.mount(dhost);
      await nextTick();
      const set = async (sel: string, value: string) => {
        const el = dhost.querySelector<HTMLSelectElement>(sel)!;
        el.value = value;
        el.dispatchEvent(new Event('change', { bubbles: true }));
        await nextTick();
      };
      await set('select[aria-label="這一節的開始位置"]', 'continuous');
      await set('select[aria-label="欄數"]', '2');
      [...dhost.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].find((b) => b.parentElement?.textContent?.includes('分隔線'))!.click();
      await nextTick();
      dhost.querySelector<HTMLButtonElement>('.dx-primary')!.click();
      await nextTick();
      dialog.unmount();
      dhost.remove();
      expect(editor.sections().map((s) => [s.start, s.columns.count, s.columns.separator])).toEqual([['nextPage', 1, false], ['continuous', 2, true]]);
      // 分欄符號 at the start of the English paragraph.
      cursorIn(editor, 'ENGLISH-STARTS-HERE');
      await nextTick();
      (await breaksMenu(bar)).item('分欄符號').click();
      expect(notices).toEqual([]);
      const out = resolve(process.env.COLS_EXPORT!, 'cols-ui.docx');
      // (jsdom's Blob has no arrayBuffer(): the saved package is read and written as it is.)
      writeFileSync(out, await (await JSZip.loadAsync(await editor.save())).generateAsync({ type: 'nodebuffer' }));
      const xml = await savedXml(editor);
      writeFileSync(resolve(process.env.COLS_EXPORT!, 'cols-ui.document.xml'), xml);
    } finally {
      app.unmount();
      bar.remove();
      editor.destroy();
      host.remove();
    }
  });
});
