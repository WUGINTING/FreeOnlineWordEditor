// Review of persona-300 text columns: a block that does not fit in its column at the column's top.
// On a page whose columns start below other text it goes to the next page; a block taller than a
// whole column (it could only be cut off) makes its section laid out in one column, which the
// host says before making a PDF (columnsShownAsOne), until no block of the section could be.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { computeBreaks } from '../../src/papyrus/editor/pagination';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const para = (t: string, sectPr = '') => `<w:p>${sectPr ? `<w:pPr>${sectPr}</w:pPr>` : ''}<w:r><w:t>${t}</w:t></w:r></w:p>`;
const COLS = `<w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`;

/** How tall each block is on the fake page (px, by its text); others are 20 px. */
let heights: Record<string, number> = {};
let restore: () => void = () => {};

beforeEach(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const rect = (top: number, height: number) => ({ left: 0, right: 600, width: 600, top, bottom: top + height, height, x: 0, y: top, toJSON() {} }) as DOMRect;
  const spy = vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    const root = this.closest('.ProseMirror');
    const h = (el: Element) => heights[el.textContent ?? ''] ?? 20;
    if (this === root) return rect(0, Array.from(root.children).reduce((s, c) => s + h(c), 0));
    if (root && this.parentElement === root) {
      // The blocks start below the body's top padding (the first page's top margin).
      let y = parseFloat(getComputedStyle(root).paddingTop) || 0;
      for (const c of Array.from(root.children)) {
        if (c === this) break;
        y += h(c);
      }
      return rect(y, h(this));
    }
    return rect(0, 0);
  });
  restore = () => spy.mockRestore();
});

let editor: DocxEditor | null = null;
afterEach(() => {
  restore();
  editor?.destroy();
  editor = null;
  heights = {};
  document.body.innerHTML = '';
});

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr>${PAGE}</w:sectPr></w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  editor = new DocxEditor(host, {});
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return editor;
}

const layoutOf = (ed: DocxEditor) => (ed as any).layout(ed.view!.state.doc);

describe('a block that does not fit at the top of its column', () => {
  it('goes to the next page when the columns start below other text on the page', async () => {
    // A 600 px title, then the section in columns on the same page (A4: a 931 px text area).
    heights = { 標題: 600, 甲: 500 };
    const ed = await open(para('標題', `<w:sectPr>${PAGE}</w:sectPr>`) + para('甲') + para('乙', COLS) + para('丙'));
    const result = computeBreaks(ed.view!, layoutOf(ed));
    expect(result.singleColumn).toBeUndefined();
    // 標題; 甲 and 乙 in columns; 丙 (the last section starts a new page).
    expect(result.pages).toHaveLength(3);
    // A spacer before 甲 takes it to the top of page 2.
    const at = ed.view!.state.doc.child(0).nodeSize;
    const spacer = result.breaks.find((b) => b.pos === at)!;
    expect(spacer).toBeTruthy();
    // (甲 starts where the title ends: the page's text top + 600.)
    const end = result.pages[0].top + result.pages[0].textTop + 600;
    expect(Math.abs(end + spacer.height - (result.pages[1].top + result.pages[1].textTop))).toBeLessThanOrEqual(1);
  });

  it('a block taller than a whole column: the section is laid out in one column, and the host is told', async () => {
    heights = { 很高的表格: 1500 };
    const ed = await open(para('標題', `<w:sectPr>${PAGE}</w:sectPr>`) + para('甲') + para('很高的表格') + para('乙', COLS) + para('丙'));
    const layout = layoutOf(ed);
    expect(layout.sections[1].columns).toBeTruthy();
    const result = computeBreaks(ed.view!, layout);
    expect(result.singleColumn).toEqual([1]);
    expect(ed.columnsShownAsOne()).toBe(false);
    (ed as any).setSingleColumn(result.singleColumn);
    expect(ed.columnsShownAsOne()).toBe(true);
    const single = layoutOf(ed);
    expect(single.sections[1].columns).toBeNull();
    expect(single.sections[1].columnsOff).toEqual(layout.sections[1].columns);
    // Laid out in one column now: no blocks moved into a column, the section stays so ...
    const again = computeBreaks(ed.view!, single);
    expect(again.placements).toBeUndefined();
    expect(again.singleColumn).toEqual([1]);
    // ... until no block of it could be taller than a column (here: the table made short).
    heights = { 很高的表格: 200 };
    const back = computeBreaks(ed.view!, single);
    expect(back.singleColumn).toBeUndefined();
    (ed as any).setSingleColumn([]);
    expect(ed.columnsShownAsOne()).toBe(false);
    expect(layoutOf(ed).sections[1].columns).toBeTruthy();
  });

  it('a new document starts in columns again', async () => {
    heights = { 很高的表格: 1500 };
    const ed = await open(para('很高的表格') + para('乙', COLS));
    (ed as any).setSingleColumn(computeBreaks(ed.view!, layoutOf(ed)).singleColumn);
    expect(ed.columnsShownAsOne()).toBe(true);
    await ed.open(await (async () => {
      const zip = blankPackage();
      zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${para('很高的表格') + para('乙', COLS)}<w:sectPr>${PAGE}</w:sectPr></w:body></w:document>`);
      return zip.generateAsync({ type: 'uint8array' });
    })());
    expect(ed.columnsShownAsOne()).toBe(false);
  });
});
