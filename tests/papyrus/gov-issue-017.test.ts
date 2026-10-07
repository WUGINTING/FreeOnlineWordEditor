// GOV-ISSUE-017: a table row that breaks across several pages must show every line of its cells
// exactly once, on screen and in the printout: no line under a row-break cover, under the header
// rows repeated above the continuation, or in the footer.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { splitRow, rowGrowth, type RowSplitCell } from '../../src/papyrus/editor/pagination';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

// ----- the cut computation, given measured line boxes -----

/** A landscape A4 page at 96 dpi, as GOVTEST_021 uses (px). */
const PAGE = { height: 793.93, textTop: 67, textBottom: 67, gap: 24 };
const pageTop = (p: number) => p * (PAGE.height + PAGE.gap);
const area = (p: number) => ({ top: pageTop(p) + PAGE.textTop, bottom: pageTop(p) + PAGE.height - PAGE.textBottom });

/**
 * A cell of `paras` paragraphs of `perPara` lines each (19.37 px lines, like 11 pt text at single
 * spacing; widow/orphan control on), from the top of its row.
 */
function cell(paras: number, perPara: number, pitch = 19.3747): RowSplitCell {
  const lines = [];
  for (let p = 0; p < paras; p++) {
    for (let i = 0; i < perPara; i++) {
      const n = p * perPara + i;
      lines.push({ top: n * pitch, bottom: (n + 1) * pitch, index: i, count: perPara, rule: { can: true, min: 2 } });
    }
  }
  return { lines, padTop: 0, padBottom: 0, bottom: lines.length * pitch };
}

/**
 * Splits the row as computeBreaks does and returns where every line of every cell ends up
 * (page, top, bottom on the canvas), with each page's cover top and the header band below it.
 */
function layOut(cells: RowSplitCell[], rowY: number, repeatHeight: number) {
  let page = Math.floor(rowY / (PAGE.height + PAGE.gap));
  const first = page;
  const cutTop = new Map<number, number>();
  const spacers = cells.map(() => new Map<number, number>());
  const extra = splitRow(cells, rowY, repeatHeight, {
    area: () => area(page),
    breakPage: (partBottom) => {
      cutTop.set(page, Math.min(partBottom, area(page).bottom));
      page++;
    },
    place: (ci, k, height) => {
      spacers[ci].set(k, height);
      return height;
    },
  });
  const lines = cells.map((c, ci) => {
    let shift = 0;
    return c.lines.map((l, k) => {
      shift += spacers[ci].get(k) ?? 0;
      const top = rowY + l.top + shift;
      const bottom = rowY + l.bottom + shift;
      const onPage = Math.floor(top / (PAGE.height + PAGE.gap));
      return { top, bottom, page: onPage };
    });
  });
  return { first, last: page, cutTop, lines, extra };
}

/** Every line on exactly one page, inside its text area, below the repeated header, above the cover. */
function expectAllVisible(r: ReturnType<typeof layOut>, repeatHeight: number) {
  for (const cellLines of r.lines) {
    for (const l of cellLines) {
      const a = area(l.page);
      const top = l.page > r.first ? a.top + repeatHeight : a.top;
      const bottom = Math.min(a.bottom, r.cutTop.get(l.page) ?? Infinity);
      expect(l.top, `line at ${l.top.toFixed(1)} on page ${l.page}`).toBeGreaterThanOrEqual(top - 0.5);
      expect(l.bottom, `line ending at ${l.bottom.toFixed(1)} on page ${l.page}`).toBeLessThanOrEqual(bottom + 0.5);
    }
    // In order: a line never lands above the one before it.
    for (let k = 1; k < cellLines.length; k++) expect(cellLines[k].top).toBeGreaterThanOrEqual(cellLines[k - 1].bottom - 0.01);
  }
}

describe('a row breaking across N pages: every line lands on exactly one page (GOV-ISSUE-017)', () => {
  const rowY = area(0).top + 300; // the row starts part-way down the first page
  const header = 36.6;
  for (const [pages, paras] of [[2, 8], [3, 14], [4, 24], [6, 40]] as const) {
    it(`${pages} pages`, () => {
      const cells = [cell(1, 2), cell(paras, 4), cell(1, 3)];
      const r = layOut(cells, rowY, header);
      expect(r.last - r.first + 1).toBe(pages);
      expectAllVisible(r, header);
      // Each page of the row shows some of the tall cell.
      const onPages = new Set(r.lines[1].map((l) => l.page));
      expect(onPages.size).toBe(pages);
    });
  }

  it('the GOVTEST_021 row: 45 paragraphs of 4 lines, every paragraph visible', () => {
    const cells = [cell(1, 2), cell(1, 2), cell(1, 2), cell(1, 3), cell(1, 3), cell(45, 4), cell(1, 2), cell(1, 2)];
    const r = layOut(cells, area(0).top + 294, 20.4);
    expectAllVisible(r, 20.4);
    const paras = new Set<number>();
    r.lines[5].forEach((_l, k) => paras.add(Math.floor(k / 4)));
    expect(paras.size).toBe(45);
  });

  it('several cells continuing at once, and a cell whose lines all fit, stay in step', () => {
    const cells = [cell(20, 3), cell(1, 1), cell(30, 2), cell(12, 5)];
    const r = layOut(cells, area(0).top + 500, 0);
    expectAllVisible(r, 0);
  });

  it('a table in the cell (one piece that cannot be split) moves to the next page whole', () => {
    const c = cell(20, 4);
    // Paragraphs 7 and 8 (lines 24..31) become one 155 px piece, as measureRow reports a nested
    // table; it straddles the first page's end.
    const unit = { top: c.lines[24].top, bottom: c.lines[31].bottom, index: 0, count: 1, rule: { can: false, min: 0 } };
    const nested = { ...c, lines: [...c.lines.slice(0, 24), unit, ...c.lines.slice(32)] };
    const r = layOut([nested], area(0).top + 150, header);
    expect(area(0).top + 150 + unit.top).toBeLessThan(area(0).bottom);
    expect(area(0).top + 150 + unit.bottom).toBeGreaterThan(area(0).bottom);
    expectAllVisible(r, header);
    const u = r.lines[0][24];
    expect(u.bottom - u.top).toBeCloseTo(unit.bottom - unit.top, 5);
  });

  it('what the row grows by moves the content after it just below its last line', () => {
    const cells = [cell(1, 2), cell(21, 4)];
    const r = layOut(cells, rowY, header);
    const lastLine = r.lines[1][r.lines[1].length - 1];
    const grownBottom = rowY + Math.max(...cells.map((c, ci) => c.bottom + r.extra[ci] + c.padBottom));
    expect(grownBottom).toBeCloseTo(lastLine.bottom, 3);
  });
});

// ----- what a split row grows by, measured from the page -----

function rect(top: number, height: number, left = 0, width = 100): DOMRect {
  return { x: left, y: top, top, left, right: left + width, bottom: top + height, width, height, toJSON() {} } as DOMRect;
}
function at<T extends Element>(el: T, r: DOMRect): T {
  el.getBoundingClientRect = () => r;
  return el;
}

describe('a split row grows by what its spacers add, not by its borders (GOV-ISSUE-017)', () => {
  it('a 1 px border share of the row is not counted as growth', () => {
    const table = document.createElement('table');
    const tbody = document.createElement('tbody');
    const tr = document.createElement('tr');
    table.append(tbody);
    tbody.append(tr);
    document.body.append(table);
    // Cell 1: a 100 px spacer, then a 20 px paragraph. Cell 2: a 20 px paragraph.
    const td1 = document.createElement('td');
    const spacer = at(document.createElement('div'), rect(100, 100));
    spacer.className = 'dx-spacer dx-in-cell';
    const p1 = at(document.createElement('p'), rect(200, 20));
    td1.append(spacer, p1);
    const td2 = document.createElement('td');
    td2.append(at(document.createElement('p'), rect(100, 20)));
    tr.append(td1, td2);
    // The row is its content (20 px), the spacer (100 px) and its share of the borders (1 px).
    at(tr, rect(100, 121));
    const view = { posAtDOM: () => { throw new Error('no position'); }, state: {} } as any;
    expect(rowGrowth(view, tr, 1)).toBeCloseTo(100, 5);
    table.remove();
  });
});

// ----- the repeated header rows over a row's continuation -----

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

describe('header rows repeated over a split row look like the table’s own (GOV-ISSUE-017)', () => {
  it('the copy sits inside the document and under its table style and cell margins', async () => {
    const styles = `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>';
    const cellXml = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const body = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="8000" w:type="dxa"/><w:tblCellMar><w:left w:w="200" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cellXml('HEAD A')}${cellXml('HEAD B')}</w:tr><w:tr>${cellXml('long')}${cellXml('text')}</w:tr></w:tbl><w:p/>` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
    zip.file('word/styles.xml', styles);
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    const first = (ed as any).firstPage();
    const pages = [
      { ...first, cuts: [{ top: 900, bottom: first.height + 24 + first.textTop, left: 90, right: 700, border: '1px solid rgb(0, 0, 0)', header: { table: 0, rows: [1] } }] },
      { ...first, inSection: 1, number: 2, top: first.height + 24 },
    ];
    (ed as any).renderPages(pages);
    const root = ed.view!.dom.closest('.dx-root')!;
    const original = ed.view!.dom.querySelector('table')!;
    const copy = root.querySelector<HTMLElement>('.dx-cuts > .dx-cut-header')!;
    const copyTable = copy.querySelector('table')!;
    // The table style's rules are written as ".dx-doc :where(.dx-ts-X) .dx-p": the class must be
    // on an element inside a .dx-doc, as it is in the body.
    const styleClass = [...original.parentElement!.classList].find((c) => c.startsWith('dx-ts-'))!;
    expect(styleClass).toBe('dx-ts-TableGrid');
    const inBody = original.querySelector('.dx-p')!;
    const inCopy = copyTable.querySelector('.dx-p')!;
    const rule = `.dx-doc .${styleClass} .dx-p`;
    expect(inBody.matches(rule)).toBe(true);
    expect(inCopy.matches(rule)).toBe(true);
    // The table's own cell margins (custom properties on the body's wrapper) come along; the
    // wrapper's section padding does not (the copy is placed at the row's own left edge).
    const wrapperStyle = original.parentElement!.getAttribute('style') ?? '';
    expect(wrapperStyle).toContain('--');
    const copyWrapper = copyTable.parentElement!;
    for (const prop of wrapperStyle.split(';').map((s) => s.split(':')[0].trim()).filter((p) => p.startsWith('--'))) {
      expect(copyWrapper.style.getPropertyValue(prop)).toBe(original.parentElement!.style.getPropertyValue(prop));
    }
    expect(copyWrapper.style.paddingLeft).toBe('');
    ed.destroy();
    host.remove();
  });
});
