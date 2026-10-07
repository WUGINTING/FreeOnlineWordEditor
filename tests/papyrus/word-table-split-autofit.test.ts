// @vitest-environment jsdom
// Word's 分割表格 (Split Table) and 自動調整 (AutoFit), 表格版面配置 › 合併 / 儲存格大小:
//  - 分割表格 splits the table above the cursor's row into two tables with an empty paragraph
//    between them (the cursor goes there); in the first row it puts an empty paragraph before
//    the table. Both tables keep the table's properties, rows keep their own; a vertical merge
//    across the line is cut in two (the lower part restarts, empty); the innermost of nested
//    tables is split; Ctrl+Shift+Enter does it in a table (a column break elsewhere);
//  - 自動調整成內容大小 / 自動調整成視窗大小 / 固定欄寬 write the w:tblPr, w:tblGrid and w:tcW Word
//    writes, with the column widths worked out as Word does (pure function, tested here with
//    given measurements: jsdom has no layout);
//  - both are one undo step and refused while 追蹤修訂 is on; the ribbon has them.
import { beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick } from 'vue';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { undo, undoDepth } from 'prosemirror-history';
import type { Node as PMNode } from 'prosemirror-model';
import { CellSelection, TableMap } from 'prosemirror-tables';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { splitTable, tableAt } from '../../src/papyrus/editor/tableCommands';
import { insertTable } from '../../src/papyrus/editor/commands';
import { autoFitTable } from '../../src/papyrus/editor/tableAutoFit';
import { autoFitWidths, columnRanges, scaleWidths, type CellRange } from '../../src/papyrus/docx/tableLayout';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { schema } from '../../src/papyrus/editor/schema';
import { parseLayers } from '../../src/papyrus/docx/wrappers';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
/** A page whose text area is 10000 twips wide. */
const SECT = '<w:sectPr><w:pgSz w:w="13000" w:h="16838"/><w:pgMar w:top="1440" w:right="1500" w:bottom="1440" w:left="1500" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const TBLPR =
  '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/><w:tblInd w:w="10" w:type="dxa"/>' +
  '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0"/></w:tblPr>';
const GRID = '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid>';
const WIDTHS = [2000, 3000, 4000];
const tc = (i: number, text: string | null, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${WIDTHS[i]}" w:type="dxa"/>${extra}</w:tcPr>${text == null ? '<w:p/>' : `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`}</w:tc>`;
const tr = (cells: string, trPr = '') => `<w:tr>${trPr}${cells}</w:tr>`;
/**
 * 5 rows × 3 columns: a repeated header row, then column 1 merged down rows 2–5 (w:vMerge),
 * row 4 with "don't split" (w:cantSplit).
 */
const TABLE =
  `<w:tbl>${TBLPR}${GRID}` +
  tr(tc(0, 'H1') + tc(1, 'H2') + tc(2, 'H3'), '<w:trPr><w:tblHeader/></w:trPr>') +
  tr(tc(0, '合併', '<w:vMerge w:val="restart"/>') + tc(1, 'b1') + tc(2, 'c1')) +
  tr(tc(0, null, '<w:vMerge/>') + tc(1, 'b2') + tc(2, 'c2')) +
  tr(tc(0, null, '<w:vMerge/>') + tc(1, 'b3') + tc(2, 'c3'), '<w:trPr><w:cantSplit/></w:trPr>') +
  tr(tc(0, null, '<w:vMerge/>') + tc(1, 'b4') + tc(2, 'c4')) +
  '</w:tbl>';

async function open(body: string, notices: string[] = []) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return editor;
}
async function documentXml(editor: DocxEditor): Promise<string> {
  const zip = await JSZip.loadAsync(await editor.save());
  return zip.file('word/document.xml')!.async('string');
}
const bodyOf = (xml: string) => xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('<w:sectPr'));
const tables = (xml: string) => xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? [];
const rowsOf = (tbl: string) => tbl.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) ?? [];
const cellsOf = (row: string) => row.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];

/** Cursor in the text `text` (its first occurrence). */
function cursorIn(editor: DocxEditor, text: string) {
  const view = editor.view!;
  let at = -1;
  view.state.doc.descendants((node, pos) => {
    if (at < 0 && node.isText && node.text!.includes(text)) at = pos + node.text!.indexOf(text);
    return at < 0;
  });
  if (at < 0) throw new Error(`no ${text}`);
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
}
const topTypes = (doc: PMNode) => {
  const out: string[] = [];
  doc.forEach((n) => out.push(n.type.name));
  return out;
};

describe('分割表格 (Split Table)', () => {
  it('splits above the cursor row: two tables, an empty paragraph between, cursor there', async () => {
    const editor = await open(TABLE + '<w:p><w:r><w:t>之後</w:t></w:r></w:p>');
    const original = await documentXml(editor);
    cursorIn(editor, 'b3');
    expect(editor.run(splitTable)).toBe(true);
    const view = editor.view!;
    expect(topTypes(view.state.doc)).toEqual(['table', 'paragraph', 'table', 'paragraph']);
    const $from = view.state.selection.$from;
    expect($from.depth).toBe(1);
    expect($from.index(0)).toBe(1);
    expect($from.parent.content.size).toBe(0);

    const xml = await documentXml(editor);
    const body = bodyOf(xml);
    expect(body).toMatch(/<\/w:tbl><w:p\b[^>]*?(\/>|>(?:(?!<w:t\b)[\s\S])*?<\/w:p>)<w:tbl>/);
    const [upper, lower] = tables(xml);
    // Both keep the table's properties and grid as written.
    for (const t of [upper, lower]) {
      expect(t.startsWith(`<w:tbl>${TBLPR}${GRID}<w:tr`)).toBe(true);
    }
    expect(rowsOf(upper)).toHaveLength(3);
    expect(rowsOf(lower)).toHaveLength(2);
    // The header row stays in the upper table only; rows keep their own w:trPr.
    expect(upper).toContain('<w:tblHeader/>');
    expect(lower).not.toContain('tblHeader');
    expect(lower).not.toContain('H1');
    expect(rowsOf(lower)[0]).toContain('<w:cantSplit/>');
    // The merge crossing the line is cut in two: restart + continue above, restart (empty) + continue below.
    const [, u1, u2] = rowsOf(upper);
    expect(cellsOf(u1)[0]).toContain('<w:vMerge w:val="restart"/>');
    expect(cellsOf(u1)[0]).toContain('合併');
    expect(cellsOf(u2)[0]).toMatch(/<w:vMerge\/>/);
    const [l0, l1] = rowsOf(lower);
    expect(cellsOf(l0)).toHaveLength(3);
    expect(cellsOf(l0)[0]).toContain('<w:vMerge w:val="restart"/>');
    expect(cellsOf(l0)[0]).toContain(`<w:tcW w:w="2000" w:type="dxa"/>`);
    expect(cellsOf(l0)[0]).not.toContain('<w:t>');
    expect(cellsOf(l1)[0]).toMatch(/<w:vMerge\/>/);
    expect(cellsOf(l1)).toHaveLength(3);
    // No cell content lost, nothing duplicated.
    for (const t of ['H1', 'H2', 'H3', '合併', 'b1', 'c1', 'b2', 'c2', 'b3', 'c3', 'b4', 'c4', '之後']) {
      expect(xml.split(`>${t}<`).length - 1, t).toBe(1);
    }
    // One undo step gives the file back.
    expect(undoDepth(view.state)).toBe(1);
    undo(view.state, view.dispatch);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });

  it('in the first row, puts an empty paragraph before the table (the table unchanged)', async () => {
    const editor = await open(TABLE + '<w:p/>');
    const original = await documentXml(editor);
    cursorIn(editor, 'H2');
    expect(editor.run(splitTable)).toBe(true);
    const view = editor.view!;
    expect(topTypes(view.state.doc)).toEqual(['paragraph', 'table', 'paragraph']);
    expect(view.state.selection.$from.index(0)).toBe(0);
    const xml = await documentXml(editor);
    expect(bodyOf(xml).startsWith('<w:p')).toBe(true);
    expect(tables(xml)).toEqual(tables(original));
    undo(view.state, view.dispatch);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });

  it('with a selection, splits above the row where it starts', async () => {
    const editor = await open(TABLE + '<w:p/>');
    const view = editor.view!;
    let from = -1;
    let to = -1;
    view.state.doc.descendants((n, pos) => {
      if (n.isText && n.text === 'b2') from = pos;
      if (n.isText && n.text === 'c4') to = pos + 2;
      return true;
    });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
    editor.run(splitTable);
    const [upper, lower] = tables(await documentXml(editor));
    expect(rowsOf(upper)).toHaveLength(2);
    expect(rowsOf(lower)).toHaveLength(3);
    editor.destroy();
  });

  it('Ctrl+Shift+Enter splits the table in a table and still inserts a column break elsewhere', async () => {
    const editor = await open(TABLE + '<w:p><w:r><w:t>之後</w:t></w:r></w:p>');
    const view = editor.view!;
    const press = () =>
      view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    cursorIn(editor, 'b2');
    press();
    expect(topTypes(view.state.doc)).toEqual(['table', 'paragraph', 'table', 'paragraph']);
    cursorIn(editor, '之後');
    press();
    expect(topTypes(view.state.doc)).toEqual(['table', 'paragraph', 'table', 'paragraph']);
    let column = 0;
    view.state.doc.descendants((n) => {
      if (n.type.name === 'hard_break' && n.attrs.type === 'column') column++;
    });
    expect(column).toBe(1);
    editor.destroy();
  });

  it('splits the innermost of nested tables', async () => {
    const inner = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
      ['n1', 'n2', 'n3'].map((t) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc></w:tr>`).join('') +
      '</w:tbl>';
    const outer = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${inner}<w:p/></w:tc></w:tr></w:tbl>`;
    const editor = await open(outer + '<w:p/>');
    cursorIn(editor, 'n2');
    editor.run(splitTable);
    const view = editor.view!;
    expect(topTypes(view.state.doc)).toEqual(['table', 'paragraph']);
    const cell = view.state.doc.child(0).child(0).child(0);
    expect(topTypes(cell)).toEqual(['table', 'paragraph', 'table', 'paragraph']);
    const xml = await documentXml(editor);
    expect(xml.match(/<w:tbl>/g)).toHaveLength(3);
    expect(xml).toMatch(/n1<\/w:t>[\s\S]*<\/w:tbl><w:p(?:\/>|><\/w:p>)<w:tbl>[\s\S]*n2<\/w:t>/);
    editor.destroy();
  });

  it('is refused with the table notice while 追蹤修訂 is on (the first row too)', async () => {
    const notices: string[] = [];
    const editor = await open(TABLE + '<w:p/>', notices);
    const original = await documentXml(editor);
    editor.setTrackChanges(true);
    for (const at of ['b3', 'H1']) {
      cursorIn(editor, at);
      editor.run(splitTable);
    }
    expect(notices).toEqual([
      '追蹤修訂開啟時，無法插入、刪除或調整表格的結構，請先關閉「追蹤修訂」。',
      '追蹤修訂開啟時，無法插入、刪除或調整表格的結構，請先關閉「追蹤修訂」。',
    ]);
    expect(bodyOf(await documentXml(editor))).toBe(bodyOf(original));
    editor.destroy();
  });

  it('leaves one row of a merge below (or above) the cut: that row is a plain cell, no w:vMerge', async () => {
    const editor = await open(TABLE + '<w:p/>');
    cursorIn(editor, 'b4');
    editor.run(splitTable);
    let [upper, lower] = tables(await documentXml(editor));
    const [r1, r2, r3] = rowsOf(upper).slice(1);
    expect(cellsOf(r1)[0]).toContain('<w:vMerge w:val="restart"/>');
    expect(cellsOf(r2)[0]).toMatch(/<w:vMerge\/>/);
    expect(cellsOf(r3)[0]).toMatch(/<w:vMerge\/>/);
    expect(rowsOf(lower)).toHaveLength(1);
    // The row below the cut: its piece's w:tcPr without the merge, and an empty paragraph.
    expect(cellsOf(rowsOf(lower)[0])[0]).toBe('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p></w:p></w:tc>');

    // And a single row above: split at the second row of the merge.
    const view = editor.view!;
    undo(view.state, view.dispatch);
    cursorIn(editor, 'b2');
    editor.run(splitTable);
    [upper, lower] = tables(await documentXml(editor));
    expect(cellsOf(rowsOf(upper)[1])[0]).toBe('<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>合併</w:t></w:r></w:p></w:tc>');
    expect(cellsOf(rowsOf(lower)[0])[0]).toContain('<w:vMerge w:val="restart"/>');
    expect(cellsOf(rowsOf(lower)[2])[0]).toMatch(/<w:vMerge\/>/);
    editor.destroy();
  });

  it('not in read-only mode: neither Ctrl+Shift+Enter nor DocxEditor.splitTable()', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${TABLE}<w:p/>${SECT}</w:body></w:document>`);
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host, { editable: false });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    const view = editor.view!;
    cursorIn(editor, 'b3');
    const before = view.state.doc;
    const e = new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    view.dom.dispatchEvent(e);
    // The key binding itself says no, whatever the view does with keys when it is not editable.
    expect(view.someProp('handleKeyDown', (f) => f(view, e))).toBeFalsy();
    expect(editor.splitTable()).toBe(false);
    expect(view.state.doc).toBe(before);
    editor.destroy();
  });

  it('works in a header being edited (split, Ctrl+Shift+Enter, 自動調整), saved in the header part', async () => {
    const editor = await open('<w:p><w:r><w:t>本文</w:t></w:r></w:p>');
    editor.editHeaderFooter('header', 0);
    expect(editor.target).toBe('header');
    const hf = editor.activeView!;
    expect(hf).not.toBe(editor.view);
    editor.run(insertTable(4, 2, 600));
    // Text in each cell (h1 … h8).
    const cells: number[] = [];
    hf.state.doc.descendants((n, p) => {
      if (n.type.name === 'table_cell') cells.push(p);
    });
    const t = hf.state.tr;
    [...cells].reverse().forEach((p, i) => t.insertText(`h${cells.length - i}`, p + 2));
    hf.dispatch(t);
    const cursorAt = (text: string) => {
      let at = -1;
      hf.state.doc.descendants((n, p) => {
        if (at < 0 && n.isText && n.text === text) at = p;
      });
      hf.dispatch(hf.state.tr.setSelection(TextSelection.create(hf.state.doc, at)));
    };
    cursorAt('h3');
    expect(editor.splitTable()).toBe(true);
    expect(topTypes(hf.state.doc).filter((x) => x === 'table')).toHaveLength(2);
    // Ctrl+Shift+Enter in the header splits the lower table again.
    cursorAt('h7');
    hf.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    expect(topTypes(hf.state.doc).filter((x) => x === 'table')).toHaveLength(3);
    cursorAt('h7');
    expect(editor.autoFitTable('window')).toBe(true);
    editor.closeHeaderFooter();
    const zip = await JSZip.loadAsync(await editor.save());
    const parts = Object.keys(zip.files).filter((f) => /^word\/header\d*\.xml$/.test(f));
    const xml = (await Promise.all(parts.map((f) => zip.file(f)!.async('string')))).find((x) => x.includes('h1'))!;
    expect(xml.match(/<w:tbl>/g)).toHaveLength(3);
    expect(xml.match(/<\/w:tbl><w:p(?:\/>|><\/w:p>)<w:tbl>/g)).toHaveLength(2);
    // Only the table with the cursor: 100 %.
    expect(xml.match(/<w:tblW w:w="5000" w:type="pct"\/>/g)).toHaveLength(1);
    expect(xml.lastIndexOf('<w:tblW w:w="5000" w:type="pct"/>')).toBeGreaterThan(xml.lastIndexOf('<w:tbl>'));
    // The body is untouched.
    expect(tables(await documentXml(editor))).toEqual([]);
    editor.destroy();
  });

  it('does nothing outside a table, and an unedited file is saved as it was', async () => {
    const editor = await open(TABLE + '<w:p><w:r><w:t>之後</w:t></w:r></w:p>');
    const original = await documentXml(editor);
    expect(tables(original)).toEqual([TABLE]);
    cursorIn(editor, '之後');
    expect(editor.can(splitTable)).toBe(false);
    expect(editor.run(splitTable)).toBe(false);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });
});

describe('自動調整: the column widths (pure)', () => {
  it('contents that fit keep their widest widths', () => {
    expect(autoFitWidths([{ min: 500, max: 2000 }, { min: 500, max: 3000 }], 10000)).toEqual([2000, 3000]);
  });

  it('too wide: short columns stay on one line, the long ones narrowed to fit, adding up to the text width', () => {
    // A real 公文 table (measured in the browser at 12 pt): 序號 | 項目名稱 | 說明 (a long sentence) | 金額.
    const w = autoFitWidths([{ min: 456, max: 696 }, { min: 456, max: 1656 }, { min: 696, max: 7896 }, { min: 1006, max: 1006 }], 8306);
    expect(w).toEqual([696, 1656, 4948, 1006]);
    expect(w.reduce((s, x) => s + x, 0)).toBe(8306);
  });

  it('too wide: long columns share the room in proportion to their content, none below its minimum', () => {
    expect(autoFitWidths([{ min: 500, max: 9000 }, { min: 500, max: 6000 }, { min: 200, max: 300 }], 7300)).toEqual([4200, 2800, 300]);
    // The second column would get 2800 but can't go under 3000: it keeps 3000, the first takes the rest.
    const w = autoFitWidths([{ min: 1000, max: 9000 }, { min: 3000, max: 6000 }, { min: 200, max: 300 }], 7300);
    expect(w).toEqual([4000, 3000, 300]);
  });

  it('when keeping the short columns would squeeze a long one under its minimum, all are narrowed', () => {
    const w = autoFitWidths([{ min: 1000, max: 1000 }, { min: 4000, max: 5000 }, { min: 100, max: 800 }], 5200);
    expect(w.reduce((s, x) => s + x, 0)).toBe(5200);
    expect(w[0]).toBe(1000);
    expect(w[1]).toBeGreaterThanOrEqual(4000);
  });

  it('when even the minimums do not fit, every column gets its minimum (the table runs past the margin)', () => {
    expect(autoFitWidths([{ min: 2500, max: 6000 }, { min: 2000, max: 3000 }], 4000)).toEqual([2500, 2000]);
  });

  it('merged cells widen the columns they span only when those are too narrow', () => {
    const cells: CellRange[] = [
      { col: 0, span: 1, min: 100, max: 200 },
      { col: 1, span: 1, min: 100, max: 300 },
      { col: 0, span: 2, min: 300, max: 1000 },
    ];
    expect(columnRanges(cells, 2)).toEqual([{ min: 150, max: 400 }, { min: 150, max: 600 }]);
    expect(columnRanges([{ col: 0, span: 2, min: 100, max: 400 }], 2)).toEqual([{ min: 50, max: 200 }, { min: 50, max: 200 }]);
  });

  it('自動調整成視窗大小 scales the columns to the text width', () => {
    expect(scaleWidths([2000, 3000, 4000], 10000)).toEqual([2222, 3333, 4445]);
  });
});

describe('自動調整 (AutoFit): what is written', () => {
  const TEXT = 10000;
  const tblPrOf = (tbl: string) => tbl.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)![0];
  const names = (tblPr: string) => [...tblPr.matchAll(/<w:(\w+)[ />]/g)].map((m) => m[1]).filter((n) => n !== 'tblPr');
  const topNames = (tblPr: string) => names(tblPr.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '<w:tblBorders/>').replace(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/, '<w:tblCellMar/>'));
  const gridOf = (tbl: string) => [...tbl.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]));
  const tcWs = (tbl: string) => rowsOf(tbl).map((r) => cellsOf(r).map((c) => c.match(/<w:tcW ([^>]*)\/>/)![1]));

  /** Measured content ranges: column 1 short, 2 long, 3 medium. */
  const measure = (_v: unknown, table: PMNode) => {
    const ranges = [{ min: 600, max: 1200 }, { min: 900, max: 2500 }, { min: 700, max: 1800 }];
    const map = TableMap.get(table);
    return [...new Set(map.map)].map((p): CellRange => {
      const rect = map.findCell(p);
      return { col: rect.left, span: rect.right - rect.left, ...ranges[rect.left] };
    });
  };

  it('自動調整成內容大小: w:tblW auto, automatic layout, grid and w:tcW from the content', async () => {
    const editor = await open(TABLE + '<w:p/>');
    const original = await documentXml(editor);
    cursorIn(editor, 'b2');
    expect(editor.run(autoFitTable('contents', { textWidth: TEXT, measure }))).toBe(true);
    const [tbl] = tables(await documentXml(editor));
    const tblPr = tblPrOf(tbl);
    expect(tblPr).toContain('<w:tblW w:w="0" w:type="auto"/>');
    // Written, not left out: a table style may say fixed (tblLayout is in CT_TblPrBase).
    expect(tblPr).toContain('<w:tblLayout w:type="autofit"/>');
    expect(tblPr).not.toContain('fixed');
    // Every other child kept, in its place.
    expect(topNames(tblPr)).toEqual(['tblStyle', 'tblW', 'jc', 'tblInd', 'tblBorders', 'tblLayout', 'tblCellMar', 'tblLook']);
    expect(tblPr).toContain('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>');
    expect(gridOf(tbl)).toEqual([1200, 2500, 1800]);
    // Every cell, the merged cell's pieces too.
    for (const row of tcWs(tbl)) expect(row).toEqual(['w:w="1200" w:type="dxa"', 'w:w="2500" w:type="dxa"', 'w:w="1800" w:type="dxa"']);
    // The rest of the table is as it was.
    expect(tbl.replace(/<w:tblPr>[\s\S]*?<\/w:tblPr>|<w:tblGrid>[\s\S]*?<\/w:tblGrid>|<w:tcW [^>]*\/>/g, '')).toBe(
      tables(original)[0].replace(/<w:tblPr>[\s\S]*?<\/w:tblPr>|<w:tblGrid>[\s\S]*?<\/w:tblGrid>|<w:tcW [^>]*\/>/g, ''),
    );
    // Shown with the new widths.
    const colwidths: number[] = [];
    editor.view!.state.doc.child(0).child(0).forEach((c) => colwidths.push(...c.attrs.colwidth));
    expect(colwidths).toEqual([80, 167, 120]);
    // One undo step.
    const view = editor.view!;
    expect(undoDepth(view.state)).toBe(1);
    undo(view.state, view.dispatch);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });

  it('自動調整成內容大小 needs the page layout: without it (no measurement) nothing changes', async () => {
    const editor = await open(TABLE + '<w:p/>');
    const original = await documentXml(editor);
    cursorIn(editor, 'b2');
    expect(editor.autoFitTable('contents')).toBe(false);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });

  it('自動調整成視窗大小: w:tblW 100 %, automatic layout, columns scaled to the text width', async () => {
    const editor = await open(TABLE + '<w:p/>');
    cursorIn(editor, 'b2');
    expect(editor.autoFitTable('window')).toBe(true);
    const [tbl] = tables(await documentXml(editor));
    const tblPr = tblPrOf(tbl);
    expect(tblPr).toContain('<w:tblW w:w="5000" w:type="pct"/>');
    expect(tblPr).toContain('<w:tblLayout w:type="autofit"/>');
    expect(tblPr).not.toContain('fixed');
    expect(topNames(tblPr)).toEqual(['tblStyle', 'tblW', 'jc', 'tblInd', 'tblBorders', 'tblLayout', 'tblCellMar', 'tblLook']);
    expect(gridOf(tbl)).toEqual([2222, 3333, 4445]);
    for (const row of tcWs(tbl)) expect(row).toEqual(['w:w="2222" w:type="dxa"', 'w:w="3333" w:type="dxa"', 'w:w="4445" w:type="dxa"']);
    editor.destroy();
  });

  it('固定欄寬: fixed layout, the current widths in twips, w:tblW their total; inserted in schema order', async () => {
    // A 100 % table with percentage cell widths and no w:tblLayout.
    const pctTable =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="6000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="pct"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="pct"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const editor = await open(pctTable + '<w:p/>');
    const original = await documentXml(editor);
    cursorIn(editor, '甲');
    expect(editor.autoFitTable('fixed')).toBe(true);
    const [tbl] = tables(await documentXml(editor));
    const tblPr = tblPrOf(tbl);
    expect(tblPr).toContain('<w:tblW w:w="10000" w:type="dxa"/>');
    expect(tblPr).toContain('<w:tblLayout w:type="fixed"/>');
    expect(topNames(tblPr)).toEqual(['tblStyle', 'tblW', 'tblBorders', 'tblLayout', 'tblCellMar', 'tblLook']);
    expect(gridOf(tbl)).toEqual([4000, 6000]);
    expect(tcWs(tbl)).toEqual([['w:w="4000" w:type="dxa"', 'w:w="6000" w:type="dxa"']]);
    expect(tbl).toContain('<w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>');
    const view = editor.view!;
    undo(view.state, view.dispatch);
    expect(await documentXml(editor)).toBe(original);
    editor.destroy();
  });

  it('a table whose style says fixed layout gets w:tblLayout autofit of its own, in schema order', async () => {
    // The table style sets w:tblLayout fixed; the table itself has none.
    const style =
      '<w:style w:type="table" w:styleId="FixedGrid"><w:name w:val="Fixed Grid"/>' +
      '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr></w:style>';
    const table =
      '<w:tbl><w:tblPr><w:tblStyle w:val="FixedGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const zip = blankPackage();
    const styles = await zip.file('word/styles.xml')!.async('string');
    expect(styles).toContain('</w:styles>');
    zip.file('word/styles.xml', styles.replace('</w:styles>', `${style}</w:styles>`));
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${table}<w:p/>${SECT}</w:body></w:document>`);
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    for (const mode of ['contents', 'window'] as const) {
      cursorIn(editor, '甲');
      const ok = mode === 'contents' ? editor.run(autoFitTable('contents', { textWidth: TEXT, measure })) : editor.autoFitTable('window');
      expect(ok).toBe(true);
      const tblPr = tblPrOf(tables(await documentXml(editor))[0]);
      expect(tblPr, mode).toContain('<w:tblLayout w:type="autofit"/>');
      expect(topNames(tblPr), mode).toEqual(['tblStyle', 'tblW', 'tblLayout', 'tblLook']);
    }
    // 固定欄寬 turns it to fixed in place.
    editor.autoFitTable('fixed');
    const tblPr = tblPrOf(tables(await documentXml(editor))[0]);
    expect(tblPr).toContain('<w:tblLayout w:type="fixed"/>');
    expect(tblPr).not.toContain('autofit');
    editor.destroy();
  });

  it('works on a table made in the editor (no w:tblPr yet)', async () => {
    const editor = await open('<w:p/>');
    editor.run(insertTable(2, 2, 600));
    expect(editor.autoFitTable('window')).toBe(true);
    const [tbl] = tables(await documentXml(editor));
    expect(tblPrOf(tbl)).toContain('<w:tblW w:w="5000" w:type="pct"/>');
    expect(tblPrOf(tbl)).toContain('<w:tblBorders>');
    expect(gridOf(tbl)).toEqual([5000, 5000]);
    editor.destroy();
  });

  it('is refused with the table notice while 追蹤修訂 is on', async () => {
    const notices: string[] = [];
    const editor = await open(TABLE + '<w:p/>', notices);
    const original = await documentXml(editor);
    editor.setTrackChanges(true);
    cursorIn(editor, 'b2');
    editor.autoFitTable('fixed');
    editor.autoFitTable('window');
    expect(notices).toEqual([
      '追蹤修訂開啟時，無法變更表格或儲存格的格式，請先關閉「追蹤修訂」。',
      '追蹤修訂開啟時，無法變更表格或儲存格的格式，請先關閉「追蹤修訂」。',
    ]);
    expect(bodyOf(await documentXml(editor))).toBe(bodyOf(original));
    editor.destroy();
  });

  it('not in read-only mode', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${TABLE}<w:p/>${SECT}</w:body></w:document>`);
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host, { editable: false });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    cursorIn(editor, 'b2');
    expect(editor.autoFitTable('window')).toBe(false);
    editor.destroy();
  });
});

/** A table from rows of cell XML (each `[width, tcPr extra, text | null]`) over a grid. */
function tableXml(grid: number[], rows: [number, string, string | null][][], tblPr = '<w:tblW w:w="0" w:type="auto"/>', rowPr: string[] = []) {
  const cell = ([w, extra, text]: [number, string, string | null]) =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}</w:tcPr>${text == null ? '<w:p/>' : `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`}</w:tc>`;
  return `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${grid.map((g) => `<w:gridCol w:w="${g}"/>`).join('')}</w:tblGrid>` +
    rows.map((r, i) => `<w:tr>${rowPr[i] ?? ''}${r.map(cell).join('')}</w:tr>`).join('') + '</w:tbl>';
}
/** Position before the cell whose text is `text`. */
function cellPos(editor: DocxEditor, text: string): number {
  let at = -1;
  editor.view!.state.doc.descendants((n, p) => {
    if (at < 0 && n.type.name === 'table_cell' && n.textContent === text) at = p;
    return at < 0;
  });
  return at;
}

describe('分割表格: merged cells, selections and wrappers', () => {
  it('the lower half of a merge takes its own row’s looks and content, not the top cell’s', async () => {
    // The top cell is yellow and centred; the pieces below have neither. The first piece holds a bookmark.
    const YELLOW = '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:vAlign w:val="center"/>';
    const piece1 = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p><w:bookmarkStart w:id="7" w:name="lower"/><w:bookmarkEnd w:id="7"/></w:p></w:tc>';
    const piece2 = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>';
    const plain = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>${YELLOW}</w:tcPr><w:p><w:r><w:t>頂</w:t></w:r></w:p></w:tc>${plain('r0')}</w:tr>` +
      `<w:tr>${piece1}${plain('r1')}</w:tr><w:tr>${piece2}${plain('r2')}</w:tr></w:tbl>`;
    const editor = await open(table + '<w:p/>');
    cursorIn(editor, 'r1');
    editor.run(splitTable);
    const view = editor.view!;
    // On screen.
    const lowerTable = view.state.doc.child(2);
    const lowerCell = lowerTable.child(0).child(0);
    expect(lowerCell.attrs.background).toBeNull();
    expect(lowerCell.attrs.vAlign).toBeNull();
    expect(lowerCell.attrs.rowspan).toBe(2);
    let lowerPos = -1;
    view.state.doc.forEach((_n, off, i) => {
      if (i === 2) lowerPos = off;
    });
    const td = view.nodeDOM(lowerPos + 2) as HTMLElement;
    expect(td.tagName).toBe('TD');
    expect(td.style.backgroundColor).toBe('');
    expect(td.style.verticalAlign).toBe('');
    expect(view.state.doc.child(0).child(0).child(0).attrs.background).toBe('#FFFF00');
    // In the file.
    const [upper, lower] = tables(await documentXml(editor));
    expect(cellsOf(rowsOf(upper)[0])[0]).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:vAlign w:val="center"/>');
    const l0 = cellsOf(rowsOf(lower)[0])[0];
    expect(l0).toContain('<w:vMerge w:val="restart"/>');
    expect(l0).not.toContain('w:shd');
    expect(l0).not.toContain('w:vAlign');
    // The piece's content (its bookmark) is kept.
    expect(l0).toContain('<w:bookmarkStart w:id="7" w:name="lower"/><w:bookmarkEnd w:id="7"/>');
    expect(cellsOf(rowsOf(lower)[1])[0]).toBe(piece2);
    editor.destroy();
  });

  it('a block merged across and down that crosses the cut is cut in two, still two columns wide', async () => {
    const table = tableXml([2000, 2000, 2000], [
      [[4000, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>', '大'], [2000, '', 'c0']],
      [[4000, '<w:gridSpan w:val="2"/><w:vMerge/>', null], [2000, '', 'c1']],
      [[4000, '<w:gridSpan w:val="2"/><w:vMerge/>', null], [2000, '', 'c2']],
    ]);
    const editor = await open(table + '<w:p/>');
    cursorIn(editor, 'c1');
    editor.run(splitTable);
    const [upper, lower] = tables(await documentXml(editor));
    expect(rowsOf(upper)).toHaveLength(1);
    const u = cellsOf(rowsOf(upper)[0]);
    expect(u).toHaveLength(2);
    expect(u[0]).toContain('<w:gridSpan w:val="2"/>');
    expect(u[0]).not.toContain('vMerge');
    expect(u[0]).toContain('大');
    const [l0, l1] = rowsOf(lower).map(cellsOf);
    expect(l0).toHaveLength(2);
    expect(l1).toHaveLength(2);
    expect(l0[0]).toContain('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>');
    expect(l1[0]).toContain('<w:gridSpan w:val="2"/><w:vMerge/>');
    expect(l0[1]).toContain('c1');
    expect(l1[1]).toContain('c2');
    editor.destroy();
  });

  it('a cell selection across the cut splits above its top row', async () => {
    const editor = await open(TABLE + '<w:p/>');
    const view = editor.view!;
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cellPos(editor, 'b2'), cellPos(editor, 'c4'))));
    expect(view.state.selection).toBeInstanceOf(CellSelection);
    editor.run(splitTable);
    const [upper, lower] = tables(await documentXml(editor));
    expect(rowsOf(upper)).toHaveLength(2);
    expect(rowsOf(lower)).toHaveLength(3);
    expect(lower).toContain('b2');
    editor.destroy();
  });

  it('in a content control (w:sdt): saved as one control holding table, paragraph, table', async () => {
    const editor = await open(`<w:sdt><w:sdtPr><w:alias w:val="清單"/></w:sdtPr><w:sdtContent>${TABLE}</w:sdtContent></w:sdt><w:p/>`);
    cursorIn(editor, 'b3');
    editor.run(splitTable);
    const body = bodyOf(await documentXml(editor));
    expect(body.match(/<w:sdt>/g)).toHaveLength(1);
    expect(body).toMatch(/^<w:sdt><w:sdtPr><w:alias w:val="清單"\/><\/w:sdtPr><w:sdtContent><w:tbl>[\s\S]*<\/w:tbl><w:p(?:\/>|><\/w:p>)<w:tbl>[\s\S]*<\/w:tbl><\/w:sdtContent><\/w:sdt>/);
    editor.destroy();
  });

  it('the new paragraph is inside the table’s content control but not in a revision around it', async () => {
    const editor = await open(`<w:sdt><w:sdtPr><w:alias w:val="清單"/></w:sdtPr><w:sdtContent>${TABLE}</w:sdtContent></w:sdt><w:p/>`);
    const view = editor.view!;
    const table = view.state.doc.child(0);
    const layers = parseLayers(table.attrs.wrap);
    expect(layers).toHaveLength(1);
    const ins = { id: 'Lins', open: '<w:ins w:id="90" w:author="甲" w:date="2026-01-01T00:00:00Z">', close: '</w:ins>' };
    view.dispatch(view.state.tr.setNodeMarkup(0, undefined, { ...table.attrs, wrap: JSON.stringify([...layers, ins]) }));
    cursorIn(editor, 'b3');
    editor.run(splitTable);
    const para = view.state.doc.child(1);
    expect(para.type.name).toBe('paragraph');
    expect(parseLayers(para.attrs.wrap).map((l) => l.id)).toEqual([layers[0].id]);
    expect(parseLayers(view.state.doc.child(2).attrs.wrap)).toHaveLength(2);
    editor.destroy();
  });

  it('a revision around the content control (outermost) stays around the paragraph too: one w:sdt is saved', async () => {
    const sdt = `<w:sdt><w:sdtPr><w:alias w:val="清單"/><w:id w:val="4242"/></w:sdtPr><w:sdtContent>${TABLE}</w:sdtContent></w:sdt>`;
    const editor = await open(sdt + '<w:p/>');
    const view = editor.view!;
    const table = view.state.doc.child(0);
    const layers = parseLayers(table.attrs.wrap);
    // [ins, sdt]: the revision outside the control.
    const ins = { id: 'Lins2', open: '<w:ins w:id="91" w:author="甲" w:date="2026-01-01T00:00:00Z">', close: '</w:ins>' };
    view.dispatch(view.state.tr.setNodeMarkup(0, undefined, { ...table.attrs, wrap: JSON.stringify([ins, ...layers]) }));
    cursorIn(editor, 'b3');
    editor.run(splitTable);
    expect(parseLayers(view.state.doc.child(1).attrs.wrap).map((l) => l.id)).toEqual(['Lins2', layers[0].id]);
    const body = bodyOf(await documentXml(editor));
    expect(body.match(/<w:sdt>/g)).toHaveLength(1);
    expect(body.match(/<w:id w:val="4242"\/>/g)).toHaveLength(1);
    expect(body).toMatch(/<w:sdtContent><w:tbl>[\s\S]*<\/w:tbl><w:p(?:\/>|><\/w:p>)<w:tbl>[\s\S]*<\/w:tbl><\/w:sdtContent>/);
    editor.destroy();
  });
});

describe('自動調整: spans, nesting, row exceptions and measuring', () => {
  it('a cell spanning columns gets the total of their widths', async () => {
    const table = tableXml([2000, 3000, 4000], [
      [[5000, '<w:gridSpan w:val="2"/>', '跨欄'], [4000, '', 'x']],
      [[2000, '', 'a'], [3000, '', 'b'], [4000, '', 'c']],
    ]);
    const editor = await open(table + '<w:p/>');
    cursorIn(editor, 'a');
    editor.autoFitTable('window');
    const [tbl] = tables(await documentXml(editor));
    expect([...tbl.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1])).toEqual([2222, 3333, 4445]);
    expect(cellsOf(rowsOf(tbl)[0])[0]).toContain('<w:tcW w:w="5555" w:type="dxa"/><w:gridSpan w:val="2"/>');
    expect(cellsOf(rowsOf(tbl)[0])[1]).toContain('<w:tcW w:w="4445" w:type="dxa"/>');
    editor.destroy();
  });

  it('a table in a table cell fits that cell (less its margins); a selected nested table is the one fitted', async () => {
    const inner = tableXml([3000], [[[3000, '', 'in1']], [[3000, '', 'in2']]]);
    const outer =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${inner}<w:p/></w:tc>` +
      '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>外</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const grids = async () => [...(await documentXml(editor)).matchAll(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/g)].map((m) => [...m[1].matchAll(/w:w="(\d+)"/g)].map((g) => +g[1]));
    const editor = await open(outer + '<w:p/>');
    cursorIn(editor, 'in2');
    editor.autoFitTable('window');
    // 6000 less Word's default cell margins (108 + 108).
    expect(await grids()).toEqual([[6000, 2000], [5784]]);
    const view = editor.view!;
    undo(view.state, view.dispatch);
    expect(await grids()).toEqual([[6000, 2000], [3000]]);
    // The nested table selected as a whole (not the outer table).
    let innerPos = -1;
    view.state.doc.descendants((n, p) => {
      if (n.type.name === 'table' && p > 0 && innerPos < 0) innerPos = p;
      return innerPos < 0;
    });
    // The editor turns a selected table into a cell selection of it; without tableEditing (a
    // bare state) the node selection itself names the table too.
    const bare = EditorState.create({ doc: view.state.doc, selection: NodeSelection.create(view.state.doc, innerPos) });
    expect(tableAt(bare)?.pos).toBe(innerPos);
    expect(tableAt(bare)?.table.textContent).toBe('in1in2');
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, innerPos)));
    editor.autoFitTable('fixed');
    const xml = await documentXml(editor);
    const [outerTbl] = xml.match(/<w:tbl><w:tblPr>[\s\S]*?<\/w:tblPr>/g)!;
    expect(outerTbl).not.toContain('tblLayout');
    expect(xml.match(/<w:tblLayout w:type="fixed"\/>/g)).toHaveLength(1);
    expect(xml).toContain('<w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/>');
    editor.destroy();
  });

  it('a row’s own table width / layout (w:tblPrEx) is set as on the table', async () => {
    const table = tableXml([4000, 4000], [
      [[4000, '', 'a'], [4000, '', 'b']],
      [[4000, '', 'c'], [4000, '', 'd']],
    ], '<w:tblW w:w="0" w:type="auto"/>', ['', '<w:tblPrEx><w:tblW w:w="3000" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/></w:tblPrEx>']);
    const editor = await open(table + '<w:p/>');
    cursorIn(editor, 'a');
    editor.autoFitTable('window');
    const rows = rowsOf(tables(await documentXml(editor))[0]);
    expect(rows[0]).not.toContain('tblPrEx');
    expect(rows[1]).toContain('<w:tblPrEx><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblLayout w:type="autofit"/></w:tblPrEx>');
    editor.destroy();
  });

  it('a column no cell was measured in never gets less than the minimum', async () => {
    const editor = await open(tableXml([4000, 4000], [[[4000, '', 'a'], [4000, '', 'b']]]) + '<w:p/>');
    cursorIn(editor, 'a');
    const measure = () => [{ col: 0, span: 1, min: 500, max: 900 }];
    editor.run(autoFitTable('contents', { textWidth: 10000, measure }));
    const [tbl] = tables(await documentXml(editor));
    expect([...tbl.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1])).toEqual([900, 360]);
    editor.destroy();
  });

  describe('measuring in the page (getBoundingClientRect stubbed: 16 px a character)', () => {
    const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    let restore: () => void = () => {};
    /** A fake layout: cells 100 px; the measuring boxes their text at 16 px a character, pictures at their width unless capped. */
    function stubLayout() {
      const original = HTMLElement.prototype.getBoundingClientRect;
      const imgs = (el: HTMLElement) =>
        Array.from(el.querySelectorAll<HTMLElement>('img')).map((img) => {
          const capped = [img, img.closest<HTMLElement>('.dx-img-wrap')].some((e) => e && e.style.maxWidth !== 'none');
          return capped ? 0 : Number(img.getAttribute('width') ?? 0);
        });
      HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
        let width = 0;
        if (this.tagName === 'TD') width = 100;
        else if (this.style.width === 'max-content') width = (this.textContent ?? '').length * 16 + imgs(this).reduce((s, w) => s + w, 0);
        else if (this.style.width === 'min-content') width = Math.max((this.textContent ?? '').length ? 16 : 0, ...imgs(this));
        return { x: 0, y: 0, top: 0, left: 0, bottom: 0, right: width, width, height: 0, toJSON() {} } as DOMRect;
      };
      restore = () => (HTMLElement.prototype.getBoundingClientRect = original);
    }

    it('text and a picture: widths from the page, the picture at its own size; the measuring box is removed', async () => {
      const editor = await open(tableXml([4000, 4000], [[[4000, '', '序號'], [4000, '', '圖']]]) + '<w:p/>');
      const view = editor.view!;
      // A 200 px picture in place of 「圖」.
      let at = -1;
      view.state.doc.descendants((n, p) => {
        if (n.isText && n.text === '圖') at = p;
      });
      view.dispatch(view.state.tr.replaceWith(at, at + 1, schema.nodes.image.create({ src: PNG, width: 200, height: 50 })));
      cursorIn(editor, '序號');
      stubLayout();
      try {
        expect(editor.autoFitTable('contents')).toBe(true);
      } finally {
        restore();
      }
      // 序號: 32 px (+ ½ px, rounded up) = 33 px; the picture 200 px → 201 px.
      const [tbl] = tables(await documentXml(editor));
      expect([...tbl.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1])).toEqual([495, 3015]);
      expect(document.querySelectorAll('[aria-hidden="true"].dx-doc').length).toBe(0);
      editor.destroy();
    });

    it('an all-empty table gets the minimum width per column, not nothing', async () => {
      const editor = await open(tableXml([4000, 4000, 2000], [[[4000, '', null], [4000, '', null], [2000, '', null]]]) + '<w:p/>');
      const view = editor.view!;
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
      stubLayout();
      try {
        expect(editor.autoFitTable('contents')).toBe(true);
      } finally {
        restore();
      }
      const [tbl] = tables(await documentXml(editor));
      expect([...tbl.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => +m[1])).toEqual([360, 360, 360]);
      editor.destroy();
    });
  });
});

describe('ribbon: 表格版面配置 › 合併 › 分割表格, 儲存格大小 › 自動調整', () => {
  it('the buttons and the menu are there and work', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${TABLE}<w:p/>${SECT}</w:body></w:document>`);
    const src = await zip.generateAsync({ type: 'uint8array' });
    const host = document.createElement('div');
    document.body.append(host);
    let editor: DocxEditor | null = null;
    const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditor) => (editor = ed) }) });
    app.mount(host);
    const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
    for (let i = 0; i < 80 && !editor; i++) await flush();
    const ed = editor! as DocxEditor;
    cursorIn(ed, 'b3');
    await flush();
    const tab = Array.from(host.querySelectorAll<HTMLButtonElement>('[role="tab"]')).find((t) => t.textContent?.trim() === '表格版面配置')!;
    tab.click();
    await flush();
    const byTitle = (t: string) => host.querySelector(`[title="${t}"]`) as HTMLButtonElement;

    // 自動調整 › 固定欄寬 / 自動調整成視窗大小.
    const merge = byTitle('請先選取兩個以上的儲存格').closest('[role="group"]')!;
    expect(merge.getAttribute('aria-label')).toBe('合併');
    const size = byTitle('自動調整：依內容或視窗寬度調整欄寬，或固定欄寬');
    expect(size.closest('[role="group"]')!.getAttribute('aria-label')).toBe('儲存格大小');
    expect(byTitle('平均分配欄寬（選取多欄時只分配這些欄）').closest('[role="group"]')).toBe(size.closest('[role="group"]'));
    expect(size.getAttribute('aria-haspopup')).toBe('menu');
    size.click();
    await flush();
    const items = Array.from(host.querySelectorAll<HTMLButtonElement>('[role="menu"][aria-label="自動調整"] [role="menuitem"]'));
    expect(items.map((b) => b.textContent?.trim())).toEqual(['自動調整成內容大小', '自動調整成視窗大小', '固定欄寬']);
    items[1].click();
    await flush();
    expect(tables(await documentXml(ed))[0]).toContain('<w:tblW w:w="5000" w:type="pct"/>');

    // 分割表格.
    const split = byTitle('分割表格：從游標所在的列起分成兩個表格 (Ctrl+Shift+Enter)');
    expect(split.closest('[role="group"]')).toBe(merge);
    const wrapper = vi.spyOn(ed, 'splitTable');
    split.click();
    await flush();
    expect(wrapper).toHaveBeenCalledTimes(1);
    expect(tables(await documentXml(ed))).toHaveLength(2);
    app.unmount();
    host.remove();
  });
});
