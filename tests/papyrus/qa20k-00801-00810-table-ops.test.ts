import { describe, expect, it } from 'vitest';
import { beforeAll, describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { goToNextCell } from 'prosemirror-tables';
import JSZip from 'jszip';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { insertHardBreak, insertTab, tabNewRow } from '../../src/papyrus/editor/commands';
import { sortTable } from '../../src/papyrus/editor/tableCommands';
import { schema } from '../../src/papyrus/editor/schema';
import { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const row = (values: string[], header = false) => `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${values.map(cell).join('')}</w:tr>`;
const table = (...rows: string[]) => `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join('')}</w:tbl>`;

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc });
  const at = (text: string, end = false) => {
    let pos = -1;
    state.doc.descendants((n, p) => { if (pos < 0 && n.isText && n.text === text) pos = p; });
    if (pos < 0) throw new Error(`missing text: ${text}`);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + (end ? text.length : 0))));
  };
  const run = (cmd: Command) => {
    let dispatched = false;
    const accepted = cmd(state, (tr) => { state = state.apply(tr); dispatched = true; });
    expect(accepted).toBe(true);
    return dispatched;
  };
  const tableNode = () => {
    let found: ReturnType<typeof state.doc.nodeAt> = null;
    state.doc.descendants((n) => { if (!found && n.type === schema.nodes.table) found = n; });
    if (!found) throw new Error('expected a table in the document');
    return found;
  };
  const cellTexts = () => {
    const t = tableNode();
    const out: string[][] = [];
    t.forEach((r) => out.push(Array.from({ length: r.childCount }, (_, i) => r.child(i).textContent)));
    return out;
  };
  const saved = async () => JSZip.loadAsync(await writeDocx(state.doc, model));
  return { at, run, cellTexts, tableNode, saved, get state() { return state; } };
}

async function openEditor(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { model } = await readDocx(bytes);
  await editor.open(bytes);
  const pressTab = () => {
    const event = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    editor.view!.dom.dispatchEvent(event);
    return event;
  };
  const dispose = () => { editor.destroy(); host.remove(); };
  return { editor, model, pressTab, dispose };
}

function sorted(state: EditorState, column: number, descending = false, keepFirst = true) {
  let cellPos = -1;
  let cellIndex = 0;
  state.doc.descendants((n, p) => {
    if (n.type === schema.nodes.table_cell) {
      if (cellIndex === column && n.textContent) cellPos = p + 2;
      cellIndex++;
    }
  });
  if (cellPos < 0) throw new Error(`no non-empty sort key in column ${column}`);
  const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cellPos)));
  let next = selected;
  sortTable(descending, keepFirst)(selected, (tr) => { next = selected.apply(tr); });
  return next;
}

const rowsOf = (state: EditorState) => {
  let t: ReturnType<typeof state.doc.nodeAt> = null;
  state.doc.descendants((n) => { if (!t && n.type === schema.nodes.table) t = n; });
  if (!t) throw new Error('expected table');
  const out: string[][] = [];
  t.forEach((r) => out.push(Array.from({ length: r.childCount }, (_, i) => r.child(i).textContent)));
  return out;
};

describe('QA20K-00801–00810 table edit operations', () => {
  it('QA20K-00801 Tab in the final invoice cell creates a row and focuses its first cell', async () => {
    const d = await openEditor(table(row(['品項', '金額']), row(['茶水', '80'])));
    let amountPos = -1;
    d.editor.view!.state.doc.descendants((n, pos) => { if (amountPos < 0 && n.isText && n.text === '80') amountPos = pos; });
    if (amountPos < 0) throw new Error('missing invoice amount text node');
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(TextSelection.create(d.editor.view!.state.doc, amountPos + 2)));
    const event = d.pressTab();
    const doc = d.editor.view!.state.doc;
    let tablePos = -1;
    let tableNode: any;
    doc.descendants((n, pos) => { if (tablePos < 0 && n.type === schema.nodes.table) { tablePos = pos; tableNode = n; } });
    expect(event.defaultPrevented).toBe(true);
    expect(Array.from({ length: tableNode.childCount }, (_, r) => Array.from({ length: tableNode.child(r).childCount }, (_, c) => tableNode.child(r).child(c).textContent)))
      .toEqual([['品項', '金額'], ['茶水', '80'], ['', '']]);
    const firstCellOfNewRowPos = tablePos + 1 + tableNode.child(0).nodeSize + tableNode.child(1).nodeSize + 1;
    const selection = d.editor.view!.state.selection.$from;
    const cellDepth = [...Array(selection.depth + 1).keys()].reverse().find(depth => selection.node(depth).type === schema.nodes.table_cell);
    expect(cellDepth).toBeDefined();
    expect(selection.before(cellDepth!)).toBe(firstCellOfNewRowPos);
    expect(selection.parent.textContent).toBe('');
    d.dispose();
  });

  it('QA20K-00802 Tab from a non-final table cell advances to the adjacent cell without changing rows', async () => {
    const d = await open(table(row(['部門', '預算']), row(['業務', '1000'])));
    d.at('業務');
    d.run(goToNextCell(1));
    expect(d.cellTexts()).toHaveLength(2);
    expect(d.state.selection.$from.parent.textContent).toBe('1000');
    expect(d.state.selection.$from.node(d.state.selection.$from.depth - 1).type).toBe(schema.nodes.table_cell);
  });

  it('QA20K-00803 Tab outside a table inserts a tab and preserves the following table', async () => {
    const d = await openEditor(`<w:p><w:r><w:t>簽核人</w:t></w:r></w:p>${table(row(['項目', '狀態']))}`);
    let signerPos = -1;
    d.editor.view!.state.doc.descendants((n, pos) => { if (signerPos < 0 && n.isText && n.text === '簽核人') signerPos = pos; });
    if (signerPos < 0) throw new Error('missing signer text node');
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(TextSelection.create(d.editor.view!.state.doc, signerPos + '簽核人'.length)));
    const event = d.pressTab();
    const doc = d.editor.view!.state.doc;
    expect(event.defaultPrevented).toBe(true);
    expect(doc.firstChild!.textContent).toBe('簽核人');
    let insertedTab = false;
    doc.descendants((n) => { if (n.type === schema.nodes.tab) insertedTab = true; });
    expect(insertedTab).toBe(true);
    let followingTable: any;
    doc.descendants(n => { if (!followingTable && n.type === schema.nodes.table) followingTable = n; });
    expect(followingTable).toBeDefined();
    expect(followingTable.childCount).toBe(1);
    expect(followingTable.child(0).childCount).toBe(2);
    expect(Array.from({ length: followingTable.child(0).childCount }, (_, i) => followingTable.child(0).child(i).textContent)).toEqual(['項目', '狀態']);
    const xml = await JSZip.loadAsync(await writeDocx(doc, d.model)).then(z => z.file('word/document.xml')!.async('string'));
    expect(xml).toContain('<w:tab/>');
    d.dispose();
  });

  it('QA20K-00804 sorting by amount keeps two marked header rows pinned in their original order', async () => {
    const d = await open(table(row(['費用明細', '2026 年度'], true), row(['項目', '金額'], true), row(['餐費', '300']), row(['交通', '80'])));
    const s = sorted(d.state, 1);
    expect(rowsOf(s)).toEqual([['費用明細', '2026 年度'], ['項目', '金額'], ['交通', '80'], ['餐費', '300']]);
    expect(s.doc.firstChild!.child(0).attrs.header).toBe(true);
    expect(s.doc.firstChild!.child(1).attrs.header).toBe(true);
  });

  it('QA20K-00805 equal department names retain source order and keep each employee attached to its row', async () => {
    // Use two data columns in the actual table fixture: the sort key ties, companion values identify row order.
    const tie = await open(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${
      '<w:tr><w:trPr><w:tblHeader/></w:trPr>'+['部門','員工','座位'].map(cell).join('')+'</w:tr>'+
      ['研發|王小明|A1','研發|李小華|B2','行政|陳小姐|C3'].map(x=>row(x.split('|'))).join('')
    }</w:tbl>`);
    const s = sorted(tie.state, 0);
    expect(rowsOf(s)).toEqual([['部門','員工','座位'],['行政','陳小姐','C3'],['研發','王小明','A1'],['研發','李小華','B2']]);
  });

  it('QA20K-00806 an all-equal key column remains in the same row order after a sort request', async () => {
    const d = await open(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${row(['分類','備註'], true)}${row(['一般','先登錄'])}${row(['一般','後登錄'])}</w:tbl>`);
    const s = sorted(d.state, 0);
    expect(rowsOf(s)).toEqual([['分類','備註'],['一般','先登錄'],['一般','後登錄']]);
  });

  it('QA20K-00807 deleting the final line item keeps the preceding row and table width intact', async () => {
    const d = await open(table(row(['品項', '金額']), row(['筆記本', '60']), row(['文件夾', '40'])));
    d.at('文件夾');
    const { deleteRow } = await import('prosemirror-tables');
    d.run(deleteRow);
    expect(d.cellTexts()).toEqual([['品項', '金額'], ['筆記本', '60']]);
    const xml = await (await d.saved()).file('word/document.xml')!.async('string');
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect((tbl.match(/<w:gridCol /g) ?? [])).toHaveLength(2);
    expect(tbl).not.toContain('文件夾');
  });

  it('QA20K-00808 Tab from the last cell in a table row with a multi-paragraph cell still appends exactly one row', async () => {
    const d = await open(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('備註')}<w:tc><w:tcPr/><w:p><w:r><w:t>已確認</w:t></w:r></w:p><w:p><w:r><w:t>待歸檔</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
    d.at('待歸檔');
    d.run(tabNewRow);
    const tableNode = d.tableNode();
    expect(tableNode.childCount).toBe(2);
    expect(tableNode.child(0).child(1).textContent).toBe('已確認待歸檔');
    expect(tableNode.child(0).child(1).childCount).toBe(2);
    expect(tableNode.child(1).childCount).toBe(2);
    expect(d.state.selection.$from.parent.textContent).toBe('');
  });

  it('QA20K-00809 sorting rows by text preserves blank-key rows at the end while leaving them in source order', async () => {
    const d = await open(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${row(['單號','承辦'],true)}${row(['','王'])}${row(['B-02','李'])}${row(['   ','陳'])}${row(['A-01','張'])}</w:tbl>`);
    const s = sorted(d.state, 0);
    expect(rowsOf(s)).toEqual([['單號','承辦'],['A-01','張'],['B-02','李'],['','王'],['   ','陳']]);
  });

  it('QA20K-00810 Shift+Enter in a table cell inserts a line break without adding a table row', async () => {
    const d = await open(table(row(['日期', '備註']), row(['週一', '已收件待核對'])));
    d.at('已收件待核對', true);
    d.run(insertHardBreak);
    expect(d.cellTexts()).toHaveLength(2);
    const note = d.tableNode().child(1).child(1).firstChild!;
    expect(note.childCount).toBe(2);
    expect(note.child(0).textContent).toBe('已收件待核對');
    expect(note.child(1).type).toBe(schema.nodes.hard_break);
    const xml = await (await d.saved()).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:br/>');
  });
});
