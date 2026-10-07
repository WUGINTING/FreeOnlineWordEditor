import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history, redo, undo } from 'prosemirror-history';
import { CellSelection, deleteTable, mergeCells, splitCell } from 'prosemirror-tables';
import JSZip from 'jszip';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { insertTable } from '../../src/papyrus/editor/commands';
import { schema } from '../../src/papyrus/editor/schema';
import { setCantSplit, setHeaderRow } from '../../src/papyrus/editor/tableCommands';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const cell = (text: string, tcPr = '') => `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
const row = (cells: string[], trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.map((value) => cell(value)).join('')}</w:tr>`;
const table = (rows: string[], cols = 3) => `<w:tbl><w:tblPr/><w:tblGrid>${Array.from({ length: cols }, () => '<w:gridCol w:w="2000"/>').join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
const base = () => table([
  row(['A1', 'B1', 'C1']), row(['A2', 'B2', 'C2']), row(['A3', 'B3', 'C3']),
]);

async function open(body = base(), withHistory = false) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: withHistory ? [history()] : [] });
  const cellPos = (text: string) => {
    let found = -1;
    state.doc.descendants((node, pos) => { if (found < 0 && node.type.name === 'table_cell' && node.textContent === text) found = pos; });
    if (found < 0) throw new Error(`No table cell with text: ${text}`);
    return found;
  };
  const at = (text: string) => {
    const pos = cellPos(text);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + 2)));
  };
  const cells = (from: string, to: string) => {
    state = state.apply(state.tr.setSelection(new CellSelection(state.doc.resolve(cellPos(from)), state.doc.resolve(cellPos(to)))));
  };
  const run = (cmd: Command) => {
    let next = state;
    const accepted = cmd(state, (tr) => { next = state.apply(tr); });
    if (accepted) state = next;
    return accepted;
  };
  const rows = () => {
    const t = state.doc.content.content.find((node) => node.type.name === 'table');
    return t ? t.content.content.map((r) => r.content.content.map((c) => c.textContent)) : [];
  };
  const save = async () => JSZip.loadAsync(await writeDocx(state.doc, model));
  return { at, cells, run, rows, save, get state() { return state; }, set state(next: EditorState) { state = next; } };
}

describe('QA20K table structure operations 00961–00970', () => {
  it('QA20K-00961 inserts a three-column table between existing document paragraphs', async () => {
    const d = await open('<w:p><w:r><w:t>Monthly summary</w:t></w:r></w:p><w:p><w:r><w:t>Insert here</w:t></w:r></w:p><w:p><w:r><w:t>Approval notes</w:t></w:r></w:p>');
    let pos = -1;
    d.state.doc.descendants((n, at) => { if (pos < 0 && n.isText && n.text === 'Insert here') pos = at; });
    d.state = d.state.apply(d.state.tr.setSelection(TextSelection.create(d.state.doc, pos)));
    expect(d.run(insertTable(2, 3, 600))).toBe(true);
    const blocks = d.state.doc.content.content.map((node) => node.type.name === 'table' ? 'table' : node.textContent);
    expect(blocks).toEqual(['Monthly summary', 'table', 'Insert here', 'Approval notes']);
    expect(d.rows()).toEqual([['', '', ''], ['', '', '']]);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    const tableXml = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect((tableXml.match(/<w:gridCol\b/g) ?? [])).toHaveLength(3);
    expect((tableXml.match(/<w:tr\b/g) ?? [])).toHaveLength(2);
    expect((tableXml.match(/<w:tc\b/g) ?? [])).toHaveLength(6);
  });

  it('QA20K-00962 merges adjacent expense headings into a two-column spanning cell', async () => {
    const d = await open(table([row(['項目', '部門', '金額']), row(['車資', '業務', '350'])]));
    d.cells('項目', '部門');
    expect(d.run(mergeCells)).toBe(true);
    expect(d.rows()).toEqual([['項目部門', '金額'], ['車資', '業務', '350']]);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect((xml.match(/<w:gridCol\b/g) ?? [])).toHaveLength(3);
  });

  it('QA20K-00963 splits a merged heading back into original grid cells without losing its text', async () => {
    const d = await open(table([row(['項目', '部門', '金額']), row(['車資', '業務', '350'])]));
    d.cells('項目', '部門'); d.run(mergeCells);
    d.at('項目部門');
    expect(d.run(splitCell)).toBe(true);
    expect(d.rows()).toEqual([['項目部門', '', '金額'], ['車資', '業務', '350']]);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    const top = [...xml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)][0][0];
    expect((top.match(/<w:tc\b/g) ?? [])).toHaveLength(3);
    expect(top).not.toContain('gridSpan');
  });

  it('QA20K-00964 merges a rectangular two-by-two approval block with intact table bounds', async () => {
    const d = await open(table([row(['簽核人', '職稱', '狀態']), row(['林', '主管', '待簽']), row(['陳', '財務', '待簽'])]));
    d.cells('林', '財務');
    expect(d.run(mergeCells)).toBe(true);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    const tableXml = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect((tableXml.match(/<w:gridCol\b/g) ?? [])).toHaveLength(3);
    expect((tableXml.match(/<w:tr\b/g) ?? [])).toHaveLength(3);
    expect(tableXml).toContain('<w:gridSpan w:val="2"/>');
    expect(tableXml).toContain('<w:vMerge w:val="restart"/>');
    expect(['林', '主管', '陳', '財務'].every((text) => tableXml.includes(`<w:t>${text}</w:t>`))).toBe(true);
  });

  it('QA20K-00965 deletes a selected obsolete table while preserving surrounding paragraphs', async () => {
    const d = await open('<w:p><w:r><w:t>Before table</w:t></w:r></w:p>' + base() + '<w:p><w:r><w:t>After table</w:t></w:r></w:p>');
    d.at('B2');
    expect(d.run(deleteTable)).toBe(true);
    expect(d.state.doc.content.content.map((node) => [node.type.name, node.textContent])).toEqual([
      ['paragraph', 'Before table'], ['paragraph', 'After table'],
    ]);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    expect(xml).not.toContain('<w:tbl>');
    expect(xml).toContain('Before table'); expect(xml).toContain('After table');
  });

  it('QA20K-00966 undoing an accidental merge restores each original amount cell', async () => {
    const d = await open(table([row(['項目', '金額', '備註']), row(['旅費', '1200', '高鐵']), row(['郵資', '80', '寄件'])]), true);
    d.cells('1200', '高鐵'); d.run(mergeCells);
    expect(d.rows()[1]).toEqual(['旅費', '1200高鐵']);
    expect(d.run(undo)).toBe(true);
    expect(d.rows()).toEqual([['項目', '金額', '備註'], ['旅費', '1200', '高鐵'], ['郵資', '80', '寄件']]);
    expect(d.run(redo)).toBe(true);
    expect(d.rows()[1]).toEqual(['旅費', '1200高鐵']);
  });

  it('QA20K-00967 marks multiple leading report rows as repeatable headers in OOXML', async () => {
    const d = await open(table([row(['Monthly costs', '', '']), row(['品項', '部門', '金額']), row(['車資', '業務', '350'])]));
    d.at('部門'); d.run(setHeaderRow(true));
    expect(d.state.doc.firstChild!.content.content.map((r) => r.attrs.header)).toEqual([true, true, false]);
    const zip = await d.save();
    const xml = await zip.file('word/document.xml')!.async('string');
    const rows = [...xml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
    expect(rows).toHaveLength(3);
    expect(rows.map((r) => r.includes('<w:tblHeader'))).toEqual([true, true, false]);
    const reopened = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(reopened.doc.firstChild!.content.content.map((r) => r.attrs.header)).toEqual([true, true, false]);
  });

  it('QA20K-00968 prevents one confidential note row from splitting in the saved structure', async () => {
    const d = await open(base()); d.at('B2'); d.run(setCantSplit(true));
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    const rows = [...xml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
    expect(rows).toHaveLength(3);
    expect(rows.map((r) => r.includes('<w:cantSplit'))).toEqual([false, true, false]);
    expect(rows.map((r) => [...r.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((m) => m[1]))).toEqual([
      ['A1', 'B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', 'B3', 'C3'],
    ]);
  });

  it('QA20K-00969 splits a two-row vertically merged department cell back into two row cells', async () => {
    const merged = table([
      `<w:tr>${cell('Operations', '<w:vMerge w:val="restart"/>')}${cell('Request 1')}${cell('Open')}</w:tr>`,
      `<w:tr>${cell('', '<w:vMerge/>')}${cell('Request 2')}${cell('Open')}</w:tr>`,
    ]);
    const d = await open(merged); d.at('Operations');
    expect(d.run(splitCell)).toBe(true);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    const tableXml = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    const rows = [...tableXml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((m) => m[0]);
    expect(rows).toHaveLength(2);
    expect(rows.map((r) => [...r.matchAll(/<w:tc\b/g)].length)).toEqual([3, 3]);
    expect((tableXml.match(/<w:gridCol\b/g) ?? [])).toHaveLength(3);
    expect(tableXml).not.toContain('<w:vMerge');
    const rowTexts = rows.map((r) => [...r.matchAll(/<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g)].map((m) => [...m[1].matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((t) => t[1]).join('')));
    expect(rowTexts).toEqual([['Operations', 'Request 1', 'Open'], ['', 'Request 2', 'Open']]);
  });

  it('QA20K-00970 undo removes an accidental table insertion while keeping nearby text', async () => {
    const source = '<w:p><w:r><w:t>Agenda</w:t></w:r></w:p><w:p><w:r><w:t>Insert table here</w:t></w:r></w:p><w:p><w:r><w:t>Next steps</w:t></w:r></w:p>';
    const d = await open(source, true);
    let pos = -1;
    d.state.doc.descendants((n, at) => { if (pos < 0 && n.isText && n.text === 'Insert table here') pos = at; });
    d.state = d.state.apply(d.state.tr.setSelection(TextSelection.create(d.state.doc, pos)));
    d.run(insertTable(2, 2, 400));
    expect(d.rows()).toEqual([['', ''], ['', '']]);
    d.run(undo);
    expect(d.rows()).toEqual([]);
    expect(d.state.doc.content.content.map((node) => node.textContent)).toEqual(['Agenda', 'Insert table here', 'Next steps']);
    d.run(redo);
    expect(d.rows()).toEqual([['', ''], ['', '']]);
    const xml = await (await d.save()).file('word/document.xml')!.async('string');
    expect((xml.match(/<w:tbl\b/g) ?? [])).toHaveLength(1);
    expect(xml).toContain('Insert table here'); expect(xml).toContain('Next steps');
  });
});
