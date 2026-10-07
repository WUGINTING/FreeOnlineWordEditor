// Sorting table rows by a column (editor/tableCommands.ts sortTable).
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { blankPackage } from '../../src/papyrus/docx/template';
import { compareCellText, sortRefusal, sortTable } from '../../src/papyrus/editor/tableCommands';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const cell = (t: string, span = '') => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${span}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const row = (cells: string[], header = false) => `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells.map((c) => cell(c)).join('')}</w:tr>`;
const table = (rows: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl><w:p/>`;

async function stateOf(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  const { doc } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  return EditorState.create({ doc });
}
/** Cursor in the cell with this text. */
function at(state: EditorState, text: string): EditorState {
  let pos = -1;
  state.doc.descendants((n, p) => {
    if (pos < 0 && n.isText && n.text === text) pos = p;
  });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
}
function run(state: EditorState, cmd: ReturnType<typeof sortTable>): EditorState {
  let next = state;
  cmd(state, (tr) => (next = state.apply(tr)));
  return next;
}
const column = (state: EditorState, i: number) => {
  const out: string[] = [];
  state.doc.firstChild!.forEach((r) => out.push(r.child(i).textContent));
  return out;
};

describe('sorting table rows', () => {
  it('numbers by value (thousands, currency, percent), header row kept on top', async () => {
    const s = at(await stateOf(table(row(['品項', '金額'], true) + row(['甲', '1,200']) + row(['乙', '$300']) + row(['丙', '45']) + row(['丁', '']))), '$300');
    const up = run(s, sortTable(false));
    expect(column(up, 1)).toEqual(['金額', '45', '$300', '1,200', '']);
    const down = run(s, sortTable(true));
    expect(column(down, 1)).toEqual(['金額', '1,200', '$300', '45', '']);
    // Whole rows move together.
    expect(column(up, 0)).toEqual(['品項', '丙', '乙', '甲', '丁']);
  });

  it('dates (西元 and 民國) by time, and Chinese text in Traditional Chinese order', async () => {
    expect(compareCellText('115年1月2日', '2025/12/31')).toBeGreaterThan(0);
    expect(compareCellText('2026-09-24', '2026/9/3')).toBeGreaterThan(0);
    // Traditional Chinese order puts fewer strokes first: 一 (1) before 十 (2) before 王 (4).
    const s = at(await stateOf(table(row(['王']) .replace('</w:tr>', `${cell('x')}</w:tr>`) + row(['十', 'y']) + row(['一', 'z']))), '王');
    const sorted = run(s, sortTable(false, false));
    expect(column(sorted, 0)).toEqual(['一', '十', '王']);
  });

  it('keeps an unmarked first row in place when asked, and sorts it when not', async () => {
    // (In Traditional Chinese order, Chinese text comes before Latin letters: the header here is Latin.)
    const s = at(await stateOf(table(row(['Zone', 'n']) + row(['b', '2']) + row(['a', '1']))), 'b');
    expect(column(run(s, sortTable(false, true)), 0)).toEqual(['Zone', 'a', 'b']);
    expect(column(run(s, sortTable(false, false)), 0)).toEqual(['a', 'b', 'Zone']);
  });

  it('refuses a table with merged cells, with the reason', async () => {
    const merged = `<w:tr>${cell('合併', '<w:gridSpan w:val="2"/>')}</w:tr>` + row(['b', '2']) + row(['a', '1']);
    const s = at(await stateOf(table(merged)), 'b');
    expect(sortRefusal(s)).toBe('表格有合併儲存格，無法排序。');
    expect(sortTable(false)(s, () => {})).toBe(false);
  });
});
