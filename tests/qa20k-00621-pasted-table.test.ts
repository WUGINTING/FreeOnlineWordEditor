import { describe, expect, it } from 'vitest';
import { DOMParser as PMDOMParser, DOMSerializer } from 'prosemirror-model';
import { schema } from '../src/papyrus/editor/schema';

function parse(html: string) {
  const host = document.createElement('div');
  host.innerHTML = html;
  const doc = PMDOMParser.fromSchema(schema).parse(host);
  const output = document.createElement('div');
  output.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
  return { doc, output };
}

describe('QA20K office spreadsheet/table paste pilot 00621–00630', () => {
  it('QA20K-00621 keeps a merged heading cell across two pasted columns', () => {
    const { doc, output } = parse('<table><tbody><tr><th colspan="2"><p>季度營運摘要</p></th></tr><tr><td><p>營收</p></td><td><p>成本</p></td></tr></tbody></table>');
    const heading = doc.firstChild!.firstChild!.firstChild!;
    expect(heading.type.name).toBe('table_header');
    expect(heading.attrs.colspan).toBe(2);
    expect(output.textContent).toContain('季度營運摘要');
    expect(output.querySelector('th')?.getAttribute('colspan')).toBe('2');
  });

  it('QA20K-00622 keeps a vertically merged label spanning two pasted rows', () => {
    const { doc, output } = parse('<table><tbody><tr><td rowspan="2"><p>合約狀態</p></td><td><p>草擬</p></td></tr><tr><td><p>待簽核</p></td></tr></tbody></table>');
    const label = doc.firstChild!.firstChild!.firstChild!;
    expect(label.attrs.rowspan).toBe(2);
    expect(output.textContent).toBe('合約狀態草擬待簽核');
    expect(output.querySelector('td[rowspan="2"]')?.textContent).toBe('合約狀態');
  });

  it('QA20K-00623 retains empty cells as positions in a copied budget row', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>部門A</p></td><td><p></p></td><td><p>1200</p></td></tr></tbody></table>');
    const row = doc.firstChild!.firstChild!;
    expect(row.childCount).toBe(3);
    expect(row.child(1).textContent).toBe('');
    expect(output.querySelectorAll('tr td')).toHaveLength(3);
  });

  it('QA20K-00624 imports a pasted TH as a header cell, not an ordinary data cell', () => {
    const { doc, output } = parse('<table><thead><tr><th><p>日期</p></th><th><p>承辦人</p></th></tr></thead><tbody><tr><td><p>9/25</p></td><td><p>王小姐</p></td></tr></tbody></table>');
    const headerRow = doc.firstChild!.firstChild!;
    expect(headerRow.child(0).type.name).toBe('table_header');
    expect(headerRow.child(1).type.name).toBe('table_header');
    expect(output.querySelectorAll('th')).toHaveLength(2);
  });

  it('QA20K-00625 normalizes a legacy bgcolor attribute on a pasted cell', () => {
    const { doc, output } = parse('<table><tbody><tr><td bgcolor="#abc"><p>已付款</p></td></tr></tbody></table>');
    const cell = doc.firstChild!.firstChild!.firstChild!;
    expect(cell.attrs.background).toBe('#AABBCC');
    expect(output.querySelector('td')?.getAttribute('style')).toContain('background-color: rgb(170, 187, 204);');
  });

  it('QA20K-00626 converts legacy middle alignment to the editor-supported middle value', () => {
    const { doc, output } = parse('<table><tbody><tr><td valign="middle"><p>待補附件</p></td></tr></tbody></table>');
    expect(doc.firstChild!.firstChild!.firstChild!.attrs.vAlign).toBe('middle');
    expect(output.querySelector('td')?.getAttribute('style')).toContain('vertical-align: middle;');
  });

  it('QA20K-00627 preserves centered paragraph alignment inside a pasted approval table', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p style="text-align: center">核准</p></td></tr></tbody></table>');
    const paragraph = doc.firstChild!.firstChild!.firstChild!.firstChild!;
    expect(paragraph.attrs.align).toBe('center');
    expect(output.querySelector('td p')?.getAttribute('style')).toContain('text-align: center;');
  });

  it('QA20K-00628 preserves row-major reading order for a copied contact matrix', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>姓名</p></td><td><p>分機</p></td></tr><tr><td><p>陳怡君</p></td><td><p>208</p></td></tr></tbody></table>');
    expect(doc.textContent).toBe('姓名分機陳怡君208');
    expect(output.querySelectorAll('tr')).toHaveLength(2);
    expect(Array.from(output.querySelectorAll('tr')).map((row) => row.textContent)).toEqual(['姓名分機', '陳怡君208']);
  });

  it('QA20K-00629 keeps paragraph boundaries inside a pasted cell note', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>交付條件</p><p>完成雙方簽章</p></td></tr></tbody></table>');
    const cell = doc.firstChild!.firstChild!.firstChild!;
    expect(cell.childCount).toBe(2);
    expect(Array.from({ length: cell.childCount }, (_, i) => cell.child(i).textContent)).toEqual(['交付條件', '完成雙方簽章']);
    expect(output.querySelectorAll('td p')).toHaveLength(2);
  });

  it('QA20K-00630 retains explicit no-shading semantics for a white pasted cell', () => {
    const { doc, output } = parse('<table><tbody><tr><td style="background-color: rgb(255, 255, 255)"><p>一般項目</p></td></tr></tbody></table>');
    expect(doc.firstChild!.firstChild!.firstChild!.attrs.background).toBe('#FFFFFF');
    expect(output.querySelector('td')?.getAttribute('style')).toContain('background-color: rgb(255, 255, 255);');
  });
});
