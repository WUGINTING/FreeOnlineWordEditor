import { describe, expect, it } from 'vitest';
import { DOMParser as ProseMirrorDOMParser, DOMSerializer } from 'prosemirror-model';
import { schema } from '../../src/papyrus/editor/schema';

function parse(html: string) {
  const input = document.createElement('div');
  input.innerHTML = html;
  const doc = ProseMirrorDOMParser.fromSchema(schema).parse(input);
  const output = document.createElement('div');
  output.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
  return { doc, output };
}

function onlyCell(doc: ReturnType<typeof parse>['doc']) {
  return doc.firstChild!.firstChild!.firstChild!;
}

function text(doc: ReturnType<typeof parse>['doc'], value: string) {
  let found: any;
  doc.descendants(node => { if (!found && node.isText && node.text?.includes(value)) found = node; });
  if (!found) throw new Error(`Text not found: ${value}`);
  return found;
}

describe('QA20K HTML table paste inline workflows 01011–01020', () => {
  it('QA20K-01011 preserves a bold invoice identifier without bolding the cell note', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p><strong>Invoice 582</strong> approved for payment</p></td></tr></tbody></table>');
    const cell = onlyCell(doc);
    expect(cell.childCount).toBe(1);
    expect(text(doc, 'Invoice 582').marks.map(mark => mark.type.name)).toContain('bold');
    expect(text(doc, ' approved for payment').marks.map(mark => mark.type.name)).not.toContain('bold');
    expect(output.querySelector('td strong')?.textContent).toBe('Invoice 582');
    expect(output.querySelector('td')?.textContent).toBe('Invoice 582 approved for payment');
  });

  it('QA20K-01012 keeps a policy hyperlink as an active linked phrase inside a table cell', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>Read <a href="https://intranet.example.test/policies/travel">travel policy</a> before booking.</p></td></tr></tbody></table>');
    const link = text(doc, 'travel policy').marks.find(mark => mark.type.name === 'link');
    expect(link?.attrs.href).toBe('https://intranet.example.test/policies/travel');
    const anchor = output.querySelector('td a');
    expect(anchor?.getAttribute('href')).toBe('https://intranet.example.test/policies/travel');
    expect(anchor?.getAttribute('target')).toBe('_blank');
    expect(anchor?.getAttribute('rel')).toBe('noopener noreferrer');
    expect(output.querySelector('td')?.textContent).toBe('Read travel policy before booking.');
  });

  it('QA20K-01013 retains a postal address line break within one table-cell paragraph', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>88 Zhongxiao Rd.<br>Suite 240</p></td></tr></tbody></table>');
    const cell = onlyCell(doc);
    expect(cell.childCount).toBe(1);
    expect(cell.firstChild!.childCount).toBe(3);
    expect(cell.firstChild!.child(1).type.name).toBe('hard_break');
    expect(cell.textContent).toBe('88 Zhongxiao Rd.Suite 240');
    expect(output.querySelectorAll('td p')).toHaveLength(1);
    expect(output.querySelector('td br')).not.toBeNull();
  });

  it('QA20K-01014 marks a superseded delivery date as deleted without striking the replacement date', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>Old delivery <del>Friday</del>; revised delivery Monday.</p></td></tr></tbody></table>');
    expect(text(doc, 'Friday').marks.map(mark => mark.type.name)).toContain('strike');
    expect(text(doc, 'Monday').marks.map(mark => mark.type.name)).not.toContain('strike');
    expect(output.querySelector('td s')?.textContent).toBe('Friday');
    expect(output.querySelector('td')?.textContent).toBe('Old delivery Friday; revised delivery Monday.');
  });

  it('QA20K-01015 retains a nested approval matrix within the matching outer-cell paragraph', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>Approval matrix</p><table><tbody><tr><td><p>Finance</p></td><td><p>Review</p></td></tr><tr><td><p>Legal</p></td><td><p>Sign</p></td></tr></tbody></table></td><td><p>Due Friday</p></td></tr></tbody></table>');
    const outer = doc.firstChild!;
    expect(outer.childCount).toBe(1);
    expect(outer.firstChild!.childCount).toBe(2);
    const firstCell = outer.firstChild!.child(0);
    expect(firstCell.childCount).toBe(2);
    expect(firstCell.child(0).textContent).toBe('Approval matrix');
    const nested = firstCell.child(1);
    expect(nested.type.name).toBe('table');
    expect(nested.childCount).toBe(2);
    expect(nested.child(0).childCount).toBe(2);
    expect(nested.textContent).toBe('FinanceReviewLegalSign');
    expect(output.querySelectorAll('table')).toHaveLength(2);
    expect(output.querySelector('td td')?.textContent).toBe('Finance');
    expect(output.textContent).toBe('Approval matrixFinanceReviewLegalSignDue Friday');
  });

  it('QA20K-01016 keeps a superscript citation attached to the pasted cell amount', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>Tax rate 8%<sup>3</sup></p></td></tr></tbody></table>');
    expect(text(doc, '3').marks.map(mark => mark.type.name)).toContain('superscript');
    expect(onlyCell(doc).textContent).toBe('Tax rate 8%3');
    expect(output.querySelector('td sup')?.textContent).toBe('3');
    expect(output.querySelector('td')?.textContent).toBe('Tax rate 8%3');
  });

  it('QA20K-01017 preserves a subscript in a chemical inventory item', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>Carbon dioxide CO<sub>2</sub> cylinder</p></td></tr></tbody></table>');
    expect(text(doc, '2').marks.map(mark => mark.type.name)).toContain('subscript');
    expect(onlyCell(doc).textContent).toBe('Carbon dioxide CO2 cylinder');
    expect(output.querySelector('td sub')?.textContent).toBe('2');
    expect(output.querySelector('td')?.textContent).toBe('Carbon dioxide CO2 cylinder');
  });

  it('QA20K-01018 confines a red risk label color to its marked words inside the cell', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p><span style="color: rgb(192, 0, 0)">HIGH RISK</span> — supplier review required</p></td></tr></tbody></table>');
    const color = text(doc, 'HIGH RISK').marks.find(mark => mark.type.name === 'color');
    expect(color?.attrs.color).toBe('#c00000'); // pasted colours are kept as #rrggbb, like colours read from a file
    expect(text(doc, ' — supplier review required').marks.map(mark => mark.type.name)).not.toContain('color');
    expect(output.querySelector('td span')?.textContent).toBe('HIGH RISK');
    expect(output.querySelector('td span')?.getAttribute('style')).toContain('color: rgb(192, 0, 0)');
  });

  it('QA20K-01019 keeps a yellow deadline highlight separate from the cell explanation', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p><mark style="background-color: rgb(255, 255, 0)">Due today</mark> — submit the signed copy</p></td></tr></tbody></table>');
    const highlight = text(doc, 'Due today').marks.find(mark => mark.type.name === 'highlight');
    expect(highlight?.attrs.color).toBe('#ffff00'); // pasted colours are kept as #rrggbb, like colours read from a file
    expect(text(doc, ' — submit the signed copy').marks.map(mark => mark.type.name)).not.toContain('highlight');
    expect(output.querySelector('td mark')?.textContent).toBe('Due today');
    expect(output.querySelector('td mark')?.getAttribute('style')).toContain('background-color: rgb(255, 255, 0)');
  });

  it('QA20K-01020 retains Latin and East Asian font names on a bilingual cell phrase', () => {
    const { doc, output } = parse('<table><tbody><tr><td><p>請由 <span style="font-family: Arial, \'標楷體\'">Finance 財務</span> 確認。</p></td></tr></tbody></table>');
    const font = text(doc, 'Finance 財務').marks.find(mark => mark.type.name === 'font');
    expect(font?.attrs.family).toBe('Arial');
    expect(font?.attrs.eastAsia).toBe('標楷體');
    const span = output.querySelector('td span');
    expect(span?.textContent).toBe('Finance 財務');
    expect(span?.getAttribute('style')).toContain('font-family');
    expect(output.querySelector('td')?.textContent).toBe('請由 Finance 財務 確認。');
  });
});
