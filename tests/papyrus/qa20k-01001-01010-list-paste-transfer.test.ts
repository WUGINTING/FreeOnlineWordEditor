import { describe, expect, it } from 'vitest';
import { DOMParser as PMDOMParser, Fragment, type Node as PMNode, Slice } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { schema } from '../../src/papyrus/editor/schema';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';
import { emptyNumbering } from '../../src/papyrus/docx/model';
import { ListCounter } from '../../src/papyrus/docx/numbering';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const wordItem = (list: string, level: number, lfo: number, marker: string, html: string) =>
  `<p class="MsoListParagraph" style="mso-list:${list} level${level} lfo${lfo};margin-left:${18 * level}pt;text-indent:-18pt">` +
  `<!--[if !supportLists]--><span style="mso-list:Ignore">${marker}<span style="font:7pt 'Times New Roman'">&nbsp;&nbsp;</span></span><!--[endif]-->${html}</p>`;

function parse(html: string) {
  const container = document.createElement('div');
  container.innerHTML = `<html><body><!--StartFragment-->${html}<!--EndFragment--></body></html>`;
  return PMDOMParser.fromSchema(schema).parseSlice(container);
}

function paragraphs(slice: Slice): PMNode[] {
  const out: PMNode[] = [];
  slice.content.forEach((node) => out.push(node));
  return out;
}

function numberingFormat(numbering: ReturnType<typeof emptyNumbering>, para: PMNode) {
  if (!para.attrs.numId) return null;
  return numbering.abstracts[numbering.nums[para.attrs.numId].abstractId].levels[para.attrs.ilvl].fmt;
}

describe('list paste and transfer workflows', () => {
  it('QA20K-01001 remaps a copied list whose source numId does not exist in the receiving document', () => {
    const numbering = emptyNumbering();
    const sourcePara = schema.nodes.paragraph.create({ numId: 'copy:77', ilvl: 0 }, schema.text('Approval item'));
    const slice = adoptPastedLists(new Slice(Fragment.from(sourcePara), 0, 0), numbering);
    const pasted = paragraphs(slice)[0];
    expect(pasted.attrs.numId).not.toBe('77');
    expect(pasted.attrs.numId).toBeTruthy();
    expect(numberingFormat(numbering, pasted)).toBe('decimal');
    const counter = new ListCounter(numbering);
    expect(counter.next(pasted.attrs.numId, pasted.attrs.ilvl)).toBe('1.');
  });

  it('QA20K-01002 keeps two Word list instances apart when they share one abstract list but have separate lfo IDs', () => {
    const numbering = emptyNumbering();
    const slice = adoptPastedLists(parse(wordItem('l12', 1, 1, '1.', 'Scope') + wordItem('l12', 1, 2, '1.', 'Deliverables')), numbering);
    const [first, second] = paragraphs(slice);
    expect(first.textContent).toBe('Scope');
    expect(second.textContent).toBe('Deliverables');
    expect(first.attrs.numId).not.toBe(second.attrs.numId);
    expect(first.attrs.ilvl).toBe(0);
    expect(second.attrs.ilvl).toBe(0);
    const counter = new ListCounter(numbering);
    expect([counter.next(first.attrs.numId, 0), counter.next(second.attrs.numId, 0)]).toEqual(['1.', '1.']);
  });

  it('QA20K-01003 treats a second paste of the same Word checklist as a new numbering instance', () => {
    const html = wordItem('l23', 1, 1, '1.', 'Check totals') + wordItem('l23', 1, 1, '2.', 'Confirm tax');
    const numbering = emptyNumbering();
    const first = paragraphs(adoptPastedLists(parse(html), numbering));
    const second = paragraphs(adoptPastedLists(parse(html), numbering));
    expect(first.map((node) => node.textContent)).toEqual(['Check totals', 'Confirm tax']);
    expect(second.map((node) => node.textContent)).toEqual(['Check totals', 'Confirm tax']);
    expect(new Set([...first, ...second].map((node) => node.attrs.numId)).size).toBe(2);
    const counter = new ListCounter(numbering);
    expect([counter.next(first[0].attrs.numId, 0), counter.next(first[1].attrs.numId, 0)]).toEqual(['1.', '2.']);
    expect([counter.next(second[0].attrs.numId, 0), counter.next(second[1].attrs.numId, 0)]).toEqual(['1.', '2.']);
  });

  it('QA20K-01004 restarts nested letter steps under each new Word parent item', () => {
    const numbering = emptyNumbering();
    const html = wordItem('l34', 1, 1, '1.', 'Collect quotes') +
      wordItem('l34', 2, 1, 'a.', 'Supplier A') + wordItem('l34', 2, 1, 'b.', 'Supplier B') +
      wordItem('l34', 1, 1, '2.', 'Approve vendor') +
      wordItem('l34', 2, 1, 'a.', 'Check contract');
    const paras = paragraphs(adoptPastedLists(parse(html), numbering));
    const counter = new ListCounter(numbering);
    const labels = paras.map((node) => counter.next(node.attrs.numId, node.attrs.ilvl));
    expect(paras.map((node) => node.attrs.ilvl)).toEqual([0, 1, 1, 0, 1]);
    expect(labels).toEqual(['1.', 'a.', 'b.', '2.', 'a.']);
  });

  it('QA20K-01005 preserves a Roman-numeral web checklist start when pasted into the document', () => {
    const numbering = emptyNumbering();
    const paras = paragraphs(adoptPastedLists(parse('<ol type="I" start="4"><li>Review request</li><li>Approve response</li></ol>'), numbering));
    const id = paras[0].attrs.numId;
    expect(paras.map((node) => node.attrs.numId)).toEqual([id, id]);
    expect(numberingFormat(numbering, paras[0])).toBe('upperRoman');
    const counter = new ListCounter(numbering);
    expect(paras.map((node) => counter.next(id, node.attrs.ilvl))).toEqual(['IV.', 'V.']);
  });

  it('QA20K-01006 keeps a nested unordered note subordinate to its numbered action list', () => {
    const numbering = emptyNumbering();
    const paras = paragraphs(adoptPastedLists(parse('<ol><li>Prepare packet<ul><li>Include signed form</li><li>Attach receipt</li></ul></li><li>Send to finance</li></ol>'), numbering));
    expect(paras.map((node) => [node.textContent, node.attrs.ilvl, numberingFormat(numbering, node)])).toEqual([
      ['Prepare packet', 0, 'decimal'], ['Include signed form', 1, 'bullet'], ['Attach receipt', 1, 'bullet'], ['Send to finance', 0, 'decimal'],
    ]);
    expect(paras[0].attrs.numId).toBe(paras[1].attrs.numId);
    expect(paras[0].attrs.numId).toBe(paras[3].attrs.numId);
  });

  it('QA20K-01007 leaves a continuation paragraph inside a list item unnumbered', () => {
    const numbering = emptyNumbering();
    const paras = paragraphs(adoptPastedLists(parse('<ol><li><p>Required action</p><p>Explain the exception in a full sentence.</p></li><li><p>Next action</p></li></ol>'), numbering));
    expect(paras.map((node) => node.textContent)).toEqual(['Required action', 'Explain the exception in a full sentence.', 'Next action']);
    expect(paras[0].attrs.numId).toBeTruthy();
    expect(paras[1].attrs.numId).toBeNull();
    expect(paras[2].attrs.numId).toBe(paras[0].attrs.numId);
    expect(numbering.nums).toHaveProperty(paras[0].attrs.numId);
  });

  it('QA20K-01008 keeps inline emphasis and links attached to their pasted numbered step', () => {
    const numbering = emptyNumbering();
    const paras = paragraphs(adoptPastedLists(parse('<ol start="2"><li><p><strong>Verify</strong> the amount in <a href="https://example.test/invoice">the invoice</a></p></li></ol>'), numbering));
    const step = paras[0];
    const bold = step.firstChild!.marks.find((mark) => mark.type.name === 'bold');
    const linkText = step.child(2);
    expect(step.textContent).toBe('Verify the amount in the invoice');
    expect(bold).toBeDefined();
    expect(linkText.marks.find((mark) => mark.type.name === 'link')?.attrs.href).toBe('https://example.test/invoice');
    expect(new ListCounter(numbering).next(step.attrs.numId, step.attrs.ilvl)).toBe('2.');
  });

  it('QA20K-01009 strips the Word marker glyph but keeps its following explanation outside list numbering', () => {
    const numbering = emptyNumbering();
    const html = wordItem('l45', 1, 1, '3.', 'Retain invoice') + '<p class="MsoNormal">Manager approval is required.</p>' + wordItem('l45', 1, 1, '4.', 'Archive receipt');
    const paras = paragraphs(adoptPastedLists(parse(html), numbering));
    expect(paras.map((node) => node.textContent)).toEqual(['Retain invoice', 'Manager approval is required.', 'Archive receipt']);
    expect(paras.map((node) => !!node.attrs.numId)).toEqual([true, false, true]);
    expect(paras[0].textContent).not.toContain('3.');
    expect(paras[2].textContent).not.toContain('4.');
    expect(paras[0].attrs.numId).toBe(paras[2].attrs.numId);
  });

  it('QA20K-01010 pastes an imported checklist into one table cell without numbering its neighbor cell', () => {
    const numbering = emptyNumbering();
    const cellA = schema.nodes.table_cell.createAndFill(null)!;
    const cellB = schema.nodes.table_cell.createAndFill(null)!;
    const row = schema.nodes.table_row.createAndFill(null, [cellA, cellB])!;
    const table = schema.nodes.table.createAndFill(null, [row])!;
    const doc = schema.nodes.doc.create(null, [table]);
    let state = EditorState.create({ schema, doc });
    let firstEmptyParagraph = -1;
    state.doc.descendants((node, pos) => {
      if (firstEmptyParagraph < 0 && node.type === schema.nodes.paragraph && !node.textContent) firstEmptyParagraph = pos + 1;
    });
    expect(firstEmptyParagraph).toBeGreaterThan(0);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, firstEmptyParagraph)));
    const slice = adoptPastedLists(parse('<ol start="2"><li>Verify invoice</li><li>Record approval</li></ol>'), numbering);
    state = state.apply(state.tr.replaceSelection(slice));
    const cells: PMNode[] = [];
    state.doc.descendants((node) => { if (node.type === schema.nodes.table_cell) cells.push(node); });
    expect(cells).toHaveLength(2);
    expect(cells[0].textContent).toBe('Verify invoiceRecord approval');
    expect(cells[1].textContent).toBe('');
    const leftItems: PMNode[] = [];
    const rightItems: PMNode[] = [];
    cells[0].descendants((node) => { if (node.type === schema.nodes.paragraph && node.attrs.numId) leftItems.push(node); });
    cells[1].descendants((node) => { if (node.type === schema.nodes.paragraph && node.attrs.numId) rightItems.push(node); });
    expect(leftItems.map((node) => node.textContent)).toEqual(['Verify invoice', 'Record approval']);
    expect(rightItems).toEqual([]);
    const counter = new ListCounter(numbering);
    expect(leftItems.map((node) => counter.next(node.attrs.numId, node.attrs.ilvl))).toEqual(['2.', '3.']);
  });
});
