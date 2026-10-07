// Table appearance: each command patches only its property in the original w:tcPr / w:trPr,
// in ECMA-376 order; untouched tables are written back unchanged.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { parseXml } from '../../src/papyrus/docx/xml';
import {
  distributeColumns, setCantSplit, setCellBackground, setCellBorders, setCellVAlign, setHeaderRow, setRowHeight, tableInfo,
} from '../../src/papyrus/editor/tableCommands';
import { compareElements } from './canonical';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const TCPR_A1 = '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="pct10" w:color="FF0000" w:fill="D9D9D9"/><w:noWrap/></w:tcPr>';
const cell = (text: string, tcPr = '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>') => `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr w:rsidR="00A10001"><w:trPr><w:cnfStyle w:val="100000000000"/><w:jc w:val="center"/></w:trPr>${cell('A1', TCPR_A1)}${cell('B1')}${cell('C1')}</w:tr>` +
  `<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${cell('A2', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>')}${cell('B2')}${cell('C2')}</w:tr>` +
  `<w:tr>${cell('', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/><w:tcMar><w:left w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>')}${cell('B3')}${cell('C3')}</w:tr>` +
  '</w:tbl>';

async function open(body = TABLE + '<w:p/>') {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { doc, model } = await readDocx(bytes);
  let state = EditorState.create({ schema, doc });
  const run = (cmd: Command) => expect(cmd(state, (tr) => (state = state.apply(tr)))).toBe(true);
  /** Put the cursor in the cell with this text. */
  const at = (text: string) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, find(state, text)))));
  const cells = (from: string, to: string) => {
    const $a = state.doc.resolve(find(state, from) - 2);
    const $b = state.doc.resolve(find(state, to) - 2);
    state = state.apply(state.tr.setSelection(new CellSelection($a, $b)));
  };
  const save = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  return { bytes, run, at, cells, save, get state() { return state; } };
}

/** Position inside the paragraph of the cell whose text is `text`. */
function find(state: EditorState, text: string): number {
  let found = -1;
  state.doc.descendants((n, pos) => {
    if (found < 0 && n.type.name === 'table_cell' && n.textContent === text) found = pos + 2;
    return found < 0;
  });
  if (found < 0) throw new Error('no cell ' + text);
  return found;
}

const tcOf = (xml: string, text: string) => {
  const m = new RegExp(`<w:tc>(?:(?!<w:tc>).)*?<w:t>${text}</w:t>`, 's').exec(xml);
  return m ? m[0] : '';
};

describe('table appearance', () => {
  it('saving an untouched table keeps every element', async () => {
    const d = await open();
    const body = async (b: Uint8Array) => parseXml(await (await JSZip.loadAsync(b)).file('word/document.xml')!.async('string')).documentElement;
    const { doc, model } = await readDocx(d.bytes);
    expect(compareElements(await body(d.bytes), await body(await writeDocx(doc, model)))).toEqual([]);
  });

  it('reads row and cell properties for the panel', async () => {
    const d = await open();
    d.at('A2');
    expect(tableInfo(d.state)).toMatchObject({ height: 400, heightRule: null, header: false, cantSplit: false, vAlign: null });
    d.at('A1');
    expect(tableInfo(d.state)!.background).toBe('#D9D9D9');
  });

  it('QA20K-00711 shading replaces only the fill of w:shd for the selected invoice cell', async () => {
    const d = await open();
    d.at('A1');
    d.run(setCellBackground('#ffff00'));
    const xml = await d.save();
    expect(tcOf(xml, 'A1')).toContain('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:noWrap/></w:tcPr>');
    expect(tcOf(xml, 'B1')).toContain('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>');
    d.run(setCellBackground(null));
    expect(tcOf(await d.save(), 'A1')).toContain('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:noWrap/></w:tcPr>');
  });

  it('QA20K-00712 applying all borders synchronizes the selected cell shared edges', async () => {
    const d = await open();
    d.at('B2');
    d.run(setCellBorders('all'));
    const xml = await d.save();
    const single = (s: string) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    expect(tcOf(xml, 'B2')).toContain(`<w:tcBorders>${single('top')}${single('left')}${single('bottom')}${single('right')}</w:tcBorders>`);
    expect(tcOf(xml, 'B1')).toContain(`<w:tcBorders>${single('bottom')}</w:tcBorders>`);
    expect(tcOf(xml, 'C2')).toContain(`<w:tcBorders>${single('left')}</w:tcBorders>`);
    expect(tcOf(xml, 'B3')).toContain(`<w:tcBorders>${single('top')}</w:tcBorders>`);
    // The merged cell A2 (rows 2-3): its right edge is set in both pieces.
    expect(tcOf(xml, 'A2')).toContain(`<w:vMerge w:val="restart"/><w:tcBorders>${single('right')}</w:tcBorders></w:tcPr>`);
    expect(xml).toContain(`<w:vMerge/><w:tcBorders>${single('right')}</w:tcBorders><w:tcMar>`);
    // Unrelated cells untouched.
    expect(tcOf(xml, 'C1')).toContain('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>');

    d.run(setCellBorders('none'));
    expect(tcOf(await d.save(), 'B2')).toContain('<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>');
  });

  it('QA20K-00713 inner and outer borders follow a selected rectangular block', async () => {
    const d = await open();
    d.cells('B1', 'C2');
    d.run(setCellBorders('inner'));
    let xml = await d.save();
    expect(tcOf(xml, 'B1')).toContain('<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>');
    expect(tcOf(xml, 'A1')).not.toContain('tcBorders');
    d.cells('B1', 'C2');
    d.run(setCellBorders('outer'));
    xml = await d.save();
    expect(tcOf(xml, 'A1')).toContain('<w:tcBorders><w:right w:val="single"');
    expect(tcOf(xml, 'C2')).toMatch(/<w:tcBorders><w:top [^>]*\/><w:left [^>]*\/><w:bottom [^>]*\/><w:right [^>]*\/><\/w:tcBorders>/);
  });

  it('QA20K-00714 vertical alignment is written to each row piece of a merged cell', async () => {
    const d = await open();
    d.at('A2');
    d.run(setCellVAlign('middle'));
    d.run(setCellBackground('#00ff00'));
    const xml = await d.save();
    expect(tcOf(xml, 'A2')).toContain('<w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/><w:vAlign w:val="center"/></w:tcPr>');
    expect(xml).toContain('<w:vMerge/><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/><w:tcMar><w:left w:w="0" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>');
  });

  it('row height, header row and "don\'t split" patch w:trPr in order and keep the rest', async () => {
    const d = await open();
    d.at('B1');
    d.run(setRowHeight(567, 'exact'));
    d.run(setCantSplit(true));
    d.run(setHeaderRow(true));
    let xml = await d.save();
    expect(xml).toContain('<w:tr w:rsidR="00A10001"><w:trPr><w:cnfStyle w:val="100000000000"/><w:cantSplit/><w:trHeight w:val="567" w:hRule="exact"/><w:tblHeader/><w:jc w:val="center"/></w:trPr>');
    // An omitted rule means "at least": changing only the value keeps it omitted.
    d.at('B2');
    d.run(setRowHeight(600, 'atLeast'));
    xml = await d.save();
    expect(xml).toContain('<w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>');
    d.run(setRowHeight(null));
    xml = await d.save();
    expect(xml).toContain('<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>');
  });

  it('QA20K-00715 keeps only the selected shipment row together at its exact height', async () => {
    const d = await open();
    d.at('B2');
    d.run(setCantSplit(true));
    d.run(setRowHeight(720, 'exact'));
    const xml = await d.save();
    const row = /<w:tr>(?:(?!<w:tr>).)*?<w:t>B2<\/w:t>(?:(?!<w:tr>).)*?<\/w:tr>/s.exec(xml)?.[0] ?? '';
    expect(row).toContain('<w:cantSplit/>');
    expect(row).toContain('<w:trHeight w:val="720" w:hRule="exact"/>');
    expect(row).toContain('<w:t>B2</w:t>');
    const nextRow = /<w:tr>(?:(?!<w:tr>).)*?<w:t>B3<\/w:t>(?:(?!<w:tr>).)*?<\/w:tr>/s.exec(xml)?.[0] ?? '';
    expect(nextRow).not.toContain('<w:cantSplit/>');
    expect(nextRow).not.toContain('<w:trHeight w:val="720"');
  });

  it('QA20K-00716 setting a repeated header through row two marks all preceding rows only', async () => {
    const d = await open();
    d.at('B2');
    d.run(setHeaderRow(true));
    const rows = () => d.state.doc.child(0).content.content.map((r) => r.attrs.header);
    expect(rows()).toEqual([true, true, false]);
    d.at('B1');
    d.run(setHeaderRow(false));
    expect(rows()).toEqual([false, false, false]);
  });

  it('QA20K-00717 distributing unevenly sized columns gives every cell equal widths and preserves total width', async () => {
    const d = await open(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="4500"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        `<w:tr>${cell('a', '<w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>')}${cell('b', '<w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>')}${cell('c')}</w:tr></w:tbl><w:p/>`,
    );
    d.at('a');
    d.run(distributeColumns);
    const xml = await d.save();
    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>');
    expect([...xml.matchAll(/<w:tcW w:w="(\d+)" w:type="dxa"\/>/g)].map(m => Number(m[1]))).toEqual([3000,3000,3000]);
    expect([...xml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map(m => Number(m[1])).reduce((sum,width)=>sum+width,0)).toBe(9000);
  });
});
