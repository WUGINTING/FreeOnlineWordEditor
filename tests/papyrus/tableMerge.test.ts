// Writer regressions around tables and pictures: vertically merged cells whose column span
// changed, relationships used only inside kept XML (merged pieces, tcPr ...), legacy w:hMerge,
// cells added to square up ragged rows, and drawing ids of new pictures.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { CellSelection, addColumnAfter, mergeCells } from 'prosemirror-tables';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { parseXml } from '../../src/papyrus/docx/xml';
import { schema } from '../../src/papyrus/editor/schema';
import { setCellBackground, setCellBorders, setCellVAlign } from '../../src/papyrus/editor/tableCommands';
import { compareElements } from './canonical';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const tc = (text: string, tcPr = '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>', inner = '') =>
  `<w:tc>${tcPr}<w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}${inner}</w:p></w:tc>`;
const table = (cols: number, rows: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(cols)}</w:tblGrid>` +
  rows.map((r) => `<w:tr>${r}</w:tr>`).join('') +
  '</w:tbl><w:p/>';

async function open(body: string, rels = '', media: Record<string, string> = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${SECT}</w:body></w:document>`);
  for (const [path, b64] of Object.entries(media)) zip.file(path, b64, { base64: true });
  if (rels) {
    const old = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', old.replace('</Relationships>', rels + '</Relationships>'));
  }
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { doc, model } = await readDocx(bytes);
  let state = EditorState.create({ schema, doc });
  const run = (cmd: Command) => expect(cmd(state, (tr) => (state = state.apply(tr)))).toBe(true);
  const cellPos = (text: string) => {
    let found = -1;
    state.doc.descendants((n, pos) => {
      if (found < 0 && n.type.name === 'table_cell' && n.textContent === text) found = pos;
      return found < 0;
    });
    if (found < 0) throw new Error('no cell ' + text);
    return found;
  };
  const at = (text: string) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cellPos(text) + 2))));
  const cells = (from: string, to: string) =>
    (state = state.apply(state.tr.setSelection(new CellSelection(state.doc.resolve(cellPos(from)), state.doc.resolve(cellPos(to))))));
  const save = async () => JSZip.loadAsync(await writeDocx(state.doc, model));
  const xmlOf = async (z?: JSZip) => (z ?? (await save())).file('word/document.xml')!.async('string');
  return { bytes, run, at, cells, save, xmlOf, get state() { return state; } };
}

/** Grid columns each w:tr covers (gridBefore + spans + gridAfter). */
function rowWidths(xml: string): number[] {
  const doc = parseXml(xml);
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return Array.from(doc.getElementsByTagNameNS(W, 'tr')).map((tr) =>
    Array.from(tr.children)
      .filter((c) => c.localName === 'tc')
      .reduce((s, c) => s + Number(c.getElementsByTagNameNS(W, 'gridSpan')[0]?.getAttributeNS(W, 'val') ?? 1), 0),
  );
}

describe('vertically merged cells whose column span changed', () => {
  const MERGED = table(2, [
    tc('M', '<w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>'),
    tc('', '<w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge/><w:tcMar><w:left w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>'),
    tc('A3') + tc('B3'),
  ]);

  it('QA20K-00718 inserting a column through a vertical merge widens both merge pieces', async () => {
    const d = await open(MERGED);
    d.at('A3');
    d.run(addColumnAfter);
    const xml = await d.xmlOf();
    expect(rowWidths(xml)).toEqual([3, 3, 3]);
    expect(xml).toMatch(/<w:gridSpan w:val="3"\/><w:vMerge w:val="restart"\/>/);
    expect(xml).toMatch(/<w:gridSpan w:val="3"\/><w:vMerge\/>/);
  });

  it('QA20K-00719 merging a vertical span with its neighbor preserves row widths and merge continuation', async () => {
    const d = await open(table(2, [
      tc('A1', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>') + tc('B1'),
      tc('', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr>') + tc('B2'),
    ]));
    d.cells('A1', 'B2');
    d.run(mergeCells);
    const xml = await d.xmlOf();
    expect(rowWidths(xml)).toEqual([2, 2]);
    expect(xml).toMatch(/<w:tcW w:w="6000" w:type="dxa"\/><w:gridSpan w:val="2"\/><w:vMerge\/>/);
  });

  it('pieces of an unchanged merged cell are still written as they were', async () => {
    const d = await open(MERGED);
    d.at('M');
    d.run(setCellVAlign('bottom'));
    const xml = await d.xmlOf();
    expect(xml).toContain('<w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge/><w:tcMar><w:left w:w="0" w:type="dxa"/></w:tcMar><w:vAlign w:val="bottom"/></w:tcPr>');
  });
});

describe('relationships used only inside kept table XML', () => {
  it('a link and a picture in a merged piece keep their relationships and media', async () => {
    const drawing =
      '<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const link = '<w:hyperlink r:id="rIdLink"><w:r><w:t>site</w:t></w:r></w:hyperlink>';
    const d = await open(
      table(1, [
        tc('A', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>'),
        tc('', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr>', link + drawing),
      ]),
      `<Relationship Id="rIdLink" Type="${REL}hyperlink" Target="https://example.com/" TargetMode="External"/>` +
        `<Relationship Id="rIdPic" Type="${REL}image" Target="media/merged.png"/>`,
      { 'word/media/merged.png': PNG },
    );
    const zip = await d.save();
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(await d.xmlOf(zip)).toContain(link);
    expect(rels).toContain('Id="rIdLink"');
    expect(rels).toContain('Id="rIdPic"');
    expect(zip.file('word/media/merged.png')).toBeTruthy();
  });
});

describe('legacy horizontal merge (w:hMerge)', () => {
  it('an untouched table keeps its w:hMerge cells element for element', async () => {
    const body = table(2, [
      tc('A', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:hMerge w:val="restart"/></w:tcPr>') + tc('', '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:hMerge/></w:tcPr>'),
      tc('C') + tc('D'),
    ]);
    const d = await open(body);
    const xml = await d.xmlOf();
    expect(xml).toContain('<w:hMerge w:val="restart"/>');
    const main = async (b: Uint8Array) => parseXml(await (await JSZip.loadAsync(b)).file('word/document.xml')!.async('string')).documentElement;
    expect(compareElements(await main(d.bytes), parseXml(xml).documentElement)).toEqual([]);
  });
});

describe('cells added to square up ragged rows', () => {
  const RAGGED = table(3, [tc('A1') + tc('B1') + tc('C1'), tc('A2')]);

  it('cell formatting commands leave them alone (they are not saved)', async () => {
    const d = await open(RAGGED);
    const filler = () => {
      let f: Record<string, unknown> | null = null;
      d.state.doc.descendants((n) => {
        if (n.type.name === 'table_cell' && n.attrs.filler) f = n.attrs;
      });
      return f;
    };
    expect(filler()).toBeTruthy();
    // Row 1's bottom edge borders the filler of row 2.
    d.cells('A1', 'C1');
    d.run(setCellBorders('all'));
    // A rectangle holding A1, B1, A2 and the filler.
    d.cells('B1', 'A2');
    d.run(setCellBackground('#ff0000'));
    d.run(setCellVAlign('bottom'));
    expect(filler()).toMatchObject({ background: null, vAlign: null, borders: null });
    const xml = await d.xmlOf();
    expect(rowWidths(xml)).toEqual([3, 1]);
  });
});

describe('drawing ids of new pictures', () => {
  it('start after the largest wp:docPr id in the package', async () => {
    const drawing = (id: number) =>
      `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="${id}" name="p${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPic"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const d = await open(
      `<w:p>${drawing(1)}${drawing(7)}</w:p><w:p><w:r><w:t>x</w:t></w:r></w:p>`,
      `<Relationship Id="rIdPic" Type="${REL}image" Target="media/p.png"/>`,
      { 'word/media/p.png': PNG },
    );
    d.run((s, dispatch) => {
      dispatch?.(s.tr.insert(s.doc.content.size - 1, schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 10, height: 10 })));
      return true;
    });
    const ids = [...(await d.xmlOf()).matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(ids).toEqual([1, 7, 8]);
  });
});
