// Editing must change only what the user changed; everything else stays byte-identical.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { parseXml } from '../../src/papyrus/docx/xml';
import { compareElements } from './canonical';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';
const P1_PPR = '<w:pPr><w:keepNext/><w:snapToGrid w:val="0"/><w:spacing w:beforeLines="50" w:before="180" w:line="400" w:lineRule="exact"/><w:ind w:firstLineChars="200" w:firstLine="480"/><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="標楷體" w:eastAsia="標楷體"/></w:rPr></w:pPr>';
const RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:kern w:val="0"/><w:sz w:val="28"/><w:lang w:eastAsia="zh-TW"/></w:rPr>';

async function sample() {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
      `<w:p w14:paraId="11111111" w14:textId="22222222" w:rsidR="00AA0001">${P1_PPR}<w:r w:rsidRPr="00BB0002">${RPR}<w:t>第一段內容</w:t></w:r></w:p>` +
      `<w:p w14:paraId="33333333">${P1_PPR}<w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r>${RPR}<w:t>第二段</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>' +
      '</w:body></w:document>',
  );
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { bytes, ...(await readDocx(bytes)) };
}

async function body(bytes: Uint8Array): Promise<Element> {
  return parseXml(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')).documentElement;
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

describe('lossless editing', () => {
  it('saving without edits keeps every element', async () => {
    const { bytes, doc, model } = await sample();
    expect(compareElements(await body(bytes), await body(await writeDocx(doc, model)))).toEqual([]);
  });

  it('changing alignment only touches w:jc of that paragraph', async () => {
    const { doc, model } = await sample();
    const { EditorState } = await import('prosemirror-state');
    const state = EditorState.create({ schema, doc });
    const edited = state.apply(state.tr.setNodeMarkup(0, undefined, { ...doc.child(0).attrs, align: 'center' })).doc;
    const xml = await documentXml(await writeDocx(edited, model));
    const p1 = xml.slice(xml.indexOf('<w:p '), xml.indexOf('</w:p>'));
    expect(p1).toContain('<w:jc w:val="center"/>');
    expect(p1).not.toContain('w:val="both"');
    // Everything else in the paragraph properties is untouched.
    for (const kept of ['<w:keepNext/>', '<w:snapToGrid w:val="0"/>', 'w:beforeLines="50"', 'w:firstLineChars="200"', 'w:lineRule="exact"', 'w14:paraId="11111111"', 'w:rsidR="00AA0001"']) {
      expect(p1).toContain(kept);
    }
    expect(p1).toContain(RPR);
    // The second paragraph is byte-identical, bookmark included.
    expect(xml).toContain(`<w:p w14:paraId="33333333">${P1_PPR}<w:bookmarkStart w:id="0" w:name="_Toc1"/>`);
  });

  it('making part of a run bold keeps the run properties of both parts', async () => {
    const { doc, model } = await sample();
    // Bold "第一" (the first two characters of the first paragraph).
    const { EditorState } = await import('prosemirror-state');
    let state = EditorState.create({ schema, doc });
    state = state.apply(state.tr.addMark(1, 3, schema.marks.bold.create()));
    const xml = await documentXml(await writeDocx(state.doc, model));
    // Bold run: original properties plus w:b in schema order.
    expect(xml).toContain(
      '<w:r w:rsidRPr="00BB0002"><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman" w:hint="eastAsia"/><w:b/><w:kern w:val="0"/><w:sz w:val="28"/><w:lang w:eastAsia="zh-TW"/></w:rPr><w:t>第一</w:t></w:r>',
    );
    // The rest of the run is unchanged.
    expect(xml).toContain(`<w:r w:rsidRPr="00BB0002">${RPR}<w:t>段內容</w:t></w:r>`);
  });

  it('pressing Enter does not duplicate paragraph ids', async () => {
    const { doc, model } = await sample();
    const { EditorState, TextSelection } = await import('prosemirror-state');
    const { splitParagraph } = await import('../../src/papyrus/editor/commands');
    let state = EditorState.create({ schema, doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    splitParagraph(state, (tr) => (state = state.apply(tr)));
    const xml = await documentXml(await writeDocx(state.doc, model));
    expect(xml.match(/w14:paraId="11111111"/g)!.length).toBe(1);
    // The new paragraph keeps the formatting, as in Word.
    expect(xml.match(/w:firstLineChars="200"/g)!.length).toBe(3);
  });
});
