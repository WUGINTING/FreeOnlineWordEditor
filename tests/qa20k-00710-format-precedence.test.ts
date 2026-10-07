import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { readDocx } from '../src/papyrus/docx/reader';
import { blankPackage } from '../src/papyrus/docx/template';
import { schema } from '../src/papyrus/editor/schema';
import { selectionValue } from '../src/papyrus/editor/commands';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe('QA20K-00710 formatting source precedence', () => {
  it('resolves one conflicting color property from direct, character, paragraph, then document default', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
      '<w:p><w:pPr><w:pStyle w:val="GreenParagraph"/></w:pPr>' +
      '<w:r><w:t>paragraph</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="BlueCharacter"/></w:rPr><w:t>character</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="BlueCharacter"/><w:color w:val="FF0000"/></w:rPr><w:t>direct</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>default</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>');
    zip.file('word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="808080"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="GreenParagraph"><w:name w:val="GreenParagraph"/><w:rPr><w:color w:val="008000"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="BlueCharacter"><w:name w:val="BlueCharacter"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>' +
      '</w:styles>');

    const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const valueAt = (pos: number) => {
      const state = EditorState.create({ schema, doc });
      const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
      return selectionValue(selected, 'color', model.styles);
    };
    // Positions land inside the four words; character offsets do not affect color inheritance.
    const paragraphStart = 2;
    const characterStart = 2 + 'paragraph'.length + 1;
    const directStart = characterStart + 'character'.length + 1;
    const secondParagraphStart = doc.child(0).nodeSize + 2;

    expect(valueAt(paragraphStart)).toEqual({ value: '#008000', mixed: false });
    expect(valueAt(characterStart)).toEqual({ value: '#0000ff', mixed: false });
    expect(valueAt(directStart)).toEqual({ value: '#ff0000', mixed: false });
    expect(valueAt(secondParagraphStart)).toEqual({ value: '#808080', mixed: false });
  });
});
