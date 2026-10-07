// The 段落 dialog's command (setParagraphFormat): what is written to the .docx.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { setParagraphFormat } from '../../src/papyrus/editor/commands';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function open(pPr: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>段落文字</w:t></w:r></w:p><w:p><w:r><w:t>另一段</w:t></w:r></w:p></w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ doc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
  const apply = (patch: Record<string, number | string | null>) => {
    setParagraphFormat(patch)(state, (tr) => (state = state.apply(tr)));
  };
  const xml = async () => {
    const out = await JSZip.loadAsync(await writeDocx(state.doc, model));
    const s = await out.file('word/document.xml')!.async('string');
    return s.match(/<w:p\b[\s\S]*?<\/w:p>/g)!;
  };
  return { apply, xml, state: () => state };
}

describe('paragraph settings', () => {
  it('writes left indent, hanging indent, space before/after and exact line spacing', async () => {
    const d = await open('<w:jc w:val="center"/>');
    d.apply({ indLeft: 567, indLeftChars: null, indFirst: -283, indFirstChars: null, spaceBefore: 240, spaceAfter: 120, line: 360, lineRule: 'exact' });
    const [p1, p2] = await d.xml();
    expect(p1).toMatch(/<w:ind [^>]*w:left="567"/);
    expect(p1).toMatch(/<w:ind [^>]*w:hanging="283"/);
    expect(p1).toMatch(/<w:spacing [^>]*w:before="240"/);
    expect(p1).toMatch(/<w:spacing [^>]*w:after="120"/);
    expect(p1).toMatch(/<w:spacing [^>]*w:line="360"[^>]*w:lineRule="exact"|<w:spacing [^>]*w:lineRule="exact"[^>]*w:line="360"/);
    // What wasn't changed stays: the alignment, and the other paragraph.
    expect(p1).toContain('<w:jc w:val="center"/>');
    expect(p2).not.toContain('<w:ind');
  });

  it('indents in characters (字元) replace the twips ones; clearing gives the style its say again', async () => {
    const d = await open('<w:ind w:left="720"/>');
    d.apply({ indLeftChars: 200, indLeft: null, indFirstChars: 200, indFirst: null });
    let [p1] = await d.xml();
    expect(p1).toMatch(/w:leftChars="200"/);
    expect(p1).toMatch(/w:firstLineChars="200"/);
    expect(p1).not.toMatch(/w:left="720"/);
    d.apply({ indLeftChars: null, indFirstChars: null });
    [p1] = await d.xml();
    expect(p1).not.toMatch(/leftChars|firstLineChars/);
  });

  it('an unchanged value is no edit at all', async () => {
    const d = await open('<w:spacing w:after="120"/>');
    const before = d.state();
    d.apply({ spaceAfter: 120 });
    expect(d.state().doc).toBe(before.doc);
  });
});
