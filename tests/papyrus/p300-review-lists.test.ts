// persona-300 code review, lists: 編號清單 inside a 公文 list turns it into 1. 2. 3. (as Word
// does); the 公文 numbers are followed directly by their text (w:suff "nothing"), in the page as
// in Word, so 「十一、」「（十一）」 no longer send their text to the next tab stop.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { ListCounter, isListKind, parseNumbering } from '../../src/papyrus/docx/numbering';
import type { Numbering } from '../../src/papyrus/docx/model';
import { parseXml } from '../../src/papyrus/docx/xml';
import { toggleList } from '../../src/papyrus/editor/commands';
import { domStubs, setup } from './p300Helpers';

domStubs();

const P = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
function markers(doc: PMNode, numbering: Numbering): string[] {
  const counter = new ListCounter(numbering);
  const out: string[] = [];
  doc.descendants((n) => {
    if (n.type.name !== 'paragraph') return true;
    out.push(n.attrs.numId ? counter.next(n.attrs.numId, n.attrs.ilvl) ?? '?' : '');
    return false;
  });
  return out;
}
const selectAll = (view: any) => view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1)));

describe('公文 list and 編號清單', () => {
  it('編號清單 in a 公文 list turns it into 1. 2. 3. (as Word does), not into plain paragraphs', async () => {
    const d = await setup(P('甲') + P('乙'));
    const n = d.editor.model.numbering;
    selectAll(d.view);
    d.editor.run(toggleList('gongwen', n));
    expect(d.editor.snapshot()!.list).toBe('gongwen');
    const gongwen = n.abstracts[n.nums[d.view.state.doc.child(0).attrs.numId].abstractId];
    expect(isListKind(gongwen, 'decimal')).toBe(false);
    d.editor.run(toggleList('decimal', n));
    expect(markers(d.view.state.doc, n)).toEqual(['1.', '2.']);
    expect(d.editor.snapshot()!.list).toBe('decimal');
    // Pressed again: the numbers go.
    d.editor.run(toggleList('decimal', n));
    expect(markers(d.view.state.doc, n)).toEqual(['', '']);
    d.done();
  });

  it('the text follows the number (w:suff nothing), in the page as in Word; read back from the file', async () => {
    const d = await setup(P('甲'));
    selectAll(d.view);
    d.editor.run(toggleList('gongwen', d.editor.model.numbering));
    const marker = d.view.dom.querySelector('.dx-marker') as HTMLElement;
    expect(marker.classList.contains('dx-marker-nothing')).toBe(true);
    expect(marker.style.minWidth).toBe('');
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/numbering.xml')!.async('string');
    expect(xml).toContain('<w:numFmt w:val="taiwaneseCountingThousand"/><w:suff w:val="nothing"/><w:lvlText w:val="%1、"/>');
    const parsed = parseNumbering(parseXml(xml));
    expect(Object.values(parsed.abstracts)[0].levels[0].suff).toBe('nothing');
    d.done();
  });

  it('a Word list with w:suff space shows a space after its number; without w:suff, a tab to the hanging indent', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const n = parseNumbering(
      parseXml(
        `<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="%1."/></w:lvl>` +
          `<w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum></w:numbering>`,
      ),
    );
    expect(n.abstracts['1'].levels.map((l) => l.suff)).toEqual(['space', undefined]);
  });
});
