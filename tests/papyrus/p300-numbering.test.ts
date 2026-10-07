// persona-300 B-2: the 公文 list 一、（一）1.（1）甲、（甲） in Word's own formats, 重新從 1 開始編號
// and 接續編號; all saved so Word shows the same numbers, and read back the same.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import type { Node as PMNode } from 'prosemirror-model';
import { ListCounter, ensureList, formatNumber } from '../../src/papyrus/docx/numbering';
import { emptyNumbering, type Numbering } from '../../src/papyrus/docx/model';
import { continueNumbering, restartNumbering, toggleList } from '../../src/papyrus/editor/commands';
import { indent } from '../../src/papyrus/editor/commands';
import { domStubs, setup } from './p300Helpers';

domStubs();

const P = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

/** The list numbers the paragraphs show, in order ('' for a paragraph not in a list). */
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

/** Selects paragraphs `from`..`to` (0-based, top level). */
function selectParas(view: any, from: number, to = from) {
  const doc = view.state.doc;
  let a = 0;
  for (let i = 0; i < from; i++) a += doc.child(i).nodeSize;
  let b = a;
  for (let i = from; i <= to; i++) b += doc.child(i).nodeSize;
  view.dispatch(view.state.tr.setSelection(TextSelection.create(doc, a + 1, b - 1)));
}

describe('公文編號', () => {
  it('levels: 一、（一）1.（1）甲、（甲）子、（子）, in Word formats, each level starting where the one above has its text', () => {
    const n = emptyNumbering();
    const id = ensureList(n, 'gongwen');
    const levels = n.abstracts[n.nums[id].abstractId].levels;
    expect(levels.map((l) => [l.fmt, l.text])).toEqual([
      ['taiwaneseCountingThousand', '%1、'],
      ['taiwaneseCountingThousand', '（%2）'],
      ['decimal', '%3.'],
      ['decimal', '（%4）'],
      ['ideographTraditional', '%5、'],
      ['ideographTraditional', '（%6）'],
      ['ideographZodiac', '%7、'],
      ['ideographZodiac', '（%8）'],
      ['decimal', '%9)'],
    ]);
    for (let i = 1; i < 9; i++) expect(levels[i].indLeft! - levels[i].hanging!).toBe(levels[i - 1].indLeft);
    const c = new ListCounter(n);
    expect([0, 1, 2, 3, 4, 5, 6, 7].map((l) => c.next(id, l))).toEqual(['一、', '（一）', '1.', '（1）', '甲、', '（甲）', '子、', '（子）']);
    // The same list again: the buttons reuse it.
    expect(ensureList(n, 'gongwen')).toBe(id);
    expect(ensureList(n, 'decimal')).not.toBe(id);
  });

  it('Chinese numbers go on past ten as Word counts them', () => {
    expect([1, 10, 11, 20, 21, 99, 100, 101, 110, 1005].map((v) => formatNumber(v, 'taiwaneseCountingThousand'))).toEqual([
      '一', '十', '十一', '二十', '二十一', '九十九', '一百', '一百零一', '一百一十', '一千零五',
    ]);
  });

  it('applied from the ribbon, saved in Word formats and read back with the same numbers', async () => {
    const d = await setup(P('主旨') + P('說明') + P('細項') + P('細項二') + P('第二項'));
    const numbering = d.editor.model.numbering;
    selectParas(d.view, 0, 4);
    expect(d.editor.run(toggleList('gongwen', numbering))).toBe(true);
    selectParas(d.view, 2, 3);
    d.editor.run(indent(1));
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '（一）', '（二）', '三、']);
    expect(d.editor.snapshot()!.list).toBe('gongwen');
    const zip = await JSZip.loadAsync(await d.editor.save());
    const xml = await zip.file('word/numbering.xml')!.async('string');
    expect(xml).toContain('<w:numFmt w:val="taiwaneseCountingThousand"/><w:suff w:val="nothing"/><w:lvlText w:val="%1、"/>');
    expect(xml).toContain('<w:lvlText w:val="（%2）"/>');
    expect(xml).toContain('<w:numFmt w:val="ideographTraditional"/><w:suff w:val="nothing"/><w:lvlText w:val="%5、"/>');
    const doc = await zip.file('word/document.xml')!.async('string');
    expect(doc.match(/<w:numPr>/g)).toHaveLength(5);
    const again = await setup('<w:p/>');
    await again.editor.open(await d.editor.save());
    expect(markers(again.editor.view!.state.doc, again.editor.model.numbering)).toEqual(['一、', '二、', '（一）', '（二）', '三、']);
    again.done();
    d.done();
  });

  it('重新從 1 開始編號: this item and the ones after it start again; saved as a start override, one undo step', async () => {
    const d = await setup(P('甲') + P('乙') + P('丙') + P('丁'));
    const numbering = d.editor.model.numbering;
    selectParas(d.view, 0, 3);
    d.editor.run(toggleList('gongwen', numbering));
    selectParas(d.view, 2);
    d.view.dispatch(closeHistory(d.view.state.tr)); // a separate step, as when done later
    expect(d.editor.can(restartNumbering(numbering))).toBe(true);
    d.editor.run(restartNumbering(numbering));
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '一、', '二、']);
    const zip = await JSZip.loadAsync(await d.editor.save());
    const xml = await zip.file('word/numbering.xml')!.async('string');
    expect(xml).toMatch(/<w:num w:numId="\d+"><w:abstractNumId w:val="\d+"\/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"\/><\/w:lvlOverride><\/w:num>/);
    const again = await setup('<w:p/>');
    await again.editor.open(await d.editor.save());
    expect(markers(again.editor.view!.state.doc, again.editor.model.numbering)).toEqual(['一、', '二、', '一、', '二、']);
    again.done();
    d.editor.undo();
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '三、', '四、']);
    d.done();
  });

  it('接續編號: a list after a paragraph break counts on from the list before it', async () => {
    const d = await setup(P('一') + P('二') + P('中間的說明') + P('三') + P('四'));
    const numbering = d.editor.model.numbering;
    selectParas(d.view, 0, 1);
    d.editor.run(toggleList('gongwen', numbering));
    selectParas(d.view, 3, 4);
    d.editor.run(toggleList('gongwen', numbering));
    // Same definition: Word (and the editor) already count on.
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '', '三、', '四、']);
    selectParas(d.view, 3);
    d.editor.run(restartNumbering(numbering));
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '', '一、', '二、']);
    expect(d.editor.can(continueNumbering(numbering))).toBe(true);
    d.editor.run(continueNumbering(numbering));
    expect(markers(d.view.state.doc, numbering)).toEqual(['一、', '二、', '', '三、', '四、']);
    // Not in a list, or nothing before: nothing to continue.
    selectParas(d.view, 2);
    expect(d.editor.can(continueNumbering(numbering))).toBe(false);
    expect(d.editor.can(restartNumbering(numbering))).toBe(false);
    d.done();
  });

  it('a file whose list restarts with a start override shows the numbers Word shows', () => {
    const n = emptyNumbering();
    const levels = [{ fmt: 'decimal', text: '%1.', start: 1, indLeft: 720, hanging: 360 }];
    n.abstracts['1'] = { id: '1', levels, raw: '<w:abstractNum/>' };
    n.nums['1'] = { id: '1', abstractId: '1', startOverrides: {}, raw: '<w:num/>' };
    n.nums['2'] = { id: '2', abstractId: '1', startOverrides: { 0: 1 }, raw: '<w:num/>' };
    const c = new ListCounter(n);
    expect(['1', '1', '2', '2', '1'].map((id) => c.next(id, 0))).toEqual(['1.', '2.', '1.', '2.', '3.']);
  });
});
