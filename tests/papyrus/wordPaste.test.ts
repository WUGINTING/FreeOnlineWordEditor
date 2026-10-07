// Lists copied from Word (the clipboard's mso-list paragraphs, not <ol>/<ul>).
import { describe, expect, it } from 'vitest';
import { DOMParser as PMDOMParser, type Node as PMNode } from 'prosemirror-model';
import { schema, wordListNumber } from '../../src/papyrus/editor/schema';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';
import { emptyNumbering } from '../../src/papyrus/docx/model';
import { ListCounter } from '../../src/papyrus/docx/numbering';

/** One list item as Word puts it on the clipboard (Chrome keeps the conditional comments). */
const item = (list: string, level: number, lfo: number, number: string, text: string, font = '') =>
  `<p class=MsoListParagraphCxSpMiddle style='margin-left:${18 * level}pt;text-indent:-18.0pt;mso-list:${list} level${level} lfo${lfo}'>` +
  `<!--[if !supportLists]--><span lang=EN-US style='${font}'><span style='mso-list:Ignore'>${number}<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><!--[endif]-->` +
  `<span style='font-family:"微軟正黑體",sans-serif'>${text}</span></p>`;

function paste(html: string) {
  const div = document.createElement('div');
  div.innerHTML = `<html><body><!--StartFragment-->${html}<!--EndFragment--></body></html>`;
  const numbering = emptyNumbering();
  const slice = adoptPastedLists(PMDOMParser.fromSchema(schema).parseSlice(div), numbering);
  const paras: PMNode[] = [];
  slice.content.forEach((n) => paras.push(n));
  const counter = new ListCounter(numbering);
  const numbers = paras.map((p) => (p.attrs.numId ? counter.next(p.attrs.numId, p.attrs.ilvl) : ''));
  const level = (p: PMNode) => numbering.abstracts[numbering.nums[p.attrs.numId].abstractId].levels[p.attrs.ilvl];
  return { paras, numbers, level };
}

describe('lists pasted from Word', () => {
  it('a numbered list starting at 3 with a second level: numbers, levels and text (no typed-out number)', () => {
    const { paras, numbers, level } = paste(
      item('l0', 1, 1, '3.', '服務範圍') + item('l0', 2, 1, 'a.', '系統維護') + item('l0', 2, 1, 'b.', '教育訓練') + item('l0', 1, 1, '4.', '付款方式'),
    );
    expect(paras.map((p) => [p.textContent, p.attrs.ilvl])).toEqual([['服務範圍', 0], ['系統維護', 1], ['教育訓練', 1], ['付款方式', 0]]);
    expect(new Set(paras.map((p) => p.attrs.numId)).size).toBe(1);
    expect(numbers).toEqual(['3.', 'a.', 'b.', '4.']);
    expect(level(paras[1]).fmt).toBe('lowerLetter');
  });

  it('keeps the number style: (1), 一、, 甲、', () => {
    expect(paste(item('l1', 1, 2, '(1)', '第一項') + item('l1', 1, 2, '(2)', '第二項')).numbers).toEqual(['(1)', '(2)']);
    const cjk = paste(item('l2', 1, 3, '一、', '總則') + item('l2', 1, 3, '二、', '細則'));
    expect(cjk.numbers).toEqual(['一、', '二、']);
    expect(cjk.level(cjk.paras[0]).fmt).toBe('taiwaneseCountingThousand');
    expect(paste(item('l3', 1, 4, '甲、', '甲案')).numbers).toEqual(['甲、']);
  });

  it('a bulleted list (Symbol font dot) is a bullet list', () => {
    const { paras, level } = paste(item('l4', 1, 5, '·', '重點一', 'font-family:Symbol') + item('l4', 1, 5, '·', '重點二', 'font-family:Symbol'));
    expect(paras.map((p) => p.textContent)).toEqual(['重點一', '重點二']);
    expect(level(paras[0]).fmt).toBe('bullet');
  });

  it('two lists in one paste stay two lists; ordinary Word paragraphs stay ordinary', () => {
    const { paras } = paste(
      item('l0', 1, 1, '1.', 'A') + `<p class=MsoNormal>中間段落</p>` + item('l5', 1, 6, '1.', 'B'),
    );
    expect(paras.map((p) => p.textContent)).toEqual(['A', '中間段落', 'B']);
    expect(paras[1].attrs.numId).toBeNull();
    expect(paras[0].attrs.numId).not.toBe(paras[2].attrs.numId);
  });

  it('reads Word number texts', () => {
    expect(wordListNumber('3.', 0)).toEqual({ fmt: 'decimal', start: 3, text: '%1.' });
    expect(wordListNumber('(二)', 1)).toEqual({ fmt: 'taiwaneseCountingThousand', start: 2, text: '(%2)' });
    expect(wordListNumber('iv.', 0)).toEqual({ fmt: 'lowerRoman', start: 4, text: '%1.' });
    expect(wordListNumber('C)', 0)).toEqual({ fmt: 'upperLetter', start: 3, text: '%1)' });
    expect(wordListNumber('十二、', 0)).toEqual({ fmt: 'taiwaneseCountingThousand', start: 12, text: '%1、' });
    expect(wordListNumber('', 0).fmt).toBe('bullet');
  });
});
