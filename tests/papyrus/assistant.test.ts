// 小助手 (editor/assistant.ts): the document as numbered paragraphs, and a list of actions applied
// to it as one undo step without losing anything that is not text.
import { describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import {
  applyAssistantActions, assistantContext, bodyParagraphs, describeAction, numberRuns, paragraphLine, planAssistantActions, textEdits,
  type AssistantAction,
} from '../../src/papyrus/editor/assistant';
import { clearSuggestion, showSuggestion } from '../../src/papyrus/editor/assistantInline';
import { schema } from '../../src/papyrus/editor/schema';
import { domStubs, setup } from './p300Helpers';

domStubs();

const p = (text: string, pPr = '') => `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const LETTER = p('主旨：有關台北市政府辦理活動一案') + p('說明：依台北市政府來函辦理。') + p('請查照。');

const lines = (d: { view: { state: { doc: any } } }) => bodyParagraphs(d.view.state.doc).map((x) => paragraphLine(x.node));

describe('what the assistant is shown', () => {
  it('the paragraphs numbered from 1, the selection, the style names', async () => {
    const d = await setup(LETTER);
    let ctx = assistantContext(d.editor, 16000);
    expect(ctx.paragraphs.map((x) => [x.n, x.text])).toEqual([[1, '主旨：有關台北市政府辦理活動一案'], [2, '說明：依台北市政府來函辦理。'], [3, '請查照。']]);
    expect(ctx.total).toBe(3);
    expect(ctx.selection).toBeNull(); // only a cursor
    expect(ctx.partial).toEqual([]);

    const second = bodyParagraphs(d.view.state.doc)[1];
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, second.pos + 4, second.pos + 9)));
    ctx = assistantContext(d.editor, 16000);
    expect(ctx.selection).toEqual({ from: 2, to: 2, text: '依台北市政' });
    d.done();
  }, 20000); // the first editor of the file: loading it takes a few seconds

  it('a tab, a line break and a picture or field keep a place in the line; markers do not', async () => {
    const d = await setup(
      '<w:p><w:bookmarkStart w:id="1" w:name="here"/><w:r><w:t>甲</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>乙</w:t></w:r><w:bookmarkEnd w:id="1"/>'
      + '<w:r><w:br/></w:r><w:r><w:t>丙</w:t></w:r></w:p>',
    );
    expect(lines(d)).toEqual(['甲\t乙↵丙']);
    d.done();
  });

  it('a long document: the paragraphs around the selection that fit', async () => {
    const d = await setup(Array.from({ length: 30 }, (_, i) => p(`第${i + 1}段` + '字'.repeat(90))).join(''));
    const at = bodyParagraphs(d.view.state.doc)[19];
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, at.pos + 1)));
    const ctx = assistantContext(d.editor, 1000);
    expect(ctx.total).toBe(30);
    const sent = ctx.paragraphs.map((x) => x.n);
    expect(sent).toContain(20);
    expect(sent.length).toBeGreaterThan(3);
    expect(sent.length).toBeLessThan(12);
    // One run of neighbours.
    expect(sent).toEqual(Array.from({ length: sent.length }, (_, i) => sent[0] + i));
    expect(ctx.paragraphs.reduce((n, x) => n + x.text.length, 0)).toBeLessThanOrEqual(1000);
    d.done();
  });
});

describe('what the user is paying attention to (scope focus)', () => {
  // 40 paragraphs of about 100 characters: far more than the 1000 sent at a time here.
  const LONG = Array.from({ length: 40 }, (_, i) => p(`第${i + 1}段` + '字'.repeat(90))).join('');
  const numbers = (ctx: { paragraphs: { n: number }[] }) => ctx.paragraphs.map((x) => x.n);
  const cursorIn = (d: Awaited<ReturnType<typeof setup>>, n: number) => {
    const at = bodyParagraphs(d.view.state.doc)[n - 1];
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, at.pos + 1)));
  };
  const focus = { scope: 'focus' as const, focusChars: 1000, visible: null };

  it('the paragraph with the cursor and its neighbours, when nothing else is known', async () => {
    const d = await setup(LONG);
    cursorIn(d, 20);
    const ctx = assistantContext(d.editor, 16000, focus);
    expect(numbers(ctx)).toEqual([16, 17, 18, 19, 20, 21, 22, 23, 24]);
    expect(ctx.total).toBe(40);
    expect(ctx.focus).toEqual({ selected: 0, asked: 0, onScreen: 0, edited: 0 });
    // The same document read whole: all of it fits in what may be sent at most.
    expect(numbers(assistantContext(d.editor, 16000))).toHaveLength(40);
    d.done();
  });

  it('what is on screen counts, wherever the cursor was left', async () => {
    const d = await setup(LONG);
    cursorIn(d, 20);
    const paras = bodyParagraphs(d.view.state.doc);
    const ctx = assistantContext(d.editor, 16000, { ...focus, visible: { from: paras[4].pos + 1, to: paras[6].pos + 1 } });
    expect(numbers(ctx)).toEqual(expect.arrayContaining([5, 6, 7, 20]));
    expect(ctx.focus!.onScreen).toBe(3);
    // Two places, not one stretch: nothing from between them.
    expect(numbers(ctx).some((n) => n > 9 && n < 18)).toBe(false);
    expect(ctx.paragraphs.reduce((sum, x) => sum + x.text.length + 8, 0)).toBeLessThanOrEqual(1000);
    d.done();
  });

  it('what was just edited counts too, for some minutes', async () => {
    const d = await setup(LONG);
    const edited = bodyParagraphs(d.view.state.doc)[32];
    d.view.dispatch(d.view.state.tr.insertText('改', edited.pos + 1));
    cursorIn(d, 2);
    const ctx = assistantContext(d.editor, 16000, focus);
    expect(numbers(ctx)).toEqual(expect.arrayContaining([2, 33]));
    expect(ctx.focus!.edited).toBe(1);
    // An edit of a quarter of an hour ago says nothing about now.
    const later = assistantContext(d.editor, 16000, { ...focus, now: Date.now() + 15 * 60_000 });
    expect(numbers(later)).not.toContain(33);
    expect(later.focus!.edited).toBe(0);
    d.done();
  });

  it('the selection always goes, and so do the paragraphs the assistant asked for', async () => {
    const d = await setup(LONG);
    const paras = bodyParagraphs(d.view.state.doc);
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, paras[9].pos + 2, paras[10].pos + 5)));
    const ctx = assistantContext(d.editor, 16000, { ...focus, include: [{ from: 38, to: 40 }, { from: 39, to: 99 }] });
    expect(numbers(ctx)).toEqual(expect.arrayContaining([10, 11, 38, 39, 40]));
    expect(ctx.selection).toMatchObject({ from: 10, to: 11 });
    expect(ctx.focus).toMatchObject({ selected: 2, asked: 3 });
    // Asked for more than the budget for looking around: they are sent all the same (up to the most that may be sent).
    const many = assistantContext(d.editor, 16000, { ...focus, include: [{ from: 1, to: 30 }] });
    expect(numbers(many)).toEqual(expect.arrayContaining(Array.from({ length: 30 }, (_, i) => i + 1)));
    d.done();
  });

  it('an outline of the rest goes along: headings, and how 公文 mark their parts', async () => {
    const d = await setup(
      p('第一章 總則', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>') + p('主旨：有關辦理活動一案') + p('說明：') + p('一、依來函辦理。') + p('普通的一段' + '字'.repeat(300))
      + p('（二）第二點') + p('第 3 條 罰則') + p('又是普通的一段' + '字'.repeat(300)) + p('結尾' + '字'.repeat(300)),
    );
    cursorIn(d, 9);
    const ctx = assistantContext(d.editor, 16000, { ...focus, focusChars: 400 });
    expect(numbers(ctx)).toEqual([9]);
    expect(ctx.outline.map((o) => [o.n, o.level, o.text])).toEqual([
      [1, 0, '第一章 總則'], [2, null, '主旨：有關辦理活動一案'], [3, null, '說明：'], [4, null, '一、依來函辦理。'], [6, null, '（二）第二點'], [7, null, '第 3 條 罰則'],
    ]);
    d.done();
  });

  it('a document that fits is sent whole, with no outline', async () => {
    const d = await setup(LETTER);
    const ctx = assistantContext(d.editor, 16000, { scope: 'focus', focusChars: 4000, visible: null });
    expect(numbers(ctx)).toEqual([1, 2, 3]);
    expect(ctx.outline).toEqual([]);
    expect(ctx.focus).toBeNull();
    d.done();
  });

  it('paragraph numbers as runs', () => {
    expect(numberRuns([3, 4, 5, 9, 15, 16])).toBe('3～5、9、15～16');
    expect(numberRuns([7])).toBe('7');
    expect(numberRuns([])).toBe('');
  });

  it('an action about a paragraph that was not sent is not applied', async () => {
    const d = await setup(LONG);
    cursorIn(d, 20);
    const ctx = assistantContext(d.editor, 16000, focus);
    const out = applyAssistantActions(d.editor, [
      { type: 'set_paragraph_text', paragraph: 2, text: '沒有送給小助手的段落' },
      { type: 'set_paragraph_text', paragraph: 20, text: '改過的第二十段' },
    ], ctx);
    expect(out.map((o) => o.ok)).toEqual([false, true]);
    expect(lines(d)[19]).toBe('改過的第二十段');
    expect(lines(d)[1]).toContain('第2段');
    d.done();
  });
});

describe('applying the actions', () => {
  it('several kinds at once, all in one undo step', async () => {
    const d = await setup(LETTER);
    const ctx = assistantContext(d.editor, 16000);
    const actions: AssistantAction[] = [
      { type: 'replace_text', find: '台北', replace: '臺北' },
      { type: 'insert_paragraph', after: 2, text: '二、檢附活動簡章一份。' },
      { type: 'insert_paragraph', after: 2, text: '三、另請派員出席。' },
      { type: 'set_paragraph_text', paragraph: 3, text: '請 查照。' },
    ];
    const out = applyAssistantActions(d.editor, actions, ctx);
    expect(out.map((o) => o.ok)).toEqual([true, true, true, true]);
    expect(out[0].message).toBe('已取代 2 處');
    expect(lines(d)).toEqual(['主旨：有關臺北市政府辦理活動一案', '說明：依臺北市政府來函辦理。', '二、檢附活動簡章一份。', '三、另請派員出席。', '請 查照。']);
    expect(d.editor.isModified()).toBe(true);

    d.editor.undo();
    expect(lines(d)).toEqual(['主旨：有關台北市政府辦理活動一案', '說明：依台北市政府來函辦理。', '請查照。']);
    d.editor.redo();
    expect(lines(d)).toHaveLength(5);
    d.done();
  });

  it('rewriting a paragraph changes only what differs: the rest keeps its formatting, tabs and markers stay', async () => {
    const d = await setup(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>主旨：</w:t></w:r><w:r><w:tab/></w:r><w:bookmarkStart w:id="1" w:name="here"/>'
      + '<w:r><w:t>有關台北活動一案</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
    );
    const ctx = assistantContext(d.editor, 16000);
    expect(ctx.paragraphs[0].text).toBe('主旨：\t有關台北活動一案');
    const out = applyAssistantActions(d.editor, [{ type: 'set_paragraph_text', paragraph: 1, text: '主旨：\t有關臺北市活動一案' }], ctx);
    expect(out[0].ok).toBe(true);
    const para = d.view.state.doc.firstChild!;
    expect(paragraphLine(para)).toBe('主旨：\t有關臺北市活動一案');
    const kinds: string[] = [];
    para.forEach((child) => kinds.push(child.isText ? `${child.text}${schema.marks.bold.isInSet(child.marks) ? '(粗)' : ''}` : child.type.name));
    // The bold label, the tab and both bookmark ends are still there.
    expect(kinds[0]).toBe('主旨：(粗)');
    expect(kinds).toContain('tab');
    expect(kinds.filter((k) => k === 'raw_inline')).toHaveLength(2);
    expect(kinds.join('')).not.toContain('台北');

    // The model left the tab out: the tab stays all the same.
    const again = assistantContext(d.editor, 16000);
    applyAssistantActions(d.editor, [{ type: 'set_paragraph_text', paragraph: 1, text: '主旨：關於臺北市活動' }], again);
    const after: string[] = [];
    d.view.state.doc.firstChild!.forEach((child) => after.push(child.isText ? child.text! : child.type.name));
    expect(after).toContain('tab');
    expect(after.filter((k) => k === 'raw_inline')).toHaveLength(2);
    expect(d.view.state.doc.firstChild!.textContent).toBe('主旨：關於臺北市活動');
    d.done();
  });

  it('new paragraphs: before everything, and after a heading as normal text', async () => {
    const d = await setup(p('標題', '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>') + p('內文'));
    const ctx = assistantContext(d.editor, 16000);
    const out = applyAssistantActions(d.editor, [
      { type: 'insert_paragraph', after: 0, text: '最前面' },
      { type: 'insert_paragraph', after: 1, text: '標題後第一段\n標題後第二段' },
    ], ctx);
    expect(out.map((o) => o.ok)).toEqual([true, true]);
    expect(out[1].message).toBe('已新增 2 段');
    expect(lines(d)).toEqual(['最前面', '標題', '標題後第一段', '標題後第二段', '內文']);
    const paras = bodyParagraphs(d.view.state.doc);
    expect(paras[1].node.attrs.styleId).toBe('Heading1');
    expect(paras[2].node.attrs.styleId).toBeNull();
    d.done();
  });

  it('deleting: the paragraphs go; the last one left is emptied instead', async () => {
    const d = await setup(LETTER);
    let ctx = assistantContext(d.editor, 16000);
    expect(applyAssistantActions(d.editor, [{ type: 'delete_paragraphs', from: 2, to: 3 }], ctx)[0]).toEqual({ ok: true, message: '已刪除 2 段' });
    expect(lines(d)).toEqual(['主旨：有關台北市政府辦理活動一案']);
    ctx = assistantContext(d.editor, 16000);
    expect(applyAssistantActions(d.editor, [{ type: 'delete_paragraphs', from: 1, to: 1 }], ctx)[0].ok).toBe(true);
    expect(lines(d)).toEqual(['']);
    d.done();
  });

  it('formatting: alignment, bold on part of a paragraph, size and colour on whole paragraphs', async () => {
    const d = await setup(LETTER);
    const ctx = assistantContext(d.editor, 16000);
    const out = applyAssistantActions(d.editor, [
      { type: 'format_paragraphs', from: 1, to: 2, align: 'center' },
      { type: 'format_text', from: 1, to: 1, text: '主旨', bold: true },
      { type: 'format_text', from: 2, to: 3, text: '', fontSize: 14, color: 'ff0000', underline: true },
      { type: 'format_paragraphs', from: 1, to: 1, style: '沒有這個樣式' },
      { type: 'format_text', from: 3, to: 3, text: '不存在', bold: true },
    ], ctx);
    expect(out.map((o) => o.ok)).toEqual([true, true, true, false, false]);
    expect(out[3].message).toContain('沒有這個樣式');
    const paras = bodyParagraphs(d.view.state.doc);
    expect(paras.map((x) => x.node.attrs.align)).toEqual(['center', 'center', null]);
    const first: string[] = [];
    paras[0].node.forEach((child) => first.push(`${child.text}${schema.marks.bold.isInSet(child.marks) ? '(粗)' : ''}`));
    expect(first).toEqual(['主旨(粗)', '：有關台北市政府辦理活動一案']);
    const third = paras[2].node.firstChild!;
    expect(schema.marks.fontSize.isInSet(third.marks)?.attrs.pt).toBe(14);
    expect(schema.marks.color.isInSet(third.marks)?.attrs.color).toBe('#FF0000');
    expect(schema.marks.underline.isInSet(third.marks)).toBeTruthy();
    // The text itself is untouched.
    expect(lines(d)).toEqual(['主旨：有關台北市政府辦理活動一案', '說明：依台北市政府來函辦理。', '請查照。']);
    d.done();
  });

  it('a paragraph edited since the assistant read it is left alone; the other actions still apply', async () => {
    const d = await setup(LETTER);
    const ctx = assistantContext(d.editor, 16000);
    const second = bodyParagraphs(d.view.state.doc)[1];
    d.view.dispatch(d.view.state.tr.insertText('（已改）', second.pos + 1));
    const out = applyAssistantActions(d.editor, [
      { type: 'set_paragraph_text', paragraph: 2, text: '說明：依來函辦理。' },
      { type: 'set_paragraph_text', paragraph: 3, text: '請 鑒核。' },
      { type: 'set_paragraph_text', paragraph: 9, text: '沒有這一段' },
      { type: 'replace_text', find: '高雄', replace: '臺北' },
    ], ctx);
    expect(out.map((o) => o.ok)).toEqual([false, true, false, false]);
    expect(out[0].message).toContain('改過了');
    expect(out[3].message).toContain('找不到');
    expect(lines(d)).toEqual(['主旨：有關台北市政府辦理活動一案', '（已改）說明：依台北市政府來函辦理。', '請 鑒核。']);
    d.done();
  });

  it('nothing applied leaves the document unmodified', async () => {
    const d = await setup(LETTER);
    const ctx = assistantContext(d.editor, 16000);
    const out = applyAssistantActions(d.editor, [{ type: 'replace_text', find: '高雄', replace: '臺北' }], ctx);
    expect(out[0].ok).toBe(false);
    expect(d.editor.isModified()).toBe(false);
    d.done();
  });

  it('while 追蹤修訂 is on, the changes are tracked changes', async () => {
    const d = await setup(LETTER);
    d.editor.setTrackChanges(true);
    const ctx = assistantContext(d.editor, 16000);
    const out = applyAssistantActions(d.editor, [{ type: 'replace_text', find: '台北', replace: '臺北' }], ctx);
    expect(out[0].ok).toBe(true);
    const json = JSON.stringify(d.view.state.doc.toJSON());
    // The one character that changed: the new one is an insertion, the old one stays as a deletion.
    expect(lines(d)[0]).toBe('主旨：有關臺北市政府辦理活動一案');
    expect(json).toContain('<w:delText>台</w:delText>');
    expect(json).not.toContain('<w:delText>台北');
    expect(json).toMatch(/w:ins|"ins"/);
    expect(json).toMatch(/w:del|"del"/);
    d.done();
  });
});

describe('a corrected paragraph changes only where it differs', () => {
  it('finds each place: a wrong character is one small change, not everything from the first to the last', () => {
    const old = '由於此兩類錯誤均可透過確定性規則攔截，吳需依賴 LLM 自我判斷，因此採用二分類法。';
    expect(textEdits(old, old.replace('吳需', '無需'))).toEqual([{ from: 19, to: 20, text: '無' }]);
    expect(textEdits(old, old.replace('此兩類', '這兩類').replace('吳需', '無需'))).toEqual([
      { from: 2, to: 3, text: '這' }, { from: 19, to: 20, text: '無' },
    ]);
    // Added, removed, at either end.
    expect(textEdits('請查照', '請 查照。')).toEqual([{ from: 1, to: 1, text: ' ' }, { from: 3, to: 3, text: '。' }]);
    expect(textEdits('依據來函辦理辦理。', '依據來函辦理。')).toEqual([{ from: 6, to: 8, text: '' }]);
    expect(textEdits('同樣的字', '同樣的字')).toEqual([]);
    // Each character on its own, even one character apart; only changes right next to each other are one.
    expect(textEdits('台北市政府', '臺北縣政府')).toEqual([{ from: 0, to: 1, text: '臺' }, { from: 2, to: 3, text: '縣' }]);
    expect(textEdits('其二為癮蔽的邏輯錯誤', '其三為隱蔽的邏輯錯誤')).toEqual([{ from: 1, to: 2, text: '三' }, { from: 3, to: 4, text: '隱' }]);
    expect(textEdits('辦裡情形', '辦理狀況')).toEqual([{ from: 1, to: 4, text: '理狀況' }]);
  });

  it('replacing a phrase by a corrected one changes only its characters that differ, wherever the phrase is', async () => {
    const d = await setup(p('不相容；其二為癮蔽的邏輯錯誤，語法正確。') + p('再說一次：其二為癮蔽的邏輯錯誤。'));
    const ctx = assistantContext(d.editor, 16000);
    const actions: AssistantAction[] = [{ type: 'replace_text', find: '其二為癮蔽的邏輯錯誤', replace: '其二為隱蔽的邏輯錯誤' }];
    const plan = planAssistantActions(d.editor, actions, ctx);
    expect(plan.outcomes[0]).toEqual({ ok: true, message: '已取代 2 處' });
    showSuggestion(d.view, plan.tr!);
    // One character in each place, not the phrase.
    expect([...d.view.dom.querySelectorAll('.dx-ai-del')].map((el) => el.textContent)).toEqual(['癮', '癮']);
    expect([...d.view.dom.querySelectorAll('ins.dx-ai-ins')].map((el) => el.textContent)).toEqual(['隱', '隱']);
    clearSuggestion(d.view);
    applyAssistantActions(d.editor, actions, ctx);
    expect(lines(d)).toEqual(['不相容；其二為隱蔽的邏輯錯誤，語法正確。', '再說一次：其二為隱蔽的邏輯錯誤。']);
    d.editor.undo();
    expect(lines(d)[0]).toBe('不相容；其二為癮蔽的邏輯錯誤，語法正確。');
    d.done();
  });

  it('a rewrite, where little stays, is one change', () => {
    const edits = textEdits('請查照辦理並回復本府為荷。', '敬請貴機關於文到後儘速核處。');
    expect(edits).toHaveLength(1);
    expect(edits[0]).toMatchObject({ from: 0, text: '敬請貴機關於文到後儘速核處' });
  });

  it('two wrong characters far apart: two small edits, drawn as such, and the formatting between them stays', async () => {
    const d = await setup(
      '<w:p><w:r><w:t>由於此兩類錯誤均可攔截，</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>吳需依賴自我判斷</w:t></w:r><w:r><w:t>，因此採用二分類法。</w:t></w:r></w:p>',
    );
    const ctx = assistantContext(d.editor, 16000);
    const actions: AssistantAction[] = [{ type: 'set_paragraph_text', paragraph: 1, text: '由於這兩類錯誤均可攔截，無需依賴自我判斷，因此採用二分類法。' }];
    const plan = planAssistantActions(d.editor, actions, ctx);
    expect(plan.tr!.steps).toHaveLength(2);
    showSuggestion(d.view, plan.tr!);
    expect([...d.view.dom.querySelectorAll('.dx-ai-del')].map((el) => el.textContent)).toEqual(['此', '吳']);
    expect([...d.view.dom.querySelectorAll('ins.dx-ai-ins')].map((el) => el.textContent)).toEqual(['這', '無']);
    clearSuggestion(d.view);

    applyAssistantActions(d.editor, actions, ctx);
    const runs: string[] = [];
    d.view.state.doc.firstChild!.forEach((child) => runs.push(`${child.text}${schema.marks.bold.isInSet(child.marks) ? '(粗)' : ''}`));
    expect(runs).toEqual(['由於這兩類錯誤均可攔截，', '無需依賴自我判斷(粗)', '，因此採用二分類法。']);
    d.editor.undo();
    expect(lines(d)).toEqual(['由於此兩類錯誤均可攔截，吳需依賴自我判斷，因此採用二分類法。']);
    d.done();
  });
});

describe('an action in words', () => {
  it('says what will change', () => {
    expect(describeAction({ type: 'replace_text', find: '台北', replace: '臺北' })).toBe('把所有的「台北」改成「臺北」');
    expect(describeAction({ type: 'insert_paragraph', after: 0, text: '開頭' })).toBe('在最前面新增：「開頭」');
    expect(describeAction({ type: 'delete_paragraphs', from: 2, to: 4 })).toBe('刪除第 2～4 段');
    expect(describeAction({ type: 'format_paragraphs', from: 1, to: 1, align: 'center', style: '標題 1' })).toBe('第 1 段套用「標題 1」樣式、置中');
    expect(describeAction({ type: 'format_text', from: 1, to: 1, text: '主旨', bold: true, fontSize: 14 })).toBe('第 1 段的「主旨」設為粗體、14 點');
  });

  it('a corrected paragraph is said as the places it changes; a rewritten one as its new text', () => {
    const shown = {
      paragraphs: [{ n: 17, text: '由於此兩類錯誤均可攔截，吳需依賴自我判斷。', style: null }], total: 20, selection: null, styles: [], partial: [], outline: [], focus: null,
    };
    expect(describeAction({ type: 'set_paragraph_text', paragraph: 17, text: '由於這兩類錯誤均可攔截，無需依賴自我判斷。' }, shown))
      .toBe('第 17 段：「由於此兩類」→「由於這兩類」、「截，吳需依」→「截，無需依」');
    expect(describeAction({ type: 'set_paragraph_text', paragraph: 17, text: '這兩類錯誤都能用規則擋下，不必靠模型自己判斷，所以分成兩類處理。' }, shown))
      .toBe('第 17 段改為：「這兩類錯誤都能用規則擋下，不必靠模型自己判斷，所以分成兩類處理。」');
    expect(describeAction({ type: 'replace_text', find: '其二為癮蔽的邏輯錯誤', replace: '其三為隱蔽的邏輯錯誤' }))
      .toBe('「其二為癮蔽的邏輯錯誤」裡的「其二為」改成「其三為」、「為癮蔽」改成「為隱蔽」');
    // Without what the assistant was shown: the new text.
    expect(describeAction({ type: 'set_paragraph_text', paragraph: 17, text: '短' })).toBe('第 17 段改為：「短」');
  });
});
