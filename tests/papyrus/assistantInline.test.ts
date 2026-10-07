// What an assistant shows in the text without changing it (editor/assistantInline.ts): its
// suggested changes drawn in place, and a grey completion after the cursor.
import { describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { applyAssistantActions, assistantContext, bodyParagraphs, paragraphLine, planAssistantActions, type AssistantAction } from '../../src/papyrus/editor/assistant';
import {
  acceptCompletion, clearCompletion, clearSuggestion, completionOf, completionPoint, hasSuggestion, showCompletion, showSuggestion, showTarget, TYPED_EVENT,
} from '../../src/papyrus/editor/assistantInline';
import { domStubs, key, setup } from './p300Helpers';

domStubs();

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const LETTER = p('主旨：有關台北市政府辦理活動一案') + p('說明：依台北市政府來函辦理。') + p('請查照。');

type Doc = Awaited<ReturnType<typeof setup>>;
const lines = (d: Doc) => bodyParagraphs(d.view.state.doc).map((x) => paragraphLine(x.node));
const texts = (d: Doc, selector: string) => [...d.view.dom.querySelectorAll(selector)].map((el) => el.textContent);
const suggest = (d: Doc, actions: AssistantAction[]) => {
  const ctx = assistantContext(d.editor, 16000);
  const plan = planAssistantActions(d.editor, actions, ctx);
  return { ctx, plan, shown: !!plan.tr && showSuggestion(d.view, plan.tr) };
};
const cursorAtEndOf = (d: Doc, n: number) => {
  const para = bodyParagraphs(d.view.state.doc)[n - 1];
  d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, para.pos + 1 + para.node.content.size)));
  return para.pos + 1 + para.node.content.size;
};

describe('a suggestion drawn in the text', () => {
  it('shows what would go and what would come, and leaves the document as it is', async () => {
    const d = await setup(LETTER);
    const before = d.view.state.doc;
    const { shown } = suggest(d, [
      { type: 'set_paragraph_text', paragraph: 2, text: '說明：依臺北市政府來函辦理。' },
      { type: 'insert_paragraph', after: 2, text: '二、檢附活動簡章一份。' },
      { type: 'delete_paragraphs', from: 3, to: 3 },
    ]);
    expect(shown).toBe(true);
    expect(hasSuggestion(d.view.state)).toBe(true);
    // Only the character that differs is struck; the whole of the paragraph to delete is.
    expect(texts(d, '.dx-ai-del')).toEqual(['台', '請查照。']);
    expect(texts(d, 'ins.dx-ai-ins')).toEqual(['臺', '二、檢附活動簡章一份。']);
    expect(d.view.dom.querySelector('ins.dx-ai-ins-block')!.textContent).toBe('二、檢附活動簡章一份。');
    // Nothing was changed: the same document, nothing to undo, nothing to save.
    expect(d.view.state.doc).toBe(before);
    expect(lines(d)).toEqual(['主旨：有關台北市政府辦理活動一案', '說明：依台北市政府來函辦理。', '請查照。']);
    expect(d.editor.isModified()).toBe(false);

    clearSuggestion(d.view);
    expect(hasSuggestion(d.view.state)).toBe(false);
    expect(d.view.dom.querySelector('.dx-ai-del, .dx-ai-ins')).toBeNull();
    d.done();
  });

  it('highlights what would only be reformatted', async () => {
    const d = await setup(LETTER);
    suggest(d, [
      { type: 'format_paragraphs', from: 1, to: 1, align: 'center' },
      { type: 'format_text', from: 2, to: 2, text: '來函', bold: true },
    ]);
    expect(d.view.dom.querySelector('.dx-ai-fmt-block')!.textContent).toBe('主旨：有關台北市政府辦理活動一案');
    expect(texts(d, '.dx-ai-fmt')).toEqual(['來函']);
    expect(bodyParagraphs(d.view.state.doc)[0].node.attrs.align).toBeNull();
    d.done();
  });

  it('accepting makes the changes shown; an edit of the document meanwhile ends the suggestion', async () => {
    const d = await setup(LETTER);
    const actions: AssistantAction[] = [{ type: 'set_paragraph_text', paragraph: 3, text: '請 鑒核。' }];
    const first = suggest(d, actions);
    clearSuggestion(d.view);
    expect(applyAssistantActions(d.editor, actions, first.ctx)[0].ok).toBe(true);
    expect(lines(d)[2]).toBe('請 鑒核。');
    expect(d.view.dom.querySelector('.dx-ai-del, .dx-ai-ins')).toBeNull();

    suggest(d, [{ type: 'set_paragraph_text', paragraph: 1, text: '主旨：辦理活動一案' }]);
    expect(hasSuggestion(d.view.state)).toBe(true);
    d.view.dispatch(d.view.state.tr.insertText('（改）', 1));
    expect(hasSuggestion(d.view.state)).toBe(false);
    d.done();
  });

  it('marks the text a request is about while the keyboard is elsewhere', async () => {
    const d = await setup(LETTER);
    const para = bodyParagraphs(d.view.state.doc)[1];
    showTarget(d.view, para.pos + 4, para.pos + 9);
    expect(texts(d, '.dx-ai-target')).toEqual(['依台北市政']);
    clearSuggestion(d.view);
    expect(d.view.dom.querySelector('.dx-ai-target')).toBeNull();
    d.done();
  });
});

describe('a completion after the cursor', () => {
  it('is offered only at the end of a paragraph, with no selection', async () => {
    const d = await setup(LETTER);
    const end = cursorAtEndOf(d, 2);
    expect(completionPoint(d.view.state)).toBe(end);
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, end - 3)));
    expect(completionPoint(d.view.state)).toBeNull();
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, end - 3, end)));
    expect(completionPoint(d.view.state)).toBeNull();
    d.done();
  });

  it('shows as grey text that is not in the document; Tab types it', async () => {
    const d = await setup(LETTER);
    const end = cursorAtEndOf(d, 2);
    expect(showCompletion(d.view, '請依限辦理。', end)).toBe(true);
    expect(texts(d, '.dx-ai-ghost')).toEqual(['請依限辦理。']);
    expect(lines(d)[1]).toBe('說明：依台北市政府來函辦理。');
    expect(d.editor.isModified()).toBe(false);

    const e = key(d.view.dom, { key: 'Tab' });
    expect(e.defaultPrevented).toBe(true);
    expect(lines(d)[1]).toBe('說明：依台北市政府來函辦理。請依限辦理。');
    expect(completionOf(d.view.state)).toBeNull();
    expect(d.view.dom.querySelector('.dx-ai-ghost')).toBeNull();
    // One edit: one Ctrl+Z takes it back.
    d.editor.undo();
    expect(lines(d)[1]).toBe('說明：依台北市政府來函辦理。');
    d.done();
  });

  it('goes away with Escape, when the cursor moves, and when something is typed', async () => {
    const d = await setup(LETTER);
    let end = cursorAtEndOf(d, 2);
    showCompletion(d.view, '甲', end);
    key(d.view.dom, { key: 'Escape' });
    expect(completionOf(d.view.state)).toBeNull();
    expect(d.notices).toEqual([]); // Escape was the completion's: the editor did not leave the text

    showCompletion(d.view, '乙', end);
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, end - 2)));
    expect(completionOf(d.view.state)).toBeNull();

    end = cursorAtEndOf(d, 2);
    showCompletion(d.view, '丙', end);
    d.view.dispatch(d.view.state.tr.insertText('字', end));
    expect(completionOf(d.view.state)).toBeNull();
    // Asked for where the cursor was: not shown once it is elsewhere.
    expect(showCompletion(d.view, '丁', end)).toBe(false);
    clearCompletion(d.view);
    expect(acceptCompletion(d.view)).toBe(false);
    d.done();
  });

  it('a key of an input method still composing is never the completion\'s', async () => {
    const d = await setup(LETTER);
    const end = cursorAtEndOf(d, 2);
    showCompletion(d.view, '甲', end);
    key(d.view.dom, { key: 'Tab', ime: true });
    // Left to the editor (in a browser, to the input method): the completion's text is not typed.
    expect(lines(d)[1]).not.toContain('甲');
    key(d.view.dom, { key: 'Escape', ime: true });
    expect(lines(d)[1]).not.toContain('甲');
    d.done();
  });

  it('tells the page when the user typed, and only then', async () => {
    const d = await setup(LETTER);
    let typed = 0;
    d.host.addEventListener(TYPED_EVENT, () => typed++);
    const end = cursorAtEndOf(d, 2);
    expect(typed).toBe(0);
    d.view.dispatch(d.view.state.tr.insertText('字', end));
    expect(typed).toBe(1);
    // A completion shown or accepted is not typing; neither is a new paragraph or an undo.
    showCompletion(d.view, '甲', end + 1);
    acceptCompletion(d.view);
    d.editor.undo();
    expect(typed).toBe(1);
    d.done();
  });
});
