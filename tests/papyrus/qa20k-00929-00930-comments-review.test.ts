import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function mount() {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src: null, author: { name: 'Pat' }, onReady: (instance: DocxEditor) => { editor = instance; } }) });
  app.mount(host);
  for (let i = 0; i < 80 && !editor; i++) { await new Promise((resolve) => setTimeout(resolve, 0)); await nextTick(); }
  if (!editor) throw new Error('DocxEditor did not become ready');
  const flush = async () => { for (let i = 0; i < 8; i++) { await new Promise((resolve) => setTimeout(resolve, 0)); await nextTick(); } };
  const select = (text: string) => {
    let pos = -1;
    editor!.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); });
    if (pos < 0) throw new Error(`Text not found: ${text}`);
    editor!.view!.dispatch(editor!.view!.state.tr.setSelection(TextSelection.create(editor!.view!.state.doc, pos, pos + text.length)));
  };
  return { host, app, editor, flush, select, done: () => { app.unmount(); host.remove(); } };
}

describe('comment review navigation from document content', () => {
  it('QA20K-00929 opens the discussion panel when the reader clicks commented body text', async () => {
    const x = await mount();
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('Payment due Friday'));
      x.select('Friday');
      const id = x.editor.addComment('Confirm the due date')!;
      await x.flush();
      expect(x.host.querySelector('.dx-comments')).toBeNull();
      const anchor = x.host.querySelector(`[data-comment-id="${id}"].dx-comment`) as HTMLElement;
      expect(anchor).not.toBeNull();
      anchor.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await x.flush();
      expect(x.host.querySelector('.dx-comments')).not.toBeNull();
      expect(x.host.querySelector('.dx-comments')!.textContent).toContain('Confirm the due date');
    } finally { x.done(); }
  });

  it('QA20K-00930 selecting a reply card reveals its anchor and activates its parent thread', async () => {
    const x = await mount();
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('Invoice total'));
      x.select('total');
      const parent = x.editor.addComment('Check the amount')!;
      const reply = x.editor.replyComment(parent, 'Finance confirmed')!;
      await x.flush();
      expect(x.editor.comments().map((comment) => comment.text)).toEqual(['Check the amount', 'Finance confirmed']);
      const toggle = Array.from(x.host.querySelectorAll<HTMLButtonElement>('.dx-status-btn')).find((b) => b.textContent?.trim().startsWith('留言'))!;
      expect(toggle).toBeDefined();
      toggle.click(); await x.flush();
      const replyCard = x.host.querySelector<HTMLButtonElement>('.dx-comment-reply .dx-comment-main')!;
      expect(replyCard.textContent).toContain('Finance confirmed');
      replyCard.click(); await x.flush();
      expect(x.editor.view!.state.doc.textBetween(x.editor.view!.state.selection.from, x.editor.view!.state.selection.to)).toBe('total');
      const thread = x.host.querySelector('.dx-comments li')!;
      expect(thread.classList.contains('active')).toBe(true);
      expect(x.editor.comments().find((comment) => comment.id === reply)?.parentId).toBe(parent);
    } finally { x.done(); }
  });
});
