// persona-300 A-7: Enter / Esc of an input method (注音 …) still composing a word pick or drop
// its characters; they never replace, close, cancel or apply in the editor's panels.
import { describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { composing, onEnter } from '../../src/papyrus/vue/keys';
import { domStubs, key, setup, typeInto } from './p300Helpers';
import { loadSfc } from './sfc';

domStubs();

describe('keys of a composing input method', () => {
  it('isComposing or keyCode 229 is composing; onEnter runs only for a real Enter', () => {
    expect(composing({ isComposing: true, keyCode: 13 } as KeyboardEvent)).toBe(true);
    expect(composing({ isComposing: false, keyCode: 229 } as KeyboardEvent)).toBe(true);
    expect(composing({ isComposing: false, keyCode: 13 } as KeyboardEvent)).toBe(false);
    let n = 0;
    const handler = onEnter(() => n++);
    const box = document.createElement('input');
    box.addEventListener('keydown', handler as EventListener);
    expect(key(box, { key: 'Enter', ime: true }).defaultPrevented).toBe(false);
    expect(n).toBe(0);
    expect(key(box, { key: 'Enter' }).defaultPrevented).toBe(true);
    expect(n).toBe(1);
  });
});

describe('find and replace', () => {
  it('Enter picking a character in 「取代為」 does not replace; Esc dropping candidates does not close', async () => {
    let closed = 0;
    const d = await setup('<w:p><w:r><w:t>甲案與甲案</w:t></w:r></w:p>', 'FindReplace.vue', { replace: true, onClose: () => closed++ });
    await d.refresh();
    const find = d.panel.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    await typeInto(find, '甲案');
    key(find, { key: 'Enter' }); // search now
    await d.refresh();
    const repl = d.panel.querySelector('input[aria-label="取代為"]') as HTMLInputElement;
    await typeInto(repl, 'ㄧˇ'); // still composing 乙
    const before = d.view.state.doc.textContent;
    const e = key(repl, { key: 'Enter', ime: true });
    expect(e.defaultPrevented).toBe(false);
    expect(d.view.state.doc.textContent).toBe(before);
    // The input method is done: now Enter replaces.
    await typeInto(repl, '乙案');
    key(repl, { key: 'Enter' });
    expect(d.view.state.doc.textContent).toBe('乙案與甲案');
    // Esc of the input method leaves the panel open; a real Esc closes it.
    key(repl, { key: 'Escape', ime: true });
    expect(closed).toBe(0);
    expect(key(repl, { key: 'Escape' }).defaultPrevented).toBe(true);
    expect(closed).toBe(1);
    d.done();
  });

  it('Enter confirming a word in 「尋找」 does not jump to the next match', async () => {
    const { searchState } = await import('../../src/papyrus/editor/search');
    const d = await setup('<w:p><w:r><w:t>公文 公文 公文</w:t></w:r></w:p>', 'FindReplace.vue', { replace: false });
    await d.refresh();
    const find = d.panel.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    await typeInto(find, '公文');
    key(find, { key: 'Enter' }); // searches
    await d.refresh();
    key(find, { key: 'Enter' }); // selects the first
    await d.refresh();
    expect(searchState(d.view.state).current).toBe(0);
    key(find, { key: 'Enter', ime: true });
    await d.refresh();
    expect(searchState(d.view.state).current).toBe(0);
    key(find, { key: 'Enter' });
    expect(searchState(d.view.state).current).toBe(1);
    d.done();
  });
});

describe('comments panel', () => {
  it('Esc of the input method keeps the comment being written; Ctrl+Enter while composing does not send', async () => {
    const Panel = await loadSfc('src/papyrus/vue/CommentsPanel.vue');
    const host = document.createElement('div');
    document.body.append(host);
    const events: string[] = [];
    const app = createApp({
      render: () =>
        h(Panel, {
          comments: [],
          anchors: {},
          activeId: null,
          editable: true,
          author: '我',
          draft: { quote: '主旨' },
          onAdd: (t: string) => events.push('add:' + t),
          onCancelDraft: () => events.push('cancel'),
        }),
    });
    app.mount(host);
    await nextTick();
    const box = host.querySelector('textarea[aria-label="新留言"]') as HTMLTextAreaElement;
    await typeInto(box, '請再確認日期');
    key(box, { key: 'Escape', ime: true });
    key(box, { key: 'Enter', ctrlKey: true, ime: true });
    expect(events).toEqual([]);
    key(box, { key: 'Enter', ctrlKey: true });
    expect(events).toEqual(['add:請再確認日期']);
    key(box, { key: 'Escape' });
    expect(events).toEqual(['add:請再確認日期', 'cancel']);
    app.unmount();
    host.remove();
  });
});

describe('table and picture boxes', () => {
  it('Enter while composing in the row height box applies nothing', async () => {
    const TABLE =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const d = await setup(TABLE, 'TablePanel.vue');
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    const box = d.panel.querySelector('input[aria-label="列高"]') as HTMLInputElement;
    await typeInto(box, '2');
    key(box, { key: 'Enter', ime: true });
    expect(d.view.state.doc.child(0).child(0).attrs.height).toBeNull();
    key(box, { key: 'Enter' });
    expect(d.view.state.doc.child(0).child(0).attrs.height).toBe(1134);
    d.done();
  });
});
