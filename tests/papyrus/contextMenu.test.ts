// GOV-184: Word's right-click menu over the text (ContextMenu.vue in DocxEditor.vue): the same
// commands as the ribbon, the table's rows and columns in a table, the keyboard (Shift+F10, the
// arrows, Esc), and the browser's own menu with Ctrl.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const cleanups: Array<() => void> = [];
afterEach(() => {
  vi.restoreAllMocks();
  while (cleanups.length) cleanups.pop()!();
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const CELL = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;

async function sample(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<w:document ${W}><w:body>`
    + `<w:p><w:r><w:t xml:space="preserve">Alpha beta </w:t></w:r><w:hyperlink r:id="rIdL"><w:r><w:t>site</w:t></w:r></w:hyperlink></w:p>`
    + `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${CELL('A1')}${CELL('B1')}</w:tr></w:tbl>`
    + `<w:p/><w:sectPr/></w:body></w:document>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.gov.tw/" TargetMode="External"/></Relationships>'));
  return zip.generateAsync({ type: 'uint8array' });
}

const flush = async () => {
  for (let i = 0; i < 10; i++) {
    await new Promise((r) => setTimeout(r, 0));
    await nextTick();
  }
};

async function mount(props: Record<string, unknown> = {}) {
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditor | null = null;
  const src = await sample();
  const app = createApp({ render: () => h(DocxEditorVue, { src, ...props, onReady: (e: DocxEditor) => (editor = e) }) });
  app.mount(host);
  for (let i = 0; i < 50 && !editor; i++) await flush();
  await flush();
  const menu = () => document.querySelector('.dx-context-menu') as HTMLElement | null;
  const labels = () => Array.from(menu()?.querySelectorAll('[role="menuitem"]') ?? []).map((b) => b.querySelector('span')!.textContent);
  const item = (key: string) => menu()!.querySelector(`[data-key="${key}"]`) as HTMLButtonElement;
  /** A right-click on the text (`ctrl`: with Ctrl held). */
  const rightClick = async (ctrl = false) => {
    const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 80, ctrlKey: ctrl });
    editor!.view!.dom.querySelector('p')!.dispatchEvent(ev);
    await flush();
    return ev;
  };
  let unmounted = false;
  const done = () => {
    if (unmounted) return;
    unmounted = true;
    app.unmount();
    host.remove();
  };
  cleanups.push(done);
  return { host, editor: editor!, menu, labels, item, rightClick, done };
}

function select(editor: DocxEditor, text: string, collapse = false) {
  const view = editor.view!;
  let at = -1;
  view.state.doc.descendants((n, pos) => {
    if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text);
    return at < 0;
  });
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, collapse ? at : at + text.length)));
}

describe('right-click menu', () => {
  it('opens over the text with the clipboard, 段落, 超連結 and 新增留言', async () => {
    const { editor, menu, labels, item, rightClick, done } = await mount();
    select(editor, 'beta');
    const ev = await rightClick();
    expect(ev.defaultPrevented).toBe(true);
    expect(menu()!.getAttribute('role')).toBe('menu');
    expect(labels()).toEqual(['剪下', '複製', '貼上', '全選', '段落…', '超連結…', '新增留言']);
    expect(item('cut').disabled).toBe(false);
    // The first item has the keyboard.
    expect(document.activeElement).toBe(item('cut'));
    done();
  });

  it('剪下 / 複製 are off without a selection; 複製 copies as Ctrl+C does', async () => {
    const { editor, menu, item, rightClick, done } = await mount();
    select(editor, 'beta', true);
    await rightClick();
    expect(item('cut').disabled).toBe(true);
    expect(item('copy').disabled).toBe(true);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    menu()!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    select(editor, 'beta');
    await rightClick();
    const exec = vi.fn(() => true);
    (document as any).execCommand = exec;
    item('copy').click();
    await flush();
    expect(exec).toHaveBeenCalledWith('copy');
    expect(menu()).toBeNull();
    done();
  });

  it('Ctrl+right-click leaves the browser its own menu (spelling suggestions)', async () => {
    const { editor, menu, rightClick, done } = await mount();
    select(editor, 'beta');
    const ev = await rightClick(true);
    expect(ev.defaultPrevented).toBe(false);
    expect(menu()).toBeNull();
    done();
  });

  it('on a link: edit, remove and open it', async () => {
    const { editor, labels, item, rightClick, done } = await mount();
    select(editor, 'site');
    await rightClick();
    expect(labels()).toEqual(expect.arrayContaining(['編輯超連結…', '移除超連結', '開啟超連結']));
    const open = vi.spyOn(window, 'open').mockImplementation(() => null);
    // Another site opens only once the user agreed to its address (persona-300).
    const ask = vi.spyOn(window, 'confirm').mockImplementation(() => true);
    item('open-link').click();
    expect(ask).toHaveBeenCalledWith('即將開啟外部網站「example.gov.tw」：\nhttps://example.gov.tw/\n確定要開啟嗎？');
    expect(open).toHaveBeenCalledWith('https://example.gov.tw/', '_blank', 'noopener,noreferrer');
    ask.mockRestore();
    // The editor's own confirm (the host's dialog), as for a click on the link.
    const own = vi.fn(() => false);
    (editor as any).options.confirm = own;
    const browserAsk = vi.spyOn(window, 'confirm');
    await flush();
    select(editor, 'site');
    await rightClick();
    item('open-link').click();
    expect(own).toHaveBeenCalledTimes(1);
    expect(browserAsk).not.toHaveBeenCalled();
    expect(open).toHaveBeenCalledTimes(1);
    browserAsk.mockRestore();
    await flush();
    select(editor, 'site');
    await rightClick();
    item('unlink').click();
    await flush();
    const marks = editor.view!.state.doc.nodeAt(editor.view!.state.selection.from)!.marks.map((m) => m.type.name);
    expect(marks).not.toContain('link');
    done();
  });

  it('in a table: its rows and columns, as on the ribbon', async () => {
    const { editor, labels, item, rightClick, done } = await mount();
    select(editor, 'A1', true);
    await rightClick();
    expect(labels()).toEqual(expect.arrayContaining(['在上方插入列', '在下方插入列', '在左側插入欄', '在右側插入欄', '刪除列', '刪除欄', '刪除表格']));
    // Nothing to merge in one cell.
    expect(item('merge').disabled).toBe(true);
    item('row-after').click();
    await flush();
    let rows = 0;
    editor.view!.state.doc.descendants((n) => {
      if (n.type.name === 'table_row') rows++;
    });
    expect(rows).toBe(2);
    done();
  });

  it('新增留言 opens the comment box for the selection', async () => {
    const { host, editor, item, rightClick, done } = await mount({ author: { name: '我' } });
    select(editor, 'beta');
    await rightClick();
    item('comment').click();
    await flush();
    expect(host.querySelector('.dx-comment-draft blockquote')!.textContent).toBe('beta');
    done();
  });

  it('while viewing only: 複製 and 全選', async () => {
    const { editor, labels, rightClick, done } = await mount({ editable: false });
    select(editor, 'beta');
    await rightClick();
    expect(labels()).toEqual(['複製', '全選']);
    done();
  });

  it('works from the keyboard: Shift+F10 opens it, the arrows move, Esc goes back to the text', async () => {
    const { host, editor, menu, item, done } = await mount();
    select(editor, 'beta');
    editor.view!.focus();
    vi.spyOn(editor.view!, 'hasFocus').mockReturnValue(true);
    vi.spyOn(editor.view!, 'coordsAtPos').mockReturnValue({ left: 50, right: 50, top: 60, bottom: 76 });
    const f10 = new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true, cancelable: true });
    host.querySelector('.dx-vue')!.dispatchEvent(f10);
    await flush();
    expect(f10.defaultPrevented).toBe(true);
    expect(menu()!.style.top).toBe('76px');
    expect(document.activeElement).toBe(item('cut'));
    const key = (k: string) => (document.activeElement as HTMLElement).dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
    key('ArrowDown');
    expect(document.activeElement).toBe(item('copy'));
    key('End');
    expect(document.activeElement).toBe(item('comment'));
    key('ArrowDown'); // wraps
    expect(document.activeElement).toBe(item('cut'));
    key('Escape');
    await flush();
    expect(menu()).toBeNull();
    expect(editor.view!.dom.contains(document.activeElement)).toBe(true);
    done();
  });

  it('a click elsewhere closes it', async () => {
    const { editor, menu, rightClick, done } = await mount();
    select(editor, 'beta');
    await rightClick();
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await flush();
    expect(menu()).toBeNull();
    done();
  });
});
