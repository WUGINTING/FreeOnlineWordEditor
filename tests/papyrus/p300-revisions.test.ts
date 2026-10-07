// persona-300 修訂與報讀: the 修訂窗格 (who, what, text, when in 民國; go there; accept / reject
// one), 上一個 / 下一個 修訂 said through aria-live, deletions readable by screen readers but never
// copied, the editor's accessible name and language, decorative copies hidden, icon buttons named.
import { describe, expect, it } from 'vitest';
import { DOMSerializer } from 'prosemirror-model';
import { schema } from '../../src/papyrus/editor/schema';
import { rocDate } from '../../src/papyrus/docx/revisions';
import { collectRevisions, revisionExcerpt, revisionSpoken } from '../../src/papyrus/editor/review';
import { listClipboardSerializer } from '../../src/papyrus/editor/pasteLists';
import { runLang } from '../../src/papyrus/editor/runLang';
import { domStubs, setup } from './p300Helpers';

domStubs();

const A = (id: number, author: string) => `w:id="${id}" w:author="${author}" w:date="2026-09-26T14:32:00Z"`;
const BODY =
  `<w:p><w:r><w:t xml:space="preserve">本案 </w:t></w:r><w:ins ${A(1, '王小明')}><w:r><w:t>同意</w:t></w:r></w:ins>` +
  `<w:del ${A(2, '李大華')}><w:r><w:delText>不同意</w:delText></w:r></w:del><w:r><w:t>辦理。</w:t></w:r></w:p>`;

describe('修訂窗格', () => {
  it('dates in 民國; how a revision is said; its text', async () => {
    expect(rocDate('2026-09-26T14:32:00Z')).toBe('115/09/26 14:32');
    expect(rocDate(null)).toBe('');
    const d = await setup(BODY);
    const revs = collectRevisions(d.view.state.doc);
    expect(revs.map((r) => revisionSpoken(r))).toEqual(['插入，王小明，115/09/26 14:32', '刪除，李大華，115/09/26 14:32']);
    expect(revs.map((r) => revisionExcerpt(d.view.state.doc, r))).toEqual(['同意', '不同意']);
    d.done();
  });

  it('lists every change; clicking goes there; 接受 / 拒絕 act on that one only', async () => {
    const d = await setup(BODY, 'RevisionsPanel.vue');
    await d.refresh();
    const items = () => Array.from(d.panel.querySelectorAll('li'));
    expect(items().map((li) => Array.from(li.querySelectorAll('.dx-revs-head > span')).map((e) => e.textContent).join(' '))).toEqual([
      '插入 王小明 115/09/26 14:32',
      '刪除 李大華 115/09/26 14:32',
    ]);
    expect(items()[1].querySelector('.dx-revs-text')!.textContent).toBe('不同意');
    (items()[1].querySelector('.dx-revs-main') as HTMLButtonElement).click();
    await d.refresh();
    expect(items()[1].classList.contains('active')).toBe(true);
    // Reject the deletion: the text comes back; the insertion is still there.
    (items()[1].querySelector('button[aria-label="拒絕：刪除，李大華"]') as HTMLButtonElement).click();
    await d.refresh();
    expect(d.view.state.doc.textContent).toBe('本案 同意不同意辦理。');
    expect(items()).toHaveLength(1);
    (items()[0].querySelector('button[aria-label="接受：插入，王小明"]') as HTMLButtonElement).click();
    await d.refresh();
    expect(items()).toHaveLength(0);
    expect(d.panel.textContent).toContain('這部分沒有修訂');
    // An edit elsewhere (undo): the list follows once typing would have paused.
    d.editor.undo();
    await d.refresh();
    await new Promise((r) => setTimeout(r, 300));
    await d.refresh();
    expect(items()).toHaveLength(1);
    d.done();
  });

  it('上一個 / 下一個 修訂 say the author and kind through aria-live', async () => {
    const d = await setup(BODY, 'DocxToolbar.vue', { styles: [] });
    await d.refresh();
    const next = d.panel.querySelector('button[aria-label="下一個修訂"]') as HTMLButtonElement;
    next.click();
    await d.refresh();
    const live = d.panel.querySelector('.dx-sr[aria-live="polite"]')!;
    expect(live.textContent).toBe('第 1 處，共 2 處：插入，王小明，115/09/26 14:32');
    next.click();
    await d.refresh();
    expect(live.textContent).toBe('第 2 處，共 2 處：刪除，李大華，115/09/26 14:32');
    d.done();
  }, 20000);
});

describe('screen readers', () => {
  it('a deletion is read once (刪除：…（刪除結束）, no role deletion besides) but not copied; insertions have role insertion', async () => {
    const d = await setup(BODY);
    const del = d.view.dom.querySelector('.dx-rev-del')!;
    expect(del.getAttribute('role')).toBeNull();
    expect(del.textContent).toBe('刪除：不同意（刪除結束）');
    expect(d.view.dom.querySelector('.dx-rev-ins')!.getAttribute('role')).toBe('insertion');
    // Copied: the deletion carries no text.
    const clip = listClipboardSerializer(() => d.editor.model.numbering).serializeFragment(d.view.state.doc.content);
    const div = document.createElement('div');
    div.append(clip);
    expect(div.querySelector('.dx-rev-del')!.textContent).toBe('');
    // Pasted back: nothing of it becomes text.
    const doc = d.view.state.doc;
    const html = DOMSerializer.fromSchema(schema).serializeFragment(doc.content);
    const box = document.createElement('div');
    box.append(html);
    const { DOMParser: PMParser } = await import('prosemirror-model');
    expect(PMParser.fromSchema(schema).parse(box).textContent).toBe('本案 同意辦理。');
    d.done();
  });

  it('the editor is named 文件內容 and is 中文; the pages behind it are hidden from screen readers', async () => {
    const d = await setup('<w:p><w:r><w:t>內容</w:t></w:r></w:p>');
    expect(d.view.dom.getAttribute('aria-label')).toBe('文件內容');
    expect(d.view.dom.getAttribute('lang')).toBe('zh-TW');
    expect(d.host.querySelector('.dx-pages')!.getAttribute('aria-hidden')).toBe('true');
    expect(d.host.querySelector('.dx-cuts')!.getAttribute('aria-hidden')).toBe('true');
    d.editor.editHeaderFooter('header');
    expect(d.editor.activeView!.dom.getAttribute('aria-label')).toBe('頁首');
    d.done();
  });

  it('Latin text of a run in another language gets its lang; Chinese text and zh runs do not', async () => {
    expect(runLang('<w:rPr><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr>')).toBe('en-US');
    expect(runLang('<w:rPr><w:lang w:val="zh-TW"/></w:rPr>')).toBeNull();
    const d = await setup(
      '<w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Executive Yuan</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr><w:t>行政院</w:t></w:r></w:p>',
    );
    const langs = Array.from(d.view.dom.querySelectorAll('[lang]')).filter((e) => e !== d.view.dom).map((e) => [e.getAttribute('lang'), e.textContent]);
    expect(langs).toEqual([['en-US', 'Executive Yuan']]);
    // Typing keeps it right.
    d.view.dispatch(d.view.state.tr.insertText(' Taiwan', 15));
    expect(Array.from(d.view.dom.querySelectorAll('[lang="en-US"]')).map((e) => e.textContent).join('')).toContain('Taiwan');
    d.done();
  });

  it('every ribbon button has a name a screen reader can say (not only a title)', async () => {
    const d = await setup('<w:p/>', 'DocxToolbar.vue', { styles: [] });
    await d.refresh();
    const unnamed = Array.from(d.panel.querySelectorAll<HTMLElement>('button, select')).filter((b) => {
      const text = (b.getAttribute('aria-label') ?? b.textContent ?? '').trim();
      return !/[一-鿿A-Za-z0-9]{2,}|[一-鿿]/.test(text) || /^[⌫⇤⇥x₂²]$/.test(text);
    });
    expect(unnamed.map((b) => b.outerHTML.slice(0, 80))).toEqual([]);
    d.done();
  }, 20000);
});
