// @vitest-environment jsdom
// The editor toolbar from the government-user review (2026-09-26):
//  - GOV-ISSUE-014: the centre of 「清除格式」 hit the invisible highlight colour input next to it;
//  - Word's ribbon: every tool under the tab where Word has it (檔案、常用、插入、設計、版面配置、參考資料、
//    校閱、檢視, and 表格 / 圖片 tabs while a table or picture is selected), 摺疊功能區 (which
//    replaced the 精簡工具列 of GOV-FINDING-024);
//  - GOV-FINDING-011: the toolbar wraps onto more rows instead of scrolling sideways;
//  - GOV-ISSUE-013: the toolbar is one Tab stop (WAI-ARIA toolbar: roving tabindex, arrow keys),
//    Alt+F10 from the text goes to it and Escape goes back to the text.
import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor as DocxEditorType } from '../../src/papyrus/editor/core';
import { insertTable } from '../../src/papyrus/editor/commands';
import { schema } from '../../src/papyrus/editor/schema';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const KEY = 'papyrus.ribbon.collapsed';
afterEach(() => {
  vi.restoreAllMocks();
  try {
    localStorage.clear();
  } catch {
    // ignore
  }
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

async function mount(extra: Record<string, unknown> = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>工具列測試文字</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const src = await zip.generateAsync({ type: 'uint8array' });
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditorType | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditorType) => (editor = ed), ...extra }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  const toolbar = host.querySelector('[role="toolbar"]') as HTMLElement;
  const tabs = () => Array.from(host.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
  const tab = (label: string) => tabs().find((t) => t.textContent?.trim() === label)!;
  return { host, toolbar, tabs, tab, editor: editor as DocxEditorType, flush, done: () => { app.unmount(); host.remove(); } };
}

/** Hidden by v-show / the hidden attribute somewhere up to the toolbar. */
function shown(el: Element, root: Element) {
  for (let n: Element | null = el; n && n !== root; n = n.parentElement) {
    if ((n as HTMLElement).hidden || (n as HTMLElement).style?.display === 'none') return false;
  }
  return true;
}
const CONTROLS = 'button, select, input';
/** The toolbar's controls a keyboard user can reach, in order. */
function reachable(toolbar: HTMLElement) {
  return Array.from(toolbar.querySelectorAll<HTMLElement>(CONTROLS)).filter(
    (el) => !(el as HTMLButtonElement).disabled && !(el as HTMLInputElement).hidden && shown(el, toolbar),
  );
}
const byTitle = (root: Element, title: string) => root.querySelector(`[title="${title}"]`) as HTMLElement;
const key = (el: Element, k: string, init: KeyboardEventInit = {}) => {
  const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init });
  el.dispatchEvent(e);
  return e;
};

/** The rule body of `selector` in a component's <style>. */
function cssRule(file: string, selector: string): string {
  const css = (readFileSync(file, 'utf8').split('<style')[1] ?? '').replace(/\r\n/g, '\n');
  const esc = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // A rule of its own (not `button, select, .dx-color { … }`).
  const m = new RegExp(`(?<!,)\\n${esc}\\s*\\{([^}]*)\\}`).exec(css);
  return m ? m[1] : '';
}

describe('GOV-ISSUE-014: every toolbar button is hit where it is drawn', () => {
  it('the colour inputs stay inside their swatch (no overlay over 清除格式)', async () => {
    const ui = await mount();
    // The colour menus (persona-300): the browser's picker is 「其他色彩…」 inside the open menu.
    const menus = Array.from(ui.toolbar.querySelectorAll<HTMLButtonElement>('.dx-cp-button'));
    expect(menus.length).toBeGreaterThanOrEqual(2);
    for (const b of menus) {
      b.click();
      await ui.flush();
      const colours = Array.from(ui.toolbar.querySelectorAll<HTMLInputElement>('input[type="color"]'));
      expect(colours).toHaveLength(1);
      const input = colours[0];
      // The invisible input covers its own swatch label and nothing else: it is the only control
      // in that label and has its own name.
      const label = input.closest('label')!;
      expect(label).not.toBeNull();
      expect(label.querySelectorAll(CONTROLS)).toHaveLength(1);
      expect(input.getAttribute('aria-label')).toBeTruthy();
      b.click();
      await ui.flush();
    }
    // 清除格式 is its own button, not inside a swatch.
    const clear = byTitle(ui.toolbar, '清除格式 (Ctrl+Space)');
    expect(clear.tagName).toBe('BUTTON');
    expect(clear.closest('label')).toBeNull();
    ui.done();
  });

  it('the swatch CSS sizes the input to the swatch and clips it (a colour input is 50 px wide by itself)', () => {
    for (const [file, swatch] of [['src/papyrus/vue/ColorPicker.vue', '.dx-cp-other']]) {
      const box = cssRule(file, swatch);
      expect(box, `${file} ${swatch}`).toMatch(/position:\s*relative/);
      expect(box, `${file} ${swatch}`).toMatch(/overflow:\s*hidden/);
      const input = cssRule(file, `${swatch} input`);
      expect(input, `${file} ${swatch} input`).toMatch(/position:\s*absolute/);
      expect(input).toMatch(/width:\s*100%/);
      expect(input).toMatch(/height:\s*100%/);
    }
  });
});

/** Puts the cursor in the body's first table cell. */
function intoTable(editor: DocxEditorType) {
  const view = editor.view!;
  let cell = -1;
  view.state.doc.descendants((node, pos) => {
    if (cell < 0 && node.type.name === 'table') cell = pos;
    return cell < 0;
  });
  view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(cell + 1))));
}

describe("Word's ribbon: every tool under the tab where Word has it", () => {
  const TABS = ['常用', '插入', '設計', '版面配置', '參考資料', '校閱', '檢視'];
  // Where Word has each tool (by the button's title).
  const WHERE: Record<string, string[]> = {
    常用: ['復原 (Ctrl+Z)', '取消復原 (Ctrl+Y)', '貼上 (Ctrl+V)', '剪下 (Ctrl+X)', '複製 (Ctrl+C)',
      '複製格式：複製所選文字的格式，再用滑鼠選取要套用的文字（按兩下可連續套用，按 Esc 結束） (鍵盤：Ctrl+Shift+C 複製格式，再按 Ctrl+Shift+V 貼上格式)', '字型', '字型大小',
      '清除格式 (Ctrl+Space)', '粗體 (Ctrl+B)', '斜體 (Ctrl+I)', '底線 (Ctrl+U)', '刪除線', '下標', '上標', '螢光標記',
      '項目符號清單', '編號清單', '減少縮排', '增加縮排',
      '顯示／隱藏編輯標記：顯示段落標記、空格、定位字元等不會列印的符號 (Ctrl+Shift+8)', '靠左對齊', '置中', '靠右對齊', '左右對齊', '行距',
      '段落設定（縮排、段前段後距離、行距）', '段落樣式', '尋找 (Ctrl+F)', '取代 (Ctrl+H)', '全選 (Ctrl+A)'],
    插入: ['插入分頁符號 (Ctrl+Enter)', '插入表格', '插入圖片', '插入連結', '頁碼', '符號：插入鍵盤上沒有的符號，例如 ※、①、℃'],
    設計: ['浮水印：在每一頁的文字後面加上「機密」、「草稿」等淡色文字或圖片'],
    版面配置: ['邊界', '方向', '大小（紙張）', '欄（分欄）', '分隔設定：分頁符號、分欄符號、分節符號', '版面設定（紙張大小、方向、邊界）'],
    參考資料: [],
    校閱: ['字數統計', '上一個修訂', '下一個修訂', '接受這部分的所有修訂'],
    檢視: ['導覽窗格：依標題瀏覽文件', '縮放到 100%', '頁寬：頁面和視窗一樣寬'],
    // Contextual: shown while the cursor is in a table.
    表格版面配置: ['請先選取兩個以上的儲存格', '分割儲存格', '分割表格：從游標所在的列起分成兩個表格 (Ctrl+Shift+Enter)',
      '自動調整：依內容或視窗寬度調整欄寬，或固定欄寬', '平均分配欄寬（選取多欄時只分配這些欄）'],
  };

  it("has 檔案 and Word's tabs in Word's order, 常用 open", async () => {
    const ui = await mount();
    const file = ui.host.querySelector('.dx-file') as HTMLButtonElement;
    expect(file.textContent?.trim()).toBe('檔案');
    expect(ui.tabs().map((t) => t.textContent?.trim())).toEqual(TABS);
    expect(ui.tab('常用').getAttribute('aria-selected')).toBe('true');
    // One Tab stop in the tab row.
    expect(ui.tabs().filter((t) => t.tabIndex === 0)).toEqual([ui.tab('常用')]);
    const panel = document.getElementById(ui.tab('常用').getAttribute('aria-controls')!)!;
    expect(panel.getAttribute('role')).toBe('tabpanel');
    expect(panel.getAttribute('aria-labelledby')).toBe(ui.tab('常用').id);
    ui.done();
  });

  it('shows each tool only on its own tab', async () => {
    const ui = await mount();
    // In a table, so that the table tabs are there too.
    ui.editor.run(insertTable(2, 2, 600));
    intoTable(ui.editor);
    await ui.flush();
    for (const [name, titles] of Object.entries(WHERE)) {
      ui.tab(name).click();
      await ui.flush();
      expect(ui.tab(name).getAttribute('aria-selected'), name).toBe('true');
      for (const t of titles) expect(shown(byTitle(ui.toolbar, t), ui.toolbar), `${t} on ${name}`).toBe(true);
      for (const [other, rest] of Object.entries(WHERE)) {
        if (other !== name) for (const t of rest) expect(shown(byTitle(ui.toolbar, t), ui.toolbar), `${t} not on ${name}`).toBe(false);
      }
    }
    // Buttons found by their name, not only their title.
    const named = (label: string) =>
      Array.from(ui.toolbar.querySelectorAll('button')).filter((b) => b.textContent?.trim() === label && shown(b, ui.toolbar));
    ui.tab('校閱').click();
    await ui.flush();
    expect(named('新增留言')).toHaveLength(1);
    expect(named('追蹤修訂')).toHaveLength(1);
    ui.tab('參考資料').click();
    await ui.flush();
    expect(named('更新目錄')).toHaveLength(1);
    ui.done();
  });

  it('Left / Right move between the tabs and open them; Down goes into the tools', async () => {
    const ui = await mount();
    const home = ui.tab('常用');
    home.focus();
    let e = key(home, 'ArrowRight');
    expect(e.defaultPrevented).toBe(true);
    await ui.flush();
    expect(document.activeElement).toBe(ui.tab('插入'));
    expect(ui.tab('插入').getAttribute('aria-selected')).toBe('true');
    expect(shown(byTitle(ui.toolbar, '插入表格'), ui.toolbar)).toBe(true);
    key(ui.tab('插入'), 'End');
    await ui.flush();
    expect(document.activeElement).toBe(ui.tab('檢視'));
    key(ui.tab('檢視'), 'ArrowRight'); // wraps around
    await ui.flush();
    expect(document.activeElement).toBe(ui.tab('常用'));
    e = key(ui.tab('常用'), 'ArrowDown');
    expect(e.defaultPrevented).toBe(true);
    await ui.flush();
    expect(document.activeElement).toBe(reachable(ui.toolbar)[0]);
    ui.done();
  });

  it('a table brings 表格設計 and 表格版面配置 (as in Word), and inserting one opens its tools', async () => {
    const ui = await mount();
    expect(ui.tabs().map((t) => t.textContent?.trim())).not.toContain('表格設計');
    ui.tab('插入').click();
    await ui.flush();
    byTitle(ui.toolbar, '插入表格').click();
    await ui.flush();
    (ui.toolbar.querySelector('[aria-label="2 列 3 欄"]') as HTMLButtonElement).click();
    await ui.flush();
    expect(ui.tabs().map((t) => t.textContent?.trim())).toEqual([...TABS, '表格設計', '表格版面配置']);
    expect(ui.tab('表格設計').getAttribute('aria-selected')).toBe('true');
    expect(shown(ui.toolbar.querySelector('.dx-table-panel')!, ui.toolbar)).toBe(true);
    ui.tab('表格版面配置').click();
    await ui.flush();
    const rows = () => {
      let count = 0;
      ui.editor.view!.state.doc.descendants((node) => {
        if (node.type.name === 'table_row') count++;
      });
      return count;
    };
    expect(rows()).toBe(2);
    byTitle(ui.toolbar, '在下方插入列').click();
    await ui.flush();
    expect(rows()).toBe(3);
    // Out of the table: the table tabs go, back to 常用.
    const view = ui.editor.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    await ui.flush();
    expect(ui.tabs().map((t) => t.textContent?.trim())).toEqual(TABS);
    expect(ui.tab('常用').getAttribute('aria-selected')).toBe('true');
    ui.done();
  });

  it("版面配置 › 邊界 / 方向 / 大小 change the page like Word's galleries", async () => {
    const ui = await mount();
    ui.tab('版面配置').click();
    await ui.flush();
    const margins = byTitle(ui.toolbar, '邊界');
    expect(margins.getAttribute('aria-haspopup')).toBe('menu');
    margins.click();
    await ui.flush();
    expect(margins.getAttribute('aria-expanded')).toBe('true');
    const items = Array.from(ui.toolbar.querySelectorAll<HTMLButtonElement>('[role="menu"] [role^="menuitem"]'));
    expect(items.map((b) => b.querySelector('b')?.textContent ?? b.textContent?.trim())).toEqual(['標準', '窄', '適中', '寬', '自訂邊界…']);
    items[1].click();
    await ui.flush();
    expect(ui.toolbar.querySelector('[role="menu"]')).toBeNull();
    let page = ui.editor.cursorSection().page;
    expect([page.marginTop, page.marginBottom, page.marginLeft, page.marginRight]).toEqual([720, 720, 720, 720]);

    byTitle(ui.toolbar, '方向').click();
    await ui.flush();
    const landscape = Array.from(ui.toolbar.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')).find((b) => b.textContent?.trim() === '橫向')!;
    expect(landscape.getAttribute('aria-checked')).toBe('false');
    landscape.click();
    await ui.flush();
    page = ui.editor.cursorSection().page;
    expect(page.width).toBeGreaterThan(page.height);

    // The keyboard: the menu opens on its first item, Escape closes it back on its button.
    const size = byTitle(ui.toolbar, '大小（紙張）');
    size.focus();
    size.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 0 }));
    await ui.flush();
    const first = document.activeElement as HTMLElement;
    expect(first.getAttribute('role')).toBe('menuitemradio');
    expect(first.getAttribute('aria-checked')).toBe('true'); // A4
    key(first, 'ArrowDown');
    expect(document.activeElement).not.toBe(first);
    key(document.activeElement!, 'Escape');
    await ui.flush();
    expect(ui.toolbar.querySelector('[role="menu"]')).toBeNull();
    expect(document.activeElement).toBe(size);
    ui.done();
  });

  it('摺疊功能區: only the tabs; a tab shows its tools until one is used; remembered', async () => {
    const ui = await mount();
    const collapse = ui.host.querySelector('.dx-collapse') as HTMLButtonElement;
    expect(collapse.getAttribute('aria-pressed')).toBe('false');
    collapse.click();
    await ui.flush();
    expect(localStorage.getItem(KEY)).toBe('1');
    const panel = ui.toolbar.closest('[role="tabpanel"]') as HTMLElement;
    expect(panel.style.display).toBe('none');
    ui.tab('常用').click();
    await ui.flush();
    expect(panel.style.display).not.toBe('none');
    byTitle(ui.toolbar, '粗體 (Ctrl+B)').click();
    await ui.flush();
    expect(panel.style.display).toBe('none');
    ui.done();

    const again = await mount();
    const panel2 = again.toolbar.closest('[role="tabpanel"]') as HTMLElement;
    expect(panel2.style.display).toBe('none');
    // Alt+F10 still reaches the tools.
    const view = again.editor.view!;
    view.focus();
    key(view.dom, 'F10', { altKey: true });
    await again.flush();
    expect(panel2.style.display).not.toBe('none');
    expect(again.toolbar.contains(document.activeElement)).toBe(true);
    // Ctrl+F1 pins it again.
    key(view.dom, 'F1', { ctrlKey: true });
    await again.flush();
    expect(localStorage.getItem(KEY)).toBe('0');
    again.done();
  });

  it('works without browser storage (full ribbon, 摺疊功能區 still works)', async () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('blocked'); });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('blocked'); });
    const ui = await mount();
    const collapse = ui.host.querySelector('.dx-collapse') as HTMLButtonElement;
    expect(collapse.getAttribute('aria-pressed')).toBe('false');
    collapse.click();
    await ui.flush();
    expect(collapse.getAttribute('aria-pressed')).toBe('true');
    ui.done();
  });

  it("檔案 is the page's own menu when it has one; viewing only shows 檔案 and 檢視", async () => {
    let files = 0;
    const ui = await mount({ fileMenu: true, onFile: () => files++ });
    (ui.host.querySelector('.dx-file') as HTMLButtonElement).click();
    await ui.flush();
    expect(files).toBe(1);
    ui.done();

    const view = await mount({ fileMenu: true, editable: false });
    expect(view.host.querySelector('.dx-file')).not.toBeNull();
    expect(view.tabs().map((t) => t.textContent?.trim())).toEqual(['檢視']);
    expect(view.tab('檢視').getAttribute('aria-selected')).toBe('true');
    view.done();

    // Without a page menu, 檔案 has 列印 and 下載 itself.
    const plain = await mount();
    (plain.host.querySelector('.dx-file') as HTMLButtonElement).click();
    await plain.flush();
    const items = Array.from(plain.host.querySelectorAll('[role="menu"][aria-label="檔案"] [role="menuitem"]')).map((b) => b.textContent?.trim());
    expect(items).toEqual(['列印', '下載']);
    plain.done();
  });
});

describe('GOV-ISSUE-013: the toolbar is one Tab stop', () => {
  const tabStops = (tb: HTMLElement) =>
    Array.from(tb.querySelectorAll<HTMLElement>(CONTROLS)).filter((el) => el.tabIndex >= 0 && !(el as HTMLButtonElement).disabled && !(el as HTMLInputElement).hidden);

  it('only one control has tabindex 0; arrow keys, Home and End move between controls', async () => {
    const ui = await mount();
    expect(ui.toolbar.getAttribute('aria-label')).toBe('文件格式工具列');
    const items = reachable(ui.toolbar);
    expect(items.length).toBeGreaterThan(20);
    expect(tabStops(ui.toolbar)).toEqual([items[0]]);
    // 復原 / 取消復原 are disabled in a new document: not in the way.
    expect(items[0]).toBe(byTitle(ui.toolbar, '貼上 (Ctrl+V)'));

    items[0].focus();
    let e = key(items[0], 'ArrowRight');
    expect(e.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(items[1]);
    await ui.flush();
    expect(tabStops(ui.toolbar)).toEqual([items[1]]);
    key(items[1], 'ArrowLeft');
    expect(document.activeElement).toBe(items[0]);
    key(items[0], 'ArrowLeft'); // wraps around
    expect(document.activeElement).toBe(items[items.length - 1]);
    key(document.activeElement!, 'ArrowRight');
    expect(document.activeElement).toBe(items[0]);
    key(items[0], 'End');
    expect(document.activeElement).toBe(items[items.length - 1]);
    key(document.activeElement!, 'Home');
    expect(document.activeElement).toBe(items[0]);
    // A select moves on with Left / Right too (Up / Down still change its value).
    const style = byTitle(ui.toolbar, '段落樣式');
    style.focus();
    e = key(style, 'ArrowRight');
    expect(e.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(items[items.indexOf(style) + 1]);
    await ui.flush();
    expect(tabStops(ui.toolbar)).toHaveLength(1);
    ui.done();
  });

  it('a text box keeps Left / Right / Home / End for its own text', async () => {
    const ui = await mount();
    ui.editor.run(insertTable(2, 2, ui.editor.contentWidth));
    intoTable(ui.editor);
    await ui.flush();
    ui.tab('表格設計').click();
    await ui.flush();
    const height = ui.toolbar.querySelector<HTMLInputElement>('.dx-table-panel input[type="text"]')!;
    expect(height).not.toBeNull();
    height.focus();
    for (const k of ['ArrowLeft', 'ArrowRight', 'Home', 'End']) {
      const e = key(height, k);
      expect(e.defaultPrevented, k).toBe(false);
      expect(document.activeElement, k).toBe(height);
    }
    await ui.flush();
    // It became the toolbar's Tab stop; still only one.
    expect(tabStops(ui.toolbar)).toEqual([height]);
    ui.done();
  });

  it("another tab's tools are skipped; the tab stop moves to the tab shown", async () => {
    const ui = await mount();
    let items = reachable(ui.toolbar);
    expect(items.map((el) => el.title)).toContain('粗體 (Ctrl+B)');
    expect(items.map((el) => el.title)).not.toContain('插入表格');
    ui.tab('插入').click();
    await ui.flush();
    items = reachable(ui.toolbar);
    expect(items.map((el) => el.title)).toContain('插入表格');
    expect(items.map((el) => el.title)).not.toContain('粗體 (Ctrl+B)');
    expect(tabStops(ui.toolbar)).toEqual([items[0]]);
    ui.done();
  });

  it('Alt+F10 in the text goes to the toolbar; Escape goes back to the text', async () => {
    const ui = await mount();
    const view = ui.editor.view!;
    view.focus();
    const e = key(view.dom, 'F10', { altKey: true });
    expect(e.defaultPrevented).toBe(true);
    const first = reachable(ui.toolbar)[0];
    expect(document.activeElement).toBe(first);
    // Move to 粗體, leave with Escape, come back to 粗體.
    const bold = byTitle(ui.toolbar, '粗體 (Ctrl+B)');
    bold.focus();
    await ui.flush();
    key(bold, 'Escape');
    expect(document.activeElement).toBe(view.dom);
    key(view.dom, 'F10', { altKey: true });
    expect(document.activeElement).toBe(bold);
    // Escape from a select and from a colour input goes back too.
    const size = byTitle(ui.toolbar, '字型大小');
    size.focus();
    key(size, 'Escape');
    expect(document.activeElement).toBe(view.dom);
    ui.done();
  });

  it('the table size grid: arrows pick the size, Escape closes it and returns to 插入表格', async () => {
    const ui = await mount();
    ui.tab('插入').click();
    await ui.flush();
    const table = byTitle(ui.toolbar, '插入表格');
    table.focus();
    table.click();
    await ui.flush();
    const cells = Array.from(ui.toolbar.querySelectorAll<HTMLButtonElement>('.dx-cell'));
    expect(cells).toHaveLength(64);
    // Only one cell is a Tab stop, so Tab does not walk through 64 cells.
    expect(cells.filter((c) => c.tabIndex >= 0)).toHaveLength(1);
    const start = document.activeElement as HTMLElement;
    expect(start.classList.contains('dx-cell')).toBe(true);
    key(start, 'ArrowRight');
    key(document.activeElement!, 'ArrowDown');
    expect((document.activeElement as HTMLElement).getAttribute('aria-label')).toBe('2 列 2 欄');
    key(document.activeElement!, 'Escape');
    await ui.flush();
    expect(ui.toolbar.querySelector('.dx-cell')).toBeNull();
    expect(document.activeElement).toBe(table);
    ui.done();
  });
});

describe('the format toggles say whether they are on (aria-pressed, not colour alone)', () => {
  const TOGGLES = ['粗體 (Ctrl+B)', '斜體 (Ctrl+I)', '底線 (Ctrl+U)', '刪除線', '上標', '下標',
    '靠左對齊', '置中', '靠右對齊', '左右對齊', '項目符號清單', '編號清單'];

  it('every toggle has aria-pressed, in step with its highlighted state', async () => {
    const ui = await mount();
    const view = ui.editor.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1)));
    await ui.flush();
    for (const t of TOGGLES) {
      const b = byTitle(ui.toolbar, t);
      expect(b.getAttribute('aria-pressed'), t).toBe(b.classList.contains('on') ? 'true' : 'false');
    }
    // Plain left-aligned text: only 靠左對齊 is on.
    expect(byTitle(ui.toolbar, '靠左對齊').getAttribute('aria-pressed')).toBe('true');
    expect(byTitle(ui.toolbar, '粗體 (Ctrl+B)').getAttribute('aria-pressed')).toBe('false');

    for (const t of ['粗體 (Ctrl+B)', '斜體 (Ctrl+I)', '底線 (Ctrl+U)', '刪除線', '上標', '置中', '項目符號清單']) {
      byTitle(ui.toolbar, t).click();
      await ui.flush();
      expect(byTitle(ui.toolbar, t).getAttribute('aria-pressed'), t).toBe('true');
    }
    expect(byTitle(ui.toolbar, '靠左對齊').getAttribute('aria-pressed')).toBe('false');
    expect(byTitle(ui.toolbar, '編號清單').getAttribute('aria-pressed')).toBe('false');
    // 下標 replaces 上標.
    byTitle(ui.toolbar, '下標').click();
    await ui.flush();
    expect(byTitle(ui.toolbar, '下標').getAttribute('aria-pressed')).toBe('true');
    expect(byTitle(ui.toolbar, '上標').getAttribute('aria-pressed')).toBe('false');
    // Off again.
    byTitle(ui.toolbar, '粗體 (Ctrl+B)').click();
    await ui.flush();
    expect(byTitle(ui.toolbar, '粗體 (Ctrl+B)').getAttribute('aria-pressed')).toBe('false');
    for (const t of TOGGLES) {
      const b = byTitle(ui.toolbar, t);
      expect(b.getAttribute('aria-pressed'), t).toBe(b.classList.contains('on') ? 'true' : 'false');
    }
    ui.done();
  });
});
