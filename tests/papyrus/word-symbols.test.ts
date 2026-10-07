// @vitest-environment jsdom
// Word's 插入 › 符號 (Insert › Symbol), with 圍繞字元 through the enclosed characters (①, ㈠, ㊣ …).
// Word inserts the symbol as text at the selection: it replaces the selection and takes the
// formatting of the text at the cursor, as typing does, and it is an undo step of its own. A
// symbol Latin and CJK fonts share (①, ※, ±, →…) is written in its own run with
// w:rFonts/@w:hint="eastAsia", as zh-TW Word writes it, so Word draws it in the East Asian font
// (標楷體) rather than the Latin one (Times New Roman). The ribbon's 符號 button opens a gallery of
// 20 symbols, the recently used ones first; 「其他符號(M)…」 opens the 符號 dialog (a subset list, a
// grid of symbols, the character code, the recently used symbols), which stays open, with the
// keyboard, after 插入 so several symbols can be inserted.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick, reactive } from 'vue';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { withEastAsiaHint } from '../../src/papyrus/docx/props';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { insertImage, insertSymbol } from '../../src/papyrus/editor/commands';
import {
  MAX_RECENT_SYMBOLS, RIBBON_SYMBOLS, SYMBOL_CATEGORIES, isSymbol, menuSymbols, needsEastAsiaHint, parseSymbolCode, symbolCode,
} from '../../src/papyrus/editor/symbols';
import {
  RECENT_SYMBOLS_KEY, pushRecentSymbol, readRecentSymbols, rememberSymbol,
} from '../../src/papyrus/editor/recentSymbols';
import { insertSymbolAndRemember } from '../../src/papyrus/editor/symbolActions';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import SymbolDialog from '../../src/papyrus/vue/SymbolDialog.vue';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const editors: DocxEditor[] = [];
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy();
  try {
    localStorage.clear();
  } catch {
    // storage unavailable
  }
});

/** 標楷體 text in bold 16 pt, Latin in Times New Roman (as 公文 templates set it). */
const RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="DFKai-SB"/><w:b/><w:sz w:val="32"/></w:rPr>';
const RUN = `<w:r>${RPR}<w:t>公文</w:t></w:r>`;
const HINT = 'w:hint="eastAsia"';

async function open(body: string, notices: string[] = []) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  editors.push(editor);
  return editor;
}
async function part(editor: DocxEditor, name = 'word/document.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(await editor.save());
  return zip.file(name)!.async('string');
}
function posOf(view: EditorView, text: string, n = 0): number {
  let found = -1;
  view.state.doc.descendants((node, pos) => {
    if (found >= 0) return false;
    if (node.isText && node.text!.includes(text)) found = pos + node.text!.indexOf(text) + n;
    return true;
  });
  if (found < 0) throw new Error(`no ${text}`);
  return found;
}
function select(view: EditorView, from: number, to = from) {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
}
/** The w:r elements whose text contains `text`. */
const runsWith = (xml: string, text: string) => (xml.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? []).filter((r) => r.includes(text));

describe('插入符號: the command', () => {
  it('writes the symbol at the cursor with the run formatting there, in its own run with w:hint="eastAsia"', async () => {
    const editor = await open(`<w:p>${RUN}</w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p>`);
    const view = editor.view!;
    select(view, posOf(view, '公文', 1));
    expect(editor.run(insertSymbol('※'))).toBe(true);
    const xml = await part(editor);
    const [before] = runsWith(xml, '<w:t>公</w:t>');
    const [symbol] = runsWith(xml, '<w:t>※</w:t>');
    const [afterRun] = runsWith(xml, '<w:t>文</w:t>');
    expect(symbol).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="DFKai-SB" w:hint="eastAsia"/>');
    expect(symbol).toContain('<w:b/>');
    expect(symbol).toContain('<w:sz w:val="32"/>');
    // The text around it keeps its run properties exactly, without the hint.
    expect(before).toContain(RPR);
    expect(afterRun).toContain(RPR);
    expect(xml).toContain('<w:t>第二段</w:t>'); // untouched
  });

  it('text typed right after the symbol does not take the hint', async () => {
    const editor = await open(`<w:p>${RUN}</w:p>`);
    const view = editor.view!;
    select(view, posOf(view, '公文', 2));
    editor.run(insertSymbol('①'));
    view.dispatch(view.state.tr.insertText('號'));
    const [typed] = runsWith(await part(editor), '號');
    expect(typed).not.toContain(HINT);
    expect(typed).toContain('<w:b/>');
  });

  it('a plain run with an East Asian font from the defaults gets an rPr with the hint only; CJK and ASCII symbols need no hint', async () => {
    const editor = await open('<w:p><w:r><w:t>甲乙</w:t></w:r></w:p>');
    const view = editor.view!;
    const styles = editor.model.styles; // the blank template's defaults: w:eastAsia="PMingLiU"
    select(view, posOf(view, '甲乙', 1));
    editor.run(insertSymbol('→', styles));
    editor.run(insertSymbol('、', styles));
    editor.run(insertSymbol('<', styles));
    const xml = await part(editor);
    expect(runsWith(xml, '→')[0]).toContain('<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>');
    expect(xml).toContain('<w:t>、&lt;乙</w:t>');
    for (const ch of ['①', '±', '×', '§', '°', 'ˇ', '“', '—']) expect(needsEastAsiaHint(ch), ch).toBe(true);
    for (const ch of ['、', '㊣', 'ㄅ', '＄', 'A', '<', '¢', '£', '¥', '©', '«', '®', 'µ', '»', '⃐']) expect(needsEastAsiaHint(ch), ch).toBe(false);
  });

  it('no hint where the text has no East Asian font or language; the run language alone is enough', async () => {
    const editor = await open(
      '<w:p><w:r><w:t>Latin</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:lang w:val="en-US" w:eastAsia="zh-TW"/></w:rPr><w:t>中文</w:t></w:r></w:p>',
    );
    const view = editor.view!;
    select(view, posOf(view, 'Latin', 5));
    editor.run(insertSymbol('①')); // no styles given: only the run itself counts, and it has nothing
    select(view, posOf(view, '中文', 2));
    editor.run(insertSymbol('②'));
    const xml = await part(editor);
    expect(xml).toContain('<w:t>Latin①</w:t>');
    expect(runsWith(xml, '②')[0]).toContain('<w:rFonts w:hint="eastAsia"/>');
    expect(runsWith(xml, '②')[0]).toContain('<w:lang w:val="en-US" w:eastAsia="zh-TW"/>');
  });

  it('withEastAsiaHint keeps the rest of w:rPr in schema order and leaves a hinted rPr as it is', () => {
    expect(withEastAsiaHint(null)).toBe('<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>');
    expect(withEastAsiaHint('<w:rPr><w:b/><w:lang w:eastAsia="zh-TW"/></w:rPr>')).toBe(
      '<w:rPr><w:rFonts w:hint="eastAsia"/><w:b/><w:lang w:eastAsia="zh-TW"/></w:rPr>',
    );
    const hinted = '<w:rPr><w:rFonts w:eastAsia="DFKai-SB" w:hint="eastAsia"/></w:rPr>';
    expect(withEastAsiaHint(hinted)).toBe(hinted);
  });

  it('replaces the selection', async () => {
    const editor = await open(`<w:p>${RUN}</w:p>`);
    const view = editor.view!;
    select(view, posOf(view, '公文'), posOf(view, '公文') + 2);
    editor.run(insertSymbol('①'));
    expect(view.state.doc.textContent).toBe('①');
    expect(runsWith(await part(editor), '①')[0]).toContain('<w:b/>');
  });

  it('replaces a selected picture', async () => {
    const editor = await open('<w:p><w:r><w:t>前後</w:t></w:r></w:p>');
    const view = editor.view!;
    select(view, posOf(view, '前後', 1));
    editor.run(insertImage('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 10, 10, 500));
    let image = -1;
    view.state.doc.descendants((node, pos) => {
      if (node.type.name === 'image') image = pos;
    });
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, image)));
    expect(editor.run(insertSymbol('※'))).toBe(true);
    let images = 0;
    view.state.doc.descendants((node) => {
      if (node.type.name === 'image') images++;
    });
    expect(images).toBe(0);
    expect(view.state.doc.textContent).toBe('前※後');
    expect(await part(editor)).not.toContain('<w:drawing');
  });

  const TABLE =
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
    '<w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  function selectCells(view: EditorView) {
    const cells: number[] = [];
    view.state.doc.descendants((node, pos) => {
      if (node.type.name === 'table_cell') cells.push(pos);
    });
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cells[0], cells[1])));
  }
  const cellXml = (xml: string) => xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g)!;

  it('over selected table cells: their text goes and the symbol is written in the first cell', async () => {
    const editor = await open(TABLE);
    const view = editor.view!;
    selectCells(view);
    expect(editor.run(insertSymbol('※'))).toBe(true);
    const [c1, c2] = cellXml(await part(editor));
    expect(c1).toContain('<w:t>※</w:t>');
    expect(c1).not.toContain('甲');
    expect(c2).not.toContain('乙');
    expect(c2.match(/<w:p\b/g)).toHaveLength(1); // no paragraph added
    undo(view.state, view.dispatch);
    expect(view.state.doc.textContent).toBe('甲乙');
  });

  it('over selected table cells with 追蹤修訂: tracked deletions and a tracked insertion, no new paragraph', async () => {
    const notices: string[] = [];
    const editor = await open(TABLE, notices);
    const view = editor.view!;
    editor.setTrackChanges(true);
    selectCells(view);
    expect(editor.run(insertSymbol('※'))).toBe(true);
    expect(notices).toEqual([]);
    const [c1, c2] = cellXml(await part(editor));
    expect(c1).toMatch(/<w:del\b[^>]*>\s*<w:r>\s*<w:delText>甲<\/w:delText>/);
    expect(c1).toMatch(/<w:ins\b[^>]*>\s*<w:r>(?:(?!<\/w:ins>)[\s\S])*<w:t>※<\/w:t>/);
    expect(c2).toMatch(/<w:delText>乙<\/w:delText>/);
    expect(c2).not.toContain('※');
    expect(c2.match(/<w:p\b/g)).toHaveLength(1);
  });

  it('is its own undo step, before and after: typing right before or right after stays separate', async () => {
    const editor = await open('<w:p><w:r><w:t>甲</w:t></w:r></w:p>');
    const view = editor.view!;
    select(view, posOf(view, '甲', 1));
    view.dispatch(view.state.tr.insertText('乙'));
    editor.run(insertSymbol('○'));
    editor.run(insertSymbol('●'));
    view.dispatch(view.state.tr.insertText('丙'));
    expect(view.state.doc.textContent).toBe('甲乙○●丙');
    undo(view.state, view.dispatch);
    expect(view.state.doc.textContent).toBe('甲乙○●');
    undo(view.state, view.dispatch);
    expect(view.state.doc.textContent).toBe('甲乙○');
    undo(view.state, view.dispatch);
    expect(view.state.doc.textContent).toBe('甲乙');
  });

  it('with 追蹤修訂 on, it is recorded as an insertion (w:ins)', async () => {
    const editor = await open('<w:p><w:r><w:t>公文</w:t></w:r></w:p>');
    const view = editor.view!;
    editor.setTrackChanges(true);
    select(view, posOf(view, '公文', 1));
    expect(editor.run(insertSymbol('㊣', editor.model.styles))).toBe(true);
    expect(editor.run(insertSymbol('①', editor.model.styles))).toBe(true);
    const xml = await part(editor);
    expect(xml).toMatch(/<w:ins\b[^>]*w:author="[^"]+"[^>]*>\s*<w:r>(?:(?!<\/w:ins>)[\s\S])*<w:t>㊣<\/w:t>/);
    expect(xml).toMatch(/<w:ins\b[^>]*>(?:(?!<\/w:ins>)[\s\S])*<w:r><w:rPr><w:rFonts w:hint="eastAsia"\/><\/w:rPr><w:t>①<\/w:t>/);
    expect(xml).toContain('<w:t>公</w:t>');
    expect(xml).toContain('<w:t>文</w:t>');
  });

  it('works in a header being edited', async () => {
    const editor = await open('<w:p><w:r><w:t>本文</w:t></w:r></w:p>');
    const body = editor.view!;
    editor.editHeaderFooter('header', 0);
    const hv = editor.activeView!;
    expect(hv).not.toBe(body);
    hv.dispatch(hv.state.tr.insertText('頁首'));
    expect(editor.run(insertSymbol('◎'))).toBe(true);
    expect(hv.state.doc.textContent).toBe('頁首◎');
    expect(body.state.doc.textContent).toBe('本文');
    editor.closeHeaderFooter();
    const zip = await JSZip.loadAsync(await editor.save());
    const header = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
    expect(await zip.file(header)!.async('string')).toContain('◎');
  });

  it('refuses what is not one valid character', async () => {
    const editor = await open('<w:p><w:r><w:t>公文</w:t></w:r></w:p>');
    const view = editor.view!;
    const before = view.state.doc;
    for (const bad of ['', 'ab', '\u0007', '\uD800', '￿']) expect(editor.run(insertSymbol(bad)), JSON.stringify(bad)).toBe(false);
    expect(view.state.doc).toBe(before);
  });

  it('run(cmd, { focus: false }) leaves the keyboard where it is; the default puts it in the text', async () => {
    const editor = await open('<w:p><w:r><w:t>公文</w:t></w:r></w:p>');
    const view = editor.view!;
    const other = document.createElement('button');
    document.body.append(other);
    other.focus();
    editor.run(insertSymbol('※'), { focus: false });
    expect(document.activeElement).toBe(other);
    editor.run(insertSymbol('※'));
    expect(view.dom.contains(document.activeElement) || document.activeElement === view.dom).toBe(true);
    other.remove();
  });
});

describe('插入符號: the character code (字元代碼)', () => {
  it('reads a hexadecimal code, with or without U+ and leading zeros', () => {
    expect(parseSymbolCode('2605')).toEqual({ ok: true, char: '★' });
    expect(parseSymbolCode(' u+2460 ')).toEqual({ ok: true, char: '①' });
    expect(parseSymbolCode('32a3')).toEqual({ ok: true, char: '㊣' });
    expect(parseSymbolCode('1F600')).toEqual({ ok: true, char: '\u{1F600}' });
    expect(parseSymbolCode('0000002605')).toEqual({ ok: true, char: '★' });
    expect(symbolCode('①')).toBe('U+2460');
    expect(symbolCode('\u{1F600}')).toBe('U+1F600');
  });

  it('refuses what is not a character, in Chinese', () => {
    const error = (s: string) => {
      const r = parseSymbolCode(s);
      expect(r.ok, s).toBe(false);
      return r.ok ? '' : r.error;
    };
    expect(error('')).toMatch(/請輸入/);
    expect(error('26G5')).toMatch(/16 進位/);
    expect(error('110000')).toMatch(/超出/);
    expect(error('0001100000')).toMatch(/超出/);
    expect(error('D800')).toMatch(/代理/);
    expect(error('DFFF')).toMatch(/代理/);
    expect(error('0007')).toMatch(/控制字元/);
    expect(error('0000')).toMatch(/控制字元/);
    expect(error('7F')).toMatch(/控制字元/);
    expect(error('85')).toMatch(/控制字元/);
    expect(error('FFFE')).toMatch(/非字元/);
    expect(error('FDD0')).toMatch(/非字元/);
    expect(error('1FFFF')).toMatch(/非字元/);
  });

  it('isSymbol: exactly one valid character', () => {
    expect(isSymbol('※')).toBe(true);
    expect(isSymbol('\u{1F600}')).toBe(true);
    for (const bad of ['', '※※', '\n', '\uD800', '￾', 1, null, {}]) expect(isSymbol(bad), JSON.stringify(bad)).toBe(false);
  });
});

describe('插入符號: recently used (最近使用過的符號)', () => {
  it('most recent first, no duplicates, at most 20', () => {
    let list: string[] = [];
    for (const ch of ['①', '②', '③', '①']) list = pushRecentSymbol(list, ch);
    expect(list).toEqual(['①', '③', '②']);
    for (const ch of SYMBOL_CATEGORIES[1].chars.slice(0, 30)) list = pushRecentSymbol(list, ch);
    expect(list).toHaveLength(MAX_RECENT_SYMBOLS);
    expect(MAX_RECENT_SYMBOLS).toBe(20);
    expect(list[0]).toBe(SYMBOL_CATEGORIES[1].chars[29]);
  });

  it(`is kept in localStorage (${RECENT_SYMBOLS_KEY})`, () => {
    expect(RECENT_SYMBOLS_KEY).toBe('papyrus.recentSymbols');
    expect(readRecentSymbols()).toEqual([]);
    rememberSymbol('※');
    rememberSymbol('℃');
    expect(JSON.parse(localStorage.getItem(RECENT_SYMBOLS_KEY)!)).toEqual(['℃', '※']);
    expect(readRecentSymbols()).toEqual(['℃', '※']);
  });

  it('ignores garbage in storage', () => {
    for (const junk of ['not json', '{"a":1}', '42', 'null', '"※"']) {
      localStorage.setItem(RECENT_SYMBOLS_KEY, junk);
      expect(readRecentSymbols(), junk).toEqual([]);
    }
    localStorage.setItem(RECENT_SYMBOLS_KEY, JSON.stringify(['※', 'ab', 7, null, '\u0000', '\uD800', '※', '①', ...Array(40).fill('②')]));
    expect(readRecentSymbols()).toEqual(['※', '①', '②']);
    expect(rememberSymbol('③')).toEqual(['③', '※', '①', '②']);
  });

  it('works on when storage throws', () => {
    const throwing = {
      getItem: () => {
        throw new Error('denied');
      },
      setItem: () => {
        throw new Error('full');
      },
    };
    expect(readRecentSymbols(throwing)).toEqual([]);
    expect(rememberSymbol('※', throwing)).toEqual(['※']);
    expect(readRecentSymbols(null)).toEqual([]);
  });

  it('a symbol that did not go in (refused while tracking changes) is not remembered', async () => {
    const notices: string[] = [];
    const editor = await open('<w:p><w:r><w:t>前</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>後</w:t></w:r></w:p>', notices);
    const view = editor.view!;
    editor.setTrackChanges(true);
    select(view, posOf(view, '前'), posOf(view, '後') + 1); // over a page break, which can't be deleted tracked
    const before = view.state.doc;
    expect(insertSymbolAndRemember(editor, '※')).toBe(false);
    expect(view.state.doc).toBe(before);
    expect(notices.length).toBeGreaterThan(0);
    expect(readRecentSymbols()).toEqual([]);
    select(view, posOf(view, '後', 1));
    expect(insertSymbolAndRemember(editor, '※')).toBe(true);
    expect(readRecentSymbols()).toEqual(['※']);
  });

  it('the ribbon gallery: 20 symbols, recently used first, then the defaults', () => {
    expect(RIBBON_SYMBOLS).toHaveLength(20);
    expect(menuSymbols([])).toEqual(RIBBON_SYMBOLS);
    const shown = menuSymbols(['㊣', '※']);
    expect(shown).toHaveLength(20);
    expect(shown.slice(0, 3)).toEqual(['㊣', '※', '○']);
    expect(new Set(shown).size).toBe(20);
  });
});

describe('插入符號: the symbol lists', () => {
  const cat = (label: string) => {
    const c = SYMBOL_CATEGORIES.find((x) => x.label === label);
    expect(c, label).toBeTruthy();
    return c!.chars;
  };
  /** The characters from `a` to `b` (code points), all in the list, in order. */
  const range = (list: readonly string[], a: string, b: string, count: number) => {
    const from = a.codePointAt(0)!;
    const to = b.codePointAt(0)!;
    expect(to - from + 1, `${a}–${b}`).toBe(count);
    const i = list.indexOf(a);
    expect(i, a).toBeGreaterThanOrEqual(0);
    for (let k = 0; k < count; k++) expect(list[i + k], `${a}+${k}`).toBe(String.fromCodePoint(from + k));
  };

  it('has Word-like subsets, each character valid and listed once', () => {
    expect(SYMBOL_CATEGORIES.map((c) => c.label)).toEqual([
      '標點符號', '帶圈及括號字元', '數學符號', '單位', '貨幣', '箭頭', '圖形與方塊', '羅馬數字', '希臘字母', '注音符號', '全形英數字',
    ]);
    for (const c of SYMBOL_CATEGORIES) {
      expect(new Set(c.chars).size, c.label).toBe(c.chars.length);
      for (const ch of c.chars) expect(isSymbol(ch), `${c.label} ${ch}`).toBe(true);
    }
    for (const ch of RIBBON_SYMBOLS) expect(isSymbol(ch)).toBe(true);
  });

  it('標點符號', () => {
    const list = cat('標點符號');
    for (const ch of '、。，；：？！「」『』（）〔〕【】《》〈〉…—～．‧') expect(list).toContain(ch);
  });

  it('帶圈及括號字元 (圍繞字元)', () => {
    const list = cat('帶圈及括號字元');
    range(list, '①', '⑳', 20);
    range(list, '㉑', '㉟', 15);
    range(list, '㊱', '㊿', 15);
    range(list, '⑴', '⒇', 20);
    range(list, '⒈', '⒛', 20);
    range(list, '㈠', '㈩', 10);
    range(list, '㊀', '㊉', 10);
    range(list, 'Ⓐ', 'Ⓩ', 26);
    range(list, 'ⓐ', 'ⓩ', 26);
    for (const ch of '㊣㊤㊥㊦㊧㊨㊞㊟㊕㊖㊗㊘㊙㊚㊛㊜㊝') expect(list).toContain(ch);
    expect(list).toHaveLength(20 + 15 + 15 + 20 + 20 + 10 + 10 + 26 + 26 + 17);
  });

  it('數學符號, 單位, 貨幣, 箭頭, 圖形與方塊', () => {
    for (const ch of '±×÷=≠≒≈≦≧<>∞√∑∫∵∴∠⊥∩∪∈⊂⊃°‰％′″') expect(cat('數學符號')).toContain(ch);
    for (const ch of '℃℉㎎㎏㎜㎝㎞㎡㏄㏎㏑㏒㏕㎥㎖㎗㎘') expect(cat('單位')).toContain(ch);
    for (const ch of '＄￥€£¢₩') expect(cat('貨幣')).toContain(ch);
    for (const ch of '→←↑↓↗↖↘↙⇒⇔↔↕') expect(cat('箭頭')).toContain(ch);
    for (const ch of '※○●◎□■△▲▽▼◇◆☆★♀♂☐☑☒✓✔✕✗♪♠♥♦♣☎☏') expect(cat('圖形與方塊')).toContain(ch);
  });

  it('羅馬數字, 希臘字母, 注音符號, 全形英數字', () => {
    range(cat('羅馬數字'), 'Ⅰ', 'Ⅻ', 12);
    range(cat('羅馬數字'), 'ⅰ', 'ⅹ', 10);
    const greek = cat('希臘字母');
    for (const ch of 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω') expect(greek).toContain(ch);
    range(cat('注音符號'), 'ㄅ', 'ㄩ', 37);
    for (const ch of 'ˊˇˋ˙') expect(cat('注音符號')).toContain(ch);
    range(cat('全形英數字'), '０', '９', 10);
    range(cat('全形英數字'), 'Ａ', 'Ｚ', 26);
    range(cat('全形英數字'), 'ａ', 'ｚ', 26);
  });
});

// ----- UI: the ribbon's 符號 gallery and the 符號 dialog -----

async function mount(props: Record<string, unknown> = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p>${RUN}</w:p><w:sectPr/></w:body></w:document>`);
  const src = await zip.generateAsync({ type: 'uint8array' });
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditor) => (editor = ed), ...props }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  const ed = editor as DocxEditor;
  const view = ed.view!;
  if (props.editable !== false) select(view, posOf(view, '公文', 1));
  const tab = (label: string) => Array.from(host.querySelectorAll<HTMLButtonElement>('[role="tab"]')).find((t) => t.textContent?.trim() === label)!;
  const button = () => host.querySelector('[role="toolbar"] [aria-label="符號"] button[aria-haspopup="menu"]') as HTMLButtonElement;
  const menu = () => host.querySelector('[role="toolbar"] .dx-menu[aria-label="符號"]') as HTMLElement | null;
  const dialog = () => document.querySelector('[role="dialog"][aria-labelledby="dx-sy-title"]') as HTMLElement | null;
  const openMenu = async () => {
    tab('插入').click();
    await flush();
    button().click();
    await flush();
    return menu()!;
  };
  return { host, editor: ed, view, flush, tab, button, menu, dialog, openMenu, done: () => { app.unmount(); host.remove(); } };
}
type Ui = Awaited<ReturnType<typeof mount>>;
/** Runs `body` on a mounted editor, unmounting it whatever happens. */
async function withUi(props: Record<string, unknown>, body: (ui: Ui) => Promise<void>) {
  const ui = await mount(props);
  try {
    await body(ui);
  } finally {
    ui.done();
  }
}
const key = (el: Element, k: string, init: KeyboardEventInit = {}) => {
  const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init });
  el.dispatchEvent(e);
  return e;
};
const text = (ui: { view: EditorView }) => ui.view.state.doc.textContent;
function pickSubset(select: HTMLSelectElement, label: string) {
  select.value = Array.from(select.options).find((o) => o.textContent?.trim() === label)!.value;
  select.dispatchEvent(new Event('change'));
}

describe('插入符號: ribbon (插入 › 符號)', () => {
  it('is the last group of 插入, a menu of 20 symbols and 其他符號(M)…; a click inserts and closes it', () =>
    withUi({}, async (ui) => {
      ui.tab('插入').click();
      await ui.flush();
      const panel = ui.button().closest('.dx-panel')!;
      const groups = Array.from(panel.querySelectorAll('.dx-rgroup'));
      expect(groups[groups.length - 1].getAttribute('aria-label')).toBe('符號');
      const b = ui.button();
      expect(b.classList.contains('dx-big')).toBe(true);
      expect(b.getAttribute('aria-expanded')).toBe('false');
      b.click();
      await ui.flush();
      expect(b.getAttribute('aria-expanded')).toBe('true');
      const menu = ui.menu()!;
      expect(menu.getAttribute('role')).toBe('menu');
      expect(menu.classList.contains('dx-popup')).toBe(true);
      const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'));
      const symbols = items.filter((i) => i.classList.contains('dx-sym'));
      expect(symbols.map((s) => s.textContent?.trim())).toEqual(RIBBON_SYMBOLS);
      expect(items[items.length - 1].textContent?.trim()).toBe('其他符號(M)…');
      expect(document.activeElement).toBe(symbols[0]);
      // Arrow keys move in the 5-wide gallery.
      key(symbols[0], 'ArrowRight');
      expect(document.activeElement).toBe(symbols[1]);
      key(symbols[1], 'ArrowDown');
      expect(document.activeElement).toBe(symbols[6]);
      symbols.find((s) => s.textContent?.trim() === '①')!.click();
      await ui.flush();
      expect(ui.menu()).toBeNull();
      expect(text(ui)).toBe('公①文');
      expect(ui.view.dom.contains(document.activeElement)).toBe(true); // back in the text
      expect(readRecentSymbols()).toEqual(['①']);
      // Opened again: the symbol just used comes first.
      ui.button().click();
      await ui.flush();
      const again = Array.from(ui.menu()!.querySelectorAll('.dx-sym')).map((s) => s.textContent?.trim());
      expect(again[0]).toBe('①');
      expect(again).toHaveLength(20);
      key(document.activeElement!, 'Escape');
      await ui.flush();
      expect(ui.menu()).toBeNull();
      expect(document.activeElement).toBe(ui.button());
    }));

  it('ArrowDown from the last row goes to 其他符號 and ArrowUp comes back; M opens the dialog', () =>
    withUi({}, async (ui) => {
      const menu = await ui.openMenu();
      const symbols = Array.from(menu.querySelectorAll<HTMLElement>('.dx-sym'));
      const more = menu.querySelector<HTMLElement>('.dx-symmore')!;
      symbols[17].focus();
      key(symbols[17], 'ArrowDown');
      expect(document.activeElement).toBe(more);
      key(more, 'ArrowUp');
      expect(document.activeElement).toBe(symbols[15]); // the last row's first symbol
      key(symbols[15], 'ArrowUp');
      expect(document.activeElement).toBe(symbols[10]);
      symbols[2].focus();
      key(symbols[2], 'ArrowUp'); // from the first row, up to 其他符號
      expect(document.activeElement).toBe(more);
      key(more, 'ArrowDown');
      expect(document.activeElement).toBe(symbols[0]);
      const e = key(symbols[0], 'm');
      expect(e.defaultPrevented).toBe(true);
      await ui.flush();
      expect(ui.menu()).toBeNull();
      expect(ui.dialog()).not.toBeNull();
    }));

  it('in read-only mode: the button is disabled and neither the gallery nor the dialog opens', () =>
    withUi({ editable: false, fileMenu: true }, async (ui) => {
      expect(ui.button()).not.toBeNull();
      expect(ui.button().disabled).toBe(true);
      ui.button().removeAttribute('disabled');
      ui.button().click();
      await ui.flush();
      expect(ui.menu()).toBeNull();
      expect(ui.dialog()).toBeNull();
    }));

  it('the 符號 dialog shows only while editing, and closes when the document becomes read-only', async () => {
    const editor = await open('<w:p><w:r><w:t>公文</w:t></w:r></w:p>');
    const props = reactive({ open: true, editable: false });
    let closed = 0;
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(SymbolDialog, { editor, open: props.open, editable: props.editable, onClose: () => closed++ }) });
    app.mount(host);
    try {
      await nextTick();
      expect(document.querySelector('[aria-labelledby="dx-sy-title"]')).toBeNull();
      props.editable = true;
      await nextTick();
      expect(document.querySelector('[aria-labelledby="dx-sy-title"]')).not.toBeNull();
      props.editable = false;
      await nextTick();
      expect(closed).toBe(1);
    } finally {
      app.unmount();
      host.remove();
    }
  });
});

describe('插入符號: the 符號 dialog', () => {
  async function openDialog(ui: Ui) {
    const menu = await ui.openMenu();
    menu.querySelector<HTMLButtonElement>('.dx-symmore')!.click();
    await ui.flush();
    const dialog = ui.dialog()!;
    expect(dialog).not.toBeNull();
    const grid = () => dialog.querySelector('[role="grid"][aria-label="符號"]') as HTMLElement;
    const cells = () => Array.from(grid().querySelectorAll<HTMLElement>('[role="gridcell"]'));
    const recent = () => dialog.querySelector('[role="grid"][aria-labelledby="dx-sy-recent-label"]') as HTMLElement | null;
    const recentCells = () => Array.from(recent()?.querySelectorAll<HTMLElement>('[role="gridcell"]') ?? []);
    const subset = () => dialog.querySelector('select') as HTMLSelectElement;
    const code = () => dialog.querySelector('input[type="text"]') as HTMLInputElement;
    const insertButton = () => dialog.querySelector('.dx-primary') as HTMLButtonElement;
    return { dialogEl: dialog, grid, cells, recent, recentCells, subset, code, insertButton };
  }
  const dialogTest = (name: string, body: (ui: Ui, d: Awaited<ReturnType<typeof openDialog>>) => Promise<void>) =>
    it(name, () => withUi({}, async (ui) => body(ui, await openDialog(ui))));

  dialogTest('opens from 其他符號…, a modal dialog titled 符號, the focus in the grid', async (ui, d) => {
    expect(ui.menu()).toBeNull();
    expect(d.dialogEl.getAttribute('aria-modal')).toBe('true');
    expect(document.getElementById('dx-sy-title')!.textContent?.trim()).toBe('符號');
    expect(d.insertButton().textContent?.trim()).toBe('插入');
    expect(Array.from(d.dialogEl.querySelectorAll('button')).some((b) => b.textContent?.trim() === '關閉')).toBe(true);
    // One Tab stop in the grid (roving tabindex), and it has the focus.
    const stops = d.cells().filter((c) => c.tabIndex === 0);
    expect(stops).toHaveLength(1);
    expect(document.activeElement).toBe(stops[0]);
    expect(stops[0].getAttribute('aria-selected')).toBe('true');
    expect(d.dialogEl.textContent).toContain('字元代碼：U+');
    // The code is read out once: with the cell's name, not again from a live region.
    expect(d.dialogEl.querySelector('[aria-live]')).toBeNull();
    expect(document.getElementById('dx-sy-recent-label')!.textContent).toBe('最近使用過的符號');
  });

  dialogTest('switching the subset changes the grid', async (ui, d) => {
    const first = d.cells()[0].textContent;
    pickSubset(d.subset(), '數學符號');
    await ui.flush();
    expect(d.cells()[0].textContent?.trim()).toBe('±');
    expect(d.cells()[0].textContent).not.toBe(first);
    expect(d.cells()).toHaveLength(SYMBOL_CATEGORIES[2].chars.length);
  });

  dialogTest('arrow keys, Home / End, PageUp / PageDown move; Enter inserts and the keyboard stays; Esc closes back to the text', async (ui, d) => {
    pickSubset(d.subset(), '帶圈及括號字元');
    await ui.flush();
    d.cells()[0].focus();
    await ui.flush();
    expect(document.activeElement?.textContent?.trim()).toBe('①');
    key(document.activeElement!, 'ArrowRight');
    await ui.flush();
    expect(document.activeElement?.textContent?.trim()).toBe('②');
    expect(d.dialogEl.textContent).toContain('字元代碼：U+2461');
    const cols = d.grid().querySelector('[role="row"]')!.querySelectorAll('[role="gridcell"]').length;
    key(document.activeElement!, 'ArrowDown');
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[1 + cols]);
    key(document.activeElement!, 'ArrowUp');
    key(document.activeElement!, 'End');
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[cols - 1]);
    key(document.activeElement!, 'Home');
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[0]);
    key(document.activeElement!, 'PageDown');
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[8 * cols]);
    key(document.activeElement!, 'PageUp');
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[0]);
    key(document.activeElement!, 'End', { ctrlKey: true });
    await ui.flush();
    expect(document.activeElement).toBe(d.cells()[d.cells().length - 1]);
    key(document.activeElement!, 'Home', { ctrlKey: true });
    key(document.activeElement!, 'ArrowRight');
    await ui.flush();
    const e = key(document.activeElement!, 'Enter');
    expect(e.defaultPrevented).toBe(true);
    await ui.flush();
    expect(text(ui)).toBe('公②文');
    expect(ui.dialog()).not.toBeNull(); // stays open, as in Word
    expect(document.activeElement).toBe(d.cells()[1]); // the keyboard stays in the grid
    key(document.activeElement!, 'Enter');
    await ui.flush();
    expect(text(ui)).toBe('公②②文');
    expect(d.recentCells()[0].textContent?.trim()).toBe('②');
    key(document.activeElement!, 'Escape');
    await ui.flush();
    expect(ui.dialog()).toBeNull();
    expect(ui.view.dom.contains(document.activeElement) || document.activeElement === ui.view.dom).toBe(true);
  });

  dialogTest('the recently used row from the keyboard: Enter inserts, the keyboard stays on the row, Esc still closes', async (ui, d) => {
    pickSubset(d.subset(), '帶圈及括號字元');
    await ui.flush();
    d.cells()[0].focus();
    key(document.activeElement!, 'Enter'); // ①
    await ui.flush();
    key(document.activeElement!, 'ArrowRight');
    key(document.activeElement!, 'Enter'); // ②
    await ui.flush();
    expect(d.recentCells().map((c) => c.textContent?.trim())).toEqual(['②', '①']);
    expect(d.recentCells().filter((c) => c.tabIndex === 0)).toEqual([d.recentCells()[0]]);
    d.recentCells()[0].focus();
    key(document.activeElement!, 'ArrowRight');
    await ui.flush();
    expect(document.activeElement).toBe(d.recentCells()[1]);
    expect(d.code().value).toBe('2460');
    key(document.activeElement!, 'Enter');
    await ui.flush();
    expect(text(ui)).toBe('公①②①文');
    expect(d.recentCells().map((c) => c.textContent?.trim())).toEqual(['①', '②']);
    // The row put ① first; the keyboard is on it, and it is the row's Tab stop.
    expect(document.activeElement).toBe(d.recentCells()[0]);
    expect(d.recentCells()[0].tabIndex).toBe(0);
    key(document.activeElement!, 'Escape');
    await ui.flush();
    expect(ui.dialog()).toBeNull();
  });

  dialogTest('double-click on a symbol inserts it; 插入 inserts the selected one', async (ui, d) => {
    const [a, b] = d.cells();
    a.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await ui.flush();
    expect(text(ui)).toBe(`公${a.textContent?.trim()}文`);
    b.click();
    await ui.flush();
    d.insertButton().focus();
    d.insertButton().click();
    await ui.flush();
    expect(text(ui)).toBe(`公${a.textContent?.trim()}${b.textContent?.trim()}文`);
    expect(ui.dialog()).not.toBeNull();
    expect(document.activeElement).toBe(d.insertButton()); // not the text: the dialog keeps the keyboard
  });

  dialogTest('the character code: 2605 inserts ★; D800 is refused with a Chinese message; a code selects its symbol', async (ui, d) => {
    const input = d.code();
    input.value = '2605';
    input.dispatchEvent(new Event('input'));
    await ui.flush();
    expect(d.dialogEl.textContent).toContain('字元代碼：U+2605');
    d.insertButton().click();
    await ui.flush();
    expect(text(ui)).toBe('公★文');
    d.code().value = 'D800';
    d.code().dispatchEvent(new Event('input'));
    await ui.flush();
    d.insertButton().click();
    await ui.flush();
    expect(d.dialogEl.querySelector('[role="alert"]')!.textContent).toMatch(/代理/);
    expect(text(ui)).toBe('公★文');
    // Typing the code of a symbol in a subset shows that subset with the symbol selected.
    d.code().value = '2460';
    d.code().dispatchEvent(new Event('input'));
    await ui.flush();
    expect(d.dialogEl.querySelector('[role="alert"]')).toBeNull();
    expect(d.cells().find((c) => c.getAttribute('aria-selected') === 'true')!.textContent?.trim()).toBe('①');
  });

  dialogTest('a typed code is kept when a cell gets the focus; Space selects the cell; Enter while composing does nothing', async (ui, d) => {
    d.code().value = '1F600';
    d.code().dispatchEvent(new Event('input'));
    await ui.flush();
    d.code().dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true, cancelable: true }));
    await ui.flush();
    expect(text(ui)).toBe('公文');
    d.cells()[3].focus();
    await ui.flush();
    expect(d.code().value).toBe('1F600');
    const e = key(d.cells()[3], ' ');
    expect(e.defaultPrevented).toBe(true);
    await ui.flush();
    expect(d.code().value).toBe(symbolCode(d.cells()[3].textContent!.trim()).slice(2));
    expect(text(ui)).toBe('公文'); // Space selects, it doesn't insert
    d.code().value = '1F600';
    d.code().dispatchEvent(new Event('input'));
    key(d.code(), 'Enter');
    await ui.flush();
    expect(text(ui)).toBe('公\u{1F600}文');
  });

  dialogTest('關閉 closes; Tab stays inside the dialog', async (ui, d) => {
    const buttons = Array.from(d.dialogEl.querySelectorAll<HTMLButtonElement>('button'));
    const last = buttons[buttons.length - 1];
    last.focus();
    const e = key(last, 'Tab');
    expect(e.defaultPrevented).toBe(true);
    expect(d.dialogEl.contains(document.activeElement)).toBe(true);
    buttons.find((b) => b.textContent?.trim() === '關閉')!.click();
    await ui.flush();
    expect(ui.dialog()).toBeNull();
  });

  it('in a header being edited: the dialog from the ribbon inserts there, and closing returns to the header', () =>
    withUi({}, async (ui) => {
      ui.editor.editHeaderFooter('header', 0);
      await ui.flush();
      const hv = ui.editor.activeView!;
      expect(hv).not.toBe(ui.view);
      expect(ui.editor.target).toBe('header');
      const d = await openDialog(ui);
      pickSubset(d.subset(), '帶圈及括號字元');
      await ui.flush();
      d.cells()[0].focus();
      key(document.activeElement!, 'Enter');
      await ui.flush();
      expect(hv.state.doc.textContent).toContain('①');
      expect(text(ui)).toBe('公文');
      key(document.activeElement!, 'Escape');
      await ui.flush();
      expect(ui.dialog()).toBeNull();
      expect(hv.dom.contains(document.activeElement) || document.activeElement === hv.dom).toBe(true);
    }));
});
