// @vitest-environment jsdom
// 顯示／隱藏編輯標記 (Word's Show/Hide ¶, Home › 段落, Ctrl+Shift+8 = Ctrl+*): shows the marks Word
// never prints — ¶ at every paragraph end (table cells and headers/footers too), · for a space,
// ° for a non-breaking space, □ for a full-width space, → for a tab and ↵ for a manual line
// break. It is a viewer preference, not part of the document: the marks are decorations only
// (no change to the document, no undo step, nothing in the saved file or the printout), and the
// choice is remembered per browser (localStorage 'papyrus.showMarks').
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { undoDepth } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';
import { insertPageBreak } from '../../src/papyrus/editor/commands';
import { lineBoxes, pagination, paginationKey } from '../../src/papyrus/editor/pagination';
import {
  SHOW_MARKS_STORAGE, formattingMarks, formattingMarksKey, showFormattingMarks, type MarkKind,
} from '../../src/papyrus/editor/formattingMarks';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
  (document as any).elementFromPoint ??= () => null;
});

const editors: DocxEditor[] = [];
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy();
  vi.restoreAllMocks();
  try {
    localStorage.removeItem(SHOW_MARKS_STORAGE);
  } catch {
    // storage stubbed away by a test
  }
});

/**
 * Body: 「甲 乙  丙」 (3 spaces), a tab, a manual line break, a non-breaking space and a full-width
 * space in one paragraph; an empty paragraph; a 1×2 table (two cell paragraphs); a paragraph
 * split by a page break (its first piece has no ¶ of its own: the w:p goes on after the break).
 */
const BODY =
  '<w:p><w:r><w:t xml:space="preserve">甲 乙  丙</w:t><w:tab/><w:t>丁</w:t><w:br/><w:t xml:space="preserve">戊\u00a0己\u3000庚</w:t></w:r></w:p>' +
  '<w:p/>' +
  '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
  '<w:tc><w:p><w:r><w:t xml:space="preserve">格 一</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>格二</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p><w:r><w:t>前</w:t><w:br w:type="page"/><w:t>後</w:t></w:r></w:p>';

async function zipOf(body: string, header?: string): Promise<Uint8Array> {
  const zip = blankPackage();
  const sect = header ? `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>${SECT}</w:sectPr>` : '<w:sectPr/>';
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`);
  if (header) {
    zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${header}</w:hdr>`);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
      '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(body = BODY, options: ConstructorParameters<typeof DocxEditor>[1] = {}, header?: string) {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, options);
  await editor.open(await zipOf(body, header));
  editors.push(editor);
  return editor;
}

async function documentXml(editor: DocxEditor): Promise<string> {
  const zip = await JSZip.loadAsync(await editor.save());
  return zip.file('word/document.xml')!.async('string');
}

/** How many marks of each kind the view's decorations hold. */
function counts(view: EditorView): Partial<Record<MarkKind, number>> {
  const out: Partial<Record<MarkKind, number>> = {};
  for (const d of formattingMarksKey.getState(view.state)!.set.find()) {
    const kind = (d.spec as { kind: MarkKind }).kind;
    out[kind] = (out[kind] ?? 0) + 1;
  }
  return out;
}

/** The marks drawn in a view's DOM. */
function drawn(root: ParentNode) {
  return {
    paragraph: root.querySelectorAll('.dx-mk-p').length,
    space: root.querySelectorAll('.dx-mk-sp').length,
    nbsp: root.querySelectorAll('.dx-mk-nbsp').length,
    ideographicSpace: root.querySelectorAll('.dx-mk-idsp').length,
    tab: root.querySelectorAll('.dx-mk-tab').length,
    lineBreak: root.querySelectorAll('.dx-mk-br').length,
  };
}

const NONE = { paragraph: 0, space: 0, nbsp: 0, ideographicSpace: 0, tab: 0, lineBreak: 0 };
// ¶: the first paragraph, the empty one, two cells and the last piece of 「前|後」 (not 「前」).
const ALL = { paragraph: 5, space: 4, nbsp: 1, ideographicSpace: 1, tab: 1, lineBreak: 1 };

describe('顯示／隱藏編輯標記: the marks', () => {
  it('toggling adds and removes ¶ · ° □ → ↵ (table cells included)', async () => {
    const editor = await open();
    const view = editor.view!;
    expect(editor.showMarks).toBe(false);
    expect(drawn(view.dom)).toEqual(NONE);
    expect(view.dom.classList.contains('dx-show-marks')).toBe(false);

    editor.setShowMarks(true);
    expect(editor.showMarks).toBe(true);
    expect(counts(view)).toEqual(ALL);
    expect(drawn(view.dom)).toEqual(ALL);
    expect(view.dom.classList.contains('dx-show-marks')).toBe(true);
    // Each ¶ ends its paragraph; the cells' paragraphs have theirs.
    const cells = Array.from(view.dom.querySelectorAll('td'));
    expect(cells.map((c) => c.querySelectorAll('.dx-mk-p').length)).toEqual([1, 1]);
    const before = view.dom.querySelector('.dx-mk-br')!.nextSibling as HTMLElement;
    expect(before.nodeName).toBe('BR'); // ↵ sits right before the line break, on its line

    editor.toggleShowMarks();
    expect(editor.showMarks).toBe(false);
    expect(drawn(view.dom)).toEqual(NONE);
    expect(view.dom.classList.contains('dx-show-marks')).toBe(false);
  });

  it('the marks are hidden from screen readers and take no text of their own', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    for (const w of Array.from(view.dom.querySelectorAll('.dx-mk'))) {
      expect(w.getAttribute('aria-hidden')).toBe('true');
      expect(w.textContent).toBe(''); // the glyph is CSS generated content
    }
    expect(view.dom.textContent).not.toMatch(/[¶·°□→↵]/);
  });

  it('does not change the document and adds no undo step', async () => {
    const editor = await open();
    const view = editor.view!;
    const doc = view.state.doc;
    editor.setShowMarks(true);
    editor.setShowMarks(false);
    editor.setShowMarks(true);
    expect(view.state.doc).toBe(doc);
    expect(undoDepth(view.state)).toBe(0);
    expect(editor.snapshot()!.canUndo).toBe(false);
    expect(editor.isModified()).toBe(false);
  });

  it('the saved file is the same with the marks shown or hidden', async () => {
    const editor = await open();
    const hidden = await documentXml(editor);
    editor.setShowMarks(true);
    expect(await documentXml(editor)).toBe(hidden);
    expect(hidden).not.toMatch(/[¶·°□→↵]/);
  });

  it('copying with the marks shown copies the text only', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    const first = view.state.doc.firstChild!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, first.nodeSize - 1)));
    // What the editor puts on the clipboard (Ctrl+C) is made from the document, not the screen.
    const { dom, text } = view.serializeForClipboard(view.state.selection.content());
    expect(text).toContain('甲 乙  丙');
    expect(text).not.toMatch(/[¶·°□→↵]/);
    expect(dom.querySelector('.dx-mk, .dx-mk-sp')).toBeNull();
  });

  it('keeps the marks up to date while typing, and splitting and joining paragraphs', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    // Type 「 X」 at the end of the first paragraph: one more space.
    const end = view.state.doc.firstChild!.nodeSize - 1;
    view.dispatch(view.state.tr.insertText(' X', end));
    expect(counts(view)).toEqual({ ...ALL, space: 5 });
    // A new paragraph after it (Enter): one more ¶.
    view.dispatch(view.state.tr.split(end + 2));
    expect(counts(view)).toEqual({ ...ALL, space: 5, paragraph: 6 });
    expect(drawn(view.dom)).toEqual({ ...ALL, space: 5, paragraph: 6 });
    // Joined again (Backspace at its start): back to one.
    view.dispatch(view.state.tr.join(end + 3));
    expect(counts(view)).toEqual({ ...ALL, space: 5 });
    // Undo takes back the typing, not the marks.
    editor.undo();
    editor.undo();
    editor.undo();
    expect(counts(view)).toEqual(ALL);
    expect(editor.showMarks).toBe(true);
  });

  /** The view's marks, and the marks made from scratch for its document: the same after any edit. */
  function marksNow(view: EditorView) {
    const list = (state: EditorState) =>
      formattingMarksKey.getState(state)!.set.find().map((d) => `${d.from}-${d.to}-${(d.spec as { kind: string }).kind}`).sort();
    const scratch = EditorState.create({ schema, doc: view.state.doc, plugins: [formattingMarks(() => true)] });
    return { kept: list(view.state), fresh: list(scratch) };
  }
  const paragraphMarks = (view: EditorView) => counts(view).paragraph;

  it('a paragraph given a section break (setNodeAttribute) loses its ¶, and gets it back', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    // The empty paragraph (the second block) ends a section now: an AttrStep, no position moves.
    const pos = view.state.doc.firstChild!.nodeSize;
    view.dispatch(view.state.tr.setNodeAttribute(pos, 'sectPr', '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'));
    expect(paragraphMarks(view)).toBe(ALL.paragraph - 1);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
    view.dispatch(view.state.tr.setNodeAttribute(pos, 'sectPr', null));
    expect(paragraphMarks(view)).toBe(ALL.paragraph);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
  });

  it('a page break inserted into or removed from a paragraph moves its ¶ (brGroup pieces)', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    // 「前」|page break|「後」: the page break removed, 「前」 is a paragraph of its own with its ¶.
    let breakPos = -1;
    view.state.doc.forEach((node, offset) => {
      if (node.type.name === 'page_break') breakPos = offset;
    });
    view.dispatch(view.state.tr.delete(breakPos, breakPos + 1));
    expect(paragraphMarks(view)).toBe(ALL.paragraph + 1);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
    editor.undo();
    expect(paragraphMarks(view)).toBe(ALL.paragraph);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
    // Ctrl+Enter in the first paragraph: its first piece has no ¶ (the w:p goes on after the break).
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
    expect(editor.run(insertPageBreak)).toBe(true);
    expect(paragraphMarks(view)).toBe(ALL.paragraph);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
    // The break inserted after a brGroup paragraph's last piece, in the same group: its ¶ goes.
    let last = -1;
    let group: string | null = null;
    view.state.doc.forEach((node, offset) => {
      if (node.attrs.brGroup && node.textContent === '後') [last, group] = [offset + node.nodeSize, node.attrs.brGroup];
    });
    view.dispatch(view.state.tr.insert(last, schema.nodes.page_break.create({ group })));
    expect(paragraphMarks(view)).toBe(ALL.paragraph - 1);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
  });

  it('a transaction with very many steps is marked again as a whole', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    const tr = view.state.tr;
    for (let i = 0; i < 300; i++) tr.insertText(' ', 2);
    view.dispatch(tr);
    // Every textblock: the five with ¶ and 「前」.
    expect(formattingMarksKey.getState(view.state)!.rebuilt).toBe(6);
    expect(counts(view).space).toBe(ALL.space + 300);
    expect(marksNow(view).kept).toEqual(marksNow(view).fresh);
  });

  it('the printout and the PDF (printHtml) show no marks', async () => {
    const editor = await open(BODY, {}, '<w:p><w:r><w:t xml:space="preserve">頁 首</w:t></w:r></w:p>');
    editor.setShowMarks(true);
    // jsdom lays nothing out: one page holding every block (all measured at 0).
    (editor as any).pages = [{ ...(editor as any).firstPage(), section: 0, inSection: 0, number: 1, top: -1 }];
    (editor as any).renderPages((editor as any).pages);
    // On screen the page's header copy shows its marks …
    expect(document.querySelector('.dx-page .dx-header .dx-mk-p')).not.toBeNull();
    // … and so do the header rows repeated where a table row breaks (a copy like renderCuts makes).
    const cut = document.createElement('div');
    cut.className = 'dx-cut-header dx-doc dx-show-marks';
    cut.style.top = '0px';
    cut.innerHTML = '<table><tbody><tr><td><p class="dx-p">標<span class="dx-mk-sp"> </span>題<span class="dx-mk dx-mk-p" aria-hidden="true"></span></p></td></tr></tbody></table>';
    document.querySelector('.dx-cuts')!.append(cut);
    const html = editor.printHtml('t')!;
    const body = html.slice(html.indexOf('<body'));
    // … the printout holds the text and the header, without any mark.
    expect(body).toContain('甲 乙  丙');
    expect(body).toContain('頁 首');
    expect(body).toContain('標 題');
    expect(body).not.toMatch(/dx-mk|dx-show-marks/);
    expect(body).not.toMatch(/[¶·°□→↵]/);
  });

  it('works in read-only mode', async () => {
    const editor = await open(BODY, { editable: false });
    editor.setShowMarks(true);
    expect(counts(editor.view!)).toEqual(ALL);
    expect(editor.snapshot()!.showMarks).toBe(true);
  });

  it('a column break keeps its own 分欄符號 label (no ↵); a section break shows its label instead of ¶', async () => {
    const editor = await open(
      '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>一</w:t><w:br w:type="column"/><w:t>二</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>三</w:t></w:r></w:p>',
    );
    editor.setShowMarks(true);
    expect(counts(editor.view!)).toEqual({ paragraph: 1 });
  });
});

describe('顯示／隱藏編輯標記: remembered per browser', () => {
  it('is kept in localStorage and read by the next editor', async () => {
    const first = await open();
    first.setShowMarks(true);
    expect(localStorage.getItem(SHOW_MARKS_STORAGE)).toBe('1');
    const second = await open();
    expect(second.showMarks).toBe(true);
    expect(second.snapshot()!.showMarks).toBe(true);
    expect(drawn(second.view!.dom)).toEqual(ALL);
    second.setShowMarks(false);
    expect(localStorage.getItem(SHOW_MARKS_STORAGE)).toBe('0');
    expect((await open()).showMarks).toBe(false);
  });

  it('works when the storage throws (private mode, blocked site data)', async () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new DOMException('denied', 'SecurityError');
    });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new DOMException('denied', 'QuotaExceededError');
    });
    const editor = await open();
    expect(editor.showMarks).toBe(false);
    expect(() => editor.setShowMarks(true)).not.toThrow();
    expect(counts(editor.view!)).toEqual(ALL);
  });

  it('works when the storage itself cannot be reached', async () => {
    const spy = vi.spyOn(window, 'localStorage', 'get').mockImplementation(() => {
      throw new DOMException('denied', 'SecurityError');
    });
    const editor = await open();
    expect(editor.showMarks).toBe(false);
    expect(() => editor.toggleShowMarks()).not.toThrow();
    expect(editor.showMarks).toBe(true);
    spy.mockRestore();
  });
});

describe('顯示／隱藏編輯標記: keyboard', () => {
  const press = (view: EditorView, init: KeyboardEventInit) => {
    const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: true, ...init });
    view.dom.dispatchEvent(e);
    return e.defaultPrevented;
  };

  it('Ctrl+Shift+8 (Ctrl+* on the keyboard) toggles the marks', async () => {
    const editor = await open();
    const view = editor.view!;
    expect(press(view, { key: '*', code: 'Digit8', keyCode: 56, shiftKey: true })).toBe(true);
    expect(editor.showMarks).toBe(true);
    expect(drawn(view.dom)).toEqual(ALL);
    // A keyboard layout whose Shift+8 is another character: found by the key's code.
    expect(press(view, { key: '(', code: 'Digit8', keyCode: 56, shiftKey: true })).toBe(true);
    expect(editor.showMarks).toBe(false);
    // The keypad's * with Ctrl.
    expect(press(view, { key: '*', code: 'NumpadMultiply', keyCode: 106 })).toBe(true);
    expect(editor.showMarks).toBe(true);
    expect(view.state.doc.textContent).not.toContain('*');
    // Not with Alt (AltGr on Windows is Ctrl+Alt), nor other digits.
    expect(press(view, { key: '*', code: 'Digit8', keyCode: 56, shiftKey: true, altKey: true })).toBe(false);
    expect(press(view, { key: '&', code: 'Digit7', keyCode: 55, shiftKey: true })).toBe(false);
    expect(editor.showMarks).toBe(true);
  });

  it('works in read-only mode too (ProseMirror handles no keys there), in the text or with nothing focused', async () => {
    const editor = await open(BODY, { editable: false });
    const view = editor.view!;
    expect(press(view, { key: '*', code: 'Digit8', keyCode: 56, shiftKey: true })).toBe(true);
    expect(editor.showMarks).toBe(true);
    expect(counts(view)).toEqual(ALL);
    const e = new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    document.body.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(editor.showMarks).toBe(false);
  });

  const bodyKey = (init: KeyboardEventInit = {}) => {
    const e = new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true, ...init });
    document.body.dispatchEvent(e);
    return e;
  };

  it('with a read-only preview over the document (版本預覽), the key reaches the preview too', async () => {
    // The app opens a read-only editor over the main one (VersionPreview) or two side by side
    // (VersionSideBySide). A read-only editor takes no focus: the key comes from the page body.
    const main = await open();
    const preview = await open(BODY, { editable: false });
    expect(bodyKey().defaultPrevented).toBe(true);
    expect(preview.showMarks).toBe(true);
    expect(counts(preview.view!)).toEqual(ALL);
    // One browser-wide choice: the editor underneath follows, and is not toggled twice.
    expect(main.showMarks).toBe(true);
    bodyKey();
    expect([main.showMarks, preview.showMarks]).toEqual([false, false]);
    // From inside one editor, likewise.
    main.view!.dom.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    expect([main.showMarks, preview.showMarks]).toEqual([true, true]);
    // And from the button (setShowMarks on one editor).
    preview.setShowMarks(false);
    expect([main.showMarks, preview.showMarks]).toEqual([false, false]);
    expect(counts(main.view!)).toEqual({});
  });

  it('the new value is the one the visible editor did not show (a hidden editor is not asked)', async () => {
    const shown = await open(BODY, { editable: false });
    const hidden = await open();
    // The later editor is hidden (checkVisibility false) and, somehow, out of step.
    (hidden as any).root.checkVisibility = () => false;
    (hidden as any).marksOn = true;
    bodyKey();
    expect(shown.showMarks).toBe(true);
    expect(hidden.showMarks).toBe(true);
  });

  it('a held key (auto-repeat) toggles once', async () => {
    const editor = await open();
    bodyKey();
    const repeat = bodyKey({ repeat: true });
    expect(repeat.defaultPrevented).toBe(false);
    expect(editor.showMarks).toBe(true);
  });

  it('on a Mac: ⌘8 (Word for Mac), ⌘⇧8 and ⌘*; not Ctrl', async () => {
    vi.spyOn(navigator, 'platform', 'get').mockReturnValue('MacIntel');
    const editor = await open();
    bodyKey({ ctrlKey: false, metaKey: true, shiftKey: false, key: '8' });
    expect(editor.showMarks).toBe(true);
    bodyKey({ ctrlKey: false, metaKey: true });
    expect(editor.showMarks).toBe(false);
    bodyKey({ ctrlKey: false, metaKey: true, shiftKey: false, key: '*', code: 'NumpadMultiply', keyCode: 106 });
    expect(editor.showMarks).toBe(true);
    bodyKey();
    expect(editor.showMarks).toBe(true);
  });

  it('elsewhere, Ctrl+8 without Shift is not the shortcut', async () => {
    const editor = await open();
    bodyKey({ shiftKey: false, key: '8' });
    expect(editor.showMarks).toBe(false);
  });

  it('keys in other parts of the page are left alone', async () => {
    const editor = await open();
    const input = document.createElement('input');
    document.body.append(input);
    const e = new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    input.dispatchEvent(e);
    input.remove();
    expect(e.defaultPrevented).toBe(false);
    expect(editor.showMarks).toBe(false);
  });

  it('a destroyed editor no longer listens (keys, scrolling)', async () => {
    const editor = await open();
    editor.setShowMarks(true);
    const view = editor.view!;
    const dispatch = vi.spyOn(view, 'dispatch');
    editors.splice(editors.indexOf(editor), 1);
    editor.destroy();
    localStorage.removeItem(SHOW_MARKS_STORAGE);
    vi.useFakeTimers();
    try {
      document.dispatchEvent(new Event('scroll'));
      window.dispatchEvent(new Event('resize'));
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
      vi.advanceTimersByTime(1000);
    } finally {
      vi.useRealTimers();
    }
    expect(dispatch).not.toHaveBeenCalled();
    expect(localStorage.getItem(SHOW_MARKS_STORAGE)).toBeNull();
    expect(editor.showMarks).toBe(true);
  });
});

describe('顯示／隱藏編輯標記: headers and footers', () => {
  const HEADER = '<w:p><w:r><w:t xml:space="preserve">機關 名稱</w:t><w:tab/><w:t>頁</w:t></w:r></w:p>';

  it('the header being edited shows its marks, and follows the toggle', async () => {
    const editor = await open('<w:p><w:r><w:t>內文</w:t></w:r></w:p>', {}, HEADER);
    editor.setShowMarks(true);
    editor.editHeaderFooter('header', 0);
    const hf = editor.activeView!;
    expect(hf).not.toBe(editor.view);
    expect(counts(hf)).toEqual({ paragraph: 1, space: 1, tab: 1 });
    expect(drawn(hf.dom)).toMatchObject({ paragraph: 1, space: 1, tab: 1 });
    // Ctrl+Shift+8 works there too, and hides the body's as well.
    hf.dom.dispatchEvent(new KeyboardEvent('keydown', { key: '*', code: 'Digit8', keyCode: 56, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }));
    expect(editor.showMarks).toBe(false);
    expect(drawn(hf.dom).paragraph).toBe(0);
    expect(drawn(editor.view!.dom).paragraph).toBe(0);
    // Its editing state is kept when it closes: reopened after the toggle, it follows.
    editor.closeHeaderFooter();
    editor.setShowMarks(true);
    editor.editHeaderFooter('header', 0);
    expect(drawn(editor.activeView!.dom).paragraph).toBe(1);
  });

  it('the header copies drawn on the pages show the marks too', async () => {
    const editor = await open('<w:p><w:r><w:t>內文</w:t></w:r></w:p>', {}, HEADER);
    const header = () => document.querySelector('.dx-page .dx-header') as HTMLElement;
    expect(drawn(header())).toEqual(NONE);
    editor.setShowMarks(true);
    expect(header().classList.contains('dx-show-marks')).toBe(true);
    expect(drawn(header())).toEqual({ ...NONE, paragraph: 1, space: 1, tab: 1 });
    expect(header().textContent).toBe('機關 名稱\t頁'); // the text itself is unchanged
    editor.setShowMarks(false);
    expect(drawn(header())).toEqual(NONE);
  });
});

describe('顯示／隱藏編輯標記: cost on a long document', () => {
  /** 2,000 paragraphs (about 200 pages) of mixed text with spaces and tabs. */
  function bigState(on: boolean) {
    const paras = [];
    for (let i = 0; i < 2000; i++) {
      paras.push(schema.nodes.paragraph.create(null, [
        schema.text(`第 ${i} 段 本案 依 規定 辦理 ，請 查照 。 Lorem ipsum dolor sit amet ${i}`),
        schema.nodes.tab.create(),
        schema.text('附件 一'),
      ]));
    }
    const doc = schema.nodes.doc.create(null, paras);
    return EditorState.create({ schema, doc, plugins: [formattingMarks(() => on)] });
  }

  it('decorates 2,000 paragraphs once, and typing redoes only the paragraph typed in', () => {
    let state = bigState(false);
    const t0 = performance.now();
    // No window: the whole document (as when the view is not laid out, e.g. in jsdom).
    state = state.apply(state.tr.setMeta(formattingMarksKey, { on: true }));
    const built = performance.now() - t0;
    const marks = formattingMarksKey.getState(state)!;
    expect(marks.on).toBe(true);
    expect(marks.rebuilt).toBe(2000);
    // Per paragraph: ¶, 16 spaces and a tab.
    expect(marks.set.find().length).toBe(2000 * (1 + 16 + 1));

    // Typing one character in the middle re-decorates that paragraph only.
    let pos = 0;
    state.doc.forEach((node, offset, index) => {
      if (index === 1000) pos = offset + node.nodeSize - 1;
    });
    const t1 = performance.now();
    for (let i = 0; i < 20; i++) state = state.apply(state.tr.insertText(' ', state.doc.resolve(pos).end()));
    const typed = (performance.now() - t1) / 20;
    expect(formattingMarksKey.getState(state)!.rebuilt).toBe(1);
    expect(formattingMarksKey.getState(state)!.set.find().length).toBe(2000 * 18 + 20);
    // Timings for the record only (a loaded machine must not fail the test).
    // eslint-disable-next-line no-console
    console.info(`formatting marks: build ${built.toFixed(1)} ms for 2,000 paragraphs, ${typed.toFixed(2)} ms per keystroke`);
  });

  /** The document range from the start of paragraph `a` to the end of paragraph `b`. */
  function range(state: EditorState, a: number, b: number) {
    let from = 0;
    let to = 0;
    state.doc.forEach((node, offset, index) => {
      if (index === a) from = offset;
      if (index === b) to = offset + node.nodeSize;
    });
    return { from, to };
  }
  const spaces = (state: EditorState) =>
    formattingMarksKey.getState(state)!.set.find().filter((d) => (d.spec as { kind: MarkKind }).kind === 'space').length;

  it('marks spaces only in the window around the screen; ¶ and → everywhere', () => {
    let state = bigState(false);
    state = state.apply(state.tr.setMeta(formattingMarksKey, { on: true, window: range(state, 100, 110) }));
    expect(spaces(state)).toBe(11 * 16);
    expect(formattingMarksKey.getState(state)!.set.find().length).toBe(2000 * 2 + 11 * 16);
    // Scrolled on: the paragraphs leaving the window lose their spaces, those entering get them.
    state = state.apply(state.tr.setMeta(formattingMarksKey, { window: range(state, 500, 505) }));
    expect(spaces(state)).toBe(6 * 16);
    expect(formattingMarksKey.getState(state)!.rebuilt).toBe(11 + 6);
    // Typing in the window marks the new space; typing far from it does not.
    const inside = range(state, 502, 502).to - 1;
    state = state.apply(state.tr.insertText(' ', inside));
    expect(spaces(state)).toBe(6 * 16 + 1);
    state = state.apply(state.tr.insertText(' ', 5));
    expect(spaces(state)).toBe(6 * 16 + 1);
    expect(formattingMarksKey.getState(state)!.set.find().length).toBe(2000 * 2 + 6 * 16 + 1);
  });

  it('shown when a long document opens (the remembered choice): spaces at its start first', () => {
    const state = bigState(true);
    const marks = formattingMarksKey.getState(state)!;
    expect(marks.window).toEqual({ from: 0, to: 30000 });
    expect(spaces(state)).toBeGreaterThan(0);
    expect(spaces(state)).toBeLessThan(2000 * 16 / 4);
    expect(marks.set.find().filter((d) => (d.spec as { kind: MarkKind }).kind === 'paragraph').length).toBe(2000);
  });

  it('the window follows the blocks on screen as the page scrolls', async () => {
    const body = Array.from({ length: 60 }, (_, i) => `<w:p><w:r><w:t xml:space="preserve">段 ${i}</w:t></w:r></w:p>`).join('');
    const editor = await open(body);
    const view = editor.view!;
    // A laid-out page: each paragraph 100 px tall, `scroll` px scrolled; the screen 768 px (jsdom).
    // (Pagination, seeing the page fill up, adds its spacers between them: not counted.)
    let scroll = 0;
    const paragraphs = () => Array.from(view.dom.querySelectorAll(':scope > p'));
    vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
      const i = this.parentElement === view.dom ? paragraphs().indexOf(this) : -1;
      const [top, height] = this === view.dom ? [-scroll, 6000] : i >= 0 ? [i * 100 - scroll, 100] : [0, 0];
      return { top, bottom: top + height, height, left: 0, right: 500, width: 500, x: 0, y: top, toJSON() {} } as DOMRect;
    });
    const shown = () => paragraphs().flatMap((p, i) => (p.querySelector('.dx-mk-sp') ? [i] : []));
    const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
    expect(window.innerHeight).toBe(768);
    vi.useFakeTimers();
    try {
      editor.setShowMarks(true);
      // Paragraphs 0–7 are on screen; two screens below reach paragraph 23.
      expect(shown()).toEqual(range(0, 23));
      expect(view.dom.querySelectorAll('.dx-mk-p').length).toBe(60);
      // Scrolled to paragraph 40: the window moves shortly after, not at once.
      scroll = 4000;
      document.dispatchEvent(new Event('scroll'));
      expect(shown()).toEqual(range(0, 23));
      vi.advanceTimersByTime(150);
      expect(shown()).toEqual(range(24, 59));
      // A small scroll within the window changes nothing.
      const set = formattingMarksKey.getState(view.state)!.set;
      scroll = 3900;
      document.dispatchEvent(new Event('scroll'));
      vi.advanceTimersByTime(150);
      expect(formattingMarksKey.getState(view.state)!.set).toBe(set);
      // Zooming moves the blocks without a scroll event: the window is checked again.
      scroll = 0;
      editor.setZoom(0.5);
      vi.advanceTimersByTime(150);
      expect(shown()).toEqual(range(0, 23));
      // While an input method is composing, the window waits (re-marking would disturb it).
      const composing = vi.spyOn(view, 'composing', 'get').mockReturnValue(true);
      scroll = 4000;
      document.dispatchEvent(new Event('scroll'));
      vi.advanceTimersByTime(500);
      expect(shown()).toEqual(range(0, 23));
      composing.mockReturnValue(false);
      vi.advanceTimersByTime(150);
      expect(shown()).toEqual(range(24, 59));
    } finally {
      vi.useRealTimers();
    }
    expect(undoDepth(view.state)).toBe(0);
  });

  it('with the marks hidden, nothing is decorated or kept up to date', () => {
    let state = bigState(false);
    state = state.apply(state.tr.insertText('字', 2));
    expect(formattingMarksKey.getState(state)!.set.find().length).toBe(0);
    expect(formattingMarksKey.getState(state)!.rebuilt).toBe(0);
  });

  it('showFormattingMarks does nothing when the view already shows what is asked', async () => {
    const editor = await open();
    const view = editor.view!;
    const spy = vi.spyOn(view, 'dispatch');
    showFormattingMarks(view, false);
    expect(spy).not.toHaveBeenCalled();
    showFormattingMarks(view, true);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});

describe('顯示／隱藏編輯標記: pagination sees no difference', () => {
  /** A rect at `top`, `height` tall (`width` wide). */
  const rect = (top: number, height: number, width = 10) =>
    ({ top, bottom: top + height, height, left: 0, right: width, width, x: 0, y: top, toJSON() {} }) as DOMRect;

  it('a paragraph ending with a line break counts the same lines with the ¶ shown (ProseMirror separator image)', () => {
    // With a widget last in a paragraph ProseMirror adds an empty separator <img> (and a <br>);
    // it sits on the line after the break and must not count as a line of text.
    const hidden = document.createElement('p');
    hidden.innerHTML = 'text<br><br class="ProseMirror-trailingBreak">';
    const shown = document.createElement('p');
    shown.innerHTML =
      'text<br><span class="dx-mk dx-mk-p" contenteditable="false" aria-hidden="true"></span>' +
      '<img class="ProseMirror-separator" alt=""><br class="ProseMirror-trailingBreak">';
    document.body.append(hidden, shown);
    vi.spyOn(Range.prototype, 'getClientRects').mockImplementation(() => [rect(0, 20, 40)] as unknown as DOMRectList);
    vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
      return this.tagName === 'IMG' ? rect(20, 0, 0) : rect(0, 40);
    });
    try {
      expect(lineBoxes(shown)).toHaveLength(lineBoxes(hidden).length);
      expect(lineBoxes(shown)).toHaveLength(1);
      // A real picture still counts.
      const picture = document.createElement('p');
      picture.innerHTML = 'text<br><img class="dx-img" alt="">';
      vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
        return this.tagName === 'IMG' ? rect(20, 50, 50) : rect(0, 70);
      });
      document.body.append(picture);
      expect(lineBoxes(picture)).toHaveLength(2);
      picture.remove();
    } finally {
      hidden.remove();
      shown.remove();
    }
  });

  it('repeated table header rows are drawn again when the marks are toggled', () => {
    const cell = (text: string) => schema.nodes.table_cell.create(null, [schema.nodes.paragraph.create(null, [schema.text(text)])]);
    const table = schema.nodes.table.create(null, [
      schema.nodes.table_row.create({ header: true }, [cell('標 題')]),
      schema.nodes.table_row.create(null, [cell('內 容')]),
    ]);
    const doc = schema.nodes.doc.create(null, [table]);
    let state = EditorState.create({ schema, doc, plugins: [formattingMarks(() => false), pagination(() => ({}) as never, () => {})] });
    const header = 1; // the first row, inside the table
    const second = header + table.child(0).nodeSize;
    state = state.apply(state.tr.setMeta(paginationKey, { breaks: [{ pos: second, height: 10, row: { cols: 1, header: [header] } }], pages: [] }));
    const key = () =>
      paginationKey.getState(state)!.decorations.find().map((d) => (d.spec as { key?: string }).key).find((k) => k?.startsWith('rh'));
    const hidden = key();
    expect(hidden).toBeTruthy();
    state = state.apply(state.tr.setMeta(formattingMarksKey, { on: true }));
    expect(key()).not.toBe(hidden);
    state = state.apply(state.tr.setMeta(formattingMarksKey, { on: false }));
    expect(key()).toBe(hidden);
  });
});
