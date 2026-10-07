// @vitest-environment jsdom
// Word's 複製格式 (Format Painter), Home › Clipboard. Word copies the formatting of the selection
// (or of the text at the cursor) and pastes it onto the next text the user selects with the
// mouse; a single click on a word paints that word. The character formatting replaces the target's
// own; the paragraph formatting comes along when the cursor had no selection or the selection held
// whole paragraphs (the paragraph mark). One click on the button paints once; a double-click
// keeps painting until Esc or the button again. Ctrl+Shift+C / Ctrl+Shift+V copy and paste the
// formatting from the keyboard. Numbering, section and page-break settings are not formatting
// the painter copies.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick } from 'vue';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { insertImage, setAlign } from '../../src/papyrus/editor/commands';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
  // jsdom has no layout: a click maps to no document position (the test sets the selection).
  (document as any).elementFromPoint ??= () => null;
});

const editors: DocxEditor[] = [];
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy();
});

async function open(body: string, options: ConstructorParameters<typeof DocxEditor>[1] = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, options);
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  editors.push(editor);
  return editor;
}
async function part(editor: DocxEditor, name = 'word/document.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(await editor.save());
  return zip.file(name)!.async('string');
}
const paragraphs = (xml: string) => xml.match(/<w:p\b[\s\S]*?<\/w:p>/g)!;

/** Document position of the `n`th character of the text `text` (the first occurrence). */
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
const selectText = (view: EditorView, text: string) => select(view, posOf(view, text), posOf(view, text) + text.length);
/** Ctrl+Shift+<letter> in the text; whether the editor used the key. */
function ctrlShift(view: EditorView, letter: 'C' | 'V'): boolean {
  const e = new KeyboardEvent('keydown', { key: letter, code: `Key${letter}`, ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
  view.dom.dispatchEvent(e);
  return e.defaultPrevented;
}
/** A mouse selection in the text: press, (the selection the test made), release. */
async function mouse(view: EditorView) {
  view.dom.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
  document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, button: 0 }));
  await new Promise((r) => setTimeout(r, 0));
}

// Source: bold red 16 pt, in a centred paragraph with space before and a list number.
const SOURCE =
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="240"/><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="32"/></w:rPr><w:t>來源文字</w:t></w:r></w:p>';
// Target: italic, with a language and a character spacing (w:spacing is kept, not modelled).
const TARGET =
  '<w:p><w:r><w:rPr><w:i/><w:spacing w:val="20"/><w:lang w:val="zh-TW"/></w:rPr><w:t>目標文字</w:t></w:r></w:p>';

describe('複製格式: character formatting', () => {
  it('pastes bold, colour and size, replaces the target\'s own direct formatting and keeps its other run properties', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    expect(ctrlShift(view, 'C')).toBe(true);
    selectText(view, '目標文字');
    expect(ctrlShift(view, 'V')).toBe(true);
    const [, p2] = paragraphs(await part(editor));
    expect(p2).toContain('<w:b/>');
    expect(p2).toContain('<w:color w:val="FF0000"/>');
    expect(p2).toContain('<w:sz w:val="32"/>');
    expect(p2).not.toContain('<w:i/>'); // the target's own direct formatting is replaced
    expect(p2).toContain('<w:spacing w:val="20"/>'); // not formatting the editor models: kept
    expect(p2).toContain('<w:lang w:val="zh-TW"/>');
    // A partial selection copies no paragraph formatting.
    expect(p2).not.toContain('<w:jc');
    expect(p2).not.toContain('<w:numPr');
  });

  it('is one undo step', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    const before = await part(editor);
    select(view, posOf(view, '來源', 1));
    editor.copyFormat();
    selectText(view, '目標文字');
    expect(editor.pasteFormat()).toBe(true);
    expect(await part(editor)).not.toBe(before);
    undo(view.state, view.dispatch);
    expect(await part(editor)).toBe(before);
  });

  it('right after typing, the paste is still its own undo step', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    select(view, posOf(view, '來源', 1)); // the cursor: paragraph formatting too
    editor.copyFormat();
    select(view, posOf(view, '目標文字', 4));
    view.dispatch(view.state.tr.insertText('新字'));
    const typed = await part(editor);
    // The cursor is after the new text: the paste goes to that word and its paragraph.
    expect(editor.pasteFormat()).toBe(true);
    expect(paragraphs(await part(editor))[1]).toContain('<w:jc w:val="center"/>');
    undo(view.state, view.dispatch);
    expect(await part(editor)).toBe(typed);
  });

  it('pasting formatting the target already has makes no edit', async () => {
    const editor = await open(SOURCE + SOURCE.replace('來源文字', '第二來源'));
    const view = editor.view!;
    selectText(view, '來源');
    editor.copyFormat();
    selectText(view, '第二');
    const doc = view.state.doc;
    editor.pasteFormat();
    expect(view.state.doc).toBe(doc);
  });

  it('outside a word (an empty paragraph), the next text typed gets the formatting', async () => {
    const editor = await open(SOURCE + '<w:p/>');
    const view = editor.view!;
    selectText(view, '來源');
    editor.copyFormat();
    select(view, view.state.doc.content.size - 1);
    expect(editor.pasteFormat()).toBe(true);
    view.dispatch(view.state.tr.insertText('新'));
    const [, p2] = paragraphs(await part(editor));
    expect(p2).toContain('<w:b/>');
    expect(p2).toContain('<w:t>新</w:t>');
  });

  it('Ctrl+Shift+V with nothing copied is left to the browser (paste as plain text)', async () => {
    const editor = await open(TARGET);
    const view = editor.view!;
    selectText(view, '目標');
    expect(ctrlShift(view, 'V')).toBe(false);
  });
});

describe('複製格式: paragraph formatting', () => {
  it('a cursor with no selection copies the paragraph formatting too (not the numbering)', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    select(view, posOf(view, '來源', 1));
    editor.copyFormat();
    selectText(view, '目標');
    editor.pasteFormat();
    const [, p2] = paragraphs(await part(editor));
    expect(p2).toContain('<w:jc w:val="center"/>');
    expect(p2).toMatch(/<w:spacing w:before="240"\/>/);
    expect(p2).toContain('<w:b/>');
    expect(p2).not.toContain('<w:numPr');
  });

  it('a selection of the whole paragraph copies the paragraph formatting; a part of it does not', async () => {
    const editor = await open(SOURCE + TARGET + TARGET.replace('目標文字', '第三段落'));
    const view = editor.view!;
    selectText(view, '來源文字');
    editor.copyFormat();
    selectText(view, '目標');
    editor.pasteFormat();
    selectText(view, '來源');
    editor.copyFormat();
    selectText(view, '第三');
    editor.pasteFormat();
    const [, p2, p3] = paragraphs(await part(editor));
    expect(p2).toContain('<w:jc w:val="center"/>');
    expect(p3).not.toContain('<w:jc');
    expect(p3).toContain('<w:b/>');
  });

  it('paragraph formatting goes to every paragraph the target touches', async () => {
    const editor = await open(SOURCE + TARGET + TARGET.replace('目標文字', '第三段落'));
    const view = editor.view!;
    select(view, posOf(view, '來源'));
    editor.copyFormat();
    select(view, posOf(view, '標文'), posOf(view, '第三') + 1);
    editor.pasteFormat();
    const [, p2, p3] = paragraphs(await part(editor));
    expect(p2).toContain('<w:jc w:val="center"/>');
    expect(p3).toContain('<w:jc w:val="center"/>');
  });
});

describe('複製格式: the painter (mouse)', () => {
  it('one click on the button paints once: the next mouse selection, then it turns off', async () => {
    const editor = await open(SOURCE + TARGET + TARGET.replace('目標文字', '第三段落'));
    const view = editor.view!;
    const root = view.dom.closest('.dx-root')!;
    selectText(view, '來源');
    expect(editor.startFormatPainter(false)).toBe(true);
    expect(editor.snapshot()!.formatPainter).toBe('once');
    expect(root.classList.contains('dx-format-painting')).toBe(true);
    selectText(view, '目標文字');
    await mouse(view);
    expect(editor.snapshot()!.formatPainter).toBeNull();
    expect(root.classList.contains('dx-format-painting')).toBe(false);
    // Off: the next selection is not painted.
    selectText(view, '第三段落');
    await mouse(view);
    const [, p2, p3] = paragraphs(await part(editor));
    expect(p2).toContain('<w:b/>');
    expect(p3).not.toContain('<w:b/>');
  });

  it('a click (no selection) paints the word under it', async () => {
    const editor = await open(SOURCE + '<w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>');
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(false);
    select(view, posOf(view, 'beta', 2));
    await mouse(view);
    const [, p2] = paragraphs(await part(editor));
    expect(p2).toMatch(/<w:t xml:space="preserve">alpha <\/w:t><\/w:r><w:r><w:rPr><w:b\/>[\s\S]*<w:t>beta<\/w:t><\/w:r><w:r><w:t xml:space="preserve"> gamma<\/w:t>/);
  });

  it('a tab ends the word a click paints', async () => {
    const editor = await open(SOURCE + '<w:p><w:r><w:t>甲乙</w:t><w:tab/><w:t>alpha beta</w:t></w:r></w:p>');
    const view = editor.view!;
    selectText(view, '來源');
    editor.copyFormat();
    select(view, posOf(view, 'alpha', 2));
    editor.pasteFormat();
    const painted: string[] = [];
    view.state.doc.child(1).forEach((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'bold')) painted.push(n.text!);
    });
    expect(painted).toEqual(['alpha']);
  });

  it('a double-click keeps painting until Esc; the button turns it off too', async () => {
    const editor = await open(SOURCE + TARGET + TARGET.replace('目標文字', '第三段落'));
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(true);
    expect(editor.snapshot()!.formatPainter).toBe('sticky');
    selectText(view, '目標文字');
    await mouse(view);
    selectText(view, '第三段落');
    await mouse(view);
    expect(editor.snapshot()!.formatPainter).toBe('sticky');
    const esc = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
    view.dom.dispatchEvent(esc);
    expect(editor.snapshot()!.formatPainter).toBeNull();
    const [, p2, p3] = paragraphs(await part(editor));
    expect(p2).toContain('<w:b/>');
    expect(p3).toContain('<w:b/>');
    // Each application is its own undo step.
    undo(view.state, view.dispatch);
    const [, u2, u3] = paragraphs(await part(editor));
    expect(u2).toContain('<w:b/>');
    expect(u3).not.toContain('<w:b/>');
    // The button again (toggle) ends it as well.
    editor.toggleFormatPainter(true);
    expect(editor.snapshot()!.formatPainter).toBe('sticky');
    editor.toggleFormatPainter(false);
    expect(editor.snapshot()!.formatPainter).toBeNull();
  });

  it('another command ends the painter', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(true);
    editor.run(setAlign('right', editor.model.styles));
    expect(editor.snapshot()!.formatPainter).toBeNull();
  });

  it('Ctrl+Shift+C copies without arming the painter', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    ctrlShift(view, 'C');
    expect(editor.snapshot()!.formatPainter).toBeNull();
    selectText(view, '目標');
    await mouse(view);
    const [, p2] = paragraphs(await part(editor));
    expect(p2).not.toContain('<w:b/>');
  });
});

describe('複製格式: 追蹤修訂 and header/footer', () => {
  it('with 追蹤修訂 on, the formatting is recorded (w:rPrChange, w:pPrChange)', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    editor.setTrackChanges(true);
    select(view, posOf(view, '來源', 1));
    editor.copyFormat();
    selectText(view, '目標文字');
    expect(editor.pasteFormat()).toBe(true);
    const [, p2] = paragraphs(await part(editor));
    expect(p2).toContain('<w:rPrChange');
    expect(p2).toMatch(/<w:rPrChange[^>]*>\s*<w:rPr>[\s\S]*<w:i\/>/); // what it was before
    expect(p2).toContain('<w:pPrChange');
    expect(p2).toContain('<w:jc w:val="center"/>');
  });

  it('works in a header being edited, with formatting copied from the body', async () => {
    const editor = await open(SOURCE);
    const view = editor.view!;
    selectText(view, '來源');
    editor.copyFormat();
    editor.editHeaderFooter('header', 0);
    const hv = editor.activeView!;
    expect(hv).not.toBe(view);
    hv.dispatch(hv.state.tr.insertText('頁首文字'));
    selectText(hv, '頁首文字');
    expect(ctrlShift(hv, 'V')).toBe(true);
    editor.closeHeaderFooter();
    const zip = await JSZip.loadAsync(await editor.save());
    const header = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
    const xml = await zip.file(header)!.async('string');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:color w:val="FF0000"/>');
  });
});

describe('複製格式: toolbar button', () => {
  async function mount() {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${SOURCE}${TARGET}<w:sectPr/></w:body></w:document>`);
    const src = await zip.generateAsync({ type: 'uint8array' });
    const host = document.createElement('div');
    document.body.append(host);
    let editor: DocxEditor | null = null;
    const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditor) => (editor = ed) }) });
    app.mount(host);
    const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
    for (let i = 0; i < 80 && !editor; i++) await flush();
    await flush();
    if (!editor) throw new Error('DocxEditorVue did not emit ready');
    const button = [...host.querySelectorAll<HTMLButtonElement>('[role="toolbar"] button')].find((b) => b.textContent?.trim() === '複製格式')!;
    return { host, button, editor: editor as DocxEditor, flush, done: () => { app.unmount(); host.remove(); } };
  }

  it('mounts, and aria-pressed follows the painter (click: once; double-click: sticky)', async () => {
    const ui = await mount();
    expect(ui.button).not.toBeNull();
    expect(ui.button.type).toBe('button');
    expect(ui.button.title).toContain('複製格式');
    expect(ui.button.title).toContain('Ctrl+Shift+C');
    expect(ui.button.getAttribute('aria-pressed')).toBe('false');
    const view = ui.editor.view!;
    selectText(view, '來源');
    ui.button.click();
    await ui.flush();
    expect(ui.button.getAttribute('aria-pressed')).toBe('true');
    expect(ui.editor.snapshot()!.formatPainter).toBe('once');
    ui.button.click();
    await ui.flush();
    expect(ui.button.getAttribute('aria-pressed')).toBe('false');
    // A double-click: two clicks, then dblclick.
    ui.button.click();
    ui.button.click();
    ui.button.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await ui.flush();
    expect(ui.editor.snapshot()!.formatPainter).toBe('sticky');
    expect(ui.button.getAttribute('aria-pressed')).toBe('true');
    view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    await ui.flush();
    expect(ui.button.getAttribute('aria-pressed')).toBe('false');
    ui.done();
  });

  it('a dblclick event on its own starts sticky painting', async () => {
    const ui = await mount();
    selectText(ui.editor.view!, '來源');
    ui.button.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await ui.flush();
    expect(ui.editor.snapshot()!.formatPainter).toBe('sticky');
    expect(ui.button.getAttribute('aria-pressed')).toBe('true');
    ui.done();
  });
});

// Found in review of the first version.
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('甲一')}${cell('乙一')}</w:tr><w:tr>${cell('甲二')}${cell('乙二')}</w:tr></w:tbl><w:p/>`;
/** The position before the table cell holding `text`. */
function cellPos(view: EditorView, text: string): number {
  const $pos = view.state.doc.resolve(posOf(view, text));
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.spec.tableRole === 'cell') return $pos.before(d);
  throw new Error(`no cell for ${text}`);
}
/** The texts shown bold, in document order. */
const boldTexts = (view: EditorView) => {
  const out: string[] = [];
  view.state.doc.descendants((n) => {
    if (n.isText && n.marks.some((m) => m.type.name === 'bold')) out.push(n.text!);
  });
  return out;
};
/** Mouse down in the text and up on the document, without waiting for the painter. */
function press(view: EditorView, target: Element = view.dom) {
  target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
  document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, button: 0 }));
}
const tick = () => new Promise((r) => setTimeout(r, 0));
const PIXEL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

describe('複製格式: table cells and pictures', () => {
  it('a column of selected cells is painted in every cell, not just the last one', async () => {
    const editor = await open(SOURCE + TABLE);
    const view = editor.view!;
    selectText(view, '來源');
    editor.copyFormat();
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cellPos(view, '甲一'), cellPos(view, '甲二'))));
    expect(editor.pasteFormat()).toBe(true);
    expect(boldTexts(view)).toEqual(['來源文字', '甲一', '甲二']);
  });

  it('a mouse selection of cells stays a cell selection and is painted', async () => {
    const editor = await open(SOURCE + TABLE);
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(false);
    selectText(view, '甲一'); // where the browser's selection is when the mouse goes up
    press(view);
    // prosemirror-tables makes the drag a cell selection after the painter's mouseup listener.
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cellPos(view, '甲一'), cellPos(view, '乙二'))));
    await tick();
    expect(view.state.selection).toBeInstanceOf(CellSelection);
    expect(boldTexts(view)).toEqual(['來源文字', '甲一', '乙一', '甲二', '乙二']);
  });

  it('a click on a picture keeps the picture selected (the text selected before is not painted)', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    select(view, posOf(view, '目標文字', 4));
    editor.run(insertImage(PIXEL, 10, 10, 500));
    let pic = -1;
    view.state.doc.descendants((n, pos) => {
      if (n.type.name === 'image') pic = pos;
    });
    expect(pic).toBeGreaterThan(0);
    selectText(view, '來源');
    editor.startFormatPainter(false);
    selectText(view, '目標');
    press(view);
    // ProseMirror selects the clicked picture on mouseup, after the painter's listener.
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pic)));
    await tick();
    expect(view.state.selection).toBeInstanceOf(NodeSelection);
    expect(boldTexts(view)).toEqual(['來源文字']);
    expect(editor.snapshot()!.formatPainter).toBeNull();
  });

  it('a mousedown on a table column-resize handle does not use up a one-shot painter', async () => {
    const editor = await open(SOURCE + TABLE);
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(false);
    const handle = document.createElement('div');
    handle.className = 'column-resize-handle';
    view.dom.querySelector('td, th')!.append(handle);
    press(view, handle);
    await tick();
    expect(editor.snapshot()!.formatPainter).toBe('once');
  });
});

describe('複製格式: Ctrl+Shift+V pastes only what Ctrl+Shift+C copied', () => {
  it('the button does not arm Ctrl+Shift+V', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    editor.startFormatPainter(false);
    editor.stopFormatPainter();
    selectText(view, '目標');
    expect(ctrlShift(view, 'V')).toBe(false);
    expect(boldTexts(view)).toEqual(['來源文字']);
  });

  it('Ctrl+C after Ctrl+Shift+C gives Ctrl+Shift+V back to the browser (paste as plain text)', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    expect(ctrlShift(view, 'C')).toBe(true);
    view.dom.dispatchEvent(new Event('copy', { bubbles: true, cancelable: true }));
    selectText(view, '目標');
    expect(ctrlShift(view, 'V')).toBe(false);
    expect(boldTexts(view)).toEqual(['來源文字']);
  });

  it('Ctrl+X does the same', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    ctrlShift(view, 'C');
    document.dispatchEvent(new Event('cut', { bubbles: true, cancelable: true }));
    selectText(view, '目標');
    expect(ctrlShift(view, 'V')).toBe(false);
  });

  it('Ctrl+Shift+C then Ctrl+Shift+V pastes the formatting', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    selectText(view, '來源');
    ctrlShift(view, 'C');
    selectText(view, '目標');
    expect(ctrlShift(view, 'V')).toBe(true);
    expect(boldTexts(view)).toEqual(['來源文字', '目標']);
  });
});

describe('複製格式: lifecycle', () => {
  it('opening another file stops the painter and forgets the copied formatting', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    const root = view.dom.closest('.dx-root')!;
    selectText(view, '來源');
    ctrlShift(view, 'C');
    editor.startFormatPainter(true);
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${TARGET}<w:sectPr/></w:body></w:document>`);
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    expect(editor.snapshot()!.formatPainter).toBeNull();
    expect(root.classList.contains('dx-format-painting')).toBe(false);
    selectText(editor.view!, '目標');
    expect(editor.pasteFormat()).toBe(false);
  });

  it('destroy() while painting leaves no paint class and throws nothing', async () => {
    const editor = await open(SOURCE + TARGET);
    const view = editor.view!;
    const root = view.dom.closest('.dx-root')!;
    selectText(view, '來源');
    editor.startFormatPainter(true);
    view.dom.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
    expect(() => editor.destroy()).not.toThrow();
    expect(root.classList.contains('dx-format-painting')).toBe(false);
    // A mouseup after the editor is gone does nothing.
    document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, button: 0 }));
    await tick();
  });

  it('is refused in read-only mode', async () => {
    const editor = await open(SOURCE + TARGET, { editable: false });
    const view = editor.view!;
    selectText(view, '來源');
    expect(editor.startFormatPainter(false)).toBe(false);
    expect(editor.snapshot()!.formatPainter).toBeNull();
  });
});
