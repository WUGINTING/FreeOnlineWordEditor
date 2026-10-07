// Editing a text box's text in the editor (phase 8A): double-click / Enter opens an editor over
// the text box, like a header's; the text is saved into both copies of the text box; undo, find
// and replace, word count and spell check include text boxes.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { NodeSelection, type Command, type EditorState } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { findMatches, replaceMatchesTr } from '../../src/papyrus/editor/search';
import { findShape } from '../../src/papyrus/editor/shapeEdit';
import { textFrames, type ShapeModel } from '../../src/papyrus/docx/shapes';
import { schema } from '../../src/papyrus/editor/schema';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
  (document as any).elementFromPoint ??= () => null;
});

/** Typing, as its own undo step. */
const type = (text: string): Command => (state, dispatch) => {
  dispatch?.(closeHistory(state.tr.insertText(text)));
  return true;
};

let editor: DocxEditor | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  editor?.destroy();
  host?.remove();
  editor = null;
});

async function open(name: string, options = {}): Promise<DocxEditor> {
  host = document.createElement('div');
  document.body.append(host);
  editor = new DocxEditor(host, options);
  await editor.open(readFileSync(`tests/fixtures/shapes/${name}.docx`));
  return editor;
}

/** The shapes of the body with their text. */
function bodyShapes(ed: DocxEditor): { key: string; shape: ShapeModel; texts: string[] }[] {
  const out: { key: string; shape: ShapeModel; texts: string[] }[] = [];
  ed.view!.state.doc.descendants((n) => {
    const shape = n.attrs?.shape as ShapeModel | undefined;
    if (shape) out.push({ key: shape.key, shape, texts: textFrames(shape).map((f) => schema.nodeFromJSON(f.doc).textContent) });
    return true;
  });
  return out;
}

async function savedXml(ed: DocxEditor, part = 'word/document.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(await ed.save());
  return zip.file(part)!.async('string');
}

const count = (xml: string, text: string) => xml.split(text).length - 1;

describe('editing a text box', () => {
  it('edits in its own editor, writes the text into the document, one undo step in the body', async () => {
    const ed = await open('textboxes');
    const box = bodyShapes(ed).find((s) => s.texts[0] === '繞圖文字方塊')!;
    expect(ed.editShapeText(box.key, 0)).toBe(true);
    expect(ed.snapshot()!.target).toBe('textbox');
    const view = ed.activeView!;
    expect(view).not.toBe(ed.view);
    expect(view.dom.classList.contains('dx-doc')).toBe(true);
    expect(view.dom.getAttribute('aria-label')).toBe('文字方塊');
    // Page breaks don't go into a text box, and it has no comments.
    expect(ed.commentTarget()).toBeNull();
    ed.run(type('（已修改）'));
    ed.run(type('！'));
    expect(ed.isModified()).toBe(true);
    // The text box's own undo, while it is open.
    ed.undo();
    expect(ed.activeView!.state.doc.textContent).toBe('（已修改）繞圖文字方塊');
    ed.closeShapeText();
    expect(ed.snapshot()!.target).toBe('body');
    expect(bodyShapes(ed).find((s) => s.key === box.key)!.texts[0]).toBe('（已修改）繞圖文字方塊');

    const xml = await savedXml(ed);
    // In the DrawingML copy and in the VML copy for older Word.
    expect(count(xml, '（已修改）')).toBe(2);

    // One Ctrl+Z in the body takes the whole edit back: the document is as opened.
    ed.undo();
    expect(bodyShapes(ed).find((s) => s.key === box.key)!.texts[0]).toBe('繞圖文字方塊');
    expect(ed.isModified()).toBe(false);
    expect(count(await savedXml(ed), '（已修改）')).toBe(0);
  });

  it('text typed back as it was gives the text box back as read (saved unchanged)', async () => {
    const ed = await open('flowchart');
    const canvas = bodyShapes(ed).find((s) => s.shape.kind === 'canvas')!;
    const before = findShape(ed.view!.state.doc, canvas.key)!.node.attrs.xml;
    ed.editShapeText(canvas.key, 1);
    ed.run(type('甲'));
    ed.undo();
    ed.closeShapeText();
    const after = findShape(ed.view!.state.doc, canvas.key)!;
    expect(after.shape).toBe(canvas.shape);
    expect(after.node.attrs.xml).toBe(before);
  });

  it('opens with Enter on a selected shape and with a double-click; Esc goes back to the shape', async () => {
    const ed = await open('shapes');
    const rect = bodyShapes(ed).find((s) => s.texts[0] === '矩形')!;
    const found = findShape(ed.view!.state.doc, rect.key)!;
    ed.view!.dispatch(ed.view!.state.tr.setSelection(NodeSelection.create(ed.view!.state.doc, found.pos)));
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key: 'Enter' })));
    expect(ed.target).toBe('textbox');
    ed.activeView!.someProp('handleKeyDown', (f) => f(ed.activeView!, new KeyboardEvent('keydown', { key: 'Escape' })));
    expect(ed.target).toBe('body');
    expect((ed.view!.state.selection as NodeSelection).node?.attrs.shape?.key).toBe(rect.key);

    const el = ed.view!.dom.querySelector(`.dx-shape[data-shape-key="${rect.key}"] .dx-shape-text`)!;
    el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(ed.target).toBe('textbox');
    // A click elsewhere in the document closes it.
    ed.view!.dom.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(ed.target).toBe('body');
  });

  it('a shape that can’t be drawn here can’t be edited: the user is told to use Word', async () => {
    const notices: string[] = [];
    const ed = await open('shapes', { onNotice: (m: string) => notices.push(m) });
    const star = ed.view!.dom.querySelector('.dx-shape-fixed')!;
    star.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(ed.target).toBe('body');
    expect(notices).toContain('這個圖形只能在 Word 修改。');
  });

  it('read-only: no editing', async () => {
    const ed = await open('textboxes', { editable: false });
    expect(ed.editShapeText(bodyShapes(ed)[0].key, 0)).toBe(false);
  });

  it('a text box in the header is edited while the header is', async () => {
    const ed = await open('textboxes');
    ed.editHeaderFooter('header');
    const hdoc = ed.activeView!.state.doc;
    let key = '';
    hdoc.descendants((n) => {
      if (n.attrs?.shape) key = n.attrs.shape.key;
      return !key;
    });
    expect(ed.editShapeText(key, 0)).toBe(true);
    ed.run(type('新'));
    ed.closeShapeText();
    expect(ed.target).toBe('header');
    ed.closeHeaderFooter();
    const header = ed.model.headerFooters.find((h) => h.kind === 'header' && h.type === 'default')!;
    const xml = await savedXml(ed, header.part!);
    expect(count(xml, '新頁首文字方塊')).toBe(2);
  });
});

describe('find and replace, word count and spell check include text boxes', () => {
  it('lists each text box as a part, opens it when a match is there, and replaces in one undo step', async () => {
    const ed = await open('flowchart');
    const parts = ed.searchParts();
    const boxes = parts.filter((p) => p.area === 'textbox');
    expect(boxes.map((p) => p.doc.textContent)).toContain('收文登錄');
    expect(boxes.every((p) => p.label === '文字方塊')).toBe(true);
    // The body also says 流程 (its first paragraphs): replace in the body and the text boxes at once.
    const q = '文';
    const changes = parts
      .map((p) => ({ id: p.id, matches: findMatches(p.doc, q) }))
      .filter((c) => c.matches.length)
      .map((c) => ({ id: c.id, build: (state: EditorState) => replaceMatchesTr(state, findMatches(state.doc, q), '字') }));
    const changed = ed.changeSearchParts(changes);
    expect(changed.some((id) => id.startsWith('textbox|'))).toBe(true);
    expect(changed).toContain('body');
    const texts = () => bodyShapes(ed).find((s) => s.shape.kind === 'canvas')!.texts;
    expect(texts()).toContain('收字登錄');
    expect(texts()).toContain('發字');
    expect(ed.view!.state.doc.textContent).not.toContain('文');
    // One Ctrl+Z: body and text boxes back.
    ed.undo();
    expect(texts()).toContain('收文登錄');
    expect(ed.view!.state.doc.textContent).toContain('流程圖後的文字');
    expect(ed.isModified()).toBe(false);

    // Going to a match in a text box opens it.
    const id = ed.searchParts().find((p) => p.doc.textContent === '歸檔')!.id;
    const view = ed.openSearchPart(id);
    expect(view).toBe(ed.activeView);
    expect(ed.activeSearchPart()).toBe(id);
    expect(ed.target).toBe('textbox');
    // Replacing there (the part being edited).
    expect(ed.changeSearchPart(id, (s) => replaceMatchesTr(s, findMatches(s.doc, '歸檔'), '存檔'))).toBe(true);
    ed.closeShapeText();
    expect(texts()).toContain('存檔');
  });

  it('replaces in a closed text box of a header', async () => {
    const ed = await open('textboxes');
    const part = ed.searchParts().find((p) => p.area === 'textbox' && p.doc.textContent === '頁首文字方塊')!;
    expect(part.label).toBe('文字方塊（頁首）');
    expect(ed.changeSearchPart(part.id, (s) => replaceMatchesTr(s, findMatches(s.doc, '頁首'), '標題'))).toBe(true);
    expect(ed.searchParts().some((p) => p.doc.textContent === '標題文字方塊')).toBe(true);
    const header = ed.model.headerFooters.find((h) => h.kind === 'header' && h.type === 'default')!;
    expect(count(await savedXml(ed, header.part!), '標題文字方塊')).toBe(2);
  });

  it('counts the words of text boxes, as Word does', async () => {
    const ed = await open('flowchart');
    const bodyOnly = ed.view!.state.doc.textContent.replace(/\s/g, '').length;
    // 開始 收文登錄 需要會辦？ 會辦單位 發文 歸檔 結束 (21) + 填寫申請單 資料齊全？ (10)
    expect(ed.wordCount().chars).toBe(bodyOnly + 31);
  });

  it('spell check reaches the text box editor', async () => {
    const ed = await open('textboxes');
    ed.setSpellcheck(true);
    ed.editShapeText(bodyShapes(ed)[0].key, 0);
    expect(ed.activeView!.dom.getAttribute('spellcheck')).toBe('true');
    ed.setSpellcheck(false);
    expect(ed.activeView!.dom.getAttribute('spellcheck')).toBe('false');
  });
});
