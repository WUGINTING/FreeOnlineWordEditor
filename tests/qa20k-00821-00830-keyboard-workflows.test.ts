import { beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor } from '../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const p = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host); await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return { editor, host };
}

function textStart(editor: DocxEditor, text: string) {
  let found = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text!.includes(text)) found = pos + node.text!.indexOf(text);
  });
  if (found < 0) throw new Error(`Text not found: ${text}`);
  return found;
}

function select(editor: DocxEditor, text: string, from = 0, to = text.length) {
  const start = textStart(editor, text);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start + from, start + to)));
  return start;
}

function caret(editor: DocxEditor, text: string, offset: number) {
  const start = textStart(editor, text);
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start + offset)));
  return start + offset;
}

function press(editor: DocxEditor, key: string, options: { ctrl?: boolean; shift?: boolean } = {}) {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ctrlKey: options.ctrl, shiftKey: options.shift });
  editor.view!.dom.dispatchEvent(event);
  return event;
}

function runs(editor: DocxEditor) {
  const result: { text: string; marks: string[] }[] = [];
  editor.view!.state.doc.descendants((node) => {
    if (node.isText) result.push({ text: node.text!, marks: node.marks.map((mark) => mark.type.name) });
  });
  return result;
}

function dispose(editor: DocxEditor, host: HTMLElement) { editor.destroy(); host.remove(); }

describe('QA20K everyday keyboard workflows 00821–00830', () => {
  it('QA20K-00821 applies Ctrl+I to the selected approval phrase only', async () => {
    const { editor, host } = await open(p('Draft approved'));
    select(editor, 'approved');
    const event = press(editor, 'i', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Draft approved');
    expect(runs(editor)).toEqual([{ text: 'Draft ', marks: [] }, { text: 'approved', marks: ['italic'] }]);
    dispose(editor, host);
  });

  it('QA20K-00822 applies Ctrl+U to a selected policy number while preserving both neighbors', async () => {
    const { editor, host } = await open(p('See policy 4.2 today'));
    select(editor, '4.2');
    const event = press(editor, 'u', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('See policy 4.2 today');
    expect(runs(editor)).toEqual([
      { text: 'See policy ', marks: [] }, { text: '4.2', marks: ['underline'] }, { text: ' today', marks: [] },
    ]);
    dispose(editor, host);
  });

  it('QA20K-00823 uses Ctrl+Z to undo one typed character without removing its preceding word', async () => {
    const { editor, host } = await open(p('Budget'));
    const end = textStart(editor, 'Budget') + 'Budget'.length;
    editor.view!.dispatch(editor.view!.state.tr.insertText('X', end));
    expect(editor.view!.state.doc.textContent).toBe('BudgetX');
    const event = press(editor, 'z', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Budget');
    dispose(editor, host);
  });

  it('QA20K-00824 uses Ctrl+Y to redo the most recently undone insertion', async () => {
    const { editor, host } = await open(p('Invoice'));
    const end = textStart(editor, 'Invoice') + 'Invoice'.length;
    editor.view!.dispatch(editor.view!.state.tr.insertText('!', end));
    press(editor, 'z', { ctrl: true });
    expect(editor.view!.state.doc.textContent).toBe('Invoice');
    const event = press(editor, 'y', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Invoice!');
    dispose(editor, host);
  });

  it('QA20K-00825 centers only the current quarterly update paragraph with Ctrl+E', async () => {
    const { editor, host } = await open(p('Quarterly update') + p('Appendix'));
    caret(editor, 'Quarterly update', 5);
    const event = press(editor, 'e', { ctrl: true });
    const paragraphs: any[] = [];
    editor.view!.state.doc.descendants((node) => { if (node.type.name === 'paragraph') paragraphs.push(node); });
    expect(event.defaultPrevented).toBe(true);
    expect(paragraphs.map((node) => [node.textContent, node.attrs.align])).toEqual([['Quarterly update', 'center'], ['Appendix', null]]);
    dispose(editor, host);
  });

  it('QA20K-00826 justifies a selected meeting-notes paragraph with Ctrl+J', async () => {
    const { editor, host } = await open(p('Meeting notes'));
    select(editor, 'Meeting notes');
    const event = press(editor, 'j', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.firstChild!.attrs.align).toBe('justify');
    expect(editor.view!.state.doc.textContent).toBe('Meeting notes');
    dispose(editor, host);
  });

  it('QA20K-00827 resets a centered memo paragraph to left alignment with Ctrl+L', async () => {
    const { editor, host } = await open(p('Memo body', '<w:jc w:val="center"/>'));
    expect(editor.view!.state.doc.firstChild!.attrs.align).toBe('center');
    caret(editor, 'Memo body', 3);
    const event = press(editor, 'l', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.firstChild!.attrs.align).toBeNull();
    expect(editor.view!.state.doc.textContent).toBe('Memo body');
    dispose(editor, host);
  });

  it('QA20K-00828 undoes a formatting shortcut as one history step while preserving text', async () => {
    const { editor, host } = await open(p('Draft approval'));
    select(editor, 'approval');
    press(editor, 'i', { ctrl: true });
    expect(runs(editor)).toEqual([{ text: 'Draft ', marks: [] }, { text: 'approval', marks: ['italic'] }]);
    const event = press(editor, 'z', { ctrl: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Draft approval');
    expect(runs(editor)).toEqual([{ text: 'Draft approval', marks: [] }]);
    dispose(editor, host);
  });

  it('QA20K-00829 uses Shift+Tab to outdent a nested numbered item one level', async () => {
    const { editor, host } = await open(
      p('Parent', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
      p('Child', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'),
    );
    caret(editor, 'Child', 2);
    const event = press(editor, 'Tab', { shift: true });
    const paragraphs: any[] = [];
    editor.view!.state.doc.descendants((node) => { if (node.type.name === 'paragraph') paragraphs.push(node); });
    expect(event.defaultPrevented).toBe(true);
    expect(paragraphs.map((node) => [node.textContent, node.attrs.numId, node.attrs.ilvl])).toEqual([
      ['Parent', '1', 0], ['Child', '1', 0],
    ]);
    dispose(editor, host);
  });

  it('QA20K-00830 applies Ctrl+Shift+X strikethrough to only the selected obsolete amount', async () => {
    const { editor, host } = await open(p('Old 125 New 130'));
    select(editor, '125');
    const event = press(editor, 'x', { ctrl: true, shift: true });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Old 125 New 130');
    expect(runs(editor)).toEqual([
      { text: 'Old ', marks: [] }, { text: '125', marks: ['strike'] }, { text: ' New 130', marks: [] },
    ]);
    dispose(editor, host);
  });
});
