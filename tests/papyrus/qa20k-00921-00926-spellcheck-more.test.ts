import { beforeAll, describe, expect, it } from 'vitest';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} })) as any;
});

async function open(spellcheck = false) {
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host, { spellcheck });
  await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));
  const done = () => { editor.destroy(); host.remove(); };
  return { host, editor, done };
}

describe('spellcheck preference across editing surfaces', () => {
  it('QA20K-00921 applies the configured initial preference to the body and first opened footer', async () => {
    const { editor, done } = await open(true);
    try {
      expect(editor.view!.dom.getAttribute('spellcheck')).toBe('true');
      editor.editHeaderFooter('footer');
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');
      expect(editor.spellcheckOn).toBe(true);
    } finally { done(); }
  });

  it('QA20K-00922 turning spellcheck off from a header updates both editor surfaces', async () => {
    const { editor, done } = await open(true);
    try {
      editor.editHeaderFooter('header');
      editor.setSpellcheck(false);
      expect(editor.view!.dom.getAttribute('spellcheck')).toBe('false');
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('false');
      expect(editor.spellcheckOn).toBe(false);
    } finally { done(); }
  });

  it('QA20K-00923 turning spellcheck on from a footer leaves document content and modified state intact', async () => {
    const { editor, done } = await open(false);
    try {
      editor.editHeaderFooter('footer');
      const before = editor.view!.state.doc.toJSON();
      editor.setSpellcheck(true);
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');
      expect(editor.view!.dom.getAttribute('spellcheck')).toBe('true');
      expect(editor.view!.state.doc.toJSON()).toEqual(before);
      expect(editor.isModified()).toBe(false);
    } finally { done(); }
  });

  it('QA20K-00924 preserves spellcheck in independently created header and footer parts', async () => {
    const { editor, done } = await open(false);
    try {
      editor.setSpellcheck(true);
      editor.editHeaderFooter('header');
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');
      editor.run((state, dispatch) => { dispatch?.(state.tr.insertText('HEAD')); return true; });
      editor.closeHeaderFooter(false);
      editor.editHeaderFooter('footer');
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');
      editor.run((state, dispatch) => { dispatch?.(state.tr.insertText('FOOT')); return true; });
      expect(editor.model.headerFooters.map((part) => part.kind)).toEqual(['header', 'footer']);
    } finally { done(); }
  });

  it('QA20K-00925 changing spellcheck while a header is closed applies the latest choice on reopen', async () => {
    const { editor, done } = await open(false);
    try {
      editor.editHeaderFooter('header');
      editor.setSpellcheck(true);
      editor.closeHeaderFooter(false);
      editor.setSpellcheck(false);
      editor.editHeaderFooter('header');
      expect(editor.view!.dom.getAttribute('spellcheck')).toBe('false');
      expect(editor.activeView!.dom.getAttribute('spellcheck')).toBe('false');
    } finally { done(); }
  });

  it('QA20K-00926 changing spellcheck does not consume the next content undo step', async () => {
    const { editor, done } = await open(false);
    try {
      editor.run((state, dispatch) => { dispatch?.(state.tr.insertText('Draft')); return true; });
      editor.setSpellcheck(true);
      expect(editor.view!.state.doc.textContent).toContain('Draft');
      expect(editor.view!.dom.getAttribute('spellcheck')).toBe('true');
      expect(editor.snapshot()!.canUndo).toBe(true);
      editor.undo();
      expect(editor.view!.state.doc.textContent).not.toContain('Draft');
      expect(editor.spellcheckOn).toBe(true);
    } finally { done(); }
  });
});
