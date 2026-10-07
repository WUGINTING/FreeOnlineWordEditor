// Leaving a header/footer and coming back must not lose its undo history.
import { beforeAll, describe, expect, it } from 'vitest';
import type { Command } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';

beforeAll(() => {
  // jsdom has no layout; the editor only needs the API to exist.
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  // Undo scrolls the cursor into view, which measures text ranges.
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

const type = (text: string): Command => (state, dispatch) => {
  dispatch?.(state.tr.insertText(text));
  return true;
};

describe('header / footer undo', () => {
  it('keeps each header’s undo history when it is closed and reopened', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));

    editor.editHeaderFooter('header');
    editor.run(type('Draft v1'));
    expect(editor.snapshot()!.target).toBe('header');
    expect(editor.snapshot()!.canUndo).toBe(true);

    editor.closeHeaderFooter();
    expect(editor.snapshot()!.target).toBe('body');

    editor.editHeaderFooter('header');
    expect(editor.snapshot()!.canUndo).toBe(true);
    editor.undo();
    const header = editor.model.headerFooters.find((h) => h.kind === 'header')!;
    expect(header.doc.textContent).toBe('');
    // ...and redo brings it back.
    editor.redo();
    expect(header.doc.textContent).toBe('Draft v1');

    editor.destroy();
    host.remove();
  });

  it('a reopened header shows changes made to it while closed, with a fresh history', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));

    editor.editHeaderFooter('footer');
    editor.run(type('A'));
    editor.closeHeaderFooter();
    // Changed from outside (e.g. a future "replace all"): the kept state is outdated.
    const footer = editor.model.headerFooters.find((h) => h.kind === 'footer')!;
    const { schema } = await import('../../src/papyrus/editor/schema');
    footer.doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('B')));

    editor.editHeaderFooter('footer');
    expect(footer.doc.textContent).toBe('B');
    expect(editor.snapshot()!.canUndo).toBe(false);

    editor.destroy();
    host.remove();
  });
});
