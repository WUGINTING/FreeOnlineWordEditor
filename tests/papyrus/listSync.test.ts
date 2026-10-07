// Lists made in the editor are kept in step with the document (madeListsSync) only when an edit
// can change which lists are used: typing does not rescan the documents, a paste and undo do.
import { describe, expect, it, vi } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { history, redo, undo } from 'prosemirror-history';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';
import { madeListsSync } from '../../src/papyrus/editor/listMarkers';
import * as numbering from '../../src/papyrus/docx/numbering';

vi.mock('../../src/papyrus/docx/numbering', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/papyrus/docx/numbering')>();
  return { ...actual, usedNumIds: vi.fn(actual.usedNumIds) };
});

const parse = (html: string) => {
  const div = document.createElement('div');
  div.innerHTML = html;
  return PMDOMParser.fromSchema(schema).parseSlice(div);
};

describe('made lists sync', () => {
  it('typing does not rescan; a paste, deleting a list item, undo and redo do (and undo still takes the list out)', async () => {
    const used = vi.mocked(numbering.usedNumIds);
    const { doc, model } = await readDocx(await blankPackage().generateAsync({ type: 'uint8array' }));
    let state = EditorState.create({ schema, doc, plugins: [history(), madeListsSync(() => model.numbering)] });
    const originalNums = structuredClone(model.numbering.nums);
    const slice = adoptPastedLists(parse('<ol><li>x</li><li>y</li></ol>'), model.numbering);
    used.mockClear();
    state = state.apply(state.tr.replaceSelection(slice));
    expect(used).toHaveBeenCalledTimes(1); // the paste
    // Typing, in a list item and elsewhere, and Enter in a plain paragraph: no rescan.
    used.mockClear();
    state = state.apply(state.tr.insertText('abc', 2));
    state = state.apply(state.tr.insert(state.doc.content.size, schema.nodes.paragraph.create(null, schema.text('plain'))));
    state = state.apply(state.tr.insertText('def', state.doc.content.size - 2));
    state = state.apply(state.tr.split(state.doc.content.size - 3));
    expect(used).not.toHaveBeenCalled();
    // Deleting the list items: rescanned, the list is set aside.
    const listEnd = state.doc.child(0).nodeSize + state.doc.child(1).nodeSize;
    const deleted = state.apply(state.tr.delete(0, listEnd));
    expect(used).toHaveBeenCalledTimes(1);
    expect(model.numbering.nums).toEqual(originalNums);
    // Undo brings it back.
    used.mockClear();
    let back = deleted;
    undo(back, (tr) => (back = back.apply(tr)));
    expect(used).toHaveBeenCalled();
    expect(Object.keys(model.numbering.nums).length).toBeGreaterThan(Object.keys(originalNums).length);
    // Undo everything: the paste's list definition goes out again; redo brings it back.
    for (let i = 0; i < 20; i++) if (!undo(back, (tr) => (back = back.apply(tr)))) break;
    expect(back.doc.eq(doc)).toBe(true);
    expect(model.numbering.nums).toEqual(originalNums);
    redo(back, (tr) => (back = back.apply(tr)));
    expect(Object.keys(model.numbering.nums).length).toBeGreaterThan(Object.keys(originalNums).length);
  });
});
