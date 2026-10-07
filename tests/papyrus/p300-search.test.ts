// persona-300: 「114」 finds full-width 「１１４」 unless 區分全形／半形 is on; one Ctrl+Z after 全部取代
// undoes it in the body and the headers / footers alike.
import { describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { findMatches, replaceMatchesTr, searchPlugin, searchState, setSearch } from '../../src/papyrus/editor/search';
import { schema } from '../../src/papyrus/editor/schema';
import { domStubs, key, setup, typeInto } from './p300Helpers';

domStubs();

const doc = (...texts: string[]) => schema.nodes.doc.create(null, texts.map((t) => schema.nodes.paragraph.create(null, schema.text(t))));

describe('full-width and half-width', () => {
  it('match each other by default; not with 區分全形／半形', () => {
    const d = doc('府授字第１１４０００號', '114 年度', 'ＡＢＣ　abc');
    expect(findMatches(d, '114')).toHaveLength(2);
    expect(findMatches(d, '１１４')).toHaveLength(2);
    expect(findMatches(d, '114', false, true)).toHaveLength(1);
    expect(findMatches(d, 'abc abc')).toHaveLength(1); // ＡＢＣ + ideographic space + abc, case folded
    expect(findMatches(d, 'abc abc', true)).toHaveLength(0);
    // Positions stay those of the document's characters.
    const m = findMatches(d, '114')[0];
    expect(d.textBetween(m.from, m.to)).toBe('１１４');
  });

  it('the editor search follows the option, and replacing keeps the rest of the text', () => {
    let state = EditorState.create({ schema, doc: doc('第１１４號與第114號'), plugins: [searchPlugin()] });
    setSearch('114')(state, (tr) => (state = state.apply(tr)));
    expect(searchState(state).matches).toHaveLength(2);
    setSearch('114', false, true)(state, (tr) => (state = state.apply(tr)));
    expect(searchState(state).matches).toHaveLength(1);
    setSearch('114')(state, (tr) => (state = state.apply(tr)));
    state = state.apply(replaceMatchesTr(state, searchState(state).matches, '115'));
    expect(state.doc.textContent).toBe('第115號與第115號');
  });

  it('the find panel has the option, off by default', async () => {
    const d = await setup('<w:p><w:r><w:t>１１４ 與 114</w:t></w:r></w:p>', 'FindReplace.vue', { replace: false });
    await d.refresh();
    const box = d.panel.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    await typeInto(box, '114');
    key(box, { key: 'Enter' });
    expect(searchState(d.view.state).matches).toHaveLength(2);
    const option = Array.from(d.panel.querySelectorAll('label')).find((l) => l.textContent!.includes('區分全形／半形'))!.querySelector('input')!;
    expect(option.checked).toBe(false);
    option.click();
    await d.refresh();
    expect(searchState(d.view.state).matches).toHaveLength(1);
    d.done();
  });
});

describe('one undo for a change in several parts', () => {
  it('changeSearchParts: Ctrl+Z in the body undoes the body and the header; Ctrl+Y redoes both; a later edit ends it', async () => {
    const d = await setup('<w:p><w:r><w:t>甲 2025</w:t></w:r></w:p>');
    // A header with 2025.
    d.editor.editHeaderFooter('header');
    const hv = d.editor.activeView!;
    hv.dispatch(hv.state.tr.insertText('頁首 2025', 1));
    d.editor.closeHeaderFooter();
    const parts = d.editor.searchParts();
    expect(parts.map((p) => p.area)).toEqual(['body', 'header']);
    const changed = d.editor.changeSearchParts(
      parts.map((p) => ({ id: p.id, build: (s: EditorState) => replaceMatchesTr(s, findMatches(s.doc, '2025'), '2026') })),
    );
    expect(changed).toHaveLength(2);
    const texts = () => d.editor.searchParts().map((p) => p.doc.textContent);
    expect(texts()).toEqual(['甲 2026', '頁首 2026']);
    expect(d.editor.snapshot()!.canUndo).toBe(true);
    // Ctrl+Z in the body.
    key(d.view.dom, { key: 'z', ctrlKey: true });
    expect(texts()).toEqual(['甲 2025', '頁首 2025']);
    key(d.view.dom, { key: 'y', ctrlKey: true });
    expect(texts()).toEqual(['甲 2026', '頁首 2026']);
    // Typing in the body afterwards: Ctrl+Z undoes the typing only.
    d.view.dispatch(d.view.state.tr.insertText('！', d.view.state.doc.content.size - 1));
    d.editor.undo();
    expect(texts()).toEqual(['甲 2026', '頁首 2026']);
    d.done();
  });
});
