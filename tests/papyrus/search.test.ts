// Find / replace: matches inside a paragraph across formatting runs, never across structure;
// replacements keep the first character's formatting and wrappers; "replace all" is one undo.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import {
  findMatches, findNext, replaceAll, replaceCurrent, searchPlugin, searchState, setSearch,
} from '../../src/papyrus/editor/search';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: [history(), searchPlugin()] });
  const run = (cmd: Command) => cmd(state, (tr) => (state = state.apply(tr)));
  const save = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  const texts = () => findMatches(state.doc, searchState(state).query, searchState(state).caseSensitive).map((m) => state.doc.textBetween(m.from, m.to));
  return { run, save, texts, get state() { return state; }, set state(s) { state = s; } };
}

const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

describe('finding', () => {
  it('matches across formatting runs and table cells, case-insensitively by default', async () => {
    const d = await open(
      `<w:p>${r('Hel', '<w:b/>')}${r('lo world')}</w:p>` +
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${r('say HELLO')}</w:p></w:tc></w:tr></w:tbl>`,
    );
    d.run(setSearch('hello'));
    expect(d.texts()).toEqual(['Hello', 'HELLO']);
    d.run(setSearch('hello', true));
    expect(searchState(d.state).matches).toHaveLength(0);
    d.run(setSearch('Hello', true));
    expect(d.texts()).toEqual(['Hello']);
  });

  it('never matches across paragraphs, links, tracked insertions or atoms; skips bookmarks', async () => {
    const d = await open(
      `<w:p>${r('ab')}</w:p><w:p>${r('cd')}</w:p>` +
        `<w:p>${r('xy')}<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">${r('zw')}</w:ins></w:p>` +
        `<w:p>${r('pq')}<w:r><w:tab/></w:r>${r('rs')}</w:p>` +
        `<w:p>${r('mn')}<w:bookmarkStart w:id="0" w:name="bm"/>${r('op')}<w:bookmarkEnd w:id="0"/></w:p>`,
    );
    for (const q of ['bc', 'yz', 'qr']) {
      d.run(setSearch(q));
      expect(searchState(d.state).matches, q).toHaveLength(0);
    }
    d.run(setSearch('mnop'));
    expect(searchState(d.state).matches).toHaveLength(1);
  });

  it('the first match after the cursor is current; Enter / Shift+Enter walk and wrap', async () => {
    const d = await open(`<w:p>${r('a1 a2 a3')}</w:p>`);
    d.state = d.state.apply(d.state.tr.setSelection(TextSelection.create(d.state.doc, 3)));
    d.run(setSearch('a'));
    expect(searchState(d.state).current).toBe(1);
    d.run(findNext(1));
    expect(d.state.doc.textBetween(d.state.selection.from, d.state.selection.to + 1)).toBe('a2');
    d.run(findNext(1));
    d.run(findNext(1));
    expect(searchState(d.state).current).toBe(0);
    d.run(findNext(-1));
    expect(searchState(d.state).current).toBe(2);
  });

  it('recomputes the matches when the document changes', async () => {
    const d = await open(`<w:p>${r('cat')}</w:p>`);
    d.run(setSearch('cat'));
    d.state = d.state.apply(d.state.tr.insertText(' cat', d.state.doc.child(0).nodeSize - 1));
    expect(searchState(d.state).matches).toHaveLength(2);
  });
});

describe('replacing', () => {
  it('takes the first character’s formatting, keeps bookmarks, stays inside the tracked insertion', async () => {
    const d = await open(
      `<w:p>${r('Hel', '<w:b/>')}${r('lo world')}</w:p>` +
        `<w:p><w:ins w:id="7" w:author="Bob" w:date="2026-01-01T00:00:00Z">${r('old text', '<w:i/>')}</w:ins></w:p>` +
        `<w:p>${r('mn')}<w:bookmarkStart w:id="0" w:name="bm"/>${r('op')}<w:bookmarkEnd w:id="0"/></w:p>`,
    );
    d.run(setSearch('hello'));
    d.run(replaceCurrent('Goodbye'));
    d.run(setSearch('old'));
    d.run(replaceCurrent('new'));
    d.run(setSearch('mnop'));
    d.run(replaceCurrent('X'));
    const xml = await d.save();
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>Goodbye</w:t></w:r>');
    expect(xml).toContain('<w:t xml:space="preserve"> world</w:t>');
    expect(xml).toMatch(/<w:ins w:id="7" w:author="Bob"[^>]*><w:r><w:rPr><w:i\/><\/w:rPr><w:t>new text<\/w:t><\/w:r><\/w:ins>/);
    expect(xml).toContain('<w:t>X</w:t></w:r><w:bookmarkStart w:id="0" w:name="bm"/><w:bookmarkEnd w:id="0"/>');
  });

  it('replace moves on to the next match', async () => {
    const d = await open(`<w:p>${r('a-a-a')}</w:p>`);
    d.run(setSearch('a'));
    d.run(replaceCurrent('bb'));
    expect(d.state.doc.textContent).toBe('bb-a-a');
    expect(searchState(d.state).current).toBe(0);
    expect(d.state.doc.textBetween(d.state.selection.from, d.state.selection.to)).toBe('a');
    expect(d.state.selection.from).toBe(1 + 3);
  });

  it('replace all is one transaction: a single undo brings every match back', async () => {
    const d = await open(`<w:p>${r('one two one')}</w:p><w:p>${r('One more', '<w:u w:val="single"/>')}</w:p>`);
    d.state = d.state.apply(d.state.tr.insertText('!', 1)); // earlier typing: a separate undo step
    d.run(setSearch('one'));
    d.run(replaceAll('1'));
    expect(d.state.doc.textContent).toBe('!1 two 11 more');
    expect(searchState(d.state).matches).toHaveLength(0);
    d.run(undo);
    expect(d.state.doc.textContent).toBe('!one two oneOne more');
    const xml = await d.save();
    expect(xml).toContain('<w:u w:val="single"/></w:rPr><w:t>One more</w:t>');
  });
});

describe('in the editor', () => {
  it('searches the body, or the header while it is being edited; pictures get resize handles', async () => {
    (globalThis as any).ResizeObserver ??= class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
    const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= zero;
    const { DocxEditor } = await import('../../src/papyrus/editor/core');
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));
    const typeText = (text: string): Command => (s, dispatch) => (dispatch?.(s.tr.insertText(text)), true);
    const exec = (cmd: Command) => {
      const v = editor.activeView!;
      return cmd(v.state, v.dispatch, v);
    };
    editor.run(typeText('body word'));
    editor.run((s, dispatch) => {
      dispatch?.(s.tr.replaceSelectionWith(schema.nodes.image.create({ src: 'data:image/gif;base64,R0lGODlhAQABAAAAACw=', width: 10, height: 10 })));
      return true;
    });
    expect(host.querySelector('.dx-img-wrap .dx-img-handle')).toBeTruthy();
    exec(setSearch('word'));
    expect(searchState(editor.view!.state).matches).toHaveLength(1);
    expect(host.querySelectorAll('.dx-search-current')).toHaveLength(1);

    editor.editHeaderFooter('header');
    editor.run(typeText('header word word'));
    exec(setSearch('word'));
    expect(searchState(editor.activeView!.state).matches).toHaveLength(2);
    exec(replaceAll('W'));
    expect(editor.activeView!.state.doc.textContent).toBe('header W W');
    expect(editor.view!.state.doc.textContent).toContain('body word');
    editor.destroy();
    host.remove();
  });
});
