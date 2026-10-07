import { describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { blankPackage } from '../src/papyrus/docx/template';
import { readDocx } from '../src/papyrus/docx/reader';
import { schema } from '../src/papyrus/editor/schema';
import { DocxEditor } from '../src/papyrus/editor/core';
import { findNext, searchPlugin, searchState, setSearch } from '../src/papyrus/editor/search';
import { loadSfc } from './papyrus/sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

function packageWithBody(body: string): JSZip {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  return zip;
}

describe('QA20K find/replace review fixes', () => {
  it('QA20K-00652 treats hyperlink text as searchable but will not join it to neighboring plain text', async () => {
    const zip = packageWithBody(
      '<w:p><w:r><w:t>a</w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>bc</w:t></w:r></w:hyperlink><w:r><w:t>d</w:t></w:r></w:p>',
    );
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/contract" TargetMode="External"/></Relationships>',
    ));
    const { doc } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const linkHrefs: string[] = [];
    doc.descendants((node) => {
      if (node.isText) {
        const link = schema.marks.link.isInSet(node.marks);
        if (link) linkHrefs.push(link.attrs.href);
      }
    });
    expect(linkHrefs).toEqual(['https://example.test/contract']);

    let state = EditorState.create({ schema, doc, plugins: [searchPlugin()] });
    const run = (cmd: Command) =>
      cmd(state, (tr) => (state = state.apply(tr)));
    run(setSearch('ab'));
    expect(searchState(state).matches).toHaveLength(0);
    run(setSearch('cd'));
    expect(searchState(state).matches).toHaveLength(0);
    run(setSearch('bc'));
    expect(searchState(state).matches).toHaveLength(1);
    expect(state.doc.textBetween(searchState(state).matches[0].from, searchState(state).matches[0].to)).toBe('bc');
  });

  it('QA20K-00653 Enter and Shift+Enter key events advance and reverse the Find panel selection', async () => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    const host = document.createElement('div');
    const panelHost = document.createElement('div');
    document.body.append(host, panelHost);
    const snapshot = ref<unknown>(null);
    const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s) });
    let app: ReturnType<typeof createApp> | null = null;
    try {
      const zip = packageWithBody('<w:p><w:r><w:t>a1 a2 a3</w:t></w:r></w:p>');
      await editor.open(await zip.generateAsync({ type: 'uint8array' }));
      editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, 3)));
      const FindReplace = await loadSfc('src/papyrus/vue/FindReplace.vue');
      app = createApp({ render: () => h(FindReplace, { editor, snapshot: snapshot.value, replace: false }) });
      app.mount(panelHost);
      await nextTick();
      const input = panelHost.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
      input.value = 'a';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      await nextTick();
      await new Promise((resolve) => setTimeout(resolve, 200));
      expect(searchState(editor.view!.state).matches).toHaveLength(3);
      expect(searchState(editor.view!.state).current).toBe(1);

      const press = (shiftKey = false) => {
        const event = new KeyboardEvent('keydown', { key: 'Enter', shiftKey, bubbles: true, cancelable: true });
        input.dispatchEvent(event);
        return event.defaultPrevented;
      };
      const expectSelected = (index: number) => {
        const state = editor.view!.state;
        const current = searchState(state);
        expect(current.current).toBe(index);
        expect([state.selection.from, state.selection.to]).toEqual([
          current.matches[index].from,
          current.matches[index].to,
        ]);
        expect(state.doc.textBetween(state.selection.from, state.selection.to)).toBe('a');
      };
      expect(press()).toBe(true);
      expectSelected(1);
      expect(press()).toBe(true);
      expectSelected(2);
      expect(press()).toBe(true);
      expectSelected(0);
      expect(press(true)).toBe(true);
      expectSelected(2);
    } finally {
      app?.unmount();
      editor.destroy();
      host.remove();
      panelHost.remove();
    }
  });

  it('QA20K-00654 records initial match position then updates count positions and current selection after typing', async () => {
    const zip = packageWithBody('<w:p><w:r><w:t>cat</w:t></w:r></w:p>');
    const { doc } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    let state = EditorState.create({ schema, doc, plugins: [searchPlugin()] });
    const run = (cmd: Command) =>
      cmd(state, (tr) => (state = state.apply(tr)));
    const positions = () => searchState(state).matches.map(({ from, to }) => [from, to]);
    run(setSearch('cat'));
    expect(positions()).toEqual([[1, 4]]);
    expect(searchState(state).current).toBe(0);

    state = state.apply(state.tr.insertText(' cat', state.doc.child(0).nodeSize - 1));
    expect(positions()).toEqual([[1, 4], [5, 8]]);
    expect(searchState(state).current).toBe(0);
    run(findNext(1));
    run(findNext(1));
    expect(searchState(state).current).toBe(1);
    expect(state.doc.textBetween(state.selection.from, state.selection.to)).toBe('cat');
    expect(positions()).toEqual([[1, 4], [5, 8]]);
  });
});
