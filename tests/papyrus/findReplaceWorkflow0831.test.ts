import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick, ref } from 'vue';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { findMatches, replaceAll, replaceCurrent, searchPlugin, searchState, setSearch } from '../../src/papyrus/editor/search';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: [history(), searchPlugin()] });
  const run = (cmd: Command) => cmd(state, (tr) => (state = state.apply(tr)));
  const text = () => state.doc.textContent;
  const matches = () => searchState(state).matches.map(({ from, to }) => state.doc.textBetween(from, to));
  const xml = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  const bytes = async () => writeDocx(state.doc, model);
  return { run, text, matches, xml, bytes, get state() { return state; }, set state(s) { state = s; } };
}

describe('QA20K-00831–00840 find and replace workflows', () => {
  it('QA20K-00831 treats punctuation as literal text when replacing an invoice token', async () => {
    const d = await open(`<w:p>${r('A[1]. A1 A[1].')}</w:p>`);
    d.run(setSearch('A[1].'));
    expect(d.matches()).toEqual(['A[1].', 'A[1].']);
    d.run(replaceAll('INV-7'));
    expect(d.text()).toBe('INV-7 A1 INV-7');
    expect(d.matches()).toEqual([]);
  });

  it('QA20K-00832 removes every obsolete label when replacement is intentionally blank', async () => {
    const d = await open(`<w:p>${r('DRAFT — DRAFT — APPROVED')}</w:p>`);
    d.run(setSearch('DRAFT'));
    expect(d.matches()).toEqual(['DRAFT', 'DRAFT']);
    expect(d.run(replaceAll(''))).toBe(true);
    expect(d.text()).toBe(' —  — APPROVED');
    expect(searchState(d.state).matches).toHaveLength(0);
  });

  it('QA20K-00833 replaces a repeated phrase in separate memo paragraphs without merging their boundaries', async () => {
    const d = await open(`<w:p>${r('Please review')}</w:p><w:p>${r('Please review today')}</w:p><w:p>${r('Do not review')}</w:p>`);
    d.run(setSearch('Please review'));
    expect(d.matches()).toEqual(['Please review', 'Please review']);
    d.run(replaceAll('Kindly check'));
    expect(d.state.doc.childCount).toBe(3);
    expect(d.state.doc.child(0).textContent).toBe('Kindly check');
    expect(d.state.doc.child(1).textContent).toBe('Kindly check today');
    expect(d.state.doc.child(2).textContent).toBe('Do not review');
  });

  it('QA20K-00834 case-sensitive bulk correction leaves differently cased names untouched', async () => {
    const d = await open(`<w:p>${r('Acme ACME acme')}</w:p>`);
    d.run(setSearch('Acme', true));
    expect(d.matches()).toEqual(['Acme']);
    d.run(replaceAll('Northwind'));
    expect(d.text()).toBe('Northwind ACME acme');
  });

  it('QA20K-00835 confines bulk replacement to matching table cells and preserves neighboring cells', async () => {
    const d = await open(`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${r('TBD')}</w:p></w:tc><w:tc><w:p>${r('Ready')}</w:p></w:tc></w:tr><w:tr><w:tc><w:p>${r('TBD')}</w:p></w:tc><w:tc><w:p>${r('TBD note')}</w:p></w:tc></w:tr></w:tbl>`);
    d.run(setSearch('TBD'));
    expect(d.matches()).toEqual(['TBD', 'TBD', 'TBD']);
    d.run(replaceAll('Done'));
    const table = d.state.doc.child(0);
    expect(table.childCount).toBe(2);
    expect(table.child(0).child(0).textContent).toBe('Done');
    expect(table.child(0).child(1).textContent).toBe('Ready');
    expect(table.child(1).child(0).textContent).toBe('Done');
    expect(table.child(1).child(1).textContent).toBe('Done note');
  });

  it('QA20K-00836 a misspelled query cannot alter text or create a replacement result', async () => {
    const d = await open(`<w:p>${r('Quarterly close is complete')}</w:p>`);
    d.run(setSearch('Quaterly'));
    expect(searchState(d.state).matches).toHaveLength(0);
    expect(d.run(replaceCurrent('Monthly'))).toBe(false);
    expect(d.run(replaceAll('Monthly'))).toBe(false);
    expect(d.text()).toBe('Quarterly close is complete');
  });

  it('QA20K-00837 treats overlapping substring candidates as non-overlapping replacement hits', async () => {
    const d = await open(`<w:p>${r('banana')}</w:p>`);
    d.run(setSearch('ana'));
    expect(d.matches()).toEqual(['ana']);
    d.run(replaceAll('X'));
    expect(d.text()).toBe('bXna');
  });

  it('QA20K-00838 escapes replacement characters in DOCX XML and reloads them as text', async () => {
    const d = await open(`<w:p>${r('LEGAL')}</w:p>`);
    d.run(setSearch('LEGAL'));
    d.run(replaceCurrent('R&D <North>'));
    const xml = await d.xml();
    expect(xml).toContain('R&amp;D &lt;North&gt;');
    const reloaded = await readDocx(await d.bytes());
    expect(reloaded.doc.textContent).toBe('R&D <North>');
  });

  it('QA20K-00839 launching Find from a selected phrase seeds and selects that phrase in the panel', async () => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= zero;
    const { DocxEditor } = await import('../../src/papyrus/editor/core');
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p>${r('Please approve this amount')}</w:p>${SECT}</w:body></w:document>`);
    const host = document.createElement('div'); const panelHost = document.createElement('div'); document.body.append(host, panelHost);
    const snapshot = ref<unknown>(null); const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s) });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    const from = 1 + 'Please '.length; const to = from + 'approve'.length;
    editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, from, to)));
    const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
    const app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value, replace: false }) }); app.mount(panelHost);
    await nextTick(); await new Promise((resolve) => setTimeout(resolve, 170)); await nextTick();
    const input = panelHost.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    expect(input.value).toBe('approve');
    expect(input.selectionStart).toBe(0);
    expect(input.selectionEnd).toBe('approve'.length);
    expect(searchState(editor.view!.state).matches.map((m) => editor.view!.state.doc.textBetween(m.from, m.to))).toEqual(['approve']);
    app.unmount(); editor.destroy(); host.remove(); panelHost.remove();
  });

  it('QA20K-00840 Find refuses to seed a multiline query from a cross-paragraph selection', async () => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= zero;
    const { DocxEditor } = await import('../../src/papyrus/editor/core');
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p>${r('first line')}</w:p><w:p>${r('second line')}</w:p>${SECT}</w:body></w:document>`);
    const host = document.createElement('div'); const panelHost = document.createElement('div'); document.body.append(host, panelHost);
    const snapshot = ref<unknown>(null); const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s) });
    await editor.open(await zip.generateAsync({ type: 'uint8array' }));
    const firstEnd = editor.view!.state.doc.child(0).nodeSize - 1;
    editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, firstEnd - 4, firstEnd + 4)));
    expect(editor.view!.state.doc.textBetween(editor.view!.state.selection.from, editor.view!.state.selection.to, '\n', '\n')).toContain('\n');
    const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
    const app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value, replace: false }) }); app.mount(panelHost);
    await nextTick();
    const input = panelHost.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    expect(input.value).toBe('');
    expect(searchState(editor.view!.state).query).toBe('');
    app.unmount(); editor.destroy(); host.remove(); panelHost.remove();
  });
});
