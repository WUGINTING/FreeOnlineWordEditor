// Writes the QA section document after section edits, so Word can confirm the result
// (scripts/inspect-sections-in-word.ps1). Opt-in:
//   QA_SECTIONS_DOCX=<qa file> DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-sections.test.ts
import { beforeAll, describe, it } from 'vitest';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { writeDocx } from '../../src/papyrus/docx/writer';

const qa = process.env.QA_SECTIONS_DOCX;
const out = process.env.DOCX_EXPORT;

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
});

async function open() {
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host);
  await ed.open(readFileSync(qa!));
  return ed;
}

const save = async (ed: DocxEditor, name: string) => writeFileSync(join(out!, name), await writeDocx(ed.view!.state.doc, ed.model));

describe.skipIf(!qa || !out)('export section edits', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });

    // Unchanged.
    let ed = await open();
    await save(ed, 'unchanged.docx');

    // A section break inside the first paragraph of section 1.
    ed = await open();
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 4)));
    ed.insertSectionBreak();
    await save(ed, 'section-break.docx');

    // Section 1 turned landscape.
    ed = await open();
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 2)));
    const p = ed.cursorSection().page;
    ed.setPageSetup({ ...p, width: p.height, height: p.width }, 'section');
    await save(ed, 'landscape-first.docx');

    // Section 1 gets "different first page" and a first-page header.
    ed = await open();
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 2)));
    ed.setTitlePage(true);
    (ed as any).pages = [{ ...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 }];
    ed.editHeaderFooter('header', 0);
    ed.run((s, d) => (d?.(s.tr.insertText('QA_FIRST_PAGE_HEADER')), true));
    ed.closeHeaderFooter();
    await save(ed, 'first-page-header.docx');
  }, 60_000);
});
