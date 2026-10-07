import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { countDocument, countText } from '../../src/papyrus/editor/wordCount';

let editors: DocxEditor[] = [];
let hosts: HTMLElement[] = [];

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});
afterEach(() => {
  for (const ed of editors) ed.destroy();
  for (const host of hosts) host.remove();
  editors = [];
  hosts = [];
});

async function openBody(xmlBody: string) {
  const host = document.createElement('div');
  document.body.append(host);
  hosts.push(host);
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xmlBody}</w:body></w:document>`);
  const ed = new DocxEditor(host);
  editors.push(ed);
  await ed.open(await zip.generateAsync({ type: 'uint8array' }));
  return ed;
}
const p = (s: string) => `<w:p><w:r><w:t xml:space="preserve">${s}</w:t></w:r></w:p>`;
const cell = (s: string) => `<w:tc>${p(s)}</w:tc>`;

// These are distinct count-inspection jobs: selected amount, cross-paragraph selection,
// table body, tab-delimited paste, hard line break, NBSP, copied repeated whitespace,
// mixed-script selection, and count refresh after an edit.
describe('QA20K document insights workflows 01051–01060', () => {
  it('QA20K-01051 reports all three statistics for a selected invoice amount', async () => {
    const ed = await openBody(p('Invoice PO-2048 $1250 due Friday'));
    const text = 'PO-2048 $1250';
    const start = ed.view!.state.doc.textContent.indexOf(text) + 1;
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, start, start + text.length)));
    expect(ed.selectionWordCount()).toEqual({ words: 2, chars: 12, charsWithSpaces: 13 });
    expect(ed.wordCount().words).toBe(5);
  });

  it('QA20K-01052 counts a selected phrase spanning two report paragraphs', async () => {
    const ed = await openBody(p('Quarterly') + p('forecast'));
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 1, 1 + 9 + 2 + 8)));
    expect(ed.selectionWordCount()).toEqual({ words: 2, chars: 17, charsWithSpaces: 18 });
    expect(ed.wordCount().words).toBe(2);
  });

  it('QA20K-01053 includes values from a DOCX table in whole-document statistics', async () => {
    const table = `<w:tbl><w:tr>${cell('Invoice')}${cell('2026-09')}</w:tr></w:tbl>`;
    const ed = await openBody(table);
    expect(ed.view!.state.doc.textContent).toContain('Invoice');
    expect(ed.view!.state.doc.textContent).toContain('2026-09');
    expect(ed.wordCount()).toEqual({ words: 2, chars: 14, charsWithSpaces: 15 });
  });

  it('QA20K-01054 treats tab-separated pasted departments as three words', () => {
    expect(countText('North\tSouth\tEast')).toEqual({ words: 3, chars: 14, charsWithSpaces: 16 });
  });

  it('QA20K-01055 counts a manual line break as separation between approval terms', () => {
    expect(countText('Revised\nby Finance')).toEqual({ words: 3, chars: 16, charsWithSpaces: 18 });
  });

  it('QA20K-01056 treats a nonbreaking space as a separator in an account label', () => {
    expect(countText('Budget\u00a0approved')).toEqual({ words: 2, chars: 14, charsWithSpaces: 15 });
  });

  it('QA20K-01057 distinguishes visible characters from repeated pasted whitespace', () => {
    expect(countText('A \t B')).toEqual({ words: 2, chars: 2, charsWithSpaces: 5 });
  });

  it('QA20K-01058 counts each CJK label character beside an unspaced product code', async () => {
    const ed = await openBody(p('報告Q3'));
    const start = 1;
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, start, start + 4)));
    expect(ed.selectionWordCount()).toEqual({ words: 3, chars: 4, charsWithSpaces: 4 });
  });

  it('QA20K-01059 refreshes body statistics after an inserted budget status', async () => {
    const ed = await openBody(p('Budget approved'));
    expect(ed.wordCount()).toEqual({ words: 2, chars: 14, charsWithSpaces: 15 });
    ed.view!.dispatch(ed.view!.state.tr.insertText(' urgently', 15));
    expect(ed.wordCount()).toEqual({ words: 3, chars: 22, charsWithSpaces: 24 });
  });

  it('QA20K-01060 updates selection statistics when the user narrows the selected contract phrase', async () => {
    const ed = await openBody(p('Please review clause 8 today'));
    const start = ed.view!.state.doc.textContent.indexOf('clause 8') + 1;
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, start, start + 8)));
    expect(ed.selectionWordCount()).toEqual({ words: 2, chars: 7, charsWithSpaces: 8 });
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, start, start + 6)));
    expect(ed.selectionWordCount()).toEqual({ words: 1, chars: 6, charsWithSpaces: 6 });
  });
});

