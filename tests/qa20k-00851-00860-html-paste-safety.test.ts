import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
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
const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host); await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return { editor, host, done: () => { editor.destroy(); host.remove(); } };
}

function findText(editor: DocxEditor, text: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); });
  if (pos < 0) throw new Error(`Text not found: ${text}`);
  return pos;
}

function select(editor: DocxEditor, text: string, offset = 0, length = text.length) {
  const pos = findText(editor, text) + offset;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos, pos + length)));
}

function setCaret(editor: DocxEditor, pos: number) {
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
}

function paste(editor: DocxEditor, clipboard: { html?: string; text?: string }) {
  const payload = { 'text/html': clipboard.html ?? '', 'text/plain': clipboard.text ?? '', files: [] };
  const event = new Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { files: payload.files, getData: (type: string) => payload[type as keyof typeof payload] ?? '' } });
  editor.view!.dom.dispatchEvent(event);
  return event;
}

function paragraphs(editor: DocxEditor) {
  const out: any[] = [];
  editor.view!.state.doc.forEach(node => out.push(node));
  return out;
}

function marksFor(editor: DocxEditor, exactText: string) {
  const found: string[][] = [];
  editor.view!.state.doc.descendants(node => { if (node.isText && node.text === exactText) found.push(node.marks.map(mark => mark.type.name).sort()); });
  return found;
}

describe('QA20K office paste workflows 00851–00860', () => {
  it('QA20K-00851 pastes a bold approval phrase at the cursor without losing adjacent wording', async () => {
    const { editor, done } = await open(p('The clause needs review today.'));
    setCaret(editor, findText(editor, 'review'));
    const event = paste(editor, { html: '<strong>approved</strong>&nbsp;', text: 'approved ' });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('The clause needs approved\u00a0review today.');
    expect(marksFor(editor, 'approved')).toEqual([['bold']]);
    done();
  });

  it('QA20K-00852 converts a pasted plain-text email with two lines into ordered paragraphs', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { text: 'Please confirm the date.\r\nThe vendor will attend.' });
    expect(event.defaultPrevented).toBe(true);
    expect(paragraphs(editor).map(node => node.textContent)).toEqual(['Please confirm the date.', 'The vendor will attend.']);
    done();
  });

  it('QA20K-00853 keeps tab-separated spreadsheet columns when pasting plain text', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { text: '2026-09\tTravel\t120\n2026-10\tMeals\t85' });
    expect(event.defaultPrevented).toBe(true);
    expect(paragraphs(editor).map(node => node.textContent)).toEqual(['2026-09\tTravel\t120', '2026-10\tMeals\t85']);
    expect(editor.view!.state.doc.textContent).toContain('\tTravel\t120');
    done();
  });

  it('QA20K-00854 replaces only the selected stale wording with a rich-text phrase', async () => {
    const { editor, done } = await open(p('Status: pending approval; owner: Finance.'));
    select(editor, 'pending approval');
    const event = paste(editor, { html: '<em>approved today</em>', text: 'approved today' });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Status: approved today; owner: Finance.');
    expect(marksFor(editor, 'approved today')).toEqual([['italic']]);
    done();
  });

  it('QA20K-00855 preserves HTML paragraph boundaries and order when pasting an email summary', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { html: '<p>Decision: proceed.</p><p>Owner: Finance.</p>', text: 'Decision: proceed.\nOwner: Finance.' });
    expect(event.defaultPrevented).toBe(true);
    expect(paragraphs(editor).map(node => node.textContent)).toEqual(['Decision: proceed.', 'Owner: Finance.']);
    done();
  });

  it('QA20K-00856 keeps underline and strikethrough semantics from a pasted change summary', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { html: '<p><u>new deadline</u>; <s>old deadline</s></p>' });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('new deadline; old deadline');
    expect(marksFor(editor, 'new deadline')).toEqual([['underline']]);
    expect(marksFor(editor, 'old deadline')).toEqual([['strike']]);
    done();
  });

  it('QA20K-00857 retains superscript and subscript notation in a pasted technical note', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { html: '<p>Area m<sup>2</sup>; formula H<sub>2</sub>O.</p>' });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Area m2; formula H2O.');
    expect(marksFor(editor, '2')).toEqual([['superscript'], ['subscript']]);
    done();
  });

  it('QA20K-00858 preserves nonbreaking spaces and mixed-script supplier text', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const source = '宏達科技\u00a0有限公司 — ACME';
    const event = paste(editor, { html: `<p>${source}</p>` });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe(source);
    expect(editor.view!.state.doc.textContent.charCodeAt(4)).toBe(0x00a0);
    done();
  });

  it('QA20K-00859 prefers formatted HTML over a plain-text clipboard fallback', async () => {
    const { editor, done } = await open(p(''));
    setCaret(editor, 1);
    const event = paste(editor, { html: '<strong>Quarterly total</strong>', text: 'UNFORMATTED FALLBACK' });
    expect(event.defaultPrevented).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Quarterly total');
    expect(marksFor(editor, 'Quarterly total')).toEqual([['bold']]);
    expect(editor.view!.state.doc.textContent).not.toContain('UNFORMATTED FALLBACK');
    done();
  });

  it('QA20K-00860 treats an empty clipboard paste as a no-op', async () => {
    const { editor, done } = await open(p('Do not change this sentence.'));
    select(editor, 'sentence');
    const before = editor.view!.state.doc;
    const modifiedBefore = editor.isModified();
    const event = paste(editor, {});
    expect(editor.view!.state.doc.eq(before)).toBe(true);
    expect(editor.view!.state.doc.textContent).toBe('Do not change this sentence.');
    expect(editor.isModified()).toBe(modifiedBefore);
    expect(event.defaultPrevented).toBe(false);
    done();
  });
});
