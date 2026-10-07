// @vitest-environment jsdom
// createDocxEditor: the whole editor in a page element, for pages that are not Vue applications.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { createDocxEditor, DocxEditor, type CreateDocxEditorOptions, type DocxEditorHandle } from '../../src/papyrus';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function docx(text: string): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const text = (h: DocxEditorHandle) => h.editor!.view!.state.doc.textContent;
const until = async (done: () => boolean) => {
  for (let i = 0; i < 200 && !done(); i++) await new Promise((r) => setTimeout(r, 10));
  if (!done()) throw new Error('timed out');
};

let host: HTMLElement;
let handle: DocxEditorHandle | null = null;

/** The editor in a new element of the page, once its first document is open. */
async function create(options: CreateDocxEditorOptions = {}, target?: string): Promise<DocxEditorHandle> {
  host = document.createElement('div');
  host.id = 'host';
  document.body.append(host);
  let ready = false;
  handle = createDocxEditor(target ?? host, { ...options, onReady: (e) => { ready = true; options.onReady?.(e); } });
  await until(() => ready);
  return handle;
}

afterEach(() => {
  handle?.destroy();
  handle = null;
  host?.remove();
});

describe('createDocxEditor', () => {
  it('puts the editor with its ribbon in the element and opens the document given', async () => {
    let got: DocxEditor | null = null;
    const h = await create({ src: await docx('第一份文件'), onReady: (e) => (got = e) });
    expect(got).toBeInstanceOf(DocxEditor);
    expect(h.editor).toBe(got);
    expect(text(h)).toBe('第一份文件');
    expect(host.querySelector('.dx-vue')).not.toBeNull();
    expect(host.querySelectorAll('.dx-vue button').length).toBeGreaterThan(10);
  });

  it('takes a CSS selector, and says so when nothing matches', async () => {
    const h = await create({}, '#host');
    expect(host.querySelector('.dx-vue')).not.toBeNull();
    expect(h.editor).not.toBeNull();
    expect(() => createDocxEditor('#no-such-element')).toThrow(/#no-such-element/);
  });

  it('save() gives the document as a .docx', async () => {
    const h = await create({ src: await docx('要存的字') });
    const blob = await h.save();
    expect(blob).toBeInstanceOf(Blob);
    // jsdom's Blob has no arrayBuffer().
    const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(blob);
    });
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file('word/document.xml')!.async('string')).toContain('要存的字');
  });

  it('open() shows another document, the same one again, or an empty one', async () => {
    const first = await docx('第一份');
    const h = await create({ src: first });
    h.open(await docx('第二份'));
    await until(() => text(h) === '第二份');
    h.open(first);
    await until(() => text(h) === '第一份');
    h.open(null);
    await until(() => text(h) === '');
  });

  it('tells the page about changes and about a file that cannot be opened', async () => {
    let changes = 0;
    const errors: unknown[] = [];
    const h = await create({ onChange: () => changes++, onError: (e) => errors.push(e) });
    const view = h.editor!.view!;
    view.dispatch(view.state.tr.insertText('新的字', 1));
    expect(changes).toBe(1);
    h.open(new Uint8Array([1, 2, 3]));
    await until(() => errors.length > 0);
    expect(errors).toHaveLength(1);
  });

  it('without the ribbon, and read-only', async () => {
    const h = await create({ toolbar: false, editable: false });
    expect(host.querySelector('.dx-vue')).not.toBeNull();
    expect(host.querySelector('[role="toolbar"], .dx-ribbon, .dx-toolbar')).toBeNull();
    expect(h.editor!.editable).toBe(false);
  });

  it('destroy() takes the editor out of the page; using it afterwards is an error', async () => {
    const h = await create();
    h.destroy();
    expect(host.innerHTML).toBe('');
    expect(h.editor).toBeNull();
    expect(() => h.save()).toThrow(/destroyed/);
    expect(() => h.open(null)).toThrow(/destroyed/);
  });
});
