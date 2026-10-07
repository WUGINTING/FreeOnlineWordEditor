// Typing and layout speed-ups (2026-09-25): every cache and fast path gives the results the full
// computation gives, and pagination keeps its guards.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { EditorState, type Plugin } from 'prosemirror-state';
import { EditorView, type DecorationSet } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { DOMParser as PMDOMParser, DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { countDocument, countText } from '../../src/papyrus/editor/wordCount';
import { documentFields, fieldPlaceholders, fieldScans, lockedFields } from '../../src/papyrus/editor/fields';
import { listMarkers, listMarkersKey } from '../../src/papyrus/editor/listMarkers';
import { tableStyleClasses } from '../../src/papyrus/editor/tableStyles';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';
import { isTyping } from '../../src/papyrus/editor/steps';
import { MAX_PAGES, computeBreaks, pagesKey, pagination, repaginate, type Layout, type PageBox } from '../../src/papyrus/editor/pagination';
import { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
async function docOf(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  return readDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${text ? p(text) : '<w:p/>'}</w:tc>`;
const TABLE = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr>${cell('甲乙 alpha')}${cell('')}</w:tr><w:tr>${cell('beta,gamma')}${cell('丙。')}</w:tr></w:tbl>`;
const BODY = p('合約 contract 草稿') + '<w:p/>' + TABLE + p('第二段 two  words') +
  '<w:p><w:r><w:t>before</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>after</w:t></w:r></w:p>' +
  p('line &lt; x') + '<w:p><w:r><w:t>soft</w:t><w:br/><w:t>break</w:t></w:r></w:p>' + p('結尾');

/** The count as Word's rules give it for the whole text at once (the reference). */
const wholeCount = (doc: PMNode) => countText(doc.textBetween(0, doc.content.size, '\n', '￼'));

describe('word count cached per block', () => {
  it('gives the whole-text count for paragraphs, tables, page breaks, pictures and empty blocks', async () => {
    const { doc } = await docOf(BODY);
    expect(countDocument(doc)).toEqual(wholeCount(doc));
    // Block-level objects and a first block without text.
    const img = schema.nodes.image.create({ src: 'data:image/png;base64,AAAA', width: 10, height: 10 });
    const built = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(),
      schema.nodes.page_break.create(),
      schema.nodes.paragraph.create(null, [schema.text('x '), img, schema.text(' y')]),
      ...doc.content.content,
      schema.nodes.page_break.create(),
    ]);
    expect(countDocument(built)).toEqual(wholeCount(built));
    const one = schema.nodes.doc.create(null, schema.nodes.page_break.create());
    expect(countDocument(one)).toEqual(wholeCount(one));
  });

  it('stays equal through typing, splitting, joining and deleting (unchanged blocks reused)', async () => {
    const { doc } = await docOf(BODY);
    let state = EditorState.create({ schema, doc });
    const edits: ((s: EditorState) => EditorState)[] = [
      (s) => s.apply(s.tr.insertText('新增 words', 3)),
      (s) => s.apply(s.tr.split(5)),
      (s) => s.apply(s.tr.delete(4, 9)), // across the split: joins
      (s) => s.apply(s.tr.insertText(' ', 2)),
      (s) => {
        let tablePos = -1;
        s.doc.forEach((n, pos) => { if (n.type.name === 'table' && tablePos < 0) tablePos = pos; });
        return s.apply(s.tr.delete(tablePos, tablePos + s.doc.nodeAt(tablePos)!.nodeSize));
      },
      (s) => s.apply(s.tr.delete(0, s.doc.content.size - 2)),
    ];
    for (const edit of edits) {
      state = edit(state);
      expect(countDocument(state.doc)).toEqual(wholeCount(state.doc));
    }
  });

  const samples = process.env.DOCX_SAMPLES;
  it.skipIf(!samples)('matches the whole-text count on every sample document', async () => {
    const files: string[] = [];
    const walk = (dir: string) => {
      for (const name of readdirSync(dir)) {
        const path = join(dir, name);
        if (name === 'node_modules' || name.startsWith('.')) continue;
        if (statSync(path).isDirectory()) walk(path);
        else if (name.endsWith('.docx') && !name.startsWith('~$')) files.push(path);
      }
    };
    walk(samples!);
    for (const file of files) {
      const { doc } = await readDocx(readFileSync(file));
      expect(countDocument(doc), file).toEqual(wholeCount(doc));
    }
  }, 120_000);
});

describe('fields: no scan while typing in a document without fields', () => {
  const FIELD = `<w:sdt><w:sdtPr><w:alias w:val="客戶名稱"/><w:tag w:val="customer"/><w:id w:val="1"/><w:text/></w:sdtPr>` +
    `<w:sdtContent><w:r><w:t>台灣公司</w:t></w:r></w:sdtContent></w:sdt>`;

  it('skips the scan for typing, and finds a field brought in by a paste', async () => {
    const { doc } = await docOf(p('一般文字') + p('第二段'));
    const plugins: Plugin[] = [fieldPlaceholders(), lockedFields()];
    let state = EditorState.create({ schema, doc, plugins });
    expect(documentFields(state.doc)).toEqual([]);
    const scans = fieldScans.count;
    for (const ch of 'abc') state = state.apply(state.tr.insertText(ch, 2));
    state = state.apply(state.tr.split(3));
    expect(documentFields(state.doc)).toEqual([]);
    expect(fieldScans.count).toBe(scans); // no scan for the edits

    // A paragraph with a field, pasted: scanned, and found.
    const withField = (await docOf(`<w:p>${FIELD}</w:p>`)).doc;
    state = state.apply(state.tr.replaceWith(state.doc.content.size, state.doc.content.size, withField.content));
    const fields = documentFields(state.doc);
    expect(fields.map((f) => f.title)).toEqual(['客戶名稱']);
    expect(fieldScans.count).toBe(scans + 1);
    // The fast path's answer is the full scan's.
    expect(fields).toEqual(documentFields(state.doc.type.create(state.doc.attrs, state.doc.content)));
  });

  it('documents with fields are still scanned, with the same results', async () => {
    const { doc } = await docOf(`<w:p>${FIELD}</w:p>` + p('其他'));
    let state = EditorState.create({ schema, doc, plugins: [fieldPlaceholders(), lockedFields()] });
    const before = documentFields(state.doc);
    state = state.apply(state.tr.insertText('X', state.doc.content.size - 1));
    const after = documentFields(state.doc);
    expect(after.map((f) => [f.title, f.text])).toEqual(before.map((f) => [f.title, f.text]));
  });
});

/** Decorations as comparable data. */
function decoData(set: DecorationSet | undefined, doc: PMNode) {
  return (set?.find(0, doc.content.size) ?? []).map((d: any) => [d.from, d.to, d.spec.key ?? '', JSON.stringify(d.type.attrs ?? null)]);
}

describe('list markers and table styles are mapped when an edit does not concern them', () => {
  const parse = (html: string) => {
    const div = document.createElement('div');
    div.innerHTML = html;
    return PMDOMParser.fromSchema(schema).parseSlice(div);
  };

  it('list markers equal a fresh build after every kind of edit', async () => {
    const { doc, model } = await readDocx(await blankPackage().generateAsync({ type: 'uint8array' }));
    const plugin = () => listMarkers(() => model.numbering);
    let state = EditorState.create({ schema, doc, plugins: [history(), plugin()] });
    const slice = adoptPastedLists(parse('<ol><li>one</li><li>two<ol type="a"><li>sub</li></ol></li><li>three</li></ol>'), model.numbering);
    state = state.apply(state.tr.replaceSelection(slice));
    // Plain paragraphs before and after the list (added without touching it).
    state = state.apply(state.tr.insert(0, schema.nodes.paragraph.create(null, schema.text('intro'))));
    state = state.apply(state.tr.insert(state.doc.content.size, schema.nodes.paragraph.create(null, schema.text('outro'))));
    const check = () => {
      const fresh = EditorState.create({ schema, doc: state.doc, plugins: [plugin()] });
      expect(decoData(listMarkersKey.getState(state), state.doc)).toEqual(decoData(listMarkersKey.getState(fresh), state.doc));
    };
    check();
    const listItem = () => {
      let at = -1;
      state.doc.forEach((n, pos) => { if (n.attrs.numId && at < 0) at = pos; });
      return at;
    };
    // Typing inside a list item and outside of lists: mapped.
    state = state.apply(state.tr.insertText('X', listItem() + 2));
    check();
    state = state.apply(state.tr.insertText('Y', 2));
    check();
    // Enter in a list item, deleting a list item, undo: rebuilt.
    state = state.apply(state.tr.split(listItem() + 3));
    check();
    const at = listItem();
    state = state.apply(state.tr.delete(at, at + state.doc.nodeAt(at)!.nodeSize));
    check();
    undo(state, (tr) => (state = state.apply(tr)));
    check();
    // Changing a list paragraph's level: rebuilt.
    const item = listItem();
    state = state.apply(state.tr.setNodeMarkup(item, undefined, { ...state.doc.nodeAt(item)!.attrs, ilvl: 1 }));
    check();
  });

  it('table style classes equal a fresh build after every kind of edit', async () => {
    const { doc } = await docOf(p('intro') + TABLE + p('outro'));
    let state = EditorState.create({ schema, doc, plugins: [tableStyleClasses()] });
    const plugin = state.plugins[0];
    const check = () => {
      const fresh = EditorState.create({ schema, doc: state.doc, plugins: [tableStyleClasses()] });
      expect(decoData(plugin.getState(state), state.doc)).toEqual(decoData(fresh.plugins[0].getState(fresh), state.doc));
    };
    state = state.apply(state.tr.insertText('typed', 2)); // outside
    check();
    let tablePos = -1;
    state.doc.forEach((n, pos) => { if (n.type.name === 'table') tablePos = pos; });
    state = state.apply(state.tr.insertText('in a cell', tablePos + 4)); // inside a cell
    check();
    const table = state.doc.nodeAt(tablePos)!;
    state = state.apply(state.tr.setNodeMarkup(tablePos, undefined, { ...table.attrs, styleId: 'Other' }));
    check();
    state = state.apply(state.tr.insert(0, table)); // a second table
    check();
    state = state.apply(state.tr.delete(0, table.nodeSize));
    check();
  });
});

describe('typing is told apart from other edits', () => {
  it('only text typed or deleted inside one paragraph counts as typing', async () => {
    const { doc } = await docOf(p('abc') + p('def'));
    const state = EditorState.create({ schema, doc, plugins: [history()] });
    expect(isTyping(state.tr.insertText('x', 2))).toBe(true);
    expect(isTyping(state.tr.delete(2, 3))).toBe(true);
    expect(isTyping(state.tr.split(2))).toBe(false); // Enter
    expect(isTyping(state.tr.delete(3, 7))).toBe(false); // joins two paragraphs
    expect(isTyping(state.tr.insertText('x', 2).setMeta('paste', true))).toBe(false);
    expect(isTyping(state.tr.insert(2, schema.nodes.image.create({ src: 'data:image/png;base64,AAAA' })))).toBe(false);
    expect(isTyping(state.tr.addMark(1, 3, schema.marks.bold.create()))).toBe(false); // formatting
    expect(isTyping(state.tr)).toBe(false);
    const typed = state.apply(state.tr.insertText('x', 2));
    let undone: any = null;
    undo(typed, (tr) => (undone = tr));
    expect(isTyping(undone)).toBe(false);
  });
});

// ----- pagination -----

const geometry = { width: 794, height: 1123, textTop: 96, textBottom: 96, marginLeft: 96, marginRight: 96 };
function layoutFor(doc: PMNode, g: Partial<typeof geometry> = {}, gap = 24): Layout {
  return { gap, sections: [{ geometry: { ...geometry, ...g }, start: 'nextPage', pageNumberStart: null, firstBlock: 0, lastBlock: doc.childCount - 1 }] };
}
function viewOf(doc: PMNode, plugins: Plugin[] = []) {
  const host = document.createElement('div');
  document.body.append(host);
  const view = new EditorView(host, { state: EditorState.create({ schema, doc, plugins }) });
  return { view, done: () => { view.destroy(); host.remove(); } };
}

describe('pagination guards', () => {
  it('a negative page height finishes with a page instead of hanging', async () => {
    const { doc } = await docOf(p('一') + p('二') + p('三'));
    const { view, done } = viewOf(doc);
    const result = computeBreaks(view, layoutFor(doc, { height: -133 }));
    expect(result.pages.length).toBeGreaterThanOrEqual(1);
    expect(result.pages.length).toBeLessThan(10);
    const zero = computeBreaks(view, layoutFor(doc, { height: 0 }, 0));
    expect(zero.pages.length).toBeLessThan(10);
    done();
  });

  it(`a block needing more than ${MAX_PAGES} pages stops at ${MAX_PAGES}`, async () => {
    const { doc } = await docOf(p('tall'));
    const { view, done } = viewOf(doc);
    const dom = view.nodeDOM(0) as HTMLElement;
    dom.getBoundingClientRect = () => ({ x: 0, y: 0, top: 0, left: 0, right: 100, bottom: 5e7, width: 100, height: 5e7, toJSON() {} }) as DOMRect;
    const t = Date.now();
    const result = computeBreaks(view, layoutFor(doc, { height: 200, textTop: 10, textBottom: 10 }, 0));
    expect(result.pages.length).toBe(MAX_PAGES);
    expect(Date.now() - t).toBeLessThan(5000);
    done();
  });

  it('a document with <w:pgSz w:h="-2000"/> opens and lays out', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${p('一')}${p('二')}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="-2000"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>');
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    await new Promise((r) => setTimeout(r, 120)); // the first measurements
    expect(ed.pages.length).toBeGreaterThanOrEqual(1);
    expect(ed.pages.length).toBeLessThan(10);
    ed.destroy();
    host.remove();
  });

  it('pages are drawn again only for changes of half a pixel or more', () => {
    const page: PageBox = { ...geometry, section: 0, inSection: 0, number: 1, top: 0, left: 0, cuts: [{ top: 500.1234, bottom: 620, left: 96, right: 698, border: 'none' }] };
    const noise: PageBox = { ...page, cuts: [{ ...page.cuts![0], top: 500.1244 }] };
    const moved: PageBox = { ...page, cuts: [{ ...page.cuts![0], top: 501.2 }] };
    expect(pagesKey([noise])).toBe(pagesKey([page]));
    expect(pagesKey([moved])).not.toBe(pagesKey([page]));
    expect(pagesKey([{ ...page, number: 2 }])).not.toBe(pagesKey([page]));
  });
});

describe('when pagination measures', () => {
  async function setup(getLayout?: (doc: PMNode) => Layout, onPages: (pages: PageBox[]) => void = () => {}) {
    const { doc } = await docOf(p('abc') + p('def'));
    vi.useFakeTimers();
    const layouts = vi.fn(getLayout ?? ((d: PMNode) => layoutFor(d)));
    const pagesSeen = vi.fn(onPages);
    const { view, done } = viewOf(doc, [history(), pagination(layouts, (pages) => pagesSeen(pages))]);
    vi.advanceTimersByTime(100);
    layouts.mockClear();
    return { view, layouts, pagesSeen, done };
  }

  it('typing waits for a pause (or a second into a burst); Enter and page setup are measured at once', async () => {
    const { view, layouts, done } = await setup();
    view.dispatch(view.state.tr.insertText('x', 2));
    vi.advanceTimersByTime(60);
    expect(layouts).toHaveBeenCalledTimes(0);
    vi.advanceTimersByTime(200);
    expect(layouts).toHaveBeenCalledTimes(1);

    // A burst of keystrokes 100 ms apart: one measurement within the first second, not ten.
    layouts.mockClear();
    for (let i = 0; i < 15; i++) {
      view.dispatch(view.state.tr.insertText('y', 2));
      vi.advanceTimersByTime(100);
    }
    vi.advanceTimersByTime(300);
    expect(layouts.mock.calls.length).toBeGreaterThanOrEqual(2);
    expect(layouts.mock.calls.length).toBeLessThanOrEqual(3);

    layouts.mockClear();
    view.dispatch(view.state.tr.split(3)); // Enter
    vi.advanceTimersByTime(40);
    expect(layouts).toHaveBeenCalledTimes(1);

    layouts.mockClear();
    view.dispatch(view.state.tr.insertText('z', 2));
    repaginate(view); // an explicit request right after typing: not held back by it
    vi.advanceTimersByTime(40);
    expect(layouts).toHaveBeenCalledTimes(1);
    done();
  });

  it('a layout that keeps alternating stops being measured', async () => {
    let flip = false;
    let view: EditorView | null = null;
    const { view: v, layouts, done } = await setup(
      (d) => layoutFor(d, { height: (flip = !flip) ? 1000 : 1100 }),
      () => view && repaginate(view), // like a header whose height follows the page count
    );
    view = v;
    view.dispatch(view.state.tr.insertText('x', 2));
    vi.advanceTimersByTime(20_000);
    expect(layouts.mock.calls.length).toBeLessThanOrEqual(60);
    // The next edit is measured again.
    layouts.mockClear();
    view.dispatch(view.state.tr.split(3));
    vi.advanceTimersByTime(40);
    expect(layouts).toHaveBeenCalled();
    done();
  });

  it('a failed measurement is reported and the next change measures again', async () => {
    const error = vi.spyOn(console, 'error').mockImplementation(() => {});
    let fail = false;
    const { view, pagesSeen, done } = await setup((d) => {
      if (fail) throw new Error('boom');
      return layoutFor(d, { height: 900 + d.content.size });
    });
    fail = true;
    view.dispatch(view.state.tr.split(3));
    expect(() => vi.advanceTimersByTime(40)).not.toThrow();
    expect(error).toHaveBeenCalled();
    fail = false;
    pagesSeen.mockClear();
    view.dispatch(view.state.tr.split(3));
    vi.advanceTimersByTime(40);
    expect(pagesSeen).toHaveBeenCalled();
    done();
  });
});

// ----- the page drawing (renderPages) -----

describe('pages are redrawn only where they changed', () => {
  async function withHeader() {
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await blankPackage().generateAsync({ type: 'uint8array' }));
    ed.editHeaderFooter('header');
    ed.run((s, d) => (d?.(s.tr.insertText('第 ')), true));
    ed.insertField('PAGE');
    ed.run((s, d) => (d?.(s.tr.insertText(' 頁')), true));
    ed.closeHeaderFooter();
    const page = ed.pages[0];
    const pages = (n: number, first = 1): PageBox[] =>
      Array.from({ length: n }, (_, i) => ({ ...page, inSection: i, number: first + i, top: i * (page.height + 24) }));
    const render = (list: PageBox[]) => (ed as any).renderPages(list);
    const done = () => { ed.destroy(); host.remove(); };
    return { ed, host, pages, render, done };
  }
  const headers = (host: HTMLElement) => Array.from(host.querySelectorAll<HTMLElement>('.dx-pages .dx-header'));

  it('keeps page elements and header boxes, and only patches the page numbers', async () => {
    const { host, pages, render, done } = await withHeader();
    render(pages(3));
    const pageEls = Array.from(host.querySelectorAll('.dx-page'));
    const boxes = headers(host);
    expect(boxes.map((b) => b.textContent)).toEqual(['第 1 頁', '第 2 頁', '第 3 頁']);
    render(pages(4, 5)); // one page more, numbering from 5
    expect(Array.from(host.querySelectorAll('.dx-page')).slice(0, 3)).toEqual(pageEls);
    expect(headers(host).slice(0, 3)).toEqual(boxes);
    expect(headers(host).map((b) => b.textContent)).toEqual(['第 5 頁', '第 6 頁', '第 7 頁', '第 8 頁']);
    render(pages(2));
    expect(host.querySelectorAll('.dx-page').length).toBe(2);
    expect(headers(host).map((b) => b.textContent)).toEqual(['第 1 頁', '第 2 頁']);
    done();
  });

  it('serializes a header once, measures it once, and a tall header still pushes the body down', async () => {
    const { ed, host, pages, render, done } = await withHeader();
    // The schema's own serializer (the one every caller gets from DOMSerializer.fromSchema).
    const serialize = vi.spyOn(DOMSerializer.fromSchema(schema), 'serializeFragment');
    let reads = 0;
    const desc = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetHeight')!;
    vi.spyOn(HTMLElement.prototype, 'offsetHeight', 'get').mockImplementation(function (this: HTMLElement) {
      if (this.classList.contains('dx-header')) {
        reads++;
        return 400;
      }
      return desc.get!.call(this);
    });
    const hf = ed.model.headerFooters.find((h) => h.kind === 'header')!;
    // A header document not drawn before: serialized once for all its pages, measured once.
    hf.doc = hf.doc.type.create(hf.doc.attrs, hf.doc.content); // a new document, same content
    render(pages(5));
    expect(serialize.mock.calls.filter((c) => c[0] === hf.doc.content).length).toBe(1);
    expect(reads).toBe(1);
    render(pages(6, 3));
    expect(reads).toBe(1);
    // Word moves the text area below a header taller than its margin.
    const headerTop = ed.sections()[0].page.header;
    const pushed = (ed as any).layout(ed.view!.state.doc).sections[0].geometry.textTop;
    expect(pushed).toBeCloseTo(headerTop / 15 + 400, 0);
    // A changed header is drawn and measured again.
    hf.doc = hf.doc.type.create(hf.doc.attrs, hf.doc.content); // a new document, same content
    render(pages(6, 3));
    expect(reads).toBe(2);
    expect(headers(host).length).toBe(6);
    done();
  });
});

describe('the cursor page is read once per frame', () => {
  it('an edit does not ask for the cursor position right away', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const updates: number[] = [];
    const ed = new DocxEditor(host, { onUpdate: (s) => updates.push(s.currentPage) });
    await ed.open(await blankPackage().generateAsync({ type: 'uint8array' }));
    await new Promise((r) => setTimeout(r, 100)); // the first layout
    const page = ed.pages[0];
    // Two pages, the cursor on the second (the next measurement comes after the typing pause).
    (ed as any).renderPages([page, { ...page, number: 2, inSection: 1, top: page.height + 24 }]);
    const coords = vi.spyOn(ed.view!, 'coordsAtPos').mockReturnValue({ top: page.height + 100, bottom: page.height + 110, left: 0, right: 0 });
    ed.run((s, d) => (d?.(s.tr.insertText('x')), true));
    expect(coords).not.toHaveBeenCalled();
    await new Promise((r) => setTimeout(r, 150));
    expect(coords).toHaveBeenCalled();
    expect(updates[updates.length - 1]).toBe(2); // told once the page is known
    expect(ed.snapshot()!.currentPage).toBe(2);
    ed.destroy();
    host.remove();
  });
});
