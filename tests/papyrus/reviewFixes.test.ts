// Regression tests for the tracked-changes review fixes. Expected results were checked against
// Microsoft Word (tracking off; Accept All / Reject All / Range.Revisions.RejectAll).
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { parseXml } from '../../src/papyrus/docx/xml';
import { documentSections, withPageSetup } from '../../src/papyrus/docx/sections';
import { schema } from '../../src/papyrus/editor/schema';
import { splitParagraph } from '../../src/papyrus/editor/commands';
import { describeRevision } from '../../src/papyrus/docx/revisions';
import {
  acceptAllRevisions, acceptRevision, collectRevisions, rejectAllRevisions, rejectRevision, review, revisionsAtSelection,
} from '../../src/papyrus/editor/review';
import { compareElements } from './canonical';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const A = (id: number, who = 'Alice') => `w:id="${id}" w:author="${who}" w:date="2026-01-02T03:04:00Z"`;
const R = (t: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const T = (t: string) => `<w:r><w:t>${t}</w:t></w:r>`;
const P = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const BR = '<w:r><w:br w:type="page"/></w:r>';
const TBL = (cell: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  cell + '</w:tc></w:tr></w:tbl>';

async function openBody(body: string, opts: { sect?: string; files?: Record<string, string | Uint8Array> } = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${opts.sect ?? SECT}</w:body></w:document>`);
  for (const [k, v] of Object.entries(opts.files ?? {})) zip.file(k, v);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: [history(), review()] });
  const run = (cmd: Command) => cmd(state, (tr) => (state = state.apply(tr)));
  const select = (from: number, to = from) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to))));
  const type = (text: string) => (state = state.apply(state.tr.insertText(text)));
  const xml = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  const paras = async () => {
    const b = /<w:body>([\s\S]*)<w:sectPr/.exec(await xml())![1];
    return b.split(/(?=<w:p[ >]|<w:tbl>)/);
  };
  const body0 = body;
  const unchanged = async () => {
    const orig = parseXml(`<w:document ${W}><w:body>${body0}${opts.sect ?? SECT}</w:body></w:document>`).documentElement;
    return compareElements(orig, parseXml(await xml()).documentElement).join('\n');
  };
  return { model, run, select, type, xml, paras, unchanged, get state() { return state; } };
}

function posOf(state: EditorState, text: string): number {
  let found = -1;
  state.doc.descendants((n, pos) => {
    if (found < 0 && n.isText && n.text!.includes(text)) found = pos + n.text!.indexOf(text);
    return found < 0;
  });
  if (found < 0) throw new Error('not found: ' + text);
  return found;
}

describe('1. resolving a selection maps positions safely', () => {
  it('rejecting an insertion that holds a deletion leaves the next deletion alone (Word: Range.Revisions.RejectAll)', async () => {
    const d = await openBody(P(T('s') + `<w:ins ${A(1)}>${T('abc')}<w:del ${A(2, 'Bob')}><w:r><w:delText>X</w:delText></w:r></w:del></w:ins><w:del ${A(3, 'Carol')}><w:r><w:delText>Y</w:delText></w:r></w:del>` + T('e')));
    const a = posOf(d.state, 'abc');
    d.select(a, a + 4); // "abc" and Bob's deletion, not Carol's
    expect(revisionsAtSelection(d.state).map((r) => r.author)).toEqual(['Alice', 'Bob']);
    d.run(rejectRevision);
    expect(d.state.doc.textContent).toBe('se');
    const xml = await d.xml();
    expect(xml).toContain(`<w:del ${A(3, 'Carol')}><w:r><w:delText>Y</w:delText></w:r></w:del>`);
    expect(collectRevisions(d.state.doc).map((r) => r.author)).toEqual(['Carol']);
  });
});

describe('2. Enter in a paragraph with paragraph revisions (as Word, tracking off)', () => {
  const body = P(R('Left Right'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr><w:pPrChange ${A(2)}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`) + P(T('Next'));

  it('the first half gets a new, untracked mark; the original mark stays with the last half', async () => {
    const d = await openBody(body);
    d.select(posOf(d.state, 'Left') + 4);
    d.run(splitParagraph);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Left</w:t></w:r></w:p>');
    expect(p[1]).toBe(`<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve"> Right</w:t></w:r></w:p>`);
    expect(collectRevisions(d.state.doc).map((r) => r.kind)).toEqual(['paraMark']);
    // Word's Accept All then keeps the user's paragraph break.
    d.run(acceptAllRevisions);
    expect(await d.paras()).toEqual([
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Left</w:t></w:r></w:p>',
      '<w:p><w:r><w:t xml:space="preserve"> RightNext</w:t></w:r></w:p>',
    ]);
  });

  it('Enter at the end: the new empty paragraph carries the original mark', async () => {
    const d = await openBody(body);
    d.select(posOf(d.state, 'Left') + 10);
    d.run(splitParagraph);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Left Right</w:t></w:r></w:p>');
    expect(p[1]).toBe(`<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr></w:p>`);
  });

  it('in a paragraph split at a page break, every piece of the first half drops the mark', async () => {
    const d = await openBody(P(T('X') + BR + T('YZ'), `<w:rPr><w:del ${A(1)}/></w:rPr>`) + P(T('Next')));
    d.select(posOf(d.state, 'YZ') + 1);
    d.run(splitParagraph);
    expect(await d.paras()).toEqual([
      `<w:p><w:r><w:t>X</w:t></w:r>${BR}<w:r><w:t>Y</w:t></w:r></w:p>`,
      `<w:p><w:pPr><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr><w:r><w:t>Z</w:t></w:r></w:p>`,
      '<w:p><w:r><w:t>Next</w:t></w:r></w:p>',
    ]);
    // Enter before the page break: the rest of the w:p (with its mark) stays one paragraph.
    const e = await openBody(P(T('XW') + BR + T('Y'), `<w:rPr><w:del ${A(1)}/></w:rPr>`) + P(T('Next')));
    e.select(posOf(e.state, 'XW') + 1);
    e.run(splitParagraph);
    expect(await e.paras()).toEqual([
      '<w:p><w:r><w:t>X</w:t></w:r></w:p>',
      `<w:p><w:pPr><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr><w:r><w:t>W</w:t></w:r>${BR}<w:r><w:t>Y</w:t></w:r></w:p>`,
      '<w:p><w:r><w:t>Next</w:t></w:r></w:p>',
    ]);
  });

  it('a section break stays on the last half (Word keeps w:sectPr on the last paragraph mark)', async () => {
    const S1 = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const d = await openBody(P(T('One')) + P(T('End1'), S1) + P(T('Two')));
    const before = documentSections(d.state.doc, d.model);
    expect(before.map((s) => [s.firstBlock, s.lastBlock])).toEqual([[0, 1], [2, 2]]);
    d.select(posOf(d.state, 'End1') + 4);
    d.run(splitParagraph);
    d.type('typed');
    const after = documentSections(d.state.doc, d.model);
    expect(after.map((s) => [s.firstBlock, s.lastBlock])).toEqual([[0, 2], [3, 3]]);
    expect(after[0].page.width).toBe(16838);
    const p = await d.paras();
    expect(p[1]).toBe('<w:p><w:r><w:t>End1</w:t></w:r></w:p>');
    expect(p[2]).toBe(`<w:p><w:pPr>${S1}</w:pPr><w:r><w:t>typed</w:t></w:r></w:p>`);
  });
});

describe('3. typing next to or inside a tracked insertion is not tracked (Word, tracking off)', () => {
  it('typed text goes between the insertion pieces, which keep distinct ids', async () => {
    const d = await openBody(P(`<w:ins ${A(7, 'Bob')}>${T('abc')}</w:ins>`));
    const a = posOf(d.state, 'abc');
    d.select(a + 1);
    d.type('X');
    d.select(a + 4);
    d.type('Y');
    const p = (await d.paras())[0];
    expect(p).toMatch(new RegExp(`^<w:p><w:ins ${A(7, 'Bob')}><w:r><w:t>a</w:t></w:r></w:ins><w:r><w:t>X</w:t></w:r><w:ins w:id="(\\d+)" w:author="Bob" w:date="2026-01-02T03:04:00Z"><w:r><w:t>bc</w:t></w:r></w:ins><w:r><w:t>Y</w:t></w:r></w:p>$`));
    expect(/<w:ins w:id="(\d+)" w:author="Bob"[^>]*><w:r><w:t>bc/.exec(p)![1]).not.toBe('7');
    expect(collectRevisions(d.state.doc).map((r) => d.state.doc.textBetween(r.from, r.to))).toEqual(['a', 'bc']);
  });

  it('typing at the start of a paragraph that begins with an insertion is not tracked either', async () => {
    const d = await openBody(P(`<w:ins ${A(7)}>${T('abc')}</w:ins>`));
    d.select(1);
    d.type('Q');
    expect((await d.paras())[0]).toBe(`<w:p><w:r><w:t>Q</w:t></w:r><w:ins ${A(7)}><w:r><w:t>abc</w:t></w:r></w:ins></w:p>`);
  });
});

describe('4. a page break inside a tracked insertion', () => {
  const body = P(T('x') + `<w:ins ${A(1)}>${T('a')}${BR}${T('b')}</w:ins>` + T('y'));

  it('opening and saving changes nothing (one w:ins, the break stays inside it)', async () => {
    const d = await openBody(body);
    expect(await d.unchanged()).toBe('');
    expect(collectRevisions(d.state.doc).map((r) => r.kind)).toEqual(['ins']);
  });

  it('accept keeps the break untracked; reject removes it and rejoins the paragraph', async () => {
    const a = await openBody(body);
    a.run(acceptAllRevisions);
    expect(await a.paras()).toEqual([`<w:p><w:r><w:t>xa</w:t></w:r>${BR}<w:r><w:t>by</w:t></w:r></w:p>`]);
    expect(collectRevisions(a.state.doc)).toEqual([]);
    const r = await openBody(body);
    r.run(rejectAllRevisions);
    expect(await r.paras()).toEqual(['<w:p><w:r><w:t>xy</w:t></w:r></w:p>']);
  });
});

describe('5. a paragraph split at page breaks is one paragraph for review', () => {
  const body =
    P(T('X') + BR + T('Y'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr><w:pPrChange ${A(2)}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`) + P(T('Z'));

  it('shows each paragraph revision once', async () => {
    const d = await openBody(body);
    expect(collectRevisions(d.state.doc).map((r) => r.kind)).toEqual(['paraFormat', 'paraMark']);
  });

  it('accepting the formatting change at the cursor keeps one w:p', async () => {
    const d = await openBody(body);
    d.select(posOf(d.state, 'Y'));
    expect(revisionsAtSelection(d.state).map((r) => r.kind)).toEqual(['paraFormat']);
    d.run(acceptRevision);
    expect((await d.paras())[0]).toBe(`<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr><w:r><w:t>X</w:t></w:r>${BR}<w:r><w:t>Y</w:t></w:r></w:p>`);
  });

  it('Accept All / Reject All leave no records', async () => {
    const a = await openBody(body);
    a.run(acceptAllRevisions);
    expect(await a.paras()).toEqual([`<w:p><w:r><w:t>X</w:t></w:r>${BR}<w:r><w:t>YZ</w:t></w:r></w:p>`]);
    expect(collectRevisions(a.state.doc)).toEqual([]);
    const r = await openBody(body);
    r.run(rejectAllRevisions);
    expect(await r.paras()).toEqual([
      `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>X</w:t></w:r>${BR}<w:r><w:t>Y</w:t></w:r></w:p>`,
      '<w:p><w:r><w:t>Z</w:t></w:r></w:p>',
    ]);
    expect(collectRevisions(r.state.doc)).toEqual([]);
  });
});

describe('6. paragraph marks that cannot join the next paragraph', () => {
  const tbl = TBL(P(T('cell'), '<w:jc w:val="right"/>'));

  it('a whole deleted paragraph before a table disappears on accept', async () => {
    const d = await openBody(P(T('A')) + P(`<w:del ${A(2)}><w:r><w:delText>B</w:delText></w:r></w:del>`, `<w:rPr><w:del ${A(1)}/></w:rPr>`) + tbl + P(T('after')));
    d.run(acceptAllRevisions);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
    expect(p[1].startsWith('<w:tbl>')).toBe(true);
    expect(collectRevisions(d.state.doc)).toEqual([]);
  });

  it('text before a table joins the first cell paragraph, which keeps its properties (Word)', async () => {
    const d = await openBody(P(T('A')) + P(T('B'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr>`) + tbl + P(T('after')));
    d.run(acceptAllRevisions);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:r><w:t>A</w:t></w:r></w:p>');
    expect(p[1].startsWith('<w:tbl>')).toBe(true);
    expect(await d.xml()).toContain('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Bcell</w:t></w:r></w:p>');
    expect(collectRevisions(d.state.doc)).toEqual([]);
    // Rejecting an inserted mark there does the same.
    const r = await openBody(P(T('A')) + P(T('B'), `<w:rPr><w:ins ${A(1)}/></w:rPr>`) + tbl + P(T('after')));
    r.run(rejectAllRevisions);
    expect(await r.xml()).toContain('<w:r><w:t>Bcell</w:t></w:r>');
  });

  it('the last paragraph: the record goes; an empty one goes with it', async () => {
    const t = await openBody(P(T('A')) + P(T('B'), `<w:rPr><w:del ${A(1)}/></w:rPr>`));
    t.run(acceptAllRevisions);
    expect(await t.paras()).toEqual(['<w:p><w:r><w:t>A</w:t></w:r></w:p>', '<w:p><w:r><w:t>B</w:t></w:r></w:p>']);
    expect(collectRevisions(t.state.doc)).toEqual([]);
    const e = await openBody(P(T('A')) + P('', `<w:rPr><w:del ${A(1)}/></w:rPr>`));
    e.run(acceptAllRevisions);
    expect(await e.paras()).toEqual(['<w:p><w:r><w:t>A</w:t></w:r></w:p>']);
    const i = await openBody(P(T('A')) + P('', `<w:rPr><w:ins ${A(1)}/></w:rPr>`));
    i.run(rejectAllRevisions);
    expect(await i.paras()).toEqual(['<w:p><w:r><w:t>A</w:t></w:r></w:p>']);
    const k = await openBody(P(T('A')) + P('', `<w:rPr><w:ins ${A(1)}/></w:rPr>`));
    k.run(acceptAllRevisions);
    expect(await k.paras()).toEqual(['<w:p><w:r><w:t>A</w:t></w:r></w:p>', '<w:p></w:p>']);
  });
});

describe('7. revisions nested in a deletion', () => {
  it('Reject All also rejects a formatting change inside a deletion (Word)', async () => {
    const d = await openBody(P(T('a') + `<w:del ${A(1)}><w:r><w:rPr><w:b/><w:rPrChange ${A(2, 'Bob')}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:delText>gone</w:delText></w:r></w:del>` + T('z')));
    d.run(rejectAllRevisions);
    expect(await d.paras()).toEqual(['<w:p><w:r><w:t>a</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>gone</w:t></w:r><w:r><w:t>z</w:t></w:r></w:p>']);
    expect(collectRevisions(d.state.doc)).toEqual([]);
    d.run(undo);
    expect(collectRevisions(d.state.doc).map((r) => r.kind)).toEqual(['del']);
  });
});

describe('8. accepting a deletion keeps range markers whose other end is outside it (Word)', () => {
  it('bookmarks', async () => {
    const s = await openBody(P(T('a') + `<w:del ${A(1)}><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:delText>gone</w:delText></w:r></w:del>` + T('kept') + '<w:bookmarkEnd w:id="0"/>' + T('z')));
    s.run(acceptAllRevisions);
    expect(await s.paras()).toEqual(['<w:p><w:r><w:t>a</w:t></w:r><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:t>kept</w:t></w:r><w:bookmarkEnd w:id="0"/><w:r><w:t>z</w:t></w:r></w:p>']);
    const w = await openBody(P(T('a') + `<w:del ${A(1)}><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:delText>gone</w:delText></w:r><w:bookmarkEnd w:id="0"/></w:del>` + T('z')));
    w.run(acceptAllRevisions);
    expect(await w.paras()).toEqual(['<w:p><w:r><w:t>az</w:t></w:r></w:p>']);
  });

  it('comment ranges', async () => {
    const s = await openBody(P(T('a') + `<w:del ${A(1)}><w:commentRangeStart w:id="7"/><w:r><w:delText>gone</w:delText></w:r></w:del>` + T('kept') + '<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r>' + T('z')));
    s.run(acceptAllRevisions);
    expect((await s.paras())[0]).toBe('<w:p><w:r><w:t>a</w:t></w:r><w:commentRangeStart w:id="7"/><w:r><w:t>kept</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r><w:r><w:t>z</w:t></w:r></w:p>');
    const w = await openBody(P(T('a') + `<w:del ${A(1)}><w:commentRangeStart w:id="7"/><w:r><w:delText>gone</w:delText></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:del>` + T('z')));
    w.run(acceptAllRevisions);
    expect(await w.paras()).toEqual(['<w:p><w:r><w:t>az</w:t></w:r></w:p>']);
  });
});

describe('restoring a deleted picture', () => {
  it('rejecting the deletion gives back an editable picture with its link', async () => {
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const drawing =
      '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
    const zip = blankPackage();
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const d = await openBody(P(T('a') + `<w:del ${A(1)}><w:r>${drawing}</w:r></w:del>` + T('z')), {
      files: {
        'word/_rels/document.xml.rels': rels.replace('</Relationships>', '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/></Relationships>'),
        'word/media/p.png': new Uint8Array(Buffer.from(png, 'base64')),
      },
    });
    d.run(rejectAllRevisions);
    const kinds: string[] = [];
    d.state.doc.descendants((n) => {
      if (n.isInline) kinds.push(n.type.name);
      return true;
    });
    expect(kinds).toEqual(['text', 'image', 'text']);
  });
});

describe('a section break on a paragraph that holds page breaks', () => {
  const PORTRAIT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const LANDSCAPE = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

  it('a w:p holding only a page break still ends its section', async () => {
    const d = await openBody(P(T('P1')) + P(BR, PORTRAIT) + P(T('L')), { sect: LANDSCAPE });
    const s = documentSections(d.state.doc, d.model);
    expect(s.map((x) => x.page.width)).toEqual([11906, 16838]);
    expect(s[0].lastBlock).toBe(2); // P1, the break, and the empty piece with the paragraph mark
    expect(await d.unchanged()).toBe('');
  });

  it('text around the break is one section, and a page setup change writes one w:p', async () => {
    const d = await openBody(P(T('X') + BR + T('Y'), LANDSCAPE) + P(T('Z')), { sect: PORTRAIT });
    const s = documentSections(d.state.doc, d.model);
    expect(s.map((x) => [x.firstBlock, x.lastBlock, x.page.width])).toEqual([[0, 2, 16838], [3, 3, 11906]]);
    expect(await d.unchanged()).toBe('');
    // Change the first section's margins, as the page setup dialog does.
    const node = d.state.doc.nodeAt(s[0].pos!)!;
    const sectPr = withPageSetup(node.attrs.sectPr, { ...s[0].page, marginLeft: 1000 });
    d.run((state, dispatch) => {
      dispatch?.(state.tr.setNodeMarkup(s[0].pos!, undefined, { ...node.attrs, sectPr }));
      return true;
    });
    const p = await d.paras();
    expect(p.length).toBe(2);
    expect(p[0]).toMatch(/^<w:p><w:pPr><w:sectPr>.*w:left="1000".*<\/w:sectPr><\/w:pPr><w:r><w:t>X<\/w:t><\/w:r><w:r><w:br w:type="page"\/><\/w:r><w:r><w:t>Y<\/w:t><\/w:r><\/w:p>$/);
  });
});

describe('10. revision times as Word shows them', () => {
  it('w:date digits as written (Word writes local time with a Z); w16du:dateUtc in local time', async () => {
    const U = (id: number) => `w:id="${id}" w:author="Ann" w:date="2026-09-24T12:01:00Z" w16du:dateUtc="2026-09-24T04:01:00Z"`;
    const d = await openBody(
      P(`<w:ins ${U(1)}>${T('a')}</w:ins><w:del ${U(2)}><w:r><w:delText>b</w:delText></w:r></w:del><w:ins ${A(3)}>${T('c')}</w:ins>`, `<w:rPr><w:del ${U(4)}/></w:rPr>`) + P(T('z')),
    );
    const utc = new Date('2026-09-24T04:01:00Z');
    const p = (n: number) => String(n).padStart(2, '0');
    const local = `${utc.getFullYear()}/${p(utc.getMonth() + 1)}/${p(utc.getDate())} ${p(utc.getHours())}:${p(utc.getMinutes())}`;
    const revs = collectRevisions(d.state.doc);
    expect(revs.map((r) => describeRevision('X', r))).toEqual([`X：Ann，${local}`, `X：Ann，${local}`, 'X：Alice，2026/01/02 03:04', `X：Ann，${local}`]);
  });
});
