// Government-user evaluation, second batch (2026-09-25): GOV-FINDING-009 (a UTF-8 byte order
// mark in package XML stopped the document from opening) and GOV-FINDING-003 (全部拒絕 on a
// large document). GOV-ISSUE-009 (the PDF's last table row over the next paragraph) is about
// layout and is checked in a real browser. Each case reproduces the
// reported problem first.
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { existsSync, readFileSync } from 'node:fs';
import { AllSelection, EditorState } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readComments } from '../../src/papyrus/docx/comments';
import { parseFragment, parseFragments, parseXml } from '../../src/papyrus/docx/xml';
import { schema } from '../../src/papyrus/editor/schema';
import { collectRevisions, rejectAllRevisions, rejectRevision, review } from '../../src/papyrus/editor/review';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PG = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const BOM = [0xef, 0xbb, 0xbf];
// Opt-in: the folder with the evaluation's documents (they are not part of this repository).
const EVIDENCE = process.env.DOCX_EVIDENCE ?? '';

const utf8 = (s: string) => new TextEncoder().encode(s);
const withBom = (s: string) => new Uint8Array([...BOM, ...utf8(s)]);
const startsWithBom = (b: Uint8Array) => b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf;

/**
 * A package like OfficeCLI's: a styled paragraph, a header, a numbered paragraph and a comment,
 * with the parts in `encode` stored with the bytes it gives (a byte order mark, UTF-16 ...).
 */
async function pkg(encode: Record<string, (xml: string) => Uint8Array> = {}): Promise<Uint8Array> {
  const zip = blankPackage();
  const docRels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    docRels.replace(
      '</Relationships>',
      `<Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdN1" Type="${REL}/numbering" Target="numbering.xml"/>` +
        `<Relationship Id="rIdC1" Type="${REL}/comments" Target="comments.xml"/></Relationships>`,
    ),
  );
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    ct.replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
    ),
  );
  zip.file('word/header1.xml', `${HEAD}<w:hdr ${W}><w:p><w:r><w:t>頁首 BOM</w:t></w:r></w:p></w:hdr>`);
  zip.file(
    'word/numbering.xml',
    `${HEAD}<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  );
  zip.file(
    'word/comments.xml',
    `${HEAD}<w:comments ${W}><w:comment w:id="0" w:author="審核者" w:date="2026-09-25T09:00:00Z"><w:p><w:r><w:t>請確認期限</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  zip.file(
    'word/document.xml',
    `${HEAD}<w:document ${W}><w:body>` +
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>採購需求表</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:commentRangeStart w:id="0"/><w:r><w:t>期限九個月</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' +
      `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>${PG}</w:sectPr></w:body></w:document>`,
  );
  for (const [part, enc] of Object.entries(encode)) {
    const text = await zip.file(part)!.async('string');
    zip.file(part, enc(text));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

/** Every part of a package, as bytes. */
async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const [name, f] of Object.entries(zip.files)) if (!f.dir) out.set(name, await f.async('uint8array'));
  return out;
}

const BOM_PARTS = [
  '[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels', 'word/document.xml', 'word/styles.xml',
  'word/settings.xml', 'word/header1.xml', 'word/numbering.xml', 'word/comments.xml',
];

describe('GOV-FINDING-009: a UTF-8 byte order mark in the package XML', () => {
  // Edge / Chrome 153's XML parser rejects a document string starting with U+FEFF ("Unexpected
  // characters outside the root element: ï»¿", the reported error); jsdom's accepts it. The
  // parser here behaves like the browser's.
  const original = DOMParser.prototype.parseFromString;
  beforeEach(() => {
    DOMParser.prototype.parseFromString = function (this: DOMParser, text: string, type: DOMParserSupportedType) {
      if (type === 'application/xml' && text.charCodeAt(0) === 0xfeff) {
        return original.call(this, '<parsererror>Unexpected characters outside the root element: ï»¿</parsererror>', type);
      }
      return original.call(this, text, type);
    };
  });
  afterEach(() => {
    DOMParser.prototype.parseFromString = original;
  });

  it('parseXml reads XML with a byte order mark before it', () => {
    expect(parseXml(`\uFEFF${HEAD}<a/>`).documentElement.localName).toBe('a');
    expect(parseXml('\uFEFF<a><b/></a>').documentElement.firstElementChild?.localName).toBe('b');
  });

  it('opens a package whose parts all start with one, with nothing left out', async () => {
    const bytes = await pkg(Object.fromEntries(BOM_PARTS.map((p) => [p, withBom])));
    for (const [name, b] of await parts(bytes)) if (BOM_PARTS.includes(name)) expect(startsWithBom(b), name).toBe(true);
    const { doc, model } = await readDocx(bytes);
    expect(model.brokenParts ?? []).toEqual([]);
    expect(doc.textContent).toBe('採購需求表期限九個月');
    // The relationships were read: styles, numbering, header and comments.
    expect(doc.child(0).attrs.styleId).toBe('Heading1');
    expect(model.paragraphStyles.some((s) => s.id === 'Heading1')).toBe(true);
    expect(Object.keys(model.numbering.nums)).toEqual(['1']);
    expect(model.headerFooters.map((h) => h.doc.textContent)).toEqual(['頁首 BOM']);
    expect((await readComments(model.zip)).map((c) => c.author)).toEqual(['審核者']);
  });

  it('saves untouched parts byte for byte and writes rewritten ones without a byte order mark', async () => {
    const bytes = await pkg(Object.fromEntries(BOM_PARTS.map((p) => [p, withBom])));
    const before = await parts(bytes);
    const first = await readDocx(bytes);
    const saved = await writeDocx(first.doc, first.model);
    const after = await parts(saved);
    let untouched = 0;
    for (const [name, b] of before) {
      const out = after.get(name)!;
      expect(out, name).toBeTruthy();
      if (Buffer.from(out).equals(Buffer.from(b))) untouched++;
      else expect(startsWithBom(out), `${name} rewritten with a byte order mark`).toBe(false);
    }
    // What the save does not need to touch keeps its bytes (and its byte order mark).
    for (const name of ['word/styles.xml', 'word/header1.xml', 'word/comments.xml', 'word/settings.xml']) {
      expect(Buffer.from(after.get(name)!).equals(Buffer.from(before.get(name)!)), name).toBe(true);
    }
    expect(untouched).toBeGreaterThanOrEqual(4);
    const again = await readDocx(saved);
    expect(again.model.brokenParts ?? []).toEqual([]);
    expect(again.doc.textContent).toBe(first.doc.textContent);
    expect(again.model.headerFooters.map((h) => h.doc.textContent)).toEqual(['頁首 BOM']);

    // An edited header is written again, without the byte order mark.
    for (const hf of first.model.headerFooters) hf.dirty = true;
    const edited = await parts(await writeDocx(first.doc, first.model));
    expect(startsWithBom(edited.get('word/header1.xml')!)).toBe(false);
    expect(parseXml(new TextDecoder().decode(edited.get('word/header1.xml')!)).documentElement.localName).toBe('hdr');
  });

  it.skipIf(!existsSync(`${EVIDENCE}/GOVTEST_051_PENDING_REVIEW_20260925.docx`))(
    'opens the OfficeCLI documents from the evaluation with their comments and revisions',
    async () => {
      for (const name of ['GOVTEST_051_PENDING_REVIEW_20260925.docx', 'GOVTEST_006_PROCUREMENT_TABLE_20260925.docx']) {
        const bytes = readFileSync(`${EVIDENCE}/${name}`);
        const { doc, model } = await readDocx(bytes);
        expect(model.brokenParts ?? [], name).toEqual([]);
        expect(doc.textContent.length, name).toBeGreaterThan(20);
        const saved = await readDocx(await writeDocx(doc, model));
        expect(saved.doc.textContent, name).toBe(doc.textContent);
        expect(saved.model.brokenParts ?? [], name).toEqual([]);
      }
      const gov051 = await readDocx(readFileSync(`${EVIDENCE}/GOVTEST_051_PENDING_REVIEW_20260925.docx`));
      expect((await readComments(gov051.model.zip)).length).toBe(3);
      expect(collectRevisions(gov051.doc).map((r) => r.kind).sort()).toEqual(['del', 'ins']);
    },
  );

  it('opens a package with UTF-16 parts (which Word reads) and saves it readable', async () => {
    const utf16le = (s: string) => {
      const text = '\uFEFF' + s.replace('encoding="UTF-8"', 'encoding="UTF-16"');
      const out = new Uint8Array(text.length * 2);
      for (let i = 0; i < text.length; i++) {
        out[2 * i] = text.charCodeAt(i) & 0xff;
        out[2 * i + 1] = text.charCodeAt(i) >> 8;
      }
      return out;
    };
    const utf16be = (s: string) => {
      const le = utf16le(s);
      for (let i = 0; i < le.length; i += 2) [le[i], le[i + 1]] = [le[i + 1], le[i]];
      return le;
    };
    const bytes = await pkg({ 'word/document.xml': utf16le, 'word/styles.xml': utf16be, 'word/header1.xml': utf16le, '_rels/.rels': utf16le });
    const { doc, model } = await readDocx(bytes);
    expect(model.brokenParts ?? []).toEqual([]);
    expect(doc.textContent).toBe('採購需求表期限九個月');
    expect(model.headerFooters.map((h) => h.doc.textContent)).toEqual(['頁首 BOM']);
    expect(model.paragraphStyles.some((s) => s.id === 'Heading1')).toBe(true);
    const saved = await writeDocx(doc, model);
    const out = await parts(saved);
    // Saved as UTF-8 (with a declaration that says so), readable by the editor and by Word.
    for (const name of ['word/document.xml', 'word/styles.xml', 'word/header1.xml', '_rels/.rels']) {
      const text = new TextDecoder('utf-8', { fatal: true }).decode(out.get(name)!);
      expect(text, name).not.toMatch(/encoding="UTF-16"/i);
      expect(() => parseXml(text), name).not.toThrow();
    }
    const again = await readDocx(saved);
    expect(again.doc.textContent).toBe(doc.textContent);
    expect(again.model.headerFooters.map((h) => h.doc.textContent)).toEqual(['頁首 BOM']);
  });
});

describe('GOV-FINDING-003: 全部拒絕 on a large document', () => {
  const A = (id: number) => `w:id="${id}" w:author="U${id % 7}" w:date="2026-01-02T03:04:00Z"`;
  /** Deletions of every shape Reject All restores: formatting, rsidDel, fields, links, tabs, breaks. */
  const shapes = (i: number) => [
    `<w:del ${A(10 * i)}><w:r><w:delText xml:space="preserve">deleted ${i} </w:delText></w:r></w:del>`,
    `<w:del ${A(10 * i + 1)}><w:r w:rsidDel="00AB12CD" w:rsidR="00112233"><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:delText>粗體${i}</w:delText><w:tab/><w:delText>x</w:delText></w:r></w:del>`,
    `<w:del ${A(10 * i + 2)}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>1</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>`,
    `<w:del ${A(10 * i + 3)}><w:r><w:delText>same</w:delText></w:r></w:del>`, // the same XML in many paragraphs
    `<w:del ${A(10 * i + 4)}><w:r><w:rPr><w:i/></w:rPr><w:delText>it</w:delText><w:br/><w:delText>alic</w:delText></w:r></w:del>`,
  ];
  async function docWith(paras: number, variety: boolean) {
    let body = '';
    for (let i = 0; i < paras; i++) {
      const dels = variety ? shapes(i).join(`<w:r><w:t xml:space="preserve"> | </w:t></w:r>`) : shapes(i)[0];
      body +=
        `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i} keeps </w:t></w:r>` +
        `<w:ins ${A(10 * i + 5)}><w:r><w:t xml:space="preserve">inserted ${i} </w:t></w:r></w:ins>${dels}` +
        `<w:r><w:rPr><w:b/><w:rPrChange ${A(10 * i + 6)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>bold ${i}</w:t></w:r></w:p>`;
    }
    // A deletion whose run declares the namespace itself (restored on its elements).
    body += `<w:p><w:del ${A(99991)}><w:r><w:delText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">own ns</w:delText></w:r></w:del></w:p>`;
    const zip = blankPackage();
    zip.file('word/document.xml', `${HEAD}<w:document ${W}><w:body>${body}<w:sectPr>${PG}</w:sectPr></w:body></w:document>`);
    return (await readDocx(await zip.generateAsync({ type: 'uint8array' }))).doc;
  }

  it('parseFragments gives what parseFragment gives, a few dozen per document', () => {
    const xmls = [
      '<w:r><w:t>a</w:t></w:r>',
      '<w:r><w14:x/><w:t xml:space="preserve"> b </w:t></w:r>',
      '<w:p xmlns:w="urn:other"><w:r/></w:p>',
      '<w:r><w:t>unclosed</w:r>', // not well-formed: null, the others still parse
      ...Array.from({ length: 150 }, (_, i) => `<w:r w16du:dateUtc="${i}"><w:t>${i}</w:t></w:r>`),
    ];
    const got = parseFragments(xmls);
    expect(got).toHaveLength(xmls.length);
    xmls.forEach((x, i) => {
      let one: Element | null = null;
      try {
        one = parseFragment(x);
      } catch {
        one = null;
      }
      expect(got[i]?.outerHTML ?? null, x).toBe(one?.outerHTML ?? null);
      expect(got[i]?.namespaceURI ?? null, x).toBe(one?.namespaceURI ?? null);
    });
  });

  it('gives the same document as rejecting everything one paragraph at a time, in one undo step', async () => {
    const doc = await docWith(60, true);
    const state = EditorState.create({ schema, doc, plugins: [history(), review()] });
    let all = state;
    expect(rejectAllRevisions(state, (tr) => (all = state.apply(tr)))).toBe(true);
    expect(collectRevisions(all.doc)).toEqual([]);
    // The same through the selection path (each paragraph resolved on its own, no batching).
    let each = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
    for (let k = 0; k < 5 && collectRevisions(each.doc).length; k++) {
      const s = each.apply(each.tr.setSelection(new AllSelection(each.doc)));
      rejectRevision(s, (tr) => (each = s.apply(tr)));
    }
    expect(collectRevisions(each.doc)).toEqual([]);
    expect(all.doc.toJSON()).toEqual(each.doc.toJSON());
    expect(all.doc.textContent).toContain('deleted 3  | 粗體3x | 1 | same | italic');
    expect(all.doc.textContent).toContain('own ns');
    // One undo step gives the document back.
    let back = all;
    undo(all, (tr) => (back = all.apply(tr)));
    expect(back.doc.eq(state.doc)).toBe(true);
  });

  it('restores 3000 deletions without building a document for each one', async () => {
    const doc = await docWith(3000, false);
    const state = EditorState.create({ schema, doc, plugins: [history(), review()] });
    expect(collectRevisions(state.doc).length).toBe(3000 * 3 + 1);
    const proto = DOMParser.prototype;
    const original = proto.parseFromString;
    let parses = 0;
    proto.parseFromString = function (this: DOMParser, ...args: Parameters<DOMParser['parseFromString']>) {
      parses++;
      return original.apply(this, args);
    };
    let next = state;
    const t0 = performance.now();
    try {
      rejectAllRevisions(state, (tr) => (next = state.apply(tr)));
    } finally {
      proto.parseFromString = original;
    }
    const time = performance.now() - t0;
    console.log(`Reject All, 3000 paragraphs / 9001 revisions: ${time.toFixed(0)} ms, ${parses} XML documents built`);
    expect(collectRevisions(next.doc)).toEqual([]);
    expect(next.doc.childCount).toBe(3001);
    expect(next.doc.child(2999).textContent).toBe('Paragraph 2999 keeps deleted 2999 bold 2999');
    // Was one per deletion (3000): most of Reject All's time in jsdom.
    expect(parses).toBeLessThan(300);
  }, 60000);
});
