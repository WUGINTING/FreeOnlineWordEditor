// Writes documents with tracked changes, and the same documents after the editor accepted or
// rejected them, so Word itself can compare the result with its own Accept All / Reject All
// (scripts/inspect-review-in-word.ps1). Opt-in:
//   DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-review.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { EditorState, type Command } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { acceptAllRevisions, rejectAllRevisions, review } from '../../src/papyrus/editor/review';

const out = process.env.DOCX_EXPORT;
const qa = process.env.QA_DOCX;
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const A = (id: number, who = 'Alice') => `w:id="${id}" w:author="${who}" w:date="2026-01-02T03:04:00Z"`;
const P = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const R = (t: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const BR = '<w:r><w:br w:type="page"/></w:r>';
const TBL = (cell: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + cell + '</w:tc></w:tr></w:tbl>';

const CASES: Record<string, string> = {
  ins: P(R('Keep ') + `<w:ins ${A(1)}>${R('INSERTED ', '<w:b/>')}</w:ins>` + R('end.')),
  del: P(R('X') + `<w:del ${A(1)}><w:r w:rsidDel="00112233"><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve"> gone </w:delText></w:r><w:r><w:tab/></w:r></w:del>` + R('Y')),
  nested: P(R('X') + `<w:ins ${A(2, 'U')}><w:del ${A(3, 'V')}><w:r><w:delText>both</w:delText></w:r></w:del></w:ins>` + R('Z')),
  format: P(R('Para ') + `<w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:rPrChange ${A(5)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>` + R(' me.')),
  paraFormat: P(R('Centered now'), `<w:jc w:val="center"/><w:pPrChange ${A(1)}><w:pPr><w:jc w:val="right"/><w:ind w:left="720"/></w:pPr></w:pPrChange>`) + P(R('next')),
  markDel: P(R('A'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr>`) + P(R('B'), '<w:jc w:val="right"/>') + P(R('C')),
  markIns: P(R('A'), `<w:jc w:val="center"/><w:rPr><w:ins ${A(1)}/></w:rPr>`) + P(R('B'), '<w:jc w:val="right"/>') + P(R('C')),
  wholePara: P(`<w:del ${A(2)}><w:r><w:delText>Gone paragraph</w:delText></w:r></w:del>`, `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr>`) + P(R('B'), '<w:jc w:val="right"/>'),
  move:
    P(`<w:moveFromRangeStart w:id="1" w:author="M" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveFrom ${A(2, 'M')}>${R('moved ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` + R('stay')) +
    P(R('end ') + `<w:moveToRangeStart w:id="3" w:author="M" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveTo ${A(4, 'M')}>${R('moved ')}</w:moveTo><w:moveToRangeEnd w:id="3"/>`),
  comment:
    P(R('Before ') + '<w:commentRangeStart w:id="7"/>' + `<w:ins ${A(8)}>${R('commented')}</w:ins>` + '<w:commentRangeEnd w:id="7"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="7"/></w:r>' + R(' after.')),
  // A page break inside an insertion; a paragraph split at a page break with paragraph revisions.
  pageBreakIns: P(R('x') + `<w:ins ${A(1)}>${R('a')}${BR}${R('b')}</w:ins>` + R('y')),
  brGroupMarks: P(R('X') + BR + R('Y'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr><w:pPrChange ${A(2)}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`) + P(R('Z')),
  // Paragraph marks that can't join the next paragraph: before a table, at the end.
  markBeforeTable: P(R('A')) + P(R('B'), `<w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr>`) + TBL(P(R('cell'), '<w:jc w:val="right"/>')) + P(R('after')),
  markInsBeforeTable: P(R('A')) + P(R('B'), `<w:rPr><w:ins ${A(1)}/></w:rPr>`) + TBL(P(R('cell'))) + P(R('after')),
  emptyBeforeTable: P(R('A')) + P(`<w:del ${A(2)}><w:r><w:delText>B</w:delText></w:r></w:del>`, `<w:rPr><w:del ${A(1)}/></w:rPr>`) + TBL(P(R('cell'))) + P(R('after')),
  markLast: P(R('A')) + P(R('B'), `<w:rPr><w:del ${A(1)}/></w:rPr>`),
  markLastEmpty: P(R('A')) + P('', `<w:rPr><w:del ${A(1)}/></w:rPr>`),
  // Revisions and range markers inside a deletion.
  fmtInDel: P(R('a') + `<w:del ${A(1)}><w:r><w:rPr><w:b/><w:rPrChange ${A(2, 'Bob')}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:delText>gone</w:delText></w:r></w:del>` + R('z')),
  bookmarkInDel: P(R('a') + `<w:del ${A(1)}><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:delText>gone</w:delText></w:r></w:del>` + R('kept') + '<w:bookmarkEnd w:id="0"/>' + R('z')),
  commentInDel:
    P(R('a') + `<w:del ${A(1)}><w:commentRangeStart w:id="7"/><w:r><w:delText>gone</w:delText></w:r></w:del>` + R('kept') + '<w:commentRangeEnd w:id="7"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="7"/></w:r>' + R('z')),
};
CASES.all = Object.entries(CASES).filter(([k]) => !k.startsWith('comment') && !k.startsWith('markLast')).map(([, v]) => v).join('');

async function withComments(zip: JSZip): Promise<void> {
  zip.file('word/comments.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}><w:comment w:id="7" w:author="Ann" w:date="2026-01-02T00:00:00Z" w:initials="A"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t>Please check</w:t></w:r></w:p></w:comment></w:comments>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdPxC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
}

async function source(body: string, comments: boolean): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  if (comments) await withComments(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function resolved(bytes: Uint8Array, cmd: Command): Promise<Uint8Array> {
  const { doc, model } = await readDocx(bytes);
  let state = EditorState.create({ schema, doc, plugins: [review()] });
  cmd(state, (tr) => (state = state.apply(tr)));
  return writeDocx(state.doc, model);
}

describe.skipIf(!out)('export documents after accepting / rejecting tracked changes', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });
    for (const [name, body] of Object.entries(CASES)) {
      const src = await source(body, name.startsWith('comment'));
      writeFileSync(join(out!, `${name}-src.docx`), src);
      // The source saved unchanged by the editor must still hold every revision.
      const { doc, model } = await readDocx(src);
      writeFileSync(join(out!, `${name}-saved.docx`), await writeDocx(doc, model));
      writeFileSync(join(out!, `${name}-accept.docx`), await resolved(src, acceptAllRevisions));
      writeFileSync(join(out!, `${name}-reject.docx`), await resolved(src, rejectAllRevisions));
    }
    if (qa) {
      const src = new Uint8Array(readFileSync(qa));
      writeFileSync(join(out!, 'qa-src.docx'), src);
      const { doc, model } = await readDocx(src);
      writeFileSync(join(out!, 'qa-saved.docx'), await writeDocx(doc, model));
      writeFileSync(join(out!, 'qa-accept.docx'), await resolved(src, acceptAllRevisions));
      writeFileSync(join(out!, 'qa-reject.docx'), await resolved(src, rejectAllRevisions));
    }
  });
});
