// Review in the running editor: revisions and comments show in the page, the review commands
// work through DocxEditor.run, and each accept / reject is undone in one step.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readComments } from '../../src/papyrus/docx/comments';
import { scanCompat } from '../../src/papyrus/docx/compat';
import { writeDocx } from '../../src/papyrus/docx/writer';
import {
  acceptAllRevisions, collectRevisions, goToRevision, rejectRevision, revealComment, reviewSummary, setComments,
  toggleRevisionMarks,
} from '../../src/papyrus/editor/review';

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
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const A = (id: number, who = 'QA Reviewer') => `w:id="${id}" w:author="${who}" w:date="2026-09-24T00:00:00Z"`;

async function sample(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
      `<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:ins ${A(1)}><w:r><w:t>new</w:t></w:r></w:ins><w:del ${A(2)}><w:r><w:delText>old</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>see this</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>',
  );
  zip.file('word/comments.xml', `<w:comments ${W}><w:comment w:id="7" w:author="Ann" w:date="2026-09-24T01:00:00Z"><w:p><w:r><w:t>Check</w:t></w:r></w:p></w:comment></w:comments>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}

describe('review in the editor', () => {
  it('shows revisions and comments, accepts / rejects with one undo step each', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await sample());
    const view = editor.view!;

    // Shown in the page.
    expect(host.querySelector('.dx-rev-ins')?.textContent).toBe('new');
    expect(host.querySelector('.dx-rev-del')?.getAttribute('data-text')).toBe('old');
    const comments = await readComments(editor.model.zip);
    setComments(view, comments);
    const range = host.querySelector('.dx-comment[data-comment-id="7"]')!;
    expect(range.textContent).toBe('see this');
    expect(range.getAttribute('title')).toBe('留言（Ann）：Check');
    expect(revealComment(view, '7')).toBe(true);
    expect(view.state.doc.textBetween(view.state.selection.from, view.state.selection.to)).toBe('see this');

    // Hide markup: the editor gets the class, the document doesn't change.
    editor.run(toggleRevisionMarks);
    expect(view.dom.classList.contains('dx-rev-off')).toBe(true);
    expect(editor.isModified()).toBe(false);
    editor.run(toggleRevisionMarks);

    // Reject the deletion reached with "next", then undo it in one step.
    editor.run(goToRevision(1));
    editor.run(goToRevision(1));
    expect(reviewSummary(view.state).current?.kind).toBe('del');
    editor.run(rejectRevision);
    expect(view.state.doc.textContent).toContain('Hello newold');
    expect(editor.isModified()).toBe(true);
    editor.undo();
    expect(editor.isModified()).toBe(false);

    editor.run(acceptAllRevisions);
    expect(collectRevisions(view.state.doc)).toEqual([]);
    const saved = await JSZip.loadAsync(await writeDocx(view.state.doc, editor.model));
    const xml = await saved.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:r><w:t>Hello new</w:t></w:r>');
    expect(xml).toContain('<w:commentRangeStart w:id="7"/>');
    // Comments part untouched.
    expect(await saved.file('word/comments.xml')!.async('string')).toContain('<w:t>Check</w:t>');
    editor.undo();
    expect(editor.isModified()).toBe(false);

    // The compatibility report names both.
    const report = await scanCompat(editor.model.zip);
    expect(report.items.map((i) => i.id)).toEqual(['comments', 'revisions']);

    editor.destroy();
    host.remove();
  });
});
