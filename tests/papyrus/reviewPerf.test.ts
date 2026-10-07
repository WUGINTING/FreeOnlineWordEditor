// Review on a large document (3000 paragraphs, 9000 revisions): typing must stay fast (the
// revision markup is updated for the changed paragraph only) and Accept All must not be
// quadratic. Times are for jsdom, with the editor view and its decorations.
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { acceptAllRevisions, collectRevisions, rejectAllRevisions, review, reviewSummary } from '../../src/papyrus/editor/review';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const PARAS = 3000;

async function bigDoc() {
  const A = (id: number) => `w:id="${id}" w:author="U${id % 7}" w:date="2026-01-02T03:04:00Z"`;
  let body = '';
  for (let i = 0; i < PARAS; i++) {
    body +=
      `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i} keeps </w:t></w:r>` +
      `<w:ins ${A(3 * i)}><w:r><w:t xml:space="preserve">inserted ${i} </w:t></w:r></w:ins>` +
      `<w:del ${A(3 * i + 1)}><w:r><w:delText xml:space="preserve">deleted ${i} </w:delText></w:r></w:del>` +
      `<w:r><w:rPr><w:b/><w:rPrChange ${A(3 * i + 2)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>bold ${i}</w:t></w:r></w:p>`;
  }
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  return (await readDocx(await zip.generateAsync({ type: 'uint8array' }))).doc;
}

describe('9. review performance', () => {
  it('typing and Accept / Reject All on 3000 paragraphs with 9000 revisions', async () => {
    const doc = await bigDoc();
    const state = EditorState.create({ schema, doc, plugins: [history(), review()] });
    expect(collectRevisions(state.doc).length).toBe(PARAS * 3);
    let pos = 0;
    for (let i = 0; i < PARAS / 2; i++) pos += state.doc.child(i).nodeSize;

    /** A typical keystroke in the middle of the document (timing in a shared test run is noisy). */
    const typeIn = (s: EditorState, summary: boolean) => {
      const host = document.createElement('div');
      document.body.append(host);
      const view = new EditorView(host, { state: s });
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos + 5)));
      const times: number[] = [];
      for (let k = 0; k < 8; k++) {
        const t0 = performance.now();
        view.dispatch(view.state.tr.insertText('x'));
        if (summary) reviewSummary(view.state); // as the toolbar does
        times.push(performance.now() - t0);
      }
      const marked = view.dom.querySelectorAll('.dx-rev-fmt').length;
      view.destroy();
      host.remove();
      return { time: times.sort((a, b) => a - b)[2], marked };
    };
    // The same editor without review, for how busy the machine is right now.
    const baseline = typeIn(EditorState.create({ schema, doc, plugins: [history()] }), false).time;
    const { time: typing, marked } = typeIn(state, true);
    expect(marked).toBe(PARAS);

    const timed = (cmd: typeof acceptAllRevisions) => {
      let best = Infinity;
      for (let k = 0; k < 2; k++) {
        let next = state;
        const t0 = performance.now();
        cmd(state, (tr) => (next = state.apply(tr)));
        best = Math.min(best, performance.now() - t0);
        expect(collectRevisions(next.doc)).toEqual([]);
      }
      return best;
    };
    const acceptAll = timed(acceptAllRevisions);
    const rejectAll = timed(rejectAllRevisions);

    console.log(`review perf: typing ${typing.toFixed(1)} ms (without review ${baseline.toFixed(1)} ms), Accept All ${acceptAll.toFixed(0)} ms, Reject All ${rejectAll.toFixed(0)} ms`);
    // ~100 ms on an idle machine; more headroom when other test files keep it busy.
    expect(typing).toBeLessThan(Math.max(100, baseline * 3));
    expect(acceptAll).toBeLessThan(1500);
    expect(rejectAll).toBeLessThan(4000); // restoring 3000 deletions parses their XML (slow in jsdom)
  }, 120000);
});
