// Find / replace on big documents: moving between matches and typing do not rescan the whole
// document, highlighting is capped, "replace all" is one step per paragraph. Also: hidden
// field codes end a stretch of text, and case folding handles final sigma and Turkish İ.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import {
  MAX_HIGHLIGHTS, findMatches, findNext, replaceAll, replaceCurrent, searchPlugin, searchState, setSearch,
} from '../../src/papyrus/editor/search';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: [history(), searchPlugin()] });
  let last: Transaction | null = null;
  const run = (cmd: Command) => cmd(state, (tr) => ((last = tr), (state = state.apply(tr))));
  const apply = (tr: Transaction) => (state = state.apply(tr));
  const save = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  const decorations = () => searchState(state).decorations.find();
  return { run, apply, save, decorations, get last() { return last!; }, get state() { return state; } };
}

const plain = (s: { matches: { from: number; to: number }[] }) => s.matches.map((m) => [m.from, m.to]);

describe('finding in big documents', () => {
  it('moving to the next match keeps the matches and only swaps the current highlight', async () => {
    const d = await open(`<w:p>${r('a b a b a')}</w:p>`);
    d.run(setSearch('a'));
    d.run(findNext(1)); // selects the first match
    const before = searchState(d.state);
    const decos = d.decorations();
    d.run(findNext(1));
    const after = searchState(d.state);
    expect(after.matches).toBe(before.matches);
    expect(after.current).toBe(before.current + 1);
    const now = d.decorations();
    // The highlight of the match that stayed plain is the same decoration (find() returns copies).
    const untouched = decos.find((x) => x.from !== before.matches[before.current].from && x.from !== after.matches[after.current].from)!;
    expect(now.some((x) => (x as any).type === (untouched as any).type)).toBe(true);
    expect(now.filter((x) => x.spec.current).map((x) => x.from)).toEqual([after.matches[after.current].from]);
    expect(now).toHaveLength(3);
  });

  it('typing rescans only the paragraph that changed, with the same result as a full search', async () => {
    const d = await open(`<w:p>${r('cat one')}</w:p><w:p>${r('two cat')}</w:p><w:p>${r('three')}</w:p>`);
    d.run(setSearch('cat'));
    const first = searchState(d.state).matches[0];
    // Type "cat" at the end of the last paragraph.
    d.apply(d.state.tr.insertText(' cat', d.state.doc.content.size - 1));
    const s = searchState(d.state);
    expect(s.matches[0]).toBe(first); // before the edit: kept as it was
    expect(plain(s)).toEqual(plain({ matches: findMatches(d.state.doc, 'cat') }));
    // Joining two paragraphs, marking text as a link, deleting across paragraphs: always as a full search.
    const edits: ((st: EditorState) => Transaction)[] = [
      (st) => st.tr.insertText('c', 3).insertText('at', 4),
      (st) => st.tr.delete(st.doc.child(0).nodeSize - 1, st.doc.child(0).nodeSize + 1),
      (st) => st.tr.addMark(2, 4, schema.marks.link.create({ href: 'https://x' })),
      (st) => st.tr.delete(2, st.doc.content.size - 3),
    ];
    // Highlights: exactly one per match, the current one marked.
    const highlights = () => d.decorations().map((x) => [x.from, x.to, !!x.spec.current]).sort((a, b) => (a[0] as number) - (b[0] as number));
    const wanted = () => searchState(d.state).matches.map((m, i) => [m.from, m.to, i === searchState(d.state).current]);
    expect(highlights()).toEqual(wanted());
    for (const edit of edits) {
      d.apply(edit(d.state));
      expect(plain(searchState(d.state))).toEqual(plain({ matches: findMatches(d.state.doc, 'cat') }));
      expect(highlights()).toEqual(wanted());
    }
  });

  it('random typing and deleting always agree with a full search', async () => {
    const d = await open(`<w:p>${r('aba')}</w:p><w:p>${r('bab', '<w:b/>')}</w:p><w:p>${r('ab')}</w:p>`);
    d.run(setSearch('ab'));
    let seed = 7;
    const rand = (n: number) => ((seed = (seed * 1103515245 + 12345) % 2147483648) % n);
    for (let i = 0; i < 300; i++) {
      const size = d.state.doc.content.size;
      const pos = 1 + rand(size - 1);
      const $p = d.state.doc.resolve(pos);
      if (!$p.parent.isTextblock) continue;
      if (rand(3)) d.apply(d.state.tr.insertText(['a', 'b', 'ab', ' '][rand(4)], pos));
      else if (pos + 2 < size) {
        const tr = d.state.tr.delete(pos, pos + 1 + rand(2));
        if (tr.doc.content.size > 0) d.apply(tr);
      }
      if (i % 50 === 0) d.run(findNext(1));
      expect(plain(searchState(d.state)), `step ${i}`).toEqual(plain({ matches: findMatches(d.state.doc, 'ab') }));
      const s = searchState(d.state);
      expect(d.decorations().map((x) => [x.from, x.to, !!x.spec.current]).sort((a, b) => (a[0] as number) - (b[0] as number))).toEqual(
        s.matches.map((m, k) => [m.from, m.to, k === s.current]),
      );
    }
  });

  it('highlights at most MAX_HIGHLIGHTS matches but counts them all', async () => {
    const n = MAX_HIGHLIGHTS + 1000;
    const d = await open(`<w:p>${r('x '.repeat(n))}</w:p>`);
    d.run(setSearch('x'));
    expect(searchState(d.state).matches).toHaveLength(n);
    expect(d.decorations().length).toBeLessThanOrEqual(MAX_HIGHLIGHTS + 1);
    // The current match is highlighted even beyond the cap.
    d.apply(d.state.tr.setSelection(TextSelection.create(d.state.doc, 2 * (n - 5))));
    d.run(findNext(1));
    const cur = searchState(d.state).matches[searchState(d.state).current];
    expect(d.decorations().some((x) => x.from === cur.from && x.to === cur.to)).toBe(true);
  });

  it('replace all is one replacement per paragraph, keeping formatting, markers and a single undo', async () => {
    const d = await open(
      `<w:p>${r('ab ab', '<w:b/>')}${r(' ab')}</w:p><w:p>${r('a')}<w:bookmarkStart w:id="0" w:name="bm"/>${r('b ab')}<w:bookmarkEnd w:id="0"/></w:p><w:p>${r('none')}</w:p>`,
    );
    d.run(setSearch('ab'));
    expect(searchState(d.state).matches).toHaveLength(5);
    d.run(replaceAll('XY'));
    expect(d.last.steps).toHaveLength(2);
    expect(d.state.doc.textContent).toBe('XY XY XYXY XYnone');
    const xml = await d.save();
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>XY XY</w:t></w:r><w:r><w:t xml:space="preserve"> XY</w:t></w:r>');
    expect(xml).toContain('<w:t>XY</w:t></w:r><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:t xml:space="preserve"> XY</w:t></w:r><w:bookmarkEnd w:id="0"/>');
    d.run(undo);
    expect(d.state.doc.textContent).toBe('ab ab abab abnone');
  });

  it('replace all on many matches is fast', async () => {
    const d = await open(Array.from({ length: 300 }, () => `<w:p>${r('aa '.repeat(100))}</w:p>`).join(''));
    d.run(setSearch('a'));
    expect(searchState(d.state).matches).toHaveLength(60000);
    const t = performance.now();
    d.run(replaceAll('b'));
    expect(performance.now() - t).toBeLessThan(5000);
    expect(d.last.steps).toHaveLength(300);
    expect(searchState(d.state).matches).toHaveLength(0);
  });

  it('replacing one match continues from it, as before', async () => {
    const d = await open(`<w:p>${r('a-a')}</w:p><w:p>${r('a')}</w:p>`);
    d.run(setSearch('a'));
    d.run(replaceCurrent('aa'));
    expect(d.state.doc.textContent).toBe('aa-aa');
    expect(searchState(d.state).current).toBe(2);
    expect(plain(searchState(d.state))).toEqual(plain({ matches: findMatches(d.state.doc, 'a') }));
  });
});

describe('what a match may span', () => {
  it('hidden field codes end a stretch of text', async () => {
    const d = await open(
      `<w:p>${r('ab')}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF bm </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${r('cd')}</w:p>`,
    );
    d.run(setSearch('bc'));
    expect(searchState(d.state).matches).toHaveLength(0);
    d.run(setSearch('cd'));
    expect(searchState(d.state).matches).toHaveLength(1);
  });

  it('case folding matches final sigma and Turkish dotted capital I', async () => {
    const d = await open(`<w:p>${r('ΛΟΓΟΣ λόγος İSTANBUL')}</w:p>`);
    d.run(setSearch('λογος'));
    expect(searchState(d.state).matches).toHaveLength(1);
    d.run(setSearch('ΛΟΓΟΣ'));
    expect(searchState(d.state).matches).toHaveLength(1);
    d.run(setSearch('istanbul'));
    expect(searchState(d.state).matches).toHaveLength(1);
    d.run(setSearch('İstanbul', true));
    expect(searchState(d.state).matches).toHaveLength(0);
  });
});
