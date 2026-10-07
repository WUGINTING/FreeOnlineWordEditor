// Writes documents edited with 追蹤修訂 on in the running editor (DocxEditor), with what the
// editor recorded and what its own Accept All / Reject All give, so Word can check them
// (scripts/inspect-track-changes-in-word.ps1): the revisions Word sees (type, author, date,
// text), Document.TrackRevisions, and Word's Accept All / Reject All against the editor's. Opt-in:
//   DOCX_EXPORT=<out> [WORD_REVISIONS_DOCX=<a file Word saved with revisions>] npx vitest run tests/papyrus/export-track.test.ts
import { beforeAll, describe, it } from 'vitest';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { TextSelection, type Command } from 'prosemirror-state';
import { deleteSelection, joinBackward } from 'prosemirror-commands';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { blankPackage } from '../../src/papyrus/docx/template';
import { deletedInfo } from '../../src/papyrus/docx/revisions';
import { groupId } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { acceptAllRevisions, collectRevisions, rejectAllRevisions } from '../../src/papyrus/editor/review';
import { insertPageBreak, setAlign, splitParagraph, toggleFormat } from '../../src/papyrus/editor/commands';

const out = process.env.DOCX_EXPORT;
const wordFile = process.env.WORD_REVISIONS_DOCX;
const AUTHOR = 'QA 測試員';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const P = (inner: string) => `<w:p>${inner}</w:p>`;
const R = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

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

async function pack(body: string, trackRevisions = false): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  if (trackRevisions) {
    const s = await zip.file('word/settings.xml')!.async('string');
    zip.file('word/settings.xml', s.replace('<w:defaultTabStop', '<w:trackRevisions/><w:defaultTabStop'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

/** A helper around the running editor: edits go through view.dispatch, as the user's do. */
class Session {
  constructor(readonly editor: DocxEditor) {}
  get view() {
    return this.editor.view!;
  }
  at(needle: string, offset = 0): number {
    let found = -1;
    this.view.state.doc.descendants((n, pos) => {
      if (found < 0 && n.isText && n.text!.includes(needle)) found = pos + n.text!.indexOf(needle) + offset;
      return found < 0;
    });
    if (found < 0) throw new Error('not found: ' + needle);
    return found;
  }
  cursor(pos: number) {
    this.view.dispatch(this.view.state.tr.setSelection(TextSelection.create(this.view.state.doc, pos)));
  }
  select(from: number, to: number) {
    this.view.dispatch(this.view.state.tr.setSelection(TextSelection.create(this.view.state.doc, from, to)));
  }
  type(text: string) {
    for (const ch of text) this.view.dispatch(this.view.state.tr.insertText(ch));
  }
  backspace() {
    const { empty, from, $from } = this.view.state.selection;
    if (!empty) this.run(deleteSelection);
    else if ($from.parentOffset === 0) this.run(joinBackward);
    else this.view.dispatch(this.view.state.tr.delete(from - 1, from));
  }
  run(cmd: Command) {
    cmd(this.view.state, this.view.dispatch, this.view);
  }
}

/** The text of each w:p (pieces split at page breaks together), joined by |. */
function paragraphsText(doc: PMNode): string {
  const out: string[] = [];
  let last: string | null = null;
  doc.forEach((n) => {
    if (n.type.name !== 'paragraph' && n.type.name !== 'page_break') return;
    const g = groupId(n);
    const text = n.type.name === 'paragraph' ? n.textContent : '';
    if (g && g === last) out[out.length - 1] += text;
    else out.push(text);
    last = g;
  });
  return out.join('|');
}

function revisionText(doc: PMNode, r: ReturnType<typeof collectRevisions>[number]): string {
  if (r.kind === 'del') return deletedInfo(doc.nodeAt(r.from)!.attrs.xml)?.text ?? '';
  if (r.kind === 'paraMark') return '\r';
  return doc.textBetween(r.from, r.to, '\r', '\f');
}

type Case = { body: string; track?: boolean; file?: Uint8Array; edit: (s: Session) => void | Promise<void> };

const CASES: Record<string, Case> = {
  insert: { body: P(R('Hello world')), edit: (s) => { s.cursor(s.at('Hello', 5)); s.type(' big'); } },
  delete: { body: P(R('Hello world')), edit: (s) => { s.cursor(s.at('world', 5)); s.backspace(); s.backspace(); } },
  ownInsertion: {
    body: P(R('Hello world')),
    edit: (s) => { s.cursor(s.at('Hello', 5)); s.type(' abc'); s.backspace(); },
  },
  crossParagraph: {
    body: P(R('One two')) + P(R('Three four')),
    edit: (s) => { s.select(s.at('two', 1), s.at('Three', 2)); s.backspace(); },
  },
  backspaceParagraph: { body: P(R('First')) + P(R('Second')), edit: (s) => { s.cursor(s.at('Second')); s.backspace(); } },
  enter: { body: P(R('Hello world')), edit: (s) => { s.cursor(s.at('world')); s.run(splitParagraph); s.type('New '); } },
  replace: { body: P(R('The cat sat')), edit: (s) => { s.select(s.at('cat'), s.at('cat', 3)); s.type('dog'); } },
  paste: {
    body: P(R('XY')),
    edit: (s) => {
      s.cursor(s.at('XY', 1));
      const slice = new Slice(Fragment.from([schema.nodes.paragraph.create(null, schema.text('pasted A')), schema.nodes.paragraph.create(null, schema.text('pasted B'))]), 1, 1);
      s.view.dispatch(s.view.state.tr.replaceSelection(slice).setMeta('uiEvent', 'paste'));
    },
  },
  format: { body: P(R('Hello world')), edit: (s) => { s.select(s.at('world'), s.at('world', 5)); s.run(toggleFormat('bold')); } },
  paragraphFormat: { body: P(R('Centered')), edit: (s) => { s.cursor(s.at('Centered', 2)); s.run(setAlign('center')); } },
  pageBreak: { body: P(R('Hello')), edit: (s) => { s.cursor(s.at('Hello', 2)); s.run(insertPageBreak); } },
  toggleOn: { body: P(R('Hello')), edit: (s) => { s.editor.setTrackChanges(true); s.cursor(s.at('Hello', 5)); s.type('!'); } },
  toggleOff: { body: P(R('Hello')), track: true, edit: (s) => { s.editor.setTrackChanges(false); } },
  chinese: {
    body: P(R('本公司同意付款')),
    edit: (s) => { s.cursor(s.at('同意', 2)); s.type('全額'); s.cursor(s.at('付款', 2)); s.backspace(); s.backspace(); s.type('支付'); },
  },
};

describe.skipIf(!out)('export documents edited with 追蹤修訂 for Word', () => {
  if (wordFile) {
    CASES.wordRevisions = {
      body: '',
      file: readFileSync(wordFile),
      edit: (s) => {
        // Word saved it with tracking on: it opens on.
        if (!s.editor.trackChanges()) throw new Error('w:trackRevisions not read');
        const doc = s.view.state.doc;
        let firstText = '';
        doc.descendants((n) => {
          if (!firstText && n.isText && n.text!.length > 3) firstText = n.text!;
          return !firstText;
        });
        // Inside Word's own insertion (another author), and in the text Word had before.
        s.cursor(s.at(firstText, 2));
        s.type('網頁');
        s.cursor(s.at(firstText.slice(-3), 3));
        s.backspace();
        s.cursor(s.at('第一段', 3));
        s.backspace();
        s.backspace();
        s.cursor(s.at('paragraph', 9));
        s.type(' typed on the web');
      },
    };
  }
  if (wordFile) CASES.wordUnchanged = { body: '', file: readFileSync(wordFile), edit: () => {} };
  for (const [name, c] of Object.entries(CASES)) {
    it(name, async () => {
      mkdirSync(out!, { recursive: true });
      const host = document.createElement('div');
      document.body.append(host);
      const editor = new DocxEditor(host, { author: { name: AUTHOR } });
      const input = c.file ?? (await pack(c.body, c.track));
      await editor.open(input);
      if (!['toggleOn', 'toggleOff', 'wordRevisions', 'wordUnchanged'].includes(name)) editor.setTrackChanges(true);
      const s = new Session(editor);
      await c.edit(s);
      const bytes = new Uint8Array(await (await JSZip.loadAsync(await editor.save())).generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
      writeFileSync(join(out!, `${name}.docx`), bytes);
      const settingsOf = async (b: Uint8Array) => (await JSZip.loadAsync(b)).file('word/settings.xml')!.async('string');
      const settingsSame = (await settingsOf(input)) === (await settingsOf(bytes));
      const modified = editor.isModified();
      const doc = s.view.state.doc;
      const revisions = collectRevisions(doc).map((r) => ({ kind: r.kind, author: r.author, date: r.date, text: revisionText(doc, r) }));
      s.run(acceptAllRevisions);
      const accepted = paragraphsText(s.view.state.doc);
      // Back to the tracked document (Accept All is one undo step), then Reject All.
      if (revisions.length) editor.undo();
      if (!s.view.state.doc.eq(doc)) throw new Error('undo did not restore the tracked document');
      s.run(rejectAllRevisions);
      const rejected = paragraphsText(s.view.state.doc);
      writeFileSync(join(out!, `${name}.json`), JSON.stringify({ trackRevisions: editor.trackChanges(), settingsSame, modified, revisions, accepted, rejected }, null, 1));
      editor.destroy();
    });
  }
});
