// Writes edited copies of real documents so they can be opened in Word
// (scripts/open-in-word.ps1). Opt-in:
//   DOCX_SAMPLES=<in> DOCX_EXPORT=<out> npx vitest run tests/export-edited.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import { setAlign, splitParagraph, toggle } from '../../src/papyrus/editor/commands';

const src = process.env.DOCX_SAMPLES;
const out = process.env.DOCX_EXPORT;

function collect(root: string, acc: string[] = []): string[] {
  for (const name of readdirSync(root)) {
    const p = join(root, name);
    if (statSync(p).isDirectory()) collect(p, acc);
    else if (name.endsWith('.docx') && !name.startsWith('~$')) acc.push(p);
  }
  return acc;
}

describe.skipIf(!src || !out)('export edited copies', () => {
  it('writes edited copies', async () => {
    mkdirSync(out!, { recursive: true });
    for (const [i, file] of collect(src!).entries()) {
      const { doc, model } = await readDocx(readFileSync(file));
      let state = EditorState.create({ schema, doc });
      const run = (cmd: (s: EditorState, d: (tr: any) => void) => boolean) => cmd(state, (tr) => (state = state.apply(tr)));
      // Every text paragraph: type, bold a word, center every third one, split every fifth.
      const positions: number[] = [];
      state.doc.descendants((n, pos) => {
        if (n.type.name === 'paragraph' && n.textContent.length > 4) positions.push(pos);
        return true;
      });
      for (const [k, pos] of positions.slice(0, 60).reverse().entries()) {
        state = state.apply(state.tr.insertText('✎', pos + 1));
        state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + 1, pos + 4)));
        run(toggle('bold'));
        if (k % 3 === 0) run(setAlign('center'));
        if (k % 5 === 0) {
          state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + 3)));
          run(splitParagraph);
        }
      }
      // Headers/footers: prefix text.
      for (const hf of model.headerFooters) {
        const s = EditorState.create({ schema, doc: hf.doc });
        hf.doc = s.apply(s.tr.insertText('【改】', 1)).doc;
        hf.dirty = true;
      }
      writeFileSync(join(out!, `edited-${String(i).padStart(3, '0')}.docx`), await writeDocx(state.doc, model));
    }
  }, 600_000);
});
