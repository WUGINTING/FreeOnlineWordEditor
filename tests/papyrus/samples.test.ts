// Round-trips real-world .docx files from a local folder.
// Opt-in: DOCX_SAMPLES=<folder> npx vitest run tests/samples.test.ts
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { stableIds } from './helpers';
import { compareElements } from './canonical';
import JSZip from 'jszip';
import { parseXml } from '../../src/papyrus/docx/xml';

const dir = process.env.DOCX_SAMPLES;

function collect(root: string, out: string[] = []): string[] {
  for (const name of readdirSync(root)) {
    const p = join(root, name);
    if (name === 'node_modules' || name.startsWith('.')) continue;
    const st = statSync(p);
    if (st.isDirectory()) collect(p, out);
    else if (name.endsWith('.docx') && !name.startsWith('~$')) out.push(p);
  }
  return out;
}

describe.skipIf(!dir)('real documents', () => {
  const files = dir ? collect(dir) : [];
  for (const file of files) {
    it(file.slice(dir!.length + 1), async () => {
      const first = await readDocx(readFileSync(file));
      const saved = await writeDocx(first.doc, first.model);
      const second = await readDocx(saved);
      expect(second.doc.textContent).toBe(first.doc.textContent);
      expect(stableIds(second.doc.toJSON())).toEqual(stableIds(first.doc.toJSON()));

      // Strict: the saved document.xml body must match the original element by element.
      const partXml = async (bytes: Uint8Array | Buffer, name: string) =>
        parseXml(await (await JSZip.loadAsync(bytes)).file(name)!.async('string')).documentElement;
      const diffs = compareElements(await partXml(readFileSync(file), 'word/document.xml'), await partXml(saved, 'word/document.xml'));
      expect(diffs.join('\n')).toBe('');

      // Force every header/footer to be rewritten, as if the user had edited it.
      for (const hf of first.model.headerFooters) hf.dirty = true;
      const thirdBytes = await writeDocx(first.doc, first.model);
      const third = await readDocx(thirdBytes);
      // Strict for rewritten headers/footers too.
      for (const hf of first.model.headerFooters) {
        const d = compareElements(await partXml(readFileSync(file), hf.part!), await partXml(thirdBytes, hf.part!));
        expect(d.join('\n')).toBe('');
      }
      const shape = (m: typeof first.model) =>
        stableIds(m.headerFooters.map((h) => ({ kind: h.kind, type: h.type, doc: h.doc.toJSON() })));
      expect(shape(third.model)).toEqual(shape(first.model));
      expect(third.model.titlePage).toBe(first.model.titlePage);
      expect(stableIds(third.doc.toJSON())).toEqual(stableIds(first.doc.toJSON()));
    }, 60_000);
  }
});
