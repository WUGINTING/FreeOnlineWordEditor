// Writes a round-tripped copy of every .docx in DOCX_SAMPLES into DOCX_EXPORT,
// so the copies can be compared in Microsoft Word (see scripts/compare-in-word.ps1).
// Opt-in: DOCX_SAMPLES=<in> DOCX_EXPORT=<out> npx vitest run tests/export.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';

const src = process.env.DOCX_SAMPLES;
const out = process.env.DOCX_EXPORT;

function collect(root: string, acc: string[] = []): string[] {
  for (const name of readdirSync(root)) {
    const p = join(root, name);
    if (name === 'node_modules' || name.startsWith('.')) continue;
    if (statSync(p).isDirectory()) collect(p, acc);
    else if (name.endsWith('.docx') && !name.startsWith('~$')) acc.push(p);
  }
  return acc;
}

describe.skipIf(!src || !out)('export round-tripped copies', () => {
  it('writes copies', async () => {
    mkdirSync(out!, { recursive: true });
    const index: { original: string; copy: string }[] = [];
    for (const [i, file] of collect(src!).entries()) {
      const { doc, model } = await readDocx(readFileSync(file));
      // Rewrite headers/footers too, as if the user had edited them.
      for (const hf of model.headerFooters) hf.dirty = true;
      const copy = join(out!, `${String(i).padStart(3, '0')}.docx`);
      writeFileSync(copy, await writeDocx(doc, model));
      index.push({ original: file, copy });
      console.log(relative(src!, file), '->', copy);
    }
    writeFileSync(join(out!, 'index.json'), JSON.stringify(index, null, 2));
  }, 600_000);
});
