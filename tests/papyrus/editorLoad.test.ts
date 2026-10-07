// DocxEditor.vue opening one document after another: the comments and compatibility notice
// shown must be the latest document's, even when an older document finishes opening last.
// The component is compiled here (the test setup has no Vue plugin) with a stand-in editor.
import { afterAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import JSZip from 'jszip';
import { createApp, h, nextTick, reactive } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { CompatReport } from '../../src/papyrus/docx/compat';
import { gates } from './editorLoadStubs';

const SFC = resolve(__dirname, '../../src/papyrus/vue/DocxEditor.vue');
const OUT = resolve(__dirname, `.docxEditor-${process.pid}.tmp.js`);
const slash = (p: string) => p.replace(/\\/g, '/');

afterAll(() => rmSync(OUT, { force: true }));

async function compiled(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'dx-editor-test', inlineTemplate: true });
  let code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText;
  const stubs = slash(resolve(__dirname, 'editorLoadStubs.ts'));
  code = code.replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) =>
    spec.endsWith('.vue') || spec === '../editor/core' ? `from '${stubs}'` : `from '${slash(resolve(dirname(SFC), spec))}'`,
  );
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ slash(OUT))).default;
}

async function docWithComment(): Promise<JSZip> {
  const zip = blankPackage();
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  zip.file('word/document.xml', `<w:document ${W}><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:body></w:document>`);
  zip.file('word/comments.xml', `<w:comments ${W}><w:comment w:id="1" w:author="Old"><w:p><w:r><w:t>stale</w:t></w:r></w:p></w:comment></w:comments>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  return JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
}

const flush = async () => {
  for (let i = 0; i < 20; i++) {
    await new Promise((r) => setTimeout(r, 0));
    await nextTick();
  }
};

describe('12. DocxEditor.vue loads', () => {
  // A cold transform of DocxEditor.vue can take longer than the default 5 s under load.
  it('drops the results of a document replaced while it was opening', { timeout: 20000 }, async () => {
    const Editor = await compiled();
    const A = { name: 'A', zip: await docWithComment() };
    const B = { name: 'B', zip: blankPackage() };
    const props = reactive<{ src: unknown }>({ src: A });
    const reports: CompatReport[] = [];
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(Editor, { src: props.src, toolbar: false, onCompat: (r: CompatReport) => reports.push(r) }) });
    app.mount(host);
    await flush();
    props.src = B; // a new document arrives while A is still opening
    await flush();
    // B opens quickly, A slowly (with the old code both ran at once).
    gates.get('B')?.();
    await flush();
    gates.get('A')?.();
    await flush();
    gates.get('B')?.();
    await flush();
    // Only B's report was shown: A had comments, B has none.
    expect(reports.length).toBeGreaterThan(0);
    expect(reports.every((r) => !r.items.some((i) => i.id === 'comments'))).toBe(true);
    app.unmount();
    host.remove();
  });
});
