// Writes documents edited with the formatting commands, so Word itself can confirm the result
// (scripts/inspect-intent-in-word.ps1). Opt-in:
//   DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-intent.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { clearFormatting, setAlign, toggleFormat } from '../../src/papyrus/editor/commands';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';

const out = process.env.DOCX_EXPORT;
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Banner"><w:name w:val="Banner"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
  '</w:styles>';

async function edit(body: string, steps: (s: { state: EditorState; run: (c: Command) => void; model: any }) => void): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  zip.file('word/styles.xml', STYLES);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  const box = { state: EditorState.create({ schema, doc }), model, run: (c: Command) => {} };
  box.run = (c) => void c(box.state, (tr) => (box.state = box.state.apply(tr)));
  steps(box);
  return writeDocx(box.state.doc, model);
}

const selectAll = (s: { state: EditorState }) =>
  (s.state = s.state.apply(s.state.tr.setSelection(TextSelection.create(s.state.doc, 1, s.state.doc.content.size - 1))));

describe.skipIf(!out)('export documents edited with formatting commands', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });

    // 1. Clear formatting over a tracked insertion and a content control.
    writeFileSync(join(out!, 'clear-formatting.docx'), await edit(
      '<w:p><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>added text</w:t></w:r></w:ins></w:p>' +
        '<w:p><w:sdt><w:sdtPr><w:alias w:val="客戶名稱"/><w:tag w:val="customer"/><w:id w:val="42"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:i/></w:rPr><w:t>王小明</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      (s) => {
        selectAll(s);
        s.run(clearFormatting);
      },
    ));

    // 2. A centered bold style: align left and turn bold off.
    writeFileSync(join(out!, 'style-override.docx'), await edit(
      '<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr><w:r><w:t>Banner text</w:t></w:r></w:p>',
      (s) => {
        selectAll(s);
        s.run(setAlign('left', s.model.styles));
        s.run(toggleFormat('bold', s.model.styles));
      },
    ));

    // 3. A pasted ordered list starting at 3 with a nested bullet.
    writeFileSync(join(out!, 'pasted-list.docx'), await edit('<w:p><w:r><w:t>before</w:t></w:r></w:p>', (s) => {
      const div = document.createElement('div');
      div.innerHTML = '<ol start="3"><li>three<ul><li>sub</li></ul></li><li>four</li></ol>';
      const slice = adoptPastedLists(PMDOMParser.fromSchema(schema).parseSlice(div), s.model.numbering);
      const end = s.state.doc.content.size;
      s.state = s.state.apply(s.state.tr.replace(end, end, slice));
    }));
  });
});
