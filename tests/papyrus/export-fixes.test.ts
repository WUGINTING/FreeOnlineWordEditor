// Writes documents exercising the table / picture writer fixes (merged cells whose span changed,
// relationships inside merged pieces, w:hMerge, drawing ids, media types) so Word itself can
// confirm they open. Opt-in:
//   DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-fixes.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { CellSelection, addColumnAfter, mergeCells } from 'prosemirror-tables';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { setCellBackground } from '../../src/papyrus/editor/tableCommands';

const out = process.env.DOCX_EXPORT;
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const RED = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAADUlEQVR4nGP4z8AARAAI/gH/xp559wAAAABJRU5ErkJggg==';
const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

const BORDERS = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('');
const tc = (text: string, tcPr = '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr>', inner = '') =>
  `<w:tc>${tcPr}<w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}${inner}</w:p></w:tc>`;
const table = (cols: number, rows: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${BORDERS}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
  `<w:tblGrid>${'<w:gridCol w:w="2800"/>'.repeat(cols)}</w:tblGrid>${rows.map((r) => `<w:tr>${r}</w:tr>`).join('')}</w:tbl><w:p/>`;
const drawing = (id: number, rid: string) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="p${id}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

interface Box {
  state: EditorState;
  run: (c: Command) => void;
  cell: (text: string) => number;
}

async function edit(body: string, steps: (s: Box) => void, rels = '', media: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${SECT}</w:body></w:document>`);
  for (const [path, b64] of Object.entries(media)) zip.file(path, b64, { base64: true });
  if (rels) {
    const old = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', old.replace('</Relationships>', rels + '</Relationships>'));
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', ct.replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'));
  }
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  const box: Box = {
    state: EditorState.create({ schema, doc }),
    run: (c) => void c(box.state, (tr) => (box.state = box.state.apply(tr))),
    cell: (text) => {
      let found = -1;
      box.state.doc.descendants((n, pos) => {
        if (found < 0 && n.type.name === 'table_cell' && n.textContent === text) found = pos;
        return found < 0;
      });
      return found;
    },
  };
  steps(box);
  return writeDocx(box.state.doc, model);
}

describe.skipIf(!out)('export documents for the writer fixes', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });
    const at = (s: Box, t: string) => (s.state = s.state.apply(s.state.tr.setSelection(TextSelection.create(s.state.doc, s.cell(t) + 2))));

    // 1a. A column added inside a 2x2 merged block (continuation piece must get gridSpan 3).
    writeFileSync(join(out!, 'merge-addcol.docx'), await edit(
      table(2, [
        tc('M', '<w:tcPr><w:tcW w:w="5600" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>'),
        tc('', '<w:tcPr><w:tcW w:w="5600" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>'),
        tc('A3') + tc('B3'),
      ]),
      (s) => {
        at(s, 'A3');
        s.run(addColumnAfter);
        at(s, 'M');
        s.run(setCellBackground('#ffff00'));
      },
    ));

    // 1b. A 2-row merged cell merged with the next column.
    writeFileSync(join(out!, 'merge-merge.docx'), await edit(
      table(2, [
        tc('A1', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>') + tc('B1'),
        tc('', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:vMerge/></w:tcPr>') + tc('B2'),
        tc('A3') + tc('B3'),
      ]),
      (s) => {
        s.state = s.state.apply(s.state.tr.setSelection(new CellSelection(s.state.doc.resolve(s.cell('A1')), s.state.doc.resolve(s.cell('B2')))));
        s.run(mergeCells);
      },
    ));

    // 3. A link and a picture inside a merged piece; 9. a legacy w:hMerge table.
    writeFileSync(join(out!, 'merged-rels.docx'), await edit(
      table(2, [
        tc('A', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>') + tc('B1'),
        tc('', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:vMerge/></w:tcPr>', '<w:hyperlink r:id="rIdLink"><w:r><w:t>example.com</w:t></w:r></w:hyperlink>' + drawing(3, 'rIdPic')) + tc('B2'),
      ]) +
        table(2, [
          tc('H', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:hMerge w:val="restart"/></w:tcPr>') + tc('', '<w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:hMerge/></w:tcPr>'),
          tc('C') + tc('D'),
        ]),
      (s) => {
        at(s, 'B1');
        s.run(setCellBackground('#00ff00'));
      },
      `<Relationship Id="rIdLink" Type="${REL}hyperlink" Target="https://example.com/" TargetMode="External"/>` +
        `<Relationship Id="rIdPic" Type="${REL}image" Target="media/merged.png"/>`,
      { 'word/media/merged.png': RED },
    ));

    // 10. New pictures after existing drawing ids 1 and 7; 8. media typed by their bytes.
    writeFileSync(join(out!, 'pictures.docx'), await edit(
      `<w:p>${drawing(1, 'rIdPic')}${drawing(7, 'rIdPic')}</w:p><w:p><w:r><w:t>new:</w:t></w:r></w:p>`,
      (s) => {
        const end = s.state.doc.content.size - 1;
        s.run((st, dispatch) => {
          dispatch?.(
            st.tr.insert(end, [
              schema.nodes.image.create({ src: `data:image/png;base64,${GIF}`, width: 40, height: 40, alt: 'gif labelled png' }),
              schema.nodes.image.create({ src: `data:application/octet-stream;base64,${RED}`, width: 80, height: 40, alt: 'untyped png' }),
            ]),
          );
          return true;
        });
      },
      `<Relationship Id="rIdPic" Type="${REL}image" Target="media/p.png"/>`,
      { 'word/media/p.png': RED },
    ));

    // 8. A picture in a format Word may not show is stored under its own type, not as .png.
    writeFileSync(join(out!, 'avif.docx'), await edit('<w:p><w:r><w:t>avif:</w:t></w:r></w:p>', (s) => {
      s.run((st, dispatch) => {
        const avif = btoa('\x00\x00\x00\x1cftypavif\x00\x00\x00\x00avifmif1miaf');
        dispatch?.(st.tr.insert(st.doc.content.size - 1, schema.nodes.image.create({ src: `data:image/avif;base64,${avif}`, width: 40, height: 40 })));
        return true;
      });
    }));
  });
});
