// Writes documents edited with find/replace, the table panel and the picture tools, so Word
// itself can confirm the result (scripts/inspect-features-in-word.ps1). Opt-in:
//   DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-features.test.ts
import { describe, it } from 'vitest';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { EditorState, NodeSelection, TextSelection, type Command } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { history } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { replaceAll, searchPlugin, setSearch } from '../../src/papyrus/editor/search';
import {
  distributeColumns, setCantSplit, setCellBackground, setCellBorders, setCellVAlign, setHeaderRow, setRowHeight,
} from '../../src/papyrus/editor/tableCommands';
import { replaceImage, setImageAttrs } from '../../src/papyrus/editor/imageView';

const out = process.env.DOCX_EXPORT;
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
// 2x1 red PNG and 1x2 blue PNG.
const RED = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAADUlEQVR4nGP4z8AARAAI/gH/xp559wAAAABJRU5ErkJggg==';
const BLUE = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAACCAIAAAAW4yFwAAAADklEQVR4nGNgYPjPAMQABgIB/xARhfEAAAAASUVORK5CYII=';

async function edit(
  body: string,
  steps: (s: { state: EditorState; run: (c: Command) => void; find: (text: string) => number }) => void,
  media?: (zip: ReturnType<typeof blankPackage>) => Promise<void>,
): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${SECT}</w:body></w:document>`);
  await media?.(zip);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  const box = {
    state: EditorState.create({ schema, doc, plugins: [history(), searchPlugin()] }),
    run: (_c: Command) => {},
    find: (text: string) => {
      let found = -1;
      box.state.doc.descendants((n, pos) => {
        if (found < 0 && n.type.name === 'table_cell' && n.textContent === text) found = pos;
        return found < 0;
      });
      return found;
    },
  };
  box.run = (c) => void c(box.state, (tr) => (box.state = box.state.apply(tr)));
  steps(box);
  return writeDocx(box.state.doc, model);
}

const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const tableXml = (grid: number[], rows: string[][]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
  '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>' +
  `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
  rows.map((r, i) => `<w:tr>${i === 1 ? '<w:trPr><w:trHeight w:val="400"/></w:trPr>' : ''}${r.map(cell).join('')}</w:tr>`).join('') +
  '</w:tbl>';

const drawing = (id: number, rid: string, cx: number, cy: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:docPr id="${id}" name="Picture ${id}" descr="old"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:nvPicPr><pic:cNvPr id="${id}" name="p${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  '<a:ln w="38100"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:ln></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

describe.skipIf(!out)('export documents edited with the find / table / picture tools', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });

    // 1. Table appearance.
    writeFileSync(join(out!, 'table.docx'), await edit(
      tableXml([2800, 2800, 2800], [['A1', 'B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', 'B3', 'C3']]) + '<w:p/>' +
        tableXml([1400, 4200, 2800], [['x', 'y', 'z']]) + '<w:p/>',
      (s) => {
        const at = (t: string) => (s.state = s.state.apply(s.state.tr.setSelection(TextSelection.create(s.state.doc, s.find(t) + 2))));
        at('A1');
        s.run(setCellBackground('#ffff00'));
        s.run(setRowHeight(567, 'exact'));
        s.run(setHeaderRow(true));
        s.run(setCantSplit(true));
        at('C1');
        s.run(setCellVAlign('bottom'));
        at('B2');
        s.run(setCellBorders('none'));
        s.state = s.state.apply(s.state.tr.setSelection(new CellSelection(s.state.doc.resolve(s.find('A3')), s.state.doc.resolve(s.find('B3')))));
        s.run(setCellBorders('outer'));
        at('x');
        s.run(distributeColumns);
      },
    ));

    // 2. Pictures: resized + alt text, and replaced by another picture.
    writeFileSync(join(out!, 'image.docx'), await edit(
      `<w:p>${drawing(1, 'rIdA', 952500, 476250)}</w:p><w:p>${drawing(2, 'rIdB', 952500, 476250)}</w:p>`,
      (s) => {
        const images: number[] = [];
        s.state.doc.descendants((n, pos) => void (n.type.name === 'image' && images.push(pos)));
        const pick = (pos: number) => (s.state = s.state.apply(s.state.tr.setSelection(NodeSelection.create(s.state.doc, pos))));
        pick(images[0]);
        s.run(setImageAttrs({ width: 5 * (96 / 2.54), height: 2.5 * (96 / 2.54), alt: '紅色方塊' }));
        pick(images[1]);
        s.run(replaceImage(`data:image/png;base64,${BLUE}`, 1, 2));
      },
      async (zip) => {
        zip.file('word/media/a.png', RED, { base64: true });
        zip.file('word/media/b.png', RED, { base64: true });
        const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
        zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
          '<Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>' +
          '<Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/b.png"/></Relationships>'));
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        zip.file('[Content_Types].xml', ct.replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'));
      },
    ));

    // 3. Replace all across runs and inside a tracked insertion.
    writeFileSync(join(out!, 'replace.docx'), await edit(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Col</w:t></w:r><w:r><w:t xml:space="preserve">or and colour; </w:t></w:r>' +
        '<w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t>color</w:t></w:r></w:ins></w:p>',
      (s) => {
        s.run(setSearch('color'));
        s.run(replaceAll('hue'));
      },
    ));
  });
});
