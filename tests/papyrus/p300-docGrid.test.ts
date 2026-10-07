// persona-300: Word's line grid (w:docGrid w:type="lines") snaps the body's lines; the paragraphs of
// such a section get the grid class and pitch, and every line spacing carries the values the grid
// CSS needs. Measured rules in docGrid.ts; the layout itself is checked in the browser.
import { afterEach, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { domStubs, W } from './p300Helpers';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { linePitchOf } from '../../src/papyrus/editor/docGrid';
import { lineGridVars, paragraphStyle } from '../../src/papyrus/editor/schema';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';

domStubs();

const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const p = (t: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${t}</w:t></w:r></w:p>`;

async function docx(body: string, grid: string): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr>${PAGE}${grid}</w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

let editor: DocxEditor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
});

async function open(bytes: Uint8Array) {
  const host = document.createElement('div');
  document.body.append(host);
  editor = new DocxEditor(host);
  await editor.open(bytes);
  return host;
}

describe('the line grid', () => {
  it('reads the pitch only for a grid with lines', () => {
    expect(linePitchOf('<w:sectPr><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')).toBe(24);
    expect(linePitchOf('<w:sectPr><w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"/></w:sectPr>')).toBeCloseTo(20.8);
    // Word's "no grid" (what most generators write) and a character grid alone.
    expect(linePitchOf('<w:sectPr><w:docGrid w:linePitch="360"/></w:sectPr>')).toBeNull();
    expect(linePitchOf('<w:sectPr><w:docGrid w:type="snapToChars" w:linePitch="360"/></w:sectPr>')).toBeNull();
    expect(linePitchOf(null)).toBeNull();
  });

  it('marks the body paragraphs of a section with a line grid (not tables, not snapToGrid off)', async () => {
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${p('格')}</w:tc></w:tr></w:tbl>`;
    const host = await open(await docx(p('一') + p('不對齊', '<w:snapToGrid w:val="0"/>') + table + p('二'), '<w:docGrid w:type="lines" w:linePitch="360"/>'));
    const grid = [...host.querySelectorAll('.dx-grid')] as HTMLElement[];
    expect(grid.map((e) => e.textContent)).toEqual(['一', '二']);
    expect(grid[0].style.getPropertyValue('--dx-pitch')).toBe('24px');
  });

  it('marks nothing for a document without a line grid, and saves it unchanged', async () => {
    const bytes = await docx(p('一') + p('二'), '<w:docGrid w:linePitch="360"/>');
    const host = await open(bytes);
    expect(host.querySelectorAll('.dx-grid')).toHaveLength(0);
    const { doc, model } = await readDocx(bytes);
    const out = await JSZip.loadAsync(await writeDocx(doc, model));
    expect(await out.file('word/document.xml')!.async('string')).toContain('<w:docGrid w:linePitch="360"/>');
  });

  it('follows the section: a grid set in one section only', async () => {
    const first = `<w:sectPr>${PAGE}<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>`;
    const host = await open(await docx(p('甲') + p('乙', first) + p('丙'), '<w:docGrid w:linePitch="360"/>'));
    expect([...host.querySelectorAll('.dx-grid')].map((e) => e.textContent)).toEqual(['甲', '乙']);
  });

  it('typing does not work the sections out again; a new paragraph or section break does', async () => {
    const { EditorState } = await import('prosemirror-state');
    const { schema } = await import('../../src/papyrus/editor/schema');
    const { docGrid } = await import('../../src/papyrus/editor/docGrid');
    let asked = 0;
    const grid = '<w:sectPr><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>';
    const plugin = docGrid((d) => (asked++, [{ firstBlock: 0, lastBlock: d.childCount - 1, xml: grid }]));
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text('一')), schema.nodes.paragraph.create(null, schema.text('二'))]);
    let state = EditorState.create({ doc, plugins: [plugin] });
    const classes = () => plugin.getState(state)!.set.find().map((d) => (d as any).type.attrs.class);
    expect(classes()).toEqual(['dx-grid', 'dx-grid']);
    asked = 0;
    state = state.apply(state.tr.insertText('字', 2));
    state = state.apply(state.tr.insertText('字', 2));
    expect(asked).toBe(0);
    expect(classes()).toEqual(['dx-grid', 'dx-grid']);
    state = state.apply(state.tr.split(3));
    expect(asked).toBe(1);
    expect(classes()).toHaveLength(3);
    state = state.apply(state.tr.setNodeMarkup(0, undefined, { ...state.doc.child(0).attrs, pPr: '<w:pPr><w:snapToGrid w:val="0"/></w:pPr>' }));
    expect(asked).toBe(2);
    expect(classes()).toHaveLength(2);
  });

  it('every line spacing carries the grid values (auto, at least, exact)', () => {
    expect(lineGridVars(360, 'auto')).toEqual({ '--dx-lm': '1.5', '--dx-lmin': '0px', '--dx-lfix': 'initial' });
    expect(lineGridVars(300, 'atLeast')).toEqual({ '--dx-lm': '1', '--dx-lmin': '20px', '--dx-lfix': 'initial' });
    expect(lineGridVars(400, 'exact')).toEqual({ '--dx-lfix': `${400 / 15}px` });
    expect(lineGridVars(null, null)).toBeNull();
    expect(paragraphStyle({ line: 480, lineRule: 'auto' })).toContain('--dx-lm:2');
  });

  it('the CSS: whole grid lines, at least the "at least" height, an exact height as it is', () => {
    const css = readFileSync('src/papyrus/editor/docGrid.css', 'utf8');
    // Only in a browser with CSS round() (elsewhere the paragraph keeps its own line spacing), and
    // not for the text of a text box or shape in the paragraph.
    expect(css).toMatch(/@supports \(line-height: round\(up, 1px, 1px\)\) \{\s*\.dx-doc \.dx-p\.dx-grid,\s*\.dx-doc \.dx-p\.dx-grid :not\(\.dx-shape, \.dx-shape \*\) \{/);
    expect(css.match(/line-height:/g)).toHaveLength(2); // in the @supports test and the one rule
    expect(css).toMatch(/var\(\s*--dx-lfix,\s*max\(var\(--dx-lmin, 0px\), round\(up, calc\(var\(--dx-lh, 1.15\) \* 1em\), var\(--dx-pitch\)\), calc\(var\(--dx-lm, 1\) \* var\(--dx-pitch\)\)\)\s*\) !important/);
  });
});
