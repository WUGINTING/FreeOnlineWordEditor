// Hardening of the DOCX reader / writer against hostile or damaged files and odd editor
// content (security + stability audits, 2026-09-25), and a guard on the cost of a save.
// Every fix keeps untouched XML byte for byte.
import { describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { TableMap } from 'prosemirror-tables';
import { EditorState } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx, type WriteWarning } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { parseXml, parseStats } from '../../src/papyrus/docx/xml';
import { readComments } from '../../src/papyrus/docx/comments';
import { ListCounter, ensureList, formatNumber, parseNumbering } from '../../src/papyrus/docx/numbering';
import { documentSections } from '../../src/papyrus/docx/sections';
import { readRunModel } from '../../src/papyrus/docx/props';

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const WR = W + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const NSX =
  W +
  ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"' +
  ' xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

function documentXml(body: string, sect = SECT): string {
  return `${HEAD}<w:document ${WR}><w:body>${body}${sect}</w:body></w:document>`;
}

async function pkg(body: string, edit?: (zip: JSZip) => Promise<void> | void, sect = SECT): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', documentXml(body, sect));
  await edit?.(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function addRel(zip: JSZip, id: string, type: string, target: string): Promise<void> {
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/></Relationships>`));
}

async function part(bytes: Uint8Array, name: string): Promise<string | null> {
  const f = (await JSZip.loadAsync(bytes)).file(name);
  return f ? f.async('string') : null;
}

const marksOf = (doc: any, name: string) => {
  const out: any[] = [];
  doc.descendants((n: any) => {
    for (const m of n.marks) if (m.type.name === name) out.push(m.attrs);
    return true;
  });
  return out;
};

// ---------------------------------------------------------------------------
describe('security: w:highlight / w:color values reaching CSS', () => {
  const run = (rPr: string) => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>x</w:t></w:r></w:p>`;

  it('only named highlight colours (own properties) or hex colours become a highlight', async () => {
    const bad = ['url(https://attacker.example/beacon)', '__proto__', 'constructor', 'toString', 'hasOwnProperty', 'red;x:y'];
    for (const val of bad) {
      const { doc } = await readDocx(await pkg(run(`<w:highlight w:val="${val}"/>`)));
      expect(marksOf(doc, 'highlight'), val).toEqual([]);
    }
    const named = await readDocx(await pkg(run('<w:highlight w:val="yellow"/>')));
    expect(marksOf(named.doc, 'highlight')).toEqual([{ color: '#ffff00' }]);
    const hex = await readDocx(await pkg(run('<w:highlight w:val="00FF00"/>')));
    expect(marksOf(hex.doc, 'highlight')).toEqual([{ color: '#00ff00' }]);
  });

  it('run colours must be hex colours too', async () => {
    const { doc } = await readDocx(await pkg(run('<w:color w:val="000;background:url(https://attacker.example/x)"/>')));
    expect(marksOf(doc, 'color')).toEqual([]);
    const ok = await readDocx(await pkg(run('<w:color w:val="1F4E79"/>')));
    expect(marksOf(ok.doc, 'color')).toEqual([{ color: '#1F4E79' }]);
    expect(readRunModel(parseXml(`<w:rPr ${W}><w:highlight w:val="__proto__"/></w:rPr>`).documentElement).highlight).toBeNull();
  });

  it('the rejected value is still saved as written when the run is untouched', async () => {
    const rPr = '<w:rPr><w:color w:val="url(x)"/><w:highlight w:val="url(https://attacker.example/beacon)"/></w:rPr>';
    const first = await readDocx(await pkg(`<w:p><w:r>${rPr}<w:t>x</w:t></w:r></w:p>`));
    const xml = await part(await writeDocx(first.doc, first.model), 'word/document.xml');
    expect(xml).toContain(rPr);
  });
});

// ---------------------------------------------------------------------------
describe('security: page size and margins used for layout are clamped', () => {
  const sect = (pgSz: string, pgMar: string) => `<w:sectPr><w:pgSz ${pgSz}/><w:pgMar ${pgMar}/></w:sectPr>`;

  it('a negative or huge page size and impossible margins stay in Word’s range', async () => {
    const bytes = await pkg(
      '<w:p/>',
      undefined,
      sect('w:w="99999999" w:h="-500"', 'w:top="-20000" w:right="-5" w:bottom="20000" w:left="40000" w:header="-3" w:footer="99999999" w:gutter="0"'),
    );
    const { doc, model } = await readDocx(bytes);
    const p = model.page;
    expect(p.width).toBe(31680);
    expect(p.height).toBe(144);
    for (const k of ['marginTop', 'marginBottom', 'marginLeft', 'marginRight', 'header', 'footer'] as const) {
      expect(p[k], k).toBeGreaterThanOrEqual(0);
    }
    expect(p.marginTop + p.marginBottom).toBeLessThan(p.height);
    expect(p.marginLeft + p.marginRight).toBeLessThan(p.width);
    expect(documentSections(doc, model)[0].page.height).toBe(144);

    // The original w:sectPr is saved as it was.
    const xml = await part(await writeDocx(doc, model), 'word/document.xml');
    expect(xml).toContain('<w:pgSz w:w="99999999" w:h="-500"/>');
    expect(xml).toContain('w:top="-20000" w:right="-5" w:bottom="20000" w:left="40000"');
  });

  it('section breaks are clamped too, normal pages are not touched', async () => {
    const brk = `<w:p><w:pPr>${sect('w:w="0" w:h="16838"', 'w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"')}</w:pPr></w:p>`;
    const { doc, model } = await readDocx(await pkg(brk + '<w:p/>'));
    const sections = documentSections(doc, model);
    expect(sections[0].page.width).toBe(144);
    expect(sections[0].page.marginLeft + sections[0].page.marginRight).toBeLessThan(144);
    expect(sections[1].page).toEqual({ width: 11906, height: 16838, marginTop: 1440, marginBottom: 1440, marginLeft: 1800, marginRight: 1800, header: 851, footer: 992 });
  });
});

// ---------------------------------------------------------------------------
describe('security: characters XML forbids never reach a saved part', () => {
  const NASTY = 'a\u0000b\u0001c\u000Bd\u000Ce\u001Ff\uFFFEg\uFFFFh\uD800i\uDC00j😀k';

  it('typed / pasted control characters are stripped; VT is a line break and FF a page break', async () => {
    const first = await readDocx(await pkg('<w:p/>'));
    const para = schema.nodes.paragraph.create(null, [
      schema.text(NASTY),
      schema.nodes.field.create({ instr: 'PAGE', text: '1\u0002' }),
      schema.nodes.image.create({ src: 'data:image/png;base64,iVBORw0KGgo=', alt: 'alt\u0003text', width: 10, height: 10 }),
    ]);
    const doc = schema.nodes.doc.create(null, [para]);
    const saved = await writeDocx(doc, first.model);
    const xml = (await part(saved, 'word/document.xml'))!;
    expect(xml).toContain('<w:t>abc</w:t><w:br/><w:t>d</w:t><w:br w:type="page"/><w:t>efghij😀k</w:t>');
    expect(xml).toContain('descr="alttext"');
    // eslint-disable-next-line no-control-regex
    expect(xml).not.toMatch(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/);
    // The saved file opens again.
    const again = await readDocx(saved);
    expect(again.doc.textContent).toContain('abc');
    expect(again.doc.textContent).toContain('efghij😀k');
  });

  it('comments with control characters stay readable', async () => {
    const first = await readDocx(await pkg('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
    const comment = { id: '0', author: 'A\u0001nn', initials: 'A', date: '2026-01-01T00:00:00Z', dateUtc: null, text: 'one\u0000two\u000Bthree\u000C', parentId: null, done: false, created: true };
    const doc = schema.nodes.doc.create({ ...first.doc.attrs, comments: [comment] }, first.doc.content);
    const saved = await writeDocx(doc, first.model);
    const zip = await JSZip.loadAsync(saved);
    for (const name of Object.keys(zip.files).filter((n) => n.endsWith('.xml'))) {
      const text = await zip.file(name)!.async('string');
      expect(() => parseXml(text), name).not.toThrow();
    }
    const back = await readComments(zip);
    expect(back[0].author).toBe('Ann');
    expect(back[0].text.replace(/\n/g, '|')).toBe('onetwo|three');
  });
});

// ---------------------------------------------------------------------------
describe('security: paraIds from the file are escaped when written back', () => {
  const EVIL = 'AB&quot; w15:evil=&quot;1';
  const COMMENTS =
    `${HEAD}<w:comments ${NSX}><w:comment w:id="0" w:author="Ann" w:date="2026-01-02T10:00:00Z" w:initials="A">` +
    `<w:p w14:paraId="${EVIL}" w14:textId="77777777"><w:r><w:annotationRef/></w:r><w:r><w:t>First</w:t></w:r></w:p></w:comment></w:comments>`;
  const EXTENDED = `${HEAD}<w15:commentsEx ${NSX}><w15:commentEx w15:paraId="${EVIL}" w15:done="0"/></w15:commentsEx>`;
  const IDS = `${HEAD}<w16cid:commentsIds ${NSX}></w16cid:commentsIds>`;

  it('commentsExtended.xml and commentsIds.xml stay well-formed', async () => {
    const bytes = await pkg(
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Alpha</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>',
      async (zip) => {
        zip.file('word/comments.xml', COMMENTS);
        zip.file('word/commentsExtended.xml', EXTENDED);
        zip.file('word/commentsIds.xml', IDS);
        await addRel(zip, 'rIdC1', 'comments', 'comments.xml');
        const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
        zip.file(
          'word/_rels/document.xml.rels',
          rels.replace(
            '</Relationships>',
            '<Relationship Id="rIdC2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' +
              '<Relationship Id="rIdC3" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/></Relationships>',
          ),
        );
      },
    );
    const first = await readDocx(bytes);
    const comments = await readComments(first.model.zip);
    expect(comments.length).toBe(1);
    const doc = schema.nodes.doc.create({ ...first.doc.attrs, comments: [{ ...comments[0], done: true }] }, first.doc.content);
    const saved = await writeDocx(doc, first.model);
    for (const name of ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml']) {
      const text = (await part(saved, name))!;
      expect(() => parseXml(text), name).not.toThrow();
      // The id stays one attribute value: no attribute was smuggled in.
      const attrs = Array.from(parseXml(text).getElementsByTagName('*')).flatMap((e) => Array.from(e.attributes).map((a) => a.localName.toLowerCase()));
      expect(attrs, name).not.toContain('evil');
    }
    expect((await readComments(await JSZip.loadAsync(saved)))[0].done).toBe(true);
  });
});

// ---------------------------------------------------------------------------
describe('stability: table grid values', () => {
  const tbl = (rows: string, grid = '<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>') =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
  const tc = (tcPr: string, text = 'x') => `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const cells = (doc: any) => {
    const out: any[] = [];
    doc.descendants((n: any) => {
      if (n.type.name === 'table_cell') out.push(n.attrs);
      return true;
    });
    return out;
  };

  it('gridSpan 0 / negative / fractional in a vertically merged cell opens (was an endless loop)', async () => {
    const rows =
      `<w:tr>${tc('')}${tc('<w:gridSpan w:val="0"/><w:vMerge w:val="restart"/>')}</w:tr>` +
      `<w:tr>${tc('')}${tc('<w:vMerge/>', '')}</w:tr>` +
      `<w:tr>${tc('<w:gridSpan w:val="-3"/>')}${tc('<w:gridSpan w:val="1.5"/>')}</w:tr>`;
    const bytes = await pkg(tbl(rows));
    const { doc, model } = await readDocx(bytes);
    for (const c of cells(doc)) expect(c.colspan).toBeGreaterThanOrEqual(1);
    expect(doc.textContent).toBe('xxxxx');
    // Untouched: the table is saved as written.
    const xml = (await part(await writeDocx(doc, model), 'word/document.xml'))!;
    expect(xml).toContain('<w:tcPr><w:gridSpan w:val="0"/><w:vMerge w:val="restart"/></w:tcPr>');
    expect(xml).toContain('<w:tcPr><w:gridSpan w:val="-3"/></w:tcPr>');
    expect(xml).toContain('<w:tcPr><w:gridSpan w:val="1.5"/></w:tcPr>');
  });

  it('gridSpan 0 without a merge opens as a one-column cell', async () => {
    const { doc, model } = await readDocx(await pkg(tbl(`<w:tr>${tc('<w:gridSpan w:val="0"/>')}${tc('')}</w:tr>`)));
    expect(cells(doc).map((c) => c.colspan)).toEqual([1, 1]);
    expect(doc.textContent).toBe('xx'); // both cells (the first one was lost)
    await expect(writeDocx(doc, model)).resolves.toBeInstanceOf(Uint8Array);
  });

  it('gridSpan and gridBefore are capped at 63 columns', async () => {
    const rows = `<w:tr><w:trPr><w:gridBefore w:val="5000"/></w:trPr>${tc('<w:gridSpan w:val="100000"/>')}</w:tr>`;
    const { doc, model } = await readDocx(await pkg(tbl(rows)));
    const all = cells(doc);
    expect(Math.max(...all.map((c) => c.colspan))).toBeLessThanOrEqual(63);
    let width = 0;
    doc.descendants((n: any) => {
      if (n.type.name === 'table') width = TableMap.get(n).width;
      return true;
    });
    expect(width).toBeLessThanOrEqual(126);
    const xml = (await part(await writeDocx(doc, model), 'word/document.xml'))!;
    expect(xml).toContain('<w:gridSpan w:val="100000"/>');
  });

  it('the writer does not look cells up once per grid slot', async () => {
    const row = `<w:tr>${Array.from({ length: 20 }, () => tc('')).join('')}</w:tr>`;
    const grid = Array.from({ length: 20 }, () => '<w:gridCol w:w="400"/>').join('');
    const { doc, model } = await readDocx(await pkg(tbl(Array.from({ length: 30 }, () => row).join(''), grid)));
    const spy = vi.spyOn(TableMap.prototype, 'findCell');
    await writeDocx(doc, model);
    // At most one lookup per cell (600 cells): not one per slot per column of the cell.
    expect(spy.mock.calls.length).toBeLessThanOrEqual(600);
    spy.mockRestore();
  });
});

// ---------------------------------------------------------------------------
describe('stability: list numbers', () => {
  const numbering = (lvl: string, override = '') =>
    `${HEAD}<w:numbering ${W}><w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/>${override}</w:num></w:numbering>`;
  const lvl = (ilvl: string, start: string, fmt: string) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${start}"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/></w:lvl>`;

  it('formatNumber never throws and stays short', () => {
    const fmts = ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'decimalZero', 'ideographTraditional', 'ideographZodiac', 'chineseCounting', 'none', 'bullet'];
    for (const fmt of fmts) {
      for (const n of [-1e12, -1e9, -30, -27, -26, -1, 0, 0.5, 1, 26, 27, 32767, 1e9, 1e12, NaN, Infinity, -Infinity]) {
        let s = '';
        expect(() => (s = formatNumber(n, fmt)), `${fmt} ${n}`).not.toThrow();
        expect(typeof s).toBe('string');
        expect(s.length, `${fmt} ${n}`).toBeLessThanOrEqual(1300);
      }
    }
    // Ordinary values are unchanged.
    expect(formatNumber(28, 'upperLetter')).toBe('BB');
    expect(formatNumber(14, 'lowerRoman')).toBe('xiv');
    expect(formatNumber(0, 'decimal')).toBe('0');
  });

  it('a bad start value or override does not break list markers', () => {
    const doc = parseXml(numbering(lvl('0', '-30', 'lowerLetter') + lvl('1', '99999999999', 'upperRoman'), '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="-1000"/></w:lvlOverride>'));
    const counter = new ListCounter(parseNumbering(doc));
    expect(() => counter.next('1', 0)).not.toThrow();
    expect(() => counter.next('1', 1)).not.toThrow();
    expect(counter.next('1', 1)!.length).toBeLessThan(100);
  });

  it('a huge w:ilvl is clamped to 0..8 for layout and counting, and saved as written', async () => {
    const huge = '1000000000';
    const bytes = await pkg(
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${huge}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`,
      async (zip) => {
        zip.file('word/numbering.xml', numbering(lvl('0', '1', 'decimal') + lvl(huge, '1', 'decimal')));
        await addRel(zip, 'rIdN', 'numbering', 'numbering.xml');
      },
    );
    const { doc, model } = await readDocx(bytes);
    expect(doc.firstChild!.attrs.ilvl).toBe(8);
    expect(model.numbering.abstracts['0'].levels.length).toBeLessThanOrEqual(9);
    const counter = new ListCounter(model.numbering);
    const t = performance.now();
    counter.next('1', 1e9);
    counter.next('1', 0);
    counter.next('1', 1e9);
    expect(performance.now() - t).toBeLessThan(1000);
    const xml = (await part(await writeDocx(doc, model), 'word/document.xml'))!;
    expect(xml).toContain(`<w:ilvl w:val="${huge}"/>`);
    // numbering.xml keeps the level as written too.
    expect(await part(await writeDocx(doc, model), 'word/numbering.xml')).toContain(`w:ilvl="${huge}"`);
  });
});

// ---------------------------------------------------------------------------
describe('stability: one broken optional part', () => {
  const BROKEN = '<w:oops ' + W + '><unclosed></w:oops>';

  it.each(['word/styles.xml', 'word/settings.xml'])('%s that is not XML: the file opens and the part is kept', async (name) => {
    const bytes = await pkg('<w:p><w:r><w:t>hello</w:t></w:r></w:p>', (zip) => void zip.file(name, BROKEN));
    const { doc, model } = await readDocx(bytes);
    expect(doc.textContent).toBe('hello');
    expect(model.brokenParts).toEqual([name]);
    // Even with track changes switched on (which edits settings.xml).
    model.trackRevisions = true;
    const warnings: WriteWarning[] = [];
    const saved = await writeDocx(doc, model, { warnings });
    expect(await part(saved, name)).toBe(BROKEN);
    if (name.includes('settings')) expect(warnings.length).toBe(1);
  });

  it('a broken numbering.xml is kept; a list made in the editor is reported, not saved into it', async () => {
    const bytes = await pkg('<w:p><w:r><w:t>hello</w:t></w:r></w:p>', async (zip) => {
      zip.file('word/numbering.xml', BROKEN);
      await addRel(zip, 'rIdN', 'numbering', 'numbering.xml');
    });
    const { doc, model } = await readDocx(bytes);
    expect(model.brokenParts).toEqual(['word/numbering.xml']);
    const numId = ensureList(model.numbering, 'decimal');
    const listed = schema.nodes.doc.create(doc.attrs, [schema.nodes.paragraph.create({ numId, ilvl: 0 }, schema.text('item'))]);
    const warnings: WriteWarning[] = [];
    const saved = await writeDocx(listed, model, { warnings });
    expect(await part(saved, 'word/numbering.xml')).toBe(BROKEN);
    expect(warnings.map((w) => w.part)).toContain('word/numbering.xml');
  });

  it('a broken header and a broken .rels file are skipped and kept', async () => {
    const bytes = await pkg(
      '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      async (zip) => {
        zip.file('word/header1.xml', BROKEN);
        zip.file('word/footer1.xml', `${HEAD}<w:ftr ${WR}><w:p><w:r><w:t>foot</w:t></w:r></w:p></w:ftr>`);
        zip.file('word/_rels/footer1.xml.rels', 'not xml at all <');
        await addRel(zip, 'rIdH', 'header', 'header1.xml');
        await addRel(zip, 'rIdF', 'footer', 'footer1.xml');
      },
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>',
    );
    const { doc, model } = await readDocx(bytes);
    expect(doc.textContent).toBe('body');
    expect([...model.brokenParts!].sort()).toEqual(['word/_rels/footer1.xml.rels', 'word/header1.xml']);
    expect(model.headerFooters.map((h) => h.part)).toEqual(['word/footer1.xml']);
    for (const hf of model.headerFooters) hf.dirty = true;
    const saved = await writeDocx(doc, model);
    expect(await part(saved, 'word/header1.xml')).toBe(BROKEN);
    expect(await part(saved, 'word/_rels/footer1.xml.rels')).toBe('not xml at all <');
    expect(await part(saved, 'word/footer1.xml')).toContain('foot');
  });

  it('a broken document.xml.rels opens the body; a broken document.xml is still an error', async () => {
    const bytes = await pkg('<w:p><w:r><w:t>body</w:t></w:r></w:p>', (zip) => void zip.file('word/_rels/document.xml.rels', 'garbage <'));
    const { doc, model } = await readDocx(bytes);
    expect(doc.textContent).toBe('body');
    expect(model.brokenParts).toEqual(['word/_rels/document.xml.rels']);
    expect(await part(await writeDocx(doc, model), 'word/_rels/document.xml.rels')).toBe('garbage <');

    const zip = blankPackage();
    zip.file('word/document.xml', '<w:document ' + W + '><w:body>');
    await expect(readDocx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow();
  });
});

// ---------------------------------------------------------------------------
describe('stability: pictures the writer cannot decode', () => {
  async function blank() {
    return readDocx(await pkg('<w:p/>'));
  }
  const withImage = (attrs: Record<string, unknown>) =>
    schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text('a'), schema.nodes.image.create(attrs)])]);

  it('an undecodable data: URL is skipped with a warning instead of failing the save', async () => {
    const { model } = await blank();
    const warnings: WriteWarning[] = [];
    const saved = await writeDocx(withImage({ src: 'data:image/png;base64,@@@not base64@@@', width: 10, height: 10 }), model, { warnings });
    expect(warnings.length).toBe(1);
    expect(warnings[0].kind).toBe('image');
    expect((await part(saved, 'word/document.xml'))!).not.toContain('<w:drawing>');
    const bad = await writeDocx(withImage({ src: 'data:image/png,%E0%A4%A', width: 10, height: 10 }), model);
    expect(bad).toBeInstanceOf(Uint8Array);
  });

  it('a non-base64 (utf8 / percent-encoded) SVG data: URL is embedded', async () => {
    const { model } = await blank();
    const svg = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="%23f00"/></svg>';
    const saved = await writeDocx(withImage({ src: svg, width: 10, height: 10 }), model);
    const zip = await JSZip.loadAsync(saved);
    const media = Object.keys(zip.files).filter((n) => n.startsWith('word/media/') && !zip.files[n].dir);
    expect(media).toEqual(['word/media/px-image1.svg']);
    expect(await zip.file(media[0])!.async('string')).toContain('fill="#f00"');
  });

  it('a replaced picture whose new data is broken keeps the original picture', async () => {
    const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    const drawing =
      '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="P"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
    const bytes = await pkg(`<w:p><w:r>${drawing}</w:r></w:p>`, async (zip) => {
      zip.file('word/media/image1.png', PNG, { base64: true });
      await addRel(zip, 'rIdImg', 'image', 'media/image1.png');
    });
    const { doc, model } = await readDocx(bytes);
    let pos = -1;
    doc.descendants((n, p) => {
      if (n.type.name === 'image') pos = p;
      return true;
    });
    expect(pos).toBeGreaterThanOrEqual(0);
    const img = doc.nodeAt(pos)!;
    const state = EditorState.create({ schema, doc });
    const edited = state.apply(state.tr.setNodeMarkup(pos, undefined, { ...img.attrs, src: 'data:image/png;base64,***' })).doc;
    const warnings: WriteWarning[] = [];
    const saved = await writeDocx(edited, model, { warnings });
    expect(warnings.length).toBe(1);
    const xml = (await part(saved, 'word/document.xml'))!;
    expect(xml).toContain('r:embed="rIdImg"');
    expect(await part(saved, 'word/_rels/document.xml.rels')).toContain('media/image1.png');
    expect((await JSZip.loadAsync(saved)).file('word/media/image1.png')).not.toBeNull();
  });
});

// ---------------------------------------------------------------------------
describe('stability: a package without [Content_Types].xml', () => {
  it('gets the template content types plus its own parts', async () => {
    const zip = blankPackage();
    zip.remove('[Content_Types].xml');
    const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const ct = (await part(await writeDocx(doc, model), '[Content_Types].xml'))!;
    expect(ct).toContain('Extension="rels"');
    expect(ct).toContain('Extension="xml"');
    expect(ct).toContain('PartName="/word/document.xml"');
    expect(ct).toContain('PartName="/word/styles.xml"');
  });
});

// ---------------------------------------------------------------------------
describe('performance guard: XML fragments parsed per save', () => {
  it('does not grow with the number of runs, paragraphs or cells that share properties', async () => {
    const para = '<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="標楷體"/><w:sz w:val="24"/></w:rPr><w:t>text</w:t></w:r></w:p>';
    const cell = `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${para}</w:tc>`;
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${`<w:tr>${cell}${cell}</w:tr>`.repeat(50)}</w:tbl>`;
    const count = async (n: number) => {
      const { doc, model } = await readDocx(await pkg(para.repeat(n) + table));
      // Bold everything and indent every paragraph, as a user would: every run's and paragraph's properties change.
      const state = EditorState.create({ schema, doc });
      const tr = state.tr.addMark(0, doc.content.size, schema.marks.bold.create());
      tr.doc.descendants((node, pos) => {
        if (node.type.name === 'paragraph') tr.setNodeMarkup(pos, undefined, { ...node.attrs, indLeft: 480 });
        return true;
      });
      const before = parseStats.fragments;
      await writeDocx(tr.doc, model);
      return parseStats.fragments - before;
    };
    await count(1); // warm the caches of other tests out of the picture
    const small = await count(20);
    const large = await count(400);
    expect(large).toBeLessThanOrEqual(small + 5);
    expect(large).toBeLessThan(50);
  });
});
