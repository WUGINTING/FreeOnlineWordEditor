import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { TextSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

async function setup() {
  const zip = blankPackage();
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host);
  await ed.open(await zip.generateAsync({ type: 'uint8array' }));
  const type = (text: string) => ed.run((s, d) => (d?.(s.tr.insertText(text)), true));
  const fields = () => {
    const found: string[] = [];
    ed.activeView?.state.doc.descendants((n) => { if (n.type.name === 'field') found.push(n.attrs.instr); });
    return found;
  };
  const save = async () => JSZip.loadAsync(await writeDocx(ed.view!.state.doc, ed.model));
  const done = () => { ed.destroy(); host.remove(); };
  return { ed, type, fields, save, done, host };
}

describe('header/footer office workflows', () => {
  it('QA20K-00931 inserts a live PAGE field from the footer toolbar and serializes it', async () => {
    const { ed, fields, save, done, host } = await setup();
    ed.editHeaderFooter('footer');
    host.querySelector<HTMLButtonElement>('button[title="在游標處插入目前頁碼"]')!.click();
    expect(fields()).toEqual(['PAGE']);
    const zip = await save();
    const footer = await zip.file('word/footer1.xml')!.async('string');
    expect(footer).toMatch(/w:fldSimple[^>]*w:instr=" PAGE "/);
    done();
  });

  it('QA20K-00932 inserts a total-page field without converting it to fixed text', async () => {
    const { ed, fields, save, done, host } = await setup();
    ed.editHeaderFooter('footer');
    host.querySelector<HTMLButtonElement>('button[title="在游標處插入總頁數"]')!.click();
    expect(fields()).toEqual(['NUMPAGES']);
    const zip = await save();
    expect(await zip.file('word/footer1.xml')!.async('string')).toMatch(/w:fldSimple[^>]*w:instr=" NUMPAGES "/);
    done();
  });

  it('QA20K-00933 keeps a human-readable Page label adjacent to the PAGE field', async () => {
    const { ed, type, fields, save, done } = await setup();
    ed.editHeaderFooter('footer'); type('頁次：'); ed.insertField('PAGE');
    expect(ed.activeView!.state.doc.textContent).toBe('頁次：');
    expect(fields()).toEqual(['PAGE']);
    const xml = await (await save()).file('word/footer1.xml')!.async('string');
    const paragraph = xml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ?? '';
    expect(paragraph).toContain('頁次：'); expect(paragraph).toContain('w:instr=" PAGE "');
    expect(paragraph.indexOf('頁次：')).toBeLessThan(paragraph.indexOf('w:instr=" PAGE "'));
    done();
  });

  it('QA20K-00934 keeps a header page field in the header story and body unchanged', async () => {
    const { ed, fields, save, done, host } = await setup();
    const bodyBefore = ed.view!.state.doc.toJSON();
    ed.editHeaderFooter('header');
    host.querySelector<HTMLButtonElement>('button[title="在游標處插入目前頁碼"]')!.click();
    expect(fields()).toEqual(['PAGE']);
    expect(ed.view!.state.doc.textContent).toBe(bodyBefore.content.map((n: any) => n.content?.map((c: any) => c.text ?? '').join('') ?? '').join(''));
    expect(ed.view!.state.doc.toJSON()).not.toMatchObject({ content: expect.arrayContaining([expect.objectContaining({ type: 'field' })]) });
    const zip = await save();
    expect(await zip.file('word/header1.xml')!.async('string')).toContain('w:instr=" PAGE "');
    expect(zip.file('word/footer1.xml')).toBeNull();
    done();
  });

  it('QA20K-00935 toggles different-first-page and creates a separate first footer part', async () => {
    const { ed, type, save, done, host } = await setup();
    ed.editHeaderFooter('footer');
    host.querySelector<HTMLInputElement>('.dx-hf-check input')!.click();
    expect(ed.sections()[0].titlePage).toBe(true);
    type('COVER ONLY');
    ed.closeHeaderFooter();
    const page = (ed as any).firstPage();
    (ed as any).pages = [{ ...page, number: 1, inSection: 0 }, { ...page, number: 2, inSection: 1, top: page.height }];
    ed.editHeaderFooter('footer', 1); type('BODY PAGES');
    const zip = await save();
    const sect = await zip.file('word/document.xml')!.async('string');
    expect(sect).toContain('w:titlePg');
    expect(sect).toContain('w:type="first"'); expect(sect).toContain('w:type="default"');
    const refTags = [...sect.matchAll(/<w:footerReference\b[^>]*\/?\s*>/g)].map((m) => m[0]);
    const attr = (tag: string, name: string) => new RegExp(`${name}="([^"]+)"`).exec(tag)?.[1];
    const firstRef = refTags.find((tag) => attr(tag, 'w:type') === 'first')!;
    const defaultRef = refTags.find((tag) => attr(tag, 'w:type') === 'default')!;
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const targetFor = (id: string) => {
      const relationship = [...rels.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map((m) => m[0]).find((tag) => attr(tag, 'Id') === id)!;
      return `word/${attr(relationship, 'Target')}`;
    };
    const firstTarget = targetFor(attr(firstRef, 'r:id')!);
    const defaultTarget = targetFor(attr(defaultRef, 'r:id')!);
    expect(firstTarget).not.toBe(defaultTarget);
    expect(await zip.file(firstTarget)!.async('string')).toContain('COVER ONLY');
    expect(await zip.file(defaultTarget)!.async('string')).toContain('BODY PAGES');
    done();
  });

  it('QA20K-00936 chooses the even-page header variant when even/odd mode and page 2 are active', async () => {
    const { ed, type, done } = await setup();
    ed.model.evenAndOdd = true;
    const first = (ed as any).firstPage();
    (ed as any).pages = [{ ...first, number: 1, inSection: 0 }, { ...first, number: 2, inSection: 1, top: first.height }];
    ed.editHeaderFooter('header', 1); type('EVEN RUNNING HEAD');
    expect(ed.model.headerFooters.some((h) => h.kind === 'header' && h.type === 'even' && h.doc.textContent === 'EVEN RUNNING HEAD')).toBe(true);
    expect(ed.model.headerFooters.some((h) => h.kind === 'header' && h.type === 'default')).toBe(false);
    done();
  });

  it('QA20K-00937 assigns a newly created footer only to the selected second section', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
    const host = document.createElement('div'); document.body.append(host); const ed = new DocxEditor(host);
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    const pg = (ed as any).firstPage(); (ed as any).pages = [{ ...pg, section: 0, inSection: 0 }, { ...pg, section: 1, inSection: 0, number: 2, top: pg.height }];
    ed.editHeaderFooter('footer', 1); ed.run((s, d) => (d?.(s.tr.insertText('SECTION TWO FOOTER')), true));
    const out = JSZip.loadAsync(await writeDocx(ed.view!.state.doc, ed.model)); const result = await out;
    const doc = await result.file('word/document.xml')!.async('string');
    expect((doc.match(/<w:footerReference/g) ?? []).length).toBe(1);
    const sections = [...doc.matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)].map((match) => match[0]);
    expect(sections).toHaveLength(2);
    expect(sections[0]).not.toContain('<w:footerReference');
    expect(sections[1]).toMatch(/<w:footerReference[^>]*w:type="default"/);
    expect(ed.model.headerFooters.filter((h) => h.kind === 'footer' && h.doc.textContent === 'SECTION TWO FOOTER')).toHaveLength(1);
    ed.destroy(); host.remove();
  });

  it('QA20K-00938 preserves independent header and footer text when switching editing surfaces', async () => {
    const { ed, type, save, done } = await setup();
    ed.editHeaderFooter('header'); type('ACME INTERNAL'); ed.closeHeaderFooter();
    ed.editHeaderFooter('footer'); type('Page '); ed.insertField('PAGE'); ed.closeHeaderFooter();
    ed.editHeaderFooter('header');
    expect(ed.activeView!.state.doc.textContent).toBe('ACME INTERNAL');
    const zip = await save();
    expect(await zip.file('word/header1.xml')!.async('string')).toContain('ACME INTERNAL');
    const footerXml = await zip.file('word/footer1.xml')!.async('string');
    const footerParagraph = footerXml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ?? '';
    expect(footerParagraph).toContain('Page ');
    expect(footerParagraph).toContain('w:instr=" PAGE "');
    expect(footerParagraph.indexOf('Page ')).toBeLessThan(footerParagraph.indexOf('w:instr=" PAGE "'));
    done();
  });

  it('QA20K-00939 undo removes only the inserted PAGE field and keeps the footer label', async () => {
    const { ed, type, fields, save, done } = await setup();
    ed.editHeaderFooter('footer'); type('Page '); ed.closeHeaderFooter();
    ed.editHeaderFooter('footer');
    const view = ed.activeView!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, view.state.doc.content.size - 1)));
    view.dispatch(closeHistory(view.state.tr));
    ed.insertField('PAGE');
    expect(fields()).toEqual(['PAGE']);
    ed.undo();
    expect(ed.activeView!.state.doc.textContent).toBe('Page ');
    expect(fields()).toEqual([]);
    expect(await (await save()).file('word/footer1.xml')!.async('string')).toContain('Page ');
    done();
  });

  it('QA20K-00940 opening then closing an empty header leaves no package part or dirty state', async () => {
    const { ed, done, save } = await setup();
    ed.editHeaderFooter('header'); expect(ed.snapshot()!.target).toBe('header'); ed.closeHeaderFooter();
    expect(ed.isModified()).toBe(false);
    expect(ed.model.headerFooters).toHaveLength(0);
    expect(Object.keys(await (await save()).files).some((name) => /word\/header\d+\.xml$/.test(name))).toBe(false);
    done();
  });
});
