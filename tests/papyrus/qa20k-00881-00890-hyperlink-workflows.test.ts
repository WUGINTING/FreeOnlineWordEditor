import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import type { DocxEditor as DocxEditorType, EditorSnapshot } from '../../src/papyrus/editor/core';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const VUE = 'src/papyrus/vue/';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => vi.restoreAllMocks());

async function setup(text: string, extraBody = '') {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>${extraBody}<w:sectPr/></w:body></w:document>`);
  return setupBytes(await zip.generateAsync({ type: 'uint8array' }));
}

async function setupBytes(bytes: Uint8Array) {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const { model } = await readDocx(bytes);
  const host = document.createElement('div');
  const toolbarHost = document.createElement('div');
  document.body.append(host, toolbarHost);
  const snapshot = ref<EditorSnapshot | null>(null);
  const editor = new DocxEditor(host, { onUpdate: (s) => { snapshot.value = s; } });
  await editor.open(bytes);
  const Toolbar = await loadSfc(VUE + 'DocxToolbar.vue');
  const app = createApp({ render: () => h(Toolbar, { editor, snapshot: snapshot.value, styles: editor.model.paragraphStyles }) });
  app.mount(toolbarHost);
  const refresh = async () => { snapshot.value = editor.snapshot(); await nextTick(); };
  const done = () => { app.unmount(); editor.destroy(); host.remove(); toolbarHost.remove(); };
  return { editor, model, toolbar: toolbarHost, refresh, done };
}

function textStart(editor: DocxEditorType, text: string) {
  let start = -1;
  editor.view!.state.doc.descendants((node, pos) => {
    if (start < 0 && node.isText && node.text!.includes(text)) start = pos + node.text!.indexOf(text);
  });
  if (start < 0) throw new Error(`missing text: ${text}`);
  return start;
}

function select(editor: DocxEditorType, text: string, offset = 0, length = text.length) {
  const start = textStart(editor, text) + offset;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, start, start + length)));
  return start;
}

function caret(editor: DocxEditorType, text: string, offset: number) {
  const pos = textStart(editor, text) + offset;
  editor.view!.dispatch(editor.view!.state.tr.setSelection(TextSelection.create(editor.view!.state.doc, pos)));
  return pos;
}

function addLinkMark(editor: DocxEditorType, text: string, href: string, bold = false) {
  const from = textStart(editor, text);
  let tr = editor.view!.state.tr.addMark(from, from + text.length, schema.marks.link.create({ href }));
  if (bold) tr = tr.addMark(from, from + text.length, schema.marks.bold.create());
  editor.view!.dispatch(tr);
  editor.view!.dispatch(closeHistory(editor.view!.state.tr));
}

function linkRuns(editor: DocxEditorType) {
  const out: { text: string; href: string | null; marks: string[] }[] = [];
  editor.view!.state.doc.descendants((node) => {
    if (node.isText) out.push({
      text: node.text!,
      href: schema.marks.link.isInSet(node.marks)?.attrs.href ?? null,
      marks: node.marks.map((mark) => mark.type.name),
    });
  });
  return out;
}

function clickLink(toolbar: HTMLElement) {
  const button = toolbar.querySelector('button[title="插入連結"]') as HTMLButtonElement;
  if (!button) throw new Error('missing Insert Link toolbar button');
  button.click();
}

describe('QA20K-00881–00890 hyperlink editing workflows', () => {
  it('QA20K-00881 undo removes the latest link mark while preserving the selected wording', async () => {
    const d = await setup('Send the signed copy today'); select(d.editor, 'signed copy');
    vi.spyOn(window, 'prompt').mockReturnValue('https://files.example/signed'); clickLink(d.toolbar);
    expect(linkRuns(d.editor).some(x => x.href)).toBe(true);
    expect(d.editor.undo()).toBe(true);
    expect(d.editor.view!.state.doc.textContent).toBe('Send the signed copy today');
    expect(linkRuns(d.editor).every(x => x.href == null)).toBe(true); d.done();
  });
  it('QA20K-00882 redo restores the link target after undoing a link insertion', async () => {
    const d = await setup('Open meeting minutes'); select(d.editor, 'meeting minutes');
    vi.spyOn(window, 'prompt').mockReturnValue('https://intranet.example/minutes'); clickLink(d.toolbar);
    expect(d.editor.undo()).toBe(true); expect(linkRuns(d.editor).every(x => !x.href)).toBe(true);
    expect(d.editor.redo()).toBe(true);
    expect(linkRuns(d.editor).filter(x => x.href)).toEqual([{ text:'meeting minutes', href:'https://intranet.example/minutes', marks:['link'] }]); d.done();
  });
  it('QA20K-00883 applies a link across mixed bold and plain formatting without flattening marks', async () => {
    const d = await setup('Please read urgent policy today');
    const a=textStart(d.editor,'urgent policy'); d.editor.view!.dispatch(d.editor.view!.state.tr.addMark(a,a+6,schema.marks.bold.create()));
    const from=textStart(d.editor,'urgent'); d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(TextSelection.create(d.editor.view!.state.doc,from,from+'urgent policy'.length))); vi.spyOn(window,'prompt').mockReturnValue('https://hr.example/policy'); clickLink(d.toolbar);
    const rows=linkRuns(d.editor); expect(rows.filter(x=>x.href).map(x=>[x.text,x.marks])).toEqual([['urgent',['link','bold']],[' policy',['link']]]);
    expect(d.editor.view!.state.doc.textContent).toBe('Please read urgent policy today'); d.done();
  });
  it('QA20K-00884 links text between two existing links without changing either boundary target', async () => {
    const d=await setup('Read policy then appendix'); addLinkMark(d.editor,'policy','https://a.example'); addLinkMark(d.editor,'appendix','https://b.example');
    select(d.editor,'then'); vi.spyOn(window,'prompt').mockReturnValue('https://c.example'); clickLink(d.toolbar);
    expect(linkRuns(d.editor).filter(x=>x.href).map(x=>[x.text,x.href])).toEqual([['policy','https://a.example'],['then','https://c.example'],['appendix','https://b.example']]); d.done();
  });
  it('QA20K-00885 updates the target URL of a bold link while retaining its label and format', async () => {
    const d=await setup('See current handbook'); const a=textStart(d.editor,'current');
    d.editor.view!.dispatch(d.editor.view!.state.tr.addMark(a,a+7,schema.marks.bold.create()));
    addLinkMark(d.editor,'current','https://old.example'); select(d.editor,'current');
    vi.spyOn(window,'prompt').mockReturnValue('https://new.example'); clickLink(d.toolbar);
    expect(linkRuns(d.editor).filter(x=>x.href).map(x=>[x.text,x.href,x.marks])).toEqual([['current','https://new.example',['link','bold']]]); d.done();
  });
  it('QA20K-00886 writes two separate linked labels to separate DOCX relationships', async () => {
    const d=await setup('Policy and schedule'); addLinkMark(d.editor,'Policy','https://corp.example/policy'); addLinkMark(d.editor,'schedule','https://corp.example/calendar');
    const zip=await JSZip.loadAsync(await writeDocx(d.editor.view!.state.doc,d.model)); const xml=await zip.file('word/document.xml')!.async('string'); const rel=await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect((xml.match(/<w:hyperlink r:id=/g)||[]).length).toBe(2); expect(rel).toContain('Target="https://corp.example/policy"'); expect(rel).toContain('Target="https://corp.example/calendar"'); d.done();
  });
  it('QA20K-00887 imports an external DOCX hyperlink relationship and retains its destination', async () => {
    const zip=blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Review the </w:t></w:r><w:hyperlink r:id="rIdImportedLink"><w:r><w:t>policy</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>`);
    const rels=await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdImportedLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://word.example/policy" TargetMode="External"/></Relationships>'));
    const d=await setupBytes(await zip.generateAsync({type:'uint8array'}));
    expect(linkRuns(d.editor).find(x=>x.text==='policy')).toMatchObject({href:'https://word.example/policy'});
    const saved=await writeDocx(d.editor.view!.state.doc,d.model); const reopened=await readDocx(saved); const found:{text:string;href:string|null}[]=[];
    reopened.doc.descendants(n=>{if(n.isText)found.push({text:n.text!,href:schema.marks.link.isInSet(n.marks)?.attrs.href??null});});
    expect(found.find(x=>x.text==='policy')).toEqual({text:'policy',href:'https://word.example/policy'}); d.done();
  });
  it('QA20K-00888 preserves an internal Appendix bookmark target through DOCX save and reopen', async () => {
    const d=await setup('See Appendix for details', '<w:p><w:bookmarkStart w:id="12" w:name="Appendix"/><w:r><w:t>Appendix content</w:t></w:r><w:bookmarkEnd w:id="12"/></w:p>');
    select(d.editor,'Appendix'); vi.spyOn(window,'prompt').mockReturnValue('#Appendix'); clickLink(d.toolbar);
    const saved=await writeDocx(d.editor.view!.state.doc,d.model); const savedZip=await JSZip.loadAsync(saved); const xml=await savedZip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:bookmarkStart w:id="12" w:name="Appendix"/>');
    expect(xml).toContain('<w:bookmarkEnd w:id="12"/>');
    expect(xml).toContain('<w:hyperlink w:anchor="Appendix">');
    const reopened=await readDocx(saved); let href:string|null=null;
    reopened.doc.descendants(n=>{if(n.isText&&n.text==='Appendix')href=schema.marks.link.isInSet(n.marks)?.attrs.href??null;});
    expect(href).toBe('#Appendix'); d.done();
  });
  it('QA20K-00889 keeps adjacent plain text outside a link after save and reopen', async () => {
    const d=await setup('Prefix deliverable suffix'); addLinkMark(d.editor,'deliverable','https://work.example/file');
    const reopened=await readDocx(await writeDocx(d.editor.view!.state.doc,d.model)); const rows:{text:string;href:string|null}[]=[];
    reopened.doc.descendants(n=>{if(n.isText)rows.push({text:n.text!,href:schema.marks.link.isInSet(n.marks)?.attrs.href??null});});
    expect(rows).toEqual([{text:'Prefix ',href:null},{text:'deliverable',href:'https://work.example/file'},{text:' suffix',href:null}]); d.done();
  });
  it('QA20K-00890 changing a link destination is a single undoable document operation', async () => {
    const d=await setup('Open handbook'); addLinkMark(d.editor,'handbook','https://old.example'); caret(d.editor,'handbook',2);
    vi.spyOn(window,'prompt').mockReturnValue('https://new.example'); clickLink(d.toolbar);
    expect(linkRuns(d.editor).find(x=>x.text==='handbook')?.href).toBe('https://new.example');
    expect(d.editor.undo()).toBe(true); expect(linkRuns(d.editor).find(x=>x.text==='handbook')?.href).toBe('https://old.example'); d.done();
  });
});


