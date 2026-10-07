import { beforeAll, describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { createApp, h as vueH, nextTick } from 'vue';
import { blankPackage } from '../src/papyrus/docx/template';
import { readDocx } from '../src/papyrus/docx/reader';
import { DocxEditor } from '../src/papyrus/editor/core';
import type { DocxEditor as DocxEditorType } from '../src/papyrus/editor/core';
import DocxEditorVue from '../src/papyrus/vue/DocxEditor.vue';
import { documentOutline, pageReferences, updatePageReferences } from '../src/papyrus/editor/outline';
import { countText } from '../src/papyrus/editor/wordCount';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const p = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const h = (text: string, name = 'Heading1') => `<w:p><w:pPr><w:pStyle w:val="${name}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const field = (bookmark: string, value: string) => `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${bookmark} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${value}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
async function read(body: string, styles = '') {
  return readDocx(await packageDoc(body, styles));
}
async function packageDoc(body: string, styles = '') {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const styleXml=styles||'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';
  zip.file('word/styles.xml', `<w:styles ${W}>${styleXml}</w:styles>`);
  return zip.generateAsync({ type: 'uint8array' });
}
function host() { const el = document.createElement('div'); document.body.append(el); return el; }
const flush = async () => { for (let i=0;i<10;i++) { await new Promise(r=>setTimeout(r,0)); await nextTick(); } };
async function mountPage(src: Uint8Array) {
  const el=host(); let editor:DocxEditorType|null=null;
  const app=createApp({render:()=>vueH(DocxEditorVue,{src,onReady:(e:DocxEditorType)=>editor=e})});app.mount(el);
  for(let i=0;i<50&&!editor;i++) await flush(); await flush();
  const nav=()=>el.querySelector('.dx-outline') as HTMLElement|null;
  const openNav=async()=>{(Array.from(el.querySelectorAll('button')).find(b=>b.textContent?.trim()===String.fromCharCode(0x5c0e,0x89bd)) as HTMLButtonElement).click();await flush();};
  return {el,app,editor:editor!,nav,openNav,done:()=>{app.unmount();el.remove();}};
}

describe('QA20K document navigation and insights 00771–00780', () => {
  it('QA20K-00771 refreshes the rendered navigation item after a heading is renamed', async () => {
    const ui=await mountPage(await packageDoc(h('Draft title')+p('Body')));await ui.openNav();
    const item=()=>Array.from(ui.nav()!.querySelectorAll('li button')).map(b=>b.textContent?.trim());
    expect(item()).toEqual(['Draft title']);
    const ed=ui.editor;const heading=ed.outline()[0];
    ed.view!.dispatch(ed.view!.state.tr.insertText('Approved title',heading.pos+1,heading.pos+1+heading.text.length));await flush();
    expect(item()).toEqual(['Approved title']);ui.done();
  });

  it('QA20K-00772 clicking the later repeated heading selects its paragraph in the editor', async () => {
    const ui=await mountPage(await packageDoc(h('Summary')+p('middle')+h('Summary')));await ui.openNav();
    const buttons=ui.nav()!.querySelectorAll('li button');expect(Array.from(buttons).map(b=>b.textContent?.trim())).toEqual(['Summary','Summary']);
    (buttons[1] as HTMLButtonElement).click();await flush();
    const heading=ui.editor.outline()[1];expect(ui.editor.view!.state.selection.$from.before()).toBe(heading.pos);
    ui.done();
  });

  it('QA20K-00773 renders a heading inside a table as an item in the navigation pane', async () => {
    const ui=await mountPage(await packageDoc(h('Before')+`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${h('Table decision')}</w:tc></w:tr></w:tbl>`+h('After')));await ui.openNav();
    expect(Array.from(ui.nav()!.querySelectorAll('li button')).map(b=>b.textContent?.trim())).toEqual(['Before','Table decision','After']);ui.done();
  });

  it('QA20K-00774 lets a direct paragraph outline level override its built-in heading level', async () => {
    const { doc, model } = await read(`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="3"/></w:pPr><w:r><w:t>Appendix</w:t></w:r></w:p>`, `<w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>`);
    expect(documentOutline(doc, model.styles).map(x=>[x.level,x.text])).toEqual([[3,'Appendix']]);
  });

  it('QA20K-00775 omits level-nine body text from the rendered navigation list', async () => {
    const body=`<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Body-like</w:t></w:r></w:p><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Top heading</w:t></w:r></w:p>`;
    const ui=await mountPage(await packageDoc(body));await ui.openNav();
    expect(Array.from(ui.nav()!.querySelectorAll('li button')).map(b=>b.textContent?.trim())).toEqual(['Top heading']);
    expect(ui.nav()!.textContent).not.toContain('Body-like');ui.done();
  });

  it('QA20K-00776 updates every TOC entry that points to the same bookmark', async () => {
    const { doc } = await read(field('SecA','4') + field('SecA','4') + `<w:p><w:bookmarkStart w:id="3" w:name="SecA"/><w:r><w:t>Policy</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>`);
    const tr=EditorState.create({doc}).tr;
    expect(pageReferences(tr.doc).map(x=>x.bookmark)).toEqual(['SecA','SecA']);
    expect(updatePageReferences(tr,()=>18)).toBe(2);
    expect(pageReferences(tr.doc).map(x=>tr.doc.textBetween(x.from,x.to))).toEqual(['18','18']);
  });

  it('QA20K-00777 calculates page locations against the unchanged layout before replacing longer earlier numbers', async () => {
    const { doc } = await read(field('A','1') + field('B','2') + `<w:p><w:bookmarkStart w:id="1" w:name="A"/><w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:p><w:bookmarkStart w:id="2" w:name="B"/><w:r><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>`);
    const tr=EditorState.create({doc}).tr; const positions:string[]=[];
    expect(updatePageReferences(tr,pos=>{ positions.push(tr.doc.resolve(pos).parent.textContent); return positions.at(-1)==='First'?123:45; })).toBe(2);
    expect(positions).toEqual(['First','Second']);
    expect(pageReferences(tr.doc).map(x=>tr.doc.textBetween(x.from,x.to))).toEqual(['123','45']);
  });

  it('QA20K-00778 clicking refresh for current page numbers creates no undo history entry', async () => {
    const body=field('Current','12')+`<w:p><w:bookmarkStart w:id="1" w:name="Current"/><w:r><w:t>Current chapter</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const ui=await mountPage(await packageDoc(body));await ui.openNav();
    (ui.editor as any).pageNumberAt=()=>12;
    const ed=ui.editor;const before=ed.view!.state.doc.textContent;
    ed.view!.dispatch(ed.view!.state.tr.insertText(' temporary',ed.view!.state.doc.content.size-2));await flush();
    (ui.nav()!.querySelector('.dx-outline-update') as HTMLButtonElement).click();await flush();
    expect(ui.el.querySelector('.dx-notice')?.textContent?.trim().length).toBeGreaterThan(0);
    expect(ed.undo()).toBe(true);await flush();
    expect(ed.view!.state.doc.textContent).toBe(before);ui.done();
  });

  it('QA20K-00779 treats tabs and nonbreaking spaces as word boundaries but counts their characters', () => {
    expect(countText('weekly\tstatus\u00a0report')).toEqual({words:3,chars:18,charsWithSpaces:20});
  });

  it('QA20K-00780 counts a cross-paragraph selection using a boundary between its words', async () => {
    const z=blankPackage(); z.file('word/document.xml',`<w:document ${W}><w:body>${p('Quarterly')} ${p('forecast')}<w:sectPr/></w:body></w:document>`);
    const el=host(); const ed=new DocxEditor(el); await ed.open(await z.generateAsync({type:'uint8array'}));
    const text=ed.view!.state.doc.textContent;
    const start=1;
    const end=ed.view!.state.doc.content.size-1;
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,start,end)));
    expect(ed.view!.state.doc.textBetween(start,end,'\n')).toContain('Quarterly');
    expect(ed.selectionWordCount()?.words).toBe(2);
    ed.destroy(); el.remove();
  });
});

