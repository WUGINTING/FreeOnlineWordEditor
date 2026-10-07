import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { history } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import { acceptAllRevisions, acceptRevision, collectRevisions, goToRevision, rejectAllRevisions, rejectRevision, review } from '../../src/papyrus/editor/review';

const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const A=(id:number)=>`w:id="${id}" w:author="Reviewer" w:date="2026-09-25T02:00:00Z"`;
const ins=(id:number,xml:string)=>`<w:ins ${A(id)}>${xml}</w:ins>`;
const del=(id:number,txt:string)=>`<w:del ${A(id)}><w:r><w:delText>${txt}</w:delText></w:r></w:del>`;
const BR='<w:r><w:br w:type="page"/></w:r>';
async function fixture(body:string){
 const zip=blankPackage(); zip.file('word/document.xml',`<w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
 const {doc,model}=await readDocx(await zip.generateAsync({type:'uint8array'}));
 let state=EditorState.create({schema,doc,plugins:[history(),review()]});
 const run=(cmd:Command)=>cmd(state,tr=>state=state.apply(tr));
 const pos=(needle:string)=>{let p=-1;state.doc.descendants((n,at)=>{if(p<0&&n.isText&&n.text!.includes(needle))p=at+n.text!.indexOf(needle)});if(p<0)throw Error('missing '+needle);return p};
 const xml=async()=>(await JSZip.loadAsync(await writeDocx(state.doc,model))).file('word/document.xml')!.async('string');
 const paras=async()=>{const b=/<w:body>([\s\S]*)<w:sectPr/.exec(await xml())![1];return b.split(/(?=<w:p[ >]|<w:tbl>)/)};
 return {run,pos,xml,paras,get state(){return state},select:(from:number,to=from)=>state=state.apply(state.tr.setSelection(TextSelection.create(state.doc,from,to)))};
}

describe('QA20K-00901–00908 tracked layout review',()=>{
 it('QA20K-00901 finalizing mixed break and text revisions preserves the accepted page boundary',async()=>{
  const d=await fixture(`<w:p><w:r><w:t xml:space="preserve">Start </w:t></w:r>${ins(1,BR)}${del(2,'obsolete')}<w:r><w:t> End</w:t></w:r></w:p>`);
  d.run(acceptAllRevisions); expect(collectRevisions(d.state.doc)).toHaveLength(0);
  expect((await d.paras())[0]).toContain('<w:br w:type="page"/>'); expect((await d.paras())[0]).not.toContain('<w:ins'); expect(d.state.doc.textContent).toBe('Start  End');
 });
 it('QA20K-00902 rejects an inserted page break without joining separate paragraphs',async()=>{
  const d=await fixture(`<w:p><w:r><w:t>A</w:t></w:r>${ins(1,BR)}<w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>Next</w:t></w:r></w:p>`);
  d.run(rejectAllRevisions);
  expect(await d.paras()).toEqual(['<w:p><w:r><w:t>AB</w:t></w:r></w:p>','<w:p><w:r><w:t>Next</w:t></w:r></w:p>']);
 });
 it('QA20K-00903 accepts a deleted page break while preserving adjacent live text',async()=>{
  const d=await fixture(`<w:p><w:r><w:t>Left</w:t></w:r><w:del ${A(1)}>${BR}</w:del><w:r><w:t>Right</w:t></w:r></w:p>`);
  d.run(goToRevision(1)); d.run(acceptRevision);
  expect((await d.paras())[0]).toBe('<w:p><w:r><w:t>LeftRight</w:t></w:r></w:p>');
  expect(collectRevisions(d.state.doc)).toHaveLength(0);
 });
 it('QA20K-00904 rejects a deleted page break and restores the explicit page boundary',async()=>{
  const d=await fixture(`<w:p><w:r><w:t>Left</w:t></w:r><w:del ${A(1)}>${BR}</w:del><w:r><w:t>Right</w:t></w:r></w:p>`);
  d.run(rejectAllRevisions);
  const sequence:string[]=[];
  d.state.doc.descendants(n=>{
   if(n.type.name==='page_break'||(n.type.name==='raw_inline'&&/<w:br\b[^>]*w:type="page"/.test(n.attrs.xml??'')))sequence.push('<PAGE_BREAK>');
   else if(n.isText&&n.text)sequence.push(n.text);
  });
  const resolvedParagraph=(await d.paras())[0];
  const xmlSequence=[...resolvedParagraph.matchAll(/<w:t(?: [^>]*)?>(.*?)<\/w:t>|<w:br w:type="page"\/>/g)].map(m=>m[0].startsWith('<w:br')?'<PAGE_BREAK>':m[1]);
  expect({model:sequence,serialized:xmlSequence}).toEqual({model:['Left','<PAGE_BREAK>','Right'],serialized:['Left','<PAGE_BREAK>','Right']});
  expect(d.state.doc.textContent).toBe('LeftRight');
  expect(collectRevisions(d.state.doc)).toHaveLength(0);
 });
 it('QA20K-00905 accepts a tracked page break inside an insertion without splitting its paragraph',async()=>{
  const d=await fixture(`<w:p><w:r><w:t>L</w:t></w:r>${ins(1,`<w:r><w:t xml:space="preserve">new </w:t></w:r>${BR}<w:r><w:t>page</w:t></w:r>`)}<w:r><w:t>R</w:t></w:r></w:p>`);
  d.run(acceptAllRevisions); const ps=await d.paras();
  expect(ps).toHaveLength(1); expect(ps[0].match(/<w:p>/g)).toHaveLength(1); expect(ps[0]).toContain('<w:br w:type="page"/>');
  expect(d.state.doc.textContent).toBe('Lnew pageR');
  expect(collectRevisions(d.state.doc)).toHaveLength(0);
 });
 it('QA20K-00906 rejects tracked text and its page break while retaining surrounding words',async()=>{
  const d=await fixture(`<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${ins(1,`<w:r><w:t>draft</w:t></w:r>${BR}`)}<w:r><w:t>After</w:t></w:r></w:p>`);
  d.run(rejectAllRevisions); expect(d.state.doc.textContent).toBe('Before After');
  expect((await d.paras())[0]).not.toContain('<w:br'); expect(collectRevisions(d.state.doc)).toHaveLength(0);
 });
 it('QA20K-00907 accepts a deleted break-only paragraph while preserving the live page break and next section',async()=>{
  const d=await fixture(`<w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:del ${A(1)}/></w:rPr></w:pPr>${del(2,'x')}${BR}</w:p><w:p><w:r><w:t>Landscape section starts</w:t></w:r></w:p>`);
  d.run(acceptAllRevisions); expect(d.state.doc.textContent).toContain('Landscape section starts');
  const ps=await d.paras();
  expect(ps).toEqual(['<w:p><w:r><w:t>Intro</w:t></w:r></w:p>','<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Landscape section starts</w:t></w:r></w:p>']);
  expect(ps.filter(x=>x.includes('<w:br w:type="page"/>'))).toHaveLength(1);
  expect(collectRevisions(d.state.doc)).toHaveLength(0);
 });
 it('QA20K-00908 saves resolved layout revisions with no revision wrappers or break markup for rejected content',async()=>{
  const d=await fixture(`<w:p><w:r><w:t xml:space="preserve">Approved </w:t></w:r>${ins(1,`<w:r><w:t>extra</w:t></w:r>${BR}`)}<w:r><w:t> text</w:t></w:r></w:p>`);
  d.run(rejectAllRevisions); const xml=await d.xml();
  expect(xml).not.toContain('<w:ins'); expect(xml).not.toContain('<w:br w:type="page"/>'); expect(xml).toContain('Approved '); expect(xml).toContain(' text');
 });
});



