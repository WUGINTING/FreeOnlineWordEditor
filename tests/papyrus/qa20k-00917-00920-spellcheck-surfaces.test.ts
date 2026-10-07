import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(()=>{
 (globalThis as any).ResizeObserver??=class{observe(){}unobserve(){}disconnect(){}};
 const zero=()=>({x:0,y:0,top:0,left:0,right:0,bottom:0,width:0,height:0,toJSON(){}}) as DOMRect;
 Range.prototype.getClientRects??=()=>Object.assign([],{item:()=>null}) as unknown as DOMRectList;
 Range.prototype.getBoundingClientRect??=zero;
});
const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
async function open(){
 const zip=blankPackage();zip.file('word/document.xml',`<w:document ${W}><w:body><w:p><w:r><w:t>Quarterly report body</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
 const host=document.createElement('div');document.body.append(host);
 const editor=new DocxEditor(host);await editor.open(await zip.generateAsync({type:'uint8array'}));
 return {host,editor,done(){editor.destroy();host.remove()}};
}
afterEach(()=>{});
describe('QA20K-00917–00920 spellcheck editing-surface propagation',()=>{
 it('QA20K-00917 an opted-in spellcheck state is applied to a header opened after the body',async()=>{
  const x=await open();try{x.editor.setSpellcheck(true);x.editor.editHeaderFooter('header');expect(x.editor.view!.dom.getAttribute('spellcheck')).toBe('true');expect(x.editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');}finally{x.done()}
 });
 it('QA20K-00918 changing spellcheck while editing a header updates both live editor surfaces',async()=>{
  const x=await open();try{x.editor.editHeaderFooter('header');expect(x.editor.activeView!.dom.getAttribute('spellcheck')).toBe('false');x.editor.setSpellcheck(true);expect(x.editor.view!.dom.getAttribute('spellcheck')).toBe('true');expect(x.editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');}finally{x.done()}
 });
 it('QA20K-00919 turning spellcheck off from an active footer updates both views without changing document text',async()=>{
  const x=await open();try{x.editor.setSpellcheck(true);x.editor.editHeaderFooter('footer');const body=x.editor.view!.state.doc.textContent;x.editor.setSpellcheck(false);expect(x.editor.activeView!.dom.getAttribute('spellcheck')).toBe('false');expect(x.editor.view!.dom.getAttribute('spellcheck')).toBe('false');expect(x.editor.view!.state.doc.textContent).toBe(body);}finally{x.done()}
 });
 it('QA20K-00920 closing and reopening the same header session keeps the current spellcheck DOM setting',async()=>{
  const x=await open();try{x.editor.editHeaderFooter('header');x.editor.setSpellcheck(true);x.editor.closeHeaderFooter(false);expect(x.editor.view!.dom.getAttribute('spellcheck')).toBe('true');x.editor.editHeaderFooter('header');expect(x.editor.activeView!.dom.getAttribute('spellcheck')).toBe('true');}finally{x.done()}
 });
});
