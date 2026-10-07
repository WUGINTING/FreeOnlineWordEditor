import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { clearFormatting, setMark, setParagraphStyle, toggle, toggleFormat } from '../../src/papyrus/editor/commands';

const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const STYLES=`<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ItalicNote"><w:name w:val="Italic Note"/><w:rPr><w:i/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Banner"><w:name w:val="Banner"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>`;
const para=(runs:string,pPr='')=>`<w:p>${pPr?`<w:pPr>${pPr}</w:pPr>`:''}${runs}</w:p>`;
const run=(text:string,rPr='')=>`<w:r>${rPr?`<w:rPr>${rPr}</w:rPr>`:''}<w:t>${text}</w:t></w:r>`;
function children(node:Element,name:string):Element[]{return Array.from(node.children).filter(child=>child.localName===name);}
function paragraphs(doc:Document):Element[]{return Array.from(doc.getElementsByTagName('w:body')[0].children).filter(node=>node.localName==='p');}
function paragraphRuns(p:Element):Element[]{return children(p,'r');}
function runText(r:Element):string{return children(r,'t').map(t=>t.textContent??'').join('');}
function runProperty(r:Element,name:string):Element|undefined{
 const rPr=children(r,'rPr')[0];return rPr?children(rPr,name)[0]:undefined;
}
function paragraphStyle(p:Element):string|null{
 const pPr=children(p,'pPr')[0],style=pPr&&children(pPr,'pStyle')[0];return style?.getAttribute('w:val')??null;
}
function parsed(xml:string):Document{
 const doc=new DOMParser().parseFromString(xml,'application/xml');
 expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);return doc;
}
async function open(body:string,styles=STYLES){
 const zip=blankPackage();zip.file('word/document.xml',`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);zip.file('word/styles.xml',styles);
 const {doc,model}=await readDocx(await zip.generateAsync({type:'uint8array'}));let state=EditorState.create({schema,doc});
 const span=(text:string,start=0,length=text.length)=>{let pos=-1;state.doc.descendants((n,p)=>{if(pos<0&&n.isText&&n.text?.includes(text))pos=p+n.text.indexOf(text)});if(pos<0)throw new Error(`missing text ${text}`);state=state.apply(state.tr.setSelection(TextSelection.create(state.doc,pos+start,pos+start+length)));};
 const runCmd=(cmd:Command)=>expect(cmd(state,tr=>{state=state.apply(tr)})).toBe(true);
 const save=async()=> (await JSZip.loadAsync(await writeDocx(state.doc,model))).file('word/document.xml')!.async('string');
 return {span,runCmd,save,get state(){return state;},model};
}

describe('QA20K everyday character and paragraph formatting 00741–00750',()=>{
 it('QA20K-00741 italicizes only the selected phrase',async()=>{
  const d=await open(para(run('Before target after')));d.span('target');d.runCmd(toggleFormat('italic',d.model.styles));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]);
  const target=runs.find(r=>runText(r)==='target')!,before=runs.find(r=>runText(r)==='Before ')!,after=runs.find(r=>runText(r)===' after')!;
  expect(runProperty(target,'i')).toBeTruthy();expect(runProperty(before,'i')).toBeUndefined();expect(runProperty(after,'i')).toBeUndefined();
 });
 it('QA20K-00742 turns off inherited italic on a selected exception',async()=>{
  const source=await open(para(run('Need review'),'<w:pStyle w:val="ItalicNote"/>'),'<w:styles '+W+'><w:style w:type="paragraph" w:styleId="ItalicNote"><w:rPr><w:i/></w:rPr></w:style></w:styles>');
  source.span('review');source.runCmd(toggleFormat('italic',source.model.styles));const xml=await source.save();
  const p=paragraphs(parsed(xml))[0],runs=paragraphRuns(p),review=runs.find(r=>runText(r)==='review')!,need=runs.find(r=>runText(r)==='Need ' )!;
  expect(paragraphStyle(p)).toBe('ItalicNote');expect(runProperty(review,'i')?.getAttribute('w:val')).toBe('0');expect(runProperty(need,'i')).toBeUndefined();expect(source.state.doc.textContent).toBe('Need review');
 });
 it('QA20K-00743 changes the selected deadline text to 18 point without resizing neighbors',async()=>{
  const d=await open(para(run('Due 30 days')));d.span('30 days');d.runCmd(setMark(schema.marks.fontSize,{pt:18}));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),deadline=runs.find(r=>runText(r)==='30 days')!,label=runs.find(r=>runText(r)==='Due ')!;
  expect(runProperty(deadline,'sz')?.getAttribute('w:val')).toBe('36');expect(runProperty(label,'sz')).toBeUndefined();
 });
 it('QA20K-00744 colors only the selected account code blue',async()=>{
  const d=await open(para(run('Account AC-42 verified')));d.span('AC-42');d.runCmd(setMark(schema.marks.color,{color:'#0070C0'}));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),code=runs.find(r=>runText(r)==='AC-42')!,left=runs.find(r=>runText(r)==='Account ')!,right=runs.find(r=>runText(r)===' verified')!;
  expect(runProperty(code,'color')?.getAttribute('w:val')).toBe('0070C0');expect(runProperty(left,'color')).toBeUndefined();expect(runProperty(right,'color')).toBeUndefined();
 });
 it('QA20K-00745 highlights only the follow-up date in yellow',async()=>{
  const d=await open(para(run('Follow up 2026-10-01 please')));d.span('2026-10-01');d.runCmd(setMark(schema.marks.highlight,{color:'#FFFF00'}));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),date=runs.find(r=>runText(r)==='2026-10-01')!,before=runs.find(r=>runText(r)==='Follow up ')!,after=runs.find(r=>runText(r)===' please')!;
  expect(runProperty(date,'highlight')?.getAttribute('w:val')).toBe('yellow');expect(runProperty(before,'highlight')).toBeUndefined();expect(runProperty(after,'highlight')).toBeUndefined();
 });
 it('QA20K-00746 applies separate Latin and East Asian font names to a selected bilingual phrase',async()=>{
  const d=await open(para(run('簽核 Sign')));d.span('簽核 Sign');d.runCmd(setMark(schema.marks.font,{family:'Arial',eastAsia:'標楷體'}));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),phrase=runs.find(r=>runText(r)==='簽核 Sign')!,fonts=runProperty(phrase,'rFonts');
  expect(fonts?.getAttribute('w:ascii')).toBe('Arial');expect(fonts?.getAttribute('w:hAnsi')).toBe('Arial');expect(fonts?.getAttribute('w:eastAsia')).toBe('標楷體');
 });
 it('QA20K-00747 underlines only the selected policy reference',async()=>{
  const d=await open(para(run('See policy 4.2 today')));d.span('4.2');d.runCmd(toggle('underline'));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),target=runs.find(r=>runText(r)==='4.2')!,left=runs.find(r=>runText(r)==='See policy ')!,right=runs.find(r=>runText(r)===' today')!;
  expect(runProperty(target,'u')?.getAttribute('w:val')).toBe('single');expect(runProperty(left,'u')).toBeUndefined();expect(runProperty(right,'u')).toBeUndefined();
 });
 it('QA20K-00748 strikes through only the superseded unit price',async()=>{
  const d=await open(para(run('Old 125 New 130')));d.span('125');d.runCmd(toggle('strike'));const xml=await d.save();
  const runs=paragraphRuns(paragraphs(parsed(xml))[0]),target=runs.find(r=>runText(r)==='125')!,left=runs.find(r=>runText(r)==='Old ')!,right=runs.find(r=>runText(r)===' New 130')!;
  expect(runProperty(target,'strike')).toBeTruthy();expect(runProperty(left,'strike')).toBeUndefined();expect(runProperty(right,'strike')).toBeUndefined();
 });
 it('QA20K-00749 clears appearance only on the selected word and keeps paragraph style and surrounding formats',async()=>{
  const body=para(run('Keep ', '<w:b/>')+run('reset', '<w:b/><w:color w:val="FF0000"/><w:sz w:val="28"/>')+run(' this', '<w:i/>'),'<w:pStyle w:val="Banner"/>');
  const d=await open(body);d.span('reset');d.runCmd(clearFormatting);const xml=await d.save();
  const p=paragraphs(parsed(xml))[0],runs=paragraphRuns(p),keep=runs.find(r=>runText(r)==='Keep ')!,reset=runs.find(r=>runText(r)==='reset')!,tail=runs.find(r=>runText(r)===' this')!;
  expect(paragraphStyle(p)).toBe('Banner');expect(runProperty(keep,'b')).toBeTruthy();expect(runProperty(tail,'i')).toBeTruthy();
  expect(runProperty(reset,'b')).toBeUndefined();expect(runProperty(reset,'color')).toBeUndefined();expect(runProperty(reset,'sz')).toBeUndefined();
 });
 it('QA20K-00750 applies a paragraph style to both selected paragraphs but not the following memo paragraph',async()=>{
  const d=await open(para(run('Objective'))+para(run('Decision'))+para(run('Appendix')));
  const end=d.state.doc.child(0).nodeSize+d.state.doc.child(1).nodeSize-1;d.span('Objective');
  const from=1,to=end;const selected=d.state.apply(d.state.tr.setSelection(TextSelection.create(d.state.doc,from,to)));
  let state=selected;expect(setParagraphStyle('Banner')(state,tr=>{state=state.apply(tr)})).toBe(true);
  const xml=await (await JSZip.loadAsync(await writeDocx(state.doc,d.model))).file('word/document.xml')!.async('string');
  const output=paragraphs(parsed(xml));
  const objective=output.find(p=>paragraphRuns(p).map(runText).join('')==='Objective')!,decision=output.find(p=>paragraphRuns(p).map(runText).join('')==='Decision')!,appendix=output.find(p=>paragraphRuns(p).map(runText).join('')==='Appendix')!;
  expect(paragraphStyle(objective)).toBe('Banner');expect(paragraphStyle(decision)).toBe('Banner');expect(paragraphStyle(appendix)).not.toBe('Banner');
 });
});
