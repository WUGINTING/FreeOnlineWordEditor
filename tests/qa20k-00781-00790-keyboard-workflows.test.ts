import { beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor } from '../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const p=(text:string,pPr='')=>`<w:p>${pPr?`<w:pPr>${pPr}</w:pPr>`:''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const boldRun=(text:string)=>`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
async function open(body:string){
  const zip=blankPackage();zip.file('word/document.xml',`<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const host=document.createElement('div');document.body.append(host);const ed=new DocxEditor(host);await ed.open(await zip.generateAsync({type:'uint8array'}));return {ed,host};
}
function press(ed:DocxEditor,key:string,opts:{ctrl?:boolean;shift?:boolean}={}){
  const e=new KeyboardEvent('keydown',{key,bubbles:true,cancelable:true,ctrlKey:opts.ctrl,shiftKey:opts.shift});ed.view!.dom.dispatchEvent(e);return e;
}
function paragraphs(ed:DocxEditor){const out:any[]=[];ed.view!.state.doc.descendants((n,pos)=>{if(n.type.name==='paragraph')out.push({node:n,pos,text:n.textContent});});return out;}
function textRuns(ed:DocxEditor){const out:{from:number;text:string;marks:string[]}[]=[];ed.view!.state.doc.descendants((n,pos)=>{if(n.isText)out.push({from:pos,text:n.text!,marks:n.marks.map(m=>m.type.name)});});return out;}
function dispose(ed:DocxEditor,host:HTMLElement){ed.destroy();host.remove();}

describe('QA20K office keyboard editing workflows 00781–00790',()=>{
  it('QA20K-00781 applies Ctrl+B to selected words without formatting adjacent text',async()=>{
    const {ed,host}=await open(p('Draft approved'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,7,15)));
    const event=press(ed,'b',{ctrl:true});const nodes:any[]=[];ed.view!.state.doc.descendants(n=>{if(n.isText)nodes.push({text:n.text,marks:n.marks.map(m=>m.type.name)});});
    expect(event.defaultPrevented).toBe(true);expect(nodes).toEqual([{text:'Draft ',marks:[]},{text:'approved',marks:['bold']}]);dispose(ed,host);
  });
  it('QA20K-00782 uses Ctrl+Space to clear character formatting on only the selection',async()=>{
    const {ed,host}=await open(`<w:p>${boldRun('Ref ')}<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>AB-19</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> keep</w:t></w:r></w:p>`);
    const start=ed.view!.state.doc.textContent.indexOf('AB-19')+1;ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,start,start+5)));
    const event=press(ed,' ',{ctrl:true});
    const selected=textRuns(ed).filter(r=>r.from<start+5&&r.from+r.text.length>start).flatMap(r=>Array.from(r.text,(ch,i)=>({ch,pos:r.from+i,marks:r.marks})) ).filter(c=>c.pos>=start&&c.pos<start+5);
    expect(event.defaultPrevented).toBe(true);expect(selected.map(x=>x.ch).join('')).toBe('AB-19');expect(selected.every(x=>!x.marks.includes('bold')&&!x.marks.includes('italic'))).toBe(true);
    const neighbors=textRuns(ed);expect(neighbors.find(r=>r.text.includes('Ref'))?.marks).toContain('bold');expect(neighbors.find(r=>r.text.includes('keep'))?.marks).toContain('italic');dispose(ed,host);
  });
  it('QA20K-00783 Enter in the middle of a sentence splits text and places the caret in the new paragraph',async()=>{
    const {ed,host}=await open(p('Submit report'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,8)));
    const event=press(ed,'Enter');expect(event.defaultPrevented).toBe(true);expect(paragraphs(ed).map(x=>x.text)).toEqual(['Submit ','report']);expect(ed.view!.state.selection.$from.parent.textContent).toBe('report');expect(ed.view!.state.selection.$from.parentOffset).toBe(0);dispose(ed,host);
  });
  it('QA20K-00784 Enter at the end of a heading starts a normal body paragraph',async()=>{
    const {ed,host}=await open(p('Decision','<w:pStyle w:val="Heading1"/>'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,9)));
    press(ed,'Enter');const ps=paragraphs(ed);expect(ps.map(x=>x.text)).toEqual(['Decision','']);expect(ps[0].node.attrs.styleId).toBe('Heading1');expect(ps[1].node.attrs.styleId).toBeNull();dispose(ed,host);
  });
  it('QA20K-00785 Enter on an empty numbered item exits the list without adding another item',async()=>{
    const {ed,host}=await open(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>`);const before=paragraphs(ed)[0];
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,before.pos+1)));const event=press(ed,'Enter');
    const after=paragraphs(ed);expect(event.defaultPrevented).toBe(true);expect(after).toHaveLength(1);expect(after[0].node.attrs.numId).toBeNull();dispose(ed,host);
  });
  it('QA20K-00786 Shift+Enter inserts a line break inside one paragraph',async()=>{
    const {ed,host}=await open(p('Name Department'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,5)));
    const event=press(ed,'Enter',{shift:true});const ps=paragraphs(ed);expect(event.defaultPrevented).toBe(true);expect(ps).toHaveLength(1);expect(ps[0].node.childCount).toBe(3);expect(ps[0].node.child(1).type.name).toBe('hard_break');dispose(ed,host);
  });
  it('QA20K-00787 Tab in a table advances from the first cell to the next cell',async()=>{
    const table='<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const {ed,host}=await open(table);let pos=-1;ed.view!.state.doc.descendants((n,p)=>{if(pos<0&&n.isText&&n.text!.includes('A1'))pos=p+n.text!.indexOf('A1');});ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,pos)));
    const event=press(ed,'Tab');expect(event.defaultPrevented).toBe(true);expect(ed.view!.state.selection.$from.parent.textContent).toBe('B1');dispose(ed,host);
  });
  it('QA20K-00788 Tab in the final table cell creates a row and focuses its first cell',async()=>{
    const table='<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Last cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const {ed,host}=await open(table);let lastCellEnd=-1;let tablePos=-1;let tableNode:any;
    ed.view!.state.doc.descendants((n,pos)=>{if(n.isText&&n.text==='Last cell')lastCellEnd=pos+n.text.length;if(tablePos<0&&n.type.name==='table'){tablePos=pos;tableNode=n;}});
    if(lastCellEnd<0||tablePos<0)throw new Error('missing final-cell text or table');
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,lastCellEnd)));
    const event=press(ed,'Tab');const doc=ed.view!.state.doc;let rows=0;let cells=0;doc.descendants(n=>{if(n.type.name==='table_row')rows++;if(n.type.name==='table_cell')cells++;});
    expect(event.defaultPrevented).toBe(true);expect(rows).toBe(2);expect(cells).toBe(2);
    const current=ed.view!.state.selection.$from;const cellDepth=[...Array(current.depth+1).keys()].reverse().find(depth=>current.node(depth).type.name==='table_cell');
    expect(cellDepth).toBeDefined();
    expect(current.before(cellDepth!)).toBe(tablePos+1+tableNode.child(0).nodeSize+1);
    expect(current.parent.textContent).toBe('');dispose(ed,host);
  });
  it('QA20K-00789 Tab outside a table inserts a tab object without splitting the paragraph',async()=>{
    const {ed,host}=await open(p('Code owner'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,5)));
    const event=press(ed,'Tab');const ps=paragraphs(ed);expect(event.defaultPrevented).toBe(true);expect(ps).toHaveLength(1);expect(ps[0].node.childCount).toBe(3);expect(ps[0].node.child(1).type.name).toBe('tab');dispose(ed,host);
  });
  it('QA20K-00790 Ctrl+Enter inserts a page break within the current paragraph',async()=>{
    const {ed,host}=await open(p('Before after'));ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc,8)));
    const event=press(ed,'Enter',{ctrl:true});const ps=paragraphs(ed);let breaks=0;const sequence:string[]=[];
    ed.view!.state.doc.descendants(n=>{if(n.type.name==='page_break'){breaks++;sequence.push('<PAGE_BREAK>');}else if(n.isText&&n.text)sequence.push(n.text);});
    expect(event.defaultPrevented).toBe(true);expect(breaks).toBe(1);expect(ps.map(x=>x.text)).toEqual(['Before ','after']);
    expect(sequence).toEqual(['Before ','<PAGE_BREAK>','after']);
    expect(ps[0].node.attrs.brGroup).toBeTruthy();expect(ps[0].node.attrs.brGroup).toBe(ps[1].node.attrs.brGroup);dispose(ed,host);
  });
});
