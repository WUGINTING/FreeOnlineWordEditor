import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow } from 'prosemirror-tables';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';

const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const cell=(text:string,width=2000,extra='')=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p>${text?`<w:r><w:t>${text}</w:t></w:r>`:''}</w:p></w:tc>`;
const row=(cells:string[])=>`<w:tr>${cells.join('')}</w:tr>`;
const table=(rows:string[])=>`<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join('')}</w:tbl>`;
const base=()=>table([
  row([cell('A1'),cell('B1'),cell('C1')]),
  row([cell('A2'),cell('B2'),cell('C2')]),
  row([cell('A3'),cell('B3'),cell('C3')]),
]);
async function open(xml=base()){
  const zip=blankPackage();
  zip.file('word/document.xml',`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${xml}<w:sectPr/></w:body></w:document>`);
  const {doc,model}=await readDocx(await zip.generateAsync({type:'uint8array'}));
  let state=EditorState.create({schema,doc});
  const at=(text:string)=>{
    let pos=-1;
    state.doc.descendants((n,p)=>{if(pos<0&&n.isText&&n.text===text)pos=p;});
    if(pos<0)throw new Error(`missing cell text ${text}`);
    state=state.apply(state.tr.setSelection(TextSelection.create(state.doc,pos)));
  };
  const run=(cmd:Command)=>expect(cmd(state,tr=>{state=state.apply(tr);})).toBe(true);
  const rows=()=>state.doc.child(0).content.content.map(r=>r.content.content.map(c=>c.textContent));
  const save=async()=>await (await import('jszip')).default.loadAsync(await writeDocx(state.doc,model));
  return {at,run,rows,save,get state(){return state;}};
}

describe('QA20K table structure operations 00721–00730',()=>{
  it('QA20K-00721 inserts a blank row after the active expense row',async()=>{
    const d=await open();d.at('B2');d.run(addRowAfter);
    expect(d.rows()).toHaveLength(4);
    expect(d.rows().map(r=>r.filter(Boolean))).toEqual([['A1','B1','C1'],['A2','B2','C2'],[],['A3','B3','C3']]);
  });
  it('QA20K-00722 inserts a blank row before the first data row without moving the header',async()=>{
    const source=table([row([cell('項目'),cell('數量'),cell('金額')]),row([cell('咖啡'),cell('2'),cell('80')]),row([cell('便當'),cell('1'),cell('120')])]);
    const d=await open(source);d.at('咖啡');d.run(addRowBefore);
    expect(d.rows()).toHaveLength(4);
    expect(d.rows()[0]).toEqual(['項目','數量','金額']);
    expect(d.rows()[2]).toEqual(['咖啡','2','80']);
  });
  it('QA20K-00723 deletes only the selected obsolete detail row',async()=>{
    const d=await open();d.at('B2');d.run(deleteRow);
    expect(d.rows()).toEqual([['A1','B1','C1'],['A3','B3','C3']]);
  });
  it('QA20K-00724 inserts a column before the first column and preserves every row',async()=>{
    const d=await open();d.at('A2');d.run(addColumnBefore);
    expect(d.rows().map(r=>r.length)).toEqual([4,4,4]);
    expect(d.rows().map(r=>r.slice(1))).toEqual([['A1','B1','C1'],['A2','B2','C2'],['A3','B3','C3']]);
  });
  it('QA20K-00725 inserts a column after the selected middle column',async()=>{
    const d=await open();d.at('B2');d.run(addColumnAfter);
    expect(d.rows().map(r=>r.length)).toEqual([4,4,4]);
    expect(d.rows()[1]).toEqual(['A2','B2','','C2']);
  });
  it('QA20K-00726 deletes only the selected final column',async()=>{
    const d=await open();d.at('C2');d.run(deleteColumn);
    expect(d.rows()).toEqual([['A1','B1'],['A2','B2'],['A3','B3']]);
  });
  it('QA20K-00727 deleting the middle row keeps rows on both sides in document order after save',async()=>{
    const d=await open();d.at('A2');d.run(deleteRow);
    const xml=await (await d.save()).file('word/document.xml')!.async('string');
    expect([...xml.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map(m=>m[1])).toEqual(['A1','B1','C1','A3','B3','C3']);
  });
  it('QA20K-00728 deleting the first column updates the saved grid to two columns',async()=>{
    const d=await open();d.at('A1');d.run(deleteColumn);
    const xml=await (await d.save()).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="1995"/><w:gridCol w:w="1995"/></w:tblGrid>');
    expect(xml).not.toContain('<w:t>A1</w:t>');
    expect(xml).toContain('<w:t>B1</w:t>');
  });
  it('QA20K-00729 inserting a row after a vertically merged row extends the merge safely',async()=>{
    const merged='<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid>'+[
      row([cell('Category',4000,'<w:vMerge w:val="restart"/>'),cell('Item')]),
      row([cell('',4000,'<w:vMerge/>'),cell('Travel')]),
    ].join('')+'</w:tbl>';
    const d=await open(merged);d.at('Item');d.run(addRowAfter);
    const xml=await (await d.save()).file('word/document.xml')!.async('string');
    const tableXml=xml.slice(xml.indexOf('<w:tbl>'),xml.indexOf('</w:tbl>'));
    const rows=[...tableXml.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].map(m=>m[1]);
    const cells=(r:string)=>[...r.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m=>m[1]);
    const cellState=(c:string)=>({text:[...c.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map(m=>m[1]).join(''),merge:c.match(/<w:vMerge(?: w:val="(restart)")?\/>/)?.[1]??(c.includes('<w:vMerge/>')?'continue':null)});
    expect((tableXml.match(/<w:gridCol /g)||[])).toHaveLength(2);
    expect(rows).toHaveLength(3);
    expect(rows.map(r=>cells(r).map(cellState))).toEqual([
      [{text:'Category',merge:'restart'},{text:'Item',merge:null}],
      [{text:'',merge:'continue'},{text:'',merge:null}],
      [{text:'',merge:'continue'},{text:'Travel',merge:null}],
    ]);
  });
  it('QA20K-00730 inserting a column creates one blank cell in each original row',async()=>{
    const d=await open();d.at('C1');d.run(addColumnAfter);
    expect(d.rows().map(r=>r[3])).toEqual(['','','']);
    expect(d.rows().map(r=>r.slice(0,3))).toEqual([['A1','B1','C1'],['A2','B2','C2'],['A3','B3','C3']]);
  });
});
