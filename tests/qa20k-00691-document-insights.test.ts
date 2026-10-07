import { beforeAll, describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { blankPackage } from '../src/papyrus/docx/template';
import { readDocx } from '../src/papyrus/docx/reader';
import { documentOutline, pageReferences, updatePageReferences } from '../src/papyrus/editor/outline';
import { countText } from '../src/papyrus/editor/wordCount';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
async function readBody(body: string, styles = '') {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  if (styles) zip.file('word/styles.xml', styles);
  return readDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const para=(text: string, pPr='')=>`<w:p>${pPr?`<w:pPr>${pPr}</w:pPr>`:''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const field=(bookmark:string,value:string)=>`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${bookmark} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>${value}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
const bookmark=(name:string,text:string)=>`<w:p><w:bookmarkStart w:id="1" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;

describe('QA20K office word-count and navigation operations 00691–00700',()=>{
  it('QA20K-00691 counts each CJK character in a bilingual invoice line',()=>{
    expect(countText('服務費 100 元')).toEqual({words:5,chars:7,charsWithSpaces:9});
  });
  it('QA20K-00692 treats a Latin phrase with punctuation as two words',()=>{
    expect(countText('Hello, world!')).toEqual({words:2,chars:12,charsWithSpaces:13});
  });
  it('QA20K-00693 counts full-width punctuation but excludes an inline object placeholder',()=>{
    expect(countText('你好，世界。\ufffc')).toEqual({words:6,chars:6,charsWithSpaces:6});
  });
  it('QA20K-00694 counts an alphanumeric product code as one token and reports spaces separately',()=>{
    expect(countText('GPT-4o 模型')).toEqual({words:3,chars:8,charsWithSpaces:9});
  });
  it('QA20K-00695 separates the whole-body count from a selected text count',async()=>{
    const zip=blankPackage();
    zip.file('word/document.xml',`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${para('合約 contract 草稿')}${para('第二段')}<w:sectPr/></w:body></w:document>`);
    const bytes=await zip.generateAsync({type:'uint8array'});
    const host=document.createElement('div');document.body.append(host);
    const {DocxEditor}=await import('../src/papyrus/editor/core');const ed=new DocxEditor(host);await ed.open(bytes);
    expect(ed.wordCount().words).toBe(8);expect(ed.selectionWordCount()).toBeNull();
    ed.view!.dispatch(ed.view!.state.tr.setSelection((await import('prosemirror-state')).TextSelection.create(ed.view!.state.doc,1,3)));
    expect(ed.selectionWordCount()?.words).toBe(2);ed.destroy();host.remove();
  });
  it('QA20K-00696 inherits custom child style outline level from its basedOn parent',async()=>{
    const styles=`<w:styles ${W}><w:style w:type="paragraph" w:styleId="ReportSubheadBase"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ReportSubhead"><w:basedOn w:val="ReportSubheadBase"/></w:style></w:styles>`;
    const {doc,model}=await readBody('<w:p><w:pPr><w:pStyle w:val="ReportSubheadBase"/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ReportSubhead"/></w:pPr><w:r><w:t>Revenue</w:t></w:r></w:p>',styles);
    expect(documentOutline(doc,model.styles).map(h=>[h.level,h.text])).toEqual([[1,'Quarterly report'],[1,'Revenue']]);
  });
  it('QA20K-00697 recognizes a direct paragraph outline level without a heading style',async()=>{
    const {doc,model}=await readBody('<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Approval detail</w:t></w:r></w:p>'+para('ordinary paragraph')+para('   ','<w:outlineLvl w:val="1"/>'));
    expect(documentOutline(doc,model.styles).map(h=>[h.level,h.text])).toEqual([[2,'Approval detail']]);
  });
  it('QA20K-00698 selecting a navigation heading moves the editor cursor into that paragraph',async()=>{
    const zip=blankPackage();
    zip.file('word/document.xml',`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Policy</w:t></w:r></w:p>${para('Body text')}<w:sectPr/></w:body></w:document>`);
    const host=document.createElement('div');document.body.append(host);
    const {DocxEditor}=await import('../src/papyrus/editor/core');
    const ed=new DocxEditor(host);await ed.open(await zip.generateAsync({type:'uint8array'}));
    const item=ed.outline()[0];
    expect(item).toMatchObject({level:0,text:'Policy'});
    ed.goToHeading(item.pos);
    expect(ed.view!.state.selection.$from.parent.textContent).toBe('Policy');
    expect(ed.view!.state.selection.$from.parentOffset).toBe(0);
    ed.destroy();host.remove();
  });
  it('QA20K-00699 refreshes multiple TOC page fields while retaining bold result marks',async()=>{
    const {doc}=await readBody(field('SectionA','9')+field('SectionB','9')+bookmark('SectionA','First')+bookmark('SectionB','Second'));
    const state=EditorState.create({doc});const tr=state.tr;
    const pages=new Map([['First',1],['Second',12]]);
    expect(pageReferences(tr.doc).map(r=>r.bookmark)).toEqual(['SectionA','SectionB']);
    expect(updatePageReferences(tr,pos=>pages.get(tr.doc.resolve(pos).parent.textContent)??null)).toBe(2);
    const refs=pageReferences(tr.doc);
    expect(refs.map(r=>tr.doc.textBetween(r.from,r.to))).toEqual(['1','12']);
    expect(tr.doc.nodeAt(refs[1].from)!.marks.map(m=>m.type.name)).toEqual(expect.arrayContaining(['bold','fieldResult']));
  });
  it('QA20K-00700 leaves an unresolved TOC bookmark result untouched',async()=>{
    const {doc}=await readBody(field('MissingSection','7'));
    const tr=EditorState.create({doc}).tr;
    expect(pageReferences(tr.doc)).toMatchObject([{bookmark:'MissingSection'}]);
    expect(updatePageReferences(tr,()=>null)).toBe(0);
    expect(tr.doc.textContent).toBe('7');
  });
});
