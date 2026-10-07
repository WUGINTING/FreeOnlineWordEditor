import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { undoDepth } from 'prosemirror-history';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';

let editors: DocxEditor[] = [];
let hosts: HTMLElement[] = [];
beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});
afterEach(() => { for (const ed of editors) ed.destroy(); for (const host of hosts) host.remove(); editors=[]; hosts=[]; });
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p=(text:string)=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
async function docx(body:string){ const zip=blankPackage(); zip.file('word/document.xml',`<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`); return new Uint8Array(await zip.generateAsync({type:'uint8array'})); }
function editor(options: Record<string,unknown>={}){const host=document.createElement('div');document.body.append(host);hosts.push(host);const ed=new DocxEditor(host,options);editors.push(ed);return ed;}

// Open lifecycle cases assert importer/editor state, not roundtrip delivery or editor UI layout.
describe('QA20K DOCX open lifecycle workflows 01061–01070',()=>{
  it('QA20K-01061 opens a locally selected DOCX Blob as a clean editable document',async()=>{
    const ed=editor(); const file=new Blob([await docx(p('Invoice for March'))],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    await ed.open(file);
    expect(ed.view!.state.doc.textContent).toBe('Invoice for March');
    expect(ed.view!.editable).toBe(true); expect(ed.isModified()).toBe(false);
  });
  it('QA20K-01062 switches from the current customer file to the newly selected file',async()=>{
    const ed=editor(); await ed.open(await docx(p('Old client draft'))); await ed.open(await docx(p('New client agreement')));
    expect(ed.view!.state.doc.textContent).toBe('New client agreement');
    expect(ed.isModified()).toBe(false);
  });
  it('QA20K-01063 starts a blank document after leaving an existing report',async()=>{
    const ed=editor(); await ed.open(await docx(p('Quarterly report'))); await ed.newDocument();
    expect(ed.view!.state.doc.textContent).toBe(''); expect(ed.isModified()).toBe(false);
  });
  it('QA20K-01064 keeps the active document usable when the selected file is not a DOCX archive',async()=>{
    const ed=editor(); await ed.open(await docx(p('Recovery copy')));
    await expect(ed.open(new Uint8Array([0x4e,0x4f,0x54,0x5a,0x49,0x50]))).rejects.toBeTruthy();
    expect(ed.view!.state.doc.textContent).toBe('Recovery copy');
    expect(ed.isModified()).toBe(false);
  });
  it('QA20K-01065 opens an empty template as an empty saved document',async()=>{
    const ed=editor(); await ed.open(await docx(''));
    expect(ed.view!.state.doc.textContent).toBe(''); expect(ed.isModified()).toBe(false);
    expect(ed.savePoint()!.body.eq(ed.view!.state.doc)).toBe(true);
  });
  it('QA20K-01066 opens an imported expense table with its row and cell structure',async()=>{
    const ed=editor(); const body='<w:tbl><w:tr><w:tc>'+p('Item')+'</w:tc><w:tc>'+p('Amount')+'</w:tc></w:tr><w:tr><w:tc>'+p('Travel')+'</w:tc><w:tc>'+p('1200')+'</w:tc></w:tr></w:tbl>';
    await ed.open(await docx(body)); const table=ed.view!.state.doc.firstChild!;
    expect(table.type.name).toBe('table'); expect(table.childCount).toBe(2);
    expect(table.child(0).childCount).toBe(2); expect(table.child(1).childCount).toBe(2);
    expect(table.textContent).toBe('ItemAmountTravel1200');
  });
  it('QA20K-01067 keeps a pre-existing hyperlink destination while opening the file',async()=>{
    const ed=editor(); const zip=blankPackage();
    zip.file('word/document.xml',`<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Open </w:t></w:r><w:hyperlink r:id="rLink"><w:r><w:t>policy portal</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`);
    const rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const rels=await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels',rels.replace('</Relationships>',`<Relationship Id="rLink" Type="${rel}" Target="https://intranet.example/policy" TargetMode="External"/></Relationships>`));
    await ed.open(new Uint8Array(await zip.generateAsync({type:'uint8array'})));
    let href:string|undefined, label=''; ed.view!.state.doc.descendants(n=>{if(n.isText&&n.text==='policy portal'){label=n.text;href=n.marks.find(m=>m.type.name==='link')?.attrs.href;}});
    expect(label).toBe('policy portal'); expect(href).toBe('https://intranet.example/policy');
  });
  it('QA20K-01068 opens a policy file in view-only mode with its text available',async()=>{
    const ed=editor({editable:false}); await ed.open(await docx(p('Policy is read only')));
    expect(ed.view!.state.doc.textContent).toBe('Policy is read only');
    expect(ed.view!.editable).toBe(false);
  });
  it('QA20K-01069 resets prior edit undo history when a replacement document opens',async()=>{
    const ed=editor(); await ed.open(await docx(p('Draft')));
    ed.view!.dispatch(ed.view!.state.tr.insertText(' v2',6));
    expect(undoDepth(ed.view!.state)).toBeGreaterThan(0);
    await ed.open(await docx(p('Approved')));
    expect(ed.view!.state.doc.textContent).toBe('Approved');
    expect(undoDepth(ed.view!.state)).toBe(0); expect(ed.isModified()).toBe(false);
  });
  it('QA20K-01070 accepts a DOCX whose package has no optional comments part',async()=>{
    const ed=editor(); await ed.open(await docx(p('Ordinary memo without comments')));
    expect(ed.view!.state.doc.textContent).toBe('Ordinary memo without comments');
    expect(ed.comments()).toEqual([]); expect(ed.isModified()).toBe(false);
  });
});

