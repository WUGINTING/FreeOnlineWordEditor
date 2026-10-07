// @vitest-environment jsdom
/**
 * The print HTML (DocxEditor.printHtml, used for 列印 and the server-made PDF) places a copy of
 * each block where its border box was measured on screen. A style's own "space before" (for
 * example `.dx-e1 .dx-doc .dx-ps-Heading1 { margin-top: 32px }`, more specific than the print
 * rule) must not move the copy: found by converting a real document, where every Heading 1 of
 * the PDF sat 32 px too low and covered the next line. The layout itself needs a real browser
 * (verified there: 861 blocks of 3 documents, 0 misplaced); jsdom checks the rule.
 */
import {describe, expect, it} from 'vitest';
import {blankPackage} from '../src/papyrus/docx/template';
import {DocxEditor} from '../src/papyrus/editor/core';

describe('print HTML', () => {
  it('takes every margin off the copied blocks, whatever the document styles say', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await blankPackage().generateAsync({type: 'uint8array'}));
    (ed as any).pages = [{...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0}];
    const html = ed.printHtml('報告')!;
    ed.destroy();
    host.remove();
    const rule = /\.dx-print-page \.dx-print-block\{([^}]*)\}/.exec(html)?.[1] ?? '';
    expect(rule).toContain('position:absolute');
    expect(rule).toContain('margin:0!important');
  });
});
