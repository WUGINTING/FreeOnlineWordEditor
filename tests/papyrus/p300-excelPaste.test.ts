// persona-300 B-9: Ctrl+V of an Excel range pastes the cells (not the picture of them Excel puts
// next to them), with the look Excel keeps in its <style> block and the <col> widths.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { inlineExcelStyles, isExcelHtml } from '../../src/papyrus/editor/pasteExcel';
import { domStubs, setup } from './p300Helpers';

domStubs();

/** What Excel (Microsoft 365, zh-TW) puts on the clipboard as text/html for a 2 × 2 range. */
const EXCEL = `<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<link id=Main-File rel=Main-File href="file:///C:/Users/a/AppData/Local/Temp/msohtmlclip1/01/clip.htm">
<style>
<!--table
	{mso-displayed-decimal-separator:"\\.";
	mso-displayed-thousand-separator:"\\,";}
@page
	{margin:.75in .7in .75in .7in;}
td
	{padding-top:1px;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-family:新細明體, serif;
	text-align:general;
	vertical-align:middle;
	border:none;
	white-space:nowrap;}
.xl65
	{font-weight:700;
	text-align:center;
	border:.5pt solid windowtext;
	background:#FFFF00;
	mso-pattern:black none;}
.xl66
	{text-align:right;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-number-format:"\\#\\,\\#\\#0";}
-->
</style>
</head>
<body link="#0563C1" vlink="#954F72">
<table border=0 cellpadding=0 cellspacing=0 width=179 style='border-collapse:collapse;width:134pt'>
<!--StartFragment-->
 <col width=107 style='mso-width-source:userset;mso-width-alt:3424;width:80pt'>
 <col width=72 style='width:54pt'>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl65 width=107 style='height:16.5pt;width:80pt'>項目</td>
  <td class=xl65 width=72 style='border-left:none;width:54pt'>金額</td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl66 style='height:16.5pt'>設備</td>
  <td class=xl66 align=right style='border-left:none'>12,000</td>
 </tr>
<!--EndFragment-->
</table>
</body>
</html>`;

/** A paste event carrying Excel's HTML, its text and the picture of the range. */
function excelPaste() {
  const e = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
  const picture = new File([new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])], 'image.png', { type: 'image/png' });
  Object.defineProperty(e, 'clipboardData', {
    value: {
      types: ['text/html', 'text/plain', 'Files'],
      files: [picture],
      getData: (t: string) => (t === 'text/html' ? EXCEL : t === 'text/plain' ? '項目\t金額\r\n設備\t12,000\r\n' : ''),
    },
  });
  return e;
}

function cells(doc: PMNode): PMNode[] {
  const out: PMNode[] = [];
  doc.descendants((n) => {
    if (n.type.name === 'table_cell' || n.type.name === 'table_header') out.push(n);
    return true;
  });
  return out;
}

describe('Excel ranges', () => {
  it('knows Excel HTML; other HTML is left as it is', () => {
    expect(isExcelHtml(EXCEL)).toBe(true);
    const html = '<table><tr><td class=a>x</td></tr></table><style>.a{background:red}</style>';
    expect(inlineExcelStyles(html)).toBe(html);
  });

  it('writes the class styles and column widths on the cells', () => {
    const out = inlineExcelStyles(EXCEL);
    const t = document.createElement('template');
    t.innerHTML = out;
    const tds = Array.from(t.content.querySelectorAll('td'));
    expect(tds.map((td) => td.getAttribute('data-colwidth'))).toEqual(['107', '72', '107', '72']);
    expect(tds[0].getAttribute('style')).toContain('font-weight:700');
    expect(tds[0].getAttribute('style')).toContain('background:#FFFF00');
    expect(tds[0].querySelector('p')!.style.textAlign).toBe('center');
    expect(JSON.parse(tds[1].getAttribute('data-dx-borders')!).left).toEqual({ val: 'nil', sz: null, color: null });
    expect(JSON.parse(tds[1].getAttribute('data-dx-borders')!).top).toEqual({ val: 'single', sz: 4, color: 'auto' });
  });

  it('Ctrl+V pastes the cells, not the picture of them, with their look and widths; saved as Word table cells', async () => {
    const d = await setup('<w:p/>');
    d.view.dom.dispatchEvent(excelPaste());
    const doc = d.view.state.doc;
    let images = 0;
    doc.descendants((n) => void (n.type.name === 'image' && images++));
    expect(images).toBe(0);
    const c = cells(doc);
    expect(c.map((n) => n.textContent)).toEqual(['項目', '金額', '設備', '12,000']);
    expect(c.map((n) => n.attrs.colwidth)).toEqual([[107], [72], [107], [72]]);
    expect(c[0].attrs.background).toBe('#FFFF00');
    expect(c[2].attrs.background).toBeNull();
    expect(c[0].firstChild!.attrs.align).toBe('center');
    expect(c[3].firstChild!.attrs.align).toBe('right');
    expect(c[0].firstChild!.firstChild!.marks.map((m) => m.type.name)).toContain('bold');
    expect(c[2].firstChild!.firstChild!.marks.map((m) => m.type.name)).not.toContain('bold');
    expect(c[1].attrs.borders.left).toEqual({ val: 'nil', sz: null, color: null });
    expect(c[2].attrs.borders.top).toEqual({ val: 'nil', sz: null, color: null });
    expect(c[2].attrs.borders.bottom).toEqual({ val: 'single', sz: 4, color: 'auto' });
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>');
    expect(xml).toMatch(/<w:tcBorders>.*<w:bottom w:val="single" w:sz="4"( w:space="0")? w:color="auto"\/>/);
    expect(xml).toMatch(/<w:jc w:val="center"\/>/);
    expect(xml).toMatch(/<w:gridCol w:w="1605"\/><w:gridCol w:w="1080"\/>/);
    d.done();
  });
});
