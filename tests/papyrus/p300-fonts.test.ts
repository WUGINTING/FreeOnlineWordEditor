// persona-300 B-3 / B-8: a Latin font from the 字型 box leaves the Chinese text's font, as in
// Word; Taiwanese government fonts fall back to similar fonts on computers without them, and
// the editor says which of the document's fonts this computer lacks.
import { afterEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { setFont } from '../../src/papyrus/editor/commands';
import { schema } from '../../src/papyrus/editor/schema';
import { fontStack, isEastAsianFont, missingFonts, resetFontChecks } from '../../src/papyrus/docx/fonts';
import { domStubs, setup } from './p300Helpers';

domStubs();
afterEach(() => {
  vi.restoreAllMocks();
  resetFontChecks();
});

const KAI_RUN = '<w:p><w:r><w:rPr><w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hAnsi="標楷體"/></w:rPr><w:t>主旨ABC</w:t></w:r></w:p>';

function selectAll(view: any) {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.child(0).nodeSize - 1)));
}

describe('字型 box, as in Word', () => {
  it('knows East Asian fonts from Latin ones', () => {
    expect(['PMingLiU', 'DFKai-SB', 'Microsoft JhengHei', '標楷體', '華康中楷體'].every(isEastAsianFont)).toBe(true);
    expect(['Times New Roman', 'Arial', 'Calibri', 'Courier New', ''].some(isEastAsianFont)).toBe(false);
  });

  it('Times New Roman on 標楷體 text changes the Latin font only (w:ascii / w:hAnsi), the Chinese text keeps 標楷體', async () => {
    const d = await setup(KAI_RUN);
    selectAll(d.view);
    d.editor.run(setFont('Times New Roman', false));
    const font = schema.marks.font.isInSet(d.view.state.doc.child(0).child(0).marks)!;
    expect(font.attrs).toEqual({ family: 'Times New Roman', eastAsia: '標楷體' });
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"\/>/);
    d.done();
  });

  it('標楷體 is set for Chinese and Latin text alike; （預設字型） removes the run font', async () => {
    const d = await setup('<w:p><w:r><w:t>主旨ABC</w:t></w:r></w:p>');
    selectAll(d.view);
    d.editor.run(setFont('DFKai-SB', true));
    expect(schema.marks.font.isInSet(d.view.state.doc.child(0).child(0).marks)!.attrs).toEqual({ family: 'DFKai-SB', eastAsia: 'DFKai-SB' });
    selectAll(d.view);
    d.editor.run(setFont(null, false));
    expect(schema.marks.font.isInSet(d.view.state.doc.child(0).child(0).marks)).toBeUndefined();
    d.done();
  });

  it('runs with different Chinese fonts each keep theirs; at the cursor the next typing gets the Latin font only', async () => {
    const d = await setup(
      '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="標楷體"/></w:rPr><w:t>甲</w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia="新細明體"/></w:rPr><w:t>乙</w:t></w:r></w:p>',
    );
    selectAll(d.view);
    d.editor.run(setFont('Arial', false));
    const para = d.view.state.doc.child(0);
    expect([0, 1].map((i) => schema.marks.font.isInSet(para.child(i).marks)!.attrs)).toEqual([
      { family: 'Arial', eastAsia: '標楷體' },
      { family: 'Arial', eastAsia: '新細明體' },
    ]);
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 2)));
    d.editor.run(setFont('Times New Roman', false));
    expect(schema.marks.font.isInSet(d.view.state.storedMarks ?? [])!.attrs).toEqual({ family: 'Times New Roman', eastAsia: '標楷體' });
    d.done();
  });

  it('the ribbon applies it that way', async () => {
    const d = await setup(KAI_RUN, 'DocxToolbar.vue', { styles: [] });
    selectAll(d.view);
    await d.refresh();
    const box = d.panel.querySelector('select[title="字型"]') as HTMLSelectElement;
    box.value = 'Times New Roman';
    box.dispatchEvent(new Event('change'));
    expect(schema.marks.font.isInSet(d.view.state.doc.child(0).child(0).marks)!.attrs).toEqual({ family: 'Times New Roman', eastAsia: '標楷體' });
    d.done();
  }, 20000);
});

describe('fonts on computers without them', () => {
  it('標楷體 and 新細明體 fall back to similar fonts, then a serif font', () => {
    expect(fontStack('DFKai-SB', true)).toBe(
      '"DFKai-SB", "標楷體", "BiauKai", "TW-Kai", "TW-Kai-98_1", "Kaiti TC", "STKaiti", "KaiTi", "AR PL UKai TW", "AR PL UKai CN", serif',
    );
    expect(fontStack('新細明體', true)).toMatch(/^"新細明體", "PMingLiU", .*"Apple LiSung".*, serif$/);
    // A Latin font: other names only, no generic family (Chinese text must reach its own font).
    expect(fontStack('Times New Roman')).toBe('"Times New Roman", "Times", "Liberation Serif", "Tinos"');
    expect(fontStack('Some Font', true)).toBe('"Some Font"');
  });

  it('the page styles and the run fonts use the lists', async () => {
    const d = await setup(KAI_RUN);
    const style = (d.view.dom.querySelector('span[style*="font-family"]') as HTMLElement).getAttribute('style')!;
    expect(style).toContain('"標楷體", "DFKai-SB", "BiauKai"');
    // A new document's defaults: 標楷體 / Times New Roman (persona-300 12).
    expect(d.editor.model.css).toContain('--dx-font-e:"標楷體", "DFKai-SB", "BiauKai"');
    d.done();
  });

  it('says which of the document fonts this computer does not have, and what is shown instead', async () => {
    // A computer with Kaiti TC but no 標楷體 (nor BiauKai, the Mac's 標楷體) or 新細明體.
    const installed = new Set(['Kaiti TC', 'Times New Roman', 'Calibri']);
    let font = '';
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
      () =>
        ({
          set font(v: string) {
            font = v;
          },
          get font() {
            return font;
          },
          measureText: () => {
            const named = /"([^"]+)"/.exec(font)?.[1];
            const generic = font.split(',').pop()!.trim().replace('72px ', '');
            return { width: named && installed.has(named) ? 500 + named.length : { monospace: 600, serif: 450, 'sans-serif': 480 }[generic]! };
          },
        }) as any,
    );
    resetFontChecks();
    expect(missingFonts(['DFKai-SB', '標楷體', 'Times New Roman', 'PMingLiU'])).toEqual([
      { font: 'DFKai-SB', label: '標楷體', substitute: 'Kaiti TC' },
      { font: 'PMingLiU', label: '新細明體', substitute: null },
    ]);
    const d = await setup(KAI_RUN);
    expect(d.editor.documentFonts()).toEqual(expect.arrayContaining(['標楷體', 'Times New Roman']));
    expect(d.editor.missingFonts()).toEqual([{ font: '標楷體', label: '標楷體', substitute: 'Kaiti TC' }]);
    d.done();
  });

  it('without a canvas nothing is reported missing', () => {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => null);
    resetFontChecks();
    expect(missingFonts(['DFKai-SB'])).toEqual([]);
  });
});
