// Word count (Word's rules) and the spell check switch.
import { beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { countText } from '../../src/papyrus/editor/wordCount';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

describe('word count', () => {
  it('counts each East Asian character, and English words and numbers between spaces', () => {
    // The expected numbers are Word's own (ComputeStatistics words / characters / with spaces).
    expect(countText('服務費 100 元')).toEqual({ words: 5, chars: 7, charsWithSpaces: 9 });
    expect(countText('Hello, world!')).toEqual({ words: 2, chars: 12, charsWithSpaces: 13 });
    // Full-width punctuation counts as a word in Word; pictures are neither word nor character.
    expect(countText('你好，世界。\ufffc')).toEqual({ words: 6, chars: 6, charsWithSpaces: 6 });
    expect(countText('GPT-4o 模型')).toEqual({ words: 3, chars: 8, charsWithSpaces: 9 });
  });

  it('the editor counts the body and the selection', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>合約 contract 草稿</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p></w:body></w:document>');
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    expect(ed.wordCount().words).toBe(8); // Word: 「合約 contract 草稿第二段」 is 8 words
    expect(ed.selectionWordCount()).toBeNull();
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 1, 3)));
    expect(ed.selectionWordCount()?.words).toBe(2);
    ed.destroy();
    host.remove();
  });
});

describe('spell check', () => {
  it('is off by default and can be turned on and off', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.newDocument();
    expect(ed.view!.dom.getAttribute('spellcheck')).toBe('false');
    ed.setSpellcheck(true);
    expect(ed.view!.dom.getAttribute('spellcheck')).toBe('true');
    expect(ed.spellcheckOn).toBe(true);
    ed.setSpellcheck(false);
    expect(ed.view!.dom.getAttribute('spellcheck')).toBe('false');
    ed.destroy();
    const on = new DocxEditor(host, { spellcheck: true });
    await on.newDocument();
    expect(on.view!.dom.getAttribute('spellcheck')).toBe('true');
    on.destroy();
    host.remove();
  });
});
