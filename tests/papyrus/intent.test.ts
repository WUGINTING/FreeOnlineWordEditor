// Every editing action must change only what the user meant: formatting commands must not
// strip Word structure, and buttons must act on what the text really shows (styles included).
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command, type Plugin } from 'prosemirror-state';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import {
  clearFormatting, effectiveAlign, selectionHas, selectionValue, setAlign, toggleFormat,
} from '../../src/papyrus/editor/commands';
import { adoptPastedLists } from '../../src/papyrus/editor/pasteLists';
import { emptyNumbering } from '../../src/papyrus/docx/model';
import { ListCounter } from '../../src/papyrus/docx/numbering';
import { review } from '../../src/papyrus/editor/review';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  // A centered, bold style, and one based on it (inheritance goes through w:basedOn).
  '<w:style w:type="paragraph" w:styleId="Banner"><w:name w:val="Banner"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="SubBanner"><w:name w:val="Sub Banner"/><w:basedOn w:val="Banner"/></w:style>' +
  '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
  '</w:styles>';

async function open(body: string, plugins: Plugin[] = []) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  zip.file('word/styles.xml', STYLES);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins });
  const run = (cmd: Command) => cmd(state, (tr) => (state = state.apply(tr)));
  const select = (from: number, to = from) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to))));
  const save = async () => (await JSZip.loadAsync(await writeDocx(state.doc, model))).file('word/document.xml')!.async('string');
  return { model, run, select, save, get state() { return state; } };
}

/** Position range of the whole text of paragraph `i` (0-based, top-level). */
function paragraphText(state: EditorState, i: number): [number, number] {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += state.doc.child(k).nodeSize;
  return [pos + 1, pos + 1 + state.doc.child(i).content.size];
}

describe('clear formatting keeps what the text means', () => {
  it('keeps tracked insertions (with author) and content controls, drops only the looks', async () => {
    const d = await open(
      '<w:p><w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="32"/><w:lang w:val="en-US"/></w:rPr><w:t>added</w:t></w:r></w:ins></w:p>' +
        '<w:p><w:sdt><w:sdtPr><w:alias w:val="客戶名稱"/><w:tag w:val="customer"/><w:id w:val="42"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t>王小明</w:t></w:r></w:sdtContent></w:sdt></w:p>',
    );
    d.select(1, d.state.doc.content.size - 1);
    d.run(clearFormatting);
    const xml = await d.save();

    // Structure: the revision with its author, the content control with its name and tag.
    expect(xml).toContain('<w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z">');
    expect(xml).toContain('<w:alias w:val="客戶名稱"/><w:tag w:val="customer"/><w:id w:val="42"/>');
    expect(xml).toContain('<w:t>added</w:t>');
    expect(xml).toContain('<w:t>王小明</w:t>');
    // Meaning kept (language), looks gone.
    expect(xml).toContain('<w:lang w:val="en-US"/>');
    for (const gone of ['<w:b/>', '<w:color', '<w:sz ', '<w:i/>', '<w:u ']) expect(xml).not.toContain(gone);
  });

  it('at the cursor, new text gets the cleared formatting and is not tracked', async () => {
    // Word with tracking off (checked in Word): text typed inside a tracked insertion is not
    // part of it; the insertion is split around it and its second part gets a new w:id.
    // (With the editor's review plugin, as in the editor.)
    const d = await open('<w:p><w:ins w:id="7" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r></w:ins></w:p>', [review()]);
    d.select(3);
    d.run(clearFormatting);
    d.run((s, dispatch) => {
      dispatch?.(s.tr.insertText('X'));
      return true;
    });
    const xml = await d.save();
    expect(xml).toMatch(
      /<w:ins w:id="7" w:author="Bob"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>ab<\/w:t><\/w:r><\/w:ins><w:r>(<w:rPr\/>)?<w:t>X<\/w:t><\/w:r><w:ins w:id="(?!7")\d+" w:author="Bob"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>c<\/w:t><\/w:r><\/w:ins>/,
    );
  });
});

describe('formatting that comes from styles', () => {
  it('shows the style’s alignment and bold, and "align left" beats a centered style', async () => {
    const d = await open('<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>');
    d.select(2);
    const styles = d.model.styles;
    expect(effectiveAlign(d.state.doc.child(0).attrs, styles)).toBe('center');
    expect(selectionHas(d.state, 'bold', styles)).toBe(true);

    d.run(setAlign('left', styles));
    expect(effectiveAlign(d.state.doc.child(0).attrs, styles)).toBe('left');
    expect(await d.save()).toContain('<w:pStyle w:val="Banner"/><w:jc w:val="left"/>');

    // Back to the style's own alignment: the explicit setting goes away again.
    d.run(setAlign('center', styles));
    const xml = await d.save();
    expect(xml).toContain('<w:pPr><w:pStyle w:val="Banner"/></w:pPr>');
  });

  it('inheritance goes through basedOn', async () => {
    const d = await open('<w:p><w:pPr><w:pStyle w:val="SubBanner"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>');
    d.select(1);
    expect(effectiveAlign(d.state.doc.child(0).attrs, d.model.styles)).toBe('center');
    expect(selectionHas(d.state, 'bold', d.model.styles)).toBe(true);
  });

  it('turning bold off inside a bold style writes an explicit "not bold", and on again removes it', async () => {
    const d = await open('<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>');
    const styles = d.model.styles;
    const [from, to] = paragraphText(d.state, 0);
    d.select(from, to);
    d.run(toggleFormat('bold', styles));
    expect(selectionHas(d.state, 'bold', styles)).toBe(false);
    expect(await d.save()).toContain('<w:r><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr><w:t>Title</w:t></w:r>');

    d.run(toggleFormat('bold', styles));
    expect(selectionHas(d.state, 'bold', styles)).toBe(true);
    expect(await d.save()).toContain('<w:r><w:t>Title</w:t></w:r>');
  });

  it('a mixed selection becomes bold everywhere with the fewest explicit settings', async () => {
    const d = await open(
      '<w:p><w:r><w:t>plain </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>direct </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>styled</w:t></w:r></w:p>',
    );
    const styles = d.model.styles;
    const [from, to] = paragraphText(d.state, 0);
    d.select(from, to);
    expect(selectionHas(d.state, 'bold', styles)).toBe(false);
    d.run(toggleFormat('bold', styles));
    expect(selectionHas(d.state, 'bold', styles)).toBe(true);
    const xml = await d.save();
    expect(xml).toContain('<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">plain </w:t></w:r>');
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">direct </w:t></w:r>');
    // Bold already comes from the character style: nothing added.
    expect(xml).toContain('<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>styled</w:t></w:r>');
  });

  it('font size, colour and font come from direct formatting, character style, paragraph style, then defaults', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
      '<w:p><w:pPr><w:pStyle w:val="Blue18"/></w:pPr><w:r><w:t>styled</w:t></w:r><w:r><w:rPr><w:sz w:val="20"/><w:color w:val="FF0000"/></w:rPr><w:t>direct</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:rStyle w:val="Big"/></w:rPr><w:t>charstyle</w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>' +
      `${SECT}</w:body></w:document>`);
    zip.file('word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="PMingLiU"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Blue18"><w:name w:val="Blue18"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="36"/><w:color w:val="0070C0"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Big"><w:name w:val="Big"/><w:rPr><w:rFonts w:eastAsia="DFKai-SB"/><w:sz w:val="48"/></w:rPr></w:style>' +
      '</w:styles>');
    const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const at = (from: number, to = from) => {
      const s = EditorState.create({ schema, doc });
      return s.apply(s.tr.setSelection(TextSelection.create(s.doc, from, to)));
    };
    const read = (st: EditorState) => ({
      size: selectionValue(st, 'fontSize', model.styles),
      color: selectionValue(st, 'color', model.styles),
      font: selectionValue(st, 'font', model.styles),
    });
    // Cursor in "styled": the paragraph style's 18 pt blue (#0070C0), the default CJK font.
    expect(read(at(3))).toEqual({
      size: { value: 18, mixed: false }, color: { value: '#0070c0', mixed: false }, font: { value: 'PMingLiU', mixed: false },
    });
    // "direct": its own 10 pt red wins over the style.
    expect(read(at(10)).size.value).toBe(10);
    expect(read(at(10)).color.value).toBe('#ff0000');
    // The whole first paragraph: two sizes and two colours → mixed.
    const whole = read(at(1, 13));
    expect(whole.size).toEqual({ value: null, mixed: true });
    expect(whole.color).toEqual({ value: null, mixed: true });
    expect(whole.font).toEqual({ value: 'PMingLiU', mixed: false });
    // Second paragraph: character style (24 pt, DFKai-SB), then plain text from the defaults (12 pt).
    const p2 = doc.child(0).nodeSize + 1;
    expect(read(at(p2 + 3)).size.value).toBe(24);
    expect(read(at(p2 + 3)).font.value).toBe('DFKai-SB');
    expect(read(at(p2 + 12)).size.value).toBe(12);
    expect(read(at(p2 + 12)).color.value).toBeNull();
  });

  it('keeps an existing explicit "not bold" from Word as it is', async () => {
    const body = '<w:p><w:pPr><w:pStyle w:val="Banner"/></w:pPr><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>normal</w:t></w:r></w:p>';
    const d = await open(body);
    d.select(2);
    expect(selectionHas(d.state, 'bold', d.model.styles)).toBe(false);
    expect(await d.save()).toContain('<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>normal</w:t></w:r>');
  });
});

describe('pasting lists', () => {
  const parse = (html: string) => {
    const div = document.createElement('div');
    div.innerHTML = html;
    return PMDOMParser.fromSchema(schema).parseSlice(div);
  };

  it('an ordered list starting at 3 stays a numbered list starting at 3', () => {
    const numbering = emptyNumbering();
    const slice = adoptPastedLists(parse('<ol start="3"><li>三</li><li>四</li></ol>'), numbering);
    const paras: any[] = [];
    slice.content.forEach((n) => paras.push(n));
    expect(paras.map((p) => p.textContent)).toEqual(['三', '四']);
    const numId = paras[0].attrs.numId;
    expect(numId).toBeTruthy();
    expect(paras[1].attrs.numId).toBe(numId);
    const counter = new ListCounter(numbering);
    expect([counter.next(numId, 0), counter.next(numId, 0)]).toEqual(['3.', '4.']);
  });

  it('nested and bulleted lists keep their levels, and each pasted list is its own list', () => {
    const numbering = emptyNumbering();
    const slice = adoptPastedLists(
      parse('<ul><li>a<ul><li>a.1</li></ul></li><li>b</li></ul><ol><li><p>one</p></li><li><p>two</p></li></ol>'),
      numbering,
    );
    const paras: any[] = [];
    slice.content.forEach((n) => paras.push(n));
    expect(paras.map((p) => [p.textContent, p.attrs.ilvl])).toEqual([['a', 0], ['a.1', 1], ['b', 0], ['one', 0], ['two', 0]]);
    const bullets = paras[0].attrs.numId;
    const numbers = paras[3].attrs.numId;
    expect(paras[1].attrs.numId).toBe(bullets);
    expect(numbers).not.toBe(bullets);
    const fmt = (numId: string, lvl: number) => numbering.abstracts[numbering.nums[numId].abstractId].levels[lvl].fmt;
    expect(fmt(bullets, 0)).toBe('bullet');
    expect(fmt(bullets, 1)).toBe('bullet');
    expect(fmt(numbers, 0)).toBe('decimal');
  });

  it('a list paragraph copied inside the document keeps its list', () => {
    const numbering = emptyNumbering();
    numbering.abstracts['9'] = { id: '9', levels: [{ fmt: 'decimal', text: '%1.', start: 1, indLeft: 720, hanging: 360 }] };
    numbering.nums['4'] = { id: '4', abstractId: '9', startOverrides: {} };
    const para = schema.nodes.paragraph.create({ numId: '4', ilvl: 0 }, schema.text('item'));
    const { DOMSerializer } = require('prosemirror-model');
    const html = DOMSerializer.fromSchema(schema).serializeNode(para) as HTMLElement;
    const slice = adoptPastedLists(parse(html.outerHTML), numbering);
    expect(slice.content.firstChild!.attrs.numId).toBe('4');
  });

  it('pasted into an empty paragraph, the middle, the end or the start of text: every item keeps its number, the text around stays plain', async () => {
    const qa = '<ol start="3">\n  <li>QA_PASTED_ITEM_THREE</li>\n  <li>QA_PASTED_ITEM_FOUR</li>\n</ol>';
    // Positions in: <p>before</p><p></p><p>after</p>
    for (const [where, pos, around] of [
      ['empty', 9, ['before', 'after']],
      ['middle', 4, ['bef', 'ore', '', 'after']],
      ['end', 7, ['before', '', 'after']],
      ['start', 1, ['before', '', 'after']],
    ] as const) {
      const d = await open('<w:p><w:r><w:t>before</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>after</w:t></w:r></w:p>');
      d.select(pos);
      const slice = adoptPastedLists(parse(qa), d.model.numbering);
      d.run((s, dispatch) => (dispatch?.(s.tr.replaceSelection(slice)), true));

        const items: any[] = [];
        const plain: string[] = [];
        const sequence: string[] = [];
        d.state.doc.forEach((n) => {
          if (n.attrs.numId) {
            items.push(n);
            sequence.push(`list:${n.textContent}`);
          } else {
            plain.push(n.textContent);
            sequence.push(`plain:${n.textContent}`);
          }
        });
        expect(items.map((n) => n.textContent), where).toEqual(['QA_PASTED_ITEM_THREE', 'QA_PASTED_ITEM_FOUR']);
        expect(plain, where).toEqual(around);
        const orderByPosition: Record<string, string[]> = {
          empty: ['plain:before', 'list:QA_PASTED_ITEM_THREE', 'list:QA_PASTED_ITEM_FOUR', 'plain:after'],
          middle: ['plain:bef', 'list:QA_PASTED_ITEM_THREE', 'list:QA_PASTED_ITEM_FOUR', 'plain:ore', 'plain:', 'plain:after'],
          end: ['plain:before', 'list:QA_PASTED_ITEM_THREE', 'list:QA_PASTED_ITEM_FOUR', 'plain:', 'plain:after'],
          start: ['list:QA_PASTED_ITEM_THREE', 'list:QA_PASTED_ITEM_FOUR', 'plain:before', 'plain:', 'plain:after'],
        };
        expect(sequence, `${where}: list/plain document order`).toEqual(orderByPosition[where]);
      const counter = new ListCounter(d.model.numbering);
      expect(items.map((n) => counter.next(n.attrs.numId, n.attrs.ilvl)), where).toEqual(['3.', '4.']);

      // Saved and opened again: still a 3. / 4. list.
      const zip = await JSZip.loadAsync(await writeDocx(d.state.doc, d.model));
      const again = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
      const reopened: any[] = [];
      again.doc.forEach((n) => n.attrs.numId && reopened.push(n));
      const c2 = new ListCounter(again.model.numbering);
      expect(reopened.map((n) => c2.next(n.attrs.numId, n.attrs.ilvl)), where).toEqual(['3.', '4.']);
    }
  });

  it('three items and a nested list pasted into existing text keep numbers and levels', async () => {
    const d = await open('<w:p><w:r><w:t>text</w:t></w:r></w:p>');
    d.select(3);
    const slice = adoptPastedLists(parse('<ol><li>a</li><li>b<ol type="a"><li>b1</li></ol></li><li>c</li></ol>'), d.model.numbering);
    d.run((s, dispatch) => (dispatch?.(s.tr.replaceSelection(slice)), true));
      const items: any[] = [];
      const paragraphs: string[] = [];
      const documentOrder: string[] = [];
      d.state.doc.forEach((n) => {
        paragraphs.push(n.textContent);
        documentOrder.push(n.attrs.numId ? `list:${n.textContent}` : `plain:${n.textContent}`);
        if (n.attrs.numId) items.push(n);
      });
      expect(paragraphs).toEqual(['te', 'a', 'b', 'b1', 'c', 'xt']);
      expect(documentOrder).toEqual(['plain:te', 'list:a', 'list:b', 'list:b1', 'list:c', 'plain:xt']);
    const counter = new ListCounter(d.model.numbering);
    expect(items.map((n) => [n.textContent, counter.next(n.attrs.numId, n.attrs.ilvl)])).toEqual([
      ['a', '1.'], ['b', '2.'], ['b1', 'a.'], ['c', '3.'],
    ]);
  });

  it('undo takes the whole paste back', async () => {
    const { history, undo, redo } = await import('prosemirror-history');
    const { madeListsSync } = await import('../../src/papyrus/editor/listMarkers');
    const zip = blankPackage();
      const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
      // The editor's own plugins for this: undo history, and keeping the numbering in step with it.
      let state = EditorState.create({ schema, doc, plugins: [history(), madeListsSync(() => model.numbering)] });
      const originalAbstracts = structuredClone(model.numbering.abstracts);
      const originalNums = structuredClone(model.numbering.nums);
      const slice = adoptPastedLists(parse('<ol start="3"><li>x</li><li>y</li></ol>'), model.numbering);
      expect(Object.keys(model.numbering.nums).length).toBeGreaterThan(Object.keys(originalNums).length);
      state = state.apply(state.tr.replaceSelection(slice));
      expect(state.doc.childCount).toBe(2);
      undo(state, (tr) => (state = state.apply(tr)));
      expect(state.doc.eq(doc)).toBe(true);
      expect(model.numbering.nums).toEqual(originalNums);
      expect(model.numbering.abstracts).toEqual(originalAbstracts);
      // Nothing of the undone paste is saved.
      const saved = await JSZip.loadAsync(await writeDocx(state.doc, model));
      expect(saved.file('word/numbering.xml')).toBeNull();
      // Redo brings the list back with its own id and definition.
      redo(state, (tr) => (state = state.apply(tr)));
      const items: any[] = [];
      state.doc.forEach((n) => items.push(n));
      const counter = new ListCounter(model.numbering);
      expect(items.map((n) => counter.next(n.attrs.numId, n.attrs.ilvl))).toEqual(['3.', '4.']);
  });

  it('plain paragraphs are left alone', () => {
    const numbering = emptyNumbering();
    const slice = adoptPastedLists(parse('<p>just text</p>'), numbering);
    expect(slice.content.firstChild!.attrs.numId).toBeNull();
    expect(Object.keys(numbering.nums)).toEqual([]);
  });
});
