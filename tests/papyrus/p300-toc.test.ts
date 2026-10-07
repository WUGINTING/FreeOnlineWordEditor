// persona-300: 參考資料 › 目錄 inserts a table of contents as Word writes one (a TOC field from the
// headings 1-3, 目錄 1-3 styles, _Toc bookmarks, links, PAGEREF page numbers, a right tab with a
// dot leader), and 「更新目錄」 rebuilds its entries from the headings, as Word's F9 does.
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { docxOf, domStubs } from './p300Helpers';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { NO_ENTRIES, tocFields } from '../../src/papyrus/editor/toc';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';

domStubs();

const para = (t: string, style?: string) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const BODY =
  para('封面') +
  para('第一章 總則', 'Heading1') +
  para('內文一') +
  para('第一節 目的', 'Heading2') +
  para('一、範圍', 'Heading3') +
  para('（一）細目', 'Heading4') +
  para('第二章 附則', 'Heading1');

let editor: DocxEditor | null = null;
let notices: string[] = [];
afterEach(() => {
  editor?.destroy();
  editor = null;
  notices = [];
});

async function open(bytes: Uint8Array | Promise<Uint8Array>) {
  const host = document.createElement('div');
  document.body.append(host);
  editor = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
  await editor.open(await bytes);
  return editor;
}

/** Puts the cursor in the paragraph with this text. */
function cursorIn(ed: DocxEditor, text: string, atEnd = false) {
  const view = ed.view!;
  let at = -1;
  view.state.doc.descendants((n, pos) => {
    if (at < 0 && n.type.name === 'paragraph' && n.textContent === text) at = atEnd ? pos + 1 + n.content.size : pos + 1;
    return at < 0;
  });
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
}

async function saved(ed: DocxEditor) {
  const zip = await JSZip.loadAsync(await ed.save());
  return { zip, doc: await zip.file('word/document.xml')!.async('string'), styles: await zip.file('word/styles.xml')!.async('string') };
}

/** The entries of the saved table of contents: [style, text before the tab, anchor]. */
function entriesOf(xml: string): [string, string, string][] {
  const out: [string, string, string][] = [];
  for (const p of xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []) {
    const style = /<w:pStyle w:val="([^"]+)"/.exec(p)?.[1] ?? '';
    const anchor = /<w:hyperlink w:anchor="([^"]+)"/.exec(p)?.[1];
    if (!anchor) continue;
    const texts = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]);
    out.push([style, texts.slice(0, -1).join(''), anchor]);
  }
  return out;
}

/** Bookmark name → the text of the paragraph it is in. */
function bookmarksOf(xml: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const p of xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []) {
    for (const m of p.matchAll(/<w:bookmarkStart w:id="\d+" w:name="([^"]+)"\/>/g)) {
      out.set(m[1], [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''));
    }
  }
  return out;
}

describe('參考資料 › 目錄 (insert)', () => {
  it('writes a TOC field from the headings 1-3, as Word does', async () => {
    const ed = await open(docxOf(BODY));
    cursorIn(ed, '封面', true);
    expect(ed.insertTableOfContents()).toBe(true);
    expect(ed.tableOfContentsCount()).toBe(1);
    const { doc, styles } = await saved(ed);
    expect(doc).toContain('<w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText>');
    const entries = entriesOf(doc);
    expect(entries.map(([s, t]) => [s, t])).toEqual([
      ['TOC1', '第一章 總則'],
      ['TOC2', '第一節 目的'],
      ['TOC3', '一、範圍'],
      ['TOC1', '第二章 附則'],
    ]);
    // Each entry links to a _Toc bookmark around its heading, and its page number is a PAGEREF to it.
    const marks = bookmarksOf(doc);
    for (const [, text, anchor] of entries) {
      expect(anchor).toMatch(/^_Toc\d{9}$/);
      expect(marks.get(anchor)).toBe(text);
      expect(doc).toContain(`<w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText>`);
    }
    // The right tab with a dot leader at the right margin (A4, margins 1800: 8306 twips; Word writes 8296).
    expect(doc).toContain('<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs>');
    // Title above, the field's end in a paragraph of its own after the entries, before 內文.
    expect(doc).toMatch(/<w:t>封面<\/w:t><\/w:r><\/w:p><w:p><w:pPr><w:pStyle w:val="TOCHeading"\/><\/w:pPr><w:r><w:t>目錄<\/w:t>/);
    expect(doc).toMatch(/<w:p><w:r><w:fldChar w:fldCharType="end"\/><\/w:r><\/w:p><w:p><w:pPr><w:pStyle w:val="Heading1"\/>/);
    // The styles the file lacked, as Word defines them (names "toc 1" … so Word knows them).
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>');
    expect(styles).toContain('w:styleId="TOC2"><w:name w:val="toc 2"/>');
    expect(styles).toContain('<w:ind w:leftChars="200" w:left="480"/>');
    expect(styles).toContain('w:styleId="TOC3"><w:name w:val="toc 3"/>');
    expect(styles).toContain('w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/>');
    expect(styles.match(/w:styleId="TOC1"/g)).toHaveLength(1);
  });

  it('reads back as the same table of contents, and saves again without changes', async () => {
    const ed = await open(docxOf(BODY));
    cursorIn(ed, '封面', true);
    ed.insertTableOfContents();
    const first = await (await JSZip.loadAsync(await ed.save())).generateAsync({ type: 'uint8array' });
    const { doc, model } = await readDocx(first);
    expect(tocFields(doc)).toHaveLength(1);
    const again = await JSZip.loadAsync(await writeDocx(doc, model));
    const once = await JSZip.loadAsync(first);
    expect(await again.file('word/document.xml')!.async('string')).toBe(await once.file('word/document.xml')!.async('string'));
    expect(await again.file('word/styles.xml')!.async('string')).toBe(await once.file('word/styles.xml')!.async('string'));
  });

  it('uses the file’s own 目錄 styles when it has them (by Word’s name), and adds none', async () => {
    const zh = async () => {
      const zip = await JSZip.loadAsync(await docxOf(BODY));
      const styles = await zip.file('word/styles.xml')!.async('string');
      zip.file(
        'word/styles.xml',
        styles.replace(
          '</w:styles>',
          '<w:style w:type="paragraph" w:styleId="11"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>' +
            '<w:style w:type="paragraph" w:styleId="21"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:leftChars="200" w:left="480"/></w:pPr></w:style>' +
            '<w:style w:type="paragraph" w:styleId="31"><w:name w:val="toc 3"/><w:basedOn w:val="Normal"/></w:style>' +
            '<w:style w:type="paragraph" w:styleId="a9"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/></w:style></w:styles>',
        ),
      );
      return zip.generateAsync({ type: 'uint8array' });
    };
    const bytes = await zh();
    const ed = await open(bytes);
    cursorIn(ed, '封面', true);
    ed.insertTableOfContents();
    const { doc, styles } = await saved(ed);
    expect(entriesOf(doc).map(([s]) => s)).toEqual(['11', '21', '31', '11']);
    expect(doc).toContain('<w:pStyle w:val="a9"/>');
    expect(styles).toBe(await (await JSZip.loadAsync(bytes)).file('word/styles.xml')!.async('string'));
  });

  it('with no headings: Word’s 「找不到目錄項目。」', async () => {
    const ed = await open(docxOf(para('只有內文')));
    cursorIn(ed, '只有內文', true);
    ed.insertTableOfContents();
    const { doc } = await saved(ed);
    expect(doc).toContain(`<w:t>${NO_ENTRIES}</w:t>`);
    expect(doc).toContain('TOC \\o &quot;1-3&quot; \\h \\z \\u');
  });

  it('is one undo step; refused in a table and while tracking changes', async () => {
    const ed = await open(docxOf(BODY + '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'));
    const before = ed.view!.state.doc;
    cursorIn(ed, '封面', true);
    ed.insertTableOfContents();
    expect(ed.view!.state.doc.eq(before)).toBe(false);
    ed.undo();
    expect(ed.view!.state.doc.eq(before)).toBe(true);
    expect(ed.isModified()).toBe(false);

    cursorIn(ed, '格');
    expect(ed.insertTableOfContents()).toBe(false);
    expect(notices.at(-1)).toBe('目錄只能插入在正文的段落中（不能在表格內）。');

    cursorIn(ed, '封面', true);
    ed.setTrackChanges(true);
    expect(ed.insertTableOfContents()).toBe(false);
    expect(notices.at(-1)).toBe('追蹤修訂開啟時，無法插入或重建目錄，請先關閉「追蹤修訂」。');
    expect(ed.view!.state.doc.eq(before)).toBe(true);
  });

  it('an empty paragraph at the cursor becomes the table of contents', async () => {
    const ed = await open(docxOf('<w:p/>' + para('第一章', 'Heading1')));
    cursorIn(ed, '');
    ed.insertTableOfContents();
    const doc = ed.view!.state.doc;
    expect(doc.child(0).textContent).toBe('目錄');
    expect(doc.lastChild!.textContent).toBe('第一章');
  });
});

describe('參考資料 › 更新目錄 (rebuild)', () => {
  async function withToc() {
    const ed = await open(docxOf(BODY));
    cursorIn(ed, '封面', true);
    ed.insertTableOfContents();
    return ed;
  }
  const headingTexts = async (ed: DocxEditor) => entriesOf((await saved(ed)).doc).map(([s, t]) => `${s}:${t}`);

  it('adds, removes and renames entries as the headings changed; old bookmarks stay', async () => {
    const ed = await withToc();
    const firstAnchors = entriesOf((await saved(ed)).doc).map(([, , a]) => a);
    const view = ed.view!;
    // Rename 第一節, delete 第二章, add a new heading 2 at the end.
    let tr = view.state.tr;
    view.state.doc.descendants((n, pos) => {
      if (n.type.name === 'paragraph' && n.textContent === '第一節 目的') tr = tr.insertText('（修正）', pos + 1 + n.content.size);
      return true;
    });
    view.dispatch(tr);
    view.state.doc.forEach((n, pos) => {
      if (n.textContent === '第二章 附則') view.dispatch(view.state.tr.delete(pos, pos + n.nodeSize));
    });
    const end = view.state.doc.content.size;
    view.dispatch(view.state.tr.insert(end, view.state.schema.nodes.paragraph.create({ styleId: 'Heading2' }, view.state.schema.text('第二節 新增'))));

    const r = ed.updateTableOfContents();
    expect(r.rebuilt).toBe(true);
    expect(await headingTexts(ed)).toEqual(['TOC1:第一章 總則', 'TOC2:第一節 目的（修正）', 'TOC3:一、範圍', 'TOC2:第二節 新增']);
    const anchors = entriesOf((await saved(ed)).doc).map(([, , a]) => a);
    expect(anchors.slice(0, 3)).toEqual(firstAnchors.slice(0, 3));
    expect(bookmarksOf((await saved(ed)).doc).get(anchors[3])).toBe('第二節 新增');
    // Exactly one field, still starting and ending as before.
    expect(tocFields(ed.view!.state.doc)).toHaveLength(1);
    // One undo step back to before the update.
    ed.undo();
    expect(await headingTexts(ed)).toEqual(['TOC1:第一章 總則', 'TOC2:第一節 目的', 'TOC3:一、範圍', 'TOC1:第二章 附則']);
  });

  it('a table of contents that lists the headings as they are is left as it is', async () => {
    const ed = await withToc();
    const before = ed.view!.state.doc;
    const r = ed.updateTableOfContents();
    expect(r.rebuilt).toBe(false);
    expect(ed.view!.state.doc).toBe(before);
  });

  it('a Word table of contents (in its content control, Word’s ids) keeps its start, end and looks', async () => {
    const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const tocPara = (style: string, text: string, bm: string, first = false) =>
      `<w:p w14:paraId="1A2B3C${style}" ${W14}><w:pPr><w:pStyle w:val="${style}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs><w:rPr><w:noProof/></w:rPr></w:pPr>` +
      (first ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : '') +
      `<w:hyperlink w:anchor="${bm}" w:history="1"><w:r w:rsidRPr="00AB12CD"><w:rPr><w:rStyle w:val="a3"/><w:noProof/></w:rPr><w:t>${text}</w:t></w:r>` +
      `<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF ${bm} \\h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r>` +
      `<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r>` +
      '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>';
    const heading = (text: string, style: string, bm: string, id: number) =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:bookmarkStart w:id="${id}" w:name="${bm}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
    const sdt =
      '<w:sdt><w:sdtPr><w:id w:val="-1"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
      '<w:p><w:pPr><w:pStyle w:val="a9"/></w:pPr><w:r><w:t>目錄</w:t></w:r></w:p>' +
      tocPara('11', '壹、計畫緣起', '_Toc100000001', true) +
      tocPara('21', '一、背景', '_Toc100000002') +
      '<w:p><w:r><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
      '</w:sdtContent></w:sdt>';
    const styles =
      '<w:style w:type="paragraph" w:styleId="11"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="21"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="a9"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>';
    const zip = await JSZip.loadAsync(
      await docxOf(sdt + heading('壹、計畫緣起', 'Heading1', '_Toc100000001', 1) + heading('一、背景', 'Heading2', '_Toc100000002', 2) + para('二、目標', 'Heading2')),
    );
    const s = await zip.file('word/styles.xml')!.async('string');
    zip.file('word/styles.xml', s.replace('</w:styles>', styles + '</w:styles>'));
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const ed = await open(bytes);
    // Untouched, it saves the table of contents as it was.
    const plain = (await saved(ed)).doc;
    expect(plain).toContain('<w:sdtContent><w:p><w:pPr><w:pStyle w:val="a9"/></w:pPr><w:r><w:t>目錄</w:t></w:r></w:p><w:p w14:paraId="1A2B3C11"');

    expect(ed.updateTableOfContents().rebuilt).toBe(true);
    const { doc } = await saved(ed);
    // Still in its content control, with Word's title, start and end paragraph.
    expect(doc).toContain('<w:sdtContent><w:p><w:pPr><w:pStyle w:val="a9"/></w:pPr><w:r><w:t>目錄</w:t></w:r></w:p>');
    expect(doc).toContain('<w:p><w:r><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>');
    expect(doc.match(/TOC \\o (&quot;|")1-3(&quot;|")/g)).toHaveLength(1);
    // Entries in the file's own styles and tab position; the new heading added with a new bookmark.
    const entries = entriesOf(doc);
    expect(entries.map(([st, t]) => `${st}:${t}`)).toEqual(['11:壹、計畫緣起', '21:一、背景', '21:二、目標']);
    expect(entries.slice(0, 2).map(([, , a]) => a)).toEqual(['_Toc100000001', '_Toc100000002']);
    expect(doc.match(/w:leader="dot" w:pos="8296"/g)).toHaveLength(3);
    const bm = entries[2][2];
    expect(doc).toMatch(new RegExp(`<w:bookmarkStart w:id="3" w:name="${bm}"/><w:r><w:t>二、目標</w:t></w:r><w:bookmarkEnd w:id="3"/>`));
  });
});

describe('untouched documents', () => {
  it('a document without a table of contents saves its styles.xml byte for byte', async () => {
    const bytes = await docxOf(BODY);
    const { doc, model } = await readDocx(bytes);
    const out = await JSZip.loadAsync(await writeDocx(doc, model));
    expect(await out.file('word/styles.xml')!.async('string')).toBe(await (await JSZip.loadAsync(bytes)).file('word/styles.xml')!.async('string'));
  });

  it('a file that uses a TOC1 it never defined keeps its styles.xml (Word shows it so), edited or not', async () => {
    const bytes = await docxOf(para('舊目錄項目', 'TOC1') + para('第一章', 'Heading1'));
    const original = await (await JSZip.loadAsync(bytes)).file('word/styles.xml')!.async('string');
    const ed = await open(bytes);
    expect((await saved(ed)).styles).toBe(original);
    ed.view!.dispatch(ed.view!.state.tr.insertText('改', 1));
    expect((await saved(ed)).styles).toBe(original);
    // A table of contents made here uses them: then they are added.
    cursorIn(ed, '第一章');
    ed.insertTableOfContents();
    expect((await saved(ed)).styles).toContain('w:styleId="TOC1"><w:name w:val="toc 1"/>');
  });
});

describe('other TOC fields (review)', () => {
  const field = (instr: string, result: string) =>
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const r = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  /** A 圖目錄 (Word's 插入圖表目錄) of one caption, and the caption with its SEQ field. */
  const FIGURES =
    para('第一章 總則', 'Heading1') +
    `<w:p>${field('TOC \\h \\z \\c &quot;圖&quot;', `<w:hyperlink w:anchor="_Toc200000001" w:history="1">${r('圖 1 組織圖')}</w:hyperlink>`)}</w:p>` +
    `<w:p><w:bookmarkStart w:id="1" w:name="_Toc200000001"/>${r('圖 ')}${field('SEQ 圖 \\* ARABIC', r('1'))}${r(' 組織圖')}<w:bookmarkEnd w:id="1"/></w:p>` +
    para('第二章 附則', 'Heading1');

  const paraTexts = (ed: DocxEditor) => {
    const out: string[] = [];
    ed.view!.state.doc.forEach((n) => out.push(n.textContent));
    return out;
  };

  it('isHeadingToc: only a table of contents of the headings is rebuilt', async () => {
    const { isHeadingToc } = await import('../../src/papyrus/editor/toc');
    expect(isHeadingToc('TOC \\o "1-3" \\h \\z \\u')).toBe(true);
    expect(isHeadingToc(' TOC \\u \\h')).toBe(true);
    expect(isHeadingToc('TOC \\h \\z \\c "圖"')).toBe(false);
    expect(isHeadingToc('TOC \\h \\z \\a "表"')).toBe(false);
    expect(isHeadingToc('TOC \\f \\h')).toBe(false);
    expect(isHeadingToc('TOC \\o "1-3" \\h \\t "標題 A,1"')).toBe(false);
    expect(isHeadingToc('TOC \\o "1-3" \\b 第一篇')).toBe(false);
    // A switch letter inside a quoted argument is not a switch.
    expect(isHeadingToc('TOC \\o "1-3" \\p "\\c"')).toBe(true);
  });

  it('更新目錄 leaves a 圖目錄 as it is (not replaced by the headings)', async () => {
    const ed = await open(docxOf(FIGURES));
    const before = paraTexts(ed);
    expect(ed.tableOfContentsCount()).toBe(0);
    const r = ed.updateTableOfContents();
    expect(r.rebuilt).toBe(false);
    expect(paraTexts(ed)).toEqual(before);
    expect(paraTexts(ed)[1]).toBe('圖 1 組織圖');
  });

  it('插入目錄 with only a 圖目錄 in the document inserts a table of contents', async () => {
    const ed = await open(docxOf(para('封面') + FIGURES));
    cursorIn(ed, '封面', true);
    expect(ed.insertTableOfContents()).toBe(true);
    expect(notices).not.toContain('這份文件已有目錄，已依目前的標題更新目錄。');
    expect(ed.tableOfContentsCount()).toBe(1);
    const texts = paraTexts(ed);
    expect(texts.slice(0, 5)).toEqual(['封面', '目錄', '第一章 總則1', '第二章 附則1', '']);
    // The 圖目錄 is still there, untouched.
    expect(texts).toContain('圖 1 組織圖');
    expect((await saved(ed)).doc).toContain('TOC \\h \\z \\c &quot;圖&quot;');
  });

  it('a TOC of other styles (\\t) keeps its entries', async () => {
    const body =
      `<w:p>${field('TOC \\o &quot;1-3&quot; \\h \\z \\t &quot;附錄標題,1&quot;', r('舊的項目'))}</w:p>` + para('第一章 總則', 'Heading1');
    const ed = await open(docxOf(body));
    expect(ed.updateTableOfContents().rebuilt).toBe(false);
    expect(paraTexts(ed)[0]).toBe('舊的項目');
  });

  it('two tables of contents are both rebuilt in place (positions found again after each)', async () => {
    const toc = (levels: string, old: string) => `<w:p>${field(`TOC \\o &quot;${levels}&quot; \\h \\z \\u`, r(old))}</w:p>`;
    const body =
      para('甲、前言', 'Heading1') + toc('1-1', '舊目錄一') + para('乙、本文', 'Heading1') + toc('1-3', '舊目錄二') + para('一、細節', 'Heading2') + para('結語');
    const ed = await open(docxOf(body));
    const r1 = ed.updateTableOfContents();
    expect(r1.rebuilt).toBe(true);
    const doc = ed.view!.state.doc;
    const fields = tocFields(doc);
    expect(fields).toHaveLength(2);
    // (An entry's text is its heading's and the page number; tabs have no text.)
    expect(paraTexts(ed)).toEqual(['甲、前言', '甲、前言1', '乙、本文1', '乙、本文', '甲、前言1', '乙、本文1', '一、細節1', '一、細節', '結語']);
    // The same bookmark per heading in both, each around its heading.
    const { doc: xml } = await saved(ed);
    const anchors = entriesOf(xml).map(([, t, a]) => `${t}>${a}`);
    expect(anchors).toHaveLength(5);
    const marks = bookmarksOf(xml);
    for (const [, t, a] of entriesOf(xml)) expect(marks.get(a)).toBe(t);
    expect(new Set(entriesOf(xml).map(([, , a]) => a)).size).toBe(3);
    expect(xml).not.toContain('舊目錄');
    // Updating again changes nothing.
    expect(ed.updateTableOfContents().rebuilt).toBe(false);
  });
});
