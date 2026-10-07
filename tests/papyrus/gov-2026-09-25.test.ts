// Government-user evaluation, second batch (2026-09-25): GOV-ISSUE-010/011/012,
// GOV-FINDING-007 and GOV-FINDING-012. Each case reproduces the reported problem first.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick, ref, type App } from 'vue';
import { NodeSelection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { findMatches } from '../../src/papyrus/editor/search';
import type { DocxEditor as DocxEditorType, EditorSnapshot } from '../../src/papyrus/editor/core';
import { loadSfc } from './sfc';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const PG = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (text: string) => `<w:p>${r(text)}</w:p>`;
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

/** A .docx with `body`, and a header / footer (word/header1.xml, word/footer1.xml) when given. */
async function docx(body: string, parts: { header?: string; footer?: string; extra?: (zip: JSZip) => Promise<void> | void } = {}): Promise<Uint8Array> {
  const zip = blankPackage();
  let refs = '';
  let rels = '';
  let types = '';
  for (const kind of ['header', 'footer'] as const) {
    const text = parts[kind];
    if (text == null) continue;
    const tag = kind === 'header' ? 'w:hdr' : 'w:ftr';
    zip.file(`word/${kind}1.xml`, `${HEAD}<${tag} ${W}>${p(text)}</${tag}>`);
    refs += `<w:${kind}Reference w:type="default" r:id="rId${kind}1"/>`;
    rels += `<Relationship Id="rId${kind}1" Type="${REL}/${kind}" Target="${kind}1.xml"/>`;
    types += `<Override PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`;
  }
  zip.file('word/document.xml', `${HEAD}<w:document ${W}><w:body>${body}<w:sectPr>${refs}${PG}</w:sectPr></w:body></w:document>`);
  const docRels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', docRels.replace('</Relationships>', rels + '</Relationships>'));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', types + '</Types>'));
  await parts.extra?.(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

/** The GOV-118 notice: six 2025 in the title, body and a table; one in the header, one in the footer. */
const GOV118_BODY =
  p('TEST118_2025 校務通知') +
  p('請各單位於 2025 年 12 月 20 日前回報。') +
  p('TEST118_BODY_2025 說明') +
  '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr><w:tc>${p('TEST118_TABLE_2025')}</w:tc><w:tc>${p('期限 2025 年 12 月 20 日')}</w:tc></w:tr></w:tbl>` +
  p('附件：2025 年度校務準備清單');
const gov118 = () => docx(GOV118_BODY, { header: 'TEST118_HEADER_2025 教育局', footer: 'TEST118_FOOTER_2025 本文件為合成測試資料' });

/** Every text of the saved package's body, header1 and footer1. */
async function partsText(bytes: Uint8Array | Blob) {
  const zip = await JSZip.loadAsync(bytes);
  const text = async (name: string) => ((await zip.file(name)?.async('string')) ?? '').replace(/<[^>]+>/g, '');
  return { body: await text('word/document.xml'), header: await text('word/header1.xml'), footer: await text('word/footer1.xml') };
}

let cleanup: (() => void)[] = [];
afterEach(() => {
  for (const f of cleanup.splice(0)) f();
  vi.restoreAllMocks();
});

async function openEditor(bytes: Uint8Array | Blob, options: Record<string, unknown> = {}) {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const host = document.createElement('div');
  document.body.append(host);
  const snapshot = ref<EditorSnapshot | null>(null);
  const notices: string[] = [];
  const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s), onNotice: (m) => notices.push(m), ...options });
  await editor.open(bytes);
  cleanup.push(() => { editor.destroy(); host.remove(); });
  return { editor, snapshot, notices, host };
}

/** The find panel next to an editor on `bytes`. */
async function findPanel(bytes: Uint8Array, replace = true) {
  const d = await openEditor(bytes);
  const Panel = await loadSfc('src/papyrus/vue/FindReplace.vue');
  const panelHost = document.createElement('div');
  document.body.append(panelHost);
  const app: App = createApp({ render: () => h(Panel, { editor: d.editor, snapshot: d.snapshot.value, replace, canReplace: true }) });
  app.mount(panelHost);
  cleanup.unshift(() => { app.unmount(); panelHost.remove(); });
  await nextTick();
  const box = (label: string) => panelHost.querySelector(`input[aria-label="${label}"]`) as HTMLInputElement;
  const type = async (label: string, value: string) => {
    const el = box(label);
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
    await new Promise((res) => setTimeout(res, 180));
    d.snapshot.value = d.editor.snapshot();
    await nextTick();
  };
  const button = (text: string | RegExp) =>
    [...panelHost.querySelectorAll('button')].find((b) => (typeof text === 'string' ? b.textContent?.trim() === text : text.test(b.textContent ?? ''))) as HTMLButtonElement | undefined;
  const click = async (el: HTMLElement | undefined | null) => {
    expect(el).toBeTruthy();
    el!.click();
    await nextTick();
    d.snapshot.value = d.editor.snapshot();
    await nextTick();
  };
  const count = () => panelHost.querySelector('.dx-count')?.textContent?.trim() ?? '';
  const message = () => panelHost.querySelector('.dx-find-msg')?.textContent?.trim() ?? '';
  return { ...d, panelHost, box, type, button, click, count, message };
}

/** Matches of `query` left in the body, the headers and the footers of a saved file. */
async function remaining(bytes: Uint8Array | Blob, query: string) {
  const { doc, model } = await readDocx(bytes);
  return {
    body: findMatches(doc, query).length,
    headers: model.headerFooters.filter((h) => h.kind === 'header').reduce((n, h) => n + findMatches(h.doc, query).length, 0),
    footers: model.headerFooters.filter((h) => h.kind === 'footer').reduce((n, h) => n + findMatches(h.doc, query).length, 0),
  };
}

describe('GOV-ISSUE-012 find / replace covers headers and footers', () => {
  it('counts the header and footer hits (each part once) and says what is searched', async () => {
    const f = await findPanel(await gov118());
    await f.type('尋找', '2025');
    // Before: 「1 / 6」 — the header and footer were not searched.
    expect(f.count()).toBe('1 / 8');
    expect(f.panelHost.textContent).toContain('正文 6、頁首 1、頁尾 1');
    const scope = f.panelHost.querySelector('.dx-find-scope')?.textContent ?? '';
    expect(scope).toContain('頁首');
    expect(scope).toContain('頁尾');
    expect(scope).toContain('文字方塊');
  });

  it('next / previous step into the header and footer, which open for editing with the hit selected', async () => {
    const f = await findPanel(await gov118());
    await f.type('尋找', '2025');
    const next = f.panelHost.querySelector('button[aria-label="下一個"]') as HTMLButtonElement;
    for (let i = 0; i < 7; i++) await f.click(next); // selects body hits 1..6, then the header
    expect(f.editor.target).toBe('header');
    const hv = f.editor.activeView!;
    expect(hv.state.doc.textBetween(hv.state.selection.from, hv.state.selection.to)).toBe('2025');
    expect(f.count()).toBe('7 / 8');
    await f.click(next);
    expect(f.editor.target).toBe('footer');
    expect(f.count()).toBe('8 / 8');
    await f.click(next); // wraps back to the body
    expect(f.editor.target).toBe('body');
    expect(f.count()).toBe('1 / 8');
    const prev = f.panelHost.querySelector('button[aria-label="上一個"]') as HTMLButtonElement;
    await f.click(prev);
    expect(f.editor.target).toBe('footer');
    expect(f.count()).toBe('8 / 8');
  });

  it('replace all changes the body, header and footer, reports where, and nothing is left after save and reopen', async () => {
    const f = await findPanel(await gov118());
    await f.type('尋找', '2025');
    await f.type('取代為', '2026');
    await f.click(f.button('全部取代'));
    await f.click(f.button(/^全部取代（8 處）$/));
    // Before: 「已取代 6 處」 and the header / footer still said 2025.
    expect(f.message()).toBe('已取代 8 處（正文 6、頁首 1、頁尾 1）');
    const saved = await f.editor.save();
    expect(await remaining(saved, '2025')).toEqual({ body: 0, headers: 0, footers: 0 });
    const text = await partsText(saved);
    expect(text.header).toContain('TEST118_HEADER_2026');
    expect(text.footer).toContain('TEST118_FOOTER_2026');
    expect(f.editor.isModified()).toBe(true);
    // Searching again finds nothing anywhere.
    await f.type('尋找', '2025');
    expect(f.count()).toBe('找不到');
  });

  it('undo: one step undoes the replace-all in the body and every header / footer part (persona-300), redo redoes it', async () => {
    const f = await findPanel(await gov118());
    await f.type('尋找', '2025');
    await f.type('取代為', '2026');
    await f.click(f.button('全部取代'));
    await f.click(f.button(/^全部取代（/));
    expect(f.editor.undo()).toBe(true);
    expect(findMatches(f.editor.view!.state.doc, '2025')).toHaveLength(6);
    const text = await partsText(await f.editor.save());
    expect(text.header).toContain('TEST118_HEADER_2025');
    expect(text.footer).toContain('TEST118_FOOTER_2025');
    expect(f.editor.redo()).toBe(true);
    const again = await partsText(await f.editor.save());
    expect(again.header).toContain('TEST118_HEADER_2026');
    expect(findMatches(f.editor.view!.state.doc, '2026')).toHaveLength(6);
    // From a part it changed (the header, opened), Ctrl+Z undoes it everywhere too.
    f.editor.editHeaderFooter('header');
    expect(f.editor.undo()).toBe(true);
    expect(f.editor.activeView!.state.doc.textContent).toContain('TEST118_HEADER_2025');
    expect(findMatches(f.editor.view!.state.doc, '2025')).toHaveLength(6);
  });

  it('a header shared by two sections is searched and replaced once', async () => {
    const body = `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdheader1"/>${PG}</w:sectPr></w:pPr>${r('A 2025')}</w:p>` + p('B 2025');
    const f = await findPanel(await docx(body, { header: 'H 2025' }));
    await f.type('尋找', '2025');
    expect(f.count()).toBe('1 / 3');
    await f.type('取代為', '2026');
    await f.click(f.button('全部取代'));
    await f.click(f.button(/^全部取代（3 處）$/));
    expect(f.message()).toBe('已取代 3 處（正文 2、頁首 1）');
    expect((await partsText(await f.editor.save())).header).toBe('H 2026');
  });
});

describe('GOV-ISSUE-010 replace all shows every match before replacing', () => {
  const CITES = p('依第十二條規定辦理。') + p('第十二條之一 另定之。') + p('參照第十二條之二及第十二條。') + p('第十二條');

  it('lists every match with its context and place; longer terms are flagged and left unchecked', async () => {
    const f = await findPanel(await docx(CITES, { header: '本規則第十二條' }));
    await f.type('尋找', '第十二條');
    await f.type('取代為', '第十三條');
    await f.click(f.button('全部取代'));
    // Before: the click replaced all six at once, 第十二條之一／之二 included.
    const review = f.panelHost.querySelector('.dx-review') as HTMLElement;
    expect(review).toBeTruthy();
    const rows = [...review.querySelectorAll('li')];
    expect(rows).toHaveLength(6);
    const boxes = rows.map((li) => li.querySelector('input[type="checkbox"]') as HTMLInputElement);
    expect(boxes.map((b) => b.checked)).toEqual([true, false, false, true, true, true]);
    expect(rows[1].textContent).toContain('第十二條之一 另定之');
    expect(rows[1].textContent).toContain('後面還有文字，可能是較長的詞');
    expect(rows[0].textContent).not.toContain('可能是較長的詞');
    expect(rows[0].textContent).toContain('正文');
    expect(rows[5].textContent).toContain('頁首');
    expect(rows[0].textContent).toContain('依');
    expect(rows[0].textContent).toContain('規定辦理');
    expect(f.button(/^全部取代（4 處）$/)).toBeTruthy();
    // Every checkbox has a name, and the list is announced as a dialog.
    expect(review.getAttribute('role')).toBe('dialog');
    for (const b of boxes) expect(b.closest('label')?.textContent?.trim()).toBeTruthy();
    // Focus goes into the list when it opens.
    expect(review.contains(document.activeElement)).toBe(true);

    await f.click(f.button(/^全部取代（4 處）$/));
    expect(f.message()).toBe('已取代 4 處（正文 3、頁首 1），略過 2 處');
    const text = await partsText(await f.editor.save());
    expect(text.body).toContain('第十二條之一');
    expect(text.body).toContain('第十二條之二');
    expect(text.body).toContain('依第十三條規定');
    expect(text.body).toContain('及第十三條。');
    expect(text.header).toBe('本規則第十三條');
  });

  it('checking a flagged match includes it; 取消 and Escape leave the document alone', async () => {
    const f = await findPanel(await docx(CITES));
    await f.type('尋找', '第十二條');
    await f.type('取代為', '第十三條');
    const before = f.editor.view!.state.doc;
    await f.click(f.button('全部取代'));
    await f.click(f.button('取消'));
    expect(f.panelHost.querySelector('.dx-review')).toBeNull();
    expect(f.editor.view!.state.doc).toBe(before);
    await f.click(f.button('全部取代'));
    const review = f.panelHost.querySelector('.dx-review') as HTMLElement;
    review.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    expect(f.panelHost.querySelector('.dx-review')).toBeNull();
    expect(f.panelHost.querySelector('[role="search"]')).toBeTruthy(); // the panel itself stays open
    expect(f.editor.view!.state.doc).toBe(before);
    await f.click(f.button('全部取代'));
    const box = f.panelHost.querySelectorAll<HTMLInputElement>('.dx-review li input[type="checkbox"]')[1];
    box.click();
    await nextTick();
    await f.click(f.button(/^全部取代（4 處）$/));
    expect(f.editor.view!.state.doc.textContent).toContain('第十三條之一');
    expect(f.editor.view!.state.doc.textContent).toContain('第十二條之二');
  });

  it('a year followed by 年 is not flagged; a number inside a longer number is', async () => {
    const f = await findPanel(await docx(p('2025年 2025 年 20251 項')));
    await f.type('尋找', '2025');
    await f.type('取代為', '2026');
    await f.click(f.button('全部取代'));
    const boxes = [...f.panelHost.querySelectorAll<HTMLInputElement>('.dx-review li input[type="checkbox"]')];
    expect(boxes.map((b) => b.checked)).toEqual([true, true, false]);
  });
});

describe('GOV-ISSUE-011 image alt text is saved', () => {
  async function imagePanel(bytes?: Uint8Array | Blob) {
    const d = await openEditor(bytes ?? (await docx('<w:p/>')));
    const Panel = await loadSfc('src/papyrus/vue/ImagePanel.vue');
    const panelHost = document.createElement('div');
    document.body.append(panelHost);
    const app = createApp({ render: () => h(Panel, { editor: d.editor, snapshot: d.snapshot.value }) });
    app.mount(panelHost);
    cleanup.unshift(() => { app.unmount(); panelHost.remove(); });
    const refresh = async () => { d.snapshot.value = d.editor.snapshot(); await nextTick(); };
    const alt = () => panelHost.querySelector('input[aria-label="替代文字"]') as HTMLInputElement;
    /** Typing without leaving the box (no change event): what the evaluator did before saving. */
    const typeAlt = async (value: string) => {
      alt().value = value;
      alt().dispatchEvent(new Event('input', { bubbles: true }));
      await refresh();
    };
    const image = () => {
      let found: any = null;
      d.editor.view!.state.doc.descendants((n, pos) => { if (n.type === schema.nodes.image) found = { node: n, pos }; });
      return found as { node: any; pos: number };
    };
    return { ...d, panelHost, refresh, alt, typeAlt, image };
  }

  function addImage(editor: DocxEditorType) {
    const tr = editor.view!.state.tr.insert(1, schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 100, height: 50, alt: '' }));
    editor.view!.dispatch(tr.setSelection(NodeSelection.create(tr.doc, 1)));
    editor.view!.dispatch(closeHistory(editor.view!.state.tr));
  }

  it('typing applies the text at once, marks the document modified and is one undo step', async () => {
    const d = await imagePanel();
    addImage(d.editor);
    d.editor.markSaved(d.editor.savePoint());
    await d.refresh();
    expect(d.editor.isModified()).toBe(false);
    await d.typeAlt('合成步');
    await d.typeAlt('合成步道與山景宣導示意圖，TEST-028');
    // Before: the image kept alt="" until the box lost focus, so a save right after typing lost it.
    expect(d.image().node.attrs.alt).toBe('合成步道與山景宣導示意圖，TEST-028');
    expect(d.editor.view!.dom.querySelector('img.dx-img')!.alt).toBe('合成步道與山景宣導示意圖，TEST-028');
    expect(d.editor.isModified()).toBe(true);
    // Still selected: the panel stays for the next keystroke.
    expect(d.alt()).toBeTruthy();
    expect(d.editor.undo()).toBe(true);
    expect(d.image().node.attrs.alt).toBe('');
    expect(d.editor.isModified()).toBe(false);
  });

  it('saves as wp:docPr/@descr, reopens with it, and clearing it saves too', async () => {
    const d = await imagePanel();
    addImage(d.editor);
    await d.refresh();
    await d.typeAlt('步道示意圖 & "logo"');
    const saved = await d.editor.save();
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<wp:docPr [^>]*descr="步道示意圖 &amp; &quot;logo&quot;"/);

    const again = await imagePanel(saved);
    expect(again.image().node.attrs.alt).toBe('步道示意圖 & "logo"');
    again.editor.view!.dispatch(again.editor.view!.state.tr.setSelection(NodeSelection.create(again.editor.view!.state.doc, again.image().pos)));
    await again.refresh();
    expect(again.alt().value).toBe('步道示意圖 & "logo"');
    await again.typeAlt('');
    expect(again.editor.isModified()).toBe(true);
    const cleared = await again.editor.save();
    const xml2 = await (await JSZip.loadAsync(cleared)).file('word/document.xml')!.async('string');
    expect(xml2).not.toContain('步道示意圖');
    const third = await openEditor(cleared);
    let alt: string | null = null;
    third.editor.view!.state.doc.descendants((n) => { if (n.type === schema.nodes.image) alt = n.attrs.alt; });
    expect(alt).toBe('');
  });

  it('is applied while 追蹤修訂 is on (Word does not track alt text) and an untouched picture keeps its XML', async () => {
    const DRAWING =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="7" name="Picture 7" descr="舊描述"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="x.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
    const bytes = await docx(`<w:p>${r('A')}<w:r>${DRAWING}</w:r></w:p><w:p><w:r>${DRAWING.replace(/id="7"/g, 'id="8"')}</w:r></w:p>`, {
      extra: async (zip) => {
        zip.file('word/media/image1.png', PNG, { base64: true });
        const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
        zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', `<Relationship Id="rIdImg" Type="${REL}/image" Target="media/image1.png"/></Relationships>`));
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        zip.file('[Content_Types].xml', ct.replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'));
      },
    });
    const d = await imagePanel(bytes);
    d.editor.setTrackChanges(true);
    const images: number[] = [];
    d.editor.view!.state.doc.descendants((n, pos) => { if (n.type === schema.nodes.image) images.push(pos); });
    d.editor.view!.dispatch(d.editor.view!.state.tr.setSelection(NodeSelection.create(d.editor.view!.state.doc, images[0])));
    await d.refresh();
    await d.typeAlt('新描述');
    expect(d.editor.view!.state.doc.nodeAt(images[0])!.attrs.alt).toBe('新描述');
    const xml = await (await JSZip.loadAsync(await d.editor.save())).file('word/document.xml')!.async('string');
    expect(xml).toContain('<wp:docPr id="7" name="Picture 7" descr="新描述"/>');
    // The second picture was not touched: its drawing is written exactly as it was read.
    expect(xml).toContain(DRAWING.replace(/id="7"/g, 'id="8"'));
  });
});

describe('GOV-FINDING-007 a non-picture chosen or dropped as a picture', () => {
  const MESSAGE = '這裡只能插入圖片（PNG、JPEG、GIF、BMP、WebP、SVG）。PDF 等附件請用公文／流程系統附加。';
  const pdf = () => new File(['%PDF-1.4\n%synthetic\n'], 'GOVTEST_012_LAND_FORM.pdf', { type: 'application/pdf' });

  it('插入圖片 with a PDF says so through onNotice, not window.alert, and changes nothing', async () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {});
    const d = await openEditor(await docx(p('採購規格')));
    const before = d.editor.view!.state.doc;
    const result = await Promise.race([
      d.editor.insertImageFile(pdf()).then(() => 'done', () => 'threw'),
      new Promise((res) => setTimeout(() => res('hung'), 300)),
    ]);
    expect(result).toBe('done');
    expect(d.notices).toEqual([MESSAGE]);
    expect(alert).not.toHaveBeenCalled();
    expect(d.editor.view!.state.doc).toBe(before);
  });

  it('dropping or pasting a PDF shows the same notice', async () => {
    const d = await openEditor(await docx(p('採購規格')));
    const view = d.editor.view!;
    const dropped = view.someProp('handleDrop', (f) => f(view, { dataTransfer: { files: [pdf()], types: ['Files'], getData: () => '' } } as any, null as any, false));
    expect(dropped).toBe(true);
    const pasted = view.someProp('handlePaste', (f) => f(view, { clipboardData: { files: [pdf()], types: ['Files'], getData: () => '' } } as any, null as any));
    expect(pasted).toBe(true);
    expect(d.notices).toEqual([MESSAGE, MESSAGE]);
    // Text pasted along with a file is pasted as usual.
    const withText = view.someProp('handlePaste', (f) => f(view, { clipboardData: { files: [pdf()], types: ['Files', 'text/plain'], getData: (t: string) => (t === 'text/plain' ? 'x' : '') } } as any, null as any));
    expect(withText).toBeFalsy();
  });
});

describe('GOV-FINDING-012 core properties title', () => {
  const coreXml = (title: string) =>
    `${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${title}<dc:creator>承辦人</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-09-01T01:02:03Z</dcterms:created></cp:coreProperties>`;
  const withCore = (title: string) =>
    docx(p('內容'), {
      extra: async (zip) => {
        zip.file('docProps/core.xml', coreXml(title));
        const rels = await zip.file('_rels/.rels')!.async('string');
        zip.file('_rels/.rels', rels.replace('</Relationships>', `<Relationship Id="rIdCore" Type="${CORE_REL}" Target="docProps/core.xml"/></Relationships>`));
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
      },
    });
  const core = async (bytes: Uint8Array | Blob) => (await JSZip.loadAsync(bytes)).file('docProps/core.xml')?.async('string');
  const TITLE = 'GOVTEST_149_首長談話稿_定稿';

  it('a generic or empty title becomes the document name; the author\'s own title stays', async () => {
    for (const generic of ['<dc:title>Word Document</dc:title>', '<dc:title/>', '<dc:title> </dc:title>', '']) {
      const { doc, model } = await readDocx(await withCore(generic));
      const xml = await core(await writeDocx(doc, model, { title: TITLE }));
      expect(xml).toContain(`<dc:title>${TITLE}</dc:title>`);
      expect(xml).toContain('<dc:creator>承辦人</dc:creator>');
      expect(xml).toContain('2026-09-01T01:02:03Z');
      expect(xml).not.toContain('Word Document');
    }
    const own = await withCore('<dc:title>局長交辦事項</dc:title>');
    const { doc, model } = await readDocx(own);
    expect(await core(await writeDocx(doc, model, { title: TITLE }))).toBe(coreXml('<dc:title>局長交辦事項</dc:title>'));
  });

  it('without the option core.xml is left exactly as it is', async () => {
    const bytes = await withCore('<dc:title>Word Document</dc:title>');
    const { doc, model } = await readDocx(bytes);
    expect(await core(await writeDocx(doc, model))).toBe(coreXml('<dc:title>Word Document</dc:title>'));
  });

  /** A package made before the template had core properties (as GOV-149's download was). */
  const noCore = () =>
    docx(p('內容'), {
      extra: async (zip) => {
        zip.remove('docProps/core.xml');
        const rels = await zip.file('_rels/.rels')!.async('string');
        zip.file('_rels/.rels', rels.replace(/<Relationship [^>]*core-properties[^>]*\/>/, ''));
        const ct = await zip.file('[Content_Types].xml')!.async('string');
        zip.file('[Content_Types].xml', ct.replace(/<Override PartName="\/docProps\/core.xml"[^>]*\/>/, ''));
      },
    });

  it('a package without core.xml gets one with the title, its relationship and content type', async () => {
    const { doc, model } = await readDocx(await noCore());
    expect(model.zip!.file('docProps/core.xml')).toBeNull();
    const out = await JSZip.loadAsync(await writeDocx(doc, model, { title: 'A & <B>' }));
    expect(await out.file('docProps/core.xml')!.async('string')).toContain('<dc:title>A &amp; &lt;B&gt;</dc:title>');
    expect(await out.file('_rels/.rels')!.async('string')).toContain(`Type="${CORE_REL}" Target="docProps/core.xml"`);
    expect(await out.file('[Content_Types].xml')!.async('string')).toContain('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
    const again = await readDocx(await writeDocx(doc, model, { title: 'A & <B>' }));
    expect(again.doc.textContent).toBe('內容');
  });

  it('the blank template records when it was created and has no generic title; DocxEditor passes the title on save', async () => {
    const zip = blankPackage();
    const xml = await zip.file('docProps/core.xml')?.async('string');
    expect(xml).toMatch(/<dcterms:created xsi:type="dcterms:W3CDTF">\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ<\/dcterms:created>/);
    expect(xml).not.toContain('Word Document');
    expect(await zip.file('_rels/.rels')!.async('string')).toContain(CORE_REL);
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/docProps/core.xml');

    const { editor } = await openEditor(await docx(p('內容')));
    expect(await core(await editor.save({ title: TITLE }))).toContain(`<dc:title>${TITLE}</dc:title>`);
    editor.setDocumentTitle('另一個標題');
    expect(await core(await editor.save())).toContain('<dc:title>另一個標題</dc:title>');
    const { editor: plain } = await openEditor(await noCore());
    expect(await core(await plain.save())).toBeUndefined();
    const created = await core(await blankPackage().generateAsync({ type: 'uint8array' }));
    const { editor: fresh } = await openEditor(await blankPackage().generateAsync({ type: 'uint8array' }));
    expect(created).toBeTruthy();
    expect(await core(await fresh.save())).not.toContain('<dc:title');
    const { editor: opt } = await openEditor(await docx(p('內容')), { documentTitle: '選項標題' });
    expect(await core(await opt.save())).toContain('<dc:title>選項標題</dc:title>');
  });
});
