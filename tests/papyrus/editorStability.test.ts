// Editor stability (audit 2026-09-25): opening, failing plugins, header/footer editing, pasted
// pictures and the run CSS cache.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { Plugin } from 'prosemirror-state';
import { DOMParser as PMDOMParser, type Node as PMNode } from 'prosemirror-model';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import {
  isSafeImageSrc, pastedColor, registerImageSources, releaseImageSources, replaceUnsafeImages, runCssCacheSize, schema,
} from '../../src/papyrus/editor/schema';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

afterEach(() => vi.restoreAllMocks());

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
async function bytes(text: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}
function mount(options = {}) {
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host, options);
  return { ed, host, done: () => { ed.destroy(); host.remove(); } };
}

describe('opening', () => {
  it('a document still opening when the editor is destroyed is dropped', async () => {
    const { ed, host } = mount();
    const opening = ed.open(await bytes('late'));
    ed.destroy();
    await opening;
    expect(ed.view).toBeNull();
    expect(host.querySelector('.ProseMirror')).toBeNull();
    host.remove();
  });

  it('only the latest of two documents opened one after the other is shown', async () => {
    const { ed, done } = mount();
    const first = ed.open(await bytes('first'));
    const second = ed.open(await bytes('second'));
    await Promise.all([first, second]);
    expect(ed.view!.state.doc.textContent).toBe('second');
    done();
  });

  it('a document that fails to open leaves the open one, and its model, as they were', async () => {
    const { ed, host, done } = mount();
    await ed.open(await bytes('kept'));
    const model = ed.model;
    const view = ed.view;
    // A list paragraph whose numbering is missing: the list markers plugin fails as it starts.
    const bad = schema.nodes.doc.create(null, schema.nodes.paragraph.create({ numId: '7' }, schema.text('new')));
    expect(() => (ed as any).setDocument(bad, { ...model, numbering: undefined })).toThrow();
    expect(ed.model).toBe(model);
    expect(ed.view).toBe(view);
    expect(ed.view!.state.doc.textContent).toBe('kept');
    expect(host.querySelectorAll('.dx-host').length).toBe(1);
    // Still usable.
    ed.run((s, d) => (d?.(s.tr.insertText('!', 1)), true));
    expect(ed.view!.state.doc.textContent).toBe('!kept');
    done();
  });
});

describe('a failing plugin', () => {
  it('leaves the document at its last good state and tells the user', async () => {
    const notices: string[] = [];
    const { ed, done } = mount({ onNotice: (m: string) => notices.push(m) });
    await ed.open(await bytes('good'));
    const view = ed.view!;
    const boom = new Plugin({
      state: {
        init: () => null,
        apply: (tr) => {
          if (tr.getMeta('boom')) throw new Error('plugin failure');
          return null;
        },
      },
    });
    view.updateState(view.state.reconfigure({ plugins: [...view.state.plugins, boom] }));
    vi.spyOn(console, 'error').mockImplementation(() => {});
    expect(() => view.dispatch(view.state.tr.insertText('bad', 1).setMeta('boom', true))).not.toThrow();
    expect(view.state.doc.textContent).toBe('good');
    expect(notices.length).toBe(1);
    // The next edit works.
    view.dispatch(view.state.tr.insertText('ok ', 1));
    expect(view.state.doc.textContent).toBe('ok good');
    done();
  });
});

describe('header/footer editing', () => {
  it('an editor that fails to start leaves no box on the page and no empty part in the model', async () => {
    const { ed, host, done } = mount();
    await ed.open(await bytes('body'));
    const parts = ed.model.headerFooters.length;
    const plugins = (ed as any).plugins;
    (ed as any).plugins = () => {
      throw new Error('cannot start');
    };
    expect(() => ed.editHeaderFooter('header')).toThrow('cannot start');
    expect(host.querySelectorAll('.dx-hf-box').length).toBe(0);
    expect(ed.model.headerFooters.length).toBe(parts);
    expect(ed.target).toBe('body');
    (ed as any).plugins = plugins;
    ed.editHeaderFooter('header'); // works again
    expect(host.querySelectorAll('.dx-hf-box').length).toBe(1);
    done();
  });
});

// A 1×1 PNG, a GIF header and an SVG.
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACw=';
const SVG = 'data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%2F%3E';

describe('pasted pictures', () => {
  const images = (html: string) => {
    const div = document.createElement('div');
    div.innerHTML = html;
    const found: string[] = [];
    PMDOMParser.fromSchema(schema).parse(div).descendants((n: PMNode) => {
      if (n.type.name === 'image') found.push(n.attrs.src);
    });
    return found;
  };

  it('only pictures that decode as PNG, JPEG, GIF, BMP, WebP or SVG are kept', () => {
    expect(images(`<p><img src="${PNG}"><img src="${GIF}"><img src="${SVG}"></p>`)).toEqual([PNG, GIF, SVG]);
    for (const src of [
      'https://evil.example/pixel.png',
      'javascript:alert(1)',
      'data:text/html;base64,PHNjcmlwdD4=',
      'data:image/png;base64,@@@@',
      'data:image/png;base64,aGVsbG8gd29ybGQ=', // "hello world", not a PNG
      'data:image/emf;base64,AQAAAGwAAAA=',
      '//evil.example/x.gif',
    ]) {
      expect(isSafeImageSrc(src), src).toBe(false);
      expect(images(`<p>a<img src="${src}">b</p>`), src).toEqual([]);
    }
  });

  it('a picture of an open document is kept whatever its format', () => {
    const emf = 'data:image/emf;base64,AQAAAGwAAAA=';
    const owner = {};
    registerImageSources(owner, [emf, PNG]);
    expect(images(`<p><img src="${emf}"></p>`)).toEqual([emf]);
    releaseImageSources(owner);
    expect(images(`<p><img src="${emf}"></p>`)).toEqual([]);
  });

  it('a refused picture becomes its alt text when pasted', async () => {
    expect(replaceUnsafeImages(`<p>a<img src="https://evil.example/x.png" alt="公司 <標誌>">b<img src="${PNG}" alt="ok"></p>`))
      .toBe(`<p>a公司 &lt;標誌&gt;b<img src="${PNG}" alt="ok"></p>`);
    const { ed, done } = mount();
    await ed.open(await bytes('x'));
    ed.view!.pasteHTML('<p>before<img src="https://evil.example/x.png" alt="logo">after</p>', new Event('paste') as ClipboardEvent);
    const doc = ed.view!.state.doc;
    let pictures = 0;
    doc.descendants((n) => { if (n.type.name === 'image') pictures++; });
    expect(pictures).toBe(0);
    expect(doc.textContent).toContain('beforelogoafter');
    done();
  });
});

describe('run CSS cache', () => {
  it('stays bounded whatever is rendered', () => {
    const serializer = (schema.marks.run.spec.toDOM as (m: unknown, inline: boolean) => unknown);
    for (let i = 0; i < 6000; i++) serializer(schema.marks.run.create({ rPr: `<w:rPr><w:spacing w:val="${i}"/></w:rPr>` }), true);
    expect(runCssCacheSize()).toBeLessThanOrEqual(4000);
    expect(runCssCacheSize()).toBeGreaterThan(0);
  });
});

describe('pasted colours', () => {
  const marks = (html: string) => {
    const div = document.createElement('div');
    div.innerHTML = html;
    const found: string[] = [];
    PMDOMParser.fromSchema(schema).parse(div).descendants((n: PMNode) => {
      for (const m of n.marks) if (m.type.name === 'color' || m.type.name === 'highlight') found.push(`${m.type.name}:${m.attrs.color}`);
    });
    return found;
  };

  it('keep hex, rgb() and Word highlight names as #rrggbb, like colours read from a file', () => {
    expect(pastedColor('rgb(31, 78, 121)')).toBe('#1f4e79');
    expect(pastedColor('#F00')).toBe('#ff0000');
    expect(pastedColor('DarkBlue')).toBe('#000080');
    expect(marks('<p><span style="color: rgb(31, 78, 121)">a</span><mark>b</mark><mark style="background-color: #00ff00">c</mark></p>'))
      .toEqual(['color:#1f4e79', 'highlight:#ffff00', 'highlight:#00ff00']);
  });

  it('drop anything else', () => {
    for (const v of ['transparent', 'inherit', 'currentcolor', 'orange', 'rgba(0, 0, 0, 0)', 'windowtext', 'expression(alert(1))', '#12345', '']) {
      expect(pastedColor(v), v).toBeNull();
    }
    expect(marks('<p><span style="color: transparent">a</span><mark style="background-color: inherit">b</mark></p>')).toEqual([]);
  });
});

describe('what the docx side reports', () => {
  it('a document with an unreadable optional part says so once when it opens', async () => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>`);
    zip.file('word/settings.xml', '<w:settings><broken');
    const notices: string[] = [];
    const { ed, done } = mount({ onNotice: (m: string) => notices.push(m) });
    await ed.open(await zip.generateAsync({ type: 'uint8array' }));
    expect(ed.model.brokenParts).toContain('word/settings.xml');
    expect(notices).toEqual(['這份文件有部分內容格式錯誤（文件設定），已略過顯示，存檔時會原樣保留。']);
    done();
  });

  it('a save that could not write a picture still succeeds, and tells the user', async () => {
    const notices: string[] = [];
    const { ed, done } = mount({ onNotice: (m: string) => notices.push(m) });
    await ed.open(await bytes('pic'));
    const img = schema.nodes.image.create({ src: 'data:image/png;base64,bm90IGEgcGljdHVyZQ==', width: 10, height: 10 });
    ed.view!.dispatch(ed.view!.state.tr.insert(1, img));
    const blob = await ed.save();
    expect(blob.size).toBeGreaterThan(0);
    expect(ed.saveWarnings.map((w) => w.kind)).toEqual(['image']);
    expect(notices).toEqual(['已存檔，但1 張圖片無法寫入，已保留原圖或略過。']);
    done();
  });
});
