import { describe, expect, it } from 'vitest';
import { DOMParser as PMDOMParser, DOMSerializer } from 'prosemirror-model';
import { schema } from '../src/papyrus/editor/schema';

function paste(html: string) {
  const host = document.createElement('div');
  host.innerHTML = `<html><body><!--StartFragment-->${html}<!--EndFragment--></body></html>`;
  const doc = PMDOMParser.fromSchema(schema).parse(host);
  const output = document.createElement('div');
  output.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
  return { doc, output };
}

describe('QA20K office clipboard safety pilot 00611–00620', () => {
  it('QA20K-00611 retains Word text colour but drops an adjacent injected CSS declaration', () => {
    const { output } = paste('<p><span style="color: rgb(31, 78, 121); background-image: url(https://evil.example/a)">合約條款</span></p>');
    expect(output.textContent).toBe('合約條款');
    expect(output.querySelector('[style]')?.getAttribute('style')).toContain('color: rgb(31, 78, 121);');
    expect(output.innerHTML).not.toMatch(/evil\.example|background-image|url\(/i);
  });

  it('QA20K-00612 rejects a pasted run colour containing a declaration breakout', () => {
    const { output } = paste('<p><span style="color: red; background-image: url(https://evil.example/b)">付款條件</span></p>');
    expect(output.textContent).toBe('付款條件');
    expect(output.innerHTML).not.toMatch(/evil\.example|background-image|url\(/i);
  });

  it('QA20K-00613 keeps pasted Latin and CJK font choices inside quoted CSS values', () => {
    const { doc, output } = paste('<p><span style="font-family: Arial, 微軟正黑體">報價 Quote</span></p>');
    const style = [...output.querySelectorAll('[style]')].map((node) => node.getAttribute('style') ?? '').join('\n');
    expect(output.textContent).toBe('報價 Quote');
    const textMarks: any[] = [];
    doc.descendants((node: any) => { if (node.isText) textMarks.push(...node.marks); });
    expect(textMarks.map((mark) => mark.attrs)).toContainEqual({ family: 'Arial', eastAsia: null });
    expect(style).toMatch(/font-family:/);
    expect(style).toContain('Arial');
    expect(style).not.toMatch(/url\(/i);
  });

  it('QA20K-00614 sanitizes quote-breaking font names pasted from a rich-text source', () => {
    const { output } = paste('<p><span style="font-family: &quot;Arial&quot;; background: url(https://evil.example/c)">通知</span></p>');
    expect(output.textContent).toBe('通知');
    expect(output.innerHTML).not.toMatch(/evil\.example|background:|url\(/i);
  });

  it('QA20K-00615 keeps a supported highlight while ignoring an injected URL property', () => {
    const { output } = paste('<p><mark style="background-color: rgb(255, 255, 0); background-image: url(https://evil.example/d)">待核對</mark></p>');
    expect(output.textContent).toBe('待核對');
    expect(output.querySelector('mark')?.getAttribute('style')).toContain('background-color: rgb(255, 255, 0);');
    expect(output.innerHTML).not.toMatch(/evil\.example|background-image|url\(/i);
  });

  it('QA20K-00616 keeps pasted cell text but discards a URL-based cell background', () => {
    const { doc, output } = paste('<table><tbody><tr><td style="background-image: url(https://evil.example/e)"><p>本月金額</p></td></tr></tbody></table>');
    const cell = doc.firstChild!.firstChild!.firstChild!;
    expect(cell.textContent).toBe('本月金額');
    expect(cell.attrs.background).toBeNull();
    expect(output.innerHTML).not.toMatch(/evil\.example|background-image|url\(/i);
  });

  it('QA20K-00617 does not turn a transparent pasted cell into black shading', () => {
    const { doc, output } = paste('<table><tbody><tr><td style="background-color: rgba(0, 0, 0, 0)"><p>未填金額</p></td></tr></tbody></table>');
    expect(doc.firstChild!.firstChild!.firstChild!.attrs.background).toBeNull();
    expect(output.textContent).toBe('未填金額');
  });

  it('QA20K-00618 ignores unsupported baseline alignment copied from a web table', () => {
    const { doc, output } = paste('<table><tbody><tr><td style="vertical-align: baseline"><p>核准</p></td></tr></tbody></table>');
    expect(doc.firstChild!.firstChild!.firstChild!.attrs.vAlign).toBeNull();
    expect(output.querySelector('td')?.getAttribute('style') ?? '').not.toContain('vertical-align');
  });

  it('QA20K-00619 retains a safe relative hyperlink from pasted meeting notes', () => {
    const { output } = paste('<p><a href="/shared/minutes/2026-09.docx">會議紀錄</a></p>');
    const link = output.querySelector('a');
    expect(link?.getAttribute('href')).toBe('/shared/minutes/2026-09.docx');
    expect(link?.getAttribute('rel')).toBe('noopener noreferrer');
    expect(link?.textContent).toBe('會議紀錄');
  });

  it('QA20K-00620 neutralizes a javascript link pasted into the document body', () => {
    const { output } = paste('<p><a href="javascript:alert(1)">開啟附件</a></p>');
    expect(output.querySelector('a')?.getAttribute('href')).toBe('#');
    expect(output.textContent).toBe('開啟附件');
  });
});
