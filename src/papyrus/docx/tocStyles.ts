// Word's built-in table of contents styles (目錄 1-3 and 目錄標題), for a table of contents made
// in the editor (editor/toc.ts) in a file that has none of them: added to styles.xml on save, as
// Word adds a built-in style the first time it is used. Only a style a table of contents made or
// rebuilt in the editor used (DocxModel.tocStylesUsed), that the document still uses and the file
// does not define, is added; styles.xml is left as it is otherwise (a file that refers to a TOC1
// it never defined keeps looking as it does in Word).
import type JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import type { DocxModel } from './model';
import { REL_TYPE, escapeXml } from './xml';
import type { Relationship } from './reader';

/** The styles this adds, by the id the editor gives them when the file has none (see toc.ts). */
export const TOC_STYLE_IDS = ['TOC1', 'TOC2', 'TOC3', 'TOCHeading'] as const;

/** Each style as Word (zh-TW) writes it; `normal` is the default paragraph style, `heading1` 標題 1. */
function definition(id: string, normal: string | null, heading1: string | null): string {
  const based = (b: string | null) => (b ? `<w:basedOn w:val="${escapeXml(b)}"/>` : '');
  const next = normal ? `<w:next w:val="${escapeXml(normal)}"/>` : '';
  const toc = (level: number, ind: string) =>
    `<w:style w:type="paragraph" w:styleId="TOC${level}"><w:name w:val="toc ${level}"/>${based(normal)}${next}` +
    `<w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>${ind ? `<w:pPr>${ind}</w:pPr>` : ''}</w:style>`;
  switch (id) {
    case 'TOC1':
      return toc(1, '');
    case 'TOC2':
      return toc(2, '<w:ind w:leftChars="200" w:left="480"/>');
    case 'TOC3':
      return toc(3, '<w:ind w:leftChars="400" w:left="960"/>');
    default:
      // 目錄標題: based on 標題 1, but not itself a heading (outline level 9 = body text).
      return (
        `<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/>${based(heading1 ?? normal)}${next}` +
        '<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>' +
        `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/><w:outlineLvl w:val="9"/></w:pPr>` +
        (heading1 ? '' : '<w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>') +
        '</w:style>'
      );
  }
}

/** The paragraph style ids a document uses. */
function usedStyleIds(doc: PMNode): Set<string> {
  const out = new Set<string>();
  doc.descendants((n) => {
    if (n.type.name === 'paragraph' && n.attrs.styleId) out.add(n.attrs.styleId as string);
    return n.type.name !== 'paragraph';
  });
  return out;
}

/**
 * Adds the table of contents styles the document uses and styles.xml lacks (see above). A file
 * without a styles part, or with one that was not readable (`broken`), is left alone.
 */
export async function addTocStyles(zip: JSZip, mainRels: Map<string, Relationship>, doc: PMNode, model: DocxModel, broken: string[]): Promise<void> {
  const made = model.tocStylesUsed;
  if (!made?.size) return;
  const used = usedStyleIds(doc);
  const wanted = TOC_STYLE_IDS.filter((id) => made.has(id) && used.has(id));
  if (!wanted.length) return;
  const part = [...mainRels.values()].find((r) => r.type === REL_TYPE.styles && !r.external)?.target;
  const file = part ? zip.file(part) : null;
  if (!part || !file || broken.includes(part)) return;
  const xml = await file.async('string');
  const defined = new Set([...xml.matchAll(/\bw:styleId="([^"]*)"/g)].map((m) => m[1]));
  const add = wanted.filter((id) => !defined.has(id));
  const close = xml.lastIndexOf('</w:styles>');
  if (!add.length || close < 0) return;
  const normal = model.styles?.defaultParagraph ?? (defined.has('Normal') ? 'Normal' : null);
  const heading1 = model.paragraphStyles.find((s) => s.name.toLowerCase() === 'heading 1')?.id ?? null;
  zip.file(part, xml.slice(0, close) + add.map((id) => definition(id, normal, heading1)).join('') + xml.slice(close));
}
