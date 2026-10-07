// Minimal package parts for a brand-new document. Written from ECMA-376 Part 1.
import JSZip from 'jszip';
import { NS } from './xml';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES =
  XML_HEAD +
  `<Types xmlns="${NS.ct}">` +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
  '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
  '</Types>';

const ROOT_RELS =
  XML_HEAD +
  `<Relationships xmlns="${NS.rel}">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '</Relationships>';

/**
 * The core properties of a new document: when it was created, and no title (a blank Word
 * template's generic "Word Document" says nothing; the editor writes the document's name as
 * the title on save, see WriteOptions.title). GOV-FINDING-012.
 */
export function newCoreXml(created = new Date()): string {
  const when = created.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return (
    XML_HEAD +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' +
    `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${when}</dcterms:created></cp:coreProperties>`
  );
}

const DOC_RELS =
  XML_HEAD +
  `<Relationships xmlns="${NS.rel}">` +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
  '</Relationships>';

const SETTINGS =
  XML_HEAD +
  `<w:settings xmlns:w="${NS.w}"><w:defaultTabStop w:val="480"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

function heading(id: string, name: string, sz: number, before: number, after: number, outline: number | null): string {
  return (
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="${after}"/>${outline != null ? `<w:outlineLvl w:val="${outline}"/>` : ''}</w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`
  );
}

// A new document is a 公文 page (persona-300), as the built-in templates (make-templates.ps1): A4,
// 標楷體 for Chinese and Times New Roman for Latin text, 16 pt (the 文書處理手冊's body size: 主旨、
// 說明、辦法, and the templates' 公文段落 style), single line spacing, no space between paragraphs;
// margins 2.5 cm top, bottom and left, 2.0 cm right, header and footer 1.2 cm from the edge.
export const GONGWEN_DEFAULTS = {
  eastAsiaFont: '標楷體',
  latinFont: 'Times New Roman',
  sizePt: 16,
  /** twips: 2.5 cm, 2.5 cm, 2.5 cm, 2.0 cm, 1.2 cm. */
  marginTop: 1417,
  marginBottom: 1417,
  marginLeft: 1417,
  marginRight: 1134,
  header: 680,
  footer: 680,
} as const;
const G = GONGWEN_DEFAULTS;

const STYLES =
  XML_HEAD +
  `<w:styles xmlns:w="${NS.w}">` +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  `<w:rFonts w:ascii="${G.latinFont}" w:hAnsi="${G.latinFont}" w:eastAsia="${G.eastAsiaFont}" w:cs="Times New Roman"/>` +
  `<w:sz w:val="${G.sizePt * 2}"/><w:szCs w:val="${G.sizePt * 2}"/><w:lang w:val="en-US" w:eastAsia="zh-TW"/>` +
  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  heading('Title', 'Title', 56, 0, 240, null) +
  heading('Heading1', 'heading 1', 36, 360, 120, 0) +
  heading('Heading2', 'heading 2', 30, 240, 120, 1) +
  heading('Heading3', 'heading 3', 26, 240, 80, 2) +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('') +
  '</w:tblBorders></w:tblPr></w:style>' +
  '</w:styles>';

export function newDocumentXml(): string {
  return (
    XML_HEAD +
    `<w:document xmlns:w="${NS.w}"><w:body><w:p/>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${G.marginTop}" w:right="${G.marginRight}" w:bottom="${G.marginBottom}" ` +
    `w:left="${G.marginLeft}" w:header="${G.header}" w:footer="${G.footer}" w:gutter="0"/></w:sectPr>` +
    '</w:body></w:document>'
  );
}

/** A complete, valid, empty .docx package, created now (docProps/core.xml). */
export function blankPackage(): JSZip {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('docProps/core.xml', newCoreXml());
  zip.file('word/document.xml', newDocumentXml());
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/settings.xml', SETTINGS);
  return zip;
}
