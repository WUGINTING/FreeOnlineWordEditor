// Which file title the online document wrote itself (persona-300, after GOV-FINDING-012).
//
// A file gets the document's name as its title (docProps/core.xml dc:title) when it has none;
// renaming the document later should change that title too, but never one the author typed in
// Word. So when the editor writes the title, it also records it in a custom document property
// (docProps/custom.xml, Word's 「摘要資訊 › 自訂」): a title still equal to that record was
// written by the system and follows the next name; any other title is the author's and stays.
import type JSZip from 'jszip';
import { escapeXml, parseXml } from './xml';
import { readRels } from './reader';

/** The custom property (Word shows it under 摘要資訊 › 自訂). */
export const TITLE_PROPERTY = '線上文件自動標題';

const REL_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CT_CUSTOM = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const NS_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
/** The format id Office gives user-defined properties. */
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export interface TitleRecord {
  /** The package path of the custom properties part (null: the package has none). */
  part: string | null;
  /** The title recorded there (null: none recorded, or the part is not readable). */
  value: string | null;
  /** The part is there but not readable XML: left alone. */
  broken: boolean;
}

/** The title the editor last wrote into this package, as recorded (see above). */
export async function readTitleRecord(zip: JSZip): Promise<TitleRecord> {
  const rel = [...(await readRels(zip, '', [])).values()].find((r) => r.type === REL_CUSTOM && !r.external);
  const file = rel ? zip.file(rel.target) : null;
  if (!rel || !file) return { part: null, value: null, broken: false };
  try {
    const doc = parseXml(await file.async('string'));
    const prop = Array.from(doc.getElementsByTagNameNS(NS_CUSTOM, 'property')).find((p) => p.getAttribute('name') === TITLE_PROPERTY);
    return { part: rel.target, value: prop ? prop.textContent ?? '' : null, broken: false };
  } catch {
    return { part: rel.target, value: null, broken: true };
  }
}

/** Whether a file's title was written by the system (it is what the record says). */
export function systemTitle(title: string | null, record: TitleRecord): boolean {
  return title != null && record.value != null && record.value.trim() === title.trim();
}

/**
 * Records `title` as the one the editor wrote: the property is changed in place (the rest of the
 * part stays byte for byte), added to the part, or the part is created with its relationship and
 * content type. Nothing changes when it already says so.
 */
export async function writeTitleRecord(zip: JSZip, title: string, record: TitleRecord, overrides: Record<string, string>): Promise<void> {
  if (record.broken || record.value === title) return;
  const value = escapeXml(title);
  if (record.part) {
    const xml = await zip.file(record.part)!.async('string');
    const doc = parseXml(xml);
    const props = Array.from(doc.getElementsByTagNameNS(NS_CUSTOM, 'property'));
    const at = props.findIndex((p) => p.getAttribute('name') === TITLE_PROPERTY);
    let out: string;
    if (at >= 0) {
      // That <property> element, rewritten with just the new text value (as Word writes it).
      const vt = props[at].firstElementChild?.prefix ?? doc.documentElement.lookupPrefix(NS_VT) ?? 'vt';
      // The same element in the text, by its place among the properties (its name may be written
      // with single quotes or character references, which the parsed XML reads the same).
      const PROPERTY = /<((?:[\w.-]+:)?property)\b([^>]*?)(\/>|>[\s\S]*?<\/\1\s*>)/g;
      if ((xml.match(PROPERTY) ?? []).length !== props.length) return;
      let k = -1;
      out = xml.replace(PROPERTY, (m, tag: string, attrs: string) =>
        ++k === at ? `<${tag}${attrs.replace(/\s*\/$/, '')}><${vt}:lpwstr>${value}</${vt}:lpwstr></${tag}>` : m,
      );
    } else {
      const pid = Math.max(1, ...props.map((p) => Number(p.getAttribute('pid')) || 0)) + 1;
      const vt = doc.documentElement.lookupPrefix(NS_VT);
      const prefix = doc.documentElement.prefix;
      const tag = prefix ? `${prefix}:property` : 'property';
      const vtTag = `${vt ?? 'vt'}:lpwstr`;
      const add = `<${tag} fmtid="${FMTID}" pid="${pid}" name="${escapeXml(TITLE_PROPERTY)}"><${vtTag}${vt ? '' : ` xmlns:vt="${NS_VT}"`}>${value}</${vtTag}></${tag}>`;
      const close = xml.lastIndexOf('</');
      const selfClosing = /<(?:[\w.-]+:)?Properties\b[^>]*\/>\s*$/.exec(xml);
      if (selfClosing) {
        const open = selfClosing[0].replace(/\s*\/>\s*$/, '>');
        out = xml.slice(0, selfClosing.index) + open + add + `</${prefix ? prefix + ':' : ''}Properties>`;
      } else if (close >= 0) {
        out = xml.slice(0, close) + add + xml.slice(close);
      } else return;
    }
    if (out !== xml) zip.file(record.part, out);
    return;
  }
  // No custom properties yet: docProps/custom.xml, pointed at from the root .rels.
  const relsFile = zip.file('_rels/.rels');
  if (!relsFile) return;
  const rels = await relsFile.async('string');
  try {
    parseXml(rels);
  } catch {
    return;
  }
  const close = rels.lastIndexOf('</Relationships>');
  if (close < 0) return;
  let part = 'docProps/custom.xml';
  for (let n = 1; zip.file(part); n++) part = `docProps/custom${n}.xml`;
  const ids = new Set((await readRels(zip, '', [])).keys());
  let n = 1;
  while (ids.has(`rIdPxCustom${n}`)) n++;
  zip.file(
    part,
    XML_HEAD +
      `<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}"><property fmtid="${FMTID}" pid="2" name="${escapeXml(TITLE_PROPERTY)}"><vt:lpwstr>${value}</vt:lpwstr></property></Properties>`,
  );
  zip.file('_rels/.rels', rels.slice(0, close) + `<Relationship Id="rIdPxCustom${n}" Type="${REL_CUSTOM}" Target="${part}"/>` + rels.slice(close));
  overrides['/' + part] = CT_CUSTOM;
}
