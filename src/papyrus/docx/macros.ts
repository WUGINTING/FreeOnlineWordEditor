import type JSZip from 'jszip';
import { REL_TYPE } from './xml';
import { findMainPart, readRels } from './reader';

/**
 * Macros in a Word package (persona-300): Word keeps a document's VBA project in
 * word/vbaProject.bin (a .docm, or a .docx that had one added), with its signature and data
 * parts beside it; ActiveX controls (word/activeX/…) run code of their own too. The page never
 * runs either and saves them as they are.
 */
const MACRO_PART = /(^|\/)vba(Project\.bin|ProjectSignature\.bin|Data\.xml)$|(^|\/)activeX\/[^/]+\.(bin|xml)$/i;

/** Whether the package holds macros or ActiveX controls. */
export function hasMacroParts(zip: JSZip | null): boolean {
  if (!zip) return false;
  return Object.keys(zip.files).some((name) => !zip.files[name].dir && MACRO_PART.test(name));
}

const ATTACHED_TEMPLATE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate';

/**
 * Whether the document is attached to a macro-enabled template (設定 w:attachedTemplate pointing
 * at a .dotm): Word runs that template's macros when it opens the document where the template is.
 */
export async function hasMacroTemplate(zip: JSZip | null): Promise<boolean> {
  if (!zip) return false;
  try {
    const main = await findMainPart(zip);
    const settings = [...(await readRels(zip, main)).values()].find((r) => r.type === REL_TYPE.settings && !r.external);
    if (!settings) return false;
    return [...(await readRels(zip, settings.target)).values()].some((r) => r.type === ATTACHED_TEMPLATE && /\.dotm$/i.test(r.target.replace(/[?#].*$/, '')));
  } catch {
    return false;
  }
}

/** Told to the user when such a document opens (the host shows it; see DocxEditor.hasMacros). */
export const MACRO_NOTICE = '這份文件含有巨集（Word 的自動程式），線上文件不會執行，下載後在 Word 開啟請小心';
