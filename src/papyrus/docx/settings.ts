// settings.xml: Word's 追蹤修訂 switch (w:trackRevisions). Word turns it on and off in the file
// (Review > Track Changes), and so does the editor; the part is left byte for byte as it is
// unless the switch changed.

import { NS, REL_TYPE, child, onOff, parseXml, serializeXml } from './xml';
import type { CommentPackage } from './comments';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

/** The children of w:settings that come before w:trackRevisions (CT_Settings order). */
const BEFORE = new Set([
  'writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries',
  'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData',
  'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges',
  'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors',
  'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter',
  'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView',
]);

/** Whether a settings part has tracked changes turned on. */
export function trackRevisionsOn(settings: Document | null): boolean {
  return onOff(child(settings?.documentElement ?? null, 'trackRevisions')) === true;
}

/**
 * Turn w:trackRevisions on or off in the package's settings.xml (undefined: leave it). Nothing
 * is written when the part already says so; a package without settings.xml gets one when the
 * switch is turned on.
 */
export async function writeTrackRevisions(on: boolean | undefined, pkg: CommentPackage): Promise<void> {
  if (on == null) return;
  const { zip, main } = pkg;
  const rel = [...pkg.mainRels.values()].find((r) => r.type === REL_TYPE.settings && !r.external);
  const file = rel ? zip.file(rel.target) : null;
  if (!rel || !file) {
    if (!on) return;
    const dir = main.includes('/') ? main.slice(0, main.lastIndexOf('/') + 1) : '';
    let part = rel?.target ?? `${dir}settings.xml`;
    for (let n = 1; !rel && zip.file(part); n++) part = `${dir}settings${n}.xml`;
    zip.file(part, `${XML_HEAD}<w:settings xmlns:w="${NS.w}"><w:trackRevisions/></w:settings>`);
    if (!rel) pkg.extraRels.push({ id: pkg.newRelId('rIdPxSettings'), type: REL_TYPE.settings, target: part, external: false });
    pkg.overrides['/' + part] = SETTINGS_TYPE;
    return;
  }
  const text = await file.async('string');
  const doc = parseXml(text);
  if (trackRevisionsOn(doc) === on) return;
  const root = doc.documentElement;
  const old = child(root, 'trackRevisions');
  if (old) root.removeChild(old);
  if (on) {
    const prefix = root.namespaceURI === NS.w ? root.prefix : root.lookupPrefix(NS.w);
    const el = doc.createElementNS(NS.w, prefix ? `${prefix}:trackRevisions` : 'trackRevisions');
    let after: Element | null = null;
    for (let c = root.firstElementChild; c; c = c.nextElementSibling) if (c.namespaceURI === NS.w && BEFORE.has(c.localName)) after = c;
    root.insertBefore(el, after ? after.nextSibling : root.firstChild);
  }
  const head = /^<\?xml[^>]*\?>/.exec(text)?.[0] ?? XML_HEAD;
  zip.file(rel.target, head + serializeXml(root));
}
