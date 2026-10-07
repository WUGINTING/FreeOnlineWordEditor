import type { StyleInheritance } from './inheritance';
import type JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';

/** Page geometry, all in twips (OOXML native unit). */
export interface PageSetup {
  width: number;
  height: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  header: number; // distance from top edge to header
  footer: number; // distance from bottom edge to footer
}

export const A4: PageSetup = {
  width: 11906,
  height: 16838,
  marginTop: 1440,
  marginBottom: 1440,
  marginLeft: 1800,
  marginRight: 1800,
  header: 851,
  footer: 992,
};

export interface NumLevel {
  fmt: string; // decimal | bullet | lowerLetter | upperLetter | lowerRoman | upperRoman | none ...
  text: string; // lvlText, e.g. "%1." or "•"
  start: number;
  indLeft: number | null; // twips
  hanging: number | null; // twips
  /** What follows the number (w:suff): a tab (Word's default, when absent), a space, or nothing. */
  suff?: 'tab' | 'space' | 'nothing';
}

export interface AbstractNum {
  id: string;
  levels: NumLevel[];
  /** Original XML; written back unchanged when present. */
  raw?: string;
  /** Made for a pasted list: never reused for the toolbar's list buttons. */
  pasted?: boolean;
}

export interface NumDef {
  id: string;
  abstractId: string;
  startOverrides: Record<number, number>;
  raw?: string;
}

export interface Numbering {
  abstracts: Record<string, AbstractNum>;
  nums: Record<string, NumDef>;
  /** Other children of w:numbering kept verbatim: picture bullets (written first) and the rest (written last). */
  head?: string[];
  tail?: string[];
  /**
   * Lists made in the editor (pasted, list buttons) that no document uses any more, e.g. after
   * undoing the paste: set aside so redo brings them back with the same ids; never written.
   * See syncMadeLists in numbering.ts.
   */
  parked?: { abstracts: Record<string, AbstractNum>; nums: Record<string, NumDef> };
}

export interface ParagraphStyleInfo {
  id: string;
  name: string;
}

export type HeaderFooterKind = 'header' | 'footer';
export type HeaderFooterType = 'default' | 'first' | 'even';

/** One header or footer part (word/header1.xml ...), referenced by one or more sections. */
export interface HeaderFooterPart {
  kind: HeaderFooterKind;
  type: HeaderFooterType;
  /** Relationship id from the main document; null until first saved. */
  relId: string | null;
  /** Package path such as "word/header1.xml"; null for one created in the editor. */
  part: string | null;
  doc: PMNode;
  /** Changed in the editor, so it must be rewritten on save. */
  dirty: boolean;
  /**
   * Made in the editor for an earlier section and not referenced by it yet: it gets its
   * reference once it has content, so opening an empty header changes nothing.
   */
  pending?: boolean;
}

/** Everything about a document that lives outside the editable body. */
export interface DocxModel {
  /** The original package; untouched parts are copied through on save. */
  zip: JSZip | null;
  page: PageSetup;
  /** Raw XML of the final (body-level) w:sectPr, reused on save. Earlier sections keep theirs in paragraphs. */
  sectPr: string | null;
  /** "Different first page" (w:titlePg) of the last section. */
  titlePage: boolean;
  /** Relationship ids of the main part when it was read, so new ones never collide. */
  relIds?: string[];
  /**
   * Package paths of optional parts (styles, numbering, settings, a header/footer, a .rels file)
   * that were not well-formed XML when opened: treated as missing, and saved untouched.
   * Absent when there are none. The UI may show them.
   */
  brokenParts?: string[];
  /** "Different odd & even pages" (w:evenAndOddHeaders in settings.xml). Display only for now. */
  evenAndOdd: boolean;
  /**
   * Word's 追蹤修訂 switch (w:trackRevisions in settings.xml): read when opening, changed with
   * DocxEditor.setTrackChanges. Saved into settings.xml only when it differs from the file.
   */
  trackRevisions?: boolean;
  /**
   * Layout rules from settings.xml that change where text sits on a page: Word's compatibility
   * mode (w:compatSetting compatibilityMode; absent = Word 2007's 12) and "suppress space
   * before after a hard page break" (w:suppressSpBfAfterPgBrk).
   */
  compat?: { mode: number; suppressSpaceAfterPageBreak: boolean };
  numbering: Numbering;
  paragraphStyles: ParagraphStyleInfo[];
  /** What paragraphs and runs inherit from styles (alignment, bold, italic). */
  styles?: StyleInheritance;
  /** CSS generated from styles.xml, scoped to `.dx-doc`. */
  css: string;
  headerFooters: HeaderFooterPart[];
  /**
   * Word's 目錄 style ids (TOC1 … TOCHeading) a table of contents made or rebuilt in the editor
   * used (editor/toc.ts): only these are added to styles.xml when the file lacks them (tocStyles.ts).
   */
  tocStylesUsed?: Set<string>;
  /** The theme's colours (theme1.xml), for shapes made in the editor. */
  theme?: Record<string, string>;
}

export function findHeaderFooter(
  model: DocxModel,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): HeaderFooterPart | undefined {
  return model.headerFooters.find((h) => h.kind === kind && h.type === type);
}

/** Which header/footer variant Word shows on a (0-based) page of a single-section document. */
export function variantForPage(model: DocxModel, pageIndex: number): HeaderFooterType {
  if (pageIndex === 0 && model.titlePage) return 'first';
  if (model.evenAndOdd && pageIndex % 2 === 1) return 'even';
  return 'default';
}

export function emptyNumbering(): Numbering {
  return { abstracts: {}, nums: {} };
}
