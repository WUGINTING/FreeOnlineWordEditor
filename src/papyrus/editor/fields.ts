import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { AddMarkStep, AddNodeMarkStep, AttrStep, DocAttrStep, RemoveMarkStep, RemoveNodeMarkStep, ReplaceAroundStep, ReplaceStep, type Step } from 'prosemirror-transform';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { parseLayers, type Layer } from '../docx/wrappers';
import { tl } from '../i18n';

/**
 * Fields to fill in: Word's content controls (w:sdt) in a template, e.g. 「客戶名稱」 or
 * 「金額」. The editor keeps them as wrapper layers (docx/wrappers.ts): on the inline
 * content (inlineWrap mark) or on whole blocks (the `wrap` attribute).
 *
 * A field is required when its tag or title says so: the tag or title contains 「必填」 or
 * "required", or the title ends with "*". Word itself has no "required" flag.
 */
export interface DocField {
  /** The wrapper's id (stable while the document is open). */
  id: string;
  /** w:alias (the title Word shows), else w:tag, else 「未命名欄位」. */
  title: string;
  tag: string;
  kind: 'text' | 'richText' | 'date' | 'list' | 'checkbox' | 'picture';
  required: boolean;
  /** Content can't be edited (w:lock contentLocked / sdtContentLocked). */
  locked: boolean;
  /** The control itself can't be deleted (w:lock sdtLocked / sdtContentLocked). */
  keepControl: boolean;
  /** Still shows its placeholder text (w:showingPlcHdr): not filled in yet. */
  placeholder: boolean;
  filled: boolean;
  text: string;
  from: number;
  to: number;
}

const SDT = /^<(?:\w+:)?sdt[\s>]/;
/** Building blocks and citations are content controls too, but not fields to fill in. */
const NOT_A_FIELD = /<(?:\w+:)?(?:docPartObj|docPartList|citation|bibliography|equation)\b/;

function attrVal(xml: string, tag: string): string | null {
  const m = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*\\bw:val="([^"]*)"`).exec(xml);
  return m ? decode(m[1]) : null;
}

function decode(s: string): string {
  return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function onOffTag(xml: string, tag: string): boolean {
  const m = new RegExp(`<(?:\\w+:)?${tag}\\b([^>]*)/?>`).exec(xml);
  return !!m && !/w:val="(0|false|off)"/.test(m[1]);
}

function kindOf(xml: string): DocField['kind'] {
  if (/<(?:\w+:)?checkbox\b/.test(xml)) return 'checkbox';
  if (/<(?:\w+:)?date\b/.test(xml)) return 'date';
  if (/<(?:\w+:)?(?:dropDownList|comboBox)\b/.test(xml)) return 'list';
  if (/<(?:\w+:)?picture\b/.test(xml)) return 'picture';
  if (/<(?:\w+:)?text\b/.test(xml)) return 'text';
  return 'richText';
}

const REQUIRED = /必填|required/i;

/** The field a wrapper layer stands for, or null when it isn't a content control to fill in. */
function fieldOf(layer: Layer): Omit<DocField, 'filled' | 'text' | 'from' | 'to'> | null {
  if (!SDT.test(layer.open)) return null;
  const pr = /<(?:\w+:)?sdtPr\b[\s\S]*?<\/(?:\w+:)?sdtPr>/.exec(layer.open)?.[0] ?? '';
  if (NOT_A_FIELD.test(pr)) return null;
  const alias = attrVal(pr, 'alias') ?? '';
  const tag = attrVal(pr, 'tag') ?? '';
  const lock = attrVal(pr, 'lock') ?? '';
  return {
    id: layer.id,
    title: alias || tag || tl('未命名欄位'),
    tag,
    kind: kindOf(pr),
    required: REQUIRED.test(tag) || REQUIRED.test(alias) || /\*\s*$/.test(alias),
    locked: lock === 'contentLocked' || lock === 'sdtContentLocked',
    keepControl: lock === 'sdtLocked' || lock === 'sdtContentLocked',
    placeholder: onOffTag(pr, 'showingPlcHdr'),
  };
}

const cache = new WeakMap<PMNode, DocField[]>();
const NO_FIELDS: DocField[] = [];

/** How many times a document was scanned for its fields (for tests). */
export const fieldScans = { count: 0 };

/** Every field of the document, in document order (documents are immutable: cached per document). */
export function documentFields(doc: PMNode): DocField[] {
  let fields = cache.get(doc);
  if (!fields) cache.set(doc, (fields = scanFields(doc)));
  return fields;
}

/**
 * The wrapper layers of a `wrap` / inlineWrap attribute, parsed once per attribute text (the
 * same text is on every run and block of a wrapper). Shared: never changed by callers.
 */
const layerCache = new Map<string, { layers: Layer[]; fields: (Omit<DocField, 'filled' | 'text' | 'from' | 'to'> | null)[] }>();
const LAYER_CACHE_MAX = 2000;

function parsedLayers(json: string) {
  let hit = layerCache.get(json);
  if (!hit) {
    if (layerCache.size >= LAYER_CACHE_MAX) layerCache.clear();
    const layers = parseLayers(json);
    hit = { layers, fields: layers.map(fieldOf) };
    layerCache.set(json, hit);
  }
  return hit;
}

function layersOf(json: string | null | undefined): Layer[] {
  return !json || json === '[]' ? [] : parsedLayers(json).layers;
}

/** Whether a wrapper attribute value has any layer. */
const hasLayers = (json: unknown): boolean => typeof json === 'string' && json !== '' && json !== '[]';

/**
 * Whether a transaction may bring in a content control: content with a wrapper, a wrapper mark
 * or a changed `wrap` attribute (unknown steps: yes). Removing content never adds a field.
 */
function mayAddWrappers(tr: Transaction): boolean {
  const wrapped = (node: PMNode) => hasLayers(node.attrs.wrap) || node.marks.some((m) => m.type.name === 'inlineWrap');
  for (const step of tr.steps) {
    if (step instanceof ReplaceStep || step instanceof ReplaceAroundStep) {
      let found = false;
      step.slice.content.descendants((node) => {
        if (found || wrapped(node)) found = true;
        return !found;
      });
      if (found) return true;
    } else if (step instanceof AddMarkStep || step instanceof AddNodeMarkStep) {
      if (step.mark.type.name === 'inlineWrap') return true;
    } else if (step instanceof AttrStep || step instanceof DocAttrStep) {
      if (step.attr === 'wrap' && hasLayers(step.value)) return true;
    } else if (!(step instanceof RemoveMarkStep || step instanceof RemoveNodeMarkStep)) {
      return true;
    }
  }
  return false;
}

/**
 * A document made from one without fields, by a transaction that adds no wrapper, has none
 * either: typing in a long document doesn't scan it for fields on every keystroke.
 */
function carryNoFields(tr: Transaction, oldDoc: PMNode, newDoc: PMNode): void {
  if (!tr.docChanged || cache.has(newDoc) || cache.get(oldDoc)?.length !== 0) return;
  if (!mayAddWrappers(tr)) cache.set(newDoc, NO_FIELDS);
}

function scanFields(doc: PMNode): DocField[] {
  fieldScans.count++;
  const found = new Map<string, Omit<DocField, 'filled' | 'text'>>();
  const add = (layers: string | null | undefined, from: number, to: number) => {
    if (!layers || layers === '[]') return;
    const parsed = parsedLayers(layers);
    parsed.layers.forEach((layer, i) => {
      const f = found.get(layer.id);
      if (f) {
        f.from = Math.min(f.from, from);
        f.to = Math.max(f.to, to);
        return;
      }
      const field = parsed.fields[i];
      if (field) found.set(layer.id, { ...field, from, to });
    });
  };
  doc.descendants((node, pos) => {
    if (node.attrs.wrap) add(node.attrs.wrap as string, pos, pos + node.nodeSize);
    if (node.isInline) {
      for (const m of node.marks) if (m.type.name === 'inlineWrap') add(m.attrs.layers, pos, pos + node.nodeSize);
    }
    return true;
  });
  if (!found.size) return NO_FIELDS;
  return [...found.values()]
    .sort((a, b) => a.from - b.from)
    .map((f) => {
      const text = doc.textBetween(f.from, f.to, '\n', '￼');
      const filled = f.kind === 'checkbox' || (!f.placeholder && text.replace(/￼/g, 'x').trim() !== '');
      return { ...f, text, filled };
    });
}

/** Required fields that are still empty (or still show their placeholder). */
export function unfilledRequired(doc: PMNode): DocField[] {
  return documentFields(doc).filter((f) => f.required && !f.filled);
}

/** Selects a field's content and scrolls to it. */
export function selectField(view: EditorView, id: string): boolean {
  const field = documentFields(view.state.doc).find((f) => f.id === id);
  if (!field) return false;
  const tr = view.state.tr.setSelection(fieldSelection(view.state.doc, field)).scrollIntoView();
  view.dispatch(tr);
  view.focus();
  return true;
}

/** The next field that isn't filled in, after the cursor (wrapping around); selects it. */
export function nextUnfilled(view: EditorView, requiredOnly = false): DocField | null {
  const fields = documentFields(view.state.doc).filter((f) => !f.filled && !f.locked && (!requiredOnly || f.required));
  if (!fields.length) return null;
  const at = view.state.selection.to;
  const next = fields.find((f) => f.from >= at) ?? fields[0];
  selectField(view, next.id);
  return next;
}

/** The text inside a field (a whole-block field: the text of its blocks). */
function fieldSelection(doc: PMNode, f: DocField): TextSelection {
  const $from = doc.resolve(f.from);
  const $to = doc.resolve(f.to);
  // A block field: from the start of its first text block to the end of its last.
  if (!$from.parent.inlineContent || !$to.parent.inlineContent) {
    const start = TextSelection.findFrom(doc.resolve(f.from), 1, true);
    const end = TextSelection.findFrom(doc.resolve(f.to), -1, true);
    if (start && end && end.from >= start.from) return TextSelection.create(doc, start.from, end.to);
  }
  return TextSelection.create(doc, f.from, f.to);
}

// ----- placeholders, as in Word -----

const fieldsKey = new PluginKey('dx-fields');

/**
 * Word's placeholder behaviour. A click in a field that still shows its placeholder selects
 * the placeholder, so typing replaces it; once the field's text changes it is no longer a
 * placeholder: w:showingPlcHdr is dropped (else Word would treat the typed text as the
 * placeholder) and so is the placeholder's character style (the grey 「按一下輸入文字」 look).
 */
export function fieldPlaceholders(): Plugin {
  return new Plugin({
    key: fieldsKey,
    state: {
      init: () => null,
      apply(tr, value, oldState, newState) {
        carryNoFields(tr, oldState.doc, newState.doc);
        return value;
      },
    },
    props: {
      handleClick(view, pos) {
        const field = documentFields(view.state.doc).find((f) => f.placeholder && !f.locked && pos >= f.from && pos <= f.to);
        if (!field) return false;
        view.dispatch(view.state.tr.setSelection(fieldSelection(view.state.doc, field)));
        return true;
      },
    },
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      const before = documentFields(oldState.doc).filter((f) => f.placeholder);
      if (!before.length) return null;
      const now = new Map(documentFields(newState.doc).map((f) => [f.id, f]));
      let tr: Transaction | null = null;
      for (const old of before) {
        const f = now.get(old.id);
        if (!f || !f.placeholder || f.text === old.text) continue;
        tr ??= newState.tr;
        dropPlaceholder(oldState, newState, tr, old, f);
      }
      // Part of the same undo step as the typing that caused it.
      return tr;
    },
  });
}

function dropPlaceholder(oldState: EditorState, state: EditorState, tr: Transaction, old: DocField, f: DocField): void {
  const unmark = (layers: string) =>
    JSON.stringify(
      parseLayers(layers).map((l) => (l.id === f.id ? { ...l, open: l.open.replace(/<(?:\w+:)?showingPlcHdr\b[^>]*\/>/g, '') } : l)),
    );
  // The character style every piece of the placeholder text had (Word's "Placeholder Text").
  let placeholderStyles: Set<string> | null = null;
  oldState.doc.nodesBetween(old.from, old.to, (node) => {
    if (!node.isText) return true;
    const ids = new Set(node.marks.filter((m) => m.type.name === 'charStyle').map((m) => String(m.attrs.id)));
    placeholderStyles = placeholderStyles ? new Set([...placeholderStyles].filter((id) => ids.has(id))) : ids;
    return false;
  });
  const styles: Set<string> = placeholderStyles ?? new Set();
  state.doc.nodesBetween(f.from, f.to, (node, pos) => {
    if (node.attrs.wrap && layersOf(node.attrs.wrap).some((l) => l.id === f.id)) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, wrap: unmark(node.attrs.wrap) }, node.marks);
    }
    if (!node.isInline) return true;
    const end = pos + node.nodeSize;
    for (const m of node.marks) {
      if (m.type.name === 'inlineWrap' && layersOf(m.attrs.layers).some((l) => l.id === f.id)) {
        tr.removeMark(pos, end, m);
        tr.addMark(pos, end, m.type.create({ ...m.attrs, layers: unmark(m.attrs.layers) }));
      }
      if (m.type.name === 'charStyle' && styles.has(String(m.attrs.id))) tr.removeMark(pos, end, m);
    }
    return false;
  });
}


// ----- locked fields, as in Word -----

/** A transaction may pass a locked field when it carries this meta (e.g. a whole-document replace by the host). */
export const PASS_LOCKS = 'dx-pass-locks';

interface Range {
  from: number;
  to: number;
}

/** What a step changes, in the positions of the document it applies to. */
function stepRanges(step: Step): Range[] {
  const out: Range[] = [];
  step.getMap().forEach((from, to) => out.push({ from, to }));
  // Mark steps change no positions but still change the content's formatting.
  const s = step as unknown as { from?: number; to?: number; pos?: number };
  if (!out.length && typeof s.from === 'number' && typeof s.to === 'number') out.push({ from: s.from, to: s.to });
  return out;
}

/**
 * Word's content control locks. "Contents cannot be edited" (contentLocked): no typing,
 * deleting or formatting inside the field. "Content control cannot be deleted" (sdtLocked):
 * its content may change but the field may not be removed as a whole. A refused edit is
 * reported through `onBlocked`.
 */
export function lockedFields(onBlocked?: (field: DocField) => void): Plugin {
  return new Plugin({
    filterTransaction(tr, state) {
      if (!tr.docChanged || tr.getMeta(PASS_LOCKS)) return true;
      const fields = documentFields(state.doc).filter((f) => f.locked || f.keepControl);
      if (!fields.length) return true;
      for (let i = 0; i < tr.steps.length; i++) {
        const before = tr.mapping.slice(0, i);
        for (const r of stepRanges(tr.steps[i])) {
          for (const f of fields) {
            const from = before.map(f.from, 1);
            const to = before.map(f.to, -1);
            const touchesInside = r.from < to && r.to > from; // overlaps the field's content
            const insertsInside = r.from === r.to && r.from > from && r.from < to;
            const removesAll = r.from <= from && r.to >= to && r.to > r.from;
            // Typing at either edge of the field adds to it when the new text carries the field.
            const growsField = r.from === r.to && (r.from === from || r.from === to) && sliceHasLayer(tr.steps[i], f.id);
            if ((f.locked && (touchesInside || insertsInside || growsField)) || (f.keepControl && removesAll)) {
              onBlocked?.(f);
              return false;
            }
          }
        }
      }
      return true;
    },
  });
}

/** Whether a step inserts content wrapped in the field `id` (its inlineWrap mark or block wrap). */
function sliceHasLayer(step: Step, id: string): boolean {
  const slice = (step as unknown as { slice?: { content: PMNode['content'] } }).slice;
  if (!slice) return false;
  let found = false;
  slice.content.descendants((node) => {
    if (found) return false;
    if (node.attrs.wrap && layersOf(node.attrs.wrap).some((l) => l.id === id)) found = true;
    for (const m of node.marks) if (m.type.name === 'inlineWrap' && layersOf(m.attrs.layers).some((l) => l.id === id)) found = true;
    return !found;
  });
  return found;
}
