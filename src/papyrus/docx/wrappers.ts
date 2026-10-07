// Wrapper elements (w:sdt, w:customXml, w:ins, w:smartTag, w:hyperlink, w:fldSimple ...)
// hold editable content. The editor flattens them: every block / inline node carries
// the list of wrappers around it ("layers", outermost first), and the writer
// re-opens and closes them around consecutive nodes that share them.

import { NS, escapeAttr, serializeXml } from './xml';

export interface Layer {
  /** Unique per wrapper instance, so two adjacent wrappers don't merge. */
  id: string;
  /** Start tag plus any property children, e.g. `<w:sdt><w:sdtPr>…</w:sdtPr><w:sdtContent>`. */
  open: string;
  close: string;
  /** Hyperlinks are regenerated from the editor's link mark (the href may have changed). */
  kind?: 'link';
}

let counter = 0;
export function newLayerId(): string {
  return 'L' + ++counter;
}

export function parseLayers(json: string | null | undefined): Layer[] {
  return json ? (JSON.parse(json) as Layer[]) : [];
}

export function addLayer(json: string | null, layer: Layer): string {
  return JSON.stringify([...parseLayers(json), layer]);
}

/** Attributes of an element as written, e.g. ` w:rsidR="00A1" w14:paraId="1F2E"`. */
export function attrString(el: Element, skip: (name: string) => boolean = () => false): string {
  let out = '';
  for (const a of Array.from(el.attributes)) if (!skip(a.name)) out += ` ${a.name}="${escapeAttr(a.value)}"`;
  return out;
}

export function startTag(el: Element, skip?: (name: string) => boolean): string {
  return `<${el.tagName}${attrString(el, skip)}>`;
}

/**
 * Describe a wrapper element. `contentHolder` is the child whose children are the
 * wrapped content (w:sdtContent for w:sdt), `propLocals` the property children kept
 * verbatim in the opening part.
 */
export function layerFor(el: Element): { layer: Layer; content: Element } {
  const id = newLayerId();
  if (el.localName === 'sdt') {
    let open = startTag(el);
    let content: Element = el;
    for (const c of Array.from(el.children)) {
      if (c.namespaceURI === NS.w && c.localName === 'sdtContent') {
        content = c;
        open += startTag(c);
        return { layer: { id, open, close: `</${c.tagName}></${el.tagName}>` }, content };
      }
      open += serializeXml(c);
    }
    return { layer: { id, open, close: `</${el.tagName}>` }, content };
  }
  // customXml / smartTag keep their *Pr child in the opening part.
  let open = startTag(el);
  for (const c of Array.from(el.children)) {
    if (c.namespaceURI === NS.w && (c.localName === 'customXmlPr' || c.localName === 'smartTagPr')) open += serializeXml(c);
  }
  return { layer: { id, open, close: `</${el.tagName}>` }, content: el };
}

/** Children of a wrapper that are content (not its property element). */
export function contentChildren(el: Element): Element[] {
  return Array.from(el.children).filter(
    (c) => !(c.namespaceURI === NS.w && (c.localName === 'customXmlPr' || c.localName === 'smartTagPr' || c.localName === 'sdtPr' || c.localName === 'sdtEndPr')),
  );
}

/**
 * Writes a sequence of items that each carry layers, opening and closing
 * wrappers so that shared layers stay shared.
 */
export class LayerStack {
  private open: Layer[] = [];
  constructor(private openTag: (l: Layer) => string = (l) => l.open) {}

  /** Markup needed to move from the currently open layers to `target`. */
  moveTo(target: Layer[]): string {
    let common = 0;
    while (common < this.open.length && common < target.length && sameLayer(this.open[common], target[common])) common++;
    let out = '';
    for (let i = this.open.length - 1; i >= common; i--) out += this.open[i].close;
    for (let i = common; i < target.length; i++) out += this.openTag(target[i]);
    this.open = target.slice();
    return out;
  }

  closeAll(): string {
    return this.moveTo([]);
  }
}

function sameLayer(a: Layer, b: Layer): boolean {
  return a.id === b.id && a.open === b.open;
}
