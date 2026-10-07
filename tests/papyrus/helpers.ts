// Test helpers for comparing editor documents.

/** Attributes / marks that only record the original XML for lossless saving. */
const BOOKKEEPING_ATTRS = new Set([
  'pPr', 'pAttrs', 'brGroup', 'wrap', 'tcPr', 'mergedTc', 'filler', 'trPr', 'trAttrs', 'tblPrEx',
  'grid', 'tblGridXml', 'tblPr', 'group', 'para', 'rPr',
]);
const BOOKKEEPING_MARKS = new Set(['run', 'inlineWrap']);

/** Document JSON without lossless bookkeeping, for comparing with a document built in the editor. */
export function content(json: any): any {
  if (Array.isArray(json)) return json.map(content);
  if (!json || typeof json !== 'object') return json;
  const out: any = {};
  for (const [k, v] of Object.entries(json)) {
    if (k === 'attrs') {
      const attrs: any = {};
      for (const [ak, av] of Object.entries(v as object)) {
        if (BOOKKEEPING_ATTRS.has(ak)) continue;
        if ((json.type === 'field' || json.type === 'image') && ['xml', 'origSrc', 'origWidth', 'origHeight'].includes(ak)) continue;
        attrs[ak] = av;
      }
      out.attrs = attrs;
    } else if (k === 'marks') {
      const marks = (v as any[]).filter((m) => !BOOKKEEPING_MARKS.has(m.type)).map(content);
      if (marks.length) out.marks = marks;
    } else out[k] = content(v);
  }
  return out;
}

/** Document JSON with generated wrapper / group ids ("L12") renumbered by first appearance. */
export function stableIds(json: any): any {
  const map = new Map<string, string>();
  const text = JSON.stringify(json).replace(/\bL\d+\b/g, (id) => {
    if (!map.has(id)) map.set(id, 'L' + map.size);
    return map.get(id)!;
  });
  return JSON.parse(text);
}
