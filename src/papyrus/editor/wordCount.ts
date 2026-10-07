import { Fragment, type Node as PMNode } from 'prosemirror-model';

/**
 * Word's word count: every East Asian character (full-width punctuation included) is a word,
 * other text counts runs between spaces (a number or an English word is one). Characters are counted with and without spaces.
 */
export interface WordCount {
  words: number;
  chars: number;
  charsWithSpaces: number;
}

// CJK ideographs, kana, hangul, full-width forms and CJK punctuation.
const EAST_ASIAN = /[⺀-鿿가-힯豈-﫿︰-﹏＀-￯]/u;

export function countText(text: string): WordCount {
  let words = 0;
  let chars = 0;
  let charsWithSpaces = 0;
  let inWord = false;
  for (const ch of text) {
    if (ch === '￼') continue; // pictures and other objects
    charsWithSpaces++;
    if (/\s/.test(ch)) {
      inWord = false;
      continue;
    }
    chars++;
    if (EAST_ASIAN.test(ch)) {
      // Full-width punctuation too (checked against Word: 「你好，世界。」 is 6 words).
      words++;
      inWord = false;
    } else if (!inWord) {
      words++;
      inWord = true;
    }
  }
  return { words, chars, charsWithSpaces };
}

interface BlockCount extends WordCount {
  /**
   * The block has a paragraph (or a block-level object) in it: the document's text puts a line
   * break before it (except before the first one), which counts as a character with spaces.
   */
  separated: boolean;
}

/** Per top-level block: documents share unchanged blocks, so an edit recounts only its own. */
const blockCache = new WeakMap<PMNode, BlockCount>();
/** And per document (the selection moving asks again for the same document). */
const docCache = new WeakMap<PMNode, WordCount>();

function countBlock(node: PMNode): BlockCount {
  let c = blockCache.get(node);
  if (c) return c;
  // The block's text as doc.textBetween gives it (without the break before its first paragraph).
  const text = Fragment.from(node).textBetween(0, node.nodeSize, '\n', '￼');
  let separated = node.isTextblock || (node.isBlock && node.isLeaf);
  if (!separated) {
    node.descendants((n) => {
      if (separated) return false;
      if (n.isTextblock || (n.isBlock && n.isLeaf)) separated = true;
      return !separated;
    });
  }
  c = { ...countText(text), separated };
  blockCache.set(node, c);
  return c;
}

/**
 * The count for a whole document: countText of the body's text (doc.textBetween with a line break
 * between paragraphs and ￼ for objects), added up per top-level block (a line break ends a word, so blocks count apart).
 */
export function countDocument(doc: PMNode): WordCount {
  const cached = docCache.get(doc);
  if (cached) return cached;
  let words = 0;
  let chars = 0;
  let charsWithSpaces = 0;
  let breaks = -1;
  doc.forEach((node) => {
    const c = countBlock(node);
    words += c.words;
    chars += c.chars;
    charsWithSpaces += c.charsWithSpaces;
    if (c.separated) breaks++;
  });
  const out = { words, chars, charsWithSpaces: charsWithSpaces + Math.max(0, breaks) };
  docCache.set(doc, out);
  return out;
}

/** Counts added up (the body and its text boxes). */
export function addCounts(counts: WordCount[]): WordCount {
  return counts.reduce((a, c) => ({ words: a.words + c.words, chars: a.chars + c.chars, charsWithSpaces: a.charsWithSpaces + c.charsWithSpaces }), { words: 0, chars: 0, charsWithSpaces: 0 });
}
