// run-187 (GOV-184): copying text anchored to a comment and pasting it put the comment marker's
// on-screen label 「註解」 into the text. Kept-as-is content must not paste as its label.
import { describe, expect, it } from 'vitest';
import { DOMParser as PMDOMParser, DOMSerializer, Fragment } from 'prosemirror-model';
import { schema } from '../../src/papyrus/editor/schema';

function roundTrip(content: Fragment) {
  const div = document.createElement('div');
  div.append(DOMSerializer.fromSchema(schema).serializeFragment(content));
  return PMDOMParser.fromSchema(schema).parse(div);
}

describe('pasting copied content with kept-as-is markers', () => {
  it('drops a comment reference marker instead of pasting 「註解」 as text', () => {
    const para = schema.nodes.paragraph.create(null, [
      schema.text('採購規格測試文字 GOV184'),
      schema.nodes.raw_inline.create({ xml: '<w:r><w:commentReference w:id="1"/></w:r>', label: '註解' }),
    ]);
    const pasted = roundTrip(Fragment.from(para));
    expect(pasted.textContent).toBe('採購規格測試文字 GOV184');
    expect(pasted.textContent).not.toContain('註解');
  });

  it('drops hidden markers and other kept content labels, and keeps plain symbols as text', () => {
    const para = schema.nodes.paragraph.create(null, [
      schema.text('A'),
      schema.nodes.raw_inline.create({ xml: '<w:bookmarkStart w:id="0" w:name="x"/>', label: 'bookmarkStart', hidden: true }),
      schema.nodes.raw_inline.create({ xml: '<w:r><w:object/></w:r>', label: '物件' }),
      schema.nodes.raw_inline.create({ xml: '<w:r><w:sym w:char="F0E0"/></w:r>', label: '→', plain: true }),
      schema.text('B'),
    ]);
    const pasted = roundTrip(Fragment.from(para));
    expect(pasted.textContent).toBe('A→B');
  });
});
