// persona-300: a right tab stop (a table of contents line) is found in the paragraph properties.
import { describe, expect, it } from 'vitest';
import { rightStop } from '../../src/papyrus/editor/tabStops';

describe('rightStop', () => {
  it('finds the rightmost right tab and its leader', () => {
    expect(rightStop('<w:pPr><w:tabs><w:tab w:val="left" w:pos="400"/><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs></w:pPr>')).toEqual({ pos: 8296, leader: 'dot' });
    expect(rightStop('<w:pPr><w:tabs><w:tab w:val="right" w:leader="none" w:pos="9000"/></w:tabs></w:pPr>')).toEqual({ pos: 9000, leader: null });
  });
  it('none without a right tab', () => {
    expect(rightStop(null)).toBeNull();
    expect(rightStop('<w:pPr><w:jc w:val="center"/></w:pPr>')).toBeNull();
    expect(rightStop('<w:pPr><w:tabs><w:tab w:val="left" w:pos="400"/></w:tabs></w:pPr>')).toBeNull();
  });
});
