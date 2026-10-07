// What a paragraph or run gets from Word styles when it does not set a property itself.
// The editor needs this to show the toolbar state that is really in effect, and to know
// when "turn off bold" or "align left" must be written explicitly to beat a style.

/** The style-inheritable properties the editor edits. */
export interface StyleProps {
  align?: string; // left | center | right | justify
  bold?: boolean;
  italic?: boolean;
  fontSize?: number; // pt
  color?: string; // "#rrggbb"
  fontFamily?: string; // Latin font
  fontEastAsia?: string; // CJK font
  /** w:outlineLvl: 0 = level 1 heading ... 8; 9 or none = body text. Shown in the navigation pane. */
  outlineLevel?: number;
}

/** Run properties that simply override along styles (unlike bold / italic). */
export type RunValueProp = 'fontSize' | 'color' | 'fontFamily' | 'fontEastAsia';

export interface StyleInheritance {
  /** From w:docDefaults. */
  defaults: StyleProps;
  /** The paragraph style used when a paragraph names none (w:default="1"). */
  defaultParagraph: string | null;
  /** Paragraph styles, already resolved along w:basedOn. */
  paragraph: Record<string, StyleProps>;
  /** Character styles, already resolved along w:basedOn. */
  character: Record<string, StyleProps>;
}

export const emptyInheritance = (): StyleInheritance => ({
  defaults: {},
  defaultParagraph: null,
  paragraph: {},
  character: {},
});

/** w:jc value -> the editor's alignment. */
export function jcToAlign(jc: string | null): string | null {
  if (!jc) return null;
  if (jc === 'both' || jc === 'distribute') return 'justify';
  if (jc === 'start') return 'left';
  if (jc === 'end') return 'right';
  return jc;
}

function paragraphStyle(inh: StyleInheritance, styleId: string | null): StyleProps | undefined {
  // A paragraph naming an unknown style is shown with the default paragraph style, as in Word.
  return (styleId ? inh.paragraph[styleId] : undefined) ?? (inh.defaultParagraph ? inh.paragraph[inh.defaultParagraph] : undefined);
}

/** The alignment a paragraph with this style has when it does not set its own. */
export function inheritedAlign(inh: StyleInheritance | undefined, styleId: string | null): string {
  if (!inh) return 'left';
  return paragraphStyle(inh, styleId)?.align ?? inh.defaults.align ?? 'left';
}

/** A run property from styles alone: character style, then paragraph style, then document defaults. */
export function inheritedValue<K extends RunValueProp>(
  inh: StyleInheritance | undefined,
  prop: K,
  paraStyleId: string | null,
  charStyleId: string | null,
): StyleProps[K] | null {
  if (!inh) return null;
  return (
    (charStyleId ? inh.character[charStyleId]?.[prop] : undefined) ??
    paragraphStyle(inh, paraStyleId)?.[prop] ??
    inh.defaults[prop] ??
    null
  );
}

/**
 * Whether text is bold / italic from styles alone. These are "toggle" properties in
 * OOXML: a character style that turns bold on inside a bold paragraph style turns it off.
 */
export function inheritedToggle(
  inh: StyleInheritance | undefined,
  prop: 'bold' | 'italic',
  paraStyleId: string | null,
  charStyleId: string | null,
): boolean {
  if (!inh) return false;
  const fromParagraph = paragraphStyle(inh, paraStyleId)?.[prop] ?? inh.defaults[prop] ?? false;
  const fromCharacter = charStyleId ? inh.character[charStyleId]?.[prop] === true : false;
  return fromParagraph !== fromCharacter;
}
