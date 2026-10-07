// Font names on their way into CSS (a <style> element and style attributes): only what a font
// name can be is let through, so a crafted document cannot add CSS of its own.

/** A font name safe inside a double-quoted CSS string ('' when nothing is left). */
export function cssFontName(name: string | null | undefined): string {
  return (name ?? '').replace(/[^\p{L}\p{N} _.\-&+@]/gu, '').trim().slice(0, 64);
}

/**
 * The font family standing for East Asian font `name` on the characters w:hint="eastAsia" moves
 * to it: an @font-face alias of the local font limited to those characters (eastAsiaFonts.ts).
 * Set as --dx-font-h wherever --dx-font-e names the font, so a hinted run can put it first.
 */
export const eastAsiaAlias = (name: string): string => `dx-ea ${cssFontName(name)}`;
