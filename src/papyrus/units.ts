// Unit conversions used by OOXML.
// twip = 1/20 pt = 1/1440 inch; CSS px = 1/96 inch; EMU = 1/914400 inch.

export const twipsToPx = (twips: number): number => twips / 15;
export const pxToTwips = (px: number): number => Math.round(px * 15);
export const emuToPx = (emu: number): number => emu / 9525;
export const pxToEmu = (px: number): number => Math.round(px * 9525);
/** w:sz is in half-points. */
export const halfPointsToPt = (hp: number): number => hp / 2;
export const ptToHalfPoints = (pt: number): number => Math.round(pt * 2);
