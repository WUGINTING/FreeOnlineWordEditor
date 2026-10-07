/** Centimetres in twips, rounded as Word writes them. */
export const cmToTwips = (cm: number) => Math.round((cm * 1440) / 2.54);
