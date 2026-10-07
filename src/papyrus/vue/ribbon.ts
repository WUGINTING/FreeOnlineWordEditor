// The ribbon's (DocxToolbar.vue) 版面配置 menus and icons.

/**
 * Paper sizes in twips (1 cm = 567 twips), portrait: the same list as PageSetupDialog.vue has
 * (kept there too, as the tests compile that dialog on its own).
 */
export const PAPERS = [
  { id: 'A4', label: 'A4（21 × 29.7 公分）', w: 11906, h: 16838 },
  { id: 'A3', label: 'A3（29.7 × 42 公分）', w: 16838, h: 23811 },
  { id: 'A5', label: 'A5（14.8 × 21 公分）', w: 8391, h: 11906 },
  { id: 'B4', label: 'B4（25.7 × 36.4 公分）', w: 14570, h: 20636 },
  { id: 'B5', label: 'B5（18.2 × 25.7 公分）', w: 10318, h: 14570 },
  { id: 'Letter', label: 'Letter（8.5 × 11 英吋）', w: 12240, h: 15840 },
  { id: 'Legal', label: 'Legal（8.5 × 14 英吋）', w: 12240, h: 20160 },
];

/** Word's 邊界 gallery (twips): top/bottom and left/right. */
export const MARGIN_PRESETS = [
  { id: 'normal', label: '標準', note: '上下 2.54 公分，左右 3.18 公分', tb: 1440, lr: 1800 },
  { id: 'narrow', label: '窄', note: '上下左右 1.27 公分', tb: 720, lr: 720 },
  { id: 'moderate', label: '適中', note: '上下 2.54 公分，左右 1.91 公分', tb: 1440, lr: 1080 },
  { id: 'wide', label: '寬', note: '上下 2.54 公分，左右 5.08 公分', tb: 1440, lr: 2880 },
];

/** Within half a millimetre (Word rounds its own values). */
export const near = (a: number, b: number) => Math.abs(a - b) < 30;

/**
 * The ribbon's icons: the inside of a 24 × 24 <svg class="dx-ico">, drawn with lines in the
 * text colour; parts with class "f" are filled.
 */
export const ICONS: Record<string, string> = {
  undo: '<path d="M8 4 3 9l5 5"/><path d="M3 9h11a6 6 0 0 1 0 12h-4"/>',
  redo: '<path d="m16 4 5 5-5 5"/><path d="M21 9H10a6 6 0 0 0 0 12h4"/>',
  paste: '<rect x="4" y="4" width="13" height="17" rx="2"/><rect x="7.5" y="2" width="6" height="4" rx="1"/><rect x="11" y="10" width="10" height="12" rx="1" class="w"/>',
  cut: '<circle cx="6" cy="18" r="3"/><circle cx="18" cy="18" r="3"/><path d="M8 16 19 3M16 16 5 3"/>',
  copy: '<rect x="8" y="8" width="12" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2h2"/>',
  formatPainter: '<rect x="3" y="3" width="14" height="6" rx="1"/><path d="M17 6h3v6h-8v3"/><rect class="f" x="10.5" y="15" width="3" height="7" rx="0.5"/>',
  find: '<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',
  replace: '<path d="M4 8h13l-3-3M20 16H7l3 3"/>',
  selectAll: '<rect x="3" y="3" width="18" height="18" rx="1" stroke-dasharray="3 2"/><path d="m8 12 3 3 5-6"/>',
  blankPage: '<path d="M6 2h9l4 4v16H6z"/><path d="M15 2v4h4"/>',
  pageBreak: '<path d="M5 2v6h14V2M5 22v-6h14v6"/><path d="M2 12h3M8 12h3M14 12h3M20 12h2"/>',
  table: '<rect x="3" y="4" width="18" height="16" rx="1"/><path d="M3 9h18M3 14h18M9 4v16M15 4v16"/>',
  // 自動調整: a table with arrows pushing its sides out.
  autoFit: '<rect x="6" y="5" width="12" height="14" rx="1"/><path d="M6 12h12M12 5v14M2 12h2M20 12h2M3 10l-1 2 1 2M21 10l1 2-1 2"/>',
  picture: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m3 18 6-5 4 3 3-2 5 4"/>',
  shapes: '<rect x="3" y="3" width="10" height="8" rx="1"/><circle cx="16" cy="16" r="5"/><path d="M8 11v4h3"/>',
  textbox: '<rect x="3" y="4" width="18" height="16" rx="1"/><path d="M7 8h10M12 8v9"/>',
  select: '<path d="M5 3v15l4-4 3 7 3-1-3-7h6z"/>',
  link: '<path d="M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1"/><path d="M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1"/>',
  comment: '<path d="M4 4h16v12H9l-5 4z"/><path d="M12 7v6M9 10h6"/>',
  comments: '<path d="M3 4h13v9H8l-5 4z"/><path d="M19 8h2v12l-4-3H9v-2"/>',
  symbol: '<path d="M4 20h5v-2a7.5 7.5 0 1 1 6 0v2h5"/>',
  // 浮水印: a page with a letter turned across it.
  watermark: '<path d="M5 2h10l4 4v16H5z"/><path d="M15 2v4h4"/><path d="m9 17.5 3-8 3 8M10.1 14.5h3.8" transform="rotate(-40 12 14)"/>',
  header: '<rect x="4" y="2" width="16" height="20" rx="1"/><rect class="f" x="7" y="5" width="10" height="3"/><path d="M7 12h10M7 15h10M7 18h6"/>',
  footer: '<rect x="4" y="2" width="16" height="20" rx="1"/><path d="M7 6h10M7 9h10M7 12h6"/><rect class="f" x="7" y="16" width="10" height="3"/>',
  pageNumber: '<rect x="4" y="2" width="16" height="20" rx="1"/><text class="f" x="12" y="19" font-size="8" text-anchor="middle">1</text>',
  margins: '<rect x="4" y="2" width="16" height="20"/><rect x="7" y="5" width="10" height="14" stroke-dasharray="2 1.5"/>',
  orientation: '<rect x="3" y="2" width="11" height="15" rx="1"/><rect x="8" y="12" width="14" height="10" rx="1" class="w"/>',
  size: '<rect x="5" y="2" width="14" height="20" rx="1"/><path d="M8 18 16 6M12 6h4v4M8 14v4h4"/>',
  columns: '<path d="M3 4h7M3 8h7M3 12h7M3 16h7M3 20h7M14 4h7M14 8h7M14 12h7M14 16h7M14 20h7"/>',
  breaks: '<path d="M5 2v6h14V2M5 22v-6h14v6"/><path d="M2 12h20" stroke-dasharray="2 2"/>',
  paragraph: '<path d="M13 4v16M17 4v16M19 4h-9a4 4 0 0 0 0 8h3"/>',
  toc: '<path d="M4 6h9M4 11h9M4 16h6"/><path d="M21 13a5 5 0 1 1-2-5M20 4v4h-4"/>',
  tocInsert: '<rect x="3" y="2" width="18" height="20" rx="1"/><path d="M6 6h8M6 10h12M8 14h10M8 18h10"/>',
  spell: '<path d="M3 16 7 5l4 11M4.5 12h5"/><path d="m13 16 3 3 6-8"/>',
  wordCount: '<text class="f" x="12" y="10" font-size="8" text-anchor="middle">ABC</text><text class="f" x="12" y="20" font-size="8" text-anchor="middle">123</text>',
  track: '<path d="M4 20h4L20 8l-4-4L4 16z"/><path d="m13 7 4 4"/>',
  markup: '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  accept: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="m7 12 3 3 7-7"/>',
  reject: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="m8 8 8 8M16 8l-8 8"/>',
  prev: '<path d="m15 5-7 7 7 7"/>',
  next: '<path d="m9 5 7 7-7 7"/>',
  nav: '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M9 3v18M5 7h2M5 11h2M5 15h2"/>',
  zoom: '<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6M10 7v6M7 10h6"/>',
  zoom100: '<rect x="2" y="5" width="20" height="14" rx="2"/><text class="f" x="12" y="15.5" font-size="7.5" text-anchor="middle">100</text>',
  pageWidth: '<rect x="6" y="3" width="12" height="18"/><path d="M2 12h20M5 9l-3 3 3 3M19 9l3 3-3 3"/>',
  print: '<path d="M6 9V3h12v6"/><rect x="3" y="9" width="18" height="8" rx="1"/><path d="M6 14h12v7H6z"/>',
  download: '<path d="M12 3v12M7 10l5 5 5-5M4 21h16"/>',
  collapse: '<path d="m6 15 6-6 6 6"/>',
  pin: '<path d="M9 3h6l-1 6 4 4H6l4-4zM12 13v8"/>',
  launcher: '<path d="M5 5h7M5 5v7M19 19 8 8M19 19h-6M19 19v-6"/>',
};
