// Outlines of Word's preset shapes (a:prstGeom) and custom geometry (a:custGeom) as SVG paths.
//
// Only the shapes documents commonly draw with (boxes, flowchart symbols, arrows, lines and
// connectors) are known; anything else returns null and is shown as Word's preview picture.
// Sizes are in any unit (the caller uses px); adjust values (a:avLst) are in Word's 1/100000.

export interface GeomPath {
  d: string;
  /** Filled (a closed outline); false for lines and the extra strokes of some symbols. */
  fill: boolean;
  /** Stroked (drawn with the outline). */
  stroke: boolean;
}

export interface CustomPath {
  w: number;
  h: number;
  fill: boolean;
  stroke: boolean;
  /** Commands in path units: M/L x y, C x1 y1 x2 y2 x y, Q x1 y1 x y, A wR hR stAng swAng (degrees), Z. */
  cmds: (string | number)[][];
}

/** A rectangle inside the shape where its text goes (Word's text rectangle of the preset), in the shape's units. */
export interface TextRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

type Av = Record<string, number> | undefined;

const f = (n: number) => String(Math.round(n * 100) / 100);
const pts = (p: [number, number][]) => 'M' + p.map(([x, y]) => `${f(x)},${f(y)}`).join('L') + 'Z';
const poly = (p: [number, number][]) => [{ d: pts(p), fill: true, stroke: true }];
const adj = (av: Av, name: string, dflt: number) => (av && Number.isFinite(av[name]) ? av[name] : dflt) / 100000;
const ellipse = (cx: number, cy: number, rx: number, ry: number) =>
  `M${f(cx - rx)},${f(cy)}A${f(rx)},${f(ry)} 0 1 0 ${f(cx + rx)},${f(cy)}A${f(rx)},${f(ry)} 0 1 0 ${f(cx - rx)},${f(cy)}Z`;

function roundRect(w: number, h: number, r: number): string {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  return `M${f(r)},0H${f(w - r)}A${f(r)},${f(r)} 0 0 1 ${f(w)},${f(r)}V${f(h - r)}A${f(r)},${f(r)} 0 0 1 ${f(w - r)},${f(h)}` +
    `H${f(r)}A${f(r)},${f(r)} 0 0 1 0,${f(h - r)}V${f(r)}A${f(r)},${f(r)} 0 0 1 ${f(r)},0Z`;
}

/** Arrow pointing right in a w × h box: shaft thickness adj1 (of h), head length adj2 (of min(w, h)). */
function rightArrow(w: number, h: number, av: Av): [number, number][] {
  const ss = Math.min(w, h);
  const a1 = Math.min(1, adj(av, 'adj1', 50000));
  const a2 = adj(av, 'adj2', 50000);
  const dy = (h * a1) / 2;
  const x1 = Math.max(0, w - ss * a2);
  return [[0, h / 2 - dy], [x1, h / 2 - dy], [x1, 0], [w, h / 2], [x1, h], [x1, h / 2 + dy], [0, h / 2 + dy]];
}

/** The same outline turned for another direction: (x, y) in a box of size (w, h) of the right-pointing one. */
function turned(points: [number, number][], map: (x: number, y: number) => [number, number]): [number, number][] {
  return points.map(([x, y]) => map(x, y));
}

const PRESETS: Record<string, (w: number, h: number, av: Av) => GeomPath[]> = {
  rect: (w, h) => poly([[0, 0], [w, 0], [w, h], [0, h]]),
  flowChartProcess: (w, h) => poly([[0, 0], [w, 0], [w, h], [0, h]]),
  roundRect: (w, h, av) => [{ d: roundRect(w, h, Math.min(w, h) * adj(av, 'adj', 16667)), fill: true, stroke: true }],
  flowChartAlternateProcess: (w, h) => [{ d: roundRect(w, h, Math.min(w, h) * 0.16667), fill: true, stroke: true }],
  ellipse: (w, h) => [{ d: ellipse(w / 2, h / 2, w / 2, h / 2), fill: true, stroke: true }],
  flowChartConnector: (w, h) => [{ d: ellipse(w / 2, h / 2, w / 2, h / 2), fill: true, stroke: true }],
  diamond: (w, h) => poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]),
  flowChartDecision: (w, h) => poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]),
  parallelogram: (w, h, av) => {
    const x = Math.min(w, Math.min(w, h) * adj(av, 'adj', 25000));
    return poly([[x, 0], [w, 0], [w - x, h], [0, h]]);
  },
  flowChartInputOutput: (w, h) => poly([[w / 5, 0], [w, 0], [(w * 4) / 5, h], [0, h]]),
  trapezoid: (w, h, av) => {
    const x = Math.min(w / 2, Math.min(w, h) * adj(av, 'adj', 25000));
    return poly([[x, 0], [w - x, 0], [w, h], [0, h]]);
  },
  flowChartManualOperation: (w, h) => poly([[0, 0], [w, 0], [(w * 4) / 5, h], [w / 5, h]]),
  flowChartManualInput: (w, h) => poly([[0, h / 5], [w, 0], [w, h], [0, h]]),
  flowChartPreparation: (w, h) => poly([[0, h / 2], [w / 5, 0], [(w * 4) / 5, 0], [w, h / 2], [(w * 4) / 5, h], [w / 5, h]]),
  hexagon: (w, h, av) => {
    const x = Math.min(w / 2, Math.min(w, h) * adj(av, 'adj', 25000));
    return poly([[0, h / 2], [x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h]]);
  },
  octagon: (w, h, av) => {
    const x = Math.min(w / 2, h / 2, Math.min(w, h) * adj(av, 'adj', 29289));
    return poly([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]]);
  },
  triangle: (w, h, av) => poly([[w * adj(av, 'adj', 50000), 0], [w, h], [0, h]]),
  flowChartExtract: (w, h) => poly([[w / 2, 0], [w, h], [0, h]]),
  flowChartMerge: (w, h) => poly([[0, 0], [w, 0], [w / 2, h]]),
  rtTriangle: (w, h) => poly([[0, 0], [w, h], [0, h]]),
  homePlate: (w, h, av) => {
    const x = Math.max(0, w - Math.min(w, h) * adj(av, 'adj', 50000));
    return poly([[0, 0], [x, 0], [w, h / 2], [x, h], [0, h]]);
  },
  chevron: (w, h, av) => {
    const d = Math.min(w, Math.min(w, h) * adj(av, 'adj', 50000));
    return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]);
  },
  flowChartOffpageConnector: (w, h) => poly([[0, 0], [w, 0], [w, (h * 4) / 5], [w / 2, h], [0, (h * 4) / 5]]),
  flowChartPunchedCard: (w, h) => poly([[w / 5, 0], [w, 0], [w, h], [0, h], [0, h / 5]]),
  flowChartTerminator: (w, h) => {
    const rx = (w * 3475) / 21600;
    const ry = h / 2;
    return [{ d: `M${f(rx)},0H${f(w - rx)}A${f(rx)},${f(ry)} 0 0 1 ${f(w - rx)},${f(h)}H${f(rx)}A${f(rx)},${f(ry)} 0 0 1 ${f(rx)},0Z`, fill: true, stroke: true }];
  },
  flowChartDocument: (w, h) => {
    const y = (v: number) => (h * v) / 21600;
    const x = (v: number) => (w * v) / 21600;
    return [{ d: `M0,0H${f(w)}V${f(y(17322))}C${f(x(10800))},${f(y(17322))} ${f(x(10800))},${f(y(23922))} 0,${f(y(20172))}Z`, fill: true, stroke: true }];
  },
  flowChartPredefinedProcess: (w, h) => [
    { d: pts([[0, 0], [w, 0], [w, h], [0, h]]), fill: true, stroke: true },
    { d: `M${f(w / 8)},0V${f(h)}M${f((w * 7) / 8)},0V${f(h)}`, fill: false, stroke: true },
  ],
  flowChartInternalStorage: (w, h) => [
    { d: pts([[0, 0], [w, 0], [w, h], [0, h]]), fill: true, stroke: true },
    { d: `M${f(w / 5)},0V${f(h)}M0,${f(h / 5)}H${f(w)}`, fill: false, stroke: true },
  ],
  rightArrow: (w, h, av) => poly(rightArrow(w, h, av)),
  leftArrow: (w, h, av) => poly(turned(rightArrow(w, h, av), (x, y) => [w - x, y])),
  downArrow: (w, h, av) => poly(turned(rightArrow(h, w, av), (x, y) => [y, x])),
  upArrow: (w, h, av) => poly(turned(rightArrow(h, w, av), (x, y) => [y, h - x])),
  leftRightArrow: (w, h, av) => {
    const ss = Math.min(w, h);
    const dy = (h * Math.min(1, adj(av, 'adj1', 50000))) / 2;
    const x1 = Math.min(w / 2, ss * adj(av, 'adj2', 50000));
    return poly([[0, h / 2], [x1, 0], [x1, h / 2 - dy], [w - x1, h / 2 - dy], [w - x1, 0], [w, h / 2], [w - x1, h], [w - x1, h / 2 + dy], [x1, h / 2 + dy], [x1, h]]);
  },
  upDownArrow: (w, h, av) => {
    const ss = Math.min(w, h);
    const dx = (w * Math.min(1, adj(av, 'adj1', 50000))) / 2;
    const y1 = Math.min(h / 2, ss * adj(av, 'adj2', 50000));
    return poly([[w / 2, 0], [w, y1], [w / 2 + dx, y1], [w / 2 + dx, h - y1], [w, h - y1], [w / 2, h], [0, h - y1], [w / 2 - dx, h - y1], [w / 2 - dx, y1], [0, y1]]);
  },
  plus: (w, h, av) => {
    const d = Math.min(w, h) * adj(av, 'adj', 25000);
    return poly([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]);
  },
  // Lines and connectors: from the top left to the bottom right (flips turn them).
  line: (w, h) => [{ d: `M0,0L${f(w)},${f(h)}`, fill: false, stroke: true }],
  straightConnector1: (w, h) => [{ d: `M0,0L${f(w)},${f(h)}`, fill: false, stroke: true }],
  bentConnector2: (w, h) => [{ d: `M0,0H${f(w)}V${f(h)}`, fill: false, stroke: true }],
  bentConnector3: (w, h, av) => {
    const x1 = w * adj(av, 'adj1', 50000);
    return [{ d: `M0,0H${f(x1)}V${f(h)}H${f(w)}`, fill: false, stroke: true }];
  },
  bentConnector4: (w, h, av) => {
    const x1 = w * adj(av, 'adj1', 50000);
    const y2 = h * adj(av, 'adj2', 50000);
    return [{ d: `M0,0H${f(x1)}V${f(y2)}H${f(w)}V${f(h)}`, fill: false, stroke: true }];
  },
  bentConnector5: (w, h, av) => {
    const x1 = w * adj(av, 'adj1', 50000);
    const y2 = h * adj(av, 'adj2', 50000);
    const x3 = w * adj(av, 'adj3', 50000);
    return [{ d: `M0,0H${f(x1)}V${f(y2)}H${f(x3)}V${f(h)}H${f(w)}`, fill: false, stroke: true }];
  },
  curvedConnector3: (w, h, av) => {
    const x2 = w * adj(av, 'adj1', 50000);
    return [{ d: `M0,0C${f(x2 / 2)},0 ${f(x2)},${f(h / 4)} ${f(x2)},${f(h / 2)}C${f(x2)},${f((h * 3) / 4)} ${f((x2 + w) / 2)},${f(h)} ${f(w)},${f(h)}`, fill: false, stroke: true }];
  },
};

/** The preset shapes that can be drawn. */
export function isKnownPreset(prst: string): boolean {
  return Object.prototype.hasOwnProperty.call(PRESETS, prst);
}

/** Whether the preset is a line (no inside, no text rectangle of its own). */
export function isLinePreset(prst: string): boolean {
  return prst === 'line' || /Connector\d*$/.test(prst) && !prst.startsWith('flowChart');
}

/** The outline of a preset shape in a w × h box; null when it can't be drawn. */
export function presetPaths(prst: string, w: number, h: number, av?: Record<string, number>): GeomPath[] | null {
  const fn = PRESETS[prst];
  return fn ? fn(w, h, av) : null;
}

/** Where a preset shape's text goes (Word's text rectangle), in the same units as w and h. */
export function presetTextRect(prst: string, w: number, h: number, av?: Record<string, number>): TextRect {
  const box = (l: number, t: number, r: number, b: number): TextRect => ({ l, t, r, b });
  switch (prst) {
    case 'ellipse':
    case 'flowChartConnector': {
      const k = (1 - Math.SQRT1_2) / 2;
      return box(w * k, h * k, w * (1 - k), h * (1 - k));
    }
    case 'diamond':
    case 'flowChartDecision':
      return box(w / 4, h / 4, (w * 3) / 4, (h * 3) / 4);
    case 'flowChartInputOutput':
      return box(w / 5, 0, (w * 4) / 5, h);
    case 'parallelogram': {
      const x = Math.min(w, Math.min(w, h) * adj(av, 'adj', 25000)) / 2;
      return box(x, 0, w - x, h);
    }
    case 'flowChartDocument':
      return box(0, 0, w, (h * 17322) / 21600);
    case 'flowChartTerminator':
      return box((w * 1018) / 21600, (h * 3163) / 21600, (w * 20582) / 21600, (h * 18437) / 21600);
    case 'flowChartPredefinedProcess':
      return box(w / 8, 0, (w * 7) / 8, h);
    case 'flowChartInternalStorage':
      return box(w / 5, h / 5, w, h);
    case 'flowChartPreparation':
      return box(w / 5, 0, (w * 4) / 5, h);
    case 'flowChartManualInput':
      return box(0, h / 5, w, h);
    case 'flowChartManualOperation':
      return box(w / 5, 0, (w * 4) / 5, h);
    case 'flowChartOffpageConnector':
      return box(0, 0, w, (h * 4) / 5);
    case 'flowChartExtract':
    case 'flowChartMerge':
    case 'triangle':
      return box(w / 4, h / 2, (w * 3) / 4, h);
    case 'roundRect':
    case 'flowChartAlternateProcess': {
      const r = Math.min(w, h) * (prst === 'roundRect' ? adj(av, 'adj', 16667) : 0.16667) * 0.29289;
      return box(r, r, w - r, h - r);
    }
    case 'rightArrow':
    case 'leftArrow': {
      const pts = rightArrow(w, h, av);
      const x1 = pts[1][0];
      return prst === 'rightArrow' ? box(0, pts[0][1], x1, pts[5][1]) : box(w - x1, pts[0][1], w, pts[5][1]);
    }
    default:
      return box(0, 0, w, h);
  }
}

/** A custom outline (a:custGeom) scaled to a w × h box. */
export function customPaths(paths: CustomPath[], w: number, h: number): GeomPath[] {
  return paths.map((p) => {
    const sx = p.w > 0 ? w / p.w : 1;
    const sy = p.h > 0 ? h / p.h : 1;
    let d = '';
    let cx = 0;
    let cy = 0;
    for (const c of p.cmds) {
      const [op, ...n] = c as [string, ...number[]];
      if (op === 'M' || op === 'L') {
        cx = n[0] * sx;
        cy = n[1] * sy;
        d += `${op}${f(cx)},${f(cy)}`;
      } else if (op === 'C') {
        d += `C${f(n[0] * sx)},${f(n[1] * sy)} ${f(n[2] * sx)},${f(n[3] * sy)} ${f(n[4] * sx)},${f(n[5] * sy)}`;
        cx = n[4] * sx;
        cy = n[5] * sy;
      } else if (op === 'Q') {
        d += `Q${f(n[0] * sx)},${f(n[1] * sy)} ${f(n[2] * sx)},${f(n[3] * sy)}`;
        cx = n[2] * sx;
        cy = n[3] * sy;
      } else if (op === 'A') {
        // DrawingML arcTo: an arc of the ellipse (wR, hR) from the current point, starting at
        // angle stAng and turning swAng (angles as seen on the ellipse).
        const rx = n[0] * sx;
        const ry = n[1] * sy;
        const param = (deg: number) => Math.atan2(rx * Math.sin((deg * Math.PI) / 180), ry * Math.cos((deg * Math.PI) / 180));
        const t1 = param(n[2]);
        const t2 = param(n[2] + n[3]);
        const ox = cx - rx * Math.cos(t1);
        const oy = cy - ry * Math.sin(t1);
        const ex = ox + rx * Math.cos(t2);
        const ey = oy + ry * Math.sin(t2);
        const large = Math.abs(n[3]) > 180 ? 1 : 0;
        const sweep = n[3] > 0 ? 1 : 0;
        d += `A${f(rx)},${f(ry)} 0 ${large} ${sweep} ${f(ex)},${f(ey)}`;
        cx = ex;
        cy = ey;
      } else if (op === 'Z') d += 'Z';
    }
    return { d, fill: p.fill, stroke: p.stroke };
  });
}

// ----- connection sites (a:cxnLst of Word's preset shape definitions) -----

/** Where a connector can attach to a shape, in the shape's box, and which way it leaves it (degrees: 0 right, 90 down, 180 left, 270 up). */
export interface Site {
  x: number;
  y: number;
  dir: number;
}

const SITES_4 = (w: number, h: number, l = 0, r = w, b = h): Site[] => [
  { x: w / 2, y: 0, dir: 270 },
  { x: l, y: h / 2, dir: 180 },
  { x: w / 2, y: b, dir: 90 },
  { x: r, y: h / 2, dir: 0 },
];

/**
 * The connection sites of a preset shape in a w × h box, in Word's order (a:stCxn / a:endCxn
 * idx): top, left, bottom, right for boxes, flowchart symbols and diamonds; eight round an
 * ellipse (from the top, anticlockwise); the two ends of a line.
 */
export function connectionSites(prst: string, w: number, h: number, av?: Record<string, number>): Site[] {
  const k = (1 - Math.SQRT1_2) / 2;
  switch (prst) {
    case 'ellipse':
    case 'flowChartConnector':
      return [
        { x: w / 2, y: 0, dir: 270 }, { x: w * k, y: h * k, dir: 270 }, { x: 0, y: h / 2, dir: 180 }, { x: w * k, y: h * (1 - k), dir: 90 },
        { x: w / 2, y: h, dir: 90 }, { x: w * (1 - k), y: h * (1 - k), dir: 90 }, { x: w, y: h / 2, dir: 0 }, { x: w * (1 - k), y: h * k, dir: 270 },
      ];
    case 'flowChartInputOutput':
      return SITES_4(w, h, w / 10, (w * 9) / 10);
    case 'parallelogram': {
      const x = Math.min(w, Math.min(w, h) * adj(av, 'adj', 25000)) / 2;
      return SITES_4(w, h, x, w - x);
    }
    case 'flowChartDocument':
      return SITES_4(w, h, 0, w, (h * 20320) / 21600);
    case 'flowChartManualOperation':
      return SITES_4(w, h, w / 10, (w * 9) / 10);
    case 'triangle':
    case 'flowChartExtract':
      return [{ x: w / 2, y: 0, dir: 270 }, { x: w / 4, y: h / 2, dir: 180 }, { x: 0, y: h, dir: 90 }, { x: w / 2, y: h, dir: 90 }, { x: w, y: h, dir: 90 }, { x: (w * 3) / 4, y: h / 2, dir: 0 }];
    case 'line':
    case 'straightConnector1':
    case 'bentConnector2':
    case 'bentConnector3':
    case 'bentConnector4':
    case 'bentConnector5':
    case 'curvedConnector3':
      return [{ x: 0, y: 0, dir: 180 }, { x: w, y: h, dir: 0 }];
    default:
      return SITES_4(w, h);
  }
}

/** A connector's box for going from `a` to `b` (straight or bent), as DrawingML writes it. */
export interface ConnectorBox {
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  av?: Record<string, number>;
}

/**
 * The box, turn and flips of a connector from `a` to `b` (leaving `a` in direction `da`, arriving
 * at `b` from direction `db`, degrees as Site.dir): a straight connector spans the two points; a
 * bent one (bentConnector3) is turned so its first segment leaves `a` the way the site faces, as
 * Word does, with its middle segment halfway.
 */
export function routeConnector(prst: string, a: { x: number; y: number }, b: { x: number; y: number }, da?: number): ConnectorBox {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const flat = (): ConnectorBox => {
    const out: ConnectorBox = { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(dx), h: Math.abs(dy) };
    if (dx < 0) out.flipH = true;
    if (dy < 0) out.flipV = true;
    return out;
  };
  if (!prst.startsWith('bentConnector') && !prst.startsWith('curvedConnector')) return flat();
  // A bent connector's path leaves its box's top left corner along its top edge: of the ways to
  // turn and flip the box that put its ends on a and b, the one whose first segment leaves `a`
  // the way the site faces (else the unturned one).
  const cx = (a.x + b.x) / 2;
  const cy = (a.y + b.y) / 2;
  const near = (p: { x: number; y: number }, q: { x: number; y: number }) => Math.abs(p.x - q.x) < 0.01 + 1e-6 * Math.abs(q.x) && Math.abs(p.y - q.y) < 0.01 + 1e-6 * Math.abs(q.y);
  for (const rot of da == null ? [0] : [0, 90, 180, 270]) {
    const turned = rot === 90 || rot === 270;
    const w = turned ? Math.abs(dy) : Math.abs(dx);
    const h = turned ? Math.abs(dx) : Math.abs(dy);
    for (const flipH of [false, true]) {
      for (const flipV of [false, true]) {
        const box: ConnectorBox = { x: cx - w / 2, y: cy - h / 2, w, h, av: { adj1: 50000 } };
        if (rot) box.rot = rot;
        if (flipH) box.flipH = true;
        if (flipV) box.flipV = true;
        const ends = connectorEnds(box);
        if (!near(ends.a, a) || !near(ends.b, b)) continue;
        const leaves = (((flipH ? 180 : 0) + rot) % 360 + 360) % 360;
        if (da == null || leaves === ((da % 360) + 360) % 360) return box;
      }
    }
  }
  return { ...flat(), av: { adj1: 50000 } };
}

/** Where a connector's ends are on the page, from its box (the inverse of routeConnector). */
export function connectorEnds(box: ConnectorBox): { a: { x: number; y: number }; b: { x: number; y: number } } {
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const local = (x: number, y: number) => {
    let lx = box.flipH ? box.w - x : x;
    let ly = box.flipV ? box.h - y : y;
    lx -= box.w / 2;
    ly -= box.h / 2;
    const t = ((box.rot ?? 0) * Math.PI) / 180;
    return { x: cx + lx * Math.cos(t) - ly * Math.sin(t), y: cy + lx * Math.sin(t) + ly * Math.cos(t) };
  };
  return { a: local(0, 0), b: local(box.w, box.h) };
}
