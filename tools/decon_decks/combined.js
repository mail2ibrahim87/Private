// Delight International – combined decontamination deck ("white stage, brand-coloured actors").
// Usage: node combined.js <out.pptx> <logo.png>    then: python3 animate.py <out.pptx> <out.anim.json> <final.pptx>
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const gi = require('react-icons/gi');
const fa = require('react-icons/fa');

const OUT = process.argv[2];
const LOGO = 'image/png;base64,' + fs.readFileSync(process.argv[3]).toString('base64');
const C = {
  soft: 'E9F2F9', tintB: 'D3E6F4', tintG: 'D9EFDC', head: '1F5F8B', dark: '1F3A52', blue: '0C74BC', green: '3AB54A', greenD: '2E8B3A',
  body: '374151', muted: '5B7085', panel: 'F3F7FA', border: 'E1E8EF', line: 'C6D3E0', eqL: '94A3B8',
  // option D: oily brown -> Delight blue
  dirty: ['5C4630', '7A5C3E', 'A08467'], streak: '3E2F20', grime: ['2F3437', '4A4F52', '6B7175'], clean: 'BFDDF2', shine: 'E3F0FA',
  h2s: 'C0392B', benz: 'E67E22', fes: '6B7280', lel: 'D4A017', steam: '0C74BC', salt: '8E9AAF', oxide: 'A0522D', water: 'BFDDF2',
  effl: 'B8A68F', glow: 'FF8C1A', oil: 'D9A441', amm: '0C74BC',
};
const W = 13.333;

async function icon(lib, name, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(lib[name], { color: '#' + color, size: '256' }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}

(async () => {
  const p = new pptxgen();
  p.layout = 'LAYOUT_WIDE';
  p.theme = { headFontFace: 'Cambria', bodyFontFace: 'Calibri' };
  p.title = 'Decontamination across the oil & gas value chain';
  p.author = 'Delight International';
  p.defineSlideMaster({ title: 'SOFT', background: { color: C.soft }, objects: [] });
  p.defineSlideMaster({ title: 'WHITE', background: { color: 'FFFFFF' }, objects: [{ image: { x: 12.33, y: 0.3, w: 0.55, h: 0.72, data: LOGO } }],
    slideNumber: { x: 12.35, y: 7.05, w: 0.6, h: 0.3, fontFace: 'Calibri', fontSize: 10, color: C.muted, align: 'right' } });

  const spec = [];
  let s, A, uid = 0;
  const slide = (master = 'WHITE', tr = 'fade') => { s = p.addSlide({ masterName: master }); A = { transition: tr, steps: [] }; spec.push(A); return s; };
  const st = (trigger, effect, targets, extra = {}) => A.steps.push({ trigger, effect, targets: [].concat(targets), ...extra });
  const shp = (type, o, name) => { const n = name || `s${++uid}`; s.addShape(type, { ...o, objectName: n }); return n; };
  const txt = (t, o, name) => { const n = name || `t${++uid}`; s.addText(t, { fontFace: 'Calibri', margin: 0, isTextBox: true, color: C.body, valign: 'middle', ...o, objectName: n }); return n; };
  const img = (d, o, name) => { const n = name || `i${++uid}`; s.addImage({ data: d, ...o, objectName: n }); return n; };
  const ln = (x1, y1, x2, y2, color, width, arrow, name) => shp(p.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1, line: { color, width, endArrowType: arrow ? 'triangle' : undefined } }, name);
  const poly = (pts, color, width, arrow) => pts.slice(1).map((pt, i) => ln(pts[i][0], pts[i][1], pt[0], pt[1], color, width, arrow && i === pts.length - 2));
  const title = (t) => txt(t, { x: 0.6, y: 0.38, w: 11.3, h: 0.85, fontFace: 'Cambria', fontSize: 28, bold: true, color: C.head }, 'title');
  const lbl = (t, x, y, w, color, align = 'left', size = 11) => txt(t, { x, y, w, h: 0.3, fontSize: size, bold: true, color, align });
  const tag = (t, x, y, w, fill, color = 'FFFFFF', name) => { const n = name || `tag${++uid}`;
    s.addText(t, { shape: p.shapes.ROUNDED_RECTANGLE, rectRadius: 0.18, x, y, w, h: 0.36, fill: { color: fill }, color, fontSize: 12, bold: true,
      align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0, objectName: n }); return n; };
  const dot = (x, y, col, d = 0.22) => shp(p.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { color: 'FFFFFF', width: 0.75 } });
  const card = (x, y, w, h, fill = C.panel) => shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { color: C.border, width: 1 },
    shadow: { type: 'outer', color: '9AA9B8', opacity: 0.22, blur: 6, offset: 2, angle: 90 } });
  const circleIcon = async (x, y, d, lib, name, fill) => [shp(p.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } }),
    img(await icon(lib, name, 'FFFFFF'), { x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 })];
  const bullets = (items, o) => txt(items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1 } })),
    { fontSize: 14, paraSpaceAfter: 6, valign: 'top', ...o });
  const notes = (t) => s.addNotes(t);
  const check = await icon(fa, 'FaCheck', 'FFFFFF');

  // ---------- shared method-slide parts ----------
  // vessel with staged cleaning (option D): base = clean Delight blue; overlays light -> mid -> dark brown on top
  const staged = (x, y, w, h, r, pal = C.dirty, streaks = true) => {
    shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r, fill: { color: C.clean }, line: { color: C.blue, width: 2 } });
    shp(p.shapes.ROUNDED_RECTANGLE, { x: x + w * 0.12, y: y + Math.min(r, h * 0.2) * 0.6, w: w * 0.18, h: h - Math.min(r, h * 0.2) * 1.2, rectRadius: w * 0.09, fill: { color: C.shine }, line: { type: 'none' } });
    const layers = [2, 1, 0].map((k) => shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: r, fill: { color: pal[k] }, line: { color: C.eqL, width: 2 } }));
    const drips = streaks ? [[0.55, 0.18, 0.24], [0.72, 0.48, 0.2], [0.42, 0.66, 0.16]].map(([fx, fy, fh]) =>
      shp(p.shapes.ROUNDED_RECTANGLE, { x: x + w * fx, y: y + h * fy, w: 0.1, h: h * fh, rectRadius: 0.05, fill: { color: C.streak }, line: { type: 'none' } })) : [];
    return { dark: layers[2], mid: layers[1], light: layers[0], drips, x, y, w, h };
  };
  const shineSweep = (v, trig = 'after') => {
    const bar = shp(p.shapes.ROUNDED_RECTANGLE, { x: v.x + v.w * 0.08, y: v.y + 0.1, w: v.w * 0.84, h: 0.18, rectRadius: 0.09, fill: { color: 'FFFFFF', transparency: 35 }, line: { type: 'none' } });
    st(trig, 'appear', bar); st('with', 'path', bar, { path: [[0, v.h - 0.45]], dur: 1100 }); st('with', 'exit', bar, { dur: 300, delay: 1000 });
  };
  const dripsAway = (v, delay = 0) => v.drips.forEach((d, k) => { st('with', 'path', d, { path: [[0, v.h * 0.25]], dur: 1200, delay: delay + k * 200 }); st('with', 'exit', d, { dur: 400, delay: delay + 900 + k * 200 }); });
  const flare = (x, y) => {
    shp(p.shapes.RECTANGLE, { x, y, w: 0.2, h: 1.3, fill: { color: 'B4C2D3' }, line: { color: C.eqL, width: 1 } });
    const f1 = shp(p.shapes.OVAL, { x: x - 0.14, y: y - 0.55, w: 0.48, h: 0.6, fill: { color: 'F7C548' }, line: { type: 'none' } });
    const f2 = shp(p.shapes.OVAL, { x: x - 0.04, y: y - 0.32, w: 0.28, h: 0.36, fill: { color: C.glow }, line: { type: 'none' } });
    txt('FLARE /\nCLOSED VENT', { x: x - 0.5, y: y + 1.35, w: 1.2, h: 0.45, fontSize: 11, bold: true, color: C.muted, align: 'center' });
    return [f1, f2];
  };
  const tankBox = (x, y, w, h, label, fillCol = C.effl) => {
    shp(p.shapes.RECTANGLE, { x, y, w, h, fill: { color: 'EEF2F6' }, line: { color: C.eqL, width: 1.5 } });
    const l1 = shp(p.shapes.RECTANGLE, { x: x + 0.05, y: y + h - 0.35, w: w - 0.1, h: 0.3, fill: { color: fillCol }, line: { type: 'none' } });
    const l2 = shp(p.shapes.RECTANGLE, { x: x + 0.05, y: y + h - 0.75, w: w - 0.1, h: 0.4, fill: { color: fillCol }, line: { type: 'none' } });
    lbl(label, x - 0.4, y + h + 0.03, w + 0.8, C.muted, 'center');
    return [l1, l2];
  };
  const gasPanel = (x, y, w, rows = ['LEL', 'H₂S', 'Benzene']) => {
    shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.55 + rows.length * 0.45, rectRadius: 0.12, fill: { color: C.panel }, line: { color: C.border, width: 1 } });
    lbl('GAS TEST AT MANWAY', x + 0.2, y + 0.12, w - 0.3, C.greenD, 'left', 12);
    return rows.map((g, k) => {
      txt(g, { x: x + 0.2, y: y + 0.55 + k * 0.45, w: 1.1, h: 0.38, fontSize: 14, bold: true, color: C.dark });
      const bad = txt('✗  above limit', { x: x + 1.3, y: y + 0.55 + k * 0.45, w: w - 1.4, h: 0.38, fontSize: 13, bold: true, color: C.h2s });
      const ok = txt('✓  OK', { x: x + 1.3, y: y + 0.55 + k * 0.45, w: w - 1.4, h: 0.38, fontSize: 13, bold: true, color: C.greenD });
      return [bad, ok];
    });
  };
  const captions = (list) => {
    const box = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.4, w: 11.5, h: 0.78, rectRadius: 0.2, fill: { color: 'EAF3FA' }, line: { type: 'none' } });
    const caps = list.map((c, k) => txt(`${k + 1}   ${c}`, { x: 0.85, y: 6.4, w: 11.0, h: 0.78, fontSize: 14, color: C.dark }));
    st('click', 'fade', [box, caps[0]], { dur: 400 });
    return caps;
  };
  const nextCap = (caps, k) => { st('click', 'vanish', caps[k - 1]); st('with', 'fade', caps[k], { dur: 300 }); };
  const badge = (t, x, y, w) => {
    const b = shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.7, rectRadius: 0.35, fill: { color: C.green }, line: { type: 'none' } });
    const bt = txt(t, { x, y, w, h: 0.7, fontSize: 15, bold: true, color: 'FFFFFF', align: 'center' });
    st('after', 'zoom', [b, bt], { dur: 500 });
    st('after', 'pulse', [b, bt], { dur: 400, scale: 1.08, repeat: 2 });
  };
  const gasPass = (meter, x, y, w, t = '✓  READY FOR SAFE ENTRY') => {
    meter.forEach(([bad, ok], k) => { st('with', 'vanish', bad, { delay: 300 + k * 600 }); st('with', 'zoom', ok, { dur: 400, delay: 300 + k * 600 }); });
    badge(t, x, y, w);
  };
  const travel = (col, pts, delay, dur, d = 0.22, trig = 'with') => {
    const [x0, y0] = pts[0];
    const n = dot(x0 - d / 2, y0 - d / 2, col, d);
    st(trig, 'appear', n, { delay }); st('with', 'path', n, { path: pts.slice(1).map(([x, y]) => [x - x0, y - y0]), dur, delay }); st('with', 'vanish', n, { delay: delay + dur });
    return n;
  };
  const packets = (col, pts, n, gap, dur, start = 0, d = 0.2) => { for (let k = 0; k < n; k++) travel(col, pts, start + k * gap, dur, d); };
  const morph = (a, aPos, b, bPos) => { st('after', 'exit', a, { dur: 450 }); st('with', 'fade', b, { dur: 450 }); };

  // ===================================================== 1 TITLE
  slide('SOFT');
  shp(p.shapes.OVAL, { x: 9.0, y: -1.4, w: 5.8, h: 5.8, fill: { color: C.tintB }, line: { type: 'none' } });
  shp(p.shapes.OVAL, { x: 10.6, y: 4.3, w: 3.8, h: 3.8, fill: { color: C.tintG }, line: { type: 'none' } });
  img(LOGO, { x: 0.75, y: 0.6, w: 1.0, h: 1.31 });
  txt('DECONTAMINATION SERVICES', { x: 0.75, y: 2.55, w: 8.5, h: 0.4, fontSize: 16, bold: true, color: C.greenD, charSpacing: 4 });
  txt('Safe, fast entry – across the oil & gas value chain', { x: 0.75, y: 3.0, w: 8.8, h: 1.7, fontFace: 'Cambria', fontSize: 40, bold: true, color: C.dark, valign: 'top' });
  txt('Delight International  ·  Maintenance Simplified', { x: 0.75, y: 6.55, w: 8, h: 0.4, fontSize: 14, color: C.muted });
  notes('Delight decontamination: where it is needed in oil & gas, how we do it, and what happens to every contaminant.');

  // ===================================================== 2 WHO WE ARE
  slide();
  title('Delight International at a glance');
  const who = [[fa, 'FaGlobeAsia', 'Regional presence', 'UAE, Saudi Arabia, Oman, Qatar and India'],
    [fa, 'FaCertificate', 'Certified QHSE', 'ISO 9001, ISO 14001 and ISO 45001'],
    [fa, 'FaIndustry', 'Sectors we serve', 'Onshore, offshore, marine, power and oil & gas'],
    [fa, 'FaLayerGroup', 'Integrated services', 'Plant maintenance, specialised services, flange management, pipeline, marine and tank cleaning']];
  for (const [i, [lib, ic, t, d]] of who.entries()) {
    const x = 0.6 + i * 3.08;
    const c1 = card(x, 1.6, 2.85, 3.6);
    const ci = await circleIcon(x + 0.9, 1.95, 1.05, lib, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x: x + 0.15, y: 3.15, w: 2.55, h: 0.5, fontSize: 17, bold: true, color: C.head, align: 'center' });
    const dd = txt(d, { x: x + 0.2, y: 3.7, w: 2.45, h: 1.3, fontSize: 14, align: 'center', valign: 'top' });
    st(i === 0 ? 'click' : 'after', 'ascend', [c1, ...ci, tt, dd], { dur: 600 });
  }
  const wb = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.6, w: 12.1, h: 0.9, rectRadius: 0.1, fill: { color: 'EAF6EC' }, line: { type: 'none' } });
  const wt = txt([{ text: 'Decontamination at Delight:  ', options: { bold: true, color: C.greenD } }, { text: 'vapour-phase, boil-out, viscosity flush, chemical circulation, packing pre- & post-treatment and tank gamma-jet – planned, executed and validated by one team.' }],
    { x: 0.85, y: 5.6, w: 11.6, h: 0.9, fontSize: 15, color: C.dark });
  st('click', 'fade', [wb, wt]);

  // ===================================================== 3–7 TREE + MORPH BRANCHES
  const BR = [
    { k: 'UPSTREAM', sub: 'Production & gas processing', icon: [gi, 'GiOilRig'], col: C.blue,
      leaves: ['Oil production stations', 'Gas processing plants', 'Produced water & tank farms'],
      rows: [['Separators, heater-treaters, desalters', [['H₂S', C.h2s], ['LEL', C.lel], ['Sludge', C.dirty[1]], ['FeS', C.fes]]],
        ['Slug catchers, amine, glycol, condensate stabilisers', [['H₂S', C.h2s], ['Mercaptans', C.salt], ['FeS', C.fes], ['Benzene', C.benz]]],
        ['Produced-water and condensate tanks', [['Sludge', C.dirty[1]], ['H₂S', C.h2s], ['LEL', C.lel]]]] },
    { k: 'MIDSTREAM', sub: 'Transport & storage', icon: [gi, 'GiPipes'], col: C.green,
      leaves: ['Pipelines & pig receivers', 'Crude terminals & tank farms', 'LPG / NGL plants & spheres'],
      rows: [['Pig traps, receivers and slug catchers', [['FeS', C.fes], ['Pigging debris', C.dirty[1]], ['LEL', C.lel]]],
        ['Crude and product storage tanks', [['Sludge', C.dirty[1]], ['Benzene', C.benz], ['LEL', C.lel], ['H₂S', C.h2s]]],
        ['Fractionation columns, LPG spheres and bullets', [['LEL', C.lel], ['Mercaptans', C.salt], ['FeS', C.fes]]]] },
    { k: 'DOWNSTREAM', sub: 'Refining', icon: [gi, 'GiFactory'], col: C.blue,
      leaves: ['Crude & vacuum units', 'Hydroprocessing & conversion', 'Sour water, amine & sulphur'],
      rows: [['CDU, VDU, desalter, pre-heat trains', [['H₂S', C.h2s], ['Benzene', C.benz], ['LEL', C.lel], ['Heavy oil', C.dirty[1]]]],
        ['Hydrotreaters, hydrocracker, FCC, coker', [['H₂S', C.h2s], ['FeS', C.fes], ['Coke', C.dark], ['LEL', C.lel]]],
        ['Sour water strippers, amine regenerators, SRU / TGT', [['H₂S', C.h2s], ['Ammonia', C.amm], ['FeS', C.fes]]]] },
    { k: 'PETROCHEMICALS', sub: 'Chemicals & polymers', icon: [gi, 'GiChemicalDrop'], col: C.green,
      leaves: ['Aromatics (BTX)', 'Olefins & polymers', 'Methanol & fertilisers'],
      rows: [['Reformer, extraction and xylene columns, aromatics tanks', [['Benzene', C.benz], ['LEL', C.lel], ['Heavy aromatics', C.dirty[1]]]],
        ['Quench towers, caustic tower, compressor KO drums', [['Tar / coke', C.dark], ['Polymer', C.salt], ['Benzene', C.benz]]],
        ['Desulphurisation, synthesis, distillation, ammonia sections', [['LEL', C.lel], ['Ammonia', C.amm], ['FeS', C.fes]]]] },
  ];
  const CX = [1.95, 5.1, 8.25, 11.4];
  const brIcons = await Promise.all(BR.map((b) => icon(b.icon[0], b.icon[1], 'FFFFFF')));
  const tree = (sc, ox, oy, opts = {}) => {
    const T = (x) => ox + x * sc, U = (y) => oy + y * sc, F = (f) => Math.max(1, Math.round(f * sc * 10) / 10);
    const nm = { root: [], bar: [], br: [[], [], [], []], leaves: [[], [], [], []], marks: [[], [], [], []], spine: [] };
    nm.root.push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(4.57), y: U(1.4), w: 4.2 * sc, h: 0.72 * sc, rectRadius: 0.36 * sc, fill: { color: C.dark }, line: { type: 'none' } }, '!!root'));
    nm.root.push(txt('OIL & GAS VALUE CHAIN', { x: T(4.57), y: U(1.4), w: 4.2 * sc, h: 0.72 * sc, fontSize: F(16), bold: true, color: 'FFFFFF', align: 'center', charSpacing: 1 }, '!!root_t'));
    nm.bar.push(ln(T(6.67), U(2.12), T(6.67), U(2.4), C.line, 2.5 * sc, false, '!!bar_v'));
    nm.bar.push(ln(T(CX[0]), U(2.4), T(CX[3]), U(2.4), C.line, 2.5 * sc, false, '!!bar_h'));
    BR.forEach((b, i) => {
      const out = opts.out === i;
      nm.spine.push(ln(T(CX[i]), U(2.4), T(CX[i]), U(out ? 2.7 : 5.95), C.line, 2.5 * sc, false, `!!spine${i}`));
      const faded = opts.highlight !== undefined && opts.highlight !== i;
      const bx = CX[i] - 1.4;
      nm.br[i].push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(bx), y: U(2.7), w: 2.8 * sc, h: 0.95 * sc, rectRadius: 0.12 * sc, fill: { color: b.col, transparency: faded ? 65 : 0 }, line: { type: 'none' } }, `!!br${i}`));
      nm.br[i].push(img(brIcons[i], { x: T(bx + 0.15), y: U(2.85), w: 0.62 * sc, h: 0.62 * sc, transparency: faded ? 60 : 0 }, `!!bri${i}`));
      nm.br[i].push(txt([{ text: b.k, options: { bold: true, fontSize: F(13.5), breakLine: true } }, { text: b.sub, options: { fontSize: F(11) } }],
        { x: T(bx + 0.85), y: U(2.7), w: 1.92 * sc, h: 0.95 * sc, color: 'FFFFFF', transparency: faded ? 50 : 0 }, `!!brt${i}`));
      if (out) return;
      b.leaves.forEach((l, k) => {
        const y = 3.95 + k * 0.72;
        nm.leaves[i].push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(bx), y: U(y), w: 2.8 * sc, h: 0.58 * sc, rectRadius: 0.29 * sc, fill: { color: 'FFFFFF', transparency: faded ? 40 : 0 }, line: { color: b.col, width: 1.25 * sc, transparency: faded ? 60 : 0 } }, `!!lf${i}_${k}`));
        nm.leaves[i].push(txt(l, { x: T(bx + 0.2), y: U(y), w: 2.15 * sc, h: 0.58 * sc, fontSize: F(12.5), bold: true, color: faded ? '9AA9B8' : C.dark }, `!!lft${i}_${k}`));
        nm.marks[i].push(shp(p.shapes.OVAL, { x: T(bx + 2.38), y: U(y + 0.11), w: 0.36 * sc, h: 0.36 * sc, fill: { color: C.green, transparency: faded ? 60 : 0 }, line: { color: 'FFFFFF', width: 1 } }, `!!mk${i}_${k}`));
        nm.marks[i].push(img(check, { x: T(bx + 2.47), y: U(y + 0.2), w: 0.18 * sc, h: 0.18 * sc }, `!!mki${i}_${k}`));
      });
    });
    return nm;
  };
  // 3 tree
  slide();
  title('Where decontamination fits in oil & gas');
  const t1 = tree(1, 0, 0);
  st('click', 'zoom', t1.root, { dur: 600 });
  st('after', 'wipe', t1.bar[0], { dir: 'up', dur: 300 });
  st('after', 'wipe', t1.bar[1], { dir: 'left', dur: 700 });
  BR.forEach((b, i) => st(i === 0 ? 'after' : 'with', 'ascend', t1.br[i], { dur: 700, delay: i * 250 }));
  BR.forEach((b, i) => {
    st('click', 'wipe', t1.spine[i], { dir: 'up', dur: 500 });
    t1.leaves[i].forEach((n, j) => { if (j % 2 === 0) st(j === 0 ? 'after' : 'with', 'ascend', [n, t1.leaves[i][j + 1]], { dur: 500, delay: (j / 2) * 200 }); });
  });
  const allMarks = t1.marks.flat();
  st('click', 'zoom', allMarks, { dur: 400 });
  const ban = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.2, w: 12.1, h: 0.6, rectRadius: 0.3, fill: { color: 'EAF6EC' }, line: { type: 'none' } });
  const bant = txt([{ text: '✓  ', options: { color: C.greenD, bold: true } }, { text: 'Every stage holds hydrocarbons, H₂S or pyrophoric scale – and needs decontamination before entry.' }],
    { x: 0.85, y: 6.2, w: 11.6, h: 0.6, fontSize: 15, color: C.dark });
  st('after', 'ascend', [ban, bant], { dur: 700 });
  allMarks.filter((n, k) => k % 2 === 0).forEach((n, k) => st(k === 0 ? 'after' : 'with', 'pulse', n, { dur: 400, scale: 1.3 }));
  notes('Build the value chain branch by branch. The green ticks mark where Delight decontaminates. The next four slides zoom into each branch.');
  // 4–7 branch zooms
  BR.forEach((b, i) => {
    slide('WHITE', 'morph');
    title(`${b.k.charAt(0) + b.k.slice(1).toLowerCase()} – ${b.sub.toLowerCase()}`);
    tree(0.3, 8.95, 0.85, { highlight: i, out: i });
    b.rows.forEach(([where, cont], k) => {
      const y = 1.6 + k * 1.55;
      shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 3.6, h: 0.8, rectRadius: 0.4, fill: { color: 'FFFFFF' }, line: { color: b.col, width: 2 } }, `!!lf${i}_${k}`);
      txt(b.leaves[k], { x: 0.85, y, w: 2.75, h: 0.8, fontSize: 16, bold: true, color: C.dark }, `!!lft${i}_${k}`);
      shp(p.shapes.OVAL, { x: 3.65, y: y + 0.17, w: 0.46, h: 0.46, fill: { color: C.green }, line: { color: 'FFFFFF', width: 1 } }, `!!mk${i}_${k}`);
      img(check, { x: 3.77, y: y + 0.29, w: 0.22, h: 0.22 }, `!!mki${i}_${k}`);
      const u = txt([{ text: 'Where: ', options: { bold: true, color: C.head } }, { text: where }], { x: 4.5, y: y - 0.3, w: 4.2, h: 0.7, fontSize: 14, valign: 'bottom' });
      let x0 = 4.5;
      const chips = cont.flatMap(([c, col]) => {
        const w = 0.35 + c.length * 0.085;
        const r = [shp(p.shapes.ROUNDED_RECTANGLE, { x: x0, y: y + 0.5, w, h: 0.36, rectRadius: 0.18, fill: { color: col }, line: { type: 'none' } }),
          txt(c, { x: x0, y: y + 0.5, w, h: 0.36, fontSize: 12, bold: true, color: 'FFFFFF', align: 'center' })];
        x0 += w + 0.12;
        return r;
      });
      st('click', 'ascend', [u, ...chips], { dur: 600 });
    });
    notes(`Morph zooms into ${b.k.toLowerCase()}: the tree shrinks to the corner and the branch's facilities fly out. Click to show where we work and what is removed.`);
  });

  // ===================================================== 8 YOUR FACILITY
  slide();
  title('Your facility');
  const fields = ['Facility / client', 'Units in scope', 'Key contaminants', 'Shutdown window', 'Utilities available (steam, N₂, water)', 'Entry criteria / permit requirements'];
  fields.forEach((f, k) => {
    const x = 0.6 + (k % 2) * 6.15, y = 1.55 + Math.floor(k / 2) * 1.6;
    const a = txt(f.toUpperCase(), { x, y, w: 5.9, h: 0.35, fontSize: 12, bold: true, color: C.greenD, charSpacing: 1 });
    const b2 = shp(p.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.4, w: 5.9, h: 0.9, rectRadius: 0.08, fill: { color: C.panel }, line: { color: C.border, width: 1 } });
    st(k === 0 ? 'click' : 'after', 'ascend', [a, b2], { dur: 400 });
  });
  txt('Filled in with the client before the meeting – their position on the value chain and what we will remove.', { x: 0.6, y: 6.45, w: 11.5, h: 0.4, fontSize: 12, italic: true, color: C.muted });
  notes('Fill in this slide for each client meeting (type into the grey boxes).');

  // ===================================================== 9 CONTAMINANTS
  slide();
  title('What has to be removed – and why');
  const hz = [['H₂S', C.h2s, 'Toxic gas – fatal at low concentrations'], ['Benzene', C.benz, 'Carcinogenic vapour – strict exposure limits'],
    ['LEL hydrocarbons', C.lel, 'Flammable / explosive atmosphere'], ['Pyrophoric FeS', C.fes, 'Self-heats and ignites when exposed to air'],
    ['Heavy oil & sludge', C.dirty[1], 'Holds gases, blocks inspection, slows mechanical cleaning'], ['Ammonia & mercaptans', C.amm, 'Toxic and odorous – sour-water and LPG units']];
  hz.forEach(([t, col, d], k) => {
    const x = 0.6 + (k % 3) * 4.1, y = 1.55 + Math.floor(k / 3) * 2.45;
    const c1 = card(x, y, 3.85, 2.15);
    const tg = tag(t, x + 0.25, y + 0.3, 2.6, col);
    const dd = txt(d, { x: x + 0.25, y: y + 0.85, w: 3.35, h: 1.1, fontSize: 15, valign: 'top' });
    st(k === 0 ? 'click' : 'after', 'ascend', [c1, tg, dd], { dur: 500 });
  });
  notes('The hazards that must be removed before entry – the colours are used for these contaminants throughout the deck.');

  // ===================================================== 10 OBJECTIVES
  slide();
  title('Decontamination objectives');
  const obj = [['De-oiling / de-sludging', 'Soften heavy deposits for easier removal; reduce hazardous sludge.'], ['Sulphide oxidation', 'Convert pyrophoric FeS and H₂S into stable, safe compounds.'],
    ['Degassing', 'Release LEL hydrocarbons and benzene to the flare / closed vent.'], ['Neutralising', 'Correct pH and treat ammonia so the effluent can be disposed of.']];
  for (const [i, [a, b]] of obj.entries()) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.55;
    const c1 = card(x, y, 5.9, 2.25);
    const ci = shp(p.shapes.OVAL, { x: x + 0.35, y: y + 0.45, w: 0.9, h: 0.9, fill: { color: i % 2 ? C.green : C.blue }, line: { type: 'none' } });
    const cn = txt(String(i + 1), { x: x + 0.35, y: y + 0.45, w: 0.9, h: 0.9, fontSize: 24, bold: true, color: 'FFFFFF', align: 'center' });
    const ta = txt(a, { x: x + 1.5, y: y + 0.4, w: 4.2, h: 0.5, fontSize: 20, bold: true, color: C.head });
    const tb = txt(b, { x: x + 1.5, y: y + 0.95, w: 4.2, h: 1.1, fontSize: 15, valign: 'top' });
    st('click', 'fade', [c1, ci, cn, ta, tb]);
  }

  // ===================================================== 11 DIVIDER methods
  const divider = (num, t, sub) => {
    slide('SOFT');
    shp(p.shapes.OVAL, { x: 9.6, y: 1.2, w: 5.0, h: 5.0, fill: { color: C.tintB }, line: { type: 'none' } });
    img(LOGO, { x: 12.33, y: 0.3, w: 0.55, h: 0.72 });
    txt(num, { x: 0.75, y: 1.75, w: 3, h: 1.2, fontFace: 'Cambria', fontSize: 66, bold: true, color: C.blue });
    const a = txt(t, { x: 0.75, y: 3.2, w: 8.5, h: 1.35, fontFace: 'Cambria', fontSize: 34, bold: true, color: C.dark, valign: 'bottom' });
    const b2 = txt(sub, { x: 0.75, y: 4.75, w: 8.5, h: 0.9, fontSize: 18, color: C.muted, valign: 'top' });
    st('with', 'ascend', [a, b2], { dur: 800 });
  };
  divider('01', 'How we decontaminate', 'Six methods – each shown with where every contaminant goes');

  // ===================================================== 12 WHICH METHOD WHERE
  slide();
  title('Which method, where?');
  const M = ['Vapour-phase', 'Boil-out', 'Viscosity flush', 'Circulation', 'Pre/post-treat', 'Gamma-jet'];
  const EQ = [['Columns & towers', [2, 0, 1, 0, 1, 0]], ['Drums, separators & desalters', [1, 2, 0, 0, 0, 0]], ['Heavy bottoms, coker & quench oil', [2, 1, 2, 0, 0, 0]],
    ['Exchangers & air coolers', [2, 0, 1, 2, 0, 0]], ['Packed beds & trays', [1, 0, 0, 0, 2, 0]], ['Tanks & temperature-limited vessels', [0, 0, 0, 2, 0, 2]]];
  const gx = 4.3, cw = 1.4;
  M.forEach((m, j) => txt(m, { x: gx + j * cw, y: 1.4, w: cw, h: 0.65, fontSize: 13, bold: true, color: C.head, align: 'center' }));
  EQ.forEach(([e, v], k) => {
    const y = 2.15 + k * 0.62;
    const row = [shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 12.1, h: 0.54, rectRadius: 0.08, fill: { color: k % 2 ? 'FFFFFF' : C.panel }, line: { type: 'none' } }),
      txt(e, { x: 0.8, y, w: 3.4, h: 0.54, fontSize: 14, bold: true, color: C.dark })];
    v.forEach((val, j) => {
      if (!val) return;
      const cx = gx + j * cw + cw / 2 - 0.15;
      row.push(shp(p.shapes.OVAL, { x: cx, y: y + 0.12, w: 0.3, h: 0.3, fill: { color: val === 2 ? C.green : 'FFFFFF' }, line: { color: C.green, width: 2 } }));
    });
    st(k === 0 ? 'click' : 'after', 'ascend', row, { dur: 450 });
  });
  const lg = txt([{ text: '●  ', options: { color: C.green } }, { text: 'main method     ' }, { text: '○  ', options: { color: C.green } }, { text: 'used in combination / where applicable' }],
    { x: 0.6, y: 6.0, w: 8, h: 0.35, fontSize: 13, color: C.muted });
  st('after', 'fade', lg);
  notes('The method is chosen per piece of equipment; most units combine two methods.');

  // ===================================================== 13 VAPOUR-PHASE
  slide();
  title('Vapour-phase – where each contaminant goes');
  const v1 = staged(5.2, 1.9, 1.3, 3.8, 0.65);
  for (let t = 1; t <= 7; t++) ln(5.25, 1.9 + 3.8 * t / 8, 6.45, 1.9 + 3.8 * t / 8, C.eqL, 1);
  lbl('Column', 5.2, 5.72, 1.3, C.muted, 'center');
  poly([[5.85, 1.9], [5.85, 1.55], [10.95, 1.55]], C.eqL, 3, true);
  const fl1 = flare(10.95, 1.25);
  poly([[5.85, 5.7], [5.85, 6.05], [9.2, 6.05], [9.2, 5.5], [9.6, 5.5]], C.eqL, 3, true);
  lbl('Condensate drain', 6.7, 5.75, 2.0, C.muted);
  const [e1, e2] = tankBox(9.6, 4.75, 1.6, 1.25, 'EFFLUENT TANK');
  lbl('STEAM', 0.6, 5.0, 0.9, C.steam, 'left', 12);
  const steamL = ln(1.3, 5.3, 5.2, 5.3, C.steam, 5, true);
  lbl('DECON CHEMISTRY', 2.2, 4.5, 1.9, C.greenD, 'left', 12);
  const chemL = ln(3.0, 4.82, 3.0, 5.3, C.green, 5, true);
  const meter = gasPanel(0.6, 1.5, 3.3);
  const caps = captions([
    'The unit is shut down, drained and isolated – the column still holds H₂S, benzene, LEL gases, pyrophoric FeS and heavy oil.',
    'Steam is introduced through the existing steam-out lines and heats the whole column.',
    'Decon chemistry is injected into the steam and reaches every surface the steam reaches.',
    'H₂S is converted into stable, non-volatile compounds that stay in the condensate and drain to the effluent tank.',
    'Pyrophoric FeS is oxidised in place to stable iron oxide – it can no longer self-heat – and is washed out.',
    'Benzene and LEL hydrocarbons are swept to the flare / closed vent; dissolved benzene drains with the condensate.',
    'Softened deposits and heavy oil drain with the condensate – the column is clean.',
    'Gas tests confirm the entry criteria – the column is handed over for safe entry.']);
  const tH = tag('H₂S', 5.35, 2.25, 1.0, C.h2s), tB = tag('Benzene', 5.35, 2.85, 1.0, C.benz), tL = tag('LEL', 5.35, 3.45, 1.0, C.lel);
  const glow = shp(p.shapes.ROUNDED_RECTANGLE, { x: 5.28, y: 4.0, w: 1.14, h: 0.5, rectRadius: 0.25, fill: { color: C.glow, transparency: 55 }, line: { color: C.glow, width: 2 } });
  const tF = tag('FeS', 5.35, 4.07, 1.0, C.fes);
  const tS = tag('stable salt', 5.25, 2.25, 1.2, C.salt), tO = tag('iron oxide', 5.25, 4.07, 1.2, C.oxide);
  st('with', 'pulse', glow, { dur: 500, scale: 1.08, repeat: 3, delay: 400 });
  const drain1 = (y) => [[5.85, y], [5.85, 6.05], [9.2, 6.05], [9.2, 5.5], [9.85, 5.5]];
  nextCap(caps, 1);
  st('with', 'wipe', steamL, { dir: 'left', dur: 700 });
  for (let k = 0; k < 7; k++) { const n = shp(p.shapes.OVAL, { x: 1.2, y: 5.2, w: 0.2, h: 0.2, fill: { color: 'FFFFFF' }, line: { color: C.steam, width: 2 } });
    const d = 700 + k * 330; st('with', 'appear', n, { delay: d }); st('with', 'path', n, { path: [[4.35, 0], [4.35, -3.0]], dur: 2200, delay: d }); st('with', 'vanish', n, { delay: d + 2200 }); }
  nextCap(caps, 2);
  st('with', 'wipe', chemL, { dir: 'down', dur: 600 });
  packets(C.green, [[3.0, 4.82], [3.0, 5.3], [5.75, 5.3], [5.75, 2.1]], 6, 360, 2400, 600);
  st('with', 'exit', v1.dark, { dur: 1500, delay: 1800 });
  nextCap(caps, 3);
  st('with', 'pulse', tH, { dur: 300, scale: 1.2, repeat: 2 });
  morph(tH, null, tS); st('after', 'exit', tS, { dur: 400, delay: 600 });
  for (let k = 0; k < 4; k++) travel(C.salt, drain1(2.43), 600 + k * 250, 2000);
  st('after', 'wipe', e1, { dir: 'down', dur: 600 });
  nextCap(caps, 4);
  st('with', 'exit', glow, { dur: 600 }); st('with', 'exit', tF, { dur: 600 }); st('with', 'fade', tO, { dur: 600 });
  dripsAway(v1, 300);
  st('after', 'exit', tO, { dur: 400, delay: 400 });
  for (let k = 0; k < 4; k++) travel(C.oxide, drain1(4.25), 400 + k * 250, 1800);
  nextCap(caps, 5);
  [[tB, 2.85], [tL, 3.45]].forEach(([n, y], k) => { st('with', 'path', n, { path: [[0, 1.37 - y], [5.4, 1.37 - y]], dur: 1800, delay: k * 400 }); st('with', 'exit', n, { dur: 400, delay: 1800 + k * 400 }); });
  st('with', 'pulse', fl1, { dur: 300, scale: 1.4, repeat: 4, delay: 1700 });
  for (let k = 0; k < 2; k++) travel(C.benz, drain1(3.03), 900 + k * 300, 1800, 0.18);
  nextCap(caps, 6);
  st('with', 'exit', v1.mid, { dur: 1400 });
  for (let k = 0; k < 4; k++) travel(C.dirty[0], drain1(5.2), k * 300, 1600);
  st('after', 'exit', v1.light, { dur: 1200 });
  shineSweep(v1);
  st('with', 'wipe', e2, { dir: 'down', dur: 700 });
  nextCap(caps, 7);
  gasPass(meter, 0.6, 3.7, 3.3);
  notes('Eight clicks. H₂S and FeS are converted in place and leave with the condensate; only the vapours go to the flare. The column cleans in stages from oily brown to Delight blue.');

  // ===================================================== 14 BOIL-OUT
  slide();
  title('Boil-out – vessels with sludge');
  const vb = staged(4.3, 2.2, 6.0, 2.4, 1.2);
  const sludge = shp(p.shapes.RECTANGLE, { x: 4.9, y: 4.0, w: 4.8, h: 0.42, fill: { color: C.dirty[0] }, line: { type: 'none' } });
  const water = shp(p.shapes.ROUNDED_RECTANGLE, { x: 4.65, y: 2.95, w: 5.3, h: 1.05, rectRadius: 0.3, fill: { color: C.water, transparency: 15 }, line: { type: 'none' } });
  lbl('Vessel / drum', 4.3, 4.62, 1.6, C.muted);
  poly([[9.4, 2.2], [9.4, 1.55], [11.6, 1.55]], C.eqL, 3, true);
  const fl2 = flare(11.6, 1.25);
  lbl('STEAM', 11.55, 4.95, 0.9, C.steam, 'left', 12);
  const header = ln(12.4, 5.3, 6.2, 5.3, C.steam, 5);
  const risers = [6.2, 7.3, 8.4].map((x) => ln(x, 5.3, x, 4.6, C.steam, 4, true));
  lbl('DECON CHEMISTRY', 9.3, 5.95, 1.9, C.greenD, 'left', 12);
  const chemB = ln(10.6, 5.9, 10.6, 5.3, C.green, 5, true);
  poly([[5.6, 4.6], [5.6, 5.55], [2.3, 5.55]], C.eqL, 3, true);
  const [bl1, bl2] = tankBox(0.7, 4.75, 1.6, 1.15, 'EFFLUENT TANK');
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 10.62, y: 2.4, w: 0.22, h: 1.9, rectRadius: 0.11, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1.5 } });
  const therm = shp(p.shapes.ROUNDED_RECTANGLE, { x: 10.66, y: 2.6, w: 0.14, h: 1.68, rectRadius: 0.07, fill: { color: C.h2s }, line: { type: 'none' } });
  shp(p.shapes.OVAL, { x: 10.55, y: 4.15, w: 0.36, h: 0.36, fill: { color: C.h2s }, line: { type: 'none' } });
  const hot = lbl('Boiling', 10.25, 2.05, 0.9, C.h2s, 'center');
  const meter2 = gasPanel(0.6, 1.5, 3.3);
  const caps2 = captions([
    'The vessel is drained and isolated – sludge on the bottom, H₂S, benzene and LEL above it, pyrophoric FeS in the deposits.',
    'Water is filled above the sludge level.',
    'Steam is injected at several points along the bottom and brings the water to the boil.',
    'Decon chemistry is injected with the steam; the rolling boil lifts and emulsifies the sludge.',
    'H₂S dissolves and is converted into stable compounds; FeS is oxidised to iron oxide – both stay in the water.',
    'Benzene and LEL vapours leave through the vent to the flare / closed vent.',
    'The vessel is drained to the effluent tank and rinsed – the vessel is clean.',
    'Gas tests confirm the entry criteria – the vessel is handed over for safe entry.']);
  const bH = tag('H₂S', 5.3, 2.4, 1.0, C.h2s), bB = tag('Benzene', 6.6, 2.4, 1.0, C.benz), bL = tag('LEL', 7.9, 2.4, 1.0, C.lel);
  const bG = shp(p.shapes.ROUNDED_RECTANGLE, { x: 8.43, y: 3.94, w: 1.14, h: 0.5, rectRadius: 0.25, fill: { color: C.glow, transparency: 55 }, line: { color: C.glow, width: 2 } });
  const bF = tag('FeS', 8.5, 4.01, 1.0, C.fes);
  const bS = tag('stable salt', 5.2, 3.3, 1.2, C.salt), bO = tag('iron oxide', 8.4, 3.3, 1.2, C.oxide);
  st('with', 'pulse', bG, { dur: 500, scale: 1.08, repeat: 3, delay: 400 });
  nextCap(caps2, 1);
  st('with', 'wipe', water, { dir: 'down', dur: 1800 });
  const bubble = (x, delay) => { const n = shp(p.shapes.OVAL, { x, y: 3.85, w: 0.16, h: 0.16, fill: { color: 'FFFFFF' }, line: { color: '7FB8E3', width: 1 } });
    st('with', 'appear', n, { delay }); st('with', 'path', n, { path: [[0.05, -0.45], [-0.03, -0.8]], dur: 1100, delay }); st('with', 'vanish', n, { delay: delay + 1100 }); };
  nextCap(caps2, 2);
  st('with', 'wipe', header, { dir: 'right', dur: 700 });
  st('after', 'wipe', risers, { dir: 'down', dur: 400 });
  packets('FFFFFF', [[12.3, 5.3], [8.4, 5.3], [8.4, 4.6]], 4, 300, 1300, 0, 0.18);
  st('after', 'wipe', therm, { dir: 'down', dur: 2000 });
  st('with', 'fade', hot, { dur: 300, delay: 1800 });
  for (let k = 0; k < 18; k++) bubble(4.9 + ((k * 0.53) % 4.9), 300 + k * 160);
  nextCap(caps2, 3);
  st('with', 'wipe', chemB, { dir: 'down', dur: 500 });
  [[6.2], [7.3], [8.4]].forEach(([x], j) => packets(C.green, [[10.6, 5.9], [10.6, 5.3], [x, 5.3], [x, 4.3]], 2, 450, 1500, 500 + j * 200));
  for (let k = 0; k < 18; k++) bubble(4.9 + ((k * 0.61) % 4.9), 900 + k * 140);
  st('with', 'exit', sludge, { dur: 2500, delay: 1200 });
  st('with', 'exit', vb.dark, { dur: 1500, delay: 1500 });
  for (let k = 0; k < 8; k++) { const n = dot(5.1 + k * 0.55, 4.05, C.dirty[1], 0.2);
    st('with', 'appear', n, { delay: 1300 + k * 150 }); st('with', 'path', n, { path: [[0.1, -0.35 - (k % 3) * 0.15]], dur: 1500, delay: 1300 + k * 150 }); }
  nextCap(caps2, 4);
  st('with', 'path', bH, { path: [[-0.1, 0.9]], dur: 900 });
  st('after', 'exit', bH, { dur: 400 }); st('with', 'fade', bS, { dur: 400 });
  st('with', 'exit', bG, { dur: 600 }); st('with', 'path', bF, { path: [[-0.1, -0.71]], dur: 800 });
  st('after', 'exit', bF, { dur: 400 }); st('with', 'fade', bO, { dur: 400 });
  dripsAway(vb, 0);
  nextCap(caps2, 5);
  [[bB, 6.6], [bL, 7.9]].forEach(([n, x], k) => { st('with', 'path', n, { path: [[8.9 - x, 0], [8.9 - x, -1.03], [11.1 - x, -1.03]], dur: 1800, delay: k * 400 }); st('with', 'exit', n, { dur: 400, delay: 1800 + k * 400 }); });
  st('with', 'pulse', fl2, { dur: 300, scale: 1.4, repeat: 4, delay: 1700 });
  nextCap(caps2, 6);
  st('with', 'wipeout', water, { dir: 'up', dur: 2200 });
  st('with', 'exit', [bS, bO], { dur: 600, delay: 300 });
  for (let k = 0; k < 6; k++) travel(k % 2 ? C.salt : C.dirty[1], [[5.6, 4.6], [5.6, 5.55], [2.15, 5.55]], 300 + k * 280, 1600);
  st('with', 'wipe', bl1, { dir: 'down', dur: 600, delay: 1200 }); st('with', 'wipe', bl2, { dir: 'down', dur: 600, delay: 2000 });
  st('with', 'exit', vb.mid, { dur: 1200, delay: 1500 });
  st('after', 'exit', vb.light, { dur: 1000 });
  shineSweep(vb);
  nextCap(caps2, 7);
  gasPass(meter2, 0.6, 3.7, 3.3);
  notes('Chemistry goes in with the steam through the bottom risers. The vessel cleans in stages from oily brown to Delight blue.');

  // ===================================================== 15 VISCOSITY FLUSH
  slide();
  title('Viscosity flush / oil wash – heavy deposits');
  const vf = staged(4.4, 1.8, 1.2, 3.8, 0.6, C.grime);
  lbl('Column bottoms', 4.1, 5.62, 1.8, C.muted, 'center');
  // exchanger with grime coating
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 7.8, y: 4.3, w: 2.2, h: 0.6, rectRadius: 0.3, fill: { color: C.clean }, line: { color: C.blue, width: 1.5 } });
  const ex3 = [2, 1, 0].map((k) => shp(p.shapes.ROUNDED_RECTANGLE, { x: 7.8, y: 4.3, w: 2.2, h: 0.6, rectRadius: 0.3, fill: { color: C.grime[k] }, line: { color: C.eqL, width: 1.5 } }));
  for (let k = 1; k <= 3; k++) ln(8.1, 4.3 + 0.15 * k, 9.8, 4.3 + 0.15 * k, C.eqL, 0.75);
  lbl('Exchanger', 7.8, 4.95, 2.2, C.muted, 'center');
  shp(p.shapes.OVAL, { x: 6.6, y: 5.75, w: 0.45, h: 0.45, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1.5 } });
  lbl('Pump', 6.45, 6.05, 0.8, C.muted, 'center', 10);
  const loopL = [poly([[5.0, 5.6], [5.0, 5.97], [6.6, 5.97]], C.oil, 4), poly([[7.05, 5.97], [10.6, 5.97], [10.6, 4.6], [10.0, 4.6]], C.oil, 4, true), poly([[7.8, 4.6], [6.2, 4.6], [6.2, 3.0], [5.6, 3.0]], C.oil, 4, true)].flat();
  // cutter stock drum
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.5, w: 1.5, h: 1.05, rectRadius: 0.2, fill: { color: 'F6E7C8' }, line: { color: C.oil, width: 1.5 } });
  txt('Flushing oil / cutter stock', { x: 0.6, y: 5.6, w: 1.9, h: 0.45, fontSize: 11, bold: true, color: C.muted, align: 'center' });
  const cutL = poly([[2.3, 5.2], [3.2, 5.2], [3.2, 5.97], [5.0, 5.97]], C.oil, 4, true);
  lbl('OIL-BASED SOLVENT', 5.9, 5.0, 2.0, C.greenD, 'left', 11);
  const solL = ln(6.82, 5.3, 6.82, 5.75, C.green, 4, true);
  // slop tank
  const slopL = poly([[10.6, 4.6], [10.6, 2.6], [11.2, 2.6]], C.oil, 4, true);
  const [s1, s2] = tankBox(11.2, 1.95, 1.4, 1.25, 'SLOP / RERUN TANK', '6B5236');
  // sample bottle panel
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.5, w: 3.4, h: 2.4, rectRadius: 0.12, fill: { color: C.panel }, line: { color: C.border, width: 1 } });
  lbl('SAMPLE – RETURNING OIL', 0.8, 1.62, 3.0, C.greenD, 'left', 12);
  shp(p.shapes.RECTANGLE, { x: 1.15, y: 2.05, w: 0.3, h: 0.25, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1 } });
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.95, y: 2.3, w: 0.7, h: 1.35, rectRadius: 0.12, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1.5 } });
  const bottle = ['F2D58A', 'C79A3C', '5A4025', '1C1F21'].map((c) => shp(p.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 2.6, w: 0.6, h: 1.0, rectRadius: 0.1, fill: { color: c }, line: { type: 'none' } }));
  const sTxt = ['Clean flushing oil', 'Dark – picking up deposits', 'Lighter with each pass', 'Clear – flush complete']
    .map((t) => txt(t, { x: 1.85, y: 2.55, w: 2.05, h: 1.0, fontSize: 14, bold: true, color: C.dark }));
  const caps3 = captions([
    'Heavy bottoms, coker and quench circuits hold tar-like deposits that steam alone cannot shift.',
    'Hot flushing oil (cutter stock) fills the bottoms circuit: column → pump → exchangers → column.',
    'The oil circulates continuously through the circuit.',
    'Oil-based solvent is added – it penetrates and softens the heavy deposits; the returning oil turns dark.',
    'Pass after pass the deposits dissolve – the exchanger and column clean up and the sample gets lighter.',
    'The loaded oil is sent to the slop / rerun tank – not to effluent.',
    'Circuit clean – followed by vapour-phase for final gas-freeing.']);
  const loopPts = [[7.05, 5.97], [10.6, 5.97], [10.6, 4.6], [6.2, 4.6], [6.2, 3.0], [5.0, 3.0], [5.0, 5.97], [6.6, 5.97]];
  // stage the bottle: start with "clean oil" visible (bottle[0], sTxt[0]) – others enter later
  nextCap(caps3, 1);
  st('with', 'wipe', cutL, { dir: 'left', dur: 500 });
  loopL.forEach((n, k) => st('after', 'wipe', n, { dir: k < 2 ? 'left' : 'down', dur: 300 }));
  nextCap(caps3, 2);
  packets(C.oil, loopPts, 9, 450, 4200, 0, 0.2);
  nextCap(caps3, 3);
  st('with', 'wipe', solL, { dir: 'up', dur: 400 });
  packets(C.green, [[6.82, 5.3], [6.82, 5.97], [10.6, 5.97], [10.6, 4.6], [6.2, 4.6]], 4, 400, 2400, 300);
  packets(C.oil, loopPts, 6, 500, 4200, 600, 0.2);
  st('with', 'fade', [bottle[3]], { dur: 800, delay: 1500 }); st('with', 'vanish', sTxt[0], { delay: 1500 }); st('with', 'fade', sTxt[1], { dur: 400, delay: 1500 });
  st('with', 'exit', [vf.dark, ex3[2]], { dur: 1500, delay: 2000 });
  nextCap(caps3, 4);
  packets(C.oil, loopPts, 8, 450, 4200, 0, 0.2);
  st('with', 'exit', bottle[3], { dur: 900, delay: 600 }); st('with', 'fade', bottle[2], { dur: 600, delay: 600 });
  st('with', 'vanish', sTxt[1], { delay: 900 }); st('with', 'fade', sTxt[2], { dur: 400, delay: 900 });
  st('with', 'exit', [vf.mid, ex3[1]], { dur: 1500, delay: 1200 });
  dripsAway(vf, 1000);
  st('with', 'exit', bottle[2], { dur: 900, delay: 3000 }); st('with', 'fade', bottle[1], { dur: 600, delay: 3000 });
  nextCap(caps3, 5);
  st('with', 'wipe', slopL, { dir: 'down', dur: 300 });
  packets('6B5236', [[10.6, 4.6], [10.6, 2.6], [11.5, 2.6]], 6, 300, 1300, 300, 0.2);
  st('with', 'wipe', s1, { dir: 'down', dur: 600, delay: 1000 }); st('with', 'wipe', s2, { dir: 'down', dur: 600, delay: 1900 });
  nextCap(caps3, 6);
  st('with', 'exit', [vf.light, ex3[0], bottle[1]], { dur: 1200 });
  st('with', 'vanish', sTxt[2], { delay: 600 }); st('with', 'fade', sTxt[3], { dur: 400, delay: 600 });
  shineSweep(vf);
  badge('NEXT: VAPOUR-PHASE GAS-FREEING', 0.6, 4.05, 3.4);
  notes('The returning-oil sample is how the end point is judged on site. Loaded oil goes to slop / rerun, then vapour-phase gas-frees the circuit.');

  // ===================================================== 16 CHEMICAL CIRCULATION
  slide();
  title('Chemical circulation – tanks & heat-limited vessels');
  const vc = staged(5.0, 2.05, 3.4, 3.15, 0.3);
  lbl('Tank / temperature-limited vessel', 4.7, 5.22, 4.0, C.muted, 'center');
  txt('No steam: lined, coated or low design temperature', { x: 4.7, y: 5.5, w: 4.0, h: 0.3, fontSize: 11, italic: true, color: C.muted, align: 'center' });
  // vent
  ln(7.6, 2.05, 7.6, 1.4, C.eqL, 3, true); lbl('to flare / closed vent', 7.75, 1.3, 2.2, C.muted);
  // skid
  const mixT = [shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 3.6, w: 1.4, h: 1.4, rectRadius: 0.15, fill: { color: 'EAF6EC' }, line: { color: C.green, width: 1.5 } }),
    txt('Mixing tank', { x: 0.6, y: 5.05, w: 1.8, h: 0.3, fontSize: 11, bold: true, color: C.muted, align: 'center' })];
  const pumpC = [shp(p.shapes.OVAL, { x: 3.0, y: 5.45, w: 0.45, h: 0.45, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1.5 } }),
    txt('Pump skid', { x: 2.75, y: 5.92, w: 1.0, h: 0.3, fontSize: 10, bold: true, color: C.muted, align: 'center' })];
  const hoses = [poly([[1.5, 5.0], [1.5, 5.67], [3.0, 5.67]], C.green, 4), poly([[3.45, 5.67], [4.6, 5.67], [4.6, 4.8], [5.0, 4.8]], C.green, 4, true), poly([[5.0, 2.6], [1.5, 2.6], [1.5, 3.6]], C.green, 4, true)].flat();
  // monitor panel
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 1.55, w: 3.5, h: 2.75, rectRadius: 0.12, fill: { color: C.panel }, line: { color: C.border, width: 1 } });
  lbl('RETURN-LINE MONITORING', 9.4, 1.67, 3.2, C.greenD, 'left', 12);
  const mon = [['H₂S', 'High', 'Low ✓'], ['pH', 'Off-spec', 'Neutral ✓'], ['Iron (Fe)', 'Rising', 'Stable ✓'], ['Gas test', 'Above limit', 'OK ✓']].map(([g, a, b], k) => {
    txt(g, { x: 9.4, y: 2.1 + k * 0.5, w: 1.2, h: 0.4, fontSize: 14, bold: true, color: C.dark });
    return [txt(a, { x: 10.7, y: 2.1 + k * 0.5, w: 1.9, h: 0.4, fontSize: 13, bold: true, color: C.h2s }), txt(b, { x: 10.7, y: 2.1 + k * 0.5, w: 1.9, h: 0.4, fontSize: 13, bold: true, color: C.greenD })];
  });
  poly([[8.4, 4.95], [10.95, 4.95]], C.eqL, 3, true);
  const [c1l, c2l] = tankBox(10.95, 4.5, 1.5, 1.15, 'EFFLUENT TANK');
  const caps4 = captions([
    'Tanks and temperature-limited vessels (lined, coated, low design temperature) cannot be steamed.',
    'A temporary circulation skid is set up: mixing tank and pump, hosed to the equipment.',
    'Decon solution is circulated continuously through the equipment and back to the mixing tank.',
    'H₂S is converted into stable compounds and FeS is oxidised; LEL / benzene vapours are vented to the flare.',
    'The return line is monitored until H₂S is low, pH is neutral and the iron level is stable.',
    'The solution is drained to the effluent tank and a neutralising rinse follows.',
    'Gas tests confirm the entry criteria – safe entry.']);
  const cH = tag('H₂S', 5.5, 2.45, 1.0, C.h2s), cB = tag('Benzene', 6.7, 2.45, 1.0, C.benz);
  const cG = shp(p.shapes.ROUNDED_RECTANGLE, { x: 5.43, y: 4.2, w: 1.14, h: 0.5, rectRadius: 0.25, fill: { color: C.glow, transparency: 55 }, line: { color: C.glow, width: 2 } });
  const cF = tag('FeS', 5.5, 4.27, 1.0, C.fes);
  const cS = tag('stable salt', 5.4, 2.45, 1.2, C.salt), cO = tag('iron oxide', 5.4, 4.27, 1.2, C.oxide);
  st('with', 'pulse', cG, { dur: 500, scale: 1.08, repeat: 3, delay: 400 });
  nextCap(caps4, 1);
  st('with', 'ascend', [...mixT, ...pumpC], { dur: 700 });
  hoses.forEach((h2, k) => st('after', 'wipe', h2, { dir: [ 'down', 'left', 'left', 'left', 'down', 'left', 'right', 'down'][k] || 'left', dur: 250 }));
  const circPts = [[1.5, 5.0], [1.5, 5.67], [4.6, 5.67], [4.6, 4.8], [5.6, 4.8], [5.6, 2.6], [1.5, 2.6], [1.5, 3.6]];
  nextCap(caps4, 2);
  packets(C.green, circPts, 9, 450, 4200);
  nextCap(caps4, 3);
  packets(C.green, circPts, 6, 500, 4200);
  st('with', 'pulse', cH, { dur: 300, scale: 1.2, repeat: 2 });
  st('with', 'exit', cH, { dur: 400, delay: 700 }); st('with', 'fade', cS, { dur: 400, delay: 700 });
  st('with', 'exit', [cG, cF], { dur: 500, delay: 1200 }); st('with', 'fade', cO, { dur: 500, delay: 1200 });
  st('with', 'path', cB, { path: [[0.9, -0.9], [0.9, -1.4]], dur: 1400, delay: 1500 }); st('with', 'exit', cB, { dur: 400, delay: 2900 });
  st('with', 'exit', vc.dark, { dur: 1500, delay: 1800 });
  nextCap(caps4, 4);
  packets(C.green, circPts, 6, 500, 4200);
  mon.slice(0, 3).forEach(([a, b], k) => { st('with', 'vanish', a, { delay: 600 + k * 700 }); st('with', 'zoom', b, { dur: 400, delay: 600 + k * 700 }); });
  st('with', 'exit', vc.mid, { dur: 1500, delay: 1000 });
  dripsAway(vc, 800);
  nextCap(caps4, 5);
  st('with', 'exit', [cS, cO], { dur: 500 });
  packets(C.salt, [[8.2, 4.95], [11.2, 4.95]], 5, 300, 1200, 200);
  st('with', 'wipe', c1l, { dir: 'down', dur: 600, delay: 900 }); st('with', 'wipe', c2l, { dir: 'down', dur: 600, delay: 1600 });
  st('with', 'exit', vc.light, { dur: 1200, delay: 1200 });
  shineSweep(vc);
  nextCap(caps4, 6);
  st('with', 'vanish', mon[3][0], { delay: 300 }); st('with', 'zoom', mon[3][1], { dur: 400, delay: 300 });
  badge('✓  READY FOR SAFE ENTRY', 5.0, 5.82, 3.4);
  notes('Used where steam is not allowed: tanks and lined / coated / low-design-temperature vessels.');

  // ===================================================== 17 PACKING PRE- & POST-TREATMENT
  slide();
  title('Packing pre- & post-treatment – packed columns');
  // phase chips (left)
  const pk_phases = ['1   PRE-TREATMENT', '2   VAPOUR-PHASE', '3   POST-TREATMENT'], pk_phCol = [C.green, C.blue, C.green];
  const pk_phOn = pk_phases.map((t, k) => {
    s.addText(t, { shape: p.shapes.ROUNDED_RECTANGLE, rectRadius: 0.19, x: 0.6, y: 1.55 + k * 0.48, w: 3.3, h: 0.38, fill: { color: C.panel }, line: { color: C.border, width: 1 },
      fontSize: 12, bold: true, color: C.muted, fontFace: 'Calibri', margin: 0, align: 'center', valign: 'middle', objectName: `ph${k}` });
    return tag(t, 0.6, 1.55 + k * 0.48, 3.3, pk_phCol[k]);
  });
  // column with packed beds
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 5.4, y: 1.85, w: 1.4, h: 3.95, rectRadius: 0.7, fill: { color: 'EEF2F6' }, line: { color: C.blue, width: 2 } });
  const beds = [2.3, 3.35, 4.4], pk_dirtB = [], pk_films = [], glows = [];
  beds.forEach((y) => {
    shp(p.shapes.RECTANGLE, { x: 5.5, y, w: 1.2, h: 0.75, fill: { color: 'C9B9A3' }, line: { color: C.eqL, width: 1 } });
    for (let k = 0; k < 5; k++) ln(5.5 + k * 0.24, y, 5.74 + k * 0.24, y + 0.75, '9C8A70', 0.75);
    pk_dirtB.push(shp(p.shapes.RECTANGLE, { x: 5.5, y, w: 1.2, h: 0.75, fill: { color: C.dirty[1], transparency: 15 }, line: { color: C.dirty[0], width: 1 } }));
    pk_films.push(shp(p.shapes.RECTANGLE, { x: 5.44, y: y - 0.04, w: 1.32, h: 0.83, fill: { color: C.green, transparency: 70 }, line: { color: C.green, width: 2.5 } }));
    glows.push(shp(p.shapes.RECTANGLE, { x: 5.42, y: y - 0.05, w: 1.36, h: 0.85, fill: { color: C.glow, transparency: 45 }, line: { color: C.glow, width: 2 } }));
  });
  lbl('Packed column', 5.1, 5.82, 2.0, C.muted, 'center');
  // top feed (pre-treat chemistry, later oxidiser)
  const pk_feedL = poly([[4.3, 1.55], [6.1, 1.55], [6.1, 1.85]], C.green, 4, true);
  const pk_fl1 = lbl('PRE-TREAT CHEMISTRY', 4.0, 1.22, 2.4, C.greenD, 'left', 12);
  const pk_fl2 = lbl('OXIDISER SOLUTION', 4.0, 1.22, 2.4, C.greenD, 'left', 12);
  // steam in at bottom, vapours to flare
  const pk_stL = poly([[4.2, 5.3], [5.4, 5.3]], C.steam, 4, true);
  const pk_stT = lbl('STEAM', 4.2, 4.98, 1.0, C.steam, 'left', 12);
  const pk_ventL = poly([[6.45, 1.9], [6.45, 1.6], [9.4, 1.6], [9.4, 2.3], [9.6, 2.3]], C.eqL, 3, true);
  const pk_fl = flare(9.6, 1.9);
  // bed tags
  const pk_filmTags = beds.map((y) => tag('Packing pre-wetted', 7.0, y + 0.2, 2.0, C.green));
  const pk_hcTags = [tag('Benzene', 7.0, beds[0] + 0.2, 1.3, C.benz), tag('LEL', 7.0, beds[1] + 0.2, 1.3, C.lel)];
  const oxTags = beds.map((y) => tag('FeS → iron oxide', 7.0, y + 0.2, 2.0, C.oxide));
  const warn = [shp(p.shapes.ISOSCELES_TRIANGLE, { x: 7.0, y: 5.05, w: 0.62, h: 0.55, fill: { color: 'F7C548' }, line: { color: C.glow, width: 1.5 } }),
    txt('!', { x: 7.0, y: 5.15, w: 0.62, h: 0.42, fontSize: 18, bold: true, color: C.dark, align: 'center' })];
  const warnT = txt('Pyrophoric FeS left on the packing – ignites in air', { x: 7.75, y: 5.0, w: 2.15, h: 0.7, fontSize: 13, bold: true, color: C.h2s });
  // drain to effluent
  poly([[6.1, 5.8], [6.1, 6.05], [9.6, 6.05], [9.6, 5.55], [10.0, 5.55]], C.eqL, 3, true);
  const [p1, p2] = tankBox(10.0, 4.75, 1.5, 1.2, 'EFFLUENT TANK');
  const meter5 = gasPanel(0.6, 3.15, 3.3);
  const caps5 = captions([
    'Packed beds hold hydrocarbon film and pyrophoric iron sulphide deep inside the packing geometry.',
    'PRE-TREATMENT: decon chemistry is fed from the top and coats every surface of the packing.',
    'With the packing pre-wetted, the vapour-phase that follows reaches deep into every bed.',
    'VAPOUR-PHASE: steam carries the chemistry up through the beds – LEL and benzene go to flare.',
    'Pyrophoric FeS can remain on the packing – it would self-heat once the column is opened.',
    'POST-TREATMENT: an oxidiser wash trickles through each bed – FeS becomes stable iron oxide.',
    'The spent wash drains to the effluent tank.',
    'Gas tests pass – the column is safe to open and the packing safe to unload.']);
  st('with', 'pulse', pk_dirtB, { dur: 500, scale: 1.04, repeat: 2, delay: 300 });
  // 2 pre-treatment
  nextCap(caps5, 1);
  st('with', 'fade', pk_phOn[0], { dur: 400 });
  st('with', 'fade', pk_fl1, { dur: 400 });
  pk_feedL.forEach((n, k) => st('with', 'wipe', n, { dir: k === 0 ? 'left' : 'down', dur: 400, delay: k * 400 }));
  for (let k = 0; k < 10; k++) travel(C.green, [[5.65 + (k % 5) * 0.22, 1.95], [5.65 + (k % 5) * 0.22, 5.4]], 900 + k * 220, 2200, 0.16);
  pk_films.forEach((f, k) => { st('with', 'fade', f, { dur: 600, delay: 1400 + k * 700 }); st('with', 'ascend', pk_filmTags[k], { dur: 500, delay: 1400 + k * 700 }); });
  // 3 enhanced vapour-phase
  nextCap(caps5, 2);
  st('with', 'pulse', pk_films, { dur: 450, scale: 1.05, repeat: 2 });
  // 4 vapour-phase
  nextCap(caps5, 3);
  st('with', 'exit', [pk_phOn[0], pk_fl1, ...pk_filmTags], { dur: 400 });
  st('with', 'fade', pk_phOn[1], { dur: 400 });
  st('with', 'wipe', [...pk_stL, pk_stT], { dir: 'left', dur: 500, delay: 300 });
  for (let k = 0; k < 12; k++) travel(C.steam, [[5.6 + (k % 4) * 0.33, 5.3], [5.6 + (k % 4) * 0.33, 2.05]], 700 + k * 200, 1800, 0.18);
  st('with', 'fade', pk_hcTags, { dur: 400, delay: 900 });
  pk_ventL.forEach((n, k) => st('with', 'wipe', n, { dir: ['up', 'left', 'down', 'left'][k], dur: 300, delay: 1200 + k * 300 }));
  pk_hcTags.forEach((t, k) => { st('with', 'path', t, { path: [[-0.55, -0.9 - k * 1.05], [2.3, -1.05 - k * 1.05], [2.6, -0.4 - k * 1.05]], dur: 1500, delay: 2600 + k * 500 });
    st('with', 'exit', t, { dur: 300, delay: 4000 + k * 500 }); });
  st('with', 'pulse', pk_fl, { dur: 350, scale: 1.25, repeat: 3, delay: 3800 });
  pk_dirtB.forEach((d, k) => st('with', 'exit', d, { dur: 800, delay: 2400 + k * 700 }));
  pk_films.forEach((f, k) => st('with', 'exit', f, { dur: 800, delay: 2400 + k * 700 }));
  meter5.forEach(([bad, ok], k) => { st('with', 'vanish', bad, { delay: 4600 + k * 400 }); st('with', 'zoom', ok, { dur: 400, delay: 4600 + k * 400 }); });
  // 5 pyrophoric residue
  nextCap(caps5, 4);
  st('with', 'fade', glows, { dur: 500 });
  st('after', 'pulse', glows, { dur: 450, scale: 1.04, repeat: 3 });
  st('with', 'zoom', [...warn, warnT], { dur: 500 });
  // 6 post-treatment
  nextCap(caps5, 5);
  st('with', 'exit', pk_phOn[1], { dur: 400 });
  st('with', 'fade', [pk_phOn[2], pk_fl2], { dur: 400 });
  for (let k = 0; k < 10; k++) travel(C.green, [[5.65 + (k % 5) * 0.22, 1.95], [5.65 + (k % 5) * 0.22, 5.4]], 400 + k * 220, 2200, 0.16);
  st('with', 'exit', [...warn, warnT], { dur: 500, delay: 800 });
  glows.forEach((g, k) => { st('with', 'exit', g, { dur: 700, delay: 1000 + k * 900 }); st('with', 'ascend', oxTags[k], { dur: 500, delay: 1000 + k * 900 }); });
  // 7 drain
  nextCap(caps5, 6);
  st('with', 'exit', oxTags, { dur: 500 });
  for (let k = 0; k < 5; k++) travel(C.oxide, [[6.1, 5.5], [6.1, 6.05], [9.6, 6.05], [9.6, 5.55], [10.25, 5.55]], 300 + k * 250, 1800);
  st('with', 'wipe', p1, { dir: 'down', dur: 600, delay: 1500 }); st('with', 'wipe', p2, { dir: 'down', dur: 600, delay: 2200 });
  // 8 handover
  nextCap(caps5, 7);
  badge('✓  SAFE TO OPEN & UNLOAD', 0.6, 5.3, 3.3);
  notes('Pre-treatment and post-treatment use the same circulation/spray route. Pre-treatment coats the packing geometry so the vapour-phase works deeper; post-treatment is the oxidation wash for pyrophoric FeS before opening.');

  // ===================================================== 18 TANK GAMMA-JET
  slide();
  title('Tank decontamination – gamma-jet circulation');
  // tank body with staged walls
  shp(p.shapes.OVAL, { x: 4.2, y: 5.35, w: 4.6, h: 0.5, fill: { color: 'DDE5EE' }, line: { color: C.eqL, width: 1.5 } });
  const vt = staged(4.2, 2.45, 4.6, 3.15, 0.05, C.dirty, false);
  shp(p.shapes.OVAL, { x: 4.2, y: 2.2, w: 4.6, h: 0.5, fill: { color: 'DDE5EE' }, line: { color: C.eqL, width: 1.5 } });
  const sl1 = shp(p.shapes.RECTANGLE, { x: 4.25, y: 4.75, w: 4.5, h: 0.8, fill: { color: C.dirty[0] }, line: { type: 'none' } });
  const sl2 = shp(p.shapes.RECTANGLE, { x: 4.25, y: 5.25, w: 4.5, h: 0.3, fill: { color: C.dirty[1] }, line: { type: 'none' } });
  lbl('Storage tank', 5.5, 5.88, 2.0, C.muted, 'center');
  const nozL = ln(6.5, 1.9, 6.5, 3.65, C.blue, 3);
  const jets = shp(p.shapes.STAR_24_POINT || p.shapes.STAR_16_POINT, { x: 5.55, y: 2.95, w: 1.9, h: 1.9, fill: { color: '7FB8E3', transparency: 45 }, line: { color: C.blue, width: 1, transparency: 30 } });
  const nozzle = shp(p.shapes.OVAL, { x: 6.38, y: 3.78, w: 0.24, h: 0.24, fill: { color: C.blue }, line: { type: 'none' } });
  lbl('Gamma-jet nozzle', 6.65, 3.3, 1.9, C.blue);
  // vent to flare (top left)
  poly([[5.0, 2.3], [5.0, 1.85], [2.7, 1.85]], C.eqL, 3, true);
  const fl6 = flare(2.5, 1.55);
  // separator + slop + recirculation
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 9.6, y: 5.05, w: 1.8, h: 0.6, rectRadius: 0.3, fill: { color: 'EEF2F6' }, line: { color: C.eqL, width: 1.5 } });
  lbl('Separator', 9.6, 5.68, 1.8, C.muted, 'center');
  const sucL = poly([[8.8, 5.35], [9.6, 5.35]], C.dirty[1], 4, true);
  const oilL = poly([[11.4, 5.35], [11.9, 5.35], [11.9, 3.3]], C.oil, 4, true);
  const [o1, o2] = tankBox(11.2, 2.05, 1.4, 1.25, 'RECOVERED OIL', '8C6A3E');
  const recL = poly([[10.5, 5.05], [10.5, 1.65], [6.5, 1.65], [6.5, 1.9]], C.blue, 4, true);
  lbl('Recirculated wash water + chemistry', 7.0, 1.3, 3.4, C.blue);
  const meter6 = gasPanel(0.6, 3.35, 3.0);
  const caps6 = captions([
    'Crude and product tanks: bottom sludge, and benzene, LEL and H₂S in the vapour space – entry is not allowed.',
    'A gamma-jet nozzle is lowered through the roof manway and wash water with decon chemistry is circulated to it.',
    'The rotating jets sweep every surface of the tank – shell, roof and floor.',
    'Sludge is broken up and pumped out; the separator sends recovered oil to the slop tank.',
    'H₂S is converted into stable compounds; benzene and LEL vapours go to the flare / closed vent.',
    'The floor and walls are clean – the wash water goes for effluent treatment.',
    'Gas tests confirm the entry criteria – safe entry.']);
  const gH = tag('H₂S', 4.6, 2.9, 1.0, C.h2s), gB = tag('Benzene', 7.4, 2.9, 1.0, C.benz), gL = tag('LEL', 7.4, 3.45, 1.0, C.lel);
  const gS = tag('stable salt', 4.5, 2.9, 1.2, C.salt);
  nextCap(caps6, 1);
  st('with', 'wipe', nozL, { dir: 'up', dur: 700 }); st('with', 'zoom', nozzle, { dur: 300, delay: 600 });
  recL.forEach((n, k) => st('after', 'wipe', n, { dir: ['down', 'right', 'up'][k] || 'down', dur: 300 }));
  nextCap(caps6, 2);
  st('with', 'zoom', jets, { dur: 500 });
  st('after', 'pulse', jets, { dur: 350, scale: 1.12, repeat: 8 });
  packets(C.blue, [[10.5, 5.05], [10.5, 1.65], [6.5, 1.65], [6.5, 3.8]], 6, 400, 1800, 0, 0.18);
  nextCap(caps6, 3);
  st('with', 'wipe', sucL, { dir: 'left', dur: 300 });
  st('with', 'pulse', jets, { dur: 350, scale: 1.12, repeat: 6 });
  packets(C.dirty[1], [[8.8, 5.35], [9.9, 5.35]], 6, 300, 900, 200, 0.2);
  st('with', 'exit', sl1, { dur: 2200, delay: 600 });
  oilL.forEach((n, k) => st('with', 'wipe', n, { dir: k === 0 ? 'left' : 'down', dur: 300, delay: 1200 + k * 300 }));
  packets(C.oil, [[11.4, 5.35], [11.9, 5.35], [11.9, 3.0]], 5, 300, 1100, 1800, 0.2);
  st('with', 'wipe', o1, { dir: 'down', dur: 600, delay: 2600 }); st('with', 'wipe', o2, { dir: 'down', dur: 600, delay: 3300 });
  nextCap(caps6, 4);
  st('with', 'pulse', gH, { dur: 300, scale: 1.2, repeat: 2 });
  st('with', 'exit', gH, { dur: 400, delay: 700 }); st('with', 'fade', gS, { dur: 400, delay: 700 });
  [[gB, 7.4, 2.9], [gL, 7.4, 3.45]].forEach(([n, x, y], k) => { st('with', 'path', n, { path: [[5.0 - 0.5 - x, 0], [5.0 - 0.5 - x, 1.67 - y], [2.0 - x, 1.67 - y]], dur: 1800, delay: 300 + k * 400 }); st('with', 'exit', n, { dur: 400, delay: 2100 + k * 400 }); });
  st('with', 'pulse', fl6, { dur: 300, scale: 1.4, repeat: 4, delay: 1900 });
  nextCap(caps6, 5);
  st('with', 'exit', [gS, sl2, jets], { dur: 800 });
  st('with', 'exit', vt.dark, { dur: 1000, delay: 300 }); st('with', 'exit', vt.mid, { dur: 1000, delay: 1100 }); st('with', 'exit', vt.light, { dur: 1000, delay: 1900 });
  shineSweep(vt);
  nextCap(caps6, 6);
  gasPass(meter6, 0.6, 5.0 + 0.25, 3.0);

  // ===================================================== 19 WHERE DOES EACH CONTAMINANT GO
  slide();
  title('Where does each contaminant go?');
  const dest = [[fa, 'FaFire', 'Flare / closed vent', 'Vapours swept out with the steam or vented', C.benz, 3.5],
    [fa, 'FaFlask', 'Converted & drained', 'Made stable in the equipment, then drained to the effluent tank', C.blue, 6.55],
    [fa, 'FaOilCan', 'Recovered oil', 'Heavy oil to the slop / rerun tank', C.dirty[1], 9.6]];
  const dPos = {};
  for (const [lib, ic, t, d, col, x] of dest) {
    card(x, 1.5, 2.9, 4.7, 'FFFFFF');
    await circleIcon(x + 0.25, 1.7, 0.7, lib, ic, col);
    txt(t, { x: x + 1.05, y: 1.7, w: 1.75, h: 0.7, fontSize: 15, bold: true, color: C.head });
    txt(d, { x: x + 0.25, y: 2.5, w: 2.45, h: 0.75, fontSize: 12, color: C.muted, valign: 'top' });
    dPos[t] = { x: x + 0.25, y: 3.35 };
  }
  const goes = [['Benzene', C.benz, 'Flare / closed vent', null], ['LEL', C.lel, 'Flare / closed vent', null],
    ['H₂S', C.h2s, 'Converted & drained', 'H₂S → stable salt'], ['FeS', C.fes, 'Converted & drained', 'FeS → iron oxide'],
    ['Ammonia', C.amm, 'Converted & drained', 'NH₃ → neutralised'], ['Mercaptans', C.salt, 'Converted & drained', 'Mercaptans → converted'],
    ['Sludge', C.dirty[1], 'Converted & drained', 'Sludge → emulsified'], ['Heavy oil', C.dirty[0], 'Recovered oil', 'Heavy oil → slop / rerun']];
  const used = {};
  const groups = [[0, 1], [2, 3, 4, 5, 6], [7]];
  goes.forEach(([t, col], k) => { goes[k].push(tag(t, 0.6, 1.55 + k * 0.58, 2.4, col)); });
  groups.forEach((g, gi2) => {
    g.forEach((k, j) => {
      const [t, col, d, after, n] = goes[k];
      const slot = used[d] = (used[d] || 0) + 1;
      const tx = dPos[d].x, ty = dPos[d].y + (slot - 1) * 0.5;
      st(j === 0 ? 'click' : 'with', 'path', n, { path: [[tx - 0.6, ty - (1.55 + k * 0.58)]], dur: 1100, delay: j * 200 });
      if (after) {
        const a = tag(after, tx, ty, 2.4, col);
        st('with', 'exit', n, { dur: 300, delay: 1100 + j * 200 }); st('with', 'fade', a, { dur: 300, delay: 1100 + j * 200 });
      }
    });
  });
  notes('Click 1: vapours to the flare. Click 2: H₂S, FeS, ammonia, mercaptans and sludge are converted / emulsified and drained. Click 3: heavy oil is recovered.');

  // ===================================================== 20 CHEMISTRY TOOLKIT
  slide();
  title('Decontamination chemistry toolkit');
  const tk = [['FaWind', 'Vapour-phase degreaser', 'Carried by steam; lifts hydrocarbons and releases LEL / benzene for safe venting.', 'Columns, overheads, exchangers'],
    ['FaTint', 'Water-based solvent', 'Emulsifies oil and sludge in water; non-flammable, low odour.', 'Boil-out, circulation, tanks'],
    ['FaOilCan', 'Oil-based solvent', 'Boosts flushing oil to penetrate heavy deposits, tar and coke.', 'Viscosity flush, coker, quench oil'],
    ['FaShieldAlt', 'Sulphide / pyrophoric oxidiser', 'Converts H₂S and pyrophoric FeS into stable compounds.', 'Post-treatment wash, packed beds, SWS, SRU'],
    ['FaBalanceScale', 'Neutraliser', 'Corrects pH and treats ammonia / acid residues for disposal.', 'Post-rinse, effluent tanks'],
    ['FaWater', 'Antifoam / dispersant', 'Controls foaming and keeps solids dispersed.', 'Amine systems, boil-outs']];
  for (const [i, [ic, t, d, u]] of tk.entries()) {
    const x = 0.6 + (i % 3) * 4.1, y = 1.5 + Math.floor(i / 3) * 2.55;
    const c1 = card(x, y, 3.85, 2.3);
    const ci = await circleIcon(x + 0.25, y + 0.25, 0.65, fa, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x: x + 1.05, y: y + 0.25, w: 2.7, h: 0.65, fontSize: 15, bold: true, color: C.head });
    const dd = txt(d, { x: x + 0.25, y: y + 1.0, w: 3.4, h: 0.85, fontSize: 13, valign: 'top' });
    const uu = txt([{ text: 'Use: ', options: { bold: true, color: C.greenD } }, { text: u }], { x: x + 0.25, y: y + 1.85, w: 3.4, h: 0.35, fontSize: 12, color: C.muted });
    st(i === 0 ? 'click' : 'after', 'fade', [c1, ...ci, tt, dd, uu], { dur: 400 });
  }
  txt('Chemistry is selected per unit after a compatibility check; product data and SDS are shared with the client.', { x: 0.6, y: 6.65, w: 11.5, h: 0.35, fontSize: 12, italic: true, color: C.muted });

  // ===================================================== 21 DIVIDER execution
  divider('02', 'Execution and why Delight', 'From site survey to handover');

  // 22 timeline
  slide();
  title('How we execute');
  const ex = [['FaSearch', 'Site survey', 'P&IDs, isometrics, data sheets, injection points'], ['FaPencilRuler', 'Decon plan', 'Method statement, chemistry, volumes, waste plan'],
    ['FaTruck', 'Mobilise & rig-up', 'Pumps, hoses, injection manifolds, effluent tanks'], ['FaCogs', 'Decontaminate', 'Vapour-phase, boil-out, flush or circulation with live monitoring'],
    ['FaCheckCircle', 'Validate & hand over', 'Gas tests against entry criteria, sign-off'], ['FaRecycle', 'Waste & demob', 'Neutralise, test and dispose via licensed contractor']];
  const tl = ln(1.2, 2.65, 12.1, 2.65, C.line, 3);
  st('click', 'wipe', tl, { dir: 'left', dur: 800 });
  for (const [i, [ic, t, d]] of ex.entries()) {
    const x = 0.7 + i * 2.05;
    const ci = await circleIcon(x + 0.45, 2.15, 1.0, fa, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x, y: 3.35, w: 1.9, h: 0.65, fontSize: 15, bold: true, color: C.head, align: 'center' });
    const dd = txt(d, { x, y: 4.0, w: 1.9, h: 1.4, fontSize: 13, align: 'center', valign: 'top' });
    st('after', 'ascend', [...ci, tt, dd], { dur: 500 });
  }
  // 23 responsibilities
  slide();
  title('Division of responsibilities');
  const rs = [['DELIGHT', C.blue, ['Decon engineering and method statement', 'Specialist supervisors and technicians', 'Decontamination chemistry', 'Pumps, skids, injection hoses and manifolds', 'Gas monitoring during decontamination', 'Effluent neutralisation and waste coordination']],
    ['CLIENT', C.green, ['Isolation, de-inventory and permits', 'Steam, water, nitrogen and power utilities', 'Access to control room / DCS readings', 'Operators to drain low points', 'Laboratory support for sample analysis', 'Single point of contact']]];
  for (const [i, [h, c, items]] of rs.entries()) {
    const x = 0.6 + i * 6.15;
    const c1 = card(x, 1.5, 5.9, 5.1);
    const hb = shp(p.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 1.8, w: 5.3, h: 0.6, rectRadius: 0.3, fill: { color: c }, line: { type: 'none' } });
    const ht = txt(h, { x: x + 0.3, y: 1.8, w: 5.3, h: 0.6, fontSize: 15, bold: true, color: 'FFFFFF', align: 'center', charSpacing: 2 });
    const bl = bullets(items, { x: x + 0.45, y: 2.7, w: 5.1, h: 3.7, fontSize: 15, paraSpaceAfter: 8 });
    st('click', 'ascend', [c1, hb, ht, bl], { dur: 700 });
  }
  // 24 monitoring
  slide();
  title('Monitoring and validation');
  const mo = [['FaTachometerAlt', 'Multi-gas meter', 'LEL, H₂S, O₂ and CO at vents and manways'], ['FaVial', 'Benzene / VOC meter', 'PID readings and detector tubes for benzene'],
    ['FaFlask', 'Gas sampling', 'Sample bombs to the client laboratory where required'], ['FaTint', 'Effluent checks', 'pH, COD and oil content before disposal']];
  for (const [i, [ic, t, d]] of mo.entries()) {
    const x = 0.6 + i * 3.08;
    const c1 = card(x, 1.55, 2.85, 3.4);
    const ci = await circleIcon(x + 0.9, 1.85, 1.05, fa, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x: x + 0.15, y: 3.05, w: 2.55, h: 0.5, fontSize: 16, bold: true, color: C.head, align: 'center' });
    const dd = txt(d, { x: x + 0.2, y: 3.6, w: 2.45, h: 1.2, fontSize: 14, align: 'center', valign: 'top' });
    st(i === 0 ? 'click' : 'after', 'ascend', [c1, ...ci, tt, dd], { dur: 600 });
  }
  const mb = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.35, w: 12.1, h: 1.0, rectRadius: 0.1, fill: { color: 'EAF6EC' }, line: { type: 'none' } });
  const mt = txt('Decontamination is complete when the client’s entry criteria are met and confirmed by validation readings – then the equipment is handed over for opening.', { x: 0.85, y: 5.35, w: 11.6, h: 1.0, fontSize: 15, color: C.dark });
  st('click', 'fade', [mb, mt]);
  // 25 waste
  slide();
  title('Less waste, less exposure');
  const we = [['FaTint', 'Less water than water washing', 'Vapour-phase uses steam already available on the unit.'], ['FaUserShield', 'Less confined-space entry', 'Equipment is gas-free and pyrophoric-safe before it is opened.'],
    ['FaFlask', 'Controlled effluent', 'Collected in holding tanks, neutralised and tested before disposal.'], ['FaOilCan', 'Oil recovered, not wasted', 'Flushing oil and tank sludge oil go to slop / rerun.']];
  for (const [i, [ic, t, d]] of we.entries()) {
    const y = 1.55 + i * 1.25;
    const ci = await circleIcon(0.8, y, 0.9, fa, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x: 2.0, y: y - 0.02, w: 9.5, h: 0.45, fontSize: 18, bold: true, color: C.head });
    const dd = txt(d, { x: 2.0, y: y + 0.45, w: 9.5, h: 0.45, fontSize: 15 });
    st('click', 'ascend', [...ci, tt, dd], { dur: 600 });
  }
  // 26 why delight
  slide();
  title('Why Delight');
  const wd = [['FaCertificate', 'Certified QHSE', 'ISO 9001, ISO 14001 and ISO 45001 management system'], ['FaGlobeAsia', 'Regional presence', 'UAE, Saudi Arabia, Oman, Qatar and India'],
    ['FaLayerGroup', 'Integrated services', 'Decontamination, chemical cleaning, hydro jetting, bundle extraction'], ['FaTruckMoving', 'Own fleet', 'HP / UHP units and 22 T – 65 T bundle extractors'],
    ['FaUsers', 'Experienced crews', 'Trained supervisors and technicians for shutdown work'], ['FaHandshake', 'One point of responsibility', 'From survey and planning to waste handling and handover']];
  for (const [i, [ic, t, d]] of wd.entries()) {
    const x = 0.6 + (i % 3) * 4.1, y = 1.5 + Math.floor(i / 3) * 2.6;
    const c1 = card(x, y, 3.85, 2.3);
    const ci = await circleIcon(x + 0.25, y + 0.3, 0.75, fa, ic, i % 2 ? C.green : C.blue);
    const tt = txt(t, { x: x + 1.15, y: y + 0.32, w: 2.6, h: 0.7, fontSize: 16, bold: true, color: C.head });
    const dd = txt(d, { x: x + 0.25, y: y + 1.2, w: 3.4, h: 0.95, fontSize: 14, valign: 'top' });
    st(i === 0 ? 'click' : 'after', 'fade', [c1, ...ci, tt, dd], { dur: 400 });
  }
  // 27 closing
  slide('SOFT');
  shp(p.shapes.OVAL, { x: 9.0, y: -1.4, w: 5.8, h: 5.8, fill: { color: C.tintB }, line: { type: 'none' } });
  shp(p.shapes.OVAL, { x: 10.6, y: 4.3, w: 3.8, h: 3.8, fill: { color: C.tintG }, line: { type: 'none' } });
  img(LOGO, { x: 0.75, y: 0.6, w: 1.0, h: 1.31 });
  txt('Thank you', { x: 0.75, y: 2.35, w: 8, h: 1.0, fontFace: 'Cambria', fontSize: 44, bold: true, color: C.dark });
  txt('Let’s plan the decontamination for your next shutdown.', { x: 0.75, y: 3.4, w: 8.5, h: 0.9, fontSize: 20, color: C.head, valign: 'top' });
  txt([{ text: 'Delight Equipment International L.L.C.', options: { bold: true, breakLine: true } }, { text: 'P.O. Box 2932, Abu Dhabi, United Arab Emirates', options: { breakLine: true } },
    { text: 'Tel +971 2 5515641  ·  info@delightintl.ae  ·  www.delightintl.ae', options: { breakLine: true } },
    { text: 'KSA (Jubail) +966 59348 1459  ·  India (Udupi) +91 9686433661' }], { x: 0.75, y: 4.7, w: 9, h: 1.6, fontSize: 14, paraSpaceAfter: 4, valign: 'top' });

  await p.writeFile({ fileName: OUT });
  fs.writeFileSync(OUT.replace('.pptx', '.anim.json'), JSON.stringify(spec));
  console.log('built', OUT, spec.length, 'slides');
})();
