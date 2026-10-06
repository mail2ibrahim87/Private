// 3-slide preview of the combined deck: animated O&G family tree, Morph zoom into a branch,
// and an animated vapour-phase decontamination.   Usage: node preview.js <out.pptx> <logo.png>
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const gi = require('react-icons/gi');
const fa = require('react-icons/fa');

const OUT = process.argv[2];
const LOGO = 'image/png;base64,' + fs.readFileSync(process.argv[3]).toString('base64');
const C = { soft: 'E9F2F9', head: '1F5F8B', dark: '1F3A52', blue: '0C74BC', green: '3AB54A', greenD: '2E8B3A', body: '374151', muted: '5B7085',
  line: 'C6D3E0', eq: 'CBD5E1', eqL: '94A3B8', dirty: '9C7A55', h2s: 'C0392B', benz: 'E67E22', fes: '6B7280', lel: 'D4A017', steam: '0C74BC' };

async function icon(lib, name, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(lib[name], { color: '#' + color, size: '256' }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}

const BR = [
  { k: 'UPSTREAM', sub: 'Production & gas processing', icon: [gi, 'GiOilRig'], col: C.blue,
    leaves: ['Oil production stations', 'Gas processing plants', 'Produced water & tank farms'] },
  { k: 'MIDSTREAM', sub: 'Transport & storage', icon: [gi, 'GiPipes'], col: C.green,
    leaves: ['Pipelines & pig receivers', 'Crude terminals & tank farms', 'LPG / NGL plants & spheres'] },
  { k: 'DOWNSTREAM', sub: 'Refining', icon: [gi, 'GiFactory'], col: C.blue,
    leaves: ['Crude & vacuum units', 'Hydroprocessing & conversion', 'Sour water, amine & sulphur'] },
  { k: 'PETROCHEMICALS', sub: 'Chemicals & polymers', icon: [gi, 'GiChemicalDrop'], col: C.green,
    leaves: ['Aromatics (BTX)', 'Olefins & polymers', 'Methanol & fertilisers'] },
];
const CX = [1.95, 5.1, 8.25, 11.4];

(async () => {
  const p = new pptxgen();
  p.layout = 'LAYOUT_WIDE';
  p.theme = { headFontFace: 'Cambria', bodyFontFace: 'Calibri' };
  const spec = [];
  let s, A;
  const slide = (tr) => { s = p.addSlide(); s.background = { color: 'FFFFFF' }; A = { transition: tr, steps: [] }; spec.push(A);
    s.addImage({ data: LOGO, x: 12.33, y: 0.3, w: 0.55, h: 0.72 }); };
  const st = (trigger, effect, targets, extra = {}) => A.steps.push({ trigger, effect, targets: [].concat(targets), ...extra });
  let uid = 0;
  const shp = (type, o, name) => { const n = name || `s${++uid}`; s.addShape(type, { ...o, objectName: n }); return n; };
  const txt = (t, o, name) => { const n = name || `t${++uid}`; s.addText(t, { fontFace: 'Calibri', margin: 0, isTextBox: true, color: C.body, valign: 'middle', ...o, objectName: n }); return n; };
  const img = (d, o, name) => { const n = name || `i${++uid}`; s.addImage({ data: d, ...o, objectName: n }); return n; };
  const ln = (x1, y1, x2, y2, color, width, name, arrow) => shp(p.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1, line: { color, width, endArrowType: arrow ? 'triangle' : undefined } }, name);
  const title = (t) => txt(t, { x: 0.6, y: 0.38, w: 11.3, h: 0.85, fontFace: 'Cambria', fontSize: 28, bold: true, color: C.head }, 'title');
  const icons = await Promise.all(BR.map((b) => icon(b.icon[0], b.icon[1], 'FFFFFF')));
  const check = await icon(fa, 'FaCheck', 'FFFFFF');

  // ---------- tree (shared by slides 1 and 2; "!!" names are matched by Morph) ----------
  const tree = (sc, ox, oy, opts = {}) => {
    const T = (x) => ox + x * sc, U = (y) => oy + y * sc, F = (f) => Math.max(1, Math.round(f * sc * 10) / 10);
    const names = { root: [], bar: [], br: [[], [], [], []], leaves: [[], [], [], []], marks: [[], [], [], []], spine: [[], [], [], []] };
    names.root.push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(4.57), y: U(1.4), w: 4.2 * sc, h: 0.72 * sc, rectRadius: 0.36 * sc, fill: { color: C.dark }, line: { type: 'none' } }, '!!root'));
    names.root.push(txt('OIL & GAS VALUE CHAIN', { x: T(4.57), y: U(1.4), w: 4.2 * sc, h: 0.72 * sc, fontSize: F(16), bold: true, color: 'FFFFFF', align: 'center', charSpacing: 1 }, '!!root_t'));
    names.bar.push(ln(T(6.67), U(2.12), T(6.67), U(2.4), C.line, 2.5 * sc, '!!bar_v'));
    names.bar.push(ln(T(CX[0]), U(2.4), T(CX[3]), U(2.4), C.line, 2.5 * sc, '!!bar_h'));
    BR.forEach((b, i) => {
      if (opts.skipLeaves !== i) names.spine[i].push(ln(T(CX[i]), U(2.4), T(CX[i]), U(5.95), C.line, 2.5 * sc, `!!spine${i}`));
      else names.spine[i].push(ln(T(CX[i]), U(2.4), T(CX[i]), U(2.7), C.line, 2.5 * sc, `!!spine${i}`));
      const faded = opts.highlight !== undefined && opts.highlight !== i;
      const bx = CX[i] - 1.4;
      names.br[i].push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(bx), y: U(2.7), w: 2.8 * sc, h: 0.95 * sc, rectRadius: 0.12 * sc,
        fill: { color: b.col, transparency: faded ? 65 : 0 }, line: { type: 'none' } }, `!!br${i}`));
      names.br[i].push(img(icons[i], { x: T(bx + 0.15), y: U(2.85), w: 0.62 * sc, h: 0.62 * sc, transparency: faded ? 60 : 0 }, `!!bri${i}`));
      names.br[i].push(txt([{ text: b.k, options: { bold: true, fontSize: F(13.5), breakLine: true } }, { text: b.sub, options: { fontSize: F(11) } }],
        { x: T(bx + 0.85), y: U(2.7), w: 1.92 * sc, h: 0.95 * sc, color: 'FFFFFF', transparency: faded ? 50 : 0 }, `!!brt${i}`));
      if (opts.skipLeaves === i) return;
      b.leaves.forEach((l, k) => {
        const y = 3.95 + k * 0.72;
        names.leaves[i].push(shp(p.shapes.ROUNDED_RECTANGLE, { x: T(bx), y: U(y), w: 2.8 * sc, h: 0.58 * sc, rectRadius: 0.29 * sc,
          fill: { color: 'FFFFFF', transparency: faded ? 40 : 0 }, line: { color: b.col, width: 1.25 * sc, transparency: faded ? 60 : 0 } }, `!!lf${i}_${k}`));
        names.leaves[i].push(txt(l, { x: T(bx + 0.2), y: U(y), w: 2.15 * sc, h: 0.58 * sc, fontSize: F(12.5), bold: true, color: faded ? '9AA9B8' : C.dark }, `!!lft${i}_${k}`));
        names.marks[i].push(shp(p.shapes.OVAL, { x: T(bx + 2.38), y: U(y + 0.11), w: 0.36 * sc, h: 0.36 * sc, fill: { color: C.green, transparency: faded ? 60 : 0 }, line: { color: 'FFFFFF', width: 1 } }, `!!mk${i}_${k}`));
        names.marks[i].push(img(check, { x: T(bx + 2.47), y: U(y + 0.2), w: 0.18 * sc, h: 0.18 * sc }, `!!mki${i}_${k}`));
      });
    });
    return names;
  };

  // ---------- slide 1: growing tree ----------
  slide('fade');
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
  s.addNotes('Build the value chain branch by branch. The green ticks mark where Delight decontaminates.');

  // ---------- slide 2: Morph zoom into Downstream ----------
  slide('morph');
  title('Downstream – refining');
  // minimap of the tree (top right); Downstream leaves fly out into the rows on the left
  tree(0.3, 8.95, 0.85, { highlight: 2, skipLeaves: 2 });
  const rows = [
    { units: 'CDU, VDU, desalter, pre-heat trains', cont: [['H₂S', C.h2s], ['Benzene', C.benz], ['LEL', C.lel], ['Heavy oil', C.dirty]] },
    { units: 'Hydrotreaters, hydrocracker, FCC, coker', cont: [['H₂S', C.h2s], ['Pyrophoric FeS', C.fes], ['Coke', C.dark], ['LEL', C.lel]] },
    { units: 'Sour water strippers, amine regenerators, SRU / TGT', cont: [['H₂S', C.h2s], ['Ammonia', C.blue], ['Pyrophoric FeS', C.fes]] },
  ];
  rows.forEach((r, k) => {
    const y = 1.55 + k * 1.55;
    shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 3.6, h: 0.8, rectRadius: 0.4, fill: { color: 'FFFFFF' }, line: { color: C.blue, width: 2 } }, `!!lf2_${k}`);
    txt(BR[2].leaves[k], { x: 0.85, y, w: 2.75, h: 0.8, fontSize: 16, bold: true, color: C.dark }, `!!lft2_${k}`);
    shp(p.shapes.OVAL, { x: 3.65, y: y + 0.17, w: 0.46, h: 0.46, fill: { color: C.green }, line: { color: 'FFFFFF', width: 1 } }, `!!mk2_${k}`);
    s.addImage({ data: check, x: 3.77, y: y + 0.29, w: 0.22, h: 0.22, objectName: `!!mki2_${k}` });
    const u = txt([{ text: 'Where: ', options: { bold: true, color: C.head } }, { text: r.units }], { x: 4.5, y: y - 0.3, w: 4.3, h: 0.7, fontSize: 14, valign: 'bottom' });
    const chips = r.cont.flatMap(([c, col], j) => {
      const w = 0.35 + c.length * 0.085;
      const x0 = 4.5 + r.cont.slice(0, j).reduce((a, [cc]) => a + 0.35 + cc.length * 0.085 + 0.12, 0);
      return [shp(p.shapes.ROUNDED_RECTANGLE, { x: x0, y: y + 0.5, w, h: 0.36, rectRadius: 0.18, fill: { color: col }, line: { type: 'none' } }),
        txt(c, { x: x0, y: y + 0.5, w, h: 0.36, fontSize: 12, bold: true, color: 'FFFFFF', align: 'center' })];
    });
    st('click', 'ascend', [u, ...chips], { dur: 600 });
  });
  const lnk = txt([{ text: 'Next: ', options: { bold: true, color: C.greenD } }, { text: 'how vapour-phase decontamination clears these units.' }], { x: 0.6, y: 6.35, w: 8.2, h: 0.45, fontSize: 14, color: C.muted });
  st('after', 'fade', lnk);
  s.addNotes('Morph zooms into the Downstream branch: the tree shrinks to the corner and the refining leaves fly into the rows. Then the units and contaminants appear.');

  // ---------- slide 3: animated vapour-phase ----------
  slide('fade');
  title('Vapour-phase decontamination – how it works');
  // flare / vent
  ln(6.65, 1.9, 6.65, 1.55, C.eqL, 3); ln(6.65, 1.55, 10.4, 1.55, C.eqL, 3, null, true);
  txt('To closed vent / flare', { x: 10.5, y: 1.4, w: 2.2, h: 0.3, fontSize: 12, bold: true, color: C.muted });
  // column (clean underneath, dirty overlay on top)
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 6.0, y: 1.9, w: 1.3, h: 4.1, rectRadius: 0.65, fill: { color: 'DCEBF7' }, line: { color: C.eqL, width: 2 } });
  for (let t = 1; t <= 8; t++) ln(6.05, 1.9 + 4.1 * t / 9, 7.25, 1.9 + 4.1 * t / 9, C.eqL, 1);
  const dirty = shp(p.shapes.ROUNDED_RECTANGLE, { x: 6.0, y: 1.9, w: 1.3, h: 4.1, rectRadius: 0.65, fill: { color: C.dirty, transparency: 30 }, line: { color: C.eqL, width: 2 } });
  txt('Column', { x: 6.0, y: 6.05, w: 1.3, h: 0.3, fontSize: 12, bold: true, color: C.muted, align: 'center' });
  const clouds = [['H₂S', C.h2s, 2.25], ['Benzene', C.benz, 3.05], ['LEL', C.lel, 3.85], ['FeS', C.fes, 4.65]].map(([t, col, y]) => {
    const n = `cl_${t}`;
    s.addText(t, { shape: p.shapes.ROUNDED_RECTANGLE, rectRadius: 0.18, x: 6.15, y, w: 1.0, h: 0.38, fill: { color: col }, color: 'FFFFFF', fontSize: 12, bold: true,
      align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0, objectName: n });
    return [n, y];
  });
  // gas test panel
  shp(p.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 2.3, w: 3.5, h: 2.75, rectRadius: 0.12, fill: { color: 'F3F7FA' }, line: { color: 'E1E8EF', width: 1 } });
  txt('GAS TEST AT MANWAY', { x: 9.45, y: 2.45, w: 3.0, h: 0.4, fontSize: 13, bold: true, color: C.greenD, charSpacing: 2 });
  const meter = ['LEL', 'H₂S', 'Benzene', 'O₂'].map((g, k) => {
    txt(g, { x: 9.45, y: 3.0 + k * 0.48, w: 1.4, h: 0.4, fontSize: 15, bold: true, color: C.dark });
    const bad = txt('✗  above limit', { x: 10.75, y: 3.0 + k * 0.48, w: 1.8, h: 0.4, fontSize: 14, bold: true, color: C.h2s }, `bad${k}`);
    const ok = txt('✓  OK', { x: 10.75, y: 3.0 + k * 0.48, w: 1.8, h: 0.4, fontSize: 14, bold: true, color: C.greenD }, `ok${k}`);
    return [bad, ok];
  });
  if (true) { /* O₂ row is "OK" from the start for realism */ }
  // steam line + chemistry line
  txt('STEAM', { x: 1.6, y: 5.45, w: 0.8, h: 0.3, fontSize: 12, bold: true, color: C.steam });
  const steamL = ln(2.4, 5.6, 6.0, 5.6, C.steam, 5, 'steamL', true);
  txt('DECON CHEMISTRY', { x: 2.6, y: 3.9, w: 2.0, h: 0.3, fontSize: 12, bold: true, color: C.greenD }, 'chemLbl');
  const chemL = ln(3.6, 4.25, 3.6, 5.6, C.green, 5, 'chemL', true);
  const steamP = Array.from({ length: 7 }, (_, k) => shp(p.shapes.OVAL, { x: 2.3, y: 5.5, w: 0.2, h: 0.2, fill: { color: 'FFFFFF' }, line: { color: C.steam, width: 2 } }, `sp${k}`));
  const chemP = Array.from({ length: 6 }, (_, k) => shp(p.shapes.OVAL, { x: 3.5, y: 4.15, w: 0.2, h: 0.2, fill: { color: C.green }, line: { type: 'none' } }, `cp${k}`));
  // captions
  const capBox = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.35, w: 12.1, h: 0.58, rectRadius: 0.29, fill: { color: 'EAF3FA' }, line: { type: 'none' } });
  const caps = ['1   The unit is shut down, drained and isolated – but the column still holds toxic and flammable residues.',
    '2   Steam is introduced through the existing steam-out lines and heats the whole column.',
    '3   Decon chemistry is injected into the steam and travels to every surface the steam reaches.',
    '4   H₂S, benzene and LEL gases are released and swept to the closed vent / flare; FeS is oxidised.',
    '5   Deposits soften and drain – the column is clean.',
    '6   Gas tests confirm the entry criteria – the column is handed over for safe entry.']
    .map((c, k) => txt(c, { x: 0.85, y: 6.35, w: 11.6, h: 0.58, fontSize: 15, color: C.dark }, `cap${k}`));
  // hide steam/chem lines, packets, OK values, later captions initially: done by entrance animations below
  // timeline
  st('click', 'fade', [capBox, caps[0]], { dur: 400 });
  st('click', 'vanish', caps[0]); st('with', 'fade', caps[1], { dur: 300 });
  st('with', 'wipe', steamL, { dir: 'left', dur: 700 });
  steamP.forEach((n, k) => {
    const d = 700 + k * 350;
    st('with', 'appear', n, { delay: d });
    st('with', 'path', n, { path: [[3.85, 0], [3.85, -3.4]], dur: 2200, delay: d });
    st('with', 'vanish', n, { delay: d + 2200 });
  });
  st('click', 'vanish', caps[1]); st('with', 'fade', caps[2], { dur: 300 });
  st('with', 'fade', 'chemLbl', { dur: 300 });
  st('with', 'wipe', chemL, { dir: 'down', dur: 600 });
  chemP.forEach((n, k) => {
    const d = 600 + k * 380;
    st('with', 'appear', n, { delay: d });
    st('with', 'path', n, { path: [[0, 1.35], [2.65, 1.35], [2.65, -2.3]], dur: 2400, delay: d });
    st('with', 'vanish', n, { delay: d + 2400 });
  });
  st('click', 'vanish', caps[2]); st('with', 'fade', caps[3], { dur: 300 });
  clouds.forEach(([n, y], k) => {
    const d = k * 450;
    st('with', 'path', n, { path: [[0, 1.36 - y], [3.75, 1.36 - y]], dur: 1800, delay: d });
    st('with', 'exit', n, { dur: 500, delay: d + 1800 });
  });
  st('click', 'vanish', caps[3]); st('with', 'fade', caps[4], { dur: 300 });
  st('with', 'exit', dirty, { dur: 2000 });
  st('click', 'vanish', caps[4]); st('with', 'fade', caps[5], { dur: 300 });
  meter.slice(0, 3).forEach(([bad, ok], k) => { st('with', 'vanish', bad, { delay: 300 + k * 600 }); st('with', 'zoom', ok, { dur: 400, delay: 300 + k * 600 }); });
  const badge = shp(p.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 5.25, w: 3.5, h: 0.75, rectRadius: 0.37, fill: { color: C.green }, line: { type: 'none' } });
  const badgeT = txt('✓  READY FOR SAFE ENTRY', { x: 9.2, y: 5.25, w: 3.5, h: 0.75, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center' });
  st('after', 'zoom', [badge, badgeT], { dur: 500 });
  st('after', 'pulse', [badge, badgeT], { dur: 400, scale: 1.08, repeat: 2 });
  s.addNotes('Click through the six steps: steam, chemistry, contaminants swept out, column cleans, gas tests pass.');

  await p.writeFile({ fileName: OUT });
  fs.writeFileSync(OUT.replace('.pptx', '.anim.json'), JSON.stringify(spec));
  console.log('built', OUT);
})();
