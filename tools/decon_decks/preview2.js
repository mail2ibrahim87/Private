// Preview 2: vapour-phase with correct contaminant destinations + boil-out.
// Usage: node preview2.js <out.pptx> <logo.png>
const pptxgen = require('pptxgenjs');
const fs = require('fs');

const OUT = process.argv[2];
const LOGO = 'image/png;base64,' + fs.readFileSync(process.argv[3]).toString('base64');
const C = { head: '1F5F8B', dark: '1F3A52', blue: '0C74BC', green: '3AB54A', greenD: '2E8B3A', body: '374151', muted: '5B7085',
  eqL: '94A3B8', clean: 'DCEBF7', dirty: '9C7A55', h2s: 'C0392B', benz: 'E67E22', fes: '6B7280', lel: 'D4A017', steam: '0C74BC',
  salt: '8E9AAF', oxide: 'A0522D', water: 'BFDDF2', effl: 'B8A68F', glow: 'FF8C1A' };

const p = new pptxgen();
p.layout = 'LAYOUT_WIDE';
p.theme = { headFontFace: 'Cambria', bodyFontFace: 'Calibri' };
const spec = [];
let s, A, uid = 0;
const slide = (tr) => { s = p.addSlide(); s.background = { color: 'FFFFFF' }; A = { transition: tr, steps: [] }; spec.push(A);
  s.addImage({ data: LOGO, x: 12.33, y: 0.3, w: 0.55, h: 0.72 }); };
const st = (trigger, effect, targets, extra = {}) => A.steps.push({ trigger, effect, targets: [].concat(targets), ...extra });
const shp = (type, o, name) => { const n = name || `s${++uid}`; s.addShape(type, { ...o, objectName: n }); return n; };
const txt = (t, o, name) => { const n = name || `t${++uid}`; s.addText(t, { fontFace: 'Calibri', margin: 0, isTextBox: true, color: C.body, valign: 'middle', ...o, objectName: n }); return n; };
const ln = (x1, y1, x2, y2, color, width, arrow, name) => shp(p.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
  flipH: x2 < x1, flipV: y2 < y1, line: { color, width, endArrowType: arrow ? 'triangle' : undefined } }, name);
const tag = (t, x, y, w, fill, color = 'FFFFFF', name) => { const n = name || `tag${++uid}`;
  s.addText(t, { shape: p.shapes.ROUNDED_RECTANGLE, rectRadius: 0.18, x, y, w, h: 0.36, fill: { color: fill }, color, fontSize: 12, bold: true,
    align: 'center', valign: 'middle', fontFace: 'Calibri', margin: 0, objectName: n }); return n; };
const dot = (x, y, col, d = 0.22) => shp(p.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { color: 'FFFFFF', width: 0.75 } });
const title = (t) => txt(t, { x: 0.6, y: 0.38, w: 10.5, h: 0.85, fontFace: 'Cambria', fontSize: 28, bold: true, color: C.head });
const lbl = (t, x, y, w, color, align = 'left', size = 11) => txt(t, { x, y, w, h: 0.3, fontSize: size, bold: true, color, align });

const flare = (x, y) => {   // stack top-left at (x, y)
  shp(p.shapes.RECTANGLE, { x, y, w: 0.2, h: 1.3, fill: { color: 'B4C2D3' }, line: { color: C.eqL, width: 1 } });
  const f1 = shp(p.shapes.OVAL, { x: x - 0.14, y: y - 0.55, w: 0.48, h: 0.6, fill: { color: 'F7C548' }, line: { type: 'none' } });
  const f2 = shp(p.shapes.OVAL, { x: x - 0.04, y: y - 0.32, w: 0.28, h: 0.36, fill: { color: C.glow }, line: { type: 'none' } });
  txt('FLARE /\nCLOSED VENT', { x: x - 0.5, y: y + 1.35, w: 1.2, h: 0.45, fontSize: 11, bold: true, color: C.muted, align: 'center' });
  return [f1, f2];
};
const effTank = (x, y, w = 1.6, h = 1.25) => {
  shp(p.shapes.RECTANGLE, { x, y, w, h, fill: { color: 'EEF2F6' }, line: { color: C.eqL, width: 1.5 } });
  const l1 = shp(p.shapes.RECTANGLE, { x: x + 0.05, y: y + h - 0.35, w: w - 0.1, h: 0.3, fill: { color: C.effl }, line: { type: 'none' } });
  const l2 = shp(p.shapes.RECTANGLE, { x: x + 0.05, y: y + h - 0.75, w: w - 0.1, h: 0.4, fill: { color: C.effl }, line: { type: 'none' } });
  lbl('EFFLUENT TANK', x - 0.3, y + h + 0.03, w + 0.6, C.muted, 'center');
  return [l1, l2];
};
const gasPanel = (x, y, w) => {
  shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 2.0, rectRadius: 0.12, fill: { color: 'F3F7FA' }, line: { color: 'E1E8EF', width: 1 } });
  lbl('GAS TEST AT MANWAY', x + 0.2, y + 0.12, w - 0.3, C.greenD, 'left', 12);
  return ['LEL', 'H₂S', 'Benzene'].map((g, k) => {
    txt(g, { x: x + 0.2, y: y + 0.55 + k * 0.45, w: 1.1, h: 0.38, fontSize: 14, bold: true, color: C.dark });
    const bad = txt('✗  above limit', { x: x + 1.3, y: y + 0.55 + k * 0.45, w: w - 1.4, h: 0.38, fontSize: 13, bold: true, color: C.h2s });
    const ok = txt('✓  OK', { x: x + 1.3, y: y + 0.55 + k * 0.45, w: w - 1.4, h: 0.38, fontSize: 13, bold: true, color: C.greenD });
    return [bad, ok];
  });
};
const captions = (list) => {
  const box = shp(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.4, w: 12.1, h: 0.78, rectRadius: 0.2, fill: { color: 'EAF3FA' }, line: { type: 'none' } });
  const caps = list.map((c, k) => txt(`${k + 1}   ${c}`, { x: 0.85, y: 6.4, w: 11.6, h: 0.78, fontSize: 14, color: C.dark }));
  return { box, caps };
};
const nextCap = (caps, k) => { st('click', 'vanish', caps[k - 1]); st('with', 'fade', caps[k], { dur: 300 }); };
const finish = (meter, x, y, w) => {
  meter.forEach(([bad, ok], k) => { st('with', 'vanish', bad, { delay: 300 + k * 600 }); st('with', 'zoom', ok, { dur: 400, delay: 300 + k * 600 }); });
  const b = shp(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.7, rectRadius: 0.35, fill: { color: C.green }, line: { type: 'none' } });
  const bt = txt('✓  READY FOR SAFE ENTRY', { x, y, w, h: 0.7, fontSize: 15, bold: true, color: 'FFFFFF', align: 'center' });
  st('after', 'zoom', [b, bt], { dur: 500 });
  st('after', 'pulse', [b, bt], { dur: 400, scale: 1.08, repeat: 2 });
};
// droplet moving along a list of absolute centre points
const travel = (col, pts, trig, delay, dur, d = 0.22) => {
  const [x0, y0] = pts[0];
  const n = dot(x0 - d / 2, y0 - d / 2, col, d);
  st(trig, 'appear', n, { delay });
  st('with', 'path', n, { path: pts.slice(1).map(([x, y]) => [x - x0, y - y0]), dur, delay });
  st('with', 'vanish', n, { delay: delay + dur });
  return n;
};

// ===================== SLIDE 1: vapour-phase =====================
slide('fade');
title('Vapour-phase – where each contaminant goes');
// column
const CXL = 5.2, CW = 1.3, CY = 1.9, CH = 3.8, CMX = CXL + CW / 2;
shp(p.shapes.ROUNDED_RECTANGLE, { x: CXL, y: CY, w: CW, h: CH, rectRadius: 0.65, fill: { color: C.clean }, line: { color: C.eqL, width: 2 } });
for (let t = 1; t <= 7; t++) ln(CXL + 0.05, CY + CH * t / 8, CXL + CW - 0.05, CY + CH * t / 8, C.eqL, 1);
const dirty = shp(p.shapes.ROUNDED_RECTANGLE, { x: CXL, y: CY, w: CW, h: CH, rectRadius: 0.65, fill: { color: C.dirty, transparency: 35 }, line: { color: C.eqL, width: 2 } });
lbl('Column', CXL, CY + CH + 0.02, CW, C.muted, 'center');
// overhead to flare
ln(CMX, CY, CMX, 1.55, C.eqL, 3); ln(CMX, 1.55, 10.95, 1.55, C.eqL, 3, true);
const flame = flare(10.95, 1.25);
// drain to effluent tank
ln(CMX, CY + CH, CMX, 6.05, C.eqL, 3); ln(CMX, 6.05, 9.2, 6.05, C.eqL, 3); ln(9.2, 6.05, 9.2, 5.5, C.eqL, 3); ln(9.2, 5.5, 9.6, 5.5, C.eqL, 3, true);
lbl('Condensate drain', 6.7, 5.75, 2.0, C.muted);
const [lvl1, lvl2] = effTank(9.6, 4.75);
// steam + chemistry
lbl('STEAM', 0.6, 5.0, 0.9, C.steam, 'left', 12);
const steamL = ln(1.3, 5.3, CXL, 5.3, C.steam, 5, true);
lbl('DECON CHEMISTRY', 2.2, 4.5, 1.9, C.greenD, 'left', 12);
const chemL = ln(3.0, 4.82, 3.0, 5.3, C.green, 5, true);
// gas panel + captions
const meter = gasPanel(0.6, 1.5, 3.3);
const { box, caps } = captions([
  'The unit is shut down, drained and isolated – the column still holds H₂S, benzene, LEL gases, pyrophoric FeS and heavy oil.',
  'Steam is introduced through the existing steam-out lines and heats the whole column.',
  'Decon chemistry is injected into the steam and reaches every surface the steam reaches.',
  'H₂S is converted into stable, non-volatile compounds that stay in the condensate and drain to the effluent tank.',
  'Pyrophoric FeS is oxidised in place to stable iron oxide – it can no longer self-heat – and is washed out.',
  'Benzene and LEL hydrocarbons are vaporised and swept to the flare / closed vent; dissolved benzene drains with the condensate.',
  'Softened deposits and heavy oil drain with the condensate – the column is clean.',
  'Gas tests confirm the entry criteria – the column is handed over for safe entry.']);
// contaminants inside the column
const TX = CXL + 0.15, TW = 1.0;
const tH2S = tag('H₂S', TX, 2.25, TW, C.h2s);
const tBenz = tag('Benzene', TX, 2.85, TW, C.benz);
const tLEL = tag('LEL', TX, 3.45, TW, C.lel);
const glow = shp(p.shapes.ROUNDED_RECTANGLE, { x: TX - 0.07, y: 4.0, w: TW + 0.14, h: 0.5, rectRadius: 0.25, fill: { color: C.glow, transparency: 55 }, line: { color: C.glow, width: 2 } });
const tFeS = tag('FeS', TX, 4.07, TW, C.fes);
const tSalt = tag('stable salt', TX - 0.1, 2.25, TW + 0.2, C.salt);
const tOx = tag('iron oxide', TX - 0.1, 4.07, TW + 0.2, C.oxide);

// --- timeline
st('click', 'fade', [box, caps[0]], { dur: 400 });
st('with', 'pulse', glow, { dur: 500, scale: 1.08, repeat: 3, delay: 400 });
// 2 steam
nextCap(caps, 1);
st('with', 'wipe', steamL, { dir: 'left', dur: 700 });
for (let k = 0; k < 7; k++) {
  const n = shp(p.shapes.OVAL, { x: 1.2, y: 5.2, w: 0.2, h: 0.2, fill: { color: 'FFFFFF' }, line: { color: C.steam, width: 2 } });
  const d = 700 + k * 330;
  st('with', 'appear', n, { delay: d }); st('with', 'path', n, { path: [[4.35, 0], [4.35, -3.0]], dur: 2200, delay: d }); st('with', 'vanish', n, { delay: d + 2200 });
}
// 3 chemistry
nextCap(caps, 2);
st('with', 'wipe', chemL, { dir: 'down', dur: 600 });
for (let k = 0; k < 6; k++) {
  const n = shp(p.shapes.OVAL, { x: 2.9, y: 4.72, w: 0.2, h: 0.2, fill: { color: C.green }, line: { type: 'none' } });
  const d = 600 + k * 360;
  st('with', 'appear', n, { delay: d }); st('with', 'path', n, { path: [[0, 0.48], [2.75, 0.48], [2.75, -3.2]], dur: 2400, delay: d }); st('with', 'vanish', n, { delay: d + 2400 });
}
// 4 H2S -> stable salt -> drain
nextCap(caps, 3);
st('with', 'pulse', tH2S, { dur: 300, scale: 1.2, repeat: 2 });
st('after', 'exit', tH2S, { dur: 500 }); st('with', 'fade', tSalt, { dur: 500 });
st('after', 'exit', tSalt, { dur: 400, delay: 600 });
const drainPts = (y) => [[CMX, y], [CMX, 6.05], [9.2, 6.05], [9.2, 5.5], [9.85, 5.5]];
for (let k = 0; k < 4; k++) travel(C.salt, drainPts(2.43), k === 0 ? 'with' : 'with', 600 + k * 250, 2000);
st('after', 'wipe', lvl1, { dir: 'down', dur: 600 });
// 5 FeS -> iron oxide -> washed out
nextCap(caps, 4);
st('with', 'exit', glow, { dur: 600 }); st('with', 'exit', tFeS, { dur: 600 }); st('with', 'fade', tOx, { dur: 600 });
st('after', 'exit', tOx, { dur: 400, delay: 600 });
for (let k = 0; k < 4; k++) travel(C.oxide, drainPts(4.25), 'with', 600 + k * 250, 1800);
// 6 benzene + LEL -> flare (part of benzene to drain)
nextCap(caps, 5);
[[tBenz, 2.85], [tLEL, 3.45]].forEach(([n, y], k) => {
  st('with', 'path', n, { path: [[0, 1.37 - y], [5.4, 1.37 - y]], dur: 1800, delay: k * 400 });
  st('with', 'exit', n, { dur: 400, delay: 1800 + k * 400 });
});
st('with', 'pulse', flame, { dur: 300, scale: 1.4, repeat: 4, delay: 1700 });
for (let k = 0; k < 2; k++) travel(C.benz, drainPts(3.03), 'with', 900 + k * 300, 1800, 0.18);
// 7 deposits drain, column clean
nextCap(caps, 6);
st('with', 'exit', dirty, { dur: 2200 });
for (let k = 0; k < 5; k++) travel(C.dirty, drainPts(5.2), 'with', k * 300, 1600);
st('after', 'wipe', lvl2, { dir: 'down', dur: 700 });
// 8 gas test
nextCap(caps, 7);
finish(meter, 0.6, 3.7, 3.3);
s.addNotes('Eight clicks. H₂S and FeS are converted in place and leave with the condensate; only the vapours go to the flare.');

// ===================== SLIDE 2: boil-out =====================
slide('fade');
title('Boil-out – vessels with sludge');
const VX = 4.3, VY = 2.2, VW = 6.0, VH = 2.4;
shp(p.shapes.ROUNDED_RECTANGLE, { x: VX, y: VY, w: VW, h: VH, rectRadius: 1.2, fill: { color: 'EEF2F6' }, line: { color: C.eqL, width: 2 } });
const sludge = shp(p.shapes.RECTANGLE, { x: 4.9, y: 4.0, w: 4.8, h: 0.42, fill: { color: C.dirty }, line: { type: 'none' } });
const water = shp(p.shapes.ROUNDED_RECTANGLE, { rectRadius: 0.3, x: 4.65, y: 2.95, w: 5.3, h: 1.05, fill: { color: C.water, transparency: 15 }, line: { type: 'none' } });
lbl('Vessel / drum', VX, VY + VH + 0.02, 1.6, C.muted);
// vent to flare (top right)
ln(9.4, VY, 9.4, 1.55, C.eqL, 3); ln(9.4, 1.55, 11.6, 1.55, C.eqL, 3, true);
const flame2 = flare(11.6, 1.25);
// steam header from the right, three risers
lbl('STEAM', 11.0, 4.95, 0.9, C.steam, 'left', 12);
const header = ln(12.4, 5.3, 6.2, 5.3, C.steam, 5);
const risers = [6.2, 7.3, 8.4].map((x) => ln(x, 5.3, x, VY + VH, C.steam, 4, true));
// drain (left) to effluent tank (bottom left)
ln(5.6, VY + VH, 5.6, 5.55, C.eqL, 3); ln(5.6, 5.55, 2.3, 5.55, C.eqL, 3, true);
const [bl1, bl2] = effTank(0.7, 4.75, 1.6, 1.15);
// thermometer
shp(p.shapes.ROUNDED_RECTANGLE, { x: 10.62, y: 2.4, w: 0.22, h: 1.9, rectRadius: 0.11, fill: { color: 'FFFFFF' }, line: { color: C.eqL, width: 1.5 } });
const therm = shp(p.shapes.ROUNDED_RECTANGLE, { x: 10.66, y: 2.6, w: 0.14, h: 1.68, rectRadius: 0.07, fill: { color: C.h2s }, line: { type: 'none' } });
shp(p.shapes.OVAL, { x: 10.55, y: 4.15, w: 0.36, h: 0.36, fill: { color: C.h2s }, line: { type: 'none' } });
const hot = lbl('Boiling', 10.35, 2.05, 0.9, C.h2s, 'center');
// chemistry from the top
lbl('DECON CHEMISTRY', 6.6, 1.15, 1.9, C.greenD, 'left', 12);
const chemL2 = ln(7.3, 1.5, 7.3, VY, C.green, 5, true);
// gas panel + captions
const meter2 = gasPanel(0.6, 1.5, 3.3);
const cap2 = captions([
  'The vessel is drained and isolated – sludge on the bottom, H₂S, benzene and LEL in the vapour space, pyrophoric FeS in the deposits.',
  'Water is filled above the sludge level.',
  'Steam is injected at several points along the bottom and brings the water to the boil.',
  'Decon chemistry is added; the rolling boil lifts and emulsifies the sludge.',
  'H₂S dissolves and is converted into stable compounds; FeS is oxidised to iron oxide – both stay in the water.',
  'Benzene and LEL vapours leave through the vent to the flare / closed vent.',
  'The vessel is drained to the effluent tank and rinsed – the vessel is clean.',
  'Gas tests confirm the entry criteria – the vessel is handed over for safe entry.']);
const vH2S = tag('H₂S', 5.3, 2.4, 1.0, C.h2s), vBenz = tag('Benzene', 6.6, 2.4, 1.0, C.benz), vLEL = tag('LEL', 7.9, 2.4, 1.0, C.lel);
const vGlow = shp(p.shapes.ROUNDED_RECTANGLE, { x: 8.43, y: 3.94, w: 1.14, h: 0.5, rectRadius: 0.25, fill: { color: C.glow, transparency: 55 }, line: { color: C.glow, width: 2 } });
const vFeS = tag('FeS', 8.5, 4.01, 1.0, C.fes);
const vSalt = tag('stable salt', 5.2, 3.3, 1.2, C.salt), vOx = tag('iron oxide', 8.4, 3.3, 1.2, C.oxide);

st('click', 'fade', [cap2.box, cap2.caps[0]], { dur: 400 });
st('with', 'pulse', vGlow, { dur: 500, scale: 1.08, repeat: 3, delay: 400 });
// 2 water fill
nextCap(cap2.caps, 1);
st('with', 'wipe', water, { dir: 'down', dur: 1800 });
// 3 steam + boiling
nextCap(cap2.caps, 2);
st('with', 'wipe', header, { dir: 'right', dur: 700 });
st('after', 'wipe', risers, { dir: 'down', dur: 400 });
st('after', 'wipe', therm, { dir: 'down', dur: 2000 });
st('with', 'fade', hot, { dur: 300, delay: 1800 });
const bubble = (x, delay) => {
  const n = shp(p.shapes.OVAL, { x, y: 3.85, w: 0.16, h: 0.16, fill: { color: 'FFFFFF' }, line: { color: '7FB8E3', width: 1 } });
  st('with', 'appear', n, { delay }); st('with', 'path', n, { path: [[0.05, -0.45], [-0.03, -0.8]], dur: 1100, delay }); st('with', 'vanish', n, { delay: delay + 1100 });
};
for (let k = 0; k < 18; k++) bubble(4.9 + ((k * 0.53) % 4.9), 300 + k * 160);
// 4 chemistry + sludge lifts
nextCap(cap2.caps, 3);
st('with', 'wipe', chemL2, { dir: 'up', dur: 500 });
for (let k = 0; k < 5; k++) travel(C.green, [[7.3, 1.45], [7.3, 2.3], [6.3 + k * 0.5, 3.4]], 'with', 500 + k * 250, 1300, 0.2);
for (let k = 0; k < 18; k++) bubble(4.9 + ((k * 0.61) % 4.9), 600 + k * 140);
st('with', 'exit', sludge, { dur: 2500, delay: 800 });
for (let k = 0; k < 8; k++) {
  const n = dot(5.1 + k * 0.55, 4.05, C.dirty, 0.2);
  st('with', 'appear', n, { delay: 900 + k * 150 });
  st('with', 'path', n, { path: [[0.1, -0.35 - (k % 3) * 0.15]], dur: 1500, delay: 900 + k * 150 });
}
// 5 H2S -> salt, FeS -> oxide (in the water)
nextCap(cap2.caps, 4);
st('with', 'path', vH2S, { path: [[-0.1, 0.9]], dur: 900 });
st('after', 'exit', vH2S, { dur: 400 }); st('with', 'fade', vSalt, { dur: 400 });
st('with', 'exit', vGlow, { dur: 600 }); st('with', 'path', vFeS, { path: [[-0.1, -0.71]], dur: 800 });
st('after', 'exit', vFeS, { dur: 400 }); st('with', 'fade', vOx, { dur: 400 });
// 6 benzene + LEL -> flare
nextCap(cap2.caps, 5);
[[vBenz, 6.6], [vLEL, 7.9]].forEach(([n, x], k) => {
  st('with', 'path', n, { path: [[9.4 - 0.5 - x, 0], [9.4 - 0.5 - x, -1.03], [11.6 - 0.5 - x, -1.03]], dur: 1800, delay: k * 400 });
  st('with', 'exit', n, { dur: 400, delay: 1800 + k * 400 });
});
st('with', 'pulse', flame2, { dur: 300, scale: 1.4, repeat: 4, delay: 1700 });
// 7 drain to effluent tank
nextCap(cap2.caps, 6);
st('with', 'wipeout', water, { dir: 'up', dur: 2200 });
st('with', 'exit', [vSalt, vOx], { dur: 600, delay: 300 });
for (let k = 0; k < 6; k++) travel(k % 2 ? C.salt : C.dirty, [[5.6, 4.6], [5.6, 5.55], [2.15, 5.55]], 'with', 300 + k * 280, 1600);
st('with', 'wipe', bl1, { dir: 'down', dur: 600, delay: 1200 });
st('with', 'wipe', bl2, { dir: 'down', dur: 600, delay: 2000 });
// 8 gas test
nextCap(cap2.caps, 7);
finish(meter2, 0.6, 3.7, 3.3);
s.addNotes('Eight clicks: fill, heat, chemistry and rolling boil, H₂S and FeS converted in the water, vapours to the flare, drain and rinse, gas test.');

(async () => {
  await p.writeFile({ fileName: OUT });
  fs.writeFileSync(OUT.replace('.pptx', '.anim.json'), JSON.stringify(spec));
  console.log('built', OUT);
})();
