// Builds one decontamination sales deck per company (Delight colours, option C2 "Soft Blue").
// Usage: node build.js <outdir> <logo.png>     then animate.py adds transitions / animations.
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const { CHEM, COMPANIES } = require('./content');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');

const OUT = process.argv[2] || 'out';
const LOGO = 'image/png;base64,' + fs.readFileSync(process.argv[3]).toString('base64');
const COL = {
  soft: 'E9F2F9', tintB: 'D3E6F4', tintG: 'D9EFDC', head: '1F5F8B', blue: '0C74BC', green: '3AB54A', greenD: '2E8B3A',
  body: '374151', muted: '5B7085', panel: 'F3F7FA', line: 'E1E8EF', white: 'FFFFFF',
  eqFill: 'CBD5E1', eqLine: '94A3B8', steam: '0C74BC', chem: '3AB54A', water: 'BFDDF2', sludge: '8B6B4A',
  h2s: 'C0392B', benz: 'E67E22', fes: '6B7280', lel: 'D4A017',
};
const THEME = { name: 'Delight Soft Blue', headFontFace: 'Cambria', bodyFontFace: 'Calibri',
  colors: { dk1: '1F2937', lt1: 'FFFFFF', dk2: '1F5F8B', lt2: 'E9F2F9', accent1: '0C74BC', accent2: '3AB54A', accent3: '1F5F8B',
            accent4: '7FB8E3', accent5: '2E8B3A', accent6: '5B7085', hlink: '0C74BC', folHlink: '1F5F8B' } };
const W = 13.333, H = 7.5;

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

async function buildDeck(co) {
  const p = new pptxgen();
  p.layout = 'LAYOUT_WIDE';
  p.theme = { headFontFace: 'Cambria', bodyFontFace: 'Calibri' };
  p.author = 'Delight International';
  p.company = 'Delight Equipment International L.L.C.';
  p.title = `Decontamination – ${co.full}`;
  const footer = `Delight International  |  Decontamination – ${co.name}`;
  p.defineSlideMaster({ title: 'SOFT', background: { color: COL.soft }, objects: [] });
  p.defineSlideMaster({
    title: 'CONTENT', background: { color: COL.white },
    objects: [
      { image: { x: 12.33, y: 0.3, w: 0.55, h: 0.72, data: LOGO } },
      { text: { text: footer, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontFace: 'Calibri', fontSize: 10, color: COL.muted, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.38, w: 11.3, h: 0.85, fontFace: 'Cambria', fontSize: 28, bold: true, color: COL.head, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
    ],
    slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontFace: 'Calibri', fontSize: 10, color: COL.muted, align: 'right' },
  });

  const spec = [];
  let s, n, A;
  const newSlide = (master, transition = 'fade') => {
    s = p.addSlide({ masterName: master });
    n = 0;
    A = { transition, steps: [] };
    spec.push(A);
    return s;
  };
  const nm = (pfx) => `${pfx}_${++n}`;
  const step = (trigger, effect, targets, extra = {}) => A.steps.push({ trigger, effect, targets: [].concat(targets), ...extra });
  const T = (text, o) => { const name = o.objectName || nm('t'); s.addText(text, { fontFace: 'Calibri', color: COL.body, margin: 0, isTextBox: true, valign: 'top', ...o, objectName: name }); return name; };
  const R = (shape, o) => { const name = o.objectName || nm('s'); s.addShape(shape, { ...o, objectName: name }); return name; };
  const I = (data, o) => { const name = nm('i'); s.addImage({ data, ...o, objectName: name }); return name; };
  const title = (t) => s.addText(t, { placeholder: 'title' });
  const card = (x, y, w, h, fill = COL.panel) => R(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { color: COL.line, width: 1 },
    shadow: { type: 'outer', color: '9AA9B8', opacity: 0.22, blur: 6, offset: 2, angle: 90 } });
  const circleIcon = async (x, y, d, iconName, fill) => [R(p.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } }),
    I(await icon(iconName, 'FFFFFF'), { x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 })];
  const line = (x1, y1, x2, y2, color, width = 3.5, arrow = false, dash) => R(p.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1, flipV: y2 < y1,
    line: { color, width, endArrowType: arrow ? 'triangle' : undefined, dashType: dash } });
  const dir = (x1, y1, x2, y2) => (x2 > x1 ? 'left' : x2 < x1 ? 'right' : y2 > y1 ? 'up' : 'down');
  // wipe-in direction for a flow segment: wipe(left) reveals from the left edge
  const flow = (pts, color, width, trig = 'after', arrowLast = true, dur = 400) => {
    const names = [];
    for (let i = 0; i < pts.length - 1; i++) {
      const [x1, y1] = pts[i], [x2, y2] = pts[i + 1];
      const nmx = line(x1, y1, x2, y2, color, width, arrowLast && i === pts.length - 2);
      const d = x2 > x1 ? 'left' : x2 < x1 ? 'right' : y2 > y1 ? 'up' : 'down';
      step(i === 0 ? trig : 'after', 'wipe', nmx, { dir: d, dur });
      names.push(nmx);
    }
    return names;
  };
  const notes = (t) => s.addNotes(t);
  const bullets = (items, o) => T(items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1 } })),
    { fontSize: 14, paraSpaceAfter: 6, ...o });

  // ---------- equipment drawings ----------
  const column = (x, y, w, h, trays = 8) => {
    const out = [R(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: w / 2, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } })];
    for (let t = 1; t <= trays; t++) out.push(line(x + 0.05, y + (h * t) / (trays + 1), x + w - 0.05, y + (h * t) / (trays + 1), COL.eqLine, 1, false, 'dash'));
    return out;
  };
  const hdrum = (x, y, w, h) => [R(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } })];
  const exch = (x, y, w, h) => {
    const o = [R(p.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } })];
    o.push(R(p.shapes.RECTANGLE, { x: x + h * 0.35, y: y - 0.05, w: 0.12, h: h + 0.1, fill: { color: COL.eqLine }, line: { type: 'none' } }));
    for (let k = 1; k <= 3; k++) o.push(line(x + h * 0.5, y + (h * k) / 4, x + w - h * 0.3, y + (h * k) / 4, COL.eqLine, 1));
    return o;
  };
  const pump = (x, y, d) => [R(p.shapes.OVAL, { x, y, w: d, h: d, fill: { color: COL.white }, line: { color: COL.eqLine, width: 1.5 } }),
    R(p.shapes.ISOSCELES_TRIANGLE, { x: x + d * 0.3, y: y + d * 0.25, w: d * 0.45, h: d * 0.5, rotate: 90, fill: { color: COL.eqLine }, line: { type: 'none' } })];
  const tank = (x, y, w, h) => [
    R(p.shapes.RECTANGLE, { x, y: y + 0.25, w, h: h - 0.5, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }),
    R(p.shapes.OVAL, { x, y: y + h - 0.5, w, h: 0.5, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }),
    R(p.shapes.RECTANGLE, { x: x + 0.02, y: y + h - 0.6, w: w - 0.04, h: 0.35, fill: { color: COL.eqFill }, line: { type: 'none' } }),
    R(p.shapes.OVAL, { x, y, w, h: 0.5, fill: { color: 'DDE5EE' }, line: { color: COL.eqLine, width: 1.5 } }),
  ];
  const legend = (x, y) => [['H₂S', COL.h2s], ['Benzene', COL.benz], ['Pyrophoric FeS', COL.fes], ['LEL', COL.lel]].flatMap(([t, c], i) => [
    R(p.shapes.OVAL, { x, y: y + i * 0.32 + 0.06, w: 0.17, h: 0.17, fill: { color: c }, line: { type: 'none' } }),
    T(t, { x: x + 0.27, y: y + i * 0.32, w: 1.8, h: 0.3, fontSize: 12, color: COL.muted })]);
  const dots = (pts) => pts.map(([x, y, c]) => R(p.shapes.OVAL, { x, y, w: 0.16, h: 0.16, fill: { color: c }, line: { type: 'none' } }));
  const lbl = (t, x, y, w, color, align = 'left') => T(t, { x, y, w, h: 0.32, fontSize: 12, bold: true, color, align });

  // unit diagram in box x 0.7..6.3, y 1.6..6.5
  const drawUnit = (type) => {
    const o = [];
    if (type === 'column' || type === 'quench') {
      o.push(...column(2.6, 1.8, 1.2, 4.4, 9));
      o.push(...hdrum(4.6, 2.0, 1.5, 0.6));
      o.push(line(3.2, 1.8, 3.2, 1.6, COL.eqLine, 2), line(3.2, 1.6, 5.35, 1.6, COL.eqLine, 2), line(5.35, 1.6, 5.35, 2.0, COL.eqLine, 2, true));
      o.push(...exch(4.4, 5.3, 1.8, 0.5), line(3.8, 5.55, 4.4, 5.55, COL.eqLine, 2, true));
      if (type === 'quench') {
        o.push(...pump(1.2, 5.6, 0.45), line(2.6, 5.9, 1.65, 5.9, COL.eqLine, 2), line(1.42, 5.6, 1.42, 3.0, COL.eqLine, 2), line(1.42, 3.0, 2.6, 3.0, COL.eqLine, 2, true));
        o.push(lbl('Quench oil circulation', 0.7, 2.55, 2.0, COL.muted));
      }
      o.push(lbl('Reboiler / pump-around', 4.2, 5.85, 2.4, COL.muted), lbl('Overhead drum', 4.6, 2.65, 1.8, COL.muted));
    } else if (type === 'reactor') {
      o.push(R(p.shapes.ROUNDED_RECTANGLE, { x: 1.4, y: 1.8, w: 1.3, h: 3.6, rectRadius: 0.6, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }));
      for (let k = 0; k < 2; k++) o.push(R(p.shapes.RECTANGLE, { x: 1.5, y: 2.5 + k * 1.3, w: 1.1, h: 0.6, fill: { color: 'B4C2D3' }, line: { type: 'none' } }));
      o.push(...exch(3.3, 4.6, 2.2, 0.55), ...hdrum(3.6, 2.4, 2.0, 0.7));
      o.push(line(2.05, 5.4, 2.05, 5.9, COL.eqLine, 2), line(2.05, 5.9, 4.4, 5.9, COL.eqLine, 2), line(4.4, 5.9, 4.4, 5.15, COL.eqLine, 2, true));
      o.push(line(5.5, 4.88, 5.9, 4.88, COL.eqLine, 2), line(5.9, 4.88, 5.9, 2.75, COL.eqLine, 2), line(5.9, 2.75, 5.6, 2.75, COL.eqLine, 2, true));
      o.push(lbl('Reactor', 1.4, 5.95, 1.3, COL.muted), lbl('Effluent exchanger', 3.3, 4.22, 2.4, COL.muted), lbl('Separator', 3.6, 3.15, 2, COL.muted));
    } else if (type === 'coker') {
      o.push(R(p.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 2.0, w: 1.1, h: 3.4, rectRadius: 0.5, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }));
      o.push(R(p.shapes.ROUNDED_RECTANGLE, { x: 2.4, y: 2.0, w: 1.1, h: 3.4, rectRadius: 0.5, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }));
      o.push(...column(4.4, 1.8, 1.1, 4.3, 8));
      o.push(line(1.55, 2.0, 1.55, 1.6, COL.eqLine, 2), line(1.55, 1.6, 4.95, 1.6, COL.eqLine, 2), line(4.95, 1.6, 4.95, 1.8, COL.eqLine, 2, true), line(2.95, 2.0, 2.95, 1.6, COL.eqLine, 2));
      o.push(lbl('Coke drums', 1.2, 5.5, 2.2, COL.muted), lbl('Main fractionator', 3.9, 6.15, 2.2, COL.muted));
    } else if (type === 'separator') {
      o.push(...hdrum(1.2, 3.0, 4.8, 1.5), R(p.shapes.RECTANGLE, { x: 4.6, y: 4.5, w: 0.6, h: 0.7, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }));
      o.push(R(p.shapes.RECTANGLE, { x: 1.6, y: 4.05, w: 4.0, h: 0.35, fill: { color: COL.sludge }, line: { type: 'none' } }));
      o.push(line(0.7, 3.4, 1.2, 3.4, COL.eqLine, 2, true), line(3.6, 3.0, 3.6, 2.4, COL.eqLine, 2), line(1.8, 4.5, 1.8, 5.2, COL.eqLine, 2));
      o.push(lbl('Sludge / heavy deposits', 1.7, 5.3, 2.8, COL.muted), lbl('Boot', 4.6, 5.3, 1, COL.muted));
    } else if (type === 'tank') {
      o.push(...tank(1.4, 2.1, 4.2, 3.8));
      o.push(R(p.shapes.RECTANGLE, { x: 1.42, y: 5.0, w: 4.16, h: 0.45, fill: { color: COL.sludge }, line: { type: 'none' } }));
      o.push(R(p.shapes.STAR_8_POINT || p.shapes.STAR_5_POINT, { x: 3.25, y: 3.4, w: 0.5, h: 0.5, fill: { color: COL.blue }, line: { type: 'none' } }));
      o.push(line(3.5, 1.6, 3.5, 3.4, COL.blue, 2.5), lbl('Gamma-jet nozzle', 3.7, 2.7, 2.2, COL.blue), lbl('Bottom sludge', 1.6, 5.95, 2.2, COL.muted));
    } else if (type === 'sphere') {
      o.push(R(p.shapes.OVAL, { x: 2.0, y: 1.8, w: 3.2, h: 3.2, fill: { color: COL.eqFill }, line: { color: COL.eqLine, width: 1.5 } }));
      for (const lx of [2.3, 3.2, 4.0, 4.9]) o.push(line(lx, 4.4, lx, 6.0, COL.eqLine, 3));
      o.push(lbl('LPG sphere', 2.9, 6.1, 1.8, COL.muted));
    } else if (type === 'exchanger') {
      o.push(...exch(1.0, 2.2, 5.0, 1.0), ...exch(1.6, 4.3, 4.0, 0.8));
      o.push(lbl('Shell & tube exchangers', 1.0, 5.4, 3, COL.muted));
    }
    return o;
  };

  // ---------- slides ----------
  // 1 Title
  newSlide('SOFT', 'fade');
  R(p.shapes.OVAL, { x: 9.0, y: -1.4, w: 5.8, h: 5.8, fill: { color: COL.tintB }, line: { type: 'none' } });
  R(p.shapes.OVAL, { x: 10.6, y: 4.3, w: 3.8, h: 3.8, fill: { color: COL.tintG }, line: { type: 'none' } });
  s.addImage({ data: LOGO, x: 0.75, y: 0.6, w: 1.0, h: 1.31 });
  T('DECONTAMINATION SERVICES', { x: 0.75, y: 2.55, w: 8.5, h: 0.4, fontSize: 16, bold: true, color: COL.greenD, charSpacing: 4 });
  T(`Safe, fast entry for ${co.name}`, { x: 0.75, y: 3.0, w: 8.6, h: 1.7, fontFace: 'Cambria', fontSize: 40, bold: true, color: '1F3A52', valign: 'top' });
  T(co.location, { x: 0.75, y: 4.9, w: 8, h: 0.4, fontSize: 18, color: COL.head });
  T('Delight International  ·  Maintenance Simplified', { x: 0.75, y: 6.55, w: 8, h: 0.4, fontSize: 14, color: COL.muted });
  notes(`Introduce Delight and the purpose: a decontamination approach tailored to ${co.full}.`);

  // 2 Agenda
  newSlide('SOFT', 'push');
  s.addImage({ data: LOGO, x: 12.33, y: 0.3, w: 0.55, h: 0.72 });
  T('Agenda', { x: 0.75, y: 0.6, w: 8, h: 0.9, fontFace: 'Cambria', fontSize: 36, bold: true, color: '1F3A52' });
  const agenda = ['Decontamination and your challenges', 'Application methods and chemistry', `${co.name} units we decontaminate`, 'Execution, monitoring and why Delight'];
  for (const [i, a] of agenda.entries()) {
    const y = 1.9 + i * 1.15;
    const c = R(p.shapes.OVAL, { x: 0.8, y, w: 0.8, h: 0.8, fill: { color: i % 2 ? COL.green : COL.blue }, line: { type: 'none' } });
    const nt = T(String(i + 1).padStart(2, '0'), { x: 0.8, y, w: 0.8, h: 0.8, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    const at = T(a, { x: 1.9, y: y + 0.12, w: 9, h: 0.6, fontSize: 22, color: '1F3A52', valign: 'middle' });
    step('click', 'appear', [c, nt, at]);
  }

  const divider = (num, t, sub) => {
    newSlide('SOFT', 'fade');
    R(p.shapes.OVAL, { x: 9.6, y: 1.2, w: 5.0, h: 5.0, fill: { color: COL.tintB }, line: { type: 'none' } });
    s.addImage({ data: LOGO, x: 12.33, y: 0.3, w: 0.55, h: 0.72 });
    T(num, { x: 0.75, y: 1.75, w: 3, h: 1.2, fontFace: 'Cambria', fontSize: 66, bold: true, color: COL.blue });
    const a = T(t, { x: 0.75, y: 3.2, w: 8.5, h: 1.35, fontFace: 'Cambria', fontSize: 34, bold: true, color: '1F3A52', valign: 'bottom' });
    const b = T(sub, { x: 0.75, y: 4.75, w: 8.5, h: 0.9, fontSize: 18, color: COL.muted });
    step('with', 'ascend', [a, b], { dur: 800 });
  };

  // 3 Divider 01
  divider('01', 'Decontamination and your challenges', `What ${co.name} needs before equipment is opened`);

  // 4 Company profile
  newSlide('CONTENT');
  title(`${co.name}: what has to be removed`);
  card(0.6, 1.5, 5.6, 4.0);
  const pin = I(await icon('FaIndustry', COL.blue), { x: 0.9, y: 1.8, w: 0.5, h: 0.5 });
  T('FACILITY', { x: 1.55, y: 1.85, w: 4, h: 0.4, fontSize: 14, bold: true, color: COL.greenD, charSpacing: 2 });
  T(co.facility, { x: 0.9, y: 2.45, w: 5.0, h: 1.6, fontSize: 16 });
  T([{ text: 'Location  ', options: { bold: true, color: COL.head } }, { text: co.location }], { x: 0.9, y: 4.3, w: 5.0, h: 0.9, fontSize: 14 });
  T('KEY CONTAMINANTS', { x: 6.7, y: 1.5, w: 5, h: 0.4, fontSize: 14, bold: true, color: COL.greenD, charSpacing: 2 });
  const chips = [];
  co.contaminants.forEach((c, i) => {
    const x = 6.7 + (i % 2) * 3.0, y = 2.0 + Math.floor(i / 2) * 0.82;
    const b = R(p.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.8, h: 0.62, rectRadius: 0.31, fill: { color: i % 2 ? 'EAF6EC' : 'EAF3FA' }, line: { color: i % 2 ? 'BFE3C4' : 'C6DDF0', width: 1 } });
    const t = T(c, { x: x + 0.2, y, w: 2.5, h: 0.62, fontSize: 14, bold: true, color: '1F3A52', valign: 'middle' });
    chips.push(b, t);
  });
  step('click', 'ascend', chips, { dur: 700 });
  const fb = R(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.75, w: 12.1, h: 0.95, rectRadius: 0.1, fill: { color: 'EAF6EC' }, line: { type: 'none' } });
  const ft = T([{ text: 'Our focus:  ', options: { bold: true, color: COL.greenD } }, { text: co.focus }], { x: 0.85, y: 5.75, w: 11.6, h: 0.95, fontSize: 15, valign: 'middle' });
  step('click', 'fade', [fb, ft]);
  notes('Site-specific contaminants and focus. Unit details are confirmed in a joint site survey.');

  // 5 What is decontamination
  newSlide('CONTENT');
  title('What is decontamination?');
  const w1 = T('Decontamination prepares process equipment for safe entry and maintenance as quickly as possible – removing toxic and flammable gases, pyrophoric deposits and hydrocarbons before the equipment is opened.', { x: 0.6, y: 1.45, w: 11.4, h: 1.0, fontSize: 18 });
  const w2 = T(`Typical challenges at ${co.name}: ${co.contaminants.slice(0, 5).join(', ')}.`, { x: 0.6, y: 2.5, w: 11.4, h: 0.6, fontSize: 15, color: COL.muted, italic: true });
  step('click', 'appear', [w1, w2]);
  T('KEY ELEMENTS WE PLAN FOR', { x: 0.6, y: 3.45, w: 12, h: 0.4, fontSize: 15, bold: true, color: COL.greenD, align: 'center', charSpacing: 2 });
  const ke = [['FaProjectDiagram', 'Method of application'], ['FaClock', 'Application time'], ['FaTools', 'Mechanical cleaning needs'], ['FaClipboardList', 'Special requirements'], ['FaRecycle', 'Waste effluent quality']];
  for (const [i, [ic, t]] of ke.entries()) {
    const x = 0.95 + i * 2.45;
    const ci = await circleIcon(x + 0.35, 4.1, 1.1, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x, y: 5.35, w: 1.8, h: 0.8, fontSize: 14, bold: true, color: '1F3A52', align: 'center' });
    step(i === 0 ? 'click' : 'after', 'ascend', [...ci, tt], { dur: 600 });
  }

  // 6 Objectives
  newSlide('CONTENT');
  title('Decontamination objectives');
  const obj = [['De-oiling / de-sludging', 'Soften heavy deposits for easier removal; reduce hazardous sludge and break up the rest.'],
    ['Sulphide oxidation', 'Neutralise pyrophoric iron sulphide (FeS) and H₂S.'],
    ['Degassing', 'Release flammable hydrocarbon gases (LEL), benzene and H₂S safely.'],
    ['Neutralising', 'Correct pH and treat ammonia so effluent can be handled and disposed of.']];
  for (const [i, [a, b]] of obj.entries()) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.55;
    const c1 = card(x, y, 5.9, 2.25);
    const ci = R(p.shapes.OVAL, { x: x + 0.35, y: y + 0.45, w: 0.9, h: 0.9, fill: { color: i % 2 ? COL.green : COL.blue }, line: { type: 'none' } });
    const cn = T(String(i + 1), { x: x + 0.35, y: y + 0.45, w: 0.9, h: 0.9, fontSize: 24, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    const ta = T(a, { x: x + 1.5, y: y + 0.4, w: 4.2, h: 0.5, fontSize: 20, bold: true, color: COL.head });
    const tb = T(b, { x: x + 1.5, y: y + 0.95, w: 4.2, h: 1.1, fontSize: 15 });
    step('click', 'fade', [c1, ci, cn, ta, tb]);
  }

  // 7 Simultaneous treatment
  newSlide('CONTENT');
  title('Simultaneous treatment in one application');
  const sim = [['Benzene', 'Below entry limit', 0], ['LEL', 'Gas-free', 0], ['H₂S & mercaptans', 'Below entry limit', 1], ['Pyrophoric FeS', 'Neutralised', 1], ['Heavy organics', 'Emulsified / removed', 2], ['Residual solids', 'Conditioned', 2]];
  const grpCol = [COL.blue, COL.green, '1F5F8B'];
  for (const [i, [c, r, g]] of sim.entries()) {
    const x = 0.75 + i * 2.0;
    const at = T(c, { x, y: 1.45, w: 1.75, h: 0.75, fontSize: 15, bold: true, color: grpCol[g], align: 'center', valign: 'bottom' });
    const a = R(p.shapes.DOWN_ARROW, { x: x + 0.2, y: 2.3, w: 1.35, h: 1.6, fill: { color: grpCol[g] }, line: { type: 'none' } });
    const rt = T(r, { x, y: 4.05, w: 1.75, h: 0.6, fontSize: 14, bold: true, color: '1F3A52', align: 'center' });
    step(i === 0 ? 'click' : 'after', 'ascend', [a, at, rt], { dur: 500 });
  }
  const gl = [['Distillation', 0], ['Oxidation', 2], ['Cleaning', 4]].flatMap(([t, k]) => {
    const x = 0.75 + k * 2.0;
    return [line(x, 4.85, x + 3.75, 4.85, grpCol[k / 2], 2), T(t, { x, y: 4.95, w: 3.75, h: 0.45, fontSize: 16, bold: true, color: grpCol[k / 2], align: 'center' })];
  });
  step('click', 'fade', gl);
  const sn = T('Final entry criteria are agreed with the client and confirmed by gas testing before handover.', { x: 0.75, y: 5.8, w: 11.8, h: 0.5, fontSize: 14, italic: true, color: COL.muted });
  step('after', 'fade', sn);

  // 8 Divider 02
  divider('02', 'Application methods and chemistry', 'Choosing the right method for each piece of equipment');

  // 9 Methods overview
  newSlide('CONTENT');
  title('Application methods');
  const meth = [['FaCloud', 'Vapour-phase', ['Chemistry is carried by steam through existing steam-out lines', 'Minimal injection points', 'Fastest route to entry']],
    ['FaSyncAlt', 'Circulation', ['For specific vessels, exchangers and piping', 'Used where steam is not available']],
    ['FaTemperatureHigh', 'Boil-out', ['For vessels with large volumes of sludge', 'Water + steam + chemistry']],
    ['FaFilter', 'Pre / post treatment', ['Oil wash / viscosity flush of heavy circuits', 'Tray and packing pre-treatment', 'Neutralising post-rinse']]];
  for (const [i, [ic, t, b]] of meth.entries()) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.5 + Math.floor(i / 2) * 2.7;
    const c1 = card(x, y, 5.9, 2.45);
    const ci = await circleIcon(x + 0.3, y + 0.3, 0.85, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x: x + 1.35, y: y + 0.38, w: 4.3, h: 0.55, fontSize: 20, bold: true, color: COL.head });
    const bb = bullets(b, { x: x + 1.35, y: y + 1.0, w: 4.35, h: 1.35 });
    step('click', 'fade', [c1, ...ci, tt, bb]);
  }

  // 10–11 Vapour-phase (two build slides)
  const vpBase = (second) => {
    newSlide('CONTENT');
    title('Vapour-phase application');
    const eq = [...column(5.0, 1.75, 1.2, 4.5, 9), ...exch(7.3, 1.95, 2.0, 0.6), ...hdrum(9.6, 3.55, 1.9, 0.75), ...pump(9.95, 4.95, 0.45)];
    lbl('Column', 5.0, 6.32, 1.2, COL.muted, 'center'); lbl('Condenser', 7.3, 2.6, 2.0, COL.muted, 'center'); lbl('Reflux drum', 9.6, 4.35, 1.9, COL.muted, 'center');
    // process lines (static)
    for (const [a, b, c, d] of [[5.6, 1.75, 5.6, 1.45], [5.6, 1.45, 8.3, 1.45], [8.3, 1.45, 8.3, 1.95], [9.3, 2.25, 10.55, 2.25], [10.55, 2.25, 10.55, 3.55], [10.17, 4.3, 10.17, 4.95], [9.95, 5.17, 7.0, 5.17], [7.0, 5.17, 7.0, 2.6], [7.0, 2.6, 6.2, 2.6]])
      line(a, b, c, d, COL.eqLine, 2);
    legend(11.0, 5.4);
    return eq;
  };
  vpBase(false);
  const t1 = bullets(['After optional pre-treatment, steam is introduced through the steam-out line.', 'Optional injection points: exchangers, heaters, strippers and drums.'], { x: 0.6, y: 1.55, w: 3.6, h: 3.0, fontSize: 15 });
  step('click', 'appear', t1);
  lbl('STEAM', 3.0, 5.62, 1.1, COL.steam);
  flow([[3.0, 5.95], [5.0, 5.95]], COL.steam, 4, 'after');
  const vd1 = dots([[5.25, 2.2, COL.h2s], [5.75, 2.9, COL.benz], [5.35, 3.8, COL.fes], [5.8, 4.6, COL.lel], [5.3, 5.3, COL.h2s], [5.75, 5.7, COL.benz]]);
  flow([[5.6, 1.75], [5.6, 1.45], [8.3, 1.45], [8.3, 1.95]], '7FB8E3', 4, 'after');
  flow([[9.3, 2.25], [10.55, 2.25], [10.55, 3.55]], '7FB8E3', 4, 'after');
  notes('Steam is introduced first; the process vapour path is shown in light blue. Contaminants are released from the column internals.');

  vpBase(true);
  line(3.0, 5.95, 5.0, 5.95, COL.steam, 4, true); lbl('STEAM', 3.0, 5.62, 1.1, COL.steam);
  line(5.6, 1.75, 5.6, 1.45, '7FB8E3', 4); line(5.6, 1.45, 8.3, 1.45, '7FB8E3', 4); line(8.3, 1.45, 8.3, 1.95, '7FB8E3', 4, true);
  line(9.3, 2.25, 10.55, 2.25, '7FB8E3', 4); line(10.55, 2.25, 10.55, 3.55, '7FB8E3', 4, true);
  const t2 = bullets(['Once the overhead is steaming, decon chemistry is injected into the steam line.', 'The vapour-phase degreaser travels with the steam to every surface the steam reaches.', 'Injection continues until gas tests meet the entry criteria.'], { x: 0.6, y: 1.55, w: 3.6, h: 3.2, fontSize: 15 });
  step('click', 'appear', t2);
  const cl = lbl('DECON CHEMISTRY', 3.0, 4.55, 2.0, COL.greenD);
  step('after', 'fade', cl);
  flow([[4.3, 4.9], [4.3, 5.9]], COL.chem, 4, 'after');
  const vd2 = dots([[8.0, 1.3, COL.h2s], [8.6, 1.3, COL.benz], [10.4, 2.7, COL.lel], [10.4, 3.1, COL.h2s]]);
  step('after', 'dissolve', vd2, { dur: 700 });
  const ok = R(p.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 5.85, w: 4.2, h: 0.6, rectRadius: 0.3, fill: { color: 'EAF6EC' }, line: { color: COL.green, width: 1.5 } });
  const okt = T('✓  Gas-free and pyrophoric-safe', { x: 6.7, y: 5.85, w: 4.2, h: 0.6, fontSize: 15, bold: true, color: COL.greenD, align: 'center', valign: 'middle' });
  step('after', 'ascend', [ok, okt], { dur: 700 });

  // 12 Boil-out
  newSlide('CONTENT');
  title('Boil-out application');
  hdrum(5.4, 2.55, 5.6, 2.4);
  R(p.shapes.RECTANGLE, { x: 5.9, y: 4.35, w: 4.6, h: 0.32, fill: { color: COL.sludge }, line: { type: 'none' } });
  lbl('Sludge', 11.15, 4.3, 1.0, COL.sludge);
  const bt = bullets(['Water filled above the sludge level.', 'Steam injected at several points to heat the vessel.', 'Water-based solvent / degreaser added at temperature.', 'Boil-out continues until gas tests are met; vapours vented safely.'], { x: 0.6, y: 1.55, w: 3.5, h: 3.8, fontSize: 15 });
  step('click', 'appear', bt);
  lbl('Add water', 4.25, 3.0, 1.2, COL.blue);
  flow([[4.3, 3.4], [5.4, 3.4]], COL.blue, 4, 'after');
  const water = R(p.shapes.ROUNDED_RECTANGLE, { x: 5.55, y: 3.35, w: 5.3, h: 1.45, rectRadius: 0.6, fill: { color: COL.water, transparency: 20 }, line: { type: 'none' } });
  step('after', 'wipe', water, { dir: 'down', dur: 1200 });
  const stl = [6.6, 8.2, 9.8].map((x) => { const l = line(x, 5.9, x, 4.95, COL.steam, 4, true); return l; });
  step('after', 'wipe', stl, { dir: 'down', dur: 500 });
  ['STEAM', 'STEAM', 'STEAM'].forEach((t, i) => lbl(t, 6.2 + i * 1.6, 5.95, 0.9, COL.steam, 'center'));
  lbl('CHEMISTRY', 7.6, 1.45, 1.6, COL.greenD, 'center');
  const chl = line(8.4, 1.8, 8.4, 2.55, COL.chem, 4, true);
  step('after', 'wipe', chl, { dir: 'up', dur: 500 });
  line(10.3, 2.75, 10.3, 1.6, COL.eqLine, 2.5); lbl('Vent', 10.45, 1.55, 0.8, COL.muted);
  const bub = [];
  for (let k = 0; k < 14; k++) bub.push(R(p.shapes.OVAL, { x: 5.9 + (k * 0.37) % 4.6, y: 3.55 + ((k * 7) % 5) * 0.14, w: 0.16, h: 0.16, fill: { color: 'FFFFFF' }, line: { color: '7FB8E3', width: 1 } }));
  bub.forEach((b, k) => step(k === 0 ? 'after' : 'with', 'dissolve', b, { dur: 600, delay: k * 120 }));
  legend(11.2, 5.45);

  // 13 Circulation
  newSlide('CONTENT');
  title('Circulation application – viscosity flush');
  column(5.6, 1.6, 1.1, 4.3, 8); exch(8.3, 3.9, 2.2, 0.6); pump(9.2, 6.0, 0.45);
  lbl('Bottoms circuit', 4.2, 5.95, 1.35, COL.muted, 'right'); lbl('Exchanger', 8.3, 4.55, 2.2, COL.muted, 'center');
  const ct = bullets(['Unit shut down and de-inventoried under plant procedures.', 'Cutter stock / flushing oil circulated through the bottoms circuit and exchangers.', 'Oil-based solvent added to penetrate and break down heavy deposits.'], { x: 0.6, y: 1.55, w: 4.0, h: 3.6, fontSize: 15 });
  step('click', 'appear', ct);
  const ol = lbl('OIL-BASED SOLVENT', 7.0, 5.0, 2.2, COL.greenD);
  step('after', 'fade', ol);
  flow([[9.42, 5.35], [9.42, 6.0]], COL.chem, 4, 'after');
  flow([[6.15, 5.9], [6.15, 6.22], [9.2, 6.22]], COL.blue, 4, 'after', true, 350);
  flow([[9.65, 6.22], [11.0, 6.22], [11.0, 4.2], [10.5, 4.2]], COL.blue, 4, 'after', true, 350);
  flow([[8.3, 4.2], [7.4, 4.2], [7.4, 2.8], [6.7, 2.8]], COL.blue, 4, 'after', true, 350);
  legend(11.4, 1.6);

  // 14 Chemistry toolkit
  newSlide('CONTENT');
  title('Decontamination chemistry toolkit');
  const tk = [['VD', 'FaWind', 'Carried by steam; dissolves and lifts hydrocarbons; releases LEL and benzene for safe venting.', 'Columns, overheads, exchangers'],
    ['WB', 'FaTint', 'Emulsifies oil and sludge in water; non-flammable, low odour.', 'Boil-out, circulation, tanks, rinses'],
    ['OB', 'FaOilCan', 'Boosts flushing oil to penetrate heavy deposits, tar and coke.', 'Bottoms circuits, coker, quench oil'],
    ['OX', 'FaShieldAlt', 'Converts pyrophoric iron sulphide and H₂S into safe compounds.', 'Packed beds, amine, SWS, SRU'],
    ['NE', 'FaBalanceScale', 'Corrects pH and treats ammonia / acid residues for disposal.', 'Post-rinse, effluent tanks'],
    ['AF', 'FaWater', 'Controls foaming and keeps solids dispersed.', 'Amine systems, boil-outs']];
  const tkn = [];
  for (const [i, [k, ic, d, u]] of tk.entries()) {
    const x = 0.6 + (i % 3) * 4.1, y = 1.5 + Math.floor(i / 3) * 2.6;
    const c1 = card(x, y, 3.85, 2.35);
    const ci = await circleIcon(x + 0.25, y + 0.25, 0.65, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(CHEM[k], { x: x + 1.05, y: y + 0.27, w: 2.7, h: 0.65, fontSize: 15, bold: true, color: COL.head, valign: 'middle' });
    const dd = T(d, { x: x + 0.25, y: y + 1.0, w: 3.4, h: 0.85, fontSize: 13 });
    const uu = T([{ text: 'Use: ', options: { bold: true, color: COL.greenD } }, { text: u }], { x: x + 0.25, y: y + 1.85, w: 3.4, h: 0.4, fontSize: 12, color: COL.muted });
    tkn.push([c1, ...ci, tt, dd, uu]);
  }
  tkn.forEach((g, i) => step(i === 0 ? 'click' : 'after', 'fade', g, { dur: 400 }));
  T('Chemistry is selected per unit after a compatibility check; product data and SDS are shared with the client.', { x: 0.6, y: 6.6, w: 11.5, h: 0.35, fontSize: 12, italic: true, color: COL.muted });

  // 15 Divider 03
  divider('03', `${co.name} units we decontaminate`, 'Common challenges, chemistry and treatment options per unit');

  // unit slides
  for (const u of co.units) {
    newSlide('CONTENT');
    title(u.name);
    drawUnit(u.draw);
    const blocks = [['COMMON CHALLENGES', u.ch], ['CHEMISTRY USED', u.chem.map((c) => CHEM[c])], ['TREATMENT OPTIONS', u.opt]];
    let y = 1.5;
    for (const [h, items] of blocks) {
      const hh = 0.38 + items.length * 0.42;
      const hn = T(h, { x: 6.9, y, w: 5.8, h: 0.35, fontSize: 14, bold: true, color: COL.greenD, charSpacing: 2 });
      const bn = bullets(items, { x: 6.9, y: y + 0.4, w: 5.8, h: hh - 0.3, fontSize: 15, paraSpaceAfter: 3 });
      step('click', 'ascend', [hn, bn], { dur: 700 });
      y += hh + 0.35;
    }
    T('Indicative – confirmed during the joint site survey.', { x: 0.6, y: 6.6, w: 6, h: 0.3, fontSize: 10, italic: true, color: COL.muted });
  }

  // Divider 04
  divider('04', 'Execution, monitoring and why Delight', 'From site survey to handover');

  // Execution timeline
  newSlide('CONTENT');
  title('How we execute');
  const ex = [['FaSearch', 'Site survey', 'P&IDs, isometrics, data sheets, injection points'], ['FaPencilRuler', 'Decon plan', 'Method statement, chemistry, volumes, waste plan'],
    ['FaTruck', 'Mobilise & rig-up', 'Pumps, hoses, injection manifolds, effluent tanks'], ['FaCogs', 'Decontaminate', 'Vapour-phase, circulation or boil-out with live monitoring'],
    ['FaCheckCircle', 'Validate & hand over', 'Gas tests against entry criteria, sign-off'], ['FaRecycle', 'Waste & demob', 'Neutralise, test and dispose via licensed contractor']];
  const tl = line(1.2, 2.65, 12.1, 2.65, COL.line, 3);
  step('click', 'wipe', tl, { dir: 'left', dur: 800 });
  for (const [i, [ic, t, d]] of ex.entries()) {
    const x = 0.7 + i * 2.05;
    const ci = await circleIcon(x + 0.45, 2.15, 1.0, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x, y: 3.35, w: 1.9, h: 0.65, fontSize: 15, bold: true, color: COL.head, align: 'center' });
    const dd = T(d, { x, y: 4.0, w: 1.9, h: 1.4, fontSize: 13, align: 'center' });
    step('after', 'ascend', [...ci, tt, dd], { dur: 500 });
  }

  // Responsibilities
  newSlide('CONTENT');
  title('Division of responsibilities');
  const rs = [['DELIGHT', COL.blue, ['Decon engineering and method statement', 'Specialist supervisors and technicians', 'Decontamination chemistry', 'Pumps, injection hoses and manifolds', 'Gas monitoring during decontamination', 'Effluent neutralisation and waste coordination']],
    [co.name.toUpperCase(), COL.green, ['Isolation, de-inventory and permits', 'Steam, water, nitrogen and power utilities', 'Access to control room / DCS readings', 'Operators to drain low points', 'Laboratory support for sample analysis', 'Single point of contact']]];
  for (const [i, [h, c, items]] of rs.entries()) {
    const x = 0.6 + i * 6.15;
    const c1 = card(x, 1.5, 5.9, 5.1);
    const hb = R(p.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: 1.8, w: 5.3, h: 0.6, rectRadius: 0.3, fill: { color: c }, line: { type: 'none' } });
    const ht = T(h, { x: x + 0.3, y: 1.8, w: 5.3, h: 0.6, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', charSpacing: 1 });
    const bl = bullets(items, { x: x + 0.45, y: 2.7, w: 5.1, h: 3.7, fontSize: 15, paraSpaceAfter: 8 });
    step('click', 'ascend', [c1, hb, ht, bl], { dur: 700 });
  }

  // Monitoring
  newSlide('CONTENT');
  title('Monitoring and validation');
  const mo = [['FaTachometerAlt', 'Multi-gas meter', 'LEL, H₂S, O₂ and CO at vents and manways'], ['FaVial', 'Benzene / VOC meter', 'PID readings and detector tubes for benzene'],
    ['FaFlask', 'Gas sampling', 'Sample bombs to the client laboratory where required'], ['FaTint', 'Effluent checks', 'pH, COD and oil content before disposal']];
  for (const [i, [ic, t, d]] of mo.entries()) {
    const x = 0.6 + i * 3.08;
    const c1 = card(x, 1.55, 2.85, 3.4);
    const ci = await circleIcon(x + 0.9, 1.85, 1.05, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x: x + 0.15, y: 3.05, w: 2.55, h: 0.5, fontSize: 16, bold: true, color: COL.head, align: 'center' });
    const dd = T(d, { x: x + 0.2, y: 3.6, w: 2.45, h: 1.2, fontSize: 14, align: 'center' });
    step(i === 0 ? 'click' : 'after', 'ascend', [c1, ...ci, tt, dd], { dur: 600 });
  }
  const mb = R(p.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.35, w: 12.1, h: 1.0, rectRadius: 0.1, fill: { color: 'EAF6EC' }, line: { type: 'none' } });
  const mt = T('Decontamination is complete when the client’s entry criteria are met and confirmed by validation readings – then the equipment is handed over for opening.', { x: 0.85, y: 5.35, w: 11.6, h: 1.0, fontSize: 15, valign: 'middle', color: '1F3A52' });
  step('click', 'fade', [mb, mt]);

  // Waste & environment
  newSlide('CONTENT');
  title('Less waste, less exposure');
  const we = [['FaTint', 'Less water than water washing', 'Vapour-phase uses steam already available on the unit.'], ['FaUserShield', 'Less confined-space entry', 'Equipment is gas-free and pyrophoric-safe before it is opened.'],
    ['FaFlask', 'Controlled effluent', 'Collected in holding tanks, neutralised and tested before disposal.'], ['FaTruck', 'Licensed disposal', 'Waste handled through approved contractors with full documentation.']];
  for (const [i, [ic, t, d]] of we.entries()) {
    const y = 1.55 + i * 1.25;
    const ci = await circleIcon(0.8, y, 0.9, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x: 2.0, y: y - 0.02, w: 9.5, h: 0.45, fontSize: 18, bold: true, color: COL.head });
    const dd = T(d, { x: 2.0, y: y + 0.45, w: 9.5, h: 0.45, fontSize: 15 });
    step('click', 'ascend', [...ci, tt, dd], { dur: 600 });
  }

  // Why Delight
  newSlide('CONTENT');
  title('Why Delight');
  const wd = [['FaCertificate', 'Certified QHSE', 'ISO 9001, ISO 14001 and ISO 45001 management system'], ['FaGlobeAsia', 'Regional presence', 'UAE, Saudi Arabia, Oman, Qatar and India'],
    ['FaLayerGroup', 'Integrated services', 'Decontamination, chemical cleaning, hydro jetting, bundle extraction'], ['FaTruckMoving', 'Own fleet', 'HP / UHP units and 22 T – 65 T bundle extractors'],
    ['FaUsers', 'Experienced crews', 'Trained supervisors and technicians for shutdown work'], ['FaHandshake', 'One point of responsibility', 'From survey and planning to waste handling and handover']];
  for (const [i, [ic, t, d]] of wd.entries()) {
    const x = 0.6 + (i % 3) * 4.1, y = 1.5 + Math.floor(i / 3) * 2.6;
    const c1 = card(x, y, 3.85, 2.3);
    const ci = await circleIcon(x + 0.25, y + 0.3, 0.75, ic, i % 2 ? COL.green : COL.blue);
    const tt = T(t, { x: x + 1.15, y: y + 0.32, w: 2.6, h: 0.7, fontSize: 16, bold: true, color: COL.head, valign: 'middle' });
    const dd = T(d, { x: x + 0.25, y: y + 1.2, w: 3.4, h: 0.95, fontSize: 14 });
    step(i === 0 ? 'click' : 'after', 'fade', [c1, ...ci, tt, dd], { dur: 400 });
  }

  // Closing
  newSlide('SOFT', 'fade');
  R(p.shapes.OVAL, { x: 9.0, y: -1.4, w: 5.8, h: 5.8, fill: { color: COL.tintB }, line: { type: 'none' } });
  R(p.shapes.OVAL, { x: 10.6, y: 4.3, w: 3.8, h: 3.8, fill: { color: COL.tintG }, line: { type: 'none' } });
  s.addImage({ data: LOGO, x: 0.75, y: 0.6, w: 1.0, h: 1.31 });
  T('Thank you', { x: 0.75, y: 2.35, w: 8, h: 1.0, fontFace: 'Cambria', fontSize: 44, bold: true, color: '1F3A52' });
  T(`Let’s plan the decontamination for your next ${co.name} shutdown.`, { x: 0.75, y: 3.4, w: 8.5, h: 0.9, fontSize: 20, color: COL.head });
  T([{ text: 'Delight Equipment International L.L.C.', options: { bold: true, breakLine: true } }, { text: 'P.O. Box 2932, Abu Dhabi, United Arab Emirates', options: { breakLine: true } },
    { text: 'Tel +971 2 5515641  ·  info@delightintl.ae  ·  www.delightintl.ae', options: { breakLine: true } },
    { text: 'KSA (Jubail) +966 59348 1459  ·  India (Udupi) +91 9686433661' }], { x: 0.75, y: 4.7, w: 9, h: 1.6, fontSize: 14, color: COL.body, paraSpaceAfter: 4 });

  const file = path.join(OUT, `Delight_Decontamination_${co.key}.pptx`);
  await p.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  fs.writeFileSync(file.replace('.pptx', '.anim.json'), JSON.stringify(spec));
  return file;
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const only = process.env.ONLY;
  for (const co of COMPANIES) {
    if (only && co.key !== only) continue;
    const f = await buildDeck(co);
    console.log('built', f);
  }
})();
