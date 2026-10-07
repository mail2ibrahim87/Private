import sys, re
sys.path.insert(0, '.')
from lib import *
D = sys.argv[1]
AMB = 'B9770E'; EQ = '94A3B8'; EQF = 'EEF2F6'; STEAM = '0C74BC'; COND = '8FC1E3'

def gas(s, bads, inj_cap, label):
    pars = []
    for k, b in enumerate(bads):
        st, en = s.shapes()[b]["span"]; xml = s.x[st:en]; i = s.nid()
        xml = re.sub(r'<p:cNvPr id="\d+" name="[^"]*"/>', '<p:cNvPr id="%d" name="dec%d"/>' % (i, i), xml, 1)
        xml = re.sub(r'<a:t>[^<]*</a:t>', '<a:t>%s</a:t>' % label, xml, 1)
        xml = re.sub(r'(<a:rPr[^>]*>\s*<a:solidFill><a:srgbClr val=")\w+', r'\g<1>' + AMB, xml, 1)
        s.x = s.x[:en] + xml + s.x[en:]
        s.retarget('exit', b, i)
        d = 1200 + k * 500
        pars += [vanish(b, d), fadein(i, d, 400), pulse(i, d + 500, 350, 1.06, 3)]
        s.build(i, [0, 1, 2])
    s.add_effects(inj_cap, pars)

def exch(s, x, y, w, h):
    out = [shape(s.nid(), 'roundRect', x, y, w, h, EQF, EQ, 1.5, 50000)]
    for k in (-1, 0, 1):
        yy = y + h / 2 + k * h * 0.2
        out.append(line(s.nid(), x + h * 0.45, yy, x + w - h * 0.45, yy, 'B4C2D3', 1))
    return out

def drum(s, x, y, w, h):
    return [shape(s.nid(), 'roundRect', x, y, w, h, EQF, EQ, 1.5, 50000),
            shape(s.nid(), 'roundRect', x + 0.06, y + h * 0.55, w - 0.12, h * 0.38, 'D6E7F5', None, 0, 50000)]

def dot(s, cx, cy, col, d=0.16):
    i = s.nid(); return i, shape(i, 'ellipse', cx - d / 2, cy - d / 2, d, d, col, 'FFFFFF', 0.75)

def travel(s, col, pts, delay, dur, d=0.16):
    i, xml = dot(s, pts[0][0], pts[0][1], col, d)
    rel = [(px_ - pts[0][0], py - pts[0][1]) for px_, py in pts[1:]]
    s.build(i, [0, 1, 3])
    return xml, [appear(i, delay), motion(i, rel, delay, dur), vanish(i, delay + dur)]

def rel_from(center, pts): return [(a - center[0], b - center[1]) for a, b in pts]

# ---------------- slide 13 vapour-phase
s = Slide(D + '/slide13.xml')
gas(s, [41, 44, 47], 52, '↓  decreasing')
s.set_geom(20, w=7.25 - 5.85)
eq = exch(s, 7.25, 1.37, 1.0, 0.36)
eq += [text(s.nid(), 'Overhead condenser', 6.85, 1.78, 1.8, 0.25, 10)]
eq += [line(s.nid(), 8.25, 1.55, 8.7, 1.55), line(s.nid(), 8.7, 1.55, 8.7, 2.3, arrow=True)]
eq += drum(s, 8.4, 2.3, 1.5, 0.6) + [text(s.nid(), 'Reflux drum', 8.3, 2.93, 1.3, 0.25, 10)]
eq += [line(s.nid(), 9.45, 2.34, 9.45, 1.55), line(s.nid(), 9.45, 1.55, 10.95, 1.55, arrow=True)]
eq += [line(s.nid(), 9.75, 2.86, 9.75, 4.75, arrow=True)]
eq += exch(s, 7.0, 4.85, 1.4, 0.5) + [text(s.nid(), 'Reboiler', 7.0, 5.37, 1.4, 0.25, 10)]
eq += [line(s.nid(), 6.43, 5.25, 7.0, 5.25, arrow=True), line(s.nid(), 7.9, 4.85, 7.9, 4.55), line(s.nid(), 7.9, 4.55, 6.5, 4.55, arrow=True)]
dots = []; e2 = []; e6 = []
for k in range(3):
    xml, p = travel(s, STEAM, [(6.45, 5.25), (7.3, 5.25), (7.9, 5.1), (7.9, 4.55), (6.55, 4.55)], 1600 + k * 600, 1800); dots.append(xml); e2 += p
for k in range(4):
    xml, p = travel(s, COND, [(9.75, 2.9), (9.75, 5.0)], 2600 + k * 350, 900); dots.append(xml); e6 += p
s.insert_before(49, ''.join(eq) + ''.join(dots))
route = [(5.85, 1.55), (8.7, 1.55), (8.7, 2.6), (9.45, 2.6), (9.45, 1.55), (10.75, 1.55)]
s.set_path(59, rel_from((5.85, 3.03), route)); s.set_path(60, rel_from((5.85, 3.63), route))
s.add_effects(51, e2); s.add_effects(55, e6)
s.save()

# ---------------- slide 14 boil-out
s = Slide(D + '/slide14.xml'); gas(s, [40, 43, 46], 52, '↓  decreasing'); s.save()

# ---------------- slide 15 viscosity flush (static overhead + reboiler)
s = Slide(D + '/slide15.xml')
eq = [line(s.nid(), 5.0, 1.8, 5.0, 1.5), line(s.nid(), 5.0, 1.5, 6.3, 1.5, arrow=True)]
eq += exch(s, 6.3, 1.32, 1.0, 0.36) + [text(s.nid(), 'Overhead condenser', 6.0, 1.72, 1.6, 0.25, 10)]
eq += [line(s.nid(), 7.3, 1.5, 8.0, 1.5), line(s.nid(), 8.0, 1.5, 8.0, 1.95, arrow=True)]
eq += drum(s, 7.7, 1.95, 1.5, 0.55) + [text(s.nid(), 'Reflux drum', 6.6, 2.1, 1.05, 0.25, 10, align='r')]
eq += [line(s.nid(), 8.95, 1.98, 8.95, 1.75), line(s.nid(), 8.95, 1.75, 9.7, 1.75, arrow=True)]
eq += [shape(s.nid(), 'rect', 9.7, 1.45, 0.2, 1.05, 'B4C2D3', EQ, 1),
       shape(s.nid(), 'ellipse', 9.56, 0.9, 0.48, 0.6, 'F7C548', None), shape(s.nid(), 'ellipse', 9.66, 1.13, 0.28, 0.36, 'FF8C1A', None),
       text(s.nid(), 'FLARE /\nCLOSED VENT'.replace('\n', ' '), 10.0, 1.6, 1.15, 0.45, 10)]
eq += [line(s.nid(), 8.5, 2.5, 8.5, 2.75, 'D9A441', 3), line(s.nid(), 8.5, 2.75, 10.6, 2.75, 'D9A441', 3)]
eq += exch(s, 3.3, 4.9, 0.95, 0.5) + [text(s.nid(), 'Reboiler', 3.2, 5.42, 1.15, 0.22, 10)]
eq += [line(s.nid(), 4.48, 5.3, 4.25, 5.3, arrow=True), line(s.nid(), 3.8, 4.9, 3.8, 4.8), line(s.nid(), 3.8, 4.8, 4.4, 4.8, arrow=True)]
s.insert_before(55, ''.join(eq)); s.save()

# ---------------- slide 16 circulation
s = Slide(D + '/slide16.xml'); gas(s, [36, 45], 56, 'Decreasing ↓'); s.save()

# ---------------- slide 17 packing
s = Slide(D + '/slide17.xml')
gas(s, [74, 77, 80], 84, '↓  decreasing')
s.set_geom(45, w=7.2 - 6.45)
eq = exch(s, 7.2, 1.42, 0.8, 0.36) + [text(s.nid(), 'Overhead condenser', 6.9, 1.15, 1.6, 0.25, 10)]
eq += [line(s.nid(), 8.0, 1.6, 8.3, 1.6), line(s.nid(), 8.3, 1.6, 8.3, 1.95, arrow=True)]
eq += drum(s, 7.6, 1.95, 1.3, 0.42) + [text(s.nid(), 'Reflux drum', 6.85, 1.98, 0.72, 0.4, 9)]
eq += [line(s.nid(), 8.65, 1.97, 8.65, 1.6), line(s.nid(), 8.65, 1.6, 9.4, 1.6)]
eq += [line(s.nid(), 8.9, 2.16, 9.04, 2.16), line(s.nid(), 9.04, 2.16, 9.04, 4.5), line(s.nid(), 9.04, 4.5, 10.75, 4.5), line(s.nid(), 10.75, 4.5, 10.75, 4.75, arrow=True)]
eq += exch(s, 4.15, 5.5, 1.05, 0.42) + [text(s.nid(), 'Reboiler', 4.15, 5.93, 1.05, 0.22, 10)]
eq += [line(s.nid(), 5.2, 5.62, 5.6, 5.62, arrow=True), line(s.nid(), 5.75, 5.72, 5.75, 5.84), line(s.nid(), 5.75, 5.84, 5.2, 5.84, arrow=True)]
dots = []; e4 = []
for k in range(3):
    xml, p = travel(s, STEAM, [(5.75, 5.75), (5.75, 5.84), (4.7, 5.84), (4.7, 5.62), (5.6, 5.62)], 800 + k * 600, 1600); dots.append(xml); e4 += p
for k in range(4):
    xml, p = travel(s, COND, [(9.04, 2.16), (9.04, 4.5), (10.75, 4.5), (10.75, 4.95)], 3200 + k * 350, 1300); dots.append(xml); e4 += p
cap_box = [i for i, d in s.shapes().items() if d['g'] and abs(d['g'][1] - 6.4) < 0.01 and abs(d['g'][2] - 11.5) < 0.01][0]
s.insert_before(cap_box, ''.join(eq) + ''.join(dots))
route = [(6.45, 1.6), (8.3, 1.6), (8.3, 2.16), (8.65, 2.16), (8.65, 1.6), (9.4, 1.6), (9.4, 2.3), (9.7, 2.3)]
s.set_path(55, rel_from((7.65, 2.68), route)); s.set_path(56, rel_from((7.65, 3.73), route))
s.add_effects(86, e4)
s.save()

# ---------------- slide 18 tank
s = Slide(D + '/slide18.xml'); gas(s, [62, 65, 68], 75, '↓  decreasing'); s.save()
print('ok')
