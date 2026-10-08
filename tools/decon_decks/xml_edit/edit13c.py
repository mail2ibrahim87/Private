import sys, re
sys.path.insert(0, '.')
from lib import *
s = Slide(sys.argv[1])

def retime(spid, route, newdur):
    g = s.shapes()[spid]['g']; c = (g[0] + g[2] / 2, g[1] + g[3] / 2)
    s.set_path(spid, [(a - c[0], b - c[1]) for a, b in route])
    t0 = s._tim()
    m = re.search(r'<p:cTn id="\d+" presetID="0" presetClass="path"[^>]*><p:stCondLst><p:cond delay="(\d+)"/></p:stCondLst><p:childTnLst><p:animMotion[^>]*><p:cBhvr><p:cTn id="\d+" dur="(\d+)"[^>]*/><p:tgtEl><p:spTgt spid="%d"/>' % spid, s.x[t0:])
    d0 = int(m.group(1)); a, b = t0 + m.start(2), t0 + m.end(2); s.x = s.x[:a] + str(newdur) + s.x[b:]
    t0 = s._tim()
    m = re.search(r'<p:cTn id="\d+" presetID="1" presetClass="exit"[^>]*><p:stCondLst><p:cond delay="(\d+)"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn[^>]*><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%d"/>' % spid, s.x[t0:])
    a, b = t0 + m.start(1), t0 + m.end(1); s.x = s.x[:a] + str(d0 + newdur) + s.x[b:]

OVH = [(5.65, 2.1), (5.85, 1.75), (5.85, 1.55), (8.7, 1.55), (8.7, 2.6)]
# chemistry: injection point -> steam line -> column -> overhead -> reflux drum -> condensed to effluent
for spid in range(72, 78):
    retime(spid, [(3.0, 4.82), (3.0, 5.3), (5.65, 5.3)] + OVH + [(9.75, 2.75), (9.75, 5.0)], 4600)
# continuous steam to the flare during steps 3-7
STEAM = [(1.3, 5.3), (5.65, 5.3)] + OVH + [(9.45, 2.6), (9.45, 1.55), (11.05, 1.55), (11.05, 1.35)]
xmls = []
for cap in (52, 53, 54, 55, 56):
    pars = []
    for k in range(4):
        i = s.nid(); xmls.append(shape(i, 'ellipse', 1.2, 5.2, 0.2, 0.2, 'FFFFFF', '0C74BC', 2))
        rel = [(a - 1.3, b - 5.3) for a, b in STEAM[1:]]
        d = 300 + k * 700
        pars += [appear(i, d), motion(i, rel, d, 4200), vanish(i, d + 4200)]
        s.build(i, [0, 1, 3])
    s.add_effects(cap, pars)
s.insert_before(49, ''.join(xmls))
s.save(); print('ok')
