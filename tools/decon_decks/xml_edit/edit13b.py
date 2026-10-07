import sys, re
sys.path.insert(0, '.')
from lib import *
s = Slide(sys.argv[1])
NEWDUR = 4200
route = [(5.65, 5.3), (5.65, 2.1), (5.85, 1.75), (5.85, 1.55), (8.7, 1.55), (8.7, 2.6), (9.45, 2.6), (9.45, 1.55), (11.05, 1.55), (11.05, 1.35)]
last = 0
for spid in range(65, 72):
    g = s.shapes()[spid]['g']; c = (g[0] + g[2] / 2, g[1] + g[3] / 2)
    s.set_path(spid, rel_from(c, route) if False else [(a - c[0], b - c[1]) for a, b in route])
    t0 = s._tim()
    # path par: change dur; read its start delay
    m = re.search(r'<p:cTn id="\d+" presetID="0" presetClass="path"[^>]*><p:stCondLst><p:cond delay="(\d+)"/></p:stCondLst><p:childTnLst><p:animMotion[^>]*><p:cBhvr><p:cTn id="\d+" dur="(\d+)"[^>]*/><p:tgtEl><p:spTgt spid="%d"/>' % spid, s.x[t0:])
    d0 = int(m.group(1)); a, b = t0 + m.start(2), t0 + m.end(2); s.x = s.x[:a] + str(NEWDUR) + s.x[b:]
    # exit par: set delay
    t0 = s._tim()
    m = re.search(r'<p:cTn id="\d+" presetID="1" presetClass="exit"[^>]*><p:stCondLst><p:cond delay="(\d+)"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn[^>]*><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%d"/>' % spid, s.x[t0:])
    a, b = t0 + m.start(1), t0 + m.end(1); s.x = s.x[:a] + str(d0 + NEWDUR) + s.x[b:]
    last = max(last, d0 + NEWDUR); first = d0 if spid == 65 else first
arr = first + NEWDUR
i = s.nid(); tag = text(i, '✓  Steam reached flare', 10.0, 3.08, 2.1, 0.3, 11, '2E8B3A')
s.insert_before(49, tag); s.build(i, [0, 1])
s.add_effects(51, [pulse(22, arr, 350, 1.25, 4), pulse(23, arr, 350, 1.25, 4), fadein(i, arr + 200, 500)])
s.add_effects(52, [vanish(i, 0)])
s.save(); print('first arrival', arr, 'last', last)
