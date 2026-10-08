import sys
sys.path.insert(0, '.')
from lib import *
D = sys.argv[1]

def move(s, spid, x=None, y=None):
    s.set_geom(spid, x=x, y=y)

def reroute(s, spid, x, y, route):
    s.set_geom(spid, x=x, y=y)
    g = s.shapes()[spid]['g']; c = (g[0] + g[2] / 2, g[1] + g[3] / 2)
    s.set_path(spid, [(a - c[0], b - c[1]) for a, b in route])

# slide 13: H2S/salt to the bottom, FeS/iron oxide to the top beds
s = Slide(D + '/slide13.xml')
move(s, 58, y=4.66); move(s, 63, y=4.66)          # H2S, stable salt
move(s, 61, y=2.2); move(s, 62, y=2.27); move(s, 64, y=2.27)   # FeS glow, FeS, iron oxide
drain = [(5.85, 6.05), (9.2, 6.05), (9.2, 5.5), (9.85, 5.5)]
for i in (78, 79, 80, 81): reroute(s, i, 5.74, 4.73, drain)   # salt packets from the bottom
for i in (82, 83, 84, 85): reroute(s, i, 5.74, 2.34, drain)   # oxide packets washed down from the top
s.save()

# slide 14 boil-out: H2S starts in the bottom sludge and is taken up into the water
s = Slide(D + '/slide14.xml')
s.set_geom(57, y=3.95)
s.set_path(57, [(-0.1, 3.3 - 3.95)])
s.save()

# slide 16 horizontal vessel: H2S / stable salt at the bottom
s = Slide(D + '/slide16.xml')
move(s, 61, x=5.05, y=3.92); move(s, 65, x=4.95, y=3.92)
s.save()

# slide 18 tank: H2S / stable salt at the floor
s = Slide(D + '/slide18.xml')
move(s, 79, y=4.85); move(s, 82, y=4.85)
s.save()
print('ok')

# --- refinements: slide 14 H2S tag in front of the sludge; slide 16 H2S/salt at bottom-right near the drain
s = Slide(D + '/slide14.xml')
S = s.shapes(); a, b = S[57]['span']; tagxml = s.x[a:b]; s.x = s.x[:a] + s.x[b:]
last = max(i for i in s.shapes() if 'Slide Number' not in s.shapes()[i]['name'])
s.x = s.x.replace('</p:spTree>', tagxml + '</p:spTree>', 1)
s.save()
s = Slide(D + '/slide16.xml')
move(s, 61, x=7.4, y=3.85); move(s, 65, x=7.32, y=3.85)
s.save()
print('refined')
