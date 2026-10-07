import re
import px
E = 914400
SW, SH = 13.333, 7.5

def emu(v): return str(int(round(v * E)))

def _body():
    return '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-AE"/></a:p></p:txBody>'

def shape(i, prst, x, y, w, h, fill=None, line='94A3B8', lw=1.5, adj=None, name=None):
    av = '<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>' % adj if adj is not None else '<a:avLst/>'
    f = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill if fill else '<a:noFill/>'
    ln = '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % (lw * 12700, line) if line else '<a:ln><a:noFill/></a:ln>'
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="%s">%s</a:prstGeom>%s%s</p:spPr>%s</p:sp>') % (i, name or 'n%d' % i, emu(x), emu(y), emu(w), emu(h), prst, av, f, ln, _body())

def line(i, x1, y1, x2, y2, color='94A3B8', lw=3, arrow=False):
    fl = []
    if x2 < x1: fl.append('flipH="1"')
    if y2 < y1: fl.append('flipV="1"')
    tail = '<a:tailEnd type="triangle"/>' if arrow else ''
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="n%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm%s><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:prstDash val="solid"/>%s</a:ln></p:spPr>%s</p:sp>') % (
        i, i, (' ' + ' '.join(fl)) if fl else '', emu(min(x1, x2)), emu(min(y1, y2)), emu(abs(x2 - x1)), emu(abs(y2 - y1)), lw * 12700, color, tail, _body())

def text(i, t, x, y, w, h, size=11, color='5B7085', bold=True, align='ctr'):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="n%d"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr marL="0" indent="0" algn="%s"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="%d" b="%d" dirty="0"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="Calibri" pitchFamily="34" charset="0"/></a:rPr><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>') % (
        i, i, emu(x), emu(y), emu(w), emu(h), align, size * 100, 1 if bold else 0, color, t)

# ---------- animation snippets (cTn ids are placeholders "X", renumbered later)
def _vis(spid, delay, cls, grp, val, node='withEffect'):
    return ('<p:par><p:cTn id="X" presetID="1" presetClass="%s" presetSubtype="0" fill="hold" grpId="%d" nodeType="%s"><p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="X" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="%s"/></p:to></p:set>'
            '</p:childTnLst></p:cTn></p:par>') % (cls, grp, node, delay, spid, val)
def appear(spid, delay=0): return _vis(spid, delay, 'entr', 0, 'visible')
def vanish(spid, delay=0): return _vis(spid, delay, 'exit', 1, 'hidden')
def fadein(spid, delay=0, dur=400):
    return ('<p:par><p:cTn id="X" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="X" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="X" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>') % (delay, spid, dur, spid)
def pulse(spid, delay=0, dur=400, scale=1.1, repeat=3):
    by = int(scale * 100000)
    return ('<p:par><p:cTn id="X" presetID="6" presetClass="emph" presetSubtype="0" repeatCount="%d" fill="hold" grpId="2" nodeType="withEffect"><p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
            '<p:animScale><p:cBhvr><p:cTn id="X" dur="%d" autoRev="1" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr><p:by x="%d" y="%d"/></p:animScale></p:childTnLst></p:cTn></p:par>') % (repeat * 1000, delay, dur, spid, by, by)
def path_str(pts):  # pts: list of (dx,dy) inches relative to start
    s = 'M 0 0 ' + ' '.join('L %.4f %.4f' % (dx / SW, dy / SH) for dx, dy in pts) + ' E'
    return s, 'A' * (len(pts) + 1)
def motion(spid, pts, delay=0, dur=1500):
    p, t = path_str(pts)
    return ('<p:par><p:cTn id="X" presetID="0" presetClass="path" presetSubtype="0" fill="hold" grpId="3" nodeType="withEffect"><p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
            '<p:animMotion origin="layout" path="%s" pathEditMode="relative" ptsTypes="%s"><p:cBhvr><p:cTn id="X" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>') % (delay, p, t, dur, spid)

class Slide:
    def __init__(self, path):
        self.path = path; self.x = open(path, encoding='utf-8').read()
        self.next = max(int(v) for v in re.findall(r'<p:cNvPr id="(\d+)"', self.x)) + 1
        self.bld = []
    def nid(self):
        i = self.next; self.next += 1; return i
    def shapes(self): return px.shapes(self.x)
    def insert_before(self, spid, xml):
        s, e = self.shapes()[spid]['span']; self.x = self.x[:s] + xml + self.x[s:]
    def set_geom(self, spid, x=None, y=None, w=None, h=None):
        S = self.shapes()[spid]; s, e = S['span']; b = self.x[s:e]; g = S['g']
        nx, ny, nw, nh = [v if v is not None else o for v, o in zip((x, y, w, h), g)]
        b = re.sub(r'<a:off x="-?\d+" y="-?\d+"/><a:ext cx="\d+" cy="\d+"/>', '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (emu(nx), emu(ny), emu(nw), emu(nh)), b, 1)
        self.x = self.x[:s] + b + self.x[e:]
    # timing
    def _tim(self): return self.x.find('<p:timing')
    def par_span(self, ctn_pos):
        s = self.x.rfind('<p:par>', 0, ctn_pos); d = 0
        for m in re.finditer(r'<p:par>|</p:par>', self.x[s:]):
            d += 1 if m.group() == '<p:par>' else -1
            if d == 0: return s, s + m.end()
    def click_effect_end(self, caption_spid):
        """end of the clickEffect par of the click that brings in caption_spid (insert withEffects there)"""
        t0 = self._tim()
        m = re.search(r'<p:cTn id="\d+" presetID="10" presetClass="entr"[^>]*nodeType="withEffect">(?:(?!</p:par>).)*?<p:spTgt spid="%d"/>' % caption_spid, self.x[t0:], re.S)
        pos = t0 + m.start()
        # walk back to the clickEffect in same sub-par
        ce = self.x.rfind('nodeType="clickEffect"', 0, pos)
        return self.par_span(ce)[1]
    def add_effects(self, caption_spid, pars):
        p = self.click_effect_end(caption_spid); self.x = self.x[:p] + ''.join(pars) + self.x[p:]
    def retarget(self, cls, old, new):
        """change the target of the (single) <cls> effect on spid old to spid new"""
        t0 = self._tim()
        for m in re.finditer(r'<p:cTn id="\d+" presetID="\d+" presetClass="%s"' % cls, self.x[t0:]):
            s, e = self.par_span(t0 + m.start())
            if 'spid="%d"' % old in self.x[s:e]:
                self.x = self.x[:s] + self.x[s:e].replace('spid="%d"' % old, 'spid="%d"' % new) + self.x[e:]
                return True
        raise Exception('no %s effect on %d' % (cls, old))
    def paths_for(self, spid):
        t0 = self._tim()
        return [(t0 + m.start(), m.group(1)) for m in re.finditer(r'<p:animMotion origin="layout" path="([^"]*)"[^>]*><p:cBhvr><p:cTn id="\d+" dur="\d+" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/>' % spid, self.x[t0:])]
    def set_path(self, spid, pts, which=0):
        pos, old = self.paths_for(spid)[which]; p, t = path_str(pts)
        seg = self.x[pos:pos + 400]
        seg2 = re.sub(r'path="[^"]*"', 'path="%s"' % p, seg, 1); seg2 = re.sub(r'ptsTypes="\w+"', 'ptsTypes="%s"' % t, seg2, 1)
        self.x = self.x[:pos] + seg2 + self.x[pos + 400:]
    def build(self, spid, grps):
        for g in grps: self.bld.append('<p:bldP spid="%d" grpId="%d" animBg="1"/>' % (spid, g))
    def save(self):
        x = self.x
        if self.bld: x = x.replace('</p:bldLst>', ''.join(self.bld) + '</p:bldLst>')
        t0 = x.find('<p:timing'); n = [0]
        def rn(m): n[0] += 1; return '<p:cTn id="%d"' % n[0]
        x = x[:t0] + re.sub(r'<p:cTn id="(\d+|X)"', rn, x[t0:])
        open(self.path, 'w', encoding='utf-8').write(x)
