"""Add native PowerPoint transitions and entrance animations to a pptxgenjs deck.

spec.json: [{"transition": "fade"|"push", "steps": [{"trigger": "click"|"with"|"after",
            "effect": "appear"|"fade"|"ascend"|"wipe"|"dissolve", "dir": "left|right|up|down",
            "dur": ms, "delay": ms, "targets": [objectName, ...]}]}, ...]   (one entry per slide)
Effects: appear, fade, dissolve, ascend, wipe, zoom (entrance); exit, vanish (exit); pulse (emphasis, "repeat");
path (motion path, "path": [[dx, dy], ...] in inches). Transitions: fade, push, morph (PowerPoint 2019 / 365), none.
Shape names starting with "!!" are matched across slides by Morph.
"""
import json
import re
import sys
import zipfile

WIPE = {"left": (8, "wipe(left)"), "up": (1, "wipe(up)"), "right": (2, "wipe(right)"), "down": (4, "wipe(down)")}


class Ids:
    def __init__(self):
        self.n = 2

    def next(self):
        self.n += 1
        return self.n


SLIDE_W, SLIDE_H = 13.333, 7.5
GRP = {"entr": 0, "exit": 1, "emph": 2, "path": 3}


def _tgt(spid):
    return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'


def _vis(ids, spid, val, delay=0):
    return (f'<p:set><p:cBhvr><p:cTn id="{ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst></p:cTn>'
            f'{_tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="{val}"/></p:to></p:set>')


def _filt(ids, spid, flt, dur, tr="in"):
    return (f'<p:animEffect transition="{tr}" filter="{flt}"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}"/>'
            f'{_tgt(spid)}</p:cBhvr></p:animEffect>')


def _anim(ids, spid, attr, a, b, dur, a_flt=False):
    va = f'<p:fltVal val="{a}"/>' if a_flt else f'<p:strVal val="{a}"/>'
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}" fill="hold"/>'
            f'{_tgt(spid)}<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
            f'<p:tav tm="0"><p:val>{va}</p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="{b}"/></p:val></p:tav>'
            f'</p:tavLst></p:anim>')


def effect_xml(ids, spid, eff, node, dur, delay, is_sp):
    """Returns (xml with {ID} placeholder, preset class)."""
    e = eff["effect"]
    extra = ""
    if e == "appear":
        cls, pid, sub, body = "entr", 1, 0, _vis(ids, spid, "visible")
    elif e in ("fade", "dissolve"):
        cls, pid, sub = "entr", 10 if e == "fade" else 9, 0
        body = _vis(ids, spid, "visible") + _filt(ids, spid, e, dur)
    elif e == "wipe":
        cls, pid = "entr", 22
        sub, flt = WIPE[eff.get("dir", "left")]
        body = _vis(ids, spid, "visible") + _filt(ids, spid, flt, dur)
    elif e == "ascend":
        cls, pid, sub = "entr", 42, 0
        body = (_vis(ids, spid, "visible") + _filt(ids, spid, "fade", dur)
                + _anim(ids, spid, "ppt_x", "#ppt_x", "#ppt_x", dur) + _anim(ids, spid, "ppt_y", "#ppt_y+.1", "#ppt_y", dur))
    elif e == "zoom":
        cls, pid, sub = "entr", 53, 16
        body = (_vis(ids, spid, "visible") + _anim(ids, spid, "ppt_w", 0, "#ppt_w", dur, True)
                + _anim(ids, spid, "ppt_h", 0, "#ppt_h", dur, True) + _filt(ids, spid, "fade", dur))
    elif e == "exit":
        cls, pid, sub = "exit", 10, 0
        body = _filt(ids, spid, "fade", dur, "out") + _vis(ids, spid, "hidden", max(dur - 1, 0))
    elif e == "wipeout":
        cls, pid = "exit", 22
        sub, flt = WIPE[eff.get("dir", "down")]
        body = _filt(ids, spid, flt, dur, "out") + _vis(ids, spid, "hidden", max(dur - 1, 0))
    elif e == "vanish":
        cls, pid, sub = "exit", 1, 0
        body = _vis(ids, spid, "hidden")
    elif e == "pulse":
        cls, pid, sub = "emph", 6, 0
        sc = int(eff.get("scale", 1.15) * 100000)
        body = (f'<p:animScale><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}" autoRev="1" fill="hold"/>{_tgt(spid)}</p:cBhvr>'
                f'<p:by x="{sc}" y="{sc}"/></p:animScale>')
        if eff.get("repeat"):
            extra = f' repeatCount="{int(eff["repeat"]) * 1000}"'
    elif e == "path":
        cls, pid, sub = "path", 0, 0
        pts = eff["path"]            # list of (dx, dy) in inches relative to start
        d = "M 0 0 " + " ".join(f"L {x / SLIDE_W:.4f} {y / SLIDE_H:.4f}" for x, y in pts) + " E"
        body = (f'<p:animMotion origin="layout" path="{d}" pathEditMode="relative" ptsTypes="{"A" * (len(pts) + 1)}">'
                f'<p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}" fill="hold"/>{_tgt(spid)}<p:attrNameLst>'
                f'<p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>')
    else:
        raise ValueError(e)
    grp = f' grpId="{GRP[cls]}"' if is_sp else ""
    return (f'<p:par><p:cTn id="{{ID}}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}"{extra} fill="hold"{grp} '
            f'nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>'), cls


def timing_xml(steps, shapes):
    ids = Ids()
    clicks = []          # list of click groups; each = list of sub-pars; each sub-par = (start_delay, [effects xml], end)
    built = []
    for st in steps:
        trig = st.get("trigger", "click")
        dur = int(st.get("dur", 500 if st.get("effect") != "ascend" else 800))
        delay = int(st.get("delay", 0))
        if trig == "click" or not clicks:
            clicks.append([[0, [], 0]])
            node = "clickEffect"
        elif trig == "after":
            prev = clicks[-1][-1]
            clicks[-1].append([prev[2], [], prev[2]])
            node = "afterEffect"
        else:
            node = "withEffect"
        sub = clicks[-1][-1]
        for i, name in enumerate(st["targets"]):
            if name not in shapes:
                raise KeyError(f"shape '{name}' not found")
            spid, is_sp = shapes[name]
            nd = node if i == 0 else "withEffect"
            sub[1].append((spid, st, nd, dur, delay, is_sp))
            sub[2] = max(sub[2], sub[0] + delay + dur)

    out = []
    for grp in clicks:
        g_id = ids.next()
        inner = []
        for start, effs, _end in grp:
            s_id = ids.next()
            ex = []
            for spid, st, nd, dur, delay, is_sp in effs:
                eid = ids.next()
                xml, cls = effect_xml(ids, spid, st, nd, dur, delay, is_sp)
                ex.append(xml.replace("{ID}", str(eid), 1))
                if is_sp and (spid, GRP[cls]) not in built:
                    built.append((spid, GRP[cls]))
            inner.append(f'<p:par><p:cTn id="{s_id}" fill="hold"><p:stCondLst><p:cond delay="{start}"/></p:stCondLst>'
                         f'<p:childTnLst>{"".join(ex)}</p:childTnLst></p:cTn></p:par>')
        out.append(f'<p:par><p:cTn id="{g_id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
                   f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')
    bld = "".join(f'<p:bldP spid="{s}" grpId="{g}" animBg="1"/>' for s, g in built)
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + "".join(out) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
            '</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
            '</p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            + (f"<p:bldLst>{bld}</p:bldLst>" if bld else "") + "</p:timing>")


def shape_map(xml):
    m = {}
    for kind in ("sp", "pic", "cxnSp"):
        for blk in re.finditer(rf"<p:{kind}>(.*?)</p:{kind}>", xml, re.S):
            c = re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"', blk.group(1))
            if c:
                m[c.group(2)] = (c.group(1), kind == "sp")
    return m


def process(src, spec_path, dst):
    spec = json.load(open(spec_path))
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", item.filename)
        if m:
            i = int(m.group(1)) - 1
            if i < len(spec):
                xml = data.decode("utf8")
                sp = spec[i]
                tr = {"fade": '<p:transition spd="med"><p:fade/></p:transition>',
                      "push": '<p:transition spd="med"><p:push dir="u"/></p:transition>',
                      "morph": ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                                '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
                                '<p:transition spd="slow"><p159:morph option="byObject"/></p:transition></mc:Choice>'
                                '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'),
                      "none": ""}.get(sp.get("transition", "fade"), "")
                tm = timing_xml(sp["steps"], shape_map(xml)) if sp.get("steps") else ""
                xml = re.sub(r"<mc:AlternateContent.*?</mc:AlternateContent>|<p:transition.*?</p:transition>|<p:timing>.*?</p:timing>", "", xml, flags=re.S)
                xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + tr + tm, 1)
                data = xml.encode("utf8")
        zout.writestr(item, data)
    zout.close()


if __name__ == "__main__":
    process(sys.argv[1], sys.argv[2], sys.argv[3])
