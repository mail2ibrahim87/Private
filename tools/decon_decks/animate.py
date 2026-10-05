"""Add native PowerPoint transitions and entrance animations to a pptxgenjs deck.

spec.json: [{"transition": "fade"|"push", "steps": [{"trigger": "click"|"with"|"after",
            "effect": "appear"|"fade"|"ascend"|"wipe"|"dissolve", "dir": "left|right|up|down",
            "dur": ms, "delay": ms, "targets": [objectName, ...]}]}, ...]   (one entry per slide)
Effects mirror the reference deck: Appear (1), Fade (10), Ascend (42), Wipe (22), Dissolve (9).
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


def effect_xml(ids, spid, eff, node, dur, delay, is_sp):
    grp = ' grpId="0"' if is_sp else ""
    show = (f'<p:set><p:cBhvr><p:cTn id="{ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>')
    e = eff["effect"]
    if e == "appear":
        pid, sub, body = 1, 0, show
    elif e == "fade":
        pid, sub = 10, 0
        body = show + (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}"/>'
                       f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
    elif e == "dissolve":
        pid, sub = 9, 0
        body = show + (f'<p:animEffect transition="in" filter="dissolve"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}"/>'
                       f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
    elif e == "wipe":
        sub, flt = WIPE[eff.get("dir", "left")]
        pid = 22
        body = show + (f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}"/>'
                       f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
    elif e == "ascend":
        pid, sub = 42, 0
        body = show + (f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}"/>'
                       f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
        for attr, a, b in (("ppt_x", "#ppt_x", "#ppt_x"), ("ppt_y", "#ppt_y+.1", "#ppt_y")):
            body += (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{ids.next()}" dur="{dur}" fill="hold"/>'
                     f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>'
                     f'</p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="{a}"/></p:val></p:tav>'
                     f'<p:tav tm="100000"><p:val><p:strVal val="{b}"/></p:val></p:tav></p:tavLst></p:anim>')
    else:
        raise ValueError(e)
    return (f'<p:par><p:cTn id="{{ID}}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" fill="hold"{grp} '
            f'nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')


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
            if is_sp and spid not in built:
                built.append(spid)
    out = []
    for grp in clicks:
        g_id = ids.next()
        inner = []
        for start, effs, _end in grp:
            s_id = ids.next()
            ex = []
            for spid, st, nd, dur, delay, is_sp in effs:
                eid = ids.next()
                ex.append(effect_xml(ids, spid, st, nd, dur, delay, is_sp).replace("{ID}", str(eid), 1))
            inner.append(f'<p:par><p:cTn id="{s_id}" fill="hold"><p:stCondLst><p:cond delay="{start}"/></p:stCondLst>'
                         f'<p:childTnLst>{"".join(ex)}</p:childTnLst></p:cTn></p:par>')
        out.append(f'<p:par><p:cTn id="{g_id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
                   f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')
    bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in built)
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
                      "push": '<p:transition spd="med"><p:push dir="u"/></p:transition>'}.get(sp.get("transition", "fade"), "")
                tm = timing_xml(sp["steps"], shape_map(xml)) if sp.get("steps") else ""
                xml = re.sub(r"<p:transition.*?</p:transition>|<p:timing>.*?</p:timing>", "", xml, flags=re.S)
                xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + tr + tm, 1)
                data = xml.encode("utf8")
        zout.writestr(item, data)
    zout.close()


if __name__ == "__main__":
    process(sys.argv[1], sys.argv[2], sys.argv[3])
