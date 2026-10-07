import re
E=914400
def blocks(x, tag):
    # top-level-ish elements of spTree: p:sp, p:cxnSp, p:pic
    out=[]
    for m in re.finditer(r'<(p:sp|p:cxnSp|p:pic)>', x):
        tg=m.group(1); end=x.find('</%s>'%tg, m.start())
        out.append((m.start(), end+len(tg)+3))
    return out
def shapes(x):
    res={}
    for s,e in blocks(x,None):
        b=x[s:e]
        m=re.search(r'<p:cNvPr id="(\d+)" name="([^"]*)"',b)
        if not m: continue
        off=re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"',b)
        g=[int(v)/E for v in off.groups()] if off else None
        prst=re.search(r'prst="(\w+)"',b)
        fl=re.search(r'<a:xfrm([^>]*)>',b)
        txt=''.join(re.findall(r'<a:t>([^<]*)</a:t>',b))
        fill=re.search(r'<p:spPr>.*?<a:solidFill><a:srgbClr val="(\w+)"',b)
        res[int(m.group(1))]=dict(name=m.group(2),g=g,prst=prst.group(1) if prst else '',txt=txt,flip=fl.group(1) if fl else '',fill=fill.group(1) if fill else '',span=(s,e))
    return res
def dump(path):
    x=open(path).read()
    for i,d in shapes(x).items():
        g=d['g']; gs=' '.join('%.2f'%v for v in g) if g else ''
        print(i,d['name'],d['prst'],gs,d['flip'].strip(),d['fill'],repr(d['txt'][:40]))
if __name__=='__main__':
    import sys; dump(sys.argv[1])
