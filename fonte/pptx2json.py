import json,re,sys
from pptx import Presentation
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
E=9525
def px(v):return round(v/E,2)
def clr(el):
    if el is None:return None
    s=el.find(A+'srgbClr')
    if s is not None:
        a=s.find(A+'alpha')
        return '#'+s.get('val')+('' if a is None else '%02X'%round(int(a.get('val'))/100000*255))
    return None
def fillof(spPr):
    if spPr is None:return None
    sf=spPr.find(A+'solidFill')
    if sf is not None:return clr(sf)
    return None
def lineof(spPr):
    ln=spPr.find(A+'ln') if spPr is not None else None
    if ln is None:return None
    if ln.find(A+'noFill') is not None:return None
    sf=ln.find(A+'solidFill')
    if sf is None:return None
    return [clr(sf),round(int(ln.get('w','12700'))/E,2),(ln.find(A+'prstDash').get('val') if ln.find(A+'prstDash') is not None else 'solid')]
def paras(tx):
    out=[]
    if tx is None:return out
    for p in tx.findall(A+'p'):
        pPr=p.find(A+'pPr');al='l'
        bu=None;marL=0;ind=0;lnsp=None;spb=0
        if pPr is not None:
            al=pPr.get('algn','l');marL=int(pPr.get('marL','0'))/E;ind=int(pPr.get('indent','0'))/E
            b=pPr.find(A+'buChar')
            if b is not None:bu=b.get('char')
            ls=pPr.find(A+'lnSpc')
            if ls is not None and ls.find(A+'spcPct') is not None:lnsp=int(ls.find(A+'spcPct').get('val'))/100000
            sb=pPr.find(A+'spcBef')
            if sb is not None and sb.find(A+'spcPts') is not None:spb=int(sb.find(A+'spcPts').get('val'))/100*96/72
        runs=[]
        for r in p:
            if r.tag==A+'r':
                rp=r.find(A+'rPr');t=r.find(A+'t').text or ''
                d={'t':t}
                if rp is not None:
                    if rp.get('sz'):d['sz']=int(rp.get('sz'))/100
                    if rp.get('b')=='1':d['b']=1
                    if rp.get('i')=='1':d['i']=1
                    c=clr(rp.find(A+'solidFill'))
                    if c:d['c']=c
                runs.append(d)
            elif r.tag==A+'br':runs.append({'t':'\n'})
        epr=p.find(A+'endParaRPr');esz=int(epr.get('sz'))/100 if epr is not None and epr.get('sz') else None
        out.append({'al':al,'runs':runs,'bu':bu,'ml':marL,'ind':ind,'ls':lnsp,'sb':spb,'esz':esz})
    return out
def body(tx):
    bp=tx.find(A+'bodyPr') if tx is not None else None
    if bp is None:return {}
    return {'anc':bp.get('anchor','t'),'wrap':bp.get('wrap','square'),'ins':[px(int(bp.get(k,'0'))) if bp.get(k) is not None else 7.2 for k in ('lIns','tIns','rIns','bIns')]}
def shape(sh):
    el=sh._element;spPr=el.find(P+'spPr')
    xf=spPr.find(A+'xfrm');off=xf.find(A+'off');ext=xf.find(A+'ext')
    d={'x':px(int(off.get('x'))),'y':px(int(off.get('y'))),'w':px(int(ext.get('cx'))),'h':px(int(ext.get('cy')))}
    g=spPr.find(A+'prstGeom');prst=g.get('prst') if g is not None else 'rect'
    d['g']=prst
    if prst=='roundRect':
        gd=g.find('.//'+A+'gd');d['adj']=int(gd.get('fmla').split()[1])/100000 if gd is not None else .1667
    f=fillof(spPr);
    if f:d['f']=f
    ln=lineof(spPr)
    if ln:d['ln']=ln
    if xf.get('flipV')=='1':d['flipV']=1
    if xf.get('flipH')=='1':d['flipH']=1
    tx=el.find(P+'txBody')
    ps=paras(tx)
    if any(r['t'] for p in ps for r in p['runs']):
        d['p']=ps;d.update(body(tx))
    d['n']=sh.name
    return d
def table(sh):
    tbl=sh._element.find('.//'+A+'tbl')
    cols=[px(int(c.get('w'))) for c in tbl.find(A+'tblGrid')]
    rows=[]
    for tr in tbl.findall(A+'tr'):
        cells=[]
        for tc in tr.findall(A+'tc'):
            tcPr=tc.find(A+'tcPr');c={}
            c['p']=paras(tc.find(A+'txBody'))
            if tc.get('gridSpan'):c['gs']=int(tc.get('gridSpan'))
            if tc.get('rowSpan'):c['rs']=int(tc.get('rowSpan'))
            if tc.get('hMerge'):c['hm']=1
            if tc.get('vMerge'):c['vm']=1
            if tcPr is not None:
                f=fillof(tcPr)
                if f:c['f']=f
                c['m']=[px(int(tcPr.get(k,'0'))) for k in ('marL','marT','marR','marB')]
                c['anc']=tcPr.get('anchor','t')
                for tag,key in (('lnL','bl'),('lnR','br'),('lnT','bt'),('lnB','bb')):
                    l=tcPr.find(A+tag)
                    if l is not None and l.find(A+'solidFill') is not None:
                        c[key]=[clr(l.find(A+'solidFill')),round(int(l.get('w','12700'))/E,2)]
            cells.append(c)
        rows.append({'h':px(int(tr.get('h'))),'c':cells})
    return {'x':px(sh.left),'y':px(sh.top),'w':px(sh.width),'h':px(sh.height),'cols':cols,'rows':rows,'n':sh.name}
prs=Presentation(sys.argv[1]);slides=[]
def bgof(slide):
    bg=slide._element.find('.//'+P+'bg')
    if bg is None:return None
    return fillof(bg.find(P+'bgPr'))
def add(sh,lst):
    if sh.shape_type==19 or sh.has_table if hasattr(sh,'has_table') else False:
        d=table(sh);d['tbl']=1;lst.append(d)
    elif getattr(sh,'has_chart',False) and sh.has_chart:
        lst.append({'chart':1,'x':px(sh.left),'y':px(sh.top),'w':px(sh.width),'h':px(sh.height),'n':sh.name})
    elif sh.shape_type==14:
        d=shape(sh) if sh._element.find(P+'spPr') is not None and sh._element.find(P+'spPr').find(A+'xfrm') is not None else None
        if d is None:
            lst.append({'x':px(sh.left),'y':px(sh.top),'w':px(sh.width),'h':px(sh.height),'slnum':1})
        else:
            d['slnum']=1;lst.append(d)
    else:
        lst.append(shape(sh))
for i,s in enumerate(prs.slides,1):
    el=[];lay=[]
    for sh in s.slide_layout.shapes:add(sh,lay)
    for sh in s.shapes:add(sh,el)
    for d in lay:d['lay']=1
    slides.append({'i':i,'bg':bgof(s) or '#FAFBFC','lay':[d for d in lay] if s.slide_layout.name!='DEFAULT' else [],'el':el})
json.dump(slides,open(sys.argv[2],'w'),ensure_ascii=False,separators=(',',':'))
print(len(json.dumps(slides)))
