#!/usr/bin/env python3
"""Atualiza o PPT da Simulação da Reforma Tributária com os números do motor (mesmo do HTML).
uso: python3 atualizar_ppt.py entrada.pptx saida.pptx [cenario.json] [--sem-lote6]
Atualiza nos slides 5-15 (empresas + consolidado): Receita Bruta, Impostos, Receita Líquida, OPEX, Lucro Bruto,
SG&A/EBITDA, EBIT, Resultado Financeiro, LAIR, IR, Lucro Líquido, % Margem e as colunas Δ e Δ%.
Slides 16-20 (fluxo mensal, movimentação, comparativos 2026/2027, fluxo por empresa) também são recalculados (flow_cli.js + flow.json). Não altera: linhas de Créd. Tributários / '↳'."""
import sys,os,json,subprocess
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
HERE=os.path.dirname(os.path.abspath(__file__))
args=[a for a in sys.argv[1:] if not a.startswith('--')]
inp,outp=args[0],args[1]; scen=args[2] if len(args)>2 else None
L6='--sem-lote6' not in sys.argv
cmd=['node',os.path.join(HERE,'calc_cli.js'),os.path.join(HERE,'bases.json'),os.path.join(HERE,'motor.js')]+([scen] if scen else [])
R=json.loads(subprocess.check_output(cmd))
# linha 'EBIT' do deck (v17): espelha o EBITDA em todas as empresas, exceto Lote 6 (EBIT real) -> consolidado soma as linhas
for _k,_ys in R.items():
    for _y,_r in _ys.items():
        _r['atual']['CT']=(_r['atual'].get('CRED') or 0); _r['rt']['CT']=_r['rt']['OPEX']-_r['atual']['OPEX']
        for _s in ('atual','rt'): _r[_s]['EBITX']=_r[_s]['EBIT' if _k=='lote6' else 'EBITDA']
ORDER=['graos','mobi','leste','mercosul','norte','sul','mcp01','welfare','rota','lote6']
def fm(v,d):
    if abs(v)<0.5*10**-d: v=0.0
    s=f"{abs(v):,.{d}f}".replace(',','X').replace('.',',').replace('X','.')
    return f"({s})" if v<0 else s
def _n(v):
    v=round(v*100,1)
    if v==0: v=0.0
    return f"{v:,.1f}".replace(',','X').replace('.',',').replace('X','.')
def pct(v): return _n(v)+'%'
def pp(v): return _n(v)+' p.p.'
FLD={'Receita Bruta':'RB','(-) Impostos':'DED','Receita Líquida':'RL','OPEX':'OPEX','Lucro Bruto':'LB','EBIT':'EBIT','LAIR':'EBT','(-) IR':'IR','Lucro Líquido':'LL','EBITDA':'EBITDA','(+) Créd. Tributários':'CRED','↳ Créd. Atual → RT¹':'CT'}
def arrow(a,t_):
    if abs(a)<0.5 and abs(t_)<0.5: return '= 0%'
    if abs(t_-a)<1e-6: return '= 0%'
    if abs(a)<0.5: return '▲ novo' if t_>0 else '▼ novo'
    v=(t_-a)/abs(a); return ('▲ ' if v>0 else '▼ ')+f"{round(abs(v)*100)}%"
def val(r,lab):
    a,t=r['atual'],r['rt']
    if lab=='Resultado Financeiro': return a['FIN'],t['FIN']
    f='EBITX' if lab=='EBIT' else FLD[lab]
    return a[f],t[f]
def cons(y):
    A={};T={}
    for k in ORDER:
        if k=='lote6' and not L6: continue
        r=R[k].get(str(y))
        if not r: continue
        for f in r['atual']: A[f]=A.get(f,0)+r['atual'][f]; T[f]=T.get(f,0)+r['rt'][f]
    return {'atual':A,'rt':T} if A else None
SMALL=False
def setv(shape,txt,color=None):
    tf=shape.text_frame; p=tf.paragraphs[0]; rs=p.runs
    if not rs: return
    rs[0].text=txt
    for r in rs[1:]: r.text=''
    if SMALL: rs[0].font.size=Pt(7)
    if color: rs[0].font.color.rgb=RGBColor.from_string(color)
RED,GRN='C00000','70AD47'
prs=Presentation(inp); BL=[1323464+i*1760177 for i in range(6)]; n=0
for si in range(4,15):
    s=prs.slides[si]; cons_slide=(si==14); k=None if cons_slide else ORDER[si-4]
    d=0; sc=1.0
    sh=[x for x in s.shapes if x.has_text_frame]
    labs=sorted([(x.top,x.text_frame.text.strip()) for x in sh if x.left==329184 and x.text_frame.text.strip()])
    # rótulo efetivo de cada linha (margem herda a linha anterior)
    eff={};prev=None
    for top,t in labs:
        if t=='% Margem': eff[top]=('M',prev)
        else: eff[top]=('V',t); prev=t
    for x in sh:
        if x.top not in eff or eff[x.top][0]=='V' and eff[x.top][1] not in list(FLD)+['Resultado Financeiro'] : 
            if x.top not in eff or eff[x.top][0]=='V': continue
        kind,lab=eff[x.top]
        for bi,b in enumerate(BL):
            ci=round((x.left-(b))/422544)
            if ci not in (0,1,2,3) or abs(x.left-(b+ci*422544))>3: continue
            y=2026+bi; r=cons(y) if cons_slide else R[k].get(str(y))
            if not r: continue
            SMALL=(cons_slide and ci<3)
            if kind=='V':
                a,t=val(r,lab); a/=sc; t/=sc; dl=t-a
                if ci==0: setv(x,fm(a,d))
                elif ci==1: setv(x,fm(t,d))
                elif ci==2: setv(x,fm(dl,d),RED if dl<-1e-9 else GRN if dl>1e-9 else None)
                elif lab.startswith('↳'): setv(x,arrow(a*sc,t*sc),RED if dl<-1e-9 else GRN if dl>1e-9 else None)
                else:
                    if abs(a)<1e-9: setv(x,'–')
                    else: setv(x,pct(dl/abs(a)),RED if dl<-1e-9 else GRN if dl>1e-9 else None)
                n+=1
            else:
                base={'Lucro Bruto':'LB','EBITDA':'EBITDA','Lucro Líquido':'LL','EBIT':'EBITX','LAIR':'EBT'}.get(lab)
                if not base: continue
                ma=r['atual'][base]/r['atual']['RL'] if r['atual']['RL'] else 0; mt=r['rt'][base]/r['rt']['RL'] if r['rt']['RL'] else 0
                if ci==0: setv(x,pct(ma))
                elif ci==1: setv(x,pct(mt))
                elif ci==2: setv(x,pp(mt-ma),GRN if mt-ma>=0 else RED)
                n+=1

# bullets "Δ Receita Bruta / Δ EBITDA / Δ Lucro Líquido" do slide 15 (um bloco por ano, 2026-2031), refeitos com os números do motor
_bl=sorted([x for x in prs.slides[14].shapes if x.has_text_frame and x.text_frame.text.strip().startswith('• Δ Receita Bruta')],key=lambda x:x.left)
for _x,_y in zip(_bl,range(2026,2032)):
    _c=cons(_y)
    if not _c: continue
    _t=[]
    for _lab,_f in (('Receita Bruta','RB'),('EBITDA','EBITDA'),('Lucro Líquido','LL')):
        _a,_r=_c['atual'][_f],_c['rt'][_f]; _d=_r-_a
        _t.append(f"• Δ {_lab}: {fm(_d,0)} mil ({pct(_d/abs(_a)) if _a else '0,0%'})")
    for _pa,_tx in zip(_x.text_frame.paragraphs,_t):
        if _pa.runs:
            _pa.runs[0].text=_tx
            for _r in _pa.runs[1:]: _r.text=''

# ===== slides 16-20: fluxo mensal 2027, movimentação, comparativos 2026/2027 e fluxo por empresa (mesmo motor do HTML) =====
from pptx.chart.data import CategoryChartData
FLW=json.loads(subprocess.check_output(['node',os.path.join(HERE,'flow_cli.js'),os.path.join(HERE,'bases.json'),os.path.join(HERE,'motor.js'),os.path.join(HERE,'flow.json')]+([scen] if scen else [])))
NAMES={'graos':'Grãos do Piauí','mobi':'CS Mobi Cuiabá','leste':'Bloco Leste','mercosul':'Rodovias Mercosul','norte':'Escolas Norte','sul':'Escolas Sul','mcp01':'MCP01','welfare':'Welfare Ambiental','rota':'Rota da Integração','lote6':'Lote 6'}
KS=[k for k in ORDER if (k!='lote6' or L6)]
KF=['deb','opex','capex','mov','des','res']
def ssum(a): return sum(a)
fcons={x:[sum(FLW[k]['rt'][x][m] for k in KS if k in FLW) for m in range(12)] for x in KF}
fat={x:sum(FLW[k]['at'][x] for k in KS if k in FLW) for x in KF}
def sc(tb,r,c,txt,color=None):
    try: cl=tb.cell(r,c)
    except Exception: return
    tf=cl.text_frame; ps=tf.paragraphs
    if not ps or not ps[0].runs: return
    ps[0].runs[0].text=txt
    for x in ps[0].runs[1:]: x.text=''
    if color: ps[0].runs[0].font.color.rgb=RGBColor.from_string(color)
def col(v): return RED if v<-1e-9 else GRN if v>1e-9 else None
def tbl(si,name): return [x for x in prs.slides[si].shapes if x.has_table and x.name==name][0].table
def bullets(si,start,texts):
    for x in prs.slides[si].shapes:
        if x.has_text_frame and x.text_frame.text.startswith(start):
            for pa,tx in zip(x.text_frame.paragraphs,texts):
                if pa.runs:
                    pa.runs[0].text=tx
                    for r_ in pa.runs[1:]: r_.text=''
            return
def setchart(si,name,cats,series):
    ch=[x for x in prs.slides[si].shapes if x.name==name][0].chart
    cd=CategoryChartData(); cd.categories=cats
    for nm,v in series: cd.add_series(nm,v)
    ch.replace_data(cd)
    if name=='Chart 513':
        try:
            ch.chart_title.text_frame.text='Atual × RT (R$ mil)'
            for pl in ch.plots:
                pl.data_labels.number_format='#,##0;(#,##0)'; pl.data_labels.number_format_is_linked=False
            for se in ch.plots[0].series:
                dl=se.data_labels; dl.show_value=True; dl.number_format='#,##0;(#,##0)'; dl.number_format_is_linked=False
            for r_ in ch.chart_title.text_frame.paragraphs[0].runs: r_.font.size=Pt(11); r_.font.bold=True
            ch.value_axis.tick_labels.number_format='#,##0;(#,##0)'; ch.value_axis.tick_labels.number_format_is_linked=False
        except Exception as e: print('chart fmt',e)
# slide 16
t=tbl(15,'Table 516')
for m in range(12): sc(t,1,1+m,'0')
sc(t,1,13,'0')
for k_,ri in (('deb',2),('opex',3),('capex',4),('mov',5),('des',6),('res',7)):
    for m in range(12): sc(t,ri,1+m,fm(fcons[k_][m],0))
    sc(t,ri,13,fm(ssum(fcons[k_]),0))
t2=tbl(15,'Table 517')
for i,k in enumerate(KS):
    sc(t2,1+i,0,NAMES[k])
    for j,x in enumerate(KF): sc(t2,1+i,1+j,fm(ssum(FLW[k]['rt'][x]),0))
for j,x in enumerate(KF): sc(t2,1+len(KS),1+j,fm(ssum(fcons[x]),0))
setchart(15,'Chart 514',list(FLW['graos']['rt']['deb'] and ['Jan/27','Fev/27','Mar/27','Abr/27','Mai/27','Jun/27','Jul/27','Ago/27','Set/27','Out/27','Nov/27','Dez/27']),[('Movimentação mensal (consolidado)',[round(v) for v in fcons['mov']])])
# slide 17
A=[fat[x] for x in KF]; T=[ssum(fcons[x]) for x in KF]
t=tbl(16,'Table 511')
d1=lambda v:'–' if abs(v)<0.5 else fm(v,0)
for i in range(6):
    sc(t,1+i,1,d1(A[i])); sc(t,1+i,2,d1(T[i])); sc(t,1+i,3,fm(T[i]-A[i],0),col(T[i]-A[i]))
setchart(16,'Chart 513',['Débitos','Créditos OPEX','Créditos CAPEX','Movimentação','Desembolso','Restituição'],[('Modelo atual',[round(v) for v in A]),('RT',[round(v) for v in T])])
big='créditos de CAPEX' if T[2]>T[1] else 'créditos de OPEX/SG&A'
m1=lambda v:('-' if v<-0.5 else '')+fm(abs(v),0)
bullets(16,'• Modelo atual',[f"• Modelo atual: débitos de {m1(A[0])} mil e créditos de {m1(A[1])} mil resultam em movimentação de {m1(A[3])} mil, integralmente desembolso de caixa.",
 f"• Reforma Tributária: débitos de {m1(T[0])} mil, compensados por créditos de OPEX/SG&A ({m1(T[1])} mil) e de CAPEX ({m1(T[2])} mil), geram movimentação líquida de {m1(T[3])} mil.",
 f"• Impacto no caixa (por CNPJ): desembolso de {m1(T[4])} mil e restituição de {m1(T[5])} mil; "+("a movimentação líquida positiva não elimina o desembolso, porque o crédito de uma empresa não paga o débito de outra." if T[3]>0 else "o crédito de uma empresa não paga o débito de outra."),
 f"• Variação da movimentação: {m1(T[3]-A[3])} mil, puxada pelos {big}."])
# slides 18-19
def nsp(a,t_):
    if abs(a)<1e-9: return '–',None
    if (a>0 and t_<0) or (a<0 and t_>0): return 'n.s.',None
    return pct((t_-a)/abs(a)),col(t_-a)
def arrow(a,t_):
    if abs(a)<0.5 and abs(t_)<0.5: return '= 0%'
    if abs(t_-a)<1e-6: return '= 0%'
    if abs(a)<0.5: return '▲ novo' if t_>0 else '▼ novo'
    v=(t_-a)/abs(a); return ('▲ ' if v>0 else '▼ ')+f"{round(abs(v)*100)}%"
def rate(v):
    return (f"{v*100:.4f}".rstrip('0').rstrip('.')).replace('.',',')+'%'
cbsrate=lambda y: rate(json.load(open(scen))['params']['cbs'][str(y)]) if scen else '11,3468%'
for si,y in ((17,2026),(18,2027)):
    t=tbl(si,'Table 511'); rows=[]
    for i,k in enumerate(KS):
        r=R[k].get(str(y)); rows.append((NAMES[k],k,r))
    tot=cons(y)
    # (a lista da tabela tem sempre as 10 linhas; sem Lote 6 a linha fica '—')
    allrows=[(NAMES[k],k,R[k].get(str(y)) if (k!='lote6' or L6) else None) for k in ORDER]
    for i,(nm,k,r) in enumerate(allrows+[('Total '+str(sum(1 for x in allrows if x[2]) if y==2027 else len(KS)-(1 if L6 else 0))+' empresas','tot',tot)]):
        ri=2+i; sc(t,ri,0,nm)
        if not r:
            for c in range(1,len(t.columns)): sc(t,ri,c,'—')
            continue
        a,b=r['atual'],r['rt']; c=[1]
        def put(x,color=None): sc(t,ri,c[0],x,color); c[0]+=1
        def pc(a_,b_): s_,cl=nsp(a_,b_); put(s_,cl)
        if y==2027: put(fm(a['RB'],0)); put(fm(b['RB'],0)); pc(a['RB'],b['RB'])
        put(fm(a['RL'],0)); put(fm(b['RL'],0)); pc(a['RL'],b['RL'])
        put(fm(a['CT'],0)); put(fm(b['CT'],0)); put(arrow(a['CT'],b['CT']),col(b['CT']-a['CT']))
        put(fm(a['EBITDA'],0)); put(fm(b['EBITDA'],0)); put(fm(b['EBITDA']-a['EBITDA'],0),col(b['EBITDA']-a['EBITDA'])); pc(a['EBITDA'],b['EBITDA'])
        put(fm(a['LL'],0)); put(fm(b['LL'],0)); put(fm(b['LL']-a['LL'],0),col(b['LL']-a['LL'])); pc(a['LL'],b['LL'])
    Tt=tot; dLL=Tt['rt']['LL']-Tt['atual']['LL']; dE=Tt['rt']['EBITDA']-Tt['atual']['EBITDA']
    rk=sorted([(R[k][str(y)]['rt']['LL']-R[k][str(y)]['atual']['LL'],k) for k in KS if R[k].get(str(y))])
    n_=3 if (y==2027 and L6) else 2
    ls=rk[:n_]; sl=sum(d for d,_ in ls); cb=cbsrate(y)
    ab=lambda v: fm(abs(v),0)
    nm_list=lambda ks_: ' e '.join([', '.join(NAMES[x] for x in ks_[:-1]),NAMES[ks_[-1]]]) if len(ks_)>1 else NAMES[ks_[0]]
    kk=[k for _,k in ls]; special=all(k in ('graos','rota','lote6') for k in kk)
    ct0,ct1=Tt['atual']['CT'],Tt['rt']['CT']; qtd=len(KS) if y==2027 else len(KS)-(1 if L6 else 0)
    b0=f"• Em {y}, o lucro líquido somado {'das '+str(qtd)+' empresas ' if y==2026 else 'das '+str(qtd)+' empresas '}{'cai' if dLL<0 else 'sobe'} {ab(dLL)} mil ({pct(dLL/abs(Tt['atual']['LL']))}) no RT, e o EBITDA {'cai' if dE<0 else 'sobe'} {ab(dE)} mil ({pct(dE/abs(Tt['atual']['EBITDA']))})."
    if y==2027:
        if L6:
            ex=cons_nol6=None
            sA={f:sum(R[k][str(y)]['atual'][f] for k in KS if k!='lote6' and R[k].get(str(y))) for f in ('LL','EBITDA')}
            sT={f:sum(R[k][str(y)]['rt'][f] for k in KS if k!='lote6' and R[k].get(str(y))) for f in ('LL','EBITDA')}
            b0+=f" Sem a Lote 6 as quedas eram de {ab(sT['LL']-sA['LL'])} mil e {ab(sT['EBITDA']-sA['EBITDA'])} mil."
        b0+=" É o primeiro ano com a CBS plena e o PIS/COFINS extinto."
    if y==2026:
        oth=[k for k in KS if R[k].get('2026') and k not in kk]; so=sum(R[k]['2026']['rt']['LL']-R[k]['2026']['atual']['LL'] for k in oth)
        bl=[b0,
         f"• {nm_list(kk)} {'perdem' if sl<0 else 'variam'} {ab(sl)} mil{', mais que a perda total' if sl<dLL else ''}: "+(f"as duas saem do PIS/COFINS cumulativo (3,65%) para CBS de {cb}." if special else "maior pressão da carga tributária no RT."),
         f"• O crédito tributário {'sobe' if ct1>=ct0 else 'cai'} de {ab(ct0)} mil para {ab(ct1)} mil ({arrow(ct0,ct1)}), mas não compensa a alíquota maior nessas duas operações.",
         f"• As outras {len(oth)} empresas com operação {'melhoram levemente (+' if so>=0 else 'pioram ('}{ab(so)} mil somados), pelo crédito sobre OPEX. MCP01 não opera em 2026."]
    else:
        dRB=Tt['rt']['RB']-Tt['atual']['RB']; dRL=Tt['rt']['RL']-Tt['atual']['RL']
        det=' ('+', '.join(f"{NAMES[k]} (–{ab(d)} mil)" for d,k in ls)+')' if False else ''
        lst=', '.join(f"{NAMES[k]} (–{ab(d)} mil)" for d,k in ls[:-1])+' e '+f"{NAMES[ls[-1][1]]} (–{ab(ls[-1][0])} mil)"
        bl=[b0,
         f"• A receita bruta {'cai' if dRB<0 else 'sobe'} {ab(dRB)} mil ({pct(dRB/abs(Tt['atual']['RB']))}), mas a receita líquida {'cai' if dRL<0 else 'sobe'} só {ab(dRL)} mil ({pct(dRL/abs(Tt['atual']['RL']))}): a maior parte da queda na receita bruta é a CBS saindo do preço (\"por fora\").",
         f"• {lst} perdem juntas {ab(sl)} mil{', mais que a perda total' if sl<dLL else ''}: "+(f"as {'três' if n_==3 else 'duas'} saem do PIS/COFINS cumulativo (3,65%) para CBS de {cb}." if special else "maior pressão da carga tributária no RT."),
         f"• O crédito tributário {'sobe' if ct1>=ct0 else 'cai'} de {ab(ct0)} mil para {ab(ct1)} mil ({arrow(ct0,ct1)}), inclusive em Grãos, Rota e Lote 6, que hoje não têm crédito; ainda assim, o ganho não compensa o aumento de alíquota nessas {'três' if n_==3 else 'duas'} operações."]
    bullets(si,'• Em 20',bl)
# slide 20
t=tbl(19,'Table 511'); rows=[]
for k in ORDER:
    if k=='lote6' and not L6: continue
    f=FLW[k]; rows.append((NAMES[k],k,f['at'],{x:ssum(f['rt'][x]) for x in KF}))
rows.append(('Total '+str(len(rows))+' empresas','tot',fat,{x:ssum(fcons[x]) for x in KF}))
for i,(nm,k,a,r) in enumerate(rows):
    ri=2+i; sc(t,ri,0,nm); c=[1]
    def put(x,color=None): sc(t,ri,c[0],x,color); c[0]+=1
    put(fm(a['deb'],0)); put(fm(r['deb'],0)); put(fm(r['deb']-a['deb'],0),col(r['deb']-a['deb']))
    put(fm(a['opex'],0)); put(fm(r['opex'],0)); put(fm(r['capex'],0))
    put(fm(a['mov'],0)); put(fm(r['mov'],0)); put(fm(r['mov']-a['mov'],0),col(r['mov']-a['mov']))
    put(fm(a['des'],0)); put(fm(r['des'],0)); put(fm(r['res'],0)); put(fm(r['des']-a['des'],0),col(r['des']-a['des']))
body=rows[:-1]; Tt=rows[-1]; mm=lambda v: ('-' if v<-0.5 else '')+fm(abs(v),0); N_=len(body)
nd=sum(1 for _,_,a,r in body if r['des']<-0.5); nr=sum(1 for _,_,a,r in body if r['res']>0.5)
both=[nm for nm,_,a,r in body if r['des']<-0.5 and r['res']>0.5]
gl=sorted([(r['des']-a['des'],nm,a,r) for nm,_,a,r in body],key=lambda z:z[0])
g=gl[0] if gl[0][0]<-0.5 else sorted(gl,key=lambda z:-abs(z[0]))[0]
rs=[(nm,r) for nm,_,a,r in body if r['res']>0.5 and r['des']>-0.5]
lst=(', '.join(x[0] for x in rs[:-1])+' e '+rs[-1][0]) if len(rs)>1 else (rs[0][0] if rs else '')
b=[f"• Modelo atual: as {N_} empresas desembolsam juntas {mm(Tt[2]['des'])} mil. No RT, o desembolso soma {mm(Tt[3]['des'])} mil em {nd} empresas e a restituição soma {mm(Tt[3]['res'])} mil em {nr}"+(f" ({' e '.join(both)} {'têm' if len(both)>1 else 'tem'} os dois ao longo do ano)" if both else "")+f". A movimentação líquida de {mm(Tt[3]['mov'])} mil esconde essa assimetria entre CNPJs.",
 f"• {g[1]} concentra o {'aumento' if g[0]<0 else 'alívio'} de saída de caixa: o desembolso vai de {mm(g[2]['des'])} mil para {mm(g[3]['des'])} mil, {mm(abs(g[0]))} mil {'a mais' if g[0]<0 else 'a menos'}"+(f", mesmo com {mm(g[3]['capex'])} mil de crédito de CAPEX no fluxo de 2027" if g[3]['capex']>0.5 else ", sem crédito de CAPEX no fluxo de 2027")+".",
 (f"• {lst} {'passam' if len(rs)>1 else 'passa'} a ter restituição líquida ({mm(sum(r['res'] for _,r in rs))} mil), puxada pelos créditos de CAPEX ({mm(sum(r['capex'] for _,r in rs))} mil). Esse caixa depende do prazo de ressarcimento." if rs else "• Nenhuma empresa fica com restituição líquida. O caixa de créditos depende do prazo de ressarcimento.")]
if L6:
    l6=[x for x in body if x[1]=='lote6'][0]
    b.append(f"• Lote 6: o desembolso vai de {mm(l6[2]['des'])} mil para {mm(l6[3]['des'])} mil, {mm(abs(l6[3]['des']-l6[2]['des']))} mil {'a mais' if l6[3]['des']<l6[2]['des'] else 'a menos'}: a CBS de {cbsrate(2027)} substitui o PIS/COFINS de 3,65%, com crédito de OPEX/SG&A de apenas {mm(l6[3]['opex'])} mil e sem CAPEX.")
bullets(19,'• Modelo atual',b)


# textos fixos com a alíquota de CBS antiga (slides 3, 4, 14, 16): troca pela CBS 2027 do cenário
_cb=cbsrate(2027)
for _s in prs.slides:
    for _x in _s.shapes:
        if _x.has_text_frame:
            for _pa in _x.text_frame.paragraphs:
                for _r in _pa.runs:
                    if 'CBS 9,25%' in _r.text or 'CBS de 9,25%' in _r.text:
                        _r.text=_r.text.replace('CBS 9,25%','CBS '+_cb).replace('CBS de 9,25%','CBS de '+_cb)

_FX=[('Fonte: planilha Aliquotas_Fluxo (aba Fluxo) e DF - Mensal das empresas, 2027 (Grãos e MCP01 atualizadas em 25/10).','Fonte: DF - Mensal 2027 de cada empresa (planilhas de origem), recalculada com as premissas da simulação.'),
 ('Créditos de CAPEX conforme a planilha (zerados no fluxo mensal de Grãos do Piauí, MCP01 e Rota da Integração).','Créditos de CAPEX conforme a DF - Mensal de cada empresa (planilhas de origem); desconsiderados para Grãos do Piauí, Rodovias Mercosul, MCP01 e Rota da Integração.'),
 ('Créditos de CAPEX conforme a DF - Mensal de cada empresa (planilhas de origem).','Créditos de CAPEX conforme a DF - Mensal de cada empresa (planilhas de origem); desconsiderados para Grãos do Piauí, Rodovias Mercosul, MCP01 e Rota da Integração.'),
 ('RT: aba Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),'),('RT: planilha Fluxo 2027,','RT: fluxo 2027 (DF - Mensal),')]
for _s in prs.slides:
    for _x in _s.shapes:
        if _x.has_text_frame:
            for _pa in _x.text_frame.paragraphs:
                for _r in _pa.runs:
                    for _a,_b in _FX:
                        if _a in _r.text: _r.text=_r.text.replace(_a,_b)

def _fixunits(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            if 'milhões' in r.text: r.text=r.text.replace('R$ milhões','R$ mil')
            if 'Os débitos do modelo atual incluem todas as deduções da DRE.' in r.text: r.text=r.text.replace('Os débitos do modelo atual incluem todas as deduções da DRE.','Os débitos (Atual e RT) incluem todas as deduções da DRE, inclusive o IIBB da Mercosul.')
for _s in prs.slides:
    for _sh in _s.shapes:
        if _sh.has_text_frame: _fixunits(_sh.text_frame)
prs.save(outp); print('células atualizadas:',n,'->',outp)
