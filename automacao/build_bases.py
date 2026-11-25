import sys,re,json,statistics; sys.path.insert(0,'/tmp/claude-0/s')
from lib import *
from extract import CFG,extract
import openpyxl
P="/mnt/user-data/uploads/05 - Reforma Tributaria_Crescento - Claude/"
NAMES={'graos':'Grãos do Piauí','mobi':'CS Mobi Cuiabá','leste':'Bloco Leste','mercosul':'Rodovias Mercosul','norte':'Escolas Norte','sul':'Escolas Sul','mcp01':'MCP01','welfare':'Welfare Ambiental','rota':'Rota da Integração','lote6':'Lote 6'}
FILES={'graos':'5700 - Grãos.xlsx','mobi':'5800 - CS Mobi.xlsx','leste':'5900 - Bloco Leste.xlsx','mercosul':'5910 - Mercosul.xlsx','norte':'5920 - Escolas Norte.xlsx','sul':'5930 - Escolas Sul.xlsx','mcp01':'5940 - MCP01.xlsx','welfare':'9904 - Welfare.xlsx','rota':'9905 - Rota da Integração.xlsx','lote6':'Lote6.xlsx (+ premissas 2027 do PPT)'}
UNIT={'leste':1000.0,'mercosul':1000.0,'welfare':1000.0}
ISS_PH0={2026:1,2027:1,2028:1,2029:.9,2030:.8,2031:.7,2032:.6,2033:0}
IBS_PH0={2026:0,2027:0,2028:0,2029:.1,2030:.2,2031:.3,2032:.4,2033:1}
CBS0=0.0925; IBS0=0.1875
LINES=['RB','DED','OPEX','CRED','SGA','DA','FIN','NONOP','IR']
LAB={'RB':'Receita Bruta Serviços','DED':'(-)Deduções','OPEX':'Opex','CRED':'(-)Creditos tributários','SGA':'SG&A','DA':'Depreciação e Amortização','FIN':'Resultado Financeiro','NONOP':'Receita Não Operacional','IR':'IR/CSLL'}
REF=re.compile(r"(?:'[^']*'!)?\$?([A-Z]{1,3})\$?(\d+)")
def num(x): return float(x) if isinstance(x,(int,float)) else 0.0
def refs_signed(f,col):
    if not isinstance(f,str) or not f.startswith('='): return []
    terms=re.findall(r'([+-]?)\s*((?:\'[^\']*\'!)?\$?[A-Z]{1,3}\$?\d+)',f[1:])
    out=[]
    for sg,tok in terms:
        m=REF.fullmatch(tok)
        if m and CI(m.group(1))==col: out.append((-1 if sg=='-' else 1,int(m.group(2))))
    return out
def leaves(gf,row,col,depth=0):
    f=gf.get((row,col))
    if isinstance(f,str) and f.startswith('=SUM(') and depth<4:
        inner=f[5:-1]; kids=[]
        for part in inner.split(','):
            m=re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)",part.strip())
            if m and CI(m.group(1))==col:
                for r in range(int(m.group(2)),int(m.group(4))+1): kids+=leaves(gf,r,col,depth+1)
            else:
                m=REF.fullmatch(part.strip())
                if m and CI(m.group(1))==col: kids+=leaves(gf,int(m.group(2)),col,depth+1)
        return kids
    return [(row,f)]
def regime_share(gf,gv,rows,col,sign_refs):
    tot=0.0;reg=0.0
    for s,j in sign_refs:
        for r,f in leaves(gf,j,col):
            v=abs(num(gv.get((r,col))))
            tot+=v
            if isinstance(f,str) and re.search(r'\$[A-Z]+\$\d+',f): reg+=v
    return (reg/tot) if tot>1e-9 else None
def build_company(k):
    d=load(k); gf=d['f']; gv=d['v']; rows,out=extract(k)
    u=UNIT.get(k,1.0)
    co={'nome':NAMES[k],'arquivo':FILES[k],'years':{}}
    # taxas IR
    for y,info in sorted(out.items()):
        pc=CI(info['rt_col']); ac=CI(info['atual_col'])
        Y={'atual':{}, 'rt':{}, 'comp':{}}
        for key in LINES:
            r=rows.get(LAB[key])
            if not r: Y['atual'][key]=0.0;Y['rt'][key]=0.0;Y['comp'][key]=None;continue
            va=num(gv.get((r,ac)))/u; vr=num(gv.get((r,pc)))/u
            Y['atual'][key]=va;Y['rt'][key]=vr
            if key in ('FIN','IR','CRED'): continue
            fr=gf.get((r,pc)); rs=refs_signed(fr,pc)
            sem=cbs=iss=ibs=0.0
            for s,j in rs:
                sem+=s*num(gv.get((j,pc-6)));cbs+=s*num(gv.get((j,pc-4)));iss+=s*num(gv.get((j,pc-3)));ibs+=s*num(gv.get((j,pc-2)))
            sem/=u;cbs/=u;iss/=u;ibs/=u
            if key=='DED':
                fixed=0.0
                if k=='mercosul': fixed=num(gv.get((22,pc)))/u   # IIBB: imposto fora do escopo da reforma
                Y['comp'][key]={'def':vr,'fixed':fixed}; continue
            res=vr-(sem+cbs+ibs)-(0 if key!='RB' else 0)
            # RB: iss col in revenue rows is excluded from POS (goes to DED) -> não somar
            c={'sem':sem,'cbs':cbs,'ibs':ibs,'res':res}
            c['regC']=regime_share(gf,gv,rows,pc-4,rs) 
            c['regI']=regime_share(gf,gv,rows,pc-2,rs)
            Y['comp'][key]=c
        # IR
        r=rows['IR/CSLL']
        frt=str(gf.get((r,pc))); fa=str(gf.get((r,ac)))
        m=re.search(r'\*\s*(\d+(?:\.\d+)?)%',frt) or re.search(r'\*\s*(0\.\d+)\b',frt)
        rate=None
        if m:
            rate=float(m.group(1)); rate=rate/100 if rate>1 else rate
        flat_a = bool(re.search(r'^=-[A-Z]+\d+\*',fa))
        Y['ir_rate']=rate; Y['ir_atual_flat']=flat_a
        co['years'][str(y)]=Y
    return co
def main():
    bases={}
    for k in CFG:
        bases[k]=build_company(k)
        print(k,{y:(round(v['ir_rate'],4) if v['ir_rate'] else None, v['ir_atual_flat']) for y,v in list(bases[k]['years'].items())[:2]})
    json.dump(bases,open('bases_raw.json','w'),default=float)
main()
