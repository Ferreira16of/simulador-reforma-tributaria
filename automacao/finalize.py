import sys,json,openpyxl,copy,re; sys.path.insert(0,'/tmp/claude-0/s')
from motor import *
from openpyxl.utils import column_index_from_string as CI
P="/mnt/user-data/uploads/05 - Reforma Tributaria_Crescento - Claude/"
raw=json.load(open('bases_raw.json'))
names=json.load(open('names.json')) if False else None
NAMES={'graos':'Grãos do Piauí','mobi':'CS Mobi Cuiabá','leste':'Bloco Leste','mercosul':'Rodovias Mercosul','norte':'Escolas Norte','sul':'Escolas Sul','mcp01':'MCP01','welfare':'Welfare Ambiental','rota':'Rota da Integração','lote6':'Lote 6'}
ORDER=['graos','mobi','leste','mercosul','norte','sul','mcp01','welfare','rota','lote6']
CRED0={'graos':.85,'mobi':.85,'leste':.85,'mercosul':.85,'norte':.85,'sul':.85,'mcp01':1.0,'welfare':1.0,'rota':.85,'lote6':1.0}
FILES={'graos':'5700 - Grãos.xlsx','mobi':'5800 - CS Mobi.xlsx','leste':'5900 - Bloco Leste.xlsx','mercosul':'5910 - Mercosul.xlsx','norte':'5920 - Escolas Norte.xlsx','sul':'5930 - Escolas Sul.xlsx','mcp01':'5940 - MCP01.xlsx','welfare':'9904 - Welfare.xlsx','rota':'9905 - Rota da Integração.xlsx','lote6':'Lote6.xlsx (2028–2033) + premissas estimadas de 2027 (PPT)'}
# ---- Lote 6
def lote6():
    ws=openpyxl.load_workbook(P+"Lote6.xlsx",data_only=True).active
    g=lambda r,c: ws.cell(r,c).value if isinstance(ws.cell(r,c).value,(int,float)) else 0.0
    years={}
    starts={2028:4,2029:12,2030:20,2031:28,2032:36,2033:44}
    for y,a in starts.items():
        Y={'atual':{},'rt':{},'comp':{},'ir_rate':0.1525,'ir_atual_flat':True}
        rows={'RB':4,'DED':5,'OPEX':8,'CRED':9,'SGA':12,'DA':15,'FIN':19,'NONOP':20,'IR':24}
        for key,r in rows.items():
            Y['atual'][key]=g(r,a)/1000*1000/1000*1000/1000 if False else g(r,a)
        # convenção do PPT: ISS fica em Impostos, não na Receita Bruta
        cb=lambda r: g(r,a+3); ib=lambda r: g(r,a+5); sem=lambda r: g(r,a+1)
        Y['comp']['RB']={'sem':sem(4),'cbs':cb(4),'ibs':ib(4),'res':0.0,'regC':0.0,'regI':0.0}
        Y['comp']['DED']={'def':g(5,a+4)+g(5,a+7) if False else g(4,a+4),'fixed':0.0}   # ISS (linha 4 da coluna ISS) 
        for key,r in (('OPEX',8),('SGA',12),('DA',15)):
            Y['comp'][key]={'sem':sem(r),'cbs':cb(r),'ibs':ib(r),'res':0.0,'regC':0.0,'regI':0.0}
        Y['comp']['NONOP']={'sem':sem(20),'cbs':cb(20),'ibs':ib(20),'res':0.0,'regC':0.0,'regI':0.0}
        Y['comp']['FIN']=None;Y['comp']['IR']=None;Y['comp']['CRED']=None
        # RT consolidado (convenção PPT) p/ conferência
        Y['rt']={'RB':sem(4)+cb(4)+ib(4),'DED':g(4,a+4),'OPEX':sem(8)+cb(8)+ib(8),'CRED':0.0,'SGA':sem(12)+cb(12)+ib(12),'DA':sem(15)+cb(15)+ib(15),'FIN':g(19,a+1),'NONOP':g(20,a+1),'IR':g(24,a+7)}
        years[str(y)]=Y
    # 2027: bloco real em Lote6.xlsx (colunas C..J): atual, sem, Pis/Cofins/ISS, CBS, ISS, IBS, total, pós
    a=3
    Y={'atual':{},'rt':{},'comp':{},'ir_rate':0.34,'ir_atual_flat':True}
    rows={'RB':4,'DED':5,'OPEX':8,'CRED':9,'SGA':12,'DA':15,'FIN':19,'NONOP':20,'IR':24}
    for key,r in rows.items(): Y['atual'][key]=g(r,a)
    cb=lambda r: g(r,a+3); ib=lambda r: g(r,a+5); sem=lambda r: g(r,a+1)
    Y['comp']['RB']={'sem':sem(4),'cbs':cb(4),'ibs':ib(4),'res':0.0,'regC':0.0,'regI':0.0}
    Y['comp']['DED']={'def':g(5,a+4),'fixed':0.0}   # ISS na linha de deduções (bloco 2027)
    for key,r in (('OPEX',8),('SGA',12),('DA',15),('NONOP',20)):
        Y['comp'][key]={'sem':sem(r),'cbs':cb(r),'ibs':ib(r),'res':0.0,'regC':0.0,'regI':0.0}
    Y['comp']['FIN']=None;Y['comp']['IR']=None;Y['comp']['CRED']=None
    Y['rt']={'RB':sem(4)+cb(4)+ib(4),'DED':g(5,a+4),'OPEX':sem(8)+cb(8)+ib(8),'CRED':0.0,'SGA':sem(12)+cb(12)+ib(12),'DA':sem(15)+cb(15)+ib(15),'FIN':g(19,a+1),'NONOP':g(20,a+1),'IR':g(24,a+7)}
    years['2027']=Y
    return {'nome':'Lote 6','years':years}
raw['lote6']=lote6()
# fill regC/regI gaps and prepare
out={}
for k in ORDER:
    co=raw[k]; co['credito0']=CRED0[k]
    pc=prepare(co); pc['credito0']=CRED0[k]
    for L in CORE:
        # preencher regC/regI ausentes com valor mais próximo da mesma linha
        vals_c=[(int(y),pc['years'][y]['comp'][L]['regC']) for y in pc['years'] if pc['years'][y]['comp'].get(L) and pc['years'][y]['comp'][L].get('regC') is not None]
        vals_i=[(int(y),pc['years'][y]['comp'][L]['regI']) for y in pc['years'] if pc['years'][y]['comp'].get(L) and pc['years'][y]['comp'][L].get('regI') is not None]
        for y in pc['years']:
            c=pc['years'][y]['comp'].get(L)
            if not c: continue
            if c.get('regC') is None: c['regC']=min(vals_c,key=lambda t:abs(t[0]-int(y)))[1] if vals_c else 0.0
            if c.get('regI') is None: c['regI']=min(vals_i,key=lambda t:abs(t[0]-int(y)))[1] if vals_i else 0.0
    out[k]=pc
json.dump(out,open('prepared.json','w'))
print("ok",[ (k,len(v['years'])) for k,v in out.items()])
