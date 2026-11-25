import openpyxl,json,re
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter as L
D="/mnt/user-data/uploads/05 - Reforma Tributaria_Crescento - Claude/"
# (arquivo, aba, coluna rótulo, coluna 2027, linha inicial (None=achar), unidade divisor, nome)
CO=[('graos','5700 - Grãos.xlsx',0,2,22,None,1,'Grãos do Piauí'),
 ('mobi','5800 - CS Mobi.xlsx',0,2,20,None,1,'CS Mobi Cuiabá'),
 ('leste','5900 - Bloco Leste.xlsx',0,2,18,None,1000,'Bloco Leste'),
 ('mercosul','5910 - Mercosul.xlsx',0,2,18,None,1000,'Rodovias Mercosul'),
 ('norte','5920 - Escolas Norte.xlsx',0,2,17,'DRE Contábil',1,'Escolas Norte'),
 ('sul','5930 - Escolas Sul.xlsx',0,2,17,'DRE Contábil',1,'Escolas Sul'),
 ('mcp01','5940 - MCP01.xlsx',0,2,10,None,1,'MCP01'),
 ('welfare','9904 - Welfare.xlsx',0,2,18,None,1000,'Welfare Ambiental'),
 ('rota','9905 - Rota da Integração.xlsx',0,2,18,None,1,'Rota da Integração'),
 ('lote6','Lote6.xlsx',0,2,3,'DRE',1,'Lote 6')]
def norm(s): return re.sub(r'\s+',' ',str(s)).strip()
data={}  # key -> list of (label, value, rownum)
for key,f,sh,lc,vc,start,div,nm in CO:
    wb=openpyxl.load_workbook(D+f,data_only=True); ws=wb.worksheets[sh]
    r0=None
    if start:
        for r in range(1,ws.max_row+1):
            if norm(ws.cell(r,lc).value)==start: r0=r+1;break
    else:
        # primeira linha 'Receita ...' após a linha 'DRE'
        for r in range(1,40):
            v=ws.cell(r,lc).value
            if v is not None and norm(v).startswith(('DRE','DRE (')): r0=r+1;break
    rows=[]
    r=r0
    while r<=ws.max_row:
        lab=ws.cell(r,lc).value
        if lab is not None and str(lab).strip() not in ('0',''):
            v=ws.cell(r,vc).value
            v=(v or 0) if isinstance(v,(int,float)) or v is None else 0
            rows.append((norm(lab),v/div,r))
            if norm(lab).lower().startswith('receita líquida') or norm(lab).lower().startswith('receita liquida'): break
        r+=1
    data[key]=rows
json.dump({k:[(a,b) for a,b,_ in v] for k,v in data.items()},open('/tmp/w/dre/raw.json','w'),ensure_ascii=False,indent=0)
for k,v in data.items(): print(k,len(v),[a for a,_,_ in v][:3],v[-1][0],round(v[-1][1]))
