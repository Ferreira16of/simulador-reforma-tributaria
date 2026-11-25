import openpyxl,json,datetime,sys
from openpyxl.utils import column_index_from_string as CI
P="/mnt/user-data/uploads/05 - Reforma Tributaria_Crescento - Claude/"
SRC={'graos':('5700 - Grãos.xlsx','DF - Mensal ',423,'BX','FK',1,(425,426,None)),
'mobi':('5800 - CS Mobi.xlsx','DF - Mensal ',333,'BL','EY',1,(335,336,337)),
'leste':('5900 - Bloco Leste.xlsx','DFs Mensais',494,'AJ','DW',1000,(496,497,498)),
'mercosul':('5910 - Mercosul.xlsx','DF - Mensal Consolidado',335,'AN','EA',1000,(337,338,339)),
'norte':('5920 - Escolas Norte.xlsx','DF - Mensal | Norte',321,'AB','DO',1,(323,324,325)),
'sul':('5930 - Escolas Sul.xlsx','DF - Mensal | Sul',316,'AB','DO',1,(318,319,320)),
'mcp01':('5940 - MCP01.xlsx','DF - Mensal',392,'V','DI',1,(394,395,None)),
'welfare':('9904 - Welfare.xlsx','DF Mensal',321,'AN','EA',1000,(323,324,325)),
'rota':('9905 - Rota da Integração.xlsx','DF - Mensal',393,'AN','EA',1,(395,396,None))}
out={}
for k,(f,sh,hr,c1,c2,div,rows) in SRC.items():
    wb=openpyxl.load_workbook(P+f,data_only=True,read_only=True)
    names=wb.sheetnames
    ws=wb[sh] if sh in names else wb[[n for n in names if n.strip()==sh.strip()][0]]
    lo,hi=CI(c1),CI(c2)
    need={hr}|{r for r in rows if r}
    data={}
    for i,row in enumerate(ws.iter_rows(min_row=min(need),max_row=max(need),values_only=True),start=min(need)):
        if i in need: data[i]=row
    hdr=data[hr]; res={'deb':[0.0]*12,'opex':[0.0]*12,'capex':[0.0]*12}
    for ci in range(lo-1,hi):
        d=hdr[ci] if ci<len(hdr) else None
        if isinstance(d,datetime.datetime) and d.year==2027:
            m=d.month-1
            for key,r in zip(('deb','opex','capex'),rows):
                if r is None: continue
                v=data[r][ci] if ci<len(data[r]) else 0
                res[key][m]+= (v if isinstance(v,(int,float)) else 0)/div
    out[k]=res
    print(k,[round(sum(res[x])) for x in ('deb','opex','capex')],flush=True)
json.dump(out,open('flow27.json','w'))
