# Extrai o fluxo mensal 2027 (débitos, créditos OPEX/SG&A e CAPEX) das abas de cada empresa de
# "2.Creditos_Antes_x_Depois_Reforma-4.xlsx" e grava flow27_v15.json (usado pelo build_v15.py).
import openpyxl,json
f='2.Creditos_Antes_x_Depois_Reforma-4.xlsx'
wb=openpyxl.load_workbook(f,read_only=True,data_only=True)
sh={'graos':'GRÃOS','mobi':'CS MOBI','leste':'BLOCO LESTE','mercosul':'MERCOSUL','norte':'ESCOLAS NORTE','sul':'ESCOLAS SUL','mcp01':'MCP01','welfare':'WELFARE','rota':'ROTA','lote6':'LOTE 6'}
out={}
for k,s in sh.items():
    rows=list(wb[s].iter_rows(min_row=1,max_row=60,min_col=2,max_col=15,values_only=True))
    g=lambda lab:[r for r in rows if isinstance(r[0],str) and r[0].strip()==lab][0]
    z=lambda r:[x if isinstance(x,(int,float)) else 0 for x in r[1:13]]
    deb=z(g('Total débitos — Depois'));cred=z(g('Total créditos — Depois'));cap=[0]*12
    for t in ('CBS','IBS','ISS'): cap=[a+b for a,b in zip(cap,z(g('CAPEX (custo de construção) — '+t)))]
    out[k]={'deb':deb,'opex':[a-b for a,b in zip(cred,cap)],'capex':cap}
json.dump(out,open('flow27_v15.json','w'))
