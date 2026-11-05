#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplica a planilha "Base_Redução" ao bases.json do simulador.
Regra (linhas da DRE de cada DF, base tributada = base do CBS da própria DF):
  - Redução de 44,2% (contraprestação): contraprestações, receita financeira IFRS 09 / remuneração do ativo financeiro (quando tributada),
    Receita de Operação - Contábil (Escolas Norte/Sul) e a Receita Bruta inteira da MCP01.
  - Redução de 70% (locação): receita acessória, locação, ABL, publicidade e naming rights.
  - Demais linhas (tarifárias, Mercosul, Welfare, Rota, construção etc.): sem redução.
Grava em empresas[k].years[ano].comp.RB: cpValor/shareCP (44,2%) e locValor/shareLoc (70%), onde share = valor / base tributada.
Uso: python3 atualizar_reducao_bases.py <pasta_das_planilhas> <bases.json>
"""
import sys,os,re,json,openpyxl
CP={'graos':['Receita Contraprestação'],'mobi':['Contraprestação Contratual','Contraprestação Pecuniária'],'leste':['Contraprestação OPEX','Receita Financeira IFRS 09'],
    'norte':['Receita de Operação - Contábil'],'sul':['Receita de Operação - Contábil'],'mcp01':['Receita Bruta']}
LOC={'graos':['Receita Acessória'],
     'mobi':['Rec. Acessórias - Locação','Rec. Acessórias - Publicidade - Outdoor','Rec. Acessórias - Publicidade - Mercado','Naming Rights'],
     'leste':['Locação Espaço p/ Construção ABL AE Carvalho','Locação Espaço p/ Construção Shopping Sacomã','Locação Espaço p/ Construção ABL Sapopemba','Locação Espaço p/ Construção Shopping Cidade Tiradentes','ABL Interna - Todos Terminais','Publicidade - Todos Terminais','Naming Right - 03 Terminais'],
     'norte':['Receita Acessória'],'sul':['Receita Acessória'],'lote6':['Receita Acessória']}
FILES={'graos':('5700 - Grãos.xlsx',1,None),'mobi':('5800 - CS Mobi.xlsx',1,None),'leste':('5900 - Bloco Leste.xlsx',1000,None),'mercosul':('5910 - Mercosul.xlsx',1000,None),
 'norte':('5920 - Escolas Norte.xlsx',1,'DRE Contábil'),'sul':('5930 - Escolas Sul.xlsx',1,'DRE Contábil'),'mcp01':('5940 - MCP01.xlsx',1,None),'welfare':('9904 - Welfare.xlsx',1000,None),'rota':('9905 - Rota da Integração.xlsx',1000/1000,None),'lote6':('9906 - Lote 6.xlsx',1,None)}
ALL_TAXED_OK={'norte','sul','mcp01'}   # toda a base tributada recebe a redução de 44,2%
def norm(s): return re.sub(r'\s+',' ',str(s)).strip()
def linhas(pasta,k):
    f,div,start=FILES[k]; ws=openpyxl.load_workbook(os.path.join(pasta,f),data_only=True).worksheets[0]
    ycol={}
    for r in range(1,13):
        tmp={}
        for c in range(1,ws.max_column+1):
            v=str(ws.cell(r,c).value).strip()
            if v in [str(y) for y in range(2026,2034)] and int(v) not in tmp: tmp[int(v)]=c
        if len(tmp)>len(ycol): ycol=tmp
    lc=4 if k=='lote6' else 2; r0=None
    for r in range(1,60):
        v=ws.cell(r,lc).value
        if start and norm(v)==start: r0=r+1;break
        if not start and v is not None and norm(v).startswith('DRE'): r0=r+1;break
    out={}; r=r0
    while r<=ws.max_row:
        lab=ws.cell(r,lc).value
        if lab is not None and str(lab).strip() not in ('0',''):
            n=norm(lab)
            if n not in out: out[n]={y:(ws.cell(r,c).value if isinstance(ws.cell(r,c).value,(int,float)) else 0)/div for y,c in ycol.items()}
            if n.lower().startswith('receita líquida'): break
        r+=1
    return out
def main(pasta,bp):
    B=json.load(open(bp,encoding='utf8'))
    for k,e in B['empresas'].items():
        L=linhas(pasta,k) if k in FILES else {}
        for ys,Y in e['years'].items():
            rb=Y['comp'].get('RB')
            if not rb: continue
            y=int(ys)
            den=(-rb['cbs']/0.0925) if rb.get('cbs') else (rb['sem'] or 0)
            cp=sum(L.get(l,{}).get(y,0) for l in CP.get(k,[])); lo=sum(L.get(l,{}).get(y,0) for l in LOC.get(k,[]))
            if k=='mobi':   # a DF tributa a Contraprestação Contratual (bruta) ou a Pecuniária (líquida do Desconto IFRS 09), conforme o ano
                lo_=lo; cands=[L.get(n,{}).get(y,0) for n in ('Contraprestação Contratual','Contraprestação Pecuniária')]
                cp=min(cands,key=lambda c:abs((c+lo_)-den)) if den else cands[0]
            if k=='lote6':   # Lote6.xlsx (2027) não tem detalhe: usa a proporção da 9906
                tot=L.get('Receita Bruta',{}).get(y,0); lo=(lo/tot*den) if tot else 0
            if k in ALL_TAXED_OK: cp=den-lo
            if den:
                s=cp+lo
                if s>den*1.0001: cp*=den/s; lo*=den/s
            rb['cpValor']=cp; rb['locValor']=lo
            rb['shareCP']=(cp/den) if den else (1.0 if k in ALL_TAXED_OK else 0.0); rb['shareLoc']=(lo/den) if den else 0.0
    json.dump(B,open(bp,'w',encoding='utf8'),ensure_ascii=False)
    for k,e in B['empresas'].items():
        print(k,{y:(round(c['comp']['RB']['shareCP']*100,1),round(c['comp']['RB']['shareLoc']*100,1)) for y,c in sorted(e['years'].items()) if c['comp'].get('RB') and y in('2026','2027','2028','2033')})
if __name__=='__main__': main(sys.argv[1],sys.argv[2])
