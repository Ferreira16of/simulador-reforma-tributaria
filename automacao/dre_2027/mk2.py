exec(open('mk_base.py').read())
import json
M=json.load(open('mot27.json')); RED_CP,RED_LOC=0.442,0.70
# portadores de tributo: linha -> (classe)  (cp=44,2% | loc=70% | none)
CAR={'graos':{'Transcerrados - CP Proteção':'cp','Transcerrados - CP Fixa':'cp','Aditivo - CP Fixa':'cp','Transcerrados':'none','Aditivo':'none','Receita Acessória':'loc'},
 'mobi':{'Contraprestação Pecuniária':'cp','Rec. Acessórias - Locação':'loc','Rec. Acessórias - Publicidade - Outdoor':'loc','Rec. Acessórias - Publicidade - Mercado':'loc','Naming Rights':'loc'},
 'leste':{'Contraprestação OPEX':'cp','Receita Financeira IFRS 09':'cp','Locação Espaço p/ Construção ABL AE Carvalho':'loc','Locação Espaço p/ Construção Shopping Sacomã':'loc','Locação Espaço p/ Construção ABL Sapopemba':'loc','Locação Espaço p/ Construção Shopping Cidade Tiradentes':'loc','ABL Interna - Todos Terminais':'loc','Publicidade - Todos Terminais':'loc','Naming Rights':'loc'},
 'mercosul':{'Receita Tarifária Brasil Veículos Leves':'none','Receita Tarifária Brasil Veículos Pesados':'none','Serviços do CUF Brasil':'none','Receita Acessoria Brasil':'none','Receita de recomposição':'none'},
 'norte':{'Receita de Operação - Contábil':'cp','Receita Financeira (IFRIC 12)':'cp'},'sul':{'Receita de Operação - Contábil':'cp','Receita Financeira (IFRIC 12)':'cp'},
 'mcp01':{'Receita Bruta':'cp'},
 'welfare':{n:'none' for n in ['Varzea Grande','Nobres','Nossa senhora do livramento','Jangada','Acorizal','Rosário do Oeste','Barra do Bugres','Porto Estrela','Cuiabá','Privado','Lodo','Campo Verde','Nova Mutum','DAE Varzea Grande','Chapada de Guimarães']},
 'rota':{'Receita Tarifária':'none','Receita Reequilíbrio':'none','(-) Obrigações Edital':'none'},
 'lote6':{}}
# pais derivados: soma dos filhos / herdam imposto
PAI={'graos':{'Receita Contraprestação':['Transcerrados - CP Proteção','Transcerrados - CP Fixa','Aditivo - CP Fixa'],'Receita Tarifária':['Transcerrados','Aditivo']},
     }   # Pecuniária = Contratual + Desconto: herda o tributo da Contratual
pos={}; chk={}
for k,_ in CO:
    m=M[k]; car=CAR[k]; tax={}
    for c,cls in car.items():
        v=vals.get((c,k),0) or 0
        red={'cp':RED_CP,'loc':RED_LOC,'none':0}[cls]
        tax[c]=m['tfull']*v*(1-red)/m['den'] if m['den'] else 0
    tot=sum(tax.values()) if car else m['trt']
    chk[k]=(round(tot,1),round(m['trt'],1))
    for p,ch in PAI.get(k,{}).items(): tax[p]=sum(tax[x] for x in ch)
    for (c,kk),v in list(vals.items()):
        pass
    P={}
    for c,_,_ in used:
        v=vals.get((c,k))
        if v is None: continue
        P[c]=v+tax.get(c,0)
    P['Receita Bruta']=vals[('Receita Bruta',k)]+tot
    if ('Receita Bruta com Construção completa',k) in vals: P['Receita Bruta com Construção completa']=vals[('Receita Bruta com Construção completa',k)]+tot
    # deduções
    dA=vals[('Deduções / Impostos (total)',k)]; dP=dA+(m['dedR']-m['dedA']); P['Deduções / Impostos (total)']=dP
    if ('IIBB',k) in vals: P['IIBB']=vals[('IIBB',k)]
    if ('ISS',k) in vals: P['ISS']=dP-P.get('IIBB',0)
    for c in ('PIS/COFINS','CBS','IBS','Crédito PIS/COFINS'):
        if (c,k) in vals: P[c]=0.0
    cred=vals.get(('Crédito PIS/COFINS',k),0)
    P['Receita Líquida']=vals[('Receita Líquida',k)]+(P['Receita Bruta']-vals[('Receita Bruta',k)])+(dP-dA)-cred
    pos[k]=P
bad={k:v for k,v in chk.items() if abs(v[0]-v[1])>1}
print('tributo por linha vs motor (ok se vazio):',bad)
# conferência dedução: soma das linhas = total
for k in pos:
    sub=sum(pos[k].get(c,0) for c in ('ISS','IIBB','PIS/COFINS','CBS','IBS') if (c,k) in vals)
    if any((c,k) in vals for c in ('ISS','IIBB','PIS/COFINS')): print('ded',k,round(sub,1),round(pos[k]['Deduções / Impostos (total)'],1))
wb=openpyxl.Workbook()
NAVY='17365D'; hf=Font(bold=True,color='FFFFFF',name='Calibri',size=10); hfill=PatternFill('solid',fgColor=NAVY)
num='#,##0;(#,##0);"–"'; thin=Side(style='thin',color='C9D3DF')
nc=len(CO); H=4
def sheet(ws,title,sub,kind):
    ws['A1']=title; ws['A1'].font=Font(bold=True,size=13,color=NAVY)
    ws['A2']=sub; ws['A2'].font=Font(italic=True,size=9,color='5B6573')
    ws.cell(H,1,'DRE 2027 (R$ mil)')
    for j,(k,n) in enumerate(CO): ws.cell(H,2+j,n)
    heads=['Total 9 empresas','Total com Lote 6']
    if kind=='atual': heads+=['Pós-Reforma 9 empresas','Pós-Reforma com Lote 6','Δ 9 empresas','Δ com Lote 6']
    for j,h in enumerate(heads): ws.cell(H,2+nc+j,h)
    last=1+nc+len(heads)
    for c in range(1,last+1):
        x=ws.cell(H,c); x.font=hf; x.fill=hfill; x.alignment=Alignment(horizontal='center' if c>1 else 'left',vertical='center',wrap_text=True)
    ws.row_dimensions[H].height=32
    return last
rowof={}
def fill(ws,kind,last):
    r=H+1
    for c,l,b in used:
        ws.cell(r,1,c).alignment=Alignment(indent=l*2); rowof[c]=r
        for j,(k,n) in enumerate(CO):
            col=L(2+j)
            if kind=='atual': v=vals.get((c,k)); v=None if v is None else round(v,3)
            elif kind=='pos': v=pos[k].get(c); v=None if v is None else round(v,3)
            else: v=f"=N('Pós-Reforma 2027'!{col}{r})-N('Receita 2027'!{col}{r})" if (c,k) in vals else None
            if v is not None: ws.cell(r,2+j,v)
            ws.cell(r,2+j).number_format=num
        a9,a10=L(2+nc),L(3+nc)
        ws.cell(r,2+nc,f'=SUM(B{r}:{L(1+nc-1)}{r})'); ws.cell(r,3+nc,f'=SUM(B{r}:{L(1+nc)}{r})')
        if kind=='atual':
            ws.cell(r,4+nc,f"='Pós-Reforma 2027'!{a9}{r}"); ws.cell(r,5+nc,f"='Pós-Reforma 2027'!{a10}{r}")
            ws.cell(r,6+nc,f"='Delta 2027'!{a9}{r}"); ws.cell(r,7+nc,f"='Delta 2027'!{a10}{r}")
        for cc in range(2+nc,last+1): ws.cell(r,cc).number_format=num; ws.cell(r,cc).font=Font(bold=True)
        if b:
            for cc in range(1,last+1): ws.cell(r,cc).font=Font(bold=True); ws.cell(r,cc).fill=PatternFill('solid',fgColor='EAF1FB')
        for cc in range(1,last+1): ws.cell(r,cc).border=Border(bottom=thin)
        r+=1
    return r
ws1=wb.active; ws1.title='Receita 2027'
ws2=wb.create_sheet('Pós-Reforma 2027'); ws3=wb.create_sheet('Delta 2027')
l1=sheet(ws1,'Receita Bruta → Receita Líquida | DRE 2027 por empresa (modelo atual, como está em cada DF)','R$ mil | Coluna "2027" da DF - Anual de cada empresa (Bloco Leste, Mercosul e Welfare convertidos de R$ para R$ mil). À direita: totais pós-reforma e delta (abas Pós-Reforma 2027 e Delta 2027).','atual')
l2=sheet(ws2,'DRE 2027 PÓS-REFORMA por empresa (cenário padrão da simulação)','R$ mil | CBS 11,3468% + IBS (sem fase em 2027), redução de 44,2% nas contraprestações/receitas financeiras e 70% na receita acessória/locação (planilha Base_Redução). ISS e IIBB mantidos; PIS/COFINS e crédito PIS/COFINS zerados.','pos')
l3=sheet(ws3,'DELTA 2027 | Pós-Reforma − Atual por empresa','R$ mil | Pós-Reforma 2027 menos modelo atual (negativo = redução da receita).','delta')
r1=fill(ws1,'atual',l1); r2=fill(ws2,'pos',l2); r3=fill(ws3,'delta',l3)
notes=['Fonte: DF - Anual de cada empresa (5700, 5800, 5900, 5910, 5920, 5930, 5940, 9904, 9905) e Lote6.xlsx, coluna 2027. Valores como estão nas planilhas.',
 'Escolas Norte e Sul: bloco "DRE Contábil" (inclui Receita de Construção e Receita Financeira IFRIC 12), o mesmo usado no simulador.',
 'Lote 6: Lote6.xlsx (2027 com início de operação), mesma base do simulador; a planilha 9906 - Lote 6 traz 2027 de ano cheio (RB 268.624) e não foi usada. Só há Receita Bruta total; o efeito da reforma aparece apenas nos totais.',
 'Receita Líquida segue a DF de cada empresa: em Bloco Leste, Escolas Norte e Escolas Sul inclui a Receita de Construção; nas demais não. O total mistura as duas bases.',
 'CS Mobi: a Receita Bruta da DF (40.091) inclui 3.233 de construção do ativo financeiro; o simulador usa 36.859.',
 'Pós-Reforma: a CBS/IBS de cada linha = tributo total do simulador distribuído pela base tributada da DF, com a redução da planilha Base_Redução (44,2% ou 70% conforme a linha). Linhas fora da base do CBS (construção, remuneração do ativo financeiro, Argentina) não mudam. Receita de Construção não é alterada pelo simulador.',
 'Linhas com recuo detalham a linha-mãe e não devem ser somadas com ela. Em CS Mobi (2027), o CBS incide sobre a Contraprestação Pecuniária (Contratual menos Desconto IFRS 09); Contratual e Desconto não recebem tributo na abertura.']
for ws,r in ((ws1,r1),(ws2,r2),(ws3,r3)):
    r+=1
    for t in notes:
        ws.cell(r,1,t).font=Font(size=9,color='5B6573'); r+=1
    ws.column_dimensions['A'].width=52
    for c in range(2,40): ws.column_dimensions[L(c)].width=15
    ws.freeze_panes=ws.cell(H+1,2); ws.sheet_view.showGridLines=False
wb.save('/tmp/w/out/DRE_2027_Receita_por_Empresa_v5.xlsx'); print('ok')
