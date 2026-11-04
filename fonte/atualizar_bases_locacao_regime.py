#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Atualiza o bases.json do simulador com duas informações vindas das planilhas da pasta:
 1) RECEITA DE LOCAÇÃO por empresa/ano, lida das DFs (CS Mobi: 'Rec. Acessórias - Locação'; Bloco Leste: 'Locação Espaço p/ Construção ...').
    Grava em comp.RB.locValor (R$ mil) e comp.RB.shareLoc (% da base tributada da receita). Sobre essa parcela vale a redução de 70% (parâmetro).
 2) % de CRÉDITO (fornecedores no Regime Normal) por empresa, lido de 'Regime Tributário ... Assertif.xlsx', coluna X ('Regime Normal', tabela de empresas).
    Grava em empresas[k].creditoRN (o motor usa como crédito padrão da simulação; credito0 continua sendo o % embutido nas DFs).
Uso: python3 atualizar_bases_locacao_regime.py <pasta_das_planilhas> <bases.json>
"""
import sys, os, re, glob, json
import openpyxl

MAP_REGIME = {'GRÃOS': 'graos', 'CS MOBI': 'mobi', 'BLOCO LESTE': 'leste', 'MERCOSUL': 'mercosul', 'ESCOLAS NORTE': 'norte',
              'ESCOLAS SUL': 'sul', 'ROTA': 'rota', 'MCP01': 'mcp01', 'WELFARE': 'welfare', 'LOTE 6': 'lote6'}


def num(x):
    if isinstance(x, (int, float)): return float(x)
    try: return float(str(x).replace(',', '.'))
    except Exception: return None


def ler_regime(pasta):
    f = glob.glob(os.path.join(pasta, 'Regime Tribut*Assertif*.xlsx'))
    f = [x for x in f if not os.path.basename(x).startswith('~')]
    if not f: sys.exit('Planilha "Regime Tributário ... Assertif" não encontrada')
    ws = openpyxl.load_workbook(f[0], data_only=True).active
    out = {}
    # tabela de empresas: cabeçalho 'Empresa' | 'Regime Normal' na coluna W/X (2ª ocorrência = lista completa incl. MCP01, Welfare e Lote 6)
    hdr = [c.row for c in ws['W'] if c.value == 'Empresa']
    r0 = hdr[-1] + 1
    for r in range(r0, r0 + 15):
        nome = ws.cell(r, 23).value; v = num(ws.cell(r, 24).value)
        if isinstance(nome, str) and nome.strip().upper() in MAP_REGIME and v is not None:
            out[MAP_REGIME[nome.strip().upper()]] = v
    return out, os.path.basename(f[0])


def loc_mobi(pasta):
    wb = openpyxl.load_workbook(os.path.join(pasta, '5800 - CS Mobi.xlsx'), read_only=True, data_only=True)
    ws = [w for w in wb if w.title.strip() == 'DF - Anual'][0]
    rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=80, values_only=True))
    hdr = next(r for r in rows if r[1] == 'Ano')
    ycol = {int(x): j for j, x in enumerate(hdr) if isinstance(x, (int, float)) and 2026 <= x <= 2033}
    lin = next(r for r in rows if isinstance(r[1], str) and r[1].strip() == 'Rec. Acessórias - Locação')
    return {y: (num(lin[j]) or 0) for y, j in ycol.items()}                 # já em R$ mil


def loc_leste(pasta):
    wb = openpyxl.load_workbook(os.path.join(pasta, '5900 - Bloco Leste.xlsx'), read_only=True, data_only=True)
    ws = wb['DFs Anuais']
    rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=80, values_only=True))
    hdr = next(r for r in rows if r[1] == 'Ano')
    ycol = {int(x): j for j, x in enumerate(hdr) if str(x).isdigit() and 2026 <= int(x) <= 2033}
    lins = [r for r in rows if isinstance(r[1], str) and r[1].strip().startswith('Locação Espaço')]
    return {y: sum(num(r[j]) or 0 for r in lins) / 1000.0 for y, j in ycol.items()}   # DF em R$ -> R$ mil


def main(pasta, bases_path):
    B = json.load(open(bases_path, encoding='utf8'))
    rn, fn = ler_regime(pasta)
    loc = {'mobi': loc_mobi(pasta), 'leste': loc_leste(pasta)}
    for k, e in B['empresas'].items():
        if k in rn: e['creditoRN'] = rn[k]
        e['fonte_creditoRN'] = fn
        for ys, Y in e['years'].items():
            rb = Y['comp'].get('RB')
            if not rb: continue
            lv = loc.get(k, {}).get(int(ys), 0.0)
            den = (rb['cpValor'] / rb['shareCP']) if rb.get('shareCP') else (-rb['cbs'] / 0.0925 if rb.get('cbs') else rb['sem'])
            rb['locValor'] = lv
            rb['shareLoc'] = min(1.0, lv / den) if den else 0.0
    json.dump(B, open(bases_path, 'w', encoding='utf8'), ensure_ascii=False)
    print('Regime Normal (crédito %):', {k: v for k, v in rn.items()})
    for k in ('mobi', 'leste'):
        print(k, {y: (round(loc[k][y], 1), round(B['empresas'][k]['years'][str(y)]['comp']['RB']['shareLoc'] * 100, 2)) for y in sorted(loc[k])})


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
