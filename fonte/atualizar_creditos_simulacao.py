#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Leva a planilha 2.Creditos_Antes_x_Depois_Reforma para as PREMISSAS DA SIMULAÇÃO:
  - alíquotas CBS 11,3468% e IBS ref. = 28% - CBS, redutor 44,2% sobre a contraprestação;
  - redução de 70% sobre a receita de LOCAÇÃO (Mobi e Bloco Leste);
  - crédito (% de fornecedores no Regime Normal) por empresa, conforme 'Regime Tributário ... Assertif' (coluna X).
Nas abas das empresas, a seção "6. DADOS DE ORIGEM (DF)" tem valores digitados às alíquotas originais das DFs
(CBS 9,25%, IBS ref. 18,75%, crédito 85%/100%, sem redutor). O script:
  1. guarda esses valores originais na seção "7. BASE ORIGINAL" (linhas 138-147; preservada em novas execuções);
  2. troca as linhas de CBS/IBS (débitos e créditos) por fórmulas = base original x fatores da aba Premissas;
  3. monta em Premissas as tabelas: alíquotas/redutores por ano, % contraprestação, % locação e crédito por empresa.
Uso: python3 atualizar_creditos_simulacao.py <2.Creditos.xlsx> <saida.xlsx> [bases.json]
"""
import sys, os, json, copy
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter as CL

SHEETS = [('graos', 'GRÃOS'), ('mobi', 'CS MOBI'), ('leste', 'BLOCO LESTE'), ('mercosul', 'MERCOSUL'), ('norte', 'ESCOLAS NORTE'),
          ('sul', 'ESCOLAS SUL'), ('mcp01', 'MCP01'), ('welfare', 'WELFARE'), ('rota', 'ROTA'), ('lote6', 'LOTE 6')]
YEARS = list(range(2026, 2034))
# linha alvo -> (tipo, usa redutor?, coluna do fator de crédito em Premissas)  tipo: C = CBS, I = IBS
TARGET = [(112, 'C', True, None), (113, 'I', True, None),
          (116, 'C', False, 'K'), (117, 'I', False, 'L'), (119, 'C', False, 'M'), (120, 'I', False, 'N'),
          (122, 'C', False, 'O'), (123, 'I', False, 'P'), (125, 'C', False, 'K'), (126, 'I', False, 'L')]   # CAPEX segue os flags do Opex
BASE0 = 138
YCOL = {'Q': 'C', 'R': 'D', 'S': 'E', 'T': 'F', 'U': 'G', 'V': 'H', 'W': 'I', 'X': 'J'}   # coluna da aba -> coluna do ano em Premissas
for c in 'CDEFGHIJKLMN':
    YCOL[c] = 'D'   # meses de 2027
R_CBS, R_IBS, R_RED, R_CBS0, R_IBS0, R_FC, R_FI, R_LOC = 10, 11, 12, 13, 14, 15, 16, 17
R_SH0, R_LS0, R_CR0 = 20, 33, 46      # primeiras linhas: % contraprestação, % locação, crédito
YEL = PatternFill('solid', fgColor='FFF2CC')
B = Font(bold=True)


def main(src, out, bases):
    wb = openpyxl.load_workbook(src)
    B_ = json.load(open(bases, encoding='utf8'))['empresas']
    pr = wb['Premissas']
    pr['B8'] = 'Alíquotas da simulação (CBS, IBS e redutores) — aplicadas às abas das empresas'; pr['B8'].font = B
    pr['B9'] = 'Ano'; pr['B9'].font = B
    labels = {R_CBS: ('Alíquota CBS da simulação', 0.113468, True), R_IBS: ('Alíquota IBS referência da simulação (28% − CBS)', None, False),
              R_RED: ('Redutor s/ contraprestação (aplicado à CBS e ao IBS da receita)', 0.442, True),
              R_CBS0: ('Alíquota CBS usada nas DFs (base)', 0.0925, True), R_IBS0: ('Alíquota IBS referência usada nas DFs (base)', 0.1875, True),
              R_FC: ('Fator CBS = simulada ÷ base', None, False), R_FI: ('Fator IBS = simulada ÷ base', None, False),
              R_LOC: ('Redução s/ receita de locação (aplicada à CBS e ao IBS da receita de locação)', 0.70, True)}
    for j, y in enumerate(YEARS):
        c = CL(3 + j)
        pr[c + '9'] = y; pr[c + '9'].font = B
        for r, (nm, val, edit) in labels.items():
            pr['B%d' % r] = nm
            cell = pr['%s%d' % (c, r)]
            if r == R_IBS: cell.value = '=0.28-%s%d' % (c, R_CBS)
            elif r == R_FC: cell.value = '=%s%d/%s%d' % (c, R_CBS, c, R_CBS0)
            elif r == R_FI: cell.value = '=%s%d/%s%d' % (c, R_IBS, c, R_IBS0)
            elif val is not None and (cell.value is None or r in (R_CBS0, R_IBS0)):
                cell.value = val          # CBS, redutores: preserva o que o usuário já editou
            if edit: cell.fill = YEL
            cell.number_format = '0.0000%' if r in (R_CBS, R_IBS, R_RED, R_CBS0, R_IBS0, R_LOC) else '0.0000'

    def tabela(r_title, titulo, r0, getter, fmt='0.0%'):
        pr['B%d' % r_title] = titulo; pr['B%d' % r_title].font = B
        pr['B%d' % (r_title + 1)] = 'Empresa'; pr['B%d' % (r_title + 1)].font = B
        for j, y in enumerate(YEARS):
            pr[CL(3 + j) + str(r_title + 1)] = y; pr[CL(3 + j) + str(r_title + 1)].font = B
        for i, (key, nome) in enumerate(SHEETS):
            r = r0 + i
            pr['B%d' % r] = nome
            for j, y in enumerate(YEARS):
                c = pr['%s%d' % (CL(3 + j), r)]
                c.value = getter(key, y); c.number_format = fmt
    def rbfield(f):
        def g(key, y):
            Y = B_[key]['years'].get(str(y)); rb = ((Y or {}).get('comp', {}).get('RB') or {})
            return float(rb.get(f) or 0.0)
        return g
    tabela(18, '% da receita que é contraprestação, por empresa (base do redutor; vem das planilhas DF via bases.json)', R_SH0, rbfield('shareCP'))
    tabela(31, '% da receita que é LOCAÇÃO, por empresa (base da redução de 70%): CS Mobi "Rec. Acessórias - Locação"; Bloco Leste "Locação Espaço p/ Construção"', R_LS0, rbfield('shareLoc'), '0.00%')
    # crédito por empresa
    pr['B44'] = 'Crédito de CBS/IBS por empresa: % de fornecedores no Regime Normal (planilha "Regime Tributário ... Assertif", coluna X)'; pr['B44'].font = B
    hdr = ['Empresa', 'Crédito base das DFs', 'Crédito simulação (Regime Normal)', 'Opex CBS', 'Opex IBS', 'SG&A CBS', 'SG&A IBS', 'D&A CBS', 'D&A IBS',
           'Fator Opex/CAPEX CBS', 'Fator Opex/CAPEX IBS', 'Fator SG&A CBS', 'Fator SG&A IBS', 'Fator D&A CBS', 'Fator D&A IBS']
    for j, h in enumerate(hdr):
        c = pr.cell(45, 2 + j); c.value = h; c.font = B
    pr['B47'] = None
    for i, (key, nome) in enumerate(SHEETS):
        r = R_CR0 + i
        e = B_[key]; y0 = e['years'].get('2028') or list(e['years'].values())[0]
        pr['B%d' % r] = nome
        pr['C%d' % r] = e['credito0']; pr['C%d' % r].number_format = '0.0%'
        cn = e.get('creditoRN', e['credito0'])
        if pr['D%d' % r].value is None: pr['D%d' % r] = cn
        pr['D%d' % r].number_format = '0.0%'; pr['D%d' % r].fill = YEL
        for j, (L, f) in enumerate([('OPEX', 'regC'), ('OPEX', 'regI'), ('SGA', 'regC'), ('SGA', 'regI'), ('DA', 'regC'), ('DA', 'regI')]):
            LL = L
            if L == 'SGA' and abs((y0['comp'].get('SGA') or {}).get('sem') or 0) < 1e-9:
                LL = 'OPEX'      # empresas sem SG&A no motor: a planilha separa parte do Opex na linha SG&A, que segue a regra do Opex
            pr.cell(r, 5 + j).value = float((y0['comp'].get(LL) or {}).get(f) or 0.0)       # E..J: flags (1 = o crédito acompanha o % de fornecedores)
            pr.cell(r, 11 + j).value = '=1+%s%d*($D%d/$C%d-1)' % (CL(5 + j), r, r, r)     # K..P: fatores
            pr.cell(r, 11 + j).number_format = '0.0000'
    pr['B57'] = 'Flags (colunas E a J): 1 = o crédito dessa linha acompanha o % de fornecedores no Regime Normal; 0 = fixo (mesma regra do simulador HTML).'
    pr.column_dimensions['B'].width = max(pr.column_dimensions['B'].width or 0, 70)
    nform = 0
    for i, (key, nome) in enumerate(SHEETS):
        ws = wb[nome]
        shr, lsr, crr = R_SH0 + i, R_LS0 + i, R_CR0 + i
        ws['B136'] = '7. BASE ORIGINAL (alíquotas e créditos das DFs: CBS 9,25%, IBS ref. 18,75%, crédito original, sem redutor) — não editar; as linhas 112-126 acima = base x fatores da aba Premissas'
        ws['B136'].font = B
        for k, (tr, kind, red, fcol) in enumerate(TARGET):
            br = BASE0 + k
            if ws.cell(br, 2).value is None:           # primeira execução: guarda a base original
                ws.cell(br, 2).value = '%s (base original)' % ws.cell(tr, 2).value
                for col in range(3, 25):
                    cell = ws.cell(tr, col)
                    v = cell.value
                    if isinstance(v, (int, float)):
                        ws.cell(br, col).value = v
                        ws.cell(br, col).number_format = cell.number_format
                        ws.cell(br, col).font = copy.copy(cell.font)
            for col in range(3, 25):
                L = CL(col)
                bc = ws.cell(br, col)
                if not isinstance(bc.value, (int, float)):
                    continue                              # coluna que já era fórmula (O, R, C..N do Lote 6) permanece
                yc = YCOL.get(L)
                if yc is None: continue
                fat = R_FC if kind == 'C' else R_FI
                f = '=%s%d*Premissas!$%s$%d' % (L, br, yc, fat)
                if red:
                    f += '*(1-Premissas!$%s$%d*Premissas!$%s$%d-Premissas!$%s$%d*Premissas!$%s$%d)' % (yc, R_RED, yc, shr, yc, R_LOC, yc, lsr)
                else:
                    f += '*Premissas!$%s$%d' % (fcol, crr)
                ws.cell(tr, col).value = f
                nform += 1
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    print('fórmulas gravadas:', nform, '->', out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'simulador', 'bases.json'))
