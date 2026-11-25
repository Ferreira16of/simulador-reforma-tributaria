#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automatiza a aba "Input_DRE  - RT" da planilha 1.Input_Orçamento_2027 ... xlsx.

Cadeia de dados
  bases das empresas (DF) -> 2.Creditos_Antes_x_Depois_Reforma-4.xlsx (abas "Consolidado com/sem Lote 6")
  -> aba Base_RT (valores, R$ mil) -> aba "Input_DRE  - RT" (fórmulas, R$) -> aba Check_RT (conferências)

O que o script faz
  1. Lê as abas "Consolidado com Lote 6" e "Consolidado sem Lote 6" da planilha 2.Creditos (valores calculados).
  2. Grava/atualiza a aba Base_RT na planilha Input (com a chave "Incluir Lote 6", padrão = 1) .
  3. Converte as células digitadas à mão da aba "Input_DRE  - RT" em fórmulas (ligadas à Base_RT e à aba Input_DRE).
  4. Cria/atualiza a aba Check_RT (diferenças que devem ser zero + resumo dos efeitos).
  5. Recalcula o arquivo (LibreOffice, se existir) para gravar os valores calculados nas novas fórmulas.
O arquivo é editado direto no XML: as demais abas (e as propriedades do Smart View) ficam intactas.

Uso:  python3 atualizar_input_dre_rt.py [pasta]      (padrão: pasta do script)
Requer: openpyxl; LibreOffice (soffice) é opcional.
"""
import sys, os, re, glob, shutil, zipfile, subprocess, tempfile, datetime
import openpyxl
from openpyxl.utils import get_column_letter as CL
from xml.sax.saxutils import escape

SHEET_RT = 'Input_DRE  - RT'
FIRST_COL = 14                      # coluna N (jan/2027) ... AC (2031)
NPER = 16                           # 12 meses de 2027 + 2028, 2029, 2030, 2031
# itens lidos da planilha 2.Creditos (rótulo, ocorrência do rótulo na aba)
ITEMS = [('CBS sobre a receita (Depois)', 'CBS', 0), ('IBS sobre a receita (Depois)', 'IBS', 0),
         ('ISS sobre a receita (Depois)', 'ISS', 1), ('Total débitos — Depois', 'Total débitos — Depois', 0),
         ('Custo (Opex) — PIS (Antes)', 'Custo (Opex) — PIS', 0), ('Custo (Opex) — COFINS (Antes)', 'Custo (Opex) — COFINS', 0),
         ('SG&A — PIS (Antes)', 'SG&A / Despesas administrativas — PIS', 0), ('SG&A — COFINS (Antes)', 'SG&A / Despesas administrativas — COFINS', 0),
         ('Custo (Opex) — CBS (crédito Depois)', 'Custo (Opex) — CBS', 0), ('Custo (Opex) — IBS (crédito Depois)', 'Custo (Opex) — IBS', 0),
         ('SG&A — CBS (crédito Depois)', 'SG&A / Despesas administrativas — CBS', 0), ('SG&A — IBS (crédito Depois)', 'SG&A / Despesas administrativas — IBS', 0),
         ('Base: receita bruta', 'Base: receita bruta', 0)]
IT = {k: i for i, k in enumerate(['CBS', 'IBS', 'ISS', 'TOTDEB', 'CPIS', 'CCOF', 'SPIS', 'SCOF', 'CCBS', 'CIBS', 'SCBS', 'SIBS', 'RB'])}
R_ATIVO, R_COM, R_SEM = 8, 25, 42   # primeira linha de cada bloco na Base_RT
HDR_ROW = 6


def find_files(d):
    inp = [f for f in glob.glob(os.path.join(d, '1.Input*.xlsx')) if not os.path.basename(f).startswith('~')]
    cre = [f for f in glob.glob(os.path.join(d, '2.Creditos*.xlsx')) if not os.path.basename(f).startswith('~')]
    if not inp or not cre:
        sys.exit('Não encontrei 1.Input*.xlsx e 2.Creditos*.xlsx em ' + d)
    inp = sorted(inp, key=lambda f: ('_automatizado' in f, f))   # parte sempre do original
    cre = sorted(cre, key=lambda f: ('_simulacao' not in f, f), reverse=False)
    sim = sorted([f for f in cre if '_simulacao' in f], reverse=True)   # a versão mais nova (_v2 > sem sufixo)
    cre = sim + [f for f in cre if f not in sim]
    return inp[0], cre[0]


def read_consolidado(cre_path):
    wb = openpyxl.load_workbook(cre_path, read_only=True, data_only=True)
    out = {}
    for nome, key in (('Consolidado com Lote 6', 'com'), ('Consolidado sem Lote 6', 'sem')):
        rows = list(wb[nome].iter_rows(min_row=1, max_row=140, min_col=1, max_col=26, values_only=True))
        # colunas: C..N = jan..dez/27 (idx 2..13); S..V = 2028..2031 (idx 18..21)
        hdr = rows[3]
        assert hdr[2] == 'Jan/27' and hdr[13] == 'Dez/27' and hdr[18] == 2028 and hdr[21] == 2031, 'layout da aba ' + nome + ' mudou'
        vals = []
        for _, rot, occ in ITEMS:
            idx = [i for i, r in enumerate(rows) if r[1] is not None and str(r[1]).strip() == rot]
            if len(idx) <= occ:
                sys.exit('Linha não encontrada em %s: %s' % (nome, rot))
            r = rows[idx[occ]]
            vals.append([float(x) if isinstance(x, (int, float)) else 0.0 for x in list(r[2:14]) + list(r[18:22])])
        out[key] = vals
    return out


YEARS5 = [2027, 2028, 2029, 2030, 2031]
EMPRESAS = [('graos', 'GRÃOS'), ('mobi', 'CS MOBI'), ('leste', 'BLOCO LESTE'), ('mercosul', 'MERCOSUL'), ('norte', 'ESCOLAS NORTE'),
            ('sul', 'ESCOLAS SUL'), ('mcp01', 'MCP01'), ('welfare', 'WELFARE'), ('rota', 'ROTA'), ('lote6', 'LOTE 6')]
def find_bases(pasta):
    cand = [os.path.join(pasta, 'bases.json'), os.path.join(pasta, 'Automacao', 'bases.json'), os.path.join(pasta, 'simulador', 'bases.json'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'simulador', 'bases.json'), os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bases.json')]
    for c in cand:
        if os.path.exists(c):
            return c
    sys.exit('bases.json (extraído das planilhas DF) não encontrado; coloque em ' + pasta)


def read_scenario(cre_path):
    """Lê as premissas da simulação direto da aba Premissas da 2.Creditos (valores calculados)."""
    wb = openpyxl.load_workbook(cre_path, read_only=True, data_only=True)
    ws = wb['Premissas']
    rows = {r: [c for c in row] for r, row in enumerate(ws.iter_rows(min_row=1, max_row=60, max_col=18, values_only=True), 1)}
    sc = {'cbs': {}, 'ibsRef': {}, 'redutor': {}, 'redutorLoc': {}}
    for j, y in enumerate(range(2026, 2034)):
        col = 2 + j      # coluna C = índice 2
        sc['cbs'][str(y)] = float(rows[10][col]); sc['ibsRef'][str(y)] = float(rows[11][col])
        sc['redutor'][str(y)] = float(rows[12][col]); sc['redutorLoc'][str(y)] = float(rows[17][col])
    emp = {}
    for i, (key, _) in enumerate(EMPRESAS):
        v = rows[46 + i][3]
        if isinstance(v, (int, float)): emp[key] = {'credito': float(v)}
    sc['empresas'] = emp
    return sc


def read_share_and_motor(bases_path, cre_path):
    """Valores anuais do motor do simulador HTML (node) com as premissas lidas da 2.Creditos."""
    import json
    motor = {k: [0] * 5 for k in ('cbs', 'ibs', 'cred', 'l6cbs', 'l6ibs', 'l6cred')}
    here = os.path.dirname(bases_path)
    mj = os.path.join(here, 'motor.js'); cli = os.path.join(here, 'calc_cli.js')
    node = shutil.which('node')
    if node and os.path.exists(mj) and os.path.exists(cli):
        scp = os.path.join(tempfile.mkdtemp(), 'sc.json')
        json.dump({'params': read_scenario(cre_path)}, open(scp, 'w'))
        res = json.loads(subprocess.run([node, cli, bases_path, mj, scp], capture_output=True, text=True, check=True).stdout)
        for j, y in enumerate(YEARS5):
            for key, _ in EMPRESAS:
                r = res[key].get(str(y))
                if not r:
                    continue
                motor['cbs'][j] += r['det']['cbsRB']; motor['ibs'][j] += r['det']['ibsRB']; motor['cred'][j] += r['det']['credOS']
                if key == 'lote6':
                    motor['l6cbs'][j] += r['det']['cbsRB']; motor['l6ibs'][j] += r['det']['ibsRB']; motor['l6cred'][j] += r['det']['credOS']
    else:
        print('Aviso: node/motor.js/calc_cli.js não encontrados; conferência com o simulador HTML ficará sem valores de referência.')
    return motor


# ---------- helpers de XML ----------
def col_idx(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def set_cell(sheet_xml, ref, formula=None, value=None, style=None):
    """Cria/substitui a célula `ref` (fórmula e/ou valor numérico) dentro da linha existente."""
    m = re.match(r'([A-Z]+)(\d+)', ref)
    col, row = m.group(1), int(m.group(2))
    rm = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % row, sheet_xml, re.S)
    if not rm:
        raise ValueError('linha %d inexistente' % row)
    rxml = rm.group(0)
    if rxml.endswith('/>'):
        rxml = rxml[:-2] + '></row>'
    cm = re.search(r'<c r="%s"(?: [^>]*?)?(?:/>|>.*?</c>)' % ref, rxml, re.S)
    st = style
    if cm and st is None:
        sm = re.search(r' s="(\d+)"', cm.group(0).split('>')[0])
        st = sm.group(1) if sm else None
    attrs = ' s="%s"' % st if st is not None else ''
    inner = ''
    if formula is not None:
        inner += '<f>%s</f>' % escape(formula)
    if value is not None:
        inner += '<v>%r</v>' % float(value)
    new = '<c r="%s"%s>%s</c>' % (ref, attrs, inner)
    if cm:
        rxml = rxml.replace(cm.group(0), new, 1)
    else:
        pos = None
        for c2 in re.finditer(r'<c r="([A-Z]+)%d"' % row, rxml):
            if col_idx(c2.group(1)) > col_idx(col):
                pos = c2.start()
                break
        rxml = rxml[:pos] + new + rxml[pos:] if pos is not None else rxml.replace('</row>', new + '</row>')
    return sheet_xml.replace(rm.group(0), rxml, 1)


def cell_xml(ref, v=None, f=None, s=None, text=None):
    a = ' s="%d"' % s if s is not None else ''
    if text is not None:
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, a, escape(text))
    inner = ''
    if f is not None:
        inner += '<f>%s</f>' % escape(f)
    if v is not None:
        if isinstance(v, str):
            return '<c r="%s"%s t="str">%s<v>%s</v></c>' % (ref, a, inner, escape(v))
        inner += '<v>%r</v>' % float(v)
    return '<c r="%s"%s>%s</c>' % (ref, a, inner)


def build_sheet(rows, widths, freeze=None):
    """rows: {linha: [(coluna(1..), xml_da_celula)]}"""
    sd = ''
    for r in sorted(rows):
        cells = ''.join(x for _, x in sorted(rows[r], key=lambda t: t[0]))
        sd += '<row r="%d">%s</row>' % (r, cells)
    cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (a, b, w) for a, b, w in widths)
    pane = ''
    if freeze:
        pane = '<pane xSplit="1" ySplit="%d" topLeftCell="B%d" activePane="bottomRight" state="frozen"/>' % (freeze, freeze + 1)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetViews><sheetView zoomScale="85" workbookViewId="0">%s</sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>%s</sheetData></worksheet>') % (pane, cols, sd)


def per_labels(ws_rt_vals):
    labs = []
    for k in range(NPER):
        c = FIRST_COL + k
        mes = ws_rt_vals.cell(3, c).value
        ano = ws_rt_vals.cell(2, c).value
        labs.append(('%s/%s' % (str(mes)[:3], str(ano)[2:])) if k < 12 else str(ano))
    return labs


def main():
    pasta = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    inp, cre = find_files(pasta)
    print('Planilha Input  :', os.path.basename(inp))
    print('Planilha Créditos:', os.path.basename(cre))
    data = read_consolidado(cre)
    motor = read_share_and_motor(find_bases(pasta), cre)

    tmp = tempfile.mkdtemp()
    zin = zipfile.ZipFile(inp)
    files = {n: zin.read(n) for n in zin.namelist()}
    infos = {i.filename: i for i in zin.infolist()}
    order = list(zin.namelist())
    zin.close()

    wb = files['xl/workbook.xml'].decode('utf8')
    rels = files['xl/_rels/workbook.xml.rels'].decode('utf8')
    ct = files['[Content_Types].xml'].decode('utf8')
    sty = files['xl/styles.xml'].decode('utf8')

    # localizar a aba RT
    sheets = re.findall(r'<sheet name="([^"]+)" sheetId="(\d+)"(?: state="\w+")? r:id="(rId\d+)"', wb)
    rid = [r for n, s, r in sheets if n == SHEET_RT][0]
    target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    rt_path = 'xl/' + target.group(1)
    rt = files[rt_path].decode('utf8')

    # estilos novos (acrescentados no fim; os existentes não mudam)
    if 'BASE_RT_STYLES' not in sty:
        nf = int(re.search(r'<fills count="(\d+)"', sty).group(1))
        sty = re.sub(r'<fills count="\d+">', '<fills count="%d">' % (nf + 1), sty)
        sty = sty.replace('</fills>', '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>', 1)
        nx = int(re.search(r'<cellXfs count="(\d+)"', sty).group(1))
        add = ('<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'          # nx   : negrito
               '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'    # nx+1 : #,##0
               '<xf numFmtId="0" fontId="2" fillId="%d" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'  # nx+2 : chave
               '<xf numFmtId="3" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'  # nx+3 : negrito #,##0
               '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>'  # nx+4 : centro
               ) % nf
        sty = re.sub(r'<cellXfs count="\d+">', '<cellXfs count="%d">' % (nx + 5), sty)
        sty = sty.replace('</cellXfs>', add + '</cellXfs>', 1)
        sty = sty.replace('</styleSheet>', '<!--BASE_RT_STYLES--></styleSheet>')
        files['xl/styles.xml'] = sty.encode('utf8')
        S_B, S_N, S_KEY, S_NB, S_C = nx, nx + 1, nx + 2, nx + 3, nx + 4
    else:
        nx = int(re.search(r'<cellXfs count="(\d+)"', sty).group(1)) - 5
        S_B, S_N, S_KEY, S_NB, S_C = nx, nx + 1, nx + 2, nx + 3, nx + 4

    # valores cacheados da aba RT para rotular períodos
    wbv = openpyxl.load_workbook(inp, data_only=True)
    labs = per_labels(wbv[SHEET_RT])

    # ---------- aba Base_RT ----------
    rows = {}
    def put(r, c, xml): rows.setdefault(r, []).append((c, xml))
    put(1, 1, cell_xml('A1', text='Base_RT — dados da Reforma Tributária (R$ mil) que alimentam a aba "Input_DRE  - RT"', s=S_B))
    put(2, 1, cell_xml('A2', text='Incluir Lote 6 na simulação? (1 = sim, 0 = não)', s=S_B))
    put(2, 2, cell_xml('B2', v=1, s=S_KEY))
    put(3, 1, cell_xml('A3', text='Fonte: %s, abas "Consolidado com Lote 6" e "Consolidado sem Lote 6" (valores calculados). Gerado por atualizar_input_dre_rt.py em %s.' % (os.path.basename(cre), datetime.date.today().strftime('%d/%m/%Y'))))
    put(4, 1, cell_xml('A4', text='Não digite nos blocos COM / SEM: eles são reescritos pelo script. A chave B2 escolhe qual bloco alimenta o Input (bloco ATIVO).'))
    for blk, r0, titulo in (('ativo', R_ATIVO, 'ATIVO (usado pelo Input_DRE  - RT) — segue a chave B2'), ('com', R_COM, 'COM Lote 6 (aba "Consolidado com Lote 6")'), ('sem', R_SEM, 'SEM Lote 6 (aba "Consolidado sem Lote 6")')):
        put(r0 - 2, 1, cell_xml('A%d' % (r0 - 2), text=titulo, s=S_B))
        put(r0 - 1, 1, cell_xml('A%d' % (r0 - 1), text='R$ mil', s=S_B))
        for k in range(NPER):
            put(r0 - 1, 2 + k, cell_xml('%s%d' % (CL(2 + k), r0 - 1), text=labs[k], s=S_B))
        for i, (nome, _, _) in enumerate(ITEMS):
            r = r0 + i
            put(r, 1, cell_xml('A%d' % r, text=nome))
            for k in range(NPER):
                ref = '%s%d' % (CL(2 + k), r)
                if blk == 'ativo':
                    f = 'IF($B$2=1,%s%d,%s%d)' % (CL(2 + k), R_COM + i, CL(2 + k), R_SEM + i)
                    v = data['com' if True else 'sem'][i][k]
                    put(r, 2 + k, cell_xml(ref, f=f, v=v, s=S_N))
                else:
                    put(r, 2 + k, cell_xml(ref, v=data[blk][i][k], s=S_N))
    base_xml = build_sheet(rows, [(1, 1, 44), (2, 17, 13)], freeze=HDR_ROW)

    # ---------- aba RT: fórmulas ----------
    def B(item, k):  # referência à Base_RT (bloco ativo)
        return 'Base_RT!%s%d' % (CL(2 + k), R_ATIVO + IT[item])
    nrt = 0
    for k in range(NPER):
        c = CL(FIRST_COL + k)
        d = data['com']
        def v(item): return d[IT[item]][k]
        fm = {
            97: ('%s*1000' % B('CBS', k), v('CBS') * 1000),
            98: ('%s*1000' % B('IBS', k), v('IBS') * 1000),
            14: ('%s*1000' % B('ISS', k), v('ISS') * 1000),
            102: ('-(%s+%s)*1000' % (B('CPIS', k), B('CCOF', k)), -(v('CPIS') + v('CCOF')) * 1000),
            103: ('%s*1000' % B('CCBS', k), v('CCBS') * 1000),
            104: ('%s*1000' % B('CIBS', k), v('CIBS') * 1000),
            109: ('%s*1000' % B('SCBS', k), v('SCBS') * 1000),
            110: ('%s*1000' % B('SIBS', k), v('SIBS') * 1000),
        }
        for r, (f, val) in fm.items():
            rt = set_cell(rt, '%s%d' % (c, r), formula=f, value=val); nrt += 1
        # linhas do orçamento original (Input_DRE) copiadas à mão -> ligadas à aba Input_DRE
        for r, src in ((96, 5), (101, 36), (108, 43), (117, 55), (6, 6), (37, 37), (47, 47), (51, 51), (52, 52)):
            rt = set_cell(rt, '%s%d' % (c, r), formula='Input_DRE!%s%d' % (c, src), value=wbv[SHEET_RT].cell(r, FIRST_COL + k).value); nrt += 1
        # IR: mesma alíquota efetiva do orçamento original sobre o LAIR da Reforma (sem circularidade)
        rt = set_cell(rt, '%s55' % c, formula='Input_DRE!%s77*(SUM(%s5:%s54)+SUM(%s56:%s67))' % (c, c, c, c, c), value=wbv[SHEET_RT].cell(55, FIRST_COL + k).value); nrt += 1
    files[rt_path] = rt.encode('utf8')

    # ---------- aba Check_RT ----------
    rows = {}
    rrt = "'%s'!" % SHEET_RT
    put(1, 1, cell_xml('A1', text='Check_RT — conferência da aba "Input_DRE  - RT" (R$; diferenças devem ser zero)', s=S_B))
    put(2, 1, cell_xml('A2', text='Resultado geral', s=S_B))
    put(3, 1, cell_xml('A3', text='Lote 6 incluído? (chave Base_RT!B2)'))
    put(3, 2, cell_xml('B3', f='IF(Base_RT!B2=1,"SIM","NÃO")', v='SIM', s=S_C))
    put(HDR_ROW - 1, 1, cell_xml('A5', text='1) Conferências (tolerância R$ 1)', s=S_B))
    put(HDR_ROW, 1, cell_xml('A6', text='Item', s=S_B))
    for k in range(NPER):
        put(HDR_ROW, 2 + k, cell_xml('%s6' % CL(2 + k), text=labs[k], s=S_B))
    put(HDR_ROW, 18, cell_xml('R6', text='Status', s=S_B))
    checks = [
        ('C1 Débitos Depois: CBS+IBS+ISS (Base) − Total débitos Depois (2.Creditos)', lambda k, c: '(%s+%s+%s-%s)*1000' % (B('CBS', k), B('IBS', k), B('ISS', k), B('TOTDEB', k))),
        ('C2 Receita RT − Receita orçamento − (CBS+IBS)', lambda k, c: '%s%s5-Input_DRE!%s5-(%s%s97+%s%s98)' % (rrt, c, c, rrt, c, rrt, c)),
        ('C3 Custo RT − Custo orçamento − (crédito antigo + CBS + IBS)', lambda k, c: '%s%s36-Input_DRE!%s36-(%s%s102+%s%s103+%s%s104)' % (rrt, c, c, rrt, c, rrt, c, rrt, c)),
        ('C4 Desp. administrativas RT − orçamento − (crédito CBS + IBS)', lambda k, c: '%s%s43-Input_DRE!%s43-(%s%s109+%s%s110)' % (rrt, c, c, rrt, c, rrt, c)),
        ('C5 ISS RT (aba) − ISS da Base_RT', lambda k, c: '%s%s14-%s*1000' % (rrt, c, B('ISS', k))),
        ('C6 Alíquota efetiva IR RT − orçamento (mesma alíquota)', lambda k, c: '(%s%s77-Input_DRE!%s77)*1000000' % (rrt, c, c)),
        ('C7 Crédito antigo de SG&A (PIS+COFINS) não mapeado deve ser zero', lambda k, c: '(%s+%s)*1000' % (B('SPIS', k), B('SCOF', k))),
        ('C8 Lucro líquido RT = LAIR RT + IR RT', lambda k, c: '%s%s72-(%s%s74+%s%s55)' % (rrt, c, rrt, c, rrt, c)),
    ]
    for i, (nome, fn) in enumerate(checks):
        r = HDR_ROW + 1 + i
        put(r, 1, cell_xml('A%d' % r, text=nome))
        for k in range(NPER):
            c = CL(FIRST_COL + k)
            put(r, 2 + k, cell_xml('%s%d' % (CL(2 + k), r), f=fn(k, c), v=0, s=S_N))
        put(r, 18, cell_xml('R%d' % r, f='IF(SUMPRODUCT(ABS(B%d:Q%d))<1,"OK","DIVERGE")' % (r, r), v='OK', s=S_C))
    last = HDR_ROW + len(checks)
    r = last + 1
    put(r, 1, cell_xml('A%d' % r, text='Info: Receita bruta da Base (consolidado) ÷ Receita do orçamento − 1 (escopos diferentes; só informativo)'))
    for k in range(NPER):
        c = CL(FIRST_COL + k)
        put(r, 2 + k, cell_xml('%s%d' % (CL(2 + k), r), f='IFERROR(%s*1000/Input_DRE!%s5-1,0)' % (B('RB', k), c), v=0, s=None))
    put(r, 18, cell_xml('R%d' % r, f='"info"', v='info', s=S_C))
    r_ok = r
    r += 1
    put(r, 1, cell_xml('A%d' % r, text='Info: contribuição do Lote 6 (COM − SEM) em débitos CBS+IBS+ISS, R$ mil'))
    for k in range(NPER):
        put(r, 2 + k, cell_xml('%s%d' % (CL(2 + k), r), f='(Base_RT!%s%d-Base_RT!%s%d)' % (CL(2 + k), R_COM + IT['TOTDEB'], CL(2 + k), R_SEM + IT['TOTDEB']), v=0, s=S_N))
    put(r, 18, cell_xml('R%d' % r, f='"info"', v='info', s=S_C))
    r_ok = r - 1
    r_last = r
    # conferência com o simulador HTML (motor), anual em R$ mil
    r = r_last + 2
    put(r, 1, cell_xml('A%d' % r, text='2) Conferência com o simulador HTML (motor), anual em R$ mil — alíquotas da simulação', s=S_B))
    r += 1
    put(r, 1, cell_xml('A%d' % r, text='Item', s=S_B))
    for j, y in enumerate(YEARS5):
        put(r, 2 + j, cell_xml('%s%d' % (CL(2 + j), r), text=str(y), s=S_B))
    put(r, 7, cell_xml('G%d' % r, text='Status', s=S_B))
    hdr = r
    simrows = []
    def annual_f(item_row, j):
        if j == 0: return 'SUM(Base_RT!B%d:M%d)' % (item_row, item_row)
        return 'Base_RT!%s%d' % (CL(13 + j), item_row)
    for i, (nome, key, itemkey) in enumerate([('CBS sobre a receita', 'cbs', 'CBS'), ('IBS sobre a receita', 'ibs', 'IBS'), ('Créditos Custo+SG&A (CBS+IBS)', 'cred', None)]):
        r1 = hdr + 1 + i * 3
        put(r1, 1, cell_xml('A%d' % r1, text=nome + ' — simulador HTML'))
        put(r1 + 1, 1, cell_xml('A%d' % (r1 + 1), text='    planilha (Base_RT)'))
        put(r1 + 2, 1, cell_xml('A%d' % (r1 + 2), text='    diferença'))
        for j in range(5):
            put(r1, 2 + j, cell_xml('%s%d' % (CL(2 + j), r1), f='%r-(1-Base_RT!$B$2)*%r' % (motor[key][j], motor['l6' + key][j]), v=motor[key][j], s=S_N))
            f = annual_f(R_ATIVO + IT[itemkey], j) if itemkey else '+'.join(annual_f(R_ATIVO + IT[x], j) for x in ('CCBS', 'CIBS', 'SCBS', 'SIBS'))
            put(r1 + 1, 2 + j, cell_xml('%s%d' % (CL(2 + j), r1 + 1), f=f, v=0, s=S_N))
            put(r1 + 2, 2 + j, cell_xml('%s%d' % (CL(2 + j), r1 + 2), f='%s%d-%s%d' % (CL(2 + j), r1 + 1, CL(2 + j), r1), v=0, s=S_N))
        put(r1 + 2, 7, cell_xml('G%d' % (r1 + 2), f='IF(SUMPRODUCT(ABS(B%d:F%d))<1,"OK","DIVERGE")' % (r1 + 2, r1 + 2), v='OK', s=S_C))
        simrows.append(r1 + 2)
    r_last = simrows[-1]
    put(2, 2, cell_xml('B2', f='IF(AND(COUNTIF(R%d:R%d,"OK")=%d,%s),"TUDO OK","VERIFICAR")' % (HDR_ROW + 1, r_ok, len(checks), ','.join('G%d="OK"' % x for x in simrows)), v='TUDO OK', s=S_KEY))
    # resumo dos efeitos
    r0 = r_last + 3
    put(r0 - 1, 1, cell_xml('A%d' % (r0 - 1), text='2) Efeitos da simulação (Reforma − orçamento original), R$ mil', s=S_B))
    put(r0, 1, cell_xml('A%d' % r0, text='R$ mil', s=S_B))
    for k in range(NPER):
        put(r0, 2 + k, cell_xml('%s%d' % (CL(2 + k), r0), text=labs[k], s=S_B))
    put(r0, 18, cell_xml('R%d' % r0, text='Total 2027', s=S_B))
    efeitos = [
        ('Δ Receita (CBS + IBS dentro da receita)', lambda c: '(%s%s5-Input_DRE!%s5)/1000' % (rrt, c, c)),
        ('Δ Deduções (ISS novo − PIS/COFINS/ISS atuais)', lambda c: '(%s%s14-(Input_DRE!%s12+Input_DRE!%s13+Input_DRE!%s14))/1000' % (rrt, c, c, c, c)),
        ('Δ Custo de vendas e serviços (créditos novos − créditos antigos)', lambda c: '(%s%s36-Input_DRE!%s36)/1000' % (rrt, c, c)),
        ('Δ Despesas administrativas (créditos novos)', lambda c: '(%s%s43-Input_DRE!%s43)/1000' % (rrt, c, c)),
        ('Δ LAIR (resultado antes do IR)', lambda c: '(%s%s74-Input_DRE!%s74)/1000' % (rrt, c, c)),
        ('Δ IR/CS', lambda c: '(%s%s55-Input_DRE!%s55)/1000' % (rrt, c, c)),
        ('Δ Lucro líquido', lambda c: '(%s%s72-Input_DRE!%s72)/1000' % (rrt, c, c)),
    ]
    for i, (nome, fn) in enumerate(efeitos):
        r = r0 + 1 + i
        put(r, 1, cell_xml('A%d' % r, text=nome, s=S_B if i == len(efeitos) - 1 else None))
        for k in range(NPER):
            put(r, 2 + k, cell_xml('%s%d' % (CL(2 + k), r), f=fn(CL(FIRST_COL + k)), v=0, s=S_N))
        put(r, 18, cell_xml('R%d' % r, f='SUM(B%d:M%d)' % (r, r), v=0, s=S_NB))
    chk_xml = build_sheet(rows, [(1, 1, 70), (2, 17, 13), (18, 18, 12)], freeze=HDR_ROW)

    # ---------- registrar as abas novas ----------
    nmax = max(int(re.search(r'sheet(\d+)\.xml', n).group(1)) for n in files if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
    sid = max(int(s) for _, s, _ in sheets)
    rmax = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels))
    for nm in ('Base_RT', 'Check_RT'):
        if re.search(r'<sheet name="%s"' % nm, wb):
            rid2 = re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"' % nm, wb).group(1)
            tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid2, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid2, rels)
            files['xl/' + tgt.group(1)] = (base_xml if nm == 'Base_RT' else chk_xml).encode('utf8')
        else:
            nmax += 1; sid += 1; rmax += 1
            p = 'xl/worksheets/sheet%d.xml' % nmax
            files[p] = (base_xml if nm == 'Base_RT' else chk_xml).encode('utf8')
            order.append(p)
            wb = wb.replace('</sheets>', '<sheet name="%s" sheetId="%d" r:id="rId%d"/></sheets>' % (nm, sid, rmax))
            rels = rels.replace('</Relationships>', '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rmax, nmax))
            ct = ct.replace('</Types>', '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % p)
    # recalcular tudo ao abrir; descartar a cadeia de cálculo (o Excel a reconstrói)
    wb = re.sub(r'<calcPr ([^>]*?)/>', lambda m: '<calcPr %s fullCalcOnLoad="1"/>' % m.group(1) if 'fullCalcOnLoad' not in m.group(1) else m.group(0), wb)
    if 'xl/calcChain.xml' in files:
        del files['xl/calcChain.xml']; order.remove('xl/calcChain.xml')
        rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', rels)
        ct = re.sub(r'<Override [^>]*calcChain[^>]*/>', '', ct)
    files['xl/workbook.xml'] = wb.encode('utf8'); files['xl/_rels/workbook.xml.rels'] = rels.encode('utf8'); files['[Content_Types].xml'] = ct.encode('utf8')

    out = os.path.join(pasta, os.path.basename(inp).replace('.xlsx', '_automatizado_v2.xlsx') if '_automatizado' not in inp else os.path.basename(inp))
    stage = os.path.join(tmp, 'stage.xlsx')
    with zipfile.ZipFile(stage, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, files[n])
    print('Fórmulas gravadas na aba "%s": %d células' % (SHEET_RT, nrt))

    # recalcular com o LibreOffice (se houver) e gravar os valores calculados nas fórmulas
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        try:
            # remove os valores em cache das fórmulas para obrigar o LibreOffice a recalcular tudo
            nov = os.path.join(tmp, 'novals.xlsx')
            with zipfile.ZipFile(nov, 'w', zipfile.ZIP_DEFLATED) as z:
                for n in order:
                    b = files[n]
                    if n.startswith('xl/worksheets/sheet'):
                        b = re.sub(r'(<f>.*?</f>)<v>.*?</v>', r'\1', b.decode('utf8'), flags=re.S).encode('utf8')
                    z.writestr(n, b)
            os.makedirs(os.path.join(tmp, 'lo'), exist_ok=True)
            subprocess.run([soffice, '--headless', '--convert-to', 'xlsx', '--outdir', os.path.join(tmp, 'lo'), nov], check=True, capture_output=True, timeout=300)
            calc = openpyxl.load_workbook(os.path.join(tmp, 'lo', 'novals.xlsx'), data_only=True)
            files = inject_values(files, order, calc, sheets_to_fix=[(rt_path, SHEET_RT), (None, 'Base_RT'), (None, 'Check_RT')], wb=files['xl/workbook.xml'].decode('utf8'), rels=files['xl/_rels/workbook.xml.rels'].decode('utf8'))
            print('Valores calculados gravados (LibreOffice).')
        except Exception as e:
            print('Aviso: não consegui recalcular via LibreOffice (%s); o Excel recalcula ao abrir.' % e)
    with zipfile.ZipFile(stage, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, files[n])
    shutil.copyfile(stage, out)
    print('Gerado:', out)


def inject_values(files, order, calc, sheets_to_fix, wb, rels):
    """Copia os valores calculados pelo LibreOffice para as tags <v> das células com fórmula."""
    for _, nome in sheets_to_fix:
        rid = re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"' % re.escape(nome), wb).group(1)
        tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
        p = 'xl/' + tgt.group(1)
        xml = files[p].decode('utf8')
        ws = calc[nome]

        def fix(m):
            ref = m.group(1)
            val = ws[ref].value
            head, f = m.group(2), m.group(3)
            if val is None:
                return m.group(0)
            if isinstance(val, str):
                head2 = re.sub(r' t="\w+"', '', head) + ' t="str"'
                return '<c r="%s"%s>%s<v>%s</v></c>' % (ref, head2, f, escape(val))
            head2 = re.sub(r' t="\w+"', '', head)
            return '<c r="%s"%s>%s<v>%r</v></c>' % (ref, head2, f, float(val))
        xml = re.sub(r'<c r="([A-Z]+\d+)"((?: [^>]*?)?)>(<f>.*?</f>)(?:<v>.*?</v>)?</c>', fix, xml, flags=re.S)
        files[p] = xml.encode('utf8')
    return files


if __name__ == '__main__':
    main()
