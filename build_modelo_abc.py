#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_modelo_abc.py — Aguiar Advisory
=====================================
Gera o MODELO FINANCEIRO PADRÃO "Grupo ABC" (piloto anonimizado).

Origem dos dados : consolidação real de um grupo de 4 empresas (fonte passada
                   como 1º argumento). Nomes e valores são ANONIMIZADOS:
                   empresas viram ALFA/BETA/GAMA/DELTA e valores são × 0,85.

Arquitetura (5 camadas):
  1. Balancete  — base única de saldos: Empresa × Conta × Período (fonte Power BI)
  2. DePara     — mapa de contas: conta → linha de DRE / linha de BP / código de nota
  3. Eliminacao — lançamentos de eliminação intercompany (Conta × Período)
  4. Quadros    — Balanco, DRE, DF_Mensal, EBITDA (100% fórmula, protegidos)
  5. CAPA/Notas — navegação, índice de notas

Regras de ouro:
  - Chave universal: EMPRESA + PERÍODO + CONTA
  - Quadros 100% por fórmula (SUMIFS/VLOOKUP) — nada digitado
  - Consolidado = Combinado − Eliminações
  - Check de amarração em todos os quadros

Uso:
  python3 build_modelo_abc.py [fonte.xlsx] [saida.xlsx]
"""

import sys, re
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule
from collections import Counter, defaultdict

# ------------------------------------------------------------------ constantes
SCALE       = 0.85
DEFAULT_SRC = 'uploads/f9357021-DERMAGE_CONSOLIDACAO_-_2025VFinal_1_.xlsx'
DEFAULT_OUT = 'artifacts/Modelo_Financeiro_Grupo_ABC.xlsx'

# Cores padrão Aguiar Advisory
NAVY, NAVY2  = '0B1526', '1C2E4A'
GOLD, GOLD_L = 'C9A227', 'F5EDD6'
GRAY_L       = 'EEF1F6'
GREEN, RED   = '1E7B34', 'B00020'
BLUE_IN      = '0000FF'

F_TITLE = Font(name='Arial', size=16, bold=True, color=NAVY)
F_SUB   = Font(name='Arial', size=10, italic=True, color='555555')
F_H     = Font(name='Arial', size=10, bold=True, color='FFFFFF')
F_SEC   = Font(name='Arial', size=10, bold=True, color='FFFFFF')
F_LIN   = Font(name='Arial', size=10)
F_TOT   = Font(name='Arial', size=10, bold=True, color=NAVY)
F_IN    = Font(name='Arial', size=10, color=BLUE_IN)
F_SM    = Font(name='Arial', size=8, color='888888')
F_CHK   = Font(name='Arial', size=10, bold=True, color=GREEN)

FILL_NAVY  = PatternFill('solid', fgColor=NAVY)
FILL_NAVY2 = PatternFill('solid', fgColor=NAVY2)
FILL_GOLD  = PatternFill('solid', fgColor=GOLD)
FILL_GOLDL = PatternFill('solid', fgColor=GOLD_L)
FILL_GRAYL = PatternFill('solid', fgColor=GRAY_L)

B_TOP = Border(top=Side(style='medium', color=NAVY))
FMT_V  = '#,##0;(#,##0)'
FMT_V2 = '#,##0.00;(#,##0.00)'

MESES36    = [f'{y}-{m:02d}' for y in (2023, 2024, 2025) for m in range(1, 13)]
MESES24    = [f'{y}-{m:02d}' for y in (2024, 2025) for m in range(1, 13)]
MESES_ELIM = [f'2025-{m:02d}' for m in range(5, 13)]
LABEL_MES  = {p: f"{['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][int(p[5:7])-1]}/{p[2:4]}"
              for p in MESES24}

EMP = {'ALFA': 'Klibra consolidado_Final', 'BETA': 'Technopharma consolidado_Final',
       'GAMA': 'Distriprime consolidado_Final', 'DELTA': 'Cosmoprime consolidado_Final'}

DRE_SET = {'ROL', 'CMV', 'Despesas com Vendas', 'Despesas Comerciais',
           'Despesas Gerais e Administrativas', 'Equivalência Patrimonial',
           'Outras Despesas Operacionais, Líquida', 'Despesas Financeiras',
           'Receitas Financeiras', 'Imposto de Renda'}

# ------------------------------------------------------------- anonimização
SUBS = [
    (r'GRUPO DERMAGE', 'GRUPO ABC'), (r'DERMAGE', 'ABC'),
    (r'KLIBRA', 'ALFA'), (r'Klibra', 'Alfa'),
    (r'TECNOPHARMA', 'BETA'), (r'Tecnopharma', 'Beta'),
    (r'DISTRIPRIME', 'GAMA'), (r'Distriprime', 'Gama'), (r'\bDISTRI\b', 'GAMA'),
    (r'COSMOPRIME', 'DELTA'), (r'Cosmoprime', 'Delta'), (r'\bCOSMO\b', 'DELTA'),
    (r'PHARMAGE', 'LOJA MODELO'), (r'EUROFARMA', 'PLANO REFERENCIA'),
    (r'BRADESCO', 'BANCO B1'), (r'\bITAU\b', 'BANCO B2'), (r'ITAÚ', 'BANCO B2'),
    (r'HSBC', 'BANCO B3'), (r'SANTANDER', 'BANCO B4'), (r'CREDIT SUISSE', 'BANCO B5'),
    (r'BARRA SHOPPING', 'LOJA 05'), (r'\bTIJUCA\b', 'LOJA 01'), (r'\bIPANEMA\b', 'LOJA 02'),
    (r'\bBOTAFOGO\b', 'LOJA 03'), (r'\bBONSUCESSO\b', 'LOJA 04'), (r'\bCOPACABANA\b', 'LOJA 06'),
    (r'\bGAVEA\b', 'LOJA 07'), (r'GÁVEA', 'LOJA 07'), (r'\bICARAI\b', 'LOJA 08'),
    (r'ICARAÍ', 'LOJA 08'), (r'\bCENTRO 2\b', 'LOJA 10'), (r'\bCENTRO\b', 'LOJA 09'),
    (r'\bBRASILIA\b', 'LOJA 11'), (r'BELO HORIZONTE', 'LOJA 12'), (r'PAGCORP', 'ESCRITORIO'),
]
BLACKLIST = ['DERMAGE', 'KLIBRA', 'TECNOPHARMA', 'DISTRIPRIME', 'COSMOPRIME', 'BRAUN',
             'BRADESCO', 'ITAU', 'HSBC', 'SANTANDER', 'SUISSE', 'TIJUCA', 'IPANEMA',
             'BOTAFOGO', 'BONSUCESSO', 'COPACABANA', 'GAVEA', 'ICARAI', 'PAGCORP',
             'EUROFARMA', 'PHARMAGE', 'RIO DE JANEIRO', 'DUQUE']

def sanitize(s):
    if not s:
        return ''
    s = re.sub(r'\d[\d.\-]{3,}', 'XXXX', str(s))          # mascara agências/contas
    for pat, rep in SUBS:
        s = re.sub(pat, rep, s)
    return s.strip()

def audit(strings, tag):
    hits = set()
    for s in strings:
        up = str(s).upper()
        for w in BLACKLIST:
            if w in up:
                hits.add(f'{tag}: "{w}" em "{str(s)[:60]}"')
    return hits

# ------------------------------------------------------------------ estilos
def head(ws, r, headers, height=26):
    for j, h in enumerate(headers, 1):
        c = ws.cell(r, j, h)
        c.font, c.fill = F_H, FILL_NAVY
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = height

def title_block(ws, t1, t2):
    ws['A1'] = t1; ws['A1'].font = F_TITLE
    ws['A2'] = t2; ws['A2'].font = F_SUB

def fmt_range(ws, ref, fmt=FMT_V):
    for row in ws[ref]:
        for c in row:
            c.number_format = fmt
            c.font = F_LIN

# ------------------------------------------------------------------ extração
def extrair(src):
    """Lê a fonte e devolve (depara, balancete). Valores já escalados ×0,85."""
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)

    # --- DePara: colunas de saldo mensal (Jan24..Dez25) e eliminação (Mai25..Dez25)
    # blocos (saldo, elim, cons): Mai=(Y,Z,AA) Jun=(AB,AC,AD) Jul=(AG,AH,AI)
    # Ago=(AL,AM,AN) Set=(AQ,AR,AS) Out=(AV,AW,AX) Nov=(BA,BB,BC) Dez=(BF,BG,BH)
    # (AF/AK/AP/AU/AZ/BE são colunas-espaçadora com fórmula de data)
    SALDO_IDX = list(range(8, 24)) + [24, 27, 32, 37, 42, 47, 52, 57]
    ELIM_IDX  = [25, 28, 33, 38, 43, 48, 53, 58]
    ws = wb['DE X PARA DERMAGE']
    depara = []
    for row in ws.iter_rows(min_row=2, max_row=696, max_col=63, values_only=True):
        if row[0] is None:
            continue
        d = {'conta': str(row[0]).strip(),
             'desc': str(row[2]).strip() if row[2] else '',
             'dre':  str(row[3]).strip() if row[3] else '',
             'bp':   str(row[4]).strip() if row[4] else '',
             'df':   str(row[5]).strip() if row[5] else ''}
        d['saldos'] = {MESES24[i]: (row[SALDO_IDX[i]] if isinstance(row[SALDO_IDX[i]], (int, float)) else 0.0)
                       for i in range(24)}
        d['elim'] = {MESES_ELIM[i]: (row[ELIM_IDX[i]] if isinstance(row[ELIM_IDX[i]], (int, float)) else 0.0)
                     for i in range(8)}
        depara.append(d)

    # --- Balancetes por empresa (somente contas do DePara; 36 meses Jan23..Dez25)
    # DEDUPLICAÇÃO: a origem às vezes lista a mesma conta 2× no mesmo mês —
    # somamos (o SUMIFS do Excel também soma). Garante 1 linha por (empresa, conta, mês).
    contas_dep = {d['conta'] for d in depara}
    acc = {}
    for emp, sheet in EMP.items():
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=6, max_row=ws.max_row, min_col=3, max_col=40, values_only=True):
            if row[0] is None:
                continue
            c = str(row[0]).strip()
            if c not in contas_dep:
                continue
            for i, mes in enumerate(MESES36):
                v = row[2 + i] if len(row) > 2 + i else None
                if isinstance(v, (int, float)) and abs(v) > 1e-9:
                    k = (emp, c, mes)
                    acc[k] = acc.get(k, 0.0) + v * SCALE
    balancete = [(e, c, mes, v) for (e, c, mes), v in acc.items()]
    return depara, balancete

# ------------------------------------------------------------ pré-verificação
def preflight(depara, balancete):
    print('--- PRÉ-VERIFICAÇÃO ---')
    unc = [d for d in depara if d['dre'] in DRE_SET and not d['df'].startswith(('U', 'O'))]
    for d in unc:
        d['df'] = 'U2.15'   # Outras Despesas e Receitas Líquidas
    print(f'1) Contas de resultado sem nota -> U2.15: {len(unc)}')
    for d in unc[:10]:
        print(f'   {d["conta"]} {d["desc"][:35]} (DRE={d["dre"][:25]})')
    r411 = next((d for d in depara if d['conta'] == '41111002'), None)
    if r411:
        print('2) Eliminação 41111002 por mês (crescente => YTD):',
              {k: round(v) for k, v in r411['elim'].items() if abs(v) > 0.5})
    contas_dep = {d['conta'] for d in depara}
    for emp in EMP:
        s = {k: 0.0 for k in '1234'}
        for e, c, mes, v in balancete:
            if e == emp and mes == '2025-12' and c in contas_dep:
                s[c[0]] += v
        # balancete: ativo = passivo + receitas + despesas (passivo positivo, despesas negativas)
        res = s['1'] - s['2'] - s['4'] - s['3']
        print(f'3) {emp}: Ativo−Passivo−Rec−Desp = {res:,.2f} '
              f'({"OK" if abs(res) < 10 else "VERIFICAR"})')
    cods = {d['df'] for d in depara if d['df'].startswith(('U', 'O'))}
    print(f'4) Códigos de nota em uso: {len(cods)}')

# ------------------------------------------------------------- normalização
AJUSTE_CONTA = '29999999'

def normalizar(depara, balancete):
    """Neutraliza o resultado do ano JÁ apropriado no PL (duplicidade P&L ×
    lucros acumulados). Onde a identidade Ativo = Passivo + Rec + Desp não
    fecha, lança entrada na conta 29999999 (linha 'Lucros acumulados') com
    valor = resíduo. Dinâmico: funciona para qualquer fonte/mês."""
    dep = {d['conta']: d for d in depara}
    from collections import defaultdict
    s = defaultdict(lambda: {k: 0.0 for k in '1234'})
    for e, c, mes, v in balancete:
        if c in dep:
            s[(e, mes)][c[0]] += v
    ajustes = []
    for (e, mes), g in sorted(s.items()):
        gap = g['1'] - g['2'] - g['4'] - g['3']
        if abs(gap) > 1:
            ajustes.append((e, mes, gap))
    if ajustes:
        depara.append({'conta': AJUSTE_CONTA,
                       'desc': 'AJUSTE — RESULTADO JA APROPRIADO NO PL',
                       'dre': '', 'bp': 'Lucros acumulados', 'df': 'T3',
                       'saldos': {}, 'elim': {}})
        for e, mes, gap in ajustes:
            balancete.append((e, AJUSTE_CONTA, mes, gap))
        print(f'--- NORMALIZAÇÃO: {len(ajustes)} ajuste(s) '
              f'(resultado já apropriado no PL) ---')
        for e, mes, gap in ajustes:
            print(f'   {e} {mes}: {gap:,.2f}')
    return depara, balancete

def niveis(conta, tamanhos=(1, 2, 3, 4, 5)):
    """Hierarquia da conta para drill-down no Power BI (Domínio: 1-2-3-4-5-8)."""
    out, pos = [], 0
    for t in tamanhos:
        if pos >= len(conta):
            break
        out.append(conta[:pos + t])
        pos += t
    if pos < len(conta):
        out.append(conta)
    return out

# ------------------------------------------------------------------ estruturas
BP_LINES = [
    ('SEC', 'ATIVO', None), ('SUB', 'Circulante', None),
    ('L', 'Caixa e equivalentes de caixa', +1), ('L', 'Títulos e valores mobiliários', +1),
    ('L', 'Contas a receber de clientes', +1), ('L', 'Estoques', +1),
    ('L', 'Impostos Recuperar', +1), ('L', 'Adiantamento a fornecedores', +1),
    ('L', 'Outros Ativos', +1),
    ('TOT', 'Total do Ativo Circulante', None),
    ('SUB', 'Não Circulante', None),
    ('L', 'Depósitos Judiciais', +1), ('L', 'Partes Relacionadas', +1),
    ('L', 'Investimentos', +1), ('L', 'Imobilizado', +1),
    ('L', 'Direito de Uso', +1), ('L', 'Intangível', +1),
    ('TOT', 'Total do Ativo Não Circulante', None),
    ('TOT', 'Total do Ativo', None),
    ('SEC', 'PASSIVO E PATRIMÔNIO LÍQUIDO', None), ('SUB', 'Passivo Circulante', None),
    ('L', 'Fornecedores', -1), ('L', 'Empréstimos', -1),
    ('L', 'Instrumentos financeiros', -1),
    ('L', 'impostos, taxas e contribuições a recolher', -1),
    ('L', 'Imposto de renda e contribuição social a pagar', -1),
    ('L', 'Adiantamento de clientes', -1), ('L', 'Salários e encargos Sociais', -1),
    ('L', 'Obrigações sociais e trabalhistas', -1), ('L', 'Outras contas a pagar', -1),
    ('L', 'Arrendamento Curto Prazo', -1), ('L', 'Outros Passivos', -1),
    ('TOT', 'Total do Passivo Circulante', None),
    ('SUB', 'Passivo Não Circulante', None),
    ('L', 'Arrendamento Longo Prazo', -1), ('L', 'Subvenção para investimentos', -1),
    ('L', 'Obrigações Tributárias', -1), ('L', 'Provisão para contingência', -1),
    ('TOT', 'Total do Passivo Não Circulante', None),
    ('SUB', 'Patrimônio Líquido', None),
    ('L', 'Capital social', -1), ('L', 'Ajuste de avaliação patrimonial', -1),
    ('L', 'Lucros acumulados', -1),
    ('TOT', 'Total do Patrimônio Líquido', None),
    ('TOT', 'TOTAL DO PASSIVO + PATRIMÔNIO LÍQUIDO', None),
    ('CHK', 'CHECK: Ativo − Passivo − PL (deve ser ~0)', None),
]

DRE_LINES = [
    ('L', 'ROL', None), ('L', 'CMV', None), ('SUB', 'Lucro Bruto', None),
    ('L', 'Despesas com Vendas', None), ('L', 'Despesas Comerciais', None),
    ('L', 'Despesas Gerais e Administrativas', None), ('L', 'Equivalência Patrimonial', None),
    ('L', 'Outras Despesas Operacionais, Líquida', None),
    ('SUB', 'Lucro antes do Resultado', None),
    ('L', 'Despesas Financeiras', None), ('L', 'Receitas Financeiras', None),
    ('SUB', 'Lucro Antes do IR / CS', None), ('L', 'Imposto de Renda', None),
    ('TOT', 'Lucro Líquido', None),
    ('CHK', 'CHECK: LL (DRE) − Σ LL mensal DF (deve ser ~0)', None),
]

DF_LINES = [
    ('L', 'U1.1', 'Receita Líquida'), ('L', 'U1.2', 'Devoluções'), ('L', 'U1.3', 'Impostos sobre Vendas'),
    ('SUB', None, 'Receita Operacional Líquida'),
    ('L', 'U2.1', 'Custos com Mercadorias Vendidas'),
    ('SUB', None, 'Lucro Bruto'),
    ('L', 'U2.2', 'Despesas com Pessoal'), ('L', 'U2.3', 'Serviços de Terceiros'),
    ('L', 'U2.5', 'Propaganda e Publicidade'), ('L', 'U2.6', 'Despesas com Vendas'),
    ('L', 'U2.7', 'Ocupação'), ('L', 'U2.8', 'Fretes'), ('L', 'U2.9', 'Correios e Encomendas'),
    ('L', 'U2.10', 'Material de Uso e Consumo'), ('L', 'U2.11', 'Outras Despesas Operacionais'),
    ('L', 'U2.14', 'Despesas Tributárias'), ('L', 'U2.15', 'Outras Despesas e Receitas Líquidas'),
    ('L', 'U2.16', 'Equivalência Patrimonial'), ('L', 'U_DEqR', 'Depreciação e Amortização'),
    ('SUB', None, 'Lucro antes do Resultado Financeiro'),
    ('L', 'U3.1', 'Despesas Financeiras (U3.1)'), ('L', 'U3.2', 'Despesas Financeiras (U3.2)'),
    ('L', 'U3.3', 'Despesas Financeiras (U3.3)'), ('L', 'U3.4', 'Despesas Financeiras (U3.4)'),
    ('L', 'U3.5', 'Despesas Financeiras (U3.5)'),
    ('L', 'U4.1', 'Receitas Financeiras (U4.1)'), ('L', 'U4.2', 'Receitas Financeiras (U4.2)'),
    ('SUB', None, 'Resultado Financeiro Líquido'),
    ('SUB', None, 'Lucro Antes do IR / CS'),
    ('L', 'O2.1', 'IRPJ e CSLL'),
    ('TOT', None, 'Lucro Líquido'),
    ('CHK', None, 'CHECK: LL − Σ linhas (deve ser ~0)'),
]

NOTAS_MAP = {
    'U1.1': ('17', 'Receita operacional líquida'), 'U1.2': ('17', 'Receita operacional líquida'),
    'U1.3': ('17', 'Receita operacional líquida'),
    'U2.1': ('18', 'Custos e despesas por natureza'), 'U2.2': ('18', 'Custos e despesas por natureza'),
    'U2.3': ('18', 'Custos e despesas por natureza'), 'U2.5': ('18', 'Custos e despesas por natureza'),
    'U2.6': ('18', 'Custos e despesas por natureza'), 'U2.7': ('18', 'Custos e despesas por natureza'),
    'U2.8': ('18', 'Custos e despesas por natureza'), 'U2.9': ('18', 'Custos e despesas por natureza'),
    'U2.10': ('18', 'Custos e despesas por natureza'), 'U2.11': ('18', 'Custos e despesas por natureza'),
    'U2.14': ('18', 'Custos e despesas por natureza'), 'U2.15': ('18', 'Custos e despesas por natureza'),
    'U2.16': ('12', 'Investimentos / Equivalência patrimonial'),
    'U_DEqR': ('7 e 8', 'Imobilizado e Intangível'),
    'U3.1': ('19', 'Resultado financeiro'), 'U3.2': ('19', 'Resultado financeiro'),
    'U3.3': ('19', 'Resultado financeiro'), 'U3.4': ('19', 'Resultado financeiro'),
    'U3.5': ('19', 'Resultado financeiro'),
    'U4.1': ('19', 'Resultado financeiro'), 'U4.2': ('19', 'Resultado financeiro'),
    'O2.1': ('13', 'Imposto de renda e contribuição social'),
    'A': ('3', 'Caixa e equivalentes de caixa'), 'C': ('4', 'Contas a receber de clientes'),
    'D': ('5', 'Estoques'), '9': ('6', 'Impostos a recuperar'),
    'E': ('20', 'Adiantamentos (instrumentos financeiros)'), 'N': ('12', 'Investimentos'),
    'N1': ('9', 'Fornecedores'), 'K': ('7', 'Imobilizado'), 'L': ('8', 'Intangível'),
    'T': ('16', 'Patrimônio líquido'), 'T1': ('16', 'Patrimônio líquido'),
    'T3': ('16', 'Patrimônio líquido'),
}

# ------------------------------------------------------------------ workbook
def build(depara, balancete, out):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    desc_of = {d['conta']: (sanitize(d['desc']) or d['conta']) for d in depara}
    dre_of  = {d['conta']: d['dre'] for d in depara}
    bp_of   = {d['conta']: d['bp'] for d in depara}
    cod_of  = {d['conta']: d['df'] for d in depara}

    # letra de BP mais comum por linha (coluna Nota do Balanço)
    letra_cnt = {}
    for d in depara:
        if d['bp'] and len(d['df']) <= 3 and not d['df'].startswith(('U', 'O')):
            letra_cnt.setdefault(d['bp'], Counter())[d['df']] += 1
    letra_bp = {k: v.most_common(1)[0][0] for k, v in letra_cnt.items()}

    # linhas de BP não cobertas pelo layout do Balanço (aviso)
    cobertas = {x[1] for x in BP_LINES if x[0] == 'L'}
    fora = {d['bp'] for d in depara
            if d['bp'] and d['bp'] not in cobertas and d['dre'] not in DRE_SET}
    if fora:
        print(f'AVISO: linhas de BP fora do layout do Balanço: {sorted(fora)}')

    # ============================================================ DePara
    dp = wb.create_sheet('DePara')
    head(dp, 1, ['Conta', 'Descrição', 'LinhaDRE', 'LinhaBP', 'CodNota', 'Grupo_BP', 'ÉResultado'])
    r = 2
    for d in sorted(depara, key=lambda x: x['conta']):
        dp.cell(r, 1, d['conta']).font = F_IN
        dp.cell(r, 2, desc_of[d['conta']])
        dp.cell(r, 3, d['dre'])
        dp.cell(r, 4, d['bp'])
        dp.cell(r, 5, d['df'])
        dp.cell(r, 6, d['df'] if (len(d['df']) <= 3 and not d['df'].startswith(('U', 'O'))) else '')
        dp.cell(r, 7, bool(d['dre'] in DRE_SET)).font = F_IN
        r += 1
    DP_LAST = r - 1
    for col, w in zip('ABCDEFG', [12, 46, 36, 36, 10, 10, 12]):
        dp.column_dimensions[col].width = w
    dp.freeze_panes = 'A2'
    dp.add_table(Table(displayName='tbl_DePara', ref=f'A1:G{DP_LAST}',
                          tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))

    # ============================================================ Balancete
    bl = wb.create_sheet('Balancete')
    head(bl, 1, ['Empresa', 'Conta', 'Descrição', 'Ano', 'Período', 'Saldo',
                 'LinhaBP', 'LinhaDRE', 'CodNota',
                 'Nível 1', 'Nível 2', 'Nível 3', 'Nível 4', 'Nível 5', 'Tipo', 'Data', 'Fluxo'])
    balancete.sort(key=lambda x: (x[0], x[1], x[2]))
    # Fluxo mensal: contas de resultado no balancete são ACUMULADAS (YTD) —
    # Fluxo = saldo(mês) − saldo(do ÚLTIMO mês EXISTENTE da conta; telescópico).
    # Em JANEIRO o YTD reinicia: fluxo(jan) = saldo próprio.
    # Telescópico garante: ΣFluxo do ano = Saldo YTD de Dez (à prova de buracos —
    # contas que pulam/somem meses na origem não quebram o total).
    _saldos = defaultdict(dict)
    for e, c, mes, v in balancete:
        _saldos[(e, c)][mes] = v
    def _fluxo(e, c, mes, v):
        if mes.endswith('-01'):
            return v
        meses_conta = _saldos[(e, c)]
        # anteriores SOMENTE dentro do mesmo ano (o YTD reinicia em janeiro)
        anteriores = [m for m in MESES36
                      if m[:4] == mes[:4] and m < mes and m in meses_conta]
        if not anteriores:
            return v
        return v - meses_conta[anteriores[-1]]
    # Contas de resultado que SUMIRAM antes de Dez (com saldo ≠ 0): gerar linha
    # "fantasma" (Saldo=0) no mês seguinte ao último existente com Fluxo = −saldo,
    # replicando o unwinding que o SUMIFS do DF_Mensal faz (mês ausente = 0).
    dre_set = DRE_SET
    fantasma = []
    for (e, c), meses_conta in _saldos.items():
        if c[0] not in '34':
            continue
        if dre_of.get(c) not in dre_set:
            continue
        if '2025-12' in meses_conta and '2024-12' in meses_conta:
            continue
        for ano in ('2024', '2025'):
            dez = f'{ano}-12'
            meses_ano = [m for m in meses_conta if m.startswith(ano)]
            if not meses_ano or dez in meses_ano:
                continue
            ultimo = max(meses_ano)
            seg = MESES36[min(MESES36.index(ultimo) + 1, len(MESES36) - 1)]
            if seg.startswith(ano):
                fantasma.append((e, c, seg, -meses_conta[ultimo]))
    for e, c, mes, v in fantasma:
        balancete.append((e, c, mes, 0.0))
        _saldos[(e, c)][mes] = 0.0
    if fantasma:
        print(f'--- LINHAS FANTASMA (unwinding de contas sumidas): {len(fantasma)} ---')
        for e, c, mes, v in fantasma[:10]:
            print(f'   {e} {c} {mes}: fluxo={v:,.2f}')
    balancete.sort(key=lambda x: (x[0], x[1], x[2]))
    r = 2
    for emp, c, mes, v in balancete:
        bl.cell(r, 1, emp).font = F_IN
        bl.cell(r, 2, c).font = F_IN
        bl.cell(r, 3, desc_of.get(c, c))
        bl.cell(r, 4, int(mes[:4])).font = F_IN
        pc = bl.cell(r, 5, mes); pc.font = F_IN; pc.number_format = '@'
        vc = bl.cell(r, 6, round(v, 2)); vc.font = F_IN; vc.number_format = FMT_V2
        bl.cell(r, 7, f'=IFERROR(VLOOKUP($B{r},DePara!$A:$E,4,FALSE),"")')
        bl.cell(r, 8, f'=IFERROR(VLOOKUP($B{r},DePara!$A:$E,3,FALSE),"")')
        bl.cell(r, 9, f'=IFERROR(VLOOKUP($B{r},DePara!$A:$E,5,FALSE),"")')
        for k, nv in enumerate(niveis(c)[:5], 10):
            bl.cell(r, k, nv).font = F_IN
        bl.cell(r, 15, 'Analitica').font = F_IN
        dcell = bl.cell(r, 16, date(int(mes[:4]), int(mes[5:7]), 1))
        dcell.font = F_IN; dcell.number_format = 'dd/mm/yyyy'
        flc = bl.cell(r, 17, round(_fluxo(emp, c, mes, v), 2))
        flc.font = F_IN; flc.number_format = FMT_V2
        r += 1
    BL_LAST = r - 1
    for col, w in zip('ABCDEFGHIJKLMNOPQ',
                      [9, 11, 44, 7, 10, 15, 32, 32, 10, 8, 8, 8, 8, 8, 10, 12, 15]):
        bl.column_dimensions[col].width = w
    bl.freeze_panes = 'A2'
    bl.add_table(Table(displayName='tbl_Balancete', ref=f'A1:Q{BL_LAST}',
                          tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))
    BF_ = f'Balancete!$F$2:$F${BL_LAST}'; BA_ = f'Balancete!$A$2:$A${BL_LAST}'
    BE_ = f'Balancete!$E$2:$E${BL_LAST}'; BG_ = f'Balancete!$G$2:$G${BL_LAST}'
    BH_ = f'Balancete!$H$2:$H${BL_LAST}'; BI_ = f'Balancete!$I$2:$I${BL_LAST}'
    print(f'Balancete: {BL_LAST - 1} lançamentos')

    # ============================================================ dCalendario
    dc = wb.create_sheet('dCalendario')
    head(dc, 1, ['Data', 'Ano', 'MêsNum', 'Mês', 'Trimestre', 'AnoMês', 'AnoMêsOrd'])
    MES_PT = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    r = 2
    from datetime import timedelta
    d0, d1 = date(2024, 1, 1), date(2025, 12, 31)
    nd = (d1 - d0).days + 1
    for i in range(nd):
        dt = d0 + timedelta(days=i)
        dc.cell(r, 1, dt).number_format = 'dd/mm/yyyy'
        dc.cell(r, 2, dt.year)
        dc.cell(r, 3, dt.month)
        dc.cell(r, 4, MES_PT[dt.month - 1])
        dc.cell(r, 5, f'T{(dt.month - 1) // 3 + 1}')
        dc.cell(r, 6, f'{dt.year}-{dt.month:02d}')
        dc.cell(r, 7, dt.year * 100 + dt.month)
        r += 1
    DC_LAST = r - 1
    dc.add_table(Table(displayName='tbl_Calendario', ref=f'A1:G{DC_LAST}',
                       tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))
    for col, w in zip('ABCDEFG', [12, 8, 8, 8, 10, 10, 10]):
        dc.column_dimensions[col].width = w
    dc.freeze_panes = 'A2'
    print(f'dCalendario: {nd} dias')

    # ============================================================ Eliminacao
    el = wb.create_sheet('Eliminacao')
    head(el, 1, ['Conta', 'Descrição', 'Período', 'Valor', 'CodNota', 'LinhaBP', 'LinhaDRE', 'Fluxo', 'Data'])
    r = 2
    for d in sorted(depara, key=lambda x: x['conta']):
        for mes in MESES_ELIM:
            v = d['elim'].get(mes, 0.0)
            if abs(v) > 1e-9:
                # eliminações também são YTD: Fluxo = valor(mês) − valor(mês anterior)
                prev = MESES_ELIM[MESES_ELIM.index(mes) - 1] if mes != MESES_ELIM[0] else None
                pv = d['elim'].get(prev, 0.0) * SCALE if prev else 0.0
                el.cell(r, 1, d['conta']).font = F_IN
                el.cell(r, 2, desc_of.get(d['conta'], d['conta']))
                pc = el.cell(r, 3, mes); pc.font = F_IN; pc.number_format = '@'
                vc = el.cell(r, 4, round(v * SCALE, 2)); vc.font = F_IN; vc.number_format = FMT_V2
                el.cell(r, 5, f'=IFERROR(VLOOKUP($A{r},DePara!$A:$E,5,FALSE),"")')
                el.cell(r, 6, f'=IFERROR(VLOOKUP($A{r},DePara!$A:$E,4,FALSE),"")')
                el.cell(r, 7, f'=IFERROR(VLOOKUP($A{r},DePara!$A:$E,3,FALSE),"")')
                flc = el.cell(r, 8, round(v * SCALE - pv, 2))
                flc.font = F_IN; flc.number_format = FMT_V2
                dcell = el.cell(r, 9, date(int(mes[:4]), int(mes[5:7]), 1))
                dcell.font = F_IN; dcell.number_format = 'dd/mm/yyyy'
                r += 1
    EL_LAST = r - 1
    for col, w in zip('ABCDEFGHI', [12, 44, 10, 15, 10, 32, 32, 15, 12]):
        el.column_dimensions[col].width = w
    el.freeze_panes = 'A2'
    el.add_table(Table(displayName='tbl_Eliminacao', ref=f'A1:I{EL_LAST}',
                          tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))
    EV_ = f'Eliminacao!$D$2:$D${EL_LAST}'; EP_ = f'Eliminacao!$C$2:$C${EL_LAST}'
    EB_ = f'Eliminacao!$F$2:$F${EL_LAST}'; ED_ = f'Eliminacao!$G$2:$G${EL_LAST}'
    EC_ = f'Eliminacao!$E$2:$E${EL_LAST}'
    print(f'Eliminacao: {EL_LAST - 1} lançamentos')

    # ============================================================ DRE
    dr = wb.create_sheet('DRE')
    title_block(dr, 'DEMONSTRAÇÃO DO RESULTADO — GRUPO ABC',
                'Exercícios 2024 × 2025 · Balancetes de Dez (acumulado YTD) · Valores fictícios (× 0,85) · R$')
    head(dr, 4, ['Linha', 'Nota', 'ALFA 2024', 'Combinado 2024', 'Eliminação 2024',
                 'Consolidado 2024', 'ALFA 2025', 'Combinado 2025', 'Eliminação 2025',
                 'Consolidado 2025'])
    dmap, r = {}, 5
    for tipo, label, _ in DRE_LINES:
        dmap[label] = r
        c = dr.cell(r, 1, label)
        if tipo == 'L':
            dr.cell(r, 2, '')
            for col, per, emp_f in (('C', '2024-12', True), ('D', '2024-12', False),
                                    ('G', '2025-12', True), ('H', '2025-12', False)):
                base = f'SUMIFS({BF_},{BH_},$A{r},{BE_},"{per}")'
                if emp_f:
                    base = f'SUMIFS({BF_},{BA_},"ALFA",{BH_},$A{r},{BE_},"{per}")'
                dr[f'{col}{r}'] = '=' + base
            dr[f'E{r}'] = f'=-SUMIFS({EV_},{ED_},$A{r},{EP_},"2024-12")'
            dr[f'F{r}'] = f'=D{r}-E{r}'
            dr[f'I{r}'] = f'=-SUMIFS({EV_},{ED_},$A{r},{EP_},"2025-12")'
            dr[f'J{r}'] = f'=H{r}-I{r}'
        elif tipo == 'CHK':
            c.font = F_CHK
        r += 1
    LL_R = dmap['Lucro Líquido']
    CHK_R = dmap[DRE_LINES[-1][1]]

    def dre_sub(label, fn):
        rr = dmap[label]
        for col in 'CDEFGHIJ':
            dr[f'{col}{rr}'] = fn(col)
    dre_sub('Lucro Bruto', lambda c: f'={c}{dmap["ROL"]}+{c}{dmap["CMV"]}')
    dre_sub('Lucro antes do Resultado',
            lambda c: f'={c}{dmap["Lucro Bruto"]}+SUM({c}{dmap["Despesas com Vendas"]}:{c}{dmap["Outras Despesas Operacionais, Líquida"]})')
    dre_sub('Lucro Antes do IR / CS',
            lambda c: f'={c}{dmap["Lucro antes do Resultado"]}+{c}{dmap["Despesas Financeiras"]}+{c}{dmap["Receitas Financeiras"]}')
    dre_sub('Lucro Líquido', lambda c: f'={c}{dmap["Lucro Antes do IR / CS"]}+{c}{dmap["Imposto de Renda"]}')
    # checks: LL anual vs soma dos meses do DF_Mensal (LL na linha 35; 2024: C..N | 2025: O..Z)
    dr[f'F{CHK_R}'] = f'=F{LL_R}-SUM(DF_Mensal!C35:N35)'
    dr[f'J{CHK_R}'] = f'=J{LL_R}-SUM(DF_Mensal!O35:Z35)'
    for col, w in zip('ABCDEFGHIJ', [42, 7, 14, 15, 15, 15, 14, 15, 15, 15]):
        dr.column_dimensions[col].width = w
    dr.freeze_panes = 'C5'
    fmt_range(dr, f'C5:J{LL_R}')
    for lab in ('Lucro Bruto', 'Lucro antes do Resultado', 'Lucro Antes do IR / CS', 'Lucro Líquido'):
        for j in range(1, 11):
            dr.cell(dmap[lab], j).font = F_TOT
            dr.cell(dmap[lab], j).border = B_TOP
    dr.cell(dmap['Lucro Líquido'], 1).fill = FILL_GOLDL
    dr.cell(CHK_R, 1).font = F_CHK

    # ============================================================ Balanco
    bq = wb.create_sheet('Balanco')
    title_block(bq, 'BALANÇO PATRIMONIAL — GRUPO ABC',
                'Data-base 31/12/2025 (auditado) · Comparativo Dez/24 · Valores fictícios (× 0,85) · R$')
    head(bq, 4, ['Linha', 'Nota', 'ALFA Dez/25', 'Combinado Dez/25', 'Eliminação',
                 'Consolidado Dez/25', 'Combinado Dez/24'])
    rowmap, r = {}, 5
    for tipo, label, sign in BP_LINES:
        rowmap[label] = r
        if tipo in ('SEC', 'SUB'):
            c = bq.cell(r, 1, label)
            c.font = F_SEC if tipo == 'SEC' else F_TOT
            fill = FILL_NAVY if tipo == 'SEC' else FILL_GOLDL
            for j in range(1, 8):
                bq.cell(r, j).fill = fill
        elif tipo == 'L':
            bq.cell(r, 1, label)
            bq.cell(r, 2, letra_bp.get(label, ''))
        elif tipo == 'TOT':
            bq.cell(r, 1, label).font = F_TOT
        elif tipo == 'CHK':
            bq.cell(r, 1, label).font = F_CHK
        r += 1
    TOT_AT = rowmap['Total do Ativo']
    TOT_PL = rowmap['TOTAL DO PASSIVO + PATRIMÔNIO LÍQUIDO']
    CHK_B = rowmap[BP_LINES[-1][1]]
    LL_ROW = LL_R  # linha do LL na DRE
    for tipo, label, sign in BP_LINES:
        if tipo != 'L':
            continue
        rr = rowmap[label]
        if label == 'Lucros acumulados':
            bq[f'C{rr}'] = f'=DRE!G{LL_ROW}+SUMIFS({BF_},{BA_},"ALFA",{BG_},$A{rr},{BE_},"2025-12")'
            bq[f'D{rr}'] = f'=DRE!H{LL_ROW}+SUMIFS({BF_},{BG_},$A{rr},{BE_},"2025-12")'
            bq[f'E{rr}'] = f'=DRE!I{LL_ROW}'
            bq[f'F{rr}'] = f'=D{rr}-E{rr}'
            bq[f'G{rr}'] = f'=DRE!D{LL_ROW}+SUMIFS({BF_},{BG_},$A{rr},{BE_},"2024-12")'
        else:
            bq[f'C{rr}'] = f'=SUMIFS({BF_},{BA_},"ALFA",{BG_},$A{rr},{BE_},"2025-12")'
            bq[f'D{rr}'] = f'=SUMIFS({BF_},{BG_},$A{rr},{BE_},"2025-12")'
            bq[f'E{rr}'] = f'={sign}*SUMIFS({EV_},{EB_},$A{rr},{EP_},"2025-12")'
            bq[f'F{rr}'] = f'=D{rr}-E{rr}'
            bq[f'G{rr}'] = f'=SUMIFS({BF_},{BG_},$A{rr},{BE_},"2024-12")'

    def bal_tot(label, comps):
        rr = rowmap[label]
        for col in 'CDEFG':
            bq[f'{col}{rr}'] = '=' + '+'.join(f'{col}{rowmap[x]}' for x in comps)
    bal_tot('Total do Ativo Circulante',
            ['Caixa e equivalentes de caixa', 'Títulos e valores mobiliários',
             'Contas a receber de clientes', 'Estoques', 'Impostos Recuperar',
             'Adiantamento a fornecedores', 'Outros Ativos'])
    bal_tot('Total do Ativo Não Circulante',
            ['Depósitos Judiciais', 'Partes Relacionadas', 'Investimentos', 'Imobilizado',
             'Direito de Uso', 'Intangível'])
    bal_tot('Total do Ativo', ['Total do Ativo Circulante', 'Total do Ativo Não Circulante'])
    bal_tot('Total do Passivo Circulante',
            ['Fornecedores', 'Empréstimos', 'Instrumentos financeiros',
             'impostos, taxas e contribuições a recolher',
             'Imposto de renda e contribuição social a pagar', 'Adiantamento de clientes',
             'Salários e encargos Sociais', 'Obrigações sociais e trabalhistas',
             'Outras contas a pagar', 'Arrendamento Curto Prazo', 'Outros Passivos'])
    bal_tot('Total do Passivo Não Circulante',
            ['Arrendamento Longo Prazo', 'Subvenção para investimentos',
             'Obrigações Tributárias', 'Provisão para contingência'])
    bal_tot('Total do Patrimônio Líquido',
            ['Capital social', 'Ajuste de avaliação patrimonial', 'Lucros acumulados'])
    bal_tot('TOTAL DO PASSIVO + PATRIMÔNIO LÍQUIDO',
            ['Total do Passivo Circulante', 'Total do Passivo Não Circulante',
             'Total do Patrimônio Líquido'])
    for col in 'CDEFG':
        bq[f'{col}{CHK_B}'] = f'={col}{TOT_AT}-{col}{TOT_PL}'
    for col, w in zip('ABCDEFG', [46, 7, 16, 17, 15, 17, 17]):
        bq.column_dimensions[col].width = w
    bq.freeze_panes = 'C5'
    fmt_range(bq, f'C5:G{CHK_B}')
    for lab in ('Total do Ativo Circulante', 'Total do Ativo Não Circulante', 'Total do Ativo',
                'Total do Passivo Circulante', 'Total do Passivo Não Circulante',
                'Total do Patrimônio Líquido', 'TOTAL DO PASSIVO + PATRIMÔNIO LÍQUIDO'):
        for j in range(1, 8):
            bq.cell(rowmap[lab], j).font = F_TOT
            bq.cell(rowmap[lab], j).border = B_TOP
    bq.cell(TOT_AT, 1).fill = FILL_GOLDL
    bq.cell(TOT_PL, 1).fill = FILL_GOLDL

    # ============================================================ DF_Mensal
    df = wb.create_sheet('DF_Mensal')
    title_block(df, 'DEMONSTRAÇÃO DO RESULTADO MENSAL POR NOTA — GRUPO ABC',
                'Consolidado · Jan/24 a Dez/25 · Jan/24–Abr/25 sem eliminações (não rastreadas na fonte) · R$')
    # linha 2 = mês anterior (chave), linha 3 = mês corrente (chave), linha 4 = rótulo
    df.cell(2, 1, 'Mês anterior (chave):').font = F_SM
    df.cell(3, 1, 'Mês (chave):').font = F_SM
    for i, mes in enumerate(MESES24):
        col = 3 + i
        prev = MESES36[MESES36.index(mes) - 1]
        df.cell(2, col, prev).font = F_SM
        df.cell(3, col, mes).font = F_SM
        c = df.cell(4, col, LABEL_MES[mes])
        c.font, c.fill = F_H, FILL_NAVY
        c.alignment = Alignment(horizontal='center')
    df.cell(4, 1, 'Código')
    df.cell(4, 2, 'Linha')
    df.cell(4, 1).font = F_H; df.cell(4, 1).fill = FILL_NAVY
    df.cell(4, 2).font = F_H; df.cell(4, 2).fill = FILL_NAVY
    fmap, r = {}, 5
    for tipo, cod, label in DF_LINES:
        fmap[label] = r
        if tipo == 'L':
            df.cell(r, 1, cod).font = F_IN
            df.cell(r, 2, label)
        else:
            c = df.cell(r, 2, label)
            c.font = F_CHK if tipo == 'CHK' else F_TOT
        r += 1
    DF_LL = fmap['Lucro Líquido']
    DF_CHK = fmap[DF_LINES[-1][2]]
    R_ROL, R_LB = fmap['Receita Operacional Líquida'], fmap['Lucro Bruto']
    R_LARF = fmap['Lucro antes do Resultado Financeiro']
    R_FIN = fmap['Resultado Financeiro Líquido']
    R_LAIR = fmap['Lucro Antes do IR / CS']
    R_IR = fmap['O2.1'] if 'O2.1' in fmap else fmap['IRPJ e CSLL']
    first_u2 = fmap['Despesas com Pessoal']
    last_u2 = fmap['Depreciação e Amortização']
    first_fin = fmap['Despesas Financeiras (U3.1)']
    last_fin = fmap['Receitas Financeiras (U4.2)']
    for tipo, cod, label in DF_LINES:
        if tipo != 'L':
            continue
        rr = fmap[label]
        for i, mes in enumerate(MESES24):
            col = get_column_letter(3 + i)
            if mes.endswith('-01'):
                # janeiro: o YTD reinicia — o saldo de janeiro JÁ é o fluxo do mês
                df[f'{col}{rr}'] = (
                    f'=SUMIFS({BF_},{BI_},$A{rr},{BE_},{col}$3)'
                    f'+SUMIFS({EV_},{EC_},$A{rr},{EP_},{col}$3)')
            else:
                df[f'{col}{rr}'] = (
                    f'=SUMIFS({BF_},{BI_},$A{rr},{BE_},{col}$3)'
                    f'-SUMIFS({BF_},{BI_},$A{rr},{BE_},{col}$2)'
                    f'+SUMIFS({EV_},{EC_},$A{rr},{EP_},{col}$3)'
                    f'-SUMIFS({EV_},{EC_},$A{rr},{EP_},{col}$2)')
    def df_sub(label, fn):
        rr = fmap[label]
        for i in range(24):
            col = get_column_letter(3 + i)
            df[f'{col}{rr}'] = fn(col)
    df_sub('Receita Operacional Líquida', lambda c: f'=SUM({c}{fmap["Receita Líquida"]}:{c}{fmap["Impostos sobre Vendas"]})')
    df_sub('Lucro Bruto', lambda c: f'={c}{R_ROL}+{c}{fmap["Custos com Mercadorias Vendidas"]}')
    df_sub('Lucro antes do Resultado Financeiro',
           lambda c: f'={c}{R_LB}+SUM({c}{first_u2}:{c}{last_u2})')
    df_sub('Resultado Financeiro Líquido', lambda c: f'=SUM({c}{first_fin}:{c}{last_fin})')
    df_sub('Lucro Antes do IR / CS', lambda c: f'={c}{R_LARF}+{c}{R_FIN}')
    df_sub('Lucro Líquido', lambda c: f'={c}{R_LAIR}+{c}{fmap["IRPJ e CSLL"]}')
    for i in range(24):
        col = get_column_letter(3 + i)
        df[f'{col}{DF_CHK}'] = (f'={col}{DF_LL}'
                                f'-(SUM({col}{fmap["Receita Líquida"]}:{col}{fmap["Impostos sobre Vendas"]})'
                                f'+{col}{fmap["Custos com Mercadorias Vendidas"]}'
                                f'+SUM({col}{first_u2}:{col}{last_u2})'
                                f'+SUM({col}{first_fin}:{col}{last_fin})'
                                f'+{col}{fmap["IRPJ e CSLL"]})')
    df.column_dimensions['A'].width = 9
    df.column_dimensions['B'].width = 36
    for i in range(24):
        df.column_dimensions[get_column_letter(3 + i)].width = 11
    df.freeze_panes = 'C5'
    fmt_range(df, f'C5:{get_column_letter(26)}{DF_LL}')
    for lab in ('Receita Operacional Líquida', 'Lucro Bruto', 'Lucro antes do Resultado Financeiro',
                'Resultado Financeiro Líquido', 'Lucro Antes do IR / CS', 'Lucro Líquido'):
        for j in range(1, 27):
            df.cell(fmap[lab], j).font = F_TOT
            df.cell(fmap[lab], j).border = B_TOP
    df.cell(DF_LL, 2).fill = FILL_GOLDL
    df.cell(DF_CHK, 2).font = F_CHK

    # ============================================================ EBITDA
    eb = wb.create_sheet('EBITDA')
    title_block(eb, 'EBITDA TRIMESTRAL — GRUPO ABC',
                'EBITDA = LL − Resultado Financeiro − IRPJ/CSLL − D&A · Consolidado · R$')
    TRIM = [('1T/24', 'C', 'E'), ('2T/24', 'F', 'H'), ('3T/24', 'I', 'K'), ('4T/24', 'L', 'N'),
            ('1T/25', 'O', 'Q'), ('2T/25', 'R', 'T'), ('3T/25', 'U', 'W'), ('4T/25', 'X', 'Z')]
    head(eb, 4, ['Linha'] + [t[0] for t in TRIM])
    eb_rows = [('Receita Operacional Líquida', R_ROL), ('Lucro Bruto', R_LB),
               ('Depreciação e Amortização', fmap['Depreciação e Amortização']),
               ('Resultado Financeiro Líquido', R_FIN), ('IRPJ e CSLL', fmap['IRPJ e CSLL']),
               ('Lucro Líquido', DF_LL)]
    r = 5
    emap = {}
    for label, src_row in eb_rows:
        emap[label] = r
        eb.cell(r, 1, label)
        for j, (t, c1, c2) in enumerate(TRIM, 2):
            eb.cell(r, j, f'=SUM(DF_Mensal!{c1}{src_row}:{c2}{src_row})')
        r += 1
    emap['EBITDA'] = r
    c = eb.cell(r, 1, 'EBITDA'); c.font = F_TOT
    for j in range(2, 10):
        col = get_column_letter(j)
        eb[f'{col}{r}'] = (f'={col}{emap["Lucro Líquido"]}-{col}{emap["Resultado Financeiro Líquido"]}'
                           f'-{col}{emap["IRPJ e CSLL"]}-{col}{emap["Depreciação e Amortização"]}')
    for col, w in zip('ABCDEFGHI', [34] + [12] * 8):
        eb.column_dimensions[col].width = w
    fmt_range(eb, f'B5:I{r}')
    for j in range(1, 10):
        eb.cell(r, j).font = F_TOT
        eb.cell(r, j).border = B_TOP
    eb.cell(r, 1).fill = FILL_GOLDL

    # ============================================================ Notas
    nt = wb.create_sheet('Notas')
    title_block(nt, 'ÍNDICE DE NOTAS EXPLICATIVAS — GRUPO ABC',
                'Vínculo código de nota (DePara) × nota do documento "Notas Explicativas Grupo ABC"')
    head(nt, 4, ['Código', 'Presente em', 'Nota (documento)', 'Título da nota'])
    r = 5
    cods_df = sorted({d['df'] for d in depara if d['df'].startswith(('U', 'O'))})
    letras = sorted({d['df'] for d in depara
                     if len(d['df']) <= 3 and not d['df'].startswith(('U', 'O'))})
    for cod in cods_df + letras:
        nota, titulo = NOTAS_MAP.get(cod, ('—', '—'))
        nt.cell(r, 1, cod).font = F_IN
        nt.cell(r, 2, 'DF / DRE' if cod.startswith(('U', 'O')) else 'Balanço')
        nt.cell(r, 3, nota)
        nt.cell(r, 4, titulo)
        r += 1
    for col, w in zip('ABCD', [10, 12, 18, 44]):
        nt.column_dimensions[col].width = w
    nt.freeze_panes = 'A5'

    # ============================================================ Parametros
    pm = wb.create_sheet('Parametros')
    title_block(pm, 'PARÂMETROS DO MODELO — GRUPO ABC', 'Premissas, empresas e regras de consolidação')
    rows = [
        ('EMPRESAS DO GRUPO', ''),
        ('ALFA', 'Controladora (holding — administra marcas, royalties e participações)'),
        ('BETA', 'Controlada — farmácia de manipulação (lojas próprias) · lucro real'),
        ('GAMA', 'Controlada — distribuidora de cosméticos (atacado) · lucro presumido'),
        ('DELTA', 'Controlada — distribuidora de cosméticos (atacado) · lucro presumido'),
        ('', ''),
        ('PREMISSAS', ''),
        ('Data-base', '31/12/2025 — auditado (interim em Nov/2025, auditoria final em Dez/2025)'),
        ('Comparativo', 'Dez/2024 (combinado — sem eliminações rastreadas)'),
        ('Escala dos valores', 'FICTÍCIOS = valores reais × 0,85 (anonimização)'),
        ('Unidade', 'R$ (reais correntes)'),
        ('', ''),
        ('REGRAS DE CONSOLIDAÇÃO', ''),
        ('Regra 1', 'Consolidado = Combinado (soma de todas) − Eliminações intercompany'),
        ('Regra 2', 'A controladora é sempre a ALFA — capital e resultado prevalecentes são os dela'),
        ('Regra 3', 'Eliminações por contas específicas: investimentos (1312xxxx), capital (22111001),'),
        ('', 'fornecedores IC (21111008), CMV de vendas IC (31211003), EQP (31521001/41131006)'),
        ('Regra 4', 'Eliminações rastreadas a partir de Mai/2025 · Jan/24–Abr/25 = combinado'),
        ('Regra 5', 'Contas de resultado no balancete são ACUMULADAS (YTD) — o DF mensal usa delta'),
        ('Regra 6', 'Em Jan o YTD reinicia: saldo de janeiro = fluxo do mês (sem delta)'),
        ('', ''),
        ('NORMALIZAÇÃO', 'Conta 29999999 neutraliza resultado já apropriado no PL'),
        ('', '(evita duplicidade P&L × lucros acumulados; lançada automaticamente'),
        ('', 'nos meses em que Ativo = Passivo + Rec + Desp não fecha)'),
        ('', ''),
        ('CHAVE UNIVERSAL', 'EMPRESA + PERÍODO + CONTA'),
        ('', ''),
        ('PERÍODOS DISPONÍVEIS', 'Jan/2023 a Dez/2025 (balancetes) · Quadros: Jan/2024 a Dez/2025'),
    ]
    r = 4
    for k, v in rows:
        a = pm.cell(r, 1, k); b = pm.cell(r, 2, v)
        if k and not v and k == k.upper():
            a.font = F_SEC; a.fill = FILL_NAVY
            pm.cell(r, 2).fill = FILL_NAVY
        else:
            a.font = F_TOT if k else F_LIN
            b.font = F_LIN
        r += 1
    pm.column_dimensions['A'].width = 26
    pm.column_dimensions['B'].width = 90

    # ============================================================ CAPA
    cp = wb.create_sheet('CAPA')
    cp.sheet_view.showGridLines = False
    cp['B2'] = 'MODELO FINANCEIRO PADRÃO — GRUPO ABC'
    cp['B2'].font = Font(name='Arial', size=20, bold=True, color=NAVY)
    cp['B3'] = 'Aguiar Advisory · Controladoria estratégica · Piloto anonimizado · Data-base Dez/2025'
    cp['B3'].font = F_SUB
    cp['B5'] = '⚠ DADOS FICTÍCIOS: nomes de empresas, bancos, lojas e sócios foram substituídos; valores = reais × 0,85. Uso interno / portfólio.'
    cp['B5'].font = Font(name='Arial', size=10, bold=True, color=RED)
    cp['B7'] = 'ÍNDICE'
    cp['B7'].font = Font(name='Arial', size=13, bold=True, color=GOLD)
    indice = [
        ('Balanco', 'Balanço patrimonial consolidado — Dez/25 (ALFA, Combinado, Eliminação, Consolidado) + Dez/24'),
        ('DRE', 'Demonstração do resultado — 2024 × 2025, com eliminações'),
        ('DF_Mensal', 'DRE mensal por nota explicativa (códigos U) — Jan/24 a Dez/25'),
        ('EBITDA', 'EBITDA trimestral — derivação transparente a partir do DF mensal'),
        ('Balancete', 'BASE ÚNICA DE SALDOS: Empresa × Conta × Período (fonte do Power BI)'),
        ('DePara', 'Mapa de contas: conta → linha de DRE / linha de BP / código de nota'),
        ('Eliminacao', 'Eliminações intercompany — Conta × Período'),
        ('Notas', 'Índice: código de nota × nota do documento'),
        ('Parametros', 'Empresas, premissas e regras de consolidação'),
    ]
    r = 8
    for sheet, desc in indice:
        c = cp.cell(r, 2, f'=HYPERLINK("#\'{sheet}\'!A1","▶ {sheet}")')
        c.font = Font(name='Arial', size=11, bold=True, color=NAVY, underline='single')
        cp.cell(r, 4, desc).font = F_LIN
        r += 1
    r += 1
    cp.cell(r, 2, 'REGRAS DE OURO — O QUE NÃO PODE MUDAR').font = Font(name='Arial', size=13, bold=True, color=GOLD)
    r += 1
    regras = [
        '1. Chave universal: EMPRESA + PERÍODO + CONTA — toda base nova entra por aqui.',
        '2. Quadros 100% por fórmula (SUMIFS/VLOOKUP no Balancete/DePara/Eliminacao) — NADA digitado.',
        '3. Consolidado = Combinado − Eliminações. A controladora é sempre a ALFA.',
        '4. Contas de resultado são acumuladas (YTD): o DF mensal subtrai o mês anterior.',
        '5. Toda conta nova DEVE nascer no DePara com destino (DRE/BP) e código de nota.',
        '6. Checks de amarração (verde) devem fechar em ~0 — se não fechar, há conta fora do mapa.',
        '7. Abas de dados (Balancete, Eliminacao) são as únicas editáveis; quadros estão protegidos.',
    ]
    for t in regras:
        cp.cell(r, 2, t).font = F_LIN
        r += 1
    r += 1
    cp.cell(r, 2, 'FÓRMULAS-CHAVE (manter)').font = Font(name='Arial', size=13, bold=True, color=GOLD)
    r += 1
    formulas = [
        'Linha de quadro:  =SUMIFS(Balancete!$F:$F; Balancete!$G:$G; linha; Balancete!$E:$E; "2025-12")',
        'Rota por empresa: adicionar  Balancete!$A:$A; "ALFA"',
        'Eliminação (BP):  =±SUMIFS(Eliminacao!$D:$D; Eliminacao!$F:$F; linha; Eliminacao!$C:$C; "2025-12")',
        'Eliminação (DRE): =−SUMIFS(... por Linha_DRE)   ·  Consolidado = Combinado − Eliminação',
        'DF mensal:        =SUMIFS(...; mês) − SUMIFS(...; mês anterior)  [delta do YTD]',
        'Roteamento:       =VLOOKUP(conta; DePara!$A:$E; 3|4|5; FALSE)  → DRE | BP | Nota',
        'Lucros acumulados: Σ(2223*) + Lucro Líquido da DRE (link DRE→BP)',
    ]
    for t in formulas:
        c = cp.cell(r, 2, t)
        c.font = Font(name='Consolas', size=9, color=NAVY2)
        r += 1
    r += 1
    cp.cell(r, 2, 'LEGENDA DE CORES').font = Font(name='Arial', size=13, bold=True, color=GOLD)
    r += 1
    legenda = [('Azul', 'Entrada de dados (editável)', F_IN),
               ('Preto', 'Fórmula (protegido — não editar)', F_LIN),
               ('Negrito navy', 'Total / subtotal', F_TOT),
               ('Verde', 'Check de amarração (deve ser ~0)', F_CHK)]
    for nome, desc, fnt in legenda:
        cp.cell(r, 2, nome).font = fnt
        cp.cell(r, 4, desc).font = F_LIN
        r += 1
    cp.column_dimensions['B'].width = 60
    cp.column_dimensions['D'].width = 95

    # ============================================================ proteção + ordem
    for name in ('Balanco', 'DRE', 'DF_Mensal', 'EBITDA', 'DePara', 'Notas', 'Parametros'):
        wb[name].protection.sheet = True   # sem senha: Revisão > Desproteger Planilha
    ordem = ['CAPA', 'Balanco', 'DRE', 'DF_Mensal', 'EBITDA',
             'Balancete', 'DePara', 'Eliminacao', 'dCalendario', 'Notas', 'Parametros']
    wb._sheets = [wb[n] for n in ordem]
    wb.active = 0

    wb.save(out)
    print(f'SALVO: {out}')

    # auditoria de anonimização
    strings = list(desc_of.values())
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith('='):
                    strings.append(c.value)
    hits = audit(strings, 'workbook')
    if hits:
        print('ATENÇÃO — vazamento potencial:')
        for h in sorted(hits)[:20]:
            print('  ', h)
    else:
        print('Auditoria de anonimização: OK (nenhum nome real encontrado)')

# ------------------------------------------------------------------ main
if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    depara, balancete = extrair(src)
    depara, balancete = normalizar(depara, balancete)
    preflight(depara, balancete)
    build(depara, balancete, out)
