# Modelo Financeiro Padrão — Grupo ABC (Aguiar Advisory)

Modelo de **bases financeiras conectadas** para controladoria: balancete único → mapa de contas (DePara) → quadros (Balanço, DRE, DF mensal, EBITDA) → notas explicativas → fonte pronta para **Power BI**.

> **Dados 100% fictícios**: nomes de empresas, bancos, lojas e sócios foram substituídos (ALFA/BETA/GAMA/DELTA, BANCO B1–B5, LOJA 01–12) e os valores foram escalados × 0,85. Nenhum dado de cliente real.

## Arquitetura (5 camadas)

| Camada | Aba | Papel |
|---|---|---|
| 1. Fonte | `Balancete` | Base única de saldos: **Empresa × Conta × Período** (formato longo, 24,5 mil lançamentos, Jan/23–Dez/25) |
| 2. Mapa | `DePara` | 562 contas (561 + conta de ajuste) → linha de DRE / linha de BP / **código de nota** (U1.1, U2.1...) |
| 3. Consolidação | `Eliminacao` | Eliminações intercompany por conta e período |
| 4. Quadros | `Balanco`, `DRE`, `DF_Mensal`, `EBITDA` | 100% fórmula (SUMIFS/VLOOKUP) — nada digitado |
| 5. Navegação | `CAPA`, `Notas`, `Parametros` | Índice com hyperlinks, regras de ouro, legenda de cores |

## Regras de ouro (o que não pode mudar)

1. **Chave universal: EMPRESA + PERÍODO + CONTA** — toda base nova entra por aqui.
2. **Quadros 100% por fórmula** — nada digitado; só as abas de dados são editáveis.
3. **Consolidado = Combinado − Eliminações**; a controladora é sempre a ALFA.
4. Contas de resultado no balancete são **acumuladas (YTD)** — o DF mensal usa delta; em janeiro o YTD reinicia (saldo de janeiro = fluxo do mês).
5. Toda conta nova **DEVE nascer no DePara** com destino (DRE/BP) e código de nota.
6. **Checks de amarração** (verde) devem fechar em ~0 — se não fechar, há conta fora do mapa.
7. Abas de quadros estão **protegidas** (sem senha: Revisão → Desproteger Planilha).

## Fórmulas-chave (manter)

```
Linha de quadro:   =SUMIFS(Balancete!$F:$F; Balancete!$G:$G; linha; Balancete!$E:$E; "2025-12")
Rota por empresa:  adicionar  Balancete!$A:$A; "ALFA"
Eliminação (BP):   =±SUMIFS(Eliminacao!$D:$D; Eliminacao!$F:$F; linha; Eliminacao!$C:$C; "2025-12")
Eliminação (DRE):  =−SUMIFS(... por Linha_DRE)   ·   Consolidado = Combinado − Eliminação
DF mensal:         =SUMIFS(...; mês) − SUMIFS(...; mês anterior)   [delta do YTD]
Roteamento:        =VLOOKUP(conta; DePara!$A:$E; 3|4|5; FALSE)  → DRE | BP | Nota
Lucros acumulados: Σ(2223*) + Lucro Líquido da DRE (link DRE→BP)
```

## Cores padrão Aguiar Advisory

- **Azul** = entrada de dados (editável) · **Preto** = fórmula (protegido) · **Navy `#0B1526`** = estrutura/totais · **Dourado `#C9A227`** = destaques · **Verde** = check de amarração.

## Como rodar

```bash
pip install openpyxl
python3 build_modelo_abc.py [fonte.xlsx] [saida.xlsx]
```

O script lê a planilha-fonte de consolidação, anonimiza (nomes + ×0,85), normaliza o balancete, monta o DePara, as eliminações, os quadros com fórmulas vivas e a CAPA — e roda uma **pré-verificação** (identidade contábil por empresa, cobertura de notas, trajetória das eliminações).

## Power BI (PBI-ready)

O Excel já sai pronto: tabelas nomeadas `tbl_Balancete`, `tbl_DePara`, `tbl_Eliminacao`, `tbl_Calendario` — carregue-as **sem Power Query**. Colunas `Data` (tipo data real) e `Fluxo` (movimento do mês; o saldo é YTD). Relacionamentos: `Calendario[Data] → Balancete[Data]`, `DePara[Conta] → Balancete[Conta]`, `Calendario[AnoMês] → Eliminacao[Período]`. Medidas DAX na coluna **Fluxo** (nunca `SUM(Saldo)` — saldo é semi-aditivo). Gabarito 2025: ROL Cons 112,5M · LL Cons 16,1M · EBITDA 20,8M · Check LL = 0.

## Particularidades documentadas

- Eliminações rastreadas a partir de Mai/2025 (Jan/24–Abr/25 = combinado).
- BETA apresenta resíduo de ~R$ 219 mil em Dez/24 na origem (contas fora do mapa na fonte) — o check do Balanço Dez/24 exibe o resíduo.
- `INVESTIMENTOS` (caixa-alta) vs `Investimentos`: SUMIFS é case-insensitive — não renomear linhas.

---
Aguiar Advisory · Controladoria estratégica, FP&A e auditoria
