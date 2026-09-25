"""Confere as funções do app contra a aba Gabarito do arquivo de exemplo."""

import pandas as pd
import pytest

import operacoes as op
from utils import avaliar_condicao_segura


def _atingimento_sudeste_janeiro(p):
    vendas, metas = p["Vendas"], p["Metas"]
    realizado = op.somases(vendas, "Valor_Total", [("Regiao", "Sudeste"), ("Data", "2026-01*")])
    meta = op.somases(metas, "Meta_Valor", [("Regiao", "Sudeste"), ("Mes", "2026-01-01")])
    return realizado / meta


def _ranking_bruno(p):
    totais = op.agrupar_resumir(p["Vendas"], ["Vendedor"], {"Valor_Total": "soma"})
    por_vendedor = dict(zip(totais["Vendedor"], totais["Valor_Total"]))
    todos = {v: por_vendedor.get(v, 0) for v in p["Vendedores"]["Vendedor"]}
    return 1 + sum(total > todos["Bruno"] for total in todos.values())


def _classificacao_linha_4(p):
    v = p["Vendas"]
    condicoes = [avaliar_condicao_segura("Valor_Total >= 1000", v), avaliar_condicao_segura("Valor_Total >= 500", v)]
    return op.funcao_se_aninhado(v, "Classe", condicoes, ["Alta", "Média"], "Baixa")["Classe"].iloc[2]


EXERCICIOS = {
    1: lambda p: p["Vendas"]["Valor_Total"].sum(),
    2: lambda p: p["Vendas"]["Valor_Total"].mean(),
    3: lambda p: p["Vendas"]["Valor_Total"].max() - p["Vendas"]["Valor_Total"].min(),
    4: lambda p: len(p["Vendas"]),
    5: lambda p: op.somase(p["Vendas"], "Valor_Total", "Regiao", "Sul"),
    6: lambda p: op.somase(p["Vendas"], "Quantidade", "Produto", "Azure"),
    7: lambda p: op.cont_ses(p["Vendas"], [("Vendedor", "Bruno"), ("Quantidade", ">=3")]),
    8: lambda p: op.executar_procv(p["Vendas"], p["Produtos"], "Produto", "Produto", ["Categoria"]).df["Categoria"].iloc[0],
    9: lambda p: op.executar_procx(p["Vendas"], p["Produtos"], "Produto", "Produto", ["Custo_Unitario"]).df["Custo_Unitario"].iloc[0],
    10: lambda p: (
        lambda d: (d["Quantidade"] * (d["Valor_Unitario"] - d["Custo_Unitario"])).sum()
    )(op.executar_procv(p["Vendas"], p["Produtos"], "Produto", "Produto", ["Custo_Unitario"]).df),
    11: lambda p: op.somase(p["Vendas"], "Valor_Total", "Vendedor", "Bruno")
    * op.executar_procv(pd.DataFrame({"Vendedor": ["Bruno"]}), p["Vendedores"], "Vendedor", "Vendedor", ["Comissao_Pct"]).df["Comissao_Pct"].iloc[0],
    12: lambda p: op.executar_procv(p["Vendas"], p["Centro_Custo"], "Cost Center", "Cost Center", ["Area"]).df["Area"].iloc[0],
    13: lambda p: op.executar_procv(p["Vendas"], p["Contas_GL"], "GL", "GL", ["Descricao_GL"]).df["Descricao_GL"].iloc[0],
    14: lambda p: op.somases(p["Vendas"], "Valor_Total", [("Data", "2026-01*")]),
    15: lambda p: op.data_extrair_componente(op.data_converter(p["Vendas"]["Data"]), "mes").iloc[0],
    16: lambda p: op.executar_procx(
        pd.DataFrame({"Vendedor": ["Paulo"]}), p["Vendedores"], "Vendedor", "Vendedor", ["Comissao_Pct"], valor_nao_encontrado="Não cadastrado"
    ).df["Comissao_Pct"].iloc[0],
    17: _classificacao_linha_4,
    18: _atingimento_sudeste_janeiro,
    19: lambda p: op.executar_procx(
        pd.DataFrame({"Atingimento": [_atingimento_sudeste_janeiro(p)]}), p["Faixas_Bonus"], "Atingimento", "Atingimento_Min", ["Bonus_Pct"], "aproximada"
    ).df["Bonus_Pct"].iloc[0],
    20: lambda p: op.somase(p["Vendas"], "Valor_Total", "Regiao", "Nordeste") / p["Vendas"]["Valor_Total"].sum(),
    21: lambda p: op.somase(
        op.executar_procv(p["Vendas"], p["Produtos"], "Produto", "Produto", ["Categoria"]).df, "Valor_Total", "Categoria", "Produtividade"
    ),
    22: _ranking_bruno,
    24: lambda p: op.somase(p["Vendas"], "Valor_Total", "Cost Center", "CC200")
    / op.executar_procx(pd.DataFrame({"CC": ["CC200"]}), p["Centro_Custo"], "CC", "Cost Center", ["Meta_Receita_Semestre"]).df["Meta_Receita_Semestre"].iloc[0],
}


@pytest.mark.parametrize("numero", sorted(EXERCICIOS))
def test_exercicio_do_gabarito(numero, planilha, gabarito):
    obtido = EXERCICIOS[numero](planilha)
    esperado = gabarito[numero]
    if isinstance(esperado, str):
        assert obtido == esperado
    else:
        assert float(obtido) == pytest.approx(float(esperado), rel=1e-9)


def test_exercicio_25_tabela_dinamica_por_mes(planilha):
    vendas = op.data_criar_coluna_periodo(planilha["Vendas"], "Data", "mes", "Mes")
    tabela = op.criar_tabela_dinamica(vendas, ["Regiao"], ["Mes"], ["Valor_Total"], "soma", None, True).set_index("Regiao")

    datas = pd.to_datetime(planilha["Vendas"]["Data"], format="%Y-%m-%d")
    esperado = planilha["Vendas"].assign(Mes=datas.dt.to_period("M").astype(str)).pivot_table(
        index="Regiao", columns="Mes", values="Valor_Total", aggfunc="sum"
    )
    for mes in esperado.columns:
        for regiao in esperado.index:
            assert tabela.loc[regiao, f"Valor_Total | {mes}"] == esperado.loc[regiao, mes]
    assert len([c for c in tabela.columns if c.startswith("Valor_Total | 2026")]) == 6
    assert tabela.loc["Total Geral", "Valor_Total | Total Geral"] == 90000
