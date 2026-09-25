"""Casos em que o app calculava valores errados sem avisar."""

import numpy as np
import pandas as pd
import pytest

import carregador
import operacoes as op
from operacoes import Criterio
from utils import arredondar_excel, avaliar_expressao_segura, converter_datas_br, converter_numero_br, interpretar_valor_digitado


# Datas ----------------------------------------------------------------------


def test_datas_iso_nao_trocam_dia_e_mes():
    datas = converter_datas_br(pd.Series(["2026-01-02", "2026-01-03 10:30:00", "02/01/2026", "13/01/2026", None, "xx"]))
    assert datas.dt.strftime("%Y-%m-%d").tolist()[:4] == ["2026-01-02", "2026-01-03", "2026-01-02", "2026-01-13"]
    assert datas.iloc[4:].isna().all()
    assert converter_datas_br("2026-01-02") == pd.Timestamp("2026-01-02")


def test_filtro_de_datas_com_valor_iso():
    df = pd.DataFrame({"D": pd.to_datetime(["2026-01-03", "2026-01-20", "2026-03-01"])})
    assert len(op.aplicar_filtros(df, [Criterio("D", "menor", "2026-01-05")])) == 1
    assert op.cont_se(df, "D", ">2026-01-05") == 2


def test_diferenca_de_datas_em_texto_iso():
    df = pd.DataFrame({"a": ["2026-01-02", "2026-01-03"], "b": ["2026-01-03", "2026-02-10"]})
    assert op.criar_coluna_diferenca_datas(df, "dias", "a", "b")["dias"].tolist() == [1, 38]


# CSV ------------------------------------------------------------------------


def _ler_csv(tmp_path, conteudo):
    caminho = tmp_path / "dados.csv"
    caminho.write_text(conteudo, encoding="utf-8")
    return carregador.carregar_csv(caminho)[0]


def test_csv_brasileiro_com_ponto_de_milhar(tmp_path):
    df = _ler_csv(tmp_path, "Valor;Preco;Qtd\n1.500;1.234,56;2\n2.300;10,50;3\n10.000;0,99;1\n")
    assert df["Valor"].tolist() == [1500, 2300, 10000]
    assert df["Preco"].tolist() == [1234.56, 10.5, 0.99]
    assert df["Qtd"].tolist() == [2, 3, 1]


def test_csv_separado_por_virgula_usa_ponto_decimal(tmp_path):
    df = _ler_csv(tmp_path, 'Taxa,Valor\n1.500,"1.234,56"\n0.25,"10,50"\n')
    assert df["Taxa"].tolist() == [1.5, 0.25]
    assert df["Valor"].tolist() == [1234.56, 10.5]


def test_csv_mantem_codigos_e_numeros_longos_como_texto(tmp_path):
    df = _ler_csv(tmp_path, "Conta;Cartao;V\n00123;4111111111111111;1\n04567;;2\n")
    assert df["Conta"].tolist() == ["00123", "04567"]
    assert df["Cartao"].iloc[0] == "4111111111111111"


def test_numero_com_zero_antes_do_ponto_nao_vira_milhar():
    assert converter_numero_br("0.500") == 0.5
    assert converter_numero_br("1.000") == 1000
    assert converter_numero_br("1.234,56") == 1234.56


# PROCV / PROCX ----------------------------------------------------------------


def test_procv_chave_inteira_contra_decimal():
    principal = pd.DataFrame({"k": [1, 2, 4]})
    consulta = pd.DataFrame({"k": [1.0, 2.0, np.nan], "v": ["a", "b", "vazio"]})
    assert op.executar_procv(principal, consulta, "k", "k", ["v"]).df["v"].tolist()[:2] == ["a", "b"]
    assert pd.isna(op.executar_procv(principal, consulta, "k", "k", ["v"]).df["v"].iloc[2])


def test_procv_e_procx_chave_vazia_nao_encontra_nada():
    principal = pd.DataFrame({"k": [1.0, np.nan]})
    consulta = pd.DataFrame({"k": [1.0, np.nan], "v": ["a", "vazio"]})
    assert pd.isna(op.executar_procv(principal, consulta, "k", "k", ["v"]).df["v"].iloc[1])
    assert pd.isna(op.executar_procx(principal, consulta, "k", "k", ["v"]).df["v"].iloc[1])


def test_procv_texto_ignora_maiusculas_e_espacos_e_mantem_codigos():
    consulta = pd.DataFrame({"k": ["Sul", "00123"], "v": [1, 2]})
    principal = pd.DataFrame({"k": ["sul", " Sul ", "123", "00123"]})
    resultado = op.executar_procv(principal, consulta, "k", "k", ["v"]).df["v"]
    assert resultado.iloc[:2].tolist() == [1, 1]
    assert pd.isna(resultado.iloc[2])
    assert resultado.iloc[3] == 2


def test_procv_mantem_ordem_e_quantidade_de_linhas():
    principal = pd.DataFrame({"k": ["b", "a", "b"], "x": [1, 2, 3]})
    consulta = pd.DataFrame({"k": ["a", "b", "b"], "v": [10, 20, 30]})
    resultado = op.executar_procv(principal, consulta, "k", "k", ["v"])
    assert resultado.df["v"].tolist() == [20, 10, 20]
    assert resultado.df["x"].tolist() == [1, 2, 3]


def test_procx_valor_nao_encontrado_numerico():
    resultado = op.executar_procx(
        pd.DataFrame({"k": [1, 9]}), pd.DataFrame({"k": [1], "v": [5.0]}), "k", "k", ["v"], valor_nao_encontrado=0
    ).df["v"]
    assert resultado.tolist() == [5.0, 0]
    assert pd.api.types.is_numeric_dtype(resultado)


# Tabela dinâmica -----------------------------------------------------------------


def test_total_geral_da_tabela_dinamica_com_vazios():
    df = pd.DataFrame({"R": ["A", "A", "B", "B"], "X": [1.0, np.nan, 3.0, 4.0], "Y": [10.0, 20.0, np.nan, 40.0]})
    tabela = op.criar_tabela_dinamica(df, ["R"], None, ["X", "Y"], "soma", None, True).set_index("R")
    assert tabela.loc["Total Geral", "X"] == 8
    assert tabela.loc["Total Geral", "Y"] == 70
    assert tabela.loc["A", "Y"] == 30


def test_tabela_dinamica_mantem_grupo_vazio_no_total():
    df = pd.DataFrame({"R": ["A", None], "X": [1.0, 4.0]})
    tabela = op.criar_tabela_dinamica(df, ["R"], None, ["X"], "soma", None, True)
    assert tabela["X"].iloc[-1] == 5


# Arredondamento ------------------------------------------------------------------


@pytest.mark.parametrize(
    "valor, casas, esperado",
    [(0.5, 0, 1), (1.5, 0, 2), (2.5, 0, 3), (-2.5, 0, -3), (0.125, 2, 0.13), (2.675, 2, 2.68), (1.005, 2, 1.01), (1234.5, -2, 1200)],
)
def test_arredondar_como_o_excel(valor, casas, esperado):
    assert arredondar_excel(valor, casas) == esperado
    df = pd.DataFrame({"x": [valor]})
    assert op.aplicar_operacao_matematica(df, "x", "arredondar", "y", casas)["y"].iloc[0] == esperado


def test_round_na_expressao_arredonda_como_o_excel():
    df = pd.DataFrame({"x": [2.5, 0.125]})
    assert avaliar_expressao_segura("round(x)", df).tolist() == [3, 0]
    assert avaliar_expressao_segura("round(x, 2)", df).tolist() == [2.5, 0.13]


# SE e valores digitados ------------------------------------------------------------


def test_valor_digitado_vira_numero_quando_e_numero():
    assert interpretar_valor_digitado("100") == 100
    assert interpretar_valor_digitado("0,05") == 0.05
    assert interpretar_valor_digitado("1.234,56") == 1234.56
    assert interpretar_valor_digitado("Alta") == "Alta"
    assert interpretar_valor_digitado('"100"') == "100"


def test_se_com_numeros_gera_coluna_numerica():
    df = pd.DataFrame({"v": [10, 1000]})
    condicao = df["v"] > 500
    resultado = op.funcao_se(df, "bonus", condicao, 100, 0)["bonus"]
    assert resultado.tolist() == [0, 100]
    assert pd.api.types.is_numeric_dtype(resultado)
    assert op.somase(op.funcao_se(df, "bonus", condicao, 100, 0), "bonus", "v", ">0") == 100


def test_se_aninhado_mistura_numero_e_texto_sem_virar_texto():
    df = pd.DataFrame({"v": [10, 600, 1000]})
    resultado = op.funcao_se_aninhado(df, "c", [df["v"] >= 1000, df["v"] >= 500], [0.08, 0.05], "sem bônus")["c"]
    assert resultado.tolist() == ["sem bônus", 0.05, 0.08]
    assert isinstance(resultado.iloc[1], float)
