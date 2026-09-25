"""Executa as opções 1 a 30 do menu com entradas simuladas e confere cada resultado de forma independente."""

import builtins
import contextlib
import copy
import io
import tempfile
import warnings
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import pytest

import interface as ui
import main as app
from estado import Sessao

warnings.filterwarnings("ignore")

TMP = Path(tempfile.mkdtemp())
XLSX = Path(__file__).resolve().parent.parent / "Dataset_Financas_Complementar.xlsx"
CSV = TMP / "atingimento.csv"
CSV.write_text("Vendedor;Atingimento;Valor;Obs\nBruno;1,153846;1.500;x\nAna;0,95;;\nCarlos;0,5;2.300,50;y\nJulia;1,2;10;\n", encoding="utf-8")
V = pd.read_excel(XLSX, "Vendas")
P = pd.read_excel(XLSX, "Produtos")
VD = pd.read_excel(XLSX, "Vendedores")


class FimEntradas(Exception):
    pass


def rodar(func, entradas, *args):
    fila = list(map(str, entradas))
    prompts = []

    def entrada_simulada(msg=""):
        prompts.append(msg)
        if not fila:
            raise FimEntradas(f"entradas acabaram no prompt: {msg!r}")
        return fila.pop(0)

    saida = io.StringIO()
    original = builtins.input
    builtins.input = entrada_simulada
    try:
        with contextlib.redirect_stdout(saida):
            func(*args)
    finally:
        builtins.input = original
    assert not fila, f"sobraram entradas {fila}; últimos prompts: {prompts[-3:]}"
    return saida.getvalue()


# Importação: xlsx (todas as abas) + CSV brasileiro
base = Sessao()
rodar(ui.fluxo_importacao, [XLSX, "s", "s", CSV, "n"], base)
ids = {d.aba or d.arquivo: d.id for d in base.listar_datasets()}
VEN, PRO, VDD, FAI, CSVID = ids["Vendas"], ids["Produtos"], ids["Vendedores"], ids["Faixas_Bonus"], ids["atingimento.csv"]
col = lambda df, nome: list(df.columns).index(nome) + 1
cv = lambda n: col(V, n)

CASOS = []
def caso(num, nome, entradas, conferir, handler=None, sessao=None):
    CASOS.append(pytest.param(num, entradas, conferir, handler, sessao, id=f"{num:02d} {nome}"))


ds = lambda s, i: s.obter_dataset(i).df
ultimo = lambda s: list(s.resultados.values())[-1]
def igual(a, b, msg=""):
    if isinstance(a, float) or isinstance(b, float):
        assert abs(float(a) - float(b)) < 1e-9, f"{msg} {a} != {b}"
    else: assert a == b, f"{msg} {a!r} != {b!r}"

# 0 importação
def c_imp(s, out):
    d = ds(s, CSVID); igual(d.Valor.tolist()[0], 1500.0); igual(d.Valor.tolist()[2], 2300.5); igual(d.Atingimento.tolist()[0], 1.153846)
    igual(len(ds(s, VEN)), 120)
caso(0, "Importar XLSX (11 abas) + CSV brasileiro", [], c_imp, handler=lambda s: None, sessao=base)

caso(1, "Catálogo", [], lambda s, o: (None if "data (texto)" in o and "numérico" in o else (_ for _ in ()).throw(AssertionError("tipos"))))
caso(2, "Visualizar + estatística", [VEN], lambda s, o: None if "750" in o else (_ for _ in ()).throw(AssertionError("média 750 ausente")))
esperado3 = V[(V.Regiao.str.lower() == "sul") & (V.Valor_Total >= 1000)]
caso(3, "Filtrar: Regiao igual 'sul' E Valor_Total >= 1.000", [VEN, cv("Regiao"), "igual", "sul", "s", cv("Valor_Total"), "maior_igual", "1.000", "n", "E", "s"],
     lambda s, o: igual(len(ds(s, VEN)), len(esperado3)) or igual(ds(s, VEN).Valor_Total.sum(), esperado3.Valor_Total.sum()))
caso(3, "Filtrar: em_lista Vendedor Bruno, Ana (OU com Quantidade 5)", [VEN, cv("Vendedor"), "em_lista", "bruno, ANA", "s", cv("Quantidade"), "igual", "5", "n", "OU", "s"],
     lambda s, o: igual(len(ds(s, VEN)), int((V.Vendedor.isin(["Bruno","Ana"]) | (V.Quantidade == 5)).sum())))
caso(4, "Ordenar por Valor_Total decrescente e Data", [VEN, f"{cv('Valor_Total')},{cv('Data')}", "n", "s", "s"],
     lambda s, o: igual(ds(s, VEN).Valor_Total.tolist(), V.sort_values(["Valor_Total","Data"], ascending=[False, True], kind="mergesort").Valor_Total.tolist()))
caso(5, "Renomear Valor_Total", [VEN, "3", cv("Valor_Total"), "VT", "n", "s"], lambda s, o: igual(ds(s, VEN).VT.sum(), 90000))
caso(5, "Remover colunas", [VEN, "4", f"{cv('GL')},{cv('Cost Center')}", "s", "s"], lambda s, o: igual(ds(s, VEN).shape[1], 7))
caso(6, "Coluna calculada: Quantidade × Valor_Unitario", [VEN, "1", "Total2", f"{cv('Quantidade')},{cv('Valor_Unitario')}", "multiplicacao", "s"],
     lambda s, o: igual((ds(s, VEN).Total2 == V.Valor_Total).all(), True))
caso(6, "Coluna calculada: percentual", [VEN, "2", "Pct", cv("Valor_Unitario"), cv("Valor_Total"), "s"],
     lambda s, o: igual(ds(s, VEN).Pct.iloc[0], V.Valor_Unitario[0] / V.Valor_Total[0] * 100))
caso(6, "Coluna calculada: SE com número", [VEN, "3", "Bonus", "Valor_Total >= 1000 and Regiao == \"Nordeste\"", "100", "0", "s"],
     lambda s, o: igual(ds(s, VEN).Bonus.sum(), 100 * int(((V.Valor_Total >= 1000) & (V.Regiao == "Nordeste")).sum())))
caso(6, "Coluna calculada: concatenação", [VEN, "4", "Chave", f"{cv('Regiao')},{cv('Produto')}", "-", "s"], lambda s, o: igual(ds(s, VEN).Chave[0], "Sudeste-Copilot"))
caso(6, "Coluna calculada: texto maiúsculas", [VEN, "5", "R", cv("Regiao"), "maiusculas", "s"], lambda s, o: igual(ds(s, VEN).R[0], "SUDESTE"))
caso(6, "Coluna calculada: diferença de datas (dias)", [VEN, "6", "Dias", cv("Data"), cv("Data"), "dias", "s"], lambda s, o: igual(int(ds(s, VEN).Dias.sum()), 0))
caso(6, "Coluna calculada: valor fixo 0,5", [VEN, "7", "Fixo", "0,5", "s"], lambda s, o: igual(ds(s, VEN).Fixo.sum(), 60.0))
caso(6, "Coluna calculada: expressão", [VEN, "8", "Lucro", "Valor_Total - Quantidade * 10", "s"], lambda s, o: igual(ds(s, VEN).Lucro.sum(), 90000 - V.Quantidade.sum() * 10))
caso(7, "Matemática: dividir por 1.000", [VEN, cv("Valor_Total"), "4", "1.000", "Mil", "s"], lambda s, o: igual(ds(s, VEN).Mil.sum(), 90.0))
caso(7, "Matemática: arredondar Valor_Total/7 com 2 casas", [CSVID, col(ds(base, CSVID), "Atingimento"), "8", "2", "Arr", "s"],
     lambda s, o: igual(ds(s, CSVID).Arr.tolist(), [1.15, 0.95, 0.5, 1.2]))
caso(8, "Texto: esquerda 3", [VEN, "2", cv("Produto"), "P3", "3", "s"], lambda s, o: igual(ds(s, VEN).P3[0], "Cop"))
caso(8, "Texto: separar Data por '-'", [VEN, "11", cv("Data"), "-", "s"], lambda s, o: igual(ds(s, VEN)["Data_parte_2"][0], "01"))
caso(8, "Texto: separar Produto por espaço", [VEN, "11", cv("Produto"), " ", "s"], lambda s, o: igual(ds(s, VEN)["Produto_parte_1"][2], "Office"))
caso(6, "Coluna calculada: concatenar com espaço", [VEN, "4", "C2", f"{cv('Vendedor')},{cv('Regiao')}", " ", "s"], lambda s, o: igual(ds(s, VEN).C2[0], "Bruno Sudeste"))
caso(9, "Data: extrair mês", [VEN, "2", cv("Data"), "mes", "Mes", "s"], lambda s, o: igual(ds(s, VEN).Mes.tolist(), pd.to_datetime(V.Data, format="%Y-%m-%d").dt.month.tolist()))
caso(9, "Data: somar 30 dias", [VEN, "4", cv("Data"), "30", "D30", "s"], lambda s, o: igual(ds(s, VEN).D30[0], pd.Timestamp("2026-02-01")))
caso(9, "Data: coluna de período mês", [VEN, "7", cv("Data"), "mes", "Per", "s"], lambda s, o: igual(ds(s, VEN).Per.value_counts().to_dict(), pd.to_datetime(V.Data).dt.to_period("M").astype(str).value_counts().to_dict()))
caso(10, "Agrupar Regiao: soma Valor_Total", [VEN, cv("Regiao"), cv("Valor_Total"), "soma", "n", "resumo", "s"],
     lambda s, o: igual(dict(zip(s.listar_datasets()[-1].df.Regiao, s.listar_datasets()[-1].df.Valor_Total)), V.groupby("Regiao").Valor_Total.sum().to_dict()))
def c11(s, o):
    t = ultimo(s).set_index("Regiao"); ref = V.pivot_table(index="Regiao", columns="Produto", values="Valor_Total", aggfunc="sum", margins=True, margins_name="Total Geral")
    for r in ref.index:
        for c in ref.columns:
            if pd.notna(ref.loc[r, c]): igual(t.loc[r, f"Valor_Total | {c}"], ref.loc[r, c], f"{r}/{c}")
caso(11, "Tabela dinâmica Regiao × Produto com totais", [VEN, cv("Regiao"), "s", cv("Produto"), cv("Valor_Total"), "soma", "n", "s", "piv", "s"], c11)
caso(12, "PROCV Vendas → Produtos (Categoria, Custo)", [VEN, PRO, cv("Produto"), 1, f"{col(P,'Categoria')},{col(P,'Custo_Unitario')}", "", "s"],
     lambda s, o: igual(ds(s, VEN).Custo_Unitario.tolist(), V.Produto.map(P.set_index("Produto").Custo_Unitario).tolist()) or igual(len(ds(s, VEN)), 120))
caso(12, "PROCV GL (número) → Contas_GL", [VEN, ids["Contas_GL"], cv("GL"), 1, 2, "", "s"], lambda s, o: igual(ds(s, VEN).Descricao_GL[0], "Receita de Serviços em Nuvem"))
caso(13, "PROCH na aba Faixas_Bonus", [FAI, 0, "Abaixo da meta", 1, "proch", "s"], lambda s, o: igual(ultimo(s).valor_retornado[0], "Próximo da meta"))
caso(14, "PROCX exato Vendas → Vendedores (Comissão, não encontrado 0)", [VEN, VDD, cv("Vendedor"), 1, col(VD, "Comissao_Pct"), "exata", "primeira", "s", "0", "s"],
     lambda s, o: igual(ds(s, VEN).Comissao_Pct.tolist(), V.Vendedor.map(VD.set_index("Vendedor").Comissao_Pct).tolist()))
caso(14, "PROCX aproximado atingimento → Faixas_Bonus", [CSVID, FAI, 2, 1, "2,3", "aproximada", "n", "s"],
     lambda s, o: igual(ds(s, CSVID).Bonus_Pct.tolist(), [0.05, 0.02, 0.0, 0.08]) or igual(ds(s, CSVID).Classificacao[1], "Próximo da meta"))
caso(15, "SOMASE Regiao 'Sul' (ex. 5)", [VEN, cv("Valor_Total"), cv("Regiao"), "Sul", "", "s"], lambda s, o: igual(ultimo(s).Resultado[0], 9000.0))
caso(16, "SOMASES Sudeste + jan/2026", [VEN, cv("Valor_Total"), cv("Regiao"), "Sudeste", "s", cv("Data"), "2026-01*", "n", "", "s"], lambda s, o: igual(ultimo(s).Resultado[0], 3000.0))
caso(17, "CONT.SE Valor_Total '>1.000'", [VEN, cv("Valor_Total"), ">1.000", "", "s"], lambda s, o: igual(ultimo(s).Resultado[0], int((V.Valor_Total > 1000).sum())))
caso(18, "CONT.SES Bruno e Quantidade >=3 (ex. 7)", [VEN, cv("Vendedor"), "Bruno", "s", cv("Quantidade"), ">=3", "n", "", "s"], lambda s, o: igual(ultimo(s).Resultado[0], 18))
caso(19, "MÉDIASE Norte", [VEN, cv("Valor_Total"), cv("Regiao"), "Norte", "", "s"], lambda s, o: igual(ultimo(s).Resultado[0], V[V.Regiao == "Norte"].Valor_Total.mean()))
caso(20, "MÉDIASES Norte e Quantidade <3", [VEN, cv("Valor_Total"), cv("Regiao"), "Norte", "s", cv("Quantidade"), "<3", "n", "", "s"],
     lambda s, o: igual(ultimo(s).Resultado[0], V[(V.Regiao == "Norte") & (V.Quantidade < 3)].Valor_Total.mean()))
caso(21, "SE Valor_Total >= 1000 → 0,05 / 0", [VEN, "Valor_Total >= 1000", "0,05", "0", "Taxa", "s"],
     lambda s, o: igual(ds(s, VEN).Taxa.sum(), 0.05 * int((V.Valor_Total >= 1000).sum())))
caso(22, "SE aninhado Alta/Média/Baixa (ex. 17)", [VEN, "Valor_Total >= 1000", "Alta", "Valor_Total >= 500", "Média", "", "Baixa", "Classe", "s"],
     lambda s, o: igual(ds(s, VEN).Classe.tolist(), np.where(V.Valor_Total >= 1000, "Alta", np.where(V.Valor_Total >= 500, "Média", "Baixa")).tolist()))
caso(23, "Preencher nulos com média", [CSVID, "2", col(ds(base, CSVID), "Valor"), "media", "s"], lambda s, o: igual(ds(s, CSVID).Valor[1], (1500 + 2300.5 + 10) / 3))
caso(23, "Preencher nulos com valor 0", [CSVID, "2", col(ds(base, CSVID), "Valor"), "valor", "0", "s"], lambda s, o: igual(ds(s, CSVID).Valor.sum(), 3810.5))
caso(23, "Remover linhas nulas (todas as colunas)", [CSVID, "1", "s", "s"], lambda s, o: igual(len(ds(s, CSVID)), 2))
caso(24, "Remover duplicados por Regiao", [VEN, "n", cv("Regiao"), "primeira", "s"], lambda s, o: igual(ds(s, VEN).Regiao.tolist(), V.drop_duplicates("Regiao").Regiao.tolist()))
caso(25, "Concatenar Vendas + Vendas", [f"{VEN},{VEN}", "", "s"], lambda s, o: igual(s.listar_datasets()[-1].df.Valor_Total.sum(), 180000))
caso(26, "Mesclar Vendas + Produtos (inner)", [VEN, PRO, cv("Produto"), 1, "inner", "", "s"],
     lambda s, o: igual(len(s.listar_datasets()[-1].df), 120) or igual(s.listar_datasets()[-1].df.Categoria.tolist(), V.Produto.map(P.set_index("Produto").Categoria).tolist()))
caso(27, "Validar qualidade", [CSVID, "s"], lambda s, o: igual(int(ultimo(s).Nulos.sum()), 3))
def c28():
    s = copy.deepcopy(base)
    rodar(ui.handler_filtrar, [VEN, cv("Regiao"), "igual", "Sul", "n", "s"], s)
    assert len(ds(s, VEN)) == 30
    return s
caso(28, "Desfazer filtro", [VEN], lambda s, o: igual(ds(s, VEN).equals(V), True), sessao=c28())
caso(29, "Histórico", [], lambda s, o: None if "Filtrar dados" in o else (_ for _ in ()).throw(AssertionError(o[:200])), sessao=c28())
def s30():
    s = copy.deepcopy(base)
    rodar(ui.handler_somase, [VEN, cv("Valor_Total"), cv("Regiao"), "Sul", "", "s"], s)
    rodar(ui.handler_se, [VEN, "Valor_Total >= 1000", "100", "0", "Bonus", "s"], s)
    return s
def c30(s, o):
    wb = openpyxl.load_workbook(TMP / "relatorio.xlsx")
    nomes = wb.sheetnames
    for obrig in ["Resumo", "Catalogo_Colunas", "Qualidade_Dados", "Historico", "Vendas_orig", "Vendas_alt", "atingimento_orig"]:
        assert obrig in nomes, f"aba {obrig} ausente: {nomes}"
    alt = pd.read_excel(TMP / "relatorio.xlsx", "Vendas_alt")
    igual(alt.Valor_Total.sum(), 90000); igual(alt.Bonus.sum(), 100 * int((V.Valor_Total >= 1000).sum()))
    assert pd.api.types.is_numeric_dtype(alt.Bonus), "Bonus gravado como texto"
    at = pd.read_excel(TMP / "relatorio.xlsx", "atingimento_orig"); igual(at.Valor.tolist()[0], 1500)
    som = [n for n in nomes if n not in ("Resumo","Catalogo_Colunas","Qualidade_Dados","Historico") and not n.endswith(("_orig","_alt"))][0]; igual(pd.read_excel(TMP / "relatorio.xlsx", som).Resultado[0], 9000)
    formulas = sum(1 for ws in wb.worksheets for r in ws.iter_rows() for c in r if c.data_type == "f")
    igual(formulas, 0, "fórmulas")
caso(30, "Salvar relatório e reler o .xlsx", [TMP / "relatorio.xlsx"], c30, sessao=s30())


@pytest.mark.parametrize("num, entradas, conferir, handler, sessao", CASOS)
def test_opcao_do_menu(num, entradas, conferir, handler, sessao):
    s = sessao or copy.deepcopy(base)
    saida = rodar(handler or app.DESPACHO[str(num)], entradas, s)
    assert "erro inesperado" not in saida.lower() and "Erro:" not in saida, saida[-400:]
    conferir(s, saida)
