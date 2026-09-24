"""Funções de transformação e análise equivalentes às principais funções do Excel.

Todas as funções desta camada são puras em relação a DataFrames: recebem um
ou mais DataFrames e parâmetros já validados, e retornam um novo DataFrame
(ou valor) sem alterar os originais. Erros de negócio (dataset inexistente,
coluna inexistente, tipos incompatíveis, divisão por zero etc.) são
sinalizados com :class:`~excel_toolkit.utils.ErroOperacao`. Avisos que não
impedem a execução (chaves duplicadas em um PROCV, valores sem
correspondência etc.) são retornados junto ao resultado para exibição na
interface e registro no histórico.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

import numpy as np
import pandas as pd

from utils import ErroOperacao, curinga_para_regex

# ---------------------------------------------------------------------------
# Utilidades de validação
# ---------------------------------------------------------------------------

AGREGACOES: dict[str, str] = {
    "soma": "sum",
    "media": "mean",
    "contagem": "count",
    "minimo": "min",
    "maximo": "max",
    "mediana": "median",
    "desvio_padrao": "std",
    "primeiro": "first",
    "ultimo": "last",
}


def validar_colunas(df: pd.DataFrame, colunas: list[str]) -> None:
    """Garante que todas as colunas informadas existem no DataFrame."""
    faltantes = [c for c in colunas if c not in df.columns]
    if faltantes:
        raise ErroOperacao(f"Coluna(s) não encontrada(s): {', '.join(str(c) for c in faltantes)}")


def validar_coluna_numerica(df: pd.DataFrame, coluna: str) -> None:
    """Garante que a coluna existe e é numérica."""
    validar_colunas(df, [coluna])
    if not pd.api.types.is_numeric_dtype(df[coluna]):
        raise ErroOperacao(f"A coluna '{coluna}' não é numérica.")


# ---------------------------------------------------------------------------
# Filtros
# ---------------------------------------------------------------------------

OPERADORES_FILTRO = (
    "igual",
    "diferente",
    "maior",
    "maior_igual",
    "menor",
    "menor_igual",
    "contem",
    "nao_contem",
    "comeca_com",
    "termina_com",
    "vazio",
    "nao_vazio",
    "em_lista",
    "entre",
)


@dataclass
class Criterio:
    """Representa um critério de filtro sobre uma coluna."""

    coluna: str
    operador: str
    valor: Any = None
    valor2: Any = None


def construir_mascara(df: pd.DataFrame, criterio: Criterio) -> pd.Series:
    """Constrói a máscara booleana correspondente a um único critério de filtro."""
    validar_colunas(df, [criterio.coluna])
    if criterio.operador not in OPERADORES_FILTRO:
        raise ErroOperacao(f"Operador de filtro desconhecido: '{criterio.operador}'")
    serie = df[criterio.coluna]

    if criterio.operador == "vazio":
        return serie.isna()
    if criterio.operador == "nao_vazio":
        return serie.notna()

    if criterio.operador == "em_lista":
        valores = criterio.valor if isinstance(criterio.valor, (list, tuple, set)) else [criterio.valor]
        return serie.isin(valores)

    if criterio.operador == "entre":
        inicio, fim = criterio.valor, criterio.valor2
        try:
            serie_num = pd.to_numeric(serie, errors="coerce")
            inicio_num, fim_num = float(inicio), float(fim)
            return serie_num.between(min(inicio_num, fim_num), max(inicio_num, fim_num))
        except (TypeError, ValueError):
            return serie.astype(str).between(str(inicio), str(fim))

    if criterio.operador in ("contem", "nao_contem", "comeca_com", "termina_com"):
        serie_texto = serie.astype(str)
        alvo = str(criterio.valor)
        if criterio.operador == "contem":
            mascara = serie_texto.str.contains(re.escape(alvo), case=False, na=False)
        elif criterio.operador == "nao_contem":
            mascara = ~serie_texto.str.contains(re.escape(alvo), case=False, na=False)
        elif criterio.operador == "comeca_com":
            mascara = serie_texto.str.startswith(alvo, na=False)
        else:
            mascara = serie_texto.str.endswith(alvo, na=False)
        return mascara

    # Operadores de comparação: tenta numérico, cai para texto/data se necessário.
    valor_comparacao = criterio.valor
    if pd.api.types.is_numeric_dtype(serie):
        try:
            valor_comparacao = float(criterio.valor)
        except (TypeError, ValueError):
            raise ErroOperacao(
                f"O valor '{criterio.valor}' não é compatível com a coluna numérica '{criterio.coluna}'."
            ) from None
    elif pd.api.types.is_datetime64_any_dtype(serie):
        valor_comparacao = pd.to_datetime(criterio.valor, errors="coerce")

    if criterio.operador == "igual":
        return serie == valor_comparacao
    if criterio.operador == "diferente":
        return serie != valor_comparacao
    if criterio.operador == "maior":
        return serie > valor_comparacao
    if criterio.operador == "maior_igual":
        return serie >= valor_comparacao
    if criterio.operador == "menor":
        return serie < valor_comparacao
    return serie <= valor_comparacao


def aplicar_filtros(df: pd.DataFrame, criterios: list[Criterio], operador_logico: str = "E") -> pd.DataFrame:
    """Aplica uma lista de critérios combinados com E (AND) ou OU (OR)."""
    if not criterios:
        raise ErroOperacao("Informe ao menos um critério de filtro.")
    mascaras = [construir_mascara(df, c) for c in criterios]
    mascara_final = mascaras[0]
    for mascara in mascaras[1:]:
        mascara_final = (mascara_final & mascara) if operador_logico.upper() == "E" else (mascara_final | mascara)
    return df[mascara_final].copy()


# ---------------------------------------------------------------------------
# Ordenação e colunas
# ---------------------------------------------------------------------------


def ordenar_dados(df: pd.DataFrame, colunas: list[str], ascendente: list[bool]) -> pd.DataFrame:
    """Ordena o DataFrame por uma ou várias colunas."""
    validar_colunas(df, colunas)
    if len(colunas) != len(ascendente):
        raise ErroOperacao("A quantidade de colunas e de direções de ordenação deve ser igual.")
    return df.sort_values(by=colunas, ascending=ascendente, kind="mergesort").reset_index(drop=True)


def selecionar_colunas(df: pd.DataFrame, colunas: list[str]) -> pd.DataFrame:
    """Retorna uma cópia do DataFrame apenas com as colunas informadas, na ordem dada."""
    validar_colunas(df, colunas)
    return df[colunas].copy()


def reordenar_colunas(df: pd.DataFrame, ordem: list[str]) -> pd.DataFrame:
    """Reordena as colunas do DataFrame conforme a lista informada (deve conter todas as colunas)."""
    if set(ordem) != set(df.columns):
        raise ErroOperacao("A nova ordem deve conter exatamente as mesmas colunas do dataset.")
    return df[ordem].copy()


def renomear_colunas(df: pd.DataFrame, mapeamento: dict[str, str]) -> pd.DataFrame:
    """Renomeia colunas conforme o mapeamento {nome_atual: novo_nome}."""
    validar_colunas(df, list(mapeamento.keys()))
    novos_nomes = list(mapeamento.values())
    if len(set(novos_nomes)) != len(novos_nomes):
        raise ErroOperacao("Os novos nomes de coluna não podem se repetir.")
    conflitos = [n for n in novos_nomes if n in df.columns and n not in mapeamento]
    if conflitos:
        raise ErroOperacao(f"Já existe(m) coluna(s) com o(s) nome(s): {', '.join(conflitos)}")
    return df.rename(columns=mapeamento)


def remover_colunas(df: pd.DataFrame, colunas: list[str]) -> pd.DataFrame:
    """Remove as colunas informadas do DataFrame."""
    validar_colunas(df, colunas)
    if len(colunas) == df.shape[1]:
        raise ErroOperacao("Não é possível remover todas as colunas do dataset.")
    return df.drop(columns=colunas)


# ---------------------------------------------------------------------------
# Coluna calculada
# ---------------------------------------------------------------------------


def criar_coluna_operacao_aritmetica(
    df: pd.DataFrame, nova_coluna: str, colunas: list[str], operacao: str
) -> pd.DataFrame:
    """Cria uma coluna combinando outras colunas numéricas: soma, subtração, multiplicação ou divisão."""
    for coluna in colunas:
        validar_coluna_numerica(df, coluna)
    if len(colunas) < 2:
        raise ErroOperacao("Informe ao menos duas colunas para a operação aritmética.")
    resultado = df.copy()
    acumulado = df[colunas[0]].astype(float)
    for coluna in colunas[1:]:
        proxima = df[coluna].astype(float)
        if operacao == "soma":
            acumulado = acumulado + proxima
        elif operacao == "subtracao":
            acumulado = acumulado - proxima
        elif operacao == "multiplicacao":
            acumulado = acumulado * proxima
        elif operacao == "divisao":
            acumulado = np.where(proxima == 0, np.nan, acumulado / proxima)
            acumulado = pd.Series(acumulado, index=df.index)
        else:
            raise ErroOperacao(f"Operação aritmética desconhecida: '{operacao}'")
    resultado[nova_coluna] = acumulado
    return resultado


def criar_coluna_percentual(
    df: pd.DataFrame, nova_coluna: str, coluna_parte: str, coluna_total: str
) -> pd.DataFrame:
    """Cria uma coluna de percentual = parte / total * 100, evitando divisão por zero."""
    validar_coluna_numerica(df, coluna_parte)
    validar_coluna_numerica(df, coluna_total)
    resultado = df.copy()
    total = df[coluna_total].astype(float)
    parte = df[coluna_parte].astype(float)
    resultado[nova_coluna] = np.where(total == 0, np.nan, parte / total * 100)
    return resultado


def criar_coluna_condicional(
    df: pd.DataFrame, nova_coluna: str, condicao: pd.Series, valor_verdadeiro: Any, valor_falso: Any
) -> pd.DataFrame:
    """Cria uma coluna a partir de uma condição booleana já calculada (ver módulo ``utils``)."""
    resultado = df.copy()
    resultado[nova_coluna] = np.where(condicao, valor_verdadeiro, valor_falso)
    return resultado


def criar_coluna_concatenacao(
    df: pd.DataFrame, nova_coluna: str, colunas: list[str], separador: str = ""
) -> pd.DataFrame:
    """Cria uma coluna concatenando o texto de outras colunas."""
    validar_colunas(df, colunas)
    resultado = df.copy()
    partes = [df[c].astype(str).fillna("") for c in colunas]
    concatenado = partes[0]
    for parte in partes[1:]:
        concatenado = concatenado + separador + parte
    resultado[nova_coluna] = concatenado
    return resultado


def criar_coluna_valor_fixo(df: pd.DataFrame, nova_coluna: str, valor: Any) -> pd.DataFrame:
    """Cria uma coluna preenchida com um valor fixo (constante) em todas as linhas."""
    resultado = df.copy()
    resultado[nova_coluna] = valor
    return resultado


def criar_coluna_expressao(df: pd.DataFrame, nova_coluna: str, valores: pd.Series) -> pd.DataFrame:
    """Cria uma coluna a partir de uma série já calculada por uma expressão validada."""
    resultado = df.copy()
    resultado[nova_coluna] = valores
    return resultado


def criar_coluna_diferenca_datas(
    df: pd.DataFrame, nova_coluna: str, coluna_inicio: str, coluna_fim: str, unidade: str = "dias"
) -> pd.DataFrame:
    """Cria uma coluna com a diferença entre duas colunas de data."""
    validar_colunas(df, [coluna_inicio, coluna_fim])
    resultado = df.copy()
    inicio = pd.to_datetime(df[coluna_inicio], errors="coerce")
    fim = pd.to_datetime(df[coluna_fim], errors="coerce")
    diferenca = fim - inicio
    if unidade == "dias":
        resultado[nova_coluna] = diferenca.dt.days
    elif unidade == "horas":
        resultado[nova_coluna] = diferenca / pd.Timedelta(hours=1)
    elif unidade == "meses":
        resultado[nova_coluna] = (fim.dt.year - inicio.dt.year) * 12 + (fim.dt.month - inicio.dt.month)
    elif unidade == "anos":
        resultado[nova_coluna] = fim.dt.year - inicio.dt.year
    else:
        raise ErroOperacao(f"Unidade desconhecida: '{unidade}'. Use dias, horas, meses ou anos.")
    return resultado


# ---------------------------------------------------------------------------
# Operações matemáticas (sobre uma coluna existente)
# ---------------------------------------------------------------------------


def aplicar_operacao_matematica(
    df: pd.DataFrame, coluna: str, operacao: str, nova_coluna: str, parametro: Optional[float] = None
) -> pd.DataFrame:
    """Aplica uma operação matemática a uma coluna numérica, gerando uma nova coluna."""
    validar_coluna_numerica(df, coluna)
    resultado = df.copy()
    serie = df[coluna].astype(float)

    if operacao == "somar_constante":
        resultado[nova_coluna] = serie + (parametro or 0)
    elif operacao == "subtrair_constante":
        resultado[nova_coluna] = serie - (parametro or 0)
    elif operacao == "multiplicar_constante":
        resultado[nova_coluna] = serie * (parametro if parametro is not None else 1)
    elif operacao == "dividir_constante":
        if not parametro:
            raise ErroOperacao("Não é possível dividir por zero.")
        resultado[nova_coluna] = serie / parametro
    elif operacao == "potencia":
        resultado[nova_coluna] = serie ** (parametro if parametro is not None else 2)
    elif operacao == "raiz_quadrada":
        if (serie < 0).any():
            raise ErroOperacao("A coluna contém valores negativos; raiz quadrada indefinida para eles.")
        resultado[nova_coluna] = np.sqrt(serie)
    elif operacao == "logaritmo":
        if (serie <= 0).any():
            raise ErroOperacao("A coluna contém valores menores ou iguais a zero; logaritmo indefinido para eles.")
        resultado[nova_coluna] = np.log(serie)
    elif operacao == "arredondar":
        casas = int(parametro) if parametro is not None else 0
        resultado[nova_coluna] = serie.round(casas)
    elif operacao == "absoluto":
        resultado[nova_coluna] = serie.abs()
    else:
        raise ErroOperacao(f"Operação matemática desconhecida: '{operacao}'")
    return resultado


# ---------------------------------------------------------------------------
# Texto
# ---------------------------------------------------------------------------


def texto_concatenar(df: pd.DataFrame, colunas: list[str], separador: str = "") -> pd.Series:
    """Concatena o texto de várias colunas."""
    validar_colunas(df, colunas)
    partes = [df[c].astype(str).fillna("") for c in colunas]
    resultado = partes[0]
    for parte in partes[1:]:
        resultado = resultado + separador + parte
    return resultado


def texto_esquerda(serie: pd.Series, quantidade: int) -> pd.Series:
    """Extrai os N caracteres mais à esquerda de cada valor (equivalente a ESQUERDA)."""
    return serie.astype(str).str.slice(0, quantidade)


def texto_direita(serie: pd.Series, quantidade: int) -> pd.Series:
    """Extrai os N caracteres mais à direita de cada valor (equivalente a DIREITA)."""
    return serie.astype(str).str.slice(-quantidade if quantidade > 0 else None)


def texto_localizar(serie: pd.Series, subtexto: str) -> pd.Series:
    """Retorna a posição (1-based) do subtexto, ou -1 se não encontrado (equivalente a LOCALIZAR)."""
    posicoes = serie.astype(str).str.find(subtexto)
    return posicoes.where(posicoes < 0, posicoes + 1)


def texto_substituir(serie: pd.Series, antigo: str, novo: str) -> pd.Series:
    """Substitui todas as ocorrências de um texto por outro (equivalente a SUBSTITUIR)."""
    return serie.astype(str).str.replace(antigo, novo, regex=False)


def texto_remover_espacos(serie: pd.Series) -> pd.Series:
    """Remove espaços nas extremidades e reduz espaços internos duplicados (equivalente a ARRUMAR)."""
    return serie.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)


def texto_maiusculas(serie: pd.Series) -> pd.Series:
    """Converte o texto para maiúsculas (equivalente a MAIÚSCULA)."""
    return serie.astype(str).str.upper()


def texto_minusculas(serie: pd.Series) -> pd.Series:
    """Converte o texto para minúsculas (equivalente a MINÚSCULA)."""
    return serie.astype(str).str.lower()


def texto_iniciais_maiusculas(serie: pd.Series) -> pd.Series:
    """Coloca a primeira letra de cada palavra em maiúscula (equivalente a PRI.MAIÚSCULA)."""
    return serie.astype(str).str.title()


def texto_contar_caracteres(serie: pd.Series) -> pd.Series:
    """Conta a quantidade de caracteres de cada valor (equivalente a NÚM.CARACT)."""
    return serie.astype(str).str.len()


def texto_separar(serie: pd.Series, delimitador: str, maximo_partes: Optional[int] = None) -> pd.DataFrame:
    """Separa o texto em várias colunas por um delimitador (equivalente a Texto para Colunas)."""
    n_max = (maximo_partes - 1) if maximo_partes else -1
    expandido = serie.astype(str).str.split(delimitador, n=n_max, expand=True)
    expandido.columns = [f"parte_{i + 1}" for i in range(expandido.shape[1])]
    return expandido


# ---------------------------------------------------------------------------
# Datas
# ---------------------------------------------------------------------------


def data_converter(serie: pd.Series, formato: Optional[str] = None) -> pd.Series:
    """Converte uma coluna para o tipo data de forma segura, com valores inválidos viram NaT."""
    if formato:
        return pd.to_datetime(serie, format=formato, errors="coerce")
    return pd.to_datetime(serie, errors="coerce", format="mixed")


_COMPONENTES_DATA: dict[str, Callable[[pd.Series], pd.Series]] = {
    "dia": lambda s: s.dt.day,
    "mes": lambda s: s.dt.month,
    "ano": lambda s: s.dt.year,
    "trimestre": lambda s: s.dt.quarter,
    "dia_semana": lambda s: s.dt.day_name(locale=None),
}


def data_extrair_componente(serie_data: pd.Series, componente: str) -> pd.Series:
    """Extrai dia, mês, ano, trimestre ou dia da semana de uma coluna de data."""
    if componente not in _COMPONENTES_DATA:
        raise ErroOperacao(f"Componente de data desconhecido: '{componente}'")
    return _COMPONENTES_DATA[componente](serie_data)


def data_diferenca(serie_inicio: pd.Series, serie_fim: pd.Series, unidade: str = "dias") -> pd.Series:
    """Calcula a diferença entre duas colunas de data na unidade informada."""
    diferenca = serie_fim - serie_inicio
    if unidade == "dias":
        return diferenca.dt.days
    if unidade == "horas":
        return diferenca / pd.Timedelta(hours=1)
    if unidade == "meses":
        return (serie_fim.dt.year - serie_inicio.dt.year) * 12 + (serie_fim.dt.month - serie_inicio.dt.month)
    if unidade == "anos":
        return serie_fim.dt.year - serie_inicio.dt.year
    raise ErroOperacao(f"Unidade desconhecida: '{unidade}'.")


def data_somar_dias(serie_data: pd.Series, dias: int) -> pd.Series:
    """Soma (ou subtrai, se negativo) uma quantidade de dias a uma coluna de data."""
    return serie_data + pd.Timedelta(days=dias)


def data_identificar_invalidas(serie_data: pd.Series) -> pd.Series:
    """Retorna uma máscara booleana indicando quais datas são inválidas (NaT)."""
    return serie_data.isna()


def data_criar_coluna_periodo(df: pd.DataFrame, coluna_data: str, periodo: str, nova_coluna: str) -> pd.DataFrame:
    """Cria uma coluna de período (mês, trimestre ou ano) a partir de uma coluna de data, para agrupamento."""
    validar_colunas(df, [coluna_data])
    resultado = df.copy()
    serie_data = data_converter(df[coluna_data])
    if periodo == "mes":
        resultado[nova_coluna] = serie_data.dt.to_period("M").astype(str)
    elif periodo == "trimestre":
        resultado[nova_coluna] = serie_data.dt.to_period("Q").astype(str)
    elif periodo == "ano":
        resultado[nova_coluna] = serie_data.dt.to_period("Y").astype(str)
    else:
        raise ErroOperacao(f"Período desconhecido: '{periodo}'. Use mes, trimestre ou ano.")
    return resultado


# ---------------------------------------------------------------------------
# Agrupar e resumir / Tabela dinâmica
# ---------------------------------------------------------------------------


def agrupar_resumir(df: pd.DataFrame, colunas_grupo: list[str], agregacoes: dict[str, str]) -> pd.DataFrame:
    """Agrupa o DataFrame pelas colunas informadas e aplica agregações por coluna."""
    validar_colunas(df, colunas_grupo)
    validar_colunas(df, list(agregacoes.keys()))
    mapa_funcoes = {}
    for coluna, nome_agregacao in agregacoes.items():
        if nome_agregacao not in AGREGACOES:
            raise ErroOperacao(f"Agregação desconhecida: '{nome_agregacao}'")
        mapa_funcoes[coluna] = AGREGACOES[nome_agregacao]
    resultado = df.groupby(colunas_grupo, dropna=False).agg(mapa_funcoes).reset_index()
    return resultado


def criar_tabela_dinamica(
    df: pd.DataFrame,
    linhas: list[str],
    colunas: Optional[list[str]],
    valores: list[str],
    funcao_agregacao: str,
    preencher_vazios: Optional[float] = None,
    totais_gerais: bool = False,
) -> pd.DataFrame:
    """Cria uma tabela dinâmica usando ``pandas.pivot_table``."""
    validar_colunas(df, linhas)
    if colunas:
        validar_colunas(df, colunas)
    validar_colunas(df, valores)
    if funcao_agregacao not in AGREGACOES:
        raise ErroOperacao(f"Agregação desconhecida: '{funcao_agregacao}'")
    tabela = pd.pivot_table(
        df,
        index=linhas,
        columns=colunas or None,
        values=valores,
        aggfunc=AGREGACOES[funcao_agregacao],
        fill_value=preencher_vazios,
        margins=totais_gerais,
        margins_name="Total Geral",
    )
    if isinstance(tabela.columns, pd.MultiIndex):
        tabela.columns = [" | ".join(str(nivel) for nivel in tupla if str(nivel) != "") for tupla in tabela.columns]
    return tabela.reset_index()


# ---------------------------------------------------------------------------
# PROCV / PROCH / PROCX
# ---------------------------------------------------------------------------


@dataclass
class ResultadoBusca:
    """Encapsula o resultado de uma operação de busca (PROCV/PROCX) com avisos coletados."""

    df: pd.DataFrame
    avisos: list[str] = field(default_factory=list)


def executar_procv(
    df_principal: pd.DataFrame,
    df_consulta: pd.DataFrame,
    chave_principal: str,
    chave_consulta: str,
    colunas_retorno: list[str],
    sufixo: str = "_procv",
) -> ResultadoBusca:
    """Reproduz o comportamento do PROCV (correspondência exata) via ``pandas.merge``."""
    validar_colunas(df_principal, [chave_principal])
    validar_colunas(df_consulta, [chave_consulta] + colunas_retorno)

    avisos: list[str] = []

    duplicadas = df_consulta[chave_consulta].duplicated().sum()
    if duplicadas > 0:
        avisos.append(
            f"A tabela de consulta possui {duplicadas} chave(s) duplicada(s) em '{chave_consulta}'; "
            "a primeira correspondência será usada para cada linha da tabela principal."
        )

    tipo_principal = df_principal[chave_principal].dtype
    tipo_consulta = df_consulta[chave_consulta].dtype
    if tipo_principal != tipo_consulta:
        avisos.append(
            f"Tipos de chave incompatíveis ('{chave_principal}': {tipo_principal} x "
            f"'{chave_consulta}': {tipo_consulta}). As chaves serão comparadas como texto."
        )

    df_esquerda = df_principal.copy()
    df_direita = df_consulta[[chave_consulta] + colunas_retorno].drop_duplicates(subset=[chave_consulta], keep="first").copy()

    coluna_auxiliar_esq = None
    coluna_auxiliar_dir = None
    if tipo_principal != tipo_consulta:
        coluna_auxiliar_esq = "__chave_comparacao__"
        coluna_auxiliar_dir = "__chave_comparacao__"
        df_esquerda[coluna_auxiliar_esq] = df_esquerda[chave_principal].astype(str)
        df_direita[coluna_auxiliar_dir] = df_direita[chave_consulta].astype(str)
        chave_merge_esquerda, chave_merge_direita = coluna_auxiliar_esq, coluna_auxiliar_dir
    else:
        chave_merge_esquerda, chave_merge_direita = chave_principal, chave_consulta

    linhas_antes = len(df_esquerda)
    resultado = df_esquerda.merge(
        df_direita,
        left_on=chave_merge_esquerda,
        right_on=chave_merge_direita,
        how="left",
        suffixes=("", sufixo),
    )

    if coluna_auxiliar_esq:
        resultado = resultado.drop(columns=[coluna_auxiliar_esq])
    if chave_consulta != chave_principal and chave_consulta in resultado.columns and chave_consulta not in df_principal.columns:
        resultado = resultado.drop(columns=[chave_consulta])

    if len(resultado) != linhas_antes:
        avisos.append(
            f"A junção alterou a quantidade de linhas: {linhas_antes} antes, {len(resultado)} depois."
        )

    for coluna in colunas_retorno:
        nome_final = coluna if coluna not in df_principal.columns else f"{coluna}{sufixo}"
        nome_final = nome_final if nome_final in resultado.columns else coluna
        if nome_final in resultado.columns:
            sem_correspondencia = resultado[nome_final].isna().sum()
            if sem_correspondencia > 0:
                avisos.append(f"{sem_correspondencia} linha(s) sem correspondência para a coluna '{coluna}'.")

    return ResultadoBusca(df=resultado, avisos=avisos)


def executar_proch(df: pd.DataFrame, linha_chave: int, valor_procurado: Any, linha_retorno: int) -> ResultadoBusca:
    """Reproduz o comportamento do PROCH (busca horizontal).

    Limitação: DataFrames do pandas são otimizados para dados organizados por
    colunas, não por linhas. Para simular o PROCH, tratamos duas linhas do
    DataFrame como se fossem um intervalo horizontal do Excel: a ``linha_chave``
    contém os valores de busca (equivalentes ao cabeçalho de uma tabela
    horizontal) e a ``linha_retorno`` contém os valores a devolver. A busca é
    feita percorrendo as colunas nessa linha específica — uma operação O(n)
    em vez de um acesso indexado, já que não há índice construído sobre linhas.
    """
    if linha_chave < 0 or linha_chave >= len(df):
        raise ErroOperacao(f"Linha de busca {linha_chave} fora do intervalo do dataset (0 a {len(df) - 1}).")
    if linha_retorno < 0 or linha_retorno >= len(df):
        raise ErroOperacao(f"Linha de retorno {linha_retorno} fora do intervalo do dataset (0 a {len(df) - 1}).")

    linha_busca = df.iloc[linha_chave]
    coluna_encontrada = None
    for coluna in df.columns:
        if str(linha_busca[coluna]) == str(valor_procurado):
            coluna_encontrada = coluna
            break

    avisos: list[str] = []
    if coluna_encontrada is None:
        avisos.append(f"Valor '{valor_procurado}' não encontrado na linha {linha_chave}.")
        valor_retornado = None
    else:
        valor_retornado = df.iloc[linha_retorno][coluna_encontrada]

    resultado = pd.DataFrame(
        [{"valor_procurado": valor_procurado, "coluna_encontrada": coluna_encontrada, "valor_retornado": valor_retornado}]
    )
    return ResultadoBusca(df=resultado, avisos=avisos)


def executar_procx(
    df_principal: pd.DataFrame,
    df_consulta: pd.DataFrame,
    coluna_busca_principal: str,
    coluna_busca_consulta: str,
    colunas_retorno: list[str],
    modo_correspondencia: str = "exata",
    ocorrencia: str = "primeira",
    valor_nao_encontrado: Any = None,
) -> ResultadoBusca:
    """Reproduz o comportamento do PROCX (XLOOKUP), com correspondência exata ou aproximada.

    Não há distinção estrutural entre "busca à esquerda" e "busca à direita" em
    um DataFrame do pandas, pois as colunas são referenciadas pelo nome e não
    pela posição relativa a um intervalo — diferente do Excel, onde o PROCX
    pode retornar colunas à esquerda da coluna de busca. Essa flexibilidade já
    é nativa aqui: basta informar qualquer coluna em ``colunas_retorno``,
    independentemente de sua posição.
    """
    validar_colunas(df_principal, [coluna_busca_principal])
    validar_colunas(df_consulta, [coluna_busca_consulta] + colunas_retorno)
    avisos: list[str] = []

    if modo_correspondencia == "exata":
        manter = "first" if ocorrencia == "primeira" else "last"
        tabela = df_consulta.drop_duplicates(subset=[coluna_busca_consulta], keep=manter)
        duplicadas = df_consulta[coluna_busca_consulta].duplicated().sum()
        if duplicadas > 0:
            avisos.append(
                f"Existem {duplicadas} chave(s) duplicada(s); mantida a ocorrência '{ocorrencia}' de cada uma."
            )
        mapa = tabela.set_index(coluna_busca_consulta)
        resultado = df_principal.copy()
        for coluna in colunas_retorno:
            valores_mapeados = df_principal[coluna_busca_principal].map(mapa[coluna])
            if valor_nao_encontrado is not None:
                valores_mapeados = valores_mapeados.fillna(valor_nao_encontrado)
            nome_novo = coluna if coluna not in resultado.columns else f"{coluna}_procx"
            resultado[nome_novo] = valores_mapeados
        nao_encontrados = df_principal[~df_principal[coluna_busca_principal].isin(mapa.index)].shape[0]
        if nao_encontrados > 0:
            avisos.append(f"{nao_encontrados} linha(s) sem correspondência exata.")
        return ResultadoBusca(df=resultado, avisos=avisos)

    if modo_correspondencia == "aproximada":
        try:
            esquerda = df_principal.copy()
            direita = df_consulta[[coluna_busca_consulta] + colunas_retorno].copy()
            esquerda["_chave_ordenacao_"] = pd.to_numeric(esquerda[coluna_busca_principal], errors="coerce")
            direita["_chave_ordenacao_"] = pd.to_numeric(direita[coluna_busca_consulta], errors="coerce")
            esquerda = esquerda.sort_values("_chave_ordenacao_")
            direita = direita.sort_values("_chave_ordenacao_")
            resultado = pd.merge_asof(esquerda, direita, on="_chave_ordenacao_", direction="nearest")
            resultado = resultado.drop(columns=["_chave_ordenacao_"])
            return ResultadoBusca(df=resultado, avisos=avisos)
        except Exception as exc:  # noqa: BLE001
            raise ErroOperacao(
                f"Não foi possível realizar a correspondência aproximada: {exc}. "
                "Verifique se as colunas de busca são numéricas ou datas."
            ) from exc

    raise ErroOperacao(f"Modo de correspondência desconhecido: '{modo_correspondencia}'.")


# ---------------------------------------------------------------------------
# Funções condicionais (SOMASE, CONT.SE, MÉDIASE e variações)
# ---------------------------------------------------------------------------


def _interpretar_criterio(criterio: str) -> tuple[str, str]:
    """Separa um critério estilo Excel (ex.: '>10', '<>0', 'abc*') em (operador, valor)."""
    criterio = str(criterio).strip()
    for operador in (">=", "<=", "<>", ">", "<", "="):
        if criterio.startswith(operador):
            return operador, criterio[len(operador):].strip()
    return "=", criterio


def construir_mascara_criterio(serie: pd.Series, criterio: str) -> pd.Series:
    """Constrói a máscara booleana equivalente a um critério estilo Excel, com suporte a curingas."""
    operador, valor = _interpretar_criterio(criterio)

    if valor.lower() in ("", "vazio"):
        return serie.isna() if operador in ("=", "") else serie.notna()

    if operador == "=" and ("*" in valor or "?" in valor):
        padrao = curinga_para_regex(valor)
        return serie.astype(str).str.match(padrao, case=False, na=False)

    if operador in (">", "<", ">=", "<="):
        serie_numerica = pd.to_numeric(serie, errors="coerce")
        try:
            valor_numerico = float(valor)
        except ValueError:
            valor_numerico = pd.to_datetime(valor, errors="coerce")
            serie_numerica = pd.to_datetime(serie, errors="coerce")
        if operador == ">":
            return serie_numerica > valor_numerico
        if operador == "<":
            return serie_numerica < valor_numerico
        if operador == ">=":
            return serie_numerica >= valor_numerico
        return serie_numerica <= valor_numerico

    try:
        valor_numerico = float(valor)
        serie_comparacao = pd.to_numeric(serie, errors="coerce")
        if operador == "<>":
            return serie_comparacao != valor_numerico
        return serie_comparacao == valor_numerico
    except ValueError:
        serie_texto = serie.astype(str)
        if operador == "<>":
            return serie_texto.str.lower() != valor.lower()
        return serie_texto.str.lower() == valor.lower()


def _mascara_multiplos_criterios(df: pd.DataFrame, pares_criterio: list[tuple[str, str]]) -> pd.Series:
    """Combina vários pares (coluna, critério) com E lógico, como fazem as funções *SES do Excel."""
    if not pares_criterio:
        raise ErroOperacao("Informe ao menos um par de coluna/critério.")
    mascara = pd.Series(True, index=df.index)
    for coluna, criterio in pares_criterio:
        validar_colunas(df, [coluna])
        mascara &= construir_mascara_criterio(df[coluna], criterio)
    return mascara


def somase(df: pd.DataFrame, coluna_soma: str, coluna_criterio: str, criterio: str) -> float:
    """Equivalente a SOMASE: soma os valores de uma coluna onde o critério é atendido."""
    validar_coluna_numerica(df, coluna_soma)
    validar_colunas(df, [coluna_criterio])
    mascara = construir_mascara_criterio(df[coluna_criterio], criterio)
    return float(df.loc[mascara, coluna_soma].sum())


def somases(df: pd.DataFrame, coluna_soma: str, pares_criterio: list[tuple[str, str]]) -> float:
    """Equivalente a SOMASES: soma os valores de uma coluna onde todos os critérios são atendidos."""
    validar_coluna_numerica(df, coluna_soma)
    mascara = _mascara_multiplos_criterios(df, pares_criterio)
    return float(df.loc[mascara, coluna_soma].sum())


def cont_se(df: pd.DataFrame, coluna_criterio: str, criterio: str) -> int:
    """Equivalente a CONT.SE: conta as linhas onde o critério é atendido."""
    validar_colunas(df, [coluna_criterio])
    mascara = construir_mascara_criterio(df[coluna_criterio], criterio)
    return int(mascara.sum())


def cont_ses(df: pd.DataFrame, pares_criterio: list[tuple[str, str]]) -> int:
    """Equivalente a CONT.SES: conta as linhas onde todos os critérios são atendidos."""
    mascara = _mascara_multiplos_criterios(df, pares_criterio)
    return int(mascara.sum())


def mediase(df: pd.DataFrame, coluna_media: str, coluna_criterio: str, criterio: str) -> float:
    """Equivalente a MÉDIASE: calcula a média de uma coluna onde o critério é atendido."""
    validar_coluna_numerica(df, coluna_media)
    validar_colunas(df, [coluna_criterio])
    mascara = construir_mascara_criterio(df[coluna_criterio], criterio)
    valores = df.loc[mascara, coluna_media]
    if len(valores) == 0:
        raise ErroOperacao("Nenhuma linha atendeu ao critério informado; média indefinida.")
    return float(valores.mean())


def mediases(df: pd.DataFrame, coluna_media: str, pares_criterio: list[tuple[str, str]]) -> float:
    """Equivalente a MÉDIASES: calcula a média de uma coluna onde todos os critérios são atendidos."""
    validar_coluna_numerica(df, coluna_media)
    mascara = _mascara_multiplos_criterios(df, pares_criterio)
    valores = df.loc[mascara, coluna_media]
    if len(valores) == 0:
        raise ErroOperacao("Nenhuma linha atendeu aos critérios informados; média indefinida.")
    return float(valores.mean())


# ---------------------------------------------------------------------------
# SE / SE aninhado
# ---------------------------------------------------------------------------


def funcao_se(
    df: pd.DataFrame, nova_coluna: str, condicao: pd.Series, valor_verdadeiro: Any, valor_falso: Any
) -> pd.DataFrame:
    """Equivalente a SE: cria uma coluna com base em uma condição booleana já calculada."""
    resultado = df.copy()
    resultado[nova_coluna] = np.where(condicao, valor_verdadeiro, valor_falso)
    return resultado


def funcao_se_aninhado(
    df: pd.DataFrame, nova_coluna: str, condicoes: list[pd.Series], valores: list[Any], valor_padrao: Any
) -> pd.DataFrame:
    """Equivalente a SE aninhado: aplica a primeira condição verdadeira, na ordem informada."""
    if len(condicoes) != len(valores):
        raise ErroOperacao("A quantidade de condições e de valores deve ser igual.")
    resultado = df.copy()
    resultado[nova_coluna] = np.select(condicoes, valores, default=valor_padrao)
    return resultado


# ---------------------------------------------------------------------------
# Nulos e duplicados
# ---------------------------------------------------------------------------


def remover_linhas_nulas(df: pd.DataFrame, colunas: Optional[list[str]] = None) -> pd.DataFrame:
    """Remove linhas com valores nulos em qualquer coluna, ou apenas nas colunas informadas."""
    if colunas:
        validar_colunas(df, colunas)
        return df.dropna(subset=colunas).reset_index(drop=True)
    return df.dropna().reset_index(drop=True)


def preencher_nulos(
    df: pd.DataFrame, colunas: list[str], metodo: str, valor: Optional[Any] = None
) -> pd.DataFrame:
    """Preenche valores nulos com um valor informado, média, mediana, moda, zero ou texto."""
    validar_colunas(df, colunas)
    resultado = df.copy()
    for coluna in colunas:
        serie = resultado[coluna]
        if metodo == "valor":
            resultado[coluna] = serie.fillna(valor)
        elif metodo == "media":
            if not pd.api.types.is_numeric_dtype(serie):
                raise ErroOperacao(f"A coluna '{coluna}' não é numérica; não é possível preencher com a média.")
            resultado[coluna] = serie.fillna(serie.mean())
        elif metodo == "mediana":
            if not pd.api.types.is_numeric_dtype(serie):
                raise ErroOperacao(f"A coluna '{coluna}' não é numérica; não é possível preencher com a mediana.")
            resultado[coluna] = serie.fillna(serie.median())
        elif metodo == "moda":
            moda = serie.mode(dropna=True)
            if moda.empty:
                raise ErroOperacao(f"Não foi possível calcular a moda da coluna '{coluna}' (sem valores válidos).")
            resultado[coluna] = serie.fillna(moda.iloc[0])
        elif metodo == "zero":
            resultado[coluna] = serie.fillna(0)
        elif metodo == "texto":
            resultado[coluna] = serie.fillna(str(valor) if valor is not None else "")
        else:
            raise ErroOperacao(f"Método de preenchimento desconhecido: '{metodo}'")
    return resultado


def remover_duplicados(df: pd.DataFrame, colunas: Optional[list[str]], manter: str = "primeira") -> pd.DataFrame:
    """Remove linhas duplicadas por todas as colunas ou por um subconjunto delas."""
    if colunas:
        validar_colunas(df, colunas)
    manter_pandas = "first" if manter == "primeira" else "last"
    return df.drop_duplicates(subset=colunas, keep=manter_pandas).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Concatenar / mesclar datasets
# ---------------------------------------------------------------------------


def concatenar_datasets(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Concatena verticalmente vários DataFrames, alinhando colunas por nome."""
    if len(dfs) < 2:
        raise ErroOperacao("Informe ao menos dois datasets para concatenar.")
    return pd.concat(dfs, ignore_index=True, sort=False)


def mesclar_datasets(
    df_esquerda: pd.DataFrame,
    df_direita: pd.DataFrame,
    chave_esquerda: str,
    chave_direita: str,
    tipo_juncao: str = "inner",
) -> pd.DataFrame:
    """Mescla dois datasets pela chave informada (inner, left, right ou outer)."""
    validar_colunas(df_esquerda, [chave_esquerda])
    validar_colunas(df_direita, [chave_direita])
    if tipo_juncao not in ("inner", "left", "right", "outer"):
        raise ErroOperacao(f"Tipo de junção desconhecido: '{tipo_juncao}'")
    return df_esquerda.merge(
        df_direita, left_on=chave_esquerda, right_on=chave_direita, how=tipo_juncao, suffixes=("", "_direita")
    )


# ---------------------------------------------------------------------------
# Qualidade dos dados
# ---------------------------------------------------------------------------


def validar_qualidade_dados(df: pd.DataFrame) -> dict[str, Any]:
    """Executa validações gerais de qualidade e retorna um resumo estruturado."""
    total_linhas = len(df)
    total_nulos = int(df.isna().sum().sum())
    duplicadas = int(df.duplicated().sum())
    colunas_totalmente_nulas = [str(c) for c in df.columns if df[c].isna().all()]
    colunas_constantes = [str(c) for c in df.columns if df[c].nunique(dropna=True) == 1]
    return {
        "total_linhas": total_linhas,
        "total_colunas": df.shape[1],
        "total_nulos": total_nulos,
        "percentual_nulos": round((total_nulos / (total_linhas * df.shape[1]) * 100) if total_linhas and df.shape[1] else 0.0, 2),
        "linhas_duplicadas": duplicadas,
        "colunas_totalmente_nulas": colunas_totalmente_nulas,
        "colunas_constantes": colunas_constantes,
    }
