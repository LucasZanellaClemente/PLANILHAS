"""Funções utilitárias compartilhadas por toda a aplicação.

Este módulo concentra rotinas pequenas e reutilizáveis que não pertencem a
nenhuma camada específica (carregamento, catálogo, operações ou exportação):
sanitização contra injeção de fórmulas, conversão de curingas estilo Excel
para expressões regulares, validação de nomes de aba e um avaliador seguro
de expressões matemáticas/lógicas informadas pelo usuário.
"""

from __future__ import annotations

import ast
import re
from typing import Any, Callable

import numpy as np
import pandas as pd

CARACTERES_INICIAIS_PERIGOSOS = ("=", "+", "-", "@")


def sanitizar_valor_formula(valor: Any) -> Any:
    """Neutraliza possível injeção de fórmula em um valor textual.

    O Excel interpreta células cujo conteúdo comece com ``=``, ``+``, ``-``
    ou ``@`` como fórmulas. Ao exportar dados de origem externa (CSV, por
    exemplo) para o relatório, prefixamos esses valores com um apóstrofo
    para que sejam tratados como texto literal.
    """
    if isinstance(valor, str) and valor and valor[0] in CARACTERES_INICIAIS_PERIGOSOS:
        return "'" + valor
    return valor


def eh_coluna_texto(serie: pd.Series) -> bool:
    """Indica se a coluna guarda texto.

    No pandas 2 o texto fica em colunas ``object``; no pandas 3 ele usa o tipo
    ``str`` (``pd.StringDtype``). As duas formas precisam ser reconhecidas.
    """
    return pd.api.types.is_object_dtype(serie) or isinstance(serie.dtype, pd.StringDtype)


def sanitizar_dataframe_formulas(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica :func:`sanitizar_valor_formula` a todas as colunas textuais de um DataFrame."""
    df_seguro = df.copy()
    for coluna in df_seguro.columns:
        if eh_coluna_texto(df_seguro[coluna]):
            df_seguro[coluna] = df_seguro[coluna].map(sanitizar_valor_formula)
    return df_seguro


def curinga_para_regex(padrao: str) -> str:
    """Converte um padrão com curingas estilo Excel (``*`` e ``?``) em regex ancorada."""
    escapado = re.escape(padrao)
    escapado = escapado.replace(r"\*", ".*").replace(r"\?", ".")
    return f"^{escapado}$"


def nome_aba_valido(nome: str, existentes: set[str]) -> str:
    """Normaliza um nome de aba para ser válido, único e com no máximo 31 caracteres."""
    caracteres_invalidos = set('[]:*?/\\')
    limpo = "".join(c for c in str(nome) if c not in caracteres_invalidos).strip()
    if not limpo:
        limpo = "Plan"
    limpo = limpo[:31]
    base = limpo
    contador = 1
    while limpo in existentes:
        sufixo = f"_{contador}"
        limpo = base[: 31 - len(sufixo)] + sufixo
        contador += 1
    existentes.add(limpo)
    return limpo


class ErroExpressaoInsegura(Exception):
    """Levantada quando uma expressão informada pelo usuário contém elementos não permitidos."""


class ErroOperacao(Exception):
    """Erro de negócio esperado (validação de dataset, coluna, tipos etc.)."""


_NOS_PERMITIDOS = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Compare,
    ast.BoolOp,
    ast.Name,
    ast.Load,
    ast.Constant,
    ast.Call,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.Mod,
    ast.Pow,
    ast.FloorDiv,
    ast.USub,
    ast.UAdd,
    ast.Not,
    ast.And,
    ast.Or,
    ast.Eq,
    ast.NotEq,
    ast.Lt,
    ast.LtE,
    ast.Gt,
    ast.GtE,
)

FUNCOES_PERMITIDAS: dict[str, Callable[..., Any]] = {
    "abs": np.abs,
    "round": np.round,
    "min": np.minimum,
    "max": np.maximum,
    "sqrt": np.sqrt,
    "log": np.log,
    "log10": np.log10,
    "exp": np.exp,
}


def _validar_arvore(arvore: ast.AST, colunas_validas: set[str]) -> None:
    """Percorre a árvore sintática validando que só existam nós, nomes e chamadas permitidos."""
    for no in ast.walk(arvore):
        if not isinstance(no, _NOS_PERMITIDOS):
            raise ErroExpressaoInsegura(f"Elemento não permitido na expressão: {type(no).__name__}")
        if isinstance(no, ast.Name) and no.id not in colunas_validas and no.id not in FUNCOES_PERMITIDAS:
            raise ErroExpressaoInsegura(f"Nome não reconhecido na expressão: '{no.id}'")
        if isinstance(no, ast.Call):
            if not isinstance(no.func, ast.Name) or no.func.id not in FUNCOES_PERMITIDAS:
                raise ErroExpressaoInsegura("Apenas as funções abs, round, min, max, sqrt, log, log10 e exp são permitidas.")


def avaliar_expressao_segura(expressao: str, df: pd.DataFrame) -> pd.Series:
    """Avalia uma expressão matemática/lógica informada pelo usuário de forma segura.

    Em vez de usar ``eval()`` diretamente sobre a entrada bruta, a expressão é
    primeiro convertida em uma árvore sintática (``ast``) e validada nó a nó:
    somente operadores aritméticos, comparações, operadores lógicos, nomes de
    colunas existentes no DataFrame e um conjunto restrito de funções
    matemáticas são aceitos. Qualquer outro elemento (atributos, subscritos,
    chamadas arbitrárias, importações etc.) é rejeitado antes da avaliação.
    """
    colunas_validas = set(str(c) for c in df.columns)
    try:
        arvore = ast.parse(expressao, mode="eval")
    except SyntaxError as exc:
        raise ErroExpressaoInsegura(f"Expressão inválida: {exc}") from exc
    _validar_arvore(arvore, colunas_validas)
    codigo = compile(arvore, "<expressao_usuario>", mode="eval")
    contexto: dict[str, Any] = {str(coluna): df[coluna] for coluna in df.columns}
    contexto.update(FUNCOES_PERMITIDAS)
    resultado = eval(codigo, {"__builtins__": {}}, contexto)  # nós já validamos a árvore acima
    if not isinstance(resultado, pd.Series):
        resultado = pd.Series([resultado] * len(df), index=df.index)
    return resultado


def avaliar_condicao_segura(expressao: str, df: pd.DataFrame) -> pd.Series:
    """Igual a :func:`avaliar_expressao_segura`, mas garante retorno booleano."""
    resultado = avaliar_expressao_segura(expressao, df)
    return resultado.astype(bool)


def formatar_percentual(valor: float) -> str:
    """Formata um número float como percentual com duas casas decimais."""
    return f"{valor:.2f}%"


def truncar_texto(texto: str, tamanho: int = 60) -> str:
    """Trunca um texto longo adicionando reticências, para exibição em tabelas."""
    texto = str(texto)
    if len(texto) <= tamanho:
        return texto
    return texto[: tamanho - 3] + "..."
