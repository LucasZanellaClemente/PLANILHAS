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


_PADRAO_NUMERO_MILHAR_BR = re.compile(r"^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$")
_PADRAO_NUMERO_DECIMAL_BR = re.compile(r"^[+-]?\d*,\d+$")


def converter_numero_br(valor: Any) -> float:
    """Converte um número digitado no formato brasileiro ou internacional para ``float``.

    Aceita ``10,5``, ``1.234,56``, ``1.000`` (milhar), ``10.5`` e ``1234``.
    Levanta ``ValueError`` quando o valor não é um número.
    """
    if isinstance(valor, (int, float, np.integer, np.floating)) and not isinstance(valor, bool):
        return float(valor)
    texto = str(valor).strip().replace(" ", "")
    if _PADRAO_NUMERO_MILHAR_BR.match(texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif _PADRAO_NUMERO_DECIMAL_BR.match(texto):
        texto = texto.replace(",", ".")
    return float(texto)


def serie_numeros_br(serie: pd.Series) -> pd.Series | None:
    """Converte uma coluna de texto com números no formato brasileiro para numérica.

    Só converte quando todos os valores preenchidos são números em formato
    brasileiro (``1.234,56``, ``10,50``); caso contrário retorna ``None``.
    """
    textos = serie.dropna().astype(str).str.strip()
    textos = textos[textos != ""]
    if textos.empty:
        return None
    eh_br = textos.str.match(_PADRAO_NUMERO_MILHAR_BR) | textos.str.match(_PADRAO_NUMERO_DECIMAL_BR)
    eh_inteiro = textos.str.match(r"^[+-]?\d+$")
    if not (eh_br | eh_inteiro).all() or not eh_br.any():
        return None
    convertida = serie.map(lambda v: converter_numero_br(v) if pd.notna(v) and str(v).strip() else np.nan)
    return convertida.astype(float)


_PADRAO_DATA_ANO_PRIMEIRO = r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T].*)?$"


def converter_datas_br(valores: Any) -> Any:
    """Converte datas priorizando o formato brasileiro (dia/mês/ano).

    Datas que começam pelo ano (``2026-02-01``) são sempre lidas como
    ano-mês-dia: no pandas 3 o ``dayfirst=True`` também inverteria essas.
    Valores inválidos viram ``NaT``.
    """
    if not isinstance(valores, (pd.Series, pd.Index, list, tuple, np.ndarray)):
        return converter_datas_br(pd.Series([valores])).iloc[0]
    serie = valores if isinstance(valores, pd.Series) else pd.Series(list(valores))
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    textos = serie.where(serie.notna(), "").astype(str).str.strip()
    ano_primeiro = textos.str.match(_PADRAO_DATA_ANO_PRIMEIRO)
    iso = pd.to_datetime(textos.where(ano_primeiro).str.replace("/", "-"), errors="coerce", format="ISO8601")
    outros = pd.to_datetime(serie.where(~ano_primeiro), errors="coerce", format="mixed", dayfirst=True)
    return iso.where(ano_primeiro, outros)


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
    ast.Invert,
    ast.And,
    ast.Or,
    ast.BitAnd,
    ast.BitOr,
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
        if isinstance(no, ast.Compare):
            for operando in [no.left, *no.comparators]:
                if isinstance(operando, ast.BinOp) and isinstance(operando.op, (ast.BitAnd, ast.BitOr)):
                    raise ErroExpressaoInsegura(
                        "Ao usar & ou |, coloque cada condição entre parênteses, ex.: "
                        "(Quantidade > 2) & (Valor_Total > 500). Ou use and/or: Quantidade > 2 and Valor_Total > 500."
                    )


def _como_booleano(valor: Any) -> Any:
    """Converte um operando lógico em booleano elemento a elemento (vazio conta como falso)."""
    if isinstance(valor, pd.Series):
        return valor.notna() & valor.astype(bool)
    return bool(valor) if not pd.isna(valor) else False


def _e_logico(*valores: Any) -> Any:
    resultado = _como_booleano(valores[0])
    for valor in valores[1:]:
        resultado = resultado & _como_booleano(valor)
    return resultado


def _ou_logico(*valores: Any) -> Any:
    resultado = _como_booleano(valores[0])
    for valor in valores[1:]:
        resultado = resultado | _como_booleano(valor)
    return resultado


def _nao_logico(valor: Any) -> Any:
    booleano = _como_booleano(valor)
    return ~booleano if isinstance(booleano, pd.Series) else not booleano


_FUNCOES_LOGICAS: dict[str, Callable[..., Any]] = {
    "__e__": _e_logico,
    "__ou__": _ou_logico,
    "__nao__": _nao_logico,
}


class _TradutorLogico(ast.NodeTransformer):
    """Troca and/or/not (e &, |, ~) por funções que operam coluna a coluna.

    O ``and``/``or`` do Python tenta converter a coluna inteira em um único
    verdadeiro/falso, o que gera "The truth value of a Series is ambiguous".
    Também desdobra comparações encadeadas (``1 < x < 5``) pelo mesmo motivo.
    """

    @staticmethod
    def _chamar(nome: str, argumentos: list[ast.expr]) -> ast.Call:
        return ast.Call(func=ast.Name(id=nome, ctx=ast.Load()), args=argumentos, keywords=[])

    def visit_BoolOp(self, no: ast.BoolOp) -> ast.AST:
        self.generic_visit(no)
        return self._chamar("__e__" if isinstance(no.op, ast.And) else "__ou__", no.values)

    def visit_BinOp(self, no: ast.BinOp) -> ast.AST:
        self.generic_visit(no)
        if isinstance(no.op, ast.BitAnd):
            return self._chamar("__e__", [no.left, no.right])
        if isinstance(no.op, ast.BitOr):
            return self._chamar("__ou__", [no.left, no.right])
        return no

    def visit_UnaryOp(self, no: ast.UnaryOp) -> ast.AST:
        self.generic_visit(no)
        if isinstance(no.op, (ast.Not, ast.Invert)):
            return self._chamar("__nao__", [no.operand])
        return no

    def visit_Compare(self, no: ast.Compare) -> ast.AST:
        self.generic_visit(no)
        if len(no.ops) == 1:
            return no
        operandos = [no.left, *no.comparators]
        partes = [
            ast.Compare(left=operandos[i], ops=[op], comparators=[operandos[i + 1]])
            for i, op in enumerate(no.ops)
        ]
        return self._chamar("__e__", partes)


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
    arvore = ast.fix_missing_locations(_TradutorLogico().visit(arvore))
    codigo = compile(arvore, "<expressao_usuario>", mode="eval")
    contexto: dict[str, Any] = {str(coluna): df[coluna] for coluna in df.columns}
    contexto.update(FUNCOES_PERMITIDAS)
    contexto.update(_FUNCOES_LOGICAS)
    try:
        resultado = eval(codigo, {"__builtins__": {}}, contexto)  # nós já validamos a árvore acima
    except Exception as exc:  # noqa: BLE001
        raise ErroExpressaoInsegura(f"Não foi possível avaliar a expressão: {exc}") from exc
    if not isinstance(resultado, pd.Series):
        resultado = pd.Series([resultado] * len(df), index=df.index)
    return resultado


def avaliar_condicao_segura(expressao: str, df: pd.DataFrame) -> pd.Series:
    """Igual a :func:`avaliar_expressao_segura`, mas garante retorno booleano."""
    resultado = avaliar_expressao_segura(expressao, df)
    return _como_booleano(resultado)


def formatar_percentual(valor: float) -> str:
    """Formata um número float como percentual com duas casas decimais."""
    return f"{valor:.2f}%"


def truncar_texto(texto: str, tamanho: int = 60) -> str:
    """Trunca um texto longo adicionando reticências, para exibição em tabelas."""
    texto = str(texto)
    if len(texto) <= tamanho:
        return texto
    return texto[: tamanho - 3] + "..."
