"""Importação e validação de arquivos de planilha (XLSX, XLS, XLSM, CSV).

As funções deste módulo são puras em relação à interface: recebem caminhos
e parâmetros, devolvem DataFrames ou metadados, e levantam
:class:`~excel_toolkit.utils.ErroOperacao` em caso de problema. A interação
com o usuário (perguntar quais abas importar, pedir codificação manual etc.)
fica a cargo do módulo ``interface``.
"""

from __future__ import annotations

import csv
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

from utils import ErroOperacao, serie_numeros_br

logger = logging.getLogger(__name__)

FORMATOS_SUPORTADOS = {".xlsx", ".xlsm", ".xls", ".csv"}
CODIFICACOES_CANDIDATAS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
DELIMITADORES_CANDIDATOS = [",", ";", "\t", "|"]


def identificar_formato(caminho: Path) -> str:
    """Identifica o formato do arquivo a partir da extensão.

    Retorna uma das strings: ``xlsx``, ``xlsm``, ``xls`` ou ``csv``.
    Levanta :class:`ErroOperacao` se a extensão não for suportada.
    """
    sufixo = caminho.suffix.lower()
    if sufixo not in FORMATOS_SUPORTADOS:
        raise ErroOperacao(
            f"Formato '{sufixo}' não suportado. Formatos aceitos: "
            f"{', '.join(sorted(FORMATOS_SUPORTADOS))}."
        )
    return sufixo.lstrip(".")


def validar_arquivo(caminho: Path) -> None:
    """Valida que o caminho existe, é um arquivo e tem extensão suportada."""
    if not caminho.exists():
        raise ErroOperacao(f"Arquivo não encontrado: {caminho}")
    if not caminho.is_file():
        raise ErroOperacao(f"O caminho informado não é um arquivo: {caminho}")
    identificar_formato(caminho)


def detectar_codificacao(caminho: Path) -> str:
    """Tenta detectar a codificação de um arquivo CSV testando codificações comuns."""
    for codificacao in CODIFICACOES_CANDIDATAS:
        try:
            with open(caminho, encoding=codificacao) as arquivo:
                arquivo.read(65536)
            return codificacao
        except (UnicodeDecodeError, UnicodeError):
            continue
    logger.warning("Não foi possível detectar a codificação de %s; usando latin-1.", caminho)
    return "latin-1"


def detectar_delimitador(caminho: Path, codificacao: str) -> str:
    """Tenta detectar o delimitador de um arquivo CSV usando ``csv.Sniffer``."""
    try:
        with open(caminho, encoding=codificacao, newline="") as arquivo:
            amostra = arquivo.read(16384)
        if not amostra.strip():
            raise ErroOperacao(f"Arquivo CSV vazio: {caminho}")
        dialeto = csv.Sniffer().sniff(amostra, delimiters="".join(DELIMITADORES_CANDIDATOS))
        return dialeto.delimiter
    except csv.Error:
        contagens = {delim: amostra.count(delim) for delim in DELIMITADORES_CANDIDATOS}
        melhor = max(contagens, key=contagens.get)
        if contagens[melhor] == 0:
            logger.warning("Não foi possível detectar o delimitador de %s; usando ','.", caminho)
            return ","
        return melhor


def listar_abas_excel(caminho: Path) -> list[str]:
    """Lista os nomes das abas de um arquivo Excel (XLSX, XLSM ou XLS)."""
    formato = identificar_formato(caminho)
    engine = "xlrd" if formato == "xls" else "openpyxl"
    try:
        planilha = pd.ExcelFile(caminho, engine=engine)
    except Exception as exc:  # noqa: BLE001 - relatamos qualquer falha de leitura ao usuário
        raise ErroOperacao(f"Falha ao abrir o arquivo Excel '{caminho.name}': {exc}") from exc
    return list(planilha.sheet_names)


def carregar_planilha_excel(caminho: Path, aba: str) -> pd.DataFrame:
    """Carrega uma única aba de um arquivo Excel como DataFrame."""
    formato = identificar_formato(caminho)
    engine = "xlrd" if formato == "xls" else "openpyxl"
    try:
        df = pd.read_excel(caminho, sheet_name=aba, engine=engine)
    except Exception as exc:  # noqa: BLE001
        raise ErroOperacao(f"Falha ao ler a aba '{aba}' do arquivo '{caminho.name}': {exc}") from exc
    return df


def carregar_csv(
    caminho: Path,
    delimitador: Optional[str] = None,
    codificacao: Optional[str] = None,
) -> tuple[pd.DataFrame, str, str]:
    """Carrega um arquivo CSV, detectando delimitador e codificação quando não informados.

    Retorna uma tupla ``(dataframe, delimitador_usado, codificacao_usada)``.
    """
    codificacao_usada = codificacao or detectar_codificacao(caminho)
    delimitador_usado = delimitador or detectar_delimitador(caminho, codificacao_usada)
    try:
        df = pd.read_csv(
            caminho,
            sep=delimitador_usado,
            encoding=codificacao_usada,
            engine="python",
        )
    except Exception as exc:  # noqa: BLE001
        raise ErroOperacao(
            f"Falha ao ler o CSV '{caminho.name}' com delimitador '{delimitador_usado}' "
            f"e codificação '{codificacao_usada}': {exc}"
        ) from exc
    if df.shape[1] == 1:
        logger.warning(
            "O CSV '%s' foi lido com apenas uma coluna; o delimitador '%s' pode estar incorreto.",
            caminho.name,
            delimitador_usado,
        )
    # Colunas com números no formato brasileiro (1.234,56 / 10,50) chegam como texto.
    for coluna in df.columns:
        if df[coluna].dtype == object:
            convertida = serie_numeros_br(df[coluna])
            if convertida is not None:
                df[coluna] = convertida
    return df, delimitador_usado, codificacao_usada
