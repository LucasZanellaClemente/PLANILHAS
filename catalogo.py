"""Construção do catálogo de datasets e colunas.

Gera as estruturas de metadados exibidas ao usuário antes e durante o uso do
menu principal: tipo de cada coluna, contagem de nulos/únicos, colunas-chave
sugeridas e indicação de duplicidades.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from utils import eh_coluna_texto


def identificar_tipo_coluna(serie: pd.Series) -> str:
    """Classifica uma coluna como numérica, texto, data, booleana ou outro tipo."""
    if pd.api.types.is_bool_dtype(serie):
        return "booleano"
    if pd.api.types.is_numeric_dtype(serie):
        return "numérico"
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "data"
    if eh_coluna_texto(serie):
        amostra = serie.dropna().astype(str).head(50)
        if len(amostra) > 0:
            convertido = pd.to_datetime(amostra, errors="coerce", format="mixed")
            if convertido.notna().mean() > 0.8:
                return "data (texto)"
        return "texto"
    return str(serie.dtype)


def categoria_coluna(tipo: str) -> str:
    """Agrupa o tipo detalhado em uma categoria simples: numérica, textual ou data."""
    if tipo == "numérico":
        return "numérica"
    if tipo in ("data", "data (texto)"):
        return "data"
    if tipo == "booleano":
        return "booleana"
    return "textual"


def detectar_colunas_chave(df: pd.DataFrame) -> list[str]:
    """Sugere colunas-chave: sem valores nulos e 100% únicas."""
    if len(df) == 0:
        return []
    chaves = []
    for coluna in df.columns:
        serie = df[coluna]
        if serie.isna().sum() == 0 and serie.is_unique:
            chaves.append(str(coluna))
    return chaves


def montar_catalogo_colunas(df: pd.DataFrame) -> list[dict[str, Any]]:
    """Monta uma lista de dicionários com metadados detalhados de cada coluna."""
    total_linhas = len(df)
    catalogo = []
    for posicao, coluna in enumerate(df.columns, start=1):
        serie = df[coluna]
        preenchidos = int(serie.notna().sum())
        nulos = int(serie.isna().sum())
        percentual_nulos = round((nulos / total_linhas * 100) if total_linhas else 0.0, 2)
        exemplos = serie.dropna().unique()[:3]
        tipo = identificar_tipo_coluna(serie)
        catalogo.append(
            {
                "posicao": posicao,
                "coluna": str(coluna),
                "tipo": tipo,
                "categoria": categoria_coluna(tipo),
                "preenchidos": preenchidos,
                "nulos": nulos,
                "percentual_nulos": percentual_nulos,
                "unicos": int(serie.nunique(dropna=True)),
                "exemplos": ", ".join(str(v) for v in exemplos),
            }
        )
    return catalogo


def montar_resumo_dataset(dataset_id: int, arquivo: str, aba: str | None, df: pd.DataFrame) -> dict[str, Any]:
    """Monta um resumo de alto nível de um dataset para o catálogo geral."""
    return {
        "id": dataset_id,
        "arquivo": arquivo,
        "aba": aba or "-",
        "linhas": len(df),
        "colunas": df.shape[1],
        "duplicadas": int(df.duplicated().sum()),
        "colunas_chave": ", ".join(detectar_colunas_chave(df)) or "nenhuma identificada",
    }


def montar_relatorio_qualidade(df: pd.DataFrame) -> pd.DataFrame:
    """Monta um DataFrame com indicadores de qualidade por coluna, usado na validação e no export."""
    total_linhas = len(df)
    linhas = []
    for coluna in df.columns:
        serie = df[coluna]
        nulos = int(serie.isna().sum())
        tipo = identificar_tipo_coluna(serie)
        linhas.append(
            {
                "Coluna": str(coluna),
                "Tipo detectado": tipo,
                "Preenchidos": int(serie.notna().sum()),
                "Nulos": nulos,
                "% Nulos": round((nulos / total_linhas * 100) if total_linhas else 0.0, 2),
                "Valores únicos": int(serie.nunique(dropna=True)),
                "Duplicados (linha inteira)": int(df.duplicated().sum()) if coluna == df.columns[0] else "",
            }
        )
    return pd.DataFrame(linhas)
