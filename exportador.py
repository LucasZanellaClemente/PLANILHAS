"""Geração do relatório final em Excel (``relatorio.xlsx``).

Constrói um arquivo com abas de resumo, catálogo de colunas, qualidade de
dados, histórico de operações, uma aba por dataset (original e, quando
alterado, também a versão alterada) e uma aba para cada resultado nomeado
(tabelas dinâmicas, buscas e agregações). Usa ``xlsxwriter`` para aplicar
formatação: cabeçalhos destacados, autofiltro, congelamento da primeira
linha, ajuste de largura de coluna, formatos numéricos/data/percentual e
tabelas nativas do Excel.
"""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from catalogo import montar_catalogo_colunas, montar_resumo_dataset
from estado import Sessao
from operacoes import validar_qualidade_dados
from utils import nome_aba_valido, sanitizar_dataframe_formulas

logger = logging.getLogger(__name__)

LARGURA_MAXIMA_COLUNA = 50
LARGURA_MINIMA_COLUNA = 10


def gerar_nome_com_timestamp(caminho: Path) -> Path:
    """Gera um novo nome de arquivo incluindo data e hora, evitando sobrescrever um existente."""
    agora = datetime.now().strftime("%Y%m%d_%H%M%S")
    return caminho.with_name(f"{caminho.stem}_{agora}{caminho.suffix}")


class _ConstrutorRelatorio:
    """Agrupa o estado necessário durante a montagem do arquivo Excel (nomes de aba já usados etc.)."""

    def __init__(self, writer: pd.ExcelWriter) -> None:
        self.writer = writer
        self.workbook = writer.book
        self.abas_existentes: set[str] = set()
        self.formato_cabecalho = self.workbook.add_format(
            {
                "bold": True,
                "bg_color": "#1F4E78",
                "font_color": "#FFFFFF",
                "border": 1,
                "text_wrap": True,
                "valign": "vcenter",
            }
        )
        self.formato_numero = self.workbook.add_format({"num_format": "#,##0.00"})
        self.formato_inteiro = self.workbook.add_format({"num_format": "#,##0"})
        self.formato_data = self.workbook.add_format({"num_format": "dd/mm/yyyy"})
        self.formato_percentual = self.workbook.add_format({"num_format": '0.00"%"'})

    def novo_nome_aba(self, nome_sugerido: str) -> str:
        """Retorna um nome de aba válido e único, registrando-o como já utilizado."""
        return nome_aba_valido(nome_sugerido, self.abas_existentes)

    def escrever_dataframe(self, nome_sugerido: str, df: pd.DataFrame, incluir_tabela_excel: bool = True) -> str:
        """Escreve um DataFrame em uma nova aba, aplicando toda a formatação padrão."""
        nome_aba = self.novo_nome_aba(nome_sugerido)
        df_seguro = sanitizar_dataframe_formulas(df) if not df.empty else df
        df_seguro.to_excel(self.writer, sheet_name=nome_aba, index=False)
        planilha = self.writer.sheets[nome_aba]

        for indice_coluna, nome_coluna in enumerate(df_seguro.columns):
            planilha.write(0, indice_coluna, str(nome_coluna), self.formato_cabecalho)

        self._formatar_colunas(planilha, df_seguro)
        planilha.freeze_panes(1, 0)

        linhas, colunas = df_seguro.shape
        if colunas > 0:
            if linhas > 0 and incluir_tabela_excel:
                # add_table já aplica seu próprio autofiltro; chamá-lo de novo geraria conflito.
                self._adicionar_tabela_excel(planilha, nome_aba, df_seguro)
            else:
                planilha.autofilter(0, 0, max(linhas, 0), colunas - 1)
        return nome_aba

    def _formatar_colunas(self, planilha: Any, df: pd.DataFrame) -> None:
        """Ajusta largura e formato numérico/data/percentual de cada coluna."""
        for indice_coluna, nome_coluna in enumerate(df.columns):
            serie = df[nome_coluna]
            tamanho_cabecalho = len(str(nome_coluna))
            if len(serie) > 0:
                # Vazios contam como 0: no pandas 3 o astype(str) mantém NaN e len() falharia.
                tamanho_valores = serie.map(lambda valor: 0 if pd.isna(valor) else len(str(valor))).max()
            else:
                tamanho_valores = 0
            largura = min(max(tamanho_cabecalho, int(tamanho_valores), LARGURA_MINIMA_COLUNA) + 2, LARGURA_MAXIMA_COLUNA)

            formato_coluna = None
            nome_texto = str(nome_coluna).lower()
            if pd.api.types.is_datetime64_any_dtype(serie):
                formato_coluna = self.formato_data
            elif "%" in nome_texto or "percentual" in nome_texto:
                formato_coluna = self.formato_percentual
            elif pd.api.types.is_float_dtype(serie):
                formato_coluna = self.formato_numero
            elif pd.api.types.is_integer_dtype(serie):
                formato_coluna = self.formato_inteiro

            planilha.set_column(indice_coluna, indice_coluna, largura, formato_coluna)

    def _adicionar_tabela_excel(self, planilha: Any, nome_aba: str, df: pd.DataFrame) -> None:
        """Registra o intervalo como uma Tabela nativa do Excel, quando possível."""
        linhas, colunas = df.shape
        try:
            colunas_tabela = [{"header": str(c)} for c in df.columns]
            planilha.add_table(
                0,
                0,
                linhas,
                colunas - 1,
                {
                    "columns": colunas_tabela,
                    "style": "Table Style Medium 9",
                    # Nome de tabela do Excel só aceita letras, números e "_" (a aba "LEIA-ME" falhava).
                    "name": f"Tbl_{re.sub(r'[^0-9A-Za-z_]', '_', nome_aba)[:200]}",
                },
            )
        except Exception as exc:  # noqa: BLE001 - formatação de tabela é best-effort
            logger.warning("Não foi possível criar tabela nativa na aba '%s': %s", nome_aba, exc)


def _montar_resumo_geral(sessao: Sessao) -> pd.DataFrame:
    """Monta o DataFrame de métricas gerais exibido na aba Resumo."""
    datasets = sessao.listar_datasets()
    total_linhas = sum(len(d.df) for d in datasets)
    total_colunas = sum(d.df.shape[1] for d in datasets)
    total_nulos = sum(int(d.df.isna().sum().sum()) for d in datasets)
    total_duplicidades = sum(int(d.df.duplicated().sum()) for d in datasets)
    abas_processadas = [d.aba for d in datasets if d.aba]

    metricas = [
        ("Data/hora de geração", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Arquivos processados", ", ".join(sorted(set(sessao.arquivos_processados))) or "-"),
        ("Quantidade de arquivos processados", len(set(sessao.arquivos_processados))),
        ("Abas processadas", ", ".join(str(a) for a in abas_processadas) or "-"),
        ("Quantidade de datasets carregados", len(datasets)),
        ("Total de linhas (todos os datasets)", total_linhas),
        ("Total de colunas (todos os datasets)", total_colunas),
        ("Total de valores nulos", total_nulos),
        ("Total de linhas duplicadas", total_duplicidades),
        ("Operações realizadas e confirmadas", len(sessao.historico)),
    ]
    return pd.DataFrame(metricas, columns=["Métrica", "Valor"])


def _montar_catalogo_colunas_geral(sessao: Sessao) -> pd.DataFrame:
    """Monta o DataFrame consolidado do catálogo de colunas de todos os datasets."""
    linhas = []
    for dataset in sessao.listar_datasets():
        for item in montar_catalogo_colunas(dataset.df):
            linha = {"Dataset": dataset.identificador_exibicao()}
            linha.update(
                {
                    "Posição": item["posicao"],
                    "Coluna": item["coluna"],
                    "Tipo": item["tipo"],
                    "Categoria": item["categoria"],
                    "Preenchidos": item["preenchidos"],
                    "Nulos": item["nulos"],
                    "% Nulos": item["percentual_nulos"],
                    "Únicos": item["unicos"],
                    "Exemplos": item["exemplos"],
                }
            )
            linhas.append(linha)
    return pd.DataFrame(linhas)


def _montar_qualidade_geral(sessao: Sessao) -> pd.DataFrame:
    """Monta o DataFrame consolidado de qualidade de dados de todos os datasets."""
    linhas = []
    for dataset in sessao.listar_datasets():
        qualidade = validar_qualidade_dados(dataset.df)
        linhas.append(
            {
                "Dataset": dataset.identificador_exibicao(),
                "Linhas": qualidade["total_linhas"],
                "Colunas": qualidade["total_colunas"],
                "Total de nulos": qualidade["total_nulos"],
                "% Nulos (geral)": qualidade["percentual_nulos"],
                "Linhas duplicadas": qualidade["linhas_duplicadas"],
                "Colunas totalmente nulas": ", ".join(qualidade["colunas_totalmente_nulas"]) or "-",
                "Colunas constantes": ", ".join(qualidade["colunas_constantes"]) or "-",
            }
        )
    return pd.DataFrame(linhas)


def _montar_historico(sessao: Sessao) -> pd.DataFrame:
    """Monta o DataFrame do histórico de operações confirmadas."""
    if not sessao.historico:
        return pd.DataFrame(
            columns=[
                "Data/Hora",
                "Dataset",
                "Operação",
                "Parâmetros",
                "Linhas (antes)",
                "Linhas (depois)",
                "Colunas (antes)",
                "Colunas (depois)",
                "Avisos",
            ]
        )
    return pd.DataFrame([entrada.como_linha() for entrada in sessao.historico])


def _nome_base_dataset(dataset: Any) -> str:
    """Nome curto que identifica o dataset numa aba: a aba de origem ou, no CSV, o nome do arquivo."""
    return str(dataset.aba) if dataset.aba else Path(dataset.arquivo).stem


def _nome_aba_com_sufixo(nome_base: str, sufixo: str) -> str:
    """Encurta o nome base para que o sufixo (``_orig``/``_alt``) caiba nos 31 caracteres do Excel."""
    return f"{nome_base[: 31 - len(sufixo)]}{sufixo}"


def gerar_relatorio(sessao: Sessao, caminho_saida: Path) -> Path:
    """Gera o arquivo ``relatorio.xlsx`` com todas as abas exigidas e retorna o caminho final."""
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        construtor = _ConstrutorRelatorio(writer)

        construtor.escrever_dataframe("Resumo", _montar_resumo_geral(sessao), incluir_tabela_excel=False)
        construtor.escrever_dataframe("Catalogo_Colunas", _montar_catalogo_colunas_geral(sessao))
        construtor.escrever_dataframe("Qualidade_Dados", _montar_qualidade_geral(sessao))
        construtor.escrever_dataframe("Historico", _montar_historico(sessao))

        datasets = sessao.listar_datasets()
        bases = [_nome_base_dataset(d) for d in datasets]
        for dataset, nome_base in zip(datasets, bases):
            if bases.count(nome_base) > 1:
                nome_base = f"{dataset.id}_{nome_base}"
            construtor.escrever_dataframe(_nome_aba_com_sufixo(nome_base, "_orig"), dataset.df_original)
            if not dataset.df.equals(dataset.df_original):
                construtor.escrever_dataframe(_nome_aba_com_sufixo(nome_base, "_alt"), dataset.df)

        for nome_resultado, df_resultado in sessao.resultados.items():
            construtor.escrever_dataframe(nome_resultado, df_resultado)

    logger.info("Relatório gerado em: %s", caminho_saida)
    return caminho_saida
