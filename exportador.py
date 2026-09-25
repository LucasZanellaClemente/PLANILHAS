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
                "Resultado",
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


def descrever_parametros(parametros: str) -> str:
    """Deixa os parâmetros gravados no histórico legíveis: ``colunas=['A', 'B']`` vira ``colunas: A, B``."""
    texto = re.sub(r"\bNone\b", "", str(parametros))
    texto = re.sub(r"[\[\]'\"]", "", texto)
    texto = re.sub(r"(\w+)=", r"\1: ", texto)
    texto = re.sub(r",\s*\)", ")", texto)
    texto = re.sub(r"\b\w+:\s*(?=,|$)", "", texto)  # parâmetros vazios
    texto = re.sub(r"(,\s*){2,}", ", ", texto)
    return re.sub(r"\s{2,}", " ", texto).strip(" ,")


def _dataset_foi_alterado(dataset: Any) -> bool:
    return dataset.origem == "derivado" or not dataset.df.equals(dataset.df_original)


def montar_operacoes_executadas(sessao: Sessao) -> pd.DataFrame:
    """Lista, em ordem, as opções executadas e confirmadas na sessão, em linguagem simples."""
    linhas = []
    for numero, entrada in enumerate(sessao.historico, start=1):
        if entrada.linhas_antes or entrada.colunas_antes:
            linhas_txt = f"{entrada.linhas_antes} → {entrada.linhas_depois}"
            colunas_txt = f"{entrada.colunas_antes} → {entrada.colunas_depois}"
        else:
            # Resultado salvo ou dataset novo: os dados de origem não mudaram.
            linhas_txt = colunas_txt = "-"
        linhas.append(
            {
                "#": numero,
                "Hora": entrada.timestamp.strftime("%H:%M:%S"),
                "Opção executada": entrada.operacao,
                "Dataset": entrada.dataset_nome,
                "O que foi feito": descrever_parametros(entrada.parametros),
                "Resultado": entrada.resultado,
                "Linhas": linhas_txt,
                "Colunas": colunas_txt,
                "Avisos": entrada.avisos,
            }
        )
    colunas = ["#", "Hora", "Opção executada", "Dataset", "O que foi feito", "Resultado", "Linhas", "Colunas", "Avisos"]
    return pd.DataFrame(linhas, columns=colunas)


def montar_resumo_sessao(sessao: Sessao) -> pd.DataFrame:
    """Resumo curto do que foi feito na sessão, usado no relatório resumido e na tela."""
    alterados = [d for d in sessao.listar_datasets() if _dataset_foi_alterado(d)]
    contagem: dict[str, int] = {}
    for entrada in sessao.historico:
        contagem[entrada.operacao] = contagem.get(entrada.operacao, 0) + 1
    metricas = [
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Arquivos importados", ", ".join(sorted(set(sessao.arquivos_processados))) or "-"),
        ("Operações executadas", len(sessao.historico)),
        ("Opções usadas", ", ".join(f"{nome} ({qtd}x)" for nome, qtd in contagem.items()) or "-"),
        (
            "Datasets alterados ou criados",
            ", ".join(f"{d.identificador_exibicao()} ({len(d.df)} linhas)" for d in alterados) or "nenhum",
        ),
        ("Resultados salvos", ", ".join(sessao.resultados) or "nenhum"),
    ]
    return pd.DataFrame(metricas, columns=["Item", "Valor"])


def _gerar_relatorio_resumido(sessao: Sessao, caminho_saida: Path) -> Path:
    """Relatório prático: resumo, opções executadas e só os dados que mudaram ou foram calculados."""
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter") as writer:
        construtor = _ConstrutorRelatorio(writer)
        construtor.escrever_dataframe("Resumo", montar_resumo_sessao(sessao), incluir_tabela_excel=False)
        construtor.escrever_dataframe("Operacoes_Executadas", montar_operacoes_executadas(sessao))
        alterados = [d for d in sessao.listar_datasets() if _dataset_foi_alterado(d)]
        bases = [_nome_base_dataset(d) for d in alterados]
        for dataset, nome_base in zip(alterados, bases):
            if bases.count(nome_base) > 1:
                nome_base = f"{dataset.id}_{nome_base}"
            construtor.escrever_dataframe(nome_base, dataset.df)
        for nome_resultado, df_resultado in sessao.resultados.items():
            construtor.escrever_dataframe(nome_resultado, df_resultado)
    logger.info("Relatório resumido gerado em: %s", caminho_saida)
    return caminho_saida


def gerar_relatorio(sessao: Sessao, caminho_saida: Path, completo: bool = False) -> Path:
    """Gera o relatório em Excel e retorna o caminho final.

    O padrão é o relatório resumido (resumo, opções executadas, datasets
    alterados e resultados). Com ``completo=True`` inclui também catálogo,
    qualidade dos dados e as versões original e alterada de todos os datasets.
    """
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    if not completo:
        return _gerar_relatorio_resumido(sessao, caminho_saida)

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
