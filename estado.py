"""Modelos de dados e estado da sessão da aplicação.

Mantém em memória os datasets carregados, os resultados nomeados gerados
pelas operações (tabelas dinâmicas, PROCV, agregações etc.) e o histórico de
operações confirmadas, além da pilha de estados anteriores usada para
desfazer alterações.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd


@dataclass
class HistoricoEntry:
    """Registro de uma operação executada e confirmada pelo usuário."""

    timestamp: datetime
    dataset_id: int
    dataset_nome: str
    operacao: str
    parametros: str
    linhas_antes: int
    linhas_depois: int
    colunas_antes: int
    colunas_depois: int
    avisos: str = ""
    resultado: str = ""

    def como_linha(self) -> dict[str, object]:
        """Converte a entrada em um dicionário pronto para virar linha de DataFrame."""
        return {
            "Data/Hora": self.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
            "Dataset": self.dataset_nome,
            "Operação": self.operacao,
            "Parâmetros": self.parametros,
            "Linhas (antes)": self.linhas_antes,
            "Linhas (depois)": self.linhas_depois,
            "Colunas (antes)": self.colunas_antes,
            "Colunas (depois)": self.colunas_depois,
            "Resultado": self.resultado,
            "Avisos": self.avisos,
        }


@dataclass
class Dataset:
    """Representa um dataset carregado ou derivado durante a sessão."""

    id: int
    nome: str
    arquivo: str
    aba: Optional[str]
    df: pd.DataFrame
    df_original: pd.DataFrame
    origem: str = "importado"
    limite_desfazer: int = 20
    pilha_desfazer: list[pd.DataFrame] = field(default_factory=list)

    def salvar_estado_para_desfazer(self) -> None:
        """Empilha uma cópia do estado atual antes de aplicar uma alteração confirmada."""
        self.pilha_desfazer.append(self.df.copy(deep=True))
        if len(self.pilha_desfazer) > self.limite_desfazer:
            self.pilha_desfazer.pop(0)

    def pode_desfazer(self) -> bool:
        """Indica se há algum estado anterior disponível para desfazer."""
        return len(self.pilha_desfazer) > 0

    def desfazer(self) -> bool:
        """Restaura o estado anterior do dataset. Retorna False se não havia nada a desfazer."""
        if not self.pilha_desfazer:
            return False
        self.df = self.pilha_desfazer.pop()
        return True

    def identificador_exibicao(self) -> str:
        """Retorna uma string amigável identificando o dataset (id, arquivo e aba)."""
        if self.aba:
            return f"[{self.id}] {self.arquivo} :: {self.aba}"
        return f"[{self.id}] {self.arquivo}"


class Sessao:
    """Mantém o estado completo da sessão: datasets, resultados nomeados e histórico."""

    def __init__(self, limite_desfazer: int = 20) -> None:
        self._proximo_id = 1
        self.limite_desfazer = limite_desfazer
        self.datasets: dict[int, Dataset] = {}
        self.resultados: dict[str, pd.DataFrame] = {}
        self.historico: list[HistoricoEntry] = []
        self.arquivos_processados: list[str] = []

    def adicionar_dataset(
        self,
        nome: str,
        arquivo: str,
        aba: Optional[str],
        df: pd.DataFrame,
        origem: str = "importado",
    ) -> Dataset:
        """Registra um novo dataset na sessão e retorna o objeto criado."""
        dataset = Dataset(
            id=self._proximo_id,
            nome=nome,
            arquivo=arquivo,
            aba=aba,
            df=df,
            df_original=df.copy(deep=True),
            origem=origem,
            limite_desfazer=self.limite_desfazer,
        )
        self.datasets[dataset.id] = dataset
        self._proximo_id += 1
        return dataset

    def obter_dataset(self, dataset_id: int) -> Optional[Dataset]:
        """Retorna o dataset pelo identificador, ou None se não existir."""
        return self.datasets.get(dataset_id)

    def listar_datasets(self) -> list[Dataset]:
        """Retorna todos os datasets carregados, ordenados por identificador."""
        return [self.datasets[chave] for chave in sorted(self.datasets)]

    def registrar_historico(self, entrada: HistoricoEntry) -> None:
        """Adiciona uma entrada ao histórico de operações confirmadas."""
        self.historico.append(entrada)

    def adicionar_resultado(self, nome: str, df: pd.DataFrame) -> str:
        """Armazena um resultado nomeado (pivô, PROCV, agregação) para exportação posterior."""
        nome_final = nome
        contador = 1
        while nome_final in self.resultados:
            contador += 1
            nome_final = f"{nome}_{contador}"
        self.resultados[nome_final] = df
        return nome_final
