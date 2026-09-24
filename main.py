"""Ponto de entrada da aplicação: configura o ambiente e conduz o menu principal.

Fluxo geral:
1. Configura o logging (arquivo ``app.log`` e console para erros).
2. Conduz a importação inicial de um ou mais arquivos.
3. Exibe o catálogo de datasets e uma prévia dos dados carregados.
4. Executa o laço do menu principal até o usuário escolher sair.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Callable

from estado import Sessao
from interface import (
    exibir_catalogo_geral,
    exibir_previas_iniciais,
    fluxo_importacao,
    handler_agrupar,
    handler_colunas,
    handler_coluna_calculada,
    handler_concatenar,
    handler_contse,
    handler_contses,
    handler_data,
    handler_desfazer,
    handler_duplicados,
    handler_filtrar,
    handler_historico,
    handler_matematica,
    handler_mediase,
    handler_mediases,
    handler_mesclar,
    handler_nulos,
    handler_ordenar,
    handler_pivot,
    handler_proch,
    handler_procv,
    handler_procx,
    handler_qualidade,
    handler_salvar,
    handler_se,
    handler_se_aninhado,
    handler_somase,
    handler_somases,
    handler_texto,
    handler_visualizar_catalogo,
    handler_visualizar_dados,
    ler_texto,
)

MENU_OPCOES: dict[str, str] = {
    "1": "Visualizar catálogo das colunas",
    "2": "Visualizar dados e resumo estatístico",
    "3": "Filtrar dados",
    "4": "Ordenar dados",
    "5": "Selecionar, renomear ou remover colunas",
    "6": "Criar coluna calculada",
    "7": "Aplicar operações matemáticas",
    "8": "Aplicar operações de texto",
    "9": "Aplicar operações de data",
    "10": "Agrupar e resumir dados",
    "11": "Criar tabela dinâmica",
    "12": "Executar PROCV",
    "13": "Executar PROCH",
    "14": "Executar PROCX",
    "15": "Executar SOMASE",
    "16": "Executar SOMASES",
    "17": "Executar CONT.SE",
    "18": "Executar CONT.SES",
    "19": "Executar MÉDIASE",
    "20": "Executar MÉDIASES",
    "21": "Executar SE",
    "22": "Executar SE aninhado",
    "23": "Tratar valores nulos",
    "24": "Remover valores duplicados",
    "25": "Concatenar datasets",
    "26": "Mesclar datasets",
    "27": "Validar qualidade dos dados",
    "28": "Desfazer última operação",
    "29": "Visualizar histórico de operações",
    "30": "Salvar projeto ou relatório",
    "0": "Sair",
}

DESPACHO: dict[str, Callable[[Sessao], None]] = {
    "1": handler_visualizar_catalogo,
    "2": handler_visualizar_dados,
    "3": handler_filtrar,
    "4": handler_ordenar,
    "5": handler_colunas,
    "6": handler_coluna_calculada,
    "7": handler_matematica,
    "8": handler_texto,
    "9": handler_data,
    "10": handler_agrupar,
    "11": handler_pivot,
    "12": handler_procv,
    "13": handler_proch,
    "14": handler_procx,
    "15": handler_somase,
    "16": handler_somases,
    "17": handler_contse,
    "18": handler_contses,
    "19": handler_mediase,
    "20": handler_mediases,
    "21": handler_se,
    "22": handler_se_aninhado,
    "23": handler_nulos,
    "24": handler_duplicados,
    "25": handler_concatenar,
    "26": handler_mesclar,
    "27": handler_qualidade,
    "28": handler_desfazer,
    "29": handler_historico,
    "30": handler_salvar,
}


def configurar_logging() -> None:
    """Configura o logging da aplicação: arquivo com todos os eventos, console apenas com erros."""
    formato = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    manipulador_arquivo = logging.FileHandler("app.log", encoding="utf-8")
    manipulador_arquivo.setLevel(logging.DEBUG)
    manipulador_arquivo.setFormatter(logging.Formatter(formato))

    manipulador_console = logging.StreamHandler(sys.stderr)
    manipulador_console.setLevel(logging.ERROR)
    manipulador_console.setFormatter(logging.Formatter(formato))

    logging.basicConfig(level=logging.DEBUG, handlers=[manipulador_arquivo, manipulador_console])


def exibir_menu_principal() -> None:
    """Imprime as opções do menu principal."""
    print("\n" + "=" * 60)
    print("MENU PRINCIPAL - Ferramenta de Análise de Planilhas")
    print("=" * 60)
    for chave in sorted(MENU_OPCOES, key=lambda k: (k != "0", int(k))):
        print(f"  {chave}. {MENU_OPCOES[chave]}")


def executar_menu_principal(sessao: Sessao) -> None:
    """Executa o laço do menu principal até o usuário escolher a opção de sair."""
    while True:
        exibir_menu_principal()
        try:
            escolha = ler_texto("\nEscolha uma opção: ")
        except EOFError:
            print("\nEntrada encerrada. Saindo da aplicação.")
            break
        if escolha == "0":
            if ler_texto("Deseja sair sem salvar o relatório? Digite 'sim' para confirmar: ", padrao="nao").lower() in ("sim", "s"):
                print("Encerrando a aplicação. Até logo!")
                break
            continue
        handler = DESPACHO.get(escolha)
        if handler is None:
            print("Opção inválida. Tente novamente.")
            continue
        try:
            handler(sessao)
        except (KeyboardInterrupt, EOFError):
            raise
        except Exception as exc:  # noqa: BLE001 - qualquer erro inesperado não deve encerrar o programa
            print(f"Ocorreu um erro inesperado ao executar a operação: {exc}")
            logging.getLogger(__name__).exception("Erro inesperado ao executar a opção '%s'", escolha)


def main() -> None:
    """Ponto de entrada da aplicação."""
    configurar_logging()
    logger = logging.getLogger(__name__)
    logger.info("Aplicação iniciada.")

    print("=" * 60)
    print("Ferramenta de Análise de Planilhas em Python")
    print("Formatos suportados: XLSX, XLS, XLSM, CSV")
    print("=" * 60)

    sessao = Sessao(limite_desfazer=20)

    try:
        fluxo_importacao(sessao)
    except KeyboardInterrupt:
        print("\nImportação interrompida pelo usuário. Encerrando.")
        return
    except EOFError:
        print("\nEntrada encerrada durante a importação. Encerrando.")
        return

    if not sessao.datasets:
        print("Nenhum dataset foi carregado. Encerrando a aplicação.")
        return

    print("\n=== Catálogo dos datasets carregados ===")
    exibir_catalogo_geral(sessao)
    print("\n=== Prévia dos dados carregados ===")
    exibir_previas_iniciais(sessao)

    try:
        executar_menu_principal(sessao)
    except KeyboardInterrupt:
        print("\nOperação interrompida pelo usuário. Encerrando a aplicação.")
    except EOFError:
        print("\nEntrada encerrada. Encerrando a aplicação.")
    except Exception:  # noqa: BLE001 - último nível de proteção contra encerramento abrupto
        logger.exception("Erro fatal não tratado na aplicação.")
        print("Ocorreu um erro inesperado e a aplicação será encerrada. Consulte 'app.log' para detalhes.")


if __name__ == "__main__":
    main()
